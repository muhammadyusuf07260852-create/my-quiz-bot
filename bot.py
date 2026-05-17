import os
import telebot
from telebot.types import ReplyKeyboardMarkup, KeyboardButton, KeyboardButtonPollType, ReplyKeyboardRemove, InlineKeyboardMarkup, InlineKeyboardButton
from dotenv import load_dotenv
import sqlite3
import database
import json
import threading
import random
import docx
import pypdf
from pptx import Presentation
import google.generativeai as genai

def extract_text_from_pptx(file_path):
    try:
        prs = Presentation(file_path)
        text_runs = []
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for paragraph in shape.text_frame.paragraphs:
                        for run in paragraph.runs:
                            if run.text:
                                text_runs.append(run.text)
                elif shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            if cell.text:
                                text_runs.append(cell.text)
        return "\n".join(text_runs).strip()
    except Exception as e:
        print(f"PPTX extract error: {e}")
        return ""

load_dotenv()
BOT_TOKEN = os.getenv('BOT_TOKEN')
bot = telebot.TeleBot(BOT_TOKEN)

# Gemini AI konfiguratsiyasi
GEMINI_API_KEY = os.getenv('GEMINI_API_KEY')
if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

# AI vaqtinchalik sessiyalari
ai_sessions = {}

def extract_text_from_pdf(file_path):
    try:
        reader = pypdf.PdfReader(file_path)
        text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
        return text.strip()
    except Exception as e:
        print(f"PDF extract error: {e}")
        return ""

def extract_text_from_docx(file_path):
    try:
        doc = docx.Document(file_path)
        text = ""
        for para in doc.paragraphs:
            if para.text:
                text += para.text + "\n"
        return text.strip()
    except Exception as e:
        print(f"Docx extract error: {e}")
        return ""

bot_info = bot.get_me()

# Telegram botda buyruqlar menyusini o'rnatish (/ bosganda chiqadi)
bot.set_my_commands([
    telebot.types.BotCommand('/start',        '🏠 Bosh menyu'),
    telebot.types.BotCommand('/newquiz',      '➕ Yangi test yaratish'),
    telebot.types.BotCommand('/add_questions','📝 Testga savol qo\'shish'),
    telebot.types.BotCommand('/stop',         '🛑 Jarayonni to\'xtatish'),
])

database.init_db()

# Xotirada saqlash
user_sessions = {} # Test yaratish uchun
taking_quiz_sessions = {} # Yakkaxon test yechish uchun
group_sessions = {} # Guruhda test yechish uchun

def get_current_quiz_id(user_id):
    session = user_sessions.get(user_id)
    if session and 'quiz_id' in session:
        return session['quiz_id']
    conn = sqlite3.connect(database.DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT quiz_id FROM quizzes WHERE creator_id = ? ORDER BY quiz_id DESC LIMIT 1", (user_id,))
    res = cursor.fetchone()
    conn.close()
    if res:
        if user_id not in user_sessions:
            user_sessions[user_id] = {}
        user_sessions[user_id]['quiz_id'] = res[0]
        return res[0]
    return None

def get_question_keyboard():
    markup = ReplyKeyboardMarkup(resize_keyboard=True, row_width=1)
    btn_poll = KeyboardButton('Savol yaratish', request_poll=KeyboardButtonPollType(type='quiz'))
    btn_done = KeyboardButton('Tugatish')
    markup.add(btn_poll, btn_done)
    return markup

def get_main_keyboard():
    markup = ReplyKeyboardMarkup(resize_keyboard=True, row_width=2)
    btn_new = KeyboardButton('➕ Yangi test yaratish')
    btn_my = KeyboardButton('📂 Mening testlarim')
    btn_results = KeyboardButton('📊 Natijalar')
    markup.add(btn_new, btn_my, btn_results)
    return markup

@bot.message_handler(commands=['start'])
def send_welcome(message):
    chat_type = message.chat.type
    text = message.text.split()
    payload = text[1] if len(text) > 1 else None

    # Agar guruhda bo'lsa
    if chat_type in ['group', 'supergroup']:
        if payload and payload.startswith('quiz_'):
            try:
                quiz_id = int(payload.split('_')[1])
                init_group_quiz(message.chat.id, quiz_id)
            except (IndexError, ValueError):
                bot.send_message(message.chat.id, "Xato link formatı.")
        return

    # Shaxsiy chat (Private)
    user_id = message.from_user.id
    username = message.from_user.username
    first_name = message.from_user.first_name
    
    conn = sqlite3.connect(database.DB_PATH)
    cursor = conn.cursor()
    cursor.execute("INSERT OR IGNORE INTO users (user_id, username, first_name) VALUES (?, ?, ?)", 
                   (user_id, username, first_name))
    conn.commit()
    conn.close()

    # Deep linking (Yakkaxon test yechish)
    if payload and payload.startswith('quiz_'):
        try:
            quiz_id = int(payload.split('_')[1])
            start_taking_quiz(message, quiz_id)
        except (IndexError, ValueError):
            bot.reply_to(message, "Xato link formati.")
        return

    database.set_user_state(user_id, 'none')
    
    welcome_text = (
        f"Assalomu alaykum, {first_name}! 👋\n\n"
        "Men Telegram Quiz Botman. Men yordamida siz turli xil testlar va viktorinalar yaratishingiz mumkin.\n\n"
        "Quyidagi menyudan kerakli bo'limni tanlang:"
    )
    bot.reply_to(message, welcome_text, reply_markup=get_main_keyboard())

def prepare_questions(quiz_id):
    conn = sqlite3.connect(database.DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT shuffle_mode FROM quizzes WHERE quiz_id = ?", (quiz_id,))
    res = cursor.fetchone()
    shuffle_mode = res[0] if res else 0

    cursor.execute("SELECT question_id, question_text, options, correct_option_id, explanation, time_limit FROM questions WHERE quiz_id = ?", (quiz_id,))
    questions = cursor.fetchall()
    conn.close()
    
    questions = [list(q) for q in questions]

    if shuffle_mode in [1, 3]: # Savollarni aralashtirish
        random.shuffle(questions)
        
    for i in range(len(questions)):
        q_id, q_text, options_json, correct_id, exp, time_limit = questions[i]
        options = json.loads(options_json)
        
        if shuffle_mode in [2, 3]: # Variantlarni aralashtirish
            correct_text = options[correct_id]
            random.shuffle(options)
            new_correct_id = options.index(correct_text)
            
            questions[i][2] = json.dumps(options)
            questions[i][3] = new_correct_id
            
    return questions

# --- YAKKAXON TEST YECHISH ---
def start_taking_quiz(message, quiz_id):
    user_id = message.from_user.id
    conn = sqlite3.connect(database.DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT title, description FROM quizzes WHERE quiz_id = ?", (quiz_id,))
    quiz = cursor.fetchone()
    conn.close()
    if not quiz:
        all_q = database.get_all_quizzes()
        all_ids = [str(q[0]) for q in all_q]
        bot.reply_to(message, f"Kechirasiz, bunday test (ID: {quiz_id}) topilmadi.\nBazadagi mavjud testlar IDlari: {', '.join(all_ids)}")
        return
        
    questions = prepare_questions(quiz_id)
    
    if not questions:
        bot.reply_to(message, "Bu testda hali savollar yo'q.")
        return
        
    title, desc = quiz
    desc_text = f"\nIzoh: {desc}" if desc else ""
    bot.reply_to(message, f"🎯 **{title}** testini boshlaymiz!{desc_text}\n\nJami savollar: {len(questions)}\nOmad!", parse_mode='Markdown')
    
    taking_quiz_sessions[user_id] = {
        'quiz_id': quiz_id,
        'questions': questions,
        'current_q_index': 0,
        'score': 0
    }
    send_next_question(user_id, message.chat.id)

def send_next_question(user_id, chat_id):
    session = taking_quiz_sessions.get(user_id)
    if not session: return
        
    idx = session['current_q_index']
    questions = session['questions']
    
    if idx < len(questions):
        q = questions[idx]
        q_id, q_text, options_json, correct_id, exp, time_limit = q
        options = json.loads(options_json)
        
        t_limit = time_limit if time_limit else 15

        msg = bot.send_poll(
            chat_id, 
            question=f"[{idx+1}/{len(questions)}] {q_text}", 
            options=options, 
            type='quiz', 
            correct_option_id=correct_id,
            explanation=exp,
            is_anonymous=False,
            open_period=t_limit
        )
        session['current_poll_id'] = msg.poll.id
        session['correct_option_id'] = correct_id
        session['chat_id'] = chat_id
        
        # Vaqt tugaguncha javob bermasa avtomatik o'tish uchun taymer
        if 'timer' in session and session['timer']:
            session['timer'].cancel()
            
        session['timer'] = threading.Timer(t_limit + 1.0, handle_question_timeout, args=[user_id, idx])
        session['timer'].start()
    else:
        score = session['score']
        total = len(questions)
        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO results (user_id, quiz_id, score) VALUES (?, ?, ?)", (user_id, session['quiz_id'], score))
        conn.commit()
        conn.close()
        
        markup = InlineKeyboardMarkup()
        start_url = f"https://t.me/{bot_info.username}?start=quiz_{session['quiz_id']}"
        group_url = f"https://t.me/{bot_info.username}?startgroup=quiz_{session['quiz_id']}"
        markup.add(InlineKeyboardButton("Yakkaxon yechish 👤", url=start_url))
        markup.add(InlineKeyboardButton("Guruhda yechish 👥", url=group_url))
        markup.add(InlineKeyboardButton("Do'stlarga ulashish ↗️", url=f"https://t.me/share/url?url={start_url}"))

        bot.send_message(chat_id, f"🏁 Test yakunlandi!\n\nSizning natijangiz: {score} / {total}\nBarakalla! 🎉", reply_markup=markup)
        del taking_quiz_sessions[user_id]

def handle_question_timeout(user_id, q_index):
    session = taking_quiz_sessions.get(user_id)
    if session and session['current_q_index'] == q_index:
        session['current_q_index'] += 1
        send_next_question(user_id, session['chat_id'])

# --- GURUHDA TEST YECHISH ---
def init_group_quiz(chat_id, quiz_id):
    conn = sqlite3.connect(database.DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT title FROM quizzes WHERE quiz_id = ?", (quiz_id,))
    quiz = cursor.fetchone()
    conn.close()
    
    if not quiz:
        all_q = database.get_all_quizzes()
        all_ids = [str(q[0]) for q in all_q]
        bot.send_message(chat_id, f"Kechirasiz, bunday test (ID: {quiz_id}) topilmadi.\nBazadagi mavjud testlar IDlari: {', '.join(all_ids)}")
        return
        
    questions = prepare_questions(quiz_id)
    
    if not questions:
        bot.send_message(chat_id, "Bu testda hali savollar yo'q.")
        return
        
    group_sessions[chat_id] = {
        'quiz_id': quiz_id,
        'title': quiz[0],
        'questions': questions,
        'current_q_index': 0,
        'participants': {},
        'status': 'waiting',
        'current_poll_id': None,
        'correct_option_id': None
    }
    
    markup = InlineKeyboardMarkup()
    markup.add(InlineKeyboardButton("Men tayyorman! (0)", callback_data="ready"))
    markup.add(InlineKeyboardButton("▶️ Boshlash", callback_data="start_group_quiz"))
    
    msg = bot.send_message(chat_id, f"📢 **{quiz[0]}** testi boshlanmoqda!\n\nQatnashish uchun pastdagi tugmani bosing.", reply_markup=markup, parse_mode='Markdown')
    group_sessions[chat_id]['message_id'] = msg.message_id

@bot.callback_query_handler(func=lambda call: call.data in ["ready", "start_group_quiz"])
def handle_group_callbacks(call):
    chat_id = call.message.chat.id
    session = group_sessions.get(chat_id)
    if not session or session['status'] != 'waiting':
        bot.answer_callback_query(call.id, "Test hozir bu holatda emas.")
        return
        
    if call.data == "ready":
        user_id = call.from_user.id
        first_name = call.from_user.first_name
        
        if user_id not in session['participants']:
            session['participants'][user_id] = {'first_name': first_name, 'score': 0}
            count = len(session['participants'])
            
            markup = InlineKeyboardMarkup()
            markup.add(InlineKeyboardButton(f"Men tayyorman! ({count})", callback_data="ready"))
            markup.add(InlineKeyboardButton("▶️ Boshlash", callback_data="start_group_quiz"))
            
            bot.edit_message_reply_markup(chat_id, call.message.message_id, reply_markup=markup)
            bot.answer_callback_query(call.id, "Siz ro'yxatga olindingiz!")
        else:
            bot.answer_callback_query(call.id, "Siz allaqachon ro'yxatdan o'tgansiz.")
            
    elif call.data == "start_group_quiz":
        if len(session['participants']) == 0:
            session['participants'][call.from_user.id] = {'first_name': call.from_user.first_name, 'score': 0}
            
        session['status'] = 'playing'
        bot.edit_message_text(f"🚀 **{session['title']}** testi boshlandi!\nDiqqat qiling, savollar kelmoqda...", chat_id=chat_id, message_id=session['message_id'], parse_mode='Markdown')
        bot.answer_callback_query(call.id, "Test boshlandi!")
        
        threading.Timer(3.0, send_next_group_question, args=[chat_id]).start()

def send_next_group_question(chat_id):
    session = group_sessions.get(chat_id)
    if not session or session['status'] != 'playing': return
    
    idx = session['current_q_index']
    questions = session['questions']
    
    if idx < len(questions):
        q = questions[idx]
        q_id, q_text, options_json, correct_id, exp, time_limit = q
        options = json.loads(options_json)
        
        t_limit = time_limit if time_limit else 15

        msg = bot.send_poll(
            chat_id, 
            question=f"[{idx+1}/{len(questions)}] {q_text}", 
            options=options, 
            type='quiz', 
            correct_option_id=correct_id,
            explanation=exp,
            is_anonymous=False,
            open_period=t_limit
        )
        session['current_poll_id'] = msg.poll.id
        session['correct_option_id'] = correct_id
        session['current_q_index'] += 1
        
        threading.Timer(t_limit + 1.0, send_next_group_question, args=[chat_id]).start()
    else:
        session['status'] = 'finished'
        participants = session['participants'].values()
        sorted_p = sorted(participants, key=lambda x: x['score'], reverse=True)
        
        text = f"🏁 **{session['title']}** testi yakunlandi!\n\n🏆 **NATIJALAR (Leaderboard):**\n\n"
        for i, p in enumerate(sorted_p):
            medal = "🥇" if i == 0 else "🥈" if i == 1 else "🥉" if i == 2 else "🔹"
            text += f"{medal} {p['first_name']} - {p['score']} ta to'g'ri\n"
            
        markup = InlineKeyboardMarkup()
        start_url = f"https://t.me/{bot_info.username}?start=quiz_{session['quiz_id']}"
        group_url = f"https://t.me/{bot_info.username}?startgroup=quiz_{session['quiz_id']}"
        markup.add(InlineKeyboardButton("Yakkaxon yechish 👤", url=start_url))
        markup.add(InlineKeyboardButton("Guruhda yechish 👥", url=group_url))
        markup.add(InlineKeyboardButton("Do'stlarga ulashish ↗️", url=f"https://t.me/share/url?url={start_url}"))

        bot.send_message(chat_id, text, parse_mode='Markdown', reply_markup=markup)
        del group_sessions[chat_id]


# --- POLL JAVOBLARINI USHLASH (Yakkaxon va Guruh) ---
@bot.poll_answer_handler()
def handle_poll_answer(poll_answer):
    user_id = poll_answer.user.id
    poll_id = poll_answer.poll_id
    selected_option = poll_answer.option_ids[0] if poll_answer.option_ids else None
    
    # 1. Yakkaxon testni tekshirish
    session = taking_quiz_sessions.get(user_id)
    if session and session.get('current_poll_id') == poll_id:
        if 'timer' in session and session['timer']:
            session['timer'].cancel()
            
        if selected_option == session['correct_option_id']:
            session['score'] += 1
            
        session['current_q_index'] += 1
        send_next_question(user_id, session['chat_id'])
        return
        
    # 2. Guruh testini tekshirish
    for chat_id, g_session in group_sessions.items():
        if g_session.get('current_poll_id') == poll_id:
            if user_id not in g_session['participants']:
                g_session['participants'][user_id] = {'first_name': poll_answer.user.first_name, 'score': 0}

            if selected_option == g_session['correct_option_id']:
                g_session['participants'][user_id]['score'] += 1
            break


# --- TEST YARATISH ---
@bot.message_handler(commands=['newquiz'])
def create_new_quiz(message):
    if message.chat.type in ['group', 'supergroup']:
        bot.reply_to(message, "Testni faqat shaxsiy chatda yaratish mumkin.")
        return
    user_id = message.from_user.id
    database.set_user_state(user_id, 'waiting_for_title')
    bot.reply_to(message, "Yangi test yaratamiz! 🎉\n\nIltimos, testning nomini yuboring (masalan: 'Matematika fanidan testlar').", reply_markup=ReplyKeyboardRemove())

@bot.message_handler(commands=['fix_my_id'])
def fix_id(message):
    user_id = message.from_user.id
    # Bazadagi Muhammadyusuf ID-si: 6559589296
    old_id = 6559589296
    database.fix_user_quizzes(old_id, user_id)
    bot.reply_to(message, f"✅ Barcha eski testlaringiz yangi ID-ingizga ({user_id}) biriktirildi! Endi 'Mening testlarim' bo'limini tekshirib ko'ring.")

@bot.message_handler(commands=['debug_me'])
def debug_me(message):
    user_id = message.from_user.id
    my_quizzes = database.get_my_quizzes(user_id)
    all_quizzes = database.get_all_quizzes()
    text = (
        f"🔍 **DEBUG MA'LUMOTLARI:**\n"
        f"Sizning ID: `{user_id}`\n"
        f"Sizning testlaringiz soni: {len(my_quizzes)}\n"
        f"Bazadagi jami testlar soni: {len(all_quizzes)}\n"
        f"Baza fayli: `{database.DB_PATH}`\n"
    )
    bot.reply_to(message, text, parse_mode='Markdown')

@bot.message_handler(commands=['add_questions'])
def add_more_questions(message):
    if message.chat.type in ['group', 'supergroup']: return
    user_id = message.from_user.id
    quiz_id = get_current_quiz_id(user_id)
    
    if quiz_id:
        database.set_user_state(user_id, 'waiting_for_questions')
        bot.reply_to(message, "OK! Oxirgi yaratgan testingizga yana savollar qo'shishingiz mumkin. Savollarni (Poll) yuboravering.\n\nTugatganingizdan so'ng yana «Tugatish» tugmasini bosing.", reply_markup=get_question_keyboard())
    else:
        bot.reply_to(message, "Sizda hali yaratilgan test yo'q.")

@bot.message_handler(commands=['stop'])
def stop_process(message):
    chat_id = message.chat.id
    user_id = message.from_user.id
    
    # 1. Yakkaxon testni to'xtatish
    if user_id in taking_quiz_sessions:
        session = taking_quiz_sessions[user_id]
        if 'timer' in session and session['timer']:
            session['timer'].cancel()
        del taking_quiz_sessions[user_id]
        bot.send_message(chat_id, "🛑 Yakkaxon test to'xtatildi.", reply_markup=ReplyKeyboardRemove())
        return

    # 2. Guruh testini to'xtatish
    if chat_id in group_sessions:
        del group_sessions[chat_id]
        bot.send_message(chat_id, "🛑 Guruhdagi test to'xtatildi.", reply_markup=ReplyKeyboardRemove())
        return

    # 3. Test yaratishni to'xtatish
    state = database.get_user_state(user_id)
    if state != 'none':
        database.set_user_state(user_id, 'none')
        bot.send_message(chat_id, "🛑 Test yaratish jarayoni to'xtatildi.", reply_markup=ReplyKeyboardRemove())
        return

    bot.send_message(chat_id, "Hozircha hech qanday faol jarayon yo'q.")

@bot.message_handler(commands=['done'])
def finish_quiz_cmd(message):
    if message.chat.type in ['group', 'supergroup']: return
    ask_for_time_limit(message)

@bot.message_handler(func=lambda message: message.text == 'Tugatish')
def finish_quiz_btn(message):
    if message.chat.type in ['group', 'supergroup']: return
    ask_for_time_limit(message)

def ask_for_time_limit(message):
    user_id = message.from_user.id
    state = database.get_user_state(user_id)
    
    if state == 'waiting_for_questions':
        database.set_user_state(user_id, 'waiting_for_time_limit')
        
        markup = InlineKeyboardMarkup()
        markup.row(InlineKeyboardButton("10 s", callback_data="set_time_10"),
                   InlineKeyboardButton("15 s", callback_data="set_time_15"),
                   InlineKeyboardButton("20 s", callback_data="set_time_20"))
        markup.row(InlineKeyboardButton("25 s", callback_data="set_time_25"),
                   InlineKeyboardButton("30 s", callback_data="set_time_30"))
        markup.row(InlineKeyboardButton("50 s", callback_data="set_time_50"),
                   InlineKeyboardButton("120 s", callback_data="set_time_120"))
                   
        bot.send_message(message.chat.id, "Bosh menyudasiz.", reply_markup=ReplyKeyboardRemove())
        bot.reply_to(message, "Iltimos, har bir savol uchun qancha vaqt berilishini tanlang:", reply_markup=markup)
    else:
        bot.reply_to(message, "Siz hozir hech qanday test yaratmayapsiz. Boshlash uchun /newquiz ni bosing.", reply_markup=ReplyKeyboardRemove())

@bot.callback_query_handler(func=lambda call: call.data.startswith("set_time_"))
def set_quiz_time_limit(call):
    user_id = call.from_user.id
    state = database.get_user_state(user_id)
    
    if state == 'waiting_for_time_limit':
        time_limit = int(call.data.split('_')[2])
        quiz_id = get_current_quiz_id(user_id)
        
        if quiz_id:
            conn = sqlite3.connect(database.DB_PATH)
            cursor = conn.cursor()
            cursor.execute("UPDATE questions SET time_limit = ? WHERE quiz_id = ?", (time_limit, quiz_id))
            conn.commit()
            conn.close()
            
            database.set_user_state(user_id, 'waiting_for_shuffle')
            
            markup = InlineKeyboardMarkup()
            markup.add(InlineKeyboardButton("Barchasi", callback_data="set_shuffle_3"))
            markup.add(InlineKeyboardButton("Faqat savollar", callback_data="set_shuffle_1"))
            markup.add(InlineKeyboardButton("Faqat variantlar", callback_data="set_shuffle_2"))
            markup.add(InlineKeyboardButton("Yo'q", callback_data="set_shuffle_0"))
            
            bot.edit_message_text(f"⏳ Vaqt {time_limit} soniya qilib belgilandi.\n\nSavollar va variantlar aralashtirilsinmi?", 
                                  chat_id=call.message.chat.id, message_id=call.message.message_id, reply_markup=markup)
        else:
            bot.answer_callback_query(call.id, "Xatolik yuz berdi. (Test topilmadi)")
    else:
        bot.answer_callback_query(call.id, "Bu tugma hozir ishlamaydi.")

@bot.callback_query_handler(func=lambda call: call.data.startswith("set_shuffle_"))
def set_quiz_shuffle(call):
    user_id = call.from_user.id
    state = database.get_user_state(user_id)
    
    if state == 'waiting_for_shuffle':
        shuffle_mode = int(call.data.split('_')[2])
        quiz_id = get_current_quiz_id(user_id)
        
        if quiz_id:
            conn = sqlite3.connect(database.DB_PATH)
            cursor = conn.cursor()
            cursor.execute("UPDATE quizzes SET shuffle_mode = ? WHERE quiz_id = ?", (shuffle_mode, quiz_id))
            conn.commit()
            conn.close()
            
            database.set_user_state(user_id, 'none')
            
            markup = InlineKeyboardMarkup()
            start_url = f"https://t.me/{bot_info.username}?start=quiz_{quiz_id}"
            group_url = f"https://t.me/{bot_info.username}?startgroup=quiz_{quiz_id}"
            
            markup.add(InlineKeyboardButton("Yakkaxon yechish 👤", url=start_url))
            markup.add(InlineKeyboardButton("Guruhda yechish 👥", url=group_url))
            markup.add(InlineKeyboardButton("Do'stlarga ulashish ↗️", url=f"https://t.me/share/url?url={start_url}"))
            
            bot.edit_message_text(f"✅ Test muvaffaqiyatli yaratildi va saqlandi!\n\nQuyidagi tugmalar orqali testni o'zingiz yechishingiz yoki guruhlarga yuborishingiz mumkin:", 
                                  chat_id=call.message.chat.id, message_id=call.message.message_id, reply_markup=markup)
            bot.answer_callback_query(call.id, "Aralashtirish rejimi saqlandi!")
        else:
            bot.answer_callback_query(call.id, "Xatolik yuz berdi.")
    else:
        bot.answer_callback_query(call.id, "Bu tugma hozir ishlamaydi.")

@bot.message_handler(commands=['skip'])
def skip_description(message):
    if message.chat.type in ['group', 'supergroup']: return
    user_id = message.from_user.id
    state = database.get_user_state(user_id)
    
    if state == 'waiting_for_description':
        database.set_user_state(user_id, 'waiting_for_questions')
        text = "Izoh o'tkazib yuborildi.\n\nEndi test uchun savollarni yuboring. Pastdagi **«Savol yaratish»** tugmasini bosing va o'z savolingizni tayyorlang.\n\n⚠️ **Muhim:** Android telefonlarda xatolik (qizil undov) chiqmasligi uchun savol yaratayotganda **«Anonim viktorina» (Анонимное голосование)** belgisini o'chirib qo'ying!\n\nBarcha savollarni yuborib bo'lgach, **«Tugatish»** tugmasini bosing."
        bot.reply_to(message, text, reply_markup=get_question_keyboard())

@bot.message_handler(content_types=['text'])
def handle_text(message):
    if message.chat.type in ['group', 'supergroup']: return
    user_id = message.from_user.id
    state = database.get_user_state(user_id)
    
    if state == 'waiting_for_title':
        title = message.text
        quiz_id = database.create_quiz(user_id, title)
        user_sessions[user_id] = {'quiz_id': quiz_id}
        database.set_user_state(user_id, 'waiting_for_description')
        bot.reply_to(message, "Ajoyib! Endi test uchun qisqacha izoh (description) yuboring.\n\nAgar izoh qo'shishni xohlamasangiz, /skip buyrug'ini bosing.")
        
    elif state == 'waiting_for_description':
        description = message.text
        quiz_id = get_current_quiz_id(user_id)
        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()
        cursor.execute("UPDATE quizzes SET description = ? WHERE quiz_id = ?", (description, quiz_id))
        conn.commit()
        conn.close()
        database.set_user_state(user_id, 'waiting_for_questions')
        text = "Izoh saqlandi.\n\nEndi test uchun savollarni yuboring. Pastdagi **«Savol yaratish»** tugmasini bosing va o'z savolingizni tayyorlang.\n\n⚠️ **Muhim:** Android telefonlarda xatolik (qizil undov) chiqmasligi uchun savol yaratayotganda **«Anonim viktorina» (Анонимное голосование)** belgisini o'chirib qo'ying!\n\nBarcha savollarni yuborib bo'lgach, **«Tugatish»** tugmasini bosing."
        bot.reply_to(message, text, reply_markup=get_question_keyboard())
    elif message.text == '➕ Yangi test yaratish':
        create_new_quiz(message)
        
    elif message.text == '📂 Mening testlarim':
        show_my_quizzes(message)
        

        
    elif message.text == '📊 Natijalar':
        bot.reply_to(message, "Natijalar bo'limi tez kunda ishga tushadi! 🔜")

    else:
        bot.reply_to(message, "Asosiy menyu:", reply_markup=get_main_keyboard())

def show_my_quizzes(message):
    user_id = message.from_user.id
    quizzes = database.get_my_quizzes(user_id)
    
    if not quizzes:
        bot.reply_to(message, "Sizda hali yaratilgan testlar yo'q. /newquiz orqali yangi test yarating.")
        return
        
    text = "📂 **Sizning testlaringiz:**\n\n"
    markup = InlineKeyboardMarkup()
    for q_id, title in quizzes:
        text += f"🔹 {title}\n"
        start_url = f"https://t.me/{bot_info.username}?start=quiz_{q_id}"
        markup.add(InlineKeyboardButton(f" Ishlash: {title}", url=start_url))
        
    bot.reply_to(message, text, reply_markup=markup, parse_mode='Markdown')



@bot.message_handler(content_types=['poll'])
def handle_poll(message):
    if message.chat.type in ['group', 'supergroup']: return
    user_id = message.from_user.id
    state = database.get_user_state(user_id)
    
    if state == 'waiting_for_questions':
        poll = message.poll
        if not poll.type == 'quiz':
            bot.reply_to(message, "Iltimos, so'rovnomani **Quiz Mode** (Viktorina rejimi) da yarating!")
            return
            
        quiz_id = get_current_quiz_id(user_id)
        question_text = poll.question
        options = json.dumps([option.text for option in poll.options])
        correct_option_id = poll.correct_option_id
        explanation = poll.explanation if poll.explanation else ""
        
        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO questions (quiz_id, question_text, options, correct_option_id, explanation)
            VALUES (?, ?, ?, ?, ?)
        """, (quiz_id, question_text, options, correct_option_id, explanation))
        conn.commit()
        conn.close()
        
        bot.reply_to(message, "✅ Savol qabul qilindi! Keyingi savolni yaratish uchun yana **«Savol yaratish»** tugmasini bosing yoki tugatish uchun **«Tugatish»** tugmasini bosing.", reply_markup=get_question_keyboard())

# --- AI ORQALI TEST YARATISH (Fayl va Rasmlar) ---

@bot.message_handler(content_types=['document'])
def handle_document(message):
    if message.chat.type in ['group', 'supergroup']: return
    user_id = message.from_user.id
    
    ADMIN_ID = 6559589296
    
    # Limit tekshirish
    if user_id != ADMIN_ID:
        usage = database.get_daily_ai_usage(user_id)
        if usage >= 3:
            bot.reply_to(message, "❌ Sizning bugungi limitingiz tugadi. Ertaga fayl yoki rasmlaringizni qaytadan yuborib ko'ring.")
            return
            
    # API key mavjudligini tekshirish
    if not os.getenv('GEMINI_API_KEY'):
        text = (
            "⚠️ **AI Quiz funksiyasi faol emas!**\n\n"
            "Ushbu funksiyadan foydalanish uchun bot egasi `.env` fayliga `GEMINI_API_KEY` kalitini qo'shishi kerak.\n\n"
            "🔑 **Kalitni qanday olish mumkin?**\n"
            "1. https://aistudio.google.com/ saytiga kiring.\n"
            "2. Bepul **Gemini API Key** oling.\n"
            "3. Botning `.env` fayliga yozing:\n"
            "`GEMINI_API_KEY=sizning_kalitingiz`"
        )
        bot.reply_to(message, text, parse_mode='Markdown')
        return

    file_name = message.document.file_name
    file_ext = os.path.splitext(file_name)[1].lower()
    
    if file_ext not in ['.pdf', '.docx', '.txt', '.pptx']:
        bot.reply_to(message, "❌ Kechirasiz, faqat `.pdf`, `.docx`, `.pptx` va `.txt` formatidagi fayllarni qabul qila olaman.")
        return
        
    processing_msg = bot.reply_to(message, "📥 **Hujjat yuklab olinmoqda va o'qilmoqda...** Iltimos, kuting ⏳", parse_mode='Markdown')
    
    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        # Faylni vaqtinchalik saqlash
        temp_path = f"temp_{user_id}{file_ext}"
        with open(temp_path, 'wb') as new_file:
            new_file.write(downloaded_file)
            
        # Matnni ajratib olish
        content_text = ""
        if file_ext == '.pdf':
            content_text = extract_text_from_pdf(temp_path)
        elif file_ext == '.docx':
            content_text = extract_text_from_docx(temp_path)
        elif file_ext == '.pptx':
            content_text = extract_text_from_pptx(temp_path)
        elif file_ext == '.txt':
            content_text = downloaded_file.decode('utf-8', errors='ignore')
            
        # Vaqtinchalik faylni o'chirish
        if os.path.exists(temp_path):
            os.remove(temp_path)
            
        if not content_text or len(content_text.strip()) < 20:
            bot.edit_message_text("❌ Fayldan matn o'qib bo'lmadi yoki undagi matn juda qisqa (kamida 20 ta belgi bo'lishi kerak).", chat_id=message.chat.id, message_id=processing_msg.message_id)
            return
            
        # Sessiyaga yozish
        ai_sessions[user_id] = {
            'type': 'text',
            'data': content_text,
            'file_name': file_name
        }
        
        # Savollar sonini tanlash tugmalari
        markup = InlineKeyboardMarkup()
        markup.row(InlineKeyboardButton("5 ta savol 📝", callback_data="ai_count_5"),
                   InlineKeyboardButton("10 ta savol 📝", callback_data="ai_count_10"))
        markup.row(InlineKeyboardButton("15 ta savol 📝", callback_data="ai_count_15"),
                   InlineKeyboardButton("20 ta savol 📝", callback_data="ai_count_20"))
        markup.row(InlineKeyboardButton("30 ta savol 📝", callback_data="ai_count_30"),
                   InlineKeyboardButton("40 ta savol 📝", callback_data="ai_count_40"))
        
        limit_text = ""
        if user_id != ADMIN_ID:
            usage = database.get_daily_ai_usage(user_id)
            limit_text = f"\n⚠️ *Eslatma: Siz bir kunda maksimal 3 tagacha fayl yoki rasm yubora olasiz. (Bugungi qolgan limitingiz: {3 - usage} ta)*\n"
        else:
            limit_text = f"\n⭐ *Siz uchun AI limiti cheksiz!*\n"
            
        success_text = (
            f"📄 **Fayl muvaffaqiyatli o'qildi!**\n"
            f"Hujjat: `{file_name}`\n"
            f"Matn hajmi: {len(content_text)} ta belgi.\n"
            f"{limit_text}\n"
            f"Ushbu fayl asosida nechta test savoli yaratmoqchisiz? Tanlang 👇"
        )
        bot.edit_message_text(success_text, chat_id=message.chat.id, message_id=processing_msg.message_id, reply_markup=markup, parse_mode='Markdown')
        
    except Exception as e:
        bot.edit_message_text(f"❌ Faylni yuklash yoki o'qishda xatolik yuz berdi: {e}", chat_id=message.chat.id, message_id=processing_msg.message_id)


@bot.message_handler(content_types=['photo'])
def handle_photo(message):
    if message.chat.type in ['group', 'supergroup']: return
    user_id = message.from_user.id
    
    ADMIN_ID = 6559589296
    
    # Limit tekshirish
    if user_id != ADMIN_ID:
        usage = database.get_daily_ai_usage(user_id)
        if usage >= 3:
            bot.reply_to(message, "❌ Sizning bugungi limitingiz tugadi. Ertaga fayl yoki rasmlaringizni qaytadan yuborib ko'ring.")
            return
            
    # API key mavjudligini tekshirish
    if not os.getenv('GEMINI_API_KEY'):
        text = (
            "⚠️ **AI Quiz funksiyasi faol emas!**\n\n"
            "Ushbu funksiyadan foydalanish uchun bot egasi `.env` fayliga `GEMINI_API_KEY` kalitini qo'shishi kerak.\n\n"
            "🔑 **Kalitni qanday olish mumkin?**\n"
            "1. https://aistudio.google.com/ saytiga kiring.\n"
            "2. Bepul **Gemini API Key** oling.\n"
            "3. Botning `.env` fayliga yozing:\n"
            "`GEMINI_API_KEY=sizning_kalitingiz`"
        )
        bot.reply_to(message, text, parse_mode='Markdown')
        return

    processing_msg = bot.reply_to(message, "📥 **Rasm yuklab olinmoqda...** Iltimos, kuting ⏳", parse_mode='Markdown')
    
    try:
        # Eng katta rasmni olish
        photo = message.photo[-1]
        file_info = bot.get_file(photo.file_id)
        downloaded_file = bot.download_file(file_info.file_path)
        
        # Sessiyaga yozish (Gemini uchun bayt ko'rinishida)
        ai_sessions[user_id] = {
            'type': 'image',
            'data': downloaded_file,
            'file_name': 'Rasm'
        }
        
        # Savollar sonini tanlash tugmalari
        markup = InlineKeyboardMarkup()
        markup.row(InlineKeyboardButton("5 ta savol 📝", callback_data="ai_count_5"),
                   InlineKeyboardButton("10 ta savol 📝", callback_data="ai_count_10"))
        markup.row(InlineKeyboardButton("15 ta savol 📝", callback_data="ai_count_15"),
                   InlineKeyboardButton("20 ta savol 📝", callback_data="ai_count_20"))
        markup.row(InlineKeyboardButton("30 ta savol 📝", callback_data="ai_count_30"),
                   InlineKeyboardButton("40 ta savol 📝", callback_data="ai_count_40"))
        
        limit_text = ""
        if user_id != ADMIN_ID:
            usage = database.get_daily_ai_usage(user_id)
            limit_text = f"\n⚠️ *Eslatma: Siz bir kunda maksimal 3 tagacha fayl yoki rasm yubora olasiz. (Bugungi qolgan limitingiz: {3 - usage} ta)*\n"
        else:
            limit_text = f"\n⭐ *Siz uchun AI limiti cheksiz!*\n"
            
        success_text = (
            f"🖼 **Rasm muvaffaqiyatli qabul qilindi!**\n"
            f"{limit_text}\n"
            f"Ushbu rasm/skrinshotdagi ma'lumotlar asosida nechta test savoli yaratmoqchisiz? Tanlang 👇"
        )
        bot.edit_message_text(success_text, chat_id=message.chat.id, message_id=processing_msg.message_id, reply_markup=markup, parse_mode='Markdown')
        
    except Exception as e:
        bot.edit_message_text(f"❌ Rasmni yuklashda xatolik yuz berdi: {e}", chat_id=message.chat.id, message_id=processing_msg.message_id)


@bot.callback_query_handler(func=lambda call: call.data.startswith("ai_count_"))
def handle_ai_count(call):
    user_id = call.from_user.id
    session = ai_sessions.get(user_id)
    
    if not session:
        bot.answer_callback_query(call.id, "Sessiya eskirgan. Qaytadan urinib ko'ring.")
        bot.edit_message_text("❌ Hech qanday faol AI sessiyasi topilmadi. Qayta urinib ko'ring (botga yangi fayl yoki rasm yuboring).", 
                              chat_id=call.message.chat.id, message_id=call.message.message_id)
        return
        
    count = int(call.data.split('_')[2])
    session['count'] = count
    bot.answer_callback_query(call.id, f"{count} ta savol tanlandi!")
    
    # Vaqtni tanlash tugmalari
    markup = InlineKeyboardMarkup()
    markup.row(InlineKeyboardButton("10 s ⏱️", callback_data="ai_time_10"),
               InlineKeyboardButton("15 s ⏱️", callback_data="ai_time_15"),
               InlineKeyboardButton("20 s ⏱️", callback_data="ai_time_20"))
    markup.row(InlineKeyboardButton("25 s ⏱️", callback_data="ai_time_25"),
               InlineKeyboardButton("50 s ⏱️", callback_data="ai_time_50"),
               InlineKeyboardButton("120 s ⏱️", callback_data="ai_time_120"))
               
    text = (
        f"🎯 Savollar soni: **{count} ta** tanlandi.\n\n"
        f"Endi, har bir savol uchun qancha vaqt berilishini tanlang 👇"
    )
    bot.edit_message_text(text, chat_id=call.message.chat.id, message_id=call.message.message_id, reply_markup=markup, parse_mode='Markdown')


@bot.callback_query_handler(func=lambda call: call.data.startswith("ai_time_"))
def handle_ai_time(call):
    user_id = call.from_user.id
    session = ai_sessions.get(user_id)
    
    if not session or 'count' not in session:
        bot.answer_callback_query(call.id, "Sessiya eskirgan. Qaytadan urinib ko'ring.")
        bot.edit_message_text("❌ Hech qanday faol AI sessiyasi topilmadi. Qayta urinib ko'ring.", 
                              chat_id=call.message.chat.id, message_id=call.message.message_id)
        return
        
    time_limit = int(call.data.split('_')[2])
    count = session['count']
    session['time_limit'] = time_limit
    
    bot.answer_callback_query(call.id, f"{time_limit} soniya tanlandi!")
    
    bot.edit_message_text(f"🤖 **AI ma'lumotlarni tahlil qilmoqda va test yaratmoqda...**\n"
                          f"Savollar soni: {count} ta\n"
                          f"Vaqt limiti: {time_limit} soniya\n\n"
                          f"Iltimos, kuting. Bu jarayon 10-30 soniya vaqt olishi mumkin ⏳", 
                          chat_id=call.message.chat.id, message_id=call.message.message_id, parse_mode='Markdown')
    
    threading.Thread(target=generate_ai_quiz_thread, args=(call.message, user_id, session, count, time_limit)).start()


def generate_ai_quiz_thread(message, user_id, session, count, time_limit):
    chat_id = message.chat.id
    message_id = message.message_id
    
    try:
        model = genai.GenerativeModel("gemini-2.5-flash")
        
        prompt = (
            f"Sizga taqdim etilgan matn yoki rasmdagi ma'lumotlardan foydalanib, roppa-rosa {count} ta "
            "ko'p variantli (multiple choice) test savollarini yarating. "
            "Barcha savollar, variantlar va tushuntirishlar faqat o'zbek tilida bo'lishi kerak. "
            "Har bir savol uchun aniq 4 ta variant (options) tayyorlang.\n\n"
            "Agar matnda tayyor savollar va to'g'ri javoblar ko'rsatilgan bo'lsa (masalan `#`, `*`, `bold` belgilar orqali), "
            "o'sha tayyor savollardan va ko'rsatilgan to'g'ri javoblardan foydalaning va to'g'ri javobning indeksini belgilang.\n\n"
            "Natijani mutlaqo qat'iy ravishda quyidagi JSON formatida qaytaring. Hech qanday boshqa matn, kirish yoki markdown formatting (masalan ```json) qo'shmang, "
            "faqat quyidagi struktura bo'yicha toza JSON yuboring:\n"
            "{\n"
            '  "title": "Mavzuga mos chiroyli test sarlavhasi (maksimal 50 ta belgi)",\n'
            '  "description": "Fayl/mavzu haqida qisqacha izoh (maksimal 150 ta belgi)",\n'
            '  "questions": [\n'
            "    {\n"
            '      "question": "Savol matni (maksimal 150 ta belgi)",\n'
            '      "options": ["1-variant", "2-variant", "3-variant", "4-variant"],\n'
            '      "correct_option_index": 0,\n'
            '      "explanation": "Nega aynan shu variant to\'g\'ri ekanligining qisqacha izohi (maksimal 100 ta belgi, ixtiyoriy)"\n'
            "    }\n"
            "  ]\n"
            "}"
        )
        
        # Gemini modeliga so'rov yuborish
        if session['type'] == 'text':
            full_content = f"Hujjat nomi: {session['file_name']}\n\nHujjat matni:\n{session['data']}"
            if len(full_content) > 60000:
                full_content = full_content[:60000] + "\n...[Matn juda uzunligi uchun kesildi]..."
                
            response = model.generate_content(
                [prompt, full_content],
                generation_config={"response_mime_type": "application/json"}
            )
        else:
            image_part = {
                "mime_type": "image/jpeg",
                "data": session['data']
            }
            response = model.generate_content(
                [prompt, image_part],
                generation_config={"response_mime_type": "application/json"}
            )
            
        raw_text = response.text.strip()
        
        # Agar rasm/matn mos kelmasa va xatolik bo'lsa
        if not raw_text:
            bot.edit_message_text("❌ Xatolik: AI matndan ma'lumot ololmadi yoki javob qaytarmadi.", chat_id=chat_id, message_id=message_id)
            return
            
        # Kutilmagan markdown belgilarni tozalash (agar response_mime_type ishlamasa)
        if raw_text.startswith("```json"):
            raw_text = raw_text.replace("```json", "", 1)
        if raw_text.endswith("```"):
            raw_text = raw_text[:-3].strip()
        raw_text = raw_text.strip()
        
        data = json.loads(raw_text)
        title = data.get('title', f"AI Test - {session['file_name']}")
        description = data.get('description', "AI tomonidan avtomatik yaratilgan test.")
        questions = data.get('questions', [])
        
        if not questions:
            bot.edit_message_text("❌ Xatolik: AI birorta ham test savoli yarata olmadi. Boshqa fayl yoki rasm yuborib ko'ring.", chat_id=chat_id, message_id=message_id)
            return
            
        # Testni bazada yaratish
        quiz_id = database.create_quiz(user_id, title)
        
        conn = sqlite3.connect(database.DB_PATH)
        cursor = conn.cursor()
        cursor.execute("UPDATE quizzes SET description = ? WHERE quiz_id = ?", (description, quiz_id))
        
        added_count = 0
        for q in questions:
            q_text = q.get('question')
            options = q.get('options')
            correct_idx = q.get('correct_option_index', 0)
            explanation = q.get('explanation', '')
            
            if q_text and options and len(options) >= 2:
                options_json = json.dumps(options[:10]) # Telegram max 10 ta variant qabul qiladi
                cursor.execute("""
                    INSERT INTO questions (quiz_id, question_text, options, correct_option_id, explanation, time_limit)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (quiz_id, q_text, options_json, correct_idx, explanation, time_limit))
                added_count += 1
                
        conn.commit()
        conn.close()
        
        # Limit hisoblagichini oshirish
        database.increment_daily_ai_usage(user_id)
        
        # Sessiyani tozalash
        if user_id in ai_sessions:
            del ai_sessions[user_id]
            
        if added_count == 0:
            bot.edit_message_text("❌ Xatolik: Savollarni tahlil qilishda yoki bazaga yozishda muammo yuz berdi.", chat_id=chat_id, message_id=message_id)
            return
            
        # O'ynash havolalarini generatsiya qilish
        start_url = f"https://t.me/{bot_info.username}?start=quiz_{quiz_id}"
        group_url = f"https://t.me/{bot_info.username}?startgroup=quiz_{quiz_id}"
        
        markup = InlineKeyboardMarkup()
        markup.add(InlineKeyboardButton("Yakkaxon yechish 👤", url=start_url))
        markup.add(InlineKeyboardButton("Guruhda yechish 👥", url=group_url))
        markup.add(InlineKeyboardButton("Do'stlarga ulashish ↗️", url=f"https://t.me/share/url?url={start_url}"))
        
        success_text = (
            f"🎉 **AI Quiz muvaffaqiyatli yaratildi!**\n\n"
            f"📌 **Sarlavha:** {title}\n"
            f"📝 **Izoh:** {description}\n"
            f"❓ **Savollar soni:** {added_count} ta\n\n"
            f"Quyidagi tugmalar orqali testni yechishingiz yoki ulashishingiz mumkin:"
        )
        bot.edit_message_text(success_text, chat_id=chat_id, message_id=message_id, reply_markup=markup, parse_mode='Markdown')
        
    except json.JSONDecodeError:
        bot.edit_message_text("❌ Xatolik: AI qaytargan ma'lumotni JSON formatida o'qib bo'lmadi. Qayta urinib ko'ring.", chat_id=chat_id, message_id=message_id)
    except Exception as e:
        error_msg = str(e)
        print(f"AI error: {error_msg}")
        if "429" in error_msg or "Quota" in error_msg:
            user_msg = "❌ Kechirasiz, ayni vaqtda botga so'rovlar juda ko'payib ketdi (AI limiti tugadi). Iltimos, birozdan so'ng (1-2 daqiqa) qayta urinib ko'ring."
        else:
            user_msg = "❌ Test yaratishda kutilmagan xatolik yuz berdi. Iltimos, qaytadan urinib ko'ring."
        bot.edit_message_text(user_msg, chat_id=chat_id, message_id=message_id)



import time
print("Bot ishga tushmoqda (Versiya: 1.1 - FixID qo'shildi)...")
print(f"Baza yo'li: {database.DB_PATH}")
while True:
    try:
        bot.polling(none_stop=True, timeout=60, long_polling_timeout=60)
    except Exception as e:
        print(f"Internet uzilishi yoki xatolik: {e}")
        time.sleep(3)
