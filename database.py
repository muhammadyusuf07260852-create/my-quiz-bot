import sqlite3
import os

# Railway Volume uchun: DATA_DIR muhit o'zgaruvchisi bo'lsa, u yerdan foydalanamiz
DATA_DIR = os.environ.get('DATA_DIR', os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(DATA_DIR, 'quizbot.db')

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Users table
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS users (
        user_id INTEGER PRIMARY KEY,
        username TEXT,
        first_name TEXT,
        state TEXT DEFAULT 'none'
    )
    ''')

    # Quizzes table
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS quizzes (
        quiz_id INTEGER PRIMARY KEY AUTOINCREMENT,
        creator_id INTEGER,
        title TEXT,
        description TEXT,
        is_published BOOLEAN DEFAULT 0,
        FOREIGN KEY(creator_id) REFERENCES users(user_id)
    )
    ''')

    # Questions table
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS questions (
        question_id INTEGER PRIMARY KEY AUTOINCREMENT,
        quiz_id INTEGER,
        question_text TEXT,
        options TEXT, -- JSON array of options
        correct_option_id INTEGER,
        explanation TEXT,
        time_limit INTEGER DEFAULT 15,
        FOREIGN KEY(quiz_id) REFERENCES quizzes(quiz_id)
    )
    ''')

    # Results table
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS results (
        result_id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        quiz_id INTEGER,
        score INTEGER,
        time_taken INTEGER,
        FOREIGN KEY(user_id) REFERENCES users(user_id),
        FOREIGN KEY(quiz_id) REFERENCES quizzes(quiz_id)
    )
    ''')

    conn.commit()
    conn.close()

def set_user_state(user_id, state):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("UPDATE users SET state = ? WHERE user_id = ?", (state, user_id))
    conn.commit()
    conn.close()

def get_user_state(user_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT state FROM users WHERE user_id = ?", (user_id,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else 'none'

def create_quiz(creator_id, title):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO quizzes (creator_id, title) VALUES (?, ?)", (creator_id, title))
    quiz_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return quiz_id

# Boshqa funksiyalar
def get_my_quizzes(user_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT quiz_id, title FROM quizzes WHERE creator_id = ?", (user_id,))
    rows = cursor.fetchall()
    conn.close()
    return rows

def get_all_quizzes():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT quiz_id, title FROM quizzes")
    rows = cursor.fetchall()
    conn.close()
    return rows

def get_quiz_by_id(quiz_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT title, description FROM quizzes WHERE quiz_id = ?", (quiz_id,))
    row = cursor.fetchone()
    conn.close()
    return row

def fix_user_quizzes(old_id, new_id):
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("UPDATE quizzes SET creator_id = ? WHERE creator_id = ?", (new_id, old_id))
    conn.commit()
    conn.close()

if __name__ == "__main__":
    init_db()
    print("Database initialized successfully.")
