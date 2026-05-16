import sqlite3

def init_db():
    conn = sqlite3.connect('quizbot.db')
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
    conn = sqlite3.connect('quizbot.db')
    cursor = conn.cursor()
    cursor.execute("UPDATE users SET state = ? WHERE user_id = ?", (state, user_id))
    conn.commit()
    conn.close()

def get_user_state(user_id):
    conn = sqlite3.connect('quizbot.db')
    cursor = conn.cursor()
    cursor.execute("SELECT state FROM users WHERE user_id = ?", (user_id,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else 'none'

def create_quiz(creator_id, title):
    conn = sqlite3.connect('quizbot.db')
    cursor = conn.cursor()
    cursor.execute("INSERT INTO quizzes (creator_id, title) VALUES (?, ?)", (creator_id, title))
    quiz_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return quiz_id

# Boshqa funksiyalarni ehtiyojga qarab shu yerga qo'shib boramiz

if __name__ == "__main__":
    init_db()
    print("Database initialized successfully.")
