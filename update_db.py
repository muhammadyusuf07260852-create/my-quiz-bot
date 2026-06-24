import sqlite3

def update():
    conn = sqlite3.connect('quizbot.db')
    cursor = conn.cursor()
    try:
        cursor.execute("ALTER TABLE quizzes ADD COLUMN shuffle_mode INTEGER DEFAULT 0")
        print("Column added")
    except sqlite3.OperationalError:
        print("Column already exists")
    conn.commit()
    conn.close()

if __name__ == "__main__":
    update()
