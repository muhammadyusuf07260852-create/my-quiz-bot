import sqlite3
import os

DB_PATH = os.path.join(r'c:\Users\user\Desktop\quiz bot', 'quizbot.db')

def test_string_id():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Try with integer
    cursor.execute("SELECT title FROM quizzes WHERE quiz_id = ?", (19,))
    res_int = cursor.fetchone()
    print(f"Result with int(19): {res_int}")
    
    # Try with string
    cursor.execute("SELECT title FROM quizzes WHERE quiz_id = ?", ("19",))
    res_str = cursor.fetchone()
    print(f"Result with str('19'): {res_str}")
    
    conn.close()

if __name__ == "__main__":
    test_string_id()
