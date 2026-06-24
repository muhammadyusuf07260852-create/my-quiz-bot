import sqlite3
import json

def check_db():
    conn = sqlite3.connect(r'c:\Users\user\Desktop\quiz bot\quizbot.db')
    cursor = conn.cursor()
    
    print("--- Quizzes ---")
    cursor.execute("SELECT * FROM quizzes")
    rows = cursor.fetchall()
    for row in rows:
        try:
            print(row)
        except:
            print("Row contains non-encodable characters")
        
    print("\n--- Users ---")
    cursor.execute("SELECT * FROM users")
    rows = cursor.fetchall()
    for row in rows:
        try:
            print(row)
        except:
            print("Row contains non-encodable characters")
        
    conn.close()

if __name__ == "__main__":
    check_db()
