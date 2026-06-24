import sqlite3
import database

def test_new_functions():
    # Database path
    db_path = r'c:\Users\user\Desktop\quiz bot\quizbot.db'
    
    # Test get_all_quizzes
    all_quizzes = database.get_all_quizzes()
    print(f"Total quizzes found: {len(all_quizzes)}")
    for q in all_quizzes:
        print(f"ID: {q[0]}, Title: {q[1]}")
        
    # Test get_my_quizzes for a known user_id from previous check
    # Let's try 6559589296
    my_quizzes = database.get_my_quizzes(6559589296)
    print(f"\nQuizzes for user 6559589296: {len(my_quizzes)}")
    for q in my_quizzes:
        print(f"ID: {q[0]}, Title: {q[1]}")

if __name__ == "__main__":
    test_new_functions()
