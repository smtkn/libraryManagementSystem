import sqlite3


class DatabaseConnection:

    def __init__(self, db_name="library.db"):
        self.db_name = db_name
        self.db_connection = None
        self.db_cursor = None
        self.connect()

    def connect(self):
        self.db_connection = sqlite3.connect(self.db_name)
        self.db_cursor = self.db_connection.cursor()

    def create_tables(self):
        self.db_cursor.execute('''
            CREATE TABLE IF NOT EXISTS members (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                surname TEXT NOT NULL,
                email TEXT UNIQUE NOT NULL,
                phone TEXT UNIQUE NOT NULL
            )
        ''')

        self.db_cursor.execute('''
            CREATE TABLE IF NOT EXISTS books (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                book_name TEXT NOT NULL,
                book_author TEXT NOT NULL,
                book_genre TEXT NOT NULL,
                book_status TEXT DEFAULT 'Available'
            )
        ''')

        self.db_cursor.execute('''
            CREATE TABLE IF NOT EXISTS borrow_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                member_id INTEGER NOT NULL,
                book_id INTEGER NOT NULL,
                borrow_date TEXT NOT NULL,
                return_date TEXT,
                FOREIGN KEY(member_id) REFERENCES members(id),
                FOREIGN KEY(book_id) REFERENCES books(id)
            )
        ''')

        self.db_connection.commit()

    def close_connection(self):
        if self.db_connection:
            self.db_connection.close()

    def __del__(self):
        self.close_connection()