import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import memberList, bookList, borrowReturnBook, libraryDB

class MainWindow(tk.Tk):

    def __init__(self):
        super().__init__()
        self.db = libraryDB.DatabaseConnection()
        self.db.create_tables()
        self.title("Library Management System")
        self.geometry("300x220+550+250")
        ttk.Style("darkly")
        self.resizable(width=False, height=False)
        self.create_widgets()
        self.create_layout()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.mainloop()

    def create_widgets(self):
        self.mainlbl = ttk.Label(self, text="Library Management System", font=("Tahoma", 16))
        self.memberListBtn = ttk.Button(self, text="Member List", command=self.member_list_window)
        self.borrowReturnBtn = ttk.Button(self, text="Borrow / Return Book", command=self.borrow_return_book_window)
        self.bookListBtn = ttk.Button(self, text="Book List", command=self.book_list_window)

    def create_layout(self):
        self.mainlbl.pack(pady=(20, 0))
        self.memberListBtn.pack(pady=(20, 0))
        self.bookListBtn.pack(pady=(20, 0))
        self.borrowReturnBtn.pack(pady=(20, 0))

    def member_list_window(self):
        self.withdraw()
        memberList.MemberList(db_connection=self.db.db_connection)

    def book_list_window(self):
        self.withdraw()
        bookList.BookList(db_connection=self.db.db_connection)

    def borrow_return_book_window(self):
        self.withdraw()
        borrowReturnBook.BorrowReturnBook(db_connection=self.db.db_connection)

    def on_close(self):
        if self.db:
            self.db.close_connection()
        self.destroy()

app = MainWindow()
