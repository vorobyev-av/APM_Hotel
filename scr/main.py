import tkinter as tk
from tkinter import ttk
from PIL import ImageTk, Image
import sqlite3

# Класс, описывающий создание окна приложения
class App(tk.Tk):
    def __init__(self, title, size):
        super().__init__()
        self.title(title)
        self.geometry(f'{size[0]}x{size[1]}')
        self.minsize(size[0], size[1])

        self.menu = Menu(self)
        #self.main = Main(self)
        self.view = View(self)
        self.about = About(self)
        self.login = Login(self)


        # Передача ссылки на экземпляр Login в Menu
        self.menu.set_login_frame(self.login)

        # Передача ссылки на экземпляр Main в Menu
        #self.menu.set_main_frame(self.main)

        # Передача ссылки на экземпляр View в Menu
        self.menu.set_view_frame(self.view)

        # Передача ссылки на экземпляр About в Menu
        self.menu.set_about_frame(self.about)

        self.mainloop()


# Класс, описывающий создание области КЛИЕНТЫ
class Clients(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        testLbl = ttk.Label(self, background='black')
        testLbl.pack(expand=True, fill='both')
        self.place(relx=0.3, y=0, relwidth=0.7, relheight=1)
        self.create_test()

    def create_test(self):
        buttonMenu5 = ttk.Button(self, text='login')
        buttonMenu5.place(relx=0.27, rely=0.8, relwidth=0.45, height=40)


# Класс, описывающий создание меню в левой части окна
class Menu(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.place(x=0, y=0, relwidth=0.3, relheight=1)
        self.create_widgets()

    def create_widgets(self):
        buttonMenu1 = ttk.Button(self, text='Test1')
        buttonMenu1.place(relx=0.27, rely=0.2, relwidth=0.45, height=40)

        buttonMenu2 = ttk.Button(self, text='Test2')
        buttonMenu2.place(relx=0.27, rely=0.3, relwidth=0.45, height=40)

        buttonMenu3 = ttk.Button(self, text='Test3')
        buttonMenu3.place(relx=0.27, rely=0.4, relwidth=0.45, height=40)

        buttonMenu4 = ttk.Button(self, text='View', command = self.show_view)
        buttonMenu4.place(relx=0.27, rely=0.5, relwidth=0.45, height=40)

        # Кнопка для перехода на экран "About"
        buttonMenu5 = ttk.Button(self, text='About', command=self.show_about)
        buttonMenu5.place(relx=0.27, rely=0.9, relwidth=0.45, height=40)

        buttonMenu6 = ttk.Button(self, text='Login', command=self.show_login)
        buttonMenu6.place(relx=0.27, rely=0.8, relwidth=0.45, height=40)


    def set_view_frame(self, view_frame):
        self.view_frame = view_frame

    def show_view(self):
        self.view_frame.tkraise()


    def set_about_frame(self, about_frame):
        self.about_frame = about_frame

    def show_about(self):
        self.about_frame.tkraise()
        

    def set_login_frame(self, login_frame):
        self.login_frame = login_frame

    def show_login(self):
        self.login_frame.tkraise()


# Класс, описывающий создание главной области
class Login(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        testLbl = ttk.Label(self, background='blue')
        testLbl.pack(expand=True, fill='both')
        self.place(relx=0.3, y=0, relwidth=0.7, relheight=1)
        self.create_widgets()

    def create_widgets(self):
        buttonLoginEnter = ttk.Button(self, text='login')
        buttonLoginEnter.place(relx=0.27, rely=0.65, relwidth=0.45, height=40)

        entryLoginLogin = ttk.Entry(self)
        entryLoginLogin.place(relx=0.27, rely=0.45, relwidth=0.45, height=40)

        entryLoginPassword = ttk.Entry(self)
        entryLoginPassword.place(relx=0.27, rely=0.55, relwidth=0.45, height=40)


# Класс, описывающий создание области "View"
class View(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        testLbl = ttk.Label(self, background='purple')
        testLbl.pack(expand=True, fill='both')
        self.place(relx=0.3, y=0, relwidth=0.7, relheight=1)
        self.create_test()

    def create_test(self):
        buttonMenu5 = ttk.Button(self, text='view_test')
        buttonMenu5.place(relx=0.27, rely=0.8, relwidth=0.45, height=40)

# Класс, описывающий создание области "About"
class About(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        testLbl = ttk.Label(self, background='red')
        testLbl.pack(expand=True, fill='both')
        self.place(relx=0.3, y=0, relwidth=0.7, relheight=1)
        self.create_test()

    def create_test(self):
        buttonMenu5 = ttk.Button(self, text='Test5')
        buttonMenu5.place(relx=0.27, rely=0.9, relwidth=0.45, height=40)

# Создание экземпляра класса
App('АРМ Космос', (800, 600))