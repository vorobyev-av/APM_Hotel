import tkinter as tk
from tkinter import ttk
from PIL import ImageTk, Image
import sqlite3
import sys, os
from datetime import datetime
import locale

locale.setlocale(locale.LC_TIME, 'russian')

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
        self.clients = Clients(self)
        self.rent = Rent(self)
        self.room = Room(self)
        self.login = Login(self)

        # Передача ссылки на экземпляр Login в Menu
        self.menu.set_login_frame(self.login)

        # Передача ссылки на экземпляр Clients в Menu
        self.menu.set_clients_frame(self.clients)

        # Передача ссылки на экземпляр Rent в Menu
        self.menu.set_rent_frame(self.rent)
       
        # Передача ссылки на экземпляр Room в Menu
        self.menu.set_room_frame(self.room)
        
        # Передача ссылки на экземпляр View в Menu
        self.menu.set_view_frame(self.view)

        # Передача ссылки на экземпляр About в Menu
        self.menu.set_about_frame(self.about)

        style = ttk.Style(self)
        style.theme_use('classic')
        style.configure('TButton', background = 'red', foreground = 'white', width = 20, borderwidth=1, focusthickness=3, focuscolor='none')
        style.map('TButton', background=[('active','#E8B4BC')])

        self.mainloop()


# Класс, описывающий создание области КЛИЕНТЫ
class Clients(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        testLbl = ttk.Label(self, background='black')
        #testLbl.pack(expand=True, fill='both')
        testLbl.place(height=999, width=999)
        self.place(relx=0.3, y=0, relwidth=0.7, relheight=1)
        self.create_test()
        self.create_widgets()
        self.fetch_data()
        self.display_data()


    def create_test(self):
        buttonMenu5 = ttk.Button(self, text='clients')
        buttonMenu5.place(relx=0.27, rely=0.8, relwidth=0.45, height=40)

    def create_widgets(self):
        buttonNewClient = ttk.Button(self, text='Новый клиент')
        buttonNewClient.place(relx=0.1, rely=0.1, relwidth=0.10, height=40)

        buttonEdit = ttk.Button(self, text='Изменить')
        buttonEdit.place(relx=0.3, rely=0.1, relwidth=0.10, height=40)

        buttonDelete = ttk.Button(self, text='Удалить')
        buttonDelete.place(relx=0.5, rely=0.1, relwidth=0.10, height=40)

        buttonRefresh = ttk.Button(self, text='Обновить', command=self.refresh_data)
        buttonRefresh.place(relx=0.5, rely=0.8, relwidth=0.10, height=40)

        self.columns = ("ID", "ФИО", "Контакт", "Паспорт", "Дата рождения")
        self.tree = ttk.Treeview(self, columns=self.columns, show="headings")
        for col in self.columns:
            self.tree.heading(col, text=col)
        self.tree.pack(expand=True)

        self.tree.column("#1", stretch=True, width=30, anchor='c')
        self.tree.column("#2", stretch=True, minwidth=200, anchor='c')
        self.tree.column("#3", stretch=True, width=130, anchor='c')
        self.tree.column("#4", stretch=True, width=130, anchor='c')
        self.tree.column("#5", stretch=True, width=130, anchor='c')

        '''
        tree.heading("ID", text="ID")
        tree.heading("ФИО", text="ФИО")
        tree.heading("Контакт", text="Контакт")
        tree.heading("Паспорт", text="Паспорт")
        tree.heading("Дата рождения", text="Дата рождения")
        tree.pack(fill="both", expand=True)
        '''

    def fetch_data(self):
        conn = sqlite3.connect('hotel.db')
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Clients")
        self.rows = cursor.fetchall()
        conn.close()

    def display_data(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        for row in self.rows:
            self.tree.insert("", "end", values=row)

    def refresh_data(self):
        self.fetch_data()
        self.display_data()


# Класс, описывающий создание области БРОНИРОВАНИЕ
class Rent(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        testLbl = ttk.Label(self, background='yellow')
        testLbl.pack(expand=True, fill='both')
        self.place(relx=0.3, y=0, relwidth=0.7, relheight=1)
        self.create_test()

    def create_test(self):
        buttonMenu5 = ttk.Button(self, text='rent')
        buttonMenu5.place(relx=0.27, rely=0.8, relwidth=0.45, height=40)


# Класс, описывающий создание области БРОНИРОВАНИЕ
class Room(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        testLbl = ttk.Label(self, background='brown')
        testLbl.pack(expand=True, fill='both')
        self.place(relx=0.3, y=0, relwidth=0.7, relheight=1)
        self.create_test()

    def create_test(self):
        buttonMenu5 = ttk.Button(self, text='room')
        buttonMenu5.place(relx=0.27, rely=0.8, relwidth=0.45, height=40)


# Класс, описывающий создание меню в левой части окна
class Menu(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        testLbl = ttk.Label(self, background='#620410')
        testLbl.pack(expand=True, fill='both')
        self.place(x=0, y=0, relwidth=0.3, relheight=1)
        self.create_widgets()
        self.update_time()
        

    def create_widgets(self):
        self.buttonMenu1 = ttk.Button(self, text='Клиенты', command=self.show_clients)
        self.buttonMenu1.place(relx=0.27, rely=0.2, relwidth=0.45, height=40)

        self.buttonMenu2 = ttk.Button(self, text='Бронирование', command=self.show_rent)
        self.buttonMenu2.place(relx=0.27, rely=0.3, relwidth=0.45, height=40)

        self.buttonMenu3 = ttk.Button(self, text='Номерной фонд', command=self.show_room)
        self.buttonMenu3.place(relx=0.27, rely=0.4, relwidth=0.45, height=40)

        self.buttonMenu4 = ttk.Button(self, text='Справочники', command = self.show_view)
        self.buttonMenu4.place(relx=0.27, rely=0.5, relwidth=0.45, height=40)

        # Кнопка для перехода на экран "About"
        self.buttonMenu5 = ttk.Button(self, text='О программе', command=self.show_about)
        self.buttonMenu5.place(relx=0.27, rely=0.9, relwidth=0.45, height=40)

        self.buttonMenu6 = ttk.Button(self, text='Вход', command=self.show_login)
        self.buttonMenu6.place(relx=0.27, rely=0.8, relwidth=0.45, height=40)


        labelHide_image_path = os.path.join(sys.path[0], '../img/test.png')

        imageHide = Image.open(labelHide_image_path)
        #imageHide = imageHide.resize((200, 200))
        self.tk_image_hide = ImageTk.PhotoImage(imageHide)
        
        self.labelHide = ttk.Label(self, image=self.tk_image_hide)
        self.labelHide.place(relx=0.27, rely=0.2, relwidth=0.45, height=30)
        #self.labelHide.pack()

        self.disable_buttons()

        # Label для отображения даты и времени
        self.time_label1 = ttk.Label(self, font=("Arial", 10, "bold"), background="#620410", foreground="yellow")
        self.time_label1.place(relx=0.1, rely=0.05, relwidth=0.8, height=30)

        self.time_label2 = ttk.Label(self, font=("Arial", 10, "bold"), background="#620410", foreground="yellow")
        self.time_label2.place(relx=0.1, rely=0.10, relwidth=0.8, height=30)

        self.time_label3 = ttk.Label(self, font=("Arial", 12, "bold"), background="#620410", foreground="yellow")
        self.time_label3.place(relx=0.1, rely=0.15, relwidth=0.8, height=30)

    def update_time(self):
        # Получаем текущую дату и время
        now = datetime.now()
        # Форматируем дату и время
        current_time1 = now.strftime("%d.%m.%Y").title()
        current_time2 = now.strftime("%A").title()
        current_time3 = now.strftime("%H:%M:%S").title()
        # Обновляем текст в Label
        self.time_label1.config(text=current_time1)
        self.time_label2.config(text=current_time2)
        self.time_label3.config(text=current_time3)
        # Вызываем эту функцию снова через 1000 мс (1 секунду)
        self.after(1000, self.update_time)

    def disable_buttons(self):
        self.labelHide.state(['disabled'])

    def enable_buttons(self):
        self.labelHide.place(x=1, y=1, width=100, height=100)

    def set_view_frame(self, view_frame):
        self.view_frame = view_frame

    def show_view(self):
        self.view_frame.tkraise()


    def set_about_frame(self, about_frame):
        self.about_frame = about_frame

    def show_about(self):
        self.about_frame.tkraise()


    def set_clients_frame(self, clients_frame):
        self.clients_frame = clients_frame

    def show_clients(self):
        self.clients_frame.tkraise()


    def set_rent_frame(self, rent_frame):
        self.rent_frame = rent_frame

    def show_rent(self):
        self.rent_frame.tkraise()


    def set_room_frame(self, room_frame):
        self.room_frame = room_frame

    def show_room(self):
        self.room_frame.tkraise()


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
        labelLogo_image_path = os.path.join(sys.path[0], '../img/test.jpg')

        imageLogo = Image.open(labelLogo_image_path)
        self.tk_image_logo = ImageTk.PhotoImage(imageLogo)
        
        self.labelLogo = ttk.Label(self, image=self.tk_image_logo)
        self.labelLogo.place(x=1, y=1, width=100, height=100)

        buttonLoginEnter = ttk.Button(self, text='login', command = self.login)
        buttonLoginEnter.place(relx=0.27, rely=0.65, relwidth=0.45, height=40)

        #buttonLoginEnter.bind = ('<Return>', self.login)

        self.entryLoginLogin = ttk.Entry(self)
        self.entryLoginLogin.place(relx=0.27, rely=0.45, relwidth=0.45, height=40)

        self.entryLoginPassword = ttk.Entry(self, show = '*')
        self.entryLoginPassword.place(relx=0.27, rely=0.55, relwidth=0.45, height=40)

        

    def login(self):
        conn = sqlite3.connect('hotel.db')
        cursor = conn.cursor()

        username = self.entryLoginLogin.get()
        password = self.entryLoginPassword.get()

        cursor.execute('SELECT * FROM Users WHERE name = ? AND password = ?;', (username, password))
        user = cursor.fetchone()

        if user:
            labelLoginCheckOn = tk.Label(self, text = 'Успешный вход', fg = 'green', bg = 'blue')
            labelLoginCheckOn.place(relx=0.27, rely=0.85, relwidth=0.45, height=40)
            
            # Включение кнопок меню после успешного входа
            self.master.menu.enable_buttons()

        else:
            labelLoginCheckOff = tk.Label(self, text = 'Ошибка входа', fg = 'red')
            labelLoginCheckOff.place(relx=0.27, rely=0.85, relwidth=0.45, height=40)

        conn.close()


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
App('АРМ Космос', (1000, 600))