import sqlite3
import os

print(f"Текущая директория: {os.getcwd()}")
print(f"Путь к hotel.db: {os.path.abspath('hotel.db')}")
print(f"Файл существует: {os.path.exists('hotel.db')}")

if os.path.exists('hotel.db'):
    print(f"Размер файла: {os.path.getsize('hotel.db')} байт")
    
    conn = sqlite3.connect('hotel.db')
    cursor = conn.cursor()
    
    cursor.execute("SELECT name, sql FROM sqlite_master WHERE type='table' ORDER BY name")
    tables = cursor.fetchall()
    
    print(f"\n=== Все таблицы ({len(tables)}) ===")
    for table in tables:
        print(f"\nТаблица: {table[0]}")
        print(f"SQL: {table[1]}")
        
        try:
            cursor.execute(f"SELECT COUNT(*) FROM '{table[0]}'")
            count = cursor.fetchone()[0]
            print(f"  Записей: {count}")
            
            if count > 0:
                cursor.execute(f"SELECT * FROM '{table[0]}' LIMIT 3")
                rows = cursor.fetchall()
                for row in rows:
                    print(f"    {row}")
        except Exception as e:
            print(f"  Ошибка при чтении: {e}")
    
    conn.close()
else:
    print("\nФайл hotel.db НЕ НАЙДЕН!")
    print("\nФайлы .db в директории:")
    for file in os.listdir('.'):
        if file.endswith('.db'):
            print(f"  {file}")