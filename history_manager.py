import sqlite3
import os
from datetime import datetime
import calendar

class HistoryManager:
    """
    Менеджер истории на базе SQLite. 
    Позволяет сохранять проблемы и быстро получать статистику для календаря.
    """
    def __init__(self, db_name="production_history.db"):
        self.db_name = db_name
        self._init_db()

    def _init_db(self):
        """Создает таблицу, если она не существует."""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS problems (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                record_date TEXT,
                shift_date TEXT,     -- Дата смены (строка YYYY-MM-DD для сортировки)
                shift_day INTEGER,   -- День месяца (число)
                shift_month INTEGER, -- Месяц (число)
                shift_year INTEGER,  -- Год (число)
                line_name TEXT,
                shift_type TEXT,
                time_val REAL,
                problem_type TEXT,
                description TEXT,
                comment TEXT
            )
        ''')
        conn.commit()
        conn.close()

    def save_problems(self, lines_data, d, m, y):
        """
        Сохраняет список проблем за конкретную дату.
        Сначала удаляет старые записи за эту дату (чтобы избежать дублей при перезаписи),
        затем вставляет актуальные.
        """
        # Преобразуем месяц (строка -> число) для удобства
        months_map = {
            "Январь": 1, "Февраль": 2, "Март": 3, "Апрель": 4, "Май": 5, "Июнь": 6,
            "Июль": 7, "Август": 8, "Сентябрь": 9, "Октябрь": 10, "Ноябрь": 11, "Декабрь": 12
        }
        month_num = months_map.get(m, 0)
        # Формат даты YYYY-MM-DD для поиска
        date_iso = f"{y}-{month_num:02d}-{d:02d}"
        
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        
        try:
            # 1. Очистка старых данных за эту дату
            cursor.execute("DELETE FROM problems WHERE shift_date = ?", (date_iso,))
            
            # 2. Вставка новых
            count = 0
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            for ld in lines_data:
                if not ld.problems:
                    continue
                for p in ld.problems:
                    cursor.execute('''
                        INSERT INTO problems (
                            record_date, shift_date, shift_day, shift_month, shift_year,
                            line_name, shift_type, time_val, problem_type, description, comment
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        now_str, date_iso, d, month_num, y,
                        ld.line_name, p.shift, p.time_val, p.type_val, p.formulation, p.comment
                    ))
                    count += 1
            
            conn.commit()
            return True, f"Записано проблем: {count}"
        except Exception as e:
            return False, str(e)
        finally:
            conn.close()

    def get_month_stats(self, month, year):
        """
        Возвращает словарь: {день: суммарные_минуты_простоя}
        Используется для раскраски календаря.
        """
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT shift_day, SUM(time_val) 
            FROM problems 
            WHERE shift_month = ? AND shift_year = ?
            GROUP BY shift_day
        ''', (month, year))
        
        rows = cursor.fetchall()
        conn.close()
        
        # Превращаем в словарь {1: 120.5, 5: 30.0, ...}
        stats = {row[0]: row[1] for row in rows}
        return stats

    def get_day_details(self, day, month, year):
        """
        Возвращает все проблемы за конкретный день.
        """
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT line_name, time_val, problem_type, description, comment
            FROM problems 
            WHERE shift_day = ? AND shift_month = ? AND shift_year = ?
            ORDER BY time_val DESC
        ''', (day, month, year))
        data = cursor.fetchall()
        conn.close()
        return data