#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Программа для обработки данных от поставщика и загрузки в комплекс.
Поддерживает форматы Excel (.xlsx, .xls) и PDF (для извлечения таблиц).
С графическим интерфейсом на базе tkinter.
"""

import pandas as pd
import numpy as np
import os
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading


def load_file(file_path):
    """
    Загрузка файла (Excel, CSV, DBF). 
    Для PDF требуется дополнительная библиотека pdfplumber.
    """
    file_ext = Path(file_path).suffix.lower()
    
    if file_ext in ['.xlsx', '.xls']:
        return pd.read_excel(file_path)
    elif file_ext == '.csv':
        return pd.read_csv(file_path)
    elif file_ext == '.dbf':
        try:
            from simpledbf import Dbf5
            # DBF файлы обычно используют cp866 (DOS) или cp1251 (Windows) кодировку
            # Пробуем сначала cp866, затем cp1251
            encodings_to_try = ['cp866', 'cp1251', 'latin1']
            df = None
            
            for encoding in encodings_to_try:
                try:
                    dbf = Dbf5(file_path, encoding=encoding)
                    df = dbf.to_dataframe()
                    print(f"DBF файл успешно прочитан с кодировкой: {encoding}")
                    break
                except UnicodeDecodeError:
                    continue
                except Exception as e:
                    # Если ошибка не связана с кодировкой, пробуем следующую
                    error_str = str(e).lower()
                    if "codec" in error_str or "decode" in error_str or "can't decode" in error_str:
                        continue
                    else:
                        # Другая ошибка - пробуем без указания кодировки
                        try:
                            dbf = Dbf5(file_path)
                            df = dbf.to_dataframe()
                            print("DBF файл прочитан с кодировкой по умолчанию")
                            break
                        except:
                            raise
            
            if df is None:
                # Если ни одна кодировка не подошла, пробуем без указания кодировки
                dbf = Dbf5(file_path)
                df = dbf.to_dataframe()
            
            return df
        except ImportError as e:
            error_msg = f"Ошибка импорта simpledbf: {str(e)}. Установите библиотеку: pip install simpledbf"
            print(error_msg)
            raise ImportError(error_msg)
        except Exception as e:
            error_msg = f"Ошибка при чтении DBF файла: {str(e)}. Проверьте целостность файла и кодировку."
            print(error_msg)
            raise ValueError(error_msg)
    elif file_ext == '.pdf':
        try:
            import pdfplumber
            tables = []
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        tables.append(table)
            if not tables:
                raise ValueError("Таблицы не найдены в PDF файле")
            # Объединяем все таблицы и создаем DataFrame
            all_rows = []
            for table in tables:
                all_rows.extend(table)
            if all_rows:
                df = pd.DataFrame(all_rows[1:], columns=all_rows[0])
                return df
            else:
                raise ValueError("Не удалось извлечь данные из PDF")
        except ImportError:
            print("Для работы с PDF установите библиотеку: pip install pdfplumber")
            raise
    else:
        raise ValueError(f"Неподдерживаемый формат файла: {file_ext}")


def process_data(supplier_df, upload_df, log_callback=None):
    """
    Обработка данных согласно требованиям:
    1. Перенос данных из столбцов SUM_N1..16, ZADOLG1..16, MZADOLG1..16
    2. Проверка GLAVA и SUM_N для установки DOGOVOR
    3. Добавление столбца "Определять тариф по площади"
    4. Соединение по KOD
    """
    
    def log(message):
        if log_callback:
            log_callback(message)
        else:
            print(message)
    
    # Создаем копию файла для загрузки
    result_df = upload_df.copy()
    
    # Проверяем наличие столбца KOD в файле для загрузки
    if 'KOD' not in upload_df.columns and 'kod' not in [c.lower() for c in upload_df.columns]:
        # Ищем столбец с похожим названием (регистронезависимо)
        kod_col = None
        for col in upload_df.columns:
            if col.lower() in ['kod', 'code', 'код', 'kodu']:
                kod_col = col
                break
        if kod_col is None:
            log("Ошибка: столбец KOD не найден в файле 'Данные для загрузки'")
            return None
    
    # Нормализуем имена столбцов (приводим к верхнему регистру для удобства)
    upload_df.columns = [str(c).upper().strip() for c in upload_df.columns]
    supplier_df.columns = [str(c).upper().strip() for c in supplier_df.columns]
    result_df.columns = [str(c).upper().strip() for c in result_df.columns]
    
    # Определяем имя столбца KOD
    kod_col_upload = 'KOD' if 'KOD' in upload_df.columns else None
    if kod_col_upload is None:
        for col in upload_df.columns:
            if col.upper() == 'KOD':
                kod_col_upload = col
                break
    
    kod_col_supplier = 'KOD' if 'KOD' in supplier_df.columns else None
    if kod_col_supplier is None:
        for col in supplier_df.columns:
            if col.upper() == 'KOD':
                kod_col_supplier = col
                break
    
    if kod_col_upload is None:
        log("Ошибка: столбец KOD не найден в файле 'Данные для загрузки'")
        return None
    
    if kod_col_supplier is None:
        log("Предупреждение: столбец KOD не найден в файле 'Данные от поставщика'. Данные не будут перенесены.")
        # Добавляем столбец "Определять тариф по площади" со значением 1
        result_df['ОПРЕДЕЛЯТЬ ТАРИФ ПО ПЛОЩАДИ'] = 1
        return result_df
    
    # Списки столбцов для обработки
    sum_cols = [f'SUM_N{i}' for i in range(1, 17)]
    zadolg_cols = [f'ZADOLG{i}' for i in range(1, 17)]
    mzadolg_cols = [f'MZADOLG{i}' for i in range(1, 17)]
    glava_cols = [f'GLAVA{i}' for i in range(1, 17)]
    dogovor_cols = [f'DOGOVOR{i}' for i in range(1, 17)]
    
    # Проверяем какие столбцы существуют в файлах
    existing_sum_supplier = [col for col in sum_cols if col in supplier_df.columns]
    existing_zadolg_supplier = [col for col in zadolg_cols if col in supplier_df.columns]
    existing_mzadolg_supplier = [col for col in mzadolg_cols if col in supplier_df.columns]
    existing_glava_supplier = [col for col in glava_cols if col in supplier_df.columns]
    
    # Группируем данные поставщика по KOD и выбираем максимальные значения для числовых полей
    supplier_dict = {}
    for idx, row in supplier_df.iterrows():
        kod_value = row[kod_col_supplier]
        if pd.notna(kod_value):
            if kod_value in supplier_dict:
                # Если KOD уже существует, выбираем максимальные значения
                existing_row = supplier_dict[kod_value]
                for col in supplier_df.columns:
                    current_val = row.get(col, 0)
                    existing_val = existing_row.get(col, 0)
                    
                    # Для числовых значений выбираем максимум
                    try:
                        if pd.api.types.is_numeric_dtype(type(current_val)) or pd.api.types.is_numeric_dtype(type(existing_val)):
                            # Пробуем преобразовать к числу
                            current_num = float(current_val) if pd.notna(current_val) else 0
                            existing_num = float(existing_val) if pd.notna(existing_val) else 0
                            if current_num > existing_num:
                                existing_row[col] = current_val
                        else:
                            # Для нечисловых значений оставляем первое непустое
                            if pd.isna(existing_val) or existing_val == '' or existing_val is None:
                                existing_row[col] = current_val
                    except (ValueError, TypeError):
                        # Если не удалось преобразовать к числу, оставляем существующее значение
                        pass
            else:
                supplier_dict[kod_value] = row.to_dict()
    
    log(f"Загружено {len(supplier_dict)} уникальных кодов из файла поставщика")
    
    missing_kods = []
    
    # Обрабатываем каждую строку в файле для загрузки
    for idx in result_df.index:
        kod_value = result_df.loc[idx, kod_col_upload]
        
        # Проверяем наличие KOD в файле для загрузки
        if pd.isna(kod_value):
            continue
        
        # Ищем соответствующую запись в данных поставщика
        if kod_value in supplier_dict:
            supplier_row = supplier_dict[kod_value]
            
            # Переносим данные SUM_N, ZADOLG, MZADOLG и обрабатываем DOGOVOR
            for i in range(1, 17):
                sum_col = f'SUM_N{i}'
                zadolg_col = f'ZADOLG{i}'
                mzadolg_col = f'MZADOLG{i}'
                glava_col = f'GLAVA{i}'
                dogovor_col = f'DOGOVOR{i}'
                
                # Переносим данные из файла поставщика только если столбцы существуют в обоих файлах
                supplier_sum_val = 0
                supplier_zadolg_val = 0
                supplier_mzadolg_val = 0
                
                if sum_col in supplier_df.columns:
                    supplier_sum_val = supplier_row.get(sum_col, 0)
                
                if zadolg_col in supplier_df.columns:
                    supplier_zadolg_val = supplier_row.get(zadolg_col, 0)
                
                if mzadolg_col in supplier_df.columns:
                    supplier_mzadolg_val = supplier_row.get(mzadolg_col, 0)
                
                # Обновляем только целевые столбцы в файле загрузки
                if sum_col in result_df.columns:
                    result_df.loc[idx, sum_col] = supplier_sum_val
                
                if mzadolg_col in result_df.columns:
                    result_df.loc[idx, mzadolg_col] = supplier_mzadolg_val
                
                # Обработка ZADOLG: ставим 1 если MZADOLG > 0 или не целое число
                if zadolg_col in result_df.columns and mzadolg_col in result_df.columns:
                    mzadolg_value = result_df.loc[idx, mzadolg_col]
                    
                    # Проверяем является ли значение числом
                    try:
                        mzadolg_num = float(mzadolg_value) if pd.notna(mzadolg_value) else 0
                        
                        # Ставим 1 если:
                        # 1. MZADOLG > 0
                        # 2. MZADOLG содержит дробную часть (не целое число)
                        if mzadolg_num > 0 or (mzadolg_num != int(mzadolg_num)):
                            result_df.loc[idx, zadolg_col] = 1
                        else:
                            result_df.loc[idx, zadolg_col] = 0
                    except (ValueError, TypeError):
                        # Если не число, проверяем наличие запятой в строковом представлении
                        mzadolg_str = str(mzadolg_value)
                        if ',' in mzadolg_str or '.' in mzadolg_str:
                            result_df.loc[idx, zadolg_col] = 1
                        else:
                            result_df.loc[idx, zadolg_col] = 0
                
                # Проверка для DOGOVOR: если GLAVA и SUM_N имеют значения, ставим 1, иначе 0
                if dogovor_col in result_df.columns:
                    glava_value = result_df.loc[idx, glava_col] if glava_col in result_df.columns else None
                    sum_value = result_df.loc[idx, sum_col] if sum_col in result_df.columns else None
                    
                    # Проверяем наличие значений (не NaN и не 0)
                    glava_has_value = pd.notna(glava_value) and glava_value != '' and glava_value != 0
                    sum_has_value = pd.notna(sum_value) and sum_value != '' and sum_value != 0
                    
                    if glava_has_value and sum_has_value:
                        result_df.loc[idx, dogovor_col] = 1
                    else:
                        result_df.loc[idx, dogovor_col] = 0
        else:
            # KOD не найден в данных поставщика
            missing_kods.append(kod_value)
            
            # Устанавливаем 0 в поля SUM_N, ZADOLG, MZADOLG, DOGOVOR
            for i in range(1, 17):
                sum_col = f'SUM_N{i}'
                zadolg_col = f'ZADOLG{i}'
                mzadolg_col = f'MZADOLG{i}'
                dogovor_col = f'DOGOVOR{i}'
                
                if sum_col in result_df.columns:
                    result_df.loc[idx, sum_col] = 0
                if mzadolg_col in result_df.columns:
                    result_df.loc[idx, mzadolg_col] = 0
                if zadolg_col in result_df.columns:
                    # Для ZADOLG тоже ставим 0 так как MZADOLG = 0
                    result_df.loc[idx, zadolg_col] = 0
                if dogovor_col in result_df.columns:
                    result_df.loc[idx, dogovor_col] = 0
    
    # Выводим сообщения об отсутствующих KOD
    if missing_kods:
        for kod in missing_kods:
            log(f"Код семьи не найден: {kod}")
    
    # Добавляем столбец "Определять тариф по площади" со значением 1
    result_df['ОПРЕДЕЛЯТЬ ТАРИФ ПО ПЛОЩАДИ'] = 1
    
    return result_df


class DataProcessorApp:
    """Графический интерфейс приложения."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Обработка данных от поставщика")
        self.root.geometry("800x600")
        
        # Переменные для хранения путей к файлам
        self.supplier_file = tk.StringVar()
        self.upload_file = tk.StringVar()
        self.output_file = tk.StringVar()
        
        self.create_widgets()
    
    def create_widgets(self):
        """Создание элементов интерфейса."""
        # Заголовок
        title_label = tk.Label(self.root, text="Программа обработки данных", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)
        
        # Фрейм для выбора файлов
        file_frame = tk.LabelFrame(self.root, text="Выбор файлов", padx=10, pady=10)
        file_frame.pack(fill="x", padx=20, pady=10)
        
        # Выбор файла поставщика
        tk.Label(file_frame, text="Файл от поставщика (Excel/DBF/PDF):").grid(row=0, column=0, sticky="w", pady=5)
        tk.Entry(file_frame, textvariable=self.supplier_file, width=50).grid(row=0, column=1, padx=10)
        tk.Button(file_frame, text="Обзор...", command=self.browse_supplier_file).grid(row=0, column=2)
        
        # Выбор файла для загрузки
        tk.Label(file_frame, text="Файл для загрузки (Excel):").grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(file_frame, textvariable=self.upload_file, width=50).grid(row=1, column=1, padx=10)
        tk.Button(file_frame, text="Обзор...", command=self.browse_upload_file).grid(row=1, column=2)
        
        # Кнопка запуска обработки
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.process_btn = tk.Button(btn_frame, text="Обработать данные", command=self.start_processing, 
                                      bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), padx=20, pady=10)
        self.process_btn.pack()
        
        # Лог операций
        log_frame = tk.LabelFrame(self.root, text="Журнал операций", padx=10, pady=10)
        log_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, wrap=tk.WORD)
        self.log_text.pack(fill="both", expand=True)
        
        # Индикатор прогресса
        self.progress_var = tk.BooleanVar(value=False)
        self.status_label = tk.Label(self.root, text="", font=("Arial", 10))
        self.status_label.pack(pady=5)
    
    def browse_supplier_file(self):
        """Выбор файла поставщика."""
        filename = filedialog.askopenfilename(
            title="Выберите файл от поставщика",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("DBF files", "*.dbf"), ("PDF files", "*.pdf"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.supplier_file.set(filename)
            self.log(f"Выбран файл поставщика: {filename}")
    
    def browse_upload_file(self):
        """Выбор файла для загрузки."""
        filename = filedialog.askopenfilename(
            title="Выберите файл для загрузки в комплекс",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.upload_file.set(filename)
            self.log(f"Выбран файл для загрузки: {filename}")
    
    def log(self, message):
        """Добавление сообщения в лог."""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def start_processing(self):
        """Запуск обработки в отдельном потоке."""
        if not self.supplier_file.get():
            messagebox.showerror("Ошибка", "Не выбран файл от поставщика!")
            return
        
        if not self.upload_file.get():
            messagebox.showerror("Ошибка", "Не выбран файл для загрузки!")
            return
        
        # Блокируем кнопку и запускаем обработку в потоке
        self.process_btn.config(state="disabled")
        self.progress_var.set(True)
        self.status_label.config(text="Обработка данных...")
        
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()
    
    def process_files(self):
        """Обработка файлов."""
        try:
            self.log("=" * 60)
            self.log("Начало обработки данных...")
            
            # Загрузка файлов
            self.log("Загрузка файла поставщика...")
            supplier_df = load_file(self.supplier_file.get())
            self.log(f"✓ Файл поставщика загружен: {len(supplier_df)} записей")
            
            self.log("Загрузка файла для загрузки...")
            upload_df = load_file(self.upload_file.get())
            self.log(f"✓ Файл для загрузки загружен: {len(upload_df)} записей")
            
            # Обработка данных
            self.log("Обработка данных...")
            result_df = process_data(supplier_df, upload_df, log_callback=self.log)
            
            if result_df is None:
                self.log("Ошибка при обработке данных!")
                self.root.after(0, lambda: messagebox.showerror("Ошибка", "Ошибка при обработке данных!"))
                return
            
            # Формирование имени выходного файла
            upload_file_path = Path(self.upload_file.get())
            output_filename = f"{upload_file_path.stem}_processed{upload_file_path.suffix}"
            output_path = upload_file_path.parent / output_filename
            
            self.output_file.set(str(output_path))
            
            # Сохранение результата
            self.log(f"Сохранение результата в: {output_path}")
            if output_path.suffix.lower() in ['.xlsx', '.xls']:
                result_df.to_excel(output_path, index=False)
            else:
                result_df.to_csv(output_path, index=False, encoding='utf-8-sig')
            
            self.log(f"✓ Файл успешно сохранен: {output_path}")
            self.log("=" * 60)
            self.log("Обработка завершена успешно!")
            
            # Показываем сообщение об успехе
            self.root.after(0, lambda: messagebox.showinfo(
                "Успех", 
                f"Обработка завершена!\n\nРезультат сохранен в:\n{output_path}\n\nВы можете скачать файл с этим именем."
            ))
            
        except Exception as e:
            error_msg = f"Ошибка: {str(e)}"
            self.log(error_msg)
            self.root.after(0, lambda: messagebox.showerror("Ошибка", error_msg))
        
        finally:
            # Разблокируем кнопку
            self.root.after(0, self.enable_button)
    
    def enable_button(self):
        """Разблокировка кнопки обработки."""
        self.process_btn.config(state="normal")
        self.progress_var.set(False)
        self.status_label.config(text="")


def main():
    """Основная функция программы."""
    root = tk.Tk()
    app = DataProcessorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
