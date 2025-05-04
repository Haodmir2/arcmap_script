# -*- coding: utf-8 -*-
import os
import sys
import logging
import traceback
import datetime
import re

# Настройка логирования
log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "select_gdb_log.txt")

# Для Python 2.7 нельзя использовать encoding в basicConfig, делаем по-другому
if sys.version_info < (3, 0):
    # Создаем обработчик с явной кодировкой для Windows
    import codecs
    # На Windows для Python 2.7 лучше использовать cp1251 вместо utf-8
    log_handler = logging.FileHandler(log_file)
    log_handler.setLevel(logging.DEBUG)
    log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    
    # Создаем логгер и добавляем обработчик
    logger = logging.getLogger('')
    logger.setLevel(logging.DEBUG)
    logger.addHandler(log_handler)
    
    # Убираем стандартные обработчики, чтобы избежать дублирования
    for hdlr in logger.handlers[:]:
        if isinstance(hdlr, logging.StreamHandler) or isinstance(hdlr, logging.FileHandler):
            if hdlr != log_handler:
                logger.removeHandler(hdlr)
    
    # Функция для перекодирования строк
    def encode_if_needed(s):
        if isinstance(s, unicode):
            return s  # Сохраняем unicode для записи в файл
        try:
            return unicode(s, 'utf-8')  # Пытаемся декодировать из utf-8
        except:
            try:
                return unicode(s, 'cp1251')  # Пытаемся декодировать из cp1251
            except:
                return unicode(str(s), errors='replace')  # Если не получается, заменяем на символы
    
    # Переопределяем все функции логирования
    def encoded_log(original_log, msg, *args, **kwargs):
        msg = encode_if_needed(msg)
        original_log(msg, *args, **kwargs)
    
    # Сохраняем оригинальные функции логирования
    original_debug = logger.debug
    original_info = logger.info
    original_warning = logger.warning
    original_error = logger.error
    original_critical = logger.critical
    
    # Переопределяем функции логирования
    logger.debug = lambda msg, *args, **kwargs: encoded_log(original_debug, msg, *args, **kwargs)
    logger.info = lambda msg, *args, **kwargs: encoded_log(original_info, msg, *args, **kwargs)
    logger.warning = lambda msg, *args, **kwargs: encoded_log(original_warning, msg, *args, **kwargs)
    logger.error = lambda msg, *args, **kwargs: encoded_log(original_error, msg, *args, **kwargs)
    logger.critical = lambda msg, *args, **kwargs: encoded_log(original_critical, msg, *args, **kwargs)
    
    # Добавляем консольный вывод для отладки
    console = logging.StreamHandler()
    console.setLevel(logging.DEBUG)
    console.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.addHandler(console)
    
else:
    # Для Python 3 можно использовать encoding
    logging.basicConfig(
        filename=log_file,
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        encoding='utf-8'
    )

# Функция для логирования ошибок с трассировкой стека
def log_exception(e, message="Произошла ошибка"):
    logging.error(message)
    logging.error(str(e))
    logging.error(traceback.format_exc())

# Запись о запуске скрипта
logging.info("="*50)
logging.info("Запуск скрипта {}".format(datetime.datetime.now()))
logging.info("Версия Python: {}".format(sys.version))
logging.info("Путь к скрипту: {}".format(os.path.abspath(__file__)))
logging.info("="*50)

# Настройка кодировки для Python 2.7
if sys.version_info[0] < 3:
    reload(sys)
    if hasattr(sys, 'setdefaultencoding'):
        sys.setdefaultencoding('utf-8')
else:
    import importlib
    if hasattr(importlib, 'reload'):
        reload = importlib.reload

# Совместимость с Python 2 и 3 для tkinter
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog
    from tkinter import ttk
    logging.info("Импортированы модули tkinter для Python 3")
except ImportError:
    # Для Python 2 (ArcGIS 10.4)
    import Tkinter as tk
    import tkFileDialog as filedialog
    import tkMessageBox as messagebox
    import tkSimpleDialog as simpledialog
    import ttk
    logging.info("Импортированы модули tkinter для Python 2")

# Проверка доступности arcpy
try:
    import arcpy
    ARCPY_AVAILABLE = True
    logging.info("Модуль arcpy успешно импортирован")
    logging.info("Версия arcpy: {}".format(arcpy.GetInstallInfo().get('Version', 'Неизвестно')))
except ImportError as e:
    ARCPY_AVAILABLE = False
    logging.warning("Модуль arcpy не найден: {}".format(str(e)))
    print("ВНИМАНИЕ: Модуль arcpy не найден. Некоторые функции проверки будут недоступны.")
    print("Для полной функциональности запустите скрипт через Python, поставляемый с ArcGIS/ArcMap.")

class GDBSelector:
    def __init__(self, master):
        self.master = master
        self.gdb_path = None
        
        # Настройка окна
        master.title("Выбор базы геоданных")
        master.geometry("500x200")
        
        # Фрейм для содержимого
        self.main_frame = tk.Frame(master, padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        self.label = tk.Label(self.main_frame, 
                             text="Выберите базу геоданных (GDB)",
                             font=("Arial", 12))
        self.label.pack(pady=10)
        
        # Предупреждение, если arcpy недоступен
        if not ARCPY_AVAILABLE:
            self.warning_label = tk.Label(self.main_frame,
                                         text="ВНИМАНИЕ: Модуль arcpy не найден. Проверка GDB будет ограничена.",
                                         fg="red",
                                         font=("Arial", 10))
            self.warning_label.pack(pady=5)
        
        # Поле для отображения пути
        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(self.main_frame, 
                                  textvariable=self.path_var,
                                  width=50)
        self.path_entry.pack(pady=10, fill=tk.X)
        
        # Кнопки
        self.button_frame = tk.Frame(self.main_frame)
        self.button_frame.pack(pady=10)
        
        self.browse_button = tk.Button(self.button_frame, 
                                     text="Обзор...", 
                                     command=self.browse_gdb)
        self.browse_button.pack(side=tk.LEFT, padx=5)
        
        self.select_button = tk.Button(self.button_frame, 
                                     text="Выбрать", 
                                     command=self.select_gdb)
        self.select_button.pack(side=tk.LEFT, padx=5)
        
        self.cancel_button = tk.Button(self.button_frame, 
                                     text="Отмена", 
                                     command=master.destroy)
        self.cancel_button.pack(side=tk.LEFT, padx=5)
    
    def browse_gdb(self):
        # Открыть диалог выбора файла с явным указанием родительского окна
        initial_dir = os.path.expanduser("~") if not self.path_var.get() else os.path.dirname(self.path_var.get())
        
        # Явно указываем родительское окно и делаем диалог модальным
        path = filedialog.askdirectory(
            parent=self.master,
            title="Выберите базу геоданных GDB",
            initialdir=initial_dir
        )
        
        # Проверка, что выбранный путь - GDB
        if path and path.strip():
            if os.path.isdir(path) and os.path.basename(path).endswith('.gdb'):
                self.path_var.set(path)
                print("Выбрана база геоданных: {}".format(path))
            else:
                messagebox.showerror("Ошибка", "Выбранный каталог не является базой геоданных GDB")
        
        # Возвращаем фокус на основное окно
        self.master.focus_set()
    
    def select_gdb(self):
        path = self.path_var.get()
        
        if not path:
            messagebox.showerror("Ошибка", "Путь к базе геоданных не указан")
            return
            
        if not os.path.exists(path):
            messagebox.showerror("Ошибка", "Указанный путь не существует")
            return
        
        # Базовая проверка формата GDB    
        if not os.path.basename(path).endswith('.gdb'):
            messagebox.showerror("Ошибка", "Путь не указывает на файл с расширением .gdb")
            return
            
        # Расширенная проверка с arcpy, если доступен
        if ARCPY_AVAILABLE:
            try:
                # Проверка валидности базы геоданных
                desc = arcpy.Describe(path)
                if desc.dataType == "Workspace":
                    self.gdb_path = path
                    messagebox.showinfo("Успех", "База геоданных выбрана: {}".format(path))
                    self.master.destroy()
                else:
                    messagebox.showerror("Ошибка", "Выбранный путь не является базой геоданных")
                    return
            except arcpy.ExecuteError:
                messagebox.showerror("Ошибка ArcPy", arcpy.GetMessages(2))
                return
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))
                return
        else:
            # Упрощенная проверка, если arcpy недоступен
            # Проверяем наличие структуры GDB (базовая проверка)
            gdb_system_files = ['a00000001.gdbindexes', 'a00000001.gdbtable', 'a00000001.gdbtablx',
                               'a00000004.CatItemsByPhysicalName.atx', 'a00000004.gdbindexes']
            
            # Проверяем хотя бы один файл из списка
            found_at_least_one = False
            for file_name in gdb_system_files:
                if os.path.exists(os.path.join(path, file_name)):
                    found_at_least_one = True
                    break
            
            if found_at_least_one:
                self.gdb_path = path
                messagebox.showinfo("Успех", "База геоданных выбрана: {}\n\n(Примечание: Полная проверка GDB недоступна без модуля arcpy)".format(path))
                self.master.destroy()
            else:
                messagebox.showwarning("Предупреждение", 
                                      "Каталог имеет расширение .gdb, но не содержит стандартных файлов GDB.\n\n"
                                      "Без модуля arcpy невозможно полностью проверить, является ли это корректной базой геоданных.")
                if messagebox.askyesno("Подтверждение", "Хотите всё равно использовать этот путь?"):
                    self.gdb_path = path
                    self.master.destroy()

class OldForestGDBSelector:
    def __init__(self, master):
        self.master = master
        self.db_path = None
        self.feature_dataset_path = None
        
        # Настройка окна
        master.title("Выбор базы данных прошлого тура")
        master.geometry("500x200")
        
        # Фрейм для содержимого
        self.main_frame = tk.Frame(master, padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        self.label = tk.Label(self.main_frame, 
                             text="Выберите базу данных прошлого тура (GDB или MDB)",
                             font=("Arial", 12))
        self.label.pack(pady=10)
        
        # Предупреждение, если arcpy недоступен
        if not ARCPY_AVAILABLE:
            self.warning_label = tk.Label(self.main_frame,
                                         text="ВНИМАНИЕ: Модуль arcpy не найден. Проверка базы данных будет ограничена.",
                                         fg="red",
                                         font=("Arial", 10))
            self.warning_label.pack(pady=5)
        
        # Поле для отображения пути
        self.path_var = tk.StringVar()
        self.path_entry = tk.Entry(self.main_frame, 
                                  textvariable=self.path_var,
                                  width=50)
        self.path_entry.pack(pady=10, fill=tk.X)
        
        # Кнопки
        self.button_frame = tk.Frame(self.main_frame)
        self.button_frame.pack(pady=10)
        
        self.browse_button = tk.Button(self.button_frame, 
                                     text="Обзор...", 
                                     command=self.browse_db)
        self.browse_button.pack(side=tk.LEFT, padx=5)
        
        self.select_button = tk.Button(self.button_frame, 
                                     text="Выбрать", 
                                     command=self.select_paths)
        self.select_button.pack(side=tk.LEFT, padx=5)
        
        self.cancel_button = tk.Button(self.button_frame, 
                                     text="Отмена", 
                                     command=master.destroy)
        self.cancel_button.pack(side=tk.LEFT, padx=5)
    
    def browse_db(self):
        # Открыть диалог выбора файла или каталога
        initial_dir = os.path.expanduser("~") if not self.path_var.get() else os.path.dirname(self.path_var.get())
        
        # Для GDB нужен выбор каталога, для MDB нужен выбор файла
        # Сначала пробуем выбрать каталог (для GDB)
        path = filedialog.askdirectory(
            parent=self.master,
            title="Выберите базу геоданных GDB прошлого тура",
            initialdir=initial_dir
        )
        
        # Если пользователь отменил выбор каталога или выбранный каталог не GDB,
        # предлагаем выбрать MDB файл
        if not path or not (os.path.isdir(path) and os.path.basename(path).endswith('.gdb')):
            # Если каталог не выбран или не является GDB, предлагаем выбрать MDB
            mdb_filetypes = [("Microsoft Access Database", "*.mdb")]
            path = filedialog.askopenfilename(
                parent=self.master,
                title="Выберите базу данных MDB прошлого тура",
                initialdir=initial_dir,
                filetypes=mdb_filetypes
            )
        
        # Проверка, что выбранный путь - GDB или MDB
        if path and path.strip():
            is_gdb = os.path.isdir(path) and os.path.basename(path).endswith('.gdb')
            is_mdb = os.path.isfile(path) and os.path.basename(path).lower().endswith('.mdb')
            
            if is_gdb or is_mdb:
                self.path_var.set(path)
                if is_gdb:
                    print("Выбрана база геоданных GDB прошлого тура: {}".format(path))
                else:
                    print("Выбрана база данных MDB прошлого тура: {}".format(path))
            else:
                messagebox.showerror("Ошибка", "Выбранный путь не является ни базой геоданных GDB, ни базой данных MDB")
        
        # Возвращаем фокус на основное окно
        self.master.focus_set()
    
    def select_paths(self):
        path = self.path_var.get()
        
        if not path:
            messagebox.showerror("Ошибка", "Путь к базе данных не указан")
            return
            
        if not os.path.exists(path):
            messagebox.showerror("Ошибка", "Указанный путь не существует")
            return
        
        # Проверка формата базы данных
        is_gdb = os.path.isdir(path) and os.path.basename(path).endswith('.gdb')
        is_mdb = os.path.isfile(path) and os.path.basename(path).lower().endswith('.mdb')
        
        if not (is_gdb or is_mdb):
            messagebox.showerror("Ошибка", "Путь не указывает на файл с расширением .gdb или .mdb")
            return
            
        # Расширенная проверка с arcpy, если доступен
        if ARCPY_AVAILABLE:
            try:
                # Проверка валидности базы данных
                desc = arcpy.Describe(path)
                if desc.dataType == "Workspace":
                    self.db_path = path
                    messagebox.showinfo("Успех", "База данных выбрана: {}".format(path))
                    self.master.destroy()
                else:
                    messagebox.showerror("Ошибка", "Выбранный путь не является базой данных")
                    return
            except arcpy.ExecuteError:
                messagebox.showerror("Ошибка ArcPy", arcpy.GetMessages(2))
                return
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))
                return
        else:
            # Упрощенная проверка, если arcpy недоступен
            messagebox.showinfo("Внимание", "Без модуля arcpy невозможно полностью проверить корректность базы данных.")
            self.db_path = path
            self.master.destroy()

class LabelClassProcessor:
    def __init__(self, db_path, shortened_name):
        self.db_path = db_path
        self.shortened_name = shortened_name
        self.labels_classes = []
        self.is_mdb = os.path.isfile(db_path) and os.path.basename(db_path).lower().endswith('.mdb')
        # Добавляем свойство для хранения пути к GDB
        self.gdb_path = None  # Будет установлено при вызове process_identity
    
    def find_label_classes(self):
        """Поиск классов с 'надпис' и цифрой в имени"""
        if not ARCPY_AVAILABLE:
            logging.error("Модуль arcpy недоступен")
            messagebox.showerror("Ошибка", "Модуль arcpy не доступен для поиска классов")
            return False, "Модуль arcpy недоступен"
        
        try:
            logging.info("Начало поиска классов надписей")
            logging.info("Параметры: База данных={}".format(self.db_path))
            
            # Очищаем рабочее пространство
            arcpy.ClearWorkspaceCache_management()
            
            # Найденные классы будем хранить в структуре {номер: путь_к_классу}
            found_classes = {}
            numbers_found = set()
            
            # Все найденные классы
            all_feature_classes = []
            
            # Классы с надписями, но с проблемами в определении номера
            problem_classes = []
            
            # Шаблоны для поиска надписей с цифрами
            pattern1 = re.compile(r'.*надпис.*\d+.*', re.IGNORECASE)  # надпис, затем цифра
            pattern2 = re.compile(r'.*\d+.*надпис.*', re.IGNORECASE)  # цифра, затем надпис
            pattern3 = re.compile(r'.*лист.*\d+.*надпис.*', re.IGNORECASE)  # слово лист, затем цифра, затем надпис
            pattern4 = re.compile(r'.*[_](\d+)[_].*надпис', re.IGNORECASE)  # цифра в имени перед надписью, разделенная подчеркиваниями
            pattern5 = re.compile(r'.*надпис.*[_](\d+)[_]', re.IGNORECASE)  # цифра в имени после надписи, разделенная подчеркиваниями
            
            # Функция для проверки имени и добавления в список найденных
            def check_and_add_class(class_path, class_name):
                logging.info("Проверка класса: {}".format(class_name))
                
                # Добавляем класс в общий список всех классов
                all_feature_classes.append((class_name, class_path))
                
                # Проверяем соответствие шаблонам
                has_nadpis = "надпис" in class_name.lower()
                matches_pattern = (pattern1.match(class_name) or pattern2.match(class_name) or 
                                   pattern3.match(class_name) or pattern4.match(class_name) or
                                   pattern5.match(class_name))
                
                if has_nadpis and matches_pattern:
                    logging.info("Найден класс с надписью: {}".format(class_name))
                    
                    # Находим все цифры в имени
                    digits_groups = re.findall(r'\d+', class_name)
                    
                    if digits_groups:
                        # Если несколько цифр в имени
                        if len(digits_groups) > 1:
                            logging.warning("Найдено несколько цифр в имени класса: {}, цифры: {}".format(class_name, digits_groups))
                            
                            # Если есть 'лист' в имени, ищем цифру сразу после него
                            if "лист" in class_name.lower():
                                # Ищем шаблон "лист_X" или "лист X" где X - цифра
                                list_number_match = re.search(r'лист[_\s]*(\d+)', class_name.lower())
                                if list_number_match:
                                    number = int(list_number_match.group(1))
                                    logging.info("Найден номер листа: {}".format(number))
                                else:
                                    # Добавляем в список проблемных классов
                                    problem_classes.append((class_name, class_path))
                                    return
                            else:
                                # Если нет явного указания на номер листа, это проблемный класс
                                problem_classes.append((class_name, class_path))
                                return
                        else:
                            # Если только одна цифра
                            number = int(digits_groups[0])
                        
                        # Проверяем на дубликаты
                        if number in numbers_found:
                            logging.warning("Найден дубликат номера {}: {}".format(number, class_name))
                        
                        numbers_found.add(number)
                        found_classes[number] = class_path
                    else:
                        # Если нет цифр, но есть слово "надпис"
                        logging.warning("Класс содержит 'надпис', но не содержит цифр: {}".format(class_name))
                        problem_classes.append((class_name, class_path))
            
            # Проверка является ли имя подходящим (содержит "надпис" и цифру)
            def is_matching_name(name):
                return (pattern1.match(name) or pattern2.match(name) or pattern3.match(name) or
                        pattern4.match(name) or pattern5.match(name))
            
            # Специальная обработка для MDB-файлов
            if self.is_mdb:
                logging.info("Используем специальную обработку для MDB файлов")
                
                # Пытаемся использовать arcpy.da.Walk для обхода всех данных в MDB
                try:
                    logging.info("Метод 1: Используем arcpy.da.Walk для обхода MDB")
                    for dirpath, dirnames, filenames in arcpy.da.Walk(
                            self.db_path,
                            datatype=["FeatureClass", "Table"]):
                        
                        workspace_path = dirpath
                        logging.info("Обход каталога: {}".format(workspace_path))
                        
                        for fc in filenames:
                            try:
                                # Полный путь к объекту
                                fc_path = os.path.join(workspace_path, fc)
                                logging.info("Найден объект: {}".format(fc))
                                
                                # Проверяем подходит ли имя
                                if "надпис" in fc.lower() or "лист" in fc.lower():
                                    check_and_add_class(fc_path, fc)
                                else:
                                    # Добавляем в общий список всех классов
                                    all_feature_classes.append((fc, fc_path))
                            except Exception as fc_err:
                                logging.warning("Ошибка при обработке объекта {}: {}".format(fc, str(fc_err)))
                
                except Exception as walk_err:
                    logging.warning("Ошибка при использовании arcpy.da.Walk: {}".format(str(walk_err)))
                
                # Если первый метод не сработал, пробуем другой подход
                if not found_classes:
                    try:
                        logging.info("Метод 2: Прямой перебор имен классов")
                        # Сначала ищем подкаталог ОАО_Пружанское
                        direct_ws = os.path.join(self.db_path, "ОАО_Пружанское")
                        arcpy.env.workspace = direct_ws
                        logging.info("Установлено рабочее пространство: {}".format(direct_ws))
                        
                        # Получаем список всех классов объектов
                        try:
                            all_fc = arcpy.ListFeatureClasses()
                            if all_fc:
                                logging.info("Все найденные классы объектов: {}".format(", ".join(all_fc)))
                                
                                # Проверяем каждый класс на соответствие шаблону
                                for fc in all_fc:
                                    if "надпис" in fc.lower() or "лист" in fc.lower():
                                        fc_path = os.path.join(direct_ws, fc)
                                        check_and_add_class(fc_path, fc)
                            else:
                                logging.info("Не найдено классов объектов в {}".format(direct_ws))
                        except Exception as fc_err:
                            logging.warning("Ошибка при получении списка классов объектов: {}".format(str(fc_err)))
                        
                        # Если не нашли, то перебираем возможные имена напрямую
                        if not found_classes:
                            logging.info("Метод 3: Перебор возможных имен классов")
                            
                            # Перебираем возможные имена, основываясь на скриншоте
                            possible_names = [
                                "Land_Пружанское_лист_1_надписи", 
                                "Land_Пружанское_лист_2_надписи", 
                                "Land_Пружанское_лист_3_надписи",
                                "Land_Пружанское_лист_3надписи"  # Заметим, что тут нет подчеркивания
                            ]
                            
                            # Расширяем список возможных имен
                            for i in range(1, 10):
                                for pattern in ["Land_Пружанское_лист_{}_надписи", "Land_Пружанское_лист_{}надписи"]:
                                    name = pattern.format(i)
                                    if name not in possible_names:
                                        possible_names.append(name)
                            
                            # Проверяем каждое возможное имя
                            for name in possible_names:
                                try:
                                    test_path = os.path.join(direct_ws, name)
                                    if arcpy.Exists(test_path):
                                        logging.info("Найден класс с прямым именем: {}".format(name))
                                        digits = re.findall(r'\d+', name)
                                        if digits:
                                            number = int(digits[0])
                                            if number not in numbers_found:
                                                numbers_found.add(number)
                                                found_classes[number] = test_path
                                except Exception as name_err:
                                    continue
                    
                    except Exception as direct_err:
                        logging.warning("Ошибка при прямом поиске: {}".format(str(direct_err)))
                
                # Если и второй метод не сработал, последняя попытка - использовать каталог
                if not found_classes:
                    logging.info("Метод 4: Пытаемся получить доступ через каталог MDB")
                    try:
                        # Перечисляем пути, которые мы хотим проверить
                        possible_paths = [
                            self.db_path,
                            os.path.join(self.db_path, "ОАО_Пружанское")
                        ]
                        
                        # Создаем список типовых имен с "надписи" из снимка экрана
                        typical_names = []
                        for i in range(1, 10):
                            typical_names.extend([
                                "Land_Пружанское_лист_{}_надписи".format(i),
                                "Land_Пружанское_лист_{}надписи".format(i)
                            ])
                        
                        # Проверяем существование по типовым именам в каждом из возможных путей
                        for path in possible_paths:
                            for name in typical_names:
                                try:
                                    full_path = os.path.join(path, name)
                                    if arcpy.Exists(full_path):
                                        logging.info("Найден класс в каталоге: {}".format(full_path))
                                        digits = re.findall(r'\d+', name)
                                        if digits:
                                            number = int(digits[0])
                                            if number not in numbers_found:
                                                numbers_found.add(number)
                                                found_classes[number] = full_path
                                except:
                                    continue
                    
                    except Exception as catalog_err:
                        logging.warning("Ошибка при поиске через каталог: {}".format(str(catalog_err)))
                
                # Восстанавливаем рабочее пространство
                arcpy.env.workspace = self.db_path
                
            else:
                # Стандартная обработка для GDB
                logging.info("Используем стандартный метод для GDB")
                arcpy.env.workspace = self.db_path
                
                # Получаем список наборов данных
                datasets = []
                try:
                    datasets = arcpy.ListDatasets()
                except:
                    datasets = []
                    logging.info("Не удалось получить список наборов данных, возможно их нет")
                
                # Проверяем классы объектов в корне базы данных
                try:
                    for fc in arcpy.ListFeatureClasses():
                        fc_path = os.path.join(self.db_path, fc)
                        check_and_add_class(fc_path, fc)
                except:
                    logging.warning("Ошибка при поиске классов в корне базы данных")
                
                # Проверяем классы объектов в наборах данных
                for ds in datasets:
                    try:
                        ds_path = os.path.join(self.db_path, ds)
                        arcpy.env.workspace = ds_path
                        
                        for fc in arcpy.ListFeatureClasses():
                            fc_path = os.path.join(ds_path, fc)
                            check_and_add_class(fc_path, fc)
                    except:
                        logging.warning("Ошибка при поиске классов в наборе данных {}".format(ds))
                
                # Возвращаем рабочее пространство в исходное состояние
                arcpy.env.workspace = self.db_path
                
            # Проверяем результаты поиска
            if not found_classes and not problem_classes:
                # Если не найдено классов и нет проблемных, предлагаем ручной выбор
                manual_selection = self.show_manual_selection_dialog(all_feature_classes)
                
                if manual_selection:
                    self.labels_classes = [path for _, path in manual_selection]
                    numbers = [i+1 for i in range(len(self.labels_classes))]
                    
                    logging.info("Пользователь вручную выбрал {} классов: {}".format(
                        len(self.labels_classes), 
                        ", ".join(name for name, _ in manual_selection)
                    ))
                    
                    return True, "Выбрано {} классов надписей вручную".format(len(self.labels_classes))
                else:
                    logging.warning("Пользователь отменил ручной выбор классов")
                    return False, "Не найдено классов, содержащих 'надпис' и цифру"
            
            # Если есть проблемные классы, показываем диалог с предупреждением
            if problem_classes:
                logging.warning("Найдено {} классов с проблемными именами".format(len(problem_classes)))
                warning_message = "Обнаружены классы с 'надпис', но с проблемами в определении номера:\n\n"
                warning_message += "\n".join([name for name, _ in problem_classes[:5]])
                if len(problem_classes) > 5:
                    warning_message += "\n... и еще {} классов".format(len(problem_classes) - 5)
                
                # Показываем диалог ручного выбора с предупреждением
                manual_selection = self.show_manual_selection_dialog(
                    all_feature_classes, 
                    warning_message=warning_message,
                    preselected=problem_classes + [(name, path) for number, path in found_classes.items() 
                                                  for fc in all_feature_classes if fc[1] == path for name, _ in [fc]]
                )
                
                if manual_selection:
                    self.labels_classes = [path for _, path in manual_selection]
                    logging.info("Пользователь вручную выбрал {} классов: {}".format(
                        len(self.labels_classes), 
                        ", ".join(name for name, _ in manual_selection)
                    ))
                    
                    return True, "Выбрано {} классов надписей вручную".format(len(self.labels_classes))
                else:
                    # Если пользователь отменил ручной выбор, продолжаем с тем что нашли автоматически
                    logging.info("Пользователь отменил ручной выбор, используем автоматически найденные классы")
            
            # Проверяем найденные автоматически классы
            if not found_classes:
                logging.warning("Не найдено классов, содержащих 'надпис' и цифру")
                return False, "Не найдено классов, содержащих 'надпис' и цифру"
            
            # Проверяем дубликаты номеров
            duplicate_numbers = []
            for num in numbers_found:
                count = 0
                for fc_name in found_classes.values():
                    if re.search(r'{}(?!\d)'.format(num), os.path.basename(fc_name)):
                        count += 1
                if count > 1:
                    duplicate_numbers.append(num)
            
            if duplicate_numbers:
                duplicate_message = "Найдены дубликаты номеров: {}".format(", ".join(map(str, duplicate_numbers)))
                logging.warning(duplicate_message)
                return False, duplicate_message
            
            # Сохраняем найденные классы
            self.labels_classes = [found_classes[num] for num in sorted(found_classes.keys())]
            
            # Проверяем наличие цифр в именах
            if not any(re.search(r'\d+', os.path.basename(fc)) for fc in self.labels_classes):
                no_numbers_message = "В именах классов надписей не найдены цифры"
                logging.warning(no_numbers_message)
                return False, no_numbers_message
            
            logging.info("Найдено {} классов с надписями: {}".format(
                len(self.labels_classes), 
                ", ".join(os.path.basename(fc) for fc in self.labels_classes)
            ))
            
            return True, "Найдено {} классов с надписями".format(len(self.labels_classes))
            
        except Exception as e:
            error_message = "Ошибка при поиске классов надписей: {}".format(str(e))
            logging.error(error_message)
            log_exception(e, "Ошибка при поиске классов надписей")
            return False, error_message
    
    def show_manual_selection_dialog(self, all_classes, warning_message=None, preselected=None):
        """Показывает диалог для ручного выбора классов надписей"""
        dialog = tk.Toplevel()
        dialog.title("Выбор классов надписей")
        dialog.geometry("600x500")
        dialog.transient(tk.Tk().winfo_toplevel())
        dialog.grab_set()
        
        # Результат выбора
        selected_classes = []
        result = {"action": None}
        
        # Фрейм для содержимого
        main_frame = tk.Frame(dialog, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        title_label = tk.Label(main_frame, 
                            text="Выберите классы надписей для обработки",
                            font=("Arial", 12, "bold"))
        title_label.pack(pady=10)
        
        # Если есть предупреждение
        if warning_message:
            warning_label = tk.Label(main_frame, 
                                  text=warning_message,
                                  font=("Arial", 10),
                                  fg="red",
                                  wraplength=550,
                                  justify=tk.LEFT)
            warning_label.pack(pady=10, fill=tk.X)
        
        # Инструкция
        hint_label = tk.Label(main_frame, 
                           text="Выберите классы содержащие надписи из списка ниже.\n"
                                "Используйте Ctrl+Click для выбора нескольких элементов.",
                           font=("Arial", 10),
                           fg="blue",
                           wraplength=550)
        hint_label.pack(pady=10)
        
        # Фрейм для списка с прокруткой
        list_frame = tk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Полоса прокрутки
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Список классов
        classes_listbox = tk.Listbox(list_frame, 
                                  selectmode=tk.MULTIPLE, 
                                  yscrollcommand=scrollbar.set,
                                  font=("Courier New", 10),
                                  height=15)
        scrollbar.config(command=classes_listbox.yview)
        classes_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Заполняем список классов
        for i, (name, path) in enumerate(all_classes):
            if "надпис" in name.lower():
                display_name = "* {}".format(name)
            else:
                display_name = "  {}".format(name)
            classes_listbox.insert(tk.END, display_name)
            
            # Если класс в списке предварительно выбранных
            if preselected and (name, path) in preselected:
                classes_listbox.selection_set(i)
        
        # Функции для кнопок
        def on_continue():
            selected_indices = classes_listbox.curselection()
            for i in selected_indices:
                name, path = all_classes[i]
                selected_classes.append((name, path))
            result["action"] = "continue"
            dialog.destroy()
        
        def on_cancel():
            result["action"] = "cancel"
            dialog.destroy()
        
        # Кнопки
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=15)
        
        continue_button = tk.Button(button_frame, 
                                  text="Продолжить", 
                                  command=on_continue,
                                  bg="#4CAF50",
                                  fg="white",
                                  font=("Arial", 10, "bold"),
                                  width=12)
        continue_button.pack(side=tk.LEFT, padx=10)
        
        cancel_button = tk.Button(button_frame, 
                                text="Отменить", 
                                command=on_cancel,
                                bg="#f44336",
                                fg="white",
                                font=("Arial", 10, "bold"),
                                width=12)
        cancel_button.pack(side=tk.LEFT, padx=10)
        
        # Ждем действия пользователя
        dialog.wait_window()
        
        # Возвращаем список выбранных классов или None, если отменено
        return selected_classes if result["action"] == "continue" else None
    
    def show_verification_dialog(self, message, has_duplicates=False, has_no_numbers=False):
        """Показывает диалог проверки с кнопками Повторить/Продолжить и Закончить"""
        dialog = tk.Toplevel()
        dialog.title("Проверка классов надписей")
        dialog.geometry("500x300")
        dialog.transient(tk.Tk().winfo_toplevel())  # Делаем окно модальным
        dialog.grab_set()  # Захватываем фокус
        
        # Переменная для хранения результата
        result = {"action": None}
        
        # Фрейм для содержимого
        main_frame = tk.Frame(dialog, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        title_label = tk.Label(main_frame, 
                             text="Результаты проверки классов надписей",
                             font=("Arial", 12, "bold"))
        title_label.pack(pady=10)
        
        # Сообщение
        message_label = tk.Label(main_frame, 
                              text=message,
                              font=("Arial", 10),
                              wraplength=460,
                              justify=tk.LEFT)
        message_label.pack(pady=10, fill=tk.X)
        
        # Подсказка пользователю
        if has_duplicates:
            hint_text = ("Обнаружены дубликаты номеров в классах надписей.\n"
                       "Рекомендуется изменить имена классов вручную, чтобы номера не повторялись.\n"
                       "После исправления нажмите 'Повторить'.")
        elif has_no_numbers:
            hint_text = ("В именах классов надписей не найдены цифры.\n"
                       "Пожалуйста, проверьте имена классов и добавьте цифры если необходимо.\n"
                       "Нажмите 'Продолжить', если считаете что всё в порядке, или 'Закончить' для отмены.")
        else:
            hint_text = "Пожалуйста, проверьте результаты поиска классов надписей."
        
        hint_label = tk.Label(main_frame, 
                           text=hint_text,
                           font=("Arial", 9),
                           fg="blue",
                           wraplength=460,
                           justify=tk.LEFT)
        hint_label.pack(pady=10, fill=tk.X)
        
        # Список найденных классов надписей
        if self.labels_classes:
            classes_frame = tk.Frame(main_frame)
            classes_frame.pack(fill=tk.BOTH, expand=True, pady=10)
            
            classes_label = tk.Label(classes_frame, text="Найденные классы надписей:", anchor="w")
            classes_label.pack(fill=tk.X)
            
            # Создаем скроллируемый список
            classes_listbox = tk.Listbox(classes_frame, height=5)
            scrollbar = tk.Scrollbar(classes_frame, orient="vertical", command=classes_listbox.yview)
            classes_listbox.configure(yscrollcommand=scrollbar.set)
            
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            classes_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            # Заполняем список найденными классами
            for fc_path in self.labels_classes:
                classes_listbox.insert(tk.END, os.path.basename(fc_path))
        
        # Кнопки
        button_frame = tk.Frame(main_frame)
        button_frame.pack(pady=15)
        
        # Выбираем текст для первой кнопки
        first_button_text = "Повторить" if has_duplicates else "Продолжить"
        
        # Функция для первой кнопки (Повторить/Продолжить)
        def on_retry_continue():
            result["action"] = "retry" if has_duplicates else "continue"
            dialog.destroy()
        
        # Функция для второй кнопки (Закончить)
        def on_finish():
            result["action"] = "finish"
            dialog.destroy()
        
        retry_continue_button = tk.Button(button_frame, 
                                       text=first_button_text, 
                                       command=on_retry_continue,
                                       bg="#4CAF50",
                                       fg="white",
                                       font=("Arial", 10, "bold"),
                                       width=10)
        retry_continue_button.pack(side=tk.LEFT, padx=10)
        
        finish_button = tk.Button(button_frame, 
                               text="Закончить", 
                               command=on_finish,
                               bg="#f44336",
                               fg="white",
                               font=("Arial", 10, "bold"),
                               width=10)
        finish_button.pack(side=tk.LEFT, padx=10)
        
        # Ждем, пока пользователь не закроет диалог
        dialog.wait_window()
        
        return result["action"]
    
    def process_identity(self, target_fc_path):
        """Выполняет операцию идентичности между целевым слоем и классами надписей"""
        if not ARCPY_AVAILABLE:
            logging.error("Модуль arcpy недоступен")
            messagebox.showerror("Ошибка", "Модуль arcpy не доступен для операций идентичности")
            return False
        
        if not self.labels_classes:
            logging.error("Не найдены классы надписей для операции идентичности")
            messagebox.showerror("Ошибка", "Не найдены классы надписей для операции идентичности")
            return False
        
        try:
            logging.info("Начало операции идентичности")
            logging.info("Входной класс: {}".format(target_fc_path))
            
            # Формируем имя выходного класса
            dataset_path = os.path.dirname(target_fc_path)
            gdb_path = os.path.dirname(dataset_path)
            
            # Сохраняем путь к GDB для дальнейшего использования
            self.gdb_path = gdb_path
            
            # Определяем входное имя класса (без расширения пути)
            input_name = os.path.basename(target_fc_path)
            
            # Формируем имя выходного класса
            output_name = input_name.replace("_контур", "").replace("контур", "") + "_сетка"
            output_path = os.path.join(dataset_path, output_name)
            
            logging.info("Выходной класс: {}".format(output_path))
            
            # Очищаем кэш рабочих пространств
            arcpy.ClearWorkspaceCache_management()
            
            # Проверяем, существует ли выходной класс
            if arcpy.Exists(output_path):
                logging.warning("Выходной класс {} уже существует".format(output_name))
                if messagebox.askyesno("Предупреждение", 
                                     "Класс объектов {} уже существует. Хотите заменить его?".format(output_name)):
                    try:
                        arcpy.Delete_management(output_path)
                        logging.info("Удалён существующий класс: {}".format(output_path))
                    except Exception as e:
                        logging.error("Ошибка при удалении существующего класса: {}".format(str(e)))
                        messagebox.showerror("Ошибка", 
                                          "Не удалось удалить существующий класс. Возможно, он используется.")
                        return False
                else:
                    logging.info("Операция идентичности отменена пользователем")
                    return False
            
            # Создаем копию целевого слоя как основу
            arcpy.CopyFeatures_management(target_fc_path, output_path)
            logging.info("Создана копия целевого слоя как основа для Identity: {}".format(output_path))

            # Проверяем пространственные привязки
            target_sr = arcpy.Describe(target_fc_path).spatialReference
            logging.info("Пространственная привязка целевого класса: {}".format(target_sr.name))
            
            # Обрабатываем каждый класс прошлого тура по очереди
            for i, label_class in enumerate(self.labels_classes):
                try:
                    logging.info("Обработка класса надписей {}/{}: {}".format(
                        i+1, len(self.labels_classes), os.path.basename(label_class)))
                    
                    # Создаем временный слой для текущего класса надписей
                    temp_label_layer = "temp_label_layer_{}".format(i)
                    # Если есть поле NPP, применяем фильтр NPP > 0
                    field_names = [f.name for f in arcpy.ListFields(label_class)]
                    where_clause = "NPP > 0" if "NPP" in field_names else None
                    
                    arcpy.MakeFeatureLayer_management(label_class, temp_label_layer, where_clause)
                    
                    # Создаем временный слой для текущего результата
                    temp_output_layer = "temp_output_layer_{}".format(i)
                    arcpy.MakeFeatureLayer_management(output_path, temp_output_layer)
                    
                    # Создаем временный результат для текущей операции Identity
                    temp_result = "in_memory\\temp_identity_{}".format(i)
                    
                    # Выполняем Identity для текущего класса надписей
                    arcpy.Identity_analysis(
                        in_features=temp_output_layer,
                        identity_features=temp_label_layer,
                        out_feature_class=temp_result,
                        join_attributes="ALL",
                        cluster_tolerance="0.001 Meters"
                    )
                    
                    # Если операция успешна, заменяем текущий результат
                    if arcpy.Exists(temp_result):
                        # Удаляем предыдущий результат
                        arcpy.Delete_management(output_path)
                        # Копируем новый результат
                        arcpy.CopyFeatures_management(temp_result, output_path)
                        logging.info("Успешно выполнена операция Identity с классом {}".format(
                            os.path.basename(label_class)))
                    
                    # Очищаем временные данные
                    for temp_layer in [temp_label_layer, temp_output_layer, temp_result]:
                        if arcpy.Exists(temp_layer):
                            arcpy.Delete_management(temp_layer)
                
                except Exception as e:
                    logging.error("Ошибка при обработке класса {}: {}".format(
                        os.path.basename(label_class), str(e)))
                    logging.error(traceback.format_exc())
            
            # Проверяем результат операции
            if arcpy.Exists(output_path):
                result_count = int(arcpy.GetCount_management(output_path).getOutput(0))
                logging.info("Итоговый класс сетки содержит {} объектов".format(result_count))
                
                if result_count > 0:
                    # Применяем инструмент MultipartToSinglepart (Раздробить составной объект)
                    try:
                        logging.info("Применение инструмента 'Раздробить составной объект' к слою {}".format(output_path))
                        
                        # Создаем временный класс объектов для результата
                        temp_singlepart_path = "in_memory\\temp_singlepart"
                        
                        # Применяем инструмент MultipartToSinglepart
                        arcpy.MultipartToSinglepart_management(
                            output_path,
                            temp_singlepart_path
                        )
                        
                        # Получаем количество объектов после раздробления
                        singlepart_count = int(arcpy.GetCount_management(temp_singlepart_path).getOutput(0))
                        logging.info("После раздробления слой содержит {} объектов".format(singlepart_count))
                        
                        # Удаляем исходный слой и заменяем его на слой с раздробленными объектами
                        arcpy.Delete_management(output_path)
                        arcpy.CopyFeatures_management(temp_singlepart_path, output_path)
                        
                        # Очищаем временные данные
                        arcpy.Delete_management(temp_singlepart_path)
                        
                        logging.info("Инструмент 'Раздробить составной объект' успешно применен")
                    except Exception as multipart_err:
                        logging.error("Ошибка при применении инструмента 'Раздробить составной объект': {}".format(str(multipart_err)))
                        logging.error(traceback.format_exc())
                        messagebox.showwarning("Предупреждение", 
                                            "Ошибка при применении инструмента 'Раздробить составной объект':\n{}".format(str(multipart_err)))
                    
                    # Обрабатываем поля в слое Land_"Сокр"_сетка
                    self.process_fields(output_path)
                    
                    # Сравниваем с границами Lots_"сокр" и удаляем объекты за пределами контура
                    self.filter_by_lots_boundary(output_path)

                    # Копируем данные из Land_"Сокр"_контур в Land_"Сокр"_сетка
                    try:
                        logging.info("Копирование данных из Land_\"Сокр\"_контур в Land_\"Сокр\"_сетка...")
                        
                        # Ищем путь к Land_"Сокр"_контур
                        land_contour_name = os.path.basename(target_fc_path)
                        
                        # Убеждаемся, что это Land_"Сокр"
                        if not land_contour_name.endswith("_контур"):
                            # Если входной слой был Land_"Сокр", то формируем имя для Land_"Сокр"_контур
                            land_contour_name = land_contour_name + "_контур"
                        
                        land_contour_path = os.path.join(dataset_path, land_contour_name)
                        
                        if arcpy.Exists(land_contour_path):
                            logging.info("Найден класс Land_\"Сокр\"_контур: {}".format(land_contour_path))
                            
                            # Создаем временный слой для Land_"Сокр"_контур
                            temp_contour_layer = "temp_contour_layer"
                            arcpy.MakeFeatureLayer_management(land_contour_path, temp_contour_layer)
                            
                            # Подсчитываем количество объектов в Land_"Сокр"_контур
                            contour_count = int(arcpy.GetCount_management(temp_contour_layer).getOutput(0))
                            logging.info("Количество объектов в Land_\"Сокр\"_контур: {}".format(contour_count))
                            
                            if contour_count > 0:
                                # Добавляем данные из Land_"Сокр"_контур в Land_"Сокр"_сетка
                                arcpy.Append_management(
                                    inputs=temp_contour_layer,
                                    target=output_path,
                                    schema_type="NO_TEST"  # Не проверяем схему данных
                                )
                                logging.info("Данные из Land_\"Сокр\"_контур скопированы в Land_\"Сокр\"_сетка")
                                
                                # Проверяем итоговое количество объектов
                                final_output_count = int(arcpy.GetCount_management(output_path).getOutput(0))
                                logging.info("Итоговое количество объектов в Land_\"Сокр\"_сетка после копирования: {}".format(final_output_count))
                            else:
                                logging.info("Land_\"Сокр\"_контур не содержит объектов для копирования")
                            
                            # Удаляем временный слой
                            arcpy.Delete_management(temp_contour_layer)
                        else:
                            logging.warning("Не найден класс Land_\"Сокр\"_контур: {}".format(land_contour_path))
                            messagebox.showwarning("Предупреждение", 
                                                 "Не найден класс Land_\"Сокр\"_контур для копирования данных")
                    
                    except Exception as copy_err:
                        logging.error("Ошибка при копировании данных из Land_\"Сокр\"_контур: {}".format(str(copy_err)))
                        logging.error(traceback.format_exc())
                        messagebox.showwarning("Предупреждение", 
                                             "Ошибка при копировании данных из Land_\"Сокр\"_контур:\n{}".format(str(copy_err)))
                    
                    # Добавляем класс в таблицу содержания
                    try:
                        # Пытаемся определить продукт ArcGIS
                        product_info = arcpy.ProductInfo()
                        logging.info("Определен продукт ArcGIS: {}".format(product_info))
                        
                        # Проверяем, запущен ли скрипт из ArcMap (более широкий список возможных значений)
                        if product_info in ["ArcView", "ArcEditor", "ArcInfo", "Desktop"]:
                            logging.info("Добавление слоя в таблицу содержания ArcMap...")
                            try:
                                import arcpy.mapping as mapping
                                
                                # Получаем текущий документ карты
                                mxd = mapping.MapDocument("CURRENT")
                                logging.info("Получен документ карты")
                                
                                # Получаем активный фрейм данных
                                if mapping.ListDataFrames(mxd):
                                    df = mapping.ListDataFrames(mxd)[0]  # Берем первый фрейм данных
                                    logging.info("Получен фрейм данных: {}".format(df.name))
                                    
                                    # Создаем слой из класса объектов
                                    try:
                                        new_layer = mapping.Layer(output_path)
                                        logging.info("Создан слой из класса объектов")
                                        
                                        # Добавляем слой в документ карты
                                        mapping.AddLayer(df, new_layer, "TOP")
                                        logging.info("Слой добавлен в фрейм данных")
                                        
                                        # Сохраняем документ карты
                                        mxd.save()
                                        logging.info("Документ карты сохранен")
                                        
                                        # Обновляем вид
                                        arcpy.RefreshTOC()
                                        arcpy.RefreshActiveView()
                                        logging.info("Вид обновлен")
                                        
                                        # Показываем сообщение пользователю
                                        messagebox.showinfo("Успех", "Слой Land_\"Сокр\"_сетка успешно добавлен в таблицу содержания")
                                    except Exception as layer_err:
                                        logging.error("Ошибка при создании слоя: {}".format(str(layer_err)))
                                        
                                        # Пробуем альтернативный способ добавления
                                        try:
                                            # Добавляем через MakeFeatureLayer и добавление временного слоя
                                            temp_layer_name = "temp_siatka_layer"
                                            arcpy.MakeFeatureLayer_management(output_path, temp_layer_name)
                                            logging.info("Создан временный слой: {}".format(temp_layer_name))
                                            
                                            # Добавляем временный слой в таблицу содержания
                                            result = arcpy.mapping.Layer(temp_layer_name)
                                            arcpy.mapping.AddLayer(df, result, "TOP")
                                            logging.info("Временный слой добавлен в таблицу содержания")
                                            
                                            # Сохраняем документ карты
                                            mxd.save()
                                            arcpy.RefreshTOC()
                                            arcpy.RefreshActiveView()
                                            
                                            # Показываем сообщение пользователю
                                            messagebox.showinfo("Успех", "Слой Land_\"Сокр\"_сетка успешно добавлен в таблицу содержания")
                                        except Exception as temp_err:
                                            logging.error("Ошибка при альтернативном методе добавления слоя: {}".format(str(temp_err)))
                                            messagebox.showwarning("Предупреждение", 
                                                                "Не удалось автоматически добавить слой в таблицу содержания.\n"
                                                                "Пожалуйста, добавьте его вручную из: {}".format(output_path))
                                else:
                                    logging.warning("Не найдены фреймы данных в документе карты")
                                    messagebox.showwarning("Предупреждение", 
                                                         "Не найдены фреймы данных в документе карты.\n"
                                                         "Пожалуйста, добавьте слой вручную из: {}".format(output_path))
                            except ImportError as ie:
                                logging.error("Ошибка импорта модуля arcpy.mapping: {}".format(str(ie)))
                                messagebox.showwarning("Предупреждение", 
                                                     "Не удалось импортировать модуль для работы с картой.\n"
                                                     "Пожалуйста, добавьте слой вручную из: {}".format(output_path))
                        
                        elif product_info in ["ArcGISPro", "Pro"]:  # Если это ArcGIS Pro
                            logging.info("Добавление слоя в таблицу содержания ArcGIS Pro...")
                            try:
                                import arcpy.mp as mp
                                
                                # Получаем текущий проект и активную карту
                                aprx = mp.ArcGISProject("CURRENT")
                                
                                if aprx.listMaps():
                                    m = aprx.activeMap
                                    logging.info("Получена активная карта: {}".format(m.name))
                                    
                                    # Добавляем слой
                                    m.addDataFromPath(output_path)
                                    logging.info("Слой добавлен в карту")
                                    
                                    # Сохраняем проект
                                    aprx.save()
                                    logging.info("Проект сохранен")
                                    
                                    # Показываем сообщение пользователю
                                    messagebox.showinfo("Успех", "Слой Land_\"Сокр\"_сетка успешно добавлен в таблицу содержания")
                                else:
                                    logging.warning("Не найдены карты в проекте")
                                    messagebox.showwarning("Предупреждение", 
                                                         "Не найдены карты в проекте.\n"
                                                         "Пожалуйста, добавьте слой вручную из: {}".format(output_path))
                            except ImportError as ie:
                                logging.error("Ошибка импорта модуля arcpy.mp: {}".format(str(ie)))
                                messagebox.showwarning("Предупреждение", 
                                                     "Не удалось импортировать модуль для работы с проектом.\n"
                                                     "Пожалуйста, добавьте слой вручную из: {}".format(output_path))
                        else:
                            # Если не удалось определить продукт, пробуем универсальный подход для ArcMap, 
                            # так как по логам видно, что это ArcMap 10.4
                            logging.warning("Не определен продукт ArcGIS ({}), пробуем универсальный подход".format(product_info))
                            try:
                                # Пробуем использовать arcpy.mapping напрямую
                                import arcpy.mapping as mapping
                                
                                try:
                                    # Создаем временный слой
                                    temp_layer_name = "temp_siatka_layer_universal"
                                    arcpy.MakeFeatureLayer_management(output_path, temp_layer_name)
                                    logging.info("Создан временный слой: {}".format(temp_layer_name))
                                    
                                    # Получаем текущий документ
                                    try:
                                        mxd = mapping.MapDocument("CURRENT")
                                        df = mapping.ListDataFrames(mxd)[0]
                                        
                                        # Создаем слой из временного слоя
                                        temp_layer = mapping.Layer(temp_layer_name)
                                        mapping.AddLayer(df, temp_layer, "TOP")
                                        
                                        # Сохраняем документ и обновляем вид
                                        mxd.save()
                                        arcpy.RefreshTOC()
                                        arcpy.RefreshActiveView()
                                        
                                        logging.info("Успешно добавлен слой в таблицу содержания универсальным методом")
                                        messagebox.showinfo("Успех", "Слой Land_\"Сокр\"_сетка успешно добавлен в таблицу содержания")
                                    except Exception as mxd_err:
                                        logging.error("Ошибка при работе с документом карты: {}".format(str(mxd_err)))
                                        # Сообщаем пользователю о необходимости ручного добавления
                                        messagebox.showwarning("Предупреждение", 
                                                            "Не удалось автоматически добавить слой в таблицу содержания.\n"
                                                            "Пожалуйста, добавьте его вручную из:\n{}".format(output_path))
                                        
                                except Exception as universal_err:
                                    logging.error("Ошибка при универсальном методе добавления слоя: {}".format(str(universal_err)))
                                    messagebox.showwarning("Предупреждение", 
                                                        "Не удалось автоматически добавить слой в таблицу содержания.\n"
                                                        "Пожалуйста, добавьте его вручную из:\n{}".format(output_path))
                            except ImportError:
                                logging.error("Не удалось импортировать модуль arcpy.mapping для универсального метода")
                                messagebox.showwarning("Предупреждение", 
                                                    "Не удалось автоматически добавить слой в таблицу содержания.\n"
                                                    "Пожалуйста, добавьте его вручную из:\n{}".format(output_path))
                    
                    except Exception as add_err:
                        logging.error("Общая ошибка при добавлении слоя в таблицу содержания: {}".format(str(add_err)))
                        logging.error(traceback.format_exc())
                        messagebox.showwarning("Предупреждение", 
                                             "Произошла ошибка при добавлении слоя в таблицу содержания.\n"
                                             "Пожалуйста, добавьте его вручную из: {}".format(output_path))
                    
                    # Применяем дополнительную обработку участков в самом конце
                    try:
                        logging.info("Начало применения дополнительной обработки участков...")
                        
                        # Получаем путь к слою Admi_"Сокр"
                        admi_clip_name = "Admi_{}".format(self.shortened_name)
                        dataset_path = os.path.dirname(output_path)
                        admi_clip_path = os.path.join(dataset_path, admi_clip_name)
                        
                        # Проверяем существование Admi_"Сокр"
                        if not arcpy.Exists(admi_clip_path):
                            logging.warning("Слой '{}' не найден, обработка будет выполнена без учета административных границ".format(admi_clip_name))
                            admi_clip_path = None
                        
                        # Создаем и запускаем обработчик участков
                        land_processor = LandProcessor(self.gdb_path, output_path, admi_clip_path)
                        process_success = land_processor.process_land_parcels()
                        
                        if process_success:
                            logging.info("Дополнительная обработка участков успешно завершена")
                            messagebox.showinfo("Успех", "Дополнительная обработка участков успешно завершена")
                        else:
                            logging.warning("Дополнительная обработка участков не была выполнена или выполнена с ошибками")
                            messagebox.showwarning("Предупреждение", 
                                                 "Дополнительная обработка участков не была выполнена или выполнена с ошибками.\nПодробности в логе.")
                    except Exception as processor_err:
                        logging.error("Ошибка при выполнении дополнительной обработки участков: {}".format(str(processor_err)))
                        logging.error(traceback.format_exc())
                        messagebox.showwarning("Предупреждение", 
                                             "Ошибка при выполнении дополнительной обработки участков:\n{}".format(str(processor_err)))
                    
                    return True
                else:
                    logging.error("Операция идентичности не создала объектов в выходном классе")
                    messagebox.showerror("Ошибка", "Операция идентичности не создала объектов")
                    return False
            else:
                logging.error("Не удалось создать выходной класс {}".format(output_name))
                messagebox.showerror("Ошибка", "Не удалось создать выходной класс {}".format(output_name))
                return False
            
        except Exception as e:
            error_message = "Ошибка при выполнении операции идентичности: {}".format(str(e))
            logging.error(error_message)
            log_exception(e, "Ошибка при выполнении операции идентичности")
            messagebox.showerror("Ошибка", error_message)
            return False

    def save_named_mxd(self, mxd, output_path, method_name=""):
        """Сохраняет именованную копию MXD файла на основе пути к выходному классу"""
        try:
            # Формируем путь к MXD на основе имени набора данных
            dataset_path = os.path.dirname(output_path)
            gdb_path = os.path.dirname(dataset_path)
            dataset_name = os.path.basename(dataset_path)
            mxd_path = os.path.join(os.path.dirname(gdb_path), "{}.mxd".format(dataset_name))
            
            # Сохраняем копию карты
            mxd.saveACopy(mxd_path)
            logging.info("Сохранена копия карты с добавленным слоем сетки{}: {}".format(
                " ({})".format(method_name) if method_name else "", mxd_path))
            return True
        except Exception as save_err:
            logging.error("Ошибка при сохранении копии карты{}: {}".format(
                " ({})".format(method_name) if method_name else "", str(save_err)))
            return False

    def process_fields(self, output_path):
        """Обрабатывает поля в выходном слое Land_"Сокр"_сетка:
        1. Находит три поля NPP
        2. Очищает нулевые значения (заменяет на "")
        3. Создает новое обобщенное поле NPP типа Short Integer
        4. Копирует данные из трех полей в одно
        5. Удаляет все лишние поля, кроме LandType и LandCode
        """
        try:
            logging.info("Начало обработки полей в слое: {}".format(output_path))
            
            # Получаем список всех полей
            fields = arcpy.ListFields(output_path)
            field_names = [field.name for field in fields]
            logging.info("Найдено полей: {}".format(len(field_names)))
            
            # Находим все поля NPP (они могут иметь разные названия после операции Identity)
            npp_fields = [field.name for field in fields if "NPP" in field.name.upper()]
            logging.info("Найдены поля NPP: {}".format(", ".join(npp_fields)))
            
            if not npp_fields:
                logging.warning("Не найдено полей NPP для обработки")
                messagebox.showwarning("Предупреждение", "Не найдено полей NPP для обработки в слое")
                return False
            
            # Создаем новое поле NPP_Combined для объединения данных - типа Short Integer
            new_field_name = "NPP_Combined"
            try:
                # Проверяем, существует ли поле
                if new_field_name in field_names:
                    logging.info("Поле {} уже существует, удаляем его".format(new_field_name))
                    arcpy.DeleteField_management(output_path, new_field_name)
                
                # Создаем новое поле Short Integer
                arcpy.AddField_management(
                    output_path,
                    new_field_name,
                    "SHORT"  # Short Integer
                )
                logging.info("Создано новое поле: {} (тип: SHORT)".format(new_field_name))
            except Exception as add_field_err:
                logging.error("Ошибка при создании нового поля {}: {}".format(new_field_name, str(add_field_err)))
                messagebox.showerror("Ошибка", "Не удалось создать новое поле NPP_Combined: {}".format(str(add_field_err)))
                return False
            
            # Копируем данные из полей NPP в новое поле
            try:
                # Создаем курсор для обновления данных
                fields_to_read = npp_fields + [new_field_name]
                with arcpy.da.UpdateCursor(output_path, fields_to_read) as cursor:
                    row_count = 0
                    for row in cursor:
                        # Берем первое ненулевое значение
                        value = None
                        for i in range(len(npp_fields)):
                            if row[i] is not None and row[i] != 0:
                                value = row[i]
                                break
                        
                        # Записываем в новое поле
                        row[-1] = value
                        cursor.updateRow(row)
                        row_count += 1
                
                logging.info("Данные из полей NPP скопированы в поле {}, обработано строк: {}".format(new_field_name, row_count))
            except Exception as update_err:
                logging.error("Ошибка при копировании данных из полей NPP: {}".format(str(update_err)))
                messagebox.showerror("Ошибка", "Не удалось скопировать данные из полей NPP: {}".format(str(update_err)))
                return False
            
            # Определяем список полей для сохранения
            system_fields = ["OBJECTID", "Shape", "SHAPE", "FID", "OID", "SHAPE_Length", "SHAPE_Area"]
            # Добавляем LandType и LandCode в список сохраняемых полей
            additional_fields = ["LandType", "LandCode"]
            fields_to_keep = system_fields + [new_field_name] + additional_fields
            
            # Создаем список полей для удаления
            fields_to_delete = [field.name for field in fields 
                               if field.name not in fields_to_keep and not field.required]
            
            # Удаляем все лишние поля
            try:
                if fields_to_delete:
                    logging.info("Удаление {} лишних полей: {}".format(len(fields_to_delete), ", ".join(fields_to_delete[:10]) + ("..." if len(fields_to_delete) > 10 else "")))
                    arcpy.DeleteField_management(output_path, fields_to_delete)
                    logging.info("Удалены все лишние поля, оставлены только NPP_Combined, LandType, LandCode и системные поля")
                else:
                    logging.info("Нет полей для удаления")
            except Exception as del_field_err:
                logging.error("Ошибка при удалении лишних полей: {}".format(str(del_field_err)))
                messagebox.showwarning("Предупреждение", "Не удалось удалить все лишние поля: {}".format(str(del_field_err)))
            
            # Переименовываем поле NPP_Combined в NPP
            try:
                # Проверяем, нет ли уже поля NPP
                if "NPP" in field_names and "NPP" != new_field_name:
                    logging.info("Поле NPP уже существует, удаляем его")
                    arcpy.DeleteField_management(output_path, "NPP")
                
                # Переименовываем NPP_Combined в NPP
                arcpy.AlterField_management(
                    output_path,
                    new_field_name,
                    "NPP",
                    "NPP"
                )
                logging.info("Поле {} переименовано в NPP".format(new_field_name))
            except Exception as rename_err:
                logging.error("Ошибка при переименовании поля {} в NPP: {}".format(new_field_name, str(rename_err)))
                logging.info("Пробуем альтернативный метод переименования")
                
                try:
                    # Альтернативный способ: создать новое поле NPP и скопировать данные
                    arcpy.AddField_management(output_path, "NPP", "SHORT")  # Short Integer
                    arcpy.CalculateField_management(output_path, "NPP", "!{}!".format(new_field_name), "PYTHON_9.3")
                    arcpy.DeleteField_management(output_path, new_field_name)
                    logging.info("Поле успешно переименовано альтернативным методом")
                except Exception as alt_rename_err:
                    logging.error("Ошибка при альтернативном переименовании: {}".format(str(alt_rename_err)))
                    messagebox.showwarning("Предупреждение", "Не удалось переименовать поле в NPP. Оставлено название {}".format(new_field_name))
            
            # Очищаем значения в поле NPP на основе значений LandType
            try:
                # Определяем имя поля NPP (это может быть NPP или NPP_Combined, в зависимости от того, удалось ли переименование)
                npp_field_name = "NPP" if "NPP" in [f.name for f in arcpy.ListFields(output_path)] else new_field_name
                
                # Проверяем наличие поля LandType
                if "LandType" in [f.name for f in arcpy.ListFields(output_path)]:
                    logging.info("Начало очистки поля NPP на основе значений LandType")
                    
                    # Разрешенные значения LandType, при которых NPP не очищается
                    allowed_types = [101, 102, 103]
                    logging.info("Разрешенные значения LandType: {}".format(", ".join(map(str, allowed_types))))
                    
                    # Создаем курсор обновления с полями NPP и LandType
                    cleared_count = 0
                    with arcpy.da.UpdateCursor(output_path, [npp_field_name, "LandType"]) as cursor:
                        for row in cursor:
                            npp_value = row[0]
                            land_type = row[1]
                            
                            # Если LandType не входит в список разрешенных, очищаем NPP
                            if land_type is not None and land_type not in allowed_types:
                                row[0] = None  # Устанавливаем NPP в NULL
                                cursor.updateRow(row)
                                cleared_count += 1
                    
                    logging.info("Очищено {} значений в поле NPP, где LandType не входит в список разрешенных".format(cleared_count))
                else:
                    logging.warning("Поле LandType не найдено, очистка NPP не выполнена")
            except Exception as clear_err:
                logging.error("Ошибка при очистке поля NPP на основе LandType: {}".format(str(clear_err)))
                messagebox.showwarning("Предупреждение", "Не удалось выполнить очистку поля NPP на основе LandType: {}".format(str(clear_err)))
            
            logging.info("Обработка полей в слое {} завершена успешно".format(output_path))
            return True
            
        except Exception as e:
            error_message = "Общая ошибка при обработке полей: {}".format(str(e))
            logging.error(error_message)
            log_exception(e, "Ошибка при обработке полей в слое")
            messagebox.showerror("Ошибка", error_message)
            return False

    def filter_by_lots_boundary(self, grid_path):
        """Сравнивает объекты Land_"Сокр"_сетка с границами Lots_"Сокр" и удаляет объекты за пределами контура
        
        Args:
            grid_path (str): Путь к слою Land_"Сокр"_сетка
        """
        try:
            logging.info("Начало сравнения слоя Land_\"Сокр\"_сетка с границами Lots_\"Сокр\"")
            
            # Находим путь к слою Lots_"Сокр"
            dataset_path = os.path.dirname(grid_path)
            gdb_path = os.path.dirname(dataset_path)
            
            # Формируем ожидаемое имя класса Lots_"Сокр"
            lots_name = "Lots_{}".format(self.shortened_name)
            lots_path = os.path.join(dataset_path, lots_name)
            
            if not arcpy.Exists(lots_path):
                logging.warning("Класс объектов '{}' не найден".format(lots_name))
                messagebox.showwarning("Предупреждение", 
                                     "Не удалось найти класс Lots_\"{}\". Фильтрация за пределами контура не выполнена.".format(self.shortened_name))
                return False
            
            logging.info("Найден класс Lots_\"Сокр\": {}".format(lots_path))
            
            # Создаем временные слои для работы
            grid_layer = "temp_grid_layer"
            lots_layer = "temp_lots_layer"
            
            # Создаем временный слой для Lots_"Сокр"
            arcpy.MakeFeatureLayer_management(lots_path, lots_layer)
            logging.info("Создан временный слой для Lots_\"Сокр\"")
            
            # Создаем временный слой для Land_"Сокр"_сетка
            arcpy.MakeFeatureLayer_management(grid_path, grid_layer)
            logging.info("Создан временный слой для Land_\"Сокр\"_сетка")
            
            # Коды LandCode, которые нужно проверять за пределами контура
            target_codes = [123, 3, 6, 7]
            logging.info("Целевые коды LandCode для проверки за пределами контура: {}".format(", ".join(map(str, target_codes))))
            
            # Строим SQL-выражение для выборки целевых кодов
            land_code_clause = "LandCode IN ({})".format(",".join(map(str, target_codes)))
            logging.info("SQL-выражение для выборки: {}".format(land_code_clause))
            
            # Выбираем объекты с целевыми кодами LandCode
            arcpy.SelectLayerByAttribute_management(
                grid_layer,
                "NEW_SELECTION",
                land_code_clause
            )
            
            # Получаем количество выбранных объектов
            selected_count = int(arcpy.GetCount_management(grid_layer).getOutput(0))
            logging.info("Выбрано объектов с целевыми кодами LandCode: {}".format(selected_count))
            
            if selected_count == 0:
                logging.info("Нет объектов с целевыми кодами LandCode для проверки за пределами контура")
                return True
            
            # Создаем временный слой для хранения только объектов с целевыми кодами
            target_codes_layer = "temp_target_codes_layer"
            arcpy.CopyFeatures_management(grid_layer, "in_memory\\target_codes_objects")
            arcpy.MakeFeatureLayer_management("in_memory\\target_codes_objects", target_codes_layer)
            logging.info("Создан временный слой только с объектами целевых кодов")
            
            # Выбираем объекты с целевыми кодами, которые пересекаются с Lots_"Сокр"
            arcpy.SelectLayerByLocation_management(
                target_codes_layer,
                "INTERSECT",  # Пространственное отношение
                lots_layer,   # Слой, с которым проверяется пересечение
                "#",          # Дистанция поиска (# означает "по умолчанию")
                "NEW_SELECTION"  # Создаем новую выборку
            )
            
            # Инвертируем выборку, чтобы получить объекты целевых кодов, которые НЕ пересекаются с контуром
            arcpy.SelectLayerByAttribute_management(
                target_codes_layer,
                "SWITCH_SELECTION"  # Инвертируем выборку
            )
            
            # Получаем количество выбранных объектов (не пересекающихся)
            outside_count = int(arcpy.GetCount_management(target_codes_layer).getOutput(0))
            logging.info("Количество объектов с целевыми кодами за пределами контура Lots_\"Сокр\": {}".format(outside_count))
            
            if outside_count == 0:
                logging.info("Нет объектов с целевыми кодами за пределами контура Lots_\"Сокр\"")
                arcpy.Delete_management("in_memory\\target_codes_objects")
                arcpy.Delete_management(target_codes_layer)
                return True
            
            # Теперь выберем эти же объекты в исходном слое grid_layer
            # Создаем временную таблицу с идентификаторами объектов для удаления
            arcpy.CopyFeatures_management(target_codes_layer, "in_memory\\objects_to_delete")
            
            # Очищаем текущую выборку в grid_layer
            arcpy.SelectLayerByAttribute_management(grid_layer, "CLEAR_SELECTION")
            
            # Выбираем объекты в grid_layer, которые совпадают с объектами в objects_to_delete
            arcpy.SelectLayerByLocation_management(
                grid_layer,
                "ARE_IDENTICAL_TO",  # Объекты должны быть идентичны
                "in_memory\\objects_to_delete",
                "#",  # Дистанция поиска
                "NEW_SELECTION"  # Создаем новую выборку
            )
            
            # Проверяем количество объектов, выбранных для удаления
            to_delete_count = int(arcpy.GetCount_management(grid_layer).getOutput(0))
            logging.info("Количество объектов выбранных для удаления: {}".format(to_delete_count))
            
            # Удаляем выбранные объекты с целевыми кодами, находящиеся за пределами контура
            arcpy.DeleteFeatures_management(grid_layer)
            logging.info("Удалено {} объектов с кодами {} за пределами контура Lots_\"Сокр\"".format(
                to_delete_count, ", ".join(map(str, target_codes))))
            
            # Снимаем выборку
            arcpy.SelectLayerByAttribute_management(grid_layer, "CLEAR_SELECTION")
            
            # Очищаем временные слои и данные
            arcpy.Delete_management("in_memory\\target_codes_objects")
            arcpy.Delete_management("in_memory\\objects_to_delete")
            arcpy.Delete_management(target_codes_layer)
            arcpy.Delete_management(grid_layer)
            arcpy.Delete_management(lots_layer)
            
            # Обновляем количество объектов после удаления
            final_count = int(arcpy.GetCount_management(grid_path).getOutput(0))
            logging.info("Итоговое количество объектов в слое Land_\"Сокр\"_сетка после фильтрации: {}".format(final_count))
            
            return True
            
        except Exception as e:
            error_message = "Ошибка при сравнении с границами Lots_\"Сокр\": {}".format(str(e))
            logging.error(error_message)
            log_exception(e, "Ошибка при фильтрации объектов за границами контура")
            messagebox.showwarning("Предупреждение", error_message)
            return False

class ValueSelector:
    def __init__(self, master, gdb_path):
        self.master = master
        self.gdb_path = gdb_path
        self.selected_value = None
        self.shortened_name = None
        
        # Настройка окна
        master.title("Выбор значения из Lots")
        master.geometry("600x500")  # Увеличим высоту окна
        
        # Основной фрейм
        self.main_frame = tk.Frame(master, padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Заголовок
        self.label = tk.Label(self.main_frame, 
                             text="Выберите значение из поля 'UsName_1' набора 'Lots'",
                             font=("Arial", 12))
        self.label.pack(pady=10)
        
        # Инструкция
        self.instruction = tk.Label(self.main_frame,
                                  text="Дважды щелкните на элементе списка или выделите элемент и нажмите кнопку 'Выбрать'",
                                  font=("Arial", 9),
                                  fg="blue")
        self.instruction.pack(pady=5)
        
        # Фрейм для списка
        self.list_frame = tk.Frame(self.main_frame)
        self.list_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Скроллбар для списка
        self.scrollbar = tk.Scrollbar(self.list_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Список значений
        self.listbox = tk.Listbox(self.list_frame, 
                                 yscrollcommand=self.scrollbar.set,
                                 font=("Arial", 10),
                                 height=15,
                                 width=70,
                                 selectmode=tk.SINGLE)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.listbox.yview)
        
        # Привязка двойного клика для выбора
        self.listbox.bind('<Double-1>', lambda event: self.select_value())
        
        # Привязка клавиши Enter для выбора
        self.listbox.bind('<Return>', lambda event: self.select_value())
        
        # Привязка обработки нажатия клавиши Enter для всего окна
        self.master.bind('<Return>', lambda event: self.select_value())
        
        # Кнопки в отдельном фрейме с явным указанием размеров
        self.button_frame = tk.Frame(self.main_frame, height=50)
        self.button_frame.pack(fill=tk.X, pady=20)
        
        # Принудительно устанавливаем размер для фрейма кнопок
        self.button_frame.pack_propagate(False)
        
        # Кнопка выбора
        self.select_button = tk.Button(self.button_frame, 
                                     text="ВЫБРАТЬ", 
                                     command=self.select_value,
                                     bg="#4CAF50",  # Зеленый цвет
                                     fg="white",
                                     font=("Arial", 10, "bold"),
                                     width=15,
                                     height=2)
        self.select_button.pack(side=tk.LEFT, padx=10, expand=True)
        
        # Кнопка отмены
        self.cancel_button = tk.Button(self.button_frame, 
                                     text="ОТМЕНА", 
                                     command=master.destroy,
                                     bg="#f44336",  # Красный цвет
                                     fg="white",
                                     font=("Arial", 10, "bold"),
                                     width=15,
                                     height=2)
        self.cancel_button.pack(side=tk.RIGHT, padx=10, expand=True)
        
        # Заполнить список значениями
        self.load_values()
        
        # Фокус на список
        self.listbox.focus_set()
    
    def load_values(self):
        if not ARCPY_AVAILABLE:
            messagebox.showerror("Ошибка", "Для работы с данными необходим модуль arcpy")
            self.master.destroy()
            return
        
        try:
            # Устанавливаем рабочее пространство
            arcpy.env.workspace = self.gdb_path
            
            # Ищем набор "Lots"
            fc_name = "Lots"
            feature_classes = arcpy.ListFeatureClasses()
            
            if fc_name not in feature_classes:
                # Проверяем в наборах данных
                datasets = arcpy.ListDatasets()
                found = False
                
                for dataset in datasets:
                    arcpy.env.workspace = os.path.join(self.gdb_path, dataset)
                    if fc_name in arcpy.ListFeatureClasses():
                        fc_name = os.path.join(dataset, fc_name)
                        found = True
                        break
                
                arcpy.env.workspace = self.gdb_path
                
                if not found:
                    messagebox.showerror("Ошибка", "Класс объектов 'Lots' не найден в базе геоданных")
                    self.master.destroy()
                    return
            
            # Получаем уникальные значения поля UsName_1
            values = set()
            with arcpy.da.SearchCursor(fc_name, ["UsName_1"]) as cursor:
                for row in cursor:
                    if row[0]:  # Проверка на None и пустые значения
                        values.add(row[0])
            
            # Сортируем значения
            sorted_values = sorted(values)
            
            # Добавляем в список
            for value in sorted_values:
                self.listbox.insert(tk.END, value)
                
        except Exception as e:
            messagebox.showerror("Ошибка", "Не удалось загрузить значения: {}".format(str(e)))
            self.master.destroy()
    
    def select_value(self):
        # Получить выбранное значение
        selection = self.listbox.curselection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Пожалуйста, выберите значение из списка")
            return
        
        # Получить выбранное значение
        self.selected_value = self.listbox.get(selection[0])
        
        # Запросить сокращенное название
        short_name = simpledialog.askstring(
            "Ввод сокращения", 
            "Введите сокращенное название для '{}'".format(self.selected_value),
            parent=self.master
        )
        
        if short_name:
            self.shortened_name = short_name
            messagebox.showinfo("Успех", 
                                "Выбрано: {}\nСокращение: {}".format(self.selected_value, self.shortened_name))
            self.master.destroy()
        else:
            messagebox.showwarning("Предупреждение", "Необходимо ввести сокращенное название")

class DataProcessor:
    def __init__(self, gdb_path, selected_value, shortened_name):
        self.gdb_path = gdb_path
        self.selected_value = selected_value
        self.shortened_name = shortened_name
        
    def process_data(self):
        """Основной метод обработки данных"""
        if not ARCPY_AVAILABLE:
            logging.error("Модуль arcpy недоступен")
            messagebox.showerror("Ошибка", "Модуль arcpy не доступен для обработки данных")
            return False
        
        try:
            logging.info("Начало обработки данных")
            logging.info("Параметры: GDB={}, выбранное значение={}, сокращение={}".format(
                self.gdb_path, self.selected_value, self.shortened_name))
            
            # Очищаем все рабочие пространства перед началом работы
            arcpy.ClearWorkspaceCache_management()
            
            # Устанавливаем рабочее пространство
            arcpy.env.workspace = self.gdb_path
            logging.info("Установлено рабочее пространство: {}".format(self.gdb_path))
            
            # Формируем имя нового набора данных на основе значения UsName_1
            # Заменяем пробелы на "_" и удаляем кавычки
            cleaned_usname = self.selected_value.replace(" ", "_").replace("\"", "").replace("'", "")
            new_dataset_name = "_{}".format(cleaned_usname)
            logging.info("Имя нового набора данных: {}".format(new_dataset_name))
            
            # Проверяем существование набора "Копия"
            logging.info("Проверка наличия набора данных 'Копия'")
            datasets = arcpy.ListDatasets()
            
            if "Копия" not in datasets:
                logging.error("Набор 'Копия' не найден в базе геоданных")
                messagebox.showerror("Ошибка", "Набор данных 'Копия' не найден в базе геоданных")
                return False
            
            # Проверяем, существует ли уже набор с новым именем
            if new_dataset_name in datasets:
                logging.warning("Набор данных '{}' уже существует".format(new_dataset_name))
                confirm = messagebox.askyesno(
                    "Предупреждение", 
                    "Набор данных '{}' уже существует. Хотите заменить его?".format(new_dataset_name)
                )
                
                if confirm:
                    try:
                        # Удаляем существующий набор данных
                        arcpy.Delete_management(os.path.join(self.gdb_path, new_dataset_name))
                        logging.info("Удален существующий набор данных: {}".format(new_dataset_name))
                    except Exception as delete_err:
                        logging.error("Ошибка при удалении набора данных: {}".format(str(delete_err)))
                        messagebox.showerror(
                            "Ошибка", 
                            "Не удалось удалить существующий набор данных. Возможно, он используется."
                        )
                        return False
                else:
                    logging.info("Пользователь отменил замену существующего набора данных")
                    return False
            
            # Копируем набор "Копия" с новым именем
            try:
                logging.info("Копирование набора 'Копия' с новым именем '{}'".format(new_dataset_name))
                
                # Получаем информацию о пространственной привязке исходного набора
                source_path = os.path.join(self.gdb_path, "Копия")
                desc = arcpy.Describe(source_path)
                spatial_reference = desc.spatialReference
                logging.info("Пространственная привязка: {}".format(spatial_reference.name))
                
                # Создаем новый набор данных с той же пространственной привязкой
                create_result = arcpy.CreateFeatureDataset_management(
                    self.gdb_path,
                    new_dataset_name,
                    spatial_reference
                )
                
                logging.info("Создан новый набор данных: {}".format(create_result.getOutput(0)))
                
                # Обновляем кэш и получаем список классов объектов в исходном наборе
                arcpy.ClearWorkspaceCache_management()
                arcpy.env.workspace = source_path
                source_fcs = arcpy.ListFeatureClasses()
                logging.info("Классы объектов в исходном наборе: {}".format(", ".join(source_fcs) if source_fcs else "нет"))
                
                # Возвращаемся к корневому рабочему пространству
                arcpy.env.workspace = self.gdb_path
                
            except Exception as copy_err:
                logging.error("Ошибка при создании нового набора данных: {}".format(str(copy_err)))
                messagebox.showerror("Ошибка", "Не удалось создать новый набор данных: {}".format(str(copy_err)))
                return False
            
            # Найдем класс объектов Lots в базе данных
            try:
                logging.info("Поиск класса объектов 'Lots'")
                
                # Ищем Lots в корне базы данных
                arcpy.env.workspace = self.gdb_path
                lots_path = None
                root_fcs = arcpy.ListFeatureClasses()
                
                if "Lots" in root_fcs:
                    lots_path = os.path.join(self.gdb_path, "Lots")
                    logging.info("Найден класс объектов 'Lots' в корне: {}".format(lots_path))
                else:
                    # Ищем Lots в наборах данных
                    for ds in datasets:
                        if ds == new_dataset_name:
                            continue  # Пропускаем только что созданный набор
                        
                        arcpy.env.workspace = os.path.join(self.gdb_path, ds)
                        ds_fcs = arcpy.ListFeatureClasses()
                        
                        if "Lots" in ds_fcs:
                            lots_path = os.path.join(self.gdb_path, ds, "Lots")
                            logging.info("Найден класс объектов 'Lots' в наборе '{}': {}".format(ds, lots_path))
                            break
                
                # Сбрасываем рабочее пространство
                arcpy.env.workspace = self.gdb_path
                
                if not lots_path:
                    logging.error("Класс объектов 'Lots' не найден в базе данных")
                    messagebox.showerror("Ошибка", "Класс объектов 'Lots' не найден в базе данных")
                    return False
                
            except Exception as find_err:
                logging.error("Ошибка при поиске класса объектов 'Lots': {}".format(str(find_err)))
                messagebox.showerror("Ошибка", "Не удалось найти класс объектов 'Lots': {}".format(str(find_err)))
                return False
            
            # Используем инструмент Select_analysis для извлечения данных
            try:
                # Формируем имя для нового класса объектов
                new_fc_name = "Lots_{}".format(self.shortened_name)
                target_fc_path = os.path.join(self.gdb_path, new_dataset_name, new_fc_name)
                logging.info("Путь для нового класса объектов: {}".format(target_fc_path))
                
                # Проверяем существование класса объектов
                if arcpy.Exists(target_fc_path):
                    logging.warning("Класс объектов '{}' уже существует".format(target_fc_path))
                    if messagebox.askyesno("Предупреждение", "Класс объектов '{}' уже существует. Заменить?".format(new_fc_name)):
                        arcpy.Delete_management(target_fc_path)
                        logging.info("Удален существующий класс объектов: {}".format(target_fc_path))
                    else:
                        logging.info("Пользователь отменил замену существующего класса объектов")
                        return False
                
                # Формируем выражение выборки
                where_clause = "\"UsName_1\" = '{}'".format(self.selected_value)
                logging.info("Выражение выборки: {}".format(where_clause))
                
                # Используем инструмент Select_analysis для извлечения данных
                logging.info("Запуск инструмента Select_analysis")
                logging.info("Входные данные: {}".format(lots_path))
                logging.info("Выходные данные: {}".format(target_fc_path))
                
                # Устанавливаем параметр перезаписи существующих данных
                arcpy.env.overwriteOutput = True
                
                # Выполняем инструмент Select_analysis
                select_result = arcpy.Select_analysis(
                    lots_path,
                    target_fc_path,
                    where_clause
                )
                
                logging.info("Создан новый класс объектов: {}".format(select_result.getOutput(0)))
                
                # Проверяем количество извлеченных объектов
                count_result = arcpy.GetCount_management(target_fc_path)
                feature_count = int(count_result.getOutput(0))
                logging.info("Количество извлеченных объектов: {}".format(feature_count))
                
                # Создаем копию слоя с названием Lots_"Сокр"_контур
                try:
                    contour_fc_name = "Lots_{}_контур".format(self.shortened_name)
                    contour_fc_path = os.path.join(self.gdb_path, new_dataset_name, contour_fc_name)
                    logging.info("Создание копии слоя с названием: {}".format(contour_fc_name))
                    
                    # Проверяем существование класса объектов
                    if arcpy.Exists(contour_fc_path):
                        logging.warning("Класс объектов '{}' уже существует".format(contour_fc_path))
                        if messagebox.askyesno("Предупреждение", "Класс объектов '{}' уже существует. Заменить?".format(contour_fc_name)):
                            arcpy.Delete_management(contour_fc_path)
                            logging.info("Удален существующий класс объектов: {}".format(contour_fc_path))
                        else:
                            logging.info("Пользователь отменил замену существующего класса объектов - контур")
                            # Продолжаем выполнение без создания контура
                    
                    if not arcpy.Exists(contour_fc_path):
                        # Копируем класс объектов с новым именем
                        arcpy.Copy_management(target_fc_path, contour_fc_path)
                        logging.info("Создана копия класса объектов: {}".format(contour_fc_path))
                        
                        # Проверяем количество объектов в новом классе
                        contour_count = int(arcpy.GetCount_management(contour_fc_path).getOutput(0))
                        logging.info("Количество объектов в контуре: {}".format(contour_count))
                        
                        # Обработка контурного слоя: удаление данных из таблицы атрибутов и создание буфера
                        try:
                            logging.info("Начало обработки контурного слоя...")
                            
                            # 1. Удаляем все данные из таблицы атрибутов контурного слоя
                            arcpy.DeleteRows_management(contour_fc_path)
                            logging.info("Данные из таблицы атрибутов контурного слоя удалены")
                            
                            # 2. Создаем буфер вокруг исходных полигонов (0.5 км = 500 м)
                            buffer_distance = "500 Meters"
                            temp_buffer_path = os.path.join("in_memory", "temp_buffer")
                            
                            logging.info("Создание буфера с расстоянием {} вокруг участков".format(buffer_distance))
                            arcpy.Buffer_analysis(
                                target_fc_path,
                                temp_buffer_path,
                                buffer_distance,
                                "FULL", 
                                "ROUND", 
                                "ALL"  # Объединяем все полигоны
                            )
                            
                            # 3. Объединяем полигоны в пределах 2 км друг от друга
                            # Сначала создаем буфер 2 км, затем растворяем его и снова создаем буфер внутрь на 1.5 км
                            dissolve_buffer_path = os.path.join("in_memory", "dissolve_buffer")
                            
                            # Создаем буфер 2 км для определения близлежащих участков
                            logging.info("Создание буфера 2 км для определения близлежащих участков")
                            arcpy.Buffer_analysis(
                                target_fc_path,
                                dissolve_buffer_path,
                                "2000 Meters",
                                "FULL", 
                                "ROUND", 
                                "ALL"  # Объединяем все полигоны
                            )
                            
                            # Создаем отрицательный буфер -1.5 км (2 км - 0.5 км), чтобы получить контур с отступом 0.5 км
                            final_buffer_path = os.path.join("in_memory", "final_buffer")
                            logging.info("Создание итогового контура с отступом 0.5 км")
                            arcpy.Buffer_analysis(
                                dissolve_buffer_path,
                                final_buffer_path,
                                "-1500 Meters",
                                "FULL", 
                                "ROUND", 
                                "ALL"
                            )
                            
                            # 4. Копируем результат в контурный слой
                            logging.info("Копирование результата в контурный слой")
                            arcpy.Append_management(
                                final_buffer_path,
                                contour_fc_path,
                                "NO_TEST"  # Не проверяем схему данных
                            )
                            
                            # 5. Очищаем временные данные
                            arcpy.Delete_management("in_memory")
                            
                            logging.info("Обработка контурного слоя завершена успешно")
                            
                            # 6. Вырезаем данные из класса Land по границам контурного слоя
                            try:
                                logging.info("Начало вырезания данных из класса Land...")
                                
                                # Ищем класс Land в базе геоданных
                                land_path = None
                                
                                # Сначала ищем в корне базы данных
                                arcpy.env.workspace = self.gdb_path
                                root_fcs = arcpy.ListFeatureClasses()
                                
                                if "Land" in root_fcs:
                                    land_path = os.path.join(self.gdb_path, "Land")
                                    logging.info("Найден класс объектов 'Land' в корне: {}".format(land_path))
                                else:
                                    # Ищем Land в наборах данных
                                    datasets = arcpy.ListDatasets()
                                    for ds in datasets:
                                        arcpy.env.workspace = os.path.join(self.gdb_path, ds)
                                        ds_fcs = arcpy.ListFeatureClasses()
                                        
                                        if "Land" in ds_fcs:
                                            land_path = os.path.join(self.gdb_path, ds, "Land")
                                            logging.info("Найден класс объектов 'Land' в наборе '{}': {}".format(ds, land_path))
                                            break
                                
                                # Сбрасываем рабочее пространство
                                arcpy.env.workspace = self.gdb_path
                                
                                if not land_path:
                                    logging.error("Класс объектов 'Land' не найден в базе данных")
                                    messagebox.showwarning("Предупреждение", "Класс объектов 'Land' не найден в базе данных")
                                else:
                                    # Формируем имя для нового класса объектов
                                    land_clip_name = "Land_{}".format(self.shortened_name)
                                    land_clip_path = os.path.join(self.gdb_path, new_dataset_name, land_clip_name)
                                    logging.info("Путь для нового класса объектов: {}".format(land_clip_path))
                                    
                                    # Проверяем существование класса объектов
                                    if arcpy.Exists(land_clip_path):
                                        logging.warning("Класс объектов '{}' уже существует".format(land_clip_path))
                                        if messagebox.askyesno("Предупреждение", "Класс объектов '{}' уже существует. Заменить?".format(land_clip_name)):
                                            arcpy.Delete_management(land_clip_path)
                                            logging.info("Удален существующий класс объектов: {}".format(land_clip_path))
                                        else:
                                            logging.info("Пользователь отменил замену существующего класса объектов")
                                            # Продолжаем без вырезания
                                    
                                    if not arcpy.Exists(land_clip_path):
                                        # Вырезаем данные из Land по контуру
                                        logging.info("Вырезание данных из '{}' по '{}', сохранение в '{}'".format(
                                            land_path, target_fc_path, land_clip_path))
                                        
                                        # Используем инструмент Clip
                                        arcpy.Clip_analysis(
                                            land_path,  # Входной класс
                                            target_fc_path,  # Вырезающий класс
                                            land_clip_path  # Выходной класс
                                        )
                                        
                                        logging.info("Вырезание данных завершено успешно")
                                        
                                        # Получаем количество объектов в результате
                                        clip_count = int(arcpy.GetCount_management(land_clip_path).getOutput(0))
                                        logging.info("Количество объектов в результате вырезания: {}".format(clip_count))
                                        
                                        # Глобальная переменная для использования в других частях кода
                                        self.land_clip_path = land_clip_path
                                        self.land_clip_name = land_clip_name
                                        
                                        # Оставляем только объекты с LandType 101, 102, 103
                                        try:
                                            logging.info("Оставляем только объекты с LandType 101, 102, 103...")
                                            
                                            # Создаем временный слой
                                            temp_land_layer = "temp_land_layer"
                                            arcpy.MakeFeatureLayer_management(land_clip_path, temp_land_layer)
                                            
                                            # Формируем SQL-выражение для выбора нужных типов
                                            landtype_sql = "LandType IN (101, 102, 103)"
                                            
                                            # Выбираем объекты с нужными типами
                                            arcpy.SelectLayerByAttribute_management(
                                                temp_land_layer,
                                                "NEW_SELECTION",
                                                landtype_sql
                                            )
                                            
                                            # Проверяем количество выбранных объектов для сохранения
                                            to_keep_count = int(arcpy.GetCount_management(temp_land_layer).getOutput(0))
                                            logging.info("Количество объектов для сохранения (LandType 101, 102, 103): {}".format(to_keep_count))
                                            
                                            # Инвертируем выборку, чтобы выбрать все, кроме нужных типов
                                            arcpy.SelectLayerByAttribute_management(
                                                temp_land_layer,
                                                "SWITCH_SELECTION"
                                            )
                                            
                                            # Проверяем количество выбранных объектов для удаления
                                            to_delete_count = int(arcpy.GetCount_management(temp_land_layer).getOutput(0))
                                            logging.info("Количество объектов для удаления (не LandType 101, 102, 103): {}".format(to_delete_count))
                                            
                                            # Удаляем выбранные объекты
                                            if to_delete_count > 0:
                                                arcpy.DeleteFeatures_management(temp_land_layer)
                                                logging.info("Удалено {} объектов с LandType не равным 101, 102, 103".format(to_delete_count))
                                            
                                            # Снимаем выборку
                                            arcpy.SelectLayerByAttribute_management(temp_land_layer, "CLEAR_SELECTION")
                                            
                                            # Удаляем временный слой
                                            arcpy.Delete_management(temp_land_layer)
                                            
                                            # Получаем итоговое количество объектов
                                            final_land_count = int(arcpy.GetCount_management(land_clip_path).getOutput(0))
                                            logging.info("Итоговое количество объектов в Land_\"Сокр\" после фильтрации: {}".format(final_land_count))
                                            
                                        except Exception as filter_err:
                                            logging.error("Ошибка при фильтрации объектов Land_\"Сокр\": {}".format(str(filter_err)))
                                            logging.error(traceback.format_exc())
                                            messagebox.showwarning("Предупреждение", 
                                                                "Ошибка при фильтрации объектов Land_\"Сокр\":\n{}".format(str(filter_err)))
                                            
                                        # Создаем Land_"Сокр"_контур напрямую из Land по границам Lots_"Сокр"_контур
                                        land_contour_name = "Land_{}_контур".format(self.shortened_name)
                                        land_contour_path = os.path.join(self.gdb_path, new_dataset_name, land_contour_name)
                                        logging.info("Создание Land_\"Сокр\"_контур напрямую из Land по границам Lots_\"Сокр\"_контур...")
                                        
                                        # Проверяем существование класса объектов
                                        if arcpy.Exists(land_contour_path):
                                            logging.warning("Класс объектов '{}' уже существует".format(land_contour_path))
                                            if messagebox.askyesno("Предупреждение", "Класс объектов '{}' уже существует. Заменить?".format(land_contour_name)):
                                                arcpy.Delete_management(land_contour_path)
                                                logging.info("Удален существующий класс объектов: {}".format(land_contour_path))
                                            else:
                                                logging.info("Пользователь отменил замену существующего класса объектов - контур")
                                                # Продолжаем вырезание других классов
                                        
                                        if not arcpy.Exists(land_contour_path):
                                            # Вырезаем данные из Land по контуру Lots_"Сокр"_контур
                                            logging.info("Вырезание данных из '{}' по '{}', сохранение в '{}'".format(
                                                land_path, contour_fc_path, land_contour_path))
                                            
                                            # Используем инструмент Clip
                                            arcpy.Clip_analysis(
                                                land_path,  # Входной класс (Land)
                                                contour_fc_path,  # Вырезающий класс (Lots_"Сокр"_контур)
                                                land_contour_path  # Выходной класс (Land_"Сокр"_контур)
                                            )
                                            
                                            logging.info("Вырезание данных Land_\"Сокр\"_контур завершено успешно")
                                            
                                            # Удаляем объекты с LandType 101, 102, 103 и LandCode 326 из Land_"Сокр"_контур
                                            try:
                                                logging.info("Удаляем объекты с LandType 101, 102, 103 и LandCode 326 из Land_\"Сокр\"_контур...")
                                                
                                                # Создаем временный слой
                                                temp_contour_layer = "temp_contour_layer"
                                                arcpy.MakeFeatureLayer_management(land_contour_path, temp_contour_layer)
                                                
                                                # Формируем SQL-выражение для выбора объектов с LandType 101, 102, 103
                                                landtype_sql = "LandType IN (101, 102, 103)"
                                                
                                                # Выбираем объекты с LandType 101, 102, 103
                                                arcpy.SelectLayerByAttribute_management(
                                                    temp_contour_layer,
                                                    "NEW_SELECTION",
                                                    landtype_sql
                                                )
                                                
                                                # Проверяем количество выбранных объектов
                                                landtype_count = int(arcpy.GetCount_management(temp_contour_layer).getOutput(0))
                                                logging.info("Выбрано {} объектов с LandType 101, 102, 103".format(landtype_count))
                                                
                                                # Удаляем выбранные объекты
                                                if landtype_count > 0:
                                                    arcpy.DeleteFeatures_management(temp_contour_layer)
                                                    logging.info("Удалено {} объектов с LandType 101, 102, 103".format(landtype_count))
                                                
                                                # Снимаем выборку
                                                arcpy.SelectLayerByAttribute_management(temp_contour_layer, "CLEAR_SELECTION")
                                                
                                                # Выбираем объекты с LandCode 326
                                                landcode_sql = "LandCode = 326"
                                                arcpy.SelectLayerByAttribute_management(
                                                    temp_contour_layer,
                                                    "NEW_SELECTION",
                                                    landcode_sql
                                                )
                                                
                                                # Проверяем количество выбранных объектов
                                                landcode_count = int(arcpy.GetCount_management(temp_contour_layer).getOutput(0))
                                                logging.info("Выбрано {} объектов с LandCode 326".format(landcode_count))
                                                
                                                # Удаляем выбранные объекты
                                                if landcode_count > 0:
                                                    arcpy.DeleteFeatures_management(temp_contour_layer)
                                                    logging.info("Удалено {} объектов с LandCode 326".format(landcode_count))
                                                
                                                # Снимаем выборку
                                                arcpy.SelectLayerByAttribute_management(temp_contour_layer, "CLEAR_SELECTION")
                                                
                                                # Удаляем временный слой
                                                arcpy.Delete_management(temp_contour_layer)
                                                
                                                # Получаем итоговое количество объектов
                                                final_contour_count = int(arcpy.GetCount_management(land_contour_path).getOutput(0))
                                                logging.info("Итоговое количество объектов в Land_\"Сокр\"_контур после фильтрации: {}".format(final_contour_count))
                                                
                                                # Сохраняем ссылки на пути
                                                self.land_contour_path = land_contour_path
                                                self.land_contour_name = land_contour_name
                                                
                                            except Exception as contour_filter_err:
                                                logging.error("Ошибка при фильтрации объектов Land_\"Сокр\"_контур: {}".format(str(contour_filter_err)))
                                                logging.error(traceback.format_exc())
                                                messagebox.showwarning("Предупреждение", 
                                                                    "Ошибка при фильтрации объектов Land_\"Сокр\"_контур:\n{}".format(str(contour_filter_err)))
                            
                            except Exception as clip_err:
                                logging.error("Ошибка при вырезании данных из Land: {}".format(str(clip_err)))
                                logging.error(traceback.format_exc())
                                messagebox.showwarning("Предупреждение", 
                                                    "Ошибка при вырезании данных из Land:\n{}".format(str(clip_err)))
                                # Продолжаем выполнение
                                
                            # 7. Вырезаем данные из класса Admi по границам контурного слоя
                            try:
                                logging.info("Начало вырезания данных из класса Admi...")
                                
                                # Ищем класс Admi в базе геоданных
                                admi_path = None
                                
                                # Сначала ищем в корне базы данных
                                arcpy.env.workspace = self.gdb_path
                                root_fcs = arcpy.ListFeatureClasses()
                                
                                if "Admi" in root_fcs:
                                    admi_path = os.path.join(self.gdb_path, "Admi")
                                    logging.info("Найден класс объектов 'Admi' в корне: {}".format(admi_path))
                                else:
                                    # Ищем Admi в наборах данных
                                    datasets = arcpy.ListDatasets()
                                    for ds in datasets:
                                        arcpy.env.workspace = os.path.join(self.gdb_path, ds)
                                        ds_fcs = arcpy.ListFeatureClasses()
                                        
                                        if "Admi" in ds_fcs:
                                            admi_path = os.path.join(self.gdb_path, ds, "Admi")
                                            logging.info("Найден класс объектов 'Admi' в наборе '{}': {}".format(ds, admi_path))
                                            break
                                
                                # Сбрасываем рабочее пространство
                                arcpy.env.workspace = self.gdb_path
                                
                                if not admi_path:
                                    logging.error("Класс объектов 'Admi' не найден в базе данных")
                                    messagebox.showwarning("Предупреждение", "Класс объектов 'Admi' не найден в базе данных")
                                else:
                                    # Формируем имя для нового класса объектов
                                    admi_clip_name = "Admi_{}".format(self.shortened_name)
                                    admi_clip_path = os.path.join(self.gdb_path, new_dataset_name, admi_clip_name)
                                    logging.info("Путь для нового класса объектов: {}".format(admi_clip_path))
                                    
                                    # Проверяем существование класса объектов
                                    if arcpy.Exists(admi_clip_path):
                                        logging.warning("Класс объектов '{}' уже существует".format(admi_clip_path))
                                        if messagebox.askyesno("Предупреждение", "Класс объектов '{}' уже существует. Заменить?".format(admi_clip_name)):
                                            arcpy.Delete_management(admi_clip_path)
                                            logging.info("Удален существующий класс объектов: {}".format(admi_clip_path))
                                        else:
                                            logging.info("Пользователь отменил замену существующего класса объектов")
                                            # Продолжаем без вырезания
                                    
                                    if not arcpy.Exists(admi_clip_path):
                                        # Вырезаем данные из Admi по Lots_"Сокр"_контур
                                        logging.info("Вырезание данных из '{}' по '{}', сохранение в '{}'".format(
                                            admi_path, contour_fc_path, admi_clip_path))
                                        
                                        # Используем инструмент Clip
                                        arcpy.Clip_analysis(
                                            admi_path,  # Входной класс
                                            contour_fc_path,  # Вырезающий класс (Lots_"Сокр"_контур)
                                            admi_clip_path  # Выходной класс
                                        )
                                        
                                        logging.info("Вырезание данных Admi завершено успешно")
                                        
                                        # Получаем количество объектов в результате
                                        admi_clip_count = int(arcpy.GetCount_management(admi_clip_path).getOutput(0))
                                        logging.info("Количество объектов в результате вырезания Admi: {}".format(admi_clip_count))
                                        
                                        # Глобальная переменная для использования в других частях кода
                                        self.admi_clip_path = admi_clip_path
                                        self.admi_clip_name = admi_clip_name
                                        self.admi_clip_count = admi_clip_count
                                            
                            except Exception as admi_clip_err:
                                logging.error("Ошибка при вырезании данных из Admi: {}".format(str(admi_clip_err)))
                                logging.error(traceback.format_exc())
                                messagebox.showwarning("Предупреждение", 
                                                      "Ошибка при вырезании данных из Admi:\n{}".format(str(admi_clip_err)))
                                # Продолжаем выполнение
                        
                        except Exception as process_err:
                            logging.error("Ошибка при обработке контурного слоя: {}".format(str(process_err)))
                            logging.error(traceback.format_exc())
                            messagebox.showwarning(
                                "Предупреждение", 
                                "Ошибка при обработке контурного слоя:\n{}".format(str(process_err))
                            )
                except Exception as contour_err:
                    logging.error("Ошибка при создании контура: {}".format(str(contour_err)))
                    # Продолжаем выполнение, так как это не должно останавливать основной процесс
                
                if feature_count == 0:
                    logging.warning("Извлечено 0 объектов по условию '{}'".format(where_clause))
                    messagebox.showwarning(
                        "Предупреждение", 
                        "Не найдено объектов, соответствующих условию:\n{}".format(where_clause)
                    )
                
                # Обновляем кэш
                arcpy.ClearWorkspaceCache_management()
                arcpy.RefreshCatalog(self.gdb_path)
                arcpy.RefreshCatalog(os.path.join(self.gdb_path, new_dataset_name))
                arcpy.RefreshTOC()  # Обновляем таблицу содержания
                arcpy.RefreshActiveView()  # Обновляем активный вид
                
                # Добавляем слой в таблицу содержания текущего документа карты и отображаем на экране
                try:
                    # Проверяем, запущен ли скрипт из ArcMap или ArcGIS Pro
                    if arcpy.ProductInfo() in ['ArcView', 'ArcEditor', 'ArcInfo']:  # ArcMap
                        import arcpy.mapping as mapping
                        mxd = mapping.MapDocument("CURRENT")
                        df = mxd.activeDataFrame
                        layer = mapping.Layer(target_fc_path)
                        mapping.AddLayer(df, layer, "TOP")
                        logging.info("Слой '{}' добавлен в таблицу содержания ArcMap".format(new_fc_name))
                        
                        # Добавляем контурный слой в карту, если он существует
                        if 'contour_fc_path' in locals() and arcpy.Exists(contour_fc_path):
                            contour_layer = mapping.Layer(contour_fc_path)
                            mapping.AddLayer(df, contour_layer, "TOP")
                            logging.info("Контурный слой '{}' добавлен в таблицу содержания ArcMap".format(contour_fc_name))
                        
                        # Добавляем вырезанный слой Land в карту, если он существует
                        if hasattr(self, 'land_clip_path') and arcpy.Exists(self.land_clip_path):
                            land_clip_layer = mapping.Layer(self.land_clip_path)
                            mapping.AddLayer(df, land_clip_layer, "TOP")
                            logging.info("Вырезанный слой Land '{}' добавлен в таблицу содержания ArcMap".format(self.land_clip_name))
                        
                        # Добавляем вырезанный слой Admi в карту, если он существует
                        if hasattr(self, 'admi_clip_path') and arcpy.Exists(self.admi_clip_path):
                            admi_clip_layer = mapping.Layer(self.admi_clip_path)
                            mapping.AddLayer(df, admi_clip_layer, "TOP")
                            logging.info("Вырезанный слой Admi '{}' добавлен в таблицу содержания ArcMap".format(self.admi_clip_name))
                        
                        # Сохраняем MXD файл с именем _UsName_1.mxd (без кавычек)
                        cleaned_usname = self.selected_value.replace(" ", "_").replace("\"", "").replace("'", "")
                        mxd_name = "_{}".format(cleaned_usname)
                        mxd_path = os.path.join(os.path.dirname(self.gdb_path), "{}.mxd".format(mxd_name))
                        
                        # Проверяем, существует ли файл MXD
                        if os.path.exists(mxd_path):
                            if messagebox.askyesno("Предупреждение", "Файл карты {}.mxd уже существует. Заменить?".format(mxd_name)):
                                mxd.saveACopy(mxd_path)
                                logging.info("Файл карты сохранен: {}".format(mxd_path))
                            else:
                                logging.info("Пользователь отменил сохранение файла карты")
                        else:
                            mxd.saveACopy(mxd_path)
                            logging.info("Файл карты сохранен: {}".format(mxd_path))
                        
                        # Обновляем каталог, чтобы слой и набор данных отображались в ArcMap
                        arcpy.RefreshCatalog(self.gdb_path)
                        arcpy.RefreshCatalog(os.path.join(self.gdb_path, new_dataset_name))
                        arcpy.RefreshTOC()  # Обновляем таблицу содержания
                        arcpy.RefreshActiveView()  # Обновляем активный вид
                        
                    elif arcpy.ProductInfo() == 'ArcGISPro':  # ArcGIS Pro
                        import arcpy.mp as mp
                        aprx = mp.ArcGISProject("CURRENT")
                        m = aprx.activeMap
                        m.addDataFromPath(target_fc_path)
                        logging.info("Слой '{}' добавлен в таблицу содержания ArcGIS Pro".format(new_fc_name))
                        
                        # Добавляем контурный слой в карту ArcGIS Pro
                        if 'contour_fc_path' in locals() and arcpy.Exists(contour_fc_path):
                            m.addDataFromPath(contour_fc_path)
                            logging.info("Слой '{}' добавлен в таблицу содержания ArcGIS Pro".format(contour_fc_name))
                        
                        # Добавляем вырезанный слой Land в карту ArcGIS Pro
                        if hasattr(self, 'land_clip_path') and arcpy.Exists(self.land_clip_path):
                            m.addDataFromPath(self.land_clip_path)
                            logging.info("Слой '{}' добавлен в таблицу содержания ArcGIS Pro".format(self.land_clip_name))
                        
                        # Добавляем вырезанный слой Admi в карту ArcGIS Pro
                        if hasattr(self, 'admi_clip_path') and arcpy.Exists(self.admi_clip_path):
                            m.addDataFromPath(self.admi_clip_path)
                            logging.info("Слой '{}' добавлен в таблицу содержания ArcGIS Pro".format(self.admi_clip_name))
                        
                        # Сохраняем APRX файл с именем _UsName_1.aprx (без кавычек)
                        cleaned_usname = self.selected_value.replace(" ", "_").replace("\"", "").replace("'", "")
                        aprx_name = "_{}".format(cleaned_usname)
                        aprx_path = os.path.join(os.path.dirname(self.gdb_path), "{}.aprx".format(aprx_name))
                        
                        # Проверяем, существует ли файл APRX
                        if os.path.exists(aprx_path):
                            if messagebox.askyesno("Предупреждение", "Файл проекта {}.aprx уже существует. Заменить?".format(aprx_name)):
                                aprx.saveACopy(aprx_path)
                                logging.info("Файл проекта сохранен: {}".format(aprx_path))
                            else:
                                logging.info("Пользователь отменил сохранение файла проекта")
                        else:
                            aprx.saveACopy(aprx_path)
                            logging.info("Файл проекта сохранен: {}".format(aprx_path))
                    else:
                        logging.warning("Скрипт запущен вне ArcMap/ArcGIS Pro или не удалось определить продукт")
                except Exception as map_err:
                    logging.warning("Не удалось добавить слой в таблицу содержания или сохранить файл: {}".format(str(map_err)))
                    # Продолжаем выполнение, так как это некритичная ошибка
                
                # Отображаем сообщение об успехе
                messagebox.showinfo(
                    "Успех", 
                    "Данные успешно обработаны:\n\n"
                    "Создан новый набор данных: {}\n"
                    "Создан новый класс объектов: {}\n"
                    "Извлечено объектов: {}\n"
                    "{}\n"
                    "{}\n"
                    "{}\n"
                    "{}".format(
                        new_dataset_name, 
                        new_fc_name, 
                        feature_count,
                        "Создан контурный слой: {}".format(contour_fc_name) if 'contour_fc_path' in locals() and arcpy.Exists(contour_fc_path) else "",
                        "Создан вырезанный слой Land: {}".format(self.land_clip_name) if hasattr(self, 'land_clip_path') and arcpy.Exists(self.land_clip_path) else "",
                        "Создан вырезанный слой Admi: {}".format(self.admi_clip_name) if hasattr(self, 'admi_clip_path') and arcpy.Exists(self.admi_clip_path) else "",
                        "Удалено объектов из Land: {}".format(self.deleted_features_count) if hasattr(self, 'deleted_features_count') else ""
                    )
                )
                
                logging.info("Обработка данных завершена успешно")
                return True
                
            except Exception as select_err:
                logging.error("Ошибка при извлечении данных: {}".format(str(select_err)))
                messagebox.showerror("Ошибка", "Не удалось извлечь данные: {}".format(str(select_err)))
                return False
            
        except Exception as e:
            logging.error("Общая ошибка при обработке данных: {}".format(str(e)))
            messagebox.showerror("Ошибка", "Произошла ошибка при обработке данных: {}".format(str(e)))
            return False

def main():
    # Создание и запуск интерфейса выбора GDB
    root = tk.Tk()
    gdb_selector = GDBSelector(root)
    
    # Делаем диалог модальным
    root.grab_set()
    root.focus_set()
    root.wait_window()
    
    gdb_path = gdb_selector.gdb_path
    if not gdb_path:
        logging.info("Пользователь отменил выбор базы геоданных или возникла ошибка")
        return
    
    # Создание и запуск интерфейса выбора базы данных прошлого тура
    root = tk.Tk()
    old_forest_selector = OldForestGDBSelector(root)
    
    # Делаем диалог модальным
    root.grab_set()
    root.focus_set()
    root.wait_window()
    
    old_db_path = old_forest_selector.db_path
    if not old_db_path:
        logging.info("Пользователь отменил выбор базы данных прошлого тура или возникла ошибка")
        return
    
    # Информация о типе выбранной базы данных
    is_mdb = os.path.isfile(old_db_path) and old_db_path.lower().endswith('.mdb')
    db_type = "MDB" if is_mdb else "GDB"
    logging.info("Выбрана база данных прошлого тура: {} (тип: {})".format(old_db_path, db_type))
    
    # Продолжаем с выбором значения
    root = tk.Tk()
    value_selector = ValueSelector(root, gdb_path)
    
    # Делаем диалог модальным
    root.grab_set()
    root.focus_set()
    root.wait_window()
    
    selected_value = value_selector.selected_value
    shortened_name = value_selector.shortened_name
    
    if not selected_value or not shortened_name:
        logging.info("Пользователь отменил выбор значения или не указал сокращение")
        return
    
    # Находим классы надписей в базе данных прошлого тура
    label_processor = LabelClassProcessor(old_db_path, shortened_name)
    success, message = label_processor.find_label_classes()
    
    if not success:
        # Проверяем тип проблемы
        has_duplicates = "дубликаты" in message.lower()
        has_no_numbers = "не найдены цифры" in message.lower()
        
        # Показываем диалог с предупреждением
        action = label_processor.show_verification_dialog(message, has_duplicates, has_no_numbers)
        
        if action == "retry":
            # Повторяем поиск (вероятно, пользователь внес изменения)
            success, message = label_processor.find_label_classes()
            if not success:
                # Если снова ошибка, показываем сообщение и завершаем работу
                logging.error("Не удалось найти корректные классы надписей после повторной попытки")
                messagebox.showerror("Ошибка", "Не удалось найти корректные классы надписей после повторной попытки")
                return
        elif action == "finish":
            # Пользователь выбрал завершение работы
            logging.info("Пользователь решил завершить работу скрипта")
            return
        # Если action == "continue", продолжаем выполнение
    
    # Обработка данных с основной базой геоданных
    processor = DataProcessor(gdb_path, selected_value, shortened_name)
    success = processor.process_data()
    
    if not success:
        logging.error("Обработка данных завершилась с ошибкой")
        return
    
    # После выполнения основной обработки данных, выполняем операцию идентичности
    # Формируем путь к слою Land_"Сокр"
    try:
        # Находим класс объектов для идентичности (Land_"Сокр")
        land_clip_name = "Land_{}".format(shortened_name)
        
        # Определяем полный путь к созданному набору данных
        cleaned_usname = selected_value.replace(" ", "_").replace("\"", "").replace("'", "")
        new_dataset_name = "_{}".format(cleaned_usname)
        dataset_path = os.path.join(gdb_path, new_dataset_name)
        
        # Полный путь к слою Land_"Сокр"
        land_clip_path = os.path.join(dataset_path, land_clip_name)
        
        if not arcpy.Exists(land_clip_path):
            logging.error("Не найден класс объектов '{}'".format(land_clip_name))
            messagebox.showerror("Ошибка", "Не найден класс объектов '{}'".format(land_clip_name))
            return
        
        # Выполняем операцию идентичности с Land_"Сокр" (не с Land_"Сокр"_контур)
        success = label_processor.process_identity(land_clip_path)
        
        if success:
            logging.info("Операция идентичности успешно выполнена")
            messagebox.showinfo("Успех", "Операция идентичности успешно выполнена")
            
            # После успешного выполнения всех операций сохраняем MXD файл
            # ВАЖНО: Убраны промежуточные сохранения MXD из всех функций
            # Теперь файл карты сохраняется только один раз - здесь, в конце работы скрипта
            try:
                if ARCPY_AVAILABLE:
                    # Определяем продукт ArcGIS
                    product_info = arcpy.ProductInfo()
                    
                    # Формируем имя для MXD файла на основе выбранного значения
                    mxd_name = "_{}".format(cleaned_usname)
                    mxd_path = os.path.join(os.path.dirname(gdb_path), "{}.mxd".format(mxd_name))
                    
                    # Для ArcMap
                    if product_info in ["ArcView", "ArcEditor", "ArcInfo", "Desktop"]:
                        import arcpy.mapping as mapping
                        try:
                            # Получаем текущий документ карты
                            mxd = mapping.MapDocument("CURRENT")
                            
                            # Сохраняем копию карты
                            mxd.saveACopy(mxd_path)
                            logging.info("Финальная копия карты сохранена: {}".format(mxd_path))
                            
                            # Обновляем вид
                            arcpy.RefreshTOC()
                            arcpy.RefreshActiveView()
                        except Exception as mxd_err:
                            logging.error("Ошибка при сохранении финальной копии карты: {}".format(str(mxd_err)))
                    
                    # Для ArcGIS Pro
                    elif product_info in ["ArcGISPro", "Pro"]:
                        import arcpy.mp as mp
                        try:
                            # Получаем текущий проект
                            aprx = mp.ArcGISProject("CURRENT")
                            
                            # Формируем имя для APRX файла
                            aprx_path = os.path.join(os.path.dirname(gdb_path), "{}.aprx".format(mxd_name))
                            
                            # Сохраняем копию проекта
                            aprx.saveACopy(aprx_path)
                            logging.info("Финальная копия проекта сохранена: {}".format(aprx_path))
                        except Exception as aprx_err:
                            logging.error("Ошибка при сохранении финальной копии проекта: {}".format(str(aprx_err)))
                    
                    # Если не удалось определить продукт
                    else:
                        logging.warning("Не удалось определить продукт ArcGIS для финального сохранения документа карты")
            except Exception as final_save_err:
                logging.error("Общая ошибка при финальном сохранении документа карты: {}".format(str(final_save_err)))
        else:
            logging.error("Операция идентичности завершилась с ошибкой")
            messagebox.showerror("Ошибка", "Операция идентичности завершилась с ошибкой")
        
    except Exception as e:
        error_message = "Ошибка при выполнении операции идентичности: {}".format(str(e))
        logging.error(error_message)
        log_exception(e, "Ошибка при выполнении операции идентичности")
        messagebox.showerror("Ошибка", error_message)

if __name__ == "__main__":
    main()