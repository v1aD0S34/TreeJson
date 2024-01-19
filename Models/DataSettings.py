class ExcelSettings:
    def __init__(self, name, excel_sheet, tree_path, tag, unit, description, prefix, postfix, start_count):
        self.Name = name  # Наименование структуры
        self.ExcelSheet = excel_sheet  # Название листа Экселя
        self.TreePath = tree_path  # Имя папки в трендах
        self.Tag = tag  # Столбец Экселя
        self.Unit = unit  # Столбец Экселя
        self.Description = description  # Столбец Экселя
        self.Prefix = prefix  # Начало тега
        self.Postfix = postfix  # Конец тега
        self.StartCount = start_count  # С какой строки начинаются данные
