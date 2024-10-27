from tkinter import ttk
from tkinter import *
from tkinter import filedialog
import tkinter.font as font
import tabula
import aspose.words as aw
import numpy as np
import pandas as pd
from sklearn.preprocessing import LabelEncoder



files = {"Интересы": '',
         "Обращения": '',
         "Объемы перевозок": '',
         "Регионы": []
         }


# функция преобразования датасета интересов
def transformInterests(dataset):
    # удаление некоторых столбцов
    dataset = dataset.drop(
        ["Дата следующей активности", "Номер", "Ссылка (служебное поле для вывода на экран прочих реквизитов объекта)"],
        axis=1, errors="ignore")

    # замена столбцов категорий столбцом чисел
    for col in ["Тема", "Сценарий", "Состояние", "Подразделение", "Следующая активность", "Канал первичного интереса"]:
        dataset[col] = dataset[col].str.lower()
        dataset[col] = dataset[col].str.strip()

        le = LabelEncoder()
        mean_value = le.fit_transform(dataset[col].dropna()).mean()

        le.fit(dataset[col].dropna())
        dataset[col] = dataset[col].apply(lambda x: le.transform([x])[0] if pd.notna(x) else mean_value)
        dataset[col] = dataset[col].astype(np.int64)

    # замена неизвестных значений на среднее из всех известных и преобразование в int
    for col in ["Ожидаемая выручка", "Вероятность сделки, %"]:
        mean_value = dataset[col].dropna().mean()

        dataset[col] = dataset[col].apply(lambda x: mean_value if pd.isnull(x) else x)
        dataset[col] = dataset[col].astype(np.int64)

    # преобразование данных в тип даты и времени
    for col in ["Дата"]:
        dataset[col] = pd.to_datetime(dataset[col], format="mixed", dayfirst=True)

    # возврат итогового датасета
    return dataset


# функция преобразования датасета обращений
def transformAppeals(dataset):
    # удаление некоторых столбцов
    dataset = dataset.drop(["Тип обращения", "Номер", "Тема вопроса", "Количество доработок"], axis=1, errors="ignore")

    # замена столбцов категорий столбцом чисел
    for col in ["Тема", "Группа вопросов"]:
        dataset[col] = dataset[col].str.lower()
        dataset[col] = dataset[col].str.strip()

        le = LabelEncoder()
        mean_value = le.fit_transform(dataset[col].dropna()).mean()

        le.fit(dataset[col].dropna())
        dataset[col] = dataset[col].apply(lambda x: le.transform([x])[0] if pd.notna(x) else mean_value)
        dataset[col] = dataset[col].astype(np.int64)

    # преобразование данных в тип даты и времени
    for col in ["Дата"]:
        dataset[col] = pd.to_datetime(dataset[col], format="mixed", dayfirst=True)

    # возврат итогового датасета
    return dataset


# функция преобразования датасета маркетингового списка
def transformMarketingList(dataset):
    # замена столбцов категорий столбцом чисел
    for col in [
        "Размер компании.Наименование",
        "ОКВЭД2.Наименование", "Город фактический",
        "Город юридический", "Карточка клиента (внешний источник).Индекс платежной дисциплины Описание",
        "Карточка клиента (внешний источник).Индекс финансового риска Описание",
        "Госконтракты.Контракт", "Госконтракты.Тип контракта"
    ]:
        dataset[col] = dataset[col].str.lower()
        dataset[col] = dataset[col].str.strip()

        le = LabelEncoder()
        mean_value = le.fit_transform(dataset[col].dropna()).mean()

        le.fit(dataset[col].dropna())
        dataset[col] = dataset[col].apply(lambda x: le.transform([x])[0] if pd.notna(x) else mean_value)
        dataset[col] = dataset[col].astype(np.int64)

    # очищение ОКВЭД2.Код от точек
    dataset["ОКВЭД2.Код"] = dataset["ОКВЭД2.Код"].str.replace('.', '')
    dataset["ОКВЭД2.Код"] = dataset["ОКВЭД2.Код"].astype(np.float64)

    # замена неизвестных значений на среднее из всех известных и преобразование в int
    for col in [
        "Размер уставного капитала объявленный",
        "ОКВЭД2.Код", "Численность персонала по данным ФНС.Количество",
        "Карточка клиента (внешний источник).Индекс платежной дисциплины Значение",
        "Карточка клиента (внешний источник).Индекс финансового риска Значение"
    ]:
        mean_value = dataset[col].dropna().mean()

        dataset[col] = dataset[col].apply(lambda x: mean_value if pd.isnull(x) else x)
        dataset[col] = dataset[col].astype(np.int64)

    # замена значений "Да" на 1, "Нет" на 0
    for col in ["Находится в реестре МСП", "Грузоотправитель", "Грузополучатель"]:
        dataset[col] = dataset[col].apply(lambda x: x if pd.isnull(x) else (0 if x == "Нет" else 1))

    # замена значений ЕЛС действующего на условие его присутствия
    dataset["ЕЛС действующий"] = dataset["ЕЛС действующий"].astype("string")
    dataset["ЕЛС действующий"] = dataset["ЕЛС действующий"].apply(
        lambda x: (0 if (pd.isnull(x) or x.strip() == '') else 1))

    # возврат итогового датасета
    return dataset


def convert_xlsx(file_name):
    return pd.read_excel(file_name)


def convert_pdf(file_name):
    tabula.convert_into(file_name, f"{file_name.split('.')[0]}.csv", output_format="csv", pages='all')
    return pd.read_csv(f"{file_name.split('.')[0]}.csv")


def convert_docx(file_name):
    doc = aw.Document(file_name)

    # Save as PDF
    doc.save(f"{file_name.split('.')[0]}.pdf")
    return convert_pdf(f"{file_name.split('.')[0]}.pdf")


def convert(file_name):
    if file_name.split('.')[-1] in ("xlsx", "xls"):
        return convert_xlsx(file_name)
    elif file_name.split('.')[-1] == 'pdf':
        return convert_pdf(file_name)
    elif file_name.split('.')[-1] == 'docx':
        return convert_docx(file_name)


def close_escape(event=None):
    window.destroy()


def open_file1():
    """Открываем файл для редактирования"""
    global files
    filepath = filedialog.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=[("Лист Microsoft Excel 97–2003", "*.xls"),
                   ("Microsoft Edge PDF Document", "*.pdf"),
                   ("Документ Microsoft Word", "*.docx"),
                   ("Все файлы", "*.*")]
    )
    Label(text=filepath.split('/')[-1], font="calibre 15 bold", bg='#fff').place(relx=0.15, rely=0.37)
    files["Интересы"] = filepath


def open_file2():
    """Открываем файл для редактирования"""
    global files
    filepath = filedialog.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=[("Лист Microsoft Excel 97–2003", "*.xls"),
                   ("Microsoft Edge PDF Document", "*.pdf"),
                   ("Документ Microsoft Word", "*.docx"),
                   ("Все файлы", "*.*")]
    )
    Label(text=filepath.split('/')[-1], font="calibre 15 bold", bg='#fff').place(relx=0.15, rely=0.57)
    files['Обращения'] = filepath


def open_file3():
    global files
    """Открываем файл для редактирования"""
    filepath = filedialog.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=[("Лист Microsoft Excel 97–2003", "*.xls"),
                   ("Microsoft Edge PDF Document", "*.pdf"),
                   ("Документ Microsoft Word", "*.docx"),
                   ("Все файлы", "*.*")]
    )
    Label(text=filepath.split('/')[-1], font="calibre 15 bold", bg='#fff').place(relx=0.15, rely=0.77)
    files["Объемы перевозок"] = filepath


def open_file_mix():
    filetypes = [("Лист Microsoft Excel 97–2003", "*.xls"),
                ("Microsoft Edge PDF Document", "*.pdf"),
                ("Документ Microsoft Word", "*.docx"),
                ("Все файлы", "*.*")]

    filenames = filedialog.askopenfilenames(title='Открыть файлы',initialdir='/',filetypes=filetypes,)
    filenames = [_.split('/')[-1] for _ in filenames]
    file = StringVar(value=filenames)
    listbox = Listbox(listvariable=file, width=65, font="calibre 15 bold", bg='#fff')
    listbox.place(relx=0.45, rely=0.4)
    scrollbar = Scrollbar(orient="vertical", command=listbox.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    files["Регионы"] = filenames


def save_file():
    """Сохраняем текущий файл как новый файл."""
    filepath = filedialog.asksaveasfilename(
        defaultextension="csv",
        filetypes=[("CSV", "*.csv"), ("Все файлы", "*.*")],
    )
    if not filepath:
        return
    with open(filepath, "w") as output_file:
        text = txt_edit.get("1.0", END)
        output_file.write(text)


def upload_file():
    files["Интересы"] = transformInterests(convert(files["Интересы"]))
    files["Обращения"] = transformAppeals(convert(files["Обращения"]))
    for mlist in files["Регионы"]:
        mlist = transformMarketingList(convert(mlist))

window = Tk()
window.title("Отток клиентов")
window.attributes('-fullscreen', True)
window.bind("<Escape>", close_escape)
window.configure(bg='#fff')
txt_edit = Text(window)
Label(text="Маркетинговые списки", padx=20, pady=40, font="calibre 20 bold", bg='#fff').place(relx=0.44, rely=0.2)
Label(text="Расчет оттока клиентов для ОАО РЖД", padx=20, pady=40, font="calibre 40 bold", bg='#fff').pack()
Label(text="Интересы", padx=20, pady=40, font="calibre 20 bold", bg='#fff').place(relx=0.14, rely=0.2)
Label(text="Обращение", padx=20, pady=40, font="calibre 20 bold", bg='#fff').place(relx=0.14, rely=0.4)
Label(text="Объемы перевозок", padx=20, pady=40, font="calibre 20 bold", bg='#fff').place(relx=0.14, rely=0.6)
btn_open1 = Button(text="Открыть файл", command=open_file1, bg = '#ff2400', bd = 0, fg = '#fff', activebackground = '#fff', activeforeground = '#ff2400',cursor = 'hand2')
btn_open2 = Button(text="Открыть файл", command=open_file2, bg = '#ff2400', bd = 0, fg = '#fff', activebackground = '#fff', activeforeground = '#ff2400',cursor = 'hand2')
btn_open3 = Button(text="Открыть файл", command=open_file3, bg = '#ff2400', bd = 0, fg = '#fff', activebackground = '#fff', activeforeground = '#ff2400',cursor = 'hand2')
btn_open_mix = Button(text="Открыть файл", command=open_file_mix, bg = '#ff2400', bd = 0, fg = '#fff', activebackground = '#fff', activeforeground = '#ff2400',cursor = 'hand2')
upload_btn = Button(text="Загрузить данные", command=upload_file, bg = '#000', bd = 0, fg = '#fff', activebackground = '#fff', activeforeground = '#000',cursor = 'hand2')
upload_btn['font'] = font.Font(family='Helvetica')
btn_open1['font'] = font.Font(family='Helvetica')
btn_open2['font'] = font.Font(family='Helvetica')
btn_open3['font'] = font.Font(family='Helvetica')
btn_open_mix['font'] = font.Font(family='Helvetica')
btn_open1.place(relx=0.15, rely=0.30,)
btn_open2.place(relx=0.15, rely=0.50,)
btn_open3.place(relx=0.15, rely=0.70,)
btn_open_mix.place(relx=0.45, rely=0.30)
upload_btn.place(relx=0.3, rely=0.80)
style = ttk.Style()
style.configure(["TRadiobutton",], font=('Helvetica', 14))
style.configure(["Button",], font=('Helvetica', 14))
window.mainloop()