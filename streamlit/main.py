#Библиотеки
import streamlit as st
import tabula
import aspose.words as aw
import pandas as pd

#конвертатор из exel в data freim
def convert_xlsx(file_name):
    return pd.read_excel(file_name)

#конвертатор из pdf в data freim
def convert_pdf(file_name):
    tabula.convert_into(file_name, f"{file_name.split('.')[0]}.csv", output_format="csv", pages='all')
    return pd.read_csv(f"{file_name.split('.')[0]}.csv")

#конвертатор из docx в data freim
def convert_docx(file_name):
    doc = aw.Document(file_name)
    doc.save(f"{file_name.split('.')[0]}.pdf")
    return convert_pdf(f"{file_name.split('.')[0]}.pdf")

#Шапка сайта
st.title("Отток клиентов")
st.sidebar.title("О нас")
st.sidebar.info(
    """
    Мы команда что во время первого чек-поинта смеялись над словом банан.
    """
)
#Загрузка файлов для в которых нахоятся "Интересы" клиента
st.write("Интересы")
file = st.file_uploader("Загрузите ваши файлы", accept_multiple_files=True, type=['xlsx', 'xls', 'docx', 'pdf'], key='1')
for uploaded_file in file:
    file_contents = uploaded_file.read()
st.write("\n\n\n")
#Загрузка файлов для в которых нахоятся "Обращения" клиента
st.write("Обращение")
file1 = st.file_uploader("Загрузите ваши файлы", accept_multiple_files=True, type=['xlsx', 'xls', 'docx', 'pdf'], key='2')
for uploaded_file in file1:
    file_contents = uploaded_file.read()
st.write("\n\n\n")
#Загрузка файлов в которых нахоятся "Объемы перевозок" клиента
st.write("Объемы перевозок")
file2 = st.file_uploader("Загрузите ваши файлы", accept_multiple_files=True,type=['xlsx', 'xls', 'docx', 'pdf'], key='3')
for uploaded_file in file2:
    file_contents = uploaded_file.read()
st.write("\n \n \n")
#Загрузка файлов для файлов, в которых нахоятся "Выгрузко-маркетинговые списки" клиента
st.write("Выгрузко-маркетинговые списки")
file3 = st.file_uploader("Загрузите ваши файлы", accept_multiple_files=True,type=['xlsx', 'xls', 'docx', 'pdf'], key='4')
for uploaded_file in file3:
    file_contents = uploaded_file.read()
st.write("\n \n \n")

#Оценка от пользователей сайта
st.write("Оцените пожалуйсто наш веб сайт:")
marks = st.feedback("stars")
if marks == 4:
    st.write("Мы рады что вам понравился наш сайт")
elif marks and marks < 4:
    text = st.text_input("Напишите отзыв чтобы мы смогли исправить проблему")
    if text != '':
        st.title("Спасибо за отзыв!")