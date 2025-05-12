# -*- coding: utf-8 -*-

"""
Created on Sun May 11 23:26:42 2025

@author: PC
"""

import streamlit as st
import pandas as pd
import re

st.set\_page\_config(page\_title="Moodle User & Course CSV Generator", layout="centered")

st.title("📥 Quản lý lớp học trên Moodle")

st.markdown(
""" <style>
\[data-testid="stToolbar"] {
visibility: hidden;
} </style>
""",
unsafe\_allow\_html=True
)

def extract\_course\_info(df\_full):
course\_info\_line = df\_full.iloc\[4, 4] if not pd.isna(df\_full.iloc\[4, 4]) else ""
class\_line = df\_full.iloc\[5, 1] if not pd.isna(df\_full.iloc\[5, 1]) else ""
course\_code\_match = re.search(r'$(.*?)$', course\_info\_line)
course\_code = course\_code\_match.group(1) if course\_code\_match else ''
class\_match = re.search(r'Lớp:\s\*(.\*)', class\_line)
class\_name = class\_match.group(1).strip() if class\_match else ''
return f"{course\_code}\_{class\_name}", course\_info\_line.split(":")\[-1].strip()

def filter\_valid\_students(df):
df\['MSSV'] = df\['MSSV'].astype(str)
return df\[df\['MSSV'].str.fullmatch(r'\d{8,}')].copy()

def split\_name(full\_name):
parts = full\_name.strip().split()
return (' '.join(parts\[:-1]), parts\[-1]) if len(parts) > 1 else ('', parts\[0])

def process\_excel(uploaded\_file):
df\_full = pd.read\_excel(uploaded\_file, sheet\_name=0, header=None, engine='xlrd')
course\_identifier, course\_fullname\_base = extract\_course\_info(df\_full)
df\_raw = pd.read\_excel(uploaded\_file, header=None, skiprows=13, engine='xlrd')
cols = \['STT', 'MSSV', 'Ho', 'Ten', 'GioiTinh', 'NgaySinh', 'Lop'] + \[f'col{i}' for i in range(7, df\_raw\.shape\[1])]
df\_raw\.columns = cols\[:df\_raw\.shape\[1]]
df\_valid = filter\_valid\_students(df\_raw\[\['MSSV', 'Ho', 'Ten']].copy())
df\_valid\['Email'] = df\_valid\['MSSV'].astype(str) + '@ntt.edu.vn'
students = \[]
for \_, row in df\_valid.iterrows():
ho\_lot, ten = split\_name(row\['Ho'] + " " + row\['Ten'])
students.append({'username': row\['MSSV'], 'password': row\['MSSV'],
'firstname': ho\_lot, 'lastname': ten,
'email': row\['Email'], 'course1': course\_identifier})
return students, course\_identifier, course\_fullname\_base

tab1, tab2 = st.tabs(\["📄 Một File", "📂 Nhiều File"])

with tab1:
st.header("📄 Xử lý Một File Excel")
uploaded\_file = st.file\_uploader("Chọn file Excel", type=\["xls", "xlsx"])
username\_gv = st.text\_input("👨‍🏫 Username Giảng Viên:")
fullname\_gv = st.text\_input("👨‍🏫 Họ và Tên Giảng Viên:")
email\_gv = st.text\_input("📧 Email Giảng Viên:")
category\_id = st.text\_input("📂 Category ID:", value="15")

```
if uploaded_file and st.button("🚀 Xử lý Một File"):
    students, course_code, course_name = process_excel(uploaded_file)
    gv_ho_lot, gv_ten = split_name(fullname_gv)
    all_users = [{'username': username_gv, 'password': 'Kcntt@123456',
                  'firstname': gv_ho_lot, 'lastname': gv_ten,
                  'email': email_gv, 'course1': course_code}] + students
    df_users = pd.DataFrame(all_users)
    df_course = pd.DataFrame([{'shortname': course_code,
                               'fullname': f"{course_name}_GV: {fullname_gv}",
                               'categoryID': category_id}])
    st.dataframe(df_users)
    st.download_button("⬇️ Tải file Người Dùng", df_users.to_csv(index=False).encode('utf-8-sig'),
                       file_name="moodle_user_upload.csv", mime="text/csv")
    st.dataframe(df_course)
    st.download_button("⬇️ Tải file Lớp Học", df_course.to_csv(index=False).encode('utf-8-sig'),
                       file_name="moodle_course_upload.csv", mime="text/csv")
```

with tab2:
st.header("📂 Xử lý Nhiều File Excel")
uploaded\_files = st.file\_uploader("Chọn nhiều file Excel", type=\["xls", "xlsx"], accept\_multiple\_files=True)
username\_gv\_multi = st.text\_input("👨‍🏫 Username Giảng Viên cho Tất Cả:")
fullname\_gv\_multi = st.text\_input("👨‍🏫 Họ và Tên Giảng Viên cho Tất Cả:")
email\_gv\_multi = st.text\_input("📧 Email Giảng Viên cho Tất Cả:")
category\_id\_multi = st.text\_input("📂 Category ID cho Tất Cả:", value="14")

```
if uploaded_files and st.button("🚀 Xử lý Nhiều File"):
    all_user_records, all_course_records = [], []
    gv_ho_lot_multi, gv_ten_multi = split_name(fullname_gv_multi)

    for file in uploaded_files:
        students, course_code, course_name = process_excel(file)
        all_user_records.append({'username': username_gv_multi, 'password': 'Kcntt@2022',
                                 'firstname': gv_ho_lot_multi, 'lastname': gv_ten_multi,
                                 'email': email_gv_multi, 'course1': course_code})
        all_user_records.extend(students)
        all_course_records.append({'shortname': course_code,
                                   'fullname': f"{course_name}_GV: {fullname_gv_multi}",
                                   'category': category_id_multi})

    df_users_all = pd.DataFrame(all_user_records)
    df_courses_all = pd.DataFrame(all_course_records)

    st.dataframe(df_users_all)
    st.download_button("⬇️ Tải file Người Dùng (Tất Cả)", df_users_all.to_csv(index=False).encode('utf-8-sig'),
                       file_name="moodle_user_upload_all.csv", mime="text/csv")
    st.dataframe(df_courses_all)
    st.download_button("⬇️ Tải file Lớp Học (Tất Cả)", df_courses_all.to_csv(index=False).encode('utf-8-sig'),
                       file_name="moodle_course_upload_all.csv", mime="text/csv")
