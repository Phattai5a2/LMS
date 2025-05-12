# -*- coding: utf-8 -*-
"""
Created on Sun May 11 23:26:42 2025

@author: PC
"""

import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Moodle User & Course CSV Generator", layout="centered")

st.title("📥 Quản lý lớp học trên Moodle")

st.markdown(
    """
    <style>
    [data-testid="stToolbar"] {
            visibility: hidden;
    }
    </style>
    """,
    unsafe_allow_html=True
)

def extract_course_info(df_full):
    course_info_line = df_full.iloc[4, 4] if not pd.isna(df_full.iloc[4, 4]) else ""
    class_line = df_full.iloc[5, 1] if not pd.isna(df_full.iloc[5, 1]) else ""
    course_code_match = re.search(r'\[(.*?)\]', course_info_line)
    course_code = course_code_match.group(1) if course_code_match else ''
    class_match = re.search(r'Lớp:\s*(.*)', class_line)
    class_name = class_match.group(1).strip() if class_match else ''
    return f"{course_code}_{class_name}", course_info_line.split(":")[-1].strip()

def filter_valid_students(df):
    df['MSSV'] = df['MSSV'].astype(str)
    return df[df['MSSV'].str.fullmatch(r'\d{8,}')].copy()

def split_name(full_name):
    parts = full_name.strip().split()
    return (' '.join(parts[:-1]), parts[-1]) if len(parts) > 1 else ('', parts[0])

def process_excel(uploaded_file):
    df_full = pd.read_excel(uploaded_file, sheet_name=0, header=None, engine='xlrd')
    course_identifier, course_fullname_base = extract_course_info(df_full)
    df_raw = pd.read_excel(uploaded_file, header=None, skiprows=13, engine='xlrd')
    cols = ['STT', 'MSSV', 'Ho', 'Ten', 'GioiTinh', 'NgaySinh', 'Lop'] + [f'col{i}' for i in range(7, df_raw.shape[1])]
    df_raw.columns = cols[:df_raw.shape[1]]
    df_valid = filter_valid_students(df_raw[['MSSV', 'Ho', 'Ten']].copy())
    df_valid['Email'] = df_valid['MSSV'].astype(str) + '@ntt.edu.vn'
    students = []
    for _, row in df_valid.iterrows():
        ho_lot, ten = split_name(row['Ho'] + " " + row['Ten'])
        students.append({'username': row['MSSV'], 'password': row['MSSV'],
                         'firstname': ho_lot, 'lastname': ten,
                         'email': row['Email'], 'course1': course_identifier})
    return students, course_identifier, course_fullname_base

tab1, tab2 = st.tabs(["📄 Một File", "📂 Nhiều File"])

with tab1:
    st.header("📄 Xử lý Một File Excel")
    uploaded_file = st.file_uploader("Chọn file Excel", type=["xls", "xlsx"])
    username_gv = st.text_input("👨‍🏫 Username Giảng Viên:")
    fullname_gv = st.text_input("👨‍🏫 Họ và Tên Giảng Viên:")
    email_gv = st.text_input("📧 Email Giảng Viên:")
    category_id = st.text_input("📂 Category ID:", value="14")

    if uploaded_file and st.button("🚀 Xử lý Một File"):
        students, course_code, course_name = process_excel(uploaded_file)
        gv_ho_lot, gv_ten = split_name(fullname_gv)
        all_users = [{'username': username_gv, 'password': 'Kcntt@2022m',
                      'firstname': gv_ho_lot, 'lastname': gv_ten,
                      'email': email_gv, 'course1': course_code}] + students
        df_users = pd.DataFrame(all_users)

        parts = course_name.split("] - ")
        course_code = parts[0].strip()
        clean_line = course_code.lstrip("[")
        course_name = parts[1]  # "Chuyên đề chuyên sâu Kỹ thuật CNTT 2 (22DTH1D)"

        # Ghép lại thành chuỗi mong muốn
        course_full = f"{clean_line}_{course_name}"
        
        df_course = pd.DataFrame([{'shortname': course_code,
                                   'fullname': f"{course_full}_GV: {fullname_gv}",
                                   'category': category_id}])
        st.dataframe(df_users)
        st.download_button("⬇️ Tải file Người Dùng", df_users.to_csv(index=False).encode('utf-8-sig'),
                           file_name="moodle_user_upload.csv", mime="text/csv")
        st.dataframe(df_course)
        st.download_button("⬇️ Tải file Lớp Học", df_course.to_csv(index=False).encode('utf-8-sig'),
                           file_name="moodle_course_upload.csv", mime="text/csv")

with tab2:
    st.header("📂 Xử lý Nhiều File Excel")
    uploaded_files = st.file_uploader("Chọn nhiều file Excel", type=["xls", "xlsx"], accept_multiple_files=True)
    username_gv_multi = st.text_input("👨‍🏫 Username Giảng Viên cho Tất Cả:")
    fullname_gv_multi = st.text_input("👨‍🏫 Họ và Tên Giảng Viên cho Tất Cả:")
    email_gv_multi = st.text_input("📧 Email Giảng Viên cho Tất Cả:")
    category_id_multi = st.text_input("📂 Category ID cho Tất Cả:", value="14")

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

        st.download_button("⬇️ Tải file Lớp Học (Tất Cả)", df_courses_all.to_csv(index=False).encode('utf-8-sig'),
                           file_name="moodle_course_upload_all.csv", mime="text/csv")
