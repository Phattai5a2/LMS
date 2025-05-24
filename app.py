# -*- coding: utf-8 -*-
"""
Created on Sun May 11 23:26:42 2025
@author: PC
"""

import streamlit as st
import pandas as pd
import re
import requests
import json
import time
import magic  # Thư viện để kiểm tra định dạng file

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

# Hàm kiểm tra định dạng file
def check_file_format(file):
    try:
        file.seek(0)  # Đặt con trỏ về đầu file
        mime = magic.Magic(mime=True)
        file_type = mime.from_buffer(file.read(1024))
        file.seek(0)  # Đặt lại con trỏ
        return file_type in [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',  # .xlsx
            'application/vnd.ms-excel',  # .xls
        ]
    except Exception:
        return False

# Hàm xử lý Excel đã sửa
def process_excel(uploaded_file):
    if not check_file_format(uploaded_file):
        st.error("File không phải định dạng Excel hợp lệ (.xlsx hoặc .xls). Vui lòng kiểm tra và tải lại file.")
        return [], "", ""
    
    try:
        # Thử đọc với openpyxl trước
        try:
            df_full = pd.read_excel(uploaded_file, sheet_name=0, header=None, engine='openpyxl')
        except Exception as e_openpyxl:
            # Nếu openpyxl thất bại, thử xlrd cho .xls
            try:
                uploaded_file.seek(0)  # Đặt lại con trỏ file
                df_full = pd.read_excel(uploaded_file, sheet_name=0, header=None, engine='xlrd')
            except Exception as e_xlrd:
                st.error(f"Không thể đọc file Excel: {str(e_openpyxl)} (openpyxl) hoặc {str(e_xlrd)} (xlrd). Vui lòng kiểm tra file.")
                return [], "", ""

        if df_full.size > 1_000_000:
            st.error("File Excel quá lớn, vui lòng giảm số lượng dữ liệu.")
            return [], "", ""

        course_identifier, course_fullname_base = extract_course_info(df_full)
        if not course_identifier:
            st.error("Không tìm thấy thông tin khóa học hợp lệ trong file. Kiểm tra ô [4,4] và [5,1].")
            return [], "", ""

        # Thử đọc dữ liệu sinh viên
        try:
            df_raw = pd.read_excel(uploaded_file, header=None, skiprows=13, engine='openpyxl')
        except Exception:
            uploaded_file.seek(0)
            df_raw = pd.read_excel(uploaded_file, header=None, skiprows=13, engine='xlrd')

        cols = ['STT', 'MSSV', 'Ho', 'Ten', 'GioiTinh', 'NgaySinh', 'Lop'] + [f'col{i}' for i in range(7, df_raw.shape[1])]
        df_raw.columns = cols[:df_raw.shape[1]]

        df_valid = filter_valid_students(df_raw[['MSSV', 'Ho', 'Ten', 'NgaySinh']].copy())
        if df_valid.empty:
            st.error("Không tìm thấy sinh viên hợp lệ trong file (MSSV phải có ít nhất 8 chữ số).")
            return [], "", ""

        df_valid['Email'] = df_valid['MSSV'].astype(str) + '@ntt.edu.vn'

        students = []
        for _, row in df_valid.iterrows():
            ho_lot, ten = split_name(row['Ho'] + " " + row['Ten'])
            try:
                dob = pd.to_datetime(row['NgaySinh'], errors='coerce')
                dob_str = dob.strftime('%d%m%Y') if not pd.isna(dob) else '01011990'
            except:
                dob_str = '01011990'

            password = f"Kcntt@{dob_str}"
            students.append({
                'username': row['MSSV'],
                'password': password,
                'firstname': ho_lot,
                'lastname': ten,
                'email': row['Email'],
                'course1': course_identifier
            })

        return students, course_identifier, course_fullname_base
    except Exception as e:
        st.error(f"Lỗi khi xử lý file Excel: {str(e)}. Vui lòng kiểm tra định dạng file và thử lại.")
        return [], "", ""

# Giao diện Streamlit (chỉ hiển thị tab Một File để ngắn gọn)
tab1, tab2 = st.tabs(["📄 Một File", "📂 Nhiều File"])

with tab1:
    st.header("📄 Xử lý Một File Excel")
    moodle_url = st.text_input("🌐 URL Moodle (VD: https://your-moodle-site.com):")
    moodle_token = st.text_input("🔑 API Token:", type="password")
    
    if moodle_url and moodle_token and st.button("🔍 Kiểm tra Token"):
        if validate_token(moodle_url, moodle_token):
            st.success("Token hợp lệ!")
        else:
            st.error("Token không hợp lệ hoặc URL Moodle sai.")

    uploaded_file = st.file_uploader("Chọn file Excel", type=["xls", "xlsx"])
    username_gv = st.text_input("👨‍🏫 Username Giảng Viên:")
    fullname_gv = st.text_input("👨‍🏫 Họ và Tên Giảng Viên:")
    category_id = st.text_input("📂 Category ID:", value="15")
    role_gv = st.selectbox("👤 Vai trò Giảng Viên:", [("Giảng viên", 3), ("Sinh viên", 5)], index=0)
    role_sv = st.selectbox("👤 Vai trò Sinh Viên:", [("Sinh viên", 5), ("Giảng viên", 3)], index=0)

    if st.button("🚀 Xử lý và Cập nhật qua API"):
        if not (uploaded_file and username_gv and fullname_gv and moodle_url and moodle_token):
            st.error("Vui lòng cung cấp đầy đủ file, thông tin giảng viên, URL Moodle và API Token.")
        else:
            with st.spinner("Đang xử lý..."):
                try:
                    students, course_code, course_name = process_excel(uploaded_file)
                    if not students:
                        st.error("Không có dữ liệu hợp lệ để xử lý.")
                        st.stop()
                    
                    course_id = create_or_update_course(course_code, f"{course_name}_GV: {fullname_gv}", 
                                                      category_id, moodle_url, moodle_token)
                    if not course_id:
                        st.error("Không thể tạo hoặc cập nhật khóa học.")
                        st.stop()
                    
                    gv_ho_lot, gv_ten = split_name(fullname_gv)
                    teacher = [{
                        'username': username_gv,
                        'password': 'Kcntt@2xxx',
                        'firstname': gv_ho_lot,
                        'lastname': gv_ten,
                        'email': f"{username_gv}@ntt.edu.vn",
                        'course1': course_code
                    }]
                    enroll_users(teacher, course_id, role_gv[1], moodle_url, moodle_token)
                    enroll_users(students, course_id, role_sv[1], moodle_url, moodle_token)

                    all_users = teacher + students
                    df_users = pd.DataFrame(all_users)
                    df_course = pd.DataFrame([{
                        'shortname': course_code,
                        'fullname': f"{course_name}_GV: {fullname_gv}",
                        'category': category_id
                    }])
                    st.dataframe(df_users.head(10))
                    st.download_button(
                        "⬇️ Tải file Người Dùng",
                        df_users.to_csv(index=False).encode('utf-8-sig'),
                        file_name="moodle_user_upload.csv",
                        mime="text/csv"
                    )
                    st.dataframe(df_course)
                    st.download_button(
                        "⬇️ Tải file Lớp Học",
                        df_course.to_csv(index=False).encode('utf-8-sig'),
                        file_name="moodle_course_upload.csv",
                        mime="text/csv"
                    )
                except Exception as e:
                    st.error(f"Lỗi xử lý: {str(e)}")

# Tab Nhiều File (giữ tương tự, chỉ cập nhật gọi process_excel)
with tab2:
    st.header("📂 Xử lý Nhiều File Excel")
    moodle_url_multi = st.text_input("🌐 URL Moodle (Nhiều File):")
    moodle_token_multi = st.text_input("🔑 API Token (Nhiều File):", type="password")
    
    if moodle_url_multi and moodle_token_multi and st.button("🔍 Kiểm tra Token (Nhiều File)"):
        if validate_token(moodle_url_multi, moodle_token_multi):
            st.success("Token hợp lệ!")
        else:
            st.error("Token không hợp lệ hoặc URL Moodle sai.")

    uploaded_files = st.file_uploader("Chọn nhiều file Excel", type=["xls", "xlsx"], accept_multiple_files=True)
    username_gv_multi = st.text_input("👨‍🏫 Username Giảng Viên cho Tất Cả:")
    fullname_gv_multi = st.text_input("👨‍🏫 Họ và Tên Giảng Viên cho Tất Cả:")
    category_id_multi = st.text_input("📂 Category ID cho Tất Cả:", value="15")
    role_gv_multi = st.selectbox("👤 Vai trò Giảng Viên (Nhiều File):", [("Giảng viên", 3), ("Sinh viên", 5)], index=0)
    role_sv_multi = st.selectbox("👤 Vai trò Sinh Viên (Nhiều File):", [("Sinh viên", 5), ("Giảng viên", 3)], index=0)

    if st.button("🚀 Xử lý và Cập nhật qua API (Nhiều File)"):
        if not (uploaded_files and username_gv_multi and fullname_gv_multi and moodle_url_multi and moodle_token_multi):
            st.error("Vui lòng cung cấp đầy đủ file, thông tin giảng viên, URL Moodle và API Token.")
        else:
            with st.spinner("Đang xử lý các file..."):
                all_user_records, all_course_records = [], []
                gv_ho_lot_multi, gv_ten_multi = split_name(fullname_gv_multi)
                progress_bar = st.progress(0)
                total_files = len(uploaded_files)

                for i, file in enumerate(uploaded_files):
                    try:
                        students, course_code, course_name = process_excel(file)
                        if not students:
                            st.warning(f"File {file.name} không có sinh viên hợp lệ.")
                            continue
                        
                        course_id = create_or_update_course(course_code, f"{course_name}_GV: {fullname_gv_multi}", 
                                                           category_id_multi, moodle_url_multi, moodle_token_multi)
                        if not course_id:
                            st.error(f"Không thể tạo/cập nhật khóa học cho file {file.name}.")
                            continue
                        
                        teacher = [{
                            'username': username_gv_multi,
                            'password': 'Kcntt@2xxx',
                            'firstname': gv_ho_lot_multi,
                            'lastname': gv_ten_multi,
                            'email': f"{username_gv_multi}@ntt.edu.vn",
                            'course1': course_code
                        }]
                        enroll_users(teacher, course_id, role_gv_multi[1], moodle_url_multi, moodle_token_multi)
                        enroll_users(students, course_id, role_sv_multi[1], moodle_url_multi, moodle_token_multi)

                        all_user_records.extend(teacher + students)
                        all_course_records.append({
                            'shortname': course_code,
                            'fullname': f"{course_name}_GV: {fullname_gv_multi}",
                            'category': category_id_multi
                        })

                        progress_bar.progress((i + 1) / total_files)
                    except Exception as e:
                        st.error(f"Lỗi khi xử lý file {file.name}: {str(e)}")

                if all_user_records:
                    df_users_all = pd.DataFrame(all_user_records)
                    df_courses_all = pd.DataFrame(all_course_records)
                    st.dataframe(df_users_all.head(10))
                    st.download_button(
                        "⬇️ Tải file Người Dùng (Tất Cả)",
                        df_users_all.to_csv(index=False).encode('utf-8-sig'),
                        file_name="moodle_user_upload_all.csv",
                        mime="text/csv"
                    )
                    st.dataframe(df_courses_all)
                    st.download_button(
                        "⬇️ Tải file Lớp Học (Tất Cả)",
                        df_courses_all.to_csv(index=False).encode('utf-8-sig'),
                        file_name="moodle_course_upload_all.csv",
                        mime="text/csv"
                    )
                else:
                    st.error("Không có dữ liệu hợp lệ để tạo CSV.")
