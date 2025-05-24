# -*- coding: utf-8 -*-
"""
Created on Sun May 11 23:26:42 2025
@author: PC
"""

import streamlit as st
import pandas as pd
import re
import requests
import time

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

# Hàm xử lý Excel đã sửa
def process_excel(uploaded_file):
    # Kiểm tra phần mở rộng file
    if not uploaded_file.name.lower().endswith(('.xls', '.xlsx')):
        st.error("File không phải định dạng Excel hợp lệ (.xls hoặc .xlsx). Vui lòng kiểm tra và tải lại file.")
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
                st.error(f"Không thể đọc file Excel: {str(e_openpyxl)} (openpyxl) hoặc {str(e_xlrd)} (xlrd). Vui lòng kiểm tra file có đúng định dạng .xls/.xlsx và không bị hỏng.")
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
            uploaded_file.seek(0)
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
        st.error(f"Lỗi khi xử lý file Excel: {str(e)}. Vui lòng kiểm tra file có đúng định dạng .xls/.xlsx và không bị hỏng.")
        return [], "", ""

# Các hàm hỗ trợ khác (giữ nguyên từ mã trước)
def moodle_api_call(function_name, params, moodle_url, token):
    try:
        url = f"{moodle_url}/webservice/rest/server.php"
        params.update({
            'wstoken': token,
            'wsfunction': function_name,
            'moodlewsrestformat': 'json'
        })
        response = requests.get(url, params=params, timeout=10)
        if response.status_code == 200:
            return response.json()
        else:
            st.error(f"Lỗi khi gọi API {function_name}: {response.status_code} - {response.text}")
            return None
    except requests.exceptions.RequestException as e:
        st.error(f"Lỗi kết nối API {function_name}: {e}")
        return None

def validate_token(moodle_url, token):
    params = {'criteria[0][key]': 'username', 'criteria[0][value]': 'admin'}
    result = moodle_api_call('core_user_get_users', params, moodle_url, token)
    return result is not None

def create_or_update_course(course_code, course_name, category_id, moodle_url, token):
    params = {'courses[0][shortname]': course_code}
    existing_course = moodle_api_call('core_course_get_courses_by_field', params, moodle_url, token)
    
    if existing_course and 'courses' in existing_course and existing_course['courses']:
        course_id = existing_course['courses'][0]['id']
        params = {
            'courses[0][id]': course_id,
            'courses[0][shortname]': course_code,
            'courses[0][fullname]': course_name,
            'courses[0][categoryid]': category_id
        }
        result = moodle_api_call('core_course_update_courses', params, moodle_url, token)
        if result:
            st.success(f"Đã cập nhật khóa học: {course_name} (ID: {course_id})")
            return course_id
    else:
        params = {
            'courses[0][shortname]': course_code,
            'courses[0][fullname]': course_name,
            'courses[0][categoryid]': category_id
        }
        result = moodle_api_call('core_course_create_courses', params, moodle_url, token)
        if result and 'courses' in result:
            course_id = result['courses'][0]['id']
            st.success(f"Đã tạo khóa học mới: {course_name} (ID: {course_id})")
            return course_id
    return None

def create_user(user, moodle_url, token):
    params = {'criteria[0][key]': 'username', 'criteria[0][value]': user['username']}
    existing_user = moodle_api_call('core_user_get_users', params, moodle_url, token)
    
    if existing_user and 'users' in existing_user and existing_user['users']:
        return existing_user['users'][0]['id']
    else:
        params = {
            'users[0][username]': user['username'],
            'users[0][password]': user['password'],
            'users[0][firstname]': user['firstname'],
            'users[0][lastname]': user['lastname'],
            'users[0][email]': user['email']
        }
        result = moodle_api_call('core_user_create_users', params, moodle_url, token)
        if result and 'users' in result:
            st.success(f"Đã tạo người dùng: {user['username']}")
            return result['users'][0]['id']
        else:
            st.error(f"Lỗi khi tạo người dùng: {user['username']}")
            return None

def enroll_users(users, course_id, role_id, moodle_url, token, batch_size=50):
    success_count = 0
    for i in range(0, len(users), batch_size):
        batch = users[i:i + batch_size]
        params = {}
        for j, user in enumerate(batch):
            user_id = create_user(user, moodle_url, token)
            if user_id:
                params[f'enrolments[{j}][roleid]'] = role_id
                params[f'enrolments[{j}][userid]'] = user_id
                params[f'enrolments[{j}][courseid]'] = course_id
        if params:
            result = moodle_api_call('enrol_manual_enrol_users', params, moodle_url, token)
            if result is None:
                success_count += len(params) // 3
                st.success(f"Đã ghi danh {len(params) // 3} người dùng trong batch.")
            else:
                st.error("Lỗi khi ghi danh batch người dùng.")
        time.sleep(0.5)
    return success_count

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

# Giao diện Streamlit
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
