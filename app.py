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
import logging

logging.basicConfig(filename='app.log', level=logging.DEBUG)

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

# Hàm kiểm tra kết nối server đã sửa (dùng HTTP thay vì ping)
def check_server_connectivity(moodle_url):
    try:
        if not moodle_url.startswith(('http://', 'https://')):
            st.error("URL Moodle phải bắt đầu bằng http:// hoặc https://.")
            logging.error("URL không hợp lệ: %s", moodle_url)
            return False
        
        # Thử truy cập trang login của Moodle
        test_url = f"{moodle_url.rstrip('/')}/login/index.php"
        response = requests.get(test_url, timeout=10)
        if response.status_code in [200, 303]:  # 200: OK, 303: Redirect (thường với Moodle)
            logging.debug("Kết nối server Moodle thành công: %s", test_url)
            return True
        else:
            st.error(f"Không thể kết nối tới server Moodle (HTTP {response.status_code}). Kiểm tra URL hoặc server.")
            logging.error("Không thể kết nối tới server: HTTP %s", response.status_code)
            return False
    except requests.exceptions.RequestException as e:
        st.error(f"Lỗi kiểm tra kết nối tới server Moodle: {str(e)}. Kiểm tra mạng, firewall, hoặc URL.")
        logging.error("Lỗi kiểm tra kết nối: %s", str(e))
        return False

# Hàm gọi API Moodle
def moodle_api_call(function_name, params, moodle_url, token):
    try:
        if not moodle_url or not token:
            st.error("URL Moodle hoặc API Token rỗng. Vui lòng kiểm tra.")
            logging.error("URL hoặc token rỗng.")
            return None
        if not moodle_url.startswith(('http://', 'https://')):
            st.error("URL Moodle phải bắt đầu bằng http:// hoặc https://.")
            logging.error("URL không hợp lệ: %s", moodle_url)
            return None
        
        # Kiểm tra kết nối server
        if not check_server_connectivity(moodle_url):
            return None

        url = f"{moodle_url.rstrip('/')}/webservice/rest/server.php"
        params.update({
            'wstoken': token,
            'wsfunction': function_name,
            'moodlewsrestformat': 'json'
        })
        logging.debug("Gọi API: %s với params: %s", url, params)
        response = requests.get(url, params=params, timeout=15)
        logging.debug("Phản hồi HTTP: %s", response.status_code)
        if response.status_code == 200:
            result = response.json()
            logging.debug("Phản hồi JSON: %s", result)
            if 'exception' in result:
                error_msg = result.get('message', 'Không có thông tin lỗi')
                st.error(f"Lỗi API {function_name}: {error_msg}. "
                         f"Kiểm tra quyền token ('moodle/user:viewdetails', 'moodle/user:viewhiddendetails'), "
                         f"'Người dùng được ủy quyền', hoặc thử tạo token mới.")
                logging.error("Lỗi API %s: %s", function_name, error_msg)
                return None
            return result
        else:
            st.error(f"Lỗi khi gọi API {function_name}: HTTP {response.status_code} - {response.text}. "
                     f"Kiểm tra URL Moodle, token, hoặc dịch vụ web.")
            logging.error("Lỗi HTTP %s: %s", response.status_code, response.text)
            return None
    except requests.exceptions.RequestException as e:
        st.error(f"Lỗi kết nối API {function_name}: {str(e)}. Kiểm tra mạng hoặc server Moodle.")
        logging.error("Lỗi kết nối: %s", str(e))
        return None

# Kiểm tra token hợp lệ đã sửa
def validate_token(moodle_url, token):
    test_username = st.session_state.get('test_username', 'admin')
    if not test_username:
        st.error("Username kiểm tra token rỗng. Vui lòng nhập username hợp lệ.")
        logging.error("Username kiểm tra rỗng.")
        return False
    
    params = {'criteria[0][key]': 'username', 'criteria[0][value]': test_username}
    result = moodle_api_call('core_user_get_users', params, moodle_url, token)
    if result is None:
        st.error(f"Không thể xác thực token. Kiểm tra token, URL, quyền ('moodle/user:viewdetails'), "
                 f"và tài khoản '{test_username}' tồn tại. Nếu thất bại, thử username khác.")
        logging.error("Không thể xác thực token cho username: %s", test_username)
        return False
    elif not result.get('users'):
        st.warning(f"Tài khoản '{test_username}' không được tìm thấy. Thử username khác.")
        logging.warning("Không tìm thấy user: %s", test_username)
        return False
    st.success("Token hợp lệ!")
    return True

# Giao diện Streamlit đã sửa
tab1, tab2 = st.tabs(["📄 Một File", "📂 Nhiều File"])

with tab1:
    st.header("📄 Xử lý Một File Excel")
    moodle_url = st.text_input("🌐 URL Moodle (VD: https://your-moodle-site.com):")
    moodle_token = st.text_input("🔑 API Token:", type="password")
    test_username = st.text_input("🔑 Username kiểm tra token (mặc định: admin):", value="admin")
    st.session_state['test_username'] = test_username  # Lưu username
    
    if moodle_url and moodle_token and st.button("🔍 Kiểm tra Token"):
        if validate_token(moodle_url, moodle_token):
            st.success("Token hợp lệ!")
        else:
            st.error("Token không hợp lệ hoặc thiếu quyền. Kiểm tra token, URL, quyền ('moodle/user:viewdetails'), "
                     f"và tài khoản '{test_username}' trong Moodle.")

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
                        st.error("Không thể tạo hoặc cập nhật khóa học. Kiểm tra token, URL, hoặc category ID.")
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

# Tab Nhiều File (giữ nguyên, không hiển thị để ngắn gọn)
with tab2:
    st.header("📂 Xử lý Nhiều File Excel")
    moodle_url_multi = st.text_input("🌐 URL Moodle (Nhiều File):")
    moodle_token_multi = st.text_input("🔑 API Token (Nhiều File):", type="password")
    
    if moodle_url_multi and moodle_token_multi and st.button("🔍 Kiểm tra Token (Nhiều File)"):
        if validate_token(moodle_url_multi, moodle_token_multi):
            st.success("Token hợp lệ!")
        else:
            st.error("Token không hợp lệ hoặc URL Moodle sai. Kiểm tra URL và token trong Moodle.")

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
                            st.error(f"Không thể tạo/cập nhật khóa học cho file {file.name}. Kiểm tra token, URL, hoặc category ID.")
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
