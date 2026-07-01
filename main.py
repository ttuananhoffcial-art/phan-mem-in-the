import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont, ImageOps
import io
import base64
import uuid
import json 
import re
import datetime
import unicodedata # Thư viện xử lý chữ Tiếng Việt
from collections import Counter
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Image as RLImage
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas

# Thư mục tạm & File lưu trữ cài đặt
OUTPUT_DIR = "the_tam_thoi"
PHOI_VDV_PATH = "phoi_vdv.png"
PHOI_HLV_PATH = "phoi_hlv.png"
CONFIG_FILE = "config_the.json" 
USER_FILE = "users.json"
os.makedirs(OUTPUT_DIR, exist_ok=True)

st.set_page_config(page_title="Phần mềm Thẻ Taekwondo", page_icon="🥋", layout="wide")

# =========================================================
# CÔNG CỤ XỬ LÝ TEXT TỐI THƯỢNG (Chống lỗi font & khoảng trắng ảo)
# =========================================================
def bo_dau_tieng_viet(text):
    """Lột sạch dấu và khoảng trắng ảo để nhận diện chức vụ VIP xác suất 100%"""
    if pd.isna(text): return ""
    s = re.sub(r'\s+', ' ', str(text).strip()) # Tiêu diệt khoảng trắng ảo
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if not unicodedata.combining(c))
    return s.replace('Đ', 'D').replace('đ', 'd').upper()
    
def chuan_hoa_chu(text):
    """Chuẩn hóa font chữ in lên thẻ đẹp nhất"""
    if pd.isna(text): return ""
    text = re.sub(r'\s+', ' ', str(text).strip())
    return unicodedata.normalize('NFC', text.upper())

def load_users():
    if os.path.exists(USER_FILE):
        try:
            with open(USER_FILE, 'r', encoding='utf-8') as f: return json.load(f)
        except: pass
    default_users = {"admin": {"password": "123456", "expiry": None, "role": "admin"}}
    with open(USER_FILE, 'w', encoding='utf-8') as f: json.dump(default_users, f, ensure_ascii=False, indent=4)
    return default_users

def save_users(users):
    with open(USER_FILE, 'w', encoding='utf-8') as f: json.dump(users, f, ensure_ascii=False, indent=4)

def get_default_config(ref_w, ref_h):
    return {
        'font_name': 'Arial Bold',
        'img_x': int(ref_w * 0.08), 'img_y': int(ref_h * 0.62),
        'img_w': int(ref_w * 0.31), 'img_h': int(ref_h * 0.28),
        'l1_color': '#ED1C24', 'l1_size': int(ref_w * 0.06), 'l1_x': int(ref_w * 0.68), 'l1_y': int(ref_h * 0.68),
        'l2_color': '#ED1C24', 'l2_size': int(ref_w * 0.06), 'l2_x': int(ref_w * 0.68), 'l2_y': int(ref_h * 0.74),
        'l3_color': '#ED1C24', 'l3_size': int(ref_w * 0.06), 'l3_x': int(ref_w * 0.68), 'l3_y': int(ref_h * 0.80),
        'l4_color': '#ED1C24', 'l4_size': int(ref_w * 0.06), 'l4_x': int(ref_w * 0.68), 'l4_y': int(ref_h * 0.86),
    }

def load_config(default_cfg):
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                saved = json.load(f)
                cfg = default_cfg.copy()
                cfg.update(saved)
                return cfg
        except: return default_cfg
    return default_cfg

def save_config(cfg):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(cfg, f, ensure_ascii=False, indent=4)
    except Exception as e: print(f"Lỗi lưu cấu hình: {e}")

def extract_images_from_excel(file_buffer):
    img_dir = os.path.join(OUTPUT_DIR, "extracted_images")
    os.makedirs(img_dir, exist_ok=True)
    images_info = []
    try:
        with zipfile.ZipFile(file_buffer, 'r') as archive:
            drawings = [f for f in archive.namelist() if f.startswith('xl/drawings/') and f.endswith('.xml') and '_rels' not in f]
            for draw_file in drawings:
                try:
                    filename = os.path.basename(draw_file)
                    rel_file = f"xl/drawings/_rels/{filename}.rels"
                    image_map = {}
                    if rel_file in archive.namelist():
                        rel_root = ET.fromstring(archive.read(rel_file))
                        for child in rel_root:
                            rId = child.attrib.get('Id')
                            target = child.attrib.get('Target')
                            if rId and target: image_map[rId] = os.path.basename(target)
                    content = archive.read(draw_file)
                    root = ET.fromstring(content)
                    namespaces = {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
                    for anchor_type in ['.//xdr:twoCellAnchor', './/xdr:oneCellAnchor']:
                        for anchor in root.findall(anchor_type, namespaces):
                            from_elem = anchor.find('xdr:from', namespaces)
                            if from_elem is not None:
                                row = int(from_elem.find('xdr:row', namespaces).text)
                                col = int(from_elem.find('xdr:col', namespaces).text)
                                rowOff = int(from_elem.find('xdr:rowOff', namespaces).text)
                                rot = 0
                                for elem in anchor.iter():
                                    if elem.tag.endswith('xfrm'):
                                        rot_str = elem.attrib.get('rot')
                                        if rot_str:
                                            try: rot = int(rot_str)
                                            except: pass
                                    if elem.tag.endswith('blip'):
                                        rId = elem.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                                        if rId and rId in image_map:
                                            img_name = image_map[rId]
                                            img_path_in_zip = f"xl/media/{img_name}"
                                            if img_path_in_zip in archive.namelist():
                                                out_path = os.path.join(img_dir, f"img_{uuid.uuid4().hex}.png")
                                                with open(out_path, "wb") as f: f.write(archive.read(img_path_in_zip))
                                                images_info.append({'row': row, 'col': col, 'rowOff': rowOff, 'path': out_path, 'rot': rot})
                except Exception as ex: print("Lỗi parse:", ex)
    except Exception as e: print(f"Lỗi đọc ảnh từ Excel: {e}")
    return images_info

def lay_font_tieng_viet(font_name, size):
    size = int(size)
    font_map = {
        "Arial Bold": ["arialbd.ttf", "Arial.ttf", "arial.ttf"],
        "Times New Roman Bold": ["timesbd.ttf", "times.ttf", "timesi.ttf"],
        "Tahoma Bold": ["tahomabd.ttf", "tahoma.ttf", "Tahoma"],
        "Calibri Bold": ["calibrib.ttf", "calibri.ttf", "Calibri"]
    }
    files = font_map.get(font_name, ["arialbd.ttf"])
    for f in files:
        try: return ImageFont.truetype(f, size)
        except: pass
    try: return ImageFont.truetype(font_name, size)
    except: return ImageFont.load_default()

def ve_chu_tu_dong_co_gian(draw, text, center_x, y, font_name, initial_size, fill, max_width):
    if not text: return
    size = int(initial_size)
    font = lay_font_tieng_viet(font_name, size)
    if isinstance(fill, str) and fill.startswith('#'):
        fill = fill.lstrip('#')
        fill = tuple(int(fill[i:i+2], 16) for i in (0, 2, 4))
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        text_w = bbox[2] - bbox[0]
        while text_w > max_width and size > 20:
            size -= 2
            font = lay_font_tieng_viet(font_name, size)
            bbox = draw.textbbox((0, 0), text, font=font)
            text_w = bbox[2] - bbox[0]
        start_x = center_x - (text_w / 2)
        draw.text((start_x, y), text, fill=fill, font=font)
    except Exception as e:
        draw.text((center_x, y), text, fill=fill, font=font)

def xu_ly_text_in_the(text):
    if pd.isna(text): return ""
    text = str(text).strip()
    if text.upper() == "NAN" or text == "": return ""
    
    if "00:00:00" in text or re.match(r"^\d{4}-\d{2}-\d{2}", text): return text.split("-")[0]
    if re.match(r"^\d{1,2}/\d{1,2}/\d{4}", text): return text.split("/")[-1]
    if text.endswith(".0"): text = text[:-2]
    
    txt_upper = chuan_hoa_chu(text)
    mapping = {
        "HLV": "HUẤN LUYỆN VIÊN", "VĐV": "VẬN ĐỘNG VIÊN", "VDV": "VẬN ĐỘNG VIÊN", "BTC": "BAN TỔ CHỨC",
        "TRƯỞNG ĐOÀN": "TRƯỞNG ĐOÀN", "HLV TRƯỞNG": "HLV TRƯỞNG", "THƯ KÝ": "THƯ KÝ", "TRỌNG TÀI": "TRỌNG TÀI"
    }
    return mapping.get(txt_upper, txt_upper)

def tao_the_ca_nhan(data, img_info, chi_in_noi_dung, cfg, col_l1, col_l2, col_l3, col_l4, excel_row, idx_count, phoi_vdv, phoi_hlv):
    # Dùng list comprehension để gộp tất cả các cột thành chuỗi vô cực (Đã xử lý khoảng trắng ảo)
    all_vals = [str(val) for val in data.tolist() if pd.notna(val)]
    all_text_clean = " ".join([bo_dau_tieng_viet(val) for val in all_vals])
    
    # Keyword VIP tuyệt đối (Đã được viết không dấu)
    chuc_vu_vip = ["HLV", "HUAN LUYEN VIEN", "TRONG TAI", "BTC", "TRUONG DOAN", "BAN TO CHUC", "THU KY"]
    is_hlv = any(kw in all_text_clean for kw in chuc_vu_vip)
    
    phoi_chon = phoi_hlv if is_hlv else phoi_vdv
    phoi_goc = Image.open(phoi_chon).convert("RGBA")
    phoi_w, phoi_h = phoi_goc.size
    card = Image.new("RGBA", (phoi_w, phoi_h), (255, 255, 255, 255)) if chi_in_noi_dung else phoi_goc
    draw = ImageDraw.Draw(card)
    
    if img_info and os.path.exists(img_info['path']):
        try:
            anh_vdv = Image.open(img_info['path'])
            anh_vdv = ImageOps.exif_transpose(anh_vdv)
            rot_val = img_info.get('rot', 0)
            if rot_val != 0:
                anh_vdv = anh_vdv.rotate(-(rot_val / 60000.0), expand=True)
            anh_vdv = anh_vdv.convert("RGBA")
            if anh_vdv.width > anh_vdv.height:
                rot_choice = st.session_state.get(f"radar_rot_{excel_row}", "➖ Giữ nguyên")
                if "👈 Đầu Trái" in rot_choice: anh_vdv = anh_vdv.rotate(-90, expand=True)
                elif "👉 Đầu Phải" in rot_choice: anh_vdv = anh_vdv.rotate(90, expand=True)
            img_w, img_h = int(cfg['img_w']), int(cfg['img_h'])
            img_x, img_y = int(cfg['img_x']), int(cfg['img_y'])
            anh_vdv = ImageOps.fit(anh_vdv, (img_w, img_h), method=Image.Resampling.LANCZOS, centering=(0.5, 0.5))
            card.paste(anh_vdv, (img_x, img_y), anh_vdv)
        except Exception as e: print(f"Lỗi vẽ ảnh: {e}")

    max_text_width = int(phoi_w * 0.55) 
    if col_l1 != "--- Không in ---" and pd.notna(data.get(col_l1)):
        txt = xu_ly_text_in_the(data.get(col_l1))
        ve_chu_tu_dong_co_gian(draw, txt, cfg['l1_x'], cfg['l1_y'], cfg['font_name'], cfg['l1_size'], cfg['l1_color'], max_text_width)
    if col_l2 != "--- Không in ---" and pd.notna(data.get(col_l2)):
        txt = xu_ly_text_in_the(data.get(col_l2))
        ve_chu_tu_dong_co_gian(draw, txt, cfg['l2_x'], cfg['l2_y'], cfg['font_name'], cfg['l2_size'], cfg['l2_color'], max_text_width)
    if col_l3 != "--- Không in ---" and pd.notna(data.get(col_l3)):
        txt = xu_ly_text_in_the(data.get(col_l3))
        ve_chu_tu_dong_co_gian(draw, txt, cfg['l3_x'], cfg['l3_y'], cfg['font_name'], cfg['l3_size'], cfg['l3_color'], max_text_width)
    if col_l4 != "--- Không in ---" and pd.notna(data.get(col_l4)):
        txt = xu_ly_text_in_the(data.get(col_l4))
        ve_chu_tu_dong_co_gian(draw, txt, cfg['l4_x'], cfg['l4_y'], cfg['font_name'], cfg['l4_size'], cfg['l4_color'], max_text_width)

    path_luu_tam = os.path.join(OUTPUT_DIR, f"the_don_{idx_count}.jpg")
    card.convert("RGB").save(path_luu_tam, format="JPEG", quality=95)
    return path_luu_tam

# --- ĐIỀU KHIỂN ĐĂNG NHẬP ---
if not st.session_state.get('logged_in', False):
    st.markdown("<br><br><br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1.5, 1, 1.5]) 
    with col2:
        st.markdown("<h2 style='text-align: center; color: #1B368E;'>🥋 ĐĂNG NHẬP HỆ THỐNG</h2>", unsafe_allow_html=True)
        with st.form("Login_Form"):
            user_input = st.text_input("Tài khoản:").strip()
            pass_input = st.text_input("Mật khẩu:", type="password").strip()
            submit_btn = st.form_submit_button("ĐĂNG NHẬP", use_container_width=True)
            if submit_btn:
                all_users = load_users()
                if user_input in all_users:
                    u_data = all_users[user_input]
                    if u_data["password"] == pass_input:
                        if u_data["expiry"] is not None:
                            expiry_time = datetime.datetime.strptime(u_data["expiry"], "%Y-%m-%d %H:%M:%S")
                            if datetime.datetime.now() > expiry_time:
                                st.error("❌ Tài khoản này đã HẾT HẠN SỬ DỤNG! Vui lòng liên hệ Admin.")
                                st.stop()
                        st.session_state['logged_in'] = True
                        st.session_state['username'] = user_input
                        st.session_state['role'] = u_data.get("role", "user")
                        st.rerun()
                    else: st.error("❌ Sai mật khẩu!")
                else: st.error("❌ Tài khoản không tồn tại trên hệ thống!")
else:
    st.sidebar.markdown(f"👤 Xin chào: **{st.session_state['username'].upper()}**")
    
    menu_options = ["🏠 Trang chủ", "➕ Tạo thẻ thi đấu"]
    if st.session_state.get('role') == "admin":
        menu_options.append("🔑 Quản lý hệ thống")
        
    menu_selection = st.sidebar.radio("Chọn hành động:", menu_options)
    if st.sidebar.button("🔒 Đăng xuất"):
        st.session_state['logged_in'] = False
        st.rerun()

    if menu_selection == "🏠 Trang chủ":
        st.title("🥋 HỆ THỐNG IN THẺ TAEKWONDO TÙY BIẾN")
        st.info("Chào mừng bạn quay trở lại phần mềm quản lý thẻ.")

    elif menu_selection == "🔑 Quản lý hệ thống" and st.session_state.get('role') == "admin":
        st.title("🔑 TRUNG TÂM QUẢN LÝ TÀI KHOẢN & BẢO MẬT")
        all_users = load_users()
        tab_admin, tab_sub_users = st.tabs(["🔒 Đổi mật khẩu Admin", "👥 Cấp tài khoản người dùng"])
        
        with tab_admin:
            st.subheader("Thay đổi mật khẩu Admin")
            new_pass = st.text_input("Mật khẩu mới:", type="password", key="adm_pass")
            confirm_pass = st.text_input("Xác nhận mật khẩu mới:", type="password", key="adm_conf")
            if st.button("Cập nhật mật khẩu Admin", type="primary"):
                if new_pass == confirm_pass:
                    if len(new_pass) >= 4:
                        all_users["admin"]["password"] = new_pass
                        save_users(all_users)
                        st.success("✅ Cập nhật mật khẩu Admin thành công! Hệ thống đã ghi nhớ.")
                    else: st.error("❌ Mật khẩu phải dài từ 4 ký tự trở lên!")
                else: st.error("❌ Mật khẩu xác nhận không khớp!")

        with tab_sub_users:
            st.subheader("Tạo tài khoản dùng thử cấp tốc")
            col_u1, col_u2, col_u3 = st.columns(3)
            new_u_name = col_u1.text_input("Tên tài khoản người dùng:", placeholder="clb_huynhthanh").strip()
            new_u_pass = col_u2.text_input("Mật khẩu cấp:", type="password", placeholder="123456").strip()
            days_to_live = col_u3.number_input("Hết hạn sau số ngày:", min_value=1, value=3, step=1)
            if st.button("Tạo tài khoản người dùng"):
                if new_u_name and new_u_pass:
                    if new_u_name in all_users: st.error("❌ Tên tài khoản này đã tồn tại rồi!")
                    else:
                        expiry_date = datetime.datetime.now() + datetime.timedelta(days=int(days_to_live))
                        all_users[new_u_name] = {"password": new_u_pass, "expiry": expiry_date.strftime("%Y-%m-%d %H:%M:%S"), "role": "user"}
                        save_users(all_users)
                        st.success(f"✅ Đã cấp thành công tài khoản '{new_u_name}'. Hết hạn vào: {expiry_date.strftime('%d/%m/%Y %H:%M:%S')}")
                        st.rerun()
                else: st.error("Vui lòng nhập đầy đủ Tài khoản và Mật khẩu!")

            st.markdown("---")
            st.subheader("📋 Danh sách các tài khoản đang hoạt động")
            user_list_data = []
            for u, info in all_users.items():
                if u == "admin": continue
                exp_str = info["expiry"]
                exp_dt = datetime.datetime.strptime(exp_str, "%Y-%m-%d %H:%M:%S")
                tinh_trang = "🟢 Đang chạy" if datetime.datetime.now() < exp_dt else "🔴 Đã hết hạn"
                user_list_data.append({"Tài khoản": u, "Mật khẩu": info["password"], "Thời gian hết hạn": exp_dt.strftime("%d/%m/%Y %H:%M:%S"), "Trạng thái": tinh_trang})
            if user_list_data:
                st.dataframe(pd.DataFrame(user_list_data), use_container_width=True)
                del_user = st.selectbox("Chọn tài khoản muốn xóa bỏ hẳn:", [x["Tài khoản"] for x in user_list_data])
                if st.button("Xóa tài khoản đã chọn"):
                    if del_user in all_users:
                        del all_users[del_user]
                        save_users(all_users)
                        st.success(f"❌ Đã xóa vĩnh viễn tài khoản '{del_user}'")
                        st.rerun()
            else: st.info("Chưa có tài khoản phụ nào được cấp.")

    elif menu_selection == "➕ Tạo thẻ thi đấu":
        st.title("➕ TẠO THẺ THI ĐẤU & DÀN TRANG IN")
        
        ref_w, ref_h = 1000, 1400
        if os.path.exists(PHOI_VDV_PATH):
            try:
                with Image.open(PHOI_VDV_PATH) as img_ref: ref_w, ref_h = img_ref.size
            except: pass

        st.sidebar.markdown("### 📏 Kích thước & Khổ giấy in:")
        the_width_cm = st.sidebar.number_input("Chiều ngang thẻ (cm):", value=10.0, step=0.1)
        the_height_cm = st.sidebar.number_input("Chiều cao thẻ (cm):", value=14.0, step=0.1)
        kieu_xuat_file = st.sidebar.radio("Chọn bố cục file PDF:", ["🔲 4 thẻ / 1 trang A4", "📄 1 thẻ / 1 trang"])
        chi_in_noi_dung = (st.sidebar.radio("Tùy chọn nền phôi:", ["🖼️ In đầy đủ (Cả nền Xanh/Hồng)", "⬜ Chỉ in nội dung (Nền trắng)"]) == "⬜ Chỉ in nội dung (Nền trắng)")

        default_cfg = get_default_config(ref_w, ref_h)
        cfg = load_config(default_cfg)

        st.markdown("### ⚙️ 3. Căn chỉnh Tọa độ & Cỡ chữ:")
        danh_sach_font = ["Arial Bold", "Times New Roman Bold", "Tahoma Bold", "Calibri Bold"]
        font_hien_tai = cfg.get('font_name', "Arial Bold")
        cfg['font_name'] = st.selectbox("🔤 Chọn kiểu phông chữ:", danh_sach_font, index=danh_sach_font.index(font_hien_tai) if font_hien_tai in danh_sach_font else 0)

        with st.expander("🛠️ Bấm vào đây để KÉO ẢNH & ĐỔI MÀU CHỮ (Tự động lưu)", expanded=False):
            tab_img, tab_txt1, tab_txt2 = st.tabs(["📸 Ảnh chân dung", "🔤 Dòng 1 & Dòng 2", "🔤 Dòng 3 & Dòng 4"])
            with tab_img:
                col1, col2 = st.columns(2)
                cfg['img_x'] = col1.number_input("Vị trí ngang X (Ảnh)", value=int(cfg['img_x']), step=10)
                cfg['img_y'] = col2.number_input("Vị trí dọc Y (Ảnh)", value=int(cfg['img_y']), step=10)
                cfg['img_w'] = col1.number_input("Chiều rộng ảnh", value=int(cfg['img_w']), step=10)
                cfg['img_h'] = col2.number_input("Chiều cao ảnh", value=int(cfg['img_h']), step=10)
            with tab_txt1:
                st.markdown("**🔹 Cấu hình Dòng 1**")
                cfg['l1_color'] = st.color_picker("🎨 Màu Dòng 1:", value=cfg['l1_color'])
                col1, col2, col3 = st.columns(3)
                cfg['l1_size'] = col1.number_input("Cỡ chữ Dòng 1", value=int(cfg['l1_size']), step=5)
                cfg['l1_x'] = col2.number_input("Tâm X Dòng 1", value=int(cfg['l1_x']), step=10)
                cfg['l1_y'] = col3.number_input("Vị trí Y Dòng 1", value=int(cfg['l1_y']), step=10)
                st.markdown("---")
                st.markdown("**🔹 Cấu hình Dòng 2**")
                cfg['l2_color'] = st.color_picker("🎨 Màu Dòng 2:", value=cfg['l2_color'])
                col1, col2, col3 = st.columns(3)
                cfg['l2_size'] = col1.number_input("Cỡ chữ Dòng 2", value=int(cfg['l2_size']), step=5)
                cfg['l2_x'] = col2.number_input("Tâm X Dòng 2", value=int(cfg['l2_x']), step=10)
                cfg['l2_y'] = col3.number_input("Vị trí Y Dòng 2", value=int(cfg['l2_y']), step=10)
            with tab_txt2:
                st.markdown("**🔹 Cấu hình Dòng 3**")
                cfg['l3_color'] = st.color_picker("🎨 Màu Dòng 3:", value=cfg['l3_color'])
                col1, col2, col3 = st.columns(3)
                cfg['l3_size'] = col1.number_input("Cỡ chữ Dòng 3", value=int(cfg['l3_size']), step=5)
                cfg['l3_x'] = col2.number_input("Tâm X Dòng 3", value=int(cfg['l3_x']), step=10)
                cfg['l3_y'] = col3.number_input("Vị trí Y Dòng 3", value=int(cfg['l3_y']), step=10)
                st.markdown("---")
                st.markdown("**🔹 Cấu hình Dòng 4**")
                cfg['l4_color'] = st.color_picker("🎨 Màu Dòng 4:", value=cfg['l4_color'])
                col1, col2, col3 = st.columns(3)
                cfg['l4_size'] = col1.number_input("Cỡ chữ Dòng 4", value=int(cfg['l4_size']), step=5)
                cfg['l4_x'] = col2.number_input("Tâm X Dòng 4", value=int(cfg['l4_x']), step=10)
                cfg['l4_y'] = col3.number_input("Vị trí Y Dòng 4", value=int(cfg['l4_y']), step=10)

        save_config(cfg)

        st.markdown("---")
        col_up1, col_up2 = st.columns([3, 1])
        with col_up1:
            file_excel = st.file_uploader("📂 4. Chọn file Excel danh sách:", type=["xlsx"])
        with col_up2:
            st.markdown("<br>", unsafe_allow_html=True)
            header_row = st.number_input("⚙️ Dòng chứa Tiêu đề cột:", min_value=1, value=4, step=1)
        
        if file_excel is not None:
            # =================================================================
            # HỆ THỐNG RAM SIÊU TỐC: ĐỌC DỮ LIỆU ĐÚNG 1 LẦN DUY NHẤT RỒI LƯU LẠI
            # =================================================================
            file_bytes = file_excel.getvalue()
            file_id = f"{file_excel.name}_{file_excel.size}_{header_row}"
            
            if st.session_state.get('current_file_id') != file_id:
                st.session_state['current_file_id'] = file_id
                with st.spinner("⏳ Đang xử lý dữ liệu và bóc tách ảnh (Chỉ chạy 1 lần duy nhất để chống lag)..."):
                    skip_r = int(header_row) - 1
                    df_raw = pd.read_excel(io.BytesIO(file_bytes), skiprows=skip_r)
                    st.session_state['raw_df'] = df_raw
                    st.session_state['ban_do_anh'] = extract_images_from_excel(io.BytesIO(file_bytes))
            
            # Kéo dữ liệu đã lưu trong RAM ra dùng cực nhanh
            df_cols = st.session_state['raw_df'].copy()
            ban_do_anh = st.session_state['ban_do_anh'].copy()
            
            clean_cols = []
            for i, c in enumerate(df_cols.columns):
                c_str = str(c).strip()
                if "Unnamed" in c_str or c_str.lower() == "nan":
                    clean_cols.append(f"Cột {i+1} (Bị trống tiêu đề)")
                else:
                    clean_cols.append(c_str)
            df_cols.columns = clean_cols
            
            st.markdown("👁️ **Màn hình X-Quang: Đây là những cột máy tính đọc được (Hãy tìm cột Năm sinh của bạn ở đây):**")
            st.dataframe(df_cols.head(2), use_container_width=True)

            cols_list = ["--- Không in ---"] + list(df_cols.columns)
            
            def_l1, def_l2, def_l3, def_l4 = 0, 0, 0, 0
            for i, col in enumerate(cols_list):
                c_low = chuan_hoa_chu(col).lower()
                if "tên" in c_low or "ten" in c_low: def_l1 = i
                if "chức vụ" in c_low or "chuc vu" in c_low: def_l2 = i
                if "sinh" in c_low or "năm" in c_low: def_l3 = i
                if "đơn vị" in c_low or "don vi" in c_low or "clb" in c_low or "đv" in c_low: def_l4 = i

            st.markdown("### 📋 5. Ghép Cột Dữ Liệu Tùy Biến:")
            col_a, col_b, col_c, col_d = st.columns(4)
            col_l1 = col_a.selectbox("Dòng 1 in cột:", cols_list, index=def_l1)
            col_l2 = col_b.selectbox("Dòng 2 in cột:", cols_list, index=def_l2)
            col_l3 = col_c.selectbox("Dòng 3 in cột:", cols_list, index=def_l3)
            col_l4 = col_d.selectbox("Dòng 4 in cột:", cols_list, index=def_l4)

            col_id = col_l1 if col_l1 != "--- Không in ---" else df_cols.columns[0]
            
            df = df_cols.copy()
            df = df.dropna(subset=[col_id])
            df = df.ffill()

            if ban_do_anh:
                phobien_col = Counter([img['col'] for img in ban_do_anh]).most_common(1)[0][0]
                valid_images = [img for img in ban_do_anh if img['col'] == phobien_col]
                valid_images.sort(key=lambda x: (x['row'], x['rowOff']))
            else: valid_images = []

            persons = []
            for idx, row in df.iterrows():
                current_xml_row = idx + int(header_row)
                current_excel_row = current_xml_row + 1
                persons.append({'row_data': row, 'xml_row': current_xml_row, 'excel_row': current_excel_row, 'img_info': None})

            for p in persons:
                for i, img in enumerate(valid_images):
                    if img['row'] == p['xml_row']: p['img_info'] = valid_images.pop(i); break
            for p in persons:
                if p['img_info'] is None:
                    for i, img in enumerate(valid_images):
                        if abs(img['row'] - p['xml_row']) <= 1: p['img_info'] = valid_images.pop(i); break

            nguoi_bi_ngang = []
            for p in persons:
                if p['img_info'] and os.path.exists(p['img_info']['path']):
                    try:
                        img_test = Image.open(p['img_info']['path']).convert("RGB")
                        img_test = ImageOps.exif_transpose(img_test)
                        if img_test.width > img_test.height: nguoi_bi_ngang.append(p)
                    except: pass

            if len(nguoi_bi_ngang) > 0:
                st.markdown("---")
                st.error(f"🚨 PHÁT HIỆN {len(nguoi_bi_ngang)} ẢNH BỊ NẰM NGANG TRONG EXCEL!")
                st.info("💡 Lưới siêu tốc: Bấm để dựng thẳng ảnh.")
                cols_radar = st.columns(8)
                for i, p in enumerate(nguoi_bi_ngang):
                    with cols_radar[i % 8]:
                        st.markdown(f"<div style='padding: 5px; border: 1px solid #ff4b4b; border-radius: 5px; text-align: center; background-color: #fff9f9;'>", unsafe_allow_html=True)
                        try:
                            img_thumb = Image.open(p['img_info']['path'])
                            img_thumb = ImageOps.exif_transpose(img_thumb)
                            img_thumb.thumbnail((100, 100))
                            st.image(img_thumb, use_container_width=True)
                        except: pass
                        ten = str(p['row_data'].get(col_id, f"Dòng {p['excel_row']}"))
                        if len(ten) > 15: ten = ten[:13] + "..."
                        st.markdown(f"<p style='font-size:12px; margin-bottom:5px;'><b>{ten}</b></p>", unsafe_allow_html=True)
                        st.radio("Cách xử lý:", ["➖ Giữ nguyên", "👈 Đầu Trái", "👉 Đầu Phải"], key=f"radar_rot_{p['excel_row']}", label_visibility="collapsed")
                        st.markdown("</div><br>", unsafe_allow_html=True)

            st.markdown("---")
            if st.button("⚡ BẮT ĐẦU TẠO ẢNH THẺ COMPLETE", type="primary", use_container_width=True):
                if not os.path.exists(PHOI_VDV_PATH) or not os.path.exists(PHOI_HLV_PATH):
                    st.error("❌ Thiếu file phoi_vdv.png hoặc phoi_hlv.png trong thư mục phần mềm.")
                else:
                    danh_sach_duong_dan_the = []
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    total_rows = len(persons)
                    
                    for idx_loop, p in enumerate(persons):
                        ten_hien_tai = str(p['row_data'].get(col_id, ""))
                        status_text.text(f"⏳ Đang vẽ và kết xuất thẻ {idx_loop + 1}/{total_rows}: {ten_hien_tai}...")
                        progress_bar.progress((idx_loop + 1) / total_rows)
                        path_tam = tao_the_ca_nhan(p['row_data'], p['img_info'], chi_in_noi_dung, cfg, col_l1, col_l2, col_l3, col_l4, p['excel_row'], idx_loop, PHOI_VDV_PATH, PHOI_HLV_PATH)
                        danh_sach_duong_dan_the.append(path_tam)
                    
                    status_text.text("✅ Đang dàn trang xuất file PDF...")
                    if len(danh_sach_duong_dan_the) > 0:
                        st.markdown("---")
                        st.markdown("<h2 style='text-align: center; color: #1B368E;'>🖨️ FILE ĐÃ SẴN SÀNG ĐỂ IN</h2>", unsafe_allow_html=True)
                        pdf_buffer = io.BytesIO()
                        col_space1, col_center, col_space3 = st.columns([1, 2, 1])
                        with col_center:
                            if "4 thẻ" in kieu_xuat_file:
                                for i in range(0, len(danh_sach_duong_dan_the), 4):
                                    batch_paths = danh_sach_duong_dan_the[i:i+4]
                                    sample_card = Image.open(batch_paths[0])
                                    c_w, c_h = sample_card.size
                                    gap = int(c_w * 0.05)
                                    a4_preview = Image.new("RGB", (c_w * 2 + gap, c_h * 2 + gap), "white")
                                    for j, p in enumerate(batch_paths): a4_preview.paste(Image.open(p), ((j % 2) * (c_w + gap), (j // 2) * (c_h + gap)))
                                    st.image(a4_preview, caption=f"Trang A4 số {int(i/4) + 1}", use_container_width=True)

                                doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, leftMargin=0.4*cm, rightMargin=0.4*cm, topMargin=0.5*cm, bottomMargin=0.5*cm)
                                story, du_lieu_bang, hang_hien_tai = [], [], []
                                for i, path_the in enumerate(danh_sach_duong_dan_the):
                                    hang_hien_tai.append(RLImage(path_the, width=the_width_cm*cm, height=the_height_cm*cm))
                                    if len(hang_hien_tai) == 2 or i == len(danh_sach_duong_dan_the) - 1:
                                        if len(hang_hien_tai) == 1: hang_hien_tai.append("") 
                                        du_lieu_bang.append(hang_hien_tai)
                                        hang_hien_tai = []
                                table = Table(du_lieu_bang, colWidths=[the_width_cm*cm]*2, rowHeights=[the_height_cm*cm]*len(du_lieu_bang))
                                table.setStyle(TableStyle([('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
                                story.append(table)
                                doc.build(story)
                                ten_file = "In_The_Grid_A4.pdf"
                            else:
                                st.info(f"💡 Kích thước xuất PDF: {the_width_cm}cm x {the_height_cm}cm mỗi trang.")
                                for i, p in enumerate(danh_sach_duong_dan_the): st.image(p, caption=f"Thẻ số {i+1}", use_container_width=True)
                                c = canvas.Canvas(pdf_buffer, pagesize=(the_width_cm*cm, the_height_cm*cm))
                                for path_the in danh_sach_duong_dan_the: c.drawImage(path_the, 0, 0, width=the_width_cm*cm, height=the_height_cm*cm); c.showPage()
                                c.save()
                                ten_file = f"In_The_Don_{the_width_cm}x{the_height_cm}.pdf"
                        
                        pdf_bytes = pdf_buffer.getvalue()
                        status_text.empty(); progress_bar.empty()
                        st.download_button("🔥 BẤM ĐỂ TẢI FILE PDF IN NGAY 🔥", data=pdf_bytes, file_name=ten_file, mime="application/pdf", use_container_width=True)
                        st.balloons()