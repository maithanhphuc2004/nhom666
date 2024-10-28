from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re
from openpyxl import Workbook
import sqlite3  # Thêm thư viện SQLite

# Đường dẫn đến file chromedriver.exe
chrome_path = r'C:\Users\ACER\OneDrive\Documents\phuc\New folder\chromedriver.exe'

# Khởi tạo driver toàn cục
options = Options()
options.add_argument("--disable-infobars")
options.add_argument("--disable-notifications")
options.add_argument("--mute-audio")
options.add_experimental_option("prefs", {
    "profile.default_content_setting_values.notifications": 2
})

driver = webdriver.Chrome(service=Service(chrome_path), options=options)

# Hàm lấy thông tin chi tiết sản phẩm
def get_detailed_product_info(link):
    print(f"Đang truy cập vào: {link}")
    start_time = time.time()  # Bắt đầu tính thời gian
    driver.get(link)  # Sử dụng driver toàn cục
    wait = WebDriverWait(driver, 10)  # Giảm thời gian chờ xuống 10 giây

    # Lấy các thông tin chi tiết sản phẩm
    try:
        ma_hang = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_sku"))).text
    except Exception as e:
        ma_hang = "Không tìm thấy"
        print(f"Lỗi lấy mã hàng: {e}")

    try:
        nha_cung_cap = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_supplier"))).text
    except Exception as e:
        nha_cung_cap = "Không tìm thấy"
        print(f"Lỗi lấy nhà cung cấp: {e}")

    # Thêm đoạn mã lấy người dịch
    try:
        nguoi_dich = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_translator"))).text
    except Exception as e:
        nguoi_dich = "Không tìm thấy"
        print(f"Lỗi lấy người dịch: {e}")

    try:
        tac_gia = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_author"))).text
    except Exception as e:
        tac_gia = "Không tìm thấy"
        print(f"Lỗi lấy tác giả: {e}")

    try:
        nxb = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_publisher"))).text
    except Exception as e:
        nxb = "Không tìm thấy"
        print(f"Lỗi lấy NXB: {e}")

    try:
        nam_xb = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_publish_year"))).text
    except Exception as e:
        nam_xb = "Không tìm thấy"
        print(f"Lỗi lấy năm xuất bản: {e}")

    try:
        ngon_ngu = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_languages"))).text
    except Exception as e:
        ngon_ngu = "Không tìm thấy"
        print(f"Lỗi lấy ngôn ngữ: {e}")

    try:
        trong_luong = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_weight"))).text
    except Exception as e:
        trong_luong = "Không tìm thấy"
        print(f"Lỗi lấy trọng lượng: {e}")

    try:
        kich_thuoc = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_size"))).text
    except Exception as e:
        kich_thuoc = "Không tìm thấy"
        print(f"Lỗi lấy kích thước: {e}")

    try:
        so_trang = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_qty_of_page"))).text
    except Exception as e:
        so_trang = "Không tìm thấy"
        print(f"Lỗi lấy số trang: {e}")

    try:
        hinh_thuc = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_book_layout"))).text
    except Exception as e:
        hinh_thuc = "Không tìm thấy"
        print(f"Lỗi lấy hình thức: {e}")

    end_time = time.time()  # Kết thúc tính thời gian
    print(f"Thời gian truy cập: {end_time - start_time:.2f} giây\n")  # Hiển thị thời gian truy cập

    return {
        "ma_hang": ma_hang,
        "nha_cung_cap": nha_cung_cap,
        "nguoi_dich": nguoi_dich,
        "tac_gia": tac_gia,
        "nxb": nxb,
        "nam_xb": nam_xb,
        "ngon_ngu": ngon_ngu,
        "trong_luong": trong_luong,
        "kich_thuoc": kich_thuoc,
        "so_trang": so_trang,
        "hinh_thuc": hinh_thuc
    }
