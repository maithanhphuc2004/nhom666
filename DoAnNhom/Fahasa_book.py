from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor
import time
import re
from openpyxl import Workbook
import sqlite3

# Đường dẫn đến file chromedriver.exe
chrome_path = r'C:\Users\ACER\OneDrive\Documents\phuc\New folder\chromedriver.exe'

# Khởi tạo driver cho mỗi luồng
def create_driver():
    options = Options()
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-notifications")
    options.add_argument("--mute-audio")
    #options.add_argument("--headless")  # Chạy trong chế độ không hiển thị

    options.add_experimental_option("prefs", {
        "profile.default_content_setting_values.notifications": 2
    })
    return webdriver.Chrome(service=Service(chrome_path), options=options)


# Hàm lấy thông tin chi tiết sản phẩm
def get_detailed_product_info(product):
    link = product['link']
    driver = create_driver()
    print(f"Đang truy cập vào: {link}")
    driver.get(link)
    wait = WebDriverWait(driver, 20)

    # Cuộn xuống cuối trang để tải hết dữ liệu
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(2)  # Đợi một chút để trang tải hết dữ liệu

# Cào dữ liệu chi tiết sản phẩm

    # lấy mã hàng của sản phẩm
    try:
        ma_hang = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_sku"))).text
    except:
        ma_hang = "Không tìm thấy"

    # lấy nhà cung cấp của sản phẩm
    try:
        nha_cung_cap = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_supplier"))).text
    except:
        nha_cung_cap = "Không tìm thấy"

    #Lấy người dịch của sản phẩm
    try:
        nguoi_dich = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_translator"))).text
    except:
        nguoi_dich = "Không tìm thấy"

    #Lấy tác giả của sản phẩm
    try:
        tac_gia = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_author"))).text
    except:
        tac_gia = "Không tìm thấy"

    #Lấy NXB của sản phẩm
    try:
        nxb = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_publisher"))).text
    except:
        nxb = "Không tìm thấy"

    #Lấy năm sản xuất của sản phẩm
    try:
        nam_xb = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_publish_year"))).text
    except:
        nam_xb = "Không tìm thấy"

    #Lấy ngôn ngữ của sản phẩm
    try:
        ngon_ngu = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_languages"))).text
    except:
        ngon_ngu = "Không tìm thấy"

    #Lấy trọng lượng sản phẩm
    try:
        trong_luong = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_weight"))).text
    except:
        trong_luong = "Không tìm thấy"

    #Lấy kích thước sản phẩm
    try:
        kich_thuoc = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_size"))).text
    except:
        kich_thuoc = "Không tìm thấy"

    #Lấy số trang sản phẩm
    try:
        so_trang = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_qty_of_page"))).text
    except:
        so_trang = "Không tìm thấy"

    #Lấy hình thức của sản phẩm
    try:
        hinh_thuc = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "data_book_layout"))).text
    except:
        hinh_thuc = "Không tìm thấy"

    # Đóng driver sau khi hoàn thành
    driver.quit()

    # Thêm chi tiết vào dictionary sản phẩm
    product.update({
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
    })
    print(f"Đã cào thông tin sản phẩm: {product['title']}")
    return product


# Hàm cào dữ liệu từ trang sản phẩm chính
def scrape_main_page():
    driver = create_driver()
    products = []
    url = 'https://www.fahasa.com/sach-trong-nuoc/van-hoc-trong-nuoc/tieu-thuyet.html?order=num_orders&limit=24&p=1'
    driver.get(url)
    time.sleep(5)

    # Dừng khi đạt 500 sản phẩm
    while len(products) < 500:
        books = driver.find_elements(By.XPATH, "//div[contains(@class,'ma-box-content')]")
        if not books:
            print("Không tìm thấy sản phẩm nào trên trang hiện tại.")
            break

        # Kiểm tra điều kiện này sau mỗi lần thêm sản phẩm
        for book in books:
            if len(products) >= 500:
                break
            try:
                link = book.find_element(By.TAG_NAME, 'h2').find_element(By.TAG_NAME, 'a').get_attribute('href')
                title = book.find_element(By.TAG_NAME, 'h2').text
                price_text = book.find_element(By.CLASS_NAME, 'special-price').text
                price = re.sub(r'[^\d]', '', price_text)

                # Kiểm tra xem sản phẩm đã có trong danh sách hay chưa
                if not any(p['link'] == link for p in products):  # Sử dụng hàm any() để kiểm tra
                    products.append({"link": link, "title": title, "price": price})
                    print(f"Đã thêm sản phẩm: {title} - Giá: {price}")

                # In ra lỗi nếu sản phẩm đã có trong danh sách
                else:
                    print(f"Sản phẩm đã tồn tại: {title}")

            # In ra lỗi khi không tìm thấy được link, title, price
            except Exception as e:
                print(f"Lỗi lấy thông tin sản phẩm: {e}")
                continue

        # Kiểm tra nếu đủ sản phẩm rồi thì thoát vòng lặp
        if len(products) >= 500:
            break

        #Nhấn nút sang trang cho đến khi không còn
        try:
            next_button = driver.find_element(By.CLASS_NAME, "icon-turn-right")
            next_button.click()
            time.sleep(5)
        except Exception as e:
            print(f"Lỗi điều hướng sang trang tiếp theo: {e}")
            break

    #Tắt driver in ra số sản phẩm đã cào được
    driver.quit()
    print(f"Tổng số sản phẩm đã cào: {len(products)}")
    return products

# Hàm chạy đa luồng để cào dữ liệu chi tiết của từng sản phẩm
def scrape_product_details(products):
    with ThreadPoolExecutor(max_workers=4) as executor:#Chạy 4 tab cùng lúc để cào dũ liệu sản phẩm về
        results = list(executor.map(get_detailed_product_info, products))
    return results

# Lưu vào file Excel
def save_to_excel(products):
    wb = Workbook()
    ws = wb.active
    ws.append(["Title", "Price", "Link", "Ma Hang", "Nha Cung Cap", "Nguoi Dich", "Tac Gia", "NXB", "Nam XB", "Ngon Ngu", "Trong Luong", "Kich Thuoc", "So Trang", "Hinh Thuc"])

    for product in products:
        ws.append([
            product['title'], product['price'], product['link'], product.get('ma_hang', ''), product.get('nha_cung_cap', ''),
            product.get('nguoi_dich', ''), product.get('tac_gia', ''), product.get('nxb', ''), product.get('nam_xb', ''),
            product.get('ngon_ngu', ''), product.get('trong_luong', ''), product.get('kich_thuoc', ''),
            product.get('so_trang', ''), product.get('hinh_thuc', '')
        ])

    wb.save("books.xlsx")
    print("Đã lưu dữ liệu vào file Excel.")

# Kết nối và lưu dữ liệu vào SQLite
def save_to_database(products):
    conn = sqlite3.connect('books.db')
    cursor = conn.cursor()
    cursor.execute('DROP TABLE IF EXISTS books')
    cursor.execute('''CREATE TABLE books(
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        title TEXT,
                        price INTEGER,
                        link TEXT,
                        ma_hang TEXT,
                        nha_cung_cap TEXT,
                        nguoi_dich TEXT,
                        tac_gia TEXT,
                        nxb TEXT,
                        nam_xb TEXT,
                        ngon_ngu TEXT,
                        trong_luong TEXT,
                        kich_thuoc TEXT,
                        so_trang TEXT,
                        hinh_thuc TEXT)''')
    conn.commit()

    for product in products:
        cursor.execute('''INSERT INTO books (title, price, link, ma_hang, nha_cung_cap, nguoi_dich, tac_gia, nxb, nam_xb, ngon_ngu, trong_luong, kich_thuoc, so_trang, hinh_thuc) 
                          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                       (product['title'], product['price'], product['link'], product.get('ma_hang', ''), product.get('nha_cung_cap', ''), product.get('nguoi_dich', ''), product.get('tac_gia', ''), product.get('nxb', ''),
                        product.get('nam_xb', ''), product.get('ngon_ngu', ''), product.get('trong_luong', ''),
                        product.get('kich_thuoc', ''), product.get('so_trang', ''), product.get('hinh_thuc', '')))
        conn.commit()

    conn.close()
    print("Đã lưu dữ liệu vào cơ sở dữ liệu SQLite.")

# Chạy các hàm để cào và lưu dữ liệu
def main():
    products = scrape_main_page()  # Cào dữ liệu sản phẩm chính
    if products:  # Kiểm tra xem có sản phẩm nào không
        detailed_products = scrape_product_details(products)  # Cào dữ liệu chi tiết với đa luồng
        save_to_excel(detailed_products)  # Lưu vào file Excel
        save_to_database(detailed_products)  # Lưu vào cơ sở dữ liệu
        print("Hoàn thành việc cào dữ liệu!")
    else:
        print("Không có sản phẩm nào được cào.")

# Bắt đầu chương trình
if __name__ == "__main__":
    main()

