import requests  # Thư viện để gửi yêu cầu HTTP
import csv  # Thư viện để làm việc với tệp CSV
import os  # Thư viện để xử lý các thao tác với tệp và thư mục
import pandas as pd  # Thư viện để xử lý và phân tích dữ liệu
from bs4 import BeautifulSoup  # Thư viện để phân tích cú pháp HTML
import time  # Thư viện để sử dụng các hàm thời gian

# Hàm để thu thập dữ liệu từ trang web
def scrape_quotes(page_count):
    url_base = "https://quotes.toscrape.com"  # Địa chỉ cơ sở của trang web
    
    # Mở tệp CSV để ghi dữ liệu
    with open("Data.csv", "w", encoding='utf-8', newline='') as file:
        writer = csv.writer(file)  # Tạo đối tượng ghi vào tệp CSV
        writer.writerow(["quote", "author", "tags"])  # Ghi tiêu đề cho các cột
        
        # Vòng lặp qua số trang đã chỉ định
        for page in range(1, page_count + 1):
            url = f"{url_base}/page/{page}"  # Tạo URL cho trang hiện tại
            try:
                response = requests.get(url)  # Gửi yêu cầu GET đến URL
                response.raise_for_status()  # Kiểm tra mã trạng thái phản hồi
            except requests.RequestException as e:
                print(f"Error fetching page {page}: {e}")  # In lỗi nếu có
                continue  # Tiếp tục với trang tiếp theo
            
            soup = BeautifulSoup(response.content, 'html.parser')  # Phân tích cú pháp HTML
            quotes = soup.find_all("div", class_="quote")  # Tìm tất cả các trích dẫn
            
            # Vòng lặp qua từng trích dẫn
            for quote in quotes:
                text = quote.find("span", class_="text").text  # Lấy nội dung trích dẫn
                author = quote.find("small", class_="author").text  # Lấy tên tác giả
                
                tags = []  # Khởi tạo danh sách chứa các tag
                tags_div = quote.find("div", class_="tags")  # Tìm phần chứa các tag
                if tags_div:
                    all_tag_a = tags_div.find_all('a')  # Lấy tất cả các thẻ <a> trong phần tags
                    tags = [tag.get_text() for tag in all_tag_a]  # Lấy nội dung của các tag
                
                writer.writerow([text, author, ", ".join(tags)])  # Ghi dữ liệu vào tệp CSV
            
            time.sleep(1)  # Tạm dừng 1 giây để tránh quá tải máy chủ

# Hàm để chuyển đổi tệp CSV thành tệp Excel
def convert_to_excel():
    df = pd.read_csv("Data.csv", encoding='utf-8')  # Đọc tệp CSV vào DataFrame
    df.to_excel("Data.xlsx", index=False)  # Lưu DataFrame vào tệp Excel
    os.remove("Data.csv")  # Xóa tệp CSV sau khi chuyển đổi

if __name__ == "__main__":
    page_count = 10  # Số trang cần thu thập dữ liệu
    scrape_quotes(page_count)  # Gọi hàm thu thập dữ liệu
    convert_to_excel()  # Gọi hàm chuyển đổi tệp sang Excel
