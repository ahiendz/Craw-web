import csv
import os
import requests
import pandas as pd
from bs4 import BeautifulSoup

def fetch_courses(url):
    """Gửi yêu cầu GET tới URL và trả về nội dung HTML nếu thành công."""
    try:
        response = requests.get(url)
        response.raise_for_status()  # Kiểm tra lỗi HTTP
        return response.text
    except requests.RequestException as e:
        print(f"Lỗi khi truy cập trang web: {e}")
        return None

def parse_courses(html):
    """Phân tích HTML và trả về danh sách tiêu đề, thông tin và danh mục."""
    soup = BeautifulSoup(html, 'html.parser')
    title_link = soup.select("h2.course-loop-title.course-loop-title-collapse-2-rows a")
    info = soup.select("div.course-loop-meta-list .meta-value")
    category = soup.select(".course-loop-category a")
    
    inFo = [k.get_text(strip=True) for k in info]
    categorys = [value.get_text(strip=True) for value in category]
    
    return title_link, inFo, categorys

def save_to_csv(fn, title_link, inFo, categorys):
    """Ghi dữ liệu vào file CSV."""
    with open(fn, mode='w', encoding='utf-8', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['STT', 'Title', 'Link', 'Category', 'Lessons', 'Hours', 'Level'])
        
        for index, title in enumerate(title_link):
            l = inFo[index * 3:(index + 1) * 3]  # Lấy 3 giá trị cho mỗi khóa học
            row = [index + 1, title.get_text(strip=True), title["href"], categorys[index], *l]
            writer.writerow(row)

def convert_csv_to_excel(csv_file, excel_file):
    """Đọc file CSV và lưu dưới dạng file Excel, sau đó xóa file CSV."""
    try:
        df = pd.read_csv(csv_file, encoding='utf-8')
        df.to_excel(excel_file, index=False)
        os.remove(csv_file)
    except Exception as e:
        print(f"Lỗi khi chuyển đổi CSV sang Excel: {e}")

def main():
    base_url = "https://tuhoc.cc/home-page/courses"
    pages = [1, 2]  # Số trang cần lấy dữ liệu
    all_title_link = []
    all_inFo = []
    all_categorys = []

    for page in pages:
        url = f"{base_url}/page/{page}/"
        html = fetch_courses(url)
        
        if html:
            title_link, inFo, categorys = parse_courses(html)
            all_title_link.extend(title_link)
            all_inFo.extend(inFo)
            all_categorys.extend(categorys)

    # Kiểm tra tính hợp lệ của dữ liệu
    if len(all_title_link) == len(all_categorys) == len(all_inFo) // 3:
        csv_file = "Khóa học Tuhoc.cc.csv"
        save_to_csv(csv_file, all_title_link, all_inFo, all_categorys)
        convert_csv_to_excel(csv_file, 'Khóa học Tuhoc.cc.xlsx')
    else:
        print("Dữ liệu không hợp lệ: Vui lòng kiểm tra lại.")

if __name__ == "__main__":
    main()
