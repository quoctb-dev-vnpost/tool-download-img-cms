import os
import pandas as pd
import requests
from tqdm import tqdm
from urllib.parse import urlparse
import imghdr
from shutil import move, rmtree
from zipfile import ZipFile
from datetime import datetime

# Đọc danh sách từ tệp Excel
excel_file_path = 'duong_link.xlsx'  # Điều chỉnh tên tệp Excel của bạn
df = pd.read_excel(excel_file_path)

# Thư mục để lưu trữ các file đã tải xuống
download_folder = 'downloads'  # Điều chỉnh thư mục nếu cần thiết
os.makedirs(download_folder, exist_ok=True)

# Hàm để tải xuống file từ đường link
def download_file(url, local_filename):
    with requests.get(url, stream=True) as response:
        with open(local_filename, 'wb') as file, tqdm(
            desc=url.split("/")[-1],  # Mô tả thanh tiến trình với tên của tệp
            total=int(response.headers.get('content-length', 0)),
            unit='B',
            unit_scale=True,
            unit_divisor=1024,
        ) as bar:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    file.write(chunk)
                    bar.update(len(chunk))

# Hàm để xác định định dạng của file dựa trên nội dung
def detect_file_format(file_path):
    file_extension = imghdr.what(file_path)
    if file_extension:
        return file_extension.upper()
    else:
        return ''

# Hàm để tạo tên file từ đường link
def generate_file_name(url, index):
    parsed_url = urlparse(url)
    path_segments = parsed_url.path.split('/')
    file_name = path_segments[-1] if path_segments[-1] else path_segments[-2]
    return os.path.join(download_folder, f'file_{index + 1}_{file_name}')

# Lặp qua các đường link và tải xuống
for i, row in df.iterrows():
    link = row['Links']
    shbg_value = row['SHBG']  # Lấy giá trị của cột SHBG từ DataFrame
    
    # Bỏ qua những dòng không hợp lệ
    if not pd.notna(link) or not link.lower().startswith('http'):
        print(f'Skipping invalid link at row {i + 1}: {link}')
        continue
    
    try:
        print(f'Processing link {i + 1}: {link}')

        # Tạo thư mục tương ứng nếu chưa tồn tại
        shbg_folder = os.path.join(download_folder, str(shbg_value))
        os.makedirs(shbg_folder, exist_ok=True)

        # Tạo tên tệp với đuôi mở rộng dựa trên định dạng
        file_name = generate_file_name(link, i)
        download_file(link, file_name)

        # Đổi tên file để thêm đuôi mở rộng dựa trên định dạng
        file_format = detect_file_format(file_name)
        new_file_name = f'{file_name}.{file_format.lower()}'
        os.rename(file_name, new_file_name)

        # Di chuyển file vào thư mục tương ứng
        move(new_file_name, shbg_folder)

        print(f'Successfully downloaded: {link}, File format: {file_format.upper()}, SHBG: {shbg_value}')
    except Exception as e:
        print(f'Error downloading {link}: {e}')

# Tạo tên file zip dựa trên ngày tháng năm hiện tại
current_date = datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
zip_file_name = f'Attach_CMS_{current_date}.zip'

print(f'Zip file created: {zip_file_name}')
print('Download, organization, and compression completed.')
