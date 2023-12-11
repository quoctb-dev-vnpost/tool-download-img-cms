import os
import pandas as pd
import requests
from tqdm import tqdm
from urllib.parse import urlparse
from openpyxl import load_workbook
from shutil import move, rmtree
from zipfile import ZipFile
from datetime import datetime
import shutil
from tempfile import mkdtemp

# Đọc danh sách từ tệp Excel
excel_file_path = 'duong_link.xlsx'
df = pd.read_excel(excel_file_path)

# Thư mục để lưu trữ các file đã tải xuống
download_folder = 'downloads'
os.makedirs(download_folder, exist_ok=True)

# Hàm để tải xuống file từ đường link
def download_file(url, local_filename):
    with requests.get(url, stream=True) as response:
        with open(local_filename, 'wb') as file, tqdm(
            desc=url.split("/")[-1],
            total=int(response.headers.get('content-length', 0)),
            unit='B',
            unit_scale=True,
            unit_divisor=1024,
        ) as bar:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    file.write(chunk)
                    bar.update(len(chunk))

# Hàm để tạo tên file từ đường link
def generate_file_name(url, index):
    parsed_url = urlparse(url)
    path_segments = parsed_url.path.split('/')
    file_name = path_segments[-1] if path_segments[-1] else path_segments[-2]
    return os.path.join(download_folder, f'file_{index + 1}_{file_name}')

# Lặp qua các đường link và tải xuống
for i, row in df.iterrows():
    link = row['Links']
    shbg_value = row['SHBG']
    
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
        file_format = os.path.splitext(file_name)[1].lower()
        new_file_name = f'{file_name}.{file_format}'
        os.rename(file_name, new_file_name)

        # Di chuyển file vào thư mục tương ứng
        move(new_file_name, shbg_folder)

        # Tạo hyperlink và chèn vào cột "File"
        hyperlink_path = f'./{shbg_folder}/{os.path.basename(new_file_name)}'
        hyperlink_formula = f'=HYPERLINK("{hyperlink_path}", "View file")' #khi chạy tool thì bỏ view file để Vlookup
        df.at[i, 'File'] = hyperlink_formula

        print(f'Successfully downloaded: {link}, File format: {file_format.upper()}, SHBG: {shbg_value}')
    except Exception as e:
        print(f'Error downloading {link}: {e}')

# Tạo tên file zip dựa trên ngày tháng năm hiện tại
current_date = datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
zip_file_name = f'Send_Attach_CMS_{current_date}.zip'

# Cập nhật tệp Excel với các giá trị mới trong cột "File"
df.to_excel(excel_file_path, index=False)

temp_dir = mkdtemp()

try:
    # Di chuyển nội dung thư mục downloads vào thư mục tạm thời
    shutil.move(download_folder, temp_dir)

    # Di chuyển tệp Excel vào thư mục tạm thời
    shutil.move(excel_file_path, temp_dir)

    # Nén thư mục tạm thời thành một file zip
    shutil.make_archive(zip_file_name, 'zip', temp_dir)

    print(f'Zip file created: {zip_file_name}.zip')
    print('Download, organization, and compression completed.')
finally:
    # Dọn dẹp thư mục tạm thời sau khi đã sử dụng xong
    rmtree(temp_dir)

print(f'Zip file created: {zip_file_name}')
print('Download, organization, and compression completed.')

