import os
import requests
import zipfile


def get_zip_path():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(current_dir, './converted_accdb.zip')


def convert_accdb_to_xlsx(file_path):
    files = {
        'files[]': open(file_path, 'rb')
    }
    print('Converting...')
    r = requests.post('https://www.rebasedata.com/api/v1/convert?outputFormat=xlsx&errorResponse=json', files=files)

    zip_path = get_zip_path()
    with open(zip_path, 'wb') as local_file:
        local_file.write(r.content)
    return zip_path


def unzip_files(zip_path):
    dir_path = os.path.join(os.path.dirname(zip_path), './converted_accdb')
    with zipfile.ZipFile(zip_path, 'r') as output_zip:
        output_zip.extractall(dir_path)
    return dir_path
