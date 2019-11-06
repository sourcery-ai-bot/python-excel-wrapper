import xlrd
import csv
import time
from pyexcelerate import Workbook
import os
import hashlib

tmp_dir = './tmp'


def xlsx_to_csv(wk_path, ws_name="Sheet1"):
    wk_name = str(os.path.basename(wk_path))
    csv_filename = tmp_dir + "/" + wk_name.split('.')[0] + '.csv'
    wb = xlrd.open_workbook(wk_path)
    ws = wb.sheet_by_index(0)
    csv_file = open(csv_filename, 'w')
    wr = csv.writer(csv_file)
    for rownum in range(ws.nrows):
        wr.writerow(ws.row_values(rownum))
    csv_file.close()
    return csv_filename


def generate_csv(wk_path, ws_name="Sheet1"):
    csv_path = xlsx_to_csv(wk_path, ws_name)
    return list(csv.reader(open(csv_path)))


def md5(file_path):
    hash_md5 = hashlib.md5()
    with open(file_path, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()


def create_directory_if_needed(directory_path):
    if not os.path.isdir(directory_path):
        os.mkdir(directory_path)


def create_file(file_path):
    if not os.path.exists(file_path):
        with open(file_path, 'w'):
            pass
    return file_path


def get_saved_hash(file_path):
    with open(file_path, 'r') as file:
        saved_hash = file.readline()
    return saved_hash


def write_hash(file_path, hash):
    with open(file_path, 'w') as file:
        file.write(hash)


def get_worksheet_data(file_hash_save_path, current_hash, wk_source_path, file_source_name):
    saved_hash = get_saved_hash(file_hash_save_path)
    if saved_hash:
        if current_hash != saved_hash:
            csv_path = xlsx_to_csv(wk_source_path)
        else:
            csv_path = tmp_dir + '/' + file_source_name + '.csv'
            if not os.path.isfile(csv_path):
                csv_path = xlsx_to_csv(wk_source_path)
        ws_source = list(csv.reader(open(csv_path)))
    else:
        ws_source = generate_csv(wk_source_path)
    return ws_source


def get_saved_hash_file(saved_hash_file_path):
    if not os.path.exists(saved_hash_file_path):
        create_file(saved_hash_file_path)
    return


def init_worksheet(wk_source_path):
    file_source_name = str((os.path.basename(wk_source_path)).split('.')[0])
    current_hash = md5(wk_source_path)
    create_directory_if_needed(tmp_dir)
    saved_hash_file_path = create_file(tmp_dir + '/' + file_source_name + '.txt')
    ws_source = get_worksheet_data(saved_hash_file_path, current_hash, wk_source_path, file_source_name)
    write_hash(saved_hash_file_path, current_hash)
    return ws_source


def main():
    wk_source_path = './data_5000.xlsx'
    wk_target_path = './data_target.xlsx'

    ws_source = init_worksheet(wk_source_path)

    wb_target = Workbook()
    ws_target = wb_target.new_sheet("Sheet1")

    first_row_source = 2
    last_row_source = 6842
    row_target = 5

    # l'index de la source commence à 0, celui du target à 1
    source_to_target_mapper = {
        0: 1,
        1: 2,
        9: 2,
        17: 3,
        4: 4,
        12: 5
    }

    for line_source in range(first_row_source, last_row_source):
        for key, value in source_to_target_mapper.items():
            ws_target[row_target][value].value = ws_source[line_source][key]
        row_target = row_target + 1
    wb_target.save(wk_target_path)


if __name__ == "__main__":
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
