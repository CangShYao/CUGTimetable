import xlrd
import xlwt


# read file and return a table
def read_file(file_path):
    xlrd.Book.encoding = "utf8"
    data = xlrd.open_workbook(file_path)
    tem_table = data.sheet_by_index(0)
    return tem_table


if __name__ == '__main__':
    in_file = "timetable.xlsx"
    table = read_file(in_file)
    print(table.cell(1, 1).value)
