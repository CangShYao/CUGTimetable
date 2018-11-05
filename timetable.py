import xlrd
import xlwt


# read file and return a table
def read_file(file_path):
    xlrd.Book.encoding = "utf8"
    data = xlrd.open_workbook(file_path)
    tem_table = data.sheet_by_index(0)
    return tem_table


def handle_complex(in_cell):
    cells = in_cell.split(' ')
    # I need 01 07 and 10
    # return a list
    return [cells[1], cells[7], cells[10]]


def handle_table(in_table):
    list_table = []
    for j in range(in_table.ncols - 1):
        list_table.append([])
        for i in range(1, in_table.nrows - 3):
            cell = in_table.cell(i, j).value
            # cell(row, col) can't distinguish merge cell
            # so we need to back to the first
            if cell is '':
                k = 0
                while in_table.cell(i - k, j).value is '':
                    k = k + 1
                cell = in_table.cell(i - k, j).value
            if j == 0:
                # 星期 replace to 周
                # list_table[j].append(cell.replace("星期", "周"))
                # maybe good
                cell = "周" + cell[2]
                list_table[j].append(cell)
                continue
            if j == 1:
                # 7-8 to 7-8节
                # list_table[j].append(cell+"节")
                cell = cell + "节"
                list_table[j].append(cell)
                continue
            if j == 2:
                # this what?
                list_table[j].append(cell)
                continue
            if j == 3:
                # This part is very complicated.
                handed_cell = handle_complex(cell)
                for z in range(len(handed_cell)):
                    if z == 0:
                        tem = handed_cell[z].replace(",", "&")  # 1-6周&8-11周
                        list_table[j].append(tem)
                        continue
                    else:
                        if i == 1:
                            list_table.append([])
                        list_table[j + z].append(handed_cell[z])
    return list_table


def write_file(result_matrix):
    work_book = xlwt.Workbook()
    sheet = work_book.add_sheet("sheet1", cell_overwrite_ok=True)
    # all in title
    sheet.write(0, 0, "name")
    sheet.write(0, 1, "type")
    sheet.write(0, 2, "time")
    sheet.write(0, 3, "during")
    sheet.write(0, 4, "teacher")
    sheet.write(0, 5, "place")
    # write content
    # I think structure more convenient
    for y in range(0, len(result_matrix[0])):
        sheet.write(y + 1, 0, result_matrix[2][y])
        sheet.write(y + 1, 1, "必修")
        sheet.write(y + 1, 2, result_matrix[0][y] + result_matrix[1][y])
        sheet.write(y + 1, 3, result_matrix[3][y])
        sheet.write(y + 1, 4, result_matrix[5][y])
        sheet.write(y + 1, 5, result_matrix[4][y])
    work_book.save("t.xls")


if __name__ == '__main__':
    in_file = "timetable.xlsx"
    table = read_file(in_file)
    result = handle_table(table)
    write_file(result)
