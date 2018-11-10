# -*- coding:utf-8 -*-
import xlwt

import timetable
import re

# 目标文件列 到 keyword的映射（或者叫做 正则表达式 关键字）
mapping = {
    3: "周数",
    4: "教师",
    5: "地点"
}

# 目标文件列 到 当前数据文件列的映射
map_temp_data = {
    0: '2',
    1: '1',
    2: '0',
    3: '3',
    4: '3',
    5: '3'
}


# 回滚函数
def rollback(in_table, i, j, cell):
    if cell == '':
        k = 0
        while in_table.cell(i - k, j).value is '':
            k = k + 1
        return in_table.cell(i - k, j).value
    else:
        return cell


# 初始化文件
def init_file():
    work_book = xlwt.Workbook()
    sheet = work_book.add_sheet("sheet1", cell_overwrite_ok=True)
    # all in title
    sheet.write(0, 0, "name")
    sheet.write(0, 1, "type")
    sheet.write(0, 2, "time")
    sheet.write(0, 3, "during")
    sheet.write(0, 4, "teacher")
    sheet.write(0, 5, "place")
    return work_book, sheet


def get_token(in_table, row, keyword):
    cell = in_table.cell(row, 3).value
    # .                 匹配任意字符
    # *                 匹配前一个元字符0到多次
    # \W                匹配非数字、字母、下划线中的任意字符
    # \S                匹配非空白字符
    # [\u4E00-\u9FA5]   匹配中文
    # +                 匹配前一个元字符1到多次
    regex = ".*(" + keyword + "\W*(\S*[\u4E00-\u9FA5]+\S*)*)"
    matches = re.match(regex, cell)
    # 结果去杂，替换 , 成 &
    result = matches.group(1).replace(",", "&")
    # 取出keyword和  :空格
    return result.replace(keyword + ": ", "")


def handle(in_table, i, j):
    k = int(map_temp_data[j])
    if k == 1:
        return "必修"
    elif k == 0:
        cell = in_table.cell(i, k).value
        # 这两个都有合并单元格，所以要回滚
        cell1 = rollback(in_table, i, k, cell).replace("星期", "周")
        cell2 = rollback(in_table, i, k + 1, in_table.cell(i, k + 1).value) + "节"
        return cell1 + cell2
    elif k < 3:
        return in_table.cell(i, k).value
    else:
        return get_token(in_table, i, mapping[j])


def save_file(in_table, filename):
    work_book, sheet = init_file()
    for j in range(len(map_temp_data)):
        for i in range(1, in_table.nrows - 3):
            sheet.write(i, j, handle(in_table, i, j))
    work_book.save(filename)


if __name__ == '__main__':
    in_file = "timetable.xlsx"
    table = timetable.read_file(in_file)
    save_file(table, "t1.xls")
