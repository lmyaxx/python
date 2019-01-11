import xlwings as xw
import os
import shutil
import re

app = xw.App(visible=False, add_book=False)
app.display_alerts = False
app.screen_updating = False


def generate_report_for_one_type(file_from, file_to, group_name):
    cur_path = get_cur_path()
    file_source = cur_path + "/" + file_from
    file_target = cur_path + "/" + file_to
    del_dir_and_file(file_target)
    generate_report_by_group_name(file_source, file_target, group_name)


def get_cur_path():
    return os.path.abspath('.')


def del_dir_and_file(path):
    if os.path.isdir(path):
        shutil.rmtree(path)


def generate_report_by_group_name(source_dir, target, group_name):
    li = get_file_paths_from_dir(source_dir)
    for source_path in li:
        name, suffix = get_file_name_by_path(source_path)
        print("处理"+name+suffix+"中")
        group_name_set, various_group_name_pos = get_info_from_source(source_path, group_name)
        copy_file_to_new_dir_by_lsp(source_path, group_name_set, various_group_name_pos, target)
        print(name+suffix+"处理成功\n")


def get_file_paths_from_dir(directory):
    result = []
    if os.path.exists(directory):
        path = os.listdir(directory)
        for p in path:
            if p.endswith(".xls") or p.endswith(".xlsx"):
                result.append(directory + '/' + p)
    return result


def get_file_name_by_path(path):
    (name, suffix) = os.path.splitext(os.path.split(path)[1])
    return name, suffix


# x, y为lsp_name 对应标题的位置
def get_info_from_source(source_path, group_name):
    wb = app.books.open(source_path)
    sheet = wb.sheets[0]
    cells = sheet.cells

    row, col = get_pos_row_col(cells, group_name)
    if row == -1:
        print("未识别到相应分组标识")
        exit(1)

    column = cells.columns[col]
    li = column.value
    # 数据行开始至结束，从第start+1行至height行
    start = row + 2
    height = li.index(None, start)
    # 对数据区根据需要分组的title进行排序
    sheet.range(str(start+1)+':'+str(height)).api.Sort(Key1=column.api, Order1=1)
    wb.save()
    # 获取排序后的title对应数据列的值
    li = wb.sheets[0].cells.columns[col].value
    group_name_set = set(li[row+2: height+1])
    group_name_set.remove(None)
    # 放置需要group_name的下标范围
    various_group_name_pos = []
    div_start = start
    for i in range(row+2, height):
        if li[div_start] == li[i+1]:
            continue
        else:
            various_group_name_pos.insert(0, [li[div_start], [div_start, i]])
            div_start = i + 1
    wb.close()
    return group_name_set, various_group_name_pos


# 定位lsp与lsp name 的位置，下标从0开始，即sheet左上角第一个元素为0，0
def get_pos_row_col(cells, group_name):
    rows = cells.rows
    row_m = min(rows.count, 20)
    col_m = cells.columns.count
    for row_n in range(0, row_m):
        li = rows[row_n].value
        for col in range(0, col_m):
            name = li[col]
            if name in group_name:
                return row_n, col
    return -1, -1


def copy_file_to_new_dir_by_lsp(source_path, lsp_set, vary_lsp_pos, directory):
    name, suffix = get_file_name_by_path(source_path)
    if os.path.exists(directory):
        pass
    else:
        os.mkdir(directory)
    # 生成新
    for lsp_name in lsp_set:
        new_lsp_name = validate_title(lsp_name)
        new_dir = directory + "/" + new_lsp_name
        if os.path.exists(new_dir):
            # del_file(new_dir)
            pass
        else:
            os.mkdir(new_dir)

        new_file = new_dir + "/" + name + suffix
        shutil.copyfile(source_path, new_file)
        omit_data_by_group_name(new_file, lsp_name, vary_lsp_pos)


def validate_title(name):
    rstr = r"[\/\\\:\*\?\"\<\>\|]"  # '/ \ : * ? " < > |'
    new_name = re.sub(rstr, "_", name)  # 替换为下划线
    i = len(new_name)
    while i > 0 and (name[i - 1] == '.' or name[i - 1] == ' '):
        i = i - 1
    return new_name[:i]


def omit_data_by_group_name(file_path, lsp_name, vary_lsp_pos):
    print("  生成" + lsp_name)
    wb = app.books.open(file_path)
    sheet = wb.sheets[0]
    for name, [start, terminate] in vary_lsp_pos:
        if name != lsp_name:
            sheet.range(str(start+1)+":"+str(terminate+1)).api.EntireRow.Delete()
    wb.save()
    wb.close()


if __name__ == '__main__':
    generate_report_for_one_type("GI_GR", "GI_GR_Target", ["Company"])
    generate_report_for_one_type("source", "target", ["LSP", "LSP Name"])
    app.quit()
