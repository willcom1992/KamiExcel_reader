import openpyxl


# book:吸いだしたいシートのあるブック名, sheet:吸いだし元となるシート名
# suidashi():{日付: [[予約情報],[予約情報]...[予約情報]]}の形の辞書のリストを返す

def suidashi(book, sheet):
    print('ワークブックを開きます．．．')

    wb = openpyxl.load_workbook(book)
    sheet = wb[sheet]
    # 取得したい行の範囲を指定（全て2列～13列）
    row_area = [[2, 33], [43, 74], [82, 113], [121, 151], [155, 186]]
    # 取得したい列の範囲を指定
    column_area = [2, 13]

    # 吸いだし作業
    all_list = []
    for i in range(0, len(row_area)):
        start_row = row_area[i][0]
        last_row = row_area[i][1]
        # 　列ごとに吸い出す
        area_list = []
        for columns in list(sheet.columns):
            if not column_area[0] <= columns[0].column <= column_area[1]:  # 条件に合わない範囲の列は飛ばす
                continue
            # 　列から行ごとに吸い出す
            rows_list = []
            for cell in columns:
                if start_row <= cell.row <= last_row:  # 条件に合う範囲のセルのみ取得
                    rows_list.append(cell.value)
            area_list.append(rows_list)
        all_list.append(area_list)

    # 吸いだしたリストの整形作業1周目
    all_list2 = []
    for area in all_list:
        area_list2 = []
        for column_lis in area:
            dic_list = []
            for n in range(0, len(column_lis) - 1)[::2]:
                lis = []
                if n == 0:
                    dic_list.append(column_lis[n])
                else:
                    lis.append(column_lis[n])
                    lis.append(column_lis[n + 1])
                    dic_list.append(lis)
            area_list2.append(dic_list)
        all_list2.append(area_list2)

    # 吸いだしたリストの整形作業2周目
    all_list3 = []
    for area in all_list2:
        area_list3 = []
        for i in range(0, len(area) - 1)[::2]:
            area_list3.append(list(zip(area[i], area[i + 1])))
        all_list3.append(area_list3)

    # 吸いだしたリストの整形作業3周目(これで最後)
    all_list4 = []
    for area in all_list3:
        for column in area:
            master_data = {}
            cells_list = []
            for cell in column[1:]:
                name = cell[0][0]
                school = cell[0][1]
                memo = cell[1][0]
                time = cell[1][1]
                lis = [name, school, memo, time]
                cells_list.append(lis)
            master_data[column[0][0]] = cells_list
            if list(master_data.keys())[0] is not None:
                master_data[list(master_data.keys())[0]] = cells_list
                all_list4.append(master_data)

    return all_list4


# suidashi()で得られたリスト(data_list)を、['日付', 'name', 'school', 'memo', 'time' ]のリストに変換
def to_list(data_list):
    reserve_list = []
    for dic in data_list:
        for k, v in dic.items():
            for i in v:
                lis = [k, *i]
                reserve_list.append(lis)
    return reserve_list


# to_list()で得られたリスト(reserve_lis)を、bookの各行に書込む
def write_excel(reserve_lis, book):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in reserve_lis:
        ws.append(row)
    wb.save(book)
    print('書込み完了')


if __name__ == "__main__":
    # 'SampleBook.xlsx'の'SampleSheet'からデータを吸い出し、整形した後、'SampleBook_transform.xlsx'を新規作成し書き込む
    write_excel(to_list(suidashi('SampleBook.xlsx', 'SampleSheet')), 'SampleBook_transform.xlsx')
