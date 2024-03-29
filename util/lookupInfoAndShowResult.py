import xlwings as xw


def copySourceData(source_sheet, comparison_sheet, start_row, end_row, src_end_row):
    # Specify the source range (C4:F41) and target range on the 'comparison' sheet
    print(f"copy source data: {source_sheet, comparison_sheet, start_row, end_row, src_end_row}")
    source_range = source_sheet.range(f'K15:Q{src_end_row}')
    target_range = comparison_sheet.range(f'B{start_row}')

    source_range.api.Copy()
    target_range.api.PasteSpecial()

    # Clear the clipboard to avoid Excel freezing issues
    xw.apps.active.api.CutCopyMode = 0

    # change cell color to default (None)
    comparison_sheet.range(f'B{start_row+1}:H{end_row}').color = None

    return


def createFetchDataTable(file_path, comparison_sheet, column_info_dict, start_row, end_row):
    file_name = file_path.split('/')[2][:-5]
    comparison_sheet.range(f'B{start_row-1}').add_hyperlink(file_path, text_to_display=file_name)

    header_source_range = comparison_sheet.range(f'B{start_row}:H{end_row}')
    header_target_range = comparison_sheet.range(f'J{start_row}')
    header_source_range.api.Copy()
    header_target_range.api.PasteSpecial()

    # delete table values
    comparison_sheet.range(f'J{start_row+1}:P{end_row}').value = ""

    # Clear the clipboard to avoid Excel freezing issues
    xw.apps.active.api.CutCopyMode = 0
    return


def isSame(a, b):
    if type(a) is float:
        a = int(a)
    if type(b) is float:
        b = int(b)

    a, b = str(a).strip().upper(), str(b).strip().upper()
    strings = ("VARCHAR", "VARCHAR2")
    numbers = ("NUMBER", "INT", "DECIMAL")
    blanks = ("", None, "NONE")

    if a == b:
        return True
    elif a in strings and b in strings:
        return True
    elif a in numbers and b in numbers:
        return True
    elif a in blanks and b in blanks:
        return True

    # print(f"a: {a} / b: {b}")
    return False


def doLookup(sheet, column_info_dict, start_row, end_row):

    src_cols = ["B", "C", "E", "F", "G", "H"]
    db_cols = ["J", "K", "M", "N", "O", "P"]
    note_col = "Q"

    data_mapping = {
        "J" : "COLUMN_NAME",
        "K" : "COLUMN_COMMENT",
        "M" : "DATA_TYPE",
        "N" : "CHARACTER_MAXIMUM_LENGTH",
        "N2" : "NUMERIC_PRECISION",      # decimal type인 경우 CHARACTER_MAXIMUM_LENGTH 대신 사용
        "O" : "IS_NOT_NULLABLE",    # 필수여부
        "P" : "IS_PRIMARY_KEY",    # PK
    }

    table_row_color = (128, 128, 128)   # grey
    alert_color = (255, 0, 0)        # red
    warning_color = (255, 255, 0)    # yellow  

    table = None
    for row in range(start_row + 1, end_row + 1):
        r = str(row)
        src_key_cell = sheet.range(src_cols[0] + r)
        db_key_cell = sheet.range(db_cols[0] + r)
        src_cells = sheet.range(f"{src_cols[0]}{r}:{src_cols[-1]}{r}")
        db_cells = sheet.range(f"{db_cols[0]}{r}:{db_cols[-1]}{r}")

        table_names = list(column_info_dict.keys())
        
        src_key = src_key_cell.value

        if src_key is None:                        # 빈 값일 때
            continue
        elif src_key.upper() in table_names:       # 테이블명일 때
            src_cells.color = table_row_color
            src_cells.api.Copy()
            db_key_cell.api.PasteSpecial()

            # Clear the clipboard to avoid Excel freezing issues
            xw.apps.active.api.CutCopyMode = 0

            # set table
            table = src_key
        else:                                                 # 컬럼명일 때
            if src_key not in column_info_dict[table].keys():
                db_key_cell.value = src_key
                sheet.range(db_cols[1] + r).value = "(정보 없음)"
                db_cells.color = warning_color
                print(f"!!! 정보 없음 - {src_key}")
                continue

            target_info = column_info_dict[table][src_key]

            if len(target_info) == 0:                   # 일치하는 정보가 없을 때
                print(f"***No target Info : table - {table} / col - {src_key}")
            else:                                       # 일치하는 정보가 있을 때 -> 기존 문서 데이터와 비교
                for i in range(len(db_cols)):
                    target_info_key = data_mapping[db_cols[i]]
                    val = target_info[target_info_key]
                    if val is None and i > 0 and target_info[data_mapping[db_cols[i-1]]]:    # decimal 타입인 경우
                        val = target_info[data_mapping["N2"]]
                    src_cell = sheet.range(src_cols[i] + r)
                    target_cell = sheet.range(db_cols[i] + r)
                    src, tgt = sheet.range(src_cols[i] + r).value, val.strip() if isinstance(val, str) else val
                    target_cell.value = tgt
                    if src is None and tgt != '' and tgt is not None:
                        src_cell.color = warning_color
                        target_cell.color = warning_color
                        print(f"!!!! src: {src} / tgt: {tgt}")
                    elif not isSame(src, tgt):
                        src_cell.color = alert_color
                        target_cell.color = alert_color
                        print(f"XXXX src: {src} / tgt: {tgt}")
    return

    


def lookupInfoAndShowResult(file_path, comparison_sheet, cur_row, wb, ws, query_row, column_info_dict):
    src_end_row = query_row - 3     # 대상 엑셀의 복사해올 마지막 줄 위치
    start_row = cur_row + 5     # 시작줄(헤더)
    end_row = start_row + src_end_row - 15   #  마지막줄(표가 끝나는 줄)

    copySourceData(ws, comparison_sheet, start_row, end_row, src_end_row)

    # 타겟 엑셀 닫기
    wb.close()

    createFetchDataTable(file_path, comparison_sheet, column_info_dict, start_row, end_row)

    doLookup(comparison_sheet, column_info_dict, start_row, end_row)

    return end_row # 마지막줄 리턴 후 다음줄부터 돌리기