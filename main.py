import sys

from util.fetchColumnInfo import fetchColumnInfo
from util.getTargetCols import getTargetTableAndCols
from util.lookupInfoAndShowResult import lookupInfoAndShowResult
from util.getFileNameList import getFileNameList
from util.getFileProperties import getFileProperties
from pprint import pprint
import xlwings as xw
import datetime as dt

folder_path = './excels' 
# comp_book = None
# comp_sheet = None
# row = 0


def prepareSheetForComparison():
    now = dt.datetime.now()
    now_formatted = now.strftime('%Y%m%d%H%M%S')
    comp_book = xw.Book()
    comp_sheet = comp_book.sheets.add("비교")
    comp_book.save(f'문서_DB_스키마_비교_{now_formatted}.xlsx')
    return comp_book, comp_sheet


def main():
    # Check if at least one parameter was provided
    param = None
    if len(sys.argv) >= 2:
        param = sys.argv[1]

    # prepare wb & ws for comparison
    comp_book, comp_sheet = prepareSheetForComparison()
    row = 0

    # get target Files
    files = getFileNameList(folder_path) if param is None else [f"{folder_path}/{param}"]

    for file_name in files:
        print(f"[START] {file_name} -------------")
        # (wb, ws, endRow)
        target_wb, target_ws, end_row = getFileProperties(file_name)

        comp_sheet.range('B1').value = "기존 문서"
        comp_sheet.range('J1').value = "실제 DB"
        # comp_sheet.range('1:1').color = (255, 223, 186)
        # active_window = comp_book.app.api.ActiveWindow
        # # 첫 행 틀 고정
        # # active_window.FreezePanes = False
        # active_window.SplitRow = 0
        # active_window.FreezePanes = True

        if end_row is None:   
            print("문서 양식을 확인해주세요")
        else:
            target_cols = getTargetTableAndCols(target_wb, target_ws, end_row)

            ### fetching column information from DB
            column_info_dict = fetchColumnInfo(target_cols)
            pprint(column_info_dict)

            ### write result of comparing on new excel sheet ('비교')
            row = lookupInfoAndShowResult(file_name, comp_sheet, row, target_wb, target_ws, end_row, column_info_dict)

        # target_wb.close()
        print(f"[END] {file_name} -------------")





if __name__ == "__main__":
    main()