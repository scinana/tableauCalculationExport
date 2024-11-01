import os, pathlib
import win32com.client
import pandas as pd


def create_new_file_paths(tableau_name_substring):

    cwd = os.getcwd()
    path_string = pathlib.Path(cwd).resolve().__str__() + "\{}"

    print(path_string)

    newFileName = 'outputs\{}'.format(tableau_name_substring)

    excel_path = path_string.format(newFileName + ".xlsx")
    path_to_pdf = path_string.format(newFileName + ".pdf")

    print(excel_path)
    print(path_to_pdf)

    return (excel_path, path_to_pdf)


def mainCol(colNumber, color, writer, sheetName):

    workbook = writer.book
    worksheet = writer.sheets[sheetName]

    format_mainCol = workbook.add_format({'text_wrap': True, 'bold': True})
    format_mainCol.set_align('vcenter')
    format_mainCol.set_bg_color(color)
    format_mainCol.set_border(1)
    worksheet.set_column(colNumber,colNumber,20,format_mainCol)
    return worksheet


def normalCol(colNumber, colWidth, writer, sheetName):

    workbook = writer.book
    worksheet = writer.sheets[sheetName]


    format2 = workbook.add_format({'text_wrap': True})
    format2.set_align('vcenter')
    format2.set_border(1)
    worksheet.set_column(colNumber,colNumber,colWidth,format2)
    return worksheet


def create_excel_from_dfs(dfs_to_use, excel_path):

    writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
    
    # input: any number of dfs
    # output: an excel file with one excel sheet per df

    # code to create each sheet in excel, with the specified df and formatting each sheet as per requirements
    # also adds a header and footer to each sheet
    # all the info to be replaced below (ie. for each df) comes form the dfs_to_use list of dictionaries

    for x in dfs_to_use:
        excelSheetTitle = x['excelSheetTitle']
        df_to_use = x['df_to_use']
        normalColWidth = x['normalColWidth']
        sheetName = x['sheetName']
        papersize = x['papersize']
        footer = x['footer']
        color = x['color']

        df_to_use.to_excel(writer, sheet_name=sheetName, index=False)

        worksheet = mainCol(colNumber = 0, color = color, writer=writer, sheetName=sheetName)

        ws = 1
        for i in normalColWidth: #iterates through each column
            worksheet = normalCol(ws, i, writer=writer, sheetName=sheetName)
            ws = ws + 1

        worksheet.set_paper(papersize)  # a4
        worksheet.fit_to_pages(1, 0)  # fit to 1 page wide, n long
        worksheet.repeat_rows(0)  # repeat the first row

        header_x = '&C&"Arial,Bold"&10{}'.format(excelSheetTitle)
        footer_x = '&L{}&CPage &P of &N'.format(footer)

        worksheet.set_header(header_x)
        worksheet.set_footer(footer_x)

    #writer.save()
    writer.close()


def create_pdf_from_excel(path_excel, path_pdf, dfs_to_use):


    # this creates an index to list each excel sheet, based on the number of sheets that were created before

    for_ws_index_list = []
    for i in range(len(dfs_to_use)):
        for_ws_index_list.append(i + 1)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Open(path_excel)

    #print all the excel sheets into a single pdf
    ws_index_list = for_ws_index_list
    wb.Worksheets(ws_index_list).Select()
    wb.ActiveSheet.ExportAsFixedFormat(0, path_pdf)
    wb.Close()
    excel.Quit()
