from pandas import ExcelFile, read_excel
from re import sub

def Excel_file(path, name, target):
    xl = ExcelFile(path).sheet_names
    for sheet in xl:
        df = read_excel(path,sheetname=sheet).applymap(str)
        if not df.empty:
            i = 0
            for c in df:
                length = str(df[c].map(len).max())
                column = str(c)
                if i == 0:
                    tmp = sub("[^a-zA-Z\d-]", "", name)
                    tablename = (tmp.split(".", 1)[0]+sheet.replace(" ","_"))
                    target.write("\n\n")
                    target.write(
                        "IMPORT EXCEL TO %s \"FIL\\%s.fil\" FROM \"%s\""
                        " TABLE \"%s\" KEEPTITLE FIELD \"%s\" C WID %s"
                        " AS \"\" "
                        % (tablename, tablename, path, sheet, column, length))
                    i = i + 1
                else:
                    target.write('FIELD "%s" C WID %s AS "" ' %(column, length))
