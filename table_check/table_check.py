import docx


doc = docx.Document("TEST用起案文書様式.docx")

num = 0
for tbl in doc.tables:
    num = num + 1
    print(num, "行数=", len(tbl.rows), "列数=", len(tbl.columns))
