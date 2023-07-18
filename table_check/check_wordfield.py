import docx


doc = docx.Document("TEST用起案文書様式.docx")

tbl = doc.tables[0]
for row in tbl.rows:
    row_text = []
    for cell in row.cells:
        row_text.append(cell.text)
    print(row_text)
