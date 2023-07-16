import docx


paragraph = input("文章を入力: ")
file_name = input("ファイル名を入力してください(.docx): ")

doc = docx.Document()

doc.add_paragraph(paragraph)
doc.save(file_name)
