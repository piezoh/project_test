from docx import Document


def main():
    new_filename = input("作成するファイル名を入力: ")
    draft_content = input("起案内容を入力:")

    # 既存のWord文書を開く
    document = Document("TEST用起案文書様式.docx")

    # 文書内の表を取得する
    table = document.tables[0]

    # 表の6行目、0列目にあるセルを取得する
    cell = table.cell(6, 0)

    # セル内のテキストを上書きする
    cell.text = draft_content

    # 上書きした文書を保存
    document.save(f"{new_filename}.docx")


if __name__ == "__main__":
    main()
