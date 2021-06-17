from docx import Document


old_file = "C:/01_input/test.docx"
new_file = "C:/02_output/result.docx"


def check_and_change(doc):
    """
    注意事项：
    直接操作paragraph.text会导致字体格式丢失，所以要遍历其run操作。
    一个段落会被分成多个run，为了确保替换元内容不被分割，最好用"AA、AB"这种形式，而不要使用类似"${21}"这样的字符串。
    """

    # 遍历所有段落进行替换。
    for para in doc.paragraphs:
        for i in range(len(para.runs)):
            para.runs[i].text = para.runs[i].text.replace("AA", "■")

    # 遍历所有表格进行替换。
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        # print(run.text)
                        run.text = run.text.replace("AA", "宋宗正")
                        run.text = run.text.replace("AB", "132934198101053217")
                        run.text = run.text.replace("BA", "■")
                        run.text = run.text.replace("BB", "□")
                        run.text = run.text.replace("BC", "□")
                        run.text = run.text.replace("BD", "128")

    return doc


def main():
    doc = Document(old_file)
    doc = check_and_change(doc)
    doc.save(new_file)


if __name__ == '__main__':
    main()
