import os
import re
import sys

import pdfplumber
import pyexcel


def extract_report_number(pdf_path):
    """从PDF第一页提取报告编号"""
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()
        if text:
            match = re.search(r"报告编号\s*([\w-]+)", text)
            if match:
                return match.group(1)
    return ""


def extract_kuanhao_and_result(pdf_path):
    """从PDF表格中提取款号和测试结果"""
    kuanhao = ""
    test_result = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for column in table:
                        if "客户认定信息" in column:
                            values = column[1]
                            if values:
                                kuanhao_matches = re.findall(r'\b([A-Z]\d{9}(-A)?)\b', values)
                                if kuanhao_matches:
                                    kuanhao = ", ".join([match[0] for match in kuanhao_matches])
                    header = table[0]
                    try:
                        test_result_index = header.index("测试结果")
                        for row in table[1:3]:
                            if len(row) > test_result_index and row[test_result_index]:
                                test_result = row[test_result_index]
                                break
                        if test_result:
                            break
                    except ValueError:
                        continue
                if kuanhao and test_result:
                    break
    return kuanhao, test_result


def write_to_excel(data, output_file):
    """将数据写入Excel文件"""
    sheet = pyexcel.Sheet(data)
    sheet.save_as(output_file)


def method_name(path):
    report_no = extract_report_number(path)
    kuanhao, test_result = extract_kuanhao_and_result(path)
    return {"报告编号": report_no, "款号": kuanhao, "测试结果": test_result}


def main(args):
    results = []
    path = args[0]
    for fileName in os.listdir(path):
        if not fileName.endswith(".pdf"):
            continue
        pdf_path = os.path.join(path, fileName)
        try:
            element = method_name(pdf_path)
            results.append(element)
            print(f"{fileName}文件已处理完毕")
        except Exception as e:
            print(f"{fileName}文件处理失败已跳过: {e}")
    # 将结果转换为适合 pyexcel 的格式
    data = [["报告编号", "款号", "测试结果"]]  # 添加表头
    for result in results:
        report_no = result.get("报告编号", "")
        test_result = result.get("测试结果", "")
        kuanhao_str = result.get("款号", "")
        if kuanhao_str:
            kuanhao_list = [k.strip() for k in kuanhao_str.split(",")]
            for kuanhao in kuanhao_list:
                data.append([report_no, kuanhao, test_result])
        else:
            data.append([report_no, "", test_result])

    # 使用 pyexcel 保存为 Excel 文件
    output_file = os.path.join(path, "output.xlsx")
    write_to_excel(data, output_file)


if __name__ == '__main__':
    main(sys.argv[1:])