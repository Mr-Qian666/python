import requests
import excel_function

file_path = r'C:\CI\foundation.xlsx'


def get_foundation_info(code_number):
    """
    根据基金代码获取当前基金信息
    :param code_number: 基金代码
    :return: 基金信息，包括获取时间，净值，涨跌值
    """
    url = "http://fundgz.1234567.com.cn/js/" + str(code_number) + ".js"
    content = requests.get(url).text
    print("#####return info#####:" + content)
    text = str(content)
    gztime_info = text[content.find(r'gztime') + 9: content.find(r'}') - 1]
    gsz_info = text[content.find(r'gsz') + 6: content.find(r'gszzl') - 3]
    gszzl_info = text[content.find(r'gszzl') + 8: content.find(r'gztime') - 3]
    foundation_value = [gztime_info, gsz_info, gszzl_info]
    return foundation_value


def update_foundation_info(para_list):
    """
    将获取到的基金估值填入表格指定位置
    :param para_list: get_foundation_info返回的value
    :return: null
    """
    for i in para_list:
        code_number = excel_function.get_value_from_excel(file_path, 'Sheet1', 'A' + str(i))
        assessment_list = get_foundation_info(code_number)
        excel_function.write_save_value_to_excel(file_path, 'Sheet1', 'N' + str(i), assessment_list[0])
        excel_function.write_save_value_to_excel(file_path, 'Sheet1', 'O' + str(i), assessment_list[1])
        excel_function.write_save_value_to_excel(file_path, 'Sheet1', 'P' + str(i), assessment_list[2])


def compare_transaction_point(para_list):
    """
    根据最新净值，与目标限定值比较后标记颜色
    :param para_list: 基金行数
    :return: null
    """
    for i in para_list:
        checkpoint_5 = excel_function.get_value_from_excel(file_path, "Sheet1", "F" + str(i))
        checkpoint_10 = excel_function.get_value_from_excel(file_path, "Sheet1", "G" + str(i))
        checkpoint_15 = excel_function.get_value_from_excel(file_path, "Sheet1", "H" + str(i))
        checkpoint_20 = excel_function.get_value_from_excel(file_path, "Sheet1", "I" + str(i))
        checkpoint_25 = excel_function.get_value_from_excel(file_path, "Sheet1", "J" + str(i))
        checkpoint_30 = excel_function.get_value_from_excel(file_path, "Sheet1", "K" + str(i))
        checkpoint_35 = excel_function.get_value_from_excel(file_path, "Sheet1", "L" + str(i))
        checkpoint_40 = excel_function.get_value_from_excel(file_path, "Sheet1", "M" + str(i))
        assessment = excel_function.get_value_from_excel(file_path, "Sheet1", "O" + str(i))
        if float(checkpoint_5) > float(assessment):
            excel_function.cell_fill_color(file_path, "Sheet1", "F" + str(i), 'red')
        else:
            excel_function.cell_fill_color(file_path, "Sheet1", "F" + str(i), 'white')
        if float(checkpoint_10) > float(assessment):
            excel_function.cell_fill_color(file_path, "Sheet1", "G" + str(i), 'red')
        else:
            excel_function.cell_fill_color(file_path, "Sheet1", "G" + str(i), 'white')
        if float(checkpoint_15) > float(assessment):
            excel_function.cell_fill_color(file_path, "Sheet1", "H" + str(i), 'red')
        else:
            excel_function.cell_fill_color(file_path, "Sheet1", "H" + str(i), 'white')
        if float(checkpoint_20) > float(assessment):
            excel_function.cell_fill_color(file_path, "Sheet1", "I" + str(i), 'red')
        else:
            excel_function.cell_fill_color(file_path, "Sheet1", "I" + str(i), 'white')
        if float(checkpoint_25) > float(assessment):
            excel_function.cell_fill_color(file_path, "Sheet1", "J" + str(i), 'red')
        else:
            excel_function.cell_fill_color(file_path, "Sheet1", "J" + str(i), 'white')
        if float(checkpoint_30) > float(assessment):
            excel_function.cell_fill_color(file_path, "Sheet1", "K" + str(i), 'red')
        else:
            excel_function.cell_fill_color(file_path, "Sheet1", "K" + str(i), 'white')
        if float(checkpoint_35) > float(assessment):
            excel_function.cell_fill_color(file_path, "Sheet1", "L" + str(i), 'red')
        else:
            excel_function.cell_fill_color(file_path, "Sheet1", "L" + str(i), 'white')
        if float(checkpoint_40) > float(assessment):
            excel_function.cell_fill_color(file_path, "Sheet1", "M" + str(i), 'red')
        else:
            excel_function.cell_fill_color(file_path, "Sheet1", "M" + str(i), 'white')
    return 0
