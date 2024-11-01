#!/usr/bin/env python
# coding=utf-8
'''
作    者 : 北极星光 light22@126.com
创建时间 : 2024-06-02 18:52:17
最后修改 : 2024-11-01 17:24:24
修 改 者 : 北极星光
'''
from datetime import datetime, timedelta

# Excel日期格式化
def excel_date_format(excel_value):
    """
    Excel日期格式化
    :param excel_value: Excel日期值
    :return: datetime对象
    """
    if excel_value:
        date_obj = datetime(1899, 12, 30) + timedelta(days=excel_value)
        return date_obj
    else:
        return


# 小写金额转大写
def num_to_capital(num):
    '''数字转人民币格式'''
    #判断参数类型
    if type(num) == type(''):
        try:
            f = float(num)
        except:
            return u'金额字符串无法转换成数字！'
        if f < 0: num = num[1:]
        tmp = num.split('.')
    elif type(num) in [type(0), type(0.0)]:
        fNum = 1.0e+16  # 如果是数字大于100兆会被截尾处理,字符串不受影响
        if not 0 <= num < fNum: return f'金额必须为非负数且小于{fNum}！'
        tmp = str(num).split('.')
    else:
        return u'金额必须是数字或字符串！'
    #汉字格式字符串
    ret, cn = '', [''] * 3
    cnStr = (u'角分整', u'零壹贰叁肆伍陆柒捌玖', u'元拾佰仟万拾佰仟亿拾佰仟兆拾佰仟京拾佰仟')
    for i in range(3):
        cn[i] = [c for c in cnStr[i]]
    if len(tmp[0]) > len(cn[2]): return u'金额数值过大，已超过设置的最大值！'
    if len(tmp) > 1 and len(tmp[1]) > 2: tmp[1] = tmp[1][:2]  #小数大于2位截尾处理
    #转换整数和小数
    for i, n in enumerate(reversed(list(tmp[0]))):
        ret = cn[1][int(n)] + cn[2][i] + ret
    if len(tmp) > 1:
        if tmp[0][-1] == '0': ret += cn[1][0]
        for i, n in enumerate(list(tmp[1])):
            ret += cn[1][int(n)] + cn[0][i]
    #根据书写规则替换掉不该出现的字串
    for i in range(2):
        ret = ret.replace(cn[1][0] + cn[0][i], cn[1][0])
    for i in range(1, 4):
        ret = ret.replace(cn[1][0] + cn[2][i], cn[1][0])
    for i in range(3, 1, -1):
        ret = ret.replace(cn[1][0] * i, cn[1][0])
    for i in range(0, len(cn[2]), 4):
        ret = ret.replace(cn[1][0] + cn[2][i], cn[2][i])
    for i in range(4, 10, 4):
        ret = ret.replace(cn[2][i + 4] + cn[2][i], cn[2][i + 4])
    if ret[0] in {cn[1][0], cn[2][0]}: ret = ret[1:]
    if len(ret) > 1:
        if ret[-1] == cn[1][0]: ret = ret[:-1] + cn[0][2]
        if ret[-1] == cn[2][0]: ret += cn[0][2]
        if ret[0] in {cn[1][0], cn[2][0]}: ret = ret[1:]
    else:
        ret = cn[1][0] + cn[2][0]
    return ret


# 替换文档中的变量改写
def replace_variables(document,**kwargs):
    '''替换模板中的变量'''
    # 查找文档中的标记并替换
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            for key, value in kwargs.items():
                if f'<{key}>' in run.text:
                    run.text = run.text.replace(f'<{key}>', str(value))
    return document


# 替换工作表中的变量改写
def replace_variables_sheet(sheet, **kwargs):
    '''替换模板中的变量'''
    # 查找工作表中的标记并替换
    for row in sheet.rows:
        for cell in row:
            for key, value in kwargs.items():
                if f'<{key}>' in str(cell.value):
                    cell.value = str(cell.value).replace(f'<{key}>', str(value))
    return sheet

if __name__ == '__main__':
    pass