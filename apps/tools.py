#!/usr/bin/env python
# coding=utf-8
'''
作    者 : 北极星光 light22@126.com
创建时间 : 2024-06-02 18:52:17
最后修改 : 2024-11-01 16:30:09
修 改 者 : 北极星光
'''


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
def replace_variables(document,
                      protocol_name=None,
                      contract_number=None,
                      contract_sign_date=None,
                      project_name=None,
                      project_address=None,
                      party_a_name=None,
                      party_b_name=None,
                      start_date=None,
                      completion_date=None,
                      warranty_period=None,
                      settlement_sign_date=None,
                      settlement_amount=None,
                      settlement_amount_capital=None,
                      tax_rate=None,
                      tax_rate_percent=None,
                      settlement_amount_untax=None,
                      settlement_amount_tax=None,
                      paid_amount=None,
                      paid_amount_capital=None,
                      balance_amount=None,
                      balance_amount_capital=None,
                      settlement_amount_untax_balance=None,
                      settlement_amount_tax_balance=None,
                      payment_ratio=None,
                      payment_ratio_percent=None,
                      payment_amount=None,
                      payment_amount_capital=None,
                      payment_amount_this=None,
                      payment_amount_this_capital=None,
                      warranty_ratio=None,
                      warranty_ratio_percent=None,
                      warranty_amount=None,
                      warranty_amount_capital=None):
    '''替换模板中的变量'''
    # 查找文档中的标记并替换
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if '<protocol_name>' in run.text:
                if protocol_name is not None:
                    run.text = run.text.replace('<protocol_name>',
                                                str(protocol_name))
            if '<contract_number>' in run.text:
                if contract_number is not None:
                    run.text = run.text.replace('<contract_number>',
                                                str(contract_number))
            if '<contract_sign_date>' in run.text:
                if contract_sign_date is not None:
                    run.text = run.text.replace('<contract_sign_date>',
                                                str(contract_sign_date))
            if '<project_name>' in run.text:
                if project_name is not None:
                    run.text = run.text.replace('<project_name>',
                                                str(project_name))
            if '<project_address>' in run.text:
                if project_address is not None:
                    run.text = run.text.replace('<project_address>',
                                                str(project_address))
            if '<party_a_name>' in run.text:
                if party_a_name is not None:
                    run.text = run.text.replace('<party_a_name>',
                                                str(party_a_name))
            if '<party_b_name>' in run.text:
                if party_b_name is not None:
                    run.text = run.text.replace('<party_b_name>',
                                                str(party_b_name))
            if '<start_date>' in run.text:
                if start_date is not None:
                    run.text = run.text.replace('<start_date>',
                                                str(start_date))
            if '<completion_date>' in run.text:
                if completion_date is not None:
                    run.text = run.text.replace('<completion_date>',
                                                str(completion_date))
            if '<warranty_period>' in run.text:
                if warranty_period is not None:
                    run.text = run.text.replace('<warranty_period>',
                                                str(warranty_period))
            if '<settlement_sign_date>' in run.text:
                if settlement_sign_date is not None:
                    run.text = run.text.replace('<settlement_sign_date>',
                                                str(settlement_sign_date))
            if '<settlement_amount>' in run.text:
                if settlement_amount is not None:
                    run.text = run.text.replace('<settlement_amount>',
                                                str(settlement_amount))
            if '<settlement_amount_capital>' in run.text:
                if settlement_amount_capital is not None:
                    run.text = run.text.replace('<settlement_amount_capital>',
                                                str(settlement_amount_capital))
            if '<tax_rate>' in run.text:
                if tax_rate is not None:
                    run.text = run.text.replace('<tax_rate>', str(tax_rate))
            if '<tax_rate_percent>' in run.text:
                if tax_rate_percent is not None:
                    run.text = run.text.replace('<tax_rate_percent>',
                                                str(tax_rate_percent))
            if '<settlement_amount_untax>' in run.text:
                if settlement_amount_untax is not None:
                    run.text = run.text.replace('<settlement_amount_untax>',
                                                str(settlement_amount_untax))
            if '<settlement_amount_tax>' in run.text:
                if settlement_amount_tax is not None:
                    run.text = run.text.replace('<settlement_amount_tax>',
                                                str(settlement_amount_tax))
            if '<paid_amount>' in run.text:
                if paid_amount is not None:
                    run.text = run.text.replace('<paid_amount>',
                                                str(paid_amount))
            if '<paid_amount_capital>' in run.text:
                if paid_amount_capital is not None:
                    run.text = run.text.replace('<paid_amount_capital>',
                                                str(paid_amount_capital))
            if '<balance_amount>' in run.text:
                if balance_amount is not None:
                    run.text = run.text.replace('<balance_amount>',
                                                str(balance_amount))
            if '<balance_amount_capital>' in run.text:
                if balance_amount_capital is not None:
                    run.text = run.text.replace('<balance_amount_capital>',
                                                str(balance_amount_capital))
            if '<settlement_amount_untax_balance>' in run.text:
                if settlement_amount_untax_balance is not None:
                    run.text = run.text.replace(
                        '<settlement_amount_untax_balance>',
                        str(settlement_amount_untax_balance))
            if '<settlement_amount_tax_balance>' in run.text:
                if settlement_amount_tax_balance is not None:
                    run.text = run.text.replace(
                        '<settlement_amount_tax_balance>',
                        str(settlement_amount_tax_balance))
            if '<payment_ratio>' in run.text:
                if payment_ratio is not None:
                    run.text = run.text.replace('<payment_ratio>',
                                                str(payment_ratio))
            if '<payment_ratio_percent>' in run.text:
                if payment_ratio_percent is not None:
                    run.text = run.text.replace('<payment_ratio_percent>',
                                                str(payment_ratio_percent))
            if '<payment_amount>' in run.text:
                if payment_amount is not None:
                    run.text = run.text.replace('<payment_amount>',
                                                str(payment_amount))
            if '<payment_amount_capital>' in run.text:
                if payment_amount_capital is not None:
                    run.text = run.text.replace('<payment_amount_capital>',
                                                str(payment_amount_capital))
            if '<payment_amount_this>' in run.text:
                if payment_amount_this is not None:
                    run.text = run.text.replace('<payment_amount_this>',
                                                str(payment_amount_this))
            if '<payment_amount_this_capital>' in run.text:
                if payment_amount_this_capital is not None:
                    run.text = run.text.replace(
                        '<payment_amount_this_capital>',
                        str(payment_amount_this_capital))
            if '<warranty_ratio>' in run.text:
                if warranty_ratio is not None:
                    run.text = run.text.replace('<warranty_ratio>',
                                                str(warranty_ratio))
            if '<warranty_ratio_percent>' in run.text:
                if warranty_ratio_percent is not None:
                    run.text = run.text.replace('<warranty_ratio_percent>',
                                                str(warranty_ratio_percent))
            if '<warranty_amount>' in run.text:
                if warranty_amount is not None:
                    run.text = run.text.replace('<warranty_amount>',
                                                str(warranty_amount))
            if '<warranty_amount_capital>' in run.text:
                if warranty_amount_capital is not None:
                    run.text = run.text.replace('<warranty_amount_capital>',
                                                str(warranty_amount_capital))
    return document


# 替换工作表中的变量改写
def replace_variables_sheet(sheet,
                            protocol_name=None,
                            contract_number=None,
                            contract_sign_date=None,
                            project_name=None,
                            project_address=None,
                            party_a_name=None,
                            party_b_name=None,
                            start_date=None,
                            completion_date=None,
                            warranty_period=None,
                            settlement_sign_date=None,
                            contract_amount=None,
                            contract_amount_capital=None,
                            settlement_amount=None,
                            settlement_amount_capital=None,
                            tax_rate=None,
                            tax_rate_percent=None,
                            settlement_amount_untax=None,
                            settlement_amount_tax=None,
                            paid_amount=None,
                            paid_amount_capital=None,
                            balance_amount=None,
                            balance_amount_capital=None,
                            settlement_amount_untax_balance=None,
                            settlement_amount_tax_balance=None,
                            payment_ratio=None,
                            payment_ratio_percent=None,
                            payment_amount=None,
                            payment_amount_capital=None,
                            payment_amount_this=None,
                            payment_amount_this_capital=None,
                            warranty_ratio=None,
                            warranty_ratio_percent=None,
                            warranty_amount=None,
                            warranty_amount_capital=None):
    '''替换模板中的变量'''
    # 查找工作表中的标记并替换
    for row in sheet.rows:
        for cell in row:
            if cell.value is not None:
                if '<protocol_name>' in str(cell.value):
                    if protocol_name is not None:
                        cell.value = str(cell.value).replace(
                            '<protocol_name>', str(protocol_name))
                if '<contract_number>' in str(cell.value):
                    if contract_number is not None:
                        cell.value = str(cell.value).replace(
                            '<contract_number>', str(contract_number))
                if '<contract_sign_date>' in str(cell.value):
                    if contract_sign_date is not None:
                        cell.value = str(cell.value).replace(
                            '<contract_sign_date>', str(contract_sign_date))
                if '<project_name>' in str(cell.value):
                    if project_name is not None:
                        cell.value = str(cell.value).replace(
                            '<project_name>', str(project_name))
                if '<project_address>' in str(cell.value):
                    if project_address is not None:
                        cell.value = str(cell.value).replace(
                            '<project_address>', str(project_address))
                if '<party_a_name>' in str(cell.value):
                    if party_a_name is not None:
                        cell.value = str(cell.value).replace(
                            '<party_a_name>', str(party_a_name))
                if '<party_b_name>' in str(cell.value):
                    if party_b_name is not None:
                        cell.value = str(cell.value).replace(
                            '<party_b_name>', str(party_b_name))
                if '<start_date>' in str(cell.value):
                    if start_date is not None:
                        cell.value = str(cell.value).replace(
                            '<start_date>', str(start_date))
                if '<completion_date>' in str(cell.value):
                    if completion_date is not None:
                        cell.value = str(cell.value).replace(
                            '<completion_date>', str(completion_date))
                if '<warranty_period>' in str(cell.value):
                    if warranty_period is not None:
                        cell.value = str(cell.value).replace(
                            '<warranty_period>', str(warranty_period))
                if '<settlement_sign_date>' in str(cell.value):
                    if settlement_sign_date is not None:
                        cell.value = str(cell.value).replace(
                            '<settlement_sign_date>',
                            str(settlement_sign_date))
                if '<contract_amount>' in str(cell.value):
                    if contract_amount is not None:
                        cell.value = str(cell.value).replace(
                            '<contract_amount>', str(contract_amount))
                if '<contract_amount_capital>' in str(cell.value):
                    if contract_amount_capital is not None:
                        cell.value = str(cell.value).replace(
                            '<contract_amount_capital>',
                            str(contract_amount_capital))
                if '<settlement_amount>' in str(cell.value):
                    if settlement_amount is not None:
                        cell.value = str(cell.value).replace(
                            '<settlement_amount>', str(settlement_amount))
                if '<settlement_amount_capital>' in str(cell.value):
                    if settlement_amount_capital is not None:
                        cell.value = str(cell.value).replace(
                            '<settlement_amount_capital>',
                            str(settlement_amount_capital))
                if '<tax_rate>' in str(cell.value):
                    if tax_rate is not None:
                        cell.value = str(cell.value).replace(
                            '<tax_rate>', str(tax_rate))
                if '<tax_rate_percent>' in str(cell.value):
                    if tax_rate_percent is not None:
                        cell.value = str(cell.value).replace(
                            '<tax_rate_percent>', str(tax_rate_percent))
                if '<settlement_amount_untax>' in str(cell.value):
                    if settlement_amount_untax is not None:
                        cell.value = str(cell.value).replace(
                            '<settlement_amount_untax>',
                            str(settlement_amount_untax))
                if '<settlement_amount_tax>' in str(cell.value):
                    if settlement_amount_tax is not None:
                        cell.value = str(cell.value).replace(
                            '<settlement_amount_tax>',
                            str(settlement_amount_tax))
                if '<paid_amount>' in str(cell.value):
                    if paid_amount is not None:
                        cell.value = str(cell.value).replace(
                            '<paid_amount>', str(paid_amount))
                if '<paid_amount_capital>' in str(cell.value):
                    if paid_amount_capital is not None:
                        cell.value = str(cell.value).replace(
                            '<paid_amount_capital>', str(paid_amount_capital))
                if '<balance_amount>' in str(cell.value):
                    if balance_amount is not None:
                        cell.value = str(cell.value).replace(
                            '<balance_amount>', str(balance_amount))
                if '<balance_amount_capital>' in str(cell.value):
                    if balance_amount_capital is not None:
                        cell.value = str(cell.value).replace(
                            '<balance_amount_capital>',
                            str(balance_amount_capital))
                if '<settlement_amount_untax_balance>' in str(cell.value):
                    if settlement_amount_untax_balance is not None:
                        cell.value = str(cell.value).replace(
                            '<settlement_amount_untax_balance>',
                            str(settlement_amount_untax_balance))
                if '<settlement_amount_tax_balance>' in str(cell.value):
                    if settlement_amount_tax_balance is not None:
                        cell.value = str(cell.value).replace(
                            '<settlement_amount_tax_balance>',
                            str(settlement_amount_tax_balance))
                if '<payment_ratio>' in str(cell.value):
                    if payment_ratio is not None:
                        cell.value = str(cell.value).replace(
                            '<payment_ratio>', str(payment_ratio))
                if '<payment_ratio_percent>' in str(cell.value):
                    if payment_ratio_percent is not None:
                        cell.value = str(cell.value).replace(
                            '<payment_ratio_percent>',
                            str(payment_ratio_percent))
                if '<payment_amount>' in str(cell.value):
                    if payment_amount is not None:
                        cell.value = str(cell.value).replace(
                            '<payment_amount>', str(payment_amount))
                if '<payment_amount_capital>' in str(cell.value):
                    if payment_amount_capital is not None:
                        cell.value = str(cell.value).replace(
                            '<payment_amount_capital>',
                            str(payment_amount_capital))
                if '<payment_amount_this>' in str(cell.value):
                    if payment_amount_this is not None:
                        cell.value = str(cell.value).replace(
                            '<payment_amount_this>', str(payment_amount_this))
                if '<payment_amount_this_capital>' in str(cell.value):
                    if payment_amount_this_capital is not None:
                        cell.value = str(cell.value).replace(
                            '<payment_amount_this_capital>',
                            str(payment_amount_this_capital))
                if '<warranty_ratio>' in str(cell.value):
                    if warranty_ratio is not None:
                        cell.value = str(cell.value).replace(
                            '<warranty_ratio>', str(warranty_ratio))
                if '<warranty_ratio_percent>' in str(cell.value):
                    if warranty_ratio_percent is not None:
                        cell.value = str(cell.value).replace(
                            '<warranty_ratio_percent>',
                            str(warranty_ratio_percent))
                if '<warranty_amount>' in str(cell.value):
                    if warranty_amount is not None:
                        cell.value = str(cell.value).replace(
                            '<warranty_amount>', str(warranty_amount))
                if '<warranty_amount_capital>' in str(cell.value):
                    if warranty_amount_capital is not None:
                        cell.value = str(cell.value).replace(
                            '<warranty_amount_capital>',
                            str(warranty_amount_capital))
    return sheet
