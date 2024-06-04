#!/usr/bin/env python
# coding=utf-8
'''
作    者 : 北极星光 light22@126.com
创建时间 : 2024-06-04 20:56:34
最后修改 : 2024-06-05 00:33:04
修 改 者 : 北极星光
'''

import time
from apps.auto_settle_protocol import auto_settle_protocol
from apps.generate_settlement_report import generate_settlement_report

if __name__ == '__main__':
    choice = input("请选择功能：\n 1.生成结算协议和质保协议\n 2.生成结算申报资料\n 输入其它字符退出\n请输入：")
    if choice == "1":
        t1 = time.time()
        print("正在生成结算协议和质保协议...")
        auto_settle_protocol()
        t2 = time.time()
        print(f"生成结算协议和质保协议完成，用时：{t2 - t1:.2f} 秒")
    elif choice == "2":
        t1 = time.time()
        print("正在生成结算申报资料...")
        generate_settlement_report()
        t2 = time.time()
        print(f"生成结算申报资料完成，用时：{t2 - t1:.2f} 秒")
    else:
        print("退出程序")
