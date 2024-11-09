# 工程协议资料自动书写

## 概述

用于批量自动生成工程结算申报资料和结算协议、质保协议等文件。

作者：北极星光	E-mail：light22@126.com

github项目地址：[https://github.com/light0220/auto_write_protocol](https://github.com/light0220/auto_write_protocol)

gitee项目地址：[https://gitee.com/light22/auto_write_protocol](https://gitee.com/light22/auto_write_protocol)

## 使用方法

+ 安装python3环境: [https://www.python.org/downloads/](https://www.python.org/downloads/)
+ 安装本项目依赖的第三方库：`pip install -r requirements.txt`
+ 克隆项目到本地：

  + github：`git clone https://github.com/light0220/auto_write_protocol.git`
  + gitee：`git clone https://gitee.com/light22/auto_write_protocol.git`
+ 进入本项目目录，将 `config-示例.xlsx` 文件改名为 `config.xlsx`
+ 进入 `\data\`子目录，将目录下所有文件名的 `-示例`后缀去掉
+ 生成结算协议以及质量保修书

  + 打开 `config.xlsx` 按照示例格式，将工程结算信息数据填入 `结算协议`工作表中。
  + 运行 `main.py` 根据提示输入1，完成结算协议以及质量保修书的自动生成。
+ 生成结算申报资料

  + 打开 `config.xlsx` 按照示例格式，将工程结算信息数据填入 `结算申报`工作表中。
  + 运行 `main.py` 根据提示输入2，完成结算申报资料的自动生成。

## 开发说明

由于不同公司的资料申报格式不尽相同，本项目提供的示例模板不一定适用所有情景，因此欢迎各位有编程基础的朋友在本框架的基础上二次开发，以生成更适合的资料格式，甚至实现生成施工合同、劳动合同、投标文件等一切格式化的资料内容。开发文档如下：

+ 项目结构：

  .
  ├── /apps/							# 主要存放程序代码的目录
  ├── /apps/tool.py						# 项目中可能调用的通用工具代码
  ├── /apps/auto_settle_protocol.py		# 生成结算协议及质量保修书的代码
  ├── /apps/generate_settlement_report.py	# 生成工程结算申报资料的代码
  ├── /data/             						# 存放资料模板的目录
  ├── /output/							# 存放输出内容的目录
  ├── /requirements.txt					# 项目依赖的第三方库
  ├── /config.xlsx						# 工程信息配置文件
  ├── /main.py							# 程序主入口文件
  └── README.md						# 项目说明文档
+ 开发思路及格式要求：

  + 主要思路为通过读取 `config.xlsx`中的数据，替换模板文档中的对应数据从而达成快速生成文档、表格等资料。
  + 模板文档中需要替换的数据应提前修改成对应标记，为方便阅读以及更好地兼容本框架，统一要求将标记写成 `<变量名>`的格式，具体修改方法可参考示例模板。
  + 新增功能时可在 `/apps/`目录下新建一个.py文件，例如 `new_function.py`，并在 `main.py` 中调用，`new_function.py`中代码参考如下：

    ```
    import os
    from docx import Document
    from openpyxl import load_workbook
    from tools import *

    def new_function():  # 新增功能的函数名
        # 读取config.xlsx文件
        workbook = load_workbook('config.xlsx', data_only=True)
        sheet = workbook['XXXX']  # 此处应与config.xlsx中的工作表名称一致，可新建一个工作表，名称尽量体现用途
        # 遍历sheet
        for i in range(2, sheet.max_row + 1):
    	...  # 此步骤将config.xlsx中对应工作表中的数据遍历，写入对应变量

            # 生成Word文档类资料
            document = Document('data/XXX.docx')  # 此处为已替换好标记的格式模板文档

            # 替换模板中的变量
            replace_variables(
                document,
                keys=value,  # 此处的keys为模板文档中替换的工程信息标记<>中的变量名，value为上一步读取到的对应数据
    	    ...
    	)
            # 创建工程文件夹
            path = f'output/{project_name}XXXX'  # 此处建议根据工程名称创建文件夹
            os.makedirs(path, exist_ok=True)
            # 保存文件
            document.save(f'{path}/{project_name}XXXX.docx')

            # 生成Excel表格类资料
            wb = load_workbook('data/XXX.xlsx')  # 此处为已替换好标记的格式模板文档

            # 替换模板中的变量
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                replace_variables_sheet(
                    ws,
                    keys=value,  # 此处的keys为模板文档中替换的工程信息标记<>中的变量名，value为上一步读取到的对应数据
    	    ...
                )
            # 保存文件
            wb.save(f'output/{project_name}XXXX.xlsx')

    ```

    新增功能时，应在 `main.py` 中添加对应功能的调用，例如：

    ```
    if choice == "3":  # 此处为新增功能的选项
        new_function()  # 此处为新增功能的调用
    ```

---
