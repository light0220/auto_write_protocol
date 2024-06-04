# 工程协议资料自动书写

## 概述

用于批量自动生成工程结算申报资料和结算协议、质保协议等文件。

作者：北极星光	E-mail：light22@126.com

github项目地址：[light0220/auto_write_protocol (github.com)](https://github.com/light0220/auto_write_protocol)

gitee项目地址：[auto_write_protocol: 自动书写工程协议 (gitee.com)](https://gitee.com/light22/auto_write_protocol)


## 使用方法

+ 克隆项目到本地：
  + github：`git clone https://github.com/light0220/auto_write_protocol.git`
  + gitee：`git clone https://gitee.com/light22/auto_write_protocol.git`
+ 将 config - 示例.xlsx 文件改名为 config.xlsx
+ 打开 config.xlsx 按照示例格式，将工程信息填入表格。
+ 将协议模板中需要替换的内容替换成标签，参考 data\示例.docx 中所示，并将文件重命名为 结算协议-模板.docx 和 质量保修书-模板.docx，放入data目录下
+ 运行根目录下 main.py 根据提示选择功能
