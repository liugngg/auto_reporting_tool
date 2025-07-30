# 报告自动化工具(V2025)介绍
## 概述
工具主要用于检测/检验报告的自动化生成，由报告的模板文件和原始记录表格文件两大部分组成，主要特点： <br>
- 检验报告模板的格式和内容完全分离；
- 报告模板内置于自动化工具中，控制着报告的格式；
- 原始记录采用Excel形式，控制内容；
- 报告模板和原始记录模板后期可以分开维护和管理。<br>
**原始记录测试完成后，报告一键自动化生成，基本不需要人工录入和修改。**<br>

## aardio版本（见“aardio_reporting”项目）
- 支持命令行 和 GUI 两种方式运行， 控制开关是 reportAardio.py 中 AARDIO = False; 
- 使用aardio打包出windows exe可执行程序;
- 图形化界面采用 aardio 制作，具体内容存放在 `Aardio_AutoReport`文件夹下；
- 原始记录的Excel模板在 `Template_examples` 文件夹下。 <br>

## v2025版本（本项目）
- 使用tkinter重写UI + 打包逻辑
- 修改部分代码结构
- 增加了更新域代码、删除文档最后空白页等内容

打包命令:
```cmd

pyinstaller --icon="templates/app.ico"  --add-data "templates;./templates" --clean -w -D main_gui.py --noconfirm -n 报告自动化生成工具

mkdir release
"C:\Program Files\7-Zip\7z.exe" a -tzip .\release\release.zip .\dist\报告自动化生成工具\*

```

## 配合使用的xlsm 模板
`\\192.168.0.200\PublicData\原始记录及报告模板\数通原始记录模板——2024.12.31`

## 联系人
liugang@caict.ac.cn


