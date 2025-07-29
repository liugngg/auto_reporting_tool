# auto_reporting_tool

## 代码说明

### v1版本
支持命令行 和 GUI 两种方式运行， 控制开关是 reportAardio.py 中 `AARDIO = False`
使用aardio打包出windows exe可执行程序

### v2025版本
- 使用tkinter重写UI + 打包逻辑
- 修改部分代码结构
- 增加了更新域代码、删除文档最后空白页等内容

打包命令:
```cmd

pyinstaller --icon=auto.ico --add-data "templates;./templates" --add-data "potin.png;."  --add-data "auto.ico;." --clean -w -D report_gui_liug.py --noconfirm -n 报告自动化生成工具软件

mkdir release
"C:\Program Files\7-Zip\7z.exe" a -tzip .\release\release.zip .\dist\报告自动化生成工具软件\*

```


## 配合使用的xlsm 模板
`\\192.168.0.200\PublicData\原始记录及报告模板\数通原始记录模板——2024.12.31`


