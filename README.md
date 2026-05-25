# EML to Excel — 邮件批量解析工具

![Python](https://img.shields.io/badge/python-3.8+-blue) ![License](https://img.shields.io/badge/license-MIT-green)

> 为日本 **Nexas Inc.** 处理 1800+ 封商务邮件，将非结构化 EML 批量转换为结构化 Excel。

## 功能

- 批量解析 `.eml` 文件，提取 7 个字段：**文件名、发件人、时间、收件人、抄送、标题、正文**
- 支持 MIME 编码解码：Base64 / Quoted-Printable / RFC 2047 邮件标头
- 支持 `multipart/alternative`、`multipart/mixed` 等多种 MIME 结构
- chardet 自动检测编码，解决邮件声明 charset 与实际编码不一致的问题
- 多时区归一化至北京时间（UTC+8），输出格式 `YYYY/MM/DD HH:MM:SS`
- 输出带格式 Excel：冻结表头、自动筛选、列宽适配

## 快速开始

```bash
pip install openpyxl chardet
python eml_to_excel.py <EML目录路径> <输出Excel路径>
```

## 示例

```bash
python eml_to_excel.py ./emails ./emails_output.xlsx
```

## 背景

大二期间协助日本 Nexas 公司（东京）处理邮件数据治理任务。
原始商务邮件散落在多个文件夹中，涉及中/日/英三语，人工逐封查阅效率极低。
此脚本实现一键批量转换，输出可检索、可筛选的结构化归档文件。

## License

MIT — 随意使用，注明出处即可。
