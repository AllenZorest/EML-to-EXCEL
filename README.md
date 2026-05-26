# EML to Excel — 邮件批量解析工具

![Python](https://img.shields.io/badge/python-3.8+-blue) ![Streamlit](https://img.shields.io/badge/Streamlit-Web_App-FF4B4B?logo=streamlit) ![License](https://img.shields.io/badge/license-MIT-green)

> 为日本字体设计公司 **Nexas** 处理 1800+ 封中/日/英商务邮件，将非结构化 EML 批量转换为结构化 Excel。

## 功能

- 批量解析 `.eml` 文件，提取 7 个字段：**文件名、发件人、时间、收件人、抄送、标题、正文**
- 支持 MIME 编码解码：Base64 / Quoted-Printable / RFC 2047 邮件标头
- 支持 `multipart/alternative`、`multipart/mixed` 等多种 MIME 结构
- chardet 自动检测编码，解决邮件声明 charset 与实际编码不一致的问题
- 多时区归一化至北京时间（UTC+8），输出格式 `YYYY/MM/DD HH:MM:SS`
- 输出带格式 Excel：冻结表头、自动筛选、列宽适配

## 两种使用方式

### 方式一：Streamlit Web 界面（推荐）

```bash
pip install -r requirements.txt
streamlit run app.py
```

浏览器打开后，拖拽上传 EML 文件 → 实时预览解析结果 → 一键下载 Excel。

### 方式二：命令行工具

```bash
pip install openpyxl chardet
python eml_to_excel.py <EML目录路径> <输出Excel路径>
```

**示例：**

```bash
python eml_to_excel.py ./emails ./emails_output.xlsx
```

## 项目结构

```
EML-to-EXCEL/
├── app.py              # Streamlit Web 界面
├── eml_to_excel.py     # 命令行批量处理工具
├── requirements.txt    # 依赖列表
└── README.md
```

## 背景

大二期间通过老师对接，协助日本字体设计公司 **Nexas**（东京，在中国区设有团队）处理邮件数据治理任务。
Nexas 是小规模字体厂商，日常商务沟通涉及中/日/英三语邮件，
散落在多个目录中的人工归档方式效率极低。此脚本实现一键批量转换，
将 1800+ 封邮件输出为可检索、可筛选的结构化归档文件。

## License

MIT — 随意使用，注明出处即可。
