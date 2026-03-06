# 小雷Excel批量助手 (Xiao Lei Excel Batch Helper)

⚡ 一个基于 Streamlit 的高效 Excel 批量处理工具，旨在简化繁琐的 Excel 数据更新和提取任务。

## 📖 简介

小雷Excel批量助手是一个 Web 应用程序，提供直观的图形界面，帮助用户轻松完成 Excel 文件的批量更新和字段提取工作。无论你是需要根据一个总表更新多个分表，还是从复杂的工作表中提取特定列，这个工具都能帮你快速搞定。

## ✨ 主要功能

### 1. 批量更新 (Batch Update)
根据源文件（A表）的数据，批量更新多个目标文件（B表）中的指定列。
- **灵活匹配**：支持自定义 A 表和 B 表的匹配列（Key）和更新列（Value）。
- **多文件处理**：一次上传多个 B 表，批量执行更新操作。
- **个性化配置**：支持全局配置，也可以针对每个 B 表单独设置表头行数、起始行等参数。
- **格式保持**：
    - 如果目标文件是 `.xlsx`，更新后保持为 `.xlsx` 格式。
    - 如果目标文件是 `.xls`，更新后保持为 `.xls` 格式。

### 2. 字段提取 (Field Extraction)
从单个 Excel 文件中快速提取所需字段，生成新的 Excel 文件。
- **按需提取**：上传文件后，通过多选框选择需要保留的列。
- **样式保留**：导出时尽可能保留原单元格的字体、边框、背景色等样式。
- **智能列宽**：根据内容（包含对中文字符的支持）自动调整导出文件的列宽，方便查看。
- **统一格式**：无论输入是 `.xls` 还是 `.xlsx`，提取结果统一保存为 `.xlsx` 格式。

## 🛠️ 安装指南

### 前置要求
- Python 3.8 或更高版本

### 安装步骤

1. 克隆或下载本项目到本地。
2. 进入项目根目录。
3. 建议创建并激活虚拟环境：
   ```bash
   python -m venv venv
   # Windows
   venv\Scripts\activate
   # macOS/Linux
   source venv/bin/activate
   ```
4. 安装依赖包：
   ```bash
   pip install -r requirements.txt
   ```

## 🚀 使用方法

### 启动应用

在终端中运行以下命令启动应用：

```bash
streamlit run app.py
```

或者使用封装好的启动脚本（方便双击运行或作为入口）：

```bash
python run_app.py
```

启动后，浏览器将自动打开应用的访问地址（通常是 `http://localhost:8501`）。

### 操作流程

#### 批量更新
1. **上传文件**：
   - 上传 **A表**（数据源）。
   - 上传一个或多个 **B表**（待更新文件）。
2. **设置行号**：
   - 设定表头所在行（用于读取列名）。
   - 设定数据起始行（从哪一行开始处理数据）。
3. **配置映射**：
   - 在全局配置区，选择 A 表和 B 表的 **匹配列**（用于关联数据）和 **更新列**（需要写入数据的列）。
   - 如有特殊文件，可在“个性化清单”中单独调整配置。
4. **执行**：点击“开始批量处理”按钮。
5. **下载**：处理完成后，下载更新后的文件。

#### 字段提取
1. **上传文件**：上传需要提取数据的 Excel 文件。
2. **选择列**：勾选需要保留的字段。
3. **配置**：设置表头行和数据起始行。
4. **执行**：点击“执行提取”。
5. **下载**：下载生成的 `.xlsx` 文件。

## 📦 打包构建 (可选)

本项目包含 `run_app.spec` 配置，支持使用 PyInstaller 打包为独立的可执行文件（.exe 或 .app），无需 Python 环境即可运行。

```bash
pyinstaller run_app.spec
```

打包完成后，可在 `dist/` 目录下找到生成的可执行程序。

## 📂 项目结构

```
ExcelAssistant/
├── app.py                  # Streamlit 应用主入口
├── run_app.py              # 启动脚本封装
├── run_app.spec            # PyInstaller 打包配置文件
├── requirements.txt        # 项目依赖列表
├── README.md               # 项目说明文档
├── hooks/                  # PyInstaller 钩子配置
│   └── hook-streamlit.py   # Streamlit 打包钩子
├── pages/                  # 功能页面模块
│   ├── batch_update.py     # 批量更新功能逻辑
│   └── field_extraction.py # 字段提取功能逻辑
└── utils/                  # 工具函数
    └── excel_helpers.py    # Excel 处理辅助函数
```

## 🧰 技术栈

- **[Streamlit](https://streamlit.io/)**: 快速构建数据应用 Web 界面。
- **[Pandas](https://pandas.pydata.org/)**: 强大的数据分析和处理库。
- **[OpenPyXL](https://openpyxl.readthedocs.io/)**: 读写 Excel 2010 xlsx/xlsm/xltx/xltm 文件。
- **[xlrd](https://xlrd.readthedocs.io/)**: 读取旧版 Excel (xls) 文件。
- **[xlutils](https://xlutils.readthedocs.io/)**: 处理旧版 Excel (xls) 文件的复制和写入。
- **[PyInstaller](https://pyinstaller.org/)**: 将 Python 程序打包成独立可执行文件。

## 🤝 贡献

欢迎提交 Issue 或 Pull Request 来改进这个项目！

---
*由小雷Excel批量助手生成*
