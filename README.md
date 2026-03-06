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
- **格式支持**：兼容 `.xlsx` 和 `.xls` 格式。

### 2. 字段提取 (Field Extraction)
从单个 Excel 文件中快速提取所需字段，生成新的 Excel 文件。
- **按需提取**：上传文件后，通过多选框选择需要保留的列。
- **样式保留**：导出时尽可能保留原单元格的字体、边框、背景色等样式。
- **自动列宽**：根据内容自动调整导出文件的列宽，方便查看。

## 🛠️ 安装指南

### 前置要求
- Python 3.8 或更高版本

### 安装步骤

1. 克隆或下载本项目到本地。
2. 进入项目根目录。
3. 安装依赖包：

```bash
pip install -r requirements.txt
```

## 🚀 使用方法

### 启动应用

在终端中运行以下命令启动应用：

```bash
streamlit run app.py
```
或者使用封装好的启动脚本：
```bash
python run_app.py
```

启动后，浏览器将自动打开应用的访问地址（通常是 `http://localhost:8501`）。

### 操作流程

1. **选择功能**：在左侧侧边栏选择“批量更新”或“字段提取”。
2. **上传文件**：根据提示上传您的 Excel 文件。
3. **配置参数**：
   - 设置表头所在行、数据起始行。
   - 选择匹配列和需要更新/提取的列。
4. **执行任务**：点击执行按钮，等待处理完成。
5. **下载结果**：处理完成后，直接下载生成的新文件。

## � 打包构建 (可选)

本项目包含 `run_app.spec` 配置，支持使用 PyInstaller 打包为可执行文件。

```bash
pyinstaller run_app.spec
```
打包完成后，可在 `dist/` 目录下找到生成的可执行程序。

## �📂 项目结构

```
excelApp/
├── app.py                  # 应用主入口
├── run_app.py              # 启动脚本封装
├── requirements.txt        # 项目依赖
├── pages/                  # 功能页面模块
│   ├── batch_update.py     # 批量更新功能逻辑
│   └── field_extraction.py # 字段提取功能逻辑
├── utils/                  # 工具函数
│   └── excel_helpers.py    # Excel 处理辅助函数
└── ...
```

## 🧰 技术栈

- **[Streamlit](https://streamlit.io/)**: 用于构建 Web 界面。
- **[Pandas](https://pandas.pydata.org/)**: 数据处理核心库。
- **[OpenPyXL](https://openpyxl.readthedocs.io/)**: 处理 `.xlsx` 文件。
- **[xlrd](https://xlrd.readthedocs.io/) / [xlutils](https://xlutils.readthedocs.io/)**: 处理旧版 `.xls` 文件。
- **[PyInstaller](https://pyinstaller.org/)**: 打包为独立应用。

---
*由小雷Excel批量助手生成*
