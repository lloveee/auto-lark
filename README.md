# auto-lark

本项目基于 **Python 3.11**+ 开发，使用 `pyautogui`, `pandas`, `openpyxl` 等库实现自动化操作。  
本文档将指导你从零开始配置运行环境。

---

## 📦 一、环境要求

- Python 3.11 及以上  
- pip（Python 自带）  
- Git（用于克隆项目）

---

## 🚀 二、克隆项目

```bash
git clone https://github.com/lloveee/auto-lark.git
cd auto-lark
```

## 🧱 三、创建虚拟环境

### ▶ Windows：

```
python -m venv .venv
.venv\Scripts\activate
```

### ▶ macOS / Linux：

```
python3 -m venv .venv
source .venv/bin/activate
```

激活后，命令行前面会出现类似：

```
(.venv) $
```

------

## 📥 四、安装依赖

项目依赖在 `requirements.txt` 中定义，执行：

```
pip install --upgrade pip
pip install -r requirements.txt
```

如果网络较慢，可使用清华镜像：

```
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

------

## 📂 五、项目依赖列表

以下为主要依赖：

| 库名                  | 用途                  |
| --------------------- | --------------------- |
| pyautogui             | 自动化鼠标键盘操作    |
| pandas                | 处理 Excel 和表格数据 |
| openpyxl              | 操作 Excel 文件       |
| pyperclip / clipboard | 管理系统剪贴板        |
| Pillow                | 处理图片（PIL）       |
| keyboard              | 键盘监听              |
| tkinter               | 图形界面（标准库）    |