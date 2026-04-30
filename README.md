# SmartPicker v2.0

SmartPicker 是一个基于 Python 和 Tkinter 的智能抽取工具，支持从 Excel、CSV、TXT、Word 文档导入数据，并提供随机抽取、手动范围抽取、动画滚动、批量抽取和多语言界面功能。

## 主要功能

- 支持导入 `.xlsx`、`.xls`、`.csv`、`.txt` 和 `.docx` 数据文件
- 自动解析表格和段落文本，智能提取候选数据
- 支持手动输入起始和结束编号进行范围抽取
- 可配置抽取人数、结果字体大小、动画速度
- 支持重复抽取 / 不允许重复抽取两种模式
- 提供简体中文和 English 两种语言界面
- 支持拖拽文件导入

## 运行环境

- Python 3.8+
- Windows 系统（使用 `ctypes.windll.shcore.SetProcessDpiAwareness` 进行 DPI 适配）

## 依赖项

```bash
pip install ttkbootstrap pandas openpyxl python-docx tkinterdnd2
```

## 使用说明

1. 下载或克隆仓库到本地：

```bash
git clone <仓库地址>
cd SmartPicker
```

2. 安装依赖：

```bash
pip install ttkbootstrap pandas openpyxl python-docx tkinterdnd2
```

3. 运行程序：

```bash
python main.py
```

4. 在程序中：

- 点击“导入数据文件”选择支持的文件类型
- 或将数据文件拖拽到窗口中
- 如果没有导入数据，可手动输入“起始编号”和“结束编号”进行范围抽取
- 点击“开始抽取”进行随机抽取
- 点击“设置”可修改字体大小、抽取人数、动画速度、语言和重复抽取模式
- 点击“重置数据”清除已加载的数据和当前结果

## 支持文件格式

- Excel：`.xlsx`, `.xls`
- CSV：`.csv`
- 文本文件：`.txt`
- Word 文档：`.docx`

## 代码结构

- `main.py`：程序主入口和 GUI 实现
- `LICENSE`：项目开源许可证

## 许可证

本项目基于 MIT 许可证开源，详见 `LICENSE`。
