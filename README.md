# 文件格式转换工具 (FileConverterGUI)

一个简单易用的图形化文件格式转换工具，支持多种常见数据格式之间的相互转换。

## 功能特性

- 支持 Excel (.xlsx)、CSV (.csv)、JSON (.json)、Markdown (.md) 格式互转
- 自动识别文件格式（基于文件扩展名）
- 支持自定义编码（utf-8、gbk、gb2312、utf-16）
- 支持 Excel 工作表选择
- 支持 JSON 数据结构选择
- 实时转换日志显示
- 跨平台兼容，打包后可直接运行

## 支持转换格式

- Excel (.xlsx) -> Markdown (.md)
- Excel (.xlsx) -> CSV (.csv)
- CSV (.csv) -> Markdown (.md)
- CSV (.csv) -> Excel (.xlsx)
- JSON (.json) -> Markdown (.md)
- Markdown (.md) -> CSV (.csv)
- Markdown (.md) -> Excel (.xlsx)

## 安装依赖

运行程序前需要安装以下 Python 库：

```bash
pip install pandas openpyxl tabulate
```

## 使用方法

### 源码运行

```bash
python FileConverterGUI.py
```

### 直接运行 (Windows)

1. 下载 dist 目录中的 `FileConverterGUI.exe`
2. 双击运行即可

## 操作步骤

1. 点击"浏览"选择输入文件
2. 设置输出文件路径
3. 选择输入/输出格式（可自动识别）
4. 配置高级选项（可选）
5. 点击"开始转换"按钮
6. 查看转换日志和结果

## 打包成可执行文件

使用 PyInstaller 打包：

```bash
pip install pyinstaller
pyinstaller FileConverterGUI.spec
```

打包完成后，可执行文件位于 `dist/FileConverterGUI.exe`

## 技术栈

- Python 3
- tkinter (GUI 框架)
- pandas (数据处理)
- openpyxl (Excel 处理)
- tabulate (Markdown 表格生成)

## 项目结构

```
baiyun-assistant/
├── FileConverterGUI.py     # 主程序文件
├── FileConverterGUI.spec   # PyInstaller 配置文件
├── icon.ico                # 程序图标（可选）
└── dist/                   # 打包输出目录
    └── FileConverterGUI.exe
```

## 常见问题

### 1. 提示缺少依赖库

运行以下命令安装：
```bash
pip install pandas openpyxl tabulate
```

### 2. Excel 文件转换格式错乱

程序会自动处理单元格内的换行符，如果仍有问题请检查源文件格式。

### 3. 打包后图标不显示

确保项目目录中存在 `icon.ico` 文件，并在打包时正确配置。

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！
