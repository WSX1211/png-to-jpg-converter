# PNG 转 JPG 转换器

一个简单易用的图片格式转换工具，可以将文件夹中的 PNG 图片批量转换为 JPG 格式。

## 功能特点

- ✅ 批量转换 PNG 到 JPG
- ✅ 支持拖拽文件夹（需要安装 tkinterdnd2）
- ✅ 自动处理透明背景（转换为白色背景）
- ✅ 实时显示转换进度
- ✅ 详细的转换日志
- ✅ 图形化界面，操作简单

## 下载 Windows 版本

[![Download Windows EXE](https://img.shields.io/badge/Download-Windows%20EXE-blue)](https://github.com/你的用户名/png-to-jpg-converter/releases/latest)

或者手动下载：
1. 点击上方的 "Releases"
2. 下载最新的 `PNG转JPG转换器.exe`

## 安装依赖（开发者）

```bash
pip install -r requirements.txt
```

或者手动安装：

```bash
pip install Pillow tkinterdnd2 pyinstaller
```

## 使用方法

### 方式一：下载 exe 文件（推荐）

1. 从 [Releases](https://github.com/你的用户名/png-to-jpg-converter/releases) 下载 `PNG转JPG转换器.exe`
2. 双击运行即可使用

### 方式二：直接运行 Python 脚本

```bash
python png_to_jpg_converter.py
```

### 方式三：自己打包成 exe 文件

1. 运行打包脚本：
```bash
python build_exe.py
```

2. 打包完成后，在 `dist` 文件夹中找到 `PNG转JPG转换器.exe`

## 操作步骤

1. 运行程序
2. 拖拽包含 PNG 图片的文件夹到窗口中
   - 或者点击「选择文件夹」按钮选择文件夹
3. 程序会自动转换所有 PNG 图片
4. 转换后的 JPG 图片保存在原文件夹的 `jpg_output` 子文件夹中

## 注意事项

- 原始 PNG 图片不会被修改
- 如果 PNG 图片有透明背景，会自动转换为白色背景
- JPG 图片质量设置为 95%（高质量）
- 如果没有安装 tkinterdnd2，拖拽功能不可用，但可以通过「选择文件夹」按钮使用

## 系统要求

- Windows 7 或更高版本（exe 版本）
- Python 3.7+（如果直接运行脚本）

## 开发者说明

### 自动打包

本项目使用 GitHub Actions 自动打包 Windows exe 文件：

1. Push 代码到 GitHub
2. GitHub Actions 自动在 Windows 环境中打包
3. 打包完成后，exe 文件会出现在 Releases 中

### 手动打包

```bash
pip install pyinstaller
python build_exe.py
```

打包后的 exe 文件可以在没有 Python 环境的 Windows 电脑上运行。

## 许可证

MIT License
