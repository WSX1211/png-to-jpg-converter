# GitHub 自动打包使用说明

## 📦 自动打包已配置完成！

本项目已配置 GitHub Actions 自动打包 Windows exe 文件。

## 🚀 使用步骤

### 1. 创建 GitHub 仓库

1. 登录 GitHub
2. 点击右上角 "+" → "New repository"
3. 填写仓库信息：
   - Repository name: `png-to-jpg-converter`
   - 选择 "Public" 或 "Private"
   - 不要勾选 "Add a README file"（我们已经有了）
   - 不要添加 .gitignore（我们已经有了）
4. 点击 "Create repository"

### 2. 推送代码到 GitHub

在项目文件夹中运行：

```bash
# 初始化 Git 仓库（如果还没有）
git init

# 添加所有文件
git add .

# 创建第一次提交
git commit -m "Initial commit: PNG to JPG converter"

# 添加远程仓库（替换为你的仓库地址）
git remote add origin https://github.com/你的用户名/png-to-jpg-converter.git

# 推送到 GitHub
git push -u origin main
```

### 3. 等待自动打包

推送后，GitHub Actions 会自动开始打包：

1. 点击仓库的 "Actions" 标签
2. 查看打包进度（通常需要 2-5 分钟）
3. 打包成功后，会出现绿色的 ✓

### 4. 下载 exe 文件

有两种下载方式：

**方式一：从 Releases 下载**
1. 点击仓库的 "Releases" 标签
2. 找到最新的 release
3. 下载 `PNG转JPG转换器.exe`

**方式二：从 Artifacts 下载**
1. 点击 "Actions" 标签
2. 点击最新的成功的工作流
3. 在页面底部 "Artifacts" 中下载 `PNG转JPG转换器-Windows`

## 🔧 手动触发打包

如果需要重新打包：

1. 点击 "Actions" 标签
2. 选择 "Build EXE for Windows"
3. 点击 "Run workflow" → "Run workflow"

## ⚠ 注意事项

1. **首次推送**：第一次推送会自动创建 v1.0.0 版本的 release
2. **后续更新**：每次 push 到 main 分支都会自动打包
3. **修改版本号**：如需修改版本号，编辑 `.github/workflows/build.yml` 中的 `tag_name`

## 📝 修改 README

记得修改 README.md 中的下载链接：

将：
```
https://github.com/你的用户名/png-to-jpg-converter/releases/latest
```

改为：
```
https://github.com/你的实际用户名/png-to-jpg-converter/releases/latest
```

## 🎉 完成！

现在你可以：
- 在任何电脑上下载 Windows exe 文件
- 无需 Python 环境即可使用
- 自动更新和发布新版本

## 文件结构

```
.
├── .github/
│   └── workflows/
│       └── build.yml          # 自动打包配置
├── png_to_jpg_converter.py    # 主程序
├── requirements.txt           # 依赖列表
├── build_exe.py              # 打包脚本
├── README.md                 # 项目说明
├── .gitignore               # Git 忽略文件
└── GitHub使用说明.md         # 本文件
```
