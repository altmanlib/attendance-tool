# 考勤统计桌面工具

用于解析钉钉或飞书导出的考勤 Excel，统计出勤、迟到、缺卡和加班情况，并导出汇总报表

## 普通用户使用

Windows 用户无需安装 Python。请从发布包中解压整个 `考勤统计工具` 文件夹，并双击：

```text
考勤统计工具.exe
```

> 不要只复制 `.exe` 文件，必须保留同级的 `_internal` 文件夹。若 Windows 弹出安全提示，选择「更多信息」→「仍要运行」

启动后：

1. 点击「📂 选择考勤Excel」，选择钉钉或飞书导出的考勤文件
2. 等待程序自动解析并展示统计结果
3. 在表格中选中员工后点击「🔍 查看明细」查看每日打卡记录；点击表头可排序
4. 点击「📤 导出汇总Excel」保存统计报表

## 打包为 Windows `.exe`

> 必须在 Windows 系统上执行打包。macOS 或 Linux 打包的程序不能在 Windows 上运行

### 1. 安装 uv

在 Windows PowerShell 中执行：

```powershell
winget install --id=astral-sh.uv -e
```

安装完成后，关闭并重新打开 PowerShell

### 2. 安装项目依赖

进入项目目录后执行：

```powershell
cd C:\路径\attendance-tool
uv sync --all-groups
```

项目已将 PyInstaller 配置为开发依赖，此步骤会一并安装

### 3. 执行打包

```powershell
uv run pyinstaller --noconfirm --clean --windowed --onedir --name 考勤统计工具 main.py
```

打包结果位于：

```text
dist\考勤统计工具\
├── 考勤统计工具.exe
└── _internal\
```

将整个 `dist\考勤统计工具` 文件夹压缩后交付给用户

`--onedir` 是推荐方式：启动更快，且被安全软件误报的概率通常低于单文件程序。若必须交付单个可执行文件，可将命令中的 `--onedir` 替换为 `--onefile`；生成的文件位于 `dist\考勤统计工具.exe`

## 使用 GitHub Actions 打包

代码推送到 GitHub 后，打开仓库的「Actions」→「Build Windows executable」→「Run workflow」，即可在 GitHub 的 Windows 环境中打包，无需本地 Windows 设备

构建完成后，在该运行记录底部的「Artifacts」下载 `attendance-tool-windows`。解压下载的 ZIP 后，将完整的 `考勤统计工具` 文件夹交付给用户

也可推送形如 `v1.0.0` 的 Git 标签自动触发构建

## 本地开发运行

需要安装 [uv](https://docs.astral.sh/uv/)，然后在项目目录执行：

```powershell
uv sync --all-groups
uv run main.py
```

## 功能

- 自动解析钉钉/飞书导出的考勤 Excel
- 统计每个人的出勤率、迟到、缺卡、加班
- 顶部卡片展示关键数据
- 支持按列排序和查看单人每日打卡明细
- 支持导出汇总 Excel
- 以红、橙、绿三色标注异常程度

## 开发环境

- Python 3.14+
- pandas
- openpyxl
- PyInstaller（仅打包时需要）
