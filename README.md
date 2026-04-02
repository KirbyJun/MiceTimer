# MiceTimer

**MiceTimer** 是一款面向实验室的实验计时/计数辅助工具，基于 PySide6 构建，适用于三箱社交、自由社交等行为学范式实验。

---

## 功能特性

- 计时/计数项目管理（内置默认范式：三箱社交、自由社交、都做）
- 快捷键控制（可在设置页修改，支持全局快捷键）
- 自动保存与恢复（`data/autosave/recovery.json`）
- 模板保存/加载 + 最近模板快速加载（`data/recent_template.json`）
- 导出 Excel：
  - **一键导出**到 `data/export/`（软件所在目录）
  - **另存为**到任意位置
- 实验进行中红色状态条
- 项目名可编辑，Ctrl+Z 仅用于文本编辑撤销，不干预计时/计数

---

## 目录结构

```
MiceTimer/
├── main.py                  # 主程序入口
├── requirements.txt         # Python 依赖
├── build.bat                # Windows 打包脚本
├── version.txt              # 自动维护的版本号（打包时生成）
├── data/
│   ├── autosave/
│   │   └── recovery.json    # 自动保存（自动创建）
│   ├── export/              # 默认导出目录（自动创建）
│   └── templates/           # 用户保存的模板（自动创建）
└── releases/                # 打包输出目录（打包时生成）
```

---

## 本地运行

### 前提

- Python 3.10 或更高版本（推荐 3.11+）
- Windows 系统（快捷键功能依赖 `keyboard` 库，需 Windows/Linux；macOS 未测试）

### 步骤

```bash
# 1. 克隆仓库
git clone https://github.com/KirbyJun/MiceTimer.git
cd MiceTimer

# 2. 安装依赖
pip install -r requirements.txt

# 3. 运行程序
python main.py
```

---

## 打包（Windows）

双击或在命令行运行：

```bat
build.bat
```

打包完成后，可执行文件位于：

```
releases\MiceTimer_vX.Y.Z\MiceTimer.exe
```

> **注意**：打包使用 PyInstaller `--onefile` 模式，首次启动会有短暂解包延迟，属正常现象。

### 可选：UPX 压缩（进一步缩小体积）

1. 下载 [UPX](https://github.com/upx/upx/releases) for Windows
2. 解压后将 `upx.exe` 放到项目根目录下的 `upx/` 文件夹中：
   ```
   MiceTimer/upx/upx.exe
   ```
3. 再次运行 `build.bat`，脚本会自动检测并启用 UPX 压缩

---

## 常见问题

### bat 双击后闪退

原因：脚本遇到错误后立即退出。  
解决：在命令行（cmd）中手动执行 `build.bat`，查看完整报错信息。

### 全局热键不生效 / 需要管理员权限

`keyboard` 库在某些系统上需要管理员权限才能注册全局热键。  
解决：右键 `MiceTimer.exe` → "以管理员身份运行"。  
或在"设置"页关闭"全局快捷键"，改用窗口内快捷键。

### 杀软误报 / 被拦截

PyInstaller 打包的单文件 exe 可能触发杀软启发式检测。  
解决：将程序添加到杀软白名单，或使用源码直接运行（`python main.py`）。

### 导出的 Excel 在哪里

默认导出路径：`<程序所在目录>\data\export\`  
也可点击"另存为..."手动选择保存位置。