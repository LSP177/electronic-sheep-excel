# AI Agent 安装与操作指南

当用户把本仓库地址（`https://github.com/LSP177/electronic-sheep-excel`）发给你并要求**“帮我安装电子羊表格”**时，请严格按以下规范执行：

## 核心原则

1. **优先安装 Release 预编译包，不要从源码重新编译**：
   - 源码构建（`build/build-engine.ps1` 与 `build/build-taskpane.ps1`）需要本地预先提供 Office/VSTO 引用程序集与种子模板，仓库内不包含这些二进制依赖；
   - 普通用户安装只需使用仓库提供的 `scripts/install-from-release.ps1` 自动下载 GitHub Releases 预编译安装包并完成 SHA-256 校验与安装。
2. **严禁强杀用户正在运行的 `EXCEL.EXE` 进程**：
   - 如果检测到 `EXCEL.EXE` 正在运行，应提示用户先手动保存工作簿并关闭 Excel，再继续安装；绝不能擅自 `Stop-Process` 强杀进程以免造成表格数据丢失。
3. **仅支持 Windows 桌面版 Microsoft Excel**：
   - 不支持 WPS、macOS 或网页版 Excel；32 位 Excel 请安装兼容版（`legacy`），64 位 Excel 优先安装现代版（`modern`）。

## 一键安装命令（AI 执行）

如果用户已经克隆了本仓库，在仓库根目录直接执行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/install-from-release.ps1
```

如果用户尚未克隆仓库，可直接执行以下单行命令（无需 clone，自动探测 Excel 位数、校验 SHA-256、静默安装并输出日志）：

```powershell
irm https://raw.githubusercontent.com/LSP177/electronic-sheep-excel/main/scripts/install-from-release.ps1 | iex
```

或显式指定版本（`-Edition modern` 或 `-Edition legacy`）：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/install-from-release.ps1 -Edition modern
```

## 故障排查与日志位置

- 安装过程默认使用 `--no-launch` 静默执行（避免阻塞 GUI 弹窗）。
- 安装日志位于：`%LOCALAPPDATA%\DataHelper\SetupLogs\`。
- 卸载入口：桌面快捷方式「卸载电子羊表格 现代版」/「卸载电子羊表格 兼容版」，或运行 `%LOCALAPPDATA%\DataHelper\1.0.0\install.ps1 -Mode Uninstall`。
