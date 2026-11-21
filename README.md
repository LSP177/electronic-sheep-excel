<div align="center">
  <img src="docs/logo.png" alt="电子羊表格 Logo" width="128" />
  <h1>电子羊表格 (Electronic Sheep Excel Helper)</h1>
  <p>面向 Windows 平台 Microsoft Excel 的双前端数据清洗、核对与脱敏工具（MIT 开源免费版）</p>
  <p>
    <a href="#-一键下载与安装">一键下载安装</a> ·
    <a href="#-让-ai-助手帮你安装-for-ai-agents">AI 一键代装</a> ·
    <a href="#核心特性">核心特性</a> ·
    <a href="#目录结构">目录结构</a> ·
    <a href="LICENSE">MIT License</a>
  </p>
</div>

**电子羊表格** 内置「操作前强制基线备份 + 失败自动熔断」安全机制，专为高频处理复杂脏表格的供应链、财务、采购及运营场景设计。基于 **MIT License** 完全免费开源，无任何授权校验与次数限制。

---

## ⚡ 一键下载与安装

> ⚠️ **运行前提**：仅支持 Windows 上的桌面版 **Microsoft Excel**（暂不支持 WPS、macOS 或网页版 Excel）；安装前请先关闭所有 Excel 窗口。

### 方式一：普通用户直接点击下载

| 版本 | 适用环境 | 直接下载链接 | 说明 |
| :--- | :--- | :--- | :--- |
| **现代版（推荐）** | 64 位 Excel + .NET 4.8 + WebView2 | [下载现代版 (.zip)](https://github.com/LSP177/electronic-sheep-excel/releases/latest/download/ElectronicSheepExcel-v1.0.0-oss-Modern-Setup.zip) | 带右侧边栏交互工作台，**请先解压再双击运行 exe** |
| **兼容版** | 32 位 / 64 位 Excel 通用 | [下载兼容版 (.exe)](https://github.com/LSP177/electronic-sheep-excel/releases/latest/download/ElectronicSheepExcel-v1.0.0-oss-Legacy-Setup.exe) | 顶部 Ribbon 经典菜单，无需 WebView2，直接双击安装 |
| **校验文件** | 可选完整性核对 | [SHA256SUMS.txt](https://github.com/LSP177/electronic-sheep-excel/releases/latest/download/SHA256SUMS.txt) | 核对安装包 SHA-256 哈希值 |

### 方式二：PowerShell 一行自动安装

会自动检测本地 Excel 位数（64 位优先现代版、32 位自动选兼容版）、下载 Release 安装包并自动校验 SHA-256：

```powershell
irm https://raw.githubusercontent.com/LSP177/electronic-sheep-excel/main/scripts/install-from-release.ps1 | iex
```

---

## 🤖 让 AI 助手帮你安装 (For AI Agents)

如果你正在使用 **Antigravity / Claude Code / Codex / Cursor** 等 AI 编程助手，只需要把仓库地址发给它并说一句：

> **“读取 `https://github.com/LSP177/electronic-sheep-excel` 里的 `README.md` 和 `AGENTS.md`，帮我自动检测 Excel 版本并安装好电子羊表格”**

AI 智能体只需运行仓库内置脚本 [`scripts/install-from-release.ps1`](scripts/install-from-release.ps1) 即可完成全自动探测、校验与静默安装（详细规则见 [`AGENTS.md`](AGENTS.md)）：

```powershell
# 已克隆仓库时执行：
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/install-from-release.ps1

# 或无需克隆直接远程安装：
irm https://raw.githubusercontent.com/LSP177/electronic-sheep-excel/main/scripts/install-from-release.ps1 | iex
```

---

## 核心特性

### 1. 安全优先：防误操作熔断架构
- **处理前基线自动备份**：每次批量修改前自动备份当天首次基线至 `%LOCALAPPDATA%\ExcelDataHelper\backups\`，并校验文件完整性（`DHBackupIntegrity` / `DHSafety`）。
- **备份失败立即熔断**：若备份写入或校验异常，引擎直接中止破坏性操作，确保原始工作簿零修改。
- **审计小票与明细日志**：每次执行均生成变更数量、跳过原因及影响地址的详细小票（`DHAudit` / `DHLog`），且默认不自动覆盖保存原文件。

### 2. 双前端架构（共享同一套核心引擎）
- **现代版（VSTO + WebView2 侧边栏）**：基于 C# VSTO 外接程序与嵌入式 WebView2（HTML/JS）提供向导式任务窗格，支持选区上下文感知、预览计划与交互确认（`frontends/taskpane/`）。
- **兼容版（Ribbon + VBA）**：原生 Office Fluent Ribbon (`customUI14.xml`) + 经典窗体，无需 .NET 4.8 与 WebView2 运行库，兼容 32/64 位 Microsoft Excel（`frontends/legacy/`）。

### 3. 核心业务模块（`core/business/`）
| 类别 | 模块 | 核心能力 |
| :--- | :--- | :--- |
| **体检与备份** | `DHCheck` / `DHReports` / `DHBackupIntegrity` / `DHSafety` | 当前表/全簿体检、问题扫描报告、基线备份与完整性校验 |
| **区域与空值清理** | `DHClean` / `DHBlankRows` / `DHBlankRegions` / `DHMerges` | 不可见字符/空格清理、安全删除空行与空区块、合并单元格拆分与保护 |
| **数值、单位与日期** | `DHNumbers` / `DHUnits` | 文本型数字转换、同列多单位拆分/统一换算、常见中文/斜杠日期规范化 |
| **核对与重组** | `DHDuplicates` / `DHLookup` / `DHReconcile` / `DHAppend` / `DHUnpivot` / `DHSplitSheet` | 重复值标记与去重、跨表补全、两表对账、多表追加、宽表逆透视、按列拆表 |
| **脱敏与交付** | `DHTransform` / `DHDelivery` / `DHMarks` | 手机号/身份证/姓名一键脱敏、交表礼仪规范检查、临时标记清除 |

---

## 目录结构

```text
electronic-sheep-excel/
├── core/
│   ├── business/        # 25 个核心 Excel 清洗、核对、脱敏与安全模块 (.bas)
│   └── runtime/         # VSTO/WebView2 桥接、事件监听、快照守卫与开源门面 (.bas/.cls)
├── frontends/
│   ├── taskpane/        # 现代版前端：C# VSTO 外接程序 + WebView2 侧边栏 UI
│   └── legacy/          # 兼容版前端：customUI14 Ribbon 菜单与经典 VBA 入口
├── installer/           # PowerShell 安装/卸载脚本与自解压引导程序 (Setup.cs)
├── scripts/             # 一键从 GitHub Releases 自动检测并安装的脚本
├── build/               # 自动化编译与打包管线 (VBA 引擎注入、VSTO 编译、安装包封装)
├── docs/                # 项目 Logo 与说明资源
├── AGENTS.md            # 给 AI 智能体的安全安装与操作规范
├── version.json         # 版本与开源版元数据
└── LICENSE              # MIT License
```

---

## 从源码构建（开发者）

> 普通用户直接通过上方 Releases 安装包使用即可，无需自行编译。如需修改源码重新构建 `.xlam` 与 `.vsto`：

1. **构建核心引擎**：
   ```powershell
   powershell -ExecutionPolicy Bypass -File build/build-engine.ps1 -BaselineDir "<path-to-seed-dir>"
   ```
2. **构建现代版 TaskPane（VSTO + WebView2）**：
   ```powershell
   powershell -ExecutionPolicy Bypass -File build/build-taskpane.ps1 -CertificateThumbprint "<your-codesign-thumbprint>" -AssetsDir "<path-to-vsto-deps>"
   ```
3. **打包双版安装程序**：
   ```powershell
   powershell -ExecutionPolicy Bypass -File build/package.ps1
   ```

---

## 开源协议

本项目基于 [MIT License](LICENSE) 开源。
