<div align="center">
  <img src="docs/logo.png" alt="电子羊表格 Logo" width="128" />
  <h1>电子羊表格 (Electronic Sheep Excel Helper)</h1>
  <p>面向 Windows 平台 Microsoft Excel 的双前端数据清洗、核对与脱敏工具（MIT 开源免费版）</p>
  <p>
    <a href="https://github.com/LSP177/electronic-sheep-excel/releases">下载免编译安装包 (Releases)</a> ·
    <a href="#核心特性">核心特性</a> ·
    <a href="#目录结构">目录结构</a> ·
    <a href="LICENSE">MIT License</a>
  </p>
</div>

**电子羊表格** 内置「操作前强制基线备份 + 失败自动熔断」安全机制，专为高频处理复杂脏表格的供应链、财务、采购及运营场景设计。

> 开源版基于 **MIT License** 完全免费开放，已移除所有授权校验与次数限制。普通用户可直接前往 [Releases 页面](https://github.com/LSP177/electronic-sheep-excel/releases) 下载一键安装程序，无需自行编译。

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
├── build/               # 自动化编译与打包管线 (VBA 引擎注入、VSTO 编译、安装包封装)
├── version.json         # 版本与开源版元数据
└── LICENSE              # MIT License
```

---

## 运行环境要求

- **操作系统**：Windows 10 / Windows 11
- **办公软件**：Microsoft Excel 桌面版（暂不支持 WPS、macOS 或网页版 Excel）
- **现代版额外依赖**：.NET Framework 4.8 与 Microsoft Edge WebView2 Runtime

---

## 从源码构建（开发者）

> 构建共享 `.xlam` 引擎需要本地安装 Microsoft Excel 桌面版，并传入包含种子文件（`DataHelper_PocEngine.xlam` 与 `TaskPaneHost.xlsx`）的目录。

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
