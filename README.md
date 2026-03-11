# Word 论文标点交互式检查助手（MVP）

## 1) 推荐技术方案（先结论）

**首选：C# + WPF + Microsoft Word Interop（COM）**

原因：
1. 能直接控制本地 Word（定位、滚动、选中、高亮、替换）并保持稳定体验。
2. WPF 适合做“左侧状态 + 右侧问题详情 + 底部日志/进度”的桌面交互。
3. 与 Windows + Office 环境高度匹配，后续扩展（白名单编辑、规则配置、报告导出）成本低。

对比：
- Python + pywin32 可做，但部署和长期维护一致性略弱。
- Office Add-in/VSTO 更适合“嵌入 Word 内部”，不如独立桌面工具便于做多面板交互流程。

## 2) MVP 模块设计

- `WordInteropService`：连接运行中的 Word / 打开文档 / 读取段落 / 定位高亮 / 替换。
- `PunctuationDetector`：按“区域 + 上下文白名单”识别可疑英文标点。
- `ScanOrchestrator`：串联扫描流程，输出逐条问题（用于“发现即暂停”）。
- `MainViewModel`：管理按钮动作、状态文本、当前问题详情、问题列表与报告导出。
- `rules.default.json`：可配置规则（可疑符号、白名单正则、替换建议）。

区域策略（MVP）：
1. 中文正文（严格）
2. 图题/公式（分图 `(a)`、式 `(2-1)` 等白名单）
3. 英文段落（宽松）
4. 参考文献（宽松）

## 3) 第一版最小可运行代码框架

当前代码已实现：
- WPF 可视化界面骨架（连接 Word、打开文档、开始检测、继续、自动替换、导出报告）。
- Word COM 联动：定位、高亮、替换。
- 扫描中文正文和图表/公式区，命中问题后暂停并展示详情。
- JSON 规则配置文件，便于后续扩展。

## 项目结构

```text
src/
  ThesisPunctuationAssistant/
    Config/
      rules.default.json
    Models/
      IssueSeverity.cs
      ParagraphUnit.cs
      PunctuationIssue.cs
      RulesConfig.cs
      ZoneType.cs
    Services/
      IWordService.cs
      PunctuationDetector.cs
      RuleConfigService.cs
      ScanOrchestrator.cs
      WordInteropService.cs
    ViewModels/
      MainViewModel.cs
      RelayCommand.cs
    App.xaml
    App.xaml.cs
    MainWindow.xaml
    MainWindow.xaml.cs
    ThesisPunctuationAssistant.csproj
```

## 运行（Windows）

1. 安装 `.NET 8 SDK` 与 Microsoft Word。
2. 在项目目录执行：
   ```bash
   dotnet restore
   dotnet run --project src/ThesisPunctuationAssistant/ThesisPunctuationAssistant.csproj
   ```
3. 点击“连接当前 Word”或“打开文档”。
4. 点击“开始检测”，命中可疑标点后会定位并暂停。

## 后续增强建议

- 真实“继续扫描”游标（保存段落索引 + 字符偏移），支持逐条 next。
- Word 区域识别升级：样式名识别（标题、参考文献、脚注、英文摘要）。
- 白名单管理 UI（持久化到用户规则文件）。
- 导出 txt/csv/json 三格式，并增加问题类型统计。
- 增加“标题不得含标点”的专门规则。
