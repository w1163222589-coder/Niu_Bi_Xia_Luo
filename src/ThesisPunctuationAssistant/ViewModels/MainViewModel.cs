using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;
using Microsoft.Win32;
using ThesisPunctuationAssistant.Models;
using ThesisPunctuationAssistant.Services;

namespace ThesisPunctuationAssistant.ViewModels;

public sealed class MainViewModel : INotifyPropertyChanged
{
    private readonly ScanOrchestrator _orchestrator;
    private readonly IWordService _wordService;
    private CancellationTokenSource? _cts;

    public ObservableCollection<PunctuationIssue> Issues { get; } = new();
    public RelayCommand ConnectWordCommand { get; }
    public RelayCommand OpenDocumentCommand { get; }
    public RelayCommand StartScanCommand { get; }
    public RelayCommand NextIssueCommand { get; }
    public RelayCommand AutoReplaceCommand { get; }
    public RelayCommand ExportReportCommand { get; }

    public string StatusText { get => _statusText; set { _statusText = value; OnPropertyChanged(); } }
    public string DocumentInfo { get => _documentInfo; set { _documentInfo = value; OnPropertyChanged(); } }
    public string ZoneModeInfo { get => _zoneModeInfo; set { _zoneModeInfo = value; OnPropertyChanged(); } }
    public string CurrentIssueSummary { get => _currentIssueSummary; set { _currentIssueSummary = value; OnPropertyChanged(); } }
    public string CurrentIssueReason { get => _currentIssueReason; set { _currentIssueReason = value; OnPropertyChanged(); } }
    public string CurrentIssueContext { get => _currentIssueContext; set { _currentIssueContext = value; OnPropertyChanged(); } }
    public string ProgressText { get => _progressText; set { _progressText = value; OnPropertyChanged(); } }

    private string _statusText = "待连接";
    private string _documentInfo = "未打开文档";
    private string _zoneModeInfo = "区域策略：中文正文 / 图表公式 / 英文段落 / 参考文献";
    private string _currentIssueSummary = "暂无问题";
    private string _currentIssueReason = string.Empty;
    private string _currentIssueContext = string.Empty;
    private string _progressText = "进度：0";

    public MainViewModel(ScanOrchestrator orchestrator, IWordService wordService)
    {
        _orchestrator = orchestrator;
        _wordService = wordService;

        ConnectWordCommand = new RelayCommand(ConnectWord);
        OpenDocumentCommand = new RelayCommand(OpenDocument);
        StartScanCommand = new RelayCommand(StartScan);
        NextIssueCommand = new RelayCommand(NextIssue, () => Issues.Count > 0);
        AutoReplaceCommand = new RelayCommand(AutoReplaceCurrent, () => Issues.Count > 0);
        ExportReportCommand = new RelayCommand(ExportReport, () => Issues.Count > 0);
    }

    private void ConnectWord()
    {
        StatusText = _wordService.ConnectRunningWord() ? "已连接到运行中的 Word" : "连接失败，请先打开 Word";
        DocumentInfo = _wordService.GetDocumentInfo();
    }

    private void OpenDocument()
    {
        var dialog = new OpenFileDialog { Filter = "Word 文档|*.docx" };
        if (dialog.ShowDialog() != true) return;

        var ok = _wordService.OpenDocument(dialog.FileName);
        StatusText = ok ? "文档已打开" : "打开失败";
        DocumentInfo = _wordService.GetDocumentInfo();
    }

    private async void StartScan()
    {
        Issues.Clear();
        _cts?.Cancel();
        _cts = new CancellationTokenSource();
        StatusText = "扫描中（遇到问题自动暂停）";

        await foreach (var issue in _orchestrator.ScanAsync(_cts.Token))
        {
            Issues.Add(issue);
            ShowIssue(issue);
            _wordService.FocusIssue(issue);
            _wordService.HighlightIssue(issue);
            StatusText = "已暂停：请点击“继续下一个”";
            ProgressText = $"已发现问题：{Issues.Count}";
            return;
        }

        StatusText = "扫描完成";
    }

    private void NextIssue()
    {
        if (Issues.Count == 0)
        {
            StartScan();
            return;
        }

        var issue = Issues[^1];
        ShowIssue(issue);
        _wordService.FocusIssue(issue);
    }

    private void AutoReplaceCurrent()
    {
        if (Issues.Count == 0) return;
        var issue = Issues[^1];
        _wordService.ReplaceIssue(issue, issue.Suggestion);
        StatusText = $"已自动替换：{issue.Symbol} -> {issue.Suggestion}";
    }

    private void ExportReport()
    {
        var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "punctuation_report.json");
        var json = JsonSerializer.Serialize(Issues, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(path, json);
        StatusText = $"报告已导出：{path}";
    }

    private void ShowIssue(PunctuationIssue issue)
    {
        CurrentIssueSummary = $"问题 #{issue.Index}: 段落 {issue.ParagraphIndex}, 符号 '{issue.Symbol}'";
        CurrentIssueReason = $"判定：{issue.Reason} 建议：{issue.Suggestion}";
        CurrentIssueContext = issue.Context;
        NextIssueCommand.RaiseCanExecuteChanged();
        AutoReplaceCommand.RaiseCanExecuteChanged();
        ExportReportCommand.RaiseCanExecuteChanged();
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged([CallerMemberName] string? name = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
