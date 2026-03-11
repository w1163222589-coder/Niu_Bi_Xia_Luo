using Microsoft.Office.Interop.Word;
using ThesisPunctuationAssistant.Models;

namespace ThesisPunctuationAssistant.Services;

public sealed class WordInteropService : IWordService
{
    private Application? _app;
    private Document? _doc;

    public bool ConnectRunningWord()
    {
        try
        {
            _app = (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
            _doc = _app.ActiveDocument;
            return _doc is not null;
        }
        catch
        {
            return false;
        }
    }

    public bool OpenDocument(string filePath)
    {
        _app ??= new Application { Visible = true };
        _doc = _app.Documents.Open(filePath, ReadOnly: false, Visible: true);
        return _doc is not null;
    }

    public string GetDocumentInfo()
    {
        if (_doc is null) return "未连接 Word 文档";
        return $"文档：{_doc.Name} | 段落数：{_doc.Paragraphs.Count}";
    }

    public IReadOnlyList<ParagraphUnit> ReadParagraphs()
    {
        if (_doc is null) return Array.Empty<ParagraphUnit>();

        var result = new List<ParagraphUnit>(_doc.Paragraphs.Count);
        for (var i = 1; i <= _doc.Paragraphs.Count; i++)
        {
            var p = _doc.Paragraphs[i];
            var text = p.Range.Text?.TrimEnd('\r', '\n') ?? string.Empty;
            result.Add(new ParagraphUnit(i, text, DetectZone(text), p.Range));
        }

        return result;
    }

    public void FocusIssue(PunctuationIssue issue)
    {
        if (issue.WordRangeRef is not Range range) return;
        range.Select();
        _app?.Activate();
        _app?.Selection?.MoveRight(WdUnits.wdCharacter, issue.CharIndex, WdMovementType.wdMove);
        _app?.Selection?.Range?.Select();
    }

    public void HighlightIssue(PunctuationIssue issue)
    {
        if (issue.WordRangeRef is not Range range) return;
        var start = range.Start + issue.CharIndex;
        var end = start + 1;
        var tokenRange = range.Document.Range(start, end);
        tokenRange.HighlightColorIndex = WdColorIndex.wdYellow;
        tokenRange.Select();
    }

    public void ReplaceIssue(PunctuationIssue issue, string replacement)
    {
        if (issue.WordRangeRef is not Range range) return;
        var tokenRange = range.Document.Range(range.Start + issue.CharIndex, range.Start + issue.CharIndex + 1);
        tokenRange.Text = replacement;
    }

    private static ZoneType DetectZone(string text)
    {
        var t = text.Trim();
        if (t.StartsWith("参考文献") || t.StartsWith("References", StringComparison.OrdinalIgnoreCase))
            return ZoneType.References;
        if (t.StartsWith("图") || t.StartsWith("Fig.", StringComparison.OrdinalIgnoreCase) || t.Contains("式("))
            return ZoneType.FigureOrEquation;
        if (IsMostlyEnglish(t))
            return ZoneType.EnglishSection;
        return ZoneType.ChineseBody;
    }

    private static bool IsMostlyEnglish(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        var english = text.Count(c => c is >= 'A' and <= 'Z' or >= 'a' and <= 'z');
        return english > text.Length * 0.6;
    }
}
