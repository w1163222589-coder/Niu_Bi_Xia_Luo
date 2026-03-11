using System.Text.RegularExpressions;
using ThesisPunctuationAssistant.Models;

namespace ThesisPunctuationAssistant.Services;

public sealed class PunctuationDetector
{
    private readonly RulesConfig _config;
    private readonly Regex[] _allowPatterns;

    public PunctuationDetector(RulesConfig config)
    {
        _config = config;
        _allowPatterns = _config.AllowedRegexes.Select(r => new Regex(r, RegexOptions.Compiled)).ToArray();
    }

    public IEnumerable<PunctuationIssue> Detect(ParagraphUnit paragraph)
    {
        if (paragraph.Zone is ZoneType.EnglishSection or ZoneType.References)
            yield break;

        var text = paragraph.Text;
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (!_config.SuspiciousEnglishPunctuation.Contains(c.ToString()))
                continue;

            if (ShouldAllowByContext(text, i, paragraph.Zone))
                continue;

            var symbol = c.ToString();
            _config.SuggestionMap.TryGetValue(symbol, out var suggestion);

            yield return new PunctuationIssue
            {
                Symbol = c,
                ParagraphIndex = paragraph.Index,
                CharIndex = i,
                Context = ExtractContext(text, i),
                Zone = paragraph.Zone,
                Suggestion = suggestion ?? "请改为对应中文标点",
                Reason = BuildReason(text, i, paragraph.Zone),
                Severity = IssueSeverity.Error,
                WordRangeRef = paragraph.NativeRef
            };
        }
    }

    private bool ShouldAllowByContext(string text, int index, ZoneType zone)
    {
        var snippet = ExtractContext(text, index, 40);
        if (_allowPatterns.Any(p => p.IsMatch(snippet))) return true;

        if (zone == ZoneType.FigureOrEquation)
        {
            if (Regex.IsMatch(snippet, @"\([a-zA-Z]\)")) return true;
            if (Regex.IsMatch(snippet, @"式\(\d+-\d+\)")) return true;
        }

        if (Regex.IsMatch(snippet, @"\d+\.\d+")) return true; // 小数点
        if (Regex.IsMatch(snippet, @"^\s*\d+\.\s")) return true; // 条目编号
        if (Regex.IsMatch(snippet, @"\[\d+(,\s*\d+)*(-\d+)?\]")) return true; // 文献引用

        return false;
    }

    private static string BuildReason(string text, int index, ZoneType zone)
    {
        if (zone == ZoneType.ChineseBody)
            return "中文正文中检测到英文半角标点，且不在白名单上下文。";
        return "当前区域不建议使用该英文标点。";
    }

    private static string ExtractContext(string text, int index, int window = 12)
    {
        var start = Math.Max(0, index - window);
        var end = Math.Min(text.Length, index + window + 1);
        return text[start..end];
    }
}
