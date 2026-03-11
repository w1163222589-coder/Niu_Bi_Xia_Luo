namespace ThesisPunctuationAssistant.Models;

public sealed record PunctuationIssue
{
    public int Index { get; init; }
    public char Symbol { get; init; }
    public string Context { get; init; } = string.Empty;
    public string Reason { get; init; } = string.Empty;
    public string Suggestion { get; init; } = string.Empty;
    public ZoneType Zone { get; init; }
    public int ParagraphIndex { get; init; }
    public int CharIndex { get; init; }
    public IssueSeverity Severity { get; init; }
    public object? WordRangeRef { get; init; }
}
