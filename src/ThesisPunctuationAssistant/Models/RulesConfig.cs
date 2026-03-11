namespace ThesisPunctuationAssistant.Models;

public sealed class RulesConfig
{
    public List<string> SuspiciousEnglishPunctuation { get; set; } = new();
    public List<string> AllowedRegexes { get; set; } = new();
    public Dictionary<string, string> SuggestionMap { get; set; } = new();
}
