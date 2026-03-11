using System.Text.Json;
using ThesisPunctuationAssistant.Models;

namespace ThesisPunctuationAssistant.Services;

public sealed class RuleConfigService
{
    private readonly string _defaultPath = Path.Combine(AppContext.BaseDirectory, "Config", "rules.default.json");

    public RulesConfig LoadDefault()
    {
        var text = File.ReadAllText(_defaultPath);
        return JsonSerializer.Deserialize<RulesConfig>(text) ?? new RulesConfig();
    }
}
