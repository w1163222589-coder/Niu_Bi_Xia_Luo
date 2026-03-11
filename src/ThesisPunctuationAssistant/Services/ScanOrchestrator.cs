using ThesisPunctuationAssistant.Models;

namespace ThesisPunctuationAssistant.Services;

public sealed class ScanOrchestrator
{
    private readonly IWordService _wordService;
    private readonly PunctuationDetector _detector;

    public ScanOrchestrator(IWordService wordService, PunctuationDetector detector)
    {
        _wordService = wordService;
        _detector = detector;
    }

    public async IAsyncEnumerable<PunctuationIssue> ScanAsync([System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct)
    {
        var paragraphs = _wordService.ReadParagraphs();
        var issueIndex = 0;

        foreach (var paragraph in paragraphs)
        {
            ct.ThrowIfCancellationRequested();

            foreach (var issue in _detector.Detect(paragraph))
            {
                issueIndex++;
                yield return issue with { Index = issueIndex };
                await Task.Yield();
            }
        }
    }
}
