using ThesisPunctuationAssistant.Models;

namespace ThesisPunctuationAssistant.Services;

public interface IWordService
{
    bool ConnectRunningWord();
    bool OpenDocument(string filePath);
    string GetDocumentInfo();
    IReadOnlyList<ParagraphUnit> ReadParagraphs();
    void FocusIssue(PunctuationIssue issue);
    void HighlightIssue(PunctuationIssue issue);
    void ReplaceIssue(PunctuationIssue issue, string replacement);
}
