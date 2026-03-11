namespace ThesisPunctuationAssistant.Models;

public sealed record ParagraphUnit(int Index, string Text, ZoneType Zone, object NativeRef);
