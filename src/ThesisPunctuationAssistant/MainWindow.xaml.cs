using ThesisPunctuationAssistant.Services;
using ThesisPunctuationAssistant.ViewModels;

namespace ThesisPunctuationAssistant;

public partial class MainWindow : System.Windows.Window
{
    public MainWindow()
    {
        InitializeComponent();

        var ruleService = new RuleConfigService();
        var rules = ruleService.LoadDefault();
        var wordService = new WordInteropService();
        var detector = new PunctuationDetector(rules);
        var orchestrator = new ScanOrchestrator(wordService, detector);
        DataContext = new MainViewModel(orchestrator, wordService);
    }
}
