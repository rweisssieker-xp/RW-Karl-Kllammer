using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using ClippyRWAvalonia.Models;
using ClippyRWAvalonia.ViewModels;
using ListBox = Avalonia.Controls.ListBox;

namespace ClippyRWAvalonia.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        Opened += (_, _) => UpdateLayoutMode(Bounds.Width);
        SizeChanged += (_, e) => UpdateLayoutMode(e.NewSize.Width);
        KeyDown += OnMainWindowKeyDown;
    }

    private MainWindowViewModel ViewModel => (MainWindowViewModel)DataContext!;

    private void UpdateLayoutMode(double width)
    {
        if (DataContext is MainWindowViewModel viewModel)
        {
            viewModel.IsCompact = width < 1260;
        }
    }

    private void OnRefreshClick(object? sender, RoutedEventArgs e) => ViewModel.Refresh();
    private void OnSaveSettingsClick(object? sender, RoutedEventArgs e) => ViewModel.SaveSettings();
    private void OnKnowledgeSearchClick(object? sender, RoutedEventArgs e) => ViewModel.FilterKnowledge();
    private void OnDeleteKnowledgeClick(object? sender, RoutedEventArgs e) => ViewModel.DeleteSelectedKnowledgeDocument();
    private void OnReindexKnowledgeClick(object? sender, RoutedEventArgs e) => ViewModel.ReindexKnowledge();
    private void OnSaveRecipeClick(object? sender, RoutedEventArgs e) => ViewModel.SaveSelectedRecipe();
    private void OnDeleteRecipeClick(object? sender, RoutedEventArgs e) => ViewModel.DeleteSelectedRecipe();
    private void OnCloneRecipeClick(object? sender, RoutedEventArgs e) => ViewModel.CloneSelectedRecipe();
    private void OnArchiveRecipeClick(object? sender, RoutedEventArgs e) => ViewModel.ArchiveSelectedRecipe();
    private void OnRitualFilterClick(object? sender, RoutedEventArgs e) => ViewModel.FilterRituals();
    private void OnPromoteHistoryClick(object? sender, RoutedEventArgs e) => ViewModel.PromoteHistoryToRecipe();
    private void OnPromoteWatchClick(object? sender, RoutedEventArgs e) => ViewModel.PromoteWatchToRecipe();
    private void OnSuggestKnowledgeRecipeClick(object? sender, RoutedEventArgs e) => ViewModel.SuggestRecipeFromKnowledge();
    private void OnStartTeachModeClick(object? sender, RoutedEventArgs e) => ViewModel.StartTeachMode();
    private void OnStopTeachModeClick(object? sender, RoutedEventArgs e) => ViewModel.StopTeachMode();
    private async void OnRunSelectedRecipeClick(object? sender, RoutedEventArgs e) => await ViewModel.RunSelectedRecipeAsync();
    private async void OnRunNextSelectedRecipeStepClick(object? sender, RoutedEventArgs e) => await ViewModel.RunNextSelectedRecipeStepAsync();
    private async void OnDryRunSelectedRecipeClick(object? sender, RoutedEventArgs e) => await ViewModel.DryRunSelectedRecipeAsync();
    private void OnAddOrUpdateRecipeStepClick(object? sender, RoutedEventArgs e) => ViewModel.AddOrUpdateRecipeStep();
    private void OnRemoveRecipeStepClick(object? sender, RoutedEventArgs e) => ViewModel.RemoveSelectedRecipeStep();
    private void OnDuplicateRecipeStepClick(object? sender, RoutedEventArgs e) => ViewModel.DuplicateSelectedRecipeStep();
    private void OnMoveRecipeStepUpClick(object? sender, RoutedEventArgs e) => ViewModel.MoveSelectedRecipeStep(-1);
    private void OnMoveRecipeStepDownClick(object? sender, RoutedEventArgs e) => ViewModel.MoveSelectedRecipeStep(1);
    private void OnAddOrUpdateRecipeParameterClick(object? sender, RoutedEventArgs e) => ViewModel.AddOrUpdateRecipeParameter();
    private void OnRemoveRecipeParameterClick(object? sender, RoutedEventArgs e) => ViewModel.RemoveSelectedRecipeParameter();
    private void OnHistoryFilterClick(object? sender, RoutedEventArgs e) => ViewModel.FilterHistory();
    private void OnDiagnosticsFilterClick(object? sender, RoutedEventArgs e) => ViewModel.FilterDiagnostics();
    private void OnExportDiagnosticsClick(object? sender, RoutedEventArgs e) => ViewModel.ExportDiagnostics();
    private void OnExportSupportBundleClick(object? sender, RoutedEventArgs e) => ViewModel.ExportSupportBundle();
    private void OnClearDiagnosticsClick(object? sender, RoutedEventArgs e) => ViewModel.ClearDiagnostics();
    private void OnRefreshActiveWindowClick(object? sender, RoutedEventArgs e) => ViewModel.RefreshActiveWindow();
    private async void OnRunSelectedAgentClick(object? sender, RoutedEventArgs e) => await ViewModel.RunSelectedAgentAsync();
    private async void OnAskAssistantClick(object? sender, RoutedEventArgs e) => await ViewModel.RunAssistantAsync();
    private void OnHandoffCodexClick(object? sender, RoutedEventArgs e) => ViewModel.AppendHandoffPrefix("nimm codex");
    private void OnHandoffCodexScreenClick(object? sender, RoutedEventArgs e) => ViewModel.AppendHandoffPrefix("nimm codex mit screen");
    private void OnHandoffClaudeCodeClick(object? sender, RoutedEventArgs e) => ViewModel.AppendHandoffPrefix("nimm claude code");
    private void OnHandoffOpenClawClick(object? sender, RoutedEventArgs e) => ViewModel.AppendHandoffPrefix("nimm openclaw");
    private async void OnAssistantSmokeTestClick(object? sender, RoutedEventArgs e) => await ViewModel.RunAssistantSmokeTestAsync();
    private async void OnSpeakResponseClick(object? sender, RoutedEventArgs e) => await ViewModel.SynthesizeCurrentResponseAsync();
    private async void OnRunActionPlanClick(object? sender, RoutedEventArgs e) => await ViewModel.RunPendingActionPlanAsync();
    private async void OnRunNextActionPlanStepClick(object? sender, RoutedEventArgs e) => await ViewModel.RunNextActionPlanStepAsync();
    private void OnSavePendingPlanAsRecipeClick(object? sender, RoutedEventArgs e) => ViewModel.SavePendingActionPlanAsRecipe();
    private void OnClearActionPlanClick(object? sender, RoutedEventArgs e) => ViewModel.ClearPendingActionPlan();
    private void OnClearConversationClick(object? sender, RoutedEventArgs e) => ViewModel.ClearConversation();
    private void OnStartSpeechCaptureClick(object? sender, RoutedEventArgs e) => ViewModel.StartSpeechCapture();
    private async void OnStopSpeechCaptureClick(object? sender, RoutedEventArgs e) => await ViewModel.StopSpeechCaptureAndAskAsync();
    private void OnCancelSpeechCaptureClick(object? sender, RoutedEventArgs e) => ViewModel.CancelSpeechCapture();
    private void OnInspectorListControlsClick(object? sender, RoutedEventArgs e) => ViewModel.RunInspectorAction("list_controls");
    private void OnInspectorReadFormClick(object? sender, RoutedEventArgs e) => ViewModel.RunInspectorAction("read_form");
    private void OnInspectorReadTableClick(object? sender, RoutedEventArgs e) => ViewModel.RunInspectorAction("read_table");
    private void OnInspectorReadDialogClick(object? sender, RoutedEventArgs e) => ViewModel.RunInspectorAction("read_dialog");
    private void OnInspectorReadSelectedRowClick(object? sender, RoutedEventArgs e) => ViewModel.RunInspectorAction("read_selected_row");
    private void OnInspectorFocusWindowClick(object? sender, RoutedEventArgs e) => ViewModel.RunInspectorAction("focus_window");
    private void OnAxReadContextClick(object? sender, RoutedEventArgs e) => ViewModel.RunInspectorAction("ax.read_context");
    private void OnAxReadFormClick(object? sender, RoutedEventArgs e) => ViewModel.RunInspectorAction("ax.read_form");
    private void OnAxReadGridClick(object? sender, RoutedEventArgs e) => ViewModel.RunInspectorAction("ax.read_grid");
    private void OnAxReadSelectedRowClick(object? sender, RoutedEventArgs e) => ViewModel.RunInspectorAction("ax.read_selected_row");
    private void OnAxConfirmDialogClick(object? sender, RoutedEventArgs e) => ViewModel.RunInspectorAction("ax.confirm_dialog");
    private void OnAxCancelDialogClick(object? sender, RoutedEventArgs e) => ViewModel.RunInspectorAction("ax.cancel_dialog");
    private void OnInspectorRunCustomClick(object? sender, RoutedEventArgs e) => ViewModel.RunInspectorAction();

    private async void OnImportKnowledgeClick(object? sender, RoutedEventArgs e)
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Import knowledge files",
            AllowMultiple = true,
            FileTypeFilter =
            [
                new FilePickerFileType("Supported knowledge files")
                {
                    Patterns = ["*.txt", "*.md", "*.log", "*.json", "*.csv", "*.pdf", "*.docx"]
                }
            ]
        });

        if (files.Count == 0)
        {
            return;
        }

        ViewModel.ImportKnowledgeFiles(files.Select(file => file.TryGetLocalPath()).Where(path => !string.IsNullOrWhiteSpace(path))!);
    }

    private async void OnImportAudioClick(object? sender, RoutedEventArgs e)
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Import audio for transcription",
            AllowMultiple = false,
            FileTypeFilter =
            [
                new FilePickerFileType("Audio files")
                {
                    Patterns = ["*.wav", "*.mp3", "*.m4a", "*.ogg"]
                }
            ]
        });

        var path = files.FirstOrDefault()?.TryGetLocalPath();
        if (!string.IsNullOrWhiteSpace(path))
        {
            await ViewModel.ImportAndTranscribeAudioAsync(path);
        }
    }

    private void OnKnowledgeSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is ListBox listBox)
        {
            ViewModel.SetSelectedKnowledgeDocument(listBox.SelectedItem as KnowledgeDocumentSummary);
        }
    }

    private void OnRecipeSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is ListBox listBox)
        {
            ViewModel.SetSelectedRecipe(listBox.SelectedItem as AutomationRecipe);
        }
    }

    private void OnWatchSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is ListBox listBox)
        {
            ViewModel.SetSelectedWatchSession(listBox.SelectedItem as WatchSessionEntry);
        }
    }

    private void OnHistorySelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is ListBox listBox)
        {
            ViewModel.SetSelectedHistoryEntry(listBox.SelectedItem as ActionHistoryEntry);
        }
    }

    private void OnDiagnosticsSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is ListBox listBox)
        {
            ViewModel.SetSelectedDiagnosticEntry(listBox.SelectedItem as DiagnosticEntry);
        }
    }

    private void OnRecipeStepSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is ListBox listBox)
        {
            ViewModel.SetSelectedRecipeStep(listBox.SelectedItem as RitualStep);
        }
    }

    private void OnRecipeParameterSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (sender is ListBox listBox)
        {
            ViewModel.SetSelectedRecipeParameter(listBox.SelectedItem as RitualParameter);
        }
    }

    private void OnOpenLiveContextTabClick(object? sender, RoutedEventArgs e) => ViewModel.NavigateToLiveContext();

    private void OnDashboardNavigateSetupClick(object? sender, RoutedEventArgs e) => ViewModel.NavigateToSetup();

    private void OnDashboardNavigateKnowledgeClick(object? sender, RoutedEventArgs e) => ViewModel.NavigateToKnowledge();

    private void OnDashboardNavigateLiveContextClick(object? sender, RoutedEventArgs e) => ViewModel.NavigateToLiveContext();

    private void OnDashboardNavigateRitualsClick(object? sender, RoutedEventArgs e) => ViewModel.NavigateToRituals(true);

    private void OnDashboardNavigateHistoryClick(object? sender, RoutedEventArgs e) => ViewModel.NavigateToHistory();

    private void OnDashboardNavigateProactiveClick(object? sender, RoutedEventArgs e) => ViewModel.NavigateToAsk();

    private void OnOnboardingEnvClick(object? sender, RoutedEventArgs e) => ViewModel.OnboardingCheckEnvironment();

    private void OnOnboardingKnowledgeClick(object? sender, RoutedEventArgs e) => ViewModel.OnboardingOpenKnowledgeFolder();

    private void OnOnboardingPttClick(object? sender, RoutedEventArgs e) => ViewModel.OnboardingTryPushToTalk();

    private void OnOnboardingAskSmokeClick(object? sender, RoutedEventArgs e) => ViewModel.OnboardingOpenAskSmoke();

    private async void OnMainWindowKeyDown(object? sender, Avalonia.Input.KeyEventArgs e)
    {
        if (!e.KeyModifiers.HasFlag(KeyModifiers.Control))
        {
            return;
        }

        if (e.Key == Key.L)
        {
            ViewModel.NavigateToLiveContext();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.K)
        {
            ViewModel.NavigateToKnowledge();
            e.Handled = true;
            return;
        }

        if (e.Key == Key.Enter)
        {
            await ViewModel.RunAssistantAsync();
            e.Handled = true;
        }
    }
}
