using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Avalonia.Threading;
using CarolusNexus.Core;
using CarolusNexus.Platform.Windows;
using ClippyRWAvalonia.Models;
using ClippyRWAvalonia.Services;

namespace ClippyRWAvalonia.ViewModels;

public sealed class MainWindowViewModel : INotifyPropertyChanged
{
    private readonly OperatorWorkspaceService _workspaceService;
    private readonly LocalAgentRunService _runService;
    private readonly AssistantRuntimeService _assistantRuntimeService;
    private readonly MicrophoneRecorderService _microphoneRecorder;
    private readonly WindowsShellService _windowsShellService;
    private readonly CompanionOverlayService _companionOverlayService;
    private readonly DesktopInspectorService _desktopInspectorService;
    private readonly AxClientAutomationService _axClientAutomationService;
    private readonly RitualRuntimeService _ritualRuntimeService;
    private bool _isCompact;
    private string _statusMessage = "Ready.";
    private string _knowledgeSearch = string.Empty;
    private string _historySearch = string.Empty;
    private string _diagnosticsSearch = string.Empty;
    private string _selectedKnowledgePreview = string.Empty;
    private string _selectedHistoryDetail = string.Empty;
    private string _selectedDiagnosticDetail = string.Empty;
    private string _selectedWatchDetail = string.Empty;
    private string _provider = "anthropic";
    private string _mode = "companion";
    private string _model = string.Empty;
    private bool _speakResponses;
    private bool _useLocalKnowledge;
    private bool _suggestAutomations;
    private bool _autoRouteLocalAgents = true;
    private bool _speakAfterAsk;
    private string _selectedRecipeName = string.Empty;
    private string _selectedRecipePrompt = string.Empty;
    private string _selectedRecipeMode = "automation";
    private string _selectedRecipeDescription = string.Empty;
    private string _selectedRecipeCategory = "general";
    private string _selectedRecipeSourceType = "manual";
    private string _selectedRecipeRisk = "low";
    private string _selectedRecipeGuardApp = string.Empty;
    private string _selectedRecipeGuardForm = string.Empty;
    private string _selectedRecipeGuardDialog = string.Empty;
    private string _selectedRecipeGuardTab = string.Empty;
    private string _selectedRecipeTags = string.Empty;
    private string _selectedRecipeKnowledgeSources = string.Empty;
    private bool _selectedRecipeEnabled = true;
    private string _ritualSearch = string.Empty;
    private string _ritualCategoryFilter = "all";
    private string _ritualSourceFilter = "all";
    private string _ritualRiskFilter = "all";
    private string _ritualStats = string.Empty;
    private string _operatorMetrics = string.Empty;
    private string _proactiveSuggestionSummary = "no proactive ritual suggestion";
    private string _safetyPolicySummary = "low: full run • medium: preview first • high: step-by-step only";
    private string _selectedRitualRunLog = "no ritual run yet";
    private string _selectedRitualRunStatus = "idle";
    private int _nextRitualStepIndex;
    private string _ritualParameterName = string.Empty;
    private string _ritualParameterLabel = string.Empty;
    private string _ritualParameterDefaultValue = string.Empty;
    private bool _ritualParameterRequired = true;
    private string _ritualParameterKind = "text";
    private string _ritualStepActionType = "app";
    private string _ritualStepActionArgument = string.Empty;
    private string _ritualStepRisk = "low";
    private int _ritualStepWaitMs;
    private int _ritualStepRetryCount;
    private string _ritualStepIfApp = string.Empty;
    private string _ritualStepIfForm = string.Empty;
    private string _ritualStepIfDialog = string.Empty;
    private string _ritualStepIfTab = string.Empty;
    private string _ritualStepOnFail = "stop";
    private bool _teachModeActive;
    private string _teachModeStatus = "teach mode idle";
    private string _ritualSuggestionPreview = "no ritual suggestion yet";
    private int _teachModeBaselineHistoryCount;
    private int _selectedHistoryPromotionCount = 1;
    private string _envSummary = string.Empty;
    private string _activeAppSummary = "not loaded";
    private string _activeAppKind = "generic";
    private string _activeAppOneLiner = "active window: (refresh)";
    private string _liveContextFatClientHint = string.Empty;
    private string _uspRailScreen = "ask screen: on";
    private string _uspRailRag = "ask RAG: on";
    private string _uspRailRituals = "rituals: 0";
    private string _uspRailAx = "foreground: —";
    private string _uspRailHandoff = "auto-route: on";
    private string _uspRailVoice = "voice: off";
    private int _mainTabSelectedIndex;
    private string _axContextSummary = "no active AX context";
    private string _axSuggestedActions = "ax.* actions become available when an AX client is active";
    private string _consolePrompt = string.Empty;
    private string _selectedAgent = "codex";
    private string _consoleOutput = string.Empty;
    private string _consoleOutputPath = string.Empty;
    private bool _isConsoleBusy;
    private string _assistantPrompt = string.Empty;
    private string _assistantResponse = string.Empty;
    private string _retrievalSources = string.Empty;
    private string _transcriptText = string.Empty;
    private string _lastGeneratedAudioPath = string.Empty;
    private bool _includeScreens = true;
    private bool _useKnowledgeForAsk = true;
    private bool _isAssistantBusy;
    private bool _isRecordingSpeech;
    private string _speechCaptureState = "idle";
    private string _inspectorAction = "list_controls";
    private string _inspectorResult = string.Empty;
    private string _actionPlanPreview = "no pending action plan";
    private string _actionPlanExecutionLog = "no plan executed yet";
    private string _actionPlanState = "idle";
    private string _currentPlanStepSummary = "no active step";
    private int _nextPlanStepIndex;
    private AssistantActionPlan _pendingActionPlan = new();

    public MainWindowViewModel(
        OperatorWorkspaceService workspaceService,
        LocalAgentRunService runService,
        AssistantRuntimeService assistantRuntimeService,
        MicrophoneRecorderService microphoneRecorder,
        WindowsShellService windowsShellService,
        CompanionOverlayService companionOverlayService,
        DesktopInspectorService desktopInspectorService,
        AxClientAutomationService axClientAutomationService,
        RitualRuntimeService ritualRuntimeService)
    {
        _workspaceService = workspaceService;
        _runService = runService;
        _assistantRuntimeService = assistantRuntimeService;
        _microphoneRecorder = microphoneRecorder;
        _windowsShellService = windowsShellService;
        _companionOverlayService = companionOverlayService;
        _desktopInspectorService = desktopInspectorService;
        _axClientAutomationService = axClientAutomationService;
        _ritualRuntimeService = ritualRuntimeService;
        KnowledgeDocuments = new ObservableCollection<KnowledgeDocumentSummary>();
        Recipes = new ObservableCollection<AutomationRecipe>();
        WatchSessions = new ObservableCollection<WatchSessionEntry>();
        ActionHistory = new ObservableCollection<ActionHistoryEntry>();
        Diagnostics = new ObservableCollection<DiagnosticEntry>();
        ConversationHistory = new ObservableCollection<ConversationTurn>();
        SelectedRecipeSteps = new ObservableCollection<RitualStep>();
        SelectedRecipeParameters = new ObservableCollection<RitualParameter>();
        CompanionModes = ["companion", "agent", "automation", "watch"];
        ProviderOptions = ["anthropic", "openai", "openai-compatible"];
        AgentOptions = ["codex", "claude-code", "openclaw"];
        Refresh();
        RefreshActiveWindow();
        _companionOverlayService.SetState("ready", "operator surface ready");
    }

    public event PropertyChangedEventHandler? PropertyChanged;

public string Title => "Carolus Nexus";
public string Subtitle => "desktop operator surface for Karl Klammer";
    public string LayoutMode => _isCompact ? "compact stack" : "wide operator";

    public string RepoRoot => _workspaceService.RepoRoot;
    public string EnvPath => _workspaceService.EnvPath;

    public string StatusMessage
    {
        get => _statusMessage;
        set => SetField(ref _statusMessage, value);
    }

    public string Provider
    {
        get => _provider;
        set => SetField(ref _provider, value);
    }

    public string Mode
    {
        get => _mode;
        set => SetField(ref _mode, value);
    }

    public string Model
    {
        get => _model;
        set => SetField(ref _model, value);
    }

    public bool SpeakResponses
    {
        get => _speakResponses;
        set
        {
            if (!SetField(ref _speakResponses, value))
            {
                return;
            }

            UpdateUspRail();
        }
    }

    public bool UseLocalKnowledge
    {
        get => _useLocalKnowledge;
        set => SetField(ref _useLocalKnowledge, value);
    }

    public bool SuggestAutomations
    {
        get => _suggestAutomations;
        set => SetField(ref _suggestAutomations, value);
    }

    public bool AutoRouteLocalAgents
    {
        get => _autoRouteLocalAgents;
        set
        {
            if (!SetField(ref _autoRouteLocalAgents, value))
            {
                return;
            }

            UpdateUspRail();
        }
    }

    public bool SpeakAfterAsk
    {
        get => _speakAfterAsk;
        set
        {
            if (!SetField(ref _speakAfterAsk, value))
            {
                return;
            }

            UpdateUspRail();
        }
    }

    public string EnvironmentState { get; private set; } = "env missing";
    public string KnowledgeState { get; private set; } = "local knowledge off";
    public string SpeechState { get; private set; } = "voice responses off";
    public string AutomationState { get; private set; } = "ritual capture off";
    public string RuntimeSummary { get; private set; } = string.Empty;
    public string KnowledgeStatus { get; private set; } = string.Empty;
    public string CountsSummary { get; private set; } = string.Empty;

    public string EnvSummary
    {
        get => _envSummary;
        set => SetField(ref _envSummary, value);
    }

    public string ActiveAppSummary
    {
        get => _activeAppSummary;
        set => SetField(ref _activeAppSummary, value);
    }

    /// <summary>Compact line for hero: display name and app kind.</summary>
    public string ActiveAppOneLiner
    {
        get => _activeAppOneLiner;
        private set => SetField(ref _activeAppOneLiner, value);
    }

    public string ActiveAppKind
    {
        get => _activeAppKind;
        private set => SetField(ref _activeAppKind, value);
    }

    /// <summary>Non-AX fat-client hint for Live Context; empty when not applicable.</summary>
    public string LiveContextFatClientHint
    {
        get => _liveContextFatClientHint;
        private set => SetField(ref _liveContextFatClientHint, value);
    }

    public string UspRailScreen
    {
        get => _uspRailScreen;
        private set => SetField(ref _uspRailScreen, value);
    }

    public string UspRailRag
    {
        get => _uspRailRag;
        private set => SetField(ref _uspRailRag, value);
    }

    public string UspRailRituals
    {
        get => _uspRailRituals;
        private set => SetField(ref _uspRailRituals, value);
    }

    public string UspRailAx
    {
        get => _uspRailAx;
        private set => SetField(ref _uspRailAx, value);
    }

    public string UspRailHandoff
    {
        get => _uspRailHandoff;
        private set => SetField(ref _uspRailHandoff, value);
    }

    public string UspRailVoice
    {
        get => _uspRailVoice;
        private set => SetField(ref _uspRailVoice, value);
    }

    public int MainTabSelectedIndex
    {
        get => _mainTabSelectedIndex;
        set => SetField(ref _mainTabSelectedIndex, value);
    }

    public const int TabAsk = 0;
    public const int TabDashboard = 1;
    public const int TabSetup = 2;
    public const int TabKnowledge = 3;
    public const int TabRituals = 4;
    public const int TabHistory = 5;
    public const int TabDiagnostics = 6;
    public const int TabConsole = 7;
    public const int TabLiveContext = 8;

    public string AxContextSummary
    {
        get => _axContextSummary;
        set => SetField(ref _axContextSummary, value);
    }

    public string AxSuggestedActions
    {
        get => _axSuggestedActions;
        set => SetField(ref _axSuggestedActions, value);
    }

    public string ActionPlanPreview
    {
        get => _actionPlanPreview;
        set => SetField(ref _actionPlanPreview, value);
    }

    public string ActionPlanExecutionLog
    {
        get => _actionPlanExecutionLog;
        set => SetField(ref _actionPlanExecutionLog, value);
    }

    public string ActionPlanState
    {
        get => _actionPlanState;
        set => SetField(ref _actionPlanState, value);
    }

    public string CurrentPlanStepSummary
    {
        get => _currentPlanStepSummary;
        set => SetField(ref _currentPlanStepSummary, value);
    }

    public string RitualSearch
    {
        get => _ritualSearch;
        set => SetField(ref _ritualSearch, value);
    }

    public string RitualCategoryFilter
    {
        get => _ritualCategoryFilter;
        set => SetField(ref _ritualCategoryFilter, value);
    }

    public string RitualSourceFilter
    {
        get => _ritualSourceFilter;
        set => SetField(ref _ritualSourceFilter, value);
    }

    public string RitualRiskFilter
    {
        get => _ritualRiskFilter;
        set => SetField(ref _ritualRiskFilter, value);
    }

    public string RitualStats
    {
        get => _ritualStats;
        set => SetField(ref _ritualStats, value);
    }

    public string OperatorMetrics
    {
        get => _operatorMetrics;
        set => SetField(ref _operatorMetrics, value);
    }

    public string ProactiveSuggestionSummary
    {
        get => _proactiveSuggestionSummary;
        set => SetField(ref _proactiveSuggestionSummary, value);
    }

    public string SafetyPolicySummary
    {
        get => _safetyPolicySummary;
        set => SetField(ref _safetyPolicySummary, value);
    }

    public string SelectedRecipeDescription
    {
        get => _selectedRecipeDescription;
        set => SetField(ref _selectedRecipeDescription, value);
    }

    public string SelectedRecipeCategory
    {
        get => _selectedRecipeCategory;
        set => SetField(ref _selectedRecipeCategory, value);
    }

    public string SelectedRecipeSourceType
    {
        get => _selectedRecipeSourceType;
        set => SetField(ref _selectedRecipeSourceType, value);
    }

    public string SelectedRecipeRisk
    {
        get => _selectedRecipeRisk;
        set => SetField(ref _selectedRecipeRisk, value);
    }

    public string SelectedRecipeGuardApp
    {
        get => _selectedRecipeGuardApp;
        set => SetField(ref _selectedRecipeGuardApp, value);
    }

    public string SelectedRecipeGuardForm
    {
        get => _selectedRecipeGuardForm;
        set => SetField(ref _selectedRecipeGuardForm, value);
    }

    public string SelectedRecipeGuardDialog
    {
        get => _selectedRecipeGuardDialog;
        set => SetField(ref _selectedRecipeGuardDialog, value);
    }

    public string SelectedRecipeGuardTab
    {
        get => _selectedRecipeGuardTab;
        set => SetField(ref _selectedRecipeGuardTab, value);
    }

    public string SelectedRecipeTags
    {
        get => _selectedRecipeTags;
        set => SetField(ref _selectedRecipeTags, value);
    }

    public string SelectedRecipeKnowledgeSources
    {
        get => _selectedRecipeKnowledgeSources;
        set => SetField(ref _selectedRecipeKnowledgeSources, value);
    }

    public bool SelectedRecipeEnabled
    {
        get => _selectedRecipeEnabled;
        set => SetField(ref _selectedRecipeEnabled, value);
    }

    public string SelectedRitualRunLog
    {
        get => _selectedRitualRunLog;
        set => SetField(ref _selectedRitualRunLog, value);
    }

    public string SelectedRitualRunStatus
    {
        get => _selectedRitualRunStatus;
        set => SetField(ref _selectedRitualRunStatus, value);
    }

    public string RitualSuggestionPreview
    {
        get => _ritualSuggestionPreview;
        set => SetField(ref _ritualSuggestionPreview, value);
    }

    public bool TeachModeActive
    {
        get => _teachModeActive;
        set => SetField(ref _teachModeActive, value);
    }

    public string TeachModeStatus
    {
        get => _teachModeStatus;
        set => SetField(ref _teachModeStatus, value);
    }

    public int SelectedHistoryPromotionCount
    {
        get => _selectedHistoryPromotionCount;
        set
        {
            if (SetField(ref _selectedHistoryPromotionCount, Math.Max(1, value)))
            {
                OnPropertyChanged(nameof(SelectedHistoryPromotionCountText));
            }
        }
    }

    public string SelectedHistoryPromotionCountText
    {
        get => SelectedHistoryPromotionCount.ToString();
        set
        {
            if (int.TryParse(value, out var parsed))
            {
                SelectedHistoryPromotionCount = parsed;
            }
        }
    }

    public string RitualParameterName
    {
        get => _ritualParameterName;
        set => SetField(ref _ritualParameterName, value);
    }

    public string RitualParameterLabel
    {
        get => _ritualParameterLabel;
        set => SetField(ref _ritualParameterLabel, value);
    }

    public string RitualParameterDefaultValue
    {
        get => _ritualParameterDefaultValue;
        set => SetField(ref _ritualParameterDefaultValue, value);
    }

    public bool RitualParameterRequired
    {
        get => _ritualParameterRequired;
        set => SetField(ref _ritualParameterRequired, value);
    }

    public string RitualParameterKind
    {
        get => _ritualParameterKind;
        set => SetField(ref _ritualParameterKind, value);
    }

    public string RitualStepActionType
    {
        get => _ritualStepActionType;
        set => SetField(ref _ritualStepActionType, value);
    }

    public string RitualStepActionArgument
    {
        get => _ritualStepActionArgument;
        set => SetField(ref _ritualStepActionArgument, value);
    }

    public string RitualStepRisk
    {
        get => _ritualStepRisk;
        set => SetField(ref _ritualStepRisk, value);
    }

    public int RitualStepWaitMs
    {
        get => _ritualStepWaitMs;
        set
        {
            if (SetField(ref _ritualStepWaitMs, Math.Max(0, value)))
            {
                OnPropertyChanged(nameof(RitualStepWaitMsText));
            }
        }
    }

    public string RitualStepWaitMsText
    {
        get => RitualStepWaitMs.ToString();
        set
        {
            if (int.TryParse(value, out var parsed))
            {
                RitualStepWaitMs = parsed;
            }
        }
    }

    public int RitualStepRetryCount
    {
        get => _ritualStepRetryCount;
        set
        {
            if (SetField(ref _ritualStepRetryCount, Math.Max(0, value)))
            {
                OnPropertyChanged(nameof(RitualStepRetryCountText));
            }
        }
    }

    public string RitualStepRetryCountText
    {
        get => RitualStepRetryCount.ToString();
        set
        {
            if (int.TryParse(value, out var parsed))
            {
                RitualStepRetryCount = parsed;
            }
        }
    }

    public string RitualStepIfApp
    {
        get => _ritualStepIfApp;
        set => SetField(ref _ritualStepIfApp, value);
    }

    public string RitualStepIfForm
    {
        get => _ritualStepIfForm;
        set => SetField(ref _ritualStepIfForm, value);
    }

    public string RitualStepIfDialog
    {
        get => _ritualStepIfDialog;
        set => SetField(ref _ritualStepIfDialog, value);
    }

    public string RitualStepIfTab
    {
        get => _ritualStepIfTab;
        set => SetField(ref _ritualStepIfTab, value);
    }

    public string RitualStepOnFail
    {
        get => _ritualStepOnFail;
        set => SetField(ref _ritualStepOnFail, value);
    }

    public bool HasPendingActionPlan => _pendingActionPlan.Steps.Count > 0;
    public bool HasRemainingPlanSteps => _nextPlanStepIndex < _pendingActionPlan.Steps.Count;
    public bool HasSelectedRecipe => SelectedRecipe != null;
    public bool HasSelectedRecipeSteps => SelectedRecipeSteps.Count > 0;
    public bool HasSelectedRecipeParameters => SelectedRecipeParameters.Count > 0;

    public ObservableCollection<KnowledgeDocumentSummary> KnowledgeDocuments { get; }
    public ObservableCollection<AutomationRecipe> Recipes { get; }
    public ObservableCollection<RitualStep> SelectedRecipeSteps { get; }
    public ObservableCollection<RitualParameter> SelectedRecipeParameters { get; }
    public ObservableCollection<WatchSessionEntry> WatchSessions { get; }
    public ObservableCollection<ActionHistoryEntry> ActionHistory { get; }
    public ObservableCollection<DiagnosticEntry> Diagnostics { get; }
    public ObservableCollection<ConversationTurn> ConversationHistory { get; }

    public IReadOnlyList<string> CompanionModes { get; }
    public IReadOnlyList<string> ProviderOptions { get; }
    public IReadOnlyList<string> AgentOptions { get; }

    public KnowledgeDocumentSummary? SelectedKnowledgeDocument { get; private set; }
    public AutomationRecipe? SelectedRecipe { get; private set; }
    public RitualStep? SelectedRecipeStep { get; private set; }
    public RitualParameter? SelectedRecipeParameter { get; private set; }
    public WatchSessionEntry? SelectedWatchSession { get; private set; }
    public ActionHistoryEntry? SelectedHistoryEntry { get; private set; }
    public DiagnosticEntry? SelectedDiagnosticEntry { get; private set; }

    public string KnowledgeSearch
    {
        get => _knowledgeSearch;
        set => SetField(ref _knowledgeSearch, value);
    }

    public string HistorySearch
    {
        get => _historySearch;
        set => SetField(ref _historySearch, value);
    }

    public string DiagnosticsSearch
    {
        get => _diagnosticsSearch;
        set => SetField(ref _diagnosticsSearch, value);
    }

    public string SelectedKnowledgePreview
    {
        get => _selectedKnowledgePreview;
        set => SetField(ref _selectedKnowledgePreview, value);
    }

    public string SelectedHistoryDetail
    {
        get => _selectedHistoryDetail;
        set => SetField(ref _selectedHistoryDetail, value);
    }

    public string SelectedDiagnosticDetail
    {
        get => _selectedDiagnosticDetail;
        set => SetField(ref _selectedDiagnosticDetail, value);
    }

    public string SelectedWatchDetail
    {
        get => _selectedWatchDetail;
        set => SetField(ref _selectedWatchDetail, value);
    }

    public string SelectedRecipeName
    {
        get => _selectedRecipeName;
        set => SetField(ref _selectedRecipeName, value);
    }

    public string SelectedRecipePrompt
    {
        get => _selectedRecipePrompt;
        set => SetField(ref _selectedRecipePrompt, value);
    }

    public string SelectedRecipeMode
    {
        get => _selectedRecipeMode;
        set => SetField(ref _selectedRecipeMode, value);
    }

    public string ConsolePrompt
    {
        get => _consolePrompt;
        set => SetField(ref _consolePrompt, value);
    }

    public string SelectedAgent
    {
        get => _selectedAgent;
        set => SetField(ref _selectedAgent, value);
    }

    public string ConsoleOutput
    {
        get => _consoleOutput;
        set => SetField(ref _consoleOutput, value);
    }

    public string ConsoleOutputPath
    {
        get => _consoleOutputPath;
        set => SetField(ref _consoleOutputPath, value);
    }

    public bool IsConsoleBusy
    {
        get => _isConsoleBusy;
        set => SetField(ref _isConsoleBusy, value);
    }

    public string AssistantPrompt
    {
        get => _assistantPrompt;
        set => SetField(ref _assistantPrompt, value);
    }

    public string AssistantResponse
    {
        get => _assistantResponse;
        set => SetField(ref _assistantResponse, value);
    }

    public string RetrievalSources
    {
        get => _retrievalSources;
        set => SetField(ref _retrievalSources, value);
    }

    public string TranscriptText
    {
        get => _transcriptText;
        set => SetField(ref _transcriptText, value);
    }

    public string LastGeneratedAudioPath
    {
        get => _lastGeneratedAudioPath;
        set => SetField(ref _lastGeneratedAudioPath, value);
    }

    public bool IncludeScreens
    {
        get => _includeScreens;
        set
        {
            if (!SetField(ref _includeScreens, value))
            {
                return;
            }

            UpdateUspRail();
        }
    }

    public bool UseKnowledgeForAsk
    {
        get => _useKnowledgeForAsk;
        set
        {
            if (!SetField(ref _useKnowledgeForAsk, value))
            {
                return;
            }

            UpdateUspRail();
        }
    }

    public bool IsAssistantBusy
    {
        get => _isAssistantBusy;
        set => SetField(ref _isAssistantBusy, value);
    }

    public bool IsRecordingSpeech
    {
        get => _isRecordingSpeech;
        set => SetField(ref _isRecordingSpeech, value);
    }

    public string SpeechCaptureState
    {
        get => _speechCaptureState;
        set => SetField(ref _speechCaptureState, value);
    }

    public bool IsCompact
    {
        get => _isCompact;
        set
        {
            if (SetField(ref _isCompact, value))
            {
                OnPropertyChanged(nameof(LayoutMode));
            }
        }
    }

    public string InspectorAction
    {
        get => _inspectorAction;
        set => SetField(ref _inspectorAction, value);
    }

    public string InspectorResult
    {
        get => _inspectorResult;
        set => SetField(ref _inspectorResult, value);
    }

    public void Refresh()
    {
        var snapshot = _workspaceService.Load();
        Provider = snapshot.Provider;
        Model = snapshot.Model;
        Mode = snapshot.Mode;
        SpeakResponses = snapshot.SpeakResponses;
        UseLocalKnowledge = snapshot.UseLocalKnowledge;
        SuggestAutomations = snapshot.SuggestAutomations;
        AutoRouteLocalAgents = snapshot.AutoRouteLocalAgents;
        SpeakAfterAsk = snapshot.SpeakAfterAsk;
        EnvironmentState = snapshot.EnvExists ? "env loaded" : "env missing";
        KnowledgeState = snapshot.UseLocalKnowledge ? "local knowledge on" : "local knowledge off";
        SpeechState = snapshot.SpeakResponses
            ? (snapshot.SpeakAfterAsk ? "voice on • auto-speak after ask" : "voice responses on")
            : "voice responses off";
        AutomationState = snapshot.SuggestAutomations ? "ritual capture on" : "ritual capture off";
        RuntimeSummary = snapshot.RuntimeSummary;
        KnowledgeStatus = snapshot.KnowledgeStatus;
        CountsSummary = $"{snapshot.Recipes.Count} recipes • {snapshot.WatchSessions.Count} watch sessions • {snapshot.ActionHistory.Count} actions • {snapshot.Diagnostics.Count} diagnostics";
        EnvSummary = BuildEnvSummary(snapshot.EnvValues);

        ReplaceCollection(KnowledgeDocuments, _workspaceService.GetKnowledgeDocuments(KnowledgeSearch));
        ReplaceCollection(Recipes, FilterRecipes(snapshot.Recipes));
        ReplaceCollection(WatchSessions, snapshot.WatchSessions);
        ReplaceCollection(ActionHistory, FilterHistory(snapshot.ActionHistory, HistorySearch));
        ReplaceCollection(Diagnostics, FilterDiagnostics(snapshot.Diagnostics, DiagnosticsSearch));
        RitualStats = BuildRitualStats(snapshot.Recipes);
        OperatorMetrics = BuildOperatorMetrics(snapshot.Recipes, snapshot.ActionHistory, snapshot.KnowledgeDocuments);

        if (SelectedKnowledgeDocument == null && KnowledgeDocuments.Count > 0)
        {
            SetSelectedKnowledgeDocument(KnowledgeDocuments[0]);
        }

        if (SelectedRecipe == null && Recipes.Count > 0)
        {
            SetSelectedRecipe(Recipes[0]);
        }

        if (SelectedWatchSession == null && WatchSessions.Count > 0)
        {
            SetSelectedWatchSession(WatchSessions[0]);
        }

        if (SelectedHistoryEntry == null && ActionHistory.Count > 0)
        {
            SetSelectedHistoryEntry(ActionHistory[0]);
        }

        if (SelectedDiagnosticEntry == null && Diagnostics.Count > 0)
        {
            SetSelectedDiagnosticEntry(Diagnostics[0]);
        }

        OnPropertyChanged(nameof(EnvironmentState));
        OnPropertyChanged(nameof(KnowledgeState));
        OnPropertyChanged(nameof(SpeechState));
        OnPropertyChanged(nameof(AutomationState));
        OnPropertyChanged(nameof(RuntimeSummary));
        OnPropertyChanged(nameof(KnowledgeStatus));
        OnPropertyChanged(nameof(CountsSummary));
        OnPropertyChanged(nameof(HasPendingActionPlan));
        UpdateUspRail();
    }

    public void SaveSettings()
    {
        _workspaceService.SaveSettings(
            Provider,
            Model,
            Mode,
            SpeakResponses,
            UseLocalKnowledge,
            SuggestAutomations,
            AutoRouteLocalAgents,
            SpeakAfterAsk);
        StatusMessage = "Saved Avalonia operator settings.";
        Refresh();
    }

    public void SetSelectedKnowledgeDocument(KnowledgeDocumentSummary? document)
    {
        SelectedKnowledgeDocument = document;
        SelectedKnowledgePreview = document == null ? string.Empty : _workspaceService.GetKnowledgePreview(document.RelativePath);
        OnPropertyChanged(nameof(SelectedKnowledgeDocument));
    }

    public void SetSelectedRecipe(AutomationRecipe? recipe)
    {
        SelectedRecipe = recipe;
        SelectedRecipeName = recipe?.Name ?? string.Empty;
        SelectedRecipePrompt = recipe?.Prompt ?? string.Empty;
        SelectedRecipeMode = string.IsNullOrWhiteSpace(recipe?.CompanionMode) ? "automation" : recipe.CompanionMode;
        SelectedRecipeDescription = recipe?.Description ?? string.Empty;
        SelectedRecipeCategory = recipe?.Category ?? "general";
        SelectedRecipeSourceType = recipe?.SourceType ?? "manual";
        SelectedRecipeRisk = recipe?.RiskLevel ?? "low";
        SelectedRecipeGuardApp = recipe?.GuardApp ?? string.Empty;
        SelectedRecipeGuardForm = recipe?.GuardForm ?? string.Empty;
        SelectedRecipeGuardDialog = recipe?.GuardDialog ?? string.Empty;
        SelectedRecipeGuardTab = recipe?.GuardTab ?? string.Empty;
        SelectedRecipeTags = recipe == null ? string.Empty : string.Join(", ", recipe.Tags);
        SelectedRecipeKnowledgeSources = recipe == null ? string.Empty : string.Join(Environment.NewLine, recipe.KnowledgeSources);
        SelectedRecipeEnabled = recipe?.Enabled ?? true;
        SelectedRitualRunLog = recipe == null
            ? "no ritual run yet"
            : $"status: {(recipe.Enabled ? "enabled" : "disabled")} • risk: {recipe.RiskLevel} • confidence: {recipe.ConfidenceScore}%{Environment.NewLine}runs: {recipe.RunCount} • failures: {recipe.FailureCount} • saved minutes: {recipe.EstimatedMinutesSaved}{Environment.NewLine}last app/form: {recipe.LastActiveApp} / {recipe.LastActiveForm}{Environment.NewLine}last run: {recipe.LastRunUtc}{Environment.NewLine}last success: {recipe.LastSuccessUtc}";
        SelectedRitualRunStatus = recipe == null ? "idle" : "ready";
        ReplaceCollection(SelectedRecipeSteps, recipe?.Steps ?? []);
        ReplaceCollection(SelectedRecipeParameters, recipe?.Parameters ?? []);
        SetSelectedRecipeStep(SelectedRecipeSteps.FirstOrDefault());
        SetSelectedRecipeParameter(SelectedRecipeParameters.FirstOrDefault());
        OnPropertyChanged(nameof(SelectedRecipe));
        OnPropertyChanged(nameof(HasSelectedRecipe));
        OnPropertyChanged(nameof(HasSelectedRecipeSteps));
        OnPropertyChanged(nameof(HasSelectedRecipeParameters));
    }

    public void SetSelectedWatchSession(WatchSessionEntry? watchSession)
    {
        SelectedWatchSession = watchSession;
        SelectedWatchDetail = watchSession == null
            ? string.Empty
            : $"timestamp: {watchSession.TimestampUtc}{Environment.NewLine}provider: {watchSession.Provider}{Environment.NewLine}model: {watchSession.Model}{Environment.NewLine}app: {watchSession.ActiveApp}{Environment.NewLine}{Environment.NewLine}prompt:{Environment.NewLine}{watchSession.Prompt}{Environment.NewLine}{Environment.NewLine}assistant:{Environment.NewLine}{watchSession.AssistantResponse}{Environment.NewLine}{Environment.NewLine}screen summary:{Environment.NewLine}{watchSession.ScreenSummary}";
        OnPropertyChanged(nameof(SelectedWatchSession));
    }

    public void SetSelectedRecipeStep(RitualStep? step)
    {
        SelectedRecipeStep = step;
        RitualStepActionType = step?.ActionType ?? "app";
        RitualStepActionArgument = step?.ActionArgument ?? string.Empty;
        RitualStepRisk = step?.RiskLevel ?? "low";
        RitualStepWaitMs = step?.WaitMs ?? 0;
        RitualStepRetryCount = step?.RetryCount ?? 0;
        RitualStepIfApp = step?.IfApp ?? string.Empty;
        RitualStepIfForm = step?.IfForm ?? string.Empty;
        RitualStepIfDialog = step?.IfDialog ?? string.Empty;
        RitualStepIfTab = step?.IfTab ?? string.Empty;
        RitualStepOnFail = step?.OnFail ?? "stop";
        OnPropertyChanged(nameof(SelectedRecipeStep));
    }

    public void SetSelectedRecipeParameter(RitualParameter? parameter)
    {
        SelectedRecipeParameter = parameter;
        RitualParameterName = parameter?.Name ?? string.Empty;
        RitualParameterLabel = parameter?.Label ?? string.Empty;
        RitualParameterDefaultValue = parameter?.DefaultValue ?? string.Empty;
        RitualParameterRequired = parameter?.Required ?? true;
        RitualParameterKind = parameter?.Kind ?? "text";
        OnPropertyChanged(nameof(SelectedRecipeParameter));
    }

    public void SetSelectedHistoryEntry(ActionHistoryEntry? entry)
    {
        SelectedHistoryEntry = entry;
        SelectedHistoryDetail = entry == null
            ? string.Empty
            : $"timestamp: {entry.TimestampUtc}{Environment.NewLine}action: {entry.ActionName}{Environment.NewLine}argument: {entry.ActionArgument}{Environment.NewLine}target: {entry.TargetLabel}{Environment.NewLine}app: {entry.ActiveApp}{Environment.NewLine}category: {entry.Category}{Environment.NewLine}result:{Environment.NewLine}{entry.Result}{Environment.NewLine}{Environment.NewLine}spoken text:{Environment.NewLine}{entry.SpokenText}";
        OnPropertyChanged(nameof(SelectedHistoryEntry));
    }

    public void SetSelectedDiagnosticEntry(DiagnosticEntry? entry)
    {
        SelectedDiagnosticEntry = entry;
        SelectedDiagnosticDetail = entry == null ? string.Empty : $"{entry.SourceFile}{Environment.NewLine}{Environment.NewLine}{entry.Line}";
        OnPropertyChanged(nameof(SelectedDiagnosticEntry));
    }

    public void SaveSelectedRecipe()
    {
        if (string.IsNullOrWhiteSpace(SelectedRecipeName))
        {
            StatusMessage = "Recipe name is required.";
            return;
        }

        var recipes = Recipes.ToList();
        var existing = SelectedRecipe ?? recipes.FirstOrDefault(item => string.Equals(item.Name, SelectedRecipeName, StringComparison.OrdinalIgnoreCase));
        if (existing == null)
        {
            existing = new AutomationRecipe { CreatedAtUtc = DateTime.UtcNow.ToString("o") };
            recipes.Add(existing);
        }

        existing.Name = SelectedRecipeName.Trim();
        existing.Description = SelectedRecipeDescription.Trim();
        existing.Prompt = ExtractRecipePrompt(SelectedRecipePrompt);
        existing.CompanionMode = string.IsNullOrWhiteSpace(SelectedRecipeMode) ? "automation" : SelectedRecipeMode;
        existing.Category = string.IsNullOrWhiteSpace(SelectedRecipeCategory) ? InferRecipeCategory(existing.Prompt) : SelectedRecipeCategory.Trim().ToLowerInvariant();
        existing.SourceType = string.IsNullOrWhiteSpace(SelectedRecipeSourceType) ? "manual" : SelectedRecipeSourceType.Trim().ToLowerInvariant();
        existing.RiskLevel = string.IsNullOrWhiteSpace(SelectedRecipeRisk) ? InferRecipeRisk(existing) : SelectedRecipeRisk.Trim().ToLowerInvariant();
        existing.GuardApp = SelectedRecipeGuardApp.Trim();
        existing.GuardForm = SelectedRecipeGuardForm.Trim();
        existing.GuardDialog = SelectedRecipeGuardDialog.Trim();
        existing.GuardTab = SelectedRecipeGuardTab.Trim();
        existing.Enabled = SelectedRecipeEnabled;
        existing.Tags = SplitCommaList(SelectedRecipeTags);
        existing.KnowledgeSources = SplitLineList(SelectedRecipeKnowledgeSources);
        existing.Steps = SelectedRecipeSteps.Select(CloneStep).ToList();
        existing.Parameters = SelectedRecipeParameters.Select(CloneParameter).ToList();
        if (existing.Steps.Count == 0)
        {
            existing.Steps = BuildStepsFromPrompt(existing.Prompt);
        }

        if (existing.Category == "ax" && string.IsNullOrWhiteSpace(existing.GuardApp))
        {
            existing.GuardApp = "ax";
        }

        existing.ConfidenceScore = CalculateConfidenceScore(existing);

        _workspaceService.SaveRecipes(recipes.OrderBy(recipe => recipe.Name).ToList());
        StatusMessage = $"Saved ritual '{existing.Name}'.";
        Refresh();
        SetSelectedRecipe(Recipes.FirstOrDefault(item => item.Name == existing.Name) ?? Recipes.FirstOrDefault());
    }

    public void DeleteSelectedRecipe()
    {
        if (SelectedRecipe == null)
        {
            StatusMessage = "No ritual selected.";
            return;
        }

        var name = SelectedRecipe.Name;
        var recipes = Recipes.Where(item => !string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)).ToList();
        _workspaceService.SaveRecipes(recipes);
        StatusMessage = $"Deleted ritual '{name}'.";
        Refresh();
    }

    public void CloneSelectedRecipe()
    {
        if (SelectedRecipe == null)
        {
            StatusMessage = "No ritual selected.";
            return;
        }

        var clone = CloneRecipe(SelectedRecipe);
        clone.Id = Guid.NewGuid().ToString("n");
        clone.Name = $"{SelectedRecipe.Name} copy";
        clone.CreatedAtUtc = DateTime.UtcNow.ToString("o");
        clone.LastRunUtc = string.Empty;
        clone.LastSuccessUtc = string.Empty;
        clone.RunCount = 0;
        clone.FailureCount = 0;
        var recipes = _workspaceService.Load().Recipes.ToList();
        recipes.Add(clone);
        _workspaceService.SaveRecipes(recipes.OrderBy(item => item.Name).ToList());
        StatusMessage = $"Cloned ritual '{SelectedRecipe.Name}'.";
        Refresh();
        SetSelectedRecipe(Recipes.FirstOrDefault(item => item.Id == clone.Id));
    }

    public void ArchiveSelectedRecipe()
    {
        if (SelectedRecipe == null)
        {
            StatusMessage = "No ritual selected.";
            return;
        }

        var recipes = _workspaceService.Load().Recipes.ToList();
        var existing = recipes.FirstOrDefault(item => item.Id == SelectedRecipe.Id);
        if (existing == null)
        {
            StatusMessage = "Selected ritual no longer exists.";
            return;
        }

        existing.Archived = true;
        existing.Enabled = false;
        _workspaceService.SaveRecipes(recipes.OrderBy(item => item.Name).ToList());
        StatusMessage = $"Archived ritual '{existing.Name}'.";
        Refresh();
    }

    public void FilterKnowledge()
    {
        ReplaceCollection(KnowledgeDocuments, _workspaceService.GetKnowledgeDocuments(KnowledgeSearch));
        SetSelectedKnowledgeDocument(KnowledgeDocuments.FirstOrDefault());
    }

    public void FilterRituals()
    {
        ReplaceCollection(Recipes, FilterRecipes(_workspaceService.Load().Recipes));
        SetSelectedRecipe(Recipes.FirstOrDefault(item => SelectedRecipe != null && item.Id == SelectedRecipe.Id) ?? Recipes.FirstOrDefault());
    }

    public void FilterHistory()
    {
        ReplaceCollection(ActionHistory, FilterHistory(_workspaceService.Load().ActionHistory, HistorySearch));
        SetSelectedHistoryEntry(ActionHistory.FirstOrDefault());
    }

    public void FilterDiagnostics()
    {
        ReplaceCollection(Diagnostics, FilterDiagnostics(_workspaceService.Load().Diagnostics, DiagnosticsSearch));
        SetSelectedDiagnosticEntry(Diagnostics.FirstOrDefault());
    }

    public int ReindexKnowledge()
    {
        var chunks = _workspaceService.ReindexKnowledge();
        StatusMessage = $"Reindexed local knowledge. {chunks} chunks written.";
        Refresh();
        return chunks;
    }

    public bool DeleteSelectedKnowledgeDocument()
    {
        if (SelectedKnowledgeDocument == null)
        {
            StatusMessage = "No knowledge document selected.";
            return false;
        }

        var deleted = _workspaceService.DeleteKnowledgeDocument(SelectedKnowledgeDocument.RelativePath);
        StatusMessage = deleted
            ? $"Removed knowledge document '{SelectedKnowledgeDocument.Title}'."
            : "Could not remove selected knowledge document.";
        Refresh();
        return deleted;
    }

    public int ImportKnowledgeFiles(IEnumerable<string> sourcePaths)
    {
        var imported = _workspaceService.ImportKnowledgeFiles(sourcePaths);
        StatusMessage = imported > 0
            ? $"Imported {imported} knowledge file(s)."
            : "No supported knowledge files were imported.";
        Refresh();
        return imported;
    }

    public async Task DryRunSelectedRecipeAsync()
    {
        if (SelectedRecipe == null)
        {
            StatusMessage = "No ritual selected.";
            return;
        }

        var result = await _ritualRuntimeService.DryRunAsync(SelectedRecipe, _nextRitualStepIndex).ConfigureAwait(false);
        SelectedRitualRunStatus = result.Status;
        SelectedRitualRunLog = string.Join(Environment.NewLine, result.LogLines);
        StatusMessage = result.Summary;
    }

    public async Task RunSelectedRecipeAsync()
    {
        await RunSelectedRecipeInternalAsync(runSingleStep: false).ConfigureAwait(false);
    }

    public async Task RunNextSelectedRecipeStepAsync()
    {
        await RunSelectedRecipeInternalAsync(runSingleStep: true).ConfigureAwait(false);
    }

    public void AddOrUpdateRecipeStep()
    {
        if (SelectedRecipe == null)
        {
            StatusMessage = "No ritual selected.";
            return;
        }

        var step = SelectedRecipeStep ?? new RitualStep();
        step.ActionType = string.IsNullOrWhiteSpace(RitualStepActionType) ? "app" : RitualStepActionType.Trim().ToLowerInvariant();
        step.ActionArgument = RitualStepActionArgument.Trim();
        step.RiskLevel = string.IsNullOrWhiteSpace(RitualStepRisk) ? InferRiskLevel(step.ActionArgument) : RitualStepRisk.Trim().ToLowerInvariant();
        step.WaitMs = RitualStepWaitMs;
        step.RetryCount = RitualStepRetryCount;
        step.IfApp = RitualStepIfApp.Trim();
        step.IfForm = RitualStepIfForm.Trim();
        step.IfDialog = RitualStepIfDialog.Trim();
        step.IfTab = RitualStepIfTab.Trim();
        step.OnFail = string.IsNullOrWhiteSpace(RitualStepOnFail) ? "stop" : RitualStepOnFail.Trim().ToLowerInvariant();

        if (SelectedRecipeStep == null)
        {
            SelectedRecipeSteps.Add(step);
        }

        SetSelectedRecipeStep(step);
        StatusMessage = "Updated ritual step.";
        OnPropertyChanged(nameof(HasSelectedRecipeSteps));
    }

    public void RemoveSelectedRecipeStep()
    {
        if (SelectedRecipeStep == null)
        {
            StatusMessage = "No ritual step selected.";
            return;
        }

        var index = SelectedRecipeSteps.IndexOf(SelectedRecipeStep);
        SelectedRecipeSteps.Remove(SelectedRecipeStep);
        SetSelectedRecipeStep(index >= 0 && index < SelectedRecipeSteps.Count ? SelectedRecipeSteps[index] : SelectedRecipeSteps.FirstOrDefault());
        StatusMessage = "Removed ritual step.";
        OnPropertyChanged(nameof(HasSelectedRecipeSteps));
    }

    public void DuplicateSelectedRecipeStep()
    {
        if (SelectedRecipeStep == null)
        {
            StatusMessage = "No ritual step selected.";
            return;
        }

        var clone = CloneStep(SelectedRecipeStep);
        var index = SelectedRecipeSteps.IndexOf(SelectedRecipeStep);
        SelectedRecipeSteps.Insert(index + 1, clone);
        SetSelectedRecipeStep(clone);
        StatusMessage = "Duplicated ritual step.";
    }

    public void MoveSelectedRecipeStep(int direction)
    {
        if (SelectedRecipeStep == null)
        {
            StatusMessage = "No ritual step selected.";
            return;
        }

        var index = SelectedRecipeSteps.IndexOf(SelectedRecipeStep);
        var targetIndex = index + direction;
        if (index < 0 || targetIndex < 0 || targetIndex >= SelectedRecipeSteps.Count)
        {
            return;
        }

        SelectedRecipeSteps.Move(index, targetIndex);
        StatusMessage = "Moved ritual step.";
    }

    public void AddOrUpdateRecipeParameter()
    {
        if (SelectedRecipe == null)
        {
            StatusMessage = "No ritual selected.";
            return;
        }

        if (string.IsNullOrWhiteSpace(RitualParameterName))
        {
            StatusMessage = "Parameter name is required.";
            return;
        }

        var parameter = SelectedRecipeParameter ?? new RitualParameter();
        parameter.Name = RitualParameterName.Trim();
        parameter.Label = string.IsNullOrWhiteSpace(RitualParameterLabel) ? RitualParameterName.Trim() : RitualParameterLabel.Trim();
        parameter.DefaultValue = RitualParameterDefaultValue;
        parameter.Required = RitualParameterRequired;
        parameter.Kind = string.IsNullOrWhiteSpace(RitualParameterKind) ? "text" : RitualParameterKind.Trim().ToLowerInvariant();
        if (SelectedRecipeParameter == null)
        {
            SelectedRecipeParameters.Add(parameter);
        }

        SetSelectedRecipeParameter(parameter);
        StatusMessage = "Updated ritual parameter.";
        OnPropertyChanged(nameof(HasSelectedRecipeParameters));
    }

    public void RemoveSelectedRecipeParameter()
    {
        if (SelectedRecipeParameter == null)
        {
            StatusMessage = "No ritual parameter selected.";
            return;
        }

        var index = SelectedRecipeParameters.IndexOf(SelectedRecipeParameter);
        SelectedRecipeParameters.Remove(SelectedRecipeParameter);
        SetSelectedRecipeParameter(index >= 0 && index < SelectedRecipeParameters.Count ? SelectedRecipeParameters[index] : SelectedRecipeParameters.FirstOrDefault());
        StatusMessage = "Removed ritual parameter.";
        OnPropertyChanged(nameof(HasSelectedRecipeParameters));
    }

    public void PromoteHistoryToRecipe()
    {
        var entries = _workspaceService.Load().ActionHistory
            .Take(Math.Max(1, SelectedHistoryPromotionCount))
            .Reverse()
            .ToList();
        if (entries.Count == 0)
        {
            StatusMessage = "No history entries available for ritual promotion.";
            return;
        }

        var recipe = BuildRecipeFromHistory(entries, "history-derived");
        LoadRecipeIntoEditor(recipe);
        RitualSuggestionPreview = $"history -> ritual{Environment.NewLine}{BuildRecipeSummary(recipe)}";
        StatusMessage = $"Prepared ritual from {entries.Count} history step(s).";
    }

    public void PromoteWatchToRecipe()
    {
        if (SelectedWatchSession == null)
        {
            StatusMessage = "No watch session selected.";
            return;
        }

        var recipe = BuildRecipeFromWatch(SelectedWatchSession);
        LoadRecipeIntoEditor(recipe);
        RitualSuggestionPreview = $"watch -> ritual{Environment.NewLine}{BuildRecipeSummary(recipe)}";
        StatusMessage = "Prepared ritual from watch session.";
    }

    public void SuggestRecipeFromKnowledge()
    {
        if (SelectedKnowledgeDocument == null)
        {
            StatusMessage = "No knowledge document selected.";
            return;
        }

        var recipe = BuildRecipeFromKnowledge(SelectedKnowledgeDocument);
        LoadRecipeIntoEditor(recipe);
        RitualSuggestionPreview = $"knowledge -> ritual{Environment.NewLine}{BuildRecipeSummary(recipe)}";
        StatusMessage = $"Prepared knowledge-backed ritual from '{SelectedKnowledgeDocument.Title}'.";
    }

    public void StartTeachMode()
    {
        TeachModeActive = true;
        _teachModeBaselineHistoryCount = _workspaceService.Load().ActionHistory.Count;
        TeachModeStatus = "teach mode active: perform semantic AX/desktop actions now";
        StatusMessage = "Teach mode started.";
    }

    public void StopTeachMode()
    {
        if (!TeachModeActive)
        {
            StatusMessage = "Teach mode is not active.";
            return;
        }

        TeachModeActive = false;
        var entries = _workspaceService.Load().ActionHistory
            .Take(Math.Max(0, _workspaceService.Load().ActionHistory.Count - _teachModeBaselineHistoryCount))
            .Reverse()
            .ToList();
        TeachModeStatus = $"teach mode captured {entries.Count} action(s)";
        if (entries.Count == 0)
        {
            RitualSuggestionPreview = "teach mode ended without semantic actions";
            StatusMessage = "Teach mode stopped without captured actions.";
            return;
        }

        var recipe = BuildRecipeFromHistory(entries, "teach-mode");
        LoadRecipeIntoEditor(recipe);
        RitualSuggestionPreview = $"teach mode -> ritual{Environment.NewLine}{BuildRecipeSummary(recipe)}";
        StatusMessage = $"Teach mode captured {entries.Count} action(s) and built a ritual draft.";
    }

    public string ExportDiagnostics()
    {
        var path = _workspaceService.ExportDiagnostics(Diagnostics);
        StatusMessage = $"Exported diagnostics to {path}.";
        Refresh();
        return path;
    }

    public string ExportSupportBundle()
    {
        var path = _workspaceService.ExportSupportBundle(Diagnostics.ToList());
        StatusMessage = $"Exported support bundle to {path}.";
        _windowsShellService.UpdateTooltip(StatusMessage);
        Refresh();
        return path;
    }

    public void ClearDiagnostics()
    {
        _workspaceService.ClearDiagnosticsLogs();
        StatusMessage = "Cleared diagnostics log files.";
        Refresh();
    }

    public void RefreshActiveWindow()
    {
        var active = _workspaceService.GetActiveWindow();
        ActiveAppKind = string.IsNullOrWhiteSpace(active.AppKind) ? "generic" : active.AppKind.Trim();
        ActiveAppOneLiner = $"{active.DisplayName} — app kind: {ActiveAppKind}";
        LiveContextFatClientHint = BuildFatClientInspectorHint(ActiveAppKind);
        ActiveAppSummary = $"active window: {active.DisplayName}{Environment.NewLine}class: {active.WindowClassName}{Environment.NewLine}app kind: {active.AppKind}{Environment.NewLine}framework: {active.DesktopFramework}{Environment.NewLine}bounds: {active.Left},{active.Top} {active.Width}x{active.Height}";
        var axContext = _axClientAutomationService.CaptureActiveContext();
        AxContextSummary = axContext.Summary;
        var snapshot = _workspaceService.Load();
        AxSuggestedActions = BuildLiveContextActions(active, axContext);
        ProactiveSuggestionSummary = BuildProactiveSuggestion(active, axContext, snapshot.Recipes, snapshot.ActionHistory);
        _companionOverlayService.DockToActiveWindow(active);
        if (!string.IsNullOrWhiteSpace(ProactiveSuggestionSummary) &&
            !ProactiveSuggestionSummary.StartsWith("no proactive", StringComparison.OrdinalIgnoreCase))
        {
            _companionOverlayService.ShowTransient("ready", ProactiveSuggestionSummary, 3200, "low");
        }

        UpdateUspRail();
    }

    public void NavigateToLiveContext()
    {
        MainTabSelectedIndex = TabLiveContext;
        StatusMessage = "Switched to Live Context.";
    }

    public void NavigateToAsk()
    {
        MainTabSelectedIndex = TabAsk;
    }

    public void NavigateToKnowledge()
    {
        MainTabSelectedIndex = TabKnowledge;
    }

    public void NavigateToRituals(bool selectFirstRecipe = false)
    {
        MainTabSelectedIndex = TabRituals;
        if (selectFirstRecipe && Recipes.Count > 0)
        {
            SetSelectedRecipe(Recipes[0]);
        }
    }

    public void NavigateToSetup()
    {
        MainTabSelectedIndex = TabSetup;
    }

    public void NavigateToDiagnostics()
    {
        MainTabSelectedIndex = TabDiagnostics;
    }

    public void NavigateToHistory()
    {
        MainTabSelectedIndex = TabHistory;
    }

    public void OnboardingCheckEnvironment()
    {
        NavigateToSetup();
        StatusMessage = EnvironmentState.Contains("missing", StringComparison.OrdinalIgnoreCase)
            ? "Copy windows/.env.example to windows/.env and add keys, then refresh."
            : "Environment file present. Run provider smoke test from Ask tab.";
    }

    public void OnboardingOpenKnowledgeFolder()
    {
        NavigateToKnowledge();
        StatusMessage = "Import or drop docs into windows/data/knowledge/, then reindex.";
    }

    public void OnboardingTryPushToTalk()
    {
        NavigateToAsk();
        StatusMessage = "Use start push-to-talk / stop + ask on Ask tab, or the global hotkey from settings.";
    }

    public void OnboardingOpenAskSmoke()
    {
        NavigateToAsk();
        StatusMessage = "On Ask tab: click smoke test to verify your provider and keys.";
    }

    private void UpdateUspRail()
    {
        var snapshot = _workspaceService.Load();
        var active = _workspaceService.GetActiveWindow();
        var kind = string.IsNullOrWhiteSpace(active.AppKind) ? "generic" : active.AppKind.Trim().ToLowerInvariant();
        UspRailScreen = _includeScreens ? "screen • on" : "screen • off";
        UspRailRag = _useKnowledgeForAsk ? "RAG • on" : "RAG • off";
        UspRailRituals = $"{snapshot.Recipes.Count} rituals";
        UspRailHandoff = _autoRouteLocalAgents ? "handoff • auto" : "handoff • manual";
        UspRailVoice = snapshot.SpeakResponses
            ? (snapshot.SpeakAfterAsk ? "voice • +after ask" : "voice • on")
            : "voice • off";
        UspRailAx = kind == "ax" ? "fg • AX" : $"fg • {kind}";
    }

    private static string BuildFatClientInspectorHint(string appKind)
    {
        var k = (appKind ?? string.Empty).Trim().ToLowerInvariant();
        if (k == "ax")
        {
            return "Dynamics AX: use AX Inspector shortcuts and ax.* plan steps here; guards if_form / if_dialog / if_tab narrow execution.";
        }

        if (k is "creo" or "babtec" or "catia" or "nx")
        {
            return "Win32 fat client (non-AX): use Desktop Inspector and app| actions — do not use ax.* plan steps; use ritual ifapp=" + k + ".";
        }

        return string.Empty;
    }

    public void RunInspectorAction(string? actionArgument = null)
    {
        var action = string.IsNullOrWhiteSpace(actionArgument) ? InspectorAction : actionArgument;
        var active = _workspaceService.GetActiveWindow();
        var riskLevel = InferRiskLevel(action ?? string.Empty);
        _companionOverlayService.ShowActionPreview(riskLevel, $"inspector: {(action ?? string.Empty).Trim()}", active);
        var normalized = (action ?? string.Empty).Trim();
        InspectorResult = normalized.StartsWith("ax.", StringComparison.OrdinalIgnoreCase)
            ? _axClientAutomationService.Execute(normalized)
            : _desktopInspectorService.Execute(normalized);
        StatusMessage = $"Inspector action '{(action ?? string.Empty).Trim()}' completed.";
        _companionOverlayService.ShowTransient("ready", "desktop inspector updated", 3200, riskLevel);
        AppendActionHistory("inspector", normalized, InspectorResult, normalized, active.AppKind);
        RefreshActiveWindow();
    }

    public void SavePendingActionPlanAsRecipe()
    {
        if (_pendingActionPlan.Steps.Count == 0)
        {
            StatusMessage = "There is no pending action plan to save.";
            return;
        }

        var name = BuildSuggestedRecipeName();
        var axContext = _axClientAutomationService.CaptureActiveContext();
        var recipe = new AutomationRecipe
        {
            Id = Guid.NewGuid().ToString("n"),
            Name = name,
            Description = "Saved from assistant action plan",
            Prompt = string.Join("; ", _pendingActionPlan.Steps.Select(step => step.ActionArgument)),
            CompanionMode = "automation",
            SourceType = "saved-plan",
            Category = _pendingActionPlan.Steps.All(step => step.ActionArgument.StartsWith("ax.", StringComparison.OrdinalIgnoreCase)) ? "ax" : "general",
            RiskLevel = InferPlanRiskLevel(_pendingActionPlan),
            GuardApp = _pendingActionPlan.Steps.All(step => step.ActionArgument.StartsWith("ax.", StringComparison.OrdinalIgnoreCase)) ? "ax" : string.Empty,
            GuardForm = axContext.FormName,
            CreatedAtUtc = DateTime.UtcNow.ToString("o"),
            Steps = _pendingActionPlan.Steps.Select(step => new RitualStep
            {
                ActionType = step.ActionName,
                ActionArgument = step.ActionArgument,
                WaitMs = step.WaitMilliseconds,
                RetryCount = step.RetryCount,
                IfApp = step.RequiredAppContains,
                IfForm = step.RequiredFormContains,
                IfDialog = step.RequiredDialogContains,
                IfTab = step.RequiredTabContains,
                OnFail = step.OnFail,
                RiskLevel = InferRiskLevel(step.ActionArgument),
                TargetLabel = _pendingActionPlan.PointTag?.Label ?? string.Empty
            }).ToList(),
            KnowledgeSources = _workspaceService.GetRelevantKnowledgeChunks(AssistantPrompt).Select(chunk => chunk.Title).Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
            Tags = BuildTagsFromSteps(_pendingActionPlan.Steps.Select(step => step.ActionArgument))
        };

        var recipes = Recipes.Where(item => !string.Equals(item.Name, recipe.Name, StringComparison.OrdinalIgnoreCase)).ToList();
        recipes.Add(recipe);
        _workspaceService.SaveRecipes(recipes.OrderBy(item => item.Name).ToList());
        StatusMessage = $"Saved plan as ritual '{recipe.Name}'.";
        Refresh();
        SetSelectedRecipe(Recipes.FirstOrDefault(item => item.Name == recipe.Name));
    }

    public async Task RunSelectedAgentAsync()
    {
        if (string.IsNullOrWhiteSpace(ConsolePrompt))
        {
            StatusMessage = "Console prompt is empty.";
            return;
        }

        try
        {
            IsConsoleBusy = true;
            StatusMessage = $"Running {SelectedAgent}...";
            var result = await _runService.RunAsync(SelectedAgent, ConsolePrompt.Trim());
            ConsoleOutputPath = result.OutputFilePath;
            ConsoleOutput =
                $"agent: {result.Agent}{Environment.NewLine}" +
                $"working directory: {result.WorkingDirectory}{Environment.NewLine}" +
                $"output file: {result.OutputFilePath}{Environment.NewLine}" +
                $"exit code: {result.ExitCode}{Environment.NewLine}{Environment.NewLine}" +
                $"stdout:{Environment.NewLine}{(string.IsNullOrWhiteSpace(result.StandardOutput) ? "<empty>" : result.StandardOutput.Trim())}{Environment.NewLine}{Environment.NewLine}" +
                $"stderr:{Environment.NewLine}{(string.IsNullOrWhiteSpace(result.StandardError) ? "<empty>" : result.StandardError.Trim())}";
            StatusMessage = $"{SelectedAgent} finished successfully.";
        }
        catch (Exception exception)
        {
            ConsoleOutput = exception.ToString();
            StatusMessage = $"{SelectedAgent} failed.";
        }
        finally
        {
            IsConsoleBusy = false;
        }
    }

    public async Task RunAssistantAsync()
    {
        if (string.IsNullOrWhiteSpace(AssistantPrompt))
        {
            StatusMessage = "Assistant prompt is empty.";
            return;
        }

        var promptRaw = AssistantPrompt.Trim();

        try
        {
            IsAssistantBusy = true;
            _companionOverlayService.SetState("thinking", "routing request");

            if (AgentHandoffTriggers.IsOpenClawTriggered(promptRaw))
            {
                await CompleteLocalAgentHandoffAsync("openclaw", AgentHandoffTriggers.RemoveOpenClawPrompt(promptRaw)).ConfigureAwait(false);
                return;
            }

            if (AgentHandoffTriggers.IsClaudeCodeTriggered(promptRaw))
            {
                await CompleteLocalAgentHandoffAsync("claude-code", AgentHandoffTriggers.RemoveClaudeCodePrompt(promptRaw)).ConfigureAwait(false);
                return;
            }

            if (AgentHandoffTriggers.IsCodexTriggered(promptRaw))
            {
                await CompleteCodexHandoffAsync(promptRaw).ConfigureAwait(false);
                return;
            }

            if (AutoRouteLocalAgents)
            {
                var active = _workspaceService.GetActiveWindow();
                var route = AgentHandoffTriggers.DetectIntentRoute(promptRaw, active);
                if (string.Equals(route, "codex", StringComparison.OrdinalIgnoreCase))
                {
                    StatusMessage = "Auto-routing to Codex from context…";
                    _windowsShellService.UpdateTooltip(StatusMessage);
                    await CompleteCodexHandoffAsync("nimm codex " + promptRaw).ConfigureAwait(false);
                    return;
                }

                if (string.Equals(route, "openclaw", StringComparison.OrdinalIgnoreCase))
                {
                    StatusMessage = "Auto-routing to OpenClaw from context…";
                    _windowsShellService.UpdateTooltip(StatusMessage);
                    var routed = "nimm openclaw " + promptRaw;
                    await CompleteLocalAgentHandoffAsync("openclaw", AgentHandoffTriggers.RemoveOpenClawPrompt(routed)).ConfigureAwait(false);
                    return;
                }
            }

            StatusMessage = $"Asking {Provider}...";
            _companionOverlayService.SetState("thinking", $"asking {Provider}");
            var knowledgeChunks = UseKnowledgeForAsk && UseLocalKnowledge
                ? _workspaceService.GetRelevantKnowledgeChunks(promptRaw, 3)
                : [];

            var result = await _assistantRuntimeService.AskAsync(
                Provider,
                Model,
                Mode,
                SuggestAutomations,
                promptRaw,
                IncludeScreens,
                ConversationHistory.ToList(),
                knowledgeChunks).ConfigureAwait(false);

            AssistantResponse = string.IsNullOrWhiteSpace(result.CleanResponseText) ? result.ResponseText : result.CleanResponseText;
            _pendingActionPlan = result.ActionPlan ?? new AssistantActionPlan();
            ActionPlanPreview = BuildActionPlanPreview(result, _workspaceService.GetActiveWindow(), _axClientAutomationService.CaptureActiveContext());
            ActionPlanState = _pendingActionPlan.Steps.Count > 0 ? "pending" : "idle";
            ActionPlanExecutionLog = _pendingActionPlan.Steps.Count > 0 ? "plan ready for execution" : "no plan executed yet";
            _nextPlanStepIndex = 0;
            CurrentPlanStepSummary = BuildCurrentStepSummary();
            RetrievalSources = BuildRetrievalSources(result);
            ConversationHistory.Add(new ConversationTurn
            {
                UserTranscript = promptRaw,
                AssistantResponse = AssistantResponse
            });
            StatusMessage = $"{result.Provider} reply ready.";
            _windowsShellService.UpdateTooltip(StatusMessage);
            ApplyAssistantTargeting(result);
            _companionOverlayService.ShowTransient("ready", BuildOverlaySnippet(AssistantResponse));
            OnPropertyChanged(nameof(HasPendingActionPlan));
            OnPropertyChanged(nameof(HasRemainingPlanSteps));

            if (SpeakResponses && SpeakAfterAsk)
            {
                var spokenText = AssistantResponse;
                _ = SpeakAfterCloudAskAsync(spokenText);
            }
        }
        catch (Exception exception)
        {
            AssistantResponse = exception.ToString();
            RetrievalSources = string.Empty;
            ActionPlanPreview = "no pending action plan";
            ActionPlanExecutionLog = "no plan executed yet";
            ActionPlanState = "error";
            _pendingActionPlan = new AssistantActionPlan();
            _nextPlanStepIndex = 0;
            CurrentPlanStepSummary = "no active step";
            StatusMessage = "Assistant request failed.";
            _windowsShellService.UpdateTooltip(StatusMessage);
            _companionOverlayService.SetState("error", "assistant request failed");
            _companionOverlayService.ClearTargetAnchor();
            OnPropertyChanged(nameof(HasPendingActionPlan));
            OnPropertyChanged(nameof(HasRemainingPlanSteps));
        }
        finally
        {
            IsAssistantBusy = false;
        }
    }

    private async Task SpeakAfterCloudAskAsync(string text)
    {
        var compact = string.IsNullOrWhiteSpace(text) ? string.Empty : text.Trim();
        if (compact.Length > 12_000)
        {
            compact = compact[..12_000];
        }

        if (string.IsNullOrWhiteSpace(compact))
        {
            return;
        }

        try
        {
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                StatusMessage = "Auto-speaking reply…";
                _windowsShellService.UpdateTooltip(StatusMessage);
                _companionOverlayService.SetState("speaking", "auto voice reply");
            });

            if (_assistantRuntimeService.IsElevenLabsVoiceConfigured())
            {
                try
                {
                    var path = await _assistantRuntimeService.SynthesizeSpeechToFileAsync(compact).ConfigureAwait(false);
                    await Dispatcher.UIThread.InvokeAsync(() =>
                    {
                        LastGeneratedAudioPath = path;
                        _assistantRuntimeService.OpenFileWithShell(path);
                        StatusMessage = "Auto-speak: opened ElevenLabs audio.";
                        _windowsShellService.UpdateTooltip(StatusMessage);
                        _companionOverlayService.ShowTransient("speaking", "auto voice done");
                    });
                    return;
                }
                catch
                {
                    // SAPI below
                }
            }

            var ok = await Task.Run(() => WindowsSpeechFallback.TrySpeak(compact)).ConfigureAwait(false);
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                if (ok)
                {
                    LastGeneratedAudioPath = string.Empty;
                    StatusMessage = "Auto-speak: Windows voice.";
                }
                else
                {
                    StatusMessage = "Auto-speak skipped (no voice path worked).";
                }

                _windowsShellService.UpdateTooltip(StatusMessage);
                _companionOverlayService.ShowTransient("ready", StatusMessage);
            });
        }
        catch
        {
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                StatusMessage = "Auto-speak failed.";
                _windowsShellService.UpdateTooltip(StatusMessage);
                _companionOverlayService.SetState("ready", "auto-speak failed");
            });
        }
    }

    public void AppendHandoffPrefix(string prefix)
    {
        var insert = (prefix ?? string.Empty).Trim();
        if (insert.Length == 0)
        {
            return;
        }

        if (!insert.EndsWith(' '))
        {
            insert += " ";
        }

        var current = (AssistantPrompt ?? string.Empty).Trim();
        AssistantPrompt = string.IsNullOrEmpty(current) ? insert.TrimEnd() : insert + current;
    }

    private async Task CompleteCodexHandoffAsync(string promptForStrip)
    {
        var inner = AgentHandoffTriggers.RemoveCodexPrompt(promptForStrip);
        if (string.IsNullOrWhiteSpace(inner))
        {
            StatusMessage = "Say what Codex should do after the trigger.";
            AssistantResponse = string.Empty;
            RetrievalSources = "codex handoff: empty task";
            _windowsShellService.UpdateTooltip(StatusMessage);
            _companionOverlayService.ShowTransient("ready", StatusMessage);
            return;
        }

        IReadOnlyList<string>? images = null;
        if (AgentHandoffTriggers.ShouldAttachCodexScreens(promptForStrip))
        {
            StatusMessage = "Capturing screens for Codex…";
            _windowsShellService.UpdateTooltip(StatusMessage);
            images = _assistantRuntimeService.SaveScreenCapturesToCodexHandoffFolder();
        }

        await CompleteLocalAgentHandoffAsync("codex", inner, images).ConfigureAwait(false);
    }

    private async Task CompleteLocalAgentHandoffAsync(string agent, string strippedPrompt, IReadOnlyList<string>? codexImages = null)
    {
        if (string.IsNullOrWhiteSpace(strippedPrompt))
        {
            StatusMessage = $"Add a task for {agent}.";
            AssistantResponse = string.Empty;
            RetrievalSources = $"{agent} handoff: empty task";
            _windowsShellService.UpdateTooltip(StatusMessage);
            return;
        }

        StatusMessage = $"Running {agent}…";
        _companionOverlayService.SetState("thinking", agent);
        _windowsShellService.UpdateTooltip(StatusMessage);

        var result = await _runService.RunAsync(agent, strippedPrompt, codexImages).ConfigureAwait(false);

        AssistantResponse = result.ResponseText.Trim() + Environment.NewLine + Environment.NewLine + "---" + Environment.NewLine + "output file: " + result.OutputFilePath;
        RetrievalSources = $"local agent handoff • {agent} • exit {result.ExitCode}";
        _pendingActionPlan = new AssistantActionPlan();
        ActionPlanPreview = "no pending action plan (local agent handoff)";
        ActionPlanExecutionLog = "no plan executed yet";
        ActionPlanState = "idle";
        _nextPlanStepIndex = 0;
        CurrentPlanStepSummary = "no active step";
        ConversationHistory.Add(new ConversationTurn
        {
            UserTranscript = AssistantPrompt.Trim(),
            AssistantResponse = AssistantResponse
        });
        StatusMessage = $"{agent} finished.";
        _windowsShellService.UpdateTooltip(StatusMessage);
        _companionOverlayService.ShowTransient("ready", $"{agent} done", 4200, "low");
        _companionOverlayService.ClearTargetAnchor();
        OnPropertyChanged(nameof(HasPendingActionPlan));
        OnPropertyChanged(nameof(HasRemainingPlanSteps));
    }

    public async Task RunAssistantSmokeTestAsync()
    {
        try
        {
            IsAssistantBusy = true;
            StatusMessage = $"Running {Provider} smoke test...";
            _companionOverlayService.SetState("thinking", $"smoke testing {Provider}");
            var result = await _assistantRuntimeService.SmokeTestAsync(Provider, Model).ConfigureAwait(false);
            AssistantResponse = result;
            RetrievalSources = "smoke test: provider connectivity only";
            ActionPlanPreview = "no pending action plan";
            ActionPlanExecutionLog = "no plan executed yet";
            ActionPlanState = "idle";
            _pendingActionPlan = new AssistantActionPlan();
            _nextPlanStepIndex = 0;
            CurrentPlanStepSummary = "no active step";
            StatusMessage = $"{Provider} smoke test passed.";
            _windowsShellService.UpdateTooltip(StatusMessage);
            _companionOverlayService.ShowTransient("ready", "provider connectivity ready");
            _companionOverlayService.ClearTargetAnchor();
            OnPropertyChanged(nameof(HasPendingActionPlan));
            OnPropertyChanged(nameof(HasRemainingPlanSteps));
        }
        catch (Exception exception)
        {
            AssistantResponse = exception.ToString();
            RetrievalSources = string.Empty;
            ActionPlanPreview = "no pending action plan";
            ActionPlanExecutionLog = "no plan executed yet";
            ActionPlanState = "error";
            _pendingActionPlan = new AssistantActionPlan();
            _nextPlanStepIndex = 0;
            CurrentPlanStepSummary = "no active step";
            StatusMessage = $"{Provider} smoke test failed.";
            _windowsShellService.UpdateTooltip(StatusMessage);
            _companionOverlayService.SetState("error", "provider smoke test failed");
            _companionOverlayService.ClearTargetAnchor();
            OnPropertyChanged(nameof(HasPendingActionPlan));
            OnPropertyChanged(nameof(HasRemainingPlanSteps));
        }
        finally
        {
            IsAssistantBusy = false;
        }
    }

    public async Task ImportAndTranscribeAudioAsync(string audioFilePath)
    {
        if (string.IsNullOrWhiteSpace(audioFilePath))
        {
            return;
        }

        try
        {
            IsAssistantBusy = true;
            StatusMessage = "Transcribing audio...";
            _companionOverlayService.SetState("transcribing", "transcribing imported audio");
            TranscriptText = await _assistantRuntimeService.TranscribeAudioFileAsync(audioFilePath).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(AssistantPrompt))
            {
                AssistantPrompt = TranscriptText;
            }

            StatusMessage = "Audio transcript ready.";
            _windowsShellService.UpdateTooltip(StatusMessage);
            _companionOverlayService.ShowTransient("ready", "audio transcript ready");
        }
        catch (Exception exception)
        {
            TranscriptText = exception.ToString();
            StatusMessage = "Audio transcription failed.";
            _windowsShellService.UpdateTooltip(StatusMessage);
            _companionOverlayService.SetState("error", "audio transcription failed");
        }
        finally
        {
            IsAssistantBusy = false;
        }
    }

    public async Task SynthesizeCurrentResponseAsync()
    {
        if (string.IsNullOrWhiteSpace(AssistantResponse))
        {
            StatusMessage = "There is no assistant response to speak.";
            return;
        }

        try
        {
            IsAssistantBusy = true;
            StatusMessage = "Synthesizing speech...";
            _companionOverlayService.SetState("speaking", "rendering voice response");

            if (_assistantRuntimeService.IsElevenLabsVoiceConfigured())
            {
                try
                {
                    LastGeneratedAudioPath = await _assistantRuntimeService.SynthesizeSpeechToFileAsync(AssistantResponse).ConfigureAwait(false);
                    _assistantRuntimeService.OpenFileWithShell(LastGeneratedAudioPath);
                    StatusMessage = "Opened generated speech file.";
                    _windowsShellService.UpdateTooltip(StatusMessage);
                    _companionOverlayService.ShowTransient("speaking", "opened generated speech file");
                    return;
                }
                catch
                {
                    // try Windows SAPI below
                }
            }

            var spoken = await Task.Run(() => WindowsSpeechFallback.TrySpeak(AssistantResponse)).ConfigureAwait(false);
            if (spoken)
            {
                LastGeneratedAudioPath = string.Empty;
                StatusMessage = "Spoke response with Windows voice (offline fallback).";
                _windowsShellService.UpdateTooltip(StatusMessage);
                _companionOverlayService.ShowTransient("speaking", "windows voice playback");
                return;
            }

            throw new InvalidOperationException("Speech unavailable: add ELEVENLABS_API_KEY and ELEVENLABS_VOICE_ID to .env, or install a Windows SAPI voice.");
        }
        catch (Exception exception)
        {
            LastGeneratedAudioPath = string.Empty;
            StatusMessage = "Speech synthesis failed.";
            AssistantResponse = AssistantResponse + Environment.NewLine + Environment.NewLine + exception;
            _windowsShellService.UpdateTooltip(StatusMessage);
            _companionOverlayService.SetState("error", "speech synthesis failed");
        }
        finally
        {
            IsAssistantBusy = false;
        }
    }

    public void ClearConversation()
    {
        ConversationHistory.Clear();
        AssistantResponse = string.Empty;
        RetrievalSources = string.Empty;
        ActionPlanPreview = "no pending action plan";
        ActionPlanExecutionLog = "no plan executed yet";
        ActionPlanState = "idle";
        _pendingActionPlan = new AssistantActionPlan();
        _nextPlanStepIndex = 0;
        CurrentPlanStepSummary = "no active step";
        TranscriptText = string.Empty;
        StatusMessage = "Cleared assistant conversation.";
        _windowsShellService.UpdateTooltip(StatusMessage);
        _companionOverlayService.SetState("ready", "conversation cleared");
        _companionOverlayService.ClearTargetAnchor();
        OnPropertyChanged(nameof(HasPendingActionPlan));
        OnPropertyChanged(nameof(HasRemainingPlanSteps));
    }

    public async Task RunPendingActionPlanAsync()
    {
        if (_pendingActionPlan.Steps.Count == 0)
        {
            StatusMessage = "There is no pending action plan.";
            return;
        }

        var executed = await ExecutePlanStepsAsync(_nextPlanStepIndex, _pendingActionPlan.Steps.Count - _nextPlanStepIndex).ConfigureAwait(false);
        AppendExecutionLog(executed);
        ActionPlanState = executed.Any(line => line.Contains("failed", StringComparison.OrdinalIgnoreCase)) ? "partial" : "executed";
        StatusMessage = "Pending action plan executed.";
        _companionOverlayService.ShowTransient("ready", "action plan executed", 3200, InferPlanRiskLevel(_pendingActionPlan));
        OnPropertyChanged(nameof(HasRemainingPlanSteps));
    }

    public void ClearPendingActionPlan()
    {
        _pendingActionPlan = new AssistantActionPlan();
        ActionPlanPreview = "no pending action plan";
        ActionPlanExecutionLog = "plan cleared";
        ActionPlanState = "idle";
        _nextPlanStepIndex = 0;
        CurrentPlanStepSummary = "no active step";
        _companionOverlayService.ClearTargetAnchor();
        OnPropertyChanged(nameof(HasPendingActionPlan));
        OnPropertyChanged(nameof(HasRemainingPlanSteps));
    }

    public async Task RunNextActionPlanStepAsync()
    {
        if (_pendingActionPlan.Steps.Count == 0 || _nextPlanStepIndex >= _pendingActionPlan.Steps.Count)
        {
            StatusMessage = "There is no remaining plan step.";
            return;
        }

        var executed = await ExecutePlanStepsAsync(_nextPlanStepIndex, 1).ConfigureAwait(false);
        AppendExecutionLog(executed);
        ActionPlanState = _nextPlanStepIndex >= _pendingActionPlan.Steps.Count ? "executed" : "pending";
        StatusMessage = _nextPlanStepIndex >= _pendingActionPlan.Steps.Count ? "Last plan step executed." : "Plan step executed.";
        _companionOverlayService.ShowTransient("ready", "plan step executed", 2400, InferPlanRiskLevel(_pendingActionPlan));
        OnPropertyChanged(nameof(HasRemainingPlanSteps));
    }

    public void StartSpeechCapture()
    {
        if (IsRecordingSpeech)
        {
            return;
        }

        _microphoneRecorder.Start();
        IsRecordingSpeech = true;
        SpeechCaptureState = "recording...";
        StatusMessage = "Recording microphone input...";
        _windowsShellService.UpdateTooltip(StatusMessage);
        _companionOverlayService.SetState("listening", "push-to-talk recording");
    }

    public async Task StopSpeechCaptureAndAskAsync()
    {
        if (!IsRecordingSpeech)
        {
            return;
        }

        string? audioFilePath = null;
        try
        {
            SpeechCaptureState = "transcribing...";
            StatusMessage = "Transcribing microphone input...";
            _windowsShellService.UpdateTooltip(StatusMessage);
            _companionOverlayService.SetState("transcribing", "transcribing microphone input");
            audioFilePath = _microphoneRecorder.Stop();
            IsRecordingSpeech = false;
            TranscriptText = await _assistantRuntimeService.TranscribeAudioFileAsync(audioFilePath).ConfigureAwait(false);
            AssistantPrompt = TranscriptText;
            SpeechCaptureState = "asking...";
            await RunAssistantAsync().ConfigureAwait(false);
            SpeechCaptureState = "idle";
        }
        catch (Exception exception)
        {
            TranscriptText = exception.ToString();
            SpeechCaptureState = "error";
            StatusMessage = "Push-to-talk failed.";
            IsRecordingSpeech = false;
            _windowsShellService.UpdateTooltip(StatusMessage);
            _companionOverlayService.SetState("error", "push-to-talk failed");
        }
        finally
        {
            TryDeleteFile(audioFilePath);
        }
    }

    public void CancelSpeechCapture()
    {
        _microphoneRecorder.Cancel();
        IsRecordingSpeech = false;
        SpeechCaptureState = "idle";
        StatusMessage = "Cancelled microphone recording.";
        _windowsShellService.UpdateTooltip(StatusMessage);
        _companionOverlayService.SetState("ready", "recording cancelled");
    }

    public void AttachShell()
    {
        _windowsShellService.HotKeyPressed += (_, _) => StartSpeechCapture();
        _windowsShellService.HotKeyReleased += async (_, _) => await StopSpeechCaptureAndAskAsync();
    }

    private static IReadOnlyList<ActionHistoryEntry> FilterHistory(IReadOnlyList<ActionHistoryEntry> entries, string search)
    {
        if (string.IsNullOrWhiteSpace(search))
        {
            return entries;
        }

        return entries.Where(entry =>
            $"{entry.TimestampUtc} {entry.ActionName} {entry.ActionArgument} {entry.TargetLabel} {entry.ActiveApp} {entry.SpokenText}"
                .Contains(search, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    private static IReadOnlyList<DiagnosticEntry> FilterDiagnostics(IReadOnlyList<DiagnosticEntry> entries, string search)
    {
        if (string.IsNullOrWhiteSpace(search))
        {
            return entries;
        }

        return entries.Where(entry =>
            $"{entry.SourceFile} {entry.Line}".Contains(search, StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    private static void ReplaceCollection<T>(ObservableCollection<T> target, IReadOnlyList<T> source)
    {
        target.Clear();
        foreach (var item in source)
        {
            target.Add(item);
        }
    }

    private async Task RunSelectedRecipeInternalAsync(bool runSingleStep)
    {
        if (SelectedRecipe == null)
        {
            StatusMessage = "No ritual selected.";
            return;
        }

        var recipe = BuildRecipeFromEditor();
        if (!runSingleStep && string.Equals(recipe.RiskLevel, "high", StringComparison.OrdinalIgnoreCase))
        {
            SelectedRitualRunStatus = "blocked";
            SelectedRitualRunLog = "high-risk rituals only run step-by-step. Use 'run next step'.";
            StatusMessage = "High-risk ritual blocked for full run.";
            _companionOverlayService.SetState("error", "high-risk ritual requires step mode", "high");
            return;
        }

        var result = runSingleStep
            ? await _ritualRuntimeService.RunAsync(recipe, startIndex: _nextRitualStepIndex, count: 1).ConfigureAwait(false)
            : await _ritualRuntimeService.RunAsync(recipe, startIndex: 0).ConfigureAwait(false);
        _nextRitualStepIndex = result.NextStepIndex;
        SelectedRitualRunStatus = result.Status;
        SelectedRitualRunLog = string.Join(Environment.NewLine, result.LogLines);
        StatusMessage = result.Summary;
        _companionOverlayService.ShowTransient(
            result.Status is "error" or "blocked" ? "error" : "ready",
            runSingleStep ? "ritual next step completed" : "ritual run completed",
            3200,
            recipe.RiskLevel);
        AppendActionHistory(runSingleStep ? "ritual-step" : "ritual-run", recipe.Name, result.Summary, recipe.Name, recipe.GuardApp);
        PersistRitualRunResult(recipe, result);
        Refresh();
        SetSelectedRecipe(Recipes.FirstOrDefault(item => item.Id == recipe.Id || item.Name == recipe.Name));
    }

    private AutomationRecipe BuildRecipeFromEditor()
    {
        var current = SelectedRecipe ?? new AutomationRecipe();
        current.Name = SelectedRecipeName.Trim();
        current.Description = SelectedRecipeDescription.Trim();
        current.Prompt = ExtractRecipePrompt(SelectedRecipePrompt);
        current.CompanionMode = SelectedRecipeMode;
        current.Category = SelectedRecipeCategory;
        current.SourceType = SelectedRecipeSourceType;
        current.RiskLevel = SelectedRecipeRisk;
        current.GuardApp = SelectedRecipeGuardApp.Trim();
        current.GuardForm = SelectedRecipeGuardForm.Trim();
        current.GuardDialog = SelectedRecipeGuardDialog.Trim();
        current.GuardTab = SelectedRecipeGuardTab.Trim();
        current.Enabled = SelectedRecipeEnabled;
        current.Tags = SplitCommaList(SelectedRecipeTags);
        current.KnowledgeSources = SplitLineList(SelectedRecipeKnowledgeSources);
        current.Steps = SelectedRecipeSteps.Select(CloneStep).ToList();
        current.Parameters = SelectedRecipeParameters.Select(CloneParameter).ToList();
        if (current.Steps.Count == 0)
        {
            current.Steps = BuildStepsFromPrompt(current.Prompt);
        }

        current.RiskLevel = InferRecipeRisk(current);
        current.ConfidenceScore = CalculateConfidenceScore(current);
        return current;
    }

    private void PersistRitualRunResult(AutomationRecipe recipe, RitualRunResult result)
    {
        var recipes = _workspaceService.Load().Recipes.ToList();
        var existing = recipes.FirstOrDefault(item => item.Id == recipe.Id || item.Name == recipe.Name);
        if (existing == null)
        {
            return;
        }

        existing.LastRunUtc = DateTime.UtcNow.ToString("o");
        existing.RunCount++;
        existing.RiskLevel = recipe.RiskLevel;
        existing.Steps = recipe.Steps;
        existing.Parameters = recipe.Parameters;
        existing.KnowledgeSources = recipe.KnowledgeSources;
        existing.Tags = recipe.Tags;
        var active = _workspaceService.GetActiveWindow();
        var axContext = _axClientAutomationService.CaptureActiveContext();
        existing.LastActiveApp = active.AppKind;
        existing.LastActiveForm = axContext.FormName;
        existing.ConfidenceScore = CalculateConfidenceScore(existing);
        if (result.Status is "completed" or "partial")
        {
            existing.LastSuccessUtc = DateTime.UtcNow.ToString("o");
            existing.EstimatedMinutesSaved += EstimateSavedMinutes(existing);
        }

        if (result.Status is "error" or "blocked")
        {
            existing.FailureCount++;
        }

        _workspaceService.SaveRecipes(recipes.OrderBy(item => item.Name).ToList());
    }

    private void LoadRecipeIntoEditor(AutomationRecipe recipe)
    {
        SetSelectedRecipe(recipe);
        SelectedRecipeName = recipe.Name;
        SelectedRecipeDescription = recipe.Description;
        SelectedRecipePrompt = recipe.Prompt;
        SelectedRecipeMode = recipe.CompanionMode;
        SelectedRecipeCategory = recipe.Category;
        SelectedRecipeSourceType = recipe.SourceType;
        SelectedRecipeRisk = recipe.RiskLevel;
        SelectedRecipeGuardApp = recipe.GuardApp;
        SelectedRecipeGuardForm = recipe.GuardForm;
        SelectedRecipeGuardDialog = recipe.GuardDialog;
        SelectedRecipeGuardTab = recipe.GuardTab;
        SelectedRecipeEnabled = recipe.Enabled;
        SelectedRecipeTags = string.Join(", ", recipe.Tags);
        SelectedRecipeKnowledgeSources = string.Join(Environment.NewLine, recipe.KnowledgeSources);
        ReplaceCollection(SelectedRecipeSteps, recipe.Steps);
        ReplaceCollection(SelectedRecipeParameters, recipe.Parameters);
        SetSelectedRecipeStep(SelectedRecipeSteps.FirstOrDefault());
        SetSelectedRecipeParameter(SelectedRecipeParameters.FirstOrDefault());
    }

    private AutomationRecipe BuildRecipeFromHistory(IReadOnlyList<ActionHistoryEntry> entries, string sourceType)
    {
        var active = _workspaceService.GetActiveWindow();
        var axContext = _axClientAutomationService.CaptureActiveContext();
        var recipe = new AutomationRecipe
        {
            Id = Guid.NewGuid().ToString("n"),
            Name = $"{(entries.All(entry => entry.Category == "ax") ? "AX" : "Desktop")} Ritual {DateTime.Now:HHmmss}",
            Description = $"Derived from {entries.Count} action history entries",
            SourceType = sourceType,
            CompanionMode = "automation",
            Category = entries.All(entry => entry.Category == "ax") ? "ax" : "general",
            GuardApp = entries.All(entry => entry.Category == "ax") ? "ax" : active.AppKind,
            GuardForm = axContext.FormName,
            CreatedAtUtc = DateTime.UtcNow.ToString("o"),
            Enabled = true,
            Steps = NormalizeHistoryToSteps(entries),
            Tags = BuildTagsFromSteps(entries.Select(entry => entry.ActionArgument)),
            Prompt = string.Join("; ", entries.Select(entry => entry.ActionArgument))
        };
        recipe.RiskLevel = InferRecipeRisk(recipe);
        recipe.Parameters = InferParameters(recipe.Steps);
        recipe.LastActiveApp = active.AppKind;
        recipe.LastActiveForm = axContext.FormName;
        recipe.ConfidenceScore = CalculateConfidenceScore(recipe);
        return recipe;
    }

    private AutomationRecipe BuildRecipeFromWatch(WatchSessionEntry watchSession)
    {
        var inferredPlan = AssistantActionPlan.Parse(watchSession.AssistantResponse);
        var steps = inferredPlan.Steps.Count > 0
            ? inferredPlan.Steps.Select(step => new RitualStep
            {
                ActionType = step.ActionName,
                ActionArgument = step.ActionArgument,
                WaitMs = step.WaitMilliseconds,
                RetryCount = step.RetryCount,
                IfApp = step.RequiredAppContains,
                IfForm = step.RequiredFormContains,
                IfDialog = step.RequiredDialogContains,
                IfTab = step.RequiredTabContains,
                OnFail = step.OnFail,
                RiskLevel = InferRiskLevel(step.ActionArgument),
                TargetLabel = inferredPlan.PointTag?.Label ?? string.Empty
            }).ToList()
            : BuildStepsFromPrompt(watchSession.AssistantResponse);
        var recipe = new AutomationRecipe
        {
            Id = Guid.NewGuid().ToString("n"),
            Name = $"Watch Ritual {DateTime.Now:HHmmss}",
            Description = "Derived from selected watch session",
            Prompt = watchSession.Prompt,
            CompanionMode = "watch",
            SourceType = "watch-derived",
            Category = steps.All(step => step.ActionArgument.StartsWith("ax.", StringComparison.OrdinalIgnoreCase)) ? "ax" : "general",
            GuardApp = watchSession.ActiveApp.Contains("ax", StringComparison.OrdinalIgnoreCase) ? "ax" : watchSession.ActiveApp,
            CreatedAtUtc = DateTime.UtcNow.ToString("o"),
            Steps = steps,
            Tags = BuildTagsFromSteps(steps.Select(step => step.ActionArgument))
        };
        recipe.RiskLevel = InferRecipeRisk(recipe);
        recipe.Parameters = InferParameters(recipe.Steps);
        recipe.ConfidenceScore = CalculateConfidenceScore(recipe);
        return recipe;
    }

    private AutomationRecipe BuildRecipeFromKnowledge(KnowledgeDocumentSummary document)
    {
        var preview = _workspaceService.GetKnowledgePreview(document.RelativePath, 2200);
        var steps = BuildStepsFromKnowledgePreview(preview);
        var recipe = new AutomationRecipe
        {
            Id = Guid.NewGuid().ToString("n"),
            Name = $"Knowledge Ritual {Path.GetFileNameWithoutExtension(document.Title)}",
            Description = "Derived from selected SOP / knowledge document",
            Prompt = preview,
            CompanionMode = "automation",
            SourceType = "knowledge-derived",
            Category = steps.All(step => step.ActionArgument.StartsWith("ax.", StringComparison.OrdinalIgnoreCase)) && steps.Count > 0 ? "ax" : "general",
            GuardApp = steps.Any(step => step.ActionArgument.StartsWith("ax.", StringComparison.OrdinalIgnoreCase)) ? "ax" : string.Empty,
            CreatedAtUtc = DateTime.UtcNow.ToString("o"),
            KnowledgeSources = [document.Title],
            Steps = steps,
            Tags = BuildTagsFromSteps(steps.Select(step => step.ActionArgument))
        };
        recipe.RiskLevel = InferRecipeRisk(recipe);
        recipe.Parameters = InferParameters(recipe.Steps);
        recipe.ConfidenceScore = CalculateConfidenceScore(recipe);
        return recipe;
    }

    private static string BuildRecipeSummary(AutomationRecipe recipe)
    {
        return $"name: {recipe.Name}{Environment.NewLine}source: {recipe.SourceType}{Environment.NewLine}risk: {recipe.RiskLevel}{Environment.NewLine}confidence: {recipe.ConfidenceScore}%{Environment.NewLine}steps: {recipe.Steps.Count}{Environment.NewLine}guards: {recipe.GuardApp} / {recipe.GuardForm}";
    }

    private static List<RitualStep> NormalizeHistoryToSteps(IReadOnlyList<ActionHistoryEntry> entries)
    {
        var steps = new List<RitualStep>();
        foreach (var group in entries.Where(entry => !string.IsNullOrWhiteSpace(entry.ActionArgument))
                     .GroupBy(entry => entry.ActionArgument.Trim(), StringComparer.OrdinalIgnoreCase))
        {
            var first = group.First();
            steps.Add(new RitualStep
            {
                ActionType = first.ActionArgument.StartsWith("ax.", StringComparison.OrdinalIgnoreCase) ? "ax" : "app",
                ActionArgument = first.ActionArgument,
                RiskLevel = InferRiskLevel(first.ActionArgument),
                TargetLabel = first.TargetLabel
            });
        }

        return steps;
    }

    private static List<RitualStep> BuildStepsFromPrompt(string prompt)
    {
        if (string.IsNullOrWhiteSpace(prompt))
        {
            return [];
        }

        var plan = AssistantActionPlan.Parse("[ACTIONS:" + prompt.Replace(Environment.NewLine, ";") + "]");
        if (plan.Steps.Count > 0)
        {
            return plan.Steps.Select(step => new RitualStep
            {
                ActionType = step.ActionName,
                ActionArgument = step.ActionArgument,
                WaitMs = step.WaitMilliseconds,
                RetryCount = step.RetryCount,
                IfApp = step.RequiredAppContains,
                IfForm = step.RequiredFormContains,
                IfDialog = step.RequiredDialogContains,
                IfTab = step.RequiredTabContains,
                OnFail = step.OnFail,
                RiskLevel = InferRiskLevel(step.ActionArgument)
            }).ToList();
        }

        return prompt.Split(['\n', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(entry => new RitualStep
            {
                ActionType = entry.StartsWith("ax.", StringComparison.OrdinalIgnoreCase) ? "ax" : "app",
                ActionArgument = entry,
                RiskLevel = InferRiskLevel(entry)
            })
            .ToList();
    }

    private static List<RitualStep> BuildStepsFromKnowledgePreview(string preview)
    {
        var lines = preview.Replace("\r\n", "\n").Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var steps = lines
            .Where(line => line.StartsWith("ax.", StringComparison.OrdinalIgnoreCase) || line.StartsWith("app|", StringComparison.OrdinalIgnoreCase))
            .SelectMany(line => BuildStepsFromPrompt(line))
            .ToList();
        if (steps.Count > 0)
        {
            return steps;
        }

        return preview.Contains("AX", StringComparison.OrdinalIgnoreCase) || preview.Contains("Dynamics", StringComparison.OrdinalIgnoreCase)
            ? [new RitualStep
                {
                    ActionType = "ax",
                    ActionArgument = "ax.read_context",
                    RiskLevel = "low",
                    TargetLabel = "review SOP context before acting"
                }]
            : [new RitualStep
                {
                    ActionType = "app",
                    ActionArgument = "read_form",
                    RiskLevel = "low",
                    TargetLabel = "review document-backed context"
                }];
    }

    private static List<RitualParameter> InferParameters(IEnumerable<RitualStep> steps)
    {
        var matches = steps
            .SelectMany(step => System.Text.RegularExpressions.Regex.Matches(step.ActionArgument ?? string.Empty, @"\{\{([a-zA-Z0-9_\-]+)\}\}")
                .Select(match => match.Groups[1].Value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(name => new RitualParameter
            {
                Name = name,
                Label = name.Replace("_", " "),
                Required = true,
                Kind = "text"
            })
            .ToList();
        return matches;
    }

    private static RitualStep CloneStep(RitualStep step)
    {
        return new RitualStep
        {
            ActionType = step.ActionType,
            ActionArgument = step.ActionArgument,
            WaitMs = step.WaitMs,
            RetryCount = step.RetryCount,
            IfApp = step.IfApp,
            IfForm = step.IfForm,
            IfDialog = step.IfDialog,
            IfTab = step.IfTab,
            OnFail = step.OnFail,
            RiskLevel = step.RiskLevel,
            TargetLabel = step.TargetLabel,
            ParameterBindings = new Dictionary<string, string>(step.ParameterBindings, StringComparer.OrdinalIgnoreCase)
        };
    }

    private static RitualParameter CloneParameter(RitualParameter parameter)
    {
        return new RitualParameter
        {
            Name = parameter.Name,
            Label = parameter.Label,
            DefaultValue = parameter.DefaultValue,
            Required = parameter.Required,
            Kind = parameter.Kind
        };
    }

    private static AutomationRecipe CloneRecipe(AutomationRecipe recipe)
    {
        return new AutomationRecipe
        {
            Id = recipe.Id,
            Name = recipe.Name,
            Description = recipe.Description,
            Prompt = recipe.Prompt,
            CompanionMode = recipe.CompanionMode,
            Category = recipe.Category,
            GuardApp = recipe.GuardApp,
            GuardForm = recipe.GuardForm,
            GuardDialog = recipe.GuardDialog,
            GuardTab = recipe.GuardTab,
            SourceType = recipe.SourceType,
            RiskLevel = recipe.RiskLevel,
            Enabled = recipe.Enabled,
            Archived = recipe.Archived,
            CreatedAtUtc = recipe.CreatedAtUtc,
            LastRunUtc = recipe.LastRunUtc,
            LastSuccessUtc = recipe.LastSuccessUtc,
            RunCount = recipe.RunCount,
            FailureCount = recipe.FailureCount,
            LastActiveApp = recipe.LastActiveApp,
            LastActiveForm = recipe.LastActiveForm,
            ConfidenceScore = recipe.ConfidenceScore,
            EstimatedMinutesSaved = recipe.EstimatedMinutesSaved,
            KnowledgeSources = recipe.KnowledgeSources.ToList(),
            Tags = recipe.Tags.ToList(),
            Parameters = recipe.Parameters.Select(CloneParameter).ToList(),
            Steps = recipe.Steps.Select(CloneStep).ToList()
        };
    }

    private IReadOnlyList<AutomationRecipe> FilterRecipes(IReadOnlyList<AutomationRecipe> recipes)
    {
        IEnumerable<AutomationRecipe> filtered = recipes.Where(recipe => !recipe.Archived);
        if (!string.IsNullOrWhiteSpace(RitualSearch))
        {
            filtered = filtered.Where(recipe =>
                $"{recipe.Name} {recipe.Description} {recipe.Category} {recipe.SourceType} {recipe.GuardApp} {recipe.GuardForm} {string.Join(' ', recipe.Tags)} {string.Join(' ', recipe.Steps.Select(step => step.ActionArgument))}"
                    .Contains(RitualSearch, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.Equals(RitualCategoryFilter, "all", StringComparison.OrdinalIgnoreCase))
        {
            filtered = filtered.Where(recipe => recipe.Category.Equals(RitualCategoryFilter, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.Equals(RitualSourceFilter, "all", StringComparison.OrdinalIgnoreCase))
        {
            filtered = filtered.Where(recipe => recipe.SourceType.Equals(RitualSourceFilter, StringComparison.OrdinalIgnoreCase));
        }

        if (!string.Equals(RitualRiskFilter, "all", StringComparison.OrdinalIgnoreCase))
        {
            filtered = filtered.Where(recipe => recipe.RiskLevel.Equals(RitualRiskFilter, StringComparison.OrdinalIgnoreCase));
        }

        return filtered.OrderByDescending(recipe => recipe.LastSuccessUtc)
            .ThenByDescending(recipe => recipe.LastRunUtc)
            .ThenBy(recipe => recipe.Name)
            .ToList();
    }

    private static string BuildRitualStats(IReadOnlyList<AutomationRecipe> recipes)
    {
        var total = recipes.Count;
        var active = recipes.Count(recipe => recipe.Enabled && !recipe.Archived);
        var recentRuns = recipes.Sum(recipe => recipe.RunCount);
        var failures = recipes.Sum(recipe => recipe.FailureCount);
        var successRate = recentRuns == 0 ? 100 : Math.Max(0, (int)Math.Round(((double)(recentRuns - failures) / recentRuns) * 100));
        return $"rituals: {total} total • {active} active • runs: {recentRuns} • success rate: {successRate}%";
    }

    private static string BuildOperatorMetrics(IReadOnlyList<AutomationRecipe> recipes, IReadOnlyList<ActionHistoryEntry> history, IReadOnlyList<KnowledgeDocumentSummary> docs)
    {
        var successfulRuns = recipes.Sum(recipe => Math.Max(0, recipe.RunCount - recipe.FailureCount));
        var minutesSaved = recipes.Sum(recipe => recipe.EstimatedMinutesSaved);
        var topApp = recipes.Where(recipe => !string.IsNullOrWhiteSpace(recipe.LastActiveApp))
            .GroupBy(recipe => recipe.LastActiveApp)
            .OrderByDescending(group => group.Count())
            .Select(group => group.Key)
            .FirstOrDefault() ?? "n/a";
        return $"saved minutes: {minutesSaved} • successful ritual runs: {successfulRuns} • history entries: {history.Count} • SOP docs: {docs.Count} • top app: {topApp}";
    }

    private static int CalculateConfidenceScore(AutomationRecipe recipe)
    {
        var score = 40;
        if (!string.IsNullOrWhiteSpace(recipe.GuardApp))
        {
            score += 10;
        }

        if (!string.IsNullOrWhiteSpace(recipe.GuardForm))
        {
            score += 10;
        }

        if (recipe.Parameters.Count > 0)
        {
            score += 10;
        }

        if (recipe.KnowledgeSources.Count > 0)
        {
            score += 10;
        }

        score += Math.Min(20, Math.Max(0, recipe.RunCount - recipe.FailureCount) * 3);
        score -= Math.Min(25, recipe.FailureCount * 5);
        return Math.Clamp(score, 5, 100);
    }

    private static int EstimateSavedMinutes(AutomationRecipe recipe)
    {
        return Math.Max(1, Math.Min(6, recipe.Steps.Count));
    }

    private static string BuildLiveContextActions(ActiveWindowInfo active, AxContextSnapshot axContext)
    {
        if (axContext.IsAxClient)
        {
            return "AX: ax.focus_window, ax.read_field:<label>, ax.set_field:<label>=<value>, ax.open_lookup:<label>, ax.read_grid, ax.select_grid_row:<query>, ax.confirm_dialog";
        }

        return active.AppKind switch
        {
            "browser" => "Browser: focus_window, read_form, list_controls, web rituals, knowledge-backed SOP runs",
            "explorer" => "Explorer: focus_window, list_controls, read_form, path rituals, folder workflows",
            "mail" => "Mail: focus_window, read_form, guarded review rituals, send/post denylist active",
            "ide" => "IDE: focus_window, list_controls, codex/claude/openclaw handoff rituals",
            "creo" => "Creo/PTC: Win32 fat client — list_controls, read_form, read_dialog, app| semantic actions; use ritual guards ifapp=creo (ax.* only for Dynamics AX)",
            "babtec" => "Babtec: Win32 fat client — list_controls, read_form, app| actions; guards ifapp=babtec",
            "catia" => "CATIA: Win32 fat client — list_controls, read_form, app| actions; guards ifapp=catia",
            "nx" => "Siemens NX: Win32 fat client — list_controls, read_form, app| actions; guards ifapp=nx",
            _ => "Desktop fat client: list_controls, read_form, ritual replay, knowledge-backed plans; set ritual Guard app to creo, babtec, catia, nx, or ax as needed"
        };
    }

    private string BuildProactiveSuggestion(ActiveWindowInfo active, AxContextSnapshot axContext, IReadOnlyList<AutomationRecipe> recipes, IReadOnlyList<ActionHistoryEntry> history)
    {
        var candidates = recipes
            .Where(recipe => recipe.Enabled && !recipe.Archived)
            .Where(recipe => string.IsNullOrWhiteSpace(recipe.GuardApp) || active.AppKind.Contains(recipe.GuardApp, StringComparison.OrdinalIgnoreCase))
            .Where(recipe => string.IsNullOrWhiteSpace(recipe.GuardForm) || axContext.FormName.Contains(recipe.GuardForm, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(recipe => recipe.ConfidenceScore)
            .ThenByDescending(recipe => recipe.LastSuccessUtc)
            .Take(2)
            .ToList();

        if (candidates.Count > 0)
        {
            var top = candidates[0];
            return $"Suggested ritual: {top.Name} • risk {top.RiskLevel} • confidence {top.ConfidenceScore}%";
        }

        var recentAx = history.Take(5).Count(entry => entry.Category == "ax");
        if (axContext.IsAxClient && recentAx >= 3)
        {
            return "Du machst diesen AX-Flow oft. Ritual erstellen?";
        }

        if (axContext.IsAxClient)
        {
            return $"Für Formular '{(string.IsNullOrWhiteSpace(axContext.FormName) ? "AX" : axContext.FormName)}' gibt es Potenzial für ein Ritual.";
        }

        return "no proactive ritual suggestion";
    }

    private static List<string> BuildTagsFromSteps(IEnumerable<string> actions)
    {
        var tags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var action in actions)
        {
            if (action.StartsWith("ax.", StringComparison.OrdinalIgnoreCase))
            {
                tags.Add("ax");
            }

            if (action.Contains("lookup", StringComparison.OrdinalIgnoreCase))
            {
                tags.Add("lookup");
            }

            if (action.Contains("grid", StringComparison.OrdinalIgnoreCase))
            {
                tags.Add("grid");
            }

            if (action.Contains("post", StringComparison.OrdinalIgnoreCase))
            {
                tags.Add("posting");
            }
        }

        return tags.OrderBy(tag => tag).ToList();
    }

    private static List<string> SplitCommaList(string value)
    {
        return (value ?? string.Empty)
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<string> SplitLineList(string value)
    {
        return (value ?? string.Empty)
            .Replace("\r\n", "\n")
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string BuildEnvSummary(Dictionary<string, string> envValues)
    {
        string GetState(params string[] keys) => keys.Any(key => envValues.TryGetValue(key, out var value) && !string.IsNullOrWhiteSpace(value)) ? "ready" : "missing";

        return $"anthropic: {GetState("ANTHROPIC_API_KEY")}{Environment.NewLine}" +
               $"openai: {GetState("OPENAI_API_KEY")}{Environment.NewLine}" +
               $"elevenlabs: {GetState("ELEVENLABS_API_KEY")}{Environment.NewLine}" +
               $"voice id: {(envValues.TryGetValue("ELEVENLABS_VOICE_ID", out var voiceId) && !string.IsNullOrWhiteSpace(voiceId) ? voiceId : "not set")}{Environment.NewLine}" +
               $"stt: {(envValues.TryGetValue("STT_PROVIDER", out var sttProvider) && !string.IsNullOrWhiteSpace(sttProvider) ? sttProvider : "whisper")}{Environment.NewLine}" +
               $"push-to-talk key: {(envValues.TryGetValue("PUSH_TO_TALK_KEY", out var pushToTalkKey) && !string.IsNullOrWhiteSpace(pushToTalkKey) ? pushToTalkKey : "F8")}";
    }

    private static string BuildRetrievalSources(AssistantRunResult result)
    {
        var parts = new List<string>
        {
            $"provider: {result.Provider}",
            $"model: {result.Model}",
            $"screens: {result.Screens.Count}",
            $"knowledge chunks: {result.KnowledgeChunks.Count}"
        };

        if (result.ActionPlan.PointTag != null)
        {
            parts.Add($"point tag: screen {result.ActionPlan.PointTag.ScreenIndex} @ {result.ActionPlan.PointTag.XPercent},{result.ActionPlan.PointTag.YPercent} ({result.ActionPlan.PointTag.Label})");
        }

        if (result.ActionPlan.Steps.Count > 0)
        {
            parts.Add("action plan: " + string.Join(" -> ", result.ActionPlan.Steps.Select(step => step.ActionArgument)));
        }

        if (result.KnowledgeChunks.Count > 0)
        {
            parts.Add("sources: " + string.Join(", ", result.KnowledgeChunks.Select(chunk => chunk.Title).Distinct(StringComparer.OrdinalIgnoreCase)));
        }

        return string.Join(Environment.NewLine, parts);
    }

    private void ApplyAssistantTargeting(AssistantRunResult result)
    {
        var pointTag = result.ActionPlan.PointTag;
        if (pointTag == null)
        {
            if (result.ActionPlan.Steps.Count == 0)
            {
                _companionOverlayService.ClearTargetAnchor();
            }

            return;
        }

        var screen = result.Screens.FirstOrDefault(item => item.ScreenIndex == pointTag.ScreenIndex);
        if (screen == null || screen.Width <= 0 || screen.Height <= 0)
        {
            return;
        }

        var x = screen.Left + (int)Math.Round(screen.Width * (pointTag.XPercent / 100.0));
        var y = screen.Top + (int)Math.Round(screen.Height * (pointTag.YPercent / 100.0));
        var risk = InferPlanRiskLevel(result.ActionPlan);
        _companionOverlayService.ShowTargetAnchor(x, y, pointTag.Label, risk);
    }

    private static string BuildActionPlanPreview(AssistantRunResult result, ActiveWindowInfo active, AxContextSnapshot axContext)
    {
        var parts = new List<string>();
        parts.Add($"risk: {InferPlanRiskLevel(result.ActionPlan)}");
        parts.Add($"active app: {active.AppKind}");
        if (!string.IsNullOrWhiteSpace(axContext.FormName))
        {
            parts.Add($"active form: {axContext.FormName}");
        }

        var confirmationMode = InferPlanRiskLevel(result.ActionPlan) switch
        {
            "high" => "step-by-step only",
            "medium" => "preview first",
            _ => "full run allowed"
        };
        parts.Add($"execution mode: {confirmationMode}");

        if (result.ActionPlan.PointTag != null)
        {
            parts.Add($"target: screen {result.ActionPlan.PointTag.ScreenIndex} @ {result.ActionPlan.PointTag.XPercent},{result.ActionPlan.PointTag.YPercent} -> {result.ActionPlan.PointTag.Label}");
        }

        if (result.ActionPlan.Steps.Count > 0)
        {
            parts.AddRange(result.ActionPlan.Steps.Select((step, index) =>
            {
                var directives = new List<string>();
                if (step.WaitMilliseconds > 0)
                {
                    directives.Add($"wait {step.WaitMilliseconds}ms");
                }

                if (!string.IsNullOrWhiteSpace(step.RequiredAppContains))
                {
                    directives.Add($"ifapp {step.RequiredAppContains}");
                }

                if (!string.IsNullOrWhiteSpace(step.RequiredFormContains))
                {
                    directives.Add($"if_form {step.RequiredFormContains}");
                }

                if (!string.IsNullOrWhiteSpace(step.RequiredDialogContains))
                {
                    directives.Add($"if_dialog {step.RequiredDialogContains}");
                }

                if (!string.IsNullOrWhiteSpace(step.RequiredTabContains))
                {
                    directives.Add($"if_tab {step.RequiredTabContains}");
                }

                if (step.RetryCount > 0)
                {
                    directives.Add($"retry {step.RetryCount}");
                }

                if (!string.Equals(step.OnFail, "stop", StringComparison.OrdinalIgnoreCase))
                {
                    directives.Add($"on_fail {step.OnFail}");
                }

                var suffix = directives.Count == 0 ? string.Empty : $" [{string.Join(", ", directives)}]";
                return $"{index + 1}. {step.ActionArgument}{suffix}";
            }));
        }

        return parts.Count == 0 ? "no pending action plan" : string.Join(Environment.NewLine, parts);
    }

    private async Task<List<string>> ExecutePlanStepsAsync(int startIndex, int count)
    {
        var active = _workspaceService.GetActiveWindow();
        var executed = new List<string>();
        ActionPlanState = "running";
        _companionOverlayService.SetState("thinking", "running action plan", InferPlanRiskLevel(_pendingActionPlan));
        var endExclusive = Math.Min(_pendingActionPlan.Steps.Count, startIndex + Math.Max(1, count));
        for (var index = startIndex; index < endExclusive; index++)
        {
            var step = _pendingActionPlan.Steps[index];
            var axContext = _axClientAutomationService.CaptureActiveContext();
            CurrentPlanStepSummary = BuildStepSummary(index, step, axContext);
            _companionOverlayService.ShowActionPreview(InferRiskLevel(step.ActionArgument), CurrentPlanStepSummary, active);
            if (!string.IsNullOrWhiteSpace(step.RequiredAppContains))
            {
                var activeName = active.DisplayName;
                if (!activeName.Contains(step.RequiredAppContains, StringComparison.OrdinalIgnoreCase) &&
                    !active.AppKind.Contains(step.RequiredAppContains, StringComparison.OrdinalIgnoreCase))
                {
                    executed.Add($"step {index + 1}: skip {step.ActionArgument} (ifapp mismatch: {step.RequiredAppContains})");
                    _nextPlanStepIndex = index + 1;
                    continue;
                }
            }

            if (!string.IsNullOrWhiteSpace(step.RequiredFormContains) &&
                !axContext.FormName.Contains(step.RequiredFormContains, StringComparison.OrdinalIgnoreCase))
            {
                executed.Add($"step {index + 1}: skip {step.ActionArgument} (if_form mismatch: {step.RequiredFormContains})");
                _nextPlanStepIndex = index + 1;
                if (string.Equals(step.OnFail, "stop", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }

                continue;
            }

            if (!string.IsNullOrWhiteSpace(step.RequiredDialogContains) &&
                !axContext.DialogTitle.Contains(step.RequiredDialogContains, StringComparison.OrdinalIgnoreCase))
            {
                executed.Add($"step {index + 1}: skip {step.ActionArgument} (if_dialog mismatch: {step.RequiredDialogContains})");
                _nextPlanStepIndex = index + 1;
                if (string.Equals(step.OnFail, "stop", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }

                continue;
            }

            if (!string.IsNullOrWhiteSpace(step.RequiredTabContains) &&
                !axContext.ActiveTab.Contains(step.RequiredTabContains, StringComparison.OrdinalIgnoreCase))
            {
                executed.Add($"step {index + 1}: skip {step.ActionArgument} (if_tab mismatch: {step.RequiredTabContains})");
                _nextPlanStepIndex = index + 1;
                if (string.Equals(step.OnFail, "stop", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }

                continue;
            }

            if (step.WaitMilliseconds > 0)
            {
                executed.Add($"step {index + 1}: wait {step.WaitMilliseconds}ms");
                await Task.Delay(step.WaitMilliseconds).ConfigureAwait(false);
            }

            var attempts = Math.Max(1, step.RetryCount + 1);
            string? lastResult = null;
            var success = false;
            for (var attempt = 1; attempt <= attempts; attempt++)
            {
                lastResult = step.ActionArgument.StartsWith("ax.", StringComparison.OrdinalIgnoreCase)
                    ? _axClientAutomationService.Execute(step.ActionArgument, axContext)
                    : _desktopInspectorService.Execute(step.ActionArgument);
                if (IsSuccessfulActionResult(lastResult))
                {
                    success = true;
                    break;
                }

                if (attempt < attempts)
                {
                    await Task.Delay(220).ConfigureAwait(false);
                }
            }

            executed.Add($"step {index + 1}: {step.ActionArgument} -> {(success ? lastResult : "failed")}");
            AppendActionHistory(
                success ? "plan-step" : "plan-step-failed",
                step.ActionArgument,
                success ? lastResult ?? string.Empty : "failed",
                step.ActionArgument,
                active.AppKind);
            _nextPlanStepIndex = index + 1;
            if (!success && string.Equals(step.OnFail, "stop", StringComparison.OrdinalIgnoreCase))
            {
                ActionPlanState = "error";
                _companionOverlayService.SetState("error", BuildFailureOverlayMessage(step, axContext), InferRiskLevel(step.ActionArgument));
                break;
            }
        }

        CurrentPlanStepSummary = BuildCurrentStepSummary();
        RefreshActiveWindow();
        return executed;
    }

    private void AppendExecutionLog(List<string> executed)
    {
        var log = string.Join(Environment.NewLine, executed);
        InspectorResult = log;
        if (string.IsNullOrWhiteSpace(log))
        {
            return;
        }

        ActionPlanExecutionLog = ActionPlanExecutionLog is "no plan executed yet" or "plan ready for execution" or "plan cleared"
            ? log
            : ActionPlanExecutionLog + Environment.NewLine + log;
    }

    private string BuildCurrentStepSummary()
    {
        if (_pendingActionPlan.Steps.Count == 0)
        {
            return "no active step";
        }

        if (_nextPlanStepIndex >= _pendingActionPlan.Steps.Count)
        {
            return "all plan steps executed";
        }

        var next = _pendingActionPlan.Steps[_nextPlanStepIndex];
        return $"next step {_nextPlanStepIndex + 1}/{_pendingActionPlan.Steps.Count}: {next.ActionArgument}";
    }

    private string BuildStepSummary(int index, AssistantActionStep step, AxContextSnapshot axContext)
    {
        if (!step.ActionArgument.StartsWith("ax.", StringComparison.OrdinalIgnoreCase))
        {
            return $"step {index + 1}/{_pendingActionPlan.Steps.Count}: {step.ActionArgument}";
        }

        var suffix = new List<string>();
        if (!string.IsNullOrWhiteSpace(axContext.FormName))
        {
            suffix.Add($"form {axContext.FormName}");
        }

        if (!string.IsNullOrWhiteSpace(axContext.DialogTitle))
        {
            suffix.Add($"dialog {axContext.DialogTitle}");
        }

        if (!string.IsNullOrWhiteSpace(axContext.ActiveTab))
        {
            suffix.Add($"tab {axContext.ActiveTab}");
        }

        return suffix.Count == 0
            ? $"step {index + 1}/{_pendingActionPlan.Steps.Count}: {step.ActionArgument}"
            : $"step {index + 1}/{_pendingActionPlan.Steps.Count}: {step.ActionArgument} [{string.Join(", ", suffix)}]";
    }

    private static bool IsSuccessfulActionResult(string? result)
    {
        if (string.IsNullOrWhiteSpace(result))
        {
            return false;
        }

        return !result.Contains("no matching", StringComparison.OrdinalIgnoreCase) &&
               !result.Contains("unsupported", StringComparison.OrdinalIgnoreCase) &&
               !result.Contains("not found", StringComparison.OrdinalIgnoreCase) &&
               !result.Contains("failed", StringComparison.OrdinalIgnoreCase) &&
               !result.Contains("not editable", StringComparison.OrdinalIgnoreCase) &&
               !result.Contains("not an AX client", StringComparison.OrdinalIgnoreCase);
    }

    private static string BuildFailureOverlayMessage(AssistantActionStep step, AxContextSnapshot axContext)
    {
        if (!step.ActionArgument.StartsWith("ax.", StringComparison.OrdinalIgnoreCase))
        {
            return "action plan step failed";
        }

        if (!string.IsNullOrWhiteSpace(axContext.ValidationSummary))
        {
            return $"validation dialog detected: {axContext.ValidationSummary}";
        }

        if (step.ActionArgument.StartsWith("ax.set_field:", StringComparison.OrdinalIgnoreCase))
        {
            return "field not found or not writable";
        }

        if (step.ActionArgument.StartsWith("ax.select_grid_row:", StringComparison.OrdinalIgnoreCase))
        {
            return "grid row not found";
        }

        if (step.ActionArgument.StartsWith("ax.open_dialog:", StringComparison.OrdinalIgnoreCase))
        {
            return "lookup unresolved";
        }

        return "AX step failed";
    }

    private static string BuildOverlaySnippet(string responseText)
    {
        var compact = (responseText ?? string.Empty).Replace(Environment.NewLine, " ").Trim();
        if (compact.Length == 0)
        {
            return "reply ready";
        }

        return compact.Length <= 110 ? compact : compact[..110].TrimEnd() + "...";
    }

    private static string InferRiskLevel(string action)
    {
        var normalized = (action ?? string.Empty).Trim().ToLowerInvariant();
        if (normalized.StartsWith("type_control:", StringComparison.Ordinal) ||
            normalized.StartsWith("click_control:", StringComparison.Ordinal) ||
            normalized.StartsWith("ax.set_field:", StringComparison.Ordinal))
        {
            return "medium";
        }

        if (normalized.StartsWith("ax.click_action:", StringComparison.Ordinal))
        {
            var risky = normalized.Contains("post", StringComparison.OrdinalIgnoreCase) ||
                        normalized.Contains("delete", StringComparison.OrdinalIgnoreCase) ||
                        normalized.Contains("confirm", StringComparison.OrdinalIgnoreCase) ||
                        normalized.Contains("save", StringComparison.OrdinalIgnoreCase) ||
                        normalized.Contains("book", StringComparison.OrdinalIgnoreCase);
            return risky ? "high" : "medium";
        }

        if (normalized.StartsWith("focus_control:", StringComparison.Ordinal) ||
            normalized == "focus_window" ||
            normalized == "list_controls" ||
            normalized.StartsWith("read_", StringComparison.Ordinal) ||
            normalized.StartsWith("ax.read_", StringComparison.Ordinal) ||
            normalized == "ax.focus_window")
        {
            return "low";
        }

        if (normalized == "ax.post" || normalized.StartsWith("ax.post:", StringComparison.Ordinal))
        {
            return "high";
        }

        return "safe";
    }

    private static string InferPlanRiskLevel(AssistantActionPlan plan)
    {
        if (plan.Steps.Count == 0)
        {
            return "low";
        }

        if (plan.Steps.Any(step => InferRiskLevel(step.ActionArgument) == "high"))
        {
            return "high";
        }

        if (plan.Steps.Any(step => InferRiskLevel(step.ActionArgument) == "medium"))
        {
            return "medium";
        }

        return "low";
    }

    private static void TryDeleteFile(string? filePath)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
        catch
        {
        }
    }

    private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private static string InferRecipeCategory(string prompt)
    {
        return prompt.Contains("ax.", StringComparison.OrdinalIgnoreCase) ? "ax" : "general";
    }

    private static string InferRecipeRisk(AutomationRecipe recipe)
    {
        if (recipe.Steps.Any(step => InferRiskLevel(step.ActionArgument) == "high"))
        {
            return "high";
        }

        if (recipe.Steps.Any(step => InferRiskLevel(step.ActionArgument) == "medium"))
        {
            return "medium";
        }

        return InferRiskLevel(recipe.Prompt);
    }

    private static string ExtractRecipePrompt(string editorText)
    {
        if (string.IsNullOrWhiteSpace(editorText))
        {
            return string.Empty;
        }

        var lines = editorText.Replace("\r\n", "\n").Split('\n');
        var promptLines = lines.Where(line =>
            !line.StartsWith("category:", StringComparison.OrdinalIgnoreCase) &&
            !line.StartsWith("guard app:", StringComparison.OrdinalIgnoreCase) &&
            !line.StartsWith("guard form:", StringComparison.OrdinalIgnoreCase));
        return string.Join(Environment.NewLine, promptLines).Trim();
    }

    private static string BuildRecipeEditorText(AutomationRecipe recipe)
    {
        var header = new List<string>
        {
            $"category: {recipe.Category}"
        };

        if (!string.IsNullOrWhiteSpace(recipe.GuardApp))
        {
            header.Add($"guard app: {recipe.GuardApp}");
        }

        if (!string.IsNullOrWhiteSpace(recipe.GuardForm))
        {
            header.Add($"guard form: {recipe.GuardForm}");
        }

        if (!string.IsNullOrWhiteSpace(recipe.GuardDialog))
        {
            header.Add($"guard dialog: {recipe.GuardDialog}");
        }

        if (!string.IsNullOrWhiteSpace(recipe.GuardTab))
        {
            header.Add($"guard tab: {recipe.GuardTab}");
        }

        return string.Join(Environment.NewLine, header) + Environment.NewLine + Environment.NewLine + recipe.Prompt;
    }

    private string BuildSuggestedRecipeName()
    {
        var axContext = _axClientAutomationService.CaptureActiveContext();
        var prefix = _pendingActionPlan.Steps.All(step => step.ActionArgument.StartsWith("ax.", StringComparison.OrdinalIgnoreCase)) ? "AX" : "Plan";
        var subject = !string.IsNullOrWhiteSpace(axContext.FormName) ? axContext.FormName : "Workflow";
        return $"{prefix} {subject} {DateTime.Now:HHmmss}";
    }

    private void AppendActionHistory(string actionName, string actionArgument, string result, string targetLabel, string activeApp)
    {
        var entries = _workspaceService.Load().ActionHistory.ToList();
        entries.Insert(0, new ActionHistoryEntry
        {
            TimestampUtc = DateTime.UtcNow.ToString("o"),
            ActionName = actionName,
            ActionArgument = actionArgument,
            TargetLabel = targetLabel,
            SpokenText = AssistantPrompt,
            ActiveApp = activeApp,
            Category = actionArgument.StartsWith("ax.", StringComparison.OrdinalIgnoreCase) ? "ax" : "general",
            Result = result
        });

        _workspaceService.SaveActionHistory(entries.Take(200).ToList());
        ReplaceCollection(ActionHistory, FilterHistory(entries, HistorySearch));
        if (ActionHistory.Count > 0)
        {
            SetSelectedHistoryEntry(ActionHistory[0]);
        }
    }
}
