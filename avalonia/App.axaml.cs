using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Threading;
using ClippyRWAvalonia.Services;
using ClippyRWAvalonia.ViewModels;
using ClippyRWAvalonia.Views;

namespace ClippyRWAvalonia;

public partial class App : Avalonia.Application
{
    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var service = new OperatorWorkspaceService();
            var env = service.ReadEnvFile();
            var runService = new LocalAgentRunService(service);
            var assistantRuntimeService = new AssistantRuntimeService(service);
            var microphoneRecorder = new MicrophoneRecorderService(service.DataRoot);
            var windowsShellService = new WindowsShellService(env);
            var companionOverlayService = new CompanionOverlayService();
            var desktopInspectorService = new DesktopInspectorService();
            var axClientAutomationService = new AxClientAutomationService();
            var ritualRuntimeService = new RitualRuntimeService(service, desktopInspectorService, axClientAutomationService);
            var viewModel = new MainWindowViewModel(service, runService, assistantRuntimeService, microphoneRecorder, windowsShellService, companionOverlayService, desktopInspectorService, axClientAutomationService, ritualRuntimeService);
            var window = new MainWindow
            {
                DataContext = viewModel
            };
            desktop.MainWindow = window;

            windowsShellService.OpenRequested += (_, _) => Dispatcher.UIThread.Post(() =>
            {
                window.Show();
                window.WindowState = Avalonia.Controls.WindowState.Normal;
                window.Activate();
            });
            windowsShellService.AskRequested += async (_, _) =>
            {
                await Dispatcher.UIThread.InvokeAsync(async () =>
                {
                    window.Show();
                    window.WindowState = Avalonia.Controls.WindowState.Normal;
                    window.Activate();
                    await viewModel.RunAssistantAsync();
                });
            };
            windowsShellService.HotKeyPressed += (_, _) => Dispatcher.UIThread.Post(viewModel.StartSpeechCapture);
            windowsShellService.HotKeyReleased += async (_, _) => await Dispatcher.UIThread.InvokeAsync(async () => await viewModel.StopSpeechCaptureAndAskAsync());
            windowsShellService.ExitRequested += (_, _) => Dispatcher.UIThread.Post(() => desktop.Shutdown());
            companionOverlayService.Initialize();
            windowsShellService.Initialize();
            desktop.Exit += (_, _) =>
            {
                companionOverlayService.Dispose();
                windowsShellService.Dispose();
                microphoneRecorder.Dispose();
            };
        }

        base.OnFrameworkInitializationCompleted();
    }
}
