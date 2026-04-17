# Feature parity: Legacy WinForms vs Avalonia

Reference sources: [`windows/Clicky.Windows.cs`](../windows/Clicky.Windows.cs), [`avalonia/ViewModels/MainWindowViewModel.cs`](../avalonia/ViewModels/MainWindowViewModel.cs), [`avalonia/Services/AssistantRuntimeService.cs`](../avalonia/Services/AssistantRuntimeService.cs).

| Area | Legacy (WinForms) | Avalonia | Parity |
|------|-------------------|----------|--------|
| Ask + vision (Anthropic / OpenAI / compatible) | Yes | Yes | Full |
| Local knowledge RAG | Yes | Yes | Full |
| Rituals / recipes / watch / history / diagnostics | Yes (separate dialogs) | Yes (tabs) | Full (UX differs) |
| Setup wizard | Dedicated dialog | Setup tab (see MainWindow.axaml) | Functional equivalent |
| Codex / Claude Code / OpenClaw | Trigger phrases + dedicated flows | Console tab + `LocalAgentRunService` | Partial |
| Codex trigger in ask flow | `CodexClient.IsTriggered` before cloud ask | `AgentHandoffTriggers` + `RunAssistantAsync` | Full |
| Auto-route to Codex | `IntentRouter.DetectRoute` (IDE context + coding keywords → `codex`) | `AgentHandoffTriggers.DetectIntentRoute` | Full |
| Auto-route to OpenClaw | `IntentRouter` (agent/workflow keywords → `openclaw`) | same | Full |
| TTS primary | ElevenLabs → temp MP3 → Windows Media Player COM | ElevenLabs → MP3 + open file (`SynthesizeSpeechToFileAsync`) | Partial |
| TTS failure fallback | `SpeakLocalFallback` via `SAPI.SpVoice` | `WindowsSpeechFallback.TrySpeak` after ElevenLabs failure or missing keys | Full |
| Speak after ask | Companion navigates with optional speaking state; `PlayResponseAudio` path | `SpeakResponses` toggles setting; synthesis is explicit button (`SynthesizeCurrentResponseAsync`) | Partial |
| Push-to-talk / tray | Yes | Yes | Full |
| Companion overlay + point tags | Yes | Yes | Full (verify edge cases separately) |
| Window inspector / semantic actions | Rich legacy adapters | Desktop inspector + AX adapter (v1) | Partial (see README “Still Not Ported”) |

## Code anchors

- Legacy auto-route: `RunAskFlowAsync` calls `IntentRouter.DetectRoute` then `RunCodexFlowAsync` / `RunOpenClawFlowAsync` (approx. lines 6076–6088 in `Clicky.Windows.cs`).
- Avalonia ask entry: `MainWindowViewModel.RunAssistantAsync` calls `_assistantRuntimeService.AskAsync` directly with no CLI or intent routing.
- Legacy TTS fallback: `SpeakLocalFallback` using `SAPI.SpVoice` after WMP / ElevenLabs failure (approx. lines 7967–8033).
- Avalonia TTS: `AssistantRuntimeService.SynthesizeSpeechToFileAsync` and `MainWindowViewModel.SynthesizeCurrentResponseAsync` (no SAPI branch).

## Suggested follow-ups (if 1:1 parity is desired)

1. Port `IntentRouter` (or equivalent) ahead of `AskAsync` in Avalonia, reusing active-window app kind from `OperatorWorkspaceService.GetActiveWindow`.
2. Mirror legacy CLI trigger checks (`OpenClawClient`, `ClaudeCodeClient`, `CodexClient`) in the ask pipeline before the cloud call.
3. Add optional Windows SAPI / local playback fallback after ElevenLabs or file-open failure when `SpeakResponses` is enabled.
