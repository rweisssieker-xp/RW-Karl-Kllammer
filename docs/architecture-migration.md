# Architecture migration: core library + Windows platform layer

This document tracks the migration path described in [`avalonia/README.md`](../avalonia/README.md) (*intended migration path*).

## Current state (implemented in repo)

| Layer | Project | Responsibility |
|-------|---------|----------------|
| Cross-platform core | [`CarolusNexus.Core`](../CarolusNexus.Core/CarolusNexus.Core.csproj) | Workspace path resolution (`WorkspaceLayout`, `WorkspacePathResolver`), shared DTOs such as `ActiveWindowInfo` |
| Windows platform | [`CarolusNexus.Platform.Windows`](../CarolusNexus.Platform.Windows/CarolusNexus.Platform.Windows.csproj) | Win32 foreground window capture and app-kind / desktop-framework heuristics (`WindowsForegroundWindow`) |
| UI + composition | [`avalonia/ClippyRW.Avalonia.csproj`](../avalonia/ClippyRW.Avalonia.csproj) | Avalonia shell, view models, orchestration |

The Avalonia app references Core and Platform.Windows. Workspace path logic and active-window capture are the first extracted seams; file/knowledge/ritual logic remains in `OperatorWorkspaceService` until further splits make sense.

## Target state (incremental)

1. **Core** — Assistant protocol, knowledge indexing, ritual models, and provider-agnostic DTOs (move from `avalonia/` as touch points stabilize).
2. **Platform.Windows** — All P/Invoke, WMP/COM audio if needed later, shell tray/hotkey wrappers that are not UI.
3. **Avalonia** — Presentation only: binds to Core abstractions via small composition root (`App.axaml.cs`).

## Build

- App (existing script): [`avalonia/Build-Avalonia.cmd`](../avalonia/Build-Avalonia.cmd) — resolves project references automatically.
- Full solution (optional): `dotnet build CarolusNexus.sln -c Release`

## Next extraction candidates

- `AssistantRuntimeService` HTTP + prompt building → Core (with injectable `IHttpClientFactory` or raw `HttpClient` from host).
- `LocalAgentRunService` → Core or Platform (process spawn is mostly generic on Windows).
- `DesktopInspectorService` / `AxClientAutomationService` → Platform.Windows only.
