# AX / fat-client backlog (prioritized)

Derived from [`avalonia/README.md`](../avalonia/README.md) section *Still Not Ported* and root [`README.md`](../README.md) *Known Limitations*.

## P0 — Reliability on real AX surfaces

1. **Grid row selection** — Robust `ax.select_grid_row` (and related) across grid variants; handle virtual scrolling and selection state.  
   *Incremental (repo):* multi-grid scoring, focus + `{HOME}` + type-ahead + `{ENTER}` on best visible row match in [`AxClientAutomationService`](../avalonia/Services/AxClientAutomationService.cs).
2. **Lookup flows** — `ax.open_lookup` and post-lookup confirmation paths; consistent waits and dialog guards.  
   *Incremental:* double `Alt+Down` with delays; new `ax.confirm_lookup` (OK/Select/Übernehmen/Apply).
3. **Posting / save pipelines** — End-to-end “enter data → validate → post” with explicit risk gates for high-impact actions (aligned with existing ritual risk policy).  
   *Incremental:* `ax.post` clicks Post/Update/Buchen/Send-style actions; ritual policy still blocks high-risk/mail contexts.

## P1 — Discovery and accessibility depth

4. **UI Automation / MSAA** — Move beyond Win32-first hybrid heuristics where AX exposes richer automation trees; reduce false negatives on control discovery.
5. **Inspector parity** — Live Context / inspector results should match or exceed legacy semantic summaries for complex forms (tabs, nested containers).

## P2 — Authoring and learning

6. **Teach-once semantic recorder** — Higher-confidence capture than history-backed teach mode: stable selectors, step labels, and guard suggestions from a single demonstration.
7. **Regression fixtures** — Scripted dry-runs or recorded snapshots for a small set of canonical AX forms (internal test harness; optional but reduces regressions).

## P3 — Documentation and operations

8. **Operator playbooks** — Short SOPs for “when to use ritual vs ask”, “high-risk blocked actions”, and AX-specific troubleshooting (link from app or repo docs as needed).

### Suggested issue titles (copy into tracker)

- `AX: Harden grid row selection across AX grid types`
- `AX: End-to-end lookup + return to form`
- `AX: Posting workflow with risk policy + confirmations`
- `AX: UIA/MSAA-backed control discovery (phase 1)`
- `Teach-once recorder: semantic AX ritual draft v2`
