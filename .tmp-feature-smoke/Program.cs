using ClippyRWAvalonia.Models;
using ClippyRWAvalonia.Services;

var workspace = new OperatorWorkspaceService();
var snapshot = workspace.Load();
var env = workspace.ReadEnvFile();
Console.WriteLine($"workspace.recipes={snapshot.Recipes.Count}");
Console.WriteLine($"workspace.watch={snapshot.WatchSessions.Count}");
Console.WriteLine($"workspace.history={snapshot.ActionHistory.Count}");
Console.WriteLine($"workspace.knowledgeDocs={snapshot.KnowledgeDocuments.Count}");
Console.WriteLine($"workspace.diagnostics={snapshot.Diagnostics.Count}");
Console.WriteLine($"env.anthropic={(env.ContainsKey("ANTHROPIC_API_KEY") ? "present" : "missing")}");
Console.WriteLine($"env.openai={(env.ContainsKey("OPENAI_API_KEY") ? "present" : "missing")}");
Console.WriteLine($"env.elevenlabs={(env.ContainsKey("ELEVENLABS_API_KEY") ? "present" : "missing")}");
var active = workspace.GetActiveWindow();
Console.WriteLine($"active={active.DisplayName}|kind={active.AppKind}|framework={active.DesktopFramework}|bounds={active.Left},{active.Top},{active.Width},{active.Height}");
var inspector = new DesktopInspectorService();
Console.WriteLine("inspector.list=" + inspector.Execute("list_controls").Split(Environment.NewLine)[0]);
Console.WriteLine("inspector.form=" + inspector.Execute("read_form").Split(Environment.NewLine)[0]);
var plan = AssistantActionPlan.Parse("Do it [POINT:1|45|62|search box] [ACTIONS:app|focus_window|ifapp=crm;app|focus_control:Search|retry=2|wait=250;app|type_control:Search=Alice]");
Console.WriteLine($"plan.clean={plan.CleanText}");
Console.WriteLine($"plan.point={(plan.PointTag==null?"none":$"{plan.PointTag.ScreenIndex}:{plan.PointTag.XPercent},{plan.PointTag.YPercent}:{plan.PointTag.Label}")}");
Console.WriteLine($"plan.steps={plan.Steps.Count}");
foreach (var step in plan.Steps)
{
  Console.WriteLine($"plan.step={step.ActionName}|{step.ActionArgument}|wait={step.WaitMilliseconds}|ifapp={step.RequiredAppContains}|retry={step.RetryCount}");
}
if (env.ContainsKey("ANTHROPIC_API_KEY"))
{
  try
  {
    var runtime = new AssistantRuntimeService(workspace);
    var smoke = await runtime.SmokeTestAsync("anthropic", "");
    Console.WriteLine("smoke.anthropic=ok");
    Console.WriteLine(smoke.Split(Environment.NewLine)[0]);
  }
  catch (Exception ex)
  {
    Console.WriteLine("smoke.anthropic=fail");
    Console.WriteLine(ex.GetType().Name + ":" + ex.Message);
  }
}
if (env.ContainsKey("OPENAI_API_KEY"))
{
  try
  {
    var runtime = new AssistantRuntimeService(workspace);
    var smoke = await runtime.SmokeTestAsync("openai", "");
    Console.WriteLine("smoke.openai=ok");
    Console.WriteLine(smoke.Split(Environment.NewLine)[0]);
  }
  catch (Exception ex)
  {
    Console.WriteLine("smoke.openai=fail");
    Console.WriteLine(ex.GetType().Name + ":" + ex.Message);
  }
}
