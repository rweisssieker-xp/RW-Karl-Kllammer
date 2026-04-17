using ClippyRWAvalonia.Models;
using ClippyRWAvalonia.Services;

var sample = "Open the customer form [POINT:1|45|62|search box] [ACTIONS:app|focus_window|ifapp=crm;app|focus_control:Search|retry=2|wait=250;app|type_control:Search=Alice]";
var plan = AssistantActionPlan.Parse(sample);
Console.WriteLine($"clean={plan.CleanText}");
Console.WriteLine($"point={(plan.PointTag==null?"none":$"{plan.PointTag.ScreenIndex}:{plan.PointTag.XPercent},{plan.PointTag.YPercent}:{plan.PointTag.Label}")}");
Console.WriteLine($"steps={plan.Steps.Count}");
for (var i = 0; i < plan.Steps.Count; i++)
{
    var s = plan.Steps[i];
    Console.WriteLine($"step{i+1}={s.ActionName}|{s.ActionArgument}|wait={s.WaitMilliseconds}|ifapp={s.RequiredAppContains}|retry={s.RetryCount}");
}
var inspector = new DesktopInspectorService();
Console.WriteLine("inspector=" + inspector.Execute("list_controls").Split(Environment.NewLine)[0]);
