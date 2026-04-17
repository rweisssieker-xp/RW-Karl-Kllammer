using ClippyRWAvalonia.Services;
var workspace = new OperatorWorkspaceService();
var runtime = new AssistantRuntimeService(workspace);
try
{
    var result = await runtime.SmokeTestAsync("anthropic", "");
    Console.WriteLine("anthropic.default=ok");
    Console.WriteLine(result);
}
catch (Exception ex)
{
    Console.WriteLine("anthropic.default=fail");
    Console.WriteLine(ex.GetType().Name + ":" + ex.Message);
}
