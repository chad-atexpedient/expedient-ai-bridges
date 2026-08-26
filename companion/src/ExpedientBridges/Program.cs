using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Text.Json;

namespace Expedient.Bridges.Companion;

internal static class Program
{
    private const string Product = "Expedient AI Bridges";
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static async Task<int> Main(string[] args)
    {
        var command = args.FirstOrDefault()?.ToLowerInvariant() ?? "status";
        try
        {
            var paths = CompanionPaths.Resolve();
            paths.Create();
            return command switch
            {
                "start" => await StartClientAsync(paths),
                "--background" => await SupervisorAsync(paths),
                "stop" or "--stop" => await StopClientAsync(paths),
                "restart" => await RestartAsync(paths),
                "status" or "--status" => await StatusAsync(paths),
                "diagnostics" or "--diagnostics" => await DiagnosticsAsync(paths),
                "migrate" or "--migrate" => await Migration.RunAsync(paths),
                "--uninstall" => UninstallInstructions(paths),
                _ => Usage()
            };
        }
        catch (Exception ex) { Console.Error.WriteLine($"{Product}: {ex.Message}"); return 1; }
    }

    private static async Task<int> StartClientAsync(CompanionPaths paths)
    {
        await Migration.RunAsync(paths, quiet: true);
        var existing = await RuntimeState.ReadAsync(paths.State);
        if (existing?.IsSupervisorAlive() == true)
        {
            Console.WriteLine(existing.Ready ? "Bridges are already ready." : "Bridge supervisor is already starting.");
            return existing.Ready ? 0 : 4;
        }
        RuntimeState.DeleteStale(paths);
        Validate(paths);
        var self = Environment.ProcessPath ?? throw new InvalidOperationException("Cannot resolve companion executable.");
        var start = new ProcessStartInfo(self) { UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = paths.Program };
        start.ArgumentList.Add("--background");
        start.Environment["EXPEDIENT_SUPERVISOR_PARENT"] = Environment.ProcessId.ToString();
        Process.Start(start)?.Dispose();
        var deadline = DateTime.UtcNow.AddSeconds(20);
        while (DateTime.UtcNow < deadline)
        {
            await Task.Delay(250);
            var state = await RuntimeState.ReadAsync(paths.State);
            if (state?.Ready == true && state.AllAlive()) { Console.WriteLine("Bridges ready on loopback ports 4001 and 4002."); return 0; }
            if (state?.Failure is { Length: > 0 } failure) throw new InvalidOperationException(failure);
        }
        throw new TimeoutException($"Supervisor readiness timed out. See {paths.SupervisorLog}");
    }

    private static async Task<int> SupervisorAsync(CompanionPaths paths)
    {
        await using var lease = SingleInstanceLease.TryAcquire(paths.Lock) ?? throw new InvalidOperationException("Another supervisor owns this profile.");
        if (File.Exists(paths.StopRequest)) File.Delete(paths.StopRequest);
        var config = Validate(paths);
        var generation = Guid.NewGuid().ToString("N");
        var self = Process.GetCurrentProcess();
        var state = new RuntimeState(generation, ProcessIdentity.From(self, "supervisor", Environment.ProcessPath!, null), null, null, false, null, DateTimeOffset.UtcNow);
        await RuntimeState.WriteAtomicAsync(paths.State, state);
        var secrets = config.Where(x => SecretRedactor.IsSecretKey(x.Key)).Select(x => x.Value).Where(x => x.Length >= 4).ToArray();
        var restarts = new Queue<DateTimeOffset>();
        try
        {
            while (!File.Exists(paths.StopRequest))
            {
                Process? codex = null, anthropic = null;
                try
                {
                    var node = RuntimeLocator.FindNode(paths.Program);
                    RuntimeLocator.RequireNode22(node);
                    var ports = ConfigValidator.Ports(config);
                    foreach (var port in ports) if (await HealthProbe.HttpRespondingAsync(port)) throw new InvalidOperationException($"Port {port} is owned by another process.");
                    LogFiles.Rotate(paths.CodexLog); LogFiles.Rotate(paths.AnthropicLog);
                    codex = StartNode(node, Path.Combine(paths.Program, "app", "index.js"), config, paths.CodexLog, secrets);
                    anthropic = StartNode(node, Path.Combine(paths.Program, "app", "anthropic-bridge.js"), config, paths.AnthropicLog, secrets);
                    state = state with { Codex = ProcessIdentity.From(codex, "codex", node, Path.Combine(paths.Program, "app", "index.js")), Anthropic = ProcessIdentity.From(anthropic, "anthropic", node, Path.Combine(paths.Program, "app", "anthropic-bridge.js")), Ready = false, Failure = null };
                    await RuntimeState.WriteAtomicAsync(paths.State, state);
                    var deadline = DateTime.UtcNow.AddSeconds(15);
                    while (DateTime.UtcNow < deadline && !(await HealthProbe.HttpRespondingAsync(ports[0]) && await HealthProbe.HttpRespondingAsync(ports[1])))
                    {
                        if (codex.HasExited || anthropic.HasExited) throw new InvalidOperationException("A bridge exited during startup.");
                        await Task.Delay(250);
                    }
                    if (!(await HealthProbe.HttpRespondingAsync(ports[0]) && await HealthProbe.HttpRespondingAsync(ports[1]))) throw new TimeoutException("Application-level readiness timed out.");
                    state = state with { Ready = true }; await RuntimeState.WriteAtomicAsync(paths.State, state);
                    while (!File.Exists(paths.StopRequest) && !codex.HasExited && !anthropic.HasExited) await Task.Delay(500);
                    if (File.Exists(paths.StopRequest)) break;
                    restarts.Enqueue(DateTimeOffset.UtcNow);
                    while (restarts.Count > 0 && DateTimeOffset.UtcNow - restarts.Peek() > TimeSpan.FromMinutes(5)) restarts.Dequeue();
                    if (restarts.Count > 5) throw new InvalidOperationException("Restart budget exhausted: more than five failures in five minutes.");
                    state = state with { Ready = false, Failure = "A bridge exited unexpectedly; bounded restart scheduled." }; await RuntimeState.WriteAtomicAsync(paths.State, state);
                    await Task.Delay(TimeSpan.FromSeconds(Math.Min(5, Math.Pow(2, restarts.Count - 1))));
                }
                finally { StopOwned(codex); StopOwned(anthropic); }
            }
            return 0;
        }
        catch (Exception ex)
        {
            state = state with { Ready = false, Failure = SecretRedactor.Redact(ex.Message, secrets) };
            await RuntimeState.WriteAtomicAsync(paths.State, state); await LogFiles.AppendAsync(paths.SupervisorLog, state.Failure); return 1;
        }
        finally
        {
            if (File.Exists(paths.StopRequest)) File.Delete(paths.StopRequest);
            if (File.Exists(paths.State)) File.Delete(paths.State);
        }
    }

    private static Process StartNode(string node, string script, IReadOnlyDictionary<string,string> config, string log, string[] secrets)
    {
        if (!File.Exists(script)) throw new FileNotFoundException("Bridge bundle not found", script);
        var info = new ProcessStartInfo(node) { UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = Path.GetDirectoryName(Path.GetDirectoryName(script))!, RedirectStandardOutput = true, RedirectStandardError = true };
        info.ArgumentList.Add(script);
        foreach (var pair in config) info.Environment[pair.Key] = pair.Value;
        if (!info.Environment.ContainsKey("NODE_OPTIONS")) info.Environment["NODE_OPTIONS"] = "--max-old-space-size=8192";
        info.Environment["BRIDGE_METRICS_DIR"] = Path.GetDirectoryName(log)!;
        var p = Process.Start(info) ?? throw new InvalidOperationException($"Could not start {script}");
        _ = LogFiles.PumpAsync(p.StandardOutput, log, secrets); _ = LogFiles.PumpAsync(p.StandardError, log, secrets); return p;
    }

    private static Dictionary<string,string> Validate(CompanionPaths paths)
    {
        var config = EnvFile.Load(paths.Config);
        ConfigValidator.Validate(config);
        foreach (var script in new[] { "index.js", "anthropic-bridge.js" }) if (!File.Exists(Path.Combine(paths.Program, "app", script))) throw new FileNotFoundException("Bridge bundle missing", script);
        return config;
    }

    private static async Task<int> StopClientAsync(CompanionPaths paths)
    {
        var state = await RuntimeState.ReadAsync(paths.State);
        if (state?.IsSupervisorAlive() != true) { RuntimeState.DeleteStale(paths); Console.WriteLine("Bridges are stopped."); return 0; }
        await File.WriteAllTextAsync(paths.StopRequest, state.Generation);
        var deadline = DateTime.UtcNow.AddSeconds(15);
        while (DateTime.UtcNow < deadline && state.IsSupervisorAlive()) await Task.Delay(250);
        if (state.IsSupervisorAlive()) throw new TimeoutException("Supervisor did not stop; refusing unsafe PID-only termination.");
        Console.WriteLine("Bridges stopped."); return 0;
    }

    private static async Task<int> RestartAsync(CompanionPaths paths) { await StopClientAsync(paths); return await StartClientAsync(paths); }
    private static async Task<int> StatusAsync(CompanionPaths paths) { var state = await RuntimeState.ReadAsync(paths.State); var ready = state?.Ready == true && state.AllAlive(); Console.WriteLine(ready ? "ready" : state?.IsSupervisorAlive() == true ? "degraded" : "stopped"); return ready ? 0 : state?.IsSupervisorAlive() == true ? 4 : 3; }
    private static async Task<int> DiagnosticsAsync(CompanionPaths paths)
    {
        var config = EnvFile.Load(paths.Config, false); var state = await RuntimeState.ReadAsync(paths.State); var ports = ConfigValidator.Ports(config);
        var report = new { product = Product, version = typeof(Program).Assembly.GetName().Version?.ToString(), os = Environment.OSVersion.ToString(), architecture = System.Runtime.InteropServices.RuntimeInformation.OSArchitecture.ToString(), program = paths.Program, data = paths.Data, configPresent = File.Exists(paths.Config), upstreamConfigured = config.ContainsKey("OPENAI_BASE_URL"), state = state is null ? "stopped" : state.Ready && state.AllAlive() ? "ready" : state.IsSupervisorAlive() ? "degraded" : "stale", generation = state?.Generation, ports = await Task.WhenAll(ports.Select(async p => new { port = p, responding = await HealthProbe.HttpRespondingAsync(p) })), failure = state?.Failure };
        Console.WriteLine(JsonSerializer.Serialize(report, JsonOptions)); return 0;
    }
    private static void StopOwned(Process? p) { if (p is null) return; try { if (!p.HasExited) { p.Kill(true); p.WaitForExit(5000); } } catch { } finally { p.Dispose(); } }
    private static int UninstallInstructions(CompanionPaths p) { Console.WriteLine($"Run 'stop', remove the application, then optionally remove {p.Data}"); return 0; }
    private static int Usage() { Console.WriteLine("Usage: ExpedientBridges [start|stop|restart|status|diagnostics|migrate]"); return 2; }
}

internal sealed record CompanionPaths(string Program,string Data,string Runtime,string Logs,string Config,string State,string Lock,string StopRequest,string SupervisorLog,string CodexLog,string AnthropicLog)
{
    public static CompanionPaths Resolve() { var program=AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar); var data=OperatingSystem.IsMacOS()?Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal),"Library","Application Support","Expedient AI Bridges"):Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"ExpedientAIBridges");var runtime=Path.Combine(data,"runtime");var logs=Path.Combine(data,"logs");return new(program,data,runtime,logs,Path.Combine(data,"config.env"),Path.Combine(runtime,"supervisor.json"),Path.Combine(runtime,"supervisor.lock"),Path.Combine(runtime,"stop.request"),Path.Combine(logs,"companion.log"),Path.Combine(logs,"codex.log"),Path.Combine(logs,"anthropic.log")); }
    public void Create(){Directory.CreateDirectory(Data);Directory.CreateDirectory(Runtime);Directory.CreateDirectory(Logs);}
}
internal static class ConfigValidator
{
    public static int[] Ports(IReadOnlyDictionary<string,string> c)=>new[]{EnvFile.Integer(c,"PORT",4001),EnvFile.Integer(c,"ANTHROPIC_PORT",4002)};
    public static void Validate(IReadOnlyDictionary<string,string> c){if(!c.TryGetValue("OPENAI_BASE_URL",out var u)||!Uri.TryCreate(u,UriKind.Absolute,out var uri)||uri.Scheme is not ("http" or "https"))throw new InvalidOperationException("OPENAI_BASE_URL must be an absolute HTTP(S) URL.");var h=c.GetValueOrDefault("BRIDGE_HOST","127.0.0.1");if(h is not ("127.0.0.1" or "::1" or "localhost"))throw new InvalidOperationException("Production companion requires a loopback BRIDGE_HOST.");var p=Ports(c);if(p.Any(x=>x<1||x>65535)||p[0]==p[1])throw new InvalidOperationException("Bridge ports must be distinct values from 1 through 65535.");}
}
internal static class EnvFile
{
    public static Dictionary<string,string> Load(string path,bool required=true){if(!File.Exists(path)){if(required)throw new FileNotFoundException("Configuration missing",path);return new();}var result=new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);var number=0;foreach(var raw in File.ReadLines(path)){number++;var line=raw.Trim().TrimStart('\uFEFF');if(line.Length==0||line.StartsWith('#'))continue;var i=line.IndexOf('=');if(i<1)throw new FormatException($"Invalid config syntax at line {number}.");var key=line[..i].Trim();if(!key.All(c=>char.IsLetterOrDigit(c)||c=='_')||!(char.IsLetter(key[0])||key[0]=='_'))throw new FormatException($"Invalid config key at line {number}.");result[key]=line[(i+1)..].Trim();}return result;}
    public static int Integer(IReadOnlyDictionary<string,string> e,string k,int d)=>e.TryGetValue(k,out var v)&&int.TryParse(v,out var n)?n:d;
}
internal static class RuntimeLocator
{
    public static string FindNode(string program){var bundled=Path.Combine(program,"runtime",OperatingSystem.IsWindows()?"node.exe":"node");if(File.Exists(bundled))return bundled;var name=OperatingSystem.IsWindows()?"node.exe":"node";foreach(var d in (Environment.GetEnvironmentVariable("PATH")??"").Split(Path.PathSeparator)){var c=Path.Combine(d.Trim('"'),name);if(File.Exists(c))return c;}throw new FileNotFoundException("Node.js 22+ not found in runtime or PATH.");}
    public static void RequireNode22(string node){using var p=Process.Start(new ProcessStartInfo(node,"--version"){RedirectStandardOutput=true,UseShellExecute=false,CreateNoWindow=true})!;var s=p.StandardOutput.ReadToEnd().Trim().TrimStart('v');p.WaitForExit();if(!Version.TryParse(s,out var v)||v.Major<22)throw new InvalidOperationException($"Node.js 22+ required; found {s}.");}
}
internal sealed record ProcessIdentity(string Name,int Pid,long StartedUtcTicks,string Executable,string? Script)
{
    public static ProcessIdentity From(Process p,string n,string executable,string? script)=>new(n,p.Id,p.StartTime.ToUniversalTime().Ticks,Path.GetFullPath(executable),script is null?null:Path.GetFullPath(script));
    public bool IsSameProcess(){try{using var p=Process.GetProcessById(Pid);if(Math.Abs(p.StartTime.ToUniversalTime().Ticks-StartedUtcTicks)>=TimeSpan.FromSeconds(2).Ticks)return false;var actual=p.MainModule?.FileName;if(string.IsNullOrWhiteSpace(actual))return true;return string.Equals(Path.GetFullPath(actual),Executable,OperatingSystem.IsWindows()?StringComparison.OrdinalIgnoreCase:StringComparison.Ordinal);}catch{return false;}}
}
internal sealed record RuntimeState(string Generation,ProcessIdentity Supervisor,ProcessIdentity? Codex,ProcessIdentity? Anthropic,bool Ready,string? Failure,DateTimeOffset StartedAt)
{
    public bool IsSupervisorAlive()=>Supervisor.IsSameProcess();public bool AllAlive()=>IsSupervisorAlive()&&Codex?.IsSameProcess()==true&&Anthropic?.IsSameProcess()==true;
    public static async Task<RuntimeState?> ReadAsync(string p){if(!File.Exists(p))return null;try{return JsonSerializer.Deserialize<RuntimeState>(await File.ReadAllTextAsync(p));}catch{return null;}}
    public static async Task WriteAtomicAsync(string p,RuntimeState s){var t=p+"."+Guid.NewGuid().ToString("N")+".tmp";await File.WriteAllTextAsync(t,JsonSerializer.Serialize(s,new JsonSerializerOptions{WriteIndented=true}));File.Move(t,p,true);}
    public static void DeleteStale(CompanionPaths p){foreach(var f in new[]{p.State,p.StopRequest})try{if(File.Exists(f))File.Delete(f);}catch{}}
}
internal sealed class SingleInstanceLease:FileStream{private SingleInstanceLease(string p):base(p,FileMode.OpenOrCreate,FileAccess.ReadWrite,FileShare.None){}public static SingleInstanceLease? TryAcquire(string p){try{return new(p);}catch(IOException){return null;}}}
internal static class HealthProbe
{
    private static readonly HttpClient Client=new(){Timeout=TimeSpan.FromMilliseconds(750)};
    public static async Task<bool> HttpRespondingAsync(int p){try{using var r=await Client.GetAsync($"http://127.0.0.1:{p}/__expedient_readiness__");return true;}catch{return false;}}
}
internal static class SecretRedactor
{
    public static bool IsSecretKey(string k)=>k.Contains("KEY",StringComparison.OrdinalIgnoreCase)||k.Contains("TOKEN",StringComparison.OrdinalIgnoreCase)||k.Contains("SECRET",StringComparison.OrdinalIgnoreCase)||k.Contains("AUTH",StringComparison.OrdinalIgnoreCase);
    public static string Redact(string s,IEnumerable<string> values){foreach(var v in values)if(v.Length>=4)s=s.Replace(v,"[REDACTED]",StringComparison.Ordinal);return s;}
}
internal static class LogFiles
{
    public static void Rotate(string p){try{if(new FileInfo(p).Length<10*1024*1024)return;for(var i=4;i>=1;i--){var from=$"{p}.{i}";var to=$"{p}.{i+1}";if(File.Exists(from))File.Move(from,to,true);}File.Move(p,p+".1",true);}catch{}}
    public static async Task PumpAsync(StreamReader r,string p,string[] secrets){await using var f=new FileStream(p,FileMode.Append,FileAccess.Write,FileShare.ReadWrite);await using var w=new StreamWriter(f){AutoFlush=true};while(await r.ReadLineAsync() is{} line)await w.WriteLineAsync($"{DateTimeOffset.UtcNow:o} {SecretRedactor.Redact(line,secrets)}");}
    public static async Task AppendAsync(string p,string? s){Rotate(p);await File.AppendAllTextAsync(p,$"{DateTimeOffset.UtcNow:o} {s}{Environment.NewLine}");}
}
internal static class Migration
{
    public static async Task<int> RunAsync(CompanionPaths p,bool quiet=false){var oldPid=Path.Combine(p.Runtime,"native-pids.json");var legacy=Path.Combine(p.Runtime,"bridge-pids.json");foreach(var f in new[]{oldPid,legacy})if(File.Exists(f)){var backup=f+".legacy";if(!File.Exists(backup))File.Move(f,backup);}var marker=Path.Combine(p.Runtime,"migration-v2.json");if(!File.Exists(marker))await File.WriteAllTextAsync(marker,JsonSerializer.Serialize(new{schema=2,migratedAt=DateTimeOffset.UtcNow,owner="ExpedientBridges supervisor",rollback="stop companion; restore *.legacy PID records only after verifying no listener owns ports 4001/4002"},new JsonSerializerOptions{WriteIndented=true}));if(!quiet)Console.WriteLine("Mutable state migrated to schema 2; configuration preserved.");return 0;}
}
