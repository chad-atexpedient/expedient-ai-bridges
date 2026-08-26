using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
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
            var app = CompanionPaths.Resolve();
            Directory.CreateDirectory(app.Data);
            Directory.CreateDirectory(app.Runtime);
            Directory.CreateDirectory(app.Logs);
            return command switch
            {
                "start" or "--background" => await StartAsync(app),
                "stop" or "--stop" => await StopAsync(app),
                "restart" => await RestartAsync(app),
                "status" or "--status" => await StatusAsync(app),
                "diagnostics" or "--diagnostics" => await DiagnosticsAsync(app),
                "--choose" => await ChooseAsync(app),
                "--uninstall" => UninstallInstructions(app),
                _ => Usage()
            };
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"{Product}: {ex.Message}");
            return 1;
        }
    }

    private static async Task<int> StartAsync(CompanionPaths paths)
    {
        var config = EnvFile.Load(paths.Config);
        var baseUrl = config.GetValueOrDefault("OPENAI_BASE_URL");
        if (string.IsNullOrWhiteSpace(baseUrl))
            throw new InvalidOperationException($"OPENAI_BASE_URL is required in {paths.Config}");
        var host = config.GetValueOrDefault("BRIDGE_HOST", "127.0.0.1");
        var allowRemote = config.GetValueOrDefault("ALLOW_NON_LOOPBACK", "0") == "1";
        if (!allowRemote && host is not ("127.0.0.1" or "::1" or "localhost"))
            throw new InvalidOperationException($"Refusing non-loopback BRIDGE_HOST={host} without ALLOW_NON_LOOPBACK=1.");
        var ports = new[] { EnvFile.Integer(config, "PORT", 4001), EnvFile.Integer(config, "ANTHROPIC_PORT", 4002) };
        var current = await ProcessState.ReadAsync(paths.Pids);
        if (current is not null && current.AllRunning())
        {
            Console.WriteLine("Bridges are already running.");
            return 0;
        }
        await StopAsync(paths, quiet: true);
        foreach (var port in ports)
            if (PortProbe.IsListening(port)) throw new InvalidOperationException($"Port {port} is already in use.");
        var node = RuntimeLocator.FindNode(paths.Program);
        RuntimeLocator.RequireNode22(node);
        var common = new ProcessStartInfo(node) { UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = paths.Program };
        foreach (var pair in config) common.Environment[pair.Key] = pair.Value;
        common.Environment["NODE_OPTIONS"] = "--max-old-space-size=8192";
        common.Environment["BRIDGE_METRICS_DIR"] = paths.Logs;
        var stdout = Path.Combine(paths.Logs, "bridges-output.log");
        var stderr = Path.Combine(paths.Logs, "bridges-error.log");
        var codex = StartNode(common, Path.Combine(paths.Program, "app", "index.js"), stdout, stderr);
        var anthropic = StartNode(common, Path.Combine(paths.Program, "app", "anthropic-bridge.js"), stdout, stderr);
        var state = new ProcessState(
            ProcessIdentity.From(codex, "codex", Path.Combine(paths.Program, "app", "index.js")),
            ProcessIdentity.From(anthropic, "anthropic", Path.Combine(paths.Program, "app", "anthropic-bridge.js")));
        await ProcessState.WriteAsync(paths.Pids, state);
        var deadline = DateTime.UtcNow.AddSeconds(12);
        while (DateTime.UtcNow < deadline && !ports.All(PortProbe.IsListening))
        {
            if (codex.HasExited || anthropic.HasExited) throw new InvalidOperationException($"A bridge exited during startup. See {stderr}");
            await Task.Delay(250);
        }
        if (!ports.All(PortProbe.IsListening))
        {
            await StopAsync(paths, quiet: true);
            throw new TimeoutException($"Bridge readiness timed out. See {stderr}");
        }
        Console.WriteLine($"Bridges ready on {host}:{ports[0]} and {host}:{ports[1]}.");
        return 0;
    }

    private static Process StartNode(ProcessStartInfo common, string script, string stdout, string stderr)
    {
        if (!File.Exists(script)) throw new FileNotFoundException("Bridge bundle not found", script);
        var info = new ProcessStartInfo(common.FileName)
        {
            UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = common.WorkingDirectory,
            RedirectStandardOutput = true, RedirectStandardError = true
        };
        foreach (var pair in common.Environment) info.Environment[pair.Key] = pair.Value;
        info.ArgumentList.Add(script);
        var process = Process.Start(info) ?? throw new InvalidOperationException($"Could not start {script}");
        _ = PumpAsync(process.StandardOutput, stdout);
        _ = PumpAsync(process.StandardError, stderr);
        return process;
    }

    private static async Task PumpAsync(StreamReader reader, string path)
    {
        await using var stream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
        await using var writer = new StreamWriter(stream) { AutoFlush = true };
        while (await reader.ReadLineAsync() is { } line) await writer.WriteLineAsync($"{DateTimeOffset.UtcNow:o} {line}");
    }

    private static async Task<int> StopAsync(CompanionPaths paths, bool quiet = false)
    {
        var state = await ProcessState.ReadAsync(paths.Pids);
        if (state is not null)
            foreach (var identity in state.Items)
                identity.TryStop();
        if (File.Exists(paths.Pids)) File.Delete(paths.Pids);
        if (!quiet) Console.WriteLine("Bridges stopped.");
        return 0;
    }

    private static async Task<int> RestartAsync(CompanionPaths paths) { await StopAsync(paths, true); return await StartAsync(paths); }
    private static async Task<int> StatusAsync(CompanionPaths paths)
    {
        var state = await ProcessState.ReadAsync(paths.Pids);
        var running = state?.AllRunning() == true;
        Console.WriteLine(running ? "running" : "stopped");
        return running ? 0 : 3;
    }
    private static async Task<int> DiagnosticsAsync(CompanionPaths paths)
    {
        var config = EnvFile.Load(paths.Config, required: false);
        var state = await ProcessState.ReadAsync(paths.Pids);
        var report = new
        {
            product = Product, os = Environment.OSVersion.ToString(), architecture = System.Runtime.InteropServices.RuntimeInformation.OSArchitecture.ToString(),
            program = paths.Program, data = paths.Data, configPresent = File.Exists(paths.Config), upstreamConfigured = config.ContainsKey("OPENAI_BASE_URL"),
            ports = new[] { EnvFile.Integer(config, "PORT", 4001), EnvFile.Integer(config, "ANTHROPIC_PORT", 4002) }.Select(p => new { port = p, listening = PortProbe.IsListening(p) }),
            processes = state?.Items.Where(x => x is not null).Select(x => new { x.Name, x.Pid, running = x.IsSameProcess() })
        };
        Console.WriteLine(JsonSerializer.Serialize(report, JsonOptions));
        return 0;
    }
    private static async Task<int> ChooseAsync(CompanionPaths paths)
    {
        Console.WriteLine("1. Start bridges\n2. Stop bridges\n3. Status\n4. Diagnostics");
        return (Console.ReadLine() ?? "") switch { "1" => await StartAsync(paths), "2" => await StopAsync(paths), "3" => await StatusAsync(paths), "4" => await DiagnosticsAsync(paths), _ => 2 };
    }
    private static int UninstallInstructions(CompanionPaths paths) { Console.WriteLine($"Remove the application, then optionally remove mutable data at: {paths.Data}"); return 0; }
    private static int Usage() { Console.WriteLine("Usage: ExpedientBridges [start|stop|restart|status|diagnostics|--background|--choose]"); return 2; }
}

internal sealed record CompanionPaths(string Program, string Data, string Runtime, string Logs, string Config, string Pids)
{
    public static CompanionPaths Resolve()
    {
        var program = AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar);
        var data = OperatingSystem.IsMacOS()
            ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "Library", "Application Support", "Expedient AI Bridges")
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExpedientAIBridges");
        return new(program, data, Path.Combine(data, "runtime"), Path.Combine(data, "logs"), Path.Combine(data, "config.env"), Path.Combine(data, "runtime", "native-pids.json"));
    }
}

internal static class EnvFile
{
    public static Dictionary<string, string> Load(string path, bool required = true)
    {
        if (!File.Exists(path)) { if (required) throw new FileNotFoundException("Configuration missing", path); return new(); }
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var raw in File.ReadLines(path))
        {
            var line = raw.Trim(); if (line.Length == 0 || line.StartsWith('#')) continue;
            var equals = line.IndexOf('='); if (equals < 1) continue;
            var key = line[..equals].Trim(); if (!key.All(c => char.IsLetterOrDigit(c) || c == '_') || !(char.IsLetter(key[0]) || key[0] == '_')) continue;
            result[key] = line[(equals + 1)..].Trim();
        }
        return result;
    }
    public static int Integer(IReadOnlyDictionary<string,string> env, string key, int fallback) => env.TryGetValue(key, out var value) && int.TryParse(value, out var number) ? number : fallback;
}

internal static class RuntimeLocator
{
    public static string FindNode(string program)
    {
        var bundled = Path.Combine(program, "runtime", OperatingSystem.IsWindows() ? "node.exe" : "node");
        if (File.Exists(bundled)) return bundled;
        var pathName = OperatingSystem.IsWindows() ? "node.exe" : "node";
        foreach (var directory in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(Path.PathSeparator))
        { var candidate = Path.Combine(directory.Trim('"'), pathName); if (File.Exists(candidate)) return candidate; }
        throw new FileNotFoundException("Node.js 22+ was not found in the application bundle or PATH.");
    }
    public static void RequireNode22(string node)
    {
        using var p = Process.Start(new ProcessStartInfo(node, "--version") { RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true })!;
        var value = p.StandardOutput.ReadToEnd().Trim().TrimStart('v'); p.WaitForExit();
        if (!Version.TryParse(value, out var version) || version.Major < 22) throw new InvalidOperationException($"Node.js 22+ required; found {value}.");
    }
}

internal sealed record ProcessIdentity(string Name, int Pid, long StartedUtcTicks, string Script)
{
    public static ProcessIdentity From(Process p, string name, string script) => new(name, p.Id, p.StartTime.ToUniversalTime().Ticks, script);
    public bool IsSameProcess()
    {
        try { using var p = Process.GetProcessById(Pid); return Math.Abs(p.StartTime.ToUniversalTime().Ticks - StartedUtcTicks) < TimeSpan.FromSeconds(2).Ticks; }
        catch { return false; }
    }
    public void TryStop() { if (!IsSameProcess()) return; try { using var p = Process.GetProcessById(Pid); p.Kill(true); p.WaitForExit(5000); } catch { } }
}

internal sealed record ProcessState(ProcessIdentity Codex, ProcessIdentity Anthropic)
{
    public IEnumerable<ProcessIdentity> Items => new[] { Codex, Anthropic }.Where(x => x is not null)!;
    public bool AllRunning() => Items.Any() && Items.All(x => x.IsSameProcess());
    public static async Task<ProcessState?> ReadAsync(string path) { if (!File.Exists(path)) return null; try { return JsonSerializer.Deserialize<ProcessState>(await File.ReadAllTextAsync(path)); } catch { return null; } }
    public static Task WriteAsync(string path, ProcessState state) => File.WriteAllTextAsync(path, JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true }));
}

internal static class PortProbe
{
    public static bool IsListening(int port)
    {
        try { using var client = new TcpClient(); return client.ConnectAsync(IPAddress.Loopback, port).Wait(TimeSpan.FromMilliseconds(200)) && client.Connected; }
        catch { return false; }
    }
}
