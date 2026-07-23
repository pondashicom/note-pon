using System.Diagnostics;
using System.IO;
using System.Text.Json;

namespace NotePon;

internal static class PowerPointWorkerHost
{
    private const string WorkerArgument = "--powerpoint-worker";
    private const string ParentProcessIdArgument = "--parent-pid";

    public static bool TryGetParentProcessId(string[] args, out int parentProcessId)
    {
        parentProcessId = 0;
        if (args.Length != 3 || !string.Equals(args[0], WorkerArgument, StringComparison.Ordinal))
        {
            return false;
        }

        return string.Equals(args[1], ParentProcessIdArgument, StringComparison.Ordinal)
            && int.TryParse(args[2], out parentProcessId)
            && parentProcessId > 0;
    }

    public static int Run(int parentProcessId)
    {
        StartParentExitMonitor(parentProcessId);
        var reader = new PowerPointReader();

        try
        {
            string? command;
            while ((command = Console.ReadLine()) is not null)
            {
                if (!string.Equals(command, "poll", StringComparison.Ordinal))
                {
                    continue;
                }

                PowerPointSnapshot snapshot = reader.Poll();
                Console.Out.WriteLine(JsonSerializer.Serialize(snapshot));
                Console.Out.Flush();
            }

            return 0;
        }
        catch (Exception exception)
        {
            AppLog.Write("The PowerPoint worker terminated unexpectedly.", exception);
            return 1;
        }
    }

    private static void StartParentExitMonitor(int parentProcessId)
    {
        var monitor = new Thread(() =>
        {
            try
            {
                using Process parent = Process.GetProcessById(parentProcessId);
                parent.WaitForExit();
            }
            catch
            {
                // If the parent cannot be opened, it is no longer safe to keep this worker alive.
            }

            Environment.Exit(0);
        })
        {
            IsBackground = true,
            Name = "NOTE-PON parent monitor"
        };

        monitor.Start();
    }

    public static ProcessStartInfo CreateStartInfo(int parentProcessId)
    {
        string executablePath = Environment.ProcessPath
            ?? throw new InvalidOperationException("The NOTE-PON executable path could not be resolved.");

        var startInfo = new ProcessStartInfo
        {
            FileName = executablePath,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            WindowStyle = ProcessWindowStyle.Hidden
        };
        startInfo.ArgumentList.Add(WorkerArgument);
        startInfo.ArgumentList.Add(ParentProcessIdArgument);
        startInfo.ArgumentList.Add(parentProcessId.ToString());
        return startInfo;
    }
}

internal sealed class PowerPointWorkerClient : IDisposable
{
    private static readonly TimeSpan PollTimeout = TimeSpan.FromSeconds(2);
    private static readonly TimeSpan[] RestartBackoff =
    [
        TimeSpan.FromSeconds(2),
        TimeSpan.FromSeconds(5),
        TimeSpan.FromSeconds(10),
        TimeSpan.FromSeconds(30)
    ];
    private readonly SemaphoreSlim _pollGate = new(1, 1);

    private Process? _worker;
    private StreamWriter? _input;
    private StreamReader? _output;
    private bool _hasConnected;
    private bool _disposed;
    private int _consecutiveWorkerFailures;
    private long _nextWorkerStartAt;

    public async Task<PowerPointSnapshot> PollAsync(CancellationToken cancellationToken)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        await _pollGate.WaitAsync(cancellationToken);
        Task<string?>? responseTask = null;

        try
        {
            if (Environment.TickCount64 < _nextWorkerStartAt)
            {
                return FailureSnapshot("PowerPoint 監視の再起動を待っています");
            }

            EnsureWorker();
            await _input!.WriteLineAsync("poll");
            await _input.FlushAsync(cancellationToken);

            responseTask = _output!.ReadLineAsync(cancellationToken).AsTask();
            string response = await responseTask.WaitAsync(PollTimeout, cancellationToken)
                ?? throw new EndOfStreamException("The PowerPoint worker closed its response stream.");

            PowerPointSnapshot snapshot = JsonSerializer.Deserialize<PowerPointSnapshot>(response)
                ?? throw new InvalidDataException("The PowerPoint worker returned an empty response.");
            ResetRestartBackoff();
            if (snapshot.State == PowerPointState.Connected)
            {
                _hasConnected = true;
            }

            return snapshot;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            throw;
        }
        catch (TimeoutException exception)
        {
            AppLog.WriteThrottled(
                "worker-timeout",
                "The PowerPoint worker did not answer within two seconds and will be restarted.",
                exception);
            StopWorker();
            await ObserveCompletionAsync(responseTask);
            ScheduleWorkerRestart();
            return FailureSnapshot("PowerPoint の応答がタイムアウトしました。監視を再起動しています");
        }
        catch (Exception exception)
        {
            AppLog.WriteThrottled(
                "worker-failure",
                "The PowerPoint worker failed and will be restarted.",
                exception);
            StopWorker();
            ScheduleWorkerRestart();
            return FailureSnapshot("PowerPoint 監視を再起動しています");
        }
        finally
        {
            _pollGate.Release();
        }
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        StopWorker();
    }

    private void EnsureWorker()
    {
        if (_worker is not null && !_worker.HasExited)
        {
            return;
        }

        StopWorker();
        _worker = Process.Start(PowerPointWorkerHost.CreateStartInfo(Environment.ProcessId))
            ?? throw new InvalidOperationException("The PowerPoint worker could not be started.");
        _input = _worker.StandardInput;
        _output = _worker.StandardOutput;
    }

    private PowerPointSnapshot FailureSnapshot(string statusText) =>
        new(
            _hasConnected ? PowerPointState.Reconnecting : PowerPointState.WaitingForPowerPoint,
            statusText);

    private void ScheduleWorkerRestart()
    {
        int delayIndex = Math.Min(_consecutiveWorkerFailures, RestartBackoff.Length - 1);
        _consecutiveWorkerFailures++;
        _nextWorkerStartAt =
            Environment.TickCount64 + (long)RestartBackoff[delayIndex].TotalMilliseconds;
    }

    private void ResetRestartBackoff()
    {
        _consecutiveWorkerFailures = 0;
        _nextWorkerStartAt = 0;
    }

    private static async Task ObserveCompletionAsync(Task<string?>? responseTask)
    {
        if (responseTask is null)
        {
            return;
        }

        try
        {
            await responseTask;
        }
        catch
        {
            // The worker pipe is expected to fail after a timeout-triggered restart.
        }
    }

    private void StopWorker()
    {
        DisposePipe(_input, "input");
        DisposePipe(_output, "output");
        _input = null;
        _output = null;

        if (_worker is null)
        {
            return;
        }

        try
        {
            if (!_worker.HasExited)
            {
                _worker.Kill(entireProcessTree: true);
                _worker.WaitForExit(milliseconds: 1000);
            }
        }
        catch (Exception exception)
        {
            AppLog.Write("The PowerPoint worker could not be stopped cleanly.", exception);
        }
        finally
        {
            try
            {
                _worker.Dispose();
            }
            catch (Exception exception)
            {
                AppLog.Write("The PowerPoint worker process handle could not be disposed.", exception);
            }

            _worker = null;
        }
    }

    private static void DisposePipe(IDisposable? pipe, string pipeName)
    {
        try
        {
            pipe?.Dispose();
        }
        catch (Exception exception)
        {
            AppLog.Write($"The PowerPoint worker {pipeName} pipe could not be disposed.", exception);
        }
    }
}
