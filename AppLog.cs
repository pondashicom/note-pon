using System.IO;
using System.Text;

namespace NotePon;

internal static class AppLog
{
    private const long MaximumLogBytes = 1024 * 1024;
    private const int MaximumExceptionCharacters = 16 * 1024;
    private static readonly object SyncRoot = new();
    private static readonly Dictionary<string, long> LastThrottledWrites = new(StringComparer.Ordinal);

    public static void Write(string message, Exception? exception = null)
    {
        try
        {
            lock (SyncRoot)
            {
                string directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "NOTE-PON");
                Directory.CreateDirectory(directory);

                string logPath = Path.Combine(directory, "note-pon.log");
                RotateIfNeeded(logPath);

                string exceptionText = exception?.ToString() ?? string.Empty;
                if (exceptionText.Length > MaximumExceptionCharacters)
                {
                    exceptionText = exceptionText[..MaximumExceptionCharacters] + " [truncated]";
                }

                var entry = new StringBuilder()
                    .Append(DateTimeOffset.Now.ToString("O"))
                    .Append(" [PID ")
                    .Append(Environment.ProcessId)
                    .Append("] ")
                    .AppendLine(message);

                if (exceptionText.Length > 0)
                {
                    entry.AppendLine(exceptionText);
                }

                File.AppendAllText(logPath, entry.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            }
        }
        catch
        {
            // Diagnostics must never become another failure source.
        }
    }

    public static void WriteThrottled(string key, string message, Exception exception)
    {
        lock (SyncRoot)
        {
            long now = Environment.TickCount64;
            if (LastThrottledWrites.TryGetValue(key, out long previous) && now - previous < 10_000)
            {
                return;
            }

            LastThrottledWrites[key] = now;
        }

        Write(message, exception);
    }

    private static void RotateIfNeeded(string logPath)
    {
        var logFile = new FileInfo(logPath);
        if (!logFile.Exists || logFile.Length < MaximumLogBytes)
        {
            return;
        }

        string previousLogPath = logPath + ".old";
        File.Move(logPath, previousLogPath, overwrite: true);
    }
}
