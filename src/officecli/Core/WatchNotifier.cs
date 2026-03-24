// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.IO.Pipes;
using System.Text;
using System.Text.Json;

namespace OfficeCli.Core;

/// <summary>
/// Sends refresh notifications (with rendered HTML) to a running watch process.
/// Non-blocking, fire-and-forget. Silently does nothing if no watch is running.
/// All pipe I/O is bounded by a timeout to prevent hangs.
/// </summary>
public static class WatchNotifier
{
    private static readonly TimeSpan PipeTimeout = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Notify watch with a pre-built message.
    /// The watch server never opens the file — all rendering is done by the caller.
    /// </summary>
    public static void NotifyIfWatching(string filePath, WatchMessage message)
    {
        try
        {
            RunWithTimeout(() =>
            {
                var pipeName = WatchServer.GetWatchPipeName(filePath);
                using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
                client.Connect(100); // fast fail if no watch

                using var writer = new StreamWriter(client, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };
                using var reader = new StreamReader(client, Encoding.UTF8, leaveOpen: true);

                var json = JsonSerializer.Serialize(message, WatchMessageJsonContext.Default.WatchMessage);
                writer.WriteLine(json);
                reader.ReadLine(); // wait for ack
            }, PipeTimeout);
        }
        catch
        {
            // No watch process running, or timed out — silently ignore
        }
    }

    /// <summary>
    /// Send a close command to a running watch process.
    /// Returns true if the watch was successfully closed.
    /// </summary>
    public static bool SendClose(string filePath)
    {
        try
        {
            bool result = false;
            RunWithTimeout(() =>
            {
                var pipeName = WatchServer.GetWatchPipeName(filePath);
                using var client = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
                client.Connect(200);

                using var writer = new StreamWriter(client, Encoding.UTF8, leaveOpen: true) { AutoFlush = true };
                using var reader = new StreamReader(client, Encoding.UTF8, leaveOpen: true);

                writer.WriteLine("close");
                reader.ReadLine();
                result = true;
            }, PipeTimeout);
            return result;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Run an action on a background thread with a timeout.
    /// Prevents the calling thread from hanging if the pipe server dies mid-conversation.
    /// </summary>
    private static void RunWithTimeout(Action action, TimeSpan timeout)
    {
        var task = Task.Run(action);
        if (!task.Wait(timeout))
            throw new TimeoutException("Pipe communication timed out");
        task.GetAwaiter().GetResult(); // propagate exceptions
    }
}

/// <summary>
/// Message sent from command processes to the watch server via named pipe.
/// </summary>
public class WatchMessage
{
    /// <summary>"replace", "add", "remove", or "full"</summary>
    public string Action { get; set; } = "full";

    /// <summary>Slide number (0 for full refresh)</summary>
    public int Slide { get; set; }

    /// <summary>Single slide HTML fragment (for replace/add)</summary>
    public string? Html { get; set; }

    /// <summary>Full HTML of the entire presentation (for caching by watch server)</summary>
    public string? FullHtml { get; set; }

    public static int ExtractSlideNum(string? path)
    {
        if (string.IsNullOrEmpty(path)) return 0;
        var match = System.Text.RegularExpressions.Regex.Match(path, @"/slide\[(\d+)\]");
        if (match.Success && int.TryParse(match.Groups[1].Value, out var num))
            return num;
        return 0;
    }
}

[System.Text.Json.Serialization.JsonSerializable(typeof(WatchMessage))]
internal partial class WatchMessageJsonContext : System.Text.Json.Serialization.JsonSerializerContext { }
