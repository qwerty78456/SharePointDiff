using System.Collections.Concurrent;

namespace SharePointDiff.Core;

/// <summary>Logical application I/O, not physical SSD traffic or TLS wire bytes.</summary>
public sealed record IoStatistics(IReadOnlyDictionary<string, long> Counters);

public sealed class IoMetrics : IDisposable
{
    private static readonly AsyncLocal<IoMetrics?> Active = new();
    private static readonly AsyncLocal<string?> Phase = new();
    private readonly ConcurrentDictionary<string, long> counters = new(StringComparer.Ordinal);
    private readonly IoMetrics? previous;
    private readonly string? previousPhase;
    private IoMetrics(string phase) { previous = Active.Value; previousPhase = Phase.Value; Active.Value = this; Phase.Value = phase; }
    public static IoMetrics Begin(string phase) => new(phase);
    public IoStatistics Snapshot() => new(new SortedDictionary<string, long>(counters, StringComparer.Ordinal));
    public static void Add(string counter, long amount = 1)
    {
        if (Active.Value is not { } current) return;
        current.counters.AddOrUpdate(counter, amount, (_, old) => old + amount);
        if (Phase.Value is { } phase) current.counters.AddOrUpdate(phase + "." + counter, amount, (_, old) => old + amount);
    }
    public static IDisposable Stage(string phase) => new StageScope(phase);
    private sealed class StageScope : IDisposable
    {
        private readonly string? previous = Phase.Value;
        public StageScope(string phase) => Phase.Value = phase;
        public void Dispose() => Phase.Value = previous;
    }
    public void Dispose() { Active.Value = previous; Phase.Value = previousPhase; }
}
