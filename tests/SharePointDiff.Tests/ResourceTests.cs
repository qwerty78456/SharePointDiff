using System.Text.RegularExpressions;
using System.Xml.Linq;
using SharePointDiff.Core;
using Xunit;

namespace SharePointDiff.Tests;

public sealed class ResourceTests
{
    private static string RepoRoot()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null && !File.Exists(Path.Combine(dir.FullName, "Directory.Build.props"))) dir = dir.Parent;
        Assert.NotNull(dir); return dir!.FullName;
    }
    private static Dictionary<string, string> Resx(string name) =>
        XDocument.Load(Path.Combine(RepoRoot(), "src", "SharePointDiff.App", "Resources", name)).Root!.Elements("data")
            .ToDictionary(x => (string)x.Attribute("name")!, x => x.Element("value")!.Value);
    private static IEnumerable<(string File, string Line)> Lines(string project, string pattern) =>
        Directory.EnumerateFiles(Path.Combine(RepoRoot(), "src", project), pattern, SearchOption.AllDirectories)
            .Where(p => !p.Contains(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar) && !p.Contains(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar))
            .SelectMany(p => File.ReadLines(p).Select(line => (Path.GetFileName(p), line)));
    private static string Tokens(string s) => string.Join("|", Regex.Matches(s, @"\{\d+[^}]*\}").Select(m => m.Value));

    /// <summary>Every code that can reach the UI as "Error" + Code has a string in all three resource files.</summary>
    [Fact]
    public void EveryErrorCodeHasLocalizedString()
    {
        var direct = new[]
        {
            @"OperationException\(\s*""(\w+)""", @"(?:issues|Issues)\.Add\(new\(""(\w+)""", @"(?:Issue|problem)\s*\??=\s*new\(""(\w+)""",
            @"\?\s*new\(""(\w+)""", @"new OperationIssue\(""(\w+)""", @"FailAsync\(\w+, \w+, ""(\w+)""\)"
        };
        var codes = new SortedSet<string>(StringComparer.Ordinal);
        foreach (var (file, line) in Lines("SharePointDiff.Core", "*.cs").Concat(Lines("SharePointDiff.App", "*.cs")))
        {
            foreach (string pattern in direct) foreach (Match m in Regex.Matches(line, pattern)) codes.Add(m.Groups[1].Value);
            // Codes chosen by a ternary: look only at the expression that follows op.Code / "?.Code ??" / OperationException(status / Code: / ShowProblem(.
            foreach (string anchor in new[] { "op.Code", "?.Code ??", "OperationException(status", "Code: \"", "ShowProblem(" })
                for (int at = line.IndexOf(anchor, StringComparison.Ordinal); at >= 0; at = line.IndexOf(anchor, at + anchor.Length, StringComparison.Ordinal))
                {
                    string tail = line[at..]; int stop = tail.IndexOfAny([';', '}']); if (stop >= 0) tail = tail[..stop];
                    foreach (Match m in Regex.Matches(tail, @"""([A-Z][A-Za-z]+)""")) codes.Add(m.Groups[1].Value);
                }
        }
        Assert.Contains("WorkspaceBusy", codes); Assert.Contains("FileIoError", codes); Assert.Contains("RequestTimeout", codes); Assert.Contains("UnscannedLibrary", codes); Assert.Contains("WriteScopedApp", codes); Assert.Contains("TransferFailed", codes);
        Assert.DoesNotContain("AgentCheck", codes); Assert.DoesNotContain("TrayProblem", codes);
        foreach (string name in new[] { "Strings.resx", "Strings.en-US.resx", "Strings.vi-VN.resx" })
        {
            var keys = Resx(name);
            var missing = codes.Where(c => !keys.ContainsKey("Error" + c)).ToArray();
            Assert.True(missing.Length == 0, name + " lacks: " + string.Join(", ", missing.Select(c => "Error" + c)));
        }
    }
    [Fact]
    public void NeutralResxHasSameKeysAndTokensAsEnUs()
    {
        var neutral = Resx("Strings.resx"); var en = Resx("Strings.en-US.resx");
        Assert.Equal(en.Keys.Order(StringComparer.Ordinal), neutral.Keys.Order(StringComparer.Ordinal));
        foreach (var (key, value) in en) Assert.True(Tokens(value) == Tokens(neutral[key]), key);
    }
    /// <summary>Every key the App asks for literally exists; dynamic "Status" + DiffStatus keys and progress codes are covered too.</summary>
    [Fact]
    public void EveryReferencedResourceKeyExists()
    {
        var used = new SortedSet<string>(StringComparer.Ordinal);
        var csPatterns = new[] { @"\b[Ll]\[""(\w+)""\]", @"\bL\.Format\(""(\w+)""", @"SetStatus\(""(\w+)""", @"statusCode = ""(\w+)""", @"LocalizationService\(\)\[""(\w+)""\]", @"\bAction\(""(\w+)""" };
        foreach (var (_, line) in Lines("SharePointDiff.App", "*.cs")) foreach (string pattern in csPatterns) foreach (Match m in Regex.Matches(line, pattern)) used.Add(m.Groups[1].Value);
        // Keys chosen by a ternary inside SetStatus(...) / L.Format(...), e.g. SetStatus(cancelled ? "Cancelled" : "ApplyComplete").
        foreach (var (_, line) in Lines("SharePointDiff.App", "*.cs"))
            foreach (Match call in Regex.Matches(line, @"(?:SetStatus|L\.Format)\(([^;]*?\?[^;]*?)\)"))
                foreach (Match m in Regex.Matches(call.Groups[1].Value, @"\?\s*""(\w+)""\s*:\s*""(\w+)""")) { used.Add(m.Groups[1].Value); used.Add(m.Groups[2].Value); }
        foreach (var (_, line) in Lines("SharePointDiff.App", "*.xaml")) foreach (Match m in Regex.Matches(line, @"\bL\[(\w+)\]")) used.Add(m.Groups[1].Value);
        // Progress codes from Core are shown through SetStatus(p.Code).
        foreach (var (_, line) in Lines("SharePointDiff.Core", "*.cs")) foreach (Match m in Regex.Matches(line, @"Report\(new\(""(\w+)""")) used.Add(m.Groups[1].Value);
        foreach (string status in Enum.GetNames<DiffStatus>()) used.Add("Status" + status);
        Assert.Contains("Compare", used); Assert.Contains("Hashing", used); Assert.Contains("FilterLocallyChanged", used); Assert.Contains("ApplyComplete", used); Assert.Contains("CardSync", used); Assert.Contains("TrayPending", used);
        foreach (string name in new[] { "Strings.resx", "Strings.en-US.resx", "Strings.vi-VN.resx" })
        {
            var keys = Resx(name); var missing = used.Where(k => !keys.ContainsKey(k)).ToArray();
            Assert.True(missing.Length == 0, name + " lacks: " + string.Join(", ", missing));
        }
    }
}
