namespace SharePointDiff.Core;

internal sealed record ReuseCandidate(string Path, string? Sha256);

internal static class ContentGrouping
{
    public static string Key(RemoteFile file) => file.QuickXorHash is { Length: > 0 } hash
        ? "hash:" + file.Size.ToString(System.Globalization.CultureInfo.InvariantCulture) + ":" + hash
        : "item:" + file.DriveId + ":" + file.ItemId + ":" + file.ETag;
}
