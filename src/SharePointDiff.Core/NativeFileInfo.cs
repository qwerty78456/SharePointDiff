using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using System.Text;

namespace SharePointDiff.Core;

public sealed record FileIdentity(ulong VolumeSerial, string FileId, long Size, long CreationTime, long LastWriteTime, long ChangeTime);

internal static class NativeFileInfo
{
    [StructLayout(LayoutKind.Sequential)] private struct BasicInfo { public long Creation, Access, Write, Change; public uint Attributes; }
    [StructLayout(LayoutKind.Sequential)] private struct IdInfo { public ulong Volume, Low, High; }
    [StructLayout(LayoutKind.Sequential)] private struct StandardInfo { public long Allocation, Length; public uint Links; public byte DeletePending, Directory; }
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GetFileInformationByHandleEx(SafeFileHandle file, int kind, out BasicInfo info, uint size);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GetFileInformationByHandleEx(SafeFileHandle file, int kind, out IdInfo info, uint size);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GetFileInformationByHandleEx(SafeFileHandle file, int kind, out StandardInfo info, uint size);
    public static FileIdentity? TryRead(FileStream file)
    {
        if (!OperatingSystem.IsWindows()) return null;
        if (!GetFileInformationByHandleEx(file.SafeFileHandle, 0, out BasicInfo basic, (uint)Marshal.SizeOf<BasicInfo>()) ||
            !GetFileInformationByHandleEx(file.SafeFileHandle, 18, out IdInfo id, (uint)Marshal.SizeOf<IdInfo>()) ||
            !GetFileInformationByHandleEx(file.SafeFileHandle, 1, out StandardInfo standard, (uint)Marshal.SizeOf<StandardInfo>())) return null;
        IoMetrics.Add("Identity.Reads");
        return new(id.Volume, id.Low.ToString("x16") + id.High.ToString("x16"), standard.Length, basic.Creation, basic.Write, basic.Change);
    }
    public static bool SameObject(FileIdentity? a, FileIdentity? b) => a != null && b != null && a.VolumeSerial == b.VolumeSerial && a.FileId == b.FileId;
    public static string? FinalPath(FileStream file)
    {
        if (!OperatingSystem.IsWindows()) return null;
        var text = new StringBuilder(512);
        uint size = GetFinalPathNameByHandleW(file.SafeFileHandle, text, (uint)text.Capacity, 0);
        if (size == 0) return null;
        if (size >= text.Capacity)
        {
            if (size > 32768) return null;
            text.EnsureCapacity((int)size + 1);
            if (GetFinalPathNameByHandleW(file.SafeFileHandle, text, (uint)text.Capacity, 0) == 0) return null;
        }
        string path = text.ToString();
        return path.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase) ? @"\\" + path[8..] : path.StartsWith(@"\\?\", StringComparison.Ordinal) ? path[4..] : path;
    }
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern uint GetFinalPathNameByHandleW(SafeFileHandle file, StringBuilder path, uint size, uint flags);
    public static bool SupportsCache(string path)
    {
        if (!OperatingSystem.IsWindows()) return false;
        try { var drive = new DriveInfo(Path.GetPathRoot(path)!); return drive.DriveType == DriveType.Fixed && drive.DriveFormat == "NTFS"; }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException) { return false; }
    }
    public static int ManualHashParallelism(string root)
    {
        if (!SupportsCache(root)) return 1;
        using var volume = CreateFileW(@"\\.\" + Path.GetPathRoot(root)!.TrimEnd('\\'), 0, 7, IntPtr.Zero, 3, 0, IntPtr.Zero);
        if (volume.IsInvalid) return 1;
        byte[] query = new byte[12]; query[0] = 7; // StorageDeviceSeekPenaltyProperty, PropertyStandardQuery
        byte[] descriptor = new byte[12];
        return DeviceIoControl(volume, 0x002D1400, query, query.Length, descriptor, descriptor.Length, out int returned, IntPtr.Zero) &&
            returned >= 9 && descriptor[8] == 0 ? Math.Clamp(Environment.ProcessorCount / 2, 2, 4) : 1;
    }
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)] private static extern SafeFileHandle CreateFileW(string path, uint access, uint share, IntPtr security, uint mode, uint flags, IntPtr template);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool DeviceIoControl(SafeFileHandle device, uint code, byte[] input, int inputSize, byte[] output, int outputSize, out int returned, IntPtr overlapped);
}
