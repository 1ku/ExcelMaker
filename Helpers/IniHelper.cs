using System.Runtime.InteropServices;

namespace ExcelMaker.Helpers;

/// <summary>
/// INI 文件解析工具（使用 Windows API）
/// </summary>
public static class IniHelper
{
    [DllImport("kernel32", CharSet = CharSet.Unicode)]
    private static extern int GetPrivateProfileString(string section, string key, string def,
        StringBuilder retVal, int size, string filePath);

    [DllImport("kernel32", CharSet = CharSet.Unicode)]
    private static extern int GetPrivateProfileSection(string section, IntPtr lpReturnedString,
        int nSize, string filePath);

    [DllImport("kernel32", CharSet = CharSet.Unicode)]
    private static extern bool WritePrivateProfileString(string section, string key,
        string value, string filePath);

    public static string Read(string filePath, string section, string key, string defaultValue = "")
    {
        var sb = new StringBuilder(2048);
        GetPrivateProfileString(section, key, defaultValue, sb, sb.Capacity, filePath);
        return sb.ToString();
    }

    public static int ReadInt(string filePath, string section, string key, int defaultValue = 0)
    {
        var val = Read(filePath, section, key, defaultValue.ToString());
        return int.TryParse(val, out var result) ? result : defaultValue;
    }

    public static Dictionary<string, string> ReadSection(string filePath, string section)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var buffer = Marshal.AllocCoTaskMem(32768);
        try
        {
            var length = GetPrivateProfileSection(section, buffer, 32768, filePath);
            if (length == 0 || length >= 32767) return result;

            var chars = new char[length];
            Marshal.Copy(buffer, chars, 0, length);
            var raw = new string(chars);
            var lines = raw.Split('\0', StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                var eqIdx = line.IndexOf('=');
                if (eqIdx <= 0) continue;
                var key = line[..eqIdx].Trim();
                var value = line[(eqIdx + 1)..].Trim();
                if (!string.IsNullOrEmpty(key))
                    result[key] = value;
            }
        }
        finally
        {
            Marshal.FreeCoTaskMem(buffer);
        }
        return result;
    }

    public static void Write(string filePath, string section, string key, string value)
    {
        var dir = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        if (!WritePrivateProfileString(section, key, value, filePath))
        {
            var err = Marshal.GetLastWin32Error();
            throw new IOException($"写入INI失败: {filePath} [{section}] {key}={value} (错误码: {err})");
        }
    }

    public static void WriteInt(string filePath, string section, string key, int value)
        => Write(filePath, section, key, value.ToString());
}
