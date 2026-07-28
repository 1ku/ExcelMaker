using System.Security.Cryptography;

namespace ExcelMaker.Helpers;

public static class Md5Helper
{
    public static string StrToMD5(string input)
    {
        var hash = MD5.HashData(Encoding.UTF8.GetBytes(input));
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    public static string Compute(string input) => StrToMD5(input);
}
