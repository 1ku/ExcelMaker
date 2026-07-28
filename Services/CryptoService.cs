using System.Security.Cryptography;

namespace ExcelMaker.Services;

/// <summary>
/// AES 加解密（与参考项目 AutoSteelCheckMain 共用同一密钥，
/// 因此已有的 db.ini 密文可直接复用）。
/// </summary>
public class CryptoService
{
    // 与 AutoSteelCheckMain 保持一致的密钥，确保历史密文可解密
    private static readonly byte[] AesKey =
    {
        0x4A, 0x7F, 0x2C, 0x91, 0xE3, 0x5D, 0x8B, 0x06,
        0x1F, 0x43, 0xA8, 0xCD, 0x72, 0xFE, 0x19, 0x64,
        0x3B, 0x88, 0xD5, 0x0E, 0x97, 0x52, 0xAC, 0xF1,
        0x69, 0xC4, 0x27, 0x83, 0xBE, 0x10, 0x4D, 0xE2
    };

    private static readonly byte[] AesIv =
    {
        0x9D, 0x31, 0x7C, 0x58, 0xA6, 0xE2, 0x1F, 0x43,
        0x8B, 0xCF, 0x05, 0x69, 0xD4, 0x72, 0xBA, 0x1E
    };

    public string Encrypt(string plainText)
    {
        using var aes = Aes.Create();
        aes.Key = AesKey;
        aes.IV = AesIv;
        using var encryptor = aes.CreateEncryptor();
        var plainBytes = Encoding.UTF8.GetBytes(plainText);
        var cipherBytes = encryptor.TransformFinalBlock(plainBytes, 0, plainBytes.Length);
        return Convert.ToBase64String(cipherBytes);
    }

    public string Decrypt(string cipherText)
    {
        using var aes = Aes.Create();
        aes.Key = AesKey;
        aes.IV = AesIv;
        using var decryptor = aes.CreateDecryptor();
        var cipherBytes = Convert.FromBase64String(cipherText);
        var plainBytes = decryptor.TransformFinalBlock(cipherBytes, 0, cipherBytes.Length);
        return Encoding.UTF8.GetString(plainBytes);
    }

    /// <summary>
    /// 生成连接串密文（用于离线生成 db.ini，无需联网工具）。
    /// </summary>
    public static string GenerateCipher(string plainConnectionString)
    {
        var svc = new CryptoService();
        return svc.Encrypt(plainConnectionString);
    }
}
