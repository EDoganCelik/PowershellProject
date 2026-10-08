using System;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace RemoteServiceManager.Services
{
    /// <summary>
    /// AES-256 şifreleme. Anahtar rastgele üretilir, makine parmak izi (PBKDF2-SHA256) ile sarmalanır.
    /// </summary>
    public static class MachineBoundCrypto
    {
        private const int Pbkdf2Iterations = 100000;
        private const int AesBlockSize = 16;
        private const int AesKeySize = 32;
        private const string NoGuid = "NO-GUID";

        // RandomNumberGenerator thread-safe'tir; her çağrıda yeniden oluşturmaya gerek yok.
        private static readonly RandomNumberGenerator Rng = RandomNumberGenerator.Create();

        // MachineGuid başarıyla okunduysa parmak izi önbelleğe alınır (registry + string birleştirme maliyeti kalkar).
        // Registry okunamadıysa önbelleğe alınmaz; sonraki çağrıda tekrar denenir (önceki davranışla aynı).
        private static volatile string _cachedFingerprint;

        public static string GetMachineFingerprint()
        {
            var cached = _cachedFingerprint;
            if (cached != null)
                return cached;

            var guid = GetMachineGuid();
            var fingerprint = (Environment.MachineName ?? string.Empty).ToUpperInvariant() + "|" + guid;

            if (!string.Equals(guid, NoGuid, StringComparison.Ordinal))
                _cachedFingerprint = fingerprint;

            return fingerprint;
        }

        public static string GetMachineBindingHash()
        {
            using (var sha = SHA256.Create())
            {
                var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(GetMachineFingerprint()));
                return Convert.ToBase64String(bytes);
            }
        }

        public static bool IsKnownMachineBinding(string binding)
        {
            if (string.IsNullOrEmpty(binding))
                return false;
            return FixedTimeEquals(binding, GetMachineBindingHash());
        }

        public static byte[] DeriveMachineKey(byte[] salt)
        {
            return DeriveBytes(GetMachineFingerprint(), salt);
        }

        public static byte[] DerivePasswordKey(string password, byte[] salt)
        {
            if (string.IsNullOrEmpty(password))
                throw new ArgumentException("Şifre boş olamaz.", nameof(password));
            return DeriveBytes(password, salt);
        }

        private static byte[] DeriveBytes(string passwordOrFingerprint, byte[] salt)
        {
            if (salt == null)
                throw new ArgumentNullException(nameof(salt));

            using (var derive = new Rfc2898DeriveBytes(
                passwordOrFingerprint, salt, Pbkdf2Iterations, HashAlgorithmName.SHA256))
            {
                return derive.GetBytes(AesKeySize);
            }
        }

        public static byte[] WrapKeyWithPassword(byte[] aesKey, string password, byte[] salt)
        {
            var passwordKey = DerivePasswordKey(password, salt);
            try
            {
                return AesEncryptWithEmbeddedIv(aesKey, passwordKey);
            }
            finally
            {
                ClearBytes(passwordKey);
            }
        }

        public static byte[] UnwrapKeyWithPassword(byte[] wrapped, string password, byte[] salt)
        {
            var passwordKey = DerivePasswordKey(password, salt);
            try
            {
                return AesDecryptEmbeddedIv(wrapped, passwordKey);
            }
            finally
            {
                ClearBytes(passwordKey);
            }
        }

        public static byte[] GenerateRandomKey()
        {
            var key = new byte[AesKeySize];
            Rng.GetBytes(key);
            return key;
        }

        public static byte[] GenerateSalt(int size = 16)
        {
            var salt = new byte[size];
            Rng.GetBytes(salt);
            return salt;
        }

        public static byte[] WrapKey(byte[] aesKey, byte[] salt)
        {
            var machineKey = DeriveMachineKey(salt);
            try
            {
                return AesEncryptWithEmbeddedIv(aesKey, machineKey);
            }
            finally
            {
                ClearBytes(machineKey);
            }
        }

        public static byte[] UnwrapKey(byte[] wrapped, byte[] salt)
        {
            var machineKey = DeriveMachineKey(salt);
            try
            {
                return AesDecryptEmbeddedIv(wrapped, machineKey);
            }
            finally
            {
                ClearBytes(machineKey);
            }
        }

        /// <summary>
        /// Payload şifreleme: IV ayrı döner, cipher içinde tekrarlanmaz.
        /// </summary>
        public static string EncryptToBase64(string plainText, byte[] aesKey, out string ivBase64)
        {
            var plain = Encoding.UTF8.GetBytes(plainText ?? string.Empty);
            try
            {
                using (var aes = Aes.Create())
                {
                    aes.Mode = CipherMode.CBC;
                    aes.Padding = PaddingMode.PKCS7;
                    aes.Key = aesKey;
                    aes.GenerateIV();
                    ivBase64 = Convert.ToBase64String(aes.IV);

                    using (var encryptor = aes.CreateEncryptor())
                    {
                        return Convert.ToBase64String(encryptor.TransformFinalBlock(plain, 0, plain.Length));
                    }
                }
            }
            finally
            {
                ClearBytes(plain);
            }
        }

        public static string DecryptFromBase64(string cipherBase64, byte[] aesKey, string ivBase64)
        {
            if (string.IsNullOrWhiteSpace(cipherBase64) || string.IsNullOrWhiteSpace(ivBase64))
                throw new CryptographicException("Şifreli veri veya IV eksik.");

            var cipher = Convert.FromBase64String(cipherBase64);
            var iv = Convert.FromBase64String(ivBase64);

            var plain = AesDecryptRaw(cipher, aesKey, iv);
            string text;
            try
            {
                text = Encoding.UTF8.GetString(plain);
            }
            finally
            {
                ClearBytes(plain);
            }

            if (!LooksLikeJsonObject(text))
                throw new CryptographicException("Çözülen yapılandırma verisi geçerli değil.");

            return text;
        }

        // string.Trim() ile aynı mantık (char.IsWhiteSpace) ama yeni string üretmeden.
        private static bool LooksLikeJsonObject(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;

            int start = 0;
            int end = text.Length - 1;

            while (start <= end && char.IsWhiteSpace(text[start])) start++;
            while (end >= start && char.IsWhiteSpace(text[end])) end--;

            return end - start + 1 >= 2 && text[start] == '{' && text[end] == '}';
        }

        /// <summary>WrapKey: IV + cipher tek blob.</summary>
        private static byte[] AesEncryptWithEmbeddedIv(byte[] plain, byte[] key)
        {
            using (var aes = Aes.Create())
            {
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                aes.Key = key;
                aes.GenerateIV();
                var iv = aes.IV;

                byte[] cipher;
                using (var encryptor = aes.CreateEncryptor())
                {
                    cipher = encryptor.TransformFinalBlock(plain, 0, plain.Length);
                }

                var blob = new byte[iv.Length + cipher.Length];
                Buffer.BlockCopy(iv, 0, blob, 0, iv.Length);
                Buffer.BlockCopy(cipher, 0, blob, iv.Length, cipher.Length);
                return blob;
            }
        }

        private static byte[] AesDecryptEmbeddedIv(byte[] blob, byte[] key)
        {
            if (blob == null || blob.Length < AesBlockSize + 1)
                throw new CryptographicException("Geçersiz şifreli veri.");

            var iv = new byte[AesBlockSize];
            Buffer.BlockCopy(blob, 0, iv, 0, AesBlockSize);
            var cipher = new byte[blob.Length - AesBlockSize];
            Buffer.BlockCopy(blob, AesBlockSize, cipher, 0, cipher.Length);
            return AesDecryptRaw(cipher, key, iv);
        }

        private static byte[] AesDecryptRaw(byte[] cipher, byte[] key, byte[] iv)
        {
            using (var aes = Aes.Create())
            {
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                aes.Key = key;
                aes.IV = iv;

                using (var decryptor = aes.CreateDecryptor())
                {
                    return decryptor.TransformFinalBlock(cipher, 0, cipher.Length);
                }
            }
        }

        private static string GetMachineGuid()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Cryptography"))
                {
                    var guid = key?.GetValue("MachineGuid") as string;
                    if (!string.IsNullOrWhiteSpace(guid))
                        return guid;
                }
            }
            catch
            {
                // ignore
            }
            return NoGuid;
        }

        /// <summary>
        /// Zamanlama saldırılarına karşı sabit sürede string karşılaştırma
        /// (.NET Framework'te CryptographicOperations.FixedTimeEquals yok).
        /// </summary>
        [MethodImpl(MethodImplOptions.NoInlining | MethodImplOptions.NoOptimization)]
        private static bool FixedTimeEquals(string a, string b)
        {
            int diff = a.Length ^ b.Length;
            int n = Math.Min(a.Length, b.Length);
            for (int i = 0; i < n; i++)
                diff |= a[i] ^ b[i];
            return diff == 0;
        }

        /// <summary>Hassas byte dizisini bellekte sıfırlar.</summary>
        private static void ClearBytes(byte[] data)
        {
            if (data != null)
                Array.Clear(data, 0, data.Length);
        }
    }
}