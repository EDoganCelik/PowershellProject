using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using RemoteServiceManager.Models;

namespace RemoteServiceManager.Services
{
    /// <summary>
    /// Makineye bağlı şifreli yapılandırma deposu (.rsmcfg).
    /// Başka makinede açılamaz.
    /// </summary>
    public class ConfigStore
    {
        private const int MaxPresets = 30;
        private const int MaxBackups = 3;

        // Beklenmedik/kötü niyetli büyük dosyaların belleği şişirmesini önler.
        private const long MaxFileBytes = 16L * 1024 * 1024;

        private readonly string _configPath;
        private readonly object _sync = new object();
        private byte[] _sessionKey;
        private AppConfigPayload _payload;

        // Son diske yazılan payload'ın özeti: değişiklik yoksa şifreleme + disk yazımı atlanır.
        private string _lastSavedHash;

        // Makine bağlama hash'i oturum boyunca değişmez; her Save'de yeniden hesaplanmasın.
        private string _machineBinding;

        public ConfigStore()
        {
            var dir = ResolveConfigDirectory();
            Directory.CreateDirectory(dir); // zaten varsa hata vermez
            _configPath = Path.Combine(dir, "settings.rsmcfg");
            _payload = new AppConfigPayload();
        }

        public string ConfigPath => _configPath;

        public AppConfigPayload Payload => _payload;

        public void LoadOrCreate()
        {
            lock (_sync)
            {
                if (!File.Exists(_configPath))
                {
                    CreateFresh();
                    Save();
                    return;
                }

                try
                {
                    var raw = ReadLimited(_configPath);
                    var root = JsonHelper.Deserialize<AppConfigRoot>(raw);
                    if (root == null || string.IsNullOrEmpty(root.WrappedKey))
                        throw new InvalidDataException("Geçersiz yapılandırma dosyası.");

                    if (!MachineBoundCrypto.IsKnownMachineBinding(root.MachineBinding))
                        throw new CryptographicException(
                            "Bu yapılandırma dosyası bu makineye ait değil. Başka bir bilgisayardan kopyalanmış olabilir.");

                    var salt = Convert.FromBase64String(root.Salt);
                    var wrapped = Convert.FromBase64String(root.WrappedKey);
                    var key = MachineBoundCrypto.UnwrapKey(wrapped, salt);

                    AppConfigPayload payload;
                    try
                    {
                        var plain = MachineBoundCrypto.DecryptFromBase64(root.CipherText, key, root.Iv);
                        payload = JsonHelper.Deserialize<AppConfigPayload>(plain) ?? new AppConfigPayload();
                    }
                    catch
                    {
                        // Çözme başarısızsa anahtar oturumda kalmasın (yarım durumda Save() yapılmasın).
                        Array.Clear(key, 0, key.Length);
                        throw;
                    }

                    NormalizePayload(payload);

                    // Alanlar yalnızca tüm adımlar başarılı olunca atanır.
                    ReplaceSessionKey(key);
                    _payload = payload;
                    _lastSavedHash = null;

                    // Eski format / eski binding ile açıldıysa güncel formata yeniden yaz.
                    Save();
                }
                catch (CryptographicException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        "Yapılandırma okunamadı: " + ex.Message, ex);
                }
            }
        }

        public void Save()
        {
            lock (_sync)
            {
                if (_sessionKey == null)
                {
                    // Yüklenemeyen mevcut dosya sessizce ezilmesin; önce yedeği alınır.
                    BackupExistingConfig();
                    CreateFresh();
                }

                if (_payload.Presets != null && _payload.Presets.Count > MaxPresets)
                    _payload.Presets.RemoveRange(0, _payload.Presets.Count - MaxPresets);

                var payloadJson = JsonHelper.Serialize(_payload);

                // İçerik değişmediyse tekrar şifreleyip diske yazma.
                var hash = ComputeHash(payloadJson);
                if (_lastSavedHash != null
                    && string.Equals(_lastSavedHash, hash, StringComparison.Ordinal)
                    && File.Exists(_configPath))
                    return;

                var salt = MachineBoundCrypto.GenerateSalt();
                var wrapped = MachineBoundCrypto.WrapKey(_sessionKey, salt);
                var cipher = MachineBoundCrypto.EncryptToBase64(payloadJson, _sessionKey, out var iv);

                if (_machineBinding == null)
                    _machineBinding = MachineBoundCrypto.GetMachineBindingHash();

                var root = new AppConfigRoot
                {
                    MachineBinding = _machineBinding,
                    Salt = Convert.ToBase64String(salt),
                    WrappedKey = Convert.ToBase64String(wrapped),
                    CipherText = cipher,
                    Iv = iv
                };

                WriteAtomic(_configPath, JsonHelper.Serialize(root));
                _lastSavedHash = hash;
            }
        }

        public bool ProfileNameExists(string name, string excludeTimestamp = null)
        {
            lock (_sync)
            {
                if (string.IsNullOrWhiteSpace(name) || _payload.Presets == null)
                    return false;

                var target = name.Trim();
                return _payload.Presets.Any(p =>
                    p != null
                    && string.Equals(p.Name?.Trim(), target, StringComparison.OrdinalIgnoreCase)
                    && (excludeTimestamp == null
                        || !string.Equals(p.Timestamp, excludeTimestamp, StringComparison.Ordinal)));
            }
        }

        public string SavePreset(PresetRecord record)
        {
            if (record == null)
                throw new ArgumentNullException(nameof(record));

            if (string.IsNullOrWhiteSpace(record.Name))
                throw new ArgumentException("Profil adı gerekli.");

            lock (_sync)
            {
                if (ProfileNameExists(record.Name))
                    throw new InvalidOperationException("Bu profil ismi zaten mevcut.");

                if (_payload.Presets == null)
                    _payload.Presets = new List<PresetRecord>();

                if (string.IsNullOrWhiteSpace(record.Timestamp))
                    record.Timestamp = CreateUniqueTimestamp();

                _payload.Presets.Add(record);
                _payload.LastSelectedKey = record.Timestamp;
                Save();
                return record.Timestamp;
            }
        }

        public string UpdatePreset(string timestamp, PresetRecord updated)
        {
            if (string.IsNullOrWhiteSpace(timestamp))
                throw new ArgumentException("Güncellenecek profil seçilmedi.");
            if (updated == null)
                throw new ArgumentNullException(nameof(updated));

            lock (_sync)
            {
                var existing = (_payload.Presets ?? new List<PresetRecord>())
                    .FirstOrDefault(p => p != null && string.Equals(p.Timestamp, timestamp, StringComparison.Ordinal));
                if (existing == null)
                    throw new InvalidOperationException("Seçili profil bulunamadı.");

                if (!string.IsNullOrWhiteSpace(updated.Name)
                    && ProfileNameExists(updated.Name, excludeTimestamp: timestamp))
                    throw new InvalidOperationException("Bu profil ismi zaten mevcut.");

                existing.Name = updated.Name;
                existing.Services = updated.Services ?? new List<string>();
                existing.FilePath = updated.FilePath ?? string.Empty;
                existing.RegPath = updated.RegPath ?? string.Empty;
                existing.RegValueName = updated.RegValueName ?? string.Empty;
                existing.Options = updated.Options;
                existing.ServiceMetrics = updated.ServiceMetrics;
                existing.PatchMetrics = updated.PatchMetrics;

                _payload.LastSelectedKey = existing.Timestamp;
                Save();
                return existing.Timestamp;
            }
        }

        public void DeletePreset(string timestamp)
        {
            lock (_sync)
            {
                if (string.IsNullOrWhiteSpace(timestamp) || _payload.Presets == null)
                    throw new InvalidOperationException("Silinecek profil seçilmedi.");

                var removed = _payload.Presets.RemoveAll(p =>
                    p != null && string.Equals(p.Timestamp, timestamp, StringComparison.Ordinal));
                if (removed == 0)
                    throw new InvalidOperationException("Seçili profil bulunamadı.");

                if (string.Equals(_payload.LastSelectedKey, timestamp, StringComparison.Ordinal))
                    _payload.LastSelectedKey = _payload.Presets.LastOrDefault()?.Timestamp;

                Save();
            }
        }

        public void SetLastSelected(string key)
        {
            lock (_sync)
            {
                _payload.LastSelectedKey = key;
                Save(); // değer değişmediyse Save() diske yazmadan döner
            }
        }

        /// <summary>
        /// Yapılandırmayı kullanıcı şifresiyle taşınabilir dosyaya yazar.
        /// Hedef makinede aynı şifreyle Import edilebilir.
        /// </summary>
        public void ExportPortable(string filePath, string password)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("Dosya yolu gerekli.", nameof(filePath));
            if (string.IsNullOrEmpty(password))
                throw new ArgumentException("Şifre gerekli.", nameof(password));

            lock (_sync)
            {
                // Makine yapılandırmasının üzerine taşınabilir format yazılmasını engelle.
                var fullPath = Path.GetFullPath(filePath);
                if (string.Equals(fullPath, Path.GetFullPath(_configPath), StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException(
                        "Taşınabilir dosya, makine yapılandırma dosyasının üzerine yazılamaz.", nameof(filePath));

                if (_sessionKey == null)
                    CreateFresh();

                // Anahtar ayrımı: taşınabilir dosya, makinenin oturum anahtarını değil
                // yalnızca bu dışa aktarım için üretilen tek kullanımlık anahtarı içerir.
                var exportKey = MachineBoundCrypto.GenerateRandomKey();
                PortableExportRoot root;
                try
                {
                    var salt = MachineBoundCrypto.GenerateSalt();
                    var wrapped = MachineBoundCrypto.WrapKeyWithPassword(exportKey, password, salt);
                    var payloadJson = JsonHelper.Serialize(_payload);
                    var cipher = MachineBoundCrypto.EncryptToBase64(payloadJson, exportKey, out var iv);

                    root = new PortableExportRoot
                    {
                        Format = "RSMPortable",
                        Version = 1,
                        Salt = Convert.ToBase64String(salt),
                        WrappedKey = Convert.ToBase64String(wrapped),
                        CipherText = cipher,
                        Iv = iv
                    };
                }
                finally
                {
                    Array.Clear(exportKey, 0, exportKey.Length);
                }

                WriteAtomic(fullPath, JsonHelper.Serialize(root));
            }
        }

        /// <summary>
        /// Şifreli taşınabilir dosyayı okur, mevcut makineye bağlar ve kaydeder.
        /// </summary>
        public void ImportPortable(string filePath, string password)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException("Import dosyası bulunamadı.", filePath);
            if (string.IsNullOrEmpty(password))
                throw new ArgumentException("Şifre gerekli.", nameof(password));

            var raw = ReadLimited(filePath);
            var root = JsonHelper.Deserialize<PortableExportRoot>(raw);
            if (root == null
                || !string.Equals(root.Format, "RSMPortable", StringComparison.OrdinalIgnoreCase)
                || string.IsNullOrEmpty(root.WrappedKey))
            {
                throw new InvalidDataException("Geçersiz taşınabilir yapılandırma dosyası.");
            }

            byte[] importKey;
            try
            {
                var salt = Convert.FromBase64String(root.Salt);
                var wrapped = Convert.FromBase64String(root.WrappedKey);
                importKey = MachineBoundCrypto.UnwrapKeyWithPassword(wrapped, password, salt);
            }
            catch (CryptographicException)
            {
                throw new CryptographicException("Şifre hatalı veya dosya bozulmuş.");
            }
            catch (Exception ex)
            {
                throw new CryptographicException("Şifre hatalı veya dosya bozulmuş.", ex);
            }

            string plain;
            try
            {
                plain = MachineBoundCrypto.DecryptFromBase64(root.CipherText, importKey, root.Iv);
            }
            catch (Exception ex)
            {
                throw new CryptographicException("Veri çözülemedi. Şifre hatalı olabilir.", ex);
            }
            finally
            {
                // Dosyadan gelen anahtar yalnızca verinin çözülmesi için kullanılır.
                Array.Clear(importKey, 0, importKey.Length);
            }

            var imported = JsonHelper.Deserialize<AppConfigPayload>(plain) ?? new AppConfigPayload();
            NormalizePayload(imported);

            lock (_sync)
            {
                // Yerel oturum anahtarı taşınabilir dosyadan bağımsız, yeni üretilir;
                // Save() veriyi bu makineye bağlayarak yeniden şifreler.
                ReplaceSessionKey(MachineBoundCrypto.GenerateRandomKey());
                _payload = imported;
                _lastSavedHash = null;
                Save();
            }
        }

        private void CreateFresh()
        {
            ReplaceSessionKey(MachineBoundCrypto.GenerateRandomKey());
            _lastSavedHash = null;
            _payload = new AppConfigPayload
            {
                Presets = new List<PresetRecord>(),
                DefaultOptions = new CheckOptions(),
                Ui = new UiPreferences()
            };
        }

        // ------------------------------------------------------------------
        // Yardımcılar
        // ------------------------------------------------------------------

        /// <summary>
        /// Eski davranış (C:\Users\{kullanıcı}\Desktop\ServiceStatus) korunur; dosya orada varsa orası kullanılır.
        /// Yoksa gerçek Masaüstü klasörü (OneDrive/yönlendirilmiş profiller dahil) kullanılır.
        /// </summary>
        private static string ResolveConfigDirectory()
        {
            var legacy = Path.Combine(@"C:\Users", Environment.UserName, "Desktop", "ServiceStatus");
            if (File.Exists(Path.Combine(legacy, "settings.rsmcfg")))
                return legacy;

            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            return string.IsNullOrWhiteSpace(desktop)
                ? legacy
                : Path.Combine(desktop, "ServiceStatus");
        }

        private static void NormalizePayload(AppConfigPayload payload)
        {
            if (payload.Presets == null)
                payload.Presets = new List<PresetRecord>();
            else
                payload.Presets.RemoveAll(p => p == null); // bozuk/elle düzenlenmiş dosyalara karşı
            if (payload.DefaultOptions == null)
                payload.DefaultOptions = new CheckOptions();
            if (payload.Ui == null)
                payload.Ui = new UiPreferences();
        }

        private void ReplaceSessionKey(byte[] newKey)
        {
            var old = _sessionKey;
            _sessionKey = newKey;
            if (old != null && !ReferenceEquals(old, newKey))
                Array.Clear(old, 0, old.Length); // eski anahtarı bellekte bırakma
        }

        /// <summary>
        /// Zaman damgası profil anahtarı olarak kullanıldığından aynı saniyede eklenen iki profil
        /// aynı anahtarı almasın (aksi halde güncelle/sil her ikisini etkilerdi).
        /// </summary>
        private string CreateUniqueTimestamp()
        {
            var now = DateTime.Now;
            string ts;
            do
            {
                ts = now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                now = now.AddSeconds(1);
            }
            while (_payload.Presets.Any(p => p != null
                && string.Equals(p.Timestamp, ts, StringComparison.Ordinal)));
            return ts;
        }

        private static string ComputeHash(string text)
        {
            using (var sha = SHA256.Create())
                return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(text)));
        }

        private static string ReadLimited(string path)
        {
            if (new FileInfo(path).Length > MaxFileBytes)
                throw new InvalidDataException("Dosya beklenenden büyük; geçerli bir yapılandırma dosyası değil.");
            return File.ReadAllText(path);
        }

        /// <summary>
        /// Geçici dosyaya yazıp diske flush eder, sonra atomik olarak yerine koyar.
        /// Elektrik kesintisi/çökme durumunda yapılandırmanın yarım yazılmasını önler.
        /// Antivirüs vb. kilitler için kısa yeniden deneme içerir.
        /// </summary>
        private static void WriteAtomic(string path, string content)
        {
            var tmp = path + ".tmp";
            try
            {
                using (var fs = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None,
                    4096, FileOptions.WriteThrough))
                using (var sw = new StreamWriter(fs, new UTF8Encoding(false)))
                {
                    sw.Write(content);
                    sw.Flush();
                    fs.Flush(true);
                }

                for (int attempt = 0; ; attempt++)
                {
                    try
                    {
                        if (File.Exists(path))
                            File.Replace(tmp, path, null);
                        else
                            File.Move(tmp, path);
                        break;
                    }
                    catch (IOException) when (attempt < 2)
                    {
                        Thread.Sleep(50 * (attempt + 1));
                    }
                }
            }
            catch
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
                throw;
            }
        }

        /// <summary>
        /// Okunamayan (ör. başka makineden gelen) yapılandırma yeni dosyayla ezilmeden önce yedeklenir.
        /// En fazla MaxBackups yedek tutulur.
        /// </summary>
        private void BackupExistingConfig()
        {
            try
            {
                if (!File.Exists(_configPath)) return;

                var backup = _configPath + ".bak-" + DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
                File.Copy(_configPath, backup, false);

                var dir = Path.GetDirectoryName(_configPath);
                var backups = Directory.GetFiles(dir, Path.GetFileName(_configPath) + ".bak-*")
                    .OrderBy(f => f, StringComparer.Ordinal)
                    .ToList();
                for (int i = 0; i < backups.Count - MaxBackups; i++)
                    File.Delete(backups[i]);
            }
            catch
            {
                // Yedekleme başarısız olsa da ana işlem devam eder.
            }
        }
    }
}