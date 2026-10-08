using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Management;
using System.ServiceProcess;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using RemoteServiceManager.Models;

namespace RemoteServiceManager.Services
{
    public class RemoteChecker
    {
        // ---- Önceden derlenmiş regex'ler (her çağrıda yeniden oluşturulmaz) ----

        // Bilgisayar adı: UNC yoluna girdiği için '\', '/', ':', boşluk vb. karakterlere izin verilmez.
        private static readonly Regex HostNameRegex = new Regex(
            @"^\w[\w\-\.]{0,252}\z",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // C:\... veya C:/... biçimindeki yerel yol
        private static readonly Regex DriveLetterPathRegex = new Regex(
            @"^([A-Za-z]):[\\/](.*)$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // HKLM / HKEY_LOCAL_MACHINE / HKU / HKEY_USERS (isteğe bağlı ':' ile, ardından '\' veya satır sonu)
        private static readonly Regex RegHiveRegex = new Regex(
            @"^(?:(?<lm>HKLM|HKEY_LOCAL_MACHINE)|(?<u>HKU|HKEY_USERS)):?(?:\\|\z)",
            RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        // Servis adı (WQL'e girdiği için katı beyaz liste): harf, rakam, '_', '-', '.', '$', '@', '#', boşluk, parantez.
        // Tırnak, ters eğik çizgi, noktalı virgül, '=', '%', kontrol karakteri vb. kesinlikle kabul edilmez.
        private static readonly Regex ServiceNameRegex = new Regex(
            @"^[\w\-\.\$@#\(\) ]{1,256}\z",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // Olay günlüğü: 1074 / 6006 / 6008
        private const string ShutdownXPath = "*[System[(EventID=1074 or EventID=6006 or EventID=6008)]]";

        // Bu sayı ve üzerindeki servis isteğinde tek seferde toplu sorgu yapılır
        private const int BulkServiceThreshold = 3;

        public HostResult Check(
            string computerName,
            IList<string> serviceNames,
            string filePath,
            string regPath,
            string regValueName,
            CheckOptions options)
        {
            var prefs = MetricCatalog.FromLegacy(options);
            return Check(computerName, serviceNames, filePath, regPath, regValueName, prefs, options);
        }

        public HostResult Check(
            string computerName,
            IList<string> serviceNames,
            string filePath,
            string regPath,
            string regValueName,
            MetricPreferences prefs,
            CheckOptions legacyOptions = null)
        {
            prefs = prefs ?? MetricCatalog.DefaultServicePreferences();
            var active = MetricCatalog.ResolveActive(prefs, MetricScope.Service);
            var need = new HashSet<string>(active.Where(m => !m.IsCustom).Select(m => m.Id), StringComparer.OrdinalIgnoreCase);
            var customs = active.Where(m => m.IsCustom).ToList();
            int pingTimeout = legacyOptions != null ? legacyOptions.PingTimeoutMs : 2000;

            var row = new HostResult
            {
                ComputerName = computerName,
                Hostname = computerName,
                LastChecked = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            };

            // Güvenlik: geçersiz / zararlı karakter içeren makine adları hiçbir uzak çağrıya (ping, UNC, WMI...) girmesin.
            if (!IsValidHostName(computerName))
            {
                row.IsOnline = false;
                row.OnlineStatus = "Offline";
                row.ErrorMessage = "Geçersiz bilgisayar adı";
                return row;
            }

            if (!HostReachability.PingOk(computerName, pingTimeout))
            {
                row.IsOnline = false;
                row.OnlineStatus = "Offline";
                row.ErrorMessage = "Makine yanıt vermiyor (ping başarısız)";
                return row;
            }

            // Ping yetmez: DNS kaydı / false-positive ICMP olabilir. C$ erişimi zorunlu.
            if (!HostReachability.CanAccessAdminShare(computerName))
            {
                row.IsOnline = false;
                row.OnlineStatus = "Offline";
                row.ErrorMessage = "Makine yanıt vermiyor (C$ erişilemedi)";
                return row;
            }

            row.IsOnline = true;
            row.OnlineStatus = "Evet";

            bool needWmi = need.Contains(MetricCatalog.Uptime) || need.Contains(MetricCatalog.LastBootTime)
                || need.Contains(MetricCatalog.DiskInfo) || need.Contains(MetricCatalog.LoggedInUser)
                || need.Contains(MetricCatalog.NetworkInfo)
                || need.Contains(MetricCatalog.LastWindowsUpdate) || need.Contains(MetricCatalog.MemoryInfo)
                || need.Contains(MetricCatalog.TopProcesses)
                || customs.Any(c => c.Kind == MetricKind.WmiProperty);

            ManagementScope scope = null;
            if (needWmi)
            {
                try { scope = WmiHelper.Connect(computerName); }
                catch (Exception ex)
                {
                    row.ErrorMessage = "WMI bağlantı hatası: " + ex.Message;
                }
            }

            if ((need.Contains(MetricCatalog.Running) || need.Contains(MetricCatalog.NotRunning))
                && serviceNames != null && serviceNames.Count > 0)
                CollectServices(row, computerName, serviceNames, scope);

            if (need.Contains(MetricCatalog.LastShutdown))
                CollectLastShutdown(row, computerName);

            if (need.Contains(MetricCatalog.FilePathControl) && !string.IsNullOrWhiteSpace(filePath))
                CollectFileVersion(row, computerName, filePath);

            if (need.Contains(MetricCatalog.RegistryValue)
                && !string.IsNullOrWhiteSpace(regPath)
                && !string.IsNullOrWhiteSpace(regValueName))
                CollectRegistry(row, computerName, regPath, regValueName);

            if (need.Contains(MetricCatalog.Bsod))
            {
                var bsod = PatchChecker.GetBsodStatus(computerName);
                row.BsodFileCount = bsod.Item1.ToString();
                row.Bsod = bsod.Item1 > 2 ? "VAR" : "YOK";
            }

            if (need.Contains(MetricCatalog.PendingUpdates))
            {
                var patch = new PatchChecker();
                var details = patch.GetPatchStatusDetails(computerName);
                row.PendingUpdates = details.Summary;
                row.UpdateSummary = details.Summary;
                row.PatchDetailLines = details.Lines;
            }

            if (need.Contains(MetricCatalog.OsBuild))
                row.OsBuild = PatchChecker.GetOsBuildWithUbr(computerName);

            if (scope != null && scope.IsConnected)
            {
                bool wantOs = need.Contains(MetricCatalog.Uptime) || need.Contains(MetricCatalog.LastBootTime)
                    || need.Contains(MetricCatalog.MemoryInfo);
                if (wantOs)
                    CollectOsInfo(row, scope, need);

                if (need.Contains(MetricCatalog.DiskInfo))
                    CollectDisk(row, scope);

                if (need.Contains(MetricCatalog.LoggedInUser))
                    CollectLoggedInUser(row, scope);

                if (need.Contains(MetricCatalog.NetworkInfo))
                    CollectNetwork(row, scope);

                if (need.Contains(MetricCatalog.LastWindowsUpdate))
                    CollectLastUpdate(row, scope);

                if (need.Contains(MetricCatalog.TopProcesses))
                    CollectTopProcesses(row, scope);
            }
            else if (needWmi && string.IsNullOrEmpty(row.ErrorMessage))
            {
                row.ErrorMessage = "WMI bağlantısı kurulamadı; sistem metrikleri alınamadı.";
            }

            foreach (var cm in customs)
                row.CustomValues[cm.ResultKey] = CollectCustomMetric(computerName, cm, scope);

            return row;
        }

        private static string CollectCustomMetric(string computer, MetricDefinition m, ManagementScope scope)
        {
            try
            {
                switch (m.Kind)
                {
                    case MetricKind.Service:
                        using (var sc = new ServiceController(m.Param1, computer))
                            return sc.Status.ToString();
                    case MetricKind.FileVersion:
                        return ReadFileVersion(computer, m.Param1);
                    case MetricKind.Registry:
                        return ReadRemoteRegistryValue(computer, m.Param1, m.Param2);
                    case MetricKind.WmiProperty:
                        if (scope == null || !scope.IsConnected) return "WMI yok";
                        // Güvenlik: özel metriklerde yalnızca salt-okunur SELECT sorgularına izin ver.
                        if (!IsSelectQuery(m.Param1))
                            return "Hata: Yalnızca SELECT sorgularına izin verilir";
                        using (var results = WmiHelper.Query(scope, m.Param1))
                        {
                            foreach (ManagementObject mo in results)
                            {
                                using (mo)
                                    return WmiHelper.GetProp(mo, m.Param2) ?? "";
                            }
                        }
                        return "Bulunamadı";
                }
            }
            catch (Exception ex)
            {
                return "Hata: " + ex.Message;
            }
            return "N/A";
        }

        public static string ReadFileVersion(string computer, string localPath)
        {
            try
            {
                var unc = ToUncPath(computer, localPath);
                if (string.IsNullOrEmpty(unc) || !File.Exists(unc))
                    return "False";
                var vi = FileVersionInfo.GetVersionInfo(unc);
                return string.IsNullOrWhiteSpace(vi.FileVersion) ? "True" : vi.FileVersion;
            }
            catch
            {
                return "False";
            }
        }

        private static void CollectServices(HostResult row, string computer, IList<string> names, ManagementScope scope)
        {
            var detail = new List<ServiceStatusItem>();

            // Performans: çok sayıda servis varsa her biri için ayrı SCM bağlantısı açmak yerine
            // tek seferde tüm servisleri listele. Bulunamayanlar eski yönteme (ServiceController -> WMI) düşer.
            Dictionary<string, string> bulk = names.Count >= BulkServiceThreshold
                ? TryGetAllServiceStatuses(computer)
                : null;

            foreach (var raw in names)
            {
                var s = (raw ?? string.Empty).Trim();
                if (s.Length == 0) continue;

                string status = null;
                if (bulk != null)
                    bulk.TryGetValue(s, out status);

                if (status == null)
                {
                    try
                    {
                        using (var sc = new ServiceController(s, computer))
                        {
                            status = sc.Status.ToString();
                        }
                    }
                    catch
                    {
                        status = TryServiceViaWmi(scope, s) ?? "NotFound";
                    }
                }

                detail.Add(new ServiceStatusItem { Name = s, Status = status });
            }

            row.ServiceDetail = detail;
            row.Running = string.Join(", ", detail.Where(d => d.Status == "Running").Select(d => d.Name));
            row.NotRunning = string.Join(", ",
                detail.Where(d => d.Status != "Running").Select(d => d.Name + "(" + d.Status + ")"));
        }

        private static Dictionary<string, string> TryGetAllServiceStatuses(string computer)
        {
            try
            {
                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var all = ServiceController.GetServices(computer);
                foreach (var sc in all)
                {
                    try { map[sc.ServiceName] = sc.Status.ToString(); }
                    catch { /* tek servis okunamadıysa atla; eski yönteme düşer */ }
                    finally { sc.Dispose(); }
                }
                return map;
            }
            catch
            {
                return null;
            }
        }

        private static string TryServiceViaWmi(ManagementScope scope, string name)
        {
            if (scope == null || !scope.IsConnected) return null;

            // Güvenlik: servis adı WQL sorgusuna girmeden önce katı beyaz liste ile doğrulanır.
            // Geçersiz ad sorguya hiç sızmaz; sonuç eskisi gibi null (-> "NotFound") döner.
            if (name == null || !ServiceNameRegex.IsMatch(name)) return null;

            try
            {
                var q = "SELECT State FROM Win32_Service WHERE Name='" + name + "'";
                using (var results = WmiHelper.Query(scope, q))
                {
                    foreach (ManagementObject mo in results)
                    {
                        using (mo)
                        {
                            return WmiHelper.GetProp(mo, "State") ?? "Unknown";
                        }
                    }
                }
            }
            catch
            {
                // ignore
            }
            return null;
        }

        private static void CollectLastShutdown(HostResult row, string computer)
        {
            try
            {
                // Hızlı yol: sunucu tarafında filtrelenmiş, sondan başlayan sorgu (tek kayıt okunur).
                DateTime? found;
                if (TryQueryLastShutdown(computer, out found))
                {
                    if (found.HasValue)
                        row.LastShutdown = found.Value.ToString("yyyy-MM-dd HH:mm:ss");
                    return;
                }

                // Yedek yol: eski EventLog taraması (hızlı yol kullanılamadıysa).
                using (var log = new EventLog("System", computer))
                {
                    // Son kayıtları sondan tarayarak 1074/6006/6008 ara
                    var entries = log.Entries;
                    int count = entries.Count; // her döngüde uzak Count çağrısı yapmamak için
                    for (int i = count - 1; i >= 0 && i >= count - 5000; i--)
                    {
                        var e = entries[i];
                        // InstanceId üst bitleri (severity/qualifier) içerir; gerçek olay kimliği için maskele.
                        long id = e.InstanceId & 0x3FFFFFFF;
                        if (id == 1074 || id == 6006 || id == 6008)
                        {
                            row.LastShutdown = e.TimeGenerated.ToString("yyyy-MM-dd HH:mm:ss");
                            return;
                        }
                    }
                }
            }
            catch
            {
                row.LastShutdown = "N/A";
            }
        }

        // true: sorgu çalıştı (time boş olabilir = kayıt yok). false: yöntem kullanılamadı, yedek yola geç.
        private static bool TryQueryLastShutdown(string computer, out DateTime? time)
        {
            time = null;
            try
            {
                using (var session = new EventLogSession(computer))
                {
                    var query = new EventLogQuery("System", PathType.LogName, ShutdownXPath);
                    query.Session = session;
                    query.ReverseDirection = true; // en yeni kayıt önce

                    using (var reader = new EventLogReader(query))
                    using (var rec = reader.ReadEvent())
                    {
                        if (rec != null && rec.TimeCreated.HasValue)
                            time = rec.TimeCreated.Value;
                    }
                }
                return true;
            }
            catch
            {
                time = null;
                return false;
            }
        }

        private static void CollectFileVersion(HostResult row, string computer, string localPath)
        {
            // ReadFileVersion ile birebir aynı mantık; tekrar eden kod kaldırıldı.
            row.FilePathControl = ReadFileVersion(computer, localPath);
        }

        private static void CollectRegistry(HostResult row, string computer, string regPath, string valueName)
        {
            try
            {
                row.RegistryValue = ReadRemoteRegistryValue(computer, regPath, valueName);
            }
            catch (Exception ex)
            {
                row.RegistryValue = "Hata: " + ex.Message;
            }
        }

        private static void CollectOsInfo(HostResult row, ManagementScope scope, HashSet<string> need)
        {
            try
            {
                using (var results = WmiHelper.Query(scope,
                    "SELECT Caption, Version, BuildNumber, LastBootUpTime, TotalVisibleMemorySize, FreePhysicalMemory FROM Win32_OperatingSystem"))
                {
                    foreach (ManagementObject mo in results)
                    {
                        using (mo)
                        {
                            if (need.Contains(MetricCatalog.Uptime) || need.Contains(MetricCatalog.LastBootTime))
                            {
                                var boot = WmiHelper.ParseWmiDate(WmiHelper.GetProp(mo, "LastBootUpTime"));
                                if (boot.HasValue)
                                {
                                    row.LastBootTime = boot.Value.ToString("yyyy-MM-dd HH:mm:ss");
                                    var up = DateTime.Now - boot.Value;
                                    row.Uptime = string.Format("{0}g {1}s {2}dk",
                                        (int)up.TotalDays, up.Hours, up.Minutes);
                                }
                            }

                            if (need.Contains(MetricCatalog.MemoryInfo))
                            {
                                ulong totalKb, freeKb;
                                if (ulong.TryParse(WmiHelper.GetProp(mo, "TotalVisibleMemorySize"), out totalKb)
                                    && ulong.TryParse(WmiHelper.GetProp(mo, "FreePhysicalMemory"), out freeKb))
                                {
                                    double totalGb = totalKb / 1024.0 / 1024.0;
                                    double freeGb = freeKb / 1024.0 / 1024.0;
                                    double usedPct = totalKb == 0 ? 0 : (1.0 - (double)freeKb / totalKb) * 100.0;
                                    row.MemoryInfo = string.Format("Toplam: {0:0.00} GB | Boş: {1:0.00} GB | Kullanım: %{2:0.0}",
                                        totalGb, freeGb, usedPct);
                                }
                            }
                        }
                        break;
                    }
                }
            }
            catch
            {
                // Uptime / bellek; OS Build ayrı registry sorgusunda
            }
        }

        private static void CollectDisk(HostResult row, ManagementScope scope)
        {
            try
            {
                var parts = new List<string>();
                using (var results = WmiHelper.Query(scope,
                    "SELECT DeviceID, FreeSpace, Size FROM Win32_LogicalDisk WHERE DriveType=3"))
                {
                    foreach (ManagementObject mo in results)
                    {
                        using (mo)
                        {
                            ulong free, size;
                            if (!ulong.TryParse(WmiHelper.GetProp(mo, "FreeSpace"), out free)) continue;
                            if (!ulong.TryParse(WmiHelper.GetProp(mo, "Size"), out size) || size == 0) continue;
                            var id = WmiHelper.GetProp(mo, "DeviceID") ?? "?";
                            double freeGb = free / 1024.0 / 1024.0 / 1024.0;
                            double pctUsed = (1.0 - (double)free / size) * 100.0;
                            parts.Add(string.Format("{0} Boş:{1:0.0}GB Doluluk:%{2:0.0}", id, freeGb, pctUsed));
                        }
                    }
                }
                row.DiskInfo = parts.Count > 0 ? string.Join(" | ", parts) : "N/A";
            }
            catch (Exception ex)
            {
                row.DiskInfo = "Hata: " + ex.Message;
            }
        }

        private static void CollectLoggedInUser(HostResult row, ManagementScope scope)
        {
            try
            {
                using (var results = WmiHelper.Query(scope, "SELECT UserName FROM Win32_ComputerSystem"))
                {
                    foreach (ManagementObject mo in results)
                    {
                        using (mo)
                        {
                            var u = WmiHelper.GetProp(mo, "UserName");
                            row.LoggedInUser = string.IsNullOrWhiteSpace(u) ? "(oturum yok / bilinmiyor)" : u;
                        }
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                row.LoggedInUser = "Hata: " + ex.Message;
            }
        }

        private static void CollectNetwork(HostResult row, ManagementScope scope)
        {
            try
            {
                // IP/MAC: yapılandırma; hız: adaptör
                var adapters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                using (var adapterResults = WmiHelper.Query(scope,
                    "SELECT Index, MACAddress, Speed, NetConnectionStatus, NetEnabled FROM Win32_NetworkAdapter WHERE PhysicalAdapter=True"))
                {
                    foreach (ManagementObject mo in adapterResults)
                    {
                        using (mo)
                        {
                            var idx = WmiHelper.GetProp(mo, "Index");
                            var mac = WmiHelper.GetProp(mo, "MACAddress");
                            if (string.IsNullOrEmpty(idx) || string.IsNullOrEmpty(mac)) continue;

                            ulong speed;
                            ulong.TryParse(WmiHelper.GetProp(mo, "Speed"), out speed);
                            string speedText = speed > 0
                                ? (speed >= 1000000000
                                    ? string.Format("{0:0} Gbps", speed / 1000000000.0)
                                    : string.Format("{0:0} Mbps", speed / 1000000.0))
                                : "?";
                            adapters[idx] = mac + " @ " + speedText;
                        }
                    }
                }

                var parts = new List<string>();
                using (var configResults = WmiHelper.Query(scope,
                    "SELECT Index, IPAddress, MACAddress, IPEnabled FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True"))
                {
                    foreach (ManagementObject mo in configResults)
                    {
                        using (mo)
                        {
                            var idx = WmiHelper.GetProp(mo, "Index");
                            var mac = WmiHelper.GetProp(mo, "MACAddress") ?? string.Empty;
                            var ips = mo["IPAddress"] as string[];
                            var ip = ips != null && ips.Length > 0
                                ? string.Join(",", ips.Where(x => x != null && x.Contains(".")))
                                : string.Empty;

                            string speedText = string.Empty;
                            if (!string.IsNullOrEmpty(idx) && adapters.ContainsKey(idx))
                            {
                                var raw = adapters[idx];
                                var at = raw.LastIndexOf(" @ ", StringComparison.Ordinal);
                                if (at >= 0)
                                    speedText = raw.Substring(at + 3);
                            }

                            if (string.IsNullOrEmpty(ip) && string.IsNullOrEmpty(mac)) continue;
                            parts.Add(string.Format("IP:{0} MAC:{1} Hız:{2}",
                                string.IsNullOrEmpty(ip) ? "-" : ip,
                                string.IsNullOrEmpty(mac) ? "-" : mac,
                                string.IsNullOrEmpty(speedText) ? "?" : speedText));
                        }
                    }
                }

                if (parts.Count == 0 && adapters.Count > 0)
                {
                    foreach (var kv in adapters.Take(3))
                        parts.Add("MAC/Hız: " + kv.Value);
                }

                row.NetworkInfo = parts.Count > 0 ? string.Join(" || ", parts.Take(4)) : "N/A";
            }
            catch (Exception ex)
            {
                row.NetworkInfo = "Hata: " + ex.Message;
            }
        }

        private static void CollectLastUpdate(HostResult row, ManagementScope scope)
        {
            try
            {
                DateTime? latest = null;
                string latestId = null;
                using (var results = WmiHelper.Query(scope,
                    "SELECT HotFixID, InstalledOn FROM Win32_QuickFixEngineering"))
                {
                    foreach (ManagementObject mo in results)
                    {
                        using (mo)
                        {
                            var installed = WmiHelper.GetProp(mo, "InstalledOn");
                            DateTime dt;
                            if (TryParseInstalledOn(installed, out dt))
                            {
                                if (!latest.HasValue || dt > latest.Value)
                                {
                                    latest = dt;
                                    latestId = WmiHelper.GetProp(mo, "HotFixID");
                                }
                            }
                        }
                    }
                }

                row.LastWindowsUpdate = latest.HasValue
                    ? latest.Value.ToString("yyyy-MM-dd") + (latestId != null ? " (" + latestId + ")" : string.Empty)
                    : "N/A";
            }
            catch (Exception ex)
            {
                row.LastWindowsUpdate = "Hata: " + ex.Message;
            }
        }

        // Önce yerel kültür, olmazsa invariant (uzak makinenin tarih biçimi farklı olabilir).
        private static bool TryParseInstalledOn(string value, out DateTime dt)
        {
            dt = default(DateTime);
            if (string.IsNullOrWhiteSpace(value)) return false;
            if (DateTime.TryParse(value, out dt)) return true;
            return DateTime.TryParse(value, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out dt);
        }

        private static void CollectTopProcesses(HostResult row, ManagementScope scope)
        {
            try
            {
                // Aynı isimli süreçleri doğrudan toplayarak birleştir (ara liste + GroupBy yok).
                var totals = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);

                // Gereksiz nesneler WMI tarafında elenir; ağ üzerinden daha az veri taşınır.
                using (var results = WmiHelper.Query(scope,
                    "SELECT Name, WorkingSetSize FROM Win32_Process WHERE Name <> 'System Idle Process' AND Name <> 'System'"))
                {
                    foreach (ManagementObject mo in results)
                    {
                        using (mo)
                        {
                            var name = WmiHelper.GetProp(mo, "Name") ?? "?";
                            if (string.Equals(name, "System Idle Process", StringComparison.OrdinalIgnoreCase)
                                || string.Equals(name, "System", StringComparison.OrdinalIgnoreCase))
                                continue;

                            ulong ws;
                            ulong.TryParse(WmiHelper.GetProp(mo, "WorkingSetSize"), out ws);

                            long current;
                            totals.TryGetValue(name, out current);
                            totals[name] = current + (long)ws;
                        }
                    }
                }

                var grouped = totals
                    .OrderByDescending(kv => kv.Value)
                    .Take(3)
                    .Select(kv => string.Format("{0} ({1:0.0} MB)", kv.Key, kv.Value / 1024.0 / 1024.0));

                row.TopProcesses = string.Join(" | ", grouped);
                if (string.IsNullOrWhiteSpace(row.TopProcesses))
                    row.TopProcesses = "N/A";
            }
            catch (Exception ex)
            {
                row.TopProcesses = "Hata: " + ex.Message;
            }
        }

        public static string ToUncPath(string computerName, string localPath)
        {
            if (string.IsNullOrWhiteSpace(localPath))
                return null;

            // Güvenlik: geçersiz makine adı UNC yoluna sızdırılamaz (ör. "host\D$\..." gibi).
            if (!IsValidHostName(computerName))
                return null;

            // Güvenlik: kontrol karakterleri (satır sonu, null vb.) içeren yollar reddedilir.
            foreach (var ch in localPath)
            {
                if (ch < ' ') return null;
            }

            var path = localPath.Trim();

            var m = DriveLetterPathRegex.Match(path);
            if (m.Success)
                return @"\\" + computerName + @"\" + m.Groups[1].Value + @"$\" + m.Groups[2].Value.Replace('/', '\\');

            // Güvenlik: UNC yolu yalnızca hedef makineyi gösteriyorsa kabul edilir.
            // Aksi halde bu makine başka bir sunucuya (NTLM kimlik doğrulamasıyla) bağlanmaya zorlanabilir.
            if (path.StartsWith(@"\\", StringComparison.Ordinal))
            {
                var rest = path.Substring(2);
                int sep = rest.IndexOf('\\');
                var host = sep < 0 ? rest : rest.Substring(0, sep);
                return string.Equals(host, computerName, StringComparison.OrdinalIgnoreCase) ? path : null;
            }

            // Göreli / tanınmayan yollar: analiz yapan makinenin yerel diskinde yanlış sonuç vermesin.
            return null;
        }

        public static string ReadRemoteRegistryValue(string computer, string regPath, string valueName)
        {
            if (string.IsNullOrWhiteSpace(regPath) || string.IsNullOrWhiteSpace(valueName))
                return "N/A";

            if (!IsValidHostName(computer))
                return "Hata: Geçersiz bilgisayar adı";

            try
            {
                // 1. Root Hive tespiti (HKLM veya HKU) - tek regex ile
                RegistryHive hive;
                string subKeyPath = regPath.Trim();

                var hm = RegHiveRegex.Match(subKeyPath);
                if (!hm.Success)
                    return "Desteklenmiyor (sadece HKLM/HKU)";

                hive = hm.Groups["lm"].Success ? RegistryHive.LocalMachine : RegistryHive.Users;
                subKeyPath = subKeyPath.Substring(hm.Length);

                // 2. Uzak makineye doğrudan .NET Registry API'si ile bağlan (Process başlatma yok!)
                using (var baseKey = RegistryKey.OpenRemoteBaseKey(hive, computer, RegistryView.Default))
                using (var subKey = baseKey.OpenSubKey(subKeyPath, writable: false)) // Sadece OKUMA modu (Güvenli)
                {
                    if (subKey == null)
                        return "Bulunamadı";

                    var val = subKey.GetValue(valueName, null);
                    if (val == null)
                        return "Bulunamadı";

                    // MultiString (REG_MULTI_SZ) veya Byte dizileri kontrolü
                    if (val is string[])
                        return string.Join(", ", (string[])val);
                    if (val is byte[])
                        return BitConverter.ToString((byte[])val);

                    return val.ToString();
                }
            }
            catch (Exception ex)
            {
                return "Hata: " + ex.Message;
            }
        }

        private static bool IsValidHostName(string name)
        {
            return !string.IsNullOrEmpty(name)
                && HostNameRegex.IsMatch(name)
                && name.IndexOf("..", StringComparison.Ordinal) < 0;
        }

        // WQL yalnızca sorgu destekler, ancak yine de salt-okunur SELECT dışındakileri reddet.
        private static bool IsSelectQuery(string query)
        {
            if (string.IsNullOrWhiteSpace(query)) return false;
            var q = query.TrimStart();
            return q.StartsWith("SELECT ", StringComparison.OrdinalIgnoreCase)
                || q.StartsWith("SELECT\t", StringComparison.OrdinalIgnoreCase)
                || q.StartsWith("SELECT\r", StringComparison.OrdinalIgnoreCase)
                || q.StartsWith("SELECT\n", StringComparison.OrdinalIgnoreCase);
        }
    }
}