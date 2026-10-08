using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Net;
using System.ServiceProcess;
using System.Text.RegularExpressions;
using RemoteServiceManager.Models;

namespace RemoteServiceManager.Services
{
    /// <summary>
    /// patchtakip.txt altyapısının C# karşılığı (WMI/DCOM, UNC, Get-Service).
    /// </summary>
    public class PatchChecker
    {
        public static readonly string[] PatchServices = { "ccmexec", "wuauserv", "bits", "cryptsvc" };

        public static readonly Dictionary<string, string> CcmSchedules = new Dictionary<string, string>
        {
            { "Update Trigger", "{00000000-0000-0000-0000-000000000108}" },
            { "Update Scan Cycle", "{00000000-0000-0000-0000-000000000113}" },
            { "Machine Cycle", "{00000000-0000-0000-0000-000000000021}" },
            { "Updates Deployment Cycle", "{00000000-0000-0000-0000-000000000108}" }
        };

        private static readonly HashSet<string> PendingStates = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Available", "Submitted", "Detecting", "Downloading", "WaitInstall",
            "Installing", "PendingUpdate", "WaitingRetry"
        };

        // ------------------------------------------------------------------
        // Güvenlik: girdi doğrulama (hostname / UNC / WQL / schedule id)
        // ------------------------------------------------------------------

        /// <summary>WMI işlemleri için varsayılan zaman aşımı (donan hedeflerde thread'in takılı kalmasını önler).</summary>
        private static readonly TimeSpan WmiTimeout = TimeSpan.FromSeconds(30);

        // NetBIOS / FQDN / IPv4. Ters eğik çizgi, eğik çizgi, iki nokta, boşluk vb. kabul edilmez
        // (UNC yolu ve ManagementPath içine enjeksiyonu engeller).
        private static readonly Regex HostNameRegex = new Regex(
            @"^[A-Za-z0-9_][A-Za-z0-9_\-\.]{0,252}$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex ScheduleIdRegex = new Regex(
            @"^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex WmiPropertyRegex = new Regex(
            @"^[A-Za-z_][A-Za-z0-9_]{0,127}$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex WqlQueryRegex = new Regex(
            @"^\s*(SELECT|ASSOCIATORS\s+OF|REFERENCES\s+OF)\b",
            RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        /// <summary>Geçerliyse kırpılmış bilgisayar adını, geçersizse null döner.</summary>
        private static string NormalizeHost(string computerName)
        {
            var name = (computerName ?? string.Empty).Trim();
            return HostNameRegex.IsMatch(name) ? name : null;
        }

        private static bool IsValidServiceName(string name)
        {
            return !string.IsNullOrWhiteSpace(name)
                   && name.Length <= 256
                   && name.IndexOfAny(new[] { '\\', '/', '\0' }) < 0;
        }

        private static bool HasPathTraversal(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return true;
            if (path.IndexOf('\0') >= 0) return true;
            return path.Split('\\', '/').Any(seg => seg == "..");
        }

        private static bool IsSafeWql(string query)
        {
            return !string.IsNullOrWhiteSpace(query)
                   && query.Length <= 2000
                   && query.IndexOf(';') < 0
                   && WqlQueryRegex.IsMatch(query);
        }

        // ------------------------------------------------------------------
        // Ortak WMI bağlantısı: tek yerde, zaman aşımlı, şifreli, en az yetkili
        // ------------------------------------------------------------------
        private static ManagementScope ConnectScope(string host, string namespacePath)
        {
            var path = new ManagementPath
            {
                Server = host,
                NamespacePath = namespacePath
            };
            var options = new ConnectionOptions
            {
                Impersonation = ImpersonationLevel.Impersonate,
                Authentication = AuthenticationLevel.PacketPrivacy, // DCOM trafiğini şifrele
                EnablePrivileges = false,                           // Bu sorgular için ayrıcalık gerekmiyor
                Timeout = WmiTimeout
            };
            var scope = new ManagementScope(path, options);
            scope.Connect();
            return scope;
        }

        public HostResult Check(string computerName, MetricPreferences prefs)
        {
            var active = MetricCatalog.ResolveActive(prefs, MetricScope.Patch);
            var need = new HashSet<string>(active.Where(m => !m.IsCustom).Select(m => m.Id), StringComparer.OrdinalIgnoreCase);
            var customs = active.Where(m => m.IsCustom).ToList();

            var row = new HostResult
            {
                ComputerName = computerName,
                Hostname = computerName,
                LastChecked = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
            };

            var reach = TestReachability(computerName);
            row.OnlineStatus = reach.Status;
            row.IsOnline = reach.CanQuery;

            if (!reach.CanQuery)
            {
                row.ErrorMessage = reach.Status;
                return row;
            }

            // TestReachability başarılıysa ad doğrulanmıştır.
            var host = NormalizeHost(computerName);

            ManagementScope cimv2 = null;
            try { cimv2 = WmiHelper.Connect(host); }
            catch (Exception ex)
            {
                row.ErrorMessage = "WMI: " + ex.Message;
                return row;
            }

            if (need.Contains(MetricCatalog.PatchOsBuild) || need.Contains(MetricCatalog.OsBuild))
                row.OsBuild = GetOsBuildWithUbr(host);

            if (need.Contains(MetricCatalog.PatchBsod) || need.Contains(MetricCatalog.PatchBsodCount) || need.Contains(MetricCatalog.Bsod))
            {
                var bsod = GetBsodStatus(host);
                row.BsodFileCount = bsod.Item1.ToString(CultureInfo.InvariantCulture);
                row.Bsod = bsod.Item1 > 2 ? "VAR" : "YOK";
            }

            if (need.Contains(MetricCatalog.PatchLoggedOnUser) || need.Contains(MetricCatalog.LoggedInUser))
                row.LoggedInUser = GetLoggedOnUser(cimv2, host);

            if (need.Contains(MetricCatalog.PatchLastShutdown) || need.Contains(MetricCatalog.LastShutdown))
                row.LastShutdown = GetLastShutdownWmi(cimv2);

            if (need.Contains(MetricCatalog.UpdateSummary) || need.Contains(MetricCatalog.PendingUpdates))
            {
                var details = GetPatchStatusDetails(host);
                row.UpdateSummary = details.Summary;
                row.PendingUpdates = details.Summary;
                row.PatchDetailLines = details.Lines;
            }

            if (need.Contains(MetricCatalog.CcmExec) || need.Contains(MetricCatalog.Wuauserv)
                || need.Contains(MetricCatalog.Bits) || need.Contains(MetricCatalog.CryptSvc))
            {
                CollectPatchServices(row, host);
            }

            foreach (var cm in customs)
                row.CustomValues[cm.ResultKey] = CollectCustom(host, cm, cimv2);

            return row;
        }

        public PatchDetails GetPatchStatusDetails(string computerName)
        {
            var result = new PatchDetails
            {
                Summary = "Bekleyen update yok",
                Lines = new List<string> { "(Bekleyen update kaydı yok)" }
            };

            var host = NormalizeHost(computerName);
            if (host == null)
            {
                result.Lines = new List<string> { "(Update kaydı bulunamadı veya erişilemedi)" };
                return result;
            }

            try
            {
                var scope = ConnectScope(host, @"root\ccm\ClientSDK");

                int pending = 0, errors = 0, complete = 0, total = 0;
                var lines = new List<string>();

                // Tek geçişte işle: önce listeye toplayıp sonra tekrar dönmek yerine her nesne
                // okunur okunmaz dispose edilir. SELECT * yerine yalnızca kullanılan kolonlar çekilir.
                using (var results = WmiHelper.Query(scope,
                    "SELECT Name, EvaluationState, PercentComplete, ErrorCode FROM CCM_SoftwareUpdate"))
                {
                    foreach (ManagementObject u in results)
                    {
                        using (u)
                        {
                            total++;

                            var name = WmiHelper.GetProp(u, "Name") ?? "?";
                            int state;
                            int.TryParse(WmiHelper.GetProp(u, "EvaluationState"), NumberStyles.Integer, CultureInfo.InvariantCulture, out state);
                            var stateText = MapEvaluationState(state);
                            var percent = WmiHelper.GetProp(u, "PercentComplete") ?? "";

                            // ErrorCode HRESULT olabilir (örn. 0x80070002 = 2147942402) ve int sınırını aşar.
                            long errCode;
                            long.TryParse(WmiHelper.GetProp(u, "ErrorCode"), NumberStyles.Integer, CultureInfo.InvariantCulture, out errCode);

                            if (PendingStates.Contains(stateText)) pending++;
                            if (errCode != 0 || string.Equals(stateText, "Error", StringComparison.OrdinalIgnoreCase))
                                errors++;
                            else if (string.Equals(stateText, "InstallComplete", StringComparison.OrdinalIgnoreCase))
                                complete++;

                            lines.Add(string.Format("- {0} | Durum={1} | Yuzde={2} | ErrorCode={3}",
                                name, stateText, percent, errCode));
                        }
                    }
                }

                if (total == 0)
                    return result;

                if (pending == 0 && errors == 0)
                {
                    result.Summary = "Bekleyen update yok";
                    result.Lines = lines.Count > 0 ? lines : new List<string> { "(Bekleyen update kaydı yok)" };
                }
                else
                {
                    result.Summary = string.Format(
                        "Bekleyen update: {0} | Hatalı update: {1} | Tamamlanan: {2} | Toplam: {3}",
                        pending, errors, complete, total);
                    result.Lines = lines;
                }
            }
            catch
            {
                result.Summary = "Bekleyen update yok";
                result.Lines = new List<string> { "(Update kaydı bulunamadı veya erişilemedi)" };
            }

            return result;
        }

        public string TriggerCcmSchedule(string computerName, string scheduleId)
        {
            var host = NormalizeHost(computerName);
            if (host == null)
                return "HATA: Geçersiz bilgisayar adı";

            var id = (scheduleId ?? string.Empty).Trim();
            if (!ScheduleIdRegex.IsMatch(id))
                return "HATA: Geçersiz schedule ID";

            try
            {
                var scope = ConnectScope(host, @"root\ccm");
                using (var cls = new ManagementClass(scope, new ManagementPath("SMS_Client"), null))
                {
                    var inParams = cls.GetMethodParameters("TriggerSchedule");
                    inParams["sScheduleID"] = id;
                    cls.InvokeMethod("TriggerSchedule", inParams, null);
                }
                return "OK";
            }
            catch (Exception ex)
            {
                return "HATA: " + ex.Message;
            }
        }

        public static string MapEvaluationState(int state)
        {
            switch (state)
            {
                case 0: return "None";
                case 1: return "Available";
                case 2: return "Submitted";
                case 3: return "Detecting";
                case 6: return "Downloading";
                case 7: return "WaitInstall";
                case 8: return "Installing";
                case 9: return "PendingSoftReboot";
                case 10: return "PendingHardReboot";
                case 11: return "WaitReboot";
                case 12: return "Verifying";
                case 13: return "InstallComplete";
                case 14: return "Error";
                case 15: return "WaitServiceWindow";
                case 16: return "WaitUserLogon";
                case 17: return "WaitUserLogoff";
                case 18: return "WaitJobUserLogon";
                case 19: return "WaitUserReconnect";
                case 20: return "PendingUserLogoff";
                case 21: return "PendingUpdate";
                case 22: return "WaitingRetry";
                case 23: return "WaitPresModeOff";
                case 24: return "WaitForOrchestration";
                default: return "Bilinmiyor(" + state + ")";
            }
        }

        public ReachabilityResult TestReachability(string computerName)
        {
            // Geçersiz/şüpheli ad (UNC, yol ayracı vb.) hiç ağa çıkmadan elenir.
            // Mevcut arayüz davranışı korunsun diye "DNS'te yok" durumu döndürülür.
            var name = NormalizeHost(computerName);
            if (name == null)
                return new ReachabilityResult { Status = "DNS'te yok", CanQuery = false };

            bool inDns = false;
            try
            {
                var entry = Dns.GetHostEntry(name);
                inDns = entry != null && entry.AddressList != null && entry.AddressList.Length > 0;
            }
            catch
            {
                inDns = false;
            }

            if (!inDns)
                return new ReachabilityResult { Status = "DNS'te yok", CanQuery = false };

            bool pingOk = HostReachability.PingOk(name, 1000);
            if (!pingOk)
                return new ReachabilityResult { Status = "Offline", CanQuery = false, InDns = true, PingOk = false };

            // Ping yetmez: Get-Item \\pc\C$ eşdeğeri ile gerçek erişim doğrulanır.
            bool shareOk;
            try
            {
                shareOk = HostReachability.CanAccessAdminShare(name);
            }
            catch
            {
                shareOk = false;
            }

            if (!shareOk)
                return new ReachabilityResult { Status = "Offline", CanQuery = false, InDns = true, PingOk = true, WmiOk = false };

            return new ReachabilityResult
            {
                Status = "Evet",
                CanQuery = true,
                InDns = true,
                PingOk = true,
                WmiOk = true
            };
        }

        public static Tuple<int, bool> GetBsodStatus(string computerName)
        {
            var host = NormalizeHost(computerName);
            if (host == null)
                return Tuple.Create(0, false);

            var unc = @"\\" + host + @"\C$\Windows\Minidump";
            if (!Directory.Exists(unc))
                return Tuple.Create(0, false);
            try
            {
                // GetFiles tüm dosya adlarını dizi olarak belleğe alır; EnumerateFiles akış halinde sayar.
                var count = Directory.EnumerateFiles(unc).Count();
                return Tuple.Create(count, true);
            }
            catch
            {
                return Tuple.Create(0, false);
            }
        }

        /// <summary>
        /// Örnek: 10.0.20348.587 (Build 20348)
        /// StdRegProv: CurrentMajor/Minor + CurrentBuildNumber + UBR
        /// </summary>
        public static string GetOsBuildWithUbr(string computerName)
        {
            var host = NormalizeHost(computerName);
            if (host == null)
                return "Bilinmiyor";

            try
            {
                var scope = ConnectScope(host, @"root\default");
                using (var reg = new ManagementClass(scope, new ManagementPath("StdRegProv"), null))
                {
                    const uint hklm = 2147483650;
                    const string keyPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion";

                    var major = GetRegDword(reg, hklm, keyPath, "CurrentMajorVersionNumber");
                    var minor = GetRegDword(reg, hklm, keyPath, "CurrentMinorVersionNumber");
                    var build = GetRegString(reg, hklm, keyPath, "CurrentBuildNumber");
                    var ubr = GetRegDword(reg, hklm, keyPath, "UBR");

                    if (string.IsNullOrWhiteSpace(build))
                        return "Bilinmiyor";

                    string version;
                    if (major.HasValue && minor.HasValue)
                    {
                        version = major.Value + "." + minor.Value + "." + build;
                        if (ubr.HasValue)
                            version += "." + ubr.Value;
                    }
                    else
                    {
                        // Eski fallback: sadece build.ubr
                        version = ubr.HasValue ? (build + "." + ubr.Value) : build;
                    }

                    return version + " (Build " + build + ")";
                }
            }
            catch
            {
                return "Bilinmiyor";
            }
        }

        private static string GetRegString(ManagementClass reg, uint hklm, string keyPath, string valueName)
        {
            var p = reg.GetMethodParameters("GetStringValue");
            p["hDefKey"] = hklm;
            p["sSubKeyName"] = keyPath;
            p["sValueName"] = valueName;
            var r = reg.InvokeMethod("GetStringValue", p, null);
            return Convert.ToString(r["sValue"]);
        }

        private static uint? GetRegDword(ManagementClass reg, uint hklm, string keyPath, string valueName)
        {
            var p = reg.GetMethodParameters("GetDWORDValue");
            p["hDefKey"] = hklm;
            p["sSubKeyName"] = keyPath;
            p["sValueName"] = valueName;
            var r = reg.InvokeMethod("GetDWORDValue", p, null);
            if (r["uValue"] == null) return null;
            return Convert.ToUInt32(r["uValue"]);
        }

        private static string GetLoggedOnUser(ManagementScope scope, string computerName)
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
                            if (!string.IsNullOrWhiteSpace(u)) return u;
                        }
                    }
                }
            }
            catch { /* ignore */ }
            return "Yok";
        }

        private static string GetLastShutdownWmi(ManagementScope scope)
        {
            try
            {
                var since = ManagementDateTimeConverter.ToDmtfDateTime(DateTime.Now.AddDays(-60));
                var filter = "Logfile='System' AND (EventCode='1074' OR EventCode='6006' OR EventCode='6008') AND TimeGenerated >= '" + since + "'";
                DateTime? latest = null;
                // EventCode yalnızca filtrede kullanılıyor; SELECT'ten çıkarmak taşınan veriyi azaltır.
                using (var results = WmiHelper.Query(scope, "SELECT TimeGenerated FROM Win32_NTLogEvent WHERE " + filter))
                {
                    foreach (ManagementObject mo in results)
                    {
                        using (mo)
                        {
                            var t = WmiHelper.ParseWmiDate(WmiHelper.GetProp(mo, "TimeGenerated"));
                            if (t.HasValue && (!latest.HasValue || t.Value > latest.Value))
                                latest = t;
                        }
                    }
                }
                return latest.HasValue ? latest.Value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) : "Bulunamadı";
            }
            catch
            {
                return "N/A";
            }
        }

        private static void CollectPatchServices(HostResult row, string computer)
        {
            foreach (var name in PatchServices)
            {
                string status;
                try
                {
                    using (var sc = new ServiceController(name, computer))
                        status = sc.Status.ToString();
                }
                catch { status = ""; }

                switch (name.ToLowerInvariant())
                {
                    case "ccmexec": row.CcmExecStatus = status; break;
                    case "wuauserv": row.WuauservStatus = status; break;
                    case "bits": row.BitsStatus = status; break;
                    case "cryptsvc": row.CryptSvcStatus = status; break;
                }
            }
        }

        private static string CollectCustom(string computer, MetricDefinition m, ManagementScope scope)
        {
            try
            {
                switch (m.Kind)
                {
                    case MetricKind.Service:
                        if (!IsValidServiceName(m.Param1))
                            return "Hata: Geçersiz servis adı";
                        using (var sc = new ServiceController(m.Param1, computer))
                            return sc.Status.ToString();
                    case MetricKind.FileVersion:
                        if (HasPathTraversal(m.Param1))
                            return "Hata: Geçersiz dosya yolu";
                        return RemoteChecker.ReadFileVersion(computer, m.Param1);
                    case MetricKind.Registry:
                        return RemoteChecker.ReadRemoteRegistryValue(computer, m.Param1, m.Param2);
                    case MetricKind.WmiProperty:
                        if (scope == null) return "WMI yok";
                        // Kullanıcı tanımlı sorgu: yalnızca okuma amaçlı WQL ve düzgün property adı kabul edilir.
                        if (!IsSafeWql(m.Param1))
                            return "Hata: Geçersiz WQL sorgusu";
                        if (string.IsNullOrEmpty(m.Param2) || !WmiPropertyRegex.IsMatch(m.Param2))
                            return "Hata: Geçersiz özellik adı";
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
    }

    public class PatchDetails
    {
        public string Summary { get; set; }
        public List<string> Lines { get; set; }
    }

    public class ReachabilityResult
    {
        public bool InDns { get; set; }
        public bool PingOk { get; set; }
        public bool WmiOk { get; set; }
        public bool CanQuery { get; set; }
        public string Status { get; set; }
    }
}