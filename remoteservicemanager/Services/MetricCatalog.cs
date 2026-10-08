using System;
using System.Collections.Generic;
using System.Linq;
using RemoteServiceManager.Models;

namespace RemoteServiceManager.Services
{
    public static class MetricCatalog
    {
        public const string Hostname = "Hostname";
        public const string IsOnline = "IsOnline";
        public const string Running = "Running";
        public const string NotRunning = "NotRunning";
        public const string LastShutdown = "LastShutdown";
        public const string FilePathControl = "FilePathControl";
        public const string RegistryValue = "RegistryValue";
        public const string Uptime = "Uptime";
        public const string LastBootTime = "LastBootTime";
        public const string DiskInfo = "DiskInfo";
        public const string LoggedInUser = "LoggedInUser";
        public const string OsBuild = "OsBuild";
        public const string NetworkInfo = "NetworkInfo";
        public const string LastWindowsUpdate = "LastWindowsUpdate";
        public const string MemoryInfo = "MemoryInfo";
        public const string TopProcesses = "TopProcesses";
        public const string Bsod = "Bsod";
        public const string PendingUpdates = "PendingUpdates";
        public const string ErrorMessage = "ErrorMessage";

        // Patch-specific
        public const string PatchOnline = "PatchOnline";
        public const string PatchOsBuild = "PatchOsBuild";
        public const string PatchBsod = "PatchBsod";
        public const string PatchBsodCount = "PatchBsodCount";
        public const string PatchLoggedOnUser = "PatchLoggedOnUser";
        public const string PatchLastShutdown = "PatchLastShutdown";
        public const string UpdateSummary = "UpdateSummary";
        public const string CcmExec = "CcmExec";
        public const string Wuauserv = "Wuauserv";
        public const string Bits = "Bits";
        public const string CryptSvc = "CryptSvc";
        public const string LastChecked = "LastChecked";

        public static IList<MetricDefinition> GetBuiltInServiceMetrics()
        {
            return new List<MetricDefinition>
            {
                B(IsOnline, "Online durumu", "Online?", MetricScope.Both, true),
                B(Running, "Çalışan servisler", "Çalışan Servisler", MetricScope.Service, true),
                B(NotRunning, "Çalışmayan servisler", "Çalışmayan Servisler", MetricScope.Service, true),
                B(LastShutdown, "Son kapanma", "Son Kapanma", MetricScope.Both, true),
                B(Uptime, "Uptime", "Uptime", MetricScope.Service, true),
                B(LastBootTime, "Last boot time", "Last Boot", MetricScope.Service, true),
                B(DiskInfo, "Disk alanı", "Disk", MetricScope.Service, true),
                B(LoggedInUser, "Oturum kullanıcısı", "Oturum", MetricScope.Both, true),
                B(OsBuild, "OS sürüm / build", "OS Build", MetricScope.Both, true),
                B(NetworkInfo, "Ağ IP/MAC/hız", "Ağ", MetricScope.Service, false),
                B(LastWindowsUpdate, "Son Windows Update (QFE)", "Son QFE", MetricScope.Service, false),
                B(MemoryInfo, "RAM bilgisi", "RAM", MetricScope.Service, true),
                B(TopProcesses, "İlk 3 süreç", "Top Process", MetricScope.Service, false),
                B(FilePathControl, "Dosya sürümü (varsayılan yol)", "Dosya", MetricScope.Service, false),
                B(RegistryValue, "Registry değeri (varsayılan)", "Registry", MetricScope.Service, false),
                B(Bsod, "Mavi ekran (Minidump)", "Mavi Ekran", MetricScope.Both, true),
                B(PendingUpdates, "Bekleyen güncelleme (CCM)", "Bekleyen Update", MetricScope.Both, true),
                B(ErrorMessage, "Hata mesajı", "Hata", MetricScope.Both, true)
            };
        }

        public static IList<MetricDefinition> GetBuiltInPatchMetrics()
        {
            return new List<MetricDefinition>
            {
                B(PatchOnline, "Erişilebilirlik (DNS/Ping/C$)", "Online", MetricScope.Patch, true),
                B(PatchOsBuild, "OS Build (UBR)", "OS Build", MetricScope.Patch, true),
                B(PatchBsod, "Mavi ekran", "Mavi Ekran", MetricScope.Patch, true),
                B(PatchBsodCount, "Minidump sayısı", "BSOD Dosya", MetricScope.Patch, true),
                B(PatchLoggedOnUser, "Oturum kullanıcısı", "Oturum", MetricScope.Patch, true),
                B(PatchLastShutdown, "Son kapanma", "Son Kapanma", MetricScope.Patch, true),
                B(UpdateSummary, "Bekleyen güncelleme özeti", "Bekleyen Update", MetricScope.Patch, true),
                B(PendingUpdates, "Bekleyen güncelleme özeti", "Bekleyen Update", MetricScope.Patch, false), // alias avoid dup - disabled by default in patch list if UpdateSummary on
                B(CcmExec, "ccmexec servisi", "ccmexec", MetricScope.Patch, true),
                B(Wuauserv, "wuauserv servisi", "wuauserv", MetricScope.Patch, true),
                B(Bits, "bits servisi", "bits", MetricScope.Patch, true),
                B(CryptSvc, "cryptsvc servisi", "cryptsvc", MetricScope.Patch, true),
                B(LastChecked, "Son kontrol zamanı", "Son Kontrol", MetricScope.Patch, true)
            };
        }

        public static MetricPreferences DefaultServicePreferences()
        {
            return new MetricPreferences
            {
                EnabledBuiltInIds = GetBuiltInServiceMetrics().Where(m => m.Enabled).Select(m => m.Id).ToList(),
                CustomMetrics = new List<MetricDefinition>()
            };
        }

        public static MetricPreferences DefaultPatchPreferences()
        {
            var ids = GetBuiltInPatchMetrics()
                .Where(m => m.Enabled && m.Id != PendingUpdates)
                .Select(m => m.Id)
                .ToList();
            return new MetricPreferences
            {
                EnabledBuiltInIds = ids,
                CustomMetrics = new List<MetricDefinition>()
            };
        }

        public static IList<MetricDefinition> ResolveActive(MetricPreferences prefs, MetricScope scope)
        {
            prefs = prefs ?? (scope == MetricScope.Patch ? DefaultPatchPreferences() : DefaultServicePreferences());
            var enabled = new HashSet<string>(prefs.EnabledBuiltInIds ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
            var builtIn = (scope == MetricScope.Patch ? GetBuiltInPatchMetrics() : GetBuiltInServiceMetrics())
                .Where(m => enabled.Contains(m.Id))
                .Select(m =>
                {
                    m.Enabled = true;
                    return m;
                })
                .ToList();

            var custom = (prefs.CustomMetrics ?? new List<MetricDefinition>())
                .Where(m => m != null && m.Enabled && (m.Scope == scope || m.Scope == MetricScope.Both || scope == MetricScope.Both))
                .ToList();

            var result = new List<MetricDefinition>();
            result.AddRange(builtIn);
            result.AddRange(custom);
            return result;
        }

        /// <summary>CheckOptions (eski) → MetricPreferences geçişi.</summary>
        public static MetricPreferences FromLegacy(CheckOptions opt)
        {
            var prefs = DefaultServicePreferences();
            if (opt == null) return prefs;

            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            void Add(string id, bool on) { if (on) set.Add(id); }

            Add(IsOnline, true);
            Add(Running, opt.CheckServices);
            Add(NotRunning, opt.CheckServices);
            Add(LastShutdown, opt.CheckLastShutdown);
            Add(FilePathControl, opt.CheckFileVersion);
            Add(RegistryValue, opt.CheckRegistry);
            Add(Uptime, opt.CheckUptime);
            Add(LastBootTime, opt.CheckUptime);
            Add(DiskInfo, opt.CheckDisk);
            Add(LoggedInUser, opt.CheckLoggedInUser);
            Add(OsBuild, opt.CheckOsBuild);
            Add(NetworkInfo, opt.CheckNetwork);
            Add(LastWindowsUpdate, opt.CheckLastWindowsUpdate);
            Add(MemoryInfo, opt.CheckMemory);
            Add(TopProcesses, opt.CheckTopProcesses);
            Add(Bsod, true);
            Add(PendingUpdates, true);
            Add(ErrorMessage, true);

            prefs.EnabledBuiltInIds = set.ToList();
            return prefs;
        }

        private static MetricDefinition B(string id, string display, string header, MetricScope scope, bool enabled)
        {
            return new MetricDefinition
            {
                Id = id,
                DisplayName = display,
                ColumnHeader = header,
                Scope = scope,
                Kind = MetricKind.BuiltIn,
                Enabled = enabled
            };
        }
    }
}
