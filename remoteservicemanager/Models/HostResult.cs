using System;
using System.Collections.Generic;

namespace RemoteServiceManager.Models
{
    public class HostResult
    {
        public const string PendingText = "Bekleniyor";

        public string ComputerName { get; set; }
        public string Hostname { get; set; }
        public bool IsPending { get; set; }
        public bool IsOnline { get; set; }
        public string OnlineStatus { get; set; } = string.Empty;
        public string Running { get; set; } = string.Empty;
        public string NotRunning { get; set; } = string.Empty;
        public string LastShutdown { get; set; } = "N/A";
        public string FilePathControl { get; set; } = "N/A";
        public string RegistryValue { get; set; } = "N/A";
        public string ErrorMessage { get; set; } = string.Empty;

        public string Uptime { get; set; } = "N/A";
        public string LastBootTime { get; set; } = "N/A";
        public string DiskInfo { get; set; } = "N/A";
        public string LoggedInUser { get; set; } = "N/A";
        public string OsBuild { get; set; } = "N/A";
        public string NetworkInfo { get; set; } = "N/A";
        public string LastWindowsUpdate { get; set; } = "N/A";
        public string MemoryInfo { get; set; } = "N/A";
        public string TopProcesses { get; set; } = "N/A";

        public string Bsod { get; set; } = "N/A";
        public string BsodFileCount { get; set; } = "N/A";
        public string PendingUpdates { get; set; } = "N/A";
        public string UpdateSummary { get; set; } = "N/A";

        public string CcmExecStatus { get; set; } = "N/A";
        public string WuauservStatus { get; set; } = "N/A";
        public string BitsStatus { get; set; } = "N/A";
        public string CryptSvcStatus { get; set; } = "N/A";
        public string LastChecked { get; set; } = string.Empty;

        public List<ServiceStatusItem> ServiceDetail { get; set; } = new List<ServiceStatusItem>();
        public List<string> PatchDetailLines { get; set; } = new List<string>();
        public Dictionary<string, string> CustomValues { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public static HostResult CreatePending(string computerName)
        {
            var cn = (computerName ?? string.Empty).Trim();
            return new HostResult
            {
                ComputerName = cn,
                Hostname = cn,
                IsPending = true
            };
        }

        public static HostResult CreateError(string computerName, string message, bool patchMode = false)
        {
            var cn = (computerName ?? string.Empty).Trim();
            return new HostResult
            {
                ComputerName = cn,
                Hostname = cn,
                IsOnline = false,
                OnlineStatus = patchMode ? "Offline" : "Offline",
                ErrorMessage = message ?? string.Empty,
                LastChecked = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            };
        }

        public string GetValue(string key)
        {
            if (string.IsNullOrEmpty(key)) return string.Empty;
            if (IsPending)
            {
                if (string.Equals(key, "ComputerName", StringComparison.OrdinalIgnoreCase))
                    return ComputerName ?? string.Empty;
                if (string.Equals(key, "Hostname", StringComparison.OrdinalIgnoreCase))
                    return Hostname ?? ComputerName ?? string.Empty;
                return PendingText;
            }
            if (CustomValues != null && CustomValues.ContainsKey(key))
                return CustomValues[key];

            switch (key)
            {
                case "ComputerName": return ComputerName;
                case "Hostname": return Hostname;
                case "IsOnline": return IsOnline.ToString();
                case "PatchOnline": return string.IsNullOrEmpty(OnlineStatus) ? IsOnline.ToString() : OnlineStatus;
                case "Running": return Running;
                case "NotRunning": return NotRunning;
                case "LastShutdown":
                case "PatchLastShutdown": return LastShutdown;
                case "FilePathControl": return FilePathControl;
                case "RegistryValue": return RegistryValue;
                case "Uptime": return Uptime;
                case "LastBootTime": return LastBootTime;
                case "DiskInfo": return DiskInfo;
                case "LoggedInUser":
                case "PatchLoggedOnUser": return LoggedInUser;
                case "OsBuild":
                case "PatchOsBuild": return OsBuild;
                case "NetworkInfo": return NetworkInfo;
                case "LastWindowsUpdate": return LastWindowsUpdate;
                case "MemoryInfo": return MemoryInfo;
                case "TopProcesses": return TopProcesses;
                case "Bsod":
                case "PatchBsod": return Bsod;
                case "PatchBsodCount": return BsodFileCount;
                case "PendingUpdates": return PendingUpdates;
                case "UpdateSummary": return string.IsNullOrEmpty(UpdateSummary) || UpdateSummary == "N/A" ? PendingUpdates : UpdateSummary;
                case "CcmExec": return CcmExecStatus;
                case "Wuauserv": return WuauservStatus;
                case "Bits": return BitsStatus;
                case "CryptSvc": return CryptSvcStatus;
                case "LastChecked": return LastChecked;
                case "ErrorMessage": return ErrorMessage;
                default:
                    if (key.StartsWith("Custom_", StringComparison.OrdinalIgnoreCase) && CustomValues != null)
                    {
                        string v;
                        if (CustomValues.TryGetValue(key, out v)) return v;
                    }
                    return string.Empty;
            }
        }
    }
}
