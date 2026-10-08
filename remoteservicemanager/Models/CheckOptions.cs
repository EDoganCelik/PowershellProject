using System;

namespace RemoteServiceManager.Models
{
    /// <summary>
    /// Kullanıcının seçtiği veri toplama seçenekleri.
    /// Kapalı olan metrikler için uzak makineye ilgili sorgu gönderilmez.
    /// </summary>
    [Serializable]
    public class CheckOptions
    {
        public bool CheckServices { get; set; } = true;
        public bool CheckLastShutdown { get; set; } = true;
        public bool CheckFileVersion { get; set; }
        public bool CheckRegistry { get; set; }

        public bool CheckUptime { get; set; } = true;
        public bool CheckDisk { get; set; } = true;
        public bool CheckLoggedInUser { get; set; } = true;
        public bool CheckOsBuild { get; set; } = true;
        public bool CheckNetwork { get; set; } = true;
        public bool CheckLastWindowsUpdate { get; set; } = true;
        public bool CheckMemory { get; set; } = true;
        public bool CheckTopProcesses { get; set; } = true;

        public int MaxParallel { get; set; } = 10;
        public int PingTimeoutMs { get; set; } = 1000;
    }
}
