using System;
using System.IO;
using System.Net.NetworkInformation;

namespace RemoteServiceManager.Services
{
    /// <summary>
    /// Uzak makine erişilebilirlik kontrolleri.
    /// Ping tek başına yeterli değildir (DNS kaydı / false-positive ICMP);
    /// admin paylaşımı (C$) ile gerçekten ulaşılabilirlik doğrulanır.
    /// </summary>
    public static class HostReachability
    {
        public static bool PingOk(string computerName, int timeoutMs)
        {
            try
            {
                using (var ping = new Ping())
                {
                    var reply = ping.Send(computerName, Math.Max(500, timeoutMs));
                    return reply != null && reply.Status == IPStatus.Success;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// PowerShell'deki Get-Item \\bilgisayar\C$ eşdeğeri.
        /// Yanıt yoksa / erişilemiyorsa false — makine çevrimdışı kabul edilir.
        /// </summary>
        public static bool CanAccessAdminShare(string computerName)
        {
            var name = (computerName ?? string.Empty).Trim();
            if (name.Length == 0)
                return false;

            var root = @"\\" + name + @"\C$";
            try
            {
                // Get-Item gibi gerçek bir erişim denemesi (Exists bazen yanıltıcı olabilir)
                Directory.GetFileSystemEntries(root);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
