using System;

namespace RemoteServiceManager.Models
{
    public class FixResultItem
    {
        public string Hostname { get; set; }
        public string ServiceName { get; set; }
        public string OncekiDurum { get; set; }
        public string SonrakiDurum { get; set; }
        public string Aciklama { get; set; }
    }

    public class HostFixSummary
    {
        public string Hostname { get; set; }
        public string Status { get; set; }
    }
}
