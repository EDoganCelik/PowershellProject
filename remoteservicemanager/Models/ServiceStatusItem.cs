using System;

namespace RemoteServiceManager.Models
{
    [Serializable]
    public class ServiceStatusItem
    {
        public string Name { get; set; }
        public string Status { get; set; }
    }
}
