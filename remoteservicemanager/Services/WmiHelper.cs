using System;
using System.Management;

namespace RemoteServiceManager.Services
{
    public static class WmiHelper
    {
        public static ManagementScope Connect(string computerName)
        {
            var path = new ManagementPath
            {
                Server = computerName,
                NamespacePath = @"root\cimv2"
            };

            var options = new ConnectionOptions
            {
                Impersonation = ImpersonationLevel.Impersonate,
                EnablePrivileges = true,
                Timeout = TimeSpan.FromSeconds(30)
            };

            var scope = new ManagementScope(path, options);
            scope.Connect();
            return scope;
        }

        public static ManagementObjectCollection Query(ManagementScope scope, string wql)
        {
            using (var searcher = new ManagementObjectSearcher(scope, new ObjectQuery(wql)))
            {
                searcher.Options.Timeout = TimeSpan.FromSeconds(30);
                return searcher.Get();
            }
        }

        public static string GetProp(ManagementBaseObject obj, string name)
        {
            try
            {
                var v = obj[name];
                return v == null ? null : Convert.ToString(v);
            }
            catch
            {
                return null;
            }
        }

        public static DateTime? ParseWmiDate(string wmiDate)
        {
            if (string.IsNullOrWhiteSpace(wmiDate))
                return null;
            try
            {
                return ManagementDateTimeConverter.ToDateTime(wmiDate);
            }
            catch
            {
                return null;
            }
        }
    }
}
