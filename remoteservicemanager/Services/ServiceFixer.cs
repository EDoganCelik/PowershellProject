using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using RemoteServiceManager.Models;

namespace RemoteServiceManager.Services
{
    public class ServiceFixer
    {
        private static readonly TimeSpan StartTimeout = TimeSpan.FromSeconds(30);

        public List<FixResultItem> Fix(
            string computerName,
            IList<ServiceStatusItem> serviceDetail)
        {
            var results = new List<FixResultItem>();
            if (serviceDetail == null)
                return results;

            foreach (var s in serviceDetail)
            {
                if (s.Status == "Running")
                {
                    results.Add(new FixResultItem
                    {
                        Hostname = computerName,
                        ServiceName = s.Name,
                        OncekiDurum = s.Status,
                        SonrakiDurum = s.Status,
                        Aciklama = "Zaten çalışıyordu, işlem yapılmadı."
                    });
                    continue;
                }

                if (s.Status == "NotFound")
                {
                    results.Add(new FixResultItem
                    {
                        Hostname = computerName,
                        ServiceName = s.Name,
                        OncekiDurum = s.Status,
                        SonrakiDurum = s.Status,
                        Aciklama = "Servis bulunamadı, işlem yapılmadı."
                    });
                    continue;
                }

                try
                {
                    using (var sc = new ServiceController(s.Name, computerName))
                    {
                        sc.Start();
                        try
                        {
                            sc.WaitForStatus(ServiceControllerStatus.Running, StartTimeout);
                        }
                        catch (System.ServiceProcess.TimeoutException)
                        {
                            // Zaman aşımında mevcut durumu raporla.
                        }

                        sc.Refresh();
                        var newStatus = sc.Status.ToString();
                        var desc = newStatus == "Running"
                            ? s.Status + " durumundaydı, Running yapıldı."
                            : "Başlatıldı ama durum hala " + newStatus + ".";

                        results.Add(new FixResultItem
                        {
                            Hostname = computerName,
                            ServiceName = s.Name,
                            OncekiDurum = s.Status,
                            SonrakiDurum = newStatus,
                            Aciklama = desc
                        });
                    }
                }
                catch (Exception ex)
                {
                    results.Add(new FixResultItem
                    {
                        Hostname = computerName,
                        ServiceName = s.Name,
                        OncekiDurum = s.Status,
                        SonrakiDurum = "Bilinmiyor",
                        Aciklama = "Hata: " + ex.Message
                    });
                }
            }

            return results;
        }

        public static string GetHostSummary(IEnumerable<FixResultItem> hostResults)
        {
            var list = hostResults?.ToList() ?? new List<FixResultItem>();
            var fixedOnes = list.Where(x => x.OncekiDurum != "Running" && x.OncekiDurum != "NotFound" && x.SonrakiDurum == "Running").ToList();
            var stillStopped = list.Where(x => x.OncekiDurum != "Running" && x.OncekiDurum != "NotFound" && x.SonrakiDurum != "Running").ToList();
            var notFound = list.Where(x => x.OncekiDurum == "NotFound").ToList();

            var parts = new List<string>();
            if (fixedOnes.Count > 0)
                parts.Add(string.Join(",", fixedOnes.Select(x => x.ServiceName)) + " stopped durumda running yapıldı");
            if (stillStopped.Count > 0)
                parts.Add(string.Join(",", stillStopped.Select(x => x.ServiceName)) + " stopped durumda kaldı");
            if (notFound.Count > 0)
                parts.Add(string.Join(",", notFound.Select(x => x.ServiceName)) + " bulunamadı");

            if (parts.Count == 0)
                return "Tüm servisler çalışıyor.";
            return string.Join(", ", parts) + ".";
        }
    }
}
