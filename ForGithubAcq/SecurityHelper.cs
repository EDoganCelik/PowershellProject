using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.RegularExpressions;

namespace GetKnownIssuesForm
{
    internal static class SecurityHelper
    {
        public const int MaxUrlLength = 2048;
        public const int MaxDownloadBytes = 10 * 1024 * 1024;
        public const int DownloadTimeoutMs = 30000;
        public const int MaxRedirects = 5;
        public static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(5);

        private static readonly Regex AllowedHostRegex = new Regex(
            @"^(support|learn)\.microsoft\.com$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex SafeAnchorIdRegex = new Regex(
            @"^[A-Za-z0-9._-]+$",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        public static string ValidateAndNormalizeUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
                throw new ArgumentException("URL bos olamaz.");

            url = url.Trim();
            if (url.Length > MaxUrlLength)
                throw new InvalidOperationException("URL cok uzun.");

            if (!Uri.TryCreate(url, UriKind.Absolute, out Uri uri))
                throw new InvalidOperationException("Gecersiz URL formati.");

            if (!string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("Yalnizca HTTPS URL'leri desteklenir.");

            if (!string.IsNullOrEmpty(uri.UserInfo))
                throw new InvalidOperationException("Kimlik bilgisi iceren URL'ler kabul edilmez.");

            if (!uri.IsDefaultPort && uri.Port != 443)
                throw new InvalidOperationException("Standart disi portlar kabul edilmez.");

            string host = uri.Host;
            if (!AllowedHostRegex.IsMatch(host))
                throw new InvalidOperationException("Yalnizca Microsoft destek/ogrenme sayfalari desteklenir.");

            EnsureHostResolvesToPublicAddress(host);

            return uri.GetLeftPart(UriPartial.Path) +
                   uri.Query +
                   (string.IsNullOrEmpty(uri.Fragment) ? string.Empty : uri.Fragment);
        }

        public static string BuildSafeAnchorUrl(string baseUrl, string anchorId)
        {
            string normalized = ValidateAndNormalizeUrl(baseUrl);
            if (string.IsNullOrEmpty(anchorId))
                return normalized;

            if (!SafeAnchorIdRegex.IsMatch(anchorId))
                throw new InvalidOperationException("Gecersiz bolum kimligi (anchor id).");

            int fragmentIndex = normalized.IndexOf('#');
            if (fragmentIndex >= 0)
                normalized = normalized.Substring(0, fragmentIndex);

            return normalized + "#" + anchorId;
        }

        public static string HtmlEncode(string value)
        {
            return WebUtility.HtmlEncode(value ?? string.Empty);
        }

        public static string SanitizeHtmlFragment(string fragment)
        {
            if (string.IsNullOrEmpty(fragment))
                return string.Empty;

            string sanitized = fragment;

            sanitized = Regex.Replace(
                sanitized,
                @"<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>",
                string.Empty,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                RegexTimeout);

            sanitized = Regex.Replace(
                sanitized,
                @"<(iframe|object|embed|form|base|link|meta)\b[^>]*>.*?</\1>",
                string.Empty,
                RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant,
                RegexTimeout);

            sanitized = Regex.Replace(
                sanitized,
                @"<(iframe|object|embed|form|base|link|meta)\b[^>]*/?>",
                string.Empty,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                RegexTimeout);

            sanitized = Regex.Replace(
                sanitized,
                @"\s+on\w+\s*=\s*(""[^""]*""|'[^']*'|[^\s>]+)",
                string.Empty,
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                RegexTimeout);

            sanitized = Regex.Replace(
                sanitized,
                @"(href|src|action|formaction|background|xlink:href)\s*=\s*(""[^""]*""|'[^']*'|[^\s>]+)",
                match =>
                {
                    string attribute = match.Groups[1].Value;
                    string rawValue = match.Groups[2].Value.Trim().Trim('"', '\'');
                    if (IsDangerousUrlValue(rawValue))
                        return attribute + "=\"#\"";
                    return match.Value;
                },
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                RegexTimeout);

            return sanitized;
        }

        public static void OpenUrlInBrowser(string url)
        {
            string safeUrl = ValidateAndNormalizeUrl(url);
            Process.Start(new ProcessStartInfo
            {
                FileName = safeUrl,
                UseShellExecute = true
            });
        }

        public static string DownloadHtml(string url)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            string currentUrl = ValidateAndNormalizeUrl(url);

            for (int redirect = 0; redirect <= MaxRedirects; redirect++)
            {
                var request = (HttpWebRequest)WebRequest.Create(currentUrl);
                request.Method = "GET";
                request.AllowAutoRedirect = false;
                request.Timeout = DownloadTimeoutMs;
                request.ReadWriteTimeout = DownloadTimeoutMs;
                request.UserAgent = "GetKnownIssuesForm/1.0";
                request.Accept = "text/html";
                request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;

                using (var response = (HttpWebResponse)request.GetResponse())
                {
                    int statusCode = (int)response.StatusCode;
                    if (statusCode >= 300 && statusCode < 400)
                    {
                        string location = response.Headers["Location"];
                        if (string.IsNullOrWhiteSpace(location))
                            throw new InvalidOperationException("Gecersiz yonlendirme yaniti.");

                        currentUrl = ValidateAndNormalizeUrl(new Uri(new Uri(currentUrl), location).AbsoluteUri);
                        continue;
                    }

                    if (statusCode != 200)
                        throw new InvalidOperationException("Sayfa indirilemedi. HTTP " + statusCode);

                    string contentType = response.ContentType ?? string.Empty;
                    if (contentType.IndexOf("text/html", StringComparison.OrdinalIgnoreCase) < 0)
                        throw new InvalidOperationException("Yanit HTML degil.");

                    using (Stream stream = response.GetResponseStream())
                    {
                        return ReadLimitedHtml(stream);
                    }
                }
            }

            throw new InvalidOperationException("Cok fazla HTTP yonlendirmesi.");
        }

        private static string ReadLimitedHtml(Stream stream)
        {
            using (var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
            {
                var buffer = new char[8192];
                var builder = new StringBuilder();
                int read;
                while ((read = reader.Read(buffer, 0, buffer.Length)) > 0)
                {
                    if (builder.Length + read > MaxDownloadBytes)
                        throw new InvalidOperationException("Sayfa boyutu izin verilen sinirin ustunde.");

                    builder.Append(buffer, 0, read);
                }

                return builder.ToString();
            }
        }

        private static bool IsDangerousUrlValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            value = value.Trim();
            return value.StartsWith("javascript:", StringComparison.OrdinalIgnoreCase) ||
                   value.StartsWith("vbscript:", StringComparison.OrdinalIgnoreCase) ||
                   value.StartsWith("data:", StringComparison.OrdinalIgnoreCase) ||
                   value.StartsWith("file:", StringComparison.OrdinalIgnoreCase);
        }

        private static void EnsureHostResolvesToPublicAddress(string host)
        {
            IPAddress[] addresses;
            try
            {
                addresses = Dns.GetHostAddresses(host);
            }
            catch (SocketException)
            {
                throw new InvalidOperationException("Sunucu adresi cozumlenemedi.");
            }

            if (addresses == null || addresses.Length == 0)
                throw new InvalidOperationException("Sunucu adresi cozumlenemedi.");

            foreach (IPAddress address in addresses)
            {
                if (IsPrivateOrReservedAddress(address))
                    throw new InvalidOperationException("Yerel veya ozel ag adreslerine erisim engellendi.");
            }
        }

        private static bool IsPrivateOrReservedAddress(IPAddress address)
        {
            if (IPAddress.IsLoopback(address))
                return true;

            byte[] bytes = address.GetAddressBytes();
            if (address.AddressFamily == AddressFamily.InterNetwork)
            {
                if (bytes[0] == 10)
                    return true;
                if (bytes[0] == 172 && bytes[1] >= 16 && bytes[1] <= 31)
                    return true;
                if (bytes[0] == 192 && bytes[1] == 168)
                    return true;
                if (bytes[0] == 127)
                    return true;
                if (bytes[0] == 0)
                    return true;
                if (bytes[0] >= 224)
                    return true;
            }

            if (address.AddressFamily == AddressFamily.InterNetworkV6)
            {
                if (address.IsIPv6LinkLocal || address.IsIPv6SiteLocal)
                    return true;
            }

            return false;
        }
    }
}
