using System;
using System.Text;
using System.Text.RegularExpressions;

namespace GetKnownIssuesForm
{
    internal static class KnownIssuesExtractor
    {
        public static KnownIssuesResult Extract(string url)
        {
            string normalizedUrl = SecurityHelper.ValidateAndNormalizeUrl(url);
            string html = SecurityHelper.DownloadHtml(normalizedUrl);

            var headingOpts = RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant;
            const string headingPattern = @"<h([1-4])[^>]*>(?:(?!</h\1>).)*?known\s+issues(?:(?!</h\1>).)*?</h\1>";
            var headingRegex = new Regex(headingPattern, headingOpts, SecurityHelper.RegexTimeout);
            Match headingMatch = headingRegex.Match(html);

            if (!headingMatch.Success)
                throw new InvalidOperationException("'Known issues' basligi bu sayfada bulunamadi.");

            int headingLevel = int.Parse(headingMatch.Groups[1].Value);
            int startIndex = headingMatch.Index;
            int searchFrom = startIndex + headingMatch.Length;

            var levelsBuilder = new StringBuilder();
            for (int i = 1; i <= headingLevel; i++)
                levelsBuilder.Append(i);
            string levels = levelsBuilder.ToString();

            var nextHeadingRegex = new Regex(
                "<h[" + levels + "][^>]*>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                SecurityHelper.RegexTimeout);
            Match nextMatch = nextHeadingRegex.Match(html, searchFrom);

            int endIndex;
            if (nextMatch.Success)
            {
                endIndex = nextMatch.Index;
            }
            else
            {
                int bodyEndIdx = html.IndexOf("</body>", searchFrom, StringComparison.OrdinalIgnoreCase);
                endIndex = bodyEndIdx >= 0 ? bodyEndIdx : html.Length;
            }

            string fragment = html.Substring(startIndex, endIndex - startIndex);
            fragment = SecurityHelper.SanitizeHtmlFragment(fragment);

            string headingOpenTag = Regex.Match(
                headingMatch.Value,
                @"^<h[1-4][^>]*>",
                RegexOptions.CultureInvariant,
                SecurityHelper.RegexTimeout).Value;
            Match idMatch = Regex.Match(
                headingOpenTag,
                @"id\s*=\s*""([^""]+)""",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                SecurityHelper.RegexTimeout);
            string anchorId = idMatch.Success ? idMatch.Groups[1].Value : null;
            string anchorUrl = SecurityHelper.BuildSafeAnchorUrl(normalizedUrl, anchorId);

            Match titleMatch = Regex.Match(
                html,
                @"<h1[^>]*>(.*?)</h1>",
                headingOpts,
                SecurityHelper.RegexTimeout);
            string pageTitle;
            if (titleMatch.Success)
            {
                pageTitle = Regex.Replace(
                    titleMatch.Groups[1].Value,
                    @"<[^>]+>",
                    string.Empty,
                    RegexOptions.CultureInvariant,
                    SecurityHelper.RegexTimeout).Trim();
            }
            else
            {
                pageTitle = "Known Issues";
            }

            if (pageTitle.Length > 300)
                pageTitle = pageTitle.Substring(0, 300);

            Match kbMatch = Regex.Match(
                normalizedUrl + " " + pageTitle,
                @"KB(\d+)",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                SecurityHelper.RegexTimeout);
            string kb = kbMatch.Success ? kbMatch.Groups[1].Value : null;

            string fullHtml = BuildFullHtml(pageTitle, fragment);

            return new KnownIssuesResult
            {
                Url = normalizedUrl,
                Title = pageTitle,
                Kb = kb,
                AnchorId = anchorId,
                AnchorUrl = anchorUrl,
                FullHtml = fullHtml
            };
        }

        private static string BuildFullHtml(string pageTitle, string fragment)
        {
            string encodedTitle = SecurityHelper.HtmlEncode(pageTitle);
            return "<!DOCTYPE html>\r\n" +
                   "<html><head><meta charset=\"utf-8\">" +
                   "<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; style-src 'unsafe-inline'; img-src data: https:; base-uri 'none'; form-action 'none';\">" +
                   "<title>" + encodedTitle + "</title>\r\n" +
                   "<style>\r\n" +
                   "  body { font-family: Segoe UI, Calibri, Arial, sans-serif; font-size: 14px; margin: 12px; }\r\n" +
                   "  table { border-collapse: collapse; width: 100%; margin: 8px 0 16px 0; }\r\n" +
                   "  th, td { border: 1px solid #999; padding: 6px 8px; text-align: left; vertical-align: top; }\r\n" +
                   "  th { background-color: #f2f2f2; }\r\n" +
                   "  h1, h2, h3, h4 { color: #1a1a1a; }\r\n" +
                   "</style>\r\n" +
                   "</head>\r\n" +
                   "<body>\r\n" +
                   "<h1>" + encodedTitle + "</h1>\r\n" +
                   fragment + "\r\n" +
                   "</body></html>";
        }
    }
}
