<#
.SYNOPSIS
    KB "Known issues" bolumunu bir Windows Formu uzerinden cekip,
    kaynak bicimlendirmesini (tablolar, kalin metin) koruyarak
    gosteren arac.

.USAGE
    powershell -STA -File .\Get-KnownIssuesForm.ps1

    (WebBrowser kontrolu STA thread gerektirir; -STA parametresi onemli.)

.NOTES
    Windows PowerShell 5.1 / .NET Framework 4.x uyumludur.
    Ek modul gerekmez. Yalnizca HTTPS ve Microsoft destek/ogrenme hostlari kabul edilir.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:MaxUrlLength = 2048
$script:MaxDownloadBytes = 10MB
$script:DownloadTimeoutSec = 30
$script:MaxRedirects = 5
$script:RegexTimeoutMs = 5000
$script:AllowedHostPattern = '^(support|learn)\.microsoft\.com$'
$script:SafeAnchorIdPattern = '^[A-Za-z0-9._-]+$'

function Test-PrivateOrReservedAddress {
    param([System.Net.IPAddress]$Address)

    if ([System.Net.IPAddress]::IsLoopback($Address)) { return $true }

    $bytes = $Address.GetAddressBytes()
    if ($Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork) {
        if ($bytes[0] -eq 10) { return $true }
        if ($bytes[0] -eq 172 -and $bytes[1] -ge 16 -and $bytes[1] -le 31) { return $true }
        if ($bytes[0] -eq 192 -and $bytes[1] -eq 168) { return $true }
        if ($bytes[0] -eq 127) { return $true }
        if ($bytes[0] -eq 0) { return $true }
        if ($bytes[0] -ge 224) { return $true }
    }

    if ($Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetworkV6) {
        if ($Address.IsIPv6LinkLocal -or $Address.IsIPv6SiteLocal) { return $true }
    }

    return $false
}

function Test-KnownIssuesUrl {
    param([string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) {
        throw "URL bos olamaz."
    }

    $Url = $Url.Trim()
    if ($Url.Length -gt $script:MaxUrlLength) {
        throw "URL cok uzun."
    }

    $uri = $null
    if (-not [System.Uri]::TryCreate($Url, [System.UriKind]::Absolute, [ref]$uri)) {
        throw "Gecersiz URL formati."
    }

    if ($uri.Scheme -ne [System.Uri]::UriSchemeHttps) {
        throw "Yalnizca HTTPS URL'leri desteklenir."
    }

    if (-not [string]::IsNullOrEmpty($uri.UserInfo)) {
        throw "Kimlik bilgisi iceren URL'ler kabul edilmez."
    }

    if (-not $uri.IsDefaultPort -and $uri.Port -ne 443) {
        throw "Standart disi portlar kabul edilmez."
    }

    if ($uri.Host -notmatch $script:AllowedHostPattern) {
        throw "Yalnizca Microsoft destek/ogrenme sayfalari desteklenir."
    }

    try {
        $addresses = [System.Net.Dns]::GetHostAddresses($uri.Host)
    }
    catch {
        throw "Sunucu adresi cozumlenemedi."
    }

    if (-not $addresses -or $addresses.Count -eq 0) {
        throw "Sunucu adresi cozumlenemedi."
    }

    foreach ($address in $addresses) {
        if (Test-PrivateOrReservedAddress -Address $address) {
            throw "Yerel veya ozel ag adreslerine erisim engellendi."
        }
    }

  return ($uri.GetLeftPart([System.UriPartial]::Path) + $uri.Query + $(if ($uri.Fragment) { $uri.Fragment } else { '' }))
}

function Get-SafeAnchorUrl {
    param(
        [string]$BaseUrl,
        [string]$AnchorId
    )

    $normalized = Test-KnownIssuesUrl -Url $BaseUrl
    if ([string]::IsNullOrEmpty($AnchorId)) {
        return $normalized
    }

    if ($AnchorId -notmatch $script:SafeAnchorIdPattern) {
        throw "Gecersiz bolum kimligi (anchor id)."
    }

    $fragmentIndex = $normalized.IndexOf('#')
    if ($fragmentIndex -ge 0) {
        $normalized = $normalized.Substring(0, $fragmentIndex)
    }

    return "$normalized#$AnchorId"
}

function Remove-UnsafeHtmlFragment {
    param([string]$Fragment)

    if ([string]::IsNullOrEmpty($Fragment)) { return '' }

    $opts = [System.Text.RegularExpressions.RegexOptions]'IgnoreCase, Singleline, CultureInvariant'
    $timeout = [TimeSpan]::FromMilliseconds($script:RegexTimeoutMs)

    $sanitized = [regex]::Replace($Fragment, '<script\b[^<]*(?:(?!</script>)<[^<]*)*</script>', '', $opts, $timeout)
    $sanitized = [regex]::Replace($sanitized, '<(iframe|object|embed|form|base|link|meta)\b[^>]*>.*?</\1>', '', $opts, $timeout)
    $sanitized = [regex]::Replace($sanitized, '<(iframe|object|embed|form|base|link|meta)\b[^>]*/?>', '', $opts, $timeout)
    $sanitized = [regex]::Replace($sanitized, '\s+on\w+\s*=\s*("[^"]*"|''[^'']*''|[^\s>]+)', '', $opts, $timeout)

    $sanitized = [regex]::Replace(
        $sanitized,
        '(href|src|action|formaction|background|xlink:href)\s*=\s*("[^"]*"|''[^'']*''|[^\s>]+)',
        {
            param($match)
            $attribute = $match.Groups[1].Value
            $rawValue = $match.Groups[2].Value.Trim().Trim('"', "'")
            if ($rawValue -match '^(javascript|vbscript|data|file):') {
                return "$attribute=`"#`""
            }
            return $match.Value
        },
        $opts,
        $timeout
    )

    return $sanitized
}

function Get-SecureHtml {
    param([string]$Url)

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $currentUrl = Test-KnownIssuesUrl -Url $Url

    for ($redirect = 0; $redirect -le $script:MaxRedirects; $redirect++) {
        $request = [System.Net.HttpWebRequest]::Create($currentUrl)
        $request.Method = 'GET'
        $request.AllowAutoRedirect = $false
        $request.Timeout = $script:DownloadTimeoutSec * 1000
        $request.ReadWriteTimeout = $script:DownloadTimeoutSec * 1000
        $request.UserAgent = 'GetKnownIssuesForm/1.0'
        $request.Accept = 'text/html'
        $request.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate

        $response = $null
        try {
            $response = $request.GetResponse()
            $statusCode = [int]$response.StatusCode

            if ($statusCode -ge 300 -and $statusCode -lt 400) {
                $location = $response.Headers['Location']
                $response.Close()
                if ([string]::IsNullOrWhiteSpace($location)) {
                    throw "Gecersiz yonlendirme yaniti."
                }
                $currentUrl = Test-KnownIssuesUrl -Url ([System.Uri]::new([System.Uri]$currentUrl, $location).AbsoluteUri)
                continue
            }

            if ($statusCode -ne 200) {
                throw "Sayfa indirilemedi. HTTP $statusCode"
            }

            $contentType = $response.ContentType
            if ($contentType -notmatch 'text/html') {
                throw "Yanit HTML degil."
            }

            $stream = $response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
            $builder = New-Object System.Text.StringBuilder
            $buffer = New-Object char[] 8192

            while (($read = $reader.Read($buffer, 0, $buffer.Length)) -gt 0) {
                if (($builder.Length + $read) -gt $script:MaxDownloadBytes) {
                    throw "Sayfa boyutu izin verilen sinirin ustunde."
                }
                [void]$builder.Append($buffer, 0, $read)
            }

            $reader.Close()
            $response.Close()
            return $builder.ToString()
        }
        finally {
            if ($response) { $response.Close() }
        }
    }

    throw "Cok fazla HTTP yonlendirmesi."
}

function Start-SafeBrowserUrl {
    param([string]$Url)

    $safeUrl = Test-KnownIssuesUrl -Url $Url
    Start-Process $safeUrl
}

# ============================================================
#  Cikarma mantigi
# ============================================================
function Get-KnownIssuesData {
    param([string]$Url)

    $normalizedUrl = Test-KnownIssuesUrl -Url $Url
    $html = Get-SecureHtml -Url $normalizedUrl

    $headingOpts = [System.Text.RegularExpressions.RegexOptions]'IgnoreCase, Singleline, CultureInvariant'
    $headingPattern = '<h([1-4])[^>]*>(?:(?!</h\1>).)*?known\s+issues(?:(?!</h\1>).)*?</h\1>'
    $headingRegex = New-Object System.Text.RegularExpressions.Regex($headingPattern, $headingOpts, [TimeSpan]::FromMilliseconds($script:RegexTimeoutMs))
    $headingMatch = $headingRegex.Match($html)

    if (-not $headingMatch.Success) {
        throw "'Known issues' basligi bu sayfada bulunamadi."
    }

    $headingLevel = [int]$headingMatch.Groups[1].Value
    $startIndex = $headingMatch.Index
    $searchFrom = $startIndex + $headingMatch.Length

    $levels = 1..$headingLevel -join ""
    $nextHeadingRegex = New-Object System.Text.RegularExpressions.Regex("<h[$levels][^>]*>", 'IgnoreCase, CultureInvariant', [TimeSpan]::FromMilliseconds($script:RegexTimeoutMs))
    $nextMatch = $nextHeadingRegex.Match($html, $searchFrom)

    if ($nextMatch.Success) {
        $endIndex = $nextMatch.Index
    } else {
        $bodyEndIdx = $html.IndexOf("</body>", $searchFrom, [System.StringComparison]::OrdinalIgnoreCase)
        $endIndex = if ($bodyEndIdx -ge 0) { $bodyEndIdx } else { $html.Length }
    }

    $fragment = $html.Substring($startIndex, $endIndex - $startIndex)
    $fragment = Remove-UnsafeHtmlFragment -Fragment $fragment

    $headingOpenTag = [regex]::Match($headingMatch.Value, '^<h[1-4][^>]*>', 'CultureInvariant', [TimeSpan]::FromMilliseconds($script:RegexTimeoutMs)).Value
    $idMatch = [regex]::Match($headingOpenTag, 'id\s*=\s*"([^"]+)"', 'IgnoreCase, CultureInvariant', [TimeSpan]::FromMilliseconds($script:RegexTimeoutMs))
    $anchorId = if ($idMatch.Success) { $idMatch.Groups[1].Value } else { $null }
    $anchorUrl = Get-SafeAnchorUrl -BaseUrl $normalizedUrl -AnchorId $anchorId

    $titleMatch = [regex]::Match($html, '<h1[^>]*>(.*?)</h1>', $headingOpts, [TimeSpan]::FromMilliseconds($script:RegexTimeoutMs))
    $pageTitle = if ($titleMatch.Success) {
        [regex]::Replace($titleMatch.Groups[1].Value, '<[^>]+>', '', 'CultureInvariant', [TimeSpan]::FromMilliseconds($script:RegexTimeoutMs)).Trim()
    } else { "Known Issues" }

    if ($pageTitle.Length -gt 300) {
        $pageTitle = $pageTitle.Substring(0, 300)
    }

    $kbMatch = [regex]::Match("$normalizedUrl $pageTitle", 'KB(\d+)', 'IgnoreCase, CultureInvariant', [TimeSpan]::FromMilliseconds($script:RegexTimeoutMs))
    $kb = if ($kbMatch.Success) { $kbMatch.Groups[1].Value } else { $null }

    $encodedTitle = [System.Net.WebUtility]::HtmlEncode($pageTitle)
    $fullHtml = @"
<!DOCTYPE html>
<html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data: https:; base-uri 'none'; form-action 'none';">
<title>$encodedTitle</title>
<style>
  body { font-family: Segoe UI, Calibri, Arial, sans-serif; font-size: 14px; margin: 12px; }
  table { border-collapse: collapse; width: 100%; margin: 8px 0 16px 0; }
  th, td { border: 1px solid #999; padding: 6px 8px; text-align: left; vertical-align: top; }
  th { background-color: #f2f2f2; }
  h1, h2, h3, h4 { color: #1a1a1a; }
</style>
</head>
<body>
<h1>$encodedTitle</h1>
$fragment
</body></html>
"@

    [pscustomobject]@{
        Url       = $normalizedUrl
        Title     = $pageTitle
        Kb        = $kb
        AnchorId  = $anchorId
        AnchorUrl = $anchorUrl
        FullHtml  = $fullHtml
    }
}

# ============================================================
#  Form
# ============================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "KB Known Issues Cikarici"
$form.Size = New-Object System.Drawing.Size(1000, 720)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(700, 450)

$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = "Top"
$topPanel.Height = 104
$form.Controls.Add($topPanel)

$lblUrl = New-Object System.Windows.Forms.Label
$lblUrl.Text = "KB Sayfa URL'si:"
$lblUrl.Location = New-Object System.Drawing.Point(10, 12)
$lblUrl.AutoSize = $true
$topPanel.Controls.Add($lblUrl)

$txtUrl = New-Object System.Windows.Forms.TextBox
$txtUrl.Location = New-Object System.Drawing.Point(10, 32)
$txtUrl.Width = 960
$txtUrl.MaxLength = 2048
$txtUrl.Anchor = "Top,Left,Right"
$txtUrl.Text = "https://support.microsoft.com/en-us/servicing/os/windows-11/2026/06/june-9-2026-kb5093998-os-build-22631-7219"
$topPanel.Controls.Add($txtUrl)

$btnFetch = New-Object System.Windows.Forms.Button
$btnFetch.Text = "Getir"
$btnFetch.Location = New-Object System.Drawing.Point(10, 64)
$btnFetch.Width = 90
$btnFetch.Anchor = "Top,Left"
$topPanel.Controls.Add($btnFetch)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Tarayicida Ac"
$btnBrowse.Location = New-Object System.Drawing.Point(110, 64)
$btnBrowse.Width = 130
$btnBrowse.Anchor = "Top,Left"
$topPanel.Controls.Add($btnBrowse)

$browser = New-Object System.Windows.Forms.WebBrowser
$browser.Dock = "Fill"
$browser.ScriptErrorsSuppressed = $true
$browser.IsWebBrowserContextMenuEnabled = $false
$browser.WebBrowserShortcutsEnabled = $false
$browser.AllowWebBrowserDrop = $false
$browser.Anchor = "Top,Left,Right,Bottom"
$browser.Location = New-Object System.Drawing.Point(0, 100)
$browser.Width = 1000
$browser.Height = 800

$form.Controls.Add($browser)
$browser.BringToFront()

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Hazir."
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)

$topPanel.BringToFront()

$script:LastResult = $null

function Set-Status {
    param([string]$Text)
    $statusLabel.Text = $Text
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-ValidatedUrlFromTextBox {
  $url = $txtUrl.Text.Trim()
  if ([string]::IsNullOrWhiteSpace($url)) {
    [System.Windows.Forms.MessageBox]::Show("Lutfen bir URL girin.", "Uyari", "OK", "Warning") | Out-Null
    return $null
  }

  try {
    return (Test-KnownIssuesUrl -Url $url)
  }
  catch {
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Gecersiz URL", "OK", "Warning") | Out-Null
    return $null
  }
}

$btnFetch.Add_Click({
    $url = Get-ValidatedUrlFromTextBox
    if (-not $url) { return }

    $btnFetch.Enabled = $false
    $form.Cursor = "WaitCursor"
    Set-Status "Sayfa indiriliyor..."

    try {
        $result = Get-KnownIssuesData -Url $url
        $script:LastResult = $result

        Set-Status "Bulundu: $($result.Title) -- onizleme yukleniyor..."
        $browser.DocumentText = $result.FullHtml
        Set-Status "Tamam. '$($result.Title)' icin Known Issues bolumu yuklendi."
    }
    catch {
        Set-Status "Hata olustu."
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Hata", "OK", "Error") | Out-Null
    }
    finally {
        $form.Cursor = "Default"
        $btnFetch.Enabled = $true
    }
})

$btnBrowse.Add_Click({
    $url = Get-ValidatedUrlFromTextBox
    if (-not $url) { return }

    $form.Cursor = "WaitCursor"
    try {
        if (-not $script:LastResult -or $script:LastResult.Url -ne $url) {
            Set-Status "Bolum konumu tespit ediliyor..."
            $result = Get-KnownIssuesData -Url $url
            $script:LastResult = $result
            $browser.DocumentText = $result.FullHtml
        }

        if ($script:LastResult.AnchorId) {
            Set-Status "Tarayicida aciliyor (dogrudan bolume atlanacak)."
        } else {
            Set-Status "Sayfada 'id' bulunamadi, sayfa basindan aciliyor."
        }

        Start-SafeBrowserUrl -Url $script:LastResult.AnchorUrl
    }
    catch {
        Set-Status "Tarayici acilamadi."
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Hata", "OK", "Error") | Out-Null
    }
    finally {
        $form.Cursor = "Default"
    }
})

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($form)
