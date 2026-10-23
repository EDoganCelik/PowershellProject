#Requires -Version 5.1
<#
.SYNOPSIS
    Istemci Durum ve SCCM Islem Araci (WinForms GUI)

.ACIKLAMA
    Sol textbox: Bilgisayar isimleri (her satira bir tane).

    "Kontrol Et" her makine icin sunlari getirir:
      - DNS durumu (cozumlenemezse "No such host is known")
      - Cevrimici durumu (ping)
      - Active Directory: hesap durumu (Enabled / Disabled / Bulunamadi) ve Last Logon
      - Istenen servislerin durumu + CcmExec (SCCM ajani) durumu

    "Servisleri Baslat" -> durdurulmus servisleri baslatir.

    SCCM Istemci Eylemi (3. bolum): Machine Policy, Update Scan, Update Deployment
    Evaluation, App Deployment Evaluation, Hardware/Software Inventory tetikler.
    Hedef: gridde secili satirlar; secim yoksa soldaki listedeki tum makineler.

    SCCM Script (2. bolum): SMS Provider'a baglanip ortamdaki scriptleri listeler,
    secilen (ONAYLI) scripti hedef makinelerde calistirir ve sonucu bekler.
    Cevap gelmezse sebep tespit edilir:
      - SCCM'de ajan yok          -> "SCCM ajani bulunmuyor"
      - Makine ping vermiyor      -> "Cevrimdisi oldugu icin script calismadi"
      - Makine cevrimici          -> "Cevrimici - script calismadi"

    SEKMELER (PANELLER):
      Panel 1 - Script Ciktilari / Kontroller : DNS, Cevrimici, AD durumu, Last Logon, servis
                durumlari ve SCCM script ciktisi (sadece cikti, exit code/OperationID vb. yok;
                onlar satira cift tiklayinca Detay'da gorunur).
      Panel 2 - Update Parser (Panel 1'den BAGIMSIZ) : kendi yapistirma kutusu, kendi script gonder /
                kontrol butonlari, kendi tablosu ve kendi durum cubugu vardir; Panel 1'e hicbir sey yazmaz.
                  * "Parse Et": SCCM Run Script ciktisi ('@' ile ayrilan makine kayitlari, satirlar ';' ile
                    ayrilmis) alt alta kolonlara bolunur (client_update_parser.ps1 mantigi).
                  * Parse sonrasi her cihaz icin DNS, Online, AD (Enabled/Disabled/Last Logon), CcmExec
                    (SCCM ajan servisi) ve SCCM kaydi kontrol edilir; sonuc ekstra sutunlara yazilir.
                  * "Script Gonder (Panel 2)": scripti dogrudan SCCM'den gonderir; cikti + kontroller
                    SADECE Panel 2'ye yazilir (Panel 1'deki "Secili Scripti Gonder" ise SADECE Panel 1'e yazar).
                  * Filtre, Excel ve "CSV Olustur ve Ac" (<gun-saat>-updates.csv, kayit klasoru secilebilir).
      Ortak olanlar sadece ayarlardir: soldaki bilgisayar listesi, SMS Provider baglantisi / script secimi,
      bekleme suresi ve "tum DC" secenegi. Sonuc tablolari ve istekler ayridir; tum isler tek havuzu
      paylastigi icin toplamda yine en fazla 10 makineye ayni anda istek gider.

    AD YAPISI KONTROLU: DNS + ping + AD (var mi / Enabled-Disabled / Last Logon).
      - "SCCM Script" gonderilirken her makine icin otomatik calisir.
      - "AD Yapisi Kontrolu" butonu ile script gondermeden tek basina da calistirilabilir.

    SONUC TABLOSU: Islem baslayinca hedef makinelerin TUMU hemen tabloya eklenir; sonucu henuz
    gelmeyen hucrelerde "Bekleniyor.." yazar, cevap geldikce ilgili satir yerinde guncellenir.
    Ulasilamayan makinelerde "Ulasilamadi" ve sebebi (cevrimdisi, DNS, SCCM ajani yok vb.) yazilir.

.PARALELLIK
    Tum islemler RunspacePool ile paralel calisir; havuz SABIT 10 ile sinirlidir
    (ayni anda en fazla 10 makineye istek gider). UI donmaz.

.REMOTING YONTEMI
    WinRM / Invoke-Command / PSSession KULLANILMAZ. Sadece RPC / DCOM / WMI:
      - Servisler          : Get-Service -ComputerName / Start-Service
      - SCCM istemci eylemi: WMI root\ccm : SMS_Client.TriggerSchedule
      - SCCM script        : SMS Provider (root\SMS\site_XXX) SMS_Scripts,
                             SMS_ClientOperation.InitiateClientOperationEx (Type 135),
                             SMS_ScriptsExecutionStatus
      - Active Directory   : RSAT ActiveDirectory modulu (Get-ADComputer / Get-ADDomainController)
    SCCM konsolu veya ConfigurationManager modulu GEREKMEZ.

.GEREKSINIMLER
    - Windows PowerShell 5.1 (PowerShell 7'de Get-WmiObject yoktur, calismaz)
    - RSAT: Active Directory PowerShell modulu (ActiveDirectory) + DC'lere ADWS (TCP 9389) erisimi
    - Hedef makinelerde yerel admin (servis + root\ccm icin), RPC/DCOM erisimi
    - SCCM'de: scriptlerin ONAYLANMIS olmasi ve kullanicinin "Run Script"
      yetkisinin bulunmasi (RBAC)

.KAYITLAR
    C:\Users\<kullanici>\Desktop\ServiceStatus\ altinda:
      - ClientToolSettings.json  : SMS Provider, servis listesi, bekleme suresi vb.
      - ClientTool_yyyyMM.log    : SCCM eylem / script gonderimlerinin denetim kaydi
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.DirectoryServices
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==================== GLOBAL DEGISKENLER ====================
$script:MaxParallel  = 10      # SABIT esik: ayni anda en fazla 10 makine
$script:LastResults  = @()     # Son kontrol sonuclari (ServiceDetail dahil)
$script:Details      = @{}     # Hostname -> satir detay/log metni
$script:SmsServer    = ''
$script:SiteCode     = ''
$script:CmScripts    = @()
$script:PendingScriptGuid = ''
$script:ParserNoOutput = @{}   # Panel 2: script'ten cikti gelmeyen hostlar
$script:ParserDetails  = @{}   # Panel 2: Hostname -> detay metni (Panel 1'in $script:Details tablosu ile KARISMAZ)
$script:ParsePasteSeq  = 0     # Panel 2: "Parse Et" sayaci (her yapistirma icin ayri SourceHost etiketi)
$script:HostRegex      = '^[A-Za-z0-9][A-Za-z0-9._-]*$'
$script:SharedPool     = $null # Panel 1 + Panel 2 islerinin ORTAK havuzu (toplamda en fazla 10 makine)
$script:ActiveRuns     = New-Object System.Collections.ArrayList   # calisan paralel isler (panel bazli)
$script:AllColsLabel = '(Tum kolonlar)'
$script:ParserDefaultFolder = Join-Path $env:USERPROFILE 'Desktop\SCCM-Raporlari'

# Update Parser CSV basligi (21 kolon). Eski 19 kolonluk cikti (client_update_parser.ps1) da otomatik taninir.
$script:ParserHeaderText = 'Hostname;UpdateName;EvalState;PercentComplete;ErrorCode;LastShutdownTime;IsBlueScreen;CBS_NTSTATUS_FROM_WIN32;CBS_STATUS_SXS_FILE_HASH_MISMATCH;UD_CERTIFICATE_ERROR_0x80010002;UD_DOWNLOAD_ERROR_0x80d02002;BuildVersion;SistemDrive_FreeGB;SistemDrive_FreePercent;Wuauserv_Durum;BITS_Durum;CryptSvc_Durum;TrustedInstaller_Durum;SoftwareDistribution_BoyutMB;DISM_CheckHealth;DTS_DP_IndirmeHatalari'
$script:ParserColumns    = @($script:ParserHeaderText -split ';')

# Eski 19 kolonluk cikti (client_update_parser.ps1 basligi): satirda 19 alan gelirse kolonlar ADIYLA yerlestirilir
$script:ParserLegacyHeaderText = 'Hostname;UpdateName;EvalState;PercentComplete;ErrorCode;LastShutdownTime;IsBlueScreen;CBS_NTSTATUS_FROM_WIN32;CBS_STATUS_SXS_FILE_HASH_MISMATCH;UD_CERTIFICATE_ERROR_0x80010002;UD_DOWNLOAD_ERROR_0x80d02002;BuildVersion;SistemDrive_FreeGB;SistemDrive_FreePercent;Wuauserv_Durum;BITS_Durum;CryptSvc_Durum;TrustedInstaller_Durum;DISM_CheckHealth'
$script:ParserLegacyColumns    = @($script:ParserLegacyHeaderText -split ';')

# Panel 2'de parse sonrasi eklenen kontrol kolonlari (Panel 1'deki DNS / Online / AD / CcmExec bilgisinin karsiligi)
$script:ParserCheckColumns = @('DNS_Durumu', 'Online_mi', 'AD_Durumu', 'AD_LastLogon', 'CcmExec_Durumu', 'SCCM_Kayit')

# SCCM istemci eylemleri: gorunen ad -> TriggerSchedule ScheduleID listesi
$script:ClientActions = [ordered]@{
    'Machine Policy (Retrieval + Evaluation)'          = @('{00000000-0000-0000-0000-000000000021}', '{00000000-0000-0000-0000-000000000022}')
    'Software Updates Scan (Update Scan)'              = @('{00000000-0000-0000-0000-000000000113}')
    'Software Updates Deployment Evaluation'           = @('{00000000-0000-0000-0000-000000000108}')
    'Update Scan + Update Deployment Evaluation'       = @('{00000000-0000-0000-0000-000000000113}', '{00000000-0000-0000-0000-000000000108}')
    'Application Deployment Evaluation'                = @('{00000000-0000-0000-0000-000000000121}')
    'Hardware Inventory'                               = @('{00000000-0000-0000-0000-000000000001}')
    'Software Inventory'                               = @('{00000000-0000-0000-0000-000000000002}')
    'Discovery Data Collection (DDR)'                  = @('{00000000-0000-0000-0000-000000000003}')
}

# ==================== PARALEL RUNSPACE'LERDE DE KULLANILAN FONKSIYONLAR ====================
# NOT: Bu bolumdeki fonksiyonlar hem ana pencere hem de RunspacePool icindeki
# runspace'ler tarafindan kullanilir. Disaridan hicbir degiskene referans vermezler.
# Kaynak kodlari New-WorkerScript ile runspace scriptinin basina gomulur.

function Test-HostOnline {
    param([string]$ComputerName)
    return [bool](Test-Connection -ComputerName $ComputerName -Count 2 -Quiet -ErrorAction SilentlyContinue)
}

function Resolve-HostDns {
    # Cozumlenemezse Text = "No such host is known" (isletim sistemi dilinden bagimsiz)
    param([string]$ComputerName)
    try {
        $ips = @([System.Net.Dns]::GetHostAddresses($ComputerName))
        if ($ips.Count -eq 0) {
            return [pscustomobject]@{ Ok = $false; Text = 'No such host is known' }
        }
        $v4 = @($ips | Where-Object { $_.AddressFamily -eq 'InterNetwork' })
        $use = if ($v4.Count -gt 0) { $v4 } else { $ips }
        $txt = (($use | ForEach-Object { $_.IPAddressToString }) -join ', ')
        return [pscustomobject]@{ Ok = $true; Text = $txt }
    }
    catch {
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        if ($ex -is [System.Net.Sockets.SocketException] -and $ex.SocketErrorCode -eq [System.Net.Sockets.SocketError]::HostNotFound) {
            return [pscustomobject]@{ Ok = $false; Text = 'No such host is known' }
        }
        return [pscustomobject]@{ Ok = $false; Text = "DNS hatasi: $($ex.Message)" }
    }
}

function Convert-FileTimeSafe {
    param($Value)
    try {
        $v = [int64]$Value
        if ($v -le 0 -or $v -ge 9223372036854775807) { return $null }
        return [DateTime]::FromFileTime($v)
    }
    catch { return $null }
}

function Get-AdComputerInfo {
    # Active Directory'den bilgisayar hesabini RSAT (ActiveDirectory modulu) ile bulur.
    # Status: Enabled / Disabled / Bulunamadi / Sorgulanamadi
    # LastLogon: lastLogonTimestamp (replike, ~14 gun gecikmeli) ile sorgulanan DC'lerdeki
    # lastLogon degerlerinin en yenisi. $QueryAllDCs = $true ise (GLOBAL TARAMA) domaindeki
    # tum DC'ler tek tek sorgulanir (daha dogru ama yavas).
    # Gereksinim: RSAT - Active Directory PowerShell modulu + ADWS (TCP 9389) erisimi.
    param([string]$ComputerName, [bool]$QueryAllDCs = $false)

    $res = [pscustomobject]@{ Status = 'Bulunamadi'; LastLogon = 'N/A'; Detail = '' }

    # ---- RSAT modulu ----
    if (-not (Get-Module -Name ActiveDirectory)) {
        try { Import-Module -Name ActiveDirectory -ErrorAction Stop -Verbose:$false }
        catch {
            $res.Status = 'Sorgulanamadi'
            $res.Detail = "ActiveDirectory modulu (RSAT) yuklenemedi: $($_.Exception.Message)"
            return $res
        }
    }

    $short  = ($ComputerName -split '\.')[0]
    $safe   = $short.Replace('\', '\5c').Replace('*', '\2a').Replace('(', '\28').Replace(')', '\29')
    $filter = "(&(objectCategory=computer)(sAMAccountName=$safe`$))"
    $props  = @('Enabled', 'LastLogonTimestamp', 'LastLogon', 'DistinguishedName')

    $found = $null
    $viaGc = $false

    # ---- 1) Mevcut domain ----
    try {
        $found = Get-ADComputer -LDAPFilter $filter -Properties $props -ErrorAction Stop | Select-Object -First 1
    }
    catch {
        $res.Status = 'Sorgulanamadi'
        $res.Detail = "AD sorgusu basarisiz: $($_.Exception.Message)"
        return $res
    }

    # ---- 2) Domain'de yoksa forest Global Catalog'da ara (baska domain'deki hesaplar icin) ----
    if (-not $found) {
        try {
            $gc = Get-ADDomainController -Discover -Service GlobalCatalog -ErrorAction Stop
            $gcHost = @($gc.HostName)[0]
            $found = Get-ADComputer -LDAPFilter $filter -Server "${gcHost}:3268" `
                -Properties @('Enabled', 'LastLogonTimestamp', 'DistinguishedName') -ErrorAction Stop |
                Select-Object -First 1
            if ($found) { $viaGc = $true }
        }
        catch { }
    }

    if (-not $found) { return $res }   # Status zaten 'Bulunamadi'

    $res.Status = if ($found.Enabled) { 'Enabled' } else { 'Disabled' }

    $times = @()
    foreach ($pn in @('LastLogonTimestamp', 'LastLogon')) {
        $v = $found.$pn
        if ($null -ne $v) {
            $t = Convert-FileTimeSafe $v
            if ($t) { $times += $t }
        }
    }

    # ---- 3) GLOBAL TARAMA: domaindeki tum DC'lerden gercek lastLogon ----
    if ($QueryAllDCs -and -not $viaGc -and $found.DistinguishedName) {
        $dn = [string]$found.DistinguishedName
        $unreach = 0
        try {
            # DC listesi runspace basina 10 dk onbellekte tutulur (cok makinede tekrar tekrar sorgulanmasin)
            $cache = $global:AdDcCache
            if (-not $cache -or ((Get-Date) - $cache.Time).TotalMinutes -gt 10) {
                $list  = @(Get-ADDomainController -Filter * -ErrorAction Stop | ForEach-Object { [string]$_.HostName })
                $cache = [pscustomobject]@{ Time = (Get-Date); Hosts = $list }
                $global:AdDcCache = $cache
            }

            foreach ($dc in $cache.Hosts) {
                # Ulasilamayan DC'de uzun beklememek icin ADWS portuna 2 sn'lik hizli kontrol
                $reach = $false
                $tcp = $null
                try {
                    $tcp = New-Object System.Net.Sockets.TcpClient
                    $iar = $tcp.BeginConnect($dc, 9389, $null, $null)
                    if ($iar.AsyncWaitHandle.WaitOne(2000, $false)) { $tcp.EndConnect($iar); $reach = $true }
                }
                catch { $reach = $false }
                finally { if ($tcp) { $tcp.Close() } }
                if (-not $reach) { $unreach++; continue }

                try {
                    $one = Get-ADComputer -Identity $dn -Server $dc -Properties LastLogon -ErrorAction Stop
                    if ($null -ne $one.LastLogon) {
                        $t = Convert-FileTimeSafe $one.LastLogon
                        if ($t) { $times += $t }
                    }
                }
                catch { }
            }
        }
        catch {
            $res.Detail = "DC listesi alinamadi, global tarama yapilamadi: $($_.Exception.Message)"
        }
        if ($unreach -gt 0) {
            $msg = "$unreach DC'ye ulasilamadi (ADWS 9389); LastLogon eksik DC verisiyle hesaplandi"
            $res.Detail = if ($res.Detail) { "$($res.Detail) | $msg" } else { $msg }
        }
    }

    if ($times.Count -gt 0) {
        $best = $times | Sort-Object -Descending | Select-Object -First 1
        $res.LastLogon = $best.ToString('yyyy-MM-dd HH:mm')
    }
    else {
        $res.LastLogon = 'Hic'
    }
    return $res
}

function Get-RemoteStatus {
    param(
        [string]$ComputerName,
        [string[]]$ServiceNames,
        [bool]$QueryAllDCs = $false
    )

    $row = [pscustomobject]@{
        Hostname      = $ComputerName
        DnsStatus     = ''
        IsOnline      = $false
        AdStatus      = ''
        LastLogon     = ''
        Running       = ''
        NotRunning    = ''
        CcmExec       = ''
        ServiceDetail = @()   # Name/Status listesi -> Servis baslatma asamasinda kullanilir
        ErrorMessage  = ''
        State         = 'fail'
    }
    $errs = @()

    # ---- DNS ----
    $dns = Resolve-HostDns -ComputerName $ComputerName
    $row.DnsStatus = $dns.Text

    # ---- Active Directory ----
    $ad = Get-AdComputerInfo -ComputerName $ComputerName -QueryAllDCs $QueryAllDCs
    $row.AdStatus  = $ad.Status
    $row.LastLogon = $ad.LastLogon
    if ($ad.Detail) { $errs += $ad.Detail }

    if (-not $dns.Ok) {
        $errs += 'DNS cozumlenemedi, ping/servis kontrolu yapilmadi'
        $row.ErrorMessage = ($errs -join ' | ')
        return $row
    }

    # ---- Cevrimici (ping) ----
    $online = Test-HostOnline -ComputerName $ComputerName
    $row.IsOnline = [bool]$online
    if (-not $online) {
        $errs += 'Makine yanit vermiyor (ping basarisiz)'
        $row.ErrorMessage = ($errs -join ' | ')
        return $row
    }

    # ---- Servisler: tek RPC cagrisi ile tum liste alinir, yerelde suzulur ----
    $all = $null
    try { $all = @(Get-Service -ComputerName $ComputerName -ErrorAction Stop) }
    catch { $errs += "Servis listesi alinamadi: $($_.Exception.Message)" }

    $detail = @()
    foreach ($sn in $ServiceNames) {
        $n = "$sn".Trim()
        if (-not $n) { continue }
        if ($null -eq $all) {
            $detail += [pscustomobject]@{ Name = $n; Status = 'Erisilemedi' }
            continue
        }
        $m = $all | Where-Object { $_.Name -eq $n -or $_.DisplayName -eq $n } | Select-Object -First 1
        if ($m) { $detail += [pscustomobject]@{ Name = $n; Status = $m.Status.ToString() } }
        else    { $detail += [pscustomobject]@{ Name = $n; Status = 'NotFound' } }
    }
    $row.ServiceDetail = $detail
    $row.Running    = (($detail | Where-Object { $_.Status -eq 'Running' } | ForEach-Object { $_.Name }) -join ', ')
    $row.NotRunning = (($detail | Where-Object { $_.Status -ne 'Running' } | ForEach-Object { "$($_.Name)($($_.Status))" }) -join ', ')

    # ---- SCCM ajani (CcmExec) ----
    if ($null -eq $all) {
        $row.CcmExec = 'Erisilemedi'
    }
    else {
        $cc = $all | Where-Object { $_.Name -eq 'CcmExec' } | Select-Object -First 1
        $row.CcmExec = if ($cc) { $cc.Status.ToString() } else { 'Yok' }
    }

    $row.ErrorMessage = ($errs -join ' | ')
    $row.State = if ($row.NotRunning -or $row.CcmExec -eq 'Yok') { 'warn' } else { 'ok' }
    return $row
}

function Start-RemoteFix {
    param(
        [string]$ComputerName,
        [pscustomobject[]]$ServiceDetail   # kontrol asamasindaki ONCEKI durum
    )

    $results = @()
    foreach ($s in $ServiceDetail) {

        if ($s.Status -eq 'Running') {
            $results += [pscustomobject]@{ Hostname = $ComputerName; ServiceName = $s.Name; OncekiDurum = $s.Status; SonrakiDurum = $s.Status; Aciklama = 'Zaten calisiyordu, islem yapilmadi.' }
            continue
        }
        if ($s.Status -eq 'NotFound' -or $s.Status -eq 'Erisilemedi') {
            $results += [pscustomobject]@{ Hostname = $ComputerName; ServiceName = $s.Name; OncekiDurum = $s.Status; SonrakiDurum = $s.Status; Aciklama = "Durum $($s.Status), islem yapilmadi." }
            continue
        }

        try {
            Get-Service -ComputerName $ComputerName -Name $s.Name -ErrorAction Stop | Start-Service -ErrorAction Stop
            Start-Sleep -Milliseconds 800
            $newStatus = (Get-Service -ComputerName $ComputerName -Name $s.Name -ErrorAction Stop).Status.ToString()
            $desc = if ($newStatus -eq 'Running') { "$($s.Status) durumundaydi, Running yapildi." } else { "Baslatildi ama durum hala $newStatus." }
            $results += [pscustomobject]@{ Hostname = $ComputerName; ServiceName = $s.Name; OncekiDurum = $s.Status; SonrakiDurum = $newStatus; Aciklama = $desc }
        }
        catch {
            $results += [pscustomobject]@{ Hostname = $ComputerName; ServiceName = $s.Name; OncekiDurum = $s.Status; SonrakiDurum = 'Bilinmiyor'; Aciklama = "Hata: $($_.Exception.Message)" }
        }
    }
    return $results
}

function Get-WmiFailureReason {
    # WMI/DCOM hatasini siniflandirir: NoAgent / Rpc / Access / Other
    param($ErrorRecord)

    $kind = 'Other'
    $first = $ErrorRecord.Exception.Message
    $e = $ErrorRecord.Exception

    while ($e) {
        if ($e -is [System.Management.ManagementException]) {
            $code = $e.ErrorCode.ToString()
            if ($code -eq 'InvalidNamespace') { $kind = 'NoAgent'; break }
            if ($code -eq 'AccessDenied')     { $kind = 'Access';  break }
        }
        if ($e -is [System.UnauthorizedAccessException]) { $kind = 'Access'; break }
        if ($e -is [System.Runtime.InteropServices.COMException]) {
            $hr = $e.ErrorCode
            if ($hr -eq -2147023174 -or $hr -eq -2147023170) { $kind = 'Rpc';     break }   # 0x800706BA / 0x800706BE
            if ($hr -eq -2147024891)                         { $kind = 'Access';  break }   # 0x80070005
            if ($hr -eq -2147217394)                         { $kind = 'NoAgent'; break }   # 0x8004100E
        }
        $e = $e.InnerException
    }

    if ($kind -eq 'Other') {
        if ($first -match 'Invalid namespace|0x8004100E') { $kind = 'NoAgent' }
        elseif ($first -match 'RPC server|0x800706BA|0x800706BE') { $kind = 'Rpc' }
        elseif ($first -match 'Access is denied|0x80070005') { $kind = 'Access' }
    }

    $text = switch ($kind) {
        'NoAgent' { "SCCM ajani bulunmuyor (root\ccm namespace'i yok)." }
        'Rpc'     { 'RPC sunucusu kullanilamiyor (makine kapali/firewall).' }
        'Access'  { 'Erisim engellendi (hedefte yerel admin yetkisi gerekli).' }
        default   { $first }
    }
    return [pscustomobject]@{ Kind = $kind; Text = $text }
}

function Invoke-SccmClientAction {
    # Hedef makinede root\ccm SMS_Client.TriggerSchedule cagirir (DCOM/RPC).
    param([string]$ComputerName, [string]$ActionName, [string[]]$ScheduleIds)

    $r = [pscustomobject]@{ Hostname = $ComputerName; Action = $ActionName; State = 'fail'; Result = ''; Detail = '' }

    $dns = Resolve-HostDns -ComputerName $ComputerName
    if (-not $dns.Ok) {
        $r.Result = "DNS: $($dns.Text) - eylem gonderilemedi."
        return $r
    }
    if (-not (Test-HostOnline -ComputerName $ComputerName)) {
        $r.Result = 'Cevrimdisi - eylem gonderilemedi (ping yanit vermiyor).'
        return $r
    }

    $sent = @()
    foreach ($id in $ScheduleIds) {
        try {
            $null = Invoke-WmiMethod -ComputerName $ComputerName -Namespace 'root\ccm' -Class 'SMS_Client' -Name 'TriggerSchedule' -ArgumentList $id -ErrorAction Stop
            $sent += $id
        }
        catch {
            $why = Get-WmiFailureReason -ErrorRecord $_
            $r.Result = $why.Text
            $r.Detail = "Basarisiz ScheduleID: $id`r`n$($_.Exception.Message)"
            if ($sent.Count -gt 0) { $r.Detail += "`r`nOncesinde gonderilenler: $($sent -join ', ')" }
            return $r
        }
    }

    $r.State  = 'ok'
    $r.Result = 'Gonderildi'
    $r.Detail = "Tetiklenen ScheduleID: $($sent -join ', ')"
    return $r
}

function Get-NoResponseVerdict {
    # Script'ten cevap gelmediginde sebebi belirler:
    #   cevrimdisi / SCCM ajani yok / cevrimici ama script calismadi
    param([string]$ComputerName, [int]$TimeoutSec)

    $dns = Resolve-HostDns -ComputerName $ComputerName
    if (-not $dns.Ok) {
        return [pscustomobject]@{ State = 'fail'; Text = "Cevrimdisi oldugu icin script calismadi (DNS: $($dns.Text))." }
    }
    if (-not (Test-HostOnline -ComputerName $ComputerName)) {
        return [pscustomobject]@{ State = 'fail'; Text = 'Cevrimdisi oldugu icin script calismadi (ping yanit vermiyor).' }
    }

    $svc = $null
    try { $svc = Get-Service -ComputerName $ComputerName -Name 'CcmExec' -ErrorAction Stop }
    catch {
        if ($_.FullyQualifiedErrorId -like 'NoServiceFoundForGivenName*') {
            return [pscustomobject]@{ State = 'fail'; Text = 'SCCM ajani bulunmuyor (CcmExec servisi yok).' }
        }
        return [pscustomobject]@{ State = 'warn'; Text = "Cevrimici - script calismadi (CcmExec sorgulanamadi: $($_.Exception.Message))." }
    }

    if ($svc.Status -ne 'Running') {
        return [pscustomobject]@{ State = 'warn'; Text = "Cevrimici - script calismadi (CcmExec servisi $($svc.Status))." }
    }
    return [pscustomobject]@{ State = 'warn'; Text = "Cevrimici - script calismadi ($TimeoutSec sn icinde SCCM'den yanit gelmedi)." }
}

function Invoke-SccmScriptOnDevice {
    # SMS Provider uzerinden "Run Script" (ClientOperation Type 135) baslatir ve
    # sonucu SMS_ScriptsExecutionStatus uzerinden bekler.
    param(
        [string]$ComputerName,
        [string]$SmsProvider,
        [string]$SiteCode,
        [string]$ScriptGuid,
        [string]$ScriptName,
        [string]$ParamB64,
        [int]$TimeoutSec
    )

    $r = [pscustomobject]@{ Hostname = $ComputerName; Action = "Script: $ScriptName"; State = 'fail'; Result = ''; Output = ''; ExitCode = $null; Detail = '' }
    $ns = "root\SMS\site_$SiteCode"
    $short = ($ComputerName -split '\.')[0]

    # ---- 1) Cihaz SCCM'de kayitli mi, istemci (ajan) var mi? ----
    try {
        $cands = @(Get-WmiObject -ComputerName $SmsProvider -Namespace $ns -ErrorAction Stop `
            -Query "SELECT ResourceId, Name, Client, Active, Obsolete FROM SMS_R_System WHERE Name = '$short'")
    }
    catch {
        $r.Result = "SCCM sorgusu basarisiz: $($_.Exception.Message)"
        return $r
    }

    $dev = $cands | Where-Object { $_.Client -eq 1 -and $_.Obsolete -ne 1 } | Select-Object -First 1
    if (-not $dev) {
        if ($cands.Count -gt 0) {
            $r.Result = 'SCCM ajani bulunmuyor (cihaz kaydi var ama istemci yuklu degil/obsolete).'
        }
        else {
            $r.Result = "SCCM ajani bulunmuyor (cihaz SCCM'de kayitli degil)."
        }
        return $r
    }
    $rid = [uint32]$dev.ResourceId

    # ---- 2) Scripti gonder ----
    $op = $null
    try {
        $mc = [WmiClass]"\\$SmsProvider\root\SMS\site_${SiteCode}:SMS_ClientOperation"
        $in = $mc.psbase.GetMethodParameters('InitiateClientOperationEx')
        $in.Param = $ParamB64
        $in.RandomizationWindow = [uint32]0
        $in.TargetCollectionID = ''
        $in.TargetResourceIDs = [uint32[]]@($rid)
        $in.Type = [uint32]135
        $res = $mc.InvokeMethod('InitiateClientOperationEx', $in, $null)
        if ($res.OperationID) { $op = [uint32]$res.OperationID }
        if ($res.ReturnValue -ne 0) {
            $r.Result = "Script gonderilemedi (ReturnValue=$($res.ReturnValue))."
            return $r
        }
    }
    catch {
        $why = Get-WmiFailureReason -ErrorRecord $_
        if ($why.Kind -eq 'Access') { $r.Result = 'Script gonderilemedi: SCCM yetkisi yok (Run Script izni gerekli).' }
        else { $r.Result = "Script gonderilemedi: $($_.Exception.Message)" }
        return $r
    }

    # ---- 3) Sonucu bekle ----
    $start    = Get-Date
    $deadline = $start.AddSeconds($TimeoutSec)
    $taskId   = $null
    $status   = $null
    $seen     = $null

    while ((Get-Date) -lt $deadline -and -not $status) {
        Start-Sleep -Seconds 4

        if (-not $taskId -and $op) {
            try {
                $t = @(Get-WmiObject -ComputerName $SmsProvider -Namespace $ns -ErrorAction Stop `
                    -Query "SELECT TaskID FROM SMS_ScriptsExecutionTask WHERE ClientOperationId = $op")
                if ($t.Count -gt 0) { $taskId = [string]$t[0].TaskID }
            }
            catch { }
        }

        $rows = @()
        try {
            $rows = @(Get-WmiObject -ComputerName $SmsProvider -Namespace $ns -ErrorAction Stop `
                -Query "SELECT * FROM SMS_ScriptsExecutionStatus WHERE ScriptGuid = '$ScriptGuid' AND ResourceID = $rid")
        }
        catch { }

        foreach ($row in $rows) {
            # Bu satir BU calistirmaya mi ait? (eski calistirmalarin sonucunu almamak icin)
            $names = @($row.PSObject.Properties | ForEach-Object { $_.Name })
            $mine = $false
            if ($taskId -and ($names -contains 'TaskID') -and $row.TaskID) {
                $mine = ([string]$row.TaskID -eq $taskId)
            }
            elseif (($names -contains 'ClientOperationId') -and $row.ClientOperationId -and $op) {
                $mine = ([uint32]$row.ClientOperationId -eq $op)
            }
            elseif (($names -contains 'LastUpdateTime') -and $row.LastUpdateTime) {
                try { $mine = ([System.Management.ManagementDateTimeConverter]::ToDateTime([string]$row.LastUpdateTime) -ge $start.AddMinutes(-1)) }
                catch { $mine = $false }
            }
            if (-not $mine) { continue }

            $seen = $row
            if ($null -ne $row.ScriptExitCode -or -not [string]::IsNullOrEmpty([string]$row.ScriptOutput)) {
                $status = $row
                break
            }
        }
    }

    # ---- 4) Sonuc ----
    if ($status) {
        $code    = $status.ScriptExitCode
        $outText = ([string]$status.ScriptOutput).TrimStart([char]0xFEFF).Trim()

        # Grid'e sadece CIKTI yazilir (satirlar ' | ' ile birlestirilir); exit code / OperationID Detay'dadir.
        $oneLine = ((($outText -split "`r`n|`n|`r") | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ' | ')

        $r.Output   = $outText
        $r.ExitCode = $code
        $r.Detail   = "OperationID: $op`r`nExit code: $code`r`nCikti:`r`n$outText"

        $ok = ($null -eq $code) -or ([int]$code -eq 0)
        $r.State = if ($ok) { 'ok' } else { 'warn' }

        if ($outText)    { $r.Result = $oneLine }
        elseif ($ok)     { $r.Result = 'Script calisti, cikti bos.' }
        else             { $r.Result = "Script calisti, cikti bos (exit code $code)." }
        return $r
    }

    $v = Get-NoResponseVerdict -ComputerName $ComputerName -TimeoutSec $TimeoutSec
    $r.State  = $v.State
    $r.Result = $v.Text
    $r.Detail = "OperationID: $op"
    if ($seen) { $r.Detail += "`r`nSCCM'de durum kaydi goruldu ancak sonuc/cikti gelmedi." }
    return $r
}

function Get-AdStructureStatus {
    # AD YAPISI KONTROLU: DNS + Cevrimici (ping) + Active Directory (var mi, Enabled/Disabled, Last Logon).
    # Servis kontrolu YAPMAZ. Script gonderiminde otomatik, ayrica tek basina da calistirilabilir.
    param([string]$ComputerName, [bool]$QueryAllDCs = $false)

    $row = [pscustomobject]@{
        Hostname     = $ComputerName
        DnsStatus    = ''
        IsOnline     = $false
        AdStatus     = ''
        LastLogon    = ''
        ErrorMessage = ''
        State        = 'fail'
    }
    $errs = @()

    $dns = Resolve-HostDns -ComputerName $ComputerName
    $row.DnsStatus = $dns.Text

    $ad = Get-AdComputerInfo -ComputerName $ComputerName -QueryAllDCs $QueryAllDCs
    $row.AdStatus  = $ad.Status
    $row.LastLogon = $ad.LastLogon
    if ($ad.Detail) { $errs += $ad.Detail }

    if (-not $dns.Ok) {
        $errs += 'DNS cozumlenemedi, ping kontrolu yapilmadi'
    }
    else {
        $row.IsOnline = [bool](Test-HostOnline -ComputerName $ComputerName)
        if (-not $row.IsOnline) { $errs += 'Makine yanit vermiyor (ping basarisiz)' }
    }

    $row.ErrorMessage = ($errs -join ' | ')
    if (-not $dns.Ok -or -not $row.IsOnline) { $row.State = 'fail' }
    elseif ($ad.Status -ne 'Enabled')         { $row.State = 'warn' }
    else                                      { $row.State = 'ok' }
    return $row
}

function Get-ServiceCheck {
    # Istenen servislerin + CcmExec'in durumunu tek RPC cagrisi ile alir (Panel 1 script gonderiminde kullanilir).
    param([string]$ComputerName, [string[]]$ServiceNames)

    $r = [pscustomobject]@{ Running = ''; NotRunning = ''; CcmExec = ''; ServiceDetail = @(); Note = '' }
    $all = $null
    try { $all = @(Get-Service -ComputerName $ComputerName -ErrorAction Stop) }
    catch { $r.Note = "Servis listesi alinamadi: $($_.Exception.Message)" }

    $detail = @()
    foreach ($sn in $ServiceNames) {
        $n = "$sn".Trim()
        if (-not $n) { continue }
        if ($null -eq $all) {
            $detail += [pscustomobject]@{ Name = $n; Status = 'Erisilemedi' }
            continue
        }
        $m = $all | Where-Object { $_.Name -eq $n -or $_.DisplayName -eq $n } | Select-Object -First 1
        if ($m) { $detail += [pscustomobject]@{ Name = $n; Status = $m.Status.ToString() } }
        else    { $detail += [pscustomobject]@{ Name = $n; Status = 'NotFound' } }
    }
    $r.ServiceDetail = $detail
    $r.Running    = (($detail | Where-Object { $_.Status -eq 'Running' } | ForEach-Object { $_.Name }) -join ', ')
    $r.NotRunning = (($detail | Where-Object { $_.Status -ne 'Running' } | ForEach-Object { "$($_.Name)($($_.Status))" }) -join ', ')

    if ($null -eq $all) { $r.CcmExec = 'Erisilemedi' }
    else {
        $cc = $all | Where-Object { $_.Name -eq 'CcmExec' } | Select-Object -First 1
        $r.CcmExec = if ($cc) { $cc.Status.ToString() } else { 'Yok' }
    }
    return $r
}

function Invoke-ScriptWithAdCheck {
    # Once AD yapisi kontrolu (DNS/ping/AD) + servis kontrolu, sonra SCCM script gonderimi. Tek sonuc nesnesi doner.
    param(
        [string]$ComputerName,
        [string]$SmsProvider,
        [string]$SiteCode,
        [string]$ScriptGuid,
        [string]$ScriptName,
        [string]$ParamB64,
        [int]$TimeoutSec,
        [bool]$QueryAllDCs = $false,
        [string[]]$ServiceNames = @()
    )

    $ad = Get-AdStructureStatus -ComputerName $ComputerName -QueryAllDCs $QueryAllDCs

    if ($ad.IsOnline -and @($ServiceNames).Count -gt 0) {
        $svc = Get-ServiceCheck -ComputerName $ComputerName -ServiceNames $ServiceNames
    }
    else {
        $svc = [pscustomobject]@{ Running = ''; NotRunning = ''; CcmExec = $(if (-not $ad.IsOnline) { 'Kontrol edilemedi (cevrimdisi)' } else { '' }); ServiceDetail = @(); Note = '' }
    }

    $sr = Invoke-SccmScriptOnDevice -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode `
        -ScriptGuid $ScriptGuid -ScriptName $ScriptName -ParamB64 $ParamB64 -TimeoutSec $TimeoutSec

    return [pscustomobject]@{
        Hostname      = $ComputerName
        DnsStatus     = $ad.DnsStatus
        IsOnline      = $ad.IsOnline
        AdStatus      = $ad.AdStatus
        LastLogon     = $ad.LastLogon
        AdNote        = $ad.ErrorMessage
        Running       = $svc.Running
        NotRunning    = $svc.NotRunning
        CcmExec       = $svc.CcmExec
        ServiceDetail = $svc.ServiceDetail
        SvcNote       = $svc.Note
        Action        = $sr.Action
        State         = $sr.State
        Result        = $sr.Result
        Output        = $sr.Output
        ExitCode      = $sr.ExitCode
        Detail        = $sr.Detail
    }
}

function Get-SccmRegistration {
    # SCCM sitesinde cihaz kaydi + istemci (ajan) durumu (SMS Provider WMI sorgusu).
    param([string]$ComputerName, [string]$SmsProvider, [string]$SiteCode)

    if ([string]::IsNullOrWhiteSpace($SmsProvider) -or [string]::IsNullOrWhiteSpace($SiteCode)) {
        return 'Sorgulanmadi (SCCM baglantisi yok)'
    }
    $short = ($ComputerName -split '\.')[0]
    try {
        $cands = @(Get-WmiObject -ComputerName $SmsProvider -Namespace "root\SMS\site_$SiteCode" -ErrorAction Stop `
            -Query "SELECT ResourceId, Name, Client, Active, Obsolete FROM SMS_R_System WHERE Name = '$short'")
    }
    catch { return "Sorgulanamadi: $($_.Exception.Message)" }

    if ($cands.Count -eq 0) { return 'SCCM de kayitli degil' }
    $ok = $cands | Where-Object { $_.Client -eq 1 -and $_.Obsolete -ne 1 } | Select-Object -First 1
    if ($ok) {
        if ($ok.Active -eq 1) { return 'Client yuklu (Active)' }
        return 'Client yuklu (Inactive)'
    }
    return 'Kayit var, client yok / obsolete'
}

function Get-ParserHostCheck {
    # PANEL 2 icin kontrol: DNS + ping + AD (Enabled/Disabled/Last Logon) + CcmExec servisi + SCCM kaydi.
    # Panel 1'in sonuc tablosuna / detaylarina DOKUNMAZ; sadece nesne dondurur.
    param([string]$ComputerName, [string]$SmsProvider = '', [string]$SiteCode = '', [bool]$QueryAllDCs = $false)

    $c = [pscustomobject]@{
        Hostname  = $ComputerName
        DnsStatus = ''
        IsOnline  = $false
        AdStatus  = ''
        LastLogon = ''
        CcmExec   = ''
        SccmKayit = ''
        Note      = ''
    }
    $notes = @()

    $dns = Resolve-HostDns -ComputerName $ComputerName
    $c.DnsStatus = $dns.Text

    $ad = Get-AdComputerInfo -ComputerName $ComputerName -QueryAllDCs $QueryAllDCs
    $c.AdStatus  = $ad.Status
    $c.LastLogon = $ad.LastLogon
    if ($ad.Detail) { $notes += $ad.Detail }

    $c.SccmKayit = Get-SccmRegistration -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode

    if (-not $dns.Ok) {
        $c.CcmExec = 'Kontrol edilemedi (DNS)'
        $notes += 'DNS cozumlenemedi, ping / CcmExec kontrolu yapilmadi'
        $c.Note = ($notes -join ' | ')
        return $c
    }

    $c.IsOnline = [bool](Test-HostOnline -ComputerName $ComputerName)
    if (-not $c.IsOnline) {
        $c.CcmExec = 'Kontrol edilemedi (cevrimdisi)'
        $notes += 'Makine yanit vermiyor (ping basarisiz)'
        $c.Note = ($notes -join ' | ')
        return $c
    }

    try {
        $svc = Get-Service -ComputerName $ComputerName -Name 'CcmExec' -ErrorAction Stop
        $c.CcmExec = $svc.Status.ToString()
    }
    catch {
        if ($_.FullyQualifiedErrorId -like 'NoServiceFoundForGivenName*') { $c.CcmExec = 'Yok' }
        else {
            $c.CcmExec = 'Erisilemedi'
            $notes += "CcmExec sorgulanamadi: $($_.Exception.Message)"
        }
    }
    $c.Note = ($notes -join ' | ')
    return $c
}

function Invoke-ParserScriptRun {
    # PANEL 2: once DNS/AD/SCCM ajan kontrolu, sonra SCCM script gonderimi. Tek sonuc nesnesi doner.
    param(
        [string]$ComputerName,
        [string]$SmsProvider,
        [string]$SiteCode,
        [string]$ScriptGuid,
        [string]$ScriptName,
        [string]$ParamB64,
        [int]$TimeoutSec,
        [bool]$QueryAllDCs = $false
    )

    $chk = Get-ParserHostCheck -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode -QueryAllDCs $QueryAllDCs
    $sr  = Invoke-SccmScriptOnDevice -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode `
        -ScriptGuid $ScriptGuid -ScriptName $ScriptName -ParamB64 $ParamB64 -TimeoutSec $TimeoutSec

    return [pscustomobject]@{
        Hostname = $ComputerName
        Check    = $chk
        Action   = $sr.Action
        State    = $sr.State
        Result   = $sr.Result
        Output   = $sr.Output
        ExitCode = $sr.ExitCode
        Detail   = $sr.Detail
    }
}
# ==================== SADECE ANA PENCERE FONKSIYONLARI ====================

function New-WorkerScript {
    # Verilen fonksiyonlarin kaynak kodunu runspace scriptinin basina gomer;
    # boylece fonksiyon tanimlari TEK yerde tutulur.
    param([string]$ParamLine, [string[]]$Funcs, [string]$CallLine)

    $defs = foreach ($n in $Funcs) {
        $body = (Get-Item -Path "function:$n").ScriptBlock.ToString()
        "function $n {`n$body`n}"
    }
    return [scriptblock]::Create("param($ParamLine)`n" + ($defs -join "`n") + "`n$CallLine")
}

$script:CheckWorker = New-WorkerScript `
    -ParamLine '$ComputerName, $ServiceNames, $QueryAllDCs' `
    -Funcs @('Test-HostOnline', 'Resolve-HostDns', 'Convert-FileTimeSafe', 'Get-AdComputerInfo', 'Get-RemoteStatus') `
    -CallLine 'Get-RemoteStatus -ComputerName $ComputerName -ServiceNames $ServiceNames -QueryAllDCs $QueryAllDCs'

$script:FixWorker = New-WorkerScript `
    -ParamLine '$ComputerName, $ServiceDetail' `
    -Funcs @('Start-RemoteFix') `
    -CallLine 'Start-RemoteFix -ComputerName $ComputerName -ServiceDetail $ServiceDetail'

$script:ActionWorker = New-WorkerScript `
    -ParamLine '$ComputerName, $ActionName, $ScheduleIds' `
    -Funcs @('Test-HostOnline', 'Resolve-HostDns', 'Get-WmiFailureReason', 'Invoke-SccmClientAction') `
    -CallLine 'Invoke-SccmClientAction -ComputerName $ComputerName -ActionName $ActionName -ScheduleIds $ScheduleIds'

$script:ScriptWorker = New-WorkerScript `
    -ParamLine '$ComputerName, $SmsProvider, $SiteCode, $ScriptGuid, $ScriptName, $ParamB64, $TimeoutSec, $QueryAllDCs, $ServiceNames' `
    -Funcs @('Test-HostOnline', 'Resolve-HostDns', 'Convert-FileTimeSafe', 'Get-AdComputerInfo', 'Get-AdStructureStatus', 'Get-ServiceCheck', 'Get-WmiFailureReason', 'Get-NoResponseVerdict', 'Invoke-SccmScriptOnDevice', 'Invoke-ScriptWithAdCheck') `
    -CallLine 'Invoke-ScriptWithAdCheck -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode -ScriptGuid $ScriptGuid -ScriptName $ScriptName -ParamB64 $ParamB64 -TimeoutSec $TimeoutSec -QueryAllDCs $QueryAllDCs -ServiceNames $ServiceNames'

$script:AdWorker = New-WorkerScript `
    -ParamLine '$ComputerName, $QueryAllDCs' `
    -Funcs @('Test-HostOnline', 'Resolve-HostDns', 'Convert-FileTimeSafe', 'Get-AdComputerInfo', 'Get-AdStructureStatus') `
    -CallLine 'Get-AdStructureStatus -ComputerName $ComputerName -QueryAllDCs $QueryAllDCs'

$script:ParserCheckWorker = New-WorkerScript `
    -ParamLine '$ComputerName, $SmsProvider, $SiteCode, $QueryAllDCs' `
    -Funcs @('Test-HostOnline', 'Resolve-HostDns', 'Convert-FileTimeSafe', 'Get-AdComputerInfo', 'Get-SccmRegistration', 'Get-ParserHostCheck') `
    -CallLine 'Get-ParserHostCheck -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode -QueryAllDCs $QueryAllDCs'

$script:ParserScriptWorker = New-WorkerScript `
    -ParamLine '$ComputerName, $SmsProvider, $SiteCode, $ScriptGuid, $ScriptName, $ParamB64, $TimeoutSec, $QueryAllDCs' `
    -Funcs @('Test-HostOnline', 'Resolve-HostDns', 'Convert-FileTimeSafe', 'Get-AdComputerInfo', 'Get-SccmRegistration', 'Get-ParserHostCheck', 'Get-WmiFailureReason', 'Get-NoResponseVerdict', 'Invoke-SccmScriptOnDevice', 'Invoke-ParserScriptRun') `
    -CallLine 'Invoke-ParserScriptRun -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode -ScriptGuid $ScriptGuid -ScriptName $ScriptName -ParamB64 $ParamB64 -TimeoutSec $TimeoutSec -QueryAllDCs $QueryAllDCs'

function Get-HostSummary {
    # Bir hostun servis baslatma sonuclarini tek satirlik metne cevirir.
    param([pscustomobject[]]$HostResults)

    $skip = @('Running', 'NotFound', 'Erisilemedi')
    $fixed        = @($HostResults | Where-Object { $_.OncekiDurum -notin $skip -and $_.SonrakiDurum -eq 'Running' })
    $stillStopped = @($HostResults | Where-Object { $_.OncekiDurum -notin $skip -and $_.SonrakiDurum -ne 'Running' })
    $notfound     = @($HostResults | Where-Object { $_.OncekiDurum -eq 'NotFound' })
    $noaccess     = @($HostResults | Where-Object { $_.OncekiDurum -eq 'Erisilemedi' })

    $parts = @()
    if ($fixed.Count -gt 0)        { $parts += "$(($fixed        | ForEach-Object { $_.ServiceName }) -join ',') stopped durumda running yapildi" }
    if ($stillStopped.Count -gt 0) { $parts += "$(($stillStopped | ForEach-Object { $_.ServiceName }) -join ',') stopped durumda kaldi" }
    if ($notfound.Count -gt 0)     { $parts += "$(($notfound     | ForEach-Object { $_.ServiceName }) -join ',') bulunamadi" }
    if ($noaccess.Count -gt 0)     { $parts += "$(($noaccess     | ForEach-Object { $_.ServiceName }) -join ',') erisilemedi" }

    if ($parts.Count -eq 0) { return 'Tum servisler calisiyor.' }
    return (($parts -join ', ') + '.')
}

# ---------- Veri klasoru / ayarlar / denetim kaydi ----------

function Get-DataDir {
    # Yol bilerek sabit kuruluyor (OneDrive klasor yonlendirmesinden etkilenmesin diye).
    $dir = "C:\Users\$env:USERNAME\Desktop\ServiceStatus"
    if (-not (Test-Path $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
    return $dir
}

function Get-SettingsFile { return (Join-Path (Get-DataDir) 'ClientToolSettings.json') }

function Write-ActionLog {
    param([string]$Hostname, [string]$Action, [string]$State, [string]$Text)
    try {
        $clean = ("$Text" -replace '\s+', ' ').Trim()
        $line = "{0}`t{1}`t{2}`t{3}`t{4}`t{5}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $env:USERNAME, $Hostname, $Action, $State, $clean
        $file = Join-Path (Get-DataDir) ("ClientTool_{0}.log" -f (Get-Date -Format 'yyyyMM'))
        Add-Content -Path $file -Value $line -Encoding UTF8
    }
    catch { }
}

# ---------- Sonuc tablosu yardimcilari ----------

$script:ResultsTable = New-Object System.Data.DataTable
foreach ($c in @('Hostname', 'DnsStatus', 'IsOnline', 'AdStatus', 'LastLogon', 'Running', 'NotRunning', 'CcmExec', 'LastAction', 'ActionResult', 'ErrorMessage', 'State')) {
    $null = $script:ResultsTable.Columns.Add($c)
}
$script:BindingSrc = New-Object System.Windows.Forms.BindingSource
$script:BindingSrc.DataSource = $script:ResultsTable

# Panel 2 (Update Parser) tablosu: 21 parser kolonu + 6 kontrol kolonu (DNS/Online/AD/LastLogon/CcmExec/SCCM) + KolonUyumu + SourceHost (gizli)
$script:ParseTable = New-Object System.Data.DataTable
foreach ($c in $script:ParserColumns) { $null = $script:ParseTable.Columns.Add($c) }
foreach ($c in $script:ParserCheckColumns) { $null = $script:ParseTable.Columns.Add($c) }
$null = $script:ParseTable.Columns.Add('KolonUyumu')
$null = $script:ParseTable.Columns.Add('SourceHost')
$script:ParseBinding = New-Object System.Windows.Forms.BindingSource
$script:ParseBinding.DataSource = $script:ParseTable

function Get-GridRow {
    param([string]$Hostname)
    $found = $script:ResultsTable.Select("Hostname = '" + $Hostname.Replace("'", "''") + "'")
    if ($found.Count -gt 0) { return $found[0] }
    $row = $script:ResultsTable.NewRow()
    $row['Hostname'] = $Hostname
    $script:ResultsTable.Rows.Add($row)
    return $row
}

function Set-GridValues {
    param([string]$Hostname, [hashtable]$Values)
    $row = Get-GridRow -Hostname $Hostname
    foreach ($k in $Values.Keys) { $row[$k] = [string]$Values[$k] }
}

function Add-RowDetail {
    param([string]$Hostname, [string]$Text)
    if (-not $script:Details.ContainsKey($Hostname)) { $script:Details[$Hostname] = '' }
    $script:Details[$Hostname] += "[$(Get-Date -Format 'HH:mm:ss')] $Text`r`n`r`n"
}

# ---------- "Bekleniyor.." yer tutucu satirlari (Panel 1) ----------
$script:PendingText = 'Bekleniyor..'

function Set-GridPending {
    # Hedef makinelerin TUMUNU hemen tabloya ekler (yoksa satir olusturur, varsa yerinde gunceller);
    # verilen kolonlara 'Bekleniyor..' yazar. Sonuc gelince Update-GridFrom* fonksiyonlari ayni satiri doldurur.
    param([string[]]$Hosts, [string[]]$Columns, [switch]$ClearError)
    foreach ($h in $Hosts) {
        $row = Get-GridRow -Hostname $h
        foreach ($c in $Columns) { $row[$c] = $script:PendingText }
        if ($ClearError) { $row['ErrorMessage'] = '' }
        $row['State'] = 'pending'
    }
}

function Resolve-PendingCells {
    # Islem hata verdigi / sonuc gelmedigi icin hala 'Bekleniyor..' kalan hucreleri doldurur.
    param([string]$Hostname, [string]$Text = 'Alinamadi')
    $row = Get-GridRow -Hostname $Hostname
    foreach ($c in $script:ResultsTable.Columns) {
        if ([string]$row[$c.ColumnName] -eq $script:PendingText) { $row[$c.ColumnName] = $Text }
    }
}

# ---------- Paralel calistirici ----------
# Panel 1 ve Panel 2 istekleri BAGIMSIZ calisir (ayri durum etiketi, ayri sonuc isleyici, ayri tablo).
# Yalnizca RunspacePool ortaktir: toplamda ayni anda en fazla $script:MaxParallel makineye istek gider.

$script:RunTimer = New-Object System.Windows.Forms.Timer
$script:RunTimer.Interval = 200

function Get-SharedPool {
    if (-not $script:SharedPool -or $script:SharedPool.RunspacePoolStateInfo.State -ne 'Opened') {
        $script:SharedPool = [runspacefactory]::CreateRunspacePool(1, $script:MaxParallel)
        $script:SharedPool.Open()
    }
    return $script:SharedPool
}

function Test-PanelBusy {
    param([int]$Panel)
    return (@($script:ActiveRuns | Where-Object { $_.Panel -eq $Panel }).Count -gt 0)
}

function Start-ParallelRun {
    param(
        [string]$Title,
        [scriptblock]$Worker,
        [object[]]$Items,          # her biri: @{ Key = hostname; Params = @{...} }
        [string]$OnResultFn,
        [int]$Panel = 1            # 1 = Panel 1 (kontrol/script), 2 = Panel 2 (update parser)
    )

    $st = [pscustomobject]@{
        Panel      = $Panel
        Pool       = (Get-SharedPool)
        Jobs       = $null
        Total      = @($Items).Count
        Done       = 0
        Title      = $Title
        OnResultFn = $OnResultFn
    }
    [void]$script:ActiveRuns.Add($st)
    Set-UiBusy -Panel $Panel -Busy $true

    $jobs = New-Object System.Collections.ArrayList
    foreach ($it in $Items) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $st.Pool
        [void]$ps.AddScript($Worker).AddParameters($it.Params)
        [void]$jobs.Add([pscustomobject]@{ PS = $ps; Async = $ps.BeginInvoke(); Key = $it.Key; Done = $false })
    }
    $st.Jobs = $jobs

    $lbl = if ($Panel -eq 2) { $lblParserStatus } else { $lblStatus }
    $lbl.ForeColor = [System.Drawing.Color]::DarkBlue
    $lbl.Text = "$Title`: 0 / $($st.Total) tamamlandi (ayni anda en fazla $($script:MaxParallel) istek)"
    if (-not $script:RunTimer.Enabled) { $script:RunTimer.Start() }
}

$script:RunTimer.Add_Tick({
    foreach ($st in @($script:ActiveRuns)) {
        $lbl = if ($st.Panel -eq 2) { $lblParserStatus } else { $lblStatus }
        $pending = $false

        foreach ($job in $st.Jobs) {
            if ($job.Done) { continue }

            if ($job.Async.IsCompleted) {
                $job.Done = $true
                $out = @()
                $errText = ''
                try { $out = @($job.PS.EndInvoke($job.Async)) } catch { $errText = $_.Exception.Message }
                if ($out.Count -eq 0 -and -not $errText -and $job.PS.Streams.Error.Count -gt 0) {
                    $errText = $job.PS.Streams.Error[0].ToString()
                }
                $job.PS.Dispose()

                try {
                    $fn = $st.OnResultFn
                    & $fn -Key $job.Key -Output $out -ErrorText $errText
                }
                catch {
                    $lbl.Text = "Sonuc isleme hatasi: $($_.Exception.Message)"
                }

                $st.Done++
                $lbl.Text = "$($st.Title): $($st.Done) / $($st.Total) tamamlandi (ayni anda en fazla $($script:MaxParallel) istek)"
            }
            else {
                $pending = $true
            }
        }

        if (-not $pending) {
            [void]$script:ActiveRuns.Remove($st)
            $lbl.Text = "$($st.Title) tamamlandi. $($st.Total) makine islendi."
            Set-UiBusy -Panel $st.Panel -Busy $false
            if ($st.Panel -eq 2) { Update-ParserInfo }
        }
    }

    if ($script:ActiveRuns.Count -eq 0) { $script:RunTimer.Stop() }
})

# ---------- Sonuc isleyiciler (UI thread'inde calisir) ----------

function Update-GridFromCheck {
    param($Key, $Output, $ErrorText)
    $r = @($Output)[0]
    if (-not $r -or -not $r.PSObject.Properties['DnsStatus']) {
        Resolve-PendingCells -Hostname $Key
        Set-GridValues -Hostname $Key -Values @{ ErrorMessage = "Islem hatasi: $ErrorText"; State = 'fail' }
        return
    }
    $script:LastResults = @($script:LastResults) + $r
    # State 'fail' = DNS cozumlenemedi / ping yok -> servis bilgisi alinamadi (ulasilamadi)
    $un = ($r.State -eq 'fail')
    Set-GridValues -Hostname $r.Hostname -Values @{
        DnsStatus    = $r.DnsStatus
        IsOnline     = $(if ($r.IsOnline) { 'Evet' } else { 'Hayir' })
        AdStatus     = $r.AdStatus
        LastLogon    = $r.LastLogon
        Running      = $(if ($un) { 'Ulasilamadi' } else { $r.Running })
        NotRunning   = $(if ($un) { 'Ulasilamadi' } else { $r.NotRunning })
        CcmExec      = $(if ($un) { 'Ulasilamadi' } else { $r.CcmExec })
        ErrorMessage = $r.ErrorMessage
        State        = $r.State
    }
    Add-RowDetail -Hostname $r.Hostname -Text ("Kontrol: DNS=$($r.DnsStatus); Online=$($r.IsOnline); AD=$($r.AdStatus); LastLogon=$($r.LastLogon); CcmExec=$($r.CcmExec)" + $(if ($r.ErrorMessage) { "`r`nNot: $($r.ErrorMessage)" } else { '' }))
}

function Update-GridFromFix {
    param($Key, $Output, $ErrorText)
    $rows = @($Output | Where-Object { $_.PSObject.Properties['ServiceName'] })
    if ($rows.Count -eq 0) {
        Set-GridValues -Hostname $Key -Values @{ LastAction = 'Servis baslat'; ActionResult = "Islem hatasi: $ErrorText"; State = 'fail' }
        return
    }
    $summary = Get-HostSummary -HostResults $rows
    $newRunning    = (($rows | Where-Object { $_.SonrakiDurum -eq 'Running' } | ForEach-Object { $_.ServiceName }) -join ', ')
    $newNotRunning = (($rows | Where-Object { $_.SonrakiDurum -ne 'Running' } | ForEach-Object { "$($_.ServiceName)($($_.SonrakiDurum))" }) -join ', ')

    # Sonraki "Servisleri Baslat" tiklamasi guncel duruma gore calissin
    foreach ($e in @($script:LastResults | Where-Object { $_.Hostname -eq $Key })) {
        $e.ServiceDetail = @($rows | ForEach-Object { [pscustomobject]@{ Name = $_.ServiceName; Status = $_.SonrakiDurum } })
    }

    Set-GridValues -Hostname $Key -Values @{
        Running      = $newRunning
        NotRunning   = $newNotRunning
        LastAction   = "Servis baslat ($(Get-Date -Format 'HH:mm:ss'))"
        ActionResult = $summary
        State        = $(if ($newNotRunning) { 'warn' } else { 'ok' })
    }
    Add-RowDetail -Hostname $Key -Text ("Servis baslatma:`r`n" + (($rows | ForEach-Object { "  $($_.ServiceName): $($_.OncekiDurum) -> $($_.SonrakiDurum) ($($_.Aciklama))" }) -join "`r`n"))
    Write-ActionLog -Hostname $Key -Action 'Servis baslat' -State $(if ($newNotRunning) { 'warn' } else { 'ok' }) -Text $summary
}

function Update-GridFromAction {
    # Hem SCCM istemci eylemleri hem script gonderimi icin ortak isleyici
    param($Key, $Output, $ErrorText)
    $r = @($Output)[0]
    if (-not $r -or -not $r.PSObject.Properties['Action']) {
        $r = [pscustomobject]@{ Hostname = $Key; Action = 'Islem'; State = 'fail'; Result = "Islem hatasi: $ErrorText"; Detail = '' }
    }
    $one = ("$($r.Result)" -replace '\s+', ' ').Trim()
    Set-GridValues -Hostname $r.Hostname -Values @{
        LastAction   = "$($r.Action) ($(Get-Date -Format 'HH:mm:ss'))"
        ActionResult = $one
        State        = $r.State
    }
    $d = "$($r.Action): $($r.Result)"
    if ($r.Detail) { $d += "`r`n$($r.Detail)" }
    Add-RowDetail -Hostname $r.Hostname -Text $d
    Write-ActionLog -Hostname $r.Hostname -Action $r.Action -State $r.State -Text $r.Result
}

function Update-GridFromScript {
    # Script gonderimi sonucu (SADECE Panel 1): DNS/Online/AD/LastLogon + script CIKTISI (sadece cikti).
    # Panel 2 (Update Parser) bu isleyiciden etkilenmez; kendi isleyicisi Update-ParserFromScript'tir.
    param($Key, $Output, $ErrorText)

    $r = @($Output | Where-Object { $_ -and $_.PSObject.Properties['Output'] })[0]
    if (-not $r) {
        $msg = "Islem hatasi: $ErrorText"
        Resolve-PendingCells -Hostname $Key
        Set-GridValues -Hostname $Key -Values @{ LastAction = "Script ($(Get-Date -Format 'HH:mm:ss'))"; ActionResult = $msg; State = 'fail' }
        Add-RowDetail -Hostname $Key -Text "Script: $msg"
        Write-ActionLog -Hostname $Key -Action 'Script' -State 'fail' -Text $msg
        return
    }

    $notes = @()
    if ($r.AdNote)  { $notes += $r.AdNote }
    if ($r.SvcNote) { $notes += $r.SvcNote }
    if ($null -ne $r.ExitCode -and [int]$r.ExitCode -ne 0) { $notes += "Script exit code: $($r.ExitCode)" }

    # Servis sorunu varsa (calismayan servis / CcmExec yok) satir en az 'warn' olur
    $state = $r.State
    if ($state -eq 'ok' -and ($r.NotRunning -or $r.CcmExec -eq 'Yok')) { $state = 'warn' }

    # Makineye ulasilamadiysa servis kolonlari bos kalmasin
    $un = (-not $r.IsOnline)

    Set-GridValues -Hostname $r.Hostname -Values @{
        DnsStatus    = $r.DnsStatus
        IsOnline     = $(if ($r.IsOnline) { 'Evet' } else { 'Hayir' })
        AdStatus     = $r.AdStatus
        LastLogon    = $r.LastLogon
        Running      = $(if ($un -and -not $r.Running) { 'Ulasilamadi' } else { $r.Running })
        NotRunning   = $(if ($un -and -not $r.NotRunning) { 'Ulasilamadi' } else { $r.NotRunning })
        CcmExec      = $(if ($un -and -not $r.CcmExec) { 'Ulasilamadi' } else { $r.CcmExec })
        LastAction   = "$($r.Action) ($(Get-Date -Format 'HH:mm:ss'))"
        ActionResult = ("$($r.Result)" -replace '\s+', ' ').Trim()
        ErrorMessage = ($notes -join ' | ')
        State        = $state
    }

    # "Servisleri Baslat" butonu bu makineyi de kapsasin
    $script:LastResults = @(@($script:LastResults | Where-Object { $_.Hostname -ne $r.Hostname }) + [pscustomobject]@{
        Hostname = $r.Hostname; IsOnline = [bool]$r.IsOnline; ServiceDetail = @($r.ServiceDetail) })

    $d = "$($r.Action): $($r.Result)`r`nAD: $($r.AdStatus); LastLogon: $($r.LastLogon); DNS: $($r.DnsStatus); Online: $($r.IsOnline); CcmExec: $($r.CcmExec)"
    if (@($r.ServiceDetail).Count -gt 0) {
        $d += "`r`nServisler: " + ((@($r.ServiceDetail) | ForEach-Object { "$($_.Name)=$($_.Status)" }) -join ', ')
    }
    if ($r.Detail) { $d += "`r`n$($r.Detail)" }
    Add-RowDetail -Hostname $r.Hostname -Text $d
    Write-ActionLog -Hostname $r.Hostname -Action $r.Action -State $state -Text $r.Result
}

function Update-GridFromAd {
    # "AD Yapisi Kontrolu" sonucu (script gondermeden)
    param($Key, $Output, $ErrorText)

    $r = @($Output | Where-Object { $_ -and $_.PSObject.Properties['AdStatus'] -and $_.PSObject.Properties['DnsStatus'] })[0]
    if (-not $r) {
        $msg = "Islem hatasi: $ErrorText"
        Resolve-PendingCells -Hostname $Key
        Set-GridValues -Hostname $Key -Values @{ LastAction = "AD yapisi kontrolu ($(Get-Date -Format 'HH:mm:ss'))"; ActionResult = $msg; State = 'fail' }
        Add-RowDetail -Hostname $Key -Text "AD yapisi kontrolu: $msg"
        return
    }

    $summary = "AD: $($r.AdStatus); Last Logon: $($r.LastLogon)"
    Set-GridValues -Hostname $r.Hostname -Values @{
        DnsStatus    = $r.DnsStatus
        IsOnline     = $(if ($r.IsOnline) { 'Evet' } else { 'Hayir' })
        AdStatus     = $r.AdStatus
        LastLogon    = $r.LastLogon
        LastAction   = "AD yapisi kontrolu ($(Get-Date -Format 'HH:mm:ss'))"
        ActionResult = $summary
        ErrorMessage = $r.ErrorMessage
        State        = $r.State
    }
    Add-RowDetail -Hostname $r.Hostname -Text ("AD yapisi kontrolu: DNS=$($r.DnsStatus); Online=$($r.IsOnline); AD=$($r.AdStatus); LastLogon=$($r.LastLogon)" + $(if ($r.ErrorMessage) { "`r`nNot: $($r.ErrorMessage)" } else { '' }))
    Write-ActionLog -Hostname $r.Hostname -Action 'AD yapisi kontrolu' -State $r.State -Text $summary
}

# ---------- SCCM yardimcilari (UI thread'i) ----------

function Get-CmSiteCode {
    param([string]$Server)
    $locs = @(Get-WmiObject -ComputerName $Server -Namespace 'root\SMS' -Class 'SMS_ProviderLocation' -ErrorAction Stop)
    $loc = $locs | Where-Object { $_.ProviderForLocalSite } | Select-Object -First 1
    if (-not $loc) { $loc = $locs | Select-Object -First 1 }
    if (-not $loc) { throw 'SMS_ProviderLocation bulunamadi (SMS Provider adini kontrol edin).' }
    return [string]$loc.SiteCode
}

function Get-CmScriptList {
    param([string]$Server, [string]$SiteCode)
    $ns = "root\SMS\site_$SiteCode"
    $all = @(Get-WmiObject -ComputerName $Server -Namespace $ns -Class 'SMS_Scripts' -ErrorAction Stop)
    $list = foreach ($s in ($all | Sort-Object ScriptName)) {
        $appr = switch ([int]$s.ApprovalState) {
            0 { 'Onay bekliyor' }
            1 { 'Reddedildi' }
            3 { 'Onayli' }
            default { "Durum $($s.ApprovalState)" }
        }
        [pscustomobject]@{
            Guid         = [string]$s.ScriptGuid
            Name         = [string]$s.ScriptName
            Author       = [string]$s.Author
            Approval     = [int]$s.ApprovalState
            ApprovalText = $appr
            Version      = [string]$s.ScriptVersion
            Path         = [string]$s.__PATH
            Display      = "$($s.ScriptName)   [$appr]  v$($s.ScriptVersion)"
        }
    }
    return @($list)
}

function Get-CmScriptFull {
    # Script/ParamsDefinition "lazy" ozelliktir; tam nesne __PATH ile yeniden alinir.
    param([string]$Path)
    return [wmi]$Path
}

function Get-CmScriptText {
    param([string]$Path)
    $full = Get-CmScriptFull -Path $Path
    $b64 = [string]$full.Script
    if ([string]::IsNullOrWhiteSpace($b64)) { return '(Script icerigi okunamadi)' }
    $bytes = [Convert]::FromBase64String($b64)
    $nulCount = @($bytes | Where-Object { $_ -eq 0 }).Count
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { $text = [System.Text.Encoding]::Unicode.GetString($bytes) }
    elseif ($nulCount -gt ($bytes.Length / 4))                              { $text = [System.Text.Encoding]::Unicode.GetString($bytes) }
    else                                                                    { $text = [System.Text.Encoding]::UTF8.GetString($bytes) }
    return $text.TrimStart([char]0xFEFF)
}

function New-CmRunScriptParam {
    # InitiateClientOperationEx 'Param' degeri: Base64(UTF8(XML)). ScriptHash,
    # SMS_Scripts'ten alinir. Parametresiz calistirma.
    param([string]$Path)
    $full = Get-CmScriptFull -Path $Path
    $xml = "<ScriptContent ScriptGuid='{0}'><ScriptVersion>{1}</ScriptVersion><ScriptType>{2}</ScriptType><ScriptHash ScriptHashAlg='SHA256'>{3}</ScriptHash><ScriptParameters></ScriptParameters><ParameterGroupHash ParameterHashAlg='SHA256'></ParameterGroupHash></ScriptContent>" -f $full.ScriptGuid, $full.ScriptVersion, $full.ScriptType, $full.ScriptHash

    $hasParams = $false
    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$full.ParamsDefinition)) {
            $pxml = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String([string]$full.ParamsDefinition)).TrimStart([char]0xFEFF)
            $doc = New-Object System.Xml.XmlDocument
            $doc.LoadXml($pxml)
            if ($doc.DocumentElement -and $doc.DocumentElement.ChildNodes.Count -gt 0) { $hasParams = $true }
        }
    }
    catch { }

    return [pscustomobject]@{
        Param     = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($xml))
        HasParams = $hasParams
    }
}

# ---------- Diyaloglar / disa aktarma ----------

function Show-TextDialog {
    param([string]$Title, [string]$Text)
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.Size = New-Object System.Drawing.Size(800, 520)
    $dlg.StartPosition = 'CenterParent'

    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Multiline = $true
    $tb.ReadOnly = $true
    $tb.ScrollBars = 'Both'
    $tb.WordWrap = $false
    $tb.Font = New-Object System.Drawing.Font('Consolas', 9)
    $tb.Dock = 'Fill'
    $tb.Text = $Text

    $dlg.Controls.Add($tb)
    $dlg.Add_Shown({ $tb.SelectionStart = 0; $tb.SelectionLength = 0 }.GetNewClosure())
    [void]$dlg.ShowDialog()
}

function Export-ToExcelAndOpen {
    param([object[]]$Data, [string]$DefaultFileName = "Rapor_$(Get-Date -Format 'yyyyMMdd_HHmmss')")

    if (-not $Data -or @($Data).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Aktarilacak veri yok.', 'Bilgi') | Out-Null
        return
    }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'Excel Dosyasi (*.xlsx)|*.xlsx'
    $sfd.FileName = "$DefaultFileName.xlsx"
    if ($sfd.ShowDialog() -ne 'OK') { return }

    $excel = $null
    try { $excel = New-Object -ComObject Excel.Application }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Bu makinede Microsoft Excel bulunamadi. CSV secenegini kullanabilirsiniz.`n`nHata: $($_.Exception.Message)", 'Hata') | Out-Null
        return
    }

    try {
        $excel.Visible = $false
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)

        # Hucre hucre degil, tek seferde 2 boyutlu dizi olarak yazilir (cok satirda cok daha hizli)
        $props    = @($Data[0].PSObject.Properties.Name)
        $rowCount = @($Data).Count
        $colCount = $props.Count
        $arr = New-Object 'object[,]' ($rowCount + 1), $colCount
        for ($c = 0; $c -lt $colCount; $c++) { $arr[0, $c] = $props[$c] }
        $r = 1
        foreach ($item in $Data) {
            for ($c = 0; $c -lt $colCount; $c++) { $arr[$r, $c] = [string]$item.($props[$c]) }
            $r++
        }
        $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($rowCount + 1, $colCount)).Value2 = $arr

        $headerRange = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $colCount))
        $headerRange.Font.Bold = $true
        $null = $ws.Columns.AutoFit()
        for ($c = 1; $c -le $colCount; $c++) {
            $cw = $ws.Columns.Item($c)
            if ($cw.ColumnWidth -gt 60) { $cw.ColumnWidth = 60 }
        }
        try { $null = $headerRange.AutoFilter() } catch { }

        $wb.SaveAs($sfd.FileName)
        $excel.Visible = $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Excel'e aktarma sirasinda hata olustu: $($_.Exception.Message)", 'Hata') | Out-Null
        try { $excel.Quit() } catch { }
    }
}

function Export-GridToCsv {
    # ';' ayiracli UTF8 CSV (Turkce Excel'de dogrudan kolonlara ayrilir).
    param([object[]]$Data, [string]$DefaultFileName, [bool]$OpenAfter = $false, [string]$InitialDirectory = '')

    if (-not $Data -or @($Data).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Aktarilacak veri yok.', 'Bilgi') | Out-Null
        return
    }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'CSV Dosyasi (*.csv)|*.csv'
    $sfd.FileName = "$DefaultFileName.csv"
    if ($InitialDirectory) {
        if (-not (Test-Path $InitialDirectory)) { try { $null = New-Item -ItemType Directory -Path $InitialDirectory -Force } catch { } }
        if (Test-Path $InitialDirectory) { $sfd.InitialDirectory = $InitialDirectory }
    }
    if ($sfd.ShowDialog() -ne 'OK') { return }

    try {
        $Data | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("CSV yazilamadi: $($_.Exception.Message)", 'Hata') | Out-Null
        return
    }

    if ($OpenAfter) {
        try { Invoke-Item -Path $sfd.FileName }
        catch { [System.Windows.Forms.MessageBox]::Show("CSV olusturuldu ancak otomatik acilamadi:`n$($sfd.FileName)", 'Bilgi') | Out-Null }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Disa aktarildi: $($sfd.FileName)", 'Bilgi') | Out-Null
    }
}

function Get-GridExportData {
    # Gridde GORUNEN kolonlar (aktif filtre varsa sadece filtrelenmis satirlar).
    # -WithDetail: Panel 1'deki satir detay logunu 'Detay' kolonu olarak ekler.
    param($Grid, $Source, [bool]$WithDetail = $false, [string[]]$SkipColumns = @())

    $out = foreach ($drv in $Source) {
        $h = [ordered]@{}
        foreach ($col in $Grid.Columns) {
            if ($col.Visible -and ($SkipColumns -notcontains $col.Name)) { $h[$col.HeaderText] = [string]$drv[$col.DataPropertyName] }
        }
        if ($WithDetail) {
            $hn = [string]$drv['Hostname']
            $h['Detay'] = if ($script:Details.ContainsKey($hn)) { ($script:Details[$hn] -replace '\s*\r?\n\s*', ' | ').Trim(' ', '|') } else { '' }
        }
        [pscustomobject]$h
    }
    return @($out)
}

# ---------- Filtre yardimcilari ----------

function Format-LikeValue {
    # DataView.RowFilter LIKE degeri icin ozel karakterleri kacirir.
    param([string]$Text)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToCharArray()) {
        switch ($ch) {
            '[' { [void]$sb.Append('[[]') }
            ']' { [void]$sb.Append('[]]') }
            '%' { [void]$sb.Append('[%]') }
            '*' { [void]$sb.Append('[*]') }
            "'" { [void]$sb.Append("''") }
            default { [void]$sb.Append($ch) }
        }
    }
    return $sb.ToString()
}

function Get-ColumnMap {
    # Gorunen kolonlar: baslik metni -> veri kolon adi
    param($Grid)
    $m = [ordered]@{}
    foreach ($c in $Grid.Columns) { if ($c.Visible) { $m[$c.HeaderText] = $c.DataPropertyName } }
    return $m
}

function Set-GridFilter {
    # Secili kolonda (veya tum kolonlarda) 'icerir' filtresi uygular.
    param($Source, $ColMap, [string]$ColumnHeader, [string]$Text)

    $q = "$Text".Trim()
    if ([string]::IsNullOrWhiteSpace($q)) { $Source.RemoveFilter(); return }

    $esc = Format-LikeValue -Text $q
    if ($ColumnHeader -and $ColMap.Contains($ColumnHeader)) {
        $Source.Filter = "[$($ColMap[$ColumnHeader])] LIKE '%$esc%'"
    }
    else {
        $Source.Filter = ((@($ColMap.Values) | ForEach-Object { "[$_] LIKE '%$esc%'" }) -join ' OR ')
    }
}

# ---------- Update Parser (Panel 2) - Panel 1'den BAGIMSIZ ----------

function Add-ParserLines {
    # Script ciktisini / yapistirilan metni '@' (makine-kayit ayirici) ve satir sonlarindan boler, her satiri
    # ';' ile kolonlara ayirip Panel 2 tablosuna ALT ALTA ekler (eski client_update_parser.ps1 mantigi).
    # Ayni SourceHost icin onceki satirlar once silinir (tekrar calistirmada cift kayit olmaz).
    # Doner: Added (eklenen satir sayisi), Hosts (satirlardaki gecerli ve tekil hostlar - kontrol icin).
    param([string]$SourceHost, [string]$Text)

    $expected = $script:ParserColumns.Count
    $legacyN  = $script:ParserLegacyColumns.Count
    $colCount = $script:ParseTable.Columns.Count

    # Ayni SourceHost'un onceki satirlari (or. "Bekleniyor.." satiri) ayni konuma yeniden yazilir (sira bozulmasin)
    $pos  = -1
    $olds = @($script:ParseTable.Select("SourceHost = '" + $SourceHost.Replace("'", "''") + "'"))
    if ($olds.Count -gt 0) { $pos = $script:ParseTable.Rows.IndexOf($olds[0]) }
    foreach ($old in $olds) { $script:ParseTable.Rows.Remove($old) }
    if ($pos -gt $script:ParseTable.Rows.Count) { $pos = $script:ParseTable.Rows.Count }

    $hostList = New-Object 'System.Collections.Generic.List[string]'
    $seen     = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $added = 0

    $lines = @("$Text" -split '@|\r\n|\n|\r' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    foreach ($line in $lines) {
        if ($line -match '^Hostname\s*;') { continue }     # script kendi basligini basmissa atla

        $parts = @($line -split ';')
        $row = $script:ParseTable.NewRow()
        for ($i = 0; $i -lt $colCount; $i++) { $row[$i] = '' }

        if ($parts.Count -lt 2) {
            # ';' yok: parse edilemeyen satir (hata mesaji vb.) - kaybolmasin
            $row['Hostname']   = $SourceHost
            $row['UpdateName'] = $line
            $row['KolonUyumu'] = 'Ayristirilamadi'
        }
        elseif ($parts.Count -eq $legacyN -and $legacyN -ne $expected) {
            # Eski 19 kolonluk cikti: degerler kolon ADIYLA yerlestirilir (yeni kolonlar bos kalir)
            for ($i = 0; $i -lt $legacyN; $i++) {
                $ci = [array]::IndexOf($script:ParserColumns, $script:ParserLegacyColumns[$i])
                if ($ci -ge 0) { $row[$ci] = $parts[$i].Trim() }
            }
            $row['KolonUyumu'] = "OK ($legacyN kolon)"
        }
        else {
            for ($i = 0; $i -lt $expected; $i++) {
                $row[$i] = $(if ($i -lt $parts.Count) { $parts[$i].Trim() } else { '' })
            }
            if ($parts.Count -gt $expected) {
                # fazla kolonlar son kolona eklenir (veri kaybolmasin)
                $row[$expected - 1] = (($parts[($expected - 1)..($parts.Count - 1)]) -join ';')
            }
            $row['KolonUyumu'] = $(if ($parts.Count -eq $expected) { 'OK' } else { "Beklenen $expected, gelen $($parts.Count)" })
        }

        $row['SourceHost'] = $SourceHost
        if ($pos -ge 0) { $script:ParseTable.Rows.InsertAt($row, $pos); $pos++ }
        else { [void]$script:ParseTable.Rows.Add($row) }
        $added++

        $hn = [string]$row['Hostname']
        if ($hn -match $script:HostRegex -and $seen.Add($hn)) { $hostList.Add($hn) }
    }

    return [pscustomobject]@{ Added = $added; Hosts = @($hostList); Pos = $pos }
}

function Add-ParserInfoRow {
    # Cikti gelmeyen / islem hatasi olan cihaz icin Panel 2'ye bilgi satiri ekler (sebep UpdateName kolonunda).
    param([string]$SourceHost, [string]$Hostname, [string]$Message, [string]$Tag = 'Cikti yok', [int]$At = -1)

    $row = $script:ParseTable.NewRow()
    for ($i = 0; $i -lt $script:ParseTable.Columns.Count; $i++) { $row[$i] = '' }
    $row['Hostname']   = $Hostname
    $row['UpdateName'] = $Message
    $row['KolonUyumu'] = $Tag
    $row['SourceHost'] = $SourceHost
    if ($At -ge 0 -and $At -le $script:ParseTable.Rows.Count) { $script:ParseTable.Rows.InsertAt($row, $At) }
    else { [void]$script:ParseTable.Rows.Add($row) }
}

function Add-ParserPendingRow {
    # Panel 2: script gonderilen cihazi hemen tabloya ekler; sonuc gelene kadar "Bekleniyor.." yazar.
    # Cihazin onceki satirlari varsa ayni konumda yer tutucu ile degistirilir; sonuc gelince Add-ParserLines
    # (ayni SourceHost) bu satiri ayni konumda gercek verilerle degistirir.
    param([string]$HostName)

    $esc = $HostName.Replace("'", "''")
    $olds = @($script:ParseTable.Select("SourceHost = '$esc'"))
    $pos = -1
    if ($olds.Count -gt 0) { $pos = $script:ParseTable.Rows.IndexOf($olds[0]) }
    foreach ($o in $olds) { $script:ParseTable.Rows.Remove($o) }
    if ($pos -gt $script:ParseTable.Rows.Count) { $pos = $script:ParseTable.Rows.Count }
    $script:ParserNoOutput.Remove($HostName)

    $row = $script:ParseTable.NewRow()
    for ($i = 0; $i -lt $script:ParseTable.Columns.Count; $i++) { $row[$i] = '' }
    $row['Hostname']   = $HostName
    $row['UpdateName'] = $script:PendingText
    foreach ($c in $script:ParserCheckColumns) { $row[$c] = $script:PendingText }
    $row['KolonUyumu'] = $script:PendingText
    $row['SourceHost'] = $HostName
    if ($pos -ge 0) { $script:ParseTable.Rows.InsertAt($row, $pos) }
    else { [void]$script:ParseTable.Rows.Add($row) }
}

function Add-ParserDetail {
    param([string]$Hostname, [string]$Text)
    if (-not $script:ParserDetails.ContainsKey($Hostname)) { $script:ParserDetails[$Hostname] = '' }
    $script:ParserDetails[$Hostname] += "[$(Get-Date -Format 'HH:mm:ss')] $Text`r`n`r`n"
}

function Set-ParserCheckValues {
    # Bir hostun TUM satirlarina (alt alta yazilan kayitlarina) kontrol kolonlarini yazar.
    param([string]$HostName, [hashtable]$Values)

    $esc = $HostName.Replace("'", "''")
    foreach ($row in @($script:ParseTable.Select("[Hostname] = '$esc' OR [SourceHost] = '$esc'"))) {
        foreach ($k in $Values.Keys) { $row[$k] = [string]$Values[$k] }
    }
}

function ConvertTo-ParserCheckValues {
    param($Check)
    return @{
        DNS_Durumu     = [string]$Check.DnsStatus
        Online_mi      = $(if ($Check.IsOnline) { 'Evet' } else { 'Hayir' })
        AD_Durumu      = [string]$Check.AdStatus
        AD_LastLogon   = [string]$Check.LastLogon
        CcmExec_Durumu = [string]$Check.CcmExec
        SCCM_Kayit     = [string]$Check.SccmKayit
    }
}

function Update-ParserInfo {
    $rows = $script:ParseTable.Rows.Count
    $hostSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $bad = 0
    foreach ($row in $script:ParseTable.Rows) {
        [void]$hostSet.Add([string]$row['Hostname'])
        if ([string]$row['KolonUyumu'] -notlike 'OK*') { $bad++ }
    }
    $txt = "Update Parser (Panel 2 - Panel 1'den bagimsiz)  |  $rows satir, $($hostSet.Count) cihaz, $bad kolon uyumsuz / bilgi satiri"
    if ($script:ParserNoOutput.Count -gt 0) { $txt += "  |  $($script:ParserNoOutput.Count) cihazdan cikti gelmedi (sebep satirda yazili)" }
    $lblParserInfo.Text = $txt
}

# ---- Panel 2 sonuc isleyicileri (UI thread'inde calisir; SADECE Panel 2 tablosuna yazar) ----

function Update-ParserFromCheck {
    # "Parse sonrasi kontrol" / "Kontrolleri Calistir" sonucu: DNS + Online + AD + CcmExec + SCCM kaydi
    param($Key, $Output, $ErrorText)

    $c = @($Output | Where-Object { $_ -and $_.PSObject.Properties['SccmKayit'] })[0]
    if (-not $c) {
        $msg = "Islem hatasi: $ErrorText"
        Set-ParserCheckValues -HostName $Key -Values @{ DNS_Durumu = $msg; Online_mi = ''; AD_Durumu = ''; AD_LastLogon = ''; CcmExec_Durumu = ''; SCCM_Kayit = '' }
        Add-ParserDetail -Hostname $Key -Text "Kontrol: $msg"
        return
    }

    Set-ParserCheckValues -HostName $Key -Values (ConvertTo-ParserCheckValues -Check $c)
    $d = "Kontrol: DNS=$($c.DnsStatus); Online=$($c.IsOnline); AD=$($c.AdStatus); LastLogon=$($c.LastLogon); CcmExec=$($c.CcmExec); SCCM kaydi=$($c.SccmKayit)"
    if ($c.Note) { $d += "`r`nNot: $($c.Note)" }
    Add-ParserDetail -Hostname $Key -Text $d
}

function Update-ParserFromScript {
    # Panel 2'den gonderilen script sonucu: cikti '@' / ';' ile parse edilip alt alta yazilir,
    # ardindan DNS/AD/SCCM kontrol kolonlari doldurulur. Panel 1'e DOKUNMAZ.
    param($Key, $Output, $ErrorText)

    $r = @($Output | Where-Object { $_ -and $_.PSObject.Properties['Output'] -and $_.PSObject.Properties['Check'] })[0]
    if (-not $r) {
        $msg = "Islem hatasi: $ErrorText"
        $r0 = Add-ParserLines -SourceHost $Key -Text ''
        Add-ParserInfoRow -SourceHost $Key -Hostname $Key -Message $msg -At $r0.Pos
        $script:ParserNoOutput[$Key] = $true
        Add-ParserDetail -Hostname $Key -Text "Script (Parser): $msg"
        Write-ActionLog -Hostname $Key -Action 'Script (Parser)' -State 'fail' -Text $msg
        return
    }

    $hn  = [string]$r.Hostname
    $chk = $r.Check
    $res = Add-ParserLines -SourceHost $hn -Text ([string]$r.Output)
    if ($res.Added -eq 0) {
        $one = ("$($r.Result)" -replace '\s+', ' ').Trim()
        Add-ParserInfoRow -SourceHost $hn -Hostname $hn -Message $one -At $res.Pos
        $script:ParserNoOutput[$hn] = $true
    }
    else {
        $script:ParserNoOutput.Remove($hn)
    }

    Set-ParserCheckValues -HostName $hn -Values (ConvertTo-ParserCheckValues -Check $chk)

    $d = "$($r.Action): $($r.Result)`r`nDNS=$($chk.DnsStatus); Online=$($chk.IsOnline); AD=$($chk.AdStatus); LastLogon=$($chk.LastLogon); CcmExec=$($chk.CcmExec); SCCM kaydi=$($chk.SccmKayit)"
    if ($chk.Note) { $d += "`r`nNot: $($chk.Note)" }
    if ($r.Detail) { $d += "`r`n$($r.Detail)" }
    Add-ParserDetail -Hostname $hn -Text $d
    Write-ActionLog -Hostname $hn -Action $r.Action -State $r.State -Text $r.Result
}

function Start-ParserChecks {
    # Verilen hostlar icin DNS / AD / SCCM ajan kontrolunu (Panel 2 isi olarak) baslatir.
    param([string[]]$Hosts)

    $allDc = [bool]$chkAllDc.Checked
    foreach ($h in $Hosts) {
        $pv = @{}
        foreach ($c in $script:ParserCheckColumns) { $pv[$c] = $script:PendingText }
        Set-ParserCheckValues -HostName $h -Values $pv
    }

    $items = @(foreach ($h in $Hosts) {
        [pscustomobject]@{ Key = $h; Params = @{ ComputerName = $h; SmsProvider = [string]$script:SmsServer; SiteCode = [string]$script:SiteCode; QueryAllDCs = $allDc } }
    })
    Start-ParallelRun -Title 'Parser kontrolu (DNS / AD / SCCM ajan)' -Worker $script:ParserCheckWorker -Items $items -OnResultFn 'Update-ParserFromCheck' -Panel 2
}

# ---------- Hedef secimi ----------

function Get-ComputerList {
    $raw = @($txtComputers.Text -split '[\r\n,;\s]+' | Where-Object { $_ })
    $valid = @()
    $bad = @()
    foreach ($n in $raw) {
        if ($n -match '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
            if ($valid -notcontains $n) { $valid += $n }
        }
        else { $bad += $n }
    }
    if ($bad.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("Gecersiz bilgisayar isimleri yok sayildi:`n$($bad -join ', ')", 'Uyari') | Out-Null
    }
    if ($valid.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('En az bir gecerli bilgisayar ismi girin.', 'Uyari') | Out-Null
        return @()
    }
    return $valid
}

function Get-TargetHosts {
    # Verilen gridde secili satirlar (Hostname kolonu) varsa onlar, yoksa soldaki listedeki tum makineler.
    param($Grid = $dgv)

    $sel = @()
    foreach ($row in $Grid.SelectedRows) {
        $h = [string]$row.Cells['Hostname'].Value
        if ($h -and ($h -match $script:HostRegex) -and ($sel -notcontains $h)) { $sel += $h }
    }
    if ($sel.Count -gt 0) {
        return [pscustomobject]@{ Hosts = $sel; Source = 'gridde secili satirlar' }
    }
    $list = @(Get-ComputerList)
    if ($list.Count -eq 0) { return $null }
    return [pscustomobject]@{ Hosts = $list; Source = 'soldaki listedeki tum makineler' }
}
function Format-HostPreview {
    param([string[]]$Hosts)
    $max = 12
    $txt = (@($Hosts | Select-Object -First $max) -join ', ')
    if (@($Hosts).Count -gt $max) { $txt += " ... (+$(@($Hosts).Count - $max))" }
    return $txt
}

function Start-ScriptSend {
    # Panel 1 ve Panel 2 ayni SCCM baglantisini / script secimini kullanir; ancak sonuc SADECE istegin
    # geldigi panele yazilir (Panel 1 -> kontrol tablosu, Panel 2 -> Update Parser tablosu).
    param([int]$Panel)

    $idx = $cboScripts.SelectedIndex
    if ($idx -lt 0 -or -not $script:SiteCode) { return }
    $s = $script:CmScripts[$idx]

    if ($s.Approval -ne 3) {
        [System.Windows.Forms.MessageBox]::Show('Bu script SCCM''de ONAYLANMAMIS; calistirilamaz.', 'Uyari') | Out-Null
        return
    }

    $t = if ($Panel -eq 2) { Get-TargetHosts -Grid $dgvParse } else { Get-TargetHosts -Grid $dgv }
    if (-not $t) { return }

    $timeout = [int]$nudTimeout.Value

    $services = @($txtServices.Text -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($Panel -eq 1 -and $services.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Servis kontrolu icin 1) Kontrol bolumune en az bir servis ismi girin.', 'Uyari') | Out-Null
        return
    }

    $param = $null
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try { $param = New-CmRunScriptParam -Path $s.Path }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Script bilgisi alinamadi:`n$($_.Exception.Message)", 'Hata') | Out-Null
        return
    }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }

    if ($Panel -eq 2) {
        $extra = "DNS / AD / SCCM ajan kontrolu de calistirilacak.`nSonuclar SADECE Panel 2'ye (Update Parser) yazilacak; Panel 1 etkilenmez."
    }
    else {
        $extra = "AD yapisi kontrolu (DNS / ping / AD) ve servis kontrolu ($($services -join ', ')) de calistirilacak.`nSonuclar SADECE Panel 1'e yazilacak; Update Parser (Panel 2) etkilenmez."
    }

    $msg = "Script: $($s.Name)  (v$($s.Version))`nYazar: $($s.Author)`n`nHedef ($($t.Source)): $(@($t.Hosts).Count) makine`n$(Format-HostPreview -Hosts $t.Hosts)`n`nCevap bekleme suresi: $timeout sn`n$extra`n`nScript bu makinelerde SYSTEM hesabiyla calisacak. Gonderilsin mi?"
    if ($param.HasParams) {
        $msg += "`n`nUYARI: Bu script parametre tanimliyor; parametreler BOS gonderilecek."
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show($msg, "Onay - Script Gonder (Panel $Panel)",
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $b64   = $param.Param
    $allDc = [bool]$chkAllDc.Checked
    $items = @(foreach ($h in $t.Hosts) {
        $p = @{
            ComputerName = $h
            SmsProvider  = $script:SmsServer
            SiteCode     = $script:SiteCode
            ScriptGuid   = $s.Guid
            ScriptName   = $s.Name
            ParamB64     = $b64
            TimeoutSec   = $timeout
            QueryAllDCs  = $allDc
        }
        if ($Panel -eq 1) { $p['ServiceNames'] = $services }   # Panel 2 worker'i bu parametreyi almaz
        [pscustomobject]@{ Key = $h; Params = $p }
    })
    Save-Settings

    if ($Panel -eq 2) {
        foreach ($h in $t.Hosts) { Add-ParserPendingRow -HostName $h }
        Update-ParserInfo
        Start-ParallelRun -Title 'Script + DNS/AD/SCCM (Parser)' -Worker $script:ParserScriptWorker -Items $items -OnResultFn 'Update-ParserFromScript' -Panel 2
    }
    else {
        Set-GridPending -Hosts $t.Hosts -ClearError -Columns @('DnsStatus', 'IsOnline', 'AdStatus', 'LastLogon', 'Running', 'NotRunning', 'CcmExec', 'ActionResult')
        Start-ParallelRun -Title 'Script + AD + servis kontrolu' -Worker $script:ScriptWorker -Items $items -OnResultFn 'Update-GridFromScript' -Panel 1
    }
}
# ==================== ANA FORM ====================

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Istemci Durum ve SCCM Islem Araci'
$form.Size = New-Object System.Drawing.Size(1360, 900)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(1150, 700)

$lblComputers = New-Object System.Windows.Forms.Label
$lblComputers.Text = 'Bilgisayar Isimleri (her satira bir tane):'
$lblComputers.Location = New-Object System.Drawing.Point(10, 10)
$lblComputers.AutoSize = $true

$txtComputers = New-Object System.Windows.Forms.TextBox
$txtComputers.Multiline = $true
$txtComputers.ScrollBars = 'Vertical'
$txtComputers.Location = New-Object System.Drawing.Point(10, 30)
$txtComputers.Size = New-Object System.Drawing.Size(200, 800)
$txtComputers.Anchor = 'Top,Bottom,Left'

# ---------- 1) Kontrol ----------
$gbCheck = New-Object System.Windows.Forms.GroupBox
$gbCheck.Text = '1) Kontrol (DNS / Online / AD / Servis)'
$gbCheck.Location = New-Object System.Drawing.Point(220, 8)
$gbCheck.Size = New-Object System.Drawing.Size(520, 184)

$lblServices = New-Object System.Windows.Forms.Label
$lblServices.Text = 'Servisler (virgulle ayirin):'
$lblServices.Location = New-Object System.Drawing.Point(10, 22)
$lblServices.AutoSize = $true

$txtServices = New-Object System.Windows.Forms.TextBox
$txtServices.Location = New-Object System.Drawing.Point(10, 42)
$txtServices.Size = New-Object System.Drawing.Size(495, 23)
$txtServices.Text = 'CcmExec, BITS, wuauserv'

$chkAllDc = New-Object System.Windows.Forms.CheckBox
$chkAllDc.Text = "AD: tum DC'lerden gercek lastLogon sorgula (yavas)"
$chkAllDc.Location = New-Object System.Drawing.Point(10, 72)
$chkAllDc.AutoSize = $true

$btnCheck = New-Object System.Windows.Forms.Button
$btnCheck.Text = 'Kontrol Et'
$btnCheck.Location = New-Object System.Drawing.Point(10, 108)
$btnCheck.Size = New-Object System.Drawing.Size(110, 32)

$btnAd = New-Object System.Windows.Forms.Button
$btnAd.Text = 'AD Yapisi Kontrolu'
$btnAd.Location = New-Object System.Drawing.Point(130, 108)
$btnAd.Size = New-Object System.Drawing.Size(150, 32)

$btnFix = New-Object System.Windows.Forms.Button
$btnFix.Text = 'Servisleri Baslat'
$btnFix.Location = New-Object System.Drawing.Point(290, 108)
$btnFix.Size = New-Object System.Drawing.Size(150, 32)
$btnFix.Enabled = $false

$lblAdHint = New-Object System.Windows.Forms.Label
$lblAdHint.Text = 'AD Yapisi Kontrolu: DNS + ping + AD (var mi / Enabled-Disabled / Last Logon). Script gonderirken de otomatik calisir.'
$lblAdHint.Location = New-Object System.Drawing.Point(10, 148)
$lblAdHint.Size = New-Object System.Drawing.Size(500, 30)
$lblAdHint.ForeColor = [System.Drawing.Color]::DimGray

$gbCheck.Controls.AddRange(@($lblServices, $txtServices, $chkAllDc, $btnCheck, $btnAd, $btnFix, $lblAdHint))

# ---------- 2) SCCM Script ----------
$gbSccm = New-Object System.Windows.Forms.GroupBox
$gbSccm.Text = '2) SCCM Script'
$gbSccm.Location = New-Object System.Drawing.Point(750, 8)
$gbSccm.Size = New-Object System.Drawing.Size(580, 184)

$lblSms = New-Object System.Windows.Forms.Label
$lblSms.Text = 'SMS Provider (Site Server) adi:'
$lblSms.Location = New-Object System.Drawing.Point(10, 22)
$lblSms.AutoSize = $true

$txtSms = New-Object System.Windows.Forms.TextBox
$txtSms.Location = New-Object System.Drawing.Point(10, 42)
$txtSms.Size = New-Object System.Drawing.Size(300, 23)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = 'Baglan / Scriptleri Listele'
$btnConnect.Location = New-Object System.Drawing.Point(320, 40)
$btnConnect.Size = New-Object System.Drawing.Size(170, 27)

$lblSite = New-Object System.Windows.Forms.Label
$lblSite.Text = 'Site: -'
$lblSite.Location = New-Object System.Drawing.Point(500, 46)
$lblSite.AutoSize = $true

$lblScript = New-Object System.Windows.Forms.Label
$lblScript.Text = 'Script:'
$lblScript.Location = New-Object System.Drawing.Point(10, 74)
$lblScript.AutoSize = $true

$cboScripts = New-Object System.Windows.Forms.ComboBox
$cboScripts.Location = New-Object System.Drawing.Point(10, 94)
$cboScripts.Size = New-Object System.Drawing.Size(410, 23)
$cboScripts.DropDownStyle = 'DropDownList'

$btnShowScript = New-Object System.Windows.Forms.Button
$btnShowScript.Text = 'Icerik'
$btnShowScript.Location = New-Object System.Drawing.Point(428, 92)
$btnShowScript.Size = New-Object System.Drawing.Size(60, 27)
$btnShowScript.Enabled = $false

$lblTimeout = New-Object System.Windows.Forms.Label
$lblTimeout.Text = 'Bekleme (sn):'
$lblTimeout.Location = New-Object System.Drawing.Point(498, 74)
$lblTimeout.AutoSize = $true

$nudTimeout = New-Object System.Windows.Forms.NumericUpDown
$nudTimeout.Location = New-Object System.Drawing.Point(498, 94)
$nudTimeout.Size = New-Object System.Drawing.Size(65, 23)
$nudTimeout.Minimum = 30
$nudTimeout.Maximum = 900
$nudTimeout.Value = 90

$btnSendScript = New-Object System.Windows.Forms.Button
$btnSendScript.Text = 'Secili Scripti Gonder (Panel 1)'
$btnSendScript.Location = New-Object System.Drawing.Point(10, 124)
$btnSendScript.Size = New-Object System.Drawing.Size(230, 28)
$btnSendScript.Enabled = $false

$lblScriptInfo = New-Object System.Windows.Forms.Label
$lblScriptInfo.Text = ''
$lblScriptInfo.Location = New-Object System.Drawing.Point(250, 131)
$lblScriptInfo.AutoSize = $true
$lblScriptInfo.ForeColor = [System.Drawing.Color]::DimGray

$lblIndep = New-Object System.Windows.Forms.Label
$lblIndep.Text = "Panel 1 ve Panel 2 bagimsizdir: bu buton sadece Panel 1'e yazar. Panel 2 icin Panel 2'deki 'Script Gonder (Panel 2)' butonunu kullanin."
$lblIndep.Location = New-Object System.Drawing.Point(10, 154)
$lblIndep.Size = New-Object System.Drawing.Size(560, 28)
$lblIndep.ForeColor = [System.Drawing.Color]::DimGray

$gbSccm.Controls.AddRange(@($lblSms, $txtSms, $btnConnect, $lblSite, $lblScript, $cboScripts, $btnShowScript, $lblTimeout, $nudTimeout, $btnSendScript, $lblScriptInfo, $lblIndep))

# ---------- 3) SCCM Istemci Eylemi ----------
$gbActions = New-Object System.Windows.Forms.GroupBox
$gbActions.Text = '3) SCCM Istemci Eylemi'
$gbActions.Location = New-Object System.Drawing.Point(220, 200)
$gbActions.Size = New-Object System.Drawing.Size(1110, 60)

$cboAction = New-Object System.Windows.Forms.ComboBox
$cboAction.Location = New-Object System.Drawing.Point(10, 25)
$cboAction.Size = New-Object System.Drawing.Size(400, 23)
$cboAction.DropDownStyle = 'DropDownList'
foreach ($k in $script:ClientActions.Keys) { [void]$cboAction.Items.Add($k) }
$cboAction.SelectedIndex = 0

$btnAction = New-Object System.Windows.Forms.Button
$btnAction.Text = 'Eylemi Gonder'
$btnAction.Location = New-Object System.Drawing.Point(420, 22)
$btnAction.Size = New-Object System.Drawing.Size(170, 28)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = 'Panel 1 hedefi: gridde secili satirlar; yoksa soldaki liste (Script / AD icin de gecerli).'
$lblHint.Location = New-Object System.Drawing.Point(605, 29)
$lblHint.AutoSize = $true
$lblHint.ForeColor = [System.Drawing.Color]::DimGray

$gbActions.Controls.AddRange(@($cboAction, $btnAction, $lblHint))

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = ''
$lblStatus.Location = New-Object System.Drawing.Point(220, 270)
$lblStatus.AutoSize = $true
$lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue

# ---------- Sonuc tablosu ----------
$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Dock = 'Fill'
$dgv.ReadOnly = $true
$dgv.AllowUserToAddRows = $false
$dgv.AutoSizeColumnsMode = 'Fill'
$dgv.RowHeadersVisible = $false
$dgv.SelectionMode = 'FullRowSelect'
$dgv.MultiSelect = $true
$dgv.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'
$dgv.AutoGenerateColumns = $false

function New-GridColumn {
    param([string]$Name, [string]$Header, [float]$Weight, [bool]$Visible = $true)
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $Name
    $col.DataPropertyName = $Name
    $col.HeaderText = $Header
    $col.FillWeight = $Weight
    $col.Visible = $Visible
    return $col
}

$dgv.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@(
    (New-GridColumn 'Hostname'     'Hostname'              12),
    (New-GridColumn 'DnsStatus'    'DNS'                   13),
    (New-GridColumn 'IsOnline'     'Online mi'              6),
    (New-GridColumn 'AdStatus'     'AD Durumu'              8),
    (New-GridColumn 'LastLogon'    'AD Last Logon'         11),
    (New-GridColumn 'Running'      'Calisan Servisler'     10),
    (New-GridColumn 'NotRunning'   'Calismayan Servisler'  11),
    (New-GridColumn 'CcmExec'      'CcmExec'                7),
    (New-GridColumn 'LastAction'   'Son Islem'             13),
    (New-GridColumn 'ActionResult' 'Islem Sonucu / Script Ciktisi'          30),
    (New-GridColumn 'ErrorMessage' 'Hata / Not'            14),
    (New-GridColumn 'State'        'State'                  5 $false)
))
$dgv.DataSource = $script:BindingSrc

$dgv.add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $rowView = $dgv.Rows[$e.RowIndex].DataBoundItem
    if (-not $rowView) { return }
    switch ("$($rowView['State'])") {
        'fail' { $e.CellStyle.BackColor = [System.Drawing.Color]::LightCoral }
        'warn' { $e.CellStyle.BackColor = [System.Drawing.Color]::LightYellow }
        'ok'   { $e.CellStyle.BackColor = [System.Drawing.Color]::LightGreen }
        'pending' { $e.CellStyle.BackColor = [System.Drawing.Color]::Gainsboro }
    }
})

$dgv.add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $h = [string]$dgv.Rows[$e.RowIndex].Cells['Hostname'].Value
    $txt = if ($script:Details.ContainsKey($h)) { $script:Details[$h] } else { '(Detay yok)' }
    Show-TextDialog -Title "Detay: $h" -Text $txt
})

# ==================== SEKMELER: PANEL 1 (Kontrol / Script Ciktilari)  +  PANEL 2 (Update Parser) ====================

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(220, 294)
$tabs.Size = New-Object System.Drawing.Size(1110, 541)
$tabs.Anchor = 'Top,Bottom,Left,Right'

$tabOut = New-Object System.Windows.Forms.TabPage
$tabOut.Text = 'Panel 1 - Script Ciktilari / Kontroller'
$tabParse = New-Object System.Windows.Forms.TabPage
$tabParse.Text = 'Panel 2 - Update Parser'

# ---------- Panel 1 arac cubugu: filtre + Excel + CSV ----------
$pnlTop1 = New-Object System.Windows.Forms.Panel
$pnlTop1.Dock = 'Top'
$pnlTop1.Height = 36

$lblF1 = New-Object System.Windows.Forms.Label
$lblF1.Text = 'Filtre:'
$lblF1.Location = New-Object System.Drawing.Point(6, 11)
$lblF1.AutoSize = $true

$cboF1 = New-Object System.Windows.Forms.ComboBox
$cboF1.DropDownStyle = 'DropDownList'
$cboF1.Location = New-Object System.Drawing.Point(50, 7)
$cboF1.Size = New-Object System.Drawing.Size(150, 23)

$txtF1 = New-Object System.Windows.Forms.TextBox
$txtF1.Location = New-Object System.Drawing.Point(206, 7)
$txtF1.Size = New-Object System.Drawing.Size(200, 23)

$btnClearF1 = New-Object System.Windows.Forms.Button
$btnClearF1.Text = 'Temizle'
$btnClearF1.Location = New-Object System.Drawing.Point(412, 5)
$btnClearF1.Size = New-Object System.Drawing.Size(65, 27)

$btnExcel1 = New-Object System.Windows.Forms.Button
$btnExcel1.Text = "Excel'e Aktar"
$btnExcel1.Location = New-Object System.Drawing.Point(490, 5)
$btnExcel1.Size = New-Object System.Drawing.Size(110, 27)

$btnCsv1 = New-Object System.Windows.Forms.Button
$btnCsv1.Text = 'CSV'
$btnCsv1.Location = New-Object System.Drawing.Point(606, 5)
$btnCsv1.Size = New-Object System.Drawing.Size(70, 27)

$lblCount1 = New-Object System.Windows.Forms.Label
$lblCount1.Text = ''
$lblCount1.Location = New-Object System.Drawing.Point(822, 11)
$lblCount1.AutoSize = $true
$lblCount1.ForeColor = [System.Drawing.Color]::DimGray

$btnUnfocus1 = New-Object System.Windows.Forms.Button
$btnUnfocus1.Text = 'Focus Kaybet'
$btnUnfocus1.Location = New-Object System.Drawing.Point(688, 5)
$btnUnfocus1.Size = New-Object System.Drawing.Size(124, 27)

$pnlTop1.Controls.AddRange(@($lblF1, $cboF1, $txtF1, $btnClearF1, $btnExcel1, $btnCsv1, $btnUnfocus1, $lblCount1))

# ---------- Panel 2: Update Parser (Panel 1'den BAGIMSIZ) ----------
$lblParserInfo = New-Object System.Windows.Forms.Label
$lblParserInfo.Dock = 'Top'
$lblParserInfo.Height = 24
$lblParserInfo.TextAlign = 'MiddleLeft'
$lblParserInfo.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$lblParserInfo.ForeColor = [System.Drawing.Color]::DarkGreen
$lblParserInfo.Text = "Update Parser (Panel 2 - Panel 1'den bagimsiz)  |  0 satir, 0 cihaz, 0 kolon uyumsuz / bilgi satiri"

$lblParserStatus = New-Object System.Windows.Forms.Label
$lblParserStatus.Dock = 'Top'
$lblParserStatus.Height = 20
$lblParserStatus.TextAlign = 'MiddleLeft'
$lblParserStatus.ForeColor = [System.Drawing.Color]::DarkBlue
$lblParserStatus.Text = ''

# -- Yapistirma alani (eski client_update_parser.ps1 mantigi)
$pnlPaste = New-Object System.Windows.Forms.Panel
$pnlPaste.Dock = 'Top'
$pnlPaste.Height = 100

$lblPaste = New-Object System.Windows.Forms.Label
$lblPaste.Text = "SCCM Run Script ciktisini yapistirin ('@' ile ayrilan her makine kaydi alt alta yazilir; her satir ';' ile ayrilmis):"
$lblPaste.Dock = 'Top'
$lblPaste.Height = 20
$lblPaste.TextAlign = 'MiddleLeft'

$txtParsePaste = New-Object System.Windows.Forms.TextBox
$txtParsePaste.Multiline = $true
$txtParsePaste.AcceptsReturn = $true
$txtParsePaste.ScrollBars = 'Both'
$txtParsePaste.WordWrap = $false
$txtParsePaste.MaxLength = [int]::MaxValue
$txtParsePaste.Font = New-Object System.Drawing.Font('Consolas', 9)
$txtParsePaste.Dock = 'Fill'

# Dock sirasi: Fill kontrol ILK, Top kontrol sonra eklenir
$pnlPaste.Controls.AddRange(@($txtParsePaste, $lblPaste))

# -- Calistirma cubugu
$pnlRun2 = New-Object System.Windows.Forms.Panel
$pnlRun2.Dock = 'Top'
$pnlRun2.Height = 36

$btnParseRun = New-Object System.Windows.Forms.Button
$btnParseRun.Text = 'Parse Et'
$btnParseRun.Location = New-Object System.Drawing.Point(6, 5)
$btnParseRun.Size = New-Object System.Drawing.Size(100, 27)

$chkParseCheck = New-Object System.Windows.Forms.CheckBox
$chkParseCheck.Text = 'DNS / AD / SCCM ajan kontrolu ekle'
$chkParseCheck.Location = New-Object System.Drawing.Point(114, 9)
$chkParseCheck.AutoSize = $true
$chkParseCheck.Checked = $true

$chkParseAppend = New-Object System.Windows.Forms.CheckBox
$chkParseAppend.Text = 'Tabloya ekle'
$chkParseAppend.Location = New-Object System.Drawing.Point(376, 9)
$chkParseAppend.AutoSize = $true

$btnParseScript = New-Object System.Windows.Forms.Button
$btnParseScript.Text = 'Script Gonder (Panel 2)'
$btnParseScript.Location = New-Object System.Drawing.Point(486, 5)
$btnParseScript.Size = New-Object System.Drawing.Size(180, 27)
$btnParseScript.Enabled = $false

$btnParseRecheck = New-Object System.Windows.Forms.Button
$btnParseRecheck.Text = 'Kontrolleri Calistir'
$btnParseRecheck.Location = New-Object System.Drawing.Point(674, 5)
$btnParseRecheck.Size = New-Object System.Drawing.Size(150, 27)

$btnParseClear = New-Object System.Windows.Forms.Button
$btnParseClear.Text = 'Parser Temizle'
$btnParseClear.Location = New-Object System.Drawing.Point(832, 5)
$btnParseClear.Size = New-Object System.Drawing.Size(105, 27)

$pnlRun2.Controls.AddRange(@($btnParseRun, $chkParseCheck, $chkParseAppend, $btnParseScript, $btnParseRecheck, $btnParseClear))

$tip = New-Object System.Windows.Forms.ToolTip
$tip.SetToolTip($chkParseAppend, 'Isaretli degilse "Parse Et" mevcut tabloyu temizleyip yeni veriyi yazar (eski parser gibi).')
$tip.SetToolTip($btnParseScript, 'Ustteki 2) SCCM Script bolumundeki script secimini kullanir. Hedef: bu tabloda secili satirlar; yoksa soldaki liste. Sonuc SADECE Panel 2''ye yazilir.')
$tip.SetToolTip($btnParseRecheck, 'Tabloda secili satirlarin (yoksa tum cihazlarin) DNS / AD / SCCM ajan kontrolunu yeniden calistirir.')

# -- Filtre cubugu
$pnlTop2 = New-Object System.Windows.Forms.Panel
$pnlTop2.Dock = 'Top'
$pnlTop2.Height = 36

$lblF2 = New-Object System.Windows.Forms.Label
$lblF2.Text = 'Filtre:'
$lblF2.Location = New-Object System.Drawing.Point(6, 11)
$lblF2.AutoSize = $true

$cboF2 = New-Object System.Windows.Forms.ComboBox
$cboF2.DropDownStyle = 'DropDownList'
$cboF2.Location = New-Object System.Drawing.Point(50, 7)
$cboF2.Size = New-Object System.Drawing.Size(200, 23)

$txtF2 = New-Object System.Windows.Forms.TextBox
$txtF2.Location = New-Object System.Drawing.Point(256, 7)
$txtF2.Size = New-Object System.Drawing.Size(200, 23)

$btnClearF2 = New-Object System.Windows.Forms.Button
$btnClearF2.Text = 'Temizle'
$btnClearF2.Location = New-Object System.Drawing.Point(462, 5)
$btnClearF2.Size = New-Object System.Drawing.Size(65, 27)

$lblCount2 = New-Object System.Windows.Forms.Label
$lblCount2.Text = ''
$lblCount2.Location = New-Object System.Drawing.Point(676, 11)
$lblCount2.AutoSize = $true
$lblCount2.ForeColor = [System.Drawing.Color]::DimGray

$btnUnfocus2 = New-Object System.Windows.Forms.Button
$btnUnfocus2.Text = 'Focus Kaybet'
$btnUnfocus2.Location = New-Object System.Drawing.Point(534, 5)
$btnUnfocus2.Size = New-Object System.Drawing.Size(124, 27)

$pnlTop2.Controls.AddRange(@($lblF2, $cboF2, $txtF2, $btnClearF2, $btnUnfocus2, $lblCount2))

# -- Kayit cubugu (eski parser: kayit klasoru + Gozat + CSV Olustur ve Ac)
$pnlSave2 = New-Object System.Windows.Forms.Panel
$pnlSave2.Dock = 'Bottom'
$pnlSave2.Height = 36

$lblPath2 = New-Object System.Windows.Forms.Label
$lblPath2.Text = 'Kayit klasoru:'
$lblPath2.Location = New-Object System.Drawing.Point(6, 11)
$lblPath2.AutoSize = $true

$txtPath2 = New-Object System.Windows.Forms.TextBox
$txtPath2.Text = $script:ParserDefaultFolder
$txtPath2.Location = New-Object System.Drawing.Point(92, 7)
$txtPath2.Size = New-Object System.Drawing.Size(380, 23)

$btnBrowse2 = New-Object System.Windows.Forms.Button
$btnBrowse2.Text = 'Gozat...'
$btnBrowse2.Location = New-Object System.Drawing.Point(480, 5)
$btnBrowse2.Size = New-Object System.Drawing.Size(70, 27)

$btnExcel2 = New-Object System.Windows.Forms.Button
$btnExcel2.Text = "Excel'e Aktar"
$btnExcel2.Location = New-Object System.Drawing.Point(558, 5)
$btnExcel2.Size = New-Object System.Drawing.Size(110, 27)

$btnCsv2 = New-Object System.Windows.Forms.Button
$btnCsv2.Text = 'CSV Olustur ve Ac'
$btnCsv2.Location = New-Object System.Drawing.Point(676, 5)
$btnCsv2.Size = New-Object System.Drawing.Size(140, 27)

$pnlSave2.Controls.AddRange(@($lblPath2, $txtPath2, $btnBrowse2, $btnExcel2, $btnCsv2))

$dgvParse = New-Object System.Windows.Forms.DataGridView
$dgvParse.Dock = 'Fill'
$dgvParse.ReadOnly = $true
$dgvParse.AllowUserToAddRows = $false
$dgvParse.AllowUserToDeleteRows = $false
$dgvParse.RowHeadersVisible = $false
$dgvParse.SelectionMode = 'FullRowSelect'
$dgvParse.MultiSelect = $true
$dgvParse.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'
$dgvParse.AutoGenerateColumns = $false
$dgvParse.AutoSizeColumnsMode = 'None'

foreach ($cn in $script:ParserColumns) {
    $col = New-GridColumn $cn $cn 10
    $col.Width = $(if ($cn -eq 'UpdateName') { 300 } else { 130 })
    [void]$dgvParse.Columns.Add($col)
}
foreach ($cn in $script:ParserCheckColumns) {
    $col = New-GridColumn $cn $cn 10
    $col.Width = $(if ($cn -eq 'SCCM_Kayit') { 200 } elseif ($cn -eq 'DNS_Durumu') { 150 } elseif ($cn -eq 'AD_LastLogon') { 130 } else { 120 })
    [void]$dgvParse.Columns.Add($col)
}
$colUy = New-GridColumn 'KolonUyumu' 'KolonUyumu' 10
$colUy.Width = 170
[void]$dgvParse.Columns.Add($colUy)
[void]$dgvParse.Columns.Add((New-GridColumn 'SourceHost' 'SourceHost' 10 $false))
$dgvParse.DataSource = $script:ParseBinding

$dgvParse.add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $rowView = $dgvParse.Rows[$e.RowIndex].DataBoundItem
    if (-not $rowView) { return }

    $cn  = $dgvParse.Columns[$e.ColumnIndex].Name
    $v   = "$($rowView[$cn])"
    $red = [System.Drawing.Color]::LightCoral
    $grn = [System.Drawing.Color]::LightGreen
    $yel = [System.Drawing.Color]::LightYellow

    if ("$($rowView['KolonUyumu'])" -notlike 'OK*') { $e.CellStyle.BackColor = $yel }

    if ($cn -eq 'DNS_Durumu') {
        if ($v -like 'No such host*' -or $v -like 'DNS hatasi*' -or $v -like 'Islem hatasi*') { $e.CellStyle.BackColor = $red }
        elseif ($v -and $v -notlike 'Kontrol ediliyor*' -and $v -notlike 'Bekleniyor*') { $e.CellStyle.BackColor = $grn }
    }
    elseif ($cn -eq 'Online_mi') {
        if ($v -eq 'Evet') { $e.CellStyle.BackColor = $grn }
        elseif ($v -eq 'Hayir') { $e.CellStyle.BackColor = $red }
    }
    elseif ($cn -eq 'AD_Durumu') {
        if ($v -eq 'Enabled') { $e.CellStyle.BackColor = $grn }
        elseif ($v -eq 'Disabled' -or $v -eq 'Bulunamadi') { $e.CellStyle.BackColor = $red }
        elseif ($v) { $e.CellStyle.BackColor = $yel }
    }
    elseif ($cn -eq 'CcmExec_Durumu') {
        if ($v -eq 'Running') { $e.CellStyle.BackColor = $grn }
        elseif ($v -eq 'Yok') { $e.CellStyle.BackColor = $red }
        elseif ($v) { $e.CellStyle.BackColor = $yel }
    }
    elseif ($cn -eq 'SCCM_Kayit') {
        if ($v -like 'Client yuklu*') { $e.CellStyle.BackColor = $grn }
        elseif ($v -like 'SCCM de kayitli degil*' -or $v -like 'Kayit var*') { $e.CellStyle.BackColor = $red }
        elseif ($v) { $e.CellStyle.BackColor = $yel }
    }

    if ($v -like 'Bekleniyor*') { $e.CellStyle.BackColor = [System.Drawing.Color]::Gainsboro }
})

$dgvParse.add_CellDoubleClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $rv = $dgvParse.Rows[$e.RowIndex].DataBoundItem
    if (-not $rv) { return }
    $hn = [string]$rv['Hostname']
    $sh = [string]$rv['SourceHost']
    $txt = '(Detay yok)'
    if ($script:ParserDetails.ContainsKey($hn))     { $txt = $script:ParserDetails[$hn] }
    elseif ($script:ParserDetails.ContainsKey($sh)) { $txt = $script:ParserDetails[$sh] }
    Show-TextDialog -Title "Detay: $hn" -Text $txt
})
# Dock sirasi: Fill kontrol ILK eklenir, Top kontroller sonra (en son eklenen en ustte durur)
$tabOut.Controls.AddRange(@($dgv, $pnlTop1))
$tabParse.Controls.AddRange(@($dgvParse, $pnlSave2, $pnlTop2, $pnlRun2, $pnlPaste, $lblParserStatus, $lblParserInfo))
[void]$tabs.TabPages.Add($tabOut)
[void]$tabs.TabPages.Add($tabParse)

# Filtre kolon listeleri
$script:ColMap1 = Get-ColumnMap -Grid $dgv
$script:ColMap2 = Get-ColumnMap -Grid $dgvParse
[void]$cboF1.Items.Add($script:AllColsLabel)
foreach ($k in $script:ColMap1.Keys) { [void]$cboF1.Items.Add($k) }
$cboF1.SelectedIndex = 0
[void]$cboF2.Items.Add($script:AllColsLabel)
foreach ($k in $script:ColMap2.Keys) { [void]$cboF2.Items.Add($k) }
$cboF2.SelectedIndex = 0

$form.Controls.AddRange(@($lblComputers, $txtComputers, $gbCheck, $gbSccm, $gbActions, $lblStatus, $tabs))

# ==================== UI DURUM YONETIMI ====================

function Update-ButtonStates {
    $busy1 = Test-PanelBusy -Panel 1
    $busy2 = Test-PanelBusy -Panel 2

    $btnFix.Enabled = ((@($script:LastResults | Where-Object { $_.IsOnline }).Count -gt 0) -and -not $busy1)

    $idx = $cboScripts.SelectedIndex
    $ok = $false
    if ($script:SiteCode -and $idx -ge 0 -and $idx -lt @($script:CmScripts).Count) {
        $ok = ($script:CmScripts[$idx].Approval -eq 3)
    }
    $btnSendScript.Enabled  = ($ok -and -not $busy1)     # Panel 1
    $btnParseScript.Enabled = ($ok -and -not $busy2)     # Panel 2
    $btnShowScript.Enabled = ($script:SiteCode -and $idx -ge 0)
}

function Set-UiBusy {
    # Panel bazli: Panel 1 mesgulken Panel 2 butonlari acik kalir (ve tersi).
    param([bool]$Busy, [int]$Panel = 1)

    if ($Panel -eq 2) {
        foreach ($b in @($btnParseRun, $btnParseScript, $btnParseRecheck)) { $b.Enabled = -not $Busy }
    }
    else {
        foreach ($b in @($btnCheck, $btnAd, $btnFix, $btnAction, $btnSendScript, $btnConnect, $btnShowScript)) { $b.Enabled = -not $Busy }
    }
    if (-not $Busy) { Update-ButtonStates }
}
# ==================== AYARLAR ====================

function Import-Settings {
    try {
        $f = Get-SettingsFile
        if (-not (Test-Path $f)) { return }
        $s = Get-Content -Path $f -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($s.SmsServer) { $txtSms.Text = [string]$s.SmsServer }
        if ($s.Services)  { $txtServices.Text = [string]$s.Services }
        if ($s.Timeout) {
            $t = [int]$s.Timeout
            if ($t -ge $nudTimeout.Minimum -and $t -le $nudTimeout.Maximum) { $nudTimeout.Value = $t }
        }
        if ($null -ne $s.AllDCs) { $chkAllDc.Checked = [bool]$s.AllDCs }
        if ($s.ParserFolder) { $txtPath2.Text = [string]$s.ParserFolder }
        if ($null -ne $s.ParserCheck) { $chkParseCheck.Checked = [bool]$s.ParserCheck }
        if ($s.LastScriptGuid) { $script:PendingScriptGuid = [string]$s.LastScriptGuid }
    }
    catch { }
}

function Save-Settings {
    try {
        $guid = ''
        $idx = $cboScripts.SelectedIndex
        if ($idx -ge 0 -and $idx -lt @($script:CmScripts).Count) { $guid = $script:CmScripts[$idx].Guid }
        elseif ($script:PendingScriptGuid) { $guid = $script:PendingScriptGuid }
        $obj = [pscustomobject]@{
            SmsServer      = $txtSms.Text.Trim()
            Services       = $txtServices.Text.Trim()
            Timeout        = [int]$nudTimeout.Value
            AllDCs         = [bool]$chkAllDc.Checked
            ParserFolder   = $txtPath2.Text.Trim()
            ParserCheck    = [bool]$chkParseCheck.Checked
            LastScriptGuid = $guid
        }
        Set-Content -Path (Get-SettingsFile) -Value ($obj | ConvertTo-Json) -Encoding UTF8
    }
    catch { }
}

# ==================== OLAY YONETICILERI ====================

# ---------- Filtreler (Panel 1 ve Panel 2) ----------
$txtF1.Add_TextChanged({ Set-GridFilter -Source $script:BindingSrc -ColMap $script:ColMap1 -ColumnHeader ([string]$cboF1.SelectedItem) -Text $txtF1.Text })
$cboF1.Add_SelectedIndexChanged({ Set-GridFilter -Source $script:BindingSrc -ColMap $script:ColMap1 -ColumnHeader ([string]$cboF1.SelectedItem) -Text $txtF1.Text })
$btnClearF1.Add_Click({ $txtF1.Text = ''; $cboF1.SelectedIndex = 0 })

$txtF2.Add_TextChanged({ Set-GridFilter -Source $script:ParseBinding -ColMap $script:ColMap2 -ColumnHeader ([string]$cboF2.SelectedItem) -Text $txtF2.Text })
$cboF2.Add_SelectedIndexChanged({ Set-GridFilter -Source $script:ParseBinding -ColMap $script:ColMap2 -ColumnHeader ([string]$cboF2.SelectedItem) -Text $txtF2.Text })
$btnClearF2.Add_Click({ $txtF2.Text = ''; $cboF2.SelectedIndex = 0 })

# Gosterilen / toplam satir sayaci (filtre uygulaninca ilk deger kuculur)
$script:BindingSrc.Add_ListChanged({ $lblCount1.Text = "Satir: $($script:BindingSrc.Count) / $($script:ResultsTable.Rows.Count)" })
$script:ParseBinding.Add_ListChanged({ $lblCount2.Text = "Satir: $($script:ParseBinding.Count) / $($script:ParseTable.Rows.Count)" })

# ---------- Kontrol Et ----------
$btnCheck.Add_Click({
    $computers = @(Get-ComputerList)
    if ($computers.Count -eq 0) { return }

    $services = @($txtServices.Text -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($services.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('En az bir servis ismi girin.', 'Uyari') | Out-Null
        return
    }

    $script:ResultsTable.Rows.Clear()
    $script:LastResults = @()
    $script:Details.Clear()

    $allDc = [bool]$chkAllDc.Checked
    $items = @(foreach ($h in $computers) {
        [pscustomobject]@{ Key = $h; Params = @{ ComputerName = $h; ServiceNames = $services; QueryAllDCs = $allDc } }
    })
    Set-GridPending -Hosts $computers -ClearError -Columns @('DnsStatus', 'IsOnline', 'AdStatus', 'LastLogon', 'Running', 'NotRunning', 'CcmExec')
    # Tablo bastan olustu: otomatik secilen ilk satir sonraki islemlerin hedefini daraltmasin
    try { $dgv.ClearSelection(); $dgv.CurrentCell = $null } catch { }
    Start-ParallelRun -Title 'Kontrol' -Worker $script:CheckWorker -Items $items -OnResultFn 'Update-GridFromCheck'
})

# ---------- Servisleri Baslat ----------
$btnFix.Add_Click({
    $targets = @($script:LastResults | Where-Object {
        $_.IsOnline -and $_.ServiceDetail -and
        (@($_.ServiceDetail | Where-Object { $_.Status -ne 'Running' -and $_.Status -ne 'NotFound' -and $_.Status -ne 'Erisilemedi' })).Count -gt 0
    })

    if ($targets.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Baslatilacak bir servis bulunamadi (hepsi calisiyor / bulunamadi / erisilemedi).', 'Bilgi') | Out-Null
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "$($targets.Count) makinede durdurulmus servisler bulundu. Baslatilsin mi? (Ayni anda en fazla $($script:MaxParallel) makine islenir)",
        'Onay', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $items = @(foreach ($t in $targets) {
        [pscustomobject]@{ Key = $t.Hostname; Params = @{ ComputerName = $t.Hostname; ServiceDetail = $t.ServiceDetail } }
    })
    Set-GridPending -Hosts @($targets | ForEach-Object { $_.Hostname }) -Columns @('ActionResult')
    Start-ParallelRun -Title 'Servis baslatma' -Worker $script:FixWorker -Items $items -OnResultFn 'Update-GridFromFix'
})

# ---------- AD Yapisi Kontrolu (script gondermeden; DNS + ping + AD) ----------
$btnAd.Add_Click({
    $t = Get-TargetHosts -Grid $dgv
    if (-not $t) { return }

    $allDc = [bool]$chkAllDc.Checked
    $items = @(foreach ($h in $t.Hosts) {
        [pscustomobject]@{ Key = $h; Params = @{ ComputerName = $h; QueryAllDCs = $allDc } }
    })
    $lblStatus.Text = "AD yapisi kontrolu: $(@($t.Hosts).Count) makine ($($t.Source))"
    Set-GridPending -Hosts $t.Hosts -ClearError -Columns @('DnsStatus', 'IsOnline', 'AdStatus', 'LastLogon', 'ActionResult')
    Start-ParallelRun -Title 'AD yapisi kontrolu' -Worker $script:AdWorker -Items $items -OnResultFn 'Update-GridFromAd'
})

# ---------- SCCM istemci eylemi ----------
$btnAction.Add_Click({
    $actionName = [string]$cboAction.SelectedItem
    if (-not $actionName) { return }
    $ids = [string[]]@($script:ClientActions[$actionName])

    $t = Get-TargetHosts -Grid $dgv
    if (-not $t) { return }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Eylem: $actionName`n`nHedef ($($t.Source)): $(@($t.Hosts).Count) makine`n$(Format-HostPreview -Hosts $t.Hosts)`n`nGonderilsin mi?",
        'Onay', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $items = @(foreach ($h in $t.Hosts) {
        [pscustomobject]@{ Key = $h; Params = @{ ComputerName = $h; ActionName = $actionName; ScheduleIds = $ids } }
    })
    Set-GridPending -Hosts $t.Hosts -Columns @('ActionResult')
    Start-ParallelRun -Title 'SCCM eylemi' -Worker $script:ActionWorker -Items $items -OnResultFn 'Update-GridFromAction'
})

# ---------- SCCM'e baglan / scriptleri listele ----------
$btnConnect.Add_Click({
    $srv = $txtSms.Text.Trim()
    if ($srv -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') {
        [System.Windows.Forms.MessageBox]::Show('Gecerli bir SMS Provider / Site Server adi girin.', 'Uyari') | Out-Null
        return
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $lblStatus.Text = "SMS Provider'a baglaniliyor: $srv ..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $site = Get-CmSiteCode -Server $srv
        $list = @(Get-CmScriptList -Server $srv -SiteCode $site)

        $script:SmsServer = $srv
        $script:SiteCode  = $site
        $script:CmScripts = $list
        $lblSite.Text = "Site: $site"

        $cboScripts.Items.Clear()
        foreach ($s in $list) { [void]$cboScripts.Items.Add($s.Display) }

        $sel = -1
        if ($script:PendingScriptGuid) {
            for ($i = 0; $i -lt $list.Count; $i++) { if ($list[$i].Guid -eq $script:PendingScriptGuid) { $sel = $i; break } }
        }
        if ($sel -lt 0 -and $list.Count -gt 0) { $sel = 0 }
        if ($sel -ge 0) { $cboScripts.SelectedIndex = $sel }

        $approved = @($list | Where-Object { $_.Approval -eq 3 }).Count
        $lblStatus.Text = "SCCM baglantisi tamam. Site: $site, $($list.Count) script ($approved onayli)."
        Save-Settings
    }
    catch {
        $script:SiteCode = ''
        $script:CmScripts = @()
        $cboScripts.Items.Clear()
        $lblSite.Text = 'Site: -'
        $lblStatus.Text = 'SCCM baglantisi basarisiz.'
        [System.Windows.Forms.MessageBox]::Show("SMS Provider'a baglanilamadi veya scriptler listelenemedi:`n`n$($_.Exception.Message)", 'Hata') | Out-Null
    }
    finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        Update-ButtonStates
    }
})

$cboScripts.Add_SelectedIndexChanged({
    $idx = $cboScripts.SelectedIndex
    if ($idx -ge 0 -and $idx -lt @($script:CmScripts).Count) {
        $s = $script:CmScripts[$idx]
        $lblScriptInfo.Text = "Yazar: $($s.Author)  |  $($s.ApprovalText)  |  v$($s.Version)  |  $($s.Guid)"
    }
    else { $lblScriptInfo.Text = '' }
    Update-ButtonStates
})

$btnShowScript.Add_Click({
    $idx = $cboScripts.SelectedIndex
    if ($idx -lt 0) { return }
    $s = $script:CmScripts[$idx]
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $txt = Get-CmScriptText -Path $s.Path
        Show-TextDialog -Title "Script: $($s.Name) (v$($s.Version))" -Text $txt
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Script icerigi okunamadi:`n$($_.Exception.Message)", 'Hata') | Out-Null
    }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
})

# ---------- Secili scripti gonder (SADECE Panel 1'e yazar) ----------
$btnSendScript.Add_Click({ Start-ScriptSend -Panel 1 })
# ---------- Disa aktarma: Panel 1 ----------
$btnExcel1.Add_Click({
    if ($script:ResultsTable.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Aktarilacak sonuc yok.', 'Bilgi') | Out-Null
        return
    }
    Export-ToExcelAndOpen -Data (Get-GridExportData -Grid $dgv -Source $script:BindingSrc -WithDetail $true) -DefaultFileName "IstemciRaporu_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
})

$btnCsv1.Add_Click({
    if ($script:ResultsTable.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Aktarilacak sonuc yok.', 'Bilgi') | Out-Null
        return
    }
    Export-GridToCsv -Data (Get-GridExportData -Grid $dgv -Source $script:BindingSrc -WithDetail $true) -DefaultFileName "IstemciRaporu_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
})

# ---------- Panel 2 (Update Parser) olaylari - Panel 1'e DOKUNMAZ ----------
# Aktarilan kolonlar: 21 parser kolonu + 6 kontrol kolonu (KolonUyumu aktarilmaz).

$btnBrowse2.Add_Click({
    $fd = New-Object System.Windows.Forms.FolderBrowserDialog
    if (Test-Path $txtPath2.Text) { $fd.SelectedPath = $txtPath2.Text }
    if ($fd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtPath2.Text = $fd.SelectedPath }
})

$btnExcel2.Add_Click({
    if ($script:ParseTable.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Aktarilacak parser verisi yok.', 'Bilgi') | Out-Null
        return
    }
    Export-ToExcelAndOpen -Data (Get-GridExportData -Grid $dgvParse -Source $script:ParseBinding -SkipColumns @('KolonUyumu')) -DefaultFileName "$(Get-Date -Format 'dd-HH')-updates"
})

$btnCsv2.Add_Click({
    # Eski parser gibi: kayit klasorune "<gun-saat>-updates.csv" yazar ve otomatik acar.
    $lblParserStatus.ForeColor = [System.Drawing.Color]::DarkRed
    if ($script:ParseTable.Rows.Count -eq 0) {
        $lblParserStatus.Text = 'Aktarilacak parser verisi yok.'
        return
    }
    $folder = $txtPath2.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($folder)) {
        $lblParserStatus.Text = 'Kayit klasoru bos olamaz.'
        return
    }
    if (-not (Test-Path $folder)) {
        try { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
        catch {
            $lblParserStatus.Text = "Klasor olusturulamadi: $($_.Exception.Message)"
            return
        }
    }

    $fullPath = Join-Path $folder ("{0}-updates.csv" -f (Get-Date -Format 'dd-HH'))
    $data = Get-GridExportData -Grid $dgvParse -Source $script:ParseBinding -SkipColumns @('KolonUyumu')
    try { $data | Export-Csv -Path $fullPath -NoTypeInformation -Encoding UTF8 -Delimiter ';' -Force }
    catch {
        $lblParserStatus.Text = "CSV yazilamadi: $($_.Exception.Message)"
        return
    }

    $bad = 0
    foreach ($row in $script:ParseTable.Rows) { if ([string]$row['KolonUyumu'] -notlike 'OK*') { $bad++ } }
    if ($bad -gt 0) {
        $lblParserStatus.ForeColor = [System.Drawing.Color]::DarkOrange
        $lblParserStatus.Text = "CSV olusturuldu ($fullPath) ancak $bad satir kolon uyumsuz / bilgi satiri. Basliklarla hizasini kontrol et."
    }
    else {
        $lblParserStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        $lblParserStatus.Text = "CSV basariyla olusturuldu: $fullPath"
    }

    try { Invoke-Item -Path $fullPath }
    catch { [System.Windows.Forms.MessageBox]::Show("CSV olusturuldu ancak otomatik acilamadi:`n$fullPath", 'Bilgi') | Out-Null }
})

$btnParseRun.Add_Click({
    $lblParserStatus.ForeColor = [System.Drawing.Color]::DarkRed
    if ([string]::IsNullOrWhiteSpace($txtParsePaste.Text)) {
        $lblParserStatus.Text = 'Once metin kutusuna SCCM sonucunu yapistirin.'
        return
    }

    if (-not $chkParseAppend.Checked) {
        $script:ParseTable.Rows.Clear()
        $script:ParserNoOutput = @{}
        $script:ParserDetails.Clear()
    }

    $script:ParsePasteSeq++
    $res = Add-ParserLines -SourceHost "(yapistirilan $(Get-Date -Format 'HH:mm:ss') #$($script:ParsePasteSeq))" -Text $txtParsePaste.Text
    Update-ParserInfo
    if (-not $chkParseAppend.Checked) {
        try { $dgvParse.ClearSelection(); $dgvParse.CurrentCell = $null } catch { }
    }

    if ($res.Added -eq 0) {
        $lblParserStatus.Text = 'Gecerli bir satir bulunamadi.'
        return
    }

    $hostCount = @($res.Hosts).Count
    if ($chkParseCheck.Checked -and $hostCount -gt 0) {
        $lblParserStatus.ForeColor = [System.Drawing.Color]::DarkBlue
        $lblParserStatus.Text = "$($res.Added) satir parse edildi, $hostCount cihaz icin DNS / AD / SCCM kontrolu basliyor..."
        Start-ParserChecks -Hosts @($res.Hosts)
    }
    else {
        $lblParserStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        $lblParserStatus.Text = "$($res.Added) satir parse edildi ($hostCount cihaz)."
    }
})

$btnParseScript.Add_Click({ Start-ScriptSend -Panel 2 })

$btnParseRecheck.Add_Click({
    $hostsToCheck = @()
    $src = 'secili satirlar'
    foreach ($r in $dgvParse.SelectedRows) {
        $h = [string]$r.Cells['Hostname'].Value
        if ($h -and ($h -match $script:HostRegex) -and ($hostsToCheck -notcontains $h)) { $hostsToCheck += $h }
    }
    if ($hostsToCheck.Count -eq 0) {
        $src = 'tablodaki tum cihazlar'
        foreach ($row in $script:ParseTable.Rows) {
            $h = [string]$row['Hostname']
            if ($h -and ($h -match $script:HostRegex) -and ($hostsToCheck -notcontains $h)) { $hostsToCheck += $h }
        }
    }
    if ($hostsToCheck.Count -eq 0) {
        $lblParserStatus.ForeColor = [System.Drawing.Color]::DarkRed
        $lblParserStatus.Text = 'Kontrol edilecek gecerli bir cihaz adi bulunamadi.'
        return
    }
    $lblParserStatus.ForeColor = [System.Drawing.Color]::DarkBlue
    $lblParserStatus.Text = "Kontrol: $($hostsToCheck.Count) cihaz ($src)"
    Start-ParserChecks -Hosts $hostsToCheck
})

$btnParseClear.Add_Click({
    if ($script:ParseTable.Rows.Count -eq 0) { return }
    $c = [System.Windows.Forms.MessageBox]::Show('Update Parser tablosundaki tum satirlar silinsin mi?', 'Onay',
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($c -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    $script:ParseTable.Rows.Clear()
    $script:ParserNoOutput = @{}
    $script:ParserDetails.Clear()
    Update-ParserInfo
})
# ---------- Grid focus / secim birakma (iki panel icin) ----------
# Secili satir varsa hedef olarak o satirlar kullanilir; bu buton secimi ve aktif hucreyi kaldirir,
# odagi soldaki bilgisayar listesine tasir (hedef tekrar soldaki liste olur).
function Clear-GridFocus {
    param($Grid)
    try {
        $Grid.ClearSelection()
        $Grid.CurrentCell = $null
    }
    catch { }
    $null = $txtComputers.Focus()
}
$btnUnfocus1.Add_Click({ Clear-GridFocus -Grid $dgv })
$btnUnfocus2.Add_Click({ Clear-GridFocus -Grid $dgvParse })

# Tablonun BOS alanina (satirlarin altina) tiklayinca veya Esc'ye basinca da secim / focus kalkar
foreach ($g in @($dgv, $dgvParse)) {
    $g.Add_MouseDown({
        param($s, $e)
        if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
        $hit = $s.HitTest($e.X, $e.Y)
        if ($hit.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::None) { Clear-GridFocus -Grid $s }
    })
    $g.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            Clear-GridFocus -Grid $s
            $e.Handled = $true
        }
    })
}

$script:Tip = New-Object System.Windows.Forms.ToolTip
$script:Tip.SetToolTip($btnUnfocus1, 'Tablodaki secimi ve aktif hucreyi kaldirir; hedef tekrar soldaki liste olur. (Tablonun bos alanina tiklamak veya Esc de ayni isi yapar.)')
$script:Tip.SetToolTip($btnUnfocus2, 'Tablodaki secimi ve aktif hucreyi kaldirir; hedef tekrar soldaki liste olur. (Tablonun bos alanina tiklamak veya Esc de ayni isi yapar.)')

# ==================== FORM KAPANIRKEN TEMIZLIK ====================

$form.Add_FormClosing({
    Save-Settings
    if ($script:RunTimer.Enabled) { $script:RunTimer.Stop() }
    if ($script:SharedPool) { try { $script:SharedPool.Close(); $script:SharedPool.Dispose() } catch { } }
})

# ==================== BASLANGIC ====================

Import-Settings
[void]$form.ShowDialog()
