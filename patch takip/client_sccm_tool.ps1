#Requires -Version 5.1
<#
.SYNOPSIS
    Istemci Durum ve SCCM Islem Araci (WinForms GUI)

.ACIKLAMA
    Sol textbox: bilgisayar isimleri (her satira bir tane).

    1) Kontrol Et      : DNS, Cevrimici (ping + KIMLIK DOGRULAMA), AD (onbellekten), servisler + CcmExec.
       Servisleri Baslat: durdurulmus servisleri baslatir.  AD Yapisi Kontrolu: DNS + ping + AD (servissiz).
    2) SCCM Script     : SMS Provider'daki ONAYLI scripti hedeflerde calistirir, sonucu bekler. Cevap gelmezse sebep
                         (cevrimdisi / IP baska cihazda / SCCM ajani yok / cevrimici ama calismadi) yazilir.
    3) SCCM Istemci Eylemi: Machine Policy, Update Scan/Deployment, App Deployment, HW/SW Inventory, DDR tetikler.
       Hedef: gridde secili hucrelerin satirlari; secim yoksa soldaki listedeki tum makineler.

    CEVRIMICI KONTROLU (sahte "online" onlemi): VPN'e baglanmis makinelerin DNS kaydi eski IP'de kalabilir; o IP
    baska cihaza verilince ping cevap verir. Bu yuzden ping cevabi veren cihazin ADI okunur (NetBIOS node status,
    yoksa WMI Win32_ComputerSystem) ve beklenen adla karsilastirilir:
      Evet | Hayir | Hayir (IP baska cihazda: <ad>) | Hayir (ping var, cihaz dogrulanamadi: <IP>) -> sadece Evet'te uzak islem.
      DNS "joker" cevaplari (olmayan ada gateway IP'si donmesi) kayit yok sayilir.

    PANEL 1 - Script Ciktilari / Kontroller : DNS, Online, AD, Last Logon, servisler, script ciktisi (detay: cift tik).
    PANEL 2 - Update Parser (Panel 1'den BAGIMSIZ): yapistirilan / Panel 2'den gonderilen script ciktisini ('@' ile
              ayrilan kayitlar, ';' ile ayrilan kolonlar) alt alta kolonlara boler; her cihaz icin DNS, Online, AD,
              SERVISLER (1. bolumdeki liste), CcmExec ve SCCM kaydi kontrol edilir. Filtre, Excel, CSV.
    Tablolar: hucre secilebilir; Ctrl+C secili hucre(ler)i kopyalar, sag tik menusunde diger kopyalama secenekleri.

    AD ONBELLEGI: AD bilgisi diskteki CSV'lerden okunur (ServiceStatus\AdCache\, OU basina bir dosya); "AD Onbellegi..." penceresinde
    domain -> OU eslemesi girilir, "Kaydet ve Onbellegi Guncelle" ile OU'lar (alt OU dahil) AD'den taranir.
    Onbellek sadece kullanici isterse guncellenir. Last Logon = lastLogonTimestamp (~14 gun gecikmeli).

.MIMARI
    - WinRM / Invoke-Command / PSSession KULLANILMAZ. Sadece RPC / DCOM / WMI, ServiceController, NetBIOS (UDP 137).
    - Internete cikmaz. Kimlik bilgisi sorulmaz: scripti calistiran kullanicinin yetkileri (Integrated Windows Auth).
    - RunspacePool, SABIT en fazla 10 eszamanli hedef (Panel 1 + Panel 2 ortak). UI donmaz.
    - SCCM konsolu / ConfigurationManager modulu GEREKMEZ.

.GEREKSINIMLER
    - Windows PowerShell 5.1, STA (gerekirse script kendini STA ile yeniden baslatir)
    - RSAT ActiveDirectory modulu + ADWS (TCP 9389) -> sadece AD onbellegi guncellenirken
    - Hedeflerde yerel admin (servis + root\ccm), RPC/DCOM erisimi; SCCM'de onayli script + "Run Script" yetkisi

.KAYITLAR  (%USERPROFILE%\Desktop\ServiceStatus\)
    ClientToolSettings.json (ayarlar) | AdCache\*.csv (AD onbellegi, OU basina) | ClientTool_yyyyMM.log (denetim kaydi)
#>

# WinForms ve pano (Ctrl+C) STA ister; Get-WmiObject sadece Windows PowerShell 5.1'de var.
if ($PSVersionTable.PSEdition -eq 'Core' -or [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $winPs = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if ($PSCommandPath -and (Test-Path -LiteralPath $winPs)) {
        Start-Process -FilePath $winPs -ArgumentList @('-NoProfile', '-STA', '-File', "`"$PSCommandPath`"")
        return
    }
}

Add-Type -AssemblyName System.Windows.Forms, System.Drawing, System.ServiceProcess, System.Management
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==================== GLOBAL DEGISKENLER ====================
function New-NameMap { return (New-Object System.Collections.Hashtable ([System.StringComparer]::OrdinalIgnoreCase)) }   # tr-TR'de I/i sorunu olmayan sozluk

$script:MaxParallel    = 10           # SABIT esik: ayni anda en fazla 10 makine
$script:LastResults    = New-NameMap  # Hostname -> son kontrol sonucu (IsOnline + ServiceDetail; "Servisleri Baslat" icin)
$script:Details        = New-NameMap  # Panel 1: Hostname -> satir detay/log metni
$script:ParserNoOutput = New-NameMap  # Panel 2: script'ten cikti gelmeyen hostlar
$script:ParserDetails  = New-NameMap  # Panel 2: Hostname -> detay metni
$script:SmsServer      = ''
$script:SiteCode       = ''
$script:CmScripts      = @()
$script:PendingScriptGuid = ''
$script:ParsePasteSeq  = 0            # Panel 2: "Parse Et" sayaci (her yapistirma icin ayri SourceHost etiketi)
$script:HostRegex      = '^[A-Za-z0-9İıŞşĞğÜüÖöÇç][A-Za-z0-9İıŞşĞğÜüÖöÇç._-]*$'   # Panel 1 + Panel 2 ortak
$script:SharedPool     = $null        # Panel 1 + Panel 2 islerinin ORTAK havuzu
$script:ActiveRuns     = New-Object System.Collections.ArrayList
$script:AllColsLabel   = '(Tum kolonlar)'
$script:PendingText    = 'Bekleniyor..'
$script:CancelText     = '__IPTAL__'   # iptal edilen islerin sonuc isleyicilerine gecen isaret
$script:CancelledJobs  = New-Object System.Collections.ArrayList   # iptal edilmis, uzak cagrisi bitmeyi bekleyen isler
$script:ScriptPoll     = New-Object System.Collections.ArrayList   # gonderilmis, sonucu toplu yoklamayla beklenen scriptler
$script:PollJob        = $null        # o an calisan toplu yoklama sorgusu (ayni anda en fazla 1)
$script:PollRunspace   = $null        # yoklama icin ayri runspace (10'luk havuzu kullanmaz)
$script:LastPoll       = [datetime]::MinValue
$script:PollCount      = 0            # SMS Provider'a giden yoklama turu sayaci (durum cubugunda gosterilir)
$script:PollError      = ''
$script:DeferFlag      = $false
$script:CurrentRun     = $null
$script:Prefetch       = New-Object System.Collections.ArrayList   # V2: SCCM kaydi on-alma isleri
$script:LastOp         = @{}          # V6: panel -> son islem (RunArgs + Items) "Basarisizlari Tekrar Dene" icin
$script:StateFilter    = @{ 1 = ''; 2 = '' }   # V7: ozet cubugu durum filtresi
$script:SumButtons     = @{}
$script:RetryButtons   = @{}
$script:SumDirty       = @{ 1 = $true; 2 = $true }
$script:DirtyRows      = New-Object 'System.Collections.Generic.HashSet[object]'   # V7: Durum'u yeniden hesaplanacak Panel 2 satirlari
$script:Auto           = @{}          # V13: panel -> otomatik yenileme ayarlari
$script:AdProgress     = $null        # AD onbellek guncelleme ilerlemesi (senkronize hashtable; worker'lar yazar)
$script:AdProgOrder    = @()
$script:AdSaveSummary  = ''
$script:AdStart        = $null
$script:SccmExtra      = New-NameMap  # kisa ad -> @(Boundary Group(s), Operating System Build) (son toplu SCCM sorgusundan)
$script:ReportQueue    = New-Object System.Collections.ArrayList   # gosterilmeyi bekleyen servis baslatma raporlari
$script:ShowingReport  = $false
$script:ParserSvc      = New-NameMap  # Panel 2: Hostname -> son servis durumu (IsOnline + ServiceDetail; Panel 2 "Servisleri Baslat" icin)

# AD ONBELLEGI domain / OU eslemesi (ClientToolSettings.json'da saklanir). OU = DistinguishedName (alt OU'lar dahil taranir)
$script:AdCacheSources = @()
$script:AdCacheNew     = New-NameMap  # guncelleme sirasinda: Domain -> worker sonucu
$script:ParserDefaultFolder = Join-Path $env:USERPROFILE 'Desktop\SCCM-Raporlari'

# Update Parser basligi (21 kolon). Eski 19 kolonluk cikti (client_update_parser.ps1) kolon ADIYLA yerlestirilir.
$script:ParserColumns = @('Hostname;UpdateName;EvalState;PercentComplete;ErrorCode;LastShutdownTime;IsBlueScreen;CBS_NTSTATUS_FROM_WIN32;CBS_STATUS_SXS_FILE_HASH_MISMATCH;UD_CERTIFICATE_ERROR_0x80010002;UD_DOWNLOAD_ERROR_0x80d02002;BuildVersion;SistemDrive_FreeGB;SistemDrive_FreePercent;Wuauserv_Durum;BITS_Durum;CryptSvc_Durum;TrustedInstaller_Durum;SoftwareDistribution_BoyutMB;DISM_CheckHealth;DTS_DP_IndirmeHatalari' -split ';')
$script:ParserLegacyColumns = @($script:ParserColumns | Where-Object { $_ -ne 'SoftwareDistribution_BoyutMB' -and $_ -ne 'DTS_DP_IndirmeHatalari' })

# Panel 2'de parse sonrasi eklenen kontrol kolonlari
# SCCM'den (SMS_CombinedDeviceResources) toplu sorguyla gelen kolonlar - Panel 1: BoundaryGroups / OsBuild
$script:ParserSccmColumns = @('Boundary_Groups', 'OS_Build')
$script:ParserCheckColumns = @('DNS_Durumu', 'Online_mi', 'AD_Durumu', 'AD_LastLogon', 'Calisan_Servisler', 'Calismayan_Servisler', 'CcmExec_Durumu', 'SCCM_Kayit')

# SCCM istemci eylemleri: gorunen ad -> TriggerSchedule ScheduleID listesi
$script:ClientActions = [ordered]@{
    'Machine Policy (Retrieval + Evaluation)'    = @('{00000000-0000-0000-0000-000000000021}', '{00000000-0000-0000-0000-000000000022}')
    'Software Updates Scan (Update Scan)'        = @('{00000000-0000-0000-0000-000000000113}')
    'Software Updates Deployment Evaluation'     = @('{00000000-0000-0000-0000-000000000108}')
    'Update Scan + Update Deployment Evaluation' = @('{00000000-0000-0000-0000-000000000113}', '{00000000-0000-0000-0000-000000000108}')
    'Application Deployment Evaluation'          = @('{00000000-0000-0000-0000-000000000121}')
    'Hardware Inventory'                         = @('{00000000-0000-0000-0000-000000000001}')
    'Software Inventory'                         = @('{00000000-0000-0000-0000-000000000002}')
    'Discovery Data Collection (DDR)'            = @('{00000000-0000-0000-0000-000000000003}')
    'Istemci Onarimi (RepairClient / ccmrepair)' = @('@REPAIR')    # V11: ozel eylem, ScheduleID degil
    'CcmExec (SMS Agent Host) Yeniden Baslat'    = @('@RESTART')
}

# ==================== PARALEL RUNSPACE'LERDE DE KULLANILAN FONKSIYONLAR ====================
# NOT: Bu bolumdeki fonksiyonlar hem ana pencere hem de RunspacePool icindeki runspace'ler tarafindan kullanilir.
# Disaridan hicbir $script: degiskenine referans vermezler. Kaynak kodlari New-WorkerScript ile worker'a gomulur.

function ConvertTo-WqlString {
    # WQL sorgusuna girecek metni kacislar (\ ve ').
    param([string]$Text)
    return ($Text -replace '\\', '\\' -replace "'", "\'")
}

function Get-ErrText {
    # Istisna zincirindeki en anlamli (en ic) mesaji dondurur.
    param($Exception)
    $e = $Exception
    while ($e.InnerException) { $e = $e.InnerException }
    if ($e.Message -and $e -ne $Exception) { return "$($Exception.Message) ($($e.Message))" }
    return $Exception.Message
}

function Test-Ping {
    # Hizli ICMP kontrolu (.NET Ping). PS 5.1'deki Test-Connection (WMI Win32_PingStatus) cok daha yavastir.
    param([string]$Address, [int]$Count = 2, [int]$TimeoutMs = 1000)
    $p = New-Object System.Net.NetworkInformation.Ping
    try {
        for ($i = 0; $i -lt $Count; $i++) {
            try { if ($p.Send($Address, $TimeoutMs).Status -eq 'Success') { return $true } } catch { }
        }
        return $false
    }
    finally { $p.Dispose() }
}

function Test-TcpPort {
    param([string]$Address, [int]$Port, [int]$TimeoutMs = 1500)
    $c = New-Object System.Net.Sockets.TcpClient
    try {
        $ar = $c.BeginConnect($Address, $Port, $null, $null)
        if (-not $ar.AsyncWaitHandle.WaitOne($TimeoutMs)) { return $false }
        $c.EndConnect($ar)
        return $true
    }
    catch { return $false }
    finally { $c.Close() }
}

function Get-NetBiosName {
    # NetBIOS node status sorgusu (nbtstat -A esdegeri, UDP 137). Kimlik dogrulamasi gerektirmez;
    # IP adresindeki cihazin GERCEK adini dondurur. Cevap yoksa $null.
    param([string]$Ip, [int]$TimeoutMs = 1000)
    $u = New-Object System.Net.Sockets.UdpClient
    try {
        $u.Client.ReceiveTimeout = $TimeoutMs
        # Baslik (TxID, flags=0, QD=1) + '*' adinin kodlanmis hali (CKAAAA...) + NBSTAT(0x21) / IN(1)
        $q = [byte[]](@(0x4E, 0x42, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0x20, 0x43, 0x4B) + @(,0x41 * 30) + @(0, 0, 0x21, 0, 1))
        [void]$u.Send($q, $q.Length, $Ip, 137)
        $ep = New-Object System.Net.IPEndPoint ([System.Net.IPAddress]::Any, 0)
        $b = $u.Receive([ref]$ep)
        if ($b.Length -lt 57) { return $null }
        for ($i = 0; $i -lt $b[56]; $i++) {
            $o = 57 + $i * 18
            if ($o + 18 -gt $b.Length) { break }
            # Sonek 0x00 (workstation) ve grup olmayan (flags bit 15 = 0) ilk ad = makine adi
            if ($b[$o + 15] -eq 0 -and ($b[$o + 16] -band 0x80) -eq 0) {
                return [System.Text.Encoding]::ASCII.GetString($b, $o, 15).Trim()
            }
        }
        return $null
    }
    catch { return $null }
    finally { $u.Close() }
}

function Get-WmiFailureReason {
    # WMI/DCOM hatasini siniflandirir: NoAgent / Rpc / Access / Identity / Other
    param($ErrorRecord)

    $kind  = 'Other'
    $first = $ErrorRecord.Exception.Message
    $e     = $ErrorRecord.Exception
    while ($e) {
        if ($e -is [System.Management.ManagementException]) {
            $code = $e.ErrorCode.ToString()
            if ($code -eq 'InvalidNamespace') { $kind = 'NoAgent'; break }
            if ($code -eq 'AccessDenied')     { $kind = 'Access';  break }
        }
        if ($e -is [System.UnauthorizedAccessException]) { $kind = 'Access'; break }
        if ($e -is [System.Runtime.InteropServices.COMException]) {
            $hr = $e.ErrorCode
            if ($hr -eq -2147023174 -or $hr -eq -2147023170) { $kind = 'Rpc';      break }   # 0x800706BA / 0x800706BE
            if ($hr -eq -2147024891)                         { $kind = 'Access';   break }   # 0x80070005
            if ($hr -eq -2147217394)                         { $kind = 'NoAgent';  break }   # 0x8004100E
            if ($hr -eq -2146893022)                         { $kind = 'Identity'; break }   # 0x80090322 SEC_E_WRONG_PRINCIPAL
        }
        $e = $e.InnerException
    }
    if ($kind -eq 'Other') {
        if     ($first -match 'Invalid namespace|0x8004100E')                    { $kind = 'NoAgent' }
        elseif ($first -match 'RPC server|0x800706BA|0x800706BE')                { $kind = 'Rpc' }
        elseif ($first -match 'Access is denied|0x80070005')                     { $kind = 'Access' }
        elseif ($first -match 'target principal name is incorrect|0x80090322')   { $kind = 'Identity' }
    }
    $text = switch ($kind) {
        'NoAgent'  { "SCCM ajani bulunmuyor (root\ccm namespace'i yok)." }
        'Rpc'      { 'RPC sunucusu kullanilamiyor (makine kapali/firewall).' }
        'Access'   { 'Erisim engellendi (hedefte yerel admin yetkisi gerekli).' }
        'Identity' { 'Kerberos hedef adi uyusmuyor (IP baska cihazda olabilir).' }
        default    { $first }
    }
    return [pscustomobject]@{ Kind = $kind; Text = $text }
}

function Get-WmiComputerName {
    # DCOM/WMI ile hedefin kendi bildirdigi bilgisayar adini okur (scripti calistiran kullanicinin yetkisiyle).
    param([string]$Target, [int]$TimeoutSec = 15)
    try {
        $opt = New-Object System.Management.ConnectionOptions
        $opt.Timeout = [TimeSpan]::FromSeconds($TimeoutSec)
        $sc = New-Object System.Management.ManagementScope ("\\$Target\root\cimv2", $opt)
        $sc.Connect()
        $eo = New-Object System.Management.EnumerationOptions
        $eo.Timeout = [TimeSpan]::FromSeconds($TimeoutSec)
        $s = New-Object System.Management.ManagementObjectSearcher ($sc, (New-Object System.Management.ObjectQuery 'SELECT Name FROM Win32_ComputerSystem'), $eo)
        foreach ($o in $s.Get()) { return [pscustomobject]@{ Name = [string]$o['Name']; Kind = ''; Text = '' } }
        return [pscustomobject]@{ Name = ''; Kind = 'Other'; Text = 'Win32_ComputerSystem bos dondu' }
    }
    catch {
        $why = Get-WmiFailureReason -ErrorRecord $_
        return [pscustomobject]@{ Name = ''; Kind = $why.Kind; Text = $why.Text }
    }
}

function Get-DnsWildcardIps {
    # Var olmayan RASTGELE bir ad cozumlenebiliyorsa DNS "joker" cevap veriyor demektir (router / modem DNS'i, wildcard
    # kayit): olmayan her ada ayni IP (or. gateway 192.168.1.1) doner ve ping cevap verir. Bu IP'ler gercek kayit sayilmaz.
    # Sonuc runspace basina 10 dk onbellekte tutulur.
    $c = $global:DnsWildcard
    if ($c -and ((Get-Date) - $c.Time).TotalMinutes -lt 10) { return $c.Ips }
    $ips = @()
    try { $ips = @([System.Net.Dns]::GetHostAddresses("ct-yok-$([guid]::NewGuid().ToString('N').Substring(0, 12))") | ForEach-Object { $_.IPAddressToString }) } catch { }
    $global:DnsWildcard = [pscustomobject]@{ Time = Get-Date; Ips = $ips }
    return $ips
}

function Resolve-HostDns {
    # Cozumlenemezse Text = "No such host is known" (isletim sistemi dilinden bagimsiz). Ips: once IPv4.
    # DNS joker cevabi (Get-DnsWildcardIps) kayit YOK sayilir.
    param([string]$ComputerName)
    try {
        $ips = @([System.Net.Dns]::GetHostAddresses($ComputerName))
        if ($ips.Count -eq 0) { return [pscustomobject]@{ Ok = $false; Text = 'No such host is known'; Ips = @() } }
        $v4  = @($ips | Where-Object { $_.AddressFamily -eq 'InterNetwork' })
        $use = @($(if ($v4.Count -gt 0) { $v4 } else { $ips }) | ForEach-Object { $_.IPAddressToString })
        $wc  = @(Get-DnsWildcardIps)
        if ($wc.Count -gt 0 -and @($use | Where-Object { $wc -notcontains $_ }).Count -eq 0) {
            return [pscustomobject]@{ Ok = $false; Text = "No such host is known (DNS joker cevap: $($use -join ', '))"; Ips = @() }
        }
        return [pscustomobject]@{ Ok = $true; Text = ($use -join ', '); Ips = $use }
    }
    catch {
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        if ($ex -is [System.Net.Sockets.SocketException] -and $ex.SocketErrorCode -eq [System.Net.Sockets.SocketError]::HostNotFound) {
            return [pscustomobject]@{ Ok = $false; Text = 'No such host is known'; Ips = @() }
        }
        return [pscustomobject]@{ Ok = $false; Text = "DNS hatasi: $($ex.Message)"; Ips = @() }
    }
}

function Get-HostReach {
    # DNS + ping + KIMLIK DOGRULAMA.
    # Sorun: VPN'e baglanmis bir makinenin DNS kaydi eski IP'de kalir; o IP baska bir cihaza verildiginde ping
    # cevap gelir ve makine "cevrimici" sanilir (servis/WMI istekleri de YANLIS cihaza gider).
    # Cozum: ping cevabi veren IP'deki cihazin adi okunur ve beklenen adla karsilastirilir:
    #   1) NetBIOS node status (UDP 137, kimlik dogrulamasiz, hizli)
    #   2) Cevap yoksa: TCP 135 on kontrolu + WMI Win32_ComputerSystem.Name (DCOM, IWA). Kerberos'un
    #      "target principal name is incorrect" hatasi da uyusmazlik kaniti sayilir.
    # Online = ping VAR ve ad UYUSMAZLIGI YOK. Identity: Dogrulandi | Uyusmuyor | Dogrulanamadi | '' (ping yok)
    param([string]$ComputerName)

    $short = (([string]$ComputerName -split '\.')[0]).Trim()
    $nbExp = $short.Substring(0, [Math]::Min(15, $short.Length))
    $r = [pscustomobject]@{ DnsOk = $false; DnsText = ''; Ip = ''; Online = $false; Identity = ''; ActualName = ''; OnlineText = 'Hayir'; Note = '' }

    $dns = Resolve-HostDns -ComputerName $ComputerName
    $r.DnsOk = $dns.Ok
    $r.DnsText = $dns.Text
    if (-not $dns.Ok) { $r.Note = 'DNS cozumlenemedi, ping/servis kontrolu yapilmadi'; return $r }

    $alive = @($dns.Ips | Where-Object { Test-Ping -Address $_ })
    if ($alive.Count -eq 0) { $r.Note = 'Makine yanit vermiyor (ping basarisiz)'; return $r }

    # 1) NetBIOS: ping veren her IP icin cihazin gercek adi
    $wrong = @()
    foreach ($ip in $alive) {
        $nb = Get-NetBiosName -Ip $ip
        if (-not $nb) { continue }
        if ([string]::Equals($nb, $nbExp, [System.StringComparison]::OrdinalIgnoreCase)) {
            $r.Ip = $ip; $r.Online = $true; $r.Identity = 'Dogrulandi'; $r.ActualName = $nb; $r.OnlineText = 'Evet'
            return $r
        }
        $wrong += [pscustomobject]@{ Ip = $ip; Name = $nb }
    }
    $r.Ip = $alive[0]

    # 2) WMI (ad uzerinden; sonraki servis/WMI istekleri de bu adla gidecegi icin dogru soru budur)
    $w = $null
    if (Test-TcpPort -Address $alive[0] -Port 135) { $w = Get-WmiComputerName -Target $ComputerName }
    if ($w -and $w.Name) {
        if ([string]::Equals($w.Name, $short, [System.StringComparison]::OrdinalIgnoreCase) -or
            [string]::Equals($w.Name, $nbExp, [System.StringComparison]::OrdinalIgnoreCase)) {
            $r.Online = $true; $r.Identity = 'Dogrulandi'; $r.ActualName = $w.Name; $r.OnlineText = 'Evet'
            return $r
        }
        $wrong = @([pscustomobject]@{ Ip = $alive[0]; Name = $w.Name })
    }
    elseif ($w -and $w.Kind -eq 'Identity' -and $wrong.Count -eq 0) {
        $hint = ''
        try { $hint = ([System.Net.Dns]::GetHostEntry($alive[0]).HostName -split '\.')[0] } catch { }
        if ($hint -and -not [string]::Equals($hint, $short, [System.StringComparison]::OrdinalIgnoreCase)) { $hint = " (PTR: $hint)" } else { $hint = '' }
        $wrong = @([pscustomobject]@{ Ip = $alive[0]; Name = "bilinmeyen cihaz$hint" })
    }

    if ($wrong.Count -gt 0) {
        $x = $wrong[0]
        $r.Identity   = 'Uyusmuyor'
        $r.ActualName = $x.Name
        $r.Ip         = $x.Ip
        $r.OnlineText = "Hayir (IP baska cihazda: $($x.Name))"
        $r.Note = "Ping cevabi BASKA cihazdan geliyor: $($x.Ip) su an '$($x.Name)' (eski DNS kaydi / VPN). Makine cevrimdisi kabul edildi."
        return $r
    }

    # Ad okunamadi (NetBIOS cevap yok + WMI ile ad okunamadi): ping cevabi BU makineden mi bilinmiyor.
    # Yanlis "cevrimici" (gateway / baska cihaz / joker DNS) sonucunu onlemek icin CEVRIMDISI kabul edilir.
    $r.Online = $false
    $r.Identity = 'Dogrulanamadi'
    $r.OnlineText = "Hayir (ping var, cihaz dogrulanamadi: $($alive[0]))"
    $why = if ($w) { $w.Text } else { 'TCP 135 (RPC) kapali' }
    $r.Note = "Ping cevabi var ($($alive[0])) ancak cevap verenin bu makine oldugu dogrulanamadi (NetBIOS cevap yok; WMI: $why). Cevrimdisi kabul edildi."
    return $r
}

function Read-FastCsv {
    # Hizli CSV okuyucu (Import-Csv yerine). Import-Csv her satir icin PSObject uretir ve 100.000 satirda saniyeler surer;
    # burada satirlar dogrudan string[] olarak okunur. Bicim: ';' ayracli, alanlar "..." icinde ("" = "), ilk satir baslik
    # (Export-Csv ve Write-FastCsv ciktisi). Tirnaksiz satirlar ';' ile bolunur. Doner: Header (string[]), Rows (List[string[]])
    param([string]$Path)
    $rows = New-Object 'System.Collections.Generic.List[string[]]'
    $header = $null
    $sep = [string[]]@('";"')
    foreach ($line in [System.IO.File]::ReadLines($Path, [System.Text.Encoding]::UTF8)) {
        if ($line.Length -eq 0) { continue }
        if ($line.Length -ge 2 -and $line[0] -eq '"' -and $line[$line.Length - 1] -eq '"') {
            $f = $line.Substring(1, $line.Length - 2).Split($sep, [System.StringSplitOptions]::None)
            for ($i = 0; $i -lt $f.Length; $i++) { if ($f[$i].IndexOf('""') -ge 0) { $f[$i] = $f[$i].Replace('""', '"') } }
        }
        else { $f = $line.Split(';') }
        if ($null -eq $header) { $header = $f } else { $rows.Add($f) }
    }
    return [pscustomobject]@{ Header = $header; Rows = $rows }
}

function Write-FastCsv {
    # Hizli CSV yazici (Export-Csv yerine): StreamWriter, UTF-8 (BOM), ';' ayrac, tum alanlar tirnakli -> Turkce Excel'de
    # dogrudan kolonlara ayrilir. Rows: string[] dizileri (Header sirasinda) VEYA nesneler (Header'daki ozellik adlariyla okunur).
    # string[] satirlarda tirnak yoksa satir tek [string]::Join ile kurulur (hucre hucre dongu yok).
    param([string]$Path, [string[]]$Header, $Rows)
    $w = New-Object System.IO.StreamWriter ($Path, $false, (New-Object System.Text.UTF8Encoding ($true)))
    try {
        $w.WriteLine('"' + [string]::Join('";"', $Header) + '"')
        $n = $Header.Count
        $vals = New-Object string[] $n
        foreach ($r in $Rows) {
            if ($r -is [string[]] -and $r.Length -eq $n) { $vals = $r.Clone() }
            else { for ($i = 0; $i -lt $n; $i++) { $vals[$i] = [string]$r.($Header[$i]) } }
            for ($i = 0; $i -lt $n; $i++) { if ($null -eq $vals[$i]) { $vals[$i] = '' } elseif ($vals[$i].IndexOf('"') -ge 0) { $vals[$i] = $vals[$i].Replace('"', '""') } }
            $w.WriteLine('"' + [string]::Join('";"', $vals) + '"')
        }
    }
    finally { $w.Dispose() }
}

function Get-AdCacheFileName {
    # OU basina onbellek dosya adi: <domain>_<OU ilk adi>_<SHA256 ilk 8>.csv (uzun DN'ler icin kisa, kararli ad)
    param([string]$Domain, [string]$OU)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $h = -join ($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes("$Domain|$OU".ToLowerInvariant()))[0..3] | ForEach-Object { $_.ToString('x2') }) }
    finally { $sha.Dispose() }
    $first = if ($OU -match '^(?:OU|CN)=([^,]+)') { $Matches[1] } else { 'OU' }
    return (("{0}_{1}_{2}.csv" -f $Domain, $first, $h) -replace '[\\/:*?"<>|\s]', '_')
}

function ConvertTo-LastLogon {
    # lastLogonTimestamp: ham FileTime (yeni onbellek) veya 'yyyy-MM-dd HH:mm:ss' (eski onbellek) -> DateTime / $null
    param([string]$Value)
    if (-not $Value) { return $null }
    $n = 0L
    if ([long]::TryParse($Value, [ref]$n)) { if ($n -gt 0 -and $n -lt [long]::MaxValue) { try { return [DateTime]::FromFileTime($n) } catch { } }; return $null }
    try { return [DateTime]::ParseExact($Value, 'yyyy-MM-dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture) } catch { return $null }
}

function ConvertTo-DomainFromDn {
    # CN=PC1,OU=X,DC=dom1,DC=local -> dom1.local
    param([string]$Dn)
    return ((@([regex]::Matches($Dn, '(?i)DC=([^,]+)') | ForEach-Object { $_.Groups[1].Value })) -join '.')
}

function Get-AdCacheIndex {
    # AD onbellegi: AdCache klasorundeki OU dosyalari (her OU icin bir CSV: Name;Enabled;LastLogonTimestamp;DistinguishedName,
    # Get-ADComputer ciktisi DONUSTURULMEDEN yazilir). Doner: kisa ad -> @(DistinguishedName, Enabled, LastLogonTimestamp ham).
    # Domain DN'den, LastLogon FileTime'dan SADECE bakilan makine icin cevrilir (ConvertTo-LastLogon) - yazarken satir isleme yok.
    # Harita AppDomain'de TEK kopya (10 runspace + UI paylasir); dosyalar degisince (ad + LastWriteTime) bir kez yeniden okunur.
    # Eski tek dosyali onbellek (AdComputerCache.csv) klasor bossa okunur.
    $r = [pscustomobject]@{ Ok = $false; Error = ''; Map = $null; Dup = $null; Count = 0; Unique = 0; Files = 0; Time = $null }

    $dir = [string]$env:CLIENTTOOL_ADCACHE_DIR
    $files = @()
    if ($dir -and [System.IO.Directory]::Exists($dir)) { $files = @([System.IO.Directory]::GetFiles($dir, '*.csv')) }
    if ($files.Count -eq 0 -and $env:CLIENTTOOL_ADCACHE_FILE -and [System.IO.File]::Exists($env:CLIENTTOOL_ADCACHE_FILE)) { $files = @($env:CLIENTTOOL_ADCACHE_FILE) }
    if ($files.Count -eq 0) { $r.Error = "AD onbellegi bulunamadi ($dir). 'AD Onbellegi...' > 'Kaydet ve Onbellegi Guncelle' ile olusturun."; return $r }

    $stamp = (@($files | Sort-Object | ForEach-Object { "$_|$([System.IO.File]::GetLastWriteTimeUtc($_).Ticks)" }) -join ';')
    $ad   = [System.AppDomain]::CurrentDomain
    $lock = $ad.GetData('ClientTool.AdCacheLock')
    if ($lock) { [System.Threading.Monitor]::Enter($lock) }
    try {
        $mem = $ad.GetData('ClientTool.AdCache')
        if (-not ($mem -and $mem.Stamp -eq $stamp)) {
            $map = New-Object System.Collections.Hashtable ([System.StringComparer]::OrdinalIgnoreCase)   # tr-TR I/i
            $dup = New-Object System.Collections.Hashtable ([System.StringComparer]::OrdinalIgnoreCase)
            $count = 0; $latest = [datetime]::MinValue; $bad = @()
            foreach ($f in $files) {
                $csv = $null
                for ($try = 0; $try -lt 3 -and $null -eq $csv; $try++) { try { $csv = Read-FastCsv -Path $f } catch { Start-Sleep -Milliseconds 300 } }
                if ($null -eq $csv) { $bad += $f; continue }
                $t = [System.IO.File]::GetLastWriteTime($f); if ($t -gt $latest) { $latest = $t }
                $ix = @{}
                for ($i = 0; $i -lt $csv.Header.Count; $i++) { $ix[$csv.Header[$i]] = $i }
                $iN = if ($ix.ContainsKey('Name')) { $ix['Name'] } else { $ix['Hostname'] }   # eski tek dosya: Hostname
                $iE = $ix['Enabled']; $iL = $ix['LastLogonTimestamp']; $iD = $ix['DistinguishedName']
                if ($null -eq $iN -or $null -eq $iE -or $null -eq $iL -or $null -eq $iD) { $bad += $f; continue }
                $need = [Math]::Max([Math]::Max($iN, $iE), [Math]::Max($iL, $iD))
                foreach ($x in $csv.Rows) {
                    if ($x.Length -le $need) { continue }
                    $k = $x[$iN]
                    if (-not $k) { continue }
                    $count++
                    $v = [string[]]@($x[$iD], $x[$iE], $x[$iL])
                    $prev = $map[$k]
                    if ($null -eq $prev) { $map[$k] = $v; continue }
                    if ($prev[0] -eq $v[0]) { continue }   # ic ice OU: ayni hesap iki kez
                    # Ayni isim birden fazla domain'de (nadir): en yeni lastLogon'lu kayit kullanilir
                    $dup[$k] = 1
                    $a = ConvertTo-LastLogon $v[2]; $b = ConvertTo-LastLogon $prev[2]
                    if ($a -and (-not $b -or $a -gt $b)) { $map[$k] = $v }
                }
            }
            if ($map.Count -eq 0 -and $bad.Count -gt 0) {
                if ($mem) { $r.Ok = $true; $r.Map = $mem.Map; $r.Dup = $mem.Dup; $r.Count = $mem.Count; $r.Unique = $mem.Map.Count; $r.Files = $mem.Files; $r.Time = $mem.Time; return $r }
                $r.Error = "AD onbellegi okunamadi: $($bad -join ', ')"; return $r
            }
            $mem = [pscustomobject]@{ Stamp = $stamp; Map = $map; Dup = $dup; Count = $count; Files = $files.Count; Time = $latest }
            $ad.SetData('ClientTool.AdCache', $mem)
        }
        $r.Ok = $true; $r.Map = $mem.Map; $r.Dup = $mem.Dup; $r.Count = $mem.Count; $r.Unique = $mem.Map.Count; $r.Files = $mem.Files; $r.Time = $mem.Time
        return $r
    }
    finally { if ($lock) { [System.Threading.Monitor]::Exit($lock) } }
}

function Get-AdOuComputers {
    # Bir domain'in OU'larini (alt OU'lar dahil) AD'den okur ve HER OU'yu ayri CSV'ye DOGRUDAN yazar:
    #   $data = Get-ADComputer ... ; $data | Select-Object Name, Enabled, LastLogonTimestamp, DistinguishedName | Export-Csv
    # Satir satir donusum YOK (LastLogon / domain okurken cevrilir). Dosya once .tmp'ye yazilir, basariliysa yerine tasinir;
    # hata veren OU'nun ESKI dosyasina dokunulmaz (eski veri kendiliginden korunur).
    # LIMIT YOK: -ResultSetSize $null; -ResultPageSize 1000 sadece sayfa boyutu.
    # Ilerleme: $Progress["$Domain [$OU]"] = @(sayi, durum, sn) - durum: okunuyor / yaziliyor / bitti / hata
    param([string]$Domain, [string[]]$SearchBases, $Progress, [string]$CacheDir)

    $items = New-Object System.Collections.ArrayList
    $modErr = ''
    if (-not (Get-Module -Name ActiveDirectory)) {
        try { Import-Module -Name ActiveDirectory -ErrorAction Stop -Verbose:$false }
        catch { $modErr = "ActiveDirectory modulu (RSAT) yuklenemedi: $($_.Exception.Message)" }
    }

    foreach ($sb in @($SearchBases)) {
        $key = "$Domain [$sb]"
        $it = [pscustomobject]@{ SearchBase = [string]$sb; Ok = $false; Error = ''; Count = 0; Seconds = 0 }
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            if ($modErr) { throw $modErr }
            if ($Progress) { $Progress[$key] = @(0, 'okunuyor', 0) }
            $data = @(Get-ADComputer -Filter * -SearchBase $sb -SearchScope Subtree -Server $Domain -ResultSetSize $null -ResultPageSize 1000 `
                    -Properties Enabled, LastLogonTimestamp -ErrorAction Stop)
            $it.Count = $data.Count
            if ($Progress) { $Progress[$key] = @($data.Count, 'yaziliyor', [int]$sw.Elapsed.TotalSeconds) }

            $file = Join-Path $CacheDir (Get-AdCacheFileName -Domain $Domain -OU $sb)
            $tmp  = "$file.tmp"
            if ($data.Count -gt 0) {
                $data | Select-Object Name, Enabled, LastLogonTimestamp, DistinguishedName |
                    Export-Csv -LiteralPath $tmp -NoTypeInformation -Delimiter ';' -Encoding UTF8 -ErrorAction Stop
            }
            else { Set-Content -LiteralPath $tmp -Value '"Name";"Enabled";"LastLogonTimestamp";"DistinguishedName"' -Encoding UTF8 -ErrorAction Stop }
            # Okuyan bir runspace dosyayi o an acik tutuyorsa tasima basarisiz olabilir: birkac kez dene
            for ($try = 1; ; $try++) {
                try { Move-Item -LiteralPath $tmp -Destination $file -Force -ErrorAction Stop; break }
                catch { if ($try -ge 10) { throw }; Start-Sleep -Milliseconds 300 }
            }
            $it.Ok = $true
            $it.Seconds = [int]$sw.Elapsed.TotalSeconds
            if ($Progress) { $Progress[$key] = @($data.Count, 'bitti', $it.Seconds) }
        }
        catch {
            $it.Error = "AD sorgusu / yazma basarisiz: $($_.Exception.Message)"
            $it.Seconds = [int]$sw.Elapsed.TotalSeconds
            if ($Progress) { $Progress[$key] = @($it.Count, 'hata', $it.Seconds) }
        }
        [void]$items.Add($it)
    }
    return [pscustomobject]@{ Domain = $Domain; Items = $items }
}

function Get-AdComputerInfo {
    # AD bilgisini AD'ye SORGU ATMADAN onbellekten okur. Status: Enabled / Disabled / Bulunamadi / Sorgulanamadi.
    # LastLogon: lastLogonTimestamp (~14 gun gecikmeli) - ham FileTime burada cevrilir.
    param([string]$ComputerName)

    $res = [pscustomobject]@{ Status = 'Bulunamadi'; LastLogon = 'N/A'; Detail = '' }
    $idx = Get-AdCacheIndex
    if (-not $idx.Ok) { $res.Status = 'Sorgulanamadi'; $res.Detail = $idx.Error; return $res }

    $short = (([string]$ComputerName -split '\.')[0]).Trim()
    $v = $idx.Map[$short]
    if ($null -eq $v) {
        $res.Detail = "AD onbellegindeki OU'larda bulunamadi (onbellek: $($idx.Time.ToString('yyyy-MM-dd HH:mm')), $($idx.Unique) makine)"
        return $res
    }
    if ($idx.Dup.ContainsKey($short)) { $res.Detail = "Birden fazla domain'de bulundu; en yeni lastLogon'a sahip kayit kullanildi ($(ConvertTo-DomainFromDn $v[0]))" }
    $res.Status = if ($v[1] -eq 'True') { 'Enabled' } else { 'Disabled' }
    $t = ConvertTo-LastLogon $v[2]
    $res.LastLogon = if ($t) { $t.ToString('yyyy-MM-dd HH:mm') } else { 'Hic' }
    return $res
}

function Get-ServiceCheck {
    # Istenen servislerin + CcmExec'in durumunu TEK RPC cagrisi (ServiceController.GetServices) ile alir.
    # Ad veya gorunen ad ile eslesir; karsilastirma kulturden bagimsizdir (tr-TR'de 'bits' = 'BITS').
    param([string]$ComputerName, [string[]]$ServiceNames = @())

    $r = [pscustomobject]@{ Running = ''; NotRunning = ''; CcmExec = ''; ServiceDetail = @(); Note = '' }
    $map = $null
    try {
        $all = [System.ServiceProcess.ServiceController]::GetServices($ComputerName)
        $map = New-Object System.Collections.Hashtable ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($s in $all) {
            $st = $s.Status.ToString()
            $map[$s.ServiceName] = $st
            if (-not $map.ContainsKey($s.DisplayName)) { $map[$s.DisplayName] = $st }
            $s.Dispose()
        }
    }
    catch { $r.Note = "Servis listesi alinamadi: $(Get-ErrText $_.Exception)" }

    $detail = @(foreach ($sn in $ServiceNames) {
        $n = "$sn".Trim()
        if (-not $n) { continue }
        $st = if ($null -eq $map) { 'Erisilemedi' } elseif ($map.ContainsKey($n)) { $map[$n] } else { 'NotFound' }
        [pscustomobject]@{ Name = $n; Status = $st }
    })
    $r.ServiceDetail = $detail
    $r.Running    = (($detail | Where-Object { $_.Status -eq 'Running' } | ForEach-Object { $_.Name }) -join ', ')
    $r.NotRunning = (($detail | Where-Object { $_.Status -ne 'Running' } | ForEach-Object { "$($_.Name)($($_.Status))" }) -join ', ')
    $r.CcmExec    = if ($null -eq $map) { 'Erisilemedi' } elseif ($map.ContainsKey('CcmExec')) { $map['CcmExec'] } else { 'Yok' }
    return $r
}

function Get-RemoteStatus {
    # "Kontrol Et": DNS + ping + kimlik + AD (onbellek) + servisler. Ulasilamazsa State = 'fail'.
    param([string]$ComputerName, [string[]]$ServiceNames)

    $reach = Get-HostReach -ComputerName $ComputerName
    $ad    = Get-AdComputerInfo -ComputerName $ComputerName
    $row = [pscustomobject]@{
        Hostname = $ComputerName; DnsStatus = $reach.DnsText; IsOnline = $reach.Online; OnlineText = $reach.OnlineText
        AdStatus = $ad.Status; LastLogon = $ad.LastLogon; Running = ''; NotRunning = ''; CcmExec = ''
        ServiceDetail = @(); ErrorMessage = ''; State = 'fail'
    }
    $errs = @($ad.Detail, $reach.Note | Where-Object { $_ })
    if ($reach.Online) {
        $svc = Get-ServiceCheck -ComputerName $ComputerName -ServiceNames $ServiceNames
        $row.Running = $svc.Running; $row.NotRunning = $svc.NotRunning; $row.CcmExec = $svc.CcmExec; $row.ServiceDetail = $svc.ServiceDetail
        if ($svc.Note) { $errs += $svc.Note }
        $row.State = if ($row.NotRunning -or $row.CcmExec -eq 'Yok' -or $reach.Identity -ne 'Dogrulandi') { 'warn' } else { 'ok' }
    }
    $row.ErrorMessage = ($errs -join ' | ')
    return $row
}

function Start-RemoteFix {
    param([string]$ComputerName, [pscustomobject[]]$ServiceDetail)   # ServiceDetail: kontrol asamasindaki ONCEKI durum

    $results = @()
    foreach ($s in $ServiceDetail) {
        $o = [pscustomobject]@{ Hostname = $ComputerName; ServiceName = $s.Name; OncekiDurum = $s.Status; SonrakiDurum = $s.Status; Aciklama = '' }
        if ($s.Status -eq 'Running') { $o.Aciklama = 'Zaten calisiyordu, islem yapilmadi.'; $results += $o; continue }
        if ($s.Status -eq 'NotFound' -or $s.Status -eq 'Erisilemedi') { $o.Aciklama = "Durum $($s.Status), islem yapilmadi."; $results += $o; continue }

        $sc = $null
        try {
            $sc = New-Object System.ServiceProcess.ServiceController ($s.Name, $ComputerName)
            if ($sc.Status -ne 'Running') {
                if ($sc.Status -ne 'StartPending') { $sc.Start() }
                try { $sc.WaitForStatus('Running', [TimeSpan]::FromSeconds(20)) } catch { }
                $sc.Refresh()
            }
            $o.SonrakiDurum = $sc.Status.ToString()
            $o.Aciklama = if ($o.SonrakiDurum -eq 'Running') { "$($s.Status) durumundaydi, Running yapildi." } else { "Baslatildi ama durum hala $($o.SonrakiDurum)." }
        }
        catch { $o.SonrakiDurum = 'Bilinmiyor'; $o.Aciklama = "Hata: $(Get-ErrText $_.Exception)" }
        finally { if ($sc) { $sc.Dispose() } }
        $results += $o
    }
    return $results
}

function Invoke-SccmClientAction {
    # Hedef makinede root\ccm SMS_Client.TriggerSchedule cagirir (DCOM/RPC).
    param([string]$ComputerName, [string]$ActionName, [string[]]$ScheduleIds)

    $r = [pscustomobject]@{ Hostname = $ComputerName; Action = $ActionName; State = 'fail'; Result = ''; Detail = '' }
    $reach = Get-HostReach -ComputerName $ComputerName
    if (-not $reach.DnsOk)  { $r.Result = "DNS: $($reach.DnsText) - eylem gonderilemedi."; return $r }
    if (-not $reach.Online) { $r.Result = "Cevrimdisi - eylem gonderilemedi. $($reach.Note)"; return $r }

    $suffix = if ($reach.Identity -eq 'Dogrulandi') { '' } else { ' (cihaz kimligi dogrulanamadi)' }
    $r.State = if ($reach.Identity -eq 'Dogrulandi') { 'ok' } else { 'warn' }

    # V11 - Istemci onarimi: SMS_Client.RepairClient (ccmrepair.exe'yi arka planda baslatir; WinRM gerekmez)
    if ($ScheduleIds[0] -eq '@REPAIR') {
        try {
            $null = Invoke-WmiMethod -ComputerName $ComputerName -Namespace 'root\ccm' -Class 'SMS_Client' -Name 'RepairClient' -ErrorAction Stop
            $r.Result = "Onarim baslatildi$suffix - ccmrepair arka planda calisir (birkac dakika surer, ajan bu sirada yeniden baslar)."
            $r.Detail = 'root\ccm SMS_Client.RepairClient cagrildi. Ilerleme: hedefte C:\Windows\CCM\Logs\ccmrepair.log'
        }
        catch {
            $r.State = 'fail'
            $r.Result = (Get-WmiFailureReason -ErrorRecord $_).Text
            $r.Detail = "RepairClient: $($_.Exception.Message)`r`nAjan WMI'si cevap vermiyorsa once 'CcmExec yeniden baslat' deneyin; ajan hic yoksa ccmsetup ile kurulum gerekir."
        }
        return $r
    }
    # V11 - CcmExec (SMS Agent Host) yeniden baslatma: ServiceController (RPC/SCM)
    if ($ScheduleIds[0] -eq '@RESTART') {
        $sc = $null
        try {
            $sc = New-Object System.ServiceProcess.ServiceController ('CcmExec', $ComputerName)
            $before = $sc.Status.ToString()
            if ($sc.Status -ne 'Stopped') {
                if ($sc.Status -ne 'StopPending') { $sc.Stop() }
                $sc.WaitForStatus('Stopped', [TimeSpan]::FromSeconds(90))
            }
            $sc.Start()
            try { $sc.WaitForStatus('Running', [TimeSpan]::FromSeconds(60)) } catch { }
            $sc.Refresh()
            $after = $sc.Status.ToString()
            if ($after -ne 'Running') { $r.State = 'warn' }
            $r.Result = "CcmExec: $before -> $after$suffix"
            $r.Detail = 'SMS Agent Host (CcmExec) durdurulup baslatildi.'
        }
        catch {
            $r.State = 'fail'
            $err = Get-ErrText $_.Exception
            $r.Result = if ($err -match 'was not found|bulunamad|1060') { 'SCCM ajani bulunmuyor (CcmExec servisi yok).' } else { "CcmExec yeniden baslatilamadi: $err" }
        }
        finally { if ($sc) { $sc.Dispose() } }
        return $r
    }

    $sent = @()
    foreach ($id in $ScheduleIds) {
        try {
            $null = Invoke-WmiMethod -ComputerName $ComputerName -Namespace 'root\ccm' -Class 'SMS_Client' -Name 'TriggerSchedule' -ArgumentList $id -ErrorAction Stop
            $sent += $id
        }
        catch {
            $r.State  = 'fail'
            $r.Result = (Get-WmiFailureReason -ErrorRecord $_).Text
            $r.Detail = "Basarisiz ScheduleID: $id`r`n$($_.Exception.Message)"
            if ($sent.Count -gt 0) { $r.Detail += "`r`nOncesinde gonderilenler: $($sent -join ', ')" }
            return $r
        }
    }
    $r.Result = "Gonderildi$suffix"
    $r.Detail = "Tetiklenen ScheduleID: $($sent -join ', ')"
    return $r
}

function Get-NoResponseVerdict {
    # Script'ten cevap gelmediginde sebebi belirler: cevrimdisi / IP baska cihazda / SCCM ajani yok / cevrimici ama calismadi
    param([string]$ComputerName, [int]$TimeoutSec)

    $reach = Get-HostReach -ComputerName $ComputerName
    if (-not $reach.DnsOk)  { return [pscustomobject]@{ State = 'fail'; Text = "Cevrimdisi oldugu icin script calismadi (DNS: $($reach.DnsText))." } }
    if (-not $reach.Online) { return [pscustomobject]@{ State = 'fail'; Text = "Cevrimdisi oldugu icin script calismadi. $($reach.Note)" } }

    $cc = (Get-ServiceCheck -ComputerName $ComputerName).CcmExec
    $pre = if ($reach.Identity -eq 'Dogrulandi') { 'Cevrimici' } else { 'Ping var (kimlik dogrulanamadi)' }
    if ($cc -eq 'Yok')         { return [pscustomobject]@{ State = 'fail'; Text = 'SCCM ajani bulunmuyor (CcmExec servisi yok).' } }
    if ($cc -eq 'Erisilemedi') { return [pscustomobject]@{ State = 'warn'; Text = "$pre - script calismadi (CcmExec sorgulanamadi)." } }
    if ($cc -ne 'Running')     { return [pscustomobject]@{ State = 'warn'; Text = "$pre - script calismadi (CcmExec servisi $cc)." } }
    return [pscustomobject]@{ State = 'warn'; Text = "$pre - script calismadi ($TimeoutSec sn icinde SCCM'den yanit gelmedi)." }
}

function Get-SccmDeviceMap {
    # V2 - TOPLU SCCM KAYDI: SMS_R_System cihaz basina degil, $Chunk'lik parcalar halinde TEK sorguyla okunur
    # (WQL'de IN yok: Name = 'a' OR Name = 'b' ...). Doner: kisa ad -> Get-SccmDevice ile ayni yapida kayit
    # + BoundaryGroups / OsBuild: SMS_CombinedDeviceResources (konsoldaki "Boundary Group(s)" ve "Operating System Build"
    # kolonlari) - ayni parca icin TEK ek sorgu. BoundaryGroups ConfigMgr 2002+; eski surumde sadece DeviceOSBuild okunur.
    param([string[]]$Names, [string]$SmsProvider, [string]$SiteCode, [int]$Chunk = 100)

    $ns  = "root\SMS\site_$SiteCode"
    $map = New-Object System.Collections.Hashtable ([System.StringComparer]::OrdinalIgnoreCase)
    $shorts = @($Names | ForEach-Object { (([string]$_ -split '\.')[0]).Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    $cdrCols = 'ResourceID, Name, BoundaryGroups, DeviceOSBuild'   # ilk hatada 'ResourceID, Name, DeviceOSBuild'e duser
    for ($i = 0; $i -lt $shorts.Count; $i += $Chunk) {
        $part = @($shorts[$i..([Math]::Min($i + $Chunk, $shorts.Count) - 1)])
        $w = (($part | ForEach-Object { "Name = '$(ConvertTo-WqlString $_)'" }) -join ' OR ')
        $byName = New-Object System.Collections.Hashtable ([System.StringComparer]::OrdinalIgnoreCase)   # ad -> aday listesi
        $err = ''
        try {
            foreach ($o in @(Get-WmiObject -ComputerName $SmsProvider -Namespace $ns -ErrorAction Stop `
                        -Query "SELECT ResourceId, Name, Client, Active, Obsolete FROM SMS_R_System WHERE $w")) {
                $l = $byName[[string]$o.Name]
                if ($null -eq $l) { $l = New-Object System.Collections.ArrayList; $byName[[string]$o.Name] = $l }
                [void]$l.Add([pscustomobject]@{ ResourceId = $o.ResourceId; Name = [string]$o.Name; Client = $o.Client; Active = $o.Active; Obsolete = $o.Obsolete })
            }
        }
        catch { $err = $_.Exception.Message }

        # Boundary Group(s) + Operating System Build: ResourceID -> @(BoundaryGroups, DeviceOSBuild)
        $cdr = @{}
        $cdrErr = ''
        if (-not $err) {
            while ($true) {
                try {
                    foreach ($o in @(Get-WmiObject -ComputerName $SmsProvider -Namespace $ns -ErrorAction Stop -Query "SELECT $cdrCols FROM SMS_CombinedDeviceResources WHERE $w")) {
                        $bg = if ($o.PSObject.Properties['BoundaryGroups']) { [string]$o.BoundaryGroups } else { '(bu SCCM surumunde yok)' }
                        $cdr[[string]$o.ResourceID] = [string[]]@($bg, [string]$o.DeviceOSBuild)
                    }
                    break
                }
                catch {
                    if ($cdrCols -like '*BoundaryGroups*') { $cdrCols = 'ResourceID, Name, DeviceOSBuild'; continue }   # ConfigMgr 2002 oncesi
                    $cdrErr = $_.Exception.Message
                    break
                }
            }
        }

        foreach ($n in $part) {
            if ($err) { $map[$n] = [pscustomobject]@{ Ok = $false; Error = $err; Cands = @(); Dev = $null; BoundaryGroups = '(sorgulanamadi)'; OsBuild = '(sorgulanamadi)' }; continue }
            $c   = @(if ($byName.ContainsKey($n)) { $byName[$n] })
            $dev = $c | Where-Object { $_.Client -eq 1 -and $_.Obsolete -ne 1 } | Select-Object -First 1
            $bg = ''; $os = ''
            if ($c.Count -eq 0) { $bg = $os = '(SCCM kaydi yok)' }
            elseif ($cdrErr) { $bg = $os = '(sorgulanamadi)' }
            else {
                $ref = if ($dev) { $dev } else { $c[0] }
                $v = $cdr[[string]$ref.ResourceId]
                if ($v) { $bg = $v[0]; $os = $v[1] }
            }
            $map[$n] = [pscustomobject]@{ Ok = $true; Error = ''; Cands = $c; Dev = $dev; BoundaryGroups = $bg; OsBuild = $os }
        }
    }
    return $map
}

function Get-SccmDevice {
    # SMS_R_System'de cihaz kaydi. Doner: Ok, Error, Cands (tum adaylar), Dev (istemcisi olan aktif kayit)
    # -Info verilmisse (toplu sorgudan gelen kayit, Get-SccmDeviceMap) SMS Provider'a HIC gidilmez.
    param([string]$ComputerName, [string]$SmsProvider, [string]$SiteCode, $Info)
    if ($Info) { return $Info }
    $short = ConvertTo-WqlString (($ComputerName -split '\.')[0])
    try {
        $c = @(Get-WmiObject -ComputerName $SmsProvider -Namespace "root\SMS\site_$SiteCode" -ErrorAction Stop `
                -Query "SELECT ResourceId, Name, Client, Active, Obsolete FROM SMS_R_System WHERE Name = '$short'")
        return [pscustomobject]@{ Ok = $true; Error = ''; Cands = $c; Dev = ($c | Where-Object { $_.Client -eq 1 -and $_.Obsolete -ne 1 } | Select-Object -First 1) }
    }
    catch { return [pscustomobject]@{ Ok = $false; Error = $_.Exception.Message; Cands = @(); Dev = $null } }
}

function Send-SccmScript {
    # SMS Provider uzerinden "Run Script" (ClientOperation Type 135) BASLATIR; sonucu BEKLEMEZ.
    # Sonuc, ana penceredeki TOPLU yoklayicidan (Get-ScriptPollStatus, 4 sn'de bir tum bekleyenler icin tek sorgu) gelir.
    # Pending = $true ise gonderim basarili, ResourceId / OperationId / SentAt yoklama icin doludur.
    param([string]$ComputerName, [string]$SmsProvider, [string]$SiteCode, [string]$ScriptGuid, [string]$ScriptName, [string]$ParamB64, [int]$TimeoutSec, $SccmInfo)

    $r = [pscustomobject]@{
        Action = "Script: $ScriptName"; State = 'fail'; Result = ''; Output = ''; ExitCode = $null; Detail = ''
        Pending = $false; ResourceId = $null; OperationId = $null; SentAt = $null
        SmsProvider = $SmsProvider; SiteCode = $SiteCode; ScriptGuid = $ScriptGuid; TimeoutSec = $TimeoutSec
    }
    $ns = "root\SMS\site_$SiteCode"

    # ---- 1) Cihaz SCCM'de kayitli mi, istemci (ajan) var mi? ----
    $q = Get-SccmDevice -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode -Info $SccmInfo
    if (-not $q.Ok) { $r.Result = "SCCM sorgusu basarisiz: $($q.Error)"; return $r }
    if (-not $q.Dev) {
        $r.Result = if ($q.Cands.Count -gt 0) { 'SCCM ajani bulunmuyor (cihaz kaydi var ama istemci yuklu degil/obsolete).' } else { "SCCM ajani bulunmuyor (cihaz SCCM'de kayitli degil)." }
        return $r
    }
    $rid = [uint32]$q.Dev.ResourceId

    # ---- 2) Scripti gonder ----
    try {
        $mc = [WmiClass]"\\$SmsProvider\${ns}:SMS_ClientOperation"
        $in = $mc.psbase.GetMethodParameters('InitiateClientOperationEx')
        $in.Param = $ParamB64
        $in.RandomizationWindow = [uint32]0
        $in.TargetCollectionID = ''
        $in.TargetResourceIDs = [uint32[]]@($rid)
        $in.Type = [uint32]135
        $res = $mc.InvokeMethod('InitiateClientOperationEx', $in, $null)
        if ($res.ReturnValue -ne 0) { $r.Result = "Script gonderilemedi (ReturnValue=$($res.ReturnValue))."; return $r }
        if ($res.OperationID) { $r.OperationId = [uint32]$res.OperationID }
    }
    catch {
        $why = Get-WmiFailureReason -ErrorRecord $_
        $r.Result = if ($why.Kind -eq 'Access') { 'Script gonderilemedi: SCCM yetkisi yok (Run Script izni gerekli).' } else { "Script gonderilemedi: $($_.Exception.Message)" }
        return $r
    }
    $r.Pending = $true; $r.ResourceId = $rid; $r.SentAt = Get-Date
    $r.State = 'pending'; $r.Result = 'Gonderildi, sonuc bekleniyor'
    return $r
}

function Select-ScriptStatusRow {
    # SMS_ScriptsExecutionStatus satirlarindan BU calistirmaya ait olani secer (eski calistirmalarin sonucu alinmaz).
    # Eslesme sirasi: TaskID -> ClientOperationId -> LastUpdateTime (sema surumler arasinda farkli olabildigi icin).
    param([object[]]$Rows, $TaskId, $Op, [datetime]$Start)
    $seen = $false
    foreach ($row in $Rows) {
        if (-not $row) { continue }
        $names = @($row.PSObject.Properties | ForEach-Object { $_.Name })
        $mine = $false
        if ($TaskId -and ($names -contains 'TaskID') -and $row.TaskID) { $mine = ([string]$row.TaskID -eq [string]$TaskId) }
        elseif (($names -contains 'ClientOperationId') -and $row.ClientOperationId -and $Op) { $mine = ([uint32]$row.ClientOperationId -eq [uint32]$Op) }
        elseif (($names -contains 'LastUpdateTime') -and $row.LastUpdateTime) {
            try { $mine = ([System.Management.ManagementDateTimeConverter]::ToDateTime([string]$row.LastUpdateTime) -ge $Start.AddMinutes(-1)) } catch { }
        }
        if (-not $mine) { continue }
        $seen = $true
        if ($null -ne $row.ScriptExitCode -or -not [string]::IsNullOrEmpty([string]$row.ScriptOutput)) { return [pscustomobject]@{ Seen = $true; Status = $row } }
    }
    return [pscustomobject]@{ Seen = $seen; Status = $null }
}

function Get-ScriptPollStatus {
    # TOPLU YOKLAMA: bekleyen TUM cihazlar icin SMS Provider'a cihaz basina degil, grup basina TEK sorgu atar
    # (WQL'de IN yok: 'ResourceID = a OR ResourceID = b ...'; sorgu boyu icin $Chunk'lik parcalar).
    # Items: Id, Provider, SiteCode, Guid, Rid, Op, TaskId, Start. Doner: Id, TaskId, Seen, Done, ExitCode, Output, Error
    param([object[]]$Items, [int]$Chunk = 100)

    $out = New-Object System.Collections.ArrayList
    foreach ($g in @($Items | Group-Object -Property { "$($_.Provider)|$($_.SiteCode)|$($_.Guid)" })) {
        $f   = $g.Group[0]
        $ns  = "root\SMS\site_$($f.SiteCode)"
        $grp = @($g.Group)
        for ($i = 0; $i -lt $grp.Count; $i += $Chunk) {
            $part = @($grp[$i..([Math]::Min($i + $Chunk, $grp.Count) - 1)])

            # TaskID'si henuz bilinmeyenler icin tek sorgu
            $taskMap = @{}
            $noTask = @($part | Where-Object { -not $_.TaskId -and $_.Op })
            if ($noTask.Count -gt 0) {
                $w = (($noTask | ForEach-Object { "ClientOperationId = $([uint32]$_.Op)" }) -join ' OR ')
                try {
                    foreach ($t in @(Get-WmiObject -ComputerName $f.Provider -Namespace $ns -ErrorAction Stop -Query "SELECT TaskID, ClientOperationId FROM SMS_ScriptsExecutionTask WHERE $w")) {
                        $taskMap[[string]$t.ClientOperationId] = [string]$t.TaskID
                    }
                }
                catch { }
            }

            # Sonuc satirlari: tek sorgu
            $byRid = @{}
            $err = ''
            $w = (($part | ForEach-Object { "ResourceID = $([uint32]$_.Rid)" }) -join ' OR ')
            try {
                foreach ($row in @(Get-WmiObject -ComputerName $f.Provider -Namespace $ns -ErrorAction Stop `
                            -Query "SELECT * FROM SMS_ScriptsExecutionStatus WHERE ScriptGuid = '$(ConvertTo-WqlString $f.Guid)' AND ($w)")) {
                    $k = [string]$row.ResourceID
                    if (-not $byRid.ContainsKey($k)) { $byRid[$k] = New-Object System.Collections.ArrayList }
                    [void]$byRid[$k].Add($row)
                }
            }
            catch { $err = $_.Exception.Message }

            foreach ($it in $part) {
                $tid = if ($it.TaskId) { $it.TaskId } elseif ($it.Op -and $taskMap.ContainsKey([string]$it.Op)) { $taskMap[[string]$it.Op] } else { $null }
                $m = Select-ScriptStatusRow -Rows @($byRid[[string]$it.Rid]) -TaskId $tid -Op $it.Op -Start $it.Start
                [void]$out.Add([pscustomobject]@{
                    Id = $it.Id; TaskId = $tid; Seen = [bool]$m.Seen; Done = [bool]$m.Status; Error = $err
                    ExitCode = $(if ($m.Status) { $m.Status.ScriptExitCode } else { $null })
                    Output   = $(if ($m.Status) { [string]$m.Status.ScriptOutput } else { '' })
                })
            }
        }
    }
    return $out
}

function Get-AdStructureStatus {
    # AD YAPISI KONTROLU: DNS + Cevrimici (ping + kimlik) + AD (var mi, Enabled/Disabled, Last Logon). Servis kontrolu YAPMAZ.
    param([string]$ComputerName)

    $reach = Get-HostReach -ComputerName $ComputerName
    $ad    = Get-AdComputerInfo -ComputerName $ComputerName
    $state = if (-not $reach.Online) { 'fail' } elseif ($ad.Status -ne 'Enabled' -or $reach.Identity -ne 'Dogrulandi') { 'warn' } else { 'ok' }
    return [pscustomobject]@{
        Hostname = $ComputerName; DnsStatus = $reach.DnsText; IsOnline = $reach.Online; OnlineText = $reach.OnlineText
        Identity = $reach.Identity; AdStatus = $ad.Status; LastLogon = $ad.LastLogon
        ErrorMessage = (@($ad.Detail, $reach.Note | Where-Object { $_ }) -join ' | '); State = $state; ReachNote = $reach.Note
    }
}

function New-SkippedScript {
    # Cevrimdisi cihaza SCCM scripti GONDERILMEZ (sonucu beklemek de anlamsiz); Send-SccmScript ile ayni yapida sonuc doner.
    param([string]$ScriptName, [string]$Why)
    return [pscustomobject]@{
        Action = "Script: $ScriptName"; State = 'fail'; Result = "Script gonderilmedi - cihaz cevrimdisi ($Why)"; Output = ''; ExitCode = $null
        Detail = "Cevrimdisi oldugu icin SCCM'e gonderim yapilmadi. $Why"; Pending = $false; ResourceId = $null; OperationId = $null; SentAt = $null
    }
}

function Invoke-ScriptWithAdCheck {
    # PANEL 1: once AD yapisi kontrolu (DNS/ping/kimlik/AD) + servis kontrolu, sonra SCCM script GONDERIMI (sonuc toplu yoklamadan).
    param([string]$ComputerName, [string]$SmsProvider, [string]$SiteCode, [string]$ScriptGuid, [string]$ScriptName, [string]$ParamB64, [int]$TimeoutSec, [string[]]$ServiceNames = @(), $SccmInfo)

    $ad = Get-AdStructureStatus -ComputerName $ComputerName
    if ($ad.IsOnline -and @($ServiceNames).Count -gt 0) { $svc = Get-ServiceCheck -ComputerName $ComputerName -ServiceNames $ServiceNames }
    else { $svc = [pscustomobject]@{ Running = ''; NotRunning = ''; CcmExec = $(if (-not $ad.IsOnline) { 'Kontrol edilemedi (cevrimdisi)' } else { '' }); ServiceDetail = @(); Note = '' } }

    if ($ad.IsOnline) {
        $sr = Send-SccmScript -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode `
            -ScriptGuid $ScriptGuid -ScriptName $ScriptName -ParamB64 $ParamB64 -TimeoutSec $TimeoutSec -SccmInfo $SccmInfo
    }
    else { $sr = New-SkippedScript -ScriptName $ScriptName -Why $(if ($ad.ReachNote) { $ad.ReachNote } else { $ad.OnlineText }) }

    return [pscustomobject]@{
        Hostname = $ComputerName; DnsStatus = $ad.DnsStatus; IsOnline = $ad.IsOnline; OnlineText = $ad.OnlineText
        AdStatus = $ad.AdStatus; LastLogon = $ad.LastLogon; AdNote = $ad.ErrorMessage
        Running = $svc.Running; NotRunning = $svc.NotRunning; CcmExec = $svc.CcmExec; ServiceDetail = $svc.ServiceDetail; SvcNote = $svc.Note
        Action = $sr.Action; State = $sr.State; Result = $sr.Result; Output = $sr.Output; ExitCode = $sr.ExitCode; Detail = $sr.Detail
        Pending = $sr.Pending; ResourceId = $sr.ResourceId; OperationId = $sr.OperationId; SentAt = $sr.SentAt
        SmsProvider = $SmsProvider; SiteCode = $SiteCode; ScriptGuid = $ScriptGuid; TimeoutSec = $TimeoutSec
    }
}

function Get-SccmRegistration {
    # SCCM sitesinde cihaz kaydi + istemci (ajan) durumu (SMS Provider WMI sorgusu).
    param([string]$ComputerName, [string]$SmsProvider, [string]$SiteCode, $SccmInfo)

    if ([string]::IsNullOrWhiteSpace($SmsProvider) -or [string]::IsNullOrWhiteSpace($SiteCode)) { return 'Sorgulanmadi (SCCM baglantisi yok)' }
    $q = Get-SccmDevice -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode -Info $SccmInfo
    if (-not $q.Ok) { return "Sorgulanamadi: $($q.Error)" }
    if ($q.Cands.Count -eq 0) { return 'SCCM de kayitli degil' }
    if ($q.Dev) { if ($q.Dev.Active -eq 1) { return 'Client yuklu (Active)' } else { return 'Client yuklu (Inactive)' } }
    return 'Kayit var, client yok / obsolete'
}

function Get-ParserHostCheck {
    # PANEL 2 icin kontrol: DNS + ping + kimlik + AD + SERVISLER (1. bolumdeki liste) + CcmExec + SCCM kaydi.
    # Panel 1'in sonuc tablosuna / detaylarina DOKUNMAZ; sadece nesne dondurur.
    param([string]$ComputerName, [string]$SmsProvider = '', [string]$SiteCode = '', [string[]]$ServiceNames = @(), $SccmInfo)

    $reach = Get-HostReach -ComputerName $ComputerName
    $ad    = Get-AdComputerInfo -ComputerName $ComputerName
    $c = [pscustomobject]@{
        Hostname = $ComputerName; DnsStatus = $reach.DnsText; IsOnline = $reach.Online; OnlineText = $reach.OnlineText
        AdStatus = $ad.Status; LastLogon = $ad.LastLogon; Running = ''; NotRunning = ''; CcmExec = ''; ServiceDetail = @()
        SccmKayit = (Get-SccmRegistration -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode -SccmInfo $SccmInfo); Note = ''
    }
    $notes = @($ad.Detail, $reach.Note | Where-Object { $_ })
    $c | Add-Member -NotePropertyName ReachNote -NotePropertyValue $reach.Note

    if (-not $reach.DnsOk) { $c.CcmExec = $c.Running = $c.NotRunning = 'Kontrol edilemedi (DNS)' }
    elseif (-not $reach.Online) { $c.CcmExec = $c.Running = $c.NotRunning = 'Kontrol edilemedi (cevrimdisi)' }
    else {
        $svc = Get-ServiceCheck -ComputerName $ComputerName -ServiceNames $ServiceNames
        $c.CcmExec = $svc.CcmExec; $c.Running = $svc.Running; $c.NotRunning = $svc.NotRunning; $c.ServiceDetail = $svc.ServiceDetail
        if ($svc.Note) { $notes += $svc.Note }
    }
    $c.Note = ($notes -join ' | ')
    return $c
}

function Invoke-ParserScriptRun {
    # PANEL 2: once DNS/AD/servis/SCCM ajan kontrolu, sonra SCCM script GONDERIMI (sonuc toplu yoklamadan).
    param([string]$ComputerName, [string]$SmsProvider, [string]$SiteCode, [string]$ScriptGuid, [string]$ScriptName, [string]$ParamB64, [int]$TimeoutSec, [string[]]$ServiceNames = @(), $SccmInfo)

    $chk = Get-ParserHostCheck -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode -ServiceNames $ServiceNames -SccmInfo $SccmInfo
    if ($chk.IsOnline) {
        $sr = Send-SccmScript -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode `
            -ScriptGuid $ScriptGuid -ScriptName $ScriptName -ParamB64 $ParamB64 -TimeoutSec $TimeoutSec -SccmInfo $SccmInfo
    }
    else { $sr = New-SkippedScript -ScriptName $ScriptName -Why $(if ($chk.ReachNote) { $chk.ReachNote } else { $chk.OnlineText }) }
    return [pscustomobject]@{
        Hostname = $ComputerName; Check = $chk; Action = $sr.Action; State = $sr.State; Result = $sr.Result; Output = $sr.Output; ExitCode = $sr.ExitCode; Detail = $sr.Detail
        Pending = $sr.Pending; ResourceId = $sr.ResourceId; OperationId = $sr.OperationId; SentAt = $sr.SentAt
        SmsProvider = $SmsProvider; SiteCode = $SiteCode; ScriptGuid = $ScriptGuid; TimeoutSec = $TimeoutSec
    }
}

# ==================== SADECE ANA PENCERE FONKSIYONLARI ====================

function New-WorkerScript {
    # Ortak fonksiyonlarin kaynak kodunu worker scriptinin basina gomer; tanimlar TEK yerde tutulur.
    param([string]$ParamLine, [string]$CallLine, [string[]]$Funcs)
    $all = @('ConvertTo-WqlString', 'Get-ErrText', 'Test-Ping', 'Test-TcpPort', 'Get-NetBiosName', 'Get-WmiFailureReason', 'Get-WmiComputerName',
        'Get-DnsWildcardIps', 'Resolve-HostDns', 'Get-HostReach', 'Read-FastCsv', 'Get-AdCacheFileName', 'ConvertTo-LastLogon', 'ConvertTo-DomainFromDn', 'Get-AdCacheIndex', 'Get-AdOuComputers', 'Get-AdComputerInfo', 'Get-ServiceCheck',
        'Get-RemoteStatus', 'Start-RemoteFix', 'Invoke-SccmClientAction', 'Get-NoResponseVerdict', 'Get-SccmDevice', 'Send-SccmScript',
        'Get-AdStructureStatus', 'New-SkippedScript', 'Invoke-ScriptWithAdCheck', 'Get-SccmRegistration', 'Get-ParserHostCheck', 'Invoke-ParserScriptRun')
    $defs = foreach ($n in $(if ($Funcs) { $Funcs } else { $all })) { "function $n {`n$((Get-Item -Path "function:$n").ScriptBlock.ToString())`n}" }
    return [scriptblock]::Create("param($ParamLine)`n" + ($defs -join "`n") + "`n$CallLine")
}

$script:CheckWorker        = New-WorkerScript '$ComputerName, $ServiceNames' 'Get-RemoteStatus -ComputerName $ComputerName -ServiceNames $ServiceNames'
$script:FixWorker          = New-WorkerScript '$ComputerName, $ServiceDetail' 'Start-RemoteFix -ComputerName $ComputerName -ServiceDetail $ServiceDetail'
$script:ActionWorker       = New-WorkerScript '$ComputerName, $ActionName, $ScheduleIds' 'Invoke-SccmClientAction -ComputerName $ComputerName -ActionName $ActionName -ScheduleIds $ScheduleIds'
$script:ScriptWorker       = New-WorkerScript '$ComputerName, $SmsProvider, $SiteCode, $ScriptGuid, $ScriptName, $ParamB64, $TimeoutSec, $ServiceNames, $SccmInfo' 'Invoke-ScriptWithAdCheck -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode -ScriptGuid $ScriptGuid -ScriptName $ScriptName -ParamB64 $ParamB64 -TimeoutSec $TimeoutSec -ServiceNames $ServiceNames -SccmInfo $SccmInfo'
$script:AdWorker           = New-WorkerScript '$ComputerName' 'Get-AdStructureStatus -ComputerName $ComputerName'
$script:AdCacheWorker      = New-WorkerScript '$Domain, $SearchBases, $Progress, $CacheDir' 'Get-AdOuComputers -Domain $Domain -SearchBases $SearchBases -Progress $Progress -CacheDir $CacheDir' @('Get-AdCacheFileName', 'Get-AdOuComputers')
$script:ParserCheckWorker  = New-WorkerScript '$ComputerName, $SmsProvider, $SiteCode, $ServiceNames, $SccmInfo' 'Get-ParserHostCheck -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode -ServiceNames $ServiceNames -SccmInfo $SccmInfo'
$script:VerdictWorker      = New-WorkerScript '$ComputerName, $TimeoutSec' 'Get-NoResponseVerdict -ComputerName $ComputerName -TimeoutSec $TimeoutSec'
$script:PollWorker         = New-WorkerScript '$Items' 'Get-ScriptPollStatus -Items $Items' @('ConvertTo-WqlString', 'Select-ScriptStatusRow', 'Get-ScriptPollStatus')
$script:ParserScriptWorker = New-WorkerScript '$ComputerName, $SmsProvider, $SiteCode, $ScriptGuid, $ScriptName, $ParamB64, $TimeoutSec, $ServiceNames, $SccmInfo' 'Invoke-ParserScriptRun -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode -ScriptGuid $ScriptGuid -ScriptName $ScriptName -ParamB64 $ParamB64 -TimeoutSec $TimeoutSec -ServiceNames $ServiceNames -SccmInfo $SccmInfo'
$script:SccmMapWorker      = New-WorkerScript '$Names, $SmsProvider, $SiteCode' 'Get-SccmDeviceMap -Names $Names -SmsProvider $SmsProvider -SiteCode $SiteCode' @('ConvertTo-WqlString', 'Get-SccmDeviceMap')

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
    # Bilerek %USERPROFILE%\Desktop (GetFolderPath degil): OneDrive klasor yonlendirmesinden etkilenmesin.
    $dir = Join-Path $env:USERPROFILE 'Desktop\ServiceStatus'
    if (-not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
    return $dir
}

function Get-SettingsFile { return (Join-Path (Get-DataDir) 'ClientToolSettings.json') }
function Get-AdCacheFile  { return (Join-Path (Get-DataDir) 'AdComputerCache.csv') }   # ESKI tek dosyali onbellek (sadece gecis icin okunur)
function Get-AdCacheDir {
    $d = Join-Path (Get-DataDir) 'AdCache'   # OU basina bir CSV
    if (-not (Test-Path -LiteralPath $d)) { $null = New-Item -ItemType Directory -Path $d -Force }
    return $d
}

# Paralel runspace'ler script degiskenlerini goremez; onbellek yolunu process geneli ortam degiskeniyle paylasiriz
# (Get-AdCacheIndex bu degiskeni okur).
$env:CLIENTTOOL_ADCACHE_FILE = Get-AdCacheFile
$env:CLIENTTOOL_ADCACHE_DIR  = Get-AdCacheDir
# AD onbellek haritasi tum runspace'lerde TEK kopya (AppDomain); ayni anda tek runspace yukler
if (-not [System.AppDomain]::CurrentDomain.GetData('ClientTool.AdCacheLock')) { [System.AppDomain]::CurrentDomain.SetData('ClientTool.AdCacheLock', (New-Object System.Object)) }

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
$script:ResultsTable.Locale = [System.Globalization.CultureInfo]::InvariantCulture   # tr-TR'de I/i eslesmesi icin
foreach ($c in @('Hostname', 'DnsStatus', 'IsOnline', 'AdStatus', 'LastLogon', 'Running', 'NotRunning', 'CcmExec', 'BoundaryGroups', 'OsBuild', 'LastAction', 'ActionResult', 'ErrorMessage', 'State')) {
    $null = $script:ResultsTable.Columns.Add($c)
}
$script:ResultsTable.PrimaryKey = @($script:ResultsTable.Columns['Hostname'])   # Rows.Find ile O(1) arama (CaseSensitive = false)
$script:BindingSrc = New-Object System.Windows.Forms.BindingSource
$script:BindingSrc.DataSource = $script:ResultsTable

# Panel 2 (Update Parser) tablosu: 21 parser kolonu + 6 kontrol kolonu (DNS/Online/AD/LastLogon/CcmExec/SCCM) + KolonUyumu + SourceHost (gizli)
$script:ParseTable = New-Object System.Data.DataTable
$script:ParseTable.Locale = [System.Globalization.CultureInfo]::InvariantCulture
foreach ($c in $script:ParserColumns) { $null = $script:ParseTable.Columns.Add($c) }
foreach ($c in $script:ParserCheckColumns) { $null = $script:ParseTable.Columns.Add($c) }
foreach ($c in $script:ParserSccmColumns) { $null = $script:ParseTable.Columns.Add($c) }
$null = $script:ParseTable.Columns.Add('KolonUyumu')
$null = $script:ParseTable.Columns.Add('SourceHost')
$null = $script:ParseTable.Columns.Add('Durum')   # V7: hesaplanan satir durumu (gizli; ozet cubugu / filtre / tekrar dene)
$script:ParseBinding = New-Object System.Windows.Forms.BindingSource
$script:ParseBinding.DataSource = $script:ParseTable

function Get-GridRow {
    param([string]$Hostname)
    $found = $script:ResultsTable.Rows.Find($Hostname)
    if ($found) { return $found }
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

function Format-JobError {
    param([string]$ErrorText)
    if ($ErrorText -eq $script:CancelText) { return 'Iptal edildi' }
    return "Islem hatasi: $ErrorText"
}

function Resolve-PendingCells {
    # Islem hata verdigi / iptal edildigi icin hala 'Bekleniyor..' kalan hucreleri doldurur.
    param([string]$Hostname, [string]$ErrorText = '')
    $Text = if ($ErrorText -eq $script:CancelText) { 'Iptal edildi' } else { 'Alinamadi' }
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
    return (@($script:ActiveRuns | Where-Object { $_.Panel -eq $Panel }).Count -gt 0 -or @($script:Prefetch | Where-Object { $_.Panel -eq $Panel }).Count -gt 0)
}

function Start-ParallelRun {
    param(
        [string]$Title,
        [scriptblock]$Worker,
        [object[]]$Items,          # her biri: @{ Key = hostname; Params = @{...} }
        [string]$OnResultFn,
        [int]$Panel = 1,           # 1 = Panel 1 (kontrol/script), 2 = Panel 2 (update parser)
        [string]$Unit = 'makine'   # durum cubugundaki "N makine islendi" metni icin
    )

    $st = [pscustomobject]@{
        Panel      = $Panel
        Pool       = (Get-SharedPool)
        Jobs       = $null
        Total      = @($Items).Count
        Done       = 0
        Title      = $Title
        OnResultFn = $OnResultFn
        Unit       = $Unit
    }
    [void]$script:ActiveRuns.Add($st)
    Set-UiBusy -Panel $Panel -Busy $true
    if ($Unit -eq 'makine') {   # V6: tekrar deneme icin son islem (AD onbellek guncellemesi haric)
        $script:LastOp[$Panel] = [pscustomobject]@{ Items = @($Items); RunArgs = @{ Title = ($Title -replace ' - tekrar$', ''); Worker = $Worker; OnResultFn = $OnResultFn; Panel = $Panel; Unit = $Unit } }
    }

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

function Complete-Job {
    # Bitmis isin ciktisini / hatasini alir ve PowerShell nesnesini birakir.
    param($Job)
    $out = @(); $err = ''
    try { $out = @($Job.PS.EndInvoke($Job.Async)) } catch { $err = $_.Exception.Message }
    if ($out.Count -eq 0 -and -not $err -and $Job.PS.Streams.Error.Count -gt 0) { $err = $Job.PS.Streams.Error[0].ToString() }
    $Job.PS.Dispose()
    return [pscustomobject]@{ Out = $out; Err = $err }
}

function Invoke-JobResult {
    # Sonuc isleyiciyi cagirir. Isleyici isi ERTELERSE (script gonderildi, sonuc yoklamada) tamamlanan sayilmaz.
    param($Run, $Job, $Output, [string]$ErrorText, $Label)
    $script:DeferFlag = $false
    $script:CurrentRun = $Run
    try { & $Run.OnResultFn -Key $Job.Key -Output $Output -ErrorText $ErrorText }
    catch { $Label.Text = "Sonuc isleme hatasi: $($_.Exception.Message)" }
    if (-not $script:DeferFlag) { $Run.Done++ }
}

# ---------- TOPLU SCRIPT SONUCU YOKLAMASI (UI thread'i yonetir, sorgular arka plandaki ayri runspace'te) ----------
# Script her cihaza AYRI gonderilir (worker); gonderilen cihaz $script:ScriptPoll listesine girer ve havuzu birakir.
# 4 sn'de bir TEK yoklama isi (Get-ScriptPollStatus) bekleyen TUM cihazlarin sonucunu grup basina tek sorguyla alir;
# gelen sonuc aninda ilgili satira yazilir. Suresi dolan cihaz icin "cevap gelmedi" analizi (DNS/ping/CcmExec) cihaz
# basina havuzda calisir. Eskiden: cihaz basina 4 sn'de bir 1-2 sorgu + bekleme boyunca dolu havuz slotu.

function Receive-ScriptSend {
    # Script gonderim isinin sonucu: gonderildiyse yoklamaya alinir (is ERTELENIR), degilse son isleyiciye gider.
    param($Key, $Output, $ErrorText, [string]$FinalFn)
    $r = @($Output | Where-Object { $_ -and $_.PSObject.Properties['Pending'] })[0]
    if (-not $r -or -not $r.Pending) { & $FinalFn -Key $Key -Output $Output -ErrorText $ErrorText; return }

    $script:DeferFlag = $true
    [void]$script:ScriptPoll.Add([pscustomobject]@{
        Id = [guid]::NewGuid().ToString(); Run = $script:CurrentRun; Key = $Key; Base = $r; Phase = 'poll'; Job = $null; Seen = $false
        Provider = $r.SmsProvider; SiteCode = $r.SiteCode; Guid = $r.ScriptGuid; Rid = $r.ResourceId; Op = $r.OperationId; TaskId = $null
        Start = $r.SentAt; TimeoutSec = $r.TimeoutSec; Deadline = $r.SentAt.AddSeconds($r.TimeoutSec)
    })
    # Kontrol sonuclari (DNS / AD / servis) script sonucu beklenmeden tabloya yazilir
    $wait = "Gonderildi - sonuc bekleniyor (en fazla $($r.TimeoutSec) sn)"
    if ($FinalFn -eq 'Update-ParserFromScript') { Set-ParserCheckValues -HostName $Key -Values (ConvertTo-ParserCheckValues -Check $r.Check) }
    else {
        Set-GridValues -Hostname $Key -Values @{
            DnsStatus = $r.DnsStatus; IsOnline = $r.OnlineText; AdStatus = $r.AdStatus; LastLogon = $r.LastLogon
            Running = $r.Running; NotRunning = $r.NotRunning; CcmExec = $r.CcmExec; ActionResult = $wait
        }
    }
}
function Receive-ScriptSend1 { param($Key, $Output, $ErrorText) Receive-ScriptSend -Key $Key -Output $Output -ErrorText $ErrorText -FinalFn 'Update-GridFromScript' }
function Receive-ScriptSend2 { param($Key, $Output, $ErrorText) Receive-ScriptSend -Key $Key -Output $Output -ErrorText $ErrorText -FinalFn 'Update-ParserFromScript' }

function Complete-ScriptItem {
    # Yoklamadan sonuc geldi (-HasStatus) ya da sure doldu ve analiz bitti (-Verdict): sonuc nesnesini tamamlayip
    # ilgili panelin son isleyicisine (Update-GridFromScript / Update-ParserFromScript) verir.
    param($Item, $ExitCode, [string]$OutText, [switch]$HasStatus, $Verdict)
    $b = $Item.Base
    $b.Pending = $false
    if ($HasStatus) {
        $o  = ([string]$OutText).TrimStart([char]0xFEFF).Trim()
        $ok = ($null -eq $ExitCode) -or ([int]$ExitCode -eq 0)
        $b.Output   = $o
        $b.ExitCode = $ExitCode
        $b.Detail   = "OperationID: $($Item.Op)`r`nExit code: $ExitCode`r`nCikti:`r`n$o"
        $b.State    = if ($ok) { 'ok' } else { 'warn' }
        # Grid'e sadece CIKTI yazilir (satirlar ' | ' ile birlestirilir); exit code / OperationID Detay'dadir.
        $b.Result   = if ($o) { ((($o -split "`r`n|`n|`r") | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ' | ') }
                      elseif ($ok) { 'Script calisti, cikti bos.' } else { "Script calisti, cikti bos (exit code $ExitCode)." }
    }
    else {
        $b.State  = $Verdict.State
        $b.Result = $Verdict.Text
        $b.Detail = "OperationID: $($Item.Op)" + $(if ($Item.Seen) { "`r`nSCCM'de durum kaydi goruldu ancak sonuc/cikti gelmedi." } else { '' })
    }
    [void]$script:ScriptPoll.Remove($Item)
    $lbl = if ($Item.Run.Panel -eq 2) { $lblParserStatus } else { $lblStatus }
    Invoke-JobResult -Run $Item.Run -Job ([pscustomobject]@{ Key = $Item.Key }) -Output @($b) -ErrorText '' -Label $lbl
}

function Update-ScriptPoll {
    # RunTimer'dan cagrilir (200 ms). Agir isler (WMI) UI thread'inde DEGIL, arka planda calisir.
    # 1) Biten yoklama sorgusunun sonuclarini dagit
    if ($script:PollJob -and $script:PollJob.Async.IsCompleted) {
        $r = Complete-Job $script:PollJob
        $script:PollJob = $null
        $byId = @{}
        foreach ($x in $script:ScriptPoll) { $byId[$x.Id] = $x }
        foreach ($s in @($r.Out)) {
            $it = $byId[[string]$s.Id]
            if (-not $it -or $it.Phase -ne 'poll') { continue }   # bu arada iptal edilmis / suresi dolmus
            if ($s.TaskId) { $it.TaskId = $s.TaskId }
            if ($s.Seen) { $it.Seen = $true }
            if ($s.Done) { Complete-ScriptItem -Item $it -ExitCode $s.ExitCode -OutText $s.Output -HasStatus }
        }
        $script:PollError = $(if ($r.Err) { $r.Err } else { @($r.Out | Where-Object { $_.Error } | Select-Object -First 1 | ForEach-Object { $_.Error }) -join '' })
    }

    # 2) Suresi dolanlar: "cevap gelmedi" analizi cihaz basina (havuzda). Yoklama surerken beklenir ki son sonuc kacmasin.
    $now = Get-Date
    foreach ($it in @($script:ScriptPoll | Where-Object { -not $script:PollJob -and $_.Phase -eq 'poll' -and $now -ge $_.Deadline })) {
        $it.Phase = 'verdict'
        $ps = [powershell]::Create()
        $ps.RunspacePool = (Get-SharedPool)
        [void]$ps.AddScript($script:VerdictWorker).AddParameters(@{ ComputerName = $it.Key; TimeoutSec = $it.TimeoutSec })
        $it.Job = [pscustomobject]@{ PS = $ps; Async = $ps.BeginInvoke(); Key = $it.Key; Done = $false }
    }
    foreach ($it in @($script:ScriptPoll | Where-Object { $_.Phase -eq 'verdict' -and $_.Job.Async.IsCompleted })) {
        $r = Complete-Job $it.Job
        $v = @($r.Out | Where-Object { $_ -and $_.PSObject.Properties['Text'] })[0]
        if (-not $v) { $v = [pscustomobject]@{ State = 'warn'; Text = "$($it.TimeoutSec) sn icinde yanit gelmedi; sebep analizi yapilamadi ($($r.Err))." } }
        Complete-ScriptItem -Item $it -Verdict $v
    }

    # 3) Yeni yoklama: ayni anda tek sorgu, en fazla 4 sn'de bir
    $wait = @($script:ScriptPoll | Where-Object { $_.Phase -eq 'poll' -and ($now - $_.Start).TotalSeconds -ge 3 })
    if (-not $script:PollJob -and $wait.Count -gt 0 -and ($now - $script:LastPoll).TotalSeconds -ge 4) {
        if (-not $script:PollRunspace -or $script:PollRunspace.RunspaceStateInfo.State -ne 'Opened') {
            $script:PollRunspace = [runspacefactory]::CreateRunspace()
            $script:PollRunspace.Open()
        }
        $items = @($wait | ForEach-Object { [pscustomobject]@{ Id = $_.Id; Provider = $_.Provider; SiteCode = $_.SiteCode; Guid = $_.Guid; Rid = $_.Rid; Op = $_.Op; TaskId = $_.TaskId; Start = $_.Start } })
        $ps = [powershell]::Create()
        $ps.Runspace = $script:PollRunspace
        [void]$ps.AddScript($script:PollWorker).AddParameter('Items', $items)
        $script:PollJob = [pscustomobject]@{ PS = $ps; Async = $ps.BeginInvoke(); Key = 'poll'; Done = $false }
        $script:LastPoll = $now
        $script:PollCount++
    }
}

function Get-RunWaitCount {
    param($Run)
    return @($script:ScriptPoll | Where-Object { [object]::ReferenceEquals($_.Run, $Run) }).Count
}

function Stop-PanelRun {
    # IPTAL: paneldeki bekleyen / calisan isler durdurulur. Tamamlanmis sonuclar korunur; durdurulanlarin satirina
    # "Iptal edildi" yazilir. Uzak bir cagri (WMI/RPC) o an suruyorsa arka planda bitmesi beklenir, sonucu yok sayilir.
    # NOT: SCCM'e gonderilmis bir script / istemci eylemi geri ALINAMAZ; sadece sonucunu beklemek birakilir.
    param([int]$Panel)

    $runs = @($script:ActiveRuns | Where-Object { $_.Panel -eq $Panel })
    $pfs  = @($script:Prefetch | Where-Object { $_.Panel -eq $Panel })
    if ($runs.Count -eq 0 -and $pfs.Count -eq 0) { return }
    $left = 0
    foreach ($pf in $pfs) { $left += @($pf.Items).Count }
    foreach ($st in $runs) { $left += @($st.Jobs | Where-Object { -not $_.Done }).Count + (Get-RunWaitCount $st) }
    $c = [System.Windows.Forms.MessageBox]::Show(
        "Panel ${Panel}: $(($runs | ForEach-Object { $_.Title }) -join ', ')`n`n$left is henuz tamamlanmadi. Iptal edilsin mi?`n`nTamamlanan sonuclar tabloda kalir. SCCM'e gonderilmis script / eylemler geri alinamaz (sadece sonucu beklenmez).",
        'Iptal', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($c -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $lbl = if ($Panel -eq 2) { $lblParserStatus } else { $lblStatus }
    $n = 0
    foreach ($pf in $pfs) {   # SCCM on-almasi surerken iptal: is hic baslamadi
        [void]$script:Prefetch.Remove($pf)
        try { [void]$pf.Job.PS.BeginStop($null, $null) } catch { }
        [void]$script:CancelledJobs.Add($pf.Job)
        $fake = [pscustomobject]@{ Panel = $Panel; Done = 0; OnResultFn = $pf.RunArgs.OnResultFn }
        foreach ($it in $pf.Items) { Invoke-JobResult -Run $fake -Job $it -Output @() -ErrorText $script:CancelText -Label $lbl; $n++ }
    }
    foreach ($st in $runs) {
        foreach ($job in $st.Jobs) {
            if ($job.Done) { continue }
            $job.Done = $true
            if ($job.Async.IsCompleted) {   # o an bitmis: gercek sonucu isle
                $r = Complete-Job $job
                Invoke-JobResult -Run $st -Job $job -Output $r.Out -ErrorText $r.Err -Label $lbl
                continue
            }
            try { [void]$job.PS.BeginStop($null, $null) } catch { }
            [void]$script:CancelledJobs.Add($job)
            Invoke-JobResult -Run $st -Job $job -Output @() -ErrorText $script:CancelText -Label $lbl
            $n++
        }
        # Sonucu yoklamada bekleyen scriptler (gonderilmis; sadece beklemeyi birakiriz)
        foreach ($it in @($script:ScriptPoll | Where-Object { [object]::ReferenceEquals($_.Run, $st) })) {
            if ($it.Job -and -not $it.Job.Async.IsCompleted) { try { [void]$it.Job.PS.BeginStop($null, $null) } catch { }; [void]$script:CancelledJobs.Add($it.Job) }
            elseif ($it.Job) { $null = Complete-Job $it.Job }
            [void]$script:ScriptPoll.Remove($it)
            Invoke-JobResult -Run $st -Job ([pscustomobject]@{ Key = $it.Key }) -Output @() -ErrorText $script:CancelText -Label $lbl
            $n++
        }
        [void]$script:ActiveRuns.Remove($st)
        if ($st.Report -and $st.Report.Count -gt 0) { [void]$script:ReportQueue.Add($st) }
        Write-ActionLog -Hostname '-' -Action "$($st.Title) - IPTAL" -State 'warn' -Text "$($st.Total) isten $($st.Total - $n) tamamlandi, $n iptal edildi"
    }
    $lbl.ForeColor = [System.Drawing.Color]::DarkOrange
    $lbl.Text = "Iptal edildi: $n is durduruldu, tamamlanan sonuclar tabloda."
    Set-UiBusy -Panel $Panel -Busy $false
    if ($Panel -eq 2) { Update-ParserInfo }
    Show-PendingReports
}

$script:RunTimer.Add_Tick({
    # Iptal edilmis isler: uzak cagri bitince kaynak birakilir, sonuc yok sayilir
    foreach ($z in @($script:CancelledJobs)) {
        if ($z.Async.IsCompleted) { $null = Complete-Job $z; [void]$script:CancelledJobs.Remove($z) }
    }
    if ($script:Prefetch.Count -gt 0) { Update-Prefetch }
    if ($script:ScriptPoll.Count -gt 0 -or $script:PollJob) { Update-ScriptPoll }

    foreach ($st in @($script:ActiveRuns)) {
        $lbl = if ($st.Panel -eq 2) { $lblParserStatus } else { $lblStatus }
        $pending = $false
        foreach ($job in $st.Jobs) {
            if ($job.Done) { continue }
            if (-not $job.Async.IsCompleted) { $pending = $true; continue }
            $job.Done = $true
            $r = Complete-Job $job
            Invoke-JobResult -Run $st -Job $job -Output $r.Out -ErrorText $r.Err -Label $lbl
        }
        $w = Get-RunWaitCount $st
        $txt = "$($st.Title): $($st.Done) / $($st.Total) tamamlandi (ayni anda en fazla $($script:MaxParallel) istek)"
        if ($st.Unit -eq 'domain') { $txt = "AD onbellek guncelleme ($([int]((Get-Date) - $script:AdStart).TotalSeconds) sn)  |  $(Format-AdProgress)" }
        if ($w -gt 0) {
            $txt += "  |  $w cihazin script sonucu bekleniyor (toplu yoklama: 4 sn'de bir tek sorgu, $($script:PollCount). tur)"
            if ($script:PollError) { $txt += "  |  Son yoklama hatasi: $($script:PollError)" }
        }
        if ($lbl.Text -ne $txt) { $lbl.Text = $txt }
        if (-not $pending -and $w -eq 0) {
            [void]$script:ActiveRuns.Remove($st)
            if ($st.Report -and $st.Report.Count -gt 0) { [void]$script:ReportQueue.Add($st) }
            $lbl.Text = if ($st.Unit -eq 'domain' -and $script:AdSaveSummary) { $script:AdSaveSummary } else { "$($st.Title) tamamlandi. $($st.Total) $($st.Unit) islendi." }
            Set-UiBusy -Panel $st.Panel -Busy $false
            if ($st.Panel -eq 2) { Update-ParserInfo }
        }
    }
    if ($script:ActiveRuns.Count -eq 0 -and $script:CancelledJobs.Count -eq 0 -and $script:ScriptPoll.Count -eq 0 -and -not $script:PollJob -and $script:Prefetch.Count -eq 0) { $script:RunTimer.Stop() }
    Show-PendingReports   # en sonda: modal pencere acikken gelen tick'ler isleri islemeye devam eder
})

# ---------- Sonuc isleyiciler (UI thread'inde calisir) ----------

function Update-GridFromCheck {
    param($Key, $Output, $ErrorText)
    $r = @($Output)[0]
    if (-not $r -or -not $r.PSObject.Properties['DnsStatus']) {
        Resolve-PendingCells -Hostname $Key -ErrorText $ErrorText
        Set-GridValues -Hostname $Key -Values @{ ErrorMessage = (Format-JobError $ErrorText); State = 'fail' }
        return
    }
    $script:LastResults[$r.Hostname] = $r
    # State 'fail' = DNS cozumlenemedi / ping yok -> servis bilgisi alinamadi (ulasilamadi)
    $un = ($r.State -eq 'fail')
    Set-GridValues -Hostname $r.Hostname -Values @{
        DnsStatus    = $r.DnsStatus
        IsOnline     = $r.OnlineText
        AdStatus     = $r.AdStatus
        LastLogon    = $r.LastLogon
        Running      = $(if ($un) { 'Ulasilamadi' } else { $r.Running })
        NotRunning   = $(if ($un) { 'Ulasilamadi' } else { $r.NotRunning })
        CcmExec      = $(if ($un) { 'Ulasilamadi' } else { $r.CcmExec })
        ErrorMessage = $r.ErrorMessage
        State        = $r.State
    }
    Add-RowDetail -Hostname $r.Hostname -Text ("Kontrol: DNS=$($r.DnsStatus); Online=$($r.OnlineText); AD=$($r.AdStatus); LastLogon=$($r.LastLogon); CcmExec=$($r.CcmExec)" + $(if ($r.ErrorMessage) { "`r`nNot: $($r.ErrorMessage)" } else { '' }))
}

function ConvertFrom-FixRows {
    # Servis baslatma sonuc satirlarini (Start-RemoteFix) tablo degerlerine cevirir. $null = gecerli sonuc yok.
    param($Output)
    $rows = @($Output | Where-Object { $_ -and $_.PSObject.Properties['ServiceName'] })
    if ($rows.Count -eq 0) { return $null }
    $notRun = (($rows | Where-Object { $_.SonrakiDurum -ne 'Running' } | ForEach-Object { "$($_.ServiceName)($($_.SonrakiDurum))" }) -join ', ')
    return [pscustomobject]@{
        Running       = (($rows | Where-Object { $_.SonrakiDurum -eq 'Running' } | ForEach-Object { $_.ServiceName }) -join ', ')
        NotRunning    = $notRun
        State         = $(if ($notRun) { 'warn' } else { 'ok' })
        Summary       = (Get-HostSummary -HostResults $rows)
        Detail        = ("Servis baslatma:`r`n" + (($rows | ForEach-Object { "  $($_.ServiceName): $($_.OncekiDurum) -> $($_.SonrakiDurum) ($($_.Aciklama))" }) -join "`r`n"))
        ServiceDetail = @($rows | ForEach-Object { [pscustomobject]@{ Name = $_.ServiceName; Status = $_.SonrakiDurum } })
    }
}

function Update-GridFromFix {
    param($Key, $Output, $ErrorText)
    $f = ConvertFrom-FixRows $Output
    if (-not $f) {
        Set-GridValues -Hostname $Key -Values @{ LastAction = 'Servis baslat'; ActionResult = (Format-JobError $ErrorText); State = 'fail' }
        Add-FixReport -Hostname $Key -Status (Format-JobError $ErrorText)
        return
    }
    Add-FixReport -Hostname $Key -Status $f.Summary -Fix $f
    $e = $script:LastResults[$Key]   # sonraki "Servisleri Baslat" guncel duruma gore calissin
    if ($e) { $e.ServiceDetail = $f.ServiceDetail }
    Set-GridValues -Hostname $Key -Values @{
        Running = $f.Running; NotRunning = $f.NotRunning; State = $f.State
        LastAction = "Servis baslat ($(Get-Date -Format 'HH:mm:ss'))"; ActionResult = $f.Summary
    }
    Add-RowDetail -Hostname $Key -Text $f.Detail
    Write-ActionLog -Hostname $Key -Action 'Servis baslat' -State $f.State -Text $f.Summary
}

function Update-GridFromAction {
    # Hem SCCM istemci eylemleri hem script gonderimi icin ortak isleyici
    param($Key, $Output, $ErrorText)
    $r = @($Output)[0]
    if (-not $r -or -not $r.PSObject.Properties['Action']) {
        $r = [pscustomobject]@{ Hostname = $Key; Action = 'Islem'; State = 'fail'; Result = (Format-JobError $ErrorText); Detail = '' }
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
        $msg = (Format-JobError $ErrorText)
        Resolve-PendingCells -Hostname $Key -ErrorText $ErrorText
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
        IsOnline     = $r.OnlineText
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
    $script:LastResults[$r.Hostname] = [pscustomobject]@{ Hostname = $r.Hostname; IsOnline = [bool]$r.IsOnline; ServiceDetail = @($r.ServiceDetail) }

    $d = "$($r.Action): $($r.Result)`r`nAD: $($r.AdStatus); LastLogon: $($r.LastLogon); DNS: $($r.DnsStatus); Online: $($r.OnlineText); CcmExec: $($r.CcmExec)"
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
        $msg = (Format-JobError $ErrorText)
        Resolve-PendingCells -Hostname $Key -ErrorText $ErrorText
        Set-GridValues -Hostname $Key -Values @{ LastAction = "AD yapisi kontrolu ($(Get-Date -Format 'HH:mm:ss'))"; ActionResult = $msg; State = 'fail' }
        Add-RowDetail -Hostname $Key -Text "AD yapisi kontrolu: $msg"
        return
    }

    $summary = "AD: $($r.AdStatus); Last Logon: $($r.LastLogon)"
    Set-GridValues -Hostname $r.Hostname -Values @{
        DnsStatus    = $r.DnsStatus
        IsOnline     = $r.OnlineText
        AdStatus     = $r.AdStatus
        LastLogon    = $r.LastLogon
        LastAction   = "AD yapisi kontrolu ($(Get-Date -Format 'HH:mm:ss'))"
        ActionResult = $summary
        ErrorMessage = $r.ErrorMessage
        State        = $r.State
    }
    Add-RowDetail -Hostname $r.Hostname -Text ("AD yapisi kontrolu: DNS=$($r.DnsStatus); Online=$($r.OnlineText); AD=$($r.AdStatus); LastLogon=$($r.LastLogon)" + $(if ($r.ErrorMessage) { "`r`nNot: $($r.ErrorMessage)" } else { '' }))
    Write-ActionLog -Hostname $r.Hostname -Action 'AD yapisi kontrolu' -State $r.State -Text $summary
}

# ---------- AD onbellegi (CSV) guncelleme (UI thread'i) ----------

function Update-AdCacheFromWorker {
    # Her domain'in sonucu geldikce toplanir; hepsi gelince ozet (dosyalar worker'larda zaten yazildi).
    param($Key, $Output, $ErrorText)

    $r = @($Output | Where-Object { $_ -and $_.PSObject.Properties['Items'] -and $_.PSObject.Properties['Domain'] })[0]
    if (-not $r) {
        # Is hic sonuc donmedi: bu domain'in tum OU'lari basarisiz sayilir
        $fail = foreach ($s in @($script:AdCacheSources | Where-Object { $_.Domain -eq $Key })) {
            [pscustomobject]@{ SearchBase = [string]$s.OU; Ok = $false; Error = (Format-JobError $ErrorText); Count = 0; Seconds = 0 }
        }
        $r = [pscustomobject]@{ Domain = [string]$Key; Items = @($fail) }
    }
    $script:AdCacheNew[[string]$Key] = $r

    $domains = @($script:AdCacheSources | Group-Object -Property Domain).Count   # buyuk/kucuk harf duyarsiz
    if ($script:AdCacheNew.Count -ge $domains) { Complete-AdCacheUpdate }
}

function Complete-AdCacheUpdate {
    # Tum domain'ler bitti: ozet + artik tanimli olmayan OU dosyalarinin temizligi + etiket. Birlestirme / yeniden yazma YOK.
    $dir = Get-AdCacheDir
    $okCount = 0; $total = 0
    $parts    = New-Object 'System.Collections.Generic.List[string]'
    $problems = New-Object 'System.Collections.Generic.List[string]'
    $keep     = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($s in @($script:AdCacheSources)) {
        $dom = [string]$s.Domain; $ou = [string]$s.OU
        [void]$keep.Add((Get-AdCacheFileName -Domain $dom -OU $ou))
        $r = $null
        $dr = $script:AdCacheNew[$dom]
        if ($dr) { foreach ($x in $dr.Items) { if ($x.SearchBase -eq $ou) { $r = $x; break } } }
        if ($r -and $r.Ok) { $okCount++; $total += $r.Count; $parts.Add(("{0} [{1}]: {2:N0} ({3} sn)" -f $dom, $ou, $r.Count, $r.Seconds)) }
        else {
            $why = if ($r -and $r.Error) { $r.Error } else { 'sonuc alinamadi' }
            $old = Join-Path $dir (Get-AdCacheFileName -Domain $dom -OU $ou)
            $problems.Add("${dom} [$ou]: $why`n   ($(if (Test-Path -LiteralPath $old) { 'eski dosya korundu' } else { 'bu OU icin onbellek yok' }))")
        }
    }
    # Ayarlardan cikarilmis OU'larin dosyalari (ve yarim kalmis .tmp'ler) silinir; hepsi basariliysa eski tek dosya da.
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue)) {
        if ($f.Extension -eq '.tmp' -or ($f.Extension -eq '.csv' -and -not $keep.Contains($f.Name))) { try { Remove-Item -LiteralPath $f.FullName -Force } catch { } }
    }
    if ($okCount -eq @($script:AdCacheSources).Count -and (Test-Path -LiteralPath (Get-AdCacheFile))) { try { Remove-Item -LiteralPath (Get-AdCacheFile) -Force } catch { } }

    $idx  = Get-AdCacheIndex
    $secs = if ($script:AdStart) { [int]((Get-Date) - $script:AdStart).TotalSeconds } else { 0 }
    $uniq = if ($idx.Ok) { $idx.Unique } else { 0 }
    $script:AdSaveSummary = ("AD onbellegi guncellendi ({0} sn): AD'den {1:N0} kayit okundu, onbellekte {2:N0} tekil makine | {3}/{4} OU basarili | {5}" -f
        $secs, $total, $uniq, $okCount, @($script:AdCacheSources).Count, ($parts -join ' | '))
    if ($okCount -eq 0) { $script:AdSaveSummary = "AD onbellegi GUNCELLENMEDI: hicbir OU'dan veri alinamadi (mevcut dosyalar korundu)." }
    Update-AdCacheLabel
    $state = if ($okCount -eq 0) { 'fail' } elseif ($problems.Count -gt 0) { 'warn' } else { 'ok' }
    Write-ActionLog -Hostname '-' -Action 'AD onbellek guncelleme' -State $state -Text ((@($script:AdSaveSummary) + @($problems)) -join ' | ')
    if ($problems.Count -gt 0) {
        $title = if ($okCount -eq 0) { "Hicbir OU'dan veri alinamadi; mevcut onbellek DEGISTIRILMEDI." } else { "Onbellek guncellendi ancak bazi OU'lar alinamadi:" }
        [System.Windows.Forms.MessageBox]::Show("$title`n`n" + ($problems -join "`n`n"), 'AD Onbellegi', 'OK', $(if ($okCount -eq 0) { 'Error' } else { 'Warning' })) | Out-Null
    }
}

function Update-AdCacheLabel {
    # Onbellek durumunu (tekil makine + kayit + son guncelleme zamani) kontrol grubundaki etikete yazar.
    $idx = Get-AdCacheIndex
    if ($idx.Ok) {
        $age  = ((Get-Date) - $idx.Time).TotalDays
        $text = "Onbellek: {0:N0} makine ({1} dosya) | {2}" -f $idx.Unique, $idx.Files, $idx.Time.ToString('yyyy-MM-dd HH:mm')
        if ($age -gt 7) { $text += ' (eski)' }
        $lblAdCache.Text      = $text
        $lblAdCache.ForeColor = $(if ($age -gt 7) { [System.Drawing.Color]::DarkOrange } else { [System.Drawing.Color]::DimGray })
    }
    elseif (@($script:AdCacheSources).Count -eq 0) {
        $lblAdCache.Text      = "Domain / OU tanimli degil - 'AD Onbellegi...' ile ekleyin"
        $lblAdCache.ForeColor = [System.Drawing.Color]::Firebrick
    }
    else {
        $lblAdCache.Text      = 'Onbellek yok - AD bilgisi okunamaz'
        $lblAdCache.ForeColor = [System.Drawing.Color]::Firebrick
    }
}

function Get-AdSourcesFromGrid {
    # Dialogdaki tabloyu dogrular ve { Sources; Errors } dondurur.
    param($Grid)

    $list = @()
    $errs = @()
    $seen = New-Object System.Collections.Hashtable ([System.StringComparer]::OrdinalIgnoreCase)
    $null = $Grid.EndEdit()
    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }
        $d = ([string]$row.Cells[0].Value).Trim()
        $o = ([string]$row.Cells[1].Value).Trim()
        if (-not $d -and -not $o) { continue }
        $n = $row.Index + 1
        if ($d -notmatch '^[A-Za-z0-9][A-Za-z0-9.-]*$') { $errs += "Satir ${n}: gecerli bir domain adi girin (ornek: domain1.local)."; continue }
        if ($o -notmatch '(?i)DC=')                     { $errs += "Satir ${n}: OU, DistinguishedName olmali (ornek: OU=Workstations,DC=domain1,DC=local)."; continue }
        $k = "$d|$o"
        if ($seen.ContainsKey($k)) { $errs += "Satir ${n}: ayni domain / OU birden fazla kez girilmis."; continue }
        $seen[$k] = $true
        $list += [pscustomobject]@{ Domain = $d; OU = $o }
    }
    return [pscustomobject]@{ Sources = $list; Errors = $errs }
}

function Invoke-AdDialogApply {
    # Dialogdaki tabloyu dogrular; gecerliyse listeyi ana ayar dosyasina (JSON) kaydeder.
    $res = Get-AdSourcesFromGrid -Grid $script:AdDlgGrid
    if (@($res.Errors).Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(($res.Errors -join "`n"), 'Uyari', 'OK', 'Warning') | Out-Null
        return $false
    }
    $script:AdCacheSources = @($res.Sources)
    Save-Settings
    Update-AdCacheLabel
    return $true
}

function Show-AdCacheDialog {
    # Domain -> OU eslemesinin girildigi pencere. Liste ana ayar dosyasina (JSON) yazilir;
    # "Kaydet ve Onbellegi Guncelle" ayrica AD'den okuyup CSV onbellegini yeniden olusturur.
    $dlg = New-Ctl Form 'AD Onbellegi - Domain / OU Eslemesi' 0 0 860 450 @{ StartPosition = 'CenterParent'; MinimumSize = (New-Object System.Drawing.Size(700, 360)); ShowInTaskbar = $false }
    $lbl = New-Ctl Label ("Her satira bir domain ve o domain'de taranacak OU'yu yazin (OU, DistinguishedName olarak; ornek: OU=Workstations,DC=domain1,DC=local). " +
            "OU'nun ALTINDAKI (alt OU'lar dahil) tum bilgisayar hesaplari onbellege alinir. Ayni domain icin birden fazla satir girilebilir. " +
            "Satir silmek icin satir basligini secip Delete'e basin.") 12 10 820 48 @{ Anchor = 'Top,Left,Right' }

    $grid = New-Ctl DataGridView '' 12 64 820 260 @{
        Anchor = 'Top,Bottom,Left,Right'; AllowUserToAddRows = $true; AllowUserToDeleteRows = $true; AllowUserToResizeRows = $false
        RowHeadersVisible = $true; SelectionMode = 'RowHeaderSelect'; BackgroundColor = [System.Drawing.Color]::White; AutoSizeColumnsMode = 'Fill'
    }
    foreach ($c in @(@('Domain (DNS adi)', 30), @('OU (DistinguishedName)', 70))) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = $c[0]; $col.FillWeight = $c[1]
        [void]$grid.Columns.Add($col)
    }
    foreach ($s in @($script:AdCacheSources)) { [void]$grid.Rows.Add([string]$s.Domain, [string]$s.OU) }

    $idx = Get-AdCacheIndex
    $info = if ($idx.Ok) { "Mevcut onbellek: $($idx.Unique) makine ($($idx.Files) dosya), son guncelleme $($idx.Time.ToString('yyyy-MM-dd HH:mm'))`n$(Get-AdCacheDir)" } else { "Mevcut onbellek yok.`n$(Get-AdCacheDir)" }
    $lblInfo    = New-Ctl Label $info 12 332 820 34 @{ Anchor = 'Bottom,Left,Right'; ForeColor = [System.Drawing.Color]::DimGray }
    $btnSave    = New-Ctl Button 'Kaydet' 12 374 110 30 @{ Anchor = 'Bottom,Left' }
    $btnSaveUpd = New-Ctl Button 'Kaydet ve Onbellegi Guncelle' 130 374 210 30 @{ Anchor = 'Bottom,Left' }
    $btnClose   = New-Ctl Button 'Kapat' 732 374 100 30 @{ Anchor = 'Bottom,Right'; DialogResult = [System.Windows.Forms.DialogResult]::Cancel }

    # Olay isleyicileri closure DEGIL (GetNewClosure icinde $script: ana script yerine dinamik modulu gosterir)
    $script:AdDlg     = $dlg
    $script:AdDlgGrid = $grid
    $btnSave.Add_Click({
        if (Invoke-AdDialogApply) { $script:AdDlg.Tag = 'save'; $script:AdDlg.DialogResult = [System.Windows.Forms.DialogResult]::OK }
    })
    $btnSaveUpd.Add_Click({
        if (@($script:AdDlgGrid.Rows | Where-Object { -not $_.IsNewRow }).Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show('Once en az bir domain / OU satiri girin.', 'Uyari') | Out-Null
            return
        }
        if (Invoke-AdDialogApply) { $script:AdDlg.Tag = 'update'; $script:AdDlg.DialogResult = [System.Windows.Forms.DialogResult]::OK }
    })

    $dlg.CancelButton = $btnClose
    $dlg.Controls.AddRange(@($lbl, $grid, $lblInfo, $btnSave, $btnSaveUpd, $btnClose))
    [void]$dlg.ShowDialog($form)
    $doUpdate = ($dlg.Tag -eq 'update')
    $dlg.Dispose()
    if ($doUpdate) { Start-AdCacheUpdate }   # pencere kapandiktan sonra (arka planda, UI donmaz)
}

function Format-AdProgress {
    # Durum cubugu: "Toplam: 25.300 makine okundu (3 OU: 1 bitti, 1 okunuyor, 1 sirada) | dom1 [OU=..]: 12.500 bitti (8 sn) | ..."
    $p = $script:AdProgress
    if (-not $p) { return '' }
    $tot = 0; $st = @{}
    $parts = foreach ($k in $script:AdProgOrder) {
        $v = $p[$k]
        if (-not $v) { continue }
        $tot += [int]$v[0]
        $st[$v[1]] = 1 + [int]$st[$v[1]]
        $ou = ($k -replace '\[(OU|CN)=([^,\]]+)[^\]]*\]', '[$2]')   # durum cubugunda OU'nun sadece ilk adi
        "{0}: {1:N0} {2}{3}" -f $ou, [int]$v[0], $v[1], $(if ($v[2]) { " ($($v[2]) sn)" } else { '' })
    }
    $cnt = (@('bitti', 'yaziliyor', 'okunuyor', 'sirada', 'hata') | Where-Object { $st[$_] } | ForEach-Object { "$($st[$_]) $_" }) -join ', '
    return ("Toplam: {0:N0} makine okundu ({1} OU: {2})  |  " -f $tot, @($script:AdProgOrder).Count, $cnt) + (@($parts) -join '  |  ')
}

function Start-AdCacheUpdate {
    # Kullanici isterse: kayitli domain / OU eslemesindeki her OU AD'den taranir ve CSV onbellegi yeniden yazilir.
    # Domain basina TEK is baslatilir; o domain'in OU'lari is icinde sirayla taranir.
    # Ilerleme: $script:AdProgress (senkronize) -> durum cubugunda OU bazinda canli "N makine okundu".
    $src = @($script:AdCacheSources)
    if ($src.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Once 'AD Onbellegi...' penceresinden en az bir domain / OU girin.", 'Uyari') | Out-Null
        return
    }

    $script:AdCacheNew  = New-Object System.Collections.Hashtable ([System.StringComparer]::OrdinalIgnoreCase)
    $script:AdProgress  = [hashtable]::Synchronized(@{})
    $script:AdProgOrder = @($src | ForEach-Object { "$($_.Domain) [$($_.OU)]" })
    foreach ($k in $script:AdProgOrder) { $script:AdProgress[$k] = @(0, 'sirada', 0) }
    $script:AdSaveSummary = ''
    $script:AdStart = Get-Date
    $items = @(foreach ($g in ($src | Group-Object -Property Domain)) {
        $ous = [string[]]@($g.Group | ForEach-Object { [string]$_.OU })
        [pscustomobject]@{ Key = [string]$g.Name; Params = @{ Domain = [string]$g.Name; SearchBases = $ous; Progress = $script:AdProgress; CacheDir = (Get-AdCacheDir) } }
    })
    Start-ParallelRun -Title 'AD onbellek guncelleme' -Worker $script:AdCacheWorker -Items $items -OnResultFn 'Update-AdCacheFromWorker' -Unit 'domain'
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

function Get-ExportStamp {
    # Excel / CSV dosya adlari icin zaman damgasi: gun-ay-saat-dakika (ornek: 23-10-14-35)
    return (Get-Date -Format 'dd-MM-HH-mm')
}

function Export-ToExcelAndOpen {
    param([object[]]$Data, [string]$DefaultFileName = (Get-ExportStamp))

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
    finally {
        # COM referanslari birakilir (aksi halde kapatildiktan sonra arka planda EXCEL.EXE asili kalir)
        foreach ($o in @($headerRange, $ws, $wb, $excel)) { if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } } }
        [GC]::Collect()
    }
}

function Export-GridToCsv {
    # ';' ayiracli UTF8 CSV (Turkce Excel'de dogrudan kolonlara ayrilir). Write-FastCsv ile (Export-Csv degil).
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

    try { Write-FastCsv -Path $sfd.FileName -Header @($Data[0].PSObject.Properties.Name) -Rows $Data }
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

function Protect-CellText {
    # Excel/CSV formul enjeksiyonu onlemi: uzak makinelerden gelen script ciktisi '=', '+', '@', '-' ile basliyorsa Excel
    # onu FORMUL olarak calistirabilir (DDE vb.). Basina ' eklenir (Excel'de gorunmez, metin olarak saklanir). Sayilar dokunulmaz.
    param([string]$Text)
    if ($Text -match '^[=+@\t\r]' -or ($Text -match '^-' -and $Text -notmatch '^-\d+([.,]\d+)?$')) { return "'" + $Text }
    return $Text
}

function Get-GridExportData {
    # Gridde GORUNEN kolonlar (aktif filtre varsa sadece filtrelenmis satirlar).
    # -WithDetail: Panel 1'deki satir detay logunu 'Detay' kolonu olarak ekler.
    param($Grid, $Source, [bool]$WithDetail = $false, [string[]]$SkipColumns = @())

    $out = foreach ($drv in $Source) {
        $h = [ordered]@{}
        foreach ($col in $Grid.Columns) {
            if ($col.Visible -and ($SkipColumns -notcontains $col.Name)) { $h[$col.HeaderText] = Protect-CellText ([string]$drv[$col.DataPropertyName]) }
        }
        if ($WithDetail) {
            $hn = [string]$drv['Hostname']
            $h['Detay'] = if ($script:Details.ContainsKey($hn)) { Protect-CellText (($script:Details[$hn] -replace '\s*\r?\n\s*', ' | ').Trim(' ', '|')) } else { '' }
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
    # -StateColumn / -State: ozet cubugundan secilen durum (V7) metin filtresiyle AND'lenir.
    param($Source, $ColMap, [string]$ColumnHeader, [string]$Text, [string]$StateColumn = '', [string]$State = '')

    $parts = @()
    $q = "$Text".Trim()
    if ($q) {
        $esc = Format-LikeValue -Text $q
        if ($ColumnHeader -and $ColMap.Contains($ColumnHeader)) { $parts += "[$($ColMap[$ColumnHeader])] LIKE '%$esc%'" }
        else { $parts += '(' + ((@($ColMap.Values) | ForEach-Object { "[$_] LIKE '%$esc%'" }) -join ' OR ') + ')' }
    }
    if ($StateColumn -and $State) { $parts += "[$StateColumn] = '$State'" }
    if ($parts.Count -eq 0) { $Source.RemoveFilter() } else { $Source.Filter = ($parts -join ' AND ') }
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
        Online_mi      = [string]$Check.OnlineText
        AD_Durumu      = [string]$Check.AdStatus
        AD_LastLogon   = [string]$Check.LastLogon
        Calisan_Servisler    = [string]$Check.Running
        Calismayan_Servisler = [string]$Check.NotRunning
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
        $msg = (Format-JobError $ErrorText)
        $v = @{}; foreach ($cn in $script:ParserCheckColumns) { $v[$cn] = '' }; $v['DNS_Durumu'] = $msg
        Set-ParserCheckValues -HostName $Key -Values $v
        Add-ParserDetail -Hostname $Key -Text "Kontrol: $msg"
        return
    }

    Set-ParserCheckValues -HostName $Key -Values (ConvertTo-ParserCheckValues -Check $c)
    Set-ParserSvcState -Check $c
    Set-SccmColumns -Panel 2 -Hostname $Key
    $d = "Kontrol: DNS=$($c.DnsStatus); Online=$($c.OnlineText); AD=$($c.AdStatus); LastLogon=$($c.LastLogon); Calisan=$($c.Running); Calismayan=$($c.NotRunning); CcmExec=$($c.CcmExec); SCCM kaydi=$($c.SccmKayit)"
    if ($c.Note) { $d += "`r`nNot: $($c.Note)" }
    Add-ParserDetail -Hostname $Key -Text $d
}

function Update-ParserFromScript {
    # Panel 2'den gonderilen script sonucu: cikti '@' / ';' ile parse edilip alt alta yazilir,
    # ardindan DNS/AD/SCCM kontrol kolonlari doldurulur. Panel 1'e DOKUNMAZ.
    param($Key, $Output, $ErrorText)

    $r = @($Output | Where-Object { $_ -and $_.PSObject.Properties['Output'] -and $_.PSObject.Properties['Check'] })[0]
    if (-not $r) {
        $msg = (Format-JobError $ErrorText)
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
    Set-ParserSvcState -Check $chk
    Set-SccmColumns -Panel 2 -Hostname $hn

    $d = "$($r.Action): $($r.Result)`r`nDNS=$($chk.DnsStatus); Online=$($chk.OnlineText); AD=$($chk.AdStatus); LastLogon=$($chk.LastLogon); Calisan=$($chk.Running); Calismayan=$($chk.NotRunning); CcmExec=$($chk.CcmExec); SCCM kaydi=$($chk.SccmKayit)"
    if ($chk.Note) { $d += "`r`nNot: $($chk.Note)" }
    if ($r.Detail) { $d += "`r`n$($r.Detail)" }
    Add-ParserDetail -Hostname $hn -Text $d
    Write-ActionLog -Hostname $hn -Action $r.Action -State $r.State -Text $r.Result
}

function Set-ParserSvcState {
    # Panel 2 "Servisleri Baslat" icin hostun son servis durumunu saklar
    param($Check)
    if (-not $Check -or -not $Check.Hostname) { return }
    $script:ParserSvc[[string]$Check.Hostname] = [pscustomobject]@{ Hostname = [string]$Check.Hostname; IsOnline = [bool]$Check.IsOnline; ServiceDetail = @($Check.ServiceDetail) }
}

function Update-ParserFromFix {
    # Panel 2 "Servisleri Baslat" sonucu: hostun TUM satirlarinda servis kolonlari guncellenir. Panel 1'e DOKUNMAZ.
    param($Key, $Output, $ErrorText)
    $f = ConvertFrom-FixRows $Output
    if (-not $f) {
        $msg = Format-JobError $ErrorText
        Set-ParserCheckValues -HostName $Key -Values @{ Calismayan_Servisler = "Servis baslatma: $msg" }
        Add-ParserDetail -Hostname $Key -Text "Servis baslatma: $msg"
        Write-ActionLog -Hostname $Key -Action 'Servis baslat (Parser)' -State 'fail' -Text $msg
        Add-FixReport -Hostname $Key -Status $msg
        return
    }
    Add-FixReport -Hostname $Key -Status $f.Summary -Fix $f
    $e = $script:ParserSvc[$Key]
    if ($e) { $e.ServiceDetail = $f.ServiceDetail }
    Set-ParserCheckValues -HostName $Key -Values @{ Calisan_Servisler = $f.Running; Calismayan_Servisler = $f.NotRunning }
    Add-ParserDetail -Hostname $Key -Text "$($f.Detail)`r`nOzet: $($f.Summary)"
    Write-ActionLog -Hostname $Key -Action 'Servis baslat (Parser)' -State $f.State -Text $f.Summary
}

function Start-ParserChecks {
    # Verilen hostlar icin DNS / AD / SCCM ajan kontrolunu (Panel 2 isi olarak) baslatir.
    param([string[]]$Hosts)

    foreach ($h in $Hosts) {
        $pv = @{}
        foreach ($c in $script:ParserCheckColumns) { $pv[$c] = $script:PendingText }
        Set-ParserCheckValues -HostName $h -Values $pv
    }

    $svc = @(Get-ServiceList)
    $items = @(foreach ($h in $Hosts) {
        [pscustomobject]@{ Key = $h; Params = @{ ComputerName = $h; SmsProvider = [string]$script:SmsServer; SiteCode = [string]$script:SiteCode; ServiceNames = $svc } }
    })
    Start-SccmPrefetch -Panel 2 -Items $items -RunArgs @{ Title = 'Parser kontrolu (DNS / AD / servis / SCCM ajan)'; Worker = $script:ParserCheckWorker; OnResultFn = 'Update-ParserFromCheck'; Panel = 2 }
}

# ---------- Hedef secimi ----------

function Get-ComputerList {
    # Soldaki kutudaki gecerli, tekil adlar (sira korunur). List + HashSet: binlerce satirda da O(n).
    $valid = New-Object 'System.Collections.Generic.List[string]'
    $bad   = New-Object 'System.Collections.Generic.List[string]'
    $seen  = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($n in ($txtComputers.Text -split '[\r\n,;\s]+')) {
        if (-not $n) { continue }
        if ($n -match $script:HostRegex) { if ($seen.Add($n)) { $valid.Add($n) } }   # Turkce karakterler dahil; Panel 1 ve 2 ayni kural
        else { $bad.Add($n) }
    }
    if ($bad.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("Gecersiz bilgisayar isimleri yok sayildi:`n$($bad -join ', ')", 'Uyari') | Out-Null
    }
    if ($valid.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('En az bir gecerli bilgisayar ismi girin.', 'Uyari') | Out-Null
        return @()
    }
    return $valid.ToArray()
}

function Get-TargetHosts {
    # Verilen gridde secili hucrelerin satirlari (Hostname kolonu) varsa onlar, yoksa soldaki listedeki tum makineler.
    param($Grid = $dgv)

    $sel = @(Get-SelectedHosts -Grid $Grid)
    if ($sel.Count -gt 0) { return [pscustomobject]@{ Hosts = $sel; Source = 'gridde secili satirlar' } }
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

function Start-ServiceFix {
    # "Servisleri Baslat" (Panel 1 ve Panel 2 ayri): son kontrolde durdurulmus servisi olan cevrimici hostlar.
    # Tabloda secim varsa SADECE secili hostlar hedeflenir. Sonuc sadece istegin geldigi panele yazilir.
    param([int]$Panel)

    $grid = if ($Panel -eq 2) { $dgvParse } else { $dgv }
    $map  = if ($Panel -eq 2) { $script:ParserSvc } else { $script:LastResults }
    $sel  = New-NameMap
    foreach ($h in @(Get-SelectedHosts -Grid $grid)) { $sel[$h] = 1 }
    $targets = @($map.Values | Where-Object {
        $_.IsOnline -and ($sel.Count -eq 0 -or $sel.ContainsKey($_.Hostname)) -and
        @($_.ServiceDetail | Where-Object { $_.Status -ne 'Running' -and $_.Status -ne 'NotFound' -and $_.Status -ne 'Erisilemedi' }).Count -gt 0
    })
    if ($targets.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Baslatilacak bir servis bulunamadi ($(if ($sel.Count) { 'secili satirlarda ' })hepsi calisiyor / bulunamadi / erisilemedi ya da makine cevrimdisi).", 'Bilgi') | Out-Null
        return
    }
    $scope = if ($sel.Count) { 'secili satirlar' } else { 'tablodaki tum cihazlar' }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Panel ${Panel} ($scope): $($targets.Count) makinede durdurulmus servisler bulundu.`n$(Format-HostPreview -Hosts @($targets | ForEach-Object { $_.Hostname }))`n`nBaslatilsin mi? (Ayni anda en fazla $($script:MaxParallel) makine islenir)",
        'Onay', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $items = @(foreach ($t in $targets) { [pscustomobject]@{ Key = $t.Hostname; Params = @{ ComputerName = $t.Hostname; ServiceDetail = $t.ServiceDetail } } })
    $hosts = @($targets | ForEach-Object { $_.Hostname })
    if ($Panel -eq 2) {
        foreach ($h in $hosts) { Set-ParserCheckValues -HostName $h -Values @{ Calismayan_Servisler = $script:PendingText } }
        Start-ParallelRun -Title 'Servis baslatma (Parser)' -Worker $script:FixWorker -Items $items -OnResultFn 'Update-ParserFromFix' -Panel 2
    }
    else {
        Set-GridPending -Hosts $hosts -Columns @('ActionResult')
        Start-ParallelRun -Title 'Servis baslatma' -Worker $script:FixWorker -Items $items -OnResultFn 'Update-GridFromFix'
    }
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

    $services = @(Get-ServiceList)
    if ($services.Count -eq 0) {
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
        $extra = "DNS / AD / servis ($($services -join ', ')) / SCCM ajan kontrolu de calistirilacak.`nSonuclar SADECE Panel 2'ye (Update Parser) yazilacak; Panel 1 etkilenmez."
    }
    else {
        $extra = "AD yapisi kontrolu (DNS / ping / AD) ve servis kontrolu ($($services -join ', ')) de calistirilacak.`nSonuclar SADECE Panel 1'e yazilacak; Update Parser (Panel 2) etkilenmez."
    }

    $msg = "Script: $($s.Name)  (v$($s.Version))`nYazar: $($s.Author)`n`nHedef ($($t.Source)): $(@($t.Hosts).Count) makine`n$(Format-HostPreview -Hosts $t.Hosts)`n`nCevap bekleme suresi: $timeout sn`n$extra`n`nCevrimdisi (ya da kimligi dogrulanamayan) cihazlara script GONDERILMEZ.`nScript bu makinelerde SYSTEM hesabiyla calisacak. Gonderilsin mi?"
    if ($param.HasParams) {
        $msg += "`n`nUYARI: Bu script parametre tanimliyor; parametreler BOS gonderilecek."
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show($msg, "Onay - Script Gonder (Panel $Panel)",
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $b64   = $param.Param
    $items = @(foreach ($h in $t.Hosts) {
        $p = @{
            ComputerName = $h
            SmsProvider  = $script:SmsServer
            SiteCode     = $script:SiteCode
            ScriptGuid   = $s.Guid
            ScriptName   = $s.Name
            ParamB64     = $b64
            TimeoutSec   = $timeout
            ServiceNames = $services
        }
        [pscustomobject]@{ Key = $h; Params = $p }
    })
    Save-Settings

    if ($Panel -eq 2) {
        foreach ($h in $t.Hosts) { Add-ParserPendingRow -HostName $h }
        Update-ParserInfo
        Start-SccmPrefetch -Panel 2 -Items $items -RunArgs @{ Title = 'Script + DNS/AD/servis/SCCM (Parser)'; Worker = $script:ParserScriptWorker; OnResultFn = 'Receive-ScriptSend2'; Panel = 2 }
    }
    else {
        Set-GridPending -Hosts $t.Hosts -ClearError -Columns @('DnsStatus', 'IsOnline', 'AdStatus', 'LastLogon', 'Running', 'NotRunning', 'CcmExec', 'ActionResult')
        Start-SccmPrefetch -Panel 1 -Items $items -RunArgs @{ Title = 'Script + AD + servis kontrolu'; Worker = $script:ScriptWorker; OnResultFn = 'Receive-ScriptSend1'; Panel = 1 }
    }
}
# ---------- V2: SCCM kayitlarini toplu on-alma ----------

function Start-SccmPrefetch {
    # Ise baslamadan once tum hedeflerin SCCM kaydi (SMS_R_System) arka planda TEK (parcali) sorguyla alinir;
    # bitince her ise SccmInfo eklenip Start-ParallelRun cagrilir. Worker'lar SMS Provider'a kayit sorgusu ATMAZ.
    # SCCM baglantisi yoksa dogrudan baslatilir.
    param([int]$Panel, [object[]]$Items, [hashtable]$RunArgs)
    if (-not $script:SmsServer -or -not $script:SiteCode -or @($Items).Count -eq 0) { Start-ParallelRun @RunArgs -Items $Items; return }

    $ps = [powershell]::Create()
    $ps.RunspacePool = (Get-SharedPool)
    [void]$ps.AddScript($script:SccmMapWorker).AddParameters(@{ Names = [string[]]@($Items | ForEach-Object { $_.Key }); SmsProvider = $script:SmsServer; SiteCode = $script:SiteCode })
    [void]$script:Prefetch.Add([pscustomobject]@{ Panel = $Panel; Items = $Items; RunArgs = $RunArgs; Job = [pscustomobject]@{ PS = $ps; Async = $ps.BeginInvoke(); Key = 'sccm'; Done = $false } })
    Set-UiBusy -Panel $Panel -Busy $true
    $lbl = if ($Panel -eq 2) { $lblParserStatus } else { $lblStatus }
    $lbl.ForeColor = [System.Drawing.Color]::DarkBlue
    $lbl.Text = "SCCM kayitlari aliniyor: $(@($Items).Count) cihaz (toplu sorgu)..."
    if (-not $script:RunTimer.Enabled) { $script:RunTimer.Start() }
}

function Set-SccmColumns {
    # Boundary Group(s) / Operating System Build kolonlarini son toplu SCCM sorgusundan ($script:SccmExtra) doldurur.
    param([int]$Panel, [string]$Hostname)
    $v = $script:SccmExtra[(($Hostname -split '\.')[0])]
    if (-not $v) { return }
    if ($Panel -eq 2) { Set-ParserCheckValues -HostName $Hostname -Values @{ Boundary_Groups = $v[0]; OS_Build = $v[1] } }
    elseif ($script:ResultsTable.Rows.Find($Hostname)) { Set-GridValues -Hostname $Hostname -Values @{ BoundaryGroups = $v[0]; OsBuild = $v[1] } }
}

function Update-Prefetch {
    # RunTimer'dan: biten on-alma isinin sonucunu islere dagitip asil isi baslatir. Boundary Group(s) / OS Build
    # kolonlari burada (UI thread'inde) hemen doldurulur ve $script:SccmExtra'da saklanir (Panel 2 satirlari yeniden
    # kurulunca tekrar uygulanir).
    foreach ($pf in @($script:Prefetch)) {
        if (-not $pf.Job.Async.IsCompleted) { continue }
        [void]$script:Prefetch.Remove($pf)
        $r = Complete-Job $pf.Job
        $map = @($r.Out | Where-Object { $_ -is [hashtable] })[0]
        foreach ($it in $pf.Items) {   # kayit yoksa SccmInfo verilmez: worker eskisi gibi cihaz basina sorgular
            $k = (([string]$it.Key -split '\.')[0])
            if ($map -and $map.ContainsKey($k)) {
                $it.Params['SccmInfo'] = $map[$k]
                $script:SccmExtra[$k] = [string[]]@([string]$map[$k].BoundaryGroups, [string]$map[$k].OsBuild)
            }
            Set-SccmColumns -Panel $pf.Panel -Hostname ([string]$it.Key)
        }
        $a = $pf.RunArgs
        if ($a.Worker) { Start-ParallelRun @a -Items $pf.Items }
        else { Set-UiBusy -Panel $pf.Panel -Busy $false }   # sadece SCCM bilgisi istendi
    }
}

# ---------- V6: Basarisizlari tekrar dene ----------

function Get-FailedHosts {
    # Son islemin (LastOp) hedeflerinden sonucu HATA (kirmizi) olanlar. Panel 1: State = fail, Panel 2: Durum = fail.
    param([int]$Panel)
    $op = $script:LastOp[$Panel]
    if (-not $op) { return @() }
    $bad = New-NameMap
    if ($Panel -eq 1) {
        foreach ($row in $script:ResultsTable.Rows) { if ([string]$row['State'] -eq 'fail') { $bad[[string]$row['Hostname']] = 1 } }
    }
    else {
        Update-SummaryNow -Panel 2
        foreach ($row in $script:ParseTable.Rows) {
            if ([string]$row['Durum'] -eq 'fail') { $bad[[string]$row['Hostname']] = 1; $bad[[string]$row['SourceHost']] = 1 }
        }
    }
    return @($op.Items | Where-Object { $bad.ContainsKey([string]$_.Key) })
}

function Start-Retry {
    param([int]$Panel)
    if (Test-PanelBusy -Panel $Panel) { return }
    $op = $script:LastOp[$Panel]
    if (-not $op) { [System.Windows.Forms.MessageBox]::Show('Bu panelde tekrar denenecek bir islem yok.', 'Bilgi') | Out-Null; return }
    $items = @(Get-FailedHosts -Panel $Panel)
    if ($items.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Son islemde ($($op.RunArgs.Title)) hata alan cihaz yok.", 'Bilgi') | Out-Null; return }

    $hosts = @($items | ForEach-Object { [string]$_.Key })
    $warn = if ($op.RunArgs.OnResultFn -like 'Receive-ScriptSend*' -or $op.RunArgs.OnResultFn -in 'Update-GridFromAction', 'Update-GridFromFix', 'Update-ParserFromFix') { "`n`nDIKKAT: islem bu cihazlara TEKRAR gonderilecek (script / eylem / servis baslatma)." } else { '' }
    $c = [System.Windows.Forms.MessageBox]::Show("Islem: $($op.RunArgs.Title)`n`nHata alan (kirmizi) $($hosts.Count) cihaz:`n$(Format-HostPreview -Hosts $hosts)$warn`n`nTekrar denensin mi?",
        'Basarisizlari Tekrar Dene', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($c -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    # Ayni parametrelerle yeni is nesneleri (Params kopyalanir; ayni hashtable iki iste kullanilmasin)
    $new = @($items | ForEach-Object { [pscustomobject]@{ Key = $_.Key; Params = $_.Params.Clone() } })
    $fn = $op.RunArgs.OnResultFn
    if ($Panel -eq 2) {
        if ($fn -eq 'Receive-ScriptSend2') { foreach ($h in $hosts) { Add-ParserPendingRow -HostName $h }; Update-ParserInfo }
        elseif ($fn -eq 'Update-ParserFromFix') { foreach ($h in $hosts) { Set-ParserCheckValues -HostName $h -Values @{ Calismayan_Servisler = $script:PendingText } } }
        else { $pv = @{}; foreach ($cn in $script:ParserCheckColumns) { $pv[$cn] = $script:PendingText }; foreach ($h in $hosts) { Set-ParserCheckValues -HostName $h -Values $pv } }
    }
    else {
        $cols = if ($fn -in 'Update-GridFromCheck', 'Receive-ScriptSend1') { @('DnsStatus', 'IsOnline', 'AdStatus', 'LastLogon', 'Running', 'NotRunning', 'CcmExec', 'ActionResult') }
                elseif ($fn -eq 'Update-GridFromAd') { @('DnsStatus', 'IsOnline', 'AdStatus', 'LastLogon', 'ActionResult') } else { @('ActionResult') }
        Set-GridPending -Hosts $hosts -ClearError -Columns $cols
    }
    $a = $op.RunArgs.Clone()
    $a.Title = "$($a.Title) - tekrar"
    Start-ParallelRun @a -Items $new
}

# ---------- V7: Durum ozet cubugu ----------

function Get-ParserRowState {
    # Panel 2 satir durumu (ozet cubugu, filtre ve "tekrar dene" icin): pending / fail / warn / ok
    param($Row)
    foreach ($c in @('UpdateName') + $script:ParserCheckColumns) { if ([string]$Row[$c] -eq $script:PendingText) { return 'pending' } }
    $on = [string]$Row['Online_mi']; $dns = [string]$Row['DNS_Durumu']; $ku = [string]$Row['KolonUyumu']
    if ($on -like 'Hayir*' -or $dns -match '^(No such host|DNS hatasi|Islem hatasi|Iptal edildi)' -or $ku -eq 'Cikti yok') { return 'fail' }
    $nr = [string]$Row['Calismayan_Servisler']; $cc = [string]$Row['CcmExec_Durumu']; $ad = [string]$Row['AD_Durumu']; $sk = [string]$Row['SCCM_Kayit']
    if ($ku -notlike 'OK*' -or ($nr -and $nr -notlike 'Kontrol edilemedi*') -or ($cc -and $cc -ne 'Running') -or ($ad -and $ad -ne 'Enabled') -or
        $on -eq 'Evet (dogrulanamadi)' -or $sk -like 'SCCM de kayitli degil*' -or $sk -like 'Kayit var*') { return 'warn' }
    return 'ok'
}

function Update-SummaryNow {
    # Panel 2: degisen satirlarin Durum'u yeniden hesaplanir. Sonra sayaclar tum tablodan sayilir.
    param([int]$Panel)
    if ($Panel -eq 2) {
        $rows = @($script:DirtyRows)
        $script:DirtyRows.Clear()
        foreach ($row in $rows) {
            if ($row.RowState -eq 'Detached' -or $row.RowState -eq 'Deleted') { continue }
            $st = Get-ParserRowState $row
            if ([string]$row['Durum'] -ne $st) { $row['Durum'] = $st }
        }
        $script:DirtyRows.Clear()   # Durum yazimi satiri tekrar kirli isaretler; deger degismedigi icin gerek yok
    }
    # DIKKAT: DataTable 'if' / pipeline ciktisi olarak dondurulurse PowerShell onu SATIRLARINA acar; dogrudan atanir.
    $tbl = $script:ResultsTable; $col = 'State'
    if ($Panel -eq 2) { $tbl = $script:ParseTable; $col = 'Durum' }
    $n = @{ ok = 0; warn = 0; fail = 0; pending = 0 }
    foreach ($row in $tbl.Rows) { $v = [string]$row[$col]; if ($n.ContainsKey($v)) { $n[$v]++ } }
    $b = $script:SumButtons[$Panel]
    $b.ok.Text = "Tamam: $($n.ok)"; $b.warn.Text = "Uyari: $($n.warn)"; $b.fail.Text = "Hata: $($n.fail)"; $b.pending.Text = "Bekleyen: $($n.pending)"
    $b.all.Text = "Tumu: $($tbl.Rows.Count)"
    $script:RetryButtons[$Panel].Enabled = ($n.fail -gt 0 -and $script:LastOp[$Panel] -and -not (Test-PanelBusy -Panel $Panel))
    $script:SumDirty[$Panel] = $false
}

function Set-StateFilter {
    # Ozet dugmesine tiklama: o duruma filtrele; ayni dugmeye tekrar ya da 'Tumu' -> durum filtresi kalkar.
    param([int]$Panel, [string]$State)
    if ($State -eq 'all' -or $script:StateFilter[$Panel] -eq $State) { $State = '' }
    $script:StateFilter[$Panel] = $State
    foreach ($k in @('ok', 'warn', 'fail', 'pending', 'all')) {
        $btn = $script:SumButtons[$Panel][$k]
        $on = ($k -eq $State) -or ($k -eq 'all' -and -not $State)
        $btn.Font = New-Object System.Drawing.Font('Segoe UI', 8.25, $(if ($on) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }))
    }
    Update-Filter -Panel $Panel
}

function Update-Filter {
    param([int]$Panel)
    if ($Panel -eq 2) { Set-GridFilter -Source $script:ParseBinding -ColMap $script:ColMap2 -ColumnHeader ([string]$cboF2.SelectedItem) -Text $txtF2.Text -StateColumn 'Durum' -State $script:StateFilter[2] }
    else { Set-GridFilter -Source $script:BindingSrc -ColMap $script:ColMap1 -ColumnHeader ([string]$cboF1.SelectedItem) -Text $txtF1.Text -StateColumn 'State' -State $script:StateFilter[1] }
}

function New-SummaryBar {
    # Ikinci arac satiri: durum dugmeleri + "Basarisizlari Tekrar Dene" + otomatik yenileme
    param([int]$Panel, $Parent)
    $defs = @(@('ok', 'LightGreen'), @('warn', 'LightYellow'), @('fail', 'LightCoral'), @('pending', 'Gainsboro'), @('all', 'White'))
    $script:SumButtons[$Panel] = @{}
    $x = 6
    foreach ($d in $defs) {
        $b = New-Ctl Button '' $x 37 92 24 @{ FlatStyle = 'Flat'; BackColor = [System.Drawing.Color]::FromName($d[1]); Tag = "$Panel|$($d[0])" }
        $b.Add_Click({ $p = ([string]$this.Tag).Split('|'); Set-StateFilter -Panel ([int]$p[0]) -State $p[1] })
        $script:SumButtons[$Panel][$d[0]] = $b
        $Parent.Controls.Add($b)
        $x += 96
    }
    $r = New-Ctl Button 'Basarisizlari Tekrar Dene' ($x + 8) 36 170 26 @{ Enabled = $false; Tag = $Panel }
    $r.Add_Click({ Start-Retry -Panel ([int]$this.Tag) })
    $script:RetryButtons[$Panel] = $r
    $chk = New-Ctl CheckBox 'Otomatik yenile (dk):' ($x + 196) 40 0 0 @{ Tag = $Panel }
    $nud = New-Ctl NumericUpDown '' ($x + 340) 38 50 23 @{ Minimum = 1; Maximum = 240 }
    $nud.Value = 10
    $lbl = New-Ctl Label '' ($x + 398) 42 0 0 $gray
    $chk.Add_CheckedChanged({ Reset-AutoRefresh -Panel ([int]$this.Tag) })
    $nud.Tag = $Panel
    $nud.Add_ValueChanged({ Reset-AutoRefresh -Panel ([int]$this.Tag) })
    $script:Auto[$Panel] = [pscustomobject]@{ Check = $chk; Minutes = $nud; Label = $lbl; Next = [datetime]::MaxValue }
    $Parent.Controls.AddRange(@($r, $chk, $nud, $lbl))
    Set-StateFilter -Panel $Panel -State ''
}

# ---------- V13: Periyodik yenileme ----------

function Reset-AutoRefresh {
    param([int]$Panel)
    $a = $script:Auto[$Panel]
    if ($a.Check.Checked) { $a.Next = (Get-Date).AddMinutes([int]$a.Minutes.Value); $a.Label.Text = "Sonraki: $($a.Next.ToString('HH:mm'))" }
    else { $a.Next = [datetime]::MaxValue; $a.Label.Text = '' }
}

function Invoke-AutoRefresh {
    # AutoTimer (15 sn): zamani gelen ve mesgul olmayan panelde tablodaki TUM cihazlar icin kontrol tekrarlanir.
    # Panel 1: Kontrol (DNS / Online / AD / servis), Panel 2: Kontrolleri Calistir. Tablo silinmez, satirlar yerinde guncellenir.
    foreach ($p in 1, 2) {
        $a = $script:Auto[$p]
        if (-not $a -or -not $a.Check.Checked -or (Get-Date) -lt $a.Next) { continue }
        if (Test-PanelBusy -Panel $p) { $a.Label.Text = 'Sonraki: panel mesgul, bekleniyor'; continue }
        $tbl = $script:ResultsTable
        if ($p -eq 2) { $tbl = $script:ParseTable }   # 'if' ciktisi olarak atanmaz (DataTable satirlarina acilir)
        $seen = New-NameMap
        $hosts = @(foreach ($row in $tbl.Rows) { $h = [string]$row['Hostname']; if ($h -match $script:HostRegex -and -not $seen.ContainsKey($h)) { $seen[$h] = 1; $h } })
        $svc = @(Get-ServiceList)
        Reset-AutoRefresh -Panel $p
        if ($hosts.Count -eq 0 -or ($p -eq 1 -and $svc.Count -eq 0)) { $a.Label.Text += ' (tablo bos / servis listesi bos - atlandi)'; continue }
        Write-ActionLog -Hostname '-' -Action "Otomatik yenileme (Panel $p)" -State 'ok' -Text "$($hosts.Count) cihaz"
        if ($p -eq 2) { Start-ParserChecks -Hosts $hosts } else { Start-Check -Hosts $hosts -Services $svc }
    }
}

function Start-Check {
    # Panel 1 "Kontrol": -Reset ile tablo bastan kurulur (Kontrol Et), aksi halde satirlar yerinde guncellenir (otomatik yenileme).
    # Boundary Group(s) / OS Build icin once toplu SCCM sorgusu (SCCM'e bagliysa).
    param([string[]]$Hosts, [string[]]$Services, [switch]$Reset)
    if ($Reset) {
        $script:ResultsTable.Rows.Clear()
        $script:LastResults.Clear()
        $script:Details.Clear()
    }
    $items = @(foreach ($h in $Hosts) { [pscustomobject]@{ Key = $h; Params = @{ ComputerName = $h; ServiceNames = $Services } } })
    Set-GridPending -Hosts $Hosts -ClearError -Columns @('DnsStatus', 'IsOnline', 'AdStatus', 'LastLogon', 'Running', 'NotRunning', 'CcmExec')
    if ($Reset) { try { $dgv.ClearSelection(); $dgv.CurrentCell = $null } catch { } }   # otomatik secilen ilk satir hedefi daraltmasin
    Start-SccmPrefetch -Panel 1 -Items $items -RunArgs @{ Title = 'Kontrol'; Worker = $script:CheckWorker; OnResultFn = 'Update-GridFromCheck'; Panel = 1 }
}

# ---------- V14: Islem kaydi (denetim log) goruntuleyici ----------

function Import-LogView {
    # Secili ayin log dosyasini tabloya yukler (en yeni kayit en ustte).
    $f = [string]$script:LogCbo.SelectedItem
    $script:LogTbl.Clear()
    if (-not $f) { return }
    $path = Join-Path (Get-DataDir) $f
    $lines = $null
    for ($i = 0; $i -lt 3 -and $null -eq $lines; $i++) {
        try { $lines = [System.IO.File]::ReadAllLines($path, [System.Text.Encoding]::UTF8) } catch { Start-Sleep -Milliseconds 200 }
    }
    if ($null -eq $lines) { $script:LogLbl.Text = "Okunamadi: $path"; return }
    $script:LogTbl.BeginLoadData()
    for ($i = $lines.Count - 1; $i -ge 0; $i--) {
        if (-not $lines[$i]) { continue }
        $p = $lines[$i].Split("`t", 6)
        $row = $script:LogTbl.NewRow()
        for ($c = 0; $c -lt [Math]::Min(6, $p.Count); $c++) { $row[$c] = $p[$c] }
        $script:LogTbl.Rows.Add($row)
    }
    $script:LogTbl.EndLoadData()
    Update-LogFilter
}

function Update-LogFilter {
    Set-GridFilter -Source $script:LogBs -ColMap $script:LogColMap -ColumnHeader '' -Text $script:LogTxt.Text
    $script:LogLbl.Text = "$($script:LogBs.Count) / $($script:LogTbl.Rows.Count) kayit"
}

function Show-LogViewer {
    $dir = Get-DataDir
    $files = @(Get-ChildItem -LiteralPath $dir -Filter 'ClientTool_*.log' -ErrorAction SilentlyContinue | Sort-Object Name -Descending | ForEach-Object { $_.Name })
    if ($files.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Henuz islem kaydi yok.`n$dir", 'Bilgi') | Out-Null; return }

    $dlg = New-Ctl Form 'Islem Kayitlari (denetim kaydi)' 0 0 1150 650 @{ StartPosition = 'CenterParent'; MinimumSize = (New-Object System.Drawing.Size(800, 400)); ShowInTaskbar = $false }
    $top = New-Ctl Panel '' 0 0 0 0 @{ Dock = 'Top'; Height = 36 }
    $script:LogCbo = New-Ctl ComboBox '' 36 7 170 23 @{ DropDownStyle = 'DropDownList' }
    foreach ($f in $files) { [void]$script:LogCbo.Items.Add($f) }
    $script:LogTxt = New-Ctl TextBox '' 262 7 260 23
    $btnRef = New-Ctl Button 'Yenile' 530 5 75 27
    $btnDir = New-Ctl Button 'Klasoru Ac' 611 5 90 27
    $script:LogLbl = New-Ctl Label '' 712 11 0 0 $gray
    $top.Controls.AddRange(@((New-Ctl Label 'Ay:' 6 11), $script:LogCbo, (New-Ctl Label 'Filtre:' 214 11), $script:LogTxt, $btnRef, $btnDir, $script:LogLbl))

    $script:LogTbl = New-Object System.Data.DataTable
    $script:LogTbl.Locale = [System.Globalization.CultureInfo]::InvariantCulture
    $cols = @(@('Zaman', 11), @('Kullanici', 8), @('Hostname', 10), @('Islem', 18), @('Durum', 5), @('Aciklama', 48))
    foreach ($c in $cols) { [void]$script:LogTbl.Columns.Add($c[0]) }
    $script:LogBs = New-Object System.Windows.Forms.BindingSource
    $script:LogBs.DataSource = $script:LogTbl
    $grid = New-ResultGrid
    $grid.AutoSizeColumnsMode = 'Fill'
    foreach ($c in $cols) { [void]$grid.Columns.Add((New-GridColumn $c[0] $c[0] $c[1])) }
    $grid.DataSource = $script:LogBs
    $script:LogColMap = Get-ColumnMap -Grid $grid
    $grid.add_CellFormatting({
        param($s, $e)
        if ($e.RowIndex -lt 0) { return }
        switch ([string]$s.Rows[$e.RowIndex].Cells['Durum'].Value) {
            'fail' { $e.CellStyle.BackColor = [System.Drawing.Color]::LightCoral }
            'warn' { $e.CellStyle.BackColor = [System.Drawing.Color]::LightYellow }
        }
    })

    $script:LogCbo.Add_SelectedIndexChanged({ Import-LogView })
    $script:LogTxt.Add_TextChanged({ Update-LogFilter })
    $btnRef.Add_Click({ Import-LogView })
    $btnDir.Add_Click({ try { Invoke-Item -LiteralPath (Get-DataDir) } catch { } })

    $dlg.Controls.AddRange(@($grid, $top))
    $script:LogCbo.SelectedIndex = 0
    [void]$dlg.ShowDialog($form)
    $dlg.Dispose()
}

# ---------- Sag tik islemleri (secili satirlar) ----------

function Get-GridPanel { param($Grid) if ($Grid -eq $dgvParse) { return 2 } return 1 }

function Start-GridCheck {
    # Sag tik > Durumu Kontrol Et: secili satirlarin cihazlari icin kontrol (tablo silinmez, satirlar yerinde guncellenir).
    # Panel 1: DNS / Online / AD / servis. Panel 2: DNS / Online / AD / servis / SCCM kaydi.
    param($Grid)
    $panel = Get-GridPanel $Grid
    if (Test-PanelBusy -Panel $panel) { return }
    $hosts = @(Get-SelectedHosts -Grid $Grid)
    if ($hosts.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Secili satirlarda gecerli bir cihaz adi yok.', 'Bilgi') | Out-Null; return }
    if ($panel -eq 2) { Start-ParserChecks -Hosts $hosts; return }
    $svc = @(Get-ServiceList)
    if ($svc.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('En az bir servis ismi girin.', 'Uyari') | Out-Null; return }
    Start-Check -Hosts $hosts -Services $svc
}

function Add-FixReport {
    # Calisan servis baslatma isinin raporuna host satiri ekler (is bitince Show-FixResultDialog ile gosterilir).
    param([string]$Hostname, [string]$Status, $Fix)
    $run = $script:CurrentRun
    if (-not $run) { return }
    if (-not $run.PSObject.Properties['Report']) { $run | Add-Member -NotePropertyName Report -NotePropertyValue (New-Object System.Collections.ArrayList) }
    $detail = if ($Fix) { (($Fix.Detail -split "`r`n" | Select-Object -Skip 1 | ForEach-Object { $_.Trim() }) -join ' | ') } else { '' }
    [void]$run.Report.Add([pscustomobject]@{ Hostname = $Hostname; Status = $Status; Ayrinti = $detail })
}

function Show-PendingReports {
    # Biten servis baslatma islerinin raporlarini sirayla gosterir (ic ice pencere acilmaz).
    if ($script:ShowingReport) { return }
    $script:ShowingReport = $true
    try {
        while ($script:ReportQueue.Count -gt 0) {
            $st = $script:ReportQueue[0]
            $script:ReportQueue.RemoveAt(0)
            Show-FixResultDialog -Title "$($st.Title) - Sonuc Raporu (Panel $($st.Panel))" -Data @($st.Report | Sort-Object Hostname)
        }
    }
    finally { $script:ShowingReport = $false }
}

function Show-FixResultDialog {
    # Servis baslatma sonuc raporu (ServisKontrolAraci ile ayni bicim): Hostname | Status (+ servis bazli Ayrinti).
    # "Panoya Kopyala (Tablo)": tab ayracli -> Word / Outlook / Excel'e tablo olarak yapisir. "CSV Olarak Disa Aktar".
    param([string]$Title, [object[]]$Data)
    $dlg = New-Ctl Form $Title 0 0 1000 520 @{ StartPosition = 'CenterParent'; MinimumSize = (New-Object System.Drawing.Size(700, 350)); ShowInTaskbar = $false }
    $grid = New-Ctl DataGridView '' 10 10 965 420 @{
        Anchor = 'Top,Bottom,Left,Right'; ReadOnly = $true; AllowUserToAddRows = $false; AllowUserToDeleteRows = $false
        AutoSizeColumnsMode = 'Fill'; RowHeadersVisible = $false; SelectionMode = 'CellSelect'; ClipboardCopyMode = 'EnableWithoutHeaderText'
        BackgroundColor = [System.Drawing.Color]::White
    }
    foreach ($c in @(@('Hostname', 18), @('Status', 47), @('Ayrinti', 35))) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c[0]; $col.HeaderText = $c[0]; $col.FillWeight = $c[1]
        [void]$grid.Columns.Add($col)
    }
    foreach ($d in $Data) {
        $i = $grid.Rows.Add([string]$d.Hostname, [string]$d.Status, [string]$d.Ayrinti)
        $st = [string]$d.Status
        $grid.Rows[$i].DefaultCellStyle.BackColor = $(
            if ($st -like '*kaldi*' -or $st -like 'Islem hatasi*' -or $st -like 'Iptal*' -or $st -like '*erisilemedi*') { [System.Drawing.Color]::LightCoral }
            elseif ($st -like '*bulunamadi*') { [System.Drawing.Color]::LightYellow }
            else { [System.Drawing.Color]::LightGreen })
    }
    $script:ReportData = $Data
    $btnCopy  = New-Ctl Button 'Panoya Kopyala (Tablo)' 10 440 180 30 @{ Anchor = 'Bottom,Left' }
    $btnCsv   = New-Ctl Button 'CSV Olarak Disa Aktar' 200 440 180 30 @{ Anchor = 'Bottom,Left' }
    $btnClose = New-Ctl Button 'Kapat' 875 440 100 30 @{ Anchor = 'Bottom,Right'; DialogResult = [System.Windows.Forms.DialogResult]::Cancel }
    $btnCopy.Add_Click({
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("Hostname`tStatus`tAyrinti")
        foreach ($d in $script:ReportData) { [void]$sb.AppendLine("$($d.Hostname)`t$($d.Status)`t$($d.Ayrinti)") }
        Copy-ToClipboard $sb.ToString()
        [System.Windows.Forms.MessageBox]::Show('Tablo panoya kopyalandi. Word / Outlook / Excel gibi programlara yapistirabilirsiniz.', 'Bilgi') | Out-Null
    })
    $btnCsv.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = 'CSV Dosyasi (*.csv)|*.csv'
        $sfd.FileName = "ServisBaslatmaRaporu_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($sfd.ShowDialog() -ne 'OK') { return }
        try {
            # Not: string[] pipeline'dan gecerse acilir; bu yuzden nesne olarak verilir
            Write-FastCsv -Path $sfd.FileName -Header @('Hostname', 'Status', 'Ayrinti') -Rows @($script:ReportData | ForEach-Object {
                    [pscustomobject]@{ Hostname = $_.Hostname; Status = (Protect-CellText $_.Status); Ayrinti = (Protect-CellText $_.Ayrinti) } })
            [System.Windows.Forms.MessageBox]::Show("Disa aktarildi: $($sfd.FileName)", 'Bilgi') | Out-Null
        }
        catch { [System.Windows.Forms.MessageBox]::Show("CSV yazilamadi: $($_.Exception.Message)", 'Hata') | Out-Null }
    })
    $dlg.CancelButton = $btnClose
    $dlg.Controls.AddRange(@($grid, $btnCopy, $btnCsv, $btnClose))
    [void]$dlg.ShowDialog($form)
    $dlg.Dispose()
}

function Start-GridRepair {
    # Sag tik > Istemci Onarimi: secili cihazlarda SMS_Client.RepairClient (ccmrepair). Sonuc istegin geldigi panele yazilir.
    param($Grid)
    $panel = Get-GridPanel $Grid
    if (Test-PanelBusy -Panel $panel) { return }
    $hosts = @(Get-SelectedHosts -Grid $Grid)
    if ($hosts.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Secili satirlarda gecerli bir cihaz adi yok.', 'Bilgi') | Out-Null; return }
    $c = [System.Windows.Forms.MessageBox]::Show(
        "Istemci onarimi (ccmrepair) - Panel ${panel}, secili $($hosts.Count) cihaz:`n$(Format-HostPreview -Hosts $hosts)`n`nccmrepair hedefte arka planda calisir (birkac dakika); bu sirada SCCM ajani yeniden baslar ve eylem / script almaz.`n`nBaslatilsin mi?",
        'Onay - Istemci Onarimi', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($c -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    $name = 'Istemci Onarimi (RepairClient / ccmrepair)'
    $items = @(foreach ($h in $hosts) { [pscustomobject]@{ Key = $h; Params = @{ ComputerName = $h; ActionName = $name; ScheduleIds = [string[]]@('@REPAIR') } } })
    if ($panel -eq 2) {
        foreach ($h in $hosts) { Set-ParserCheckValues -HostName $h -Values @{ CcmExec_Durumu = $script:PendingText } }
        Start-ParallelRun -Title 'Istemci onarimi (Parser)' -Worker $script:ActionWorker -Items $items -OnResultFn 'Update-ParserFromAction' -Panel 2
    }
    else {
        Set-GridPending -Hosts $hosts -Columns @('ActionResult')
        Start-ParallelRun -Title 'Istemci onarimi' -Worker $script:ActionWorker -Items $items -OnResultFn 'Update-GridFromAction'
    }
}

function Update-ParserFromAction {
    # Panel 2'den baslatilan istemci eylemi (ccmrepair) sonucu: CcmExec_Durumu kolonuna + detaya yazilir. Panel 1'e DOKUNMAZ.
    param($Key, $Output, $ErrorText)
    $r = @($Output | Where-Object { $_ -and $_.PSObject.Properties['Action'] })[0]
    if (-not $r) { $r = [pscustomobject]@{ Hostname = $Key; Action = 'Istemci onarimi'; State = 'fail'; Result = (Format-JobError $ErrorText); Detail = '' } }
    $one = ("$($r.Result)" -replace '\s+', ' ').Trim()
    Set-ParserCheckValues -HostName $Key -Values @{ CcmExec_Durumu = "Onarim: $one" }
    Add-ParserDetail -Hostname $Key -Text ("$($r.Action): $($r.Result)" + $(if ($r.Detail) { "`r`n$($r.Detail)" } else { '' }))
    Write-ActionLog -Hostname $Key -Action "$($r.Action) (Parser)" -State $r.State -Text $r.Result
}

# ==================== ANA FORM ====================

function New-Ctl {
    # Kisa WinForms kontrol olusturucu: New-Ctl <Tip> <Metin> x y [w h] @{ Ozellik = Deger }
    # w = 0 ise Label / CheckBox AutoSize olur.
    param([string]$Type, [string]$Text = '', [int]$X = 0, [int]$Y = 0, [int]$W = 0, [int]$H = 0, [hashtable]$P = @{})
    $c = New-Object "System.Windows.Forms.$Type"
    if ($Text) { $c.Text = $Text }
    $c.Location = New-Object System.Drawing.Point($X, $Y)
    if ($W -gt 0) { $c.Size = New-Object System.Drawing.Size($W, $H) }
    elseif ($c -is [System.Windows.Forms.Label] -or $c -is [System.Windows.Forms.CheckBox]) { $c.AutoSize = $true }
    foreach ($k in $P.Keys) { $c.$k = $P[$k] }
    return $c
}
$gray = @{ ForeColor = [System.Drawing.Color]::DimGray }

$form = New-Ctl Form 'Istemci Durum ve SCCM Islem Araci' 0 0 1360 900 @{ StartPosition = 'CenterScreen'; MinimumSize = (New-Object System.Drawing.Size(1150, 700)) }

$lblComputers = New-Ctl Label 'Bilgisayar Isimleri (her satira bir tane):' 10 10
$txtComputers = New-Ctl TextBox '' 10 30 200 800 @{ Multiline = $true; ScrollBars = 'Vertical'; Anchor = 'Top,Bottom,Left' }

# ---------- 1) Kontrol ----------
$gbCheck     = New-Ctl GroupBox '1) Kontrol (DNS / Online / AD / Servis)' 220 8 520 184
$lblServices = New-Ctl Label 'Servisler (virgulle ayirin):' 10 22
$txtServices = New-Ctl TextBox 'CcmExec, BITS, wuauserv' 10 42 495 23
$btnCheck    = New-Ctl Button 'Kontrol Et' 10 78 110 32
$btnAd       = New-Ctl Button 'AD Yapisi Kontrolu' 130 78 150 32
$btnFix      = New-Ctl Button 'Servisleri Baslat' 290 78 150 32 @{ Enabled = $false }
$btnAdCache  = New-Ctl Button 'AD Onbellegi...' 10 122 165 26
$lblAdCache  = New-Ctl Label 'Onbellek: -' 182 128 328 18 $gray
$lblAdHint   = New-Ctl Label 'AD bilgisi onbellekten (CSV) okunur; Online = ping + cihaz adi dogrulamasi.' 10 158 500 16 $gray
$gbCheck.Controls.AddRange(@($lblServices, $txtServices, $btnCheck, $btnAd, $btnFix, $btnAdCache, $lblAdCache, $lblAdHint))

# ---------- 2) SCCM Script ----------
$gbSccm        = New-Ctl GroupBox '2) SCCM Script' 750 8 580 184
$lblSms        = New-Ctl Label 'SMS Provider (Site Server) adi:' 10 22
$txtSms        = New-Ctl TextBox '' 10 42 300 23
$btnConnect    = New-Ctl Button 'Baglan / Scriptleri Listele' 320 40 170 27
$lblSite       = New-Ctl Label 'Site: -' 500 46
$lblScript     = New-Ctl Label 'Script:' 10 74
$cboScripts    = New-Ctl ComboBox '' 10 94 410 23 @{ DropDownStyle = 'DropDownList' }
$btnShowScript = New-Ctl Button 'Icerik' 428 92 60 27 @{ Enabled = $false }
$lblTimeout    = New-Ctl Label 'Bekleme (sn):' 498 74
$nudTimeout    = New-Ctl NumericUpDown '' 498 94 65 23 @{ Minimum = 30; Maximum = 900 }
$nudTimeout.Value = 90
$btnSendScript = New-Ctl Button 'Secili Scripti Gonder (Panel 1)' 10 124 230 28 @{ Enabled = $false }
$lblScriptInfo = New-Ctl Label '' 250 131 0 0 $gray
$lblIndep      = New-Ctl Label "Panel 1 ve Panel 2 bagimsizdir: bu buton sadece Panel 1'e yazar. Panel 2 icin Panel 2'deki 'Script Gonder (Panel 2)' butonunu kullanin." 10 154 560 28 $gray
$gbSccm.Controls.AddRange(@($lblSms, $txtSms, $btnConnect, $lblSite, $lblScript, $cboScripts, $btnShowScript, $lblTimeout, $nudTimeout, $btnSendScript, $lblScriptInfo, $lblIndep))

# ---------- 3) SCCM Istemci Eylemi ----------
$gbActions = New-Ctl GroupBox '3) SCCM Istemci Eylemi' 220 200 1110 60
$cboAction = New-Ctl ComboBox '' 10 25 400 23 @{ DropDownStyle = 'DropDownList' }
foreach ($k in $script:ClientActions.Keys) { [void]$cboAction.Items.Add($k) }
$cboAction.SelectedIndex = 0
$btnAction = New-Ctl Button 'Eylemi Gonder' 420 22 170 28
$lblHint   = New-Ctl Label 'Hedef: gridde secili satirlar; yoksa soldaki liste.' 605 29 0 0 $gray
$btnLog    = New-Ctl Button 'Islem Kayitlari...' 980 22 120 28
$gbActions.Controls.AddRange(@($cboAction, $btnAction, $lblHint, $btnLog))

$btnCancel1 = New-Ctl Button 'Iptal Et' 220 266 90 24 @{ Enabled = $false }
$lblStatus  = New-Ctl Label '' 318 270 0 0 @{ ForeColor = [System.Drawing.Color]::DarkBlue }

# ---------- Tablo yardimcilari (Panel 1 + Panel 2 ortak) ----------

function New-GridColumn {
    param([string]$Name, [string]$Header, [float]$Weight, [bool]$Visible = $true, [int]$Width = 0)
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $Name
    $col.DataPropertyName = $Name
    $col.HeaderText = $Header
    $col.FillWeight = $Weight
    $col.Visible = $Visible
    if ($Width -gt 0) { $col.Width = $Width }
    return $col
}

function New-ResultGrid {
    # Hucre secilebilir sonuc tablosu: hucreye tiklamak hucreyi, satir basligina tiklamak tum satiri secer.
    # Ctrl+C secili hucre(ler)i (basliksiz) kopyalar; Ctrl+Shift+C secili satirlari basliklariyla kopyalar.
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = 'Fill'
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.RowHeadersVisible = $true
    $g.RowHeadersWidth = 26
    $g.RowHeadersWidthSizeMode = 'DisableResizing'
    $g.SelectionMode = 'RowHeaderSelect'
    $g.MultiSelect = $true
    $g.ClipboardCopyMode = 'EnableWithoutHeaderText'
    $g.AutoGenerateColumns = $false
    $g.ContextMenuStrip = $script:GridMenu
    return $g
}

function Copy-ToClipboard {
    param($Data)
    if ($null -eq $Data -or ($Data -is [string] -and $Data.Length -eq 0)) { return }
    try { [System.Windows.Forms.Clipboard]::SetDataObject($Data, $true, 10, 100) }
    catch { [System.Windows.Forms.MessageBox]::Show("Panoya kopyalanamadi: $($_.Exception.Message)", 'Uyari') | Out-Null }
}

function Get-SelectedRowIndexes {
    param($Grid)
    $idx = @($Grid.SelectedCells | ForEach-Object { $_.RowIndex } | Sort-Object -Unique)
    if ($idx.Count -eq 0 -and $Grid.CurrentCell) { $idx = @($Grid.CurrentCell.RowIndex) }
    return $idx
}

function Get-SelectedHosts {
    # Secili hucrelerin bulundugu satirlardaki gecerli, tekil Hostname'ler
    param($Grid)
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $out  = New-Object 'System.Collections.Generic.List[string]'
    foreach ($i in @($Grid.SelectedCells | ForEach-Object { $_.RowIndex } | Sort-Object -Unique)) {
        $h = [string]$Grid.Rows[$i].Cells['Hostname'].Value
        if ($h -and $h -match $script:HostRegex -and $seen.Add($h)) { $out.Add($h) }
    }
    return $out.ToArray()
}

function Copy-GridData {
    # Mode: Cells (secili hucreler) | Rows (secili satirlar + baslik) | Column (aktif sutun, gorunen tum satirlar) | Hosts
    param($Grid, [string]$Mode)
    $cols = @($Grid.Columns | Where-Object { $_.Visible } | Sort-Object DisplayIndex)
    switch ($Mode) {
        'Cells' { Copy-ToClipboard $Grid.GetClipboardContent() }
        'Rows' {
            $lines = New-Object 'System.Collections.Generic.List[string]'
            $lines.Add((($cols | ForEach-Object { $_.HeaderText }) -join "`t"))
            foreach ($i in (Get-SelectedRowIndexes $Grid)) {
                $lines.Add((($cols | ForEach-Object { ([string]$Grid.Rows[$i].Cells[$_.Index].Value) -replace '[\t\r\n]+', ' ' }) -join "`t"))
            }
            if ($lines.Count -gt 1) { Copy-ToClipboard ($lines -join "`r`n") }
        }
        'Column' {
            if (-not $Grid.CurrentCell) { return }
            $ci = $Grid.CurrentCell.ColumnIndex
            Copy-ToClipboard ((@($Grid.Rows | ForEach-Object { [string]$_.Cells[$ci].Value }) -join "`r`n"))
        }
        'Hosts' {
            $h = @(Get-SelectedHosts $Grid)
            if ($h.Count -eq 0) { $h = @($Grid.Rows | ForEach-Object { [string]$_.Cells['Hostname'].Value } | Where-Object { $_ } | Select-Object -Unique) }
            Copy-ToClipboard ($h -join "`r`n")
        }
    }
}

function Show-RowDetail {
    # Panel 1: $script:Details, Panel 2: $script:ParserDetails (Hostname, yoksa SourceHost)
    param($Grid, [int]$RowIndex)
    if ($RowIndex -lt 0 -or $RowIndex -ge $Grid.Rows.Count) { return }
    $rv = $Grid.Rows[$RowIndex].DataBoundItem
    if (-not $rv) { return }
    $hn = [string]$rv['Hostname']
    $map = if ($Grid -eq $dgvParse) { $script:ParserDetails } else { $script:Details }
    $txt = '(Detay yok)'
    if ($map.ContainsKey($hn)) { $txt = $map[$hn] }
    elseif ($Grid -eq $dgvParse -and $map.ContainsKey([string]$rv['SourceHost'])) { $txt = $map[[string]$rv['SourceHost']] }
    Show-TextDialog -Title "Detay: $hn" -Text $txt
}

# Sag tik menusu (iki tablo icin tek menu; hangi tablodan acildigi SourceControl'den alinir)
$script:GridMenu = New-Object System.Windows.Forms.ContextMenuStrip
# Islem ogeleri (sadece Panel 1 / Panel 2 tablolarinda; panel mesgulken pasif). Hedef: secili satirlar.
$script:GridOps = @()
foreach ($mi in @(
        @('Durumu Kontrol Et (secili satirlar)', { Start-GridCheck -Grid $script:GridMenu.SourceControl }),
        @('Servisleri Baslat (secili satirlar)', { $g = $script:GridMenu.SourceControl; if ($g.SelectedCells.Count -gt 0) { Start-ServiceFix -Panel (Get-GridPanel $g) } }),
        @('Istemci Onarimi - ccmrepair (secili satirlar)', { Start-GridRepair -Grid $script:GridMenu.SourceControl })
    )) {
    $item = New-Object System.Windows.Forms.ToolStripMenuItem ($mi[0])
    $item.Font = New-Object System.Drawing.Font($item.Font, [System.Drawing.FontStyle]::Bold)
    $item.Add_Click($mi[1])
    [void]$script:GridMenu.Items.Add($item)
    $script:GridOps += $item
}
$script:GridOpsSep = New-Object System.Windows.Forms.ToolStripSeparator
[void]$script:GridMenu.Items.Add($script:GridOpsSep)
$script:GridMenu.Add_Opening({
    $g = $script:GridMenu.SourceControl
    $main = ($g -eq $dgv -or $g -eq $dgvParse)
    $busy = $main -and (Test-PanelBusy -Panel $(if ($g -eq $dgvParse) { 2 } else { 1 }))
    foreach ($it in $script:GridOps) { $it.Visible = $main; $it.Enabled = -not $busy }
    $script:GridOpsSep.Visible = $main
})
foreach ($mi in @(
        @('Kopyala (Ctrl+C)', { Copy-GridData -Grid $script:GridMenu.SourceControl -Mode 'Cells' }),
        @('Secili satirlari basliklariyla kopyala (Ctrl+Shift+C)', { Copy-GridData -Grid $script:GridMenu.SourceControl -Mode 'Rows' }),
        @('Bu sutunu kopyala (gorunen tum satirlar)', { Copy-GridData -Grid $script:GridMenu.SourceControl -Mode 'Column' }),
        @('Hostname listesini kopyala', { Copy-GridData -Grid $script:GridMenu.SourceControl -Mode 'Hosts' }),
        @('-', $null),
        @('Hucre degerini goster (kismi secip kopyalamak icin)', {
            $g = $script:GridMenu.SourceControl
            if ($g.CurrentCell) { Show-TextDialog -Title "$($g.Columns[$g.CurrentCell.ColumnIndex].HeaderText): $([string]$g.Rows[$g.CurrentCell.RowIndex].Cells['Hostname'].Value)" -Text ([string]$g.CurrentCell.Value) }
        }),
        @('Satir detayi (cift tik)', { $g = $script:GridMenu.SourceControl; if ($g.CurrentCell) { Show-RowDetail -Grid $g -RowIndex $g.CurrentCell.RowIndex } })
    )) {
    if ($mi[0] -eq '-') { [void]$script:GridMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)); continue }
    $item = New-Object System.Windows.Forms.ToolStripMenuItem ($mi[0])
    $item.Add_Click($mi[1])
    [void]$script:GridMenu.Items.Add($item)
}

# ---------- Panel 1 sonuc tablosu ----------
$dgv = New-ResultGrid
$dgv.AutoSizeColumnsMode = 'Fill'
$dgv.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@(
    (New-GridColumn 'Hostname'     'Hostname'              12),
    (New-GridColumn 'DnsStatus'    'DNS'                   13),
    (New-GridColumn 'IsOnline'     'Online mi'              8),
    (New-GridColumn 'AdStatus'     'AD Durumu'              8),
    (New-GridColumn 'LastLogon'    'AD Last Logon'         11),
    (New-GridColumn 'Running'      'Calisan Servisler'     10),
    (New-GridColumn 'NotRunning'   'Calismayan Servisler'  11),
    (New-GridColumn 'CcmExec'      'CcmExec'                7),
    (New-GridColumn 'BoundaryGroups' 'Boundary Group(s)'   11),
    (New-GridColumn 'OsBuild'      'Operating System Build' 9),
    (New-GridColumn 'LastAction'   'Son Islem'             13),
    (New-GridColumn 'ActionResult' 'Islem Sonucu / Script Ciktisi' 30),
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
        'fail'    { $e.CellStyle.BackColor = [System.Drawing.Color]::LightCoral }
        'warn'    { $e.CellStyle.BackColor = [System.Drawing.Color]::LightYellow }
        'ok'      { $e.CellStyle.BackColor = [System.Drawing.Color]::LightGreen }
        'pending' { $e.CellStyle.BackColor = [System.Drawing.Color]::Gainsboro }
    }
})

# ==================== SEKMELER: PANEL 1 (Kontrol / Script Ciktilari)  +  PANEL 2 (Update Parser) ====================

$tabs = New-Ctl TabControl '' 220 294 1110 541 @{ Anchor = 'Top,Bottom,Left,Right' }

$tabOut   = New-Ctl TabPage 'Panel 1 - Script Ciktilari / Kontroller'
$tabParse = New-Ctl TabPage 'Panel 2 - Update Parser'

# ---------- Panel 1 arac cubugu: filtre + Excel + CSV ----------
$pnlTop1     = New-Ctl Panel '' 0 0 0 0 @{ Dock = 'Top'; Height = 66 }
$lblF1       = New-Ctl Label 'Filtre:' 6 11
$cboF1       = New-Ctl ComboBox '' 50 7 150 23 @{ DropDownStyle = 'DropDownList' }
$txtF1       = New-Ctl TextBox '' 206 7 200 23
$btnClearF1  = New-Ctl Button 'Temizle' 412 5 65 27
$btnExcel1   = New-Ctl Button "Excel'e Aktar" 490 5 110 27
$btnCsv1     = New-Ctl Button 'CSV' 606 5 70 27
$btnUnfocus1 = New-Ctl Button 'Focus Kaybet' 688 5 124 27
$lblCount1   = New-Ctl Label '' 822 11 0 0 $gray
$pnlTop1.Controls.AddRange(@($lblF1, $cboF1, $txtF1, $btnClearF1, $btnExcel1, $btnCsv1, $btnUnfocus1, $lblCount1))
New-SummaryBar -Panel 1 -Parent $pnlTop1

# ---------- Panel 2: Update Parser (Panel 1'den BAGIMSIZ) ----------
$lblParserInfo = New-Ctl Label "Update Parser (Panel 2 - Panel 1'den bagimsiz)  |  0 satir, 0 cihaz, 0 kolon uyumsuz / bilgi satiri" 0 0 0 0 @{
    AutoSize = $false; Dock = 'Top'; Height = 24; TextAlign = 'MiddleLeft'; ForeColor = [System.Drawing.Color]::DarkGreen
    Font = (New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold))
}
$lblParserStatus = New-Ctl Label '' 0 0 0 0 @{ AutoSize = $false; Dock = 'Top'; Height = 20; TextAlign = 'MiddleLeft'; ForeColor = [System.Drawing.Color]::DarkBlue }

# -- Yapistirma alani (eski client_update_parser.ps1 mantigi)
$pnlPaste = New-Ctl Panel '' 0 0 0 0 @{ Dock = 'Top'; Height = 100 }
$lblPaste = New-Ctl Label "SCCM Run Script ciktisini yapistirin ('@' ile ayrilan her makine kaydi alt alta yazilir; her satir ';' ile ayrilmis):" 0 0 0 0 @{ AutoSize = $false; Dock = 'Top'; Height = 20; TextAlign = 'MiddleLeft' }
$txtParsePaste = New-Ctl TextBox '' 0 0 0 0 @{
    Multiline = $true; AcceptsReturn = $true; ScrollBars = 'Both'; WordWrap = $false; MaxLength = [int]::MaxValue
    Font = (New-Object System.Drawing.Font('Consolas', 9)); Dock = 'Fill'
}
$pnlPaste.Controls.AddRange(@($txtParsePaste, $lblPaste))   # Dock sirasi: Fill kontrol ILK, Top kontrol sonra

# -- Calistirma cubugu
$pnlRun2         = New-Ctl Panel '' 0 0 0 0 @{ Dock = 'Top'; Height = 36 }
$btnParseRun     = New-Ctl Button 'Parse Et' 6 5 100 27
$chkParseCheck   = New-Ctl CheckBox 'DNS / AD / servis / SCCM ajan kontrolu ekle' 114 9 0 0 @{ Checked = $true }
$chkParseAppend  = New-Ctl CheckBox 'Tabloya ekle' 398 9
$btnParseScript  = New-Ctl Button 'Script Gonder (Panel 2)' 500 5 165 27 @{ Enabled = $false }
$btnParseRecheck = New-Ctl Button 'Kontrolleri Calistir' 670 5 132 27
$btnParseFix     = New-Ctl Button 'Servisleri Baslat' 807 5 120 27 @{ Enabled = $false }
$btnParseCancel  = New-Ctl Button 'Iptal Et' 932 5 70 27 @{ Enabled = $false }
$btnParseClear   = New-Ctl Button 'Temizle' 1007 5 75 27
$pnlRun2.Controls.AddRange(@($btnParseRun, $chkParseCheck, $chkParseAppend, $btnParseScript, $btnParseRecheck, $btnParseFix, $btnParseCancel, $btnParseClear))

# -- Filtre cubugu
$pnlTop2     = New-Ctl Panel '' 0 0 0 0 @{ Dock = 'Top'; Height = 66 }
$lblF2       = New-Ctl Label 'Filtre:' 6 11
$cboF2       = New-Ctl ComboBox '' 50 7 200 23 @{ DropDownStyle = 'DropDownList' }
$txtF2       = New-Ctl TextBox '' 256 7 200 23
$btnClearF2  = New-Ctl Button 'Temizle' 462 5 65 27
$btnUnfocus2 = New-Ctl Button 'Focus Kaybet' 534 5 124 27
$lblCount2   = New-Ctl Label '' 676 11 0 0 $gray
$pnlTop2.Controls.AddRange(@($lblF2, $cboF2, $txtF2, $btnClearF2, $btnUnfocus2, $lblCount2))
New-SummaryBar -Panel 2 -Parent $pnlTop2

# -- Kayit cubugu (kayit klasoru + Gozat + Excel + CSV Olustur ve Ac)
$pnlSave2   = New-Ctl Panel '' 0 0 0 0 @{ Dock = 'Bottom'; Height = 36 }
$lblPath2   = New-Ctl Label 'Kayit klasoru:' 6 11
$txtPath2   = New-Ctl TextBox $script:ParserDefaultFolder 92 7 380 23
$btnBrowse2 = New-Ctl Button 'Gozat...' 480 5 70 27
$btnExcel2  = New-Ctl Button "Excel'e Aktar" 558 5 110 27
$btnCsv2    = New-Ctl Button 'CSV Olustur ve Ac' 676 5 140 27
$pnlSave2.Controls.AddRange(@($lblPath2, $txtPath2, $btnBrowse2, $btnExcel2, $btnCsv2))

$dgvParse = New-ResultGrid
$dgvParse.AutoSizeColumnsMode = 'None'
foreach ($cn in $script:ParserColumns) { [void]$dgvParse.Columns.Add((New-GridColumn $cn $cn 10 $true $(if ($cn -eq 'UpdateName') { 300 } else { 130 }))) }
$w2 = @{ SCCM_Kayit = 200; DNS_Durumu = 150; AD_LastLogon = 130; Calisan_Servisler = 180; Calismayan_Servisler = 200 }
foreach ($cn in $script:ParserCheckColumns) { [void]$dgvParse.Columns.Add((New-GridColumn $cn $cn 10 $true $(if ($w2.ContainsKey($cn)) { $w2[$cn] } else { 120 }))) }
[void]$dgvParse.Columns.Add((New-GridColumn 'Boundary_Groups' 'Boundary Group(s)' 10 $true 200))
[void]$dgvParse.Columns.Add((New-GridColumn 'OS_Build' 'Operating System Build' 10 $true 150))
[void]$dgvParse.Columns.Add((New-GridColumn 'KolonUyumu' 'KolonUyumu' 10 $true 170))
[void]$dgvParse.Columns.Add((New-GridColumn 'SourceHost' 'SourceHost' 10 $false))
[void]$dgvParse.Columns.Add((New-GridColumn 'Durum' 'Durum' 10 $false))
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

    switch ($cn) {
        'DNS_Durumu' {
            if ($v -like 'No such host*' -or $v -like 'DNS hatasi*' -or $v -like 'Islem hatasi*' -or $v -eq 'Iptal edildi') { $e.CellStyle.BackColor = $red }
            elseif ($v -and $v -notlike 'Bekleniyor*') { $e.CellStyle.BackColor = $grn }
        }
        'Online_mi' {
            if ($v -eq 'Evet') { $e.CellStyle.BackColor = $grn }
            elseif ($v -like 'Hayir*') { $e.CellStyle.BackColor = $red }
            elseif ($v) { $e.CellStyle.BackColor = $yel }
        }
        'AD_Durumu' {
            if ($v -eq 'Enabled') { $e.CellStyle.BackColor = $grn }
            elseif ($v -eq 'Disabled' -or $v -eq 'Bulunamadi') { $e.CellStyle.BackColor = $red }
            elseif ($v) { $e.CellStyle.BackColor = $yel }
        }
        'Calismayan_Servisler' {
            if ($v -and $v -notlike 'Kontrol edilemedi*' -and $v -notlike 'Bekleniyor*') { $e.CellStyle.BackColor = $red }
        }
        'CcmExec_Durumu' {
            if ($v -eq 'Running') { $e.CellStyle.BackColor = $grn }
            elseif ($v -eq 'Yok') { $e.CellStyle.BackColor = $red }
            elseif ($v) { $e.CellStyle.BackColor = $yel }
        }
        'SCCM_Kayit' {
            if ($v -like 'Client yuklu*') { $e.CellStyle.BackColor = $grn }
            elseif ($v -like 'SCCM de kayitli degil*' -or $v -like 'Kayit var*') { $e.CellStyle.BackColor = $red }
            elseif ($v) { $e.CellStyle.BackColor = $yel }
        }
    }
    if ($v -like 'Bekleniyor*') { $e.CellStyle.BackColor = [System.Drawing.Color]::Gainsboro }
})

# Iki tablo icin ortak olaylar: cift tik = detay, sag tik = hucreyi sec + menu, Ctrl+Shift+C = satirlari basliklariyla kopyala,
# Esc / bos alana tik = secimi birak (hedef tekrar soldaki liste olur)
foreach ($g in @($dgv, $dgvParse)) {
    $g.add_CellDoubleClick({ param($s, $e) if ($e.RowIndex -ge 0) { Show-RowDetail -Grid $s -RowIndex $e.RowIndex } })
    $g.add_CellMouseDown({
        param($s, $e)
        if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right -or $e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
        $cell = $s.Rows[$e.RowIndex].Cells[$e.ColumnIndex]
        if (-not $cell.Selected) { $s.ClearSelection(); $cell.Selected = $true }
        $s.CurrentCell = $cell
    })
    $g.Add_MouseDown({
        param($s, $e)
        if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
        if ($s.HitTest($e.X, $e.Y).Type -eq [System.Windows.Forms.DataGridViewHitTestType]::None) { Clear-GridFocus -Grid $s }
    })
    $g.Add_KeyDown({
        param($s, $e)
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { Clear-GridFocus -Grid $s; $e.Handled = $true }
        elseif ($e.Control -and $e.Shift -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) { Copy-GridData -Grid $s -Mode 'Rows'; $e.Handled = $true }
    })
}

# Dock sirasi: Fill kontrol ILK eklenir, Top kontroller sonra (en son eklenen en ustte durur)
$tabOut.Controls.AddRange(@($dgv, $pnlTop1))
$tabParse.Controls.AddRange(@($dgvParse, $pnlSave2, $pnlTop2, $pnlRun2, $pnlPaste, $lblParserStatus, $lblParserInfo))
[void]$tabs.TabPages.Add($tabOut)
[void]$tabs.TabPages.Add($tabParse)

# Filtre kolon listeleri
$script:ColMap1 = Get-ColumnMap -Grid $dgv
$script:ColMap2 = Get-ColumnMap -Grid $dgvParse
foreach ($p in @(@($cboF1, $script:ColMap1), @($cboF2, $script:ColMap2))) {
    [void]$p[0].Items.Add($script:AllColsLabel)
    foreach ($k in $p[1].Keys) { [void]$p[0].Items.Add($k) }
    $p[0].SelectedIndex = 0
}

$form.Controls.AddRange(@($lblComputers, $txtComputers, $gbCheck, $gbSccm, $gbActions, $btnCancel1, $lblStatus, $tabs))

$tip = New-Object System.Windows.Forms.ToolTip
$tip.SetToolTip($chkParseAppend, 'Isaretli degilse "Parse Et" mevcut tabloyu temizleyip yeni veriyi yazar (eski parser gibi).')
$tip.SetToolTip($btnParseScript, 'Ustteki 2) SCCM Script bolumundeki script secimini kullanir. Hedef: bu tabloda secili satirlar; yoksa soldaki liste. Sonuc SADECE Panel 2''ye yazilir.')
$tip.SetToolTip($btnParseRecheck, 'Tabloda secili satirlarin (yoksa tum cihazlarin) DNS / AD / servis / SCCM ajan kontrolunu yeniden calistirir. Servis listesi: 1) Kontrol bolumu.')
$tip.SetToolTip($btnAdCache, "Domain / OU eslemesini girdiginiz pencereyi acar; eslesen OU'lardaki bilgisayar hesaplarini AD'den okuyup diskteki CSV onbellegini olusturur. Kontroller AD'ye degil bu onbellege bakar; onbellek sadece kullanici isterse guncellenir.")
$tip.SetToolTip($btnUnfocus1, 'Tablodaki secimi ve aktif hucreyi kaldirir; hedef tekrar soldaki liste olur. (Tablonun bos alanina tiklamak veya Esc de ayni isi yapar.)')
$tip.SetToolTip($btnUnfocus2, 'Tablodaki secimi ve aktif hucreyi kaldirir; hedef tekrar soldaki liste olur. (Tablonun bos alanina tiklamak veya Esc de ayni isi yapar.)')
$tip.SetToolTip($btnParseFix, 'Son kontrolde durdurulmus servisi olan cevrimici cihazlarda servisleri baslatir. Tabloda secim varsa sadece secili cihazlar. Sonuc SADECE Panel 2''ye yazilir.')
$tip.SetToolTip($btnFix, 'Son kontrolde durdurulmus servisi olan cevrimici cihazlarda servisleri baslatir. Tabloda secim varsa sadece secili cihazlar.')
$tip.SetToolTip($btnCancel1, 'Panel 1''de bekleyen / calisan isleri durdurur. Tamamlanan sonuclar korunur; SCCM''e gonderilmis script / eylemler geri alinamaz.')
$tip.SetToolTip($btnParseCancel, 'Panel 2''de bekleyen / calisan isleri durdurur. Tamamlanan sonuclar korunur; SCCM''e gonderilmis scriptler geri alinamaz.')
$tip.SetToolTip($btnLog, "ClientTool_yyyyMM.log denetim kaydini (kim, ne zaman, hangi cihaza, hangi islem, sonuc) filtrelenebilir tabloda acar.")
foreach ($p in 1, 2) {
    $tip.SetToolTip($script:RetryButtons[$p], 'Bu paneldeki SON islemi, sonucu Hata (kirmizi) olan cihazlar icin ayni ayarlarla yeniden calistirir (iptal edilenler dahil).')
    $tip.SetToolTip($script:Auto[$p].Check, $(if ($p -eq 1) { 'Secilen araliklarla tablodaki tum cihazlar icin Kontrol (DNS / Online / AD / servis) tekrarlanir. Tablo silinmez.' } else { 'Secilen araliklarla tablodaki tum cihazlar icin DNS / AD / servis / SCCM kontrolu tekrarlanir. Script tekrar GONDERILMEZ.' }))
    foreach ($k in 'ok', 'warn', 'fail', 'pending', 'all') { $tip.SetToolTip($script:SumButtons[$p][$k], 'Tikla: tabloyu bu duruma filtrele (tekrar tikla / Tumu: filtreyi kaldir). Metin filtresiyle birlikte calisir.') }
}
$tip.SetToolTip($btnCheck, "Online mi: Evet = ping + cihaz adi dogrulandi | Hayir (IP baska cihazda) = DNS kaydi eski, IP'yi baska cihaz kullaniyor | Hayir (ping var, cihaz dogrulanamadi) = cevap veren cihazin adi okunamadi (gateway / joker DNS olabilir).")

# ==================== UI DURUM YONETIMI ====================

function Update-ButtonStates {
    $busy1 = Test-PanelBusy -Panel 1
    $busy2 = Test-PanelBusy -Panel 2

    $btnFix.Enabled = ((@($script:LastResults.Values | Where-Object { $_.IsOnline }).Count -gt 0) -and -not $busy1)

    $idx = $cboScripts.SelectedIndex
    $ok = $false
    if ($script:SiteCode -and $idx -ge 0 -and $idx -lt @($script:CmScripts).Count) {
        $ok = ($script:CmScripts[$idx].Approval -eq 3)
    }
    $btnSendScript.Enabled  = ($ok -and -not $busy1)     # Panel 1
    $btnParseScript.Enabled = ($ok -and -not $busy2)     # Panel 2
    $btnParseFix.Enabled = ((@($script:ParserSvc.Values | Where-Object { $_.IsOnline }).Count -gt 0) -and -not $busy2)
    $btnShowScript.Enabled = ($script:SiteCode -and $idx -ge 0)
}

function Set-UiBusy {
    # Panel bazli: Panel 1 mesgulken Panel 2 butonlari acik kalir (ve tersi).
    param([bool]$Busy, [int]$Panel = 1)

    if ($Panel -eq 2) {
        foreach ($b in @($btnParseRun, $btnParseScript, $btnParseRecheck, $btnParseFix)) { $b.Enabled = -not $Busy }
        $btnParseCancel.Enabled = $Busy
    }
    else {
        foreach ($b in @($btnCheck, $btnAd, $btnFix, $btnAction, $btnSendScript, $btnConnect, $btnShowScript, $btnAdCache)) { $b.Enabled = -not $Busy }
        $btnCancel1.Enabled = $Busy
    }
    if ($script:RetryButtons[$Panel] -and $Busy) { $script:RetryButtons[$Panel].Enabled = $false }
    $script:SumDirty[$Panel] = $true
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
        if ($s.ParserFolder) { $txtPath2.Text = [string]$s.ParserFolder }
        if ($null -ne $s.ParserCheck) { $chkParseCheck.Checked = [bool]$s.ParserCheck }
        if ($s.LastScriptGuid) { $script:PendingScriptGuid = [string]$s.LastScriptGuid }
        if ($null -ne $s.AdCacheSources) {   # AD onbellegi domain / OU eslemesi
            $script:AdCacheSources = @($s.AdCacheSources | Where-Object { $_ -and $_.Domain -and $_.OU } |
                ForEach-Object { [pscustomobject]@{ Domain = [string]$_.Domain; OU = [string]$_.OU } })
        }
    }
    catch { }
}

function Get-ServiceList { return @($txtServices.Text -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

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
            ParserFolder   = $txtPath2.Text.Trim()
            ParserCheck    = [bool]$chkParseCheck.Checked
            LastScriptGuid = $guid
            AdCacheSources = @($script:AdCacheSources | ForEach-Object { [pscustomobject]@{ Domain = [string]$_.Domain; OU = [string]$_.OU } })
        }
        Set-Content -Path (Get-SettingsFile) -Value ($obj | ConvertTo-Json -Depth 4) -Encoding UTF8
    }
    catch { }
}

# ==================== OLAY YONETICILERI ====================

# ---------- Filtreler (Panel 1 ve Panel 2) ----------
$txtF1.Add_TextChanged({ Update-Filter -Panel 1 })
$cboF1.Add_SelectedIndexChanged({ Update-Filter -Panel 1 })
$btnClearF1.Add_Click({ $txtF1.Text = ''; $cboF1.SelectedIndex = 0 })

$txtF2.Add_TextChanged({ Update-Filter -Panel 2 })
$cboF2.Add_SelectedIndexChanged({ Update-Filter -Panel 2 })
$btnClearF2.Add_Click({ $txtF2.Text = ''; $cboF2.SelectedIndex = 0 })

# Gosterilen / toplam satir sayaci (filtre uygulaninca ilk deger kuculur)
$script:BindingSrc.Add_ListChanged({ $lblCount1.Text = "Satir: $($script:BindingSrc.Count) / $($script:ResultsTable.Rows.Count)" })
$script:ParseBinding.Add_ListChanged({ $lblCount2.Text = "Satir: $($script:ParseBinding.Count) / $($script:ParseTable.Rows.Count)" })

# V7: ozet sayaclari - tablo degisince kirli isaretlenir, SummaryTimer en fazla 0,5 sn'de bir gunceller (UI'yi yormaz)
$script:ResultsTable.Add_RowChanged({ $script:SumDirty[1] = $true })
$script:ResultsTable.Add_RowDeleted({ $script:SumDirty[1] = $true })
$script:ResultsTable.Add_TableCleared({ $script:SumDirty[1] = $true })
$script:ParseTable.Add_RowChanged({ param($s, $e) [void]$script:DirtyRows.Add($e.Row); $script:SumDirty[2] = $true })
$script:ParseTable.Add_RowDeleted({ $script:SumDirty[2] = $true })
$script:ParseTable.Add_TableCleared({ $script:DirtyRows.Clear(); $script:SumDirty[2] = $true })
$script:SummaryTimer = New-Object System.Windows.Forms.Timer
$script:SummaryTimer.Interval = 500
$script:SummaryTimer.Add_Tick({ foreach ($p in 1, 2) { if ($script:SumDirty[$p]) { Update-SummaryNow -Panel $p } } })
$script:SummaryTimer.Start()

# V13: otomatik yenileme zamanlayicisi
$script:AutoTimer = New-Object System.Windows.Forms.Timer
$script:AutoTimer.Interval = 15000
$script:AutoTimer.Add_Tick({ Invoke-AutoRefresh })
$script:AutoTimer.Start()

# V14: islem kaydi goruntuleyici
$btnLog.Add_Click({ Show-LogViewer })

# ---------- Kontrol Et ----------
$btnCheck.Add_Click({
    $computers = @(Get-ComputerList)
    if ($computers.Count -eq 0) { return }
    $services = @(Get-ServiceList)
    if ($services.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('En az bir servis ismi girin.', 'Uyari') | Out-Null
        return
    }
    Start-Check -Hosts $computers -Services $services -Reset
})

# ---------- Servisleri Baslat ----------
$btnFix.Add_Click({ Start-ServiceFix -Panel 1 })

# ---------- AD Onbellegi: domain / OU eslemesi penceresi (guncelleme sadece kullanici isterse) ----------
$btnAdCache.Add_Click({ Show-AdCacheDialog })

# ---------- AD Yapisi Kontrolu (script gondermeden; DNS + ping + AD) ----------
$btnAd.Add_Click({
    $t = Get-TargetHosts -Grid $dgv
    if (-not $t) { return }

    $items = @(foreach ($h in $t.Hosts) {
        [pscustomobject]@{ Key = $h; Params = @{ ComputerName = $h } }
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

    $note = switch ($ids[0]) {
        '@REPAIR'  { "`n`nccmrepair hedefte arka planda calisir (birkac dakika); bu sirada SCCM ajani yeniden baslar ve eylem / script almaz." }
        '@RESTART' { "`n`nCcmExec (SMS Agent Host) durdurulup baslatilacak; o anda calisan SCCM islemleri (kurulum, tarama) kesilebilir." }
        default    { '' }
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Eylem: $actionName`n`nHedef ($($t.Source)): $(@($t.Hosts).Count) makine`n$(Format-HostPreview -Hosts $t.Hosts)$note`n`nGonderilsin mi?",
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
    Export-ToExcelAndOpen -Data (Get-GridExportData -Grid $dgv -Source $script:BindingSrc -WithDetail $true) -DefaultFileName (Get-ExportStamp)
})

$btnCsv1.Add_Click({
    if ($script:ResultsTable.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Aktarilacak sonuc yok.', 'Bilgi') | Out-Null
        return
    }
    Export-GridToCsv -Data (Get-GridExportData -Grid $dgv -Source $script:BindingSrc -WithDetail $true) -DefaultFileName (Get-ExportStamp)
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
    Export-ToExcelAndOpen -Data (Get-GridExportData -Grid $dgvParse -Source $script:ParseBinding -SkipColumns @('KolonUyumu')) -DefaultFileName (Get-ExportStamp)
})

$btnCsv2.Add_Click({
    # Kayit klasorune "<gun-ay-saat-dakika>.csv" yazar ve otomatik acar (Excel ile ayni zaman damgasi).
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

    $fullPath = Join-Path $folder ("{0}.csv" -f (Get-ExportStamp))
    $data = Get-GridExportData -Grid $dgvParse -Source $script:ParseBinding -SkipColumns @('KolonUyumu')
    try { Write-FastCsv -Path $fullPath -Header @($data[0].PSObject.Properties.Name) -Rows $data }
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
        $script:ParserNoOutput.Clear()
        $script:ParserDetails.Clear()
        $script:ParserSvc.Clear()
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
$btnParseFix.Add_Click({ Start-ServiceFix -Panel 2 })
$btnParseCancel.Add_Click({ Stop-PanelRun -Panel 2 })
$btnCancel1.Add_Click({ Stop-PanelRun -Panel 1 })

$btnParseRecheck.Add_Click({
    $hostsToCheck = @(Get-SelectedHosts -Grid $dgvParse)
    $src = 'secili satirlar'
    if ($hostsToCheck.Count -eq 0) {
        $src = 'tablodaki tum cihazlar'
        $seen = New-NameMap
        $hostsToCheck = @(foreach ($row in $script:ParseTable.Rows) {
            $h = [string]$row['Hostname']
            if ($h -and ($h -match $script:HostRegex) -and -not $seen.ContainsKey($h)) { $seen[$h] = 1; $h }
        })
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
    $script:ParserNoOutput.Clear()
    $script:ParserDetails.Clear()
    $script:ParserSvc.Clear()
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

# ==================== FORM KAPANIRKEN TEMIZLIK ====================

$form.Add_FormClosing({
    Save-Settings
    if ($script:RunTimer.Enabled) { $script:RunTimer.Stop() }
    $script:SummaryTimer.Stop(); $script:AutoTimer.Stop()
    foreach ($pf in @($script:Prefetch)) { try { [void]$pf.Job.PS.BeginStop($null, $null) } catch { } }
    # Calisan isler beklenmeden durdurulur (asenkron; uzun WMI cagrisi pencerenin kapanmasini geciktirmesin)
    foreach ($st in @($script:ActiveRuns)) { foreach ($j in $st.Jobs) { if (-not $j.Done) { try { [void]$j.PS.BeginStop($null, $null) } catch { } } } }
    foreach ($j in @(@($script:ScriptPoll | ForEach-Object { $_.Job }) + $script:PollJob)) { if ($j) { try { [void]$j.PS.BeginStop($null, $null) } catch { } } }
    if ($script:PollRunspace) { try { $script:PollRunspace.CloseAsync() } catch { } }
    if ($script:SharedPool) { try { [void]$script:SharedPool.BeginClose($null, $null) } catch { } }
})

# ==================== BASLANGIC ====================

Import-Settings
Update-AdCacheLabel
[void]$form.ShowDialog()
