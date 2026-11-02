#Requires -Version 5.1
<#
.SYNOPSIS
    Task Sequence Kontrol Araci (WinForms GUI) - client_sccm_tool.ps1 Panel 1 altyapisi uzerine.

.KULLANIM (kisa)
    1) SMS Provider adini yazip "Baglan" (ayar kaydedilir, sonraki acilista otomatik baglanir).
    2) "TS Tanimlari..." ile TS ekle (Package ID ve/veya Deployment ID; SCCM'den dogrulanir). Ustteki listeden aktif TS'i sec.
    3) "Koleksiyonlar..." ile koleksiyon ekle (ID veya ad; Device Collection olmali).
    4) Sol kutuya cihaz adlari (her satira bir tane). Hedef: gridde secili satirlar; secim yoksa soldaki listenin tamami.
       - Kontrol Et      : DNS / Online (ping + kimlik) / AD / SCCM kaydi / Boundary Group(s) / OS Build (SCCM) / CcmExec /
                           OS Build (anlik, istemci kayit defteri) / koleksiyon uyelikleri (+ aktif TS durumu).
       - TS Durumu Sorgula: ISTEMCIDEN okunur (WMI): kayit defteri Execution History\System\<PackageID> + TS politikasi + calisan TS.
       - Istemci Eylemi : client_sccm_tool Panel 1'deki eylemler (Machine Policy, Update Scan / Deployment, envanter, onarim...).
       - TS Tetikle      : istemcide Software Center'daki "Yukle" ile ayni cagri (ClientSDK CCM_ProgramsManager.ExecuteProgram).
       - Koleksiyon Ekle / Cikar: Direct uyelik kurali (SMS_CollectionRuleDirect).
    5) Yazma islemleri (Tetikle / Ekle / Cikar) ONAY ister, ust sinir (varsayilan 200) asilirsa calismaz. DRY-RUN isaretliyse
       hicbir sey yazilmaz, sadece ne yapilacagi gosterilir. Her yazma islemi denetim loguna (TsTool_yyyyMM.log) yazilir.
    6) Gridde sadece KISA / SABIT ifadeler var (Excel filtresi icin). Ham / uzun bilgi: satira CIFT TIK.
       TS ham durumu tek cihazda istemciden de okunabilir: sag tik > "TS istemci detayi".

.TS_DURUM
    Kaynak (istemci, WMI / DCOM): HKLM\SOFTWARE\Microsoft\SMS\Mobile Client\Software Distribution\Execution History\System\<PackageID>\<alt anahtar>
      (_State, _RunStartTime, SuccessOrFailureCode, SuccessOrFailureReason; birden fazla calisma varsa en yeni _RunStartTime) - StdRegProv ile okunur.
    HEDEF_DEGIL  istemcide bu TS icin politika yok ve calisma gecmisi yok
    BEKLIYOR     politika alindi / planli, baslamadi           CALISIYOR  TS calisiyor / yeniden baslatma bekliyor
    HATA         basarisiz (TS_HataKodu doldurulur)              BASARILI   basariyla tamamlandi
    DURUM_YOK    cihazdan okunamadi (cevrimdisi / erisim yok) ya da ham durum tanimsiz
    Eslesme tablosu: $script:TsStateMap (tek yer; ham _State degeri). Eslesmeyen: SuccessOrFailureCode 0 degilse HATA, aksi halde DURUM_YOK.
    Ham deger cift tik detayinda gorunur; farkli bir deger gorurseniz tabloya bir satir ekleyin.
    TS_Adim: kayit defterinde adim adi yok; CALISIYOR / HATA durumunda cihaz basina SMS Provider'dan okunur (SMS_TaskSequenceExecutionStatus).
      CALISIYOR: SCCM'e bildirilen SON TAMAMLANAN adim (istemci adimi BITINCE bildirir; o an calisan adim bir sonrakidir).
      HATA     : basarisiz adim (ExitCode <> 0).
      Bulunamazsa sabit neden: 'SCCM baglantisi yok' | 'SCCM kaydi yok' | 'Adim sorgulanamadi' | 'Adim henuz SCCM''de yok' (sunucuya gecikmeli yansir).
      Tum adimlar: sag tik > "TS adim gecmisi (SCCM)".

.MIMARI
    - WinRM / Invoke-Command / PSSession KULLANILMAZ. Sadece RPC / DCOM / WMI, ServiceController, NetBIOS (UDP 137).
    - Kimlik bilgisi sorulmaz (calistiran kullanicinin yetkisi). SCCM konsolu / ConfigurationManager modulu GEREKMEZ.
    - RunspacePool, SABIT en fazla 10 eszamanli hedef. Sonuclar sadece bellekte; kalici: ayarlar JSON + denetim logu.

.KAYITLAR  (%USERPROFILE%\Desktop\ServiceStatus\)
    TsToolSettings.json (ayarlar, TS ve koleksiyon tanimlari) | TsTool_yyyyMM.log (denetim kaydi)
#>

# WinForms ve pano STA ister; Get-WmiObject sadece Windows PowerShell 5.1'de var.
if ($PSVersionTable.PSEdition -eq 'Core' -or [System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $winPs = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if ($PSCommandPath -and (Test-Path -LiteralPath $winPs)) {
        Start-Process -FilePath $winPs -ArgumentList @('-NoProfile', '-STA', '-File', "`"$PSCommandPath`"")
        return
    }
}

Add-Type -AssemblyName System.Windows.Forms, System.Drawing, System.ServiceProcess, System.Management
[System.Windows.Forms.Application]::EnableVisualStyles()

#region Global degiskenler
function New-NameMap { return (New-Object System.Collections.Hashtable ([System.StringComparer]::OrdinalIgnoreCase)) }

$script:MaxParallel   = 10            # SABIT: ayni anda en fazla 10 hedef
$script:Details       = New-NameMap   # Hostname -> satir detay metni (cift tik)
$script:SmsServer     = ''
$script:SiteCode      = ''
$script:ConnectJob    = $null
$script:HostRegex     = '^[A-Za-z0-9İıŞşĞğÜüÖöÇç][A-Za-z0-9İıŞşĞğÜüÖöÇç._-]*$'
$script:SharedPool    = $null
$script:ActiveRuns    = New-Object System.Collections.ArrayList
$script:Prefetch      = New-Object System.Collections.ArrayList   # SMS Provider toplu sorgulari (snapshot)
$script:CancelledJobs = New-Object System.Collections.ArrayList
$script:AllColsLabel  = '(Tum kolonlar)'
$script:PendingText   = 'Bekleniyor..'
$script:CancelText    = '__IPTAL__'
$script:DeferFlag     = $false
$script:CurrentRun    = $null
$script:LastOp        = @{}
$script:StateFilter   = @{ 1 = '' }
$script:SumButtons    = @{}
$script:RetryButtons  = @{}
$script:SumDirty      = @{ 1 = $true }
$script:Auto          = @{}
$script:AdData        = $null
$script:AdDataTime    = $null
$script:AdLoadJob     = $null
$global:ClientToolUi  = $true         # Get-AdComputerInfo sadece ana pencerede $AdData'ya bakar
$script:AdCacheSources = @()
$script:TsDefs        = @()           # { Name; PackageId; DeploymentId; PkgName }
$script:Collections   = @()           # { Id; Name; Short }
$script:CollOps       = @{}           # calisan koleksiyon isleri (Key -> hostlar, koleksiyon)
$script:ActiveTsName  = ''
$script:PolicyWaitSec = 90            # TS politikasi yoksa Machine Policy sonrasi en fazla bekleme

# Ham TS durumu (istemci kayit defteri Execution History ... \_State) -> normalize deger. SIRA ONEMLI, ilk eslesen kazanir.
# Farkli bir ham deger gorurseniz (cift tik detayinda "_State=") buraya bir satir ekleyin.
$script:TsStateMap = @(
    @('Success',                            'BASARILI'),
    @('Fail|Error|Abort|Cancel',            'HATA'),
    @('Run|Progress|Execut|Start|Reboot',   'CALISIYOR'),
    @('Wait|Pending|Download|Content|Queue', 'BEKLIYOR')
)
# SCCM istemci eylemleri (client_sccm_tool Panel 1 ile ayni): gorunen ad -> TriggerSchedule ScheduleID listesi
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

# Grid kolonlari (ad, baslik, genislik agirligi, gorunur)
$script:GridCols = @(
    @('Hostname', 'Hostname', 11), @('DNS_Durumu', 'DNS_Durumu', 9), @('Online_mi', 'Online_mi', 9), @('AD_Durumu', 'AD_Durumu', 6),
    @('AD_LastLogon', 'AD_LastLogon', 9), @('SCCM_Kayit', 'SCCM_Kayit', 9), @('CcmExec_Durumu', 'CcmExec', 7), @('Boundary_Groups', 'Boundary Group(s)', 9), @('OS_Build', 'Operating System Build', 7), @('OS_Build_Anlik', 'OS Build (anlik)', 7),
    @('Uye_Koleksiyonlar', 'Uye_Koleksiyonlar', 11), @('TS_Durum', 'TS_Durum', 8), @('TS_Adim', 'TS_Adim', 11),
    @('TS_HataKodu', 'TS_HataKodu', 7), @('TS_SonGuncelleme', 'TS_SonGuncelleme', 9), @('Son_Islem', 'Son Islem', 10),
    @('Islem_Sonucu', 'Islem Sonucu', 10), @('Hata_Not', 'Hata / Not', 9)
)
#endregion

#region Worker'larda da kullanilan fonksiyonlar (mevcut scriptten AYNEN)
# Bu bolumdeki fonksiyonlar $script: degiskeni kullanmaz; New-WorkerScript ile worker'lara gomulur.
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

function Test-PingOnly {
    # "Sadece ping" listesi (1. bolum): AD / WMI guveni olmayan (or. SCCM trust var, AD trust yok) cihazlar.
    # Desenler -like ile (* ve ?), virgul / noktali virgul / bosluk ile ayrilir; tam ad VEYA kisa ad eslesirse $true.
    # Liste worker'larda da okunabilsin diye process ortam degiskeninde: $env:CLIENTTOOL_PINGONLY
    param([string]$ComputerName)
    $p = [string]$env:CLIENTTOOL_PINGONLY
    if (-not $p) { return $false }
    $short = ($ComputerName -split '\.')[0]
    foreach ($x in ($p -split '[,;\s]+')) { if ($x -and ($ComputerName -like $x -or $short -like $x)) { return $true } }
    return $false
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

    # "Sadece ping" listesindeki cihaz: ping var -> cevrimici (NetBIOS / WMI kimlik dogrulamasi YAPILMAZ)
    if (Test-PingOnly -ComputerName $ComputerName) {
        $r.Ip = $alive[0]; $r.Online = $true; $r.Identity = 'Muaf'; $r.OnlineText = 'Evet (sadece ping)'
        $r.Note = 'Sadece ping listesinde: kimlik (NetBIOS / WMI) ve servis kontrolu yapilmadi.'
        return $r
    }

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

function Get-AdComputerInfo {
    # $AdData'ya sorgu: $AdData.Where({ $_.Name -eq <hostname> }, 'First'). Veri Get-ADComputer'dan geldigi gibi durur.
    # SADECE ana pencerede calisir; paralel worker'larda bos doner, AD kolonlari sonuc gelince ana pencerede (Set-AdInfo) doldurulur.
    param([string]$ComputerName)
    $res = [pscustomobject]@{ Status = ''; LastLogon = ''; Detail = '' }
    if (-not $global:ClientToolUi) { return $res }
    if ($null -eq $script:AdData) {
        $res.Status = if ($script:AdLoadJob) { 'Yukleniyor' } else { 'Sorgulanamadi' }
        $res.Detail = if ($script:AdLoadJob) { 'AD verisi yukleniyor; bitince AD kolonlari otomatik dolar' } else { "AD verisi yuklenmedi ('AD Verisi...' > 'Kaydet ve AD Verisini Yukle')" }
        return $res
    }
    $short = (([string]$ComputerName -split '\.')[0]).Trim()
    $c = $script:AdData.Where({ $_.Name -eq $short }, 'First')
    if ($c.Count -eq 0) { $res.Status = 'Bulunamadi'; $res.LastLogon = 'N/A'; $res.Detail = "AD verisinde (girilen OU'larda) yok"; return $res }
    $c = $c[0]
    $ll = $c.LastLogonTimestamp
    if ($null -eq $ll) { try { $ll = $c['LastLogonTimestamp'].Value } catch { } }   # AD modulu bu runspace'te yuklu degilse indeksleyici
    $res.Status = if ($c.Enabled) { 'Enabled' } else { 'Disabled' }
    $res.LastLogon = if ($ll) { [DateTime]::FromFileTime([long]$ll).ToString('yyyy-MM-dd HH:mm') } else { 'Hic' }
    return $res
}

function Get-ServiceCheck {
    # Istenen servislerin + CcmExec'in durumunu TEK RPC cagrisi (ServiceController.GetServices) ile alir.
    # Ad veya gorunen ad ile eslesir; karsilastirma kulturden bagimsizdir (tr-TR'de 'bits' = 'BITS').
    param([string]$ComputerName, [string[]]$ServiceNames = @())

    $r = [pscustomobject]@{ Running = ''; NotRunning = ''; CcmExec = ''; ServiceDetail = @(); Note = '' }
    if (Test-PingOnly -ComputerName $ComputerName) {   # WMI / RPC yetkisi yok: servis sorgulanmaz
        $r.Running = $r.NotRunning = $r.CcmExec = 'Kontrol edilmedi (sadece ping)'
        return $r
    }
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
#endregion

#region Worker fonksiyonlari (YENI)

function Get-WmiProp {
    # WMI nesnesinden alan okur; alan yoksa $null (surum farklarinda sorgu bozulmasin).
    param($Obj, [string]$Name)
    if ($Obj -and $Obj.PSObject.Properties[$Name]) { return $Obj.$Name }
    return $null
}

function ConvertFrom-WmiDate {
    param($Value)
    if (-not $Value) { return $null }
    try { return [System.Management.ManagementDateTimeConverter]::ToDateTime([string]$Value) } catch { return $null }
}

function Test-AccessError {
    # Yetki hatasi mi? (WMI AccessDenied 0x80041003, Win32 0x80070005, SMS Provider RBAC)
    param([string]$Text)
    return ($Text -match 'Access denied|Access is denied|Erisim engellendi|0x80041003|0x80070005|AccessDenied')
}

function Get-WmiChunked {
    # "SELECT ... WHERE <kosul>" sorgusunu ResourceID listesi icin 100'luk parcalar halinde calistirir (WQL'de IN yok).
    param([string]$SmsProvider, [string]$Ns, [string]$Query, [uint32[]]$Rids, [int]$Chunk = 100)
    $out = New-Object System.Collections.ArrayList
    $ids = @($Rids | Sort-Object -Unique)
    for ($i = 0; $i -lt $ids.Count; $i += $Chunk) {
        $part = @($ids[$i..([Math]::Min($i + $Chunk, $ids.Count) - 1)])
        $w = (($part | ForEach-Object { "ResourceID = $_" }) -join ' OR ')
        foreach ($o in @(Get-WmiObject -ComputerName $SmsProvider -Namespace $Ns -Query "$Query AND ($w)" -ErrorAction Stop)) { [void]$out.Add($o) }
    }
    return $out
}

function Test-DirectRule {
    # Koleksiyon kurali Direct mi? (SMS_CollectionRuleDirect; sadece bu kural tipinde ResourceID alani var)
    param($Rule)
    if (-not $Rule) { return $false }
    $cls = ''
    try { $cls = [string]$Rule.psbase.ClassPath.ClassName } catch { }
    if (-not $cls) { try { $cls = [string]$Rule.psbase.SystemProperties['__CLASS'].Value } catch { } }
    if ($cls) { return ($cls -eq 'SMS_CollectionRuleDirect') }
    return ($null -ne $Rule.PSObject.Properties['ResourceID'] -and $null -ne $Rule.ResourceID)
}

function Get-CollectionDirectMap {
    # Koleksiyonlarin DIRECT uyelik kurallari: CollectionID -> { Ok; Error; Name; Rids (HashSet[uint32]) }. Koleksiyon basina tek cagri.
    param([string]$SmsProvider, [string]$SiteCode, [string[]]$CollectionIds)
    $ns  = "root\SMS\site_$SiteCode"
    $map = New-Object System.Collections.Hashtable ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($id in @($CollectionIds | Where-Object { $_ } | Sort-Object -Unique)) {
        $set = New-Object 'System.Collections.Generic.HashSet[uint32]'
        try {
            $c = @(Get-WmiObject -ComputerName $SmsProvider -Namespace $ns -Class SMS_Collection -Filter "CollectionID = '$(ConvertTo-WqlString $id)'" -ErrorAction Stop)[0]
            if (-not $c) { $map[$id] = [pscustomobject]@{ Ok = $false; Error = 'Koleksiyon bulunamadi'; Name = ''; Rids = $set }; continue }
            $c.Get()   # CollectionRules lazy alan
            foreach ($r in @($c.CollectionRules)) {
                if (-not $r) { continue }
                if (Test-DirectRule $r) { [void]$set.Add([uint32]$r.ResourceID) }
            }
            $map[$id] = [pscustomobject]@{ Ok = $true; Error = ''; Name = [string]$c.Name; Rids = $set }
        }
        catch { $map[$id] = [pscustomobject]@{ Ok = $false; Error = $_.Exception.Message; Name = ''; Rids = $set } }
    }
    return $map
}

function Get-ServerSnapshot {
    # SMS Provider TOPLU sorgusu: cihaz kayitlari (+ Boundary Group(s) / OS Build) + tanimli koleksiyonlarin direct uyelikleri.
    param([string[]]$Names, [string]$SmsProvider, [string]$SiteCode, [string[]]$CollectionIds = @())
    $dev = Get-SccmDeviceMap -Names $Names -SmsProvider $SmsProvider -SiteCode $SiteCode
    $coll = $null
    if (@($CollectionIds).Count -gt 0) { $coll = Get-CollectionDirectMap -SmsProvider $SmsProvider -SiteCode $SiteCode -CollectionIds $CollectionIds }
    return [pscustomobject]@{ Snapshot = $true; Devices = $dev; Coll = $coll }
}

function Invoke-StdReg {
    # Uzak kayit defteri: WMI StdRegProv (DCOM; Remote Registry servisi gerekmez).
    param($Reg, [string]$Method, [hashtable]$P)
    $in = $Reg.GetMethodParameters($Method)
    foreach ($k in $P.Keys) { $in[$k] = $P[$k] }
    return $Reg.InvokeMethod($Method, $in, $null)
}

function Get-RegValues {
    # Anahtarin tum degerleri (ad -> metin). HKLM.
    param($Reg, [string]$Key)
    $hk = [uint32]2147483650
    $h  = [ordered]@{}
    $ev = Invoke-StdReg $Reg 'EnumValues' @{ hDefKey = $hk; sSubKeyName = $Key }
    if ([int]$ev['ReturnValue'] -ne 0 -or -not $ev['sNames']) { return $h }
    $names = @($ev['sNames']); $types = @($ev['Types'])
    for ($i = 0; $i -lt $names.Count; $i++) {
        $n = [string]$names[$i]
        $a = @{ hDefKey = $hk; sSubKeyName = $Key; sValueName = $n }
        $h[$n] = switch ([int]$types[$i]) {
            1       { [string](Invoke-StdReg $Reg 'GetStringValue' $a)['sValue'] }
            2       { [string](Invoke-StdReg $Reg 'GetExpandedStringValue' $a)['sValue'] }
            4       { [string](Invoke-StdReg $Reg 'GetDWORDValue' $a)['uValue'] }
            11      { [string](Invoke-StdReg $Reg 'GetQWORDValue' $a)['uValue'] }
            7       { (@((Invoke-StdReg $Reg 'GetMultiStringValue' $a)['sValue']) -join ' | ') }
            default { '(binary)' }
        }
    }
    return $h
}

function Get-OsBuildLive {
    # Anlik OS Build ISTEMCIDEN (SADECE OKUR): HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion (WMI StdRegProv).
    # Bicim SCCM ile ayni: <Major>.<Minor>.<CurrentBuildNumber>.<UBR>  (ornek 10.0.22631.4317)
    param([string]$ComputerName)
    $r = [pscustomobject]@{ Build = ''; Detail = '' }
    try {
        $reg = New-Object System.Management.ManagementClass ("\\$ComputerName\root\default:StdRegProv")
        $key = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        $get = {
            param([string]$Method, [string]$Name, [string]$Field)
            $o = Invoke-StdReg $reg $Method @{ hDefKey = [uint32]2147483650; sSubKeyName = $key; sValueName = $Name }
            if ([int]$o['ReturnValue'] -eq 0) { return [string]$o[$Field] }
            return ''
        }
        $maj   = & $get 'GetDWORDValue' 'CurrentMajorVersionNumber' 'uValue'
        $min   = & $get 'GetDWORDValue' 'CurrentMinorVersionNumber' 'uValue'
        $build = & $get 'GetStringValue' 'CurrentBuildNumber' 'sValue'
        $ubr   = & $get 'GetDWORDValue' 'UBR' 'uValue'
        $disp  = & $get 'GetStringValue' 'DisplayVersion' 'sValue'
        $prod  = & $get 'GetStringValue' 'ProductName' 'sValue'
        if (-not $maj) { $cv = (& $get 'GetStringValue' 'CurrentVersion' 'sValue') -split '\.'; $maj = $cv[0]; $min = $(if ($cv.Count -gt 1) { $cv[1] } else { '0' }) }
        if ($build) { $r.Build = "$maj.$min.$build" + $(if ($ubr) { ".$ubr" } else { '' }) } else { $r.Build = 'Okunamadi' }
        $r.Detail = "OS Build (anlik, kayit defteri): $($r.Build)  | $prod $disp"
    }
    catch { $r.Build = 'Okunamadi'; $r.Detail = "OS Build (anlik) okunamadi: $($_.Exception.Message)" }
    return $r
}

function Get-TsClientState {
    # TS durumu ISTEMCIDEN (SADECE OKUR): kayit defteri calisma gecmisi + TS politikasi + calisan TS.
    #   HKLM\SOFTWARE\Microsoft\SMS\Mobile Client\Software Distribution\Execution History\System\<PackageID>\<alt anahtar>
    param([string]$ComputerName, [string]$PackageId)
    $r = [pscustomobject]@{ Ok = $false; Error = ''; Policy = $false; Required = ''; AdvIds = @(); HistCount = 0; Hist = $null; HistTime = $null; Running = $false; RunningThis = $false; RunStart = $null; Step = $null; Detail = '' }
    $L = New-Object System.Collections.Generic.List[string]
    try {
        $pk  = ConvertTo-WqlString $PackageId
        $pol = @(Get-WmiObject -ComputerName $ComputerName -Namespace 'root\ccm\policy\machine\actualconfig' -Query "SELECT * FROM CCM_TaskSequence WHERE PKG_PackageID = '$pk'" -ErrorAction Stop)
        $r.Policy = ($pol.Count -gt 0)
        $r.AdvIds = @($pol | ForEach-Object { [string](Get-WmiProp $_ 'ADV_AdvertisementID') } | Where-Object { $_ } | Sort-Object -Unique)
        if ($r.Policy) { $r.Required = [string](Get-WmiProp $pol[0] 'ADV_MandatoryAssignments'); $L.Add("Politika: var (Deployment=$(Get-WmiProp $pol[0] 'ADV_AdvertisementID'); Zorunlu=$($r.Required); Tekrar=$(Get-WmiProp $pol[0] 'ADV_RepeatRunBehavior'))") }
        else { $L.Add('Politika: yok (actualconfig CCM_TaskSequence)') }

        $reg  = New-Object System.Management.ManagementClass ("\\$ComputerName\root\default:StdRegProv")
        $base = "SOFTWARE\Microsoft\SMS\Mobile Client\Software Distribution\Execution History\System\$PackageId"
        $ek   = Invoke-StdReg $reg 'EnumKey' @{ hDefKey = [uint32]2147483650; sSubKeyName = $base }
        $subs = if ([int]$ek['ReturnValue'] -eq 0 -and $ek['sNames']) { @($ek['sNames']) } else { @() }
        foreach ($s in $subs) {
            $v = Get-RegValues $reg "$base\$s"
            $r.HistCount++
            $t = $null; $dt = [datetime]::MinValue
            if ([datetime]::TryParse([string]$v['_RunStartTime'], [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { $t = $dt }
            $L.Add("Gecmis [$s]: " + (($v.Keys | ForEach-Object { "$_=$($v[$_])" }) -join '; '))
            if (-not $r.Hist -or ($t -and (-not $r.HistTime -or $t -gt $r.HistTime))) { $r.Hist = $v; $r.HistTime = $t }
        }
        if ($subs.Count -eq 0) { $L.Add("Calisma gecmisi yok (HKLM\$base)") }

        $tsm = @(Get-WmiObject -ComputerName $ComputerName -Class Win32_Process -Filter "Name = 'TsManager.exe'" -ErrorAction Stop)
        $r.Running = ($tsm.Count -gt 0)
        if ($r.Running) { $r.RunStart = ConvertFrom-WmiDate (Get-WmiProp $tsm[0] 'CreationDate') }
        if ($r.Running) {
            try {   # calisan TS bu paket mi? (istek nesnesinin herhangi bir alaninda PackageID geciyor mu)
                foreach ($q in @(Get-WmiObject -ComputerName $ComputerName -Namespace 'root\ccm\SoftMgmtAgent' -Class CCM_TSExecutionRequest -ErrorAction Stop)) {
                    foreach ($pp in $q.PSObject.Properties) { if ("$($pp.Value)" -like "*$PackageId*") { $r.RunningThis = $true } }
                }
            }
            catch { }
            $L.Add("TsManager.exe calisiyor; bu TS'e ait istek: $(if ($r.RunningThis) { 'var' } else { 'bulunamadi (baska TS olabilir)' })")
        }
        $r.Ok = $true
    }
    catch { $r.Error = $_.Exception.Message; $L.Add("Hata: $($_.Exception.Message)") }
    $r.Detail = ($L -join "`r`n")
    return $r
}

function Get-TsAdvIds {
    # TS'in deployment (Advertisement) ID'leri: istemci politikasindakiler + tanimdaki + SMS Provider'daki tum deployment'lari.
    param([string]$SmsProvider, [string]$Ns, [string]$PackageId, [string[]]$Known)
    $ids = @($Known | Where-Object { $_ })
    if ($PackageId) {
        try { $ids += @(Get-WmiObject -ComputerName $SmsProvider -Namespace $Ns -Query "SELECT AdvertisementID FROM SMS_Advertisement WHERE PackageID = '$(ConvertTo-WqlString $PackageId)'" -ErrorAction Stop | ForEach-Object { [string]$_.AdvertisementID }) } catch { }
    }
    return @($ids | Where-Object { $_ } | Sort-Object -Unique)
}

function Get-TsResourceId {
    param([string]$ComputerName, [string]$SmsProvider, [string]$Ns)
    $short = ConvertTo-WqlString ((($ComputerName -split '\.')[0]).Trim())
    $dev = @(Get-WmiObject -ComputerName $SmsProvider -Namespace $Ns -Query "SELECT ResourceId, Client, Obsolete FROM SMS_R_System WHERE Name = '$short'" -ErrorAction Stop)
    $d = @($dev | Where-Object { $_.Client -eq 1 -and $_.Obsolete -ne 1 })[0]; if (-not $d) { $d = $dev[0] }
    if ($d) { return [uint32]$d.ResourceId }
    return $null
}

function Get-TsStepInfo {
    # TS adim adlari SADECE SMS Provider'da var (SMS_TaskSequenceExecutionStatus; istemci her adim BITINCE bildirir, sunucuya gecikmeli yansir).
    # Doner: { Error; Last (en son bildirilen adim); Fail (en son ExitCode<>0 adim); Count; AdvIds }. -Since: bu calismanin baslangici.
    # Sadece ActionName dolu satirlar adimdir (or. 11134 tamamlandi, 11135 basarisiz, 11139 adim hatasi nedeniyle durdu).
    param([string]$ComputerName, [string]$SmsProvider, [string]$SiteCode, [string]$PackageId, [string[]]$AdvIds, $Since)
    $r = [pscustomobject]@{ Error = ''; Last = $null; Fail = $null; Count = 0; AdvIds = @() }
    try {
        $ns  = "root\SMS\site_$SiteCode"
        $rid = Get-TsResourceId -ComputerName $ComputerName -SmsProvider $SmsProvider -Ns $ns
        if ($null -eq $rid) { $r.Error = 'SCCM kaydi yok'; return $r }
        $r.AdvIds = Get-TsAdvIds -SmsProvider $SmsProvider -Ns $ns -PackageId $PackageId -Known $AdvIds
        if ($r.AdvIds.Count -eq 0) { $r.Error = 'Deployment bulunamadi'; return $r }
        $w = (@($r.AdvIds | ForEach-Object { "AdvertisementID = '$(ConvertTo-WqlString $_)'" }) -join ' OR ')
        foreach ($s in @(Get-WmiObject -ComputerName $SmsProvider -Namespace $ns -Query "SELECT * FROM SMS_TaskSequenceExecutionStatus WHERE ResourceID = $rid AND ($w)" -ErrorAction Stop)) {
            $t = ConvertFrom-WmiDate (Get-WmiProp $s 'ExecutionTime')
            if ($Since -and $t -and $t -lt $Since.AddMinutes(-2)) { continue }   # onceki calismalar
            $s | Add-Member -NotePropertyName _T -NotePropertyValue $t -Force
            $r.Count++
            # Adi bos satirlar adim degil, TS motorunun genel mesajlari (11140 basladi / 11141 basarisiz / 11171 bitti...): atlanir
            if (-not ([string](Get-WmiProp $s 'ActionName')).Trim()) { continue }
            if (-not $r.Last -or ($t -and $t -gt $r.Last._T)) { $r.Last = $s }
            $ec = Get-WmiProp $s 'ExitCode'
            if ($null -ne $ec -and "$ec" -ne '0' -and (-not $r.Fail -or ($t -and $t -gt $r.Fail._T))) { $r.Fail = $s }
        }
    }
    catch { $r.Error = $_.Exception.Message }
    return $r
}

function Get-TsStepHistory {
    # SADECE OKUR: cihazin bu TS icin SCCM'e bildirdigi TUM adimlar (en yeni ustte). Detay penceresi icin metin.
    param([string]$ComputerName, [string]$SmsProvider, [string]$SiteCode, [string]$PackageId, [string]$DeploymentId)
    $L = New-Object System.Collections.Generic.List[string]
    try {
        $ns  = "root\SMS\site_$SiteCode"
        $rid = Get-TsResourceId -ComputerName $ComputerName -SmsProvider $SmsProvider -Ns $ns
        if ($null -eq $rid) { return 'SCCM kaydi yok.' }
        $adv = Get-TsAdvIds -SmsProvider $SmsProvider -Ns $ns -PackageId $PackageId -Known @($DeploymentId)
        if ($adv.Count -eq 0) { return "Bu paket icin deployment bulunamadi: $PackageId" }
        $w = (@($adv | ForEach-Object { "AdvertisementID = '$(ConvertTo-WqlString $_)'" }) -join ' OR ')
        $rows = @(Get-WmiObject -ComputerName $SmsProvider -Namespace $ns -Query "SELECT * FROM SMS_TaskSequenceExecutionStatus WHERE ResourceID = $rid AND ($w)" -ErrorAction Stop |
            ForEach-Object { $_ | Add-Member -NotePropertyName _T -NotePropertyValue (ConvertFrom-WmiDate (Get-WmiProp $_ 'ExecutionTime')) -Force -PassThru } | Sort-Object _T -Descending)
        $L.Add("ResourceID=$rid  Deployment(lar)=$($adv -join ', ')  Kayit=$($rows.Count)  (istemci adimi BITINCE bildirir; sunucuya gecikmeli yansiyabilir)")
        $L.Add(('{0,-19}  {1,4}  {2,-12}  {3,-40}  {4,-25}  {5}' -f 'Zaman', 'Step', 'ExitCode', 'Adim', 'Grup', 'Mesaj'))
        foreach ($s in @($rows | Select-Object -First 300)) {
            $ec = Get-WmiProp $s 'ExitCode'
            $hex = if ($null -ne $ec -and "$ec" -ne '0') { '0x{0:X8}' -f ([int64]$ec -band 4294967295) } else { '0' }
            $L.Add(('{0,-19}  {1,4}  {2,-12}  {3,-40}  {4,-25}  {5} {6}' -f $(if ($s._T) { $s._T.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }), (Get-WmiProp $s 'Step'), $hex,
                (Get-WmiProp $s 'ActionName'), (Get-WmiProp $s 'GroupName'), (Get-WmiProp $s 'LastStatusMessageID'), (Get-WmiProp $s 'LastStatusMessageIDName')))
            $o = [string](Get-WmiProp $s 'ActionOutput')
            if ($hex -ne '0' -and $o) { $L.Add("      Cikti: $(if ($o.Length -gt 1500) { $o.Substring(0, 1500) + ' ...' } else { $o })") }
        }
    }
    catch { $L.Add("Hata: $($_.Exception.Message)") }
    return ($L -join "`r`n")
}

function Get-TsHostCheck {
    # Kontrol (cihaz basina): DNS + ping + kimlik + CcmExec (+ PackageId verilmisse TS durumu istemciden).
    # AD ana pencerede (Set-AdInfo), SCCM kaydi / Boundary / OS Build toplu sorgudan.
    param([string]$ComputerName, [string]$PackageId = '', [string]$SmsProvider = '', [string]$SiteCode = '', [string]$DeploymentId = '')
    $reach = Get-HostReach -ComputerName $ComputerName
    $r = [pscustomobject]@{
        Hostname = $ComputerName; DnsStatus = $reach.DnsText; IsOnline = $reach.Online; OnlineText = $reach.OnlineText; Identity = $reach.Identity
        AdStatus = ''; LastLogon = ''; CcmExec = ''; ErrorMessage = ''; Ts = $null; OsLive = $null
    }
    $notes = @($reach.Note | Where-Object { $_ })
    if (-not $reach.DnsOk) { $r.CcmExec = 'Kontrol edilemedi (DNS)' }
    elseif (-not $reach.Online) { $r.CcmExec = 'Kontrol edilemedi (cevrimdisi)' }
    else {
        $svc = Get-ServiceCheck -ComputerName $ComputerName
        $r.CcmExec = $svc.CcmExec
        if ($svc.Note) { $notes += $svc.Note }
        if ($reach.Identity -eq 'Dogrulandi') { $r.OsLive = Get-OsBuildLive -ComputerName $ComputerName }   # sadece ping listesi: WMI yok
    }
    if ($PackageId) {
        if ($reach.Online -and $reach.Identity -eq 'Dogrulandi') {
            $r.Ts = Get-TsClientState -ComputerName $ComputerName -PackageId $PackageId
            # Adim adi: sadece calisiyorsa ya da son calisma hatali ise SMS Provider'dan (SCCM'e bagliysa)
            $h = $r.Ts.Hist
            $failed = ($h -and ([string]$h['_State'] -match 'Fail|Error|Abort|Cancel' -or ("$($h['SuccessOrFailureCode'])" -notin '', '0' -and [string]$h['_State'] -notmatch 'Success')))
            if ($SmsProvider -and $SiteCode -and $r.Ts.Ok -and ($r.Ts.RunningThis -or $failed)) {
                $adv = @(@($r.Ts.AdvIds) + @($DeploymentId) | Where-Object { $_ } | Sort-Object -Unique)
                $r.Ts.Step = Get-TsStepInfo -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode -PackageId $PackageId -AdvIds $adv -Since $(if ($r.Ts.RunningThis) { $r.Ts.RunStart } else { $null })
            }
        }
        else {
            $why = if (-not $reach.Online) { $(if ($reach.Note) { $reach.Note } else { 'Cevrimdisi' }) } else { 'Kimlik dogrulanamadi (sadece ping listesi)' }
            $r.Ts = [pscustomobject]@{ Ok = $false; Error = $why; Detail = "TS durumu okunmadi: $why" }
        }
    }
    $r.ErrorMessage = ($notes -join ' | ')
    return $r
}

function Invoke-TsTrigger {
    # TS'i istemcide tetikler (WMI/DCOM): Software Center "Yukle" ile ayni cagri - ClientSDK CCM_ProgramsManager.ExecuteProgram.
    # Once: Online + kimlik dogrulanmis + CcmExec Running. TS istemcide (CCM_Program) yoksa Machine Policy Retrieval &
    # Evaluation tetiklenir ve en fazla $PolicyWaitSec beklenir; hala yoksa "Politika yok". Istemci politikasi DEGISTIRILMEZ.
    # Code: Tetiklendi | DryRun | Atlandi | Politika yok | Hata
    param([string]$ComputerName, [string]$PackageId, [bool]$DryRun, [int]$PolicyWaitSec = 90)
    $r = [pscustomobject]@{ Hostname = $ComputerName; DnsStatus = ''; OnlineText = ''; IsOnline = $false; CcmExec = ''; Code = 'Hata'; Reason = ''; Detail = '' }
    $d = New-Object System.Collections.Generic.List[string]
    try {
        $reach = Get-HostReach -ComputerName $ComputerName
        $r.DnsStatus = $reach.DnsText; $r.OnlineText = $reach.OnlineText; $r.IsOnline = $reach.Online
        if (-not $reach.Online) { $r.Code = 'Atlandi'; $r.Reason = $(if ($reach.Note) { $reach.Note } else { 'Cevrimdisi' }); return $r }
        if ($reach.Identity -ne 'Dogrulandi') { $r.Code = 'Atlandi'; $r.Reason = 'Kimlik dogrulanamadi (sadece ping listesi)'; return $r }

        $svc = Get-ServiceCheck -ComputerName $ComputerName
        $r.CcmExec = $svc.CcmExec
        if ($svc.CcmExec -eq 'Yok') { $r.Code = 'Atlandi'; $r.Reason = 'CcmExec servisi yok'; return $r }
        if ($svc.CcmExec -eq 'Erisilemedi') { $r.Code = 'Atlandi'; $r.Reason = "CcmExec sorgulanamadi: $($svc.Note)"; return $r }
        if ($svc.CcmExec -ne 'Running') { $r.Code = 'Atlandi'; $r.Reason = "CcmExec servisi $($svc.CcmExec)"; return $r }

        $q = "SELECT * FROM CCM_Program WHERE PackageID = '$(ConvertTo-WqlString $PackageId)'"
        $p = @(Get-WmiObject -ComputerName $ComputerName -Namespace 'root\ccm\ClientSDK' -Query $q -ErrorAction Stop)
        if ($p.Count -eq 0) {
            if ($DryRun) { $r.Code = 'Politika yok'; $r.Reason = 'TS istemcide yok (DRY-RUN: Machine Policy istenmedi)'; return $r }
            foreach ($sid in '{00000000-0000-0000-0000-000000000021}', '{00000000-0000-0000-0000-000000000022}') {
                $null = Invoke-WmiMethod -ComputerName $ComputerName -Namespace 'root\ccm' -Class 'SMS_Client' -Name 'TriggerSchedule' -ArgumentList $sid -ErrorAction Stop
            }
            $d.Add("TS istemcide yoktu; Machine Policy Retrieval + Evaluation tetiklendi, en fazla $PolicyWaitSec sn beklendi.")
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            while ($p.Count -eq 0 -and $sw.Elapsed.TotalSeconds -lt $PolicyWaitSec) {
                Start-Sleep -Seconds 10
                $p = @(Get-WmiObject -ComputerName $ComputerName -Namespace 'root\ccm\ClientSDK' -Query $q -ErrorAction Stop)
            }
            if ($p.Count -eq 0) { $r.Code = 'Politika yok'; $r.Reason = "TS politikasi $PolicyWaitSec sn icinde istemciye ulasmadi"; $r.Detail = ($d -join "`r`n"); return $r }
        }
        $p0 = $p[0]
        $d.Add("Istemcideki TS (CCM_Program): PackageID=$($p0.PackageID) ProgramID=$($p0.ProgramID) Ad=$(Get-WmiProp $p0 'Name') " +
               "LastRunStatus=$(Get-WmiProp $p0 'LastRunStatus') EvaluationState=$(Get-WmiProp $p0 'EvaluationState') ErrorCode=$(Get-WmiProp $p0 'ErrorCode')")
        if ($DryRun) { $r.Code = 'DryRun'; $r.Reason = 'Tetiklenecekti'; $r.Detail = ($d -join "`r`n"); return $r }

        $mc = New-Object System.Management.ManagementClass ("\\$ComputerName\root\ccm\ClientSDK:CCM_ProgramsManager")
        $in = $mc.GetMethodParameters('ExecuteProgram')
        $in['PackageID'] = [string]$p0.PackageID
        $in['ProgramID'] = [string]$p0.ProgramID
        $o  = $mc.InvokeMethod('ExecuteProgram', $in, $null)
        $rv = [int64]$o['ReturnValue']
        $d.Add(('ExecuteProgram ReturnValue = 0x{0:X8}' -f ($rv -band 4294967295)))
        if ($rv -eq 0) { $r.Code = 'Tetiklendi' } else { $r.Code = 'Hata'; $r.Reason = ('ExecuteProgram 0x{0:X8}' -f ($rv -band 4294967295)) }
    }
    catch {
        $why = Get-WmiFailureReason -ErrorRecord $_
        $r.Code = 'Hata'; $r.Reason = $why.Text
        $d.Add("Hata: $($_.Exception.Message)")
    }
    $r.Detail = ($d -join "`r`n")
    return $r
}

function Get-TsClientDetail {
    # SADECE OKUR: tek cihazda istemci tarafi TS bilgisi (ham, detay penceresi icin).
    param([string]$ComputerName, [string]$PackageId)
    $L = New-Object System.Collections.Generic.List[string]
    $dump = {
        param([string]$Title, [scriptblock]$Get)
        $L.Add("== $Title")
        try {
            $rows = @(& $Get)
            if ($rows.Count -eq 0) { $L.Add('   (kayit yok)') }
            foreach ($o in $rows) {
                foreach ($p in @($o.PSObject.Properties | Where-Object { $_.Name -notlike '__*' -and $_.Name -notin 'Scope', 'Path', 'Options', 'ClassPath', 'Properties', 'SystemProperties', 'Qualifiers', 'Site', 'Container', 'PSComputerName' })) {
                    if ($null -ne $p.Value -and "$($p.Value)" -ne '') { $L.Add(("   {0,-32}: {1}" -f $p.Name, (($p.Value | ForEach-Object { "$_" }) -join ' | '))) }
                }
                $L.Add('   ' + ('-' * 40))
            }
        }
        catch { $L.Add("   HATA: $($_.Exception.Message)") }
    }
    $pk = ConvertTo-WqlString $PackageId
    & $dump 'ClientSDK CCM_Program (Software Center gorunumu)' { Get-WmiObject -ComputerName $ComputerName -Namespace 'root\ccm\ClientSDK' -Query "SELECT * FROM CCM_Program WHERE PackageID = '$pk'" -ErrorAction Stop }
    & $dump 'Politika CCM_TaskSequence (actualconfig)' { Get-WmiObject -ComputerName $ComputerName -Namespace 'root\ccm\policy\machine\actualconfig' -Query "SELECT * FROM CCM_TaskSequence WHERE PKG_PackageID = '$pk'" -ErrorAction Stop | Select-Object ADV_AdvertisementID, PKG_PackageID, PKG_Name, ADV_MandatoryAssignments, ADV_RepeatRunBehavior, ADV_ActiveTime, PRG_ProgramID }
    & $dump 'Calisan TS istegi (SoftMgmtAgent CCM_TSExecutionRequest)' { Get-WmiObject -ComputerName $ComputerName -Namespace 'root\ccm\SoftMgmtAgent' -Class CCM_TSExecutionRequest -ErrorAction Stop }
    & $dump 'TS surecleri (TsManager / TSProgressUI)' { Get-WmiObject -ComputerName $ComputerName -Class Win32_Process -Filter "Name = 'TsManager.exe' OR Name = 'TSProgressUI.exe'" -ErrorAction Stop | Select-Object Name, ProcessId, CreationDate }
    $L.Add('== Kayit defteri / politika / calisan TS (Get-TsClientState)')
    $L.Add((Get-TsClientState -ComputerName $ComputerName -PackageId $PackageId).Detail)
    return ($L -join "`r`n")
}

function Invoke-CollectionChange {
    # Koleksiyona DIRECT uye ekleme / cikarma (SMS_Collection.AddMembershipRule(s) / DeleteMembershipRule(s)).
    # Idempotent: zaten uye -> eklenmez, uye degil -> cikarilmaz. Direct olmayan (sorgu / include) uye cikarilamaz.
    # Sonucta tanimli koleksiyonlarin guncel direct uyelikleri de doner (Uye_Koleksiyonlar kolonu icin).
    param([string]$SmsProvider, [string]$SiteCode, [string]$CollectionId, [string]$Mode, [string[]]$Names, [bool]$DryRun, [bool]$Refresh, [string[]]$AllCollectionIds = @())
    $ns  = "root\SMS\site_$SiteCode"
    $out = [pscustomobject]@{ CollOp = $true; Ok = $false; Error = ''; Results = @(); Devices = $null; Coll = $null; Refreshed = $false }
    $res = New-Object System.Collections.ArrayList
    $add = { param($h, $rid, $code, $det) [void]$res.Add([pscustomobject]@{ Host = $h; Rid = $rid; Code = $code; Detail = $det }) }
    try {
        $dev = Get-SccmDeviceMap -Names $Names -SmsProvider $SmsProvider -SiteCode $SiteCode
        $out.Devices = $dev
        $c = @(Get-WmiObject -ComputerName $SmsProvider -Namespace $ns -Class SMS_Collection -Filter "CollectionID = '$(ConvertTo-WqlString $CollectionId)'" -ErrorAction Stop)[0]
        if (-not $c) { throw "Koleksiyon bulunamadi: $CollectionId" }
        if ([int]$c.CollectionType -ne 2) { throw "Device Collection degil (CollectionType=$($c.CollectionType))" }
        $c.Get()
        $direct = @{}   # ResourceID -> mevcut direct kural nesnesi
        foreach ($r in @($c.CollectionRules)) {   # mevcut direct kurallar: ResourceID -> { RuleName; ResourceClassName }
            if (Test-DirectRule $r) { $direct[[uint32]$r.ResourceID] = [pscustomobject]@{ RuleName = [string]$r.RuleName; ResourceClassName = $(if ($r.ResourceClassName) { [string]$r.ResourceClassName } else { 'SMS_R_System' }) } }
        }

        $todo = New-Object System.Collections.ArrayList   # { Host; Rid; Rule }
        $notDirect = @()
        foreach ($h in $Names) {
            $k = (($h -split '\.')[0]).Trim()
            $di = $dev[$k]
            if (-not $di -or -not $di.Ok) { & $add $h $null 'Islem hatasi' "SCCM kaydi sorgulanamadi: $(if ($di) { $di.Error })"; continue }
            $ref = if ($di.Dev) { $di.Dev } elseif (@($di.Cands).Count -gt 0) { $di.Cands[0] } else { $null }
            if (-not $ref) { & $add $h $null 'SCCM kaydi yok' 'SMS_R_System kaydi yok'; continue }
            $rid = [uint32]$ref.ResourceId
            if ($Mode -eq 'Add') {
                if ($direct.ContainsKey($rid)) { & $add $h $rid 'Zaten uye' 'Direct kural zaten var'; continue }
                [void]$todo.Add([pscustomobject]@{ Host = $h; Rid = $rid })
            }
            else {
                if ($direct.ContainsKey($rid)) { [void]$todo.Add([pscustomobject]@{ Host = $h; Rid = $rid; Rule = $direct[$rid] }) }
                else { $notDirect += [pscustomobject]@{ Host = $h; Rid = $rid } }
            }
        }
        if ($notDirect.Count -gt 0) {   # direct kural yok: baska kuralla uye mi?
            $full = New-Object 'System.Collections.Generic.HashSet[uint32]'
            try { foreach ($m in (Get-WmiChunked $SmsProvider $ns "SELECT ResourceID FROM SMS_FullCollectionMembership WHERE CollectionID = '$(ConvertTo-WqlString $CollectionId)'" @($notDirect | ForEach-Object { $_.Rid }))) { [void]$full.Add([uint32]$m.ResourceID) } } catch { }
            foreach ($x in $notDirect) {
                if ($full.Contains($x.Rid)) { & $add $x.Host $x.Rid 'Direct uye degil, cikarilamaz' 'Uyelik direct kuraldan degil (sorgu / include kurali)' }
                else { & $add $x.Host $x.Rid 'Uye degil' 'Koleksiyonda direct kurali yok' }
            }
        }

        if ($todo.Count -gt 0) {
            if ($DryRun) {
                foreach ($t in $todo) { & $add $t.Host $t.Rid $(if ($Mode -eq 'Add') { 'DRY-RUN: eklenecek' } else { 'DRY-RUN: cikarilacak' }) 'DRY-RUN: SMS Provider''a yazilmadi' }
            }
            else {
                $mo = $c.psbase
                # Ekleme ve cikarma icin YENI SMS_CollectionRuleDirect ornegi (cikarmada mevcut kuralin RuleName / ResourceClassName'i kopyalanir).
                # Koleksiyondan okunan gomulu kural nesnesi metoda parametre olarak VERILMEZ (WMI tip / parametre hatasi verir).
                $rc = New-Object System.Management.ManagementClass ("\\$SmsProvider\$ns" + ':SMS_CollectionRuleDirect')
                foreach ($t in $todo) {
                    $ri = $rc.CreateInstance()
                    $ri['ResourceID'] = [uint32]$t.Rid
                    if ($Mode -eq 'Add') { $ri['ResourceClassName'] = 'SMS_R_System'; $ri['RuleName'] = [string]$t.Host }
                    else { $ri['ResourceClassName'] = $t.Rule.ResourceClassName; $ri['RuleName'] = $t.Rule.RuleName }
                    $t | Add-Member -NotePropertyName WmiRule -NotePropertyValue $ri -Force
                }
                $one = if ($Mode -eq 'Add') { 'AddMembershipRule' } else { 'DeleteMembershipRule' }
                $err = @{}; $batchErr = ''
                try {   # once toplu (tek cagri); olmazsa cihaz basina
                    $in = $mo.GetMethodParameters("${one}s")
                    $in['collectionRules'] = [System.Management.ManagementBaseObject[]]@($todo | ForEach-Object { $_.WmiRule })
                    $o = $mo.InvokeMethod("${one}s", $in, $null)
                    if ($o -and $null -ne $o['ReturnValue'] -and [int64]$o['ReturnValue'] -ne 0) { throw "${one}s ReturnValue=$($o['ReturnValue'])" }
                }
                catch {
                    $batchErr = "${one}s: $(Get-ErrText $_.Exception)"
                    foreach ($t in $todo) {
                        try {
                            $in = $mo.GetMethodParameters($one); $in['collectionRule'] = $t.WmiRule
                            $o = $mo.InvokeMethod($one, $in, $null)
                            if ($o -and $null -ne $o['ReturnValue'] -and [int64]$o['ReturnValue'] -ne 0) { throw "$one ReturnValue=$($o['ReturnValue'])" }
                        }
                        catch { $err[$t.Rid] = "$one`: $(Get-ErrText $_.Exception)" }
                    }
                }
                # Dogrulama: kural listesi yeniden okunur
                $c2 = @(Get-WmiObject -ComputerName $SmsProvider -Namespace $ns -Class SMS_Collection -Filter "CollectionID = '$(ConvertTo-WqlString $CollectionId)'" -ErrorAction Stop)[0]
                $c2.Get()
                $now = New-Object 'System.Collections.Generic.HashSet[uint32]'
                foreach ($r in @($c2.CollectionRules)) { if (Test-DirectRule $r) { [void]$now.Add([uint32]$r.ResourceID) } }
                foreach ($t in $todo) {
                    $ok = if ($Mode -eq 'Add') { $now.Contains([uint32]$t.Rid) } else { -not $now.Contains([uint32]$t.Rid) }
                    if ($ok) { & $add $t.Host $t.Rid $(if ($Mode -eq 'Add') { 'Eklendi' } else { 'Cikarildi' }) ("ResourceID=$($t.Rid)" + $(if ($batchErr) { " (toplu cagri basarisizdi, cihaz basina yapildi: $batchErr)" } else { '' })) }
                    else {
                        $m = $err[$t.Rid]; if (-not $m) { $m = 'Islem sonrasi kural listesi beklenen durumda degil' }
                        if ($batchErr) { $m += " | $batchErr" }
                        & $add $t.Host $t.Rid $(if (Test-AccessError $m) { 'Yetki yok' } else { 'Islem hatasi' }) $m
                    }
                }
                if ($Refresh -and @($res | Where-Object { $_.Code -in 'Eklendi', 'Cikarildi' }).Count -gt 0) {
                    try { $in = $mo.GetMethodParameters('RequestRefresh'); $in['includesubcollections'] = $false; $null = $mo.InvokeMethod('RequestRefresh', $in, $null); $out.Refreshed = $true } catch { }
                }
            }
        }
        $out.Ok = $true
    }
    catch {
        $out.Error = $_.Exception.Message
        $done = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($x in $res) { [void]$done.Add($x.Host) }
        foreach ($h in $Names) { if (-not $done.Contains($h)) { & $add $h $null $(if (Test-AccessError $out.Error) { 'Yetki yok' } else { 'Islem hatasi' }) $out.Error } }
    }
    $out.Results = @($res)
    try { if (@($AllCollectionIds).Count -gt 0) { $out.Coll = Get-CollectionDirectMap -SmsProvider $SmsProvider -SiteCode $SiteCode -CollectionIds $AllCollectionIds } } catch { }
    return $out
}
#endregion

#region Worker scriptleri
function New-WorkerScript {
    # Fonksiyonlarin kaynak kodunu worker scriptinin basina gomer (tanimlar TEK yerde).
    param([string]$ParamLine, [string]$CallLine, [string[]]$Funcs)
    $defs = foreach ($n in $Funcs) { "function $n {`n$((Get-Item -Path "function:$n").ScriptBlock.ToString())`n}" }
    return [scriptblock]::Create("param($ParamLine)`n" + ($defs -join "`n") + "`n$CallLine")
}
$script:ReachFuncs = @('ConvertTo-WqlString', 'Get-ErrText', 'Test-Ping', 'Test-TcpPort', 'Get-NetBiosName', 'Get-WmiFailureReason', 'Get-WmiComputerName',
    'Get-DnsWildcardIps', 'Test-PingOnly', 'Resolve-HostDns', 'Get-HostReach', 'Get-ServiceCheck', 'Get-WmiProp')
$script:ServerFuncs = @('ConvertTo-WqlString', 'Get-ErrText', 'Get-WmiProp', 'ConvertFrom-WmiDate', 'Test-AccessError', 'Get-WmiChunked', 'Get-SccmDeviceMap', 'Test-DirectRule', 'Get-CollectionDirectMap')
$script:TsClientFuncs = @('Invoke-StdReg', 'Get-RegValues', 'Get-TsClientState', 'Get-OsBuildLive', 'ConvertFrom-WmiDate', 'Get-TsAdvIds', 'Get-TsResourceId', 'Get-TsStepInfo')

$script:CheckWorker    = New-WorkerScript '$ComputerName, $PackageId, $SmsProvider, $SiteCode, $DeploymentId' 'Get-TsHostCheck -ComputerName $ComputerName -PackageId $PackageId -SmsProvider $SmsProvider -SiteCode $SiteCode -DeploymentId $DeploymentId' ($script:ReachFuncs + $script:TsClientFuncs + 'Get-TsHostCheck')
$script:ActionWorker   = New-WorkerScript '$ComputerName, $ActionName, $ScheduleIds' 'Invoke-SccmClientAction -ComputerName $ComputerName -ActionName $ActionName -ScheduleIds $ScheduleIds' ($script:ReachFuncs + 'Invoke-SccmClientAction')
$script:TriggerWorker  = New-WorkerScript '$ComputerName, $PackageId, $DryRun, $PolicyWaitSec' 'Invoke-TsTrigger -ComputerName $ComputerName -PackageId $PackageId -DryRun $DryRun -PolicyWaitSec $PolicyWaitSec' ($script:ReachFuncs + 'Invoke-TsTrigger')
$script:StepHistWorker = New-WorkerScript '$ComputerName, $SmsProvider, $SiteCode, $PackageId, $DeploymentId' 'Get-TsStepHistory -ComputerName $ComputerName -SmsProvider $SmsProvider -SiteCode $SiteCode -PackageId $PackageId -DeploymentId $DeploymentId' @('ConvertTo-WqlString', 'Get-WmiProp', 'ConvertFrom-WmiDate', 'Get-TsAdvIds', 'Get-TsResourceId', 'Get-TsStepHistory')
$script:DetailWorker   = New-WorkerScript '$ComputerName, $PackageId' 'Get-TsClientDetail -ComputerName $ComputerName -PackageId $PackageId' (@('ConvertTo-WqlString', 'Get-WmiProp') + $script:TsClientFuncs + 'Get-TsClientDetail')
$script:SnapshotWorker = New-WorkerScript '$Names, $SmsProvider, $SiteCode, $CollectionIds' 'Get-ServerSnapshot -Names $Names -SmsProvider $SmsProvider -SiteCode $SiteCode -CollectionIds $CollectionIds' ($script:ServerFuncs + 'Get-ServerSnapshot')
$script:CollWorker     = New-WorkerScript '$SmsProvider, $SiteCode, $CollectionId, $Mode, $Names, $DryRun, $Refresh, $AllCollectionIds' 'Invoke-CollectionChange -SmsProvider $SmsProvider -SiteCode $SiteCode -CollectionId $CollectionId -Mode $Mode -Names $Names -DryRun $DryRun -Refresh $Refresh -AllCollectionIds $AllCollectionIds' ($script:ServerFuncs + 'Invoke-CollectionChange')
#endregion

#region Veri klasoru / ayarlar / denetim kaydi
function Get-DataDir {
    # Bilerek %USERPROFILE%\Desktop (GetFolderPath degil): OneDrive klasor yonlendirmesinden etkilenmesin.
    $dir = Join-Path $env:USERPROFILE 'Desktop\ServiceStatus'
    if (-not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
    return $dir
}

function Write-ActionLog {
    param([string]$Hostname, [string]$Action, [string]$State, [string]$Text)
    try {
        $clean = ("$Text" -replace '\s+', ' ').Trim()
        $line = "{0}`t{1}`t{2}`t{3}`t{4}`t{5}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $env:USERNAME, $Hostname, $Action, $State, $clean
        $file = Join-Path (Get-DataDir) ("TsTool_{0}.log" -f (Get-Date -Format 'yyyyMM'))
        Add-Content -Path $file -Value $line -Encoding UTF8
    }
    catch { }
}
function Get-SettingsFile { return (Join-Path (Get-DataDir) 'TsToolSettings.json') }
#endregion

#region Sonuc tablosu
$script:ResultsTable = New-Object System.Data.DataTable
$script:ResultsTable.Locale = [System.Globalization.CultureInfo]::InvariantCulture
foreach ($c in @($script:GridCols | ForEach-Object { $_[0] }) + @('State', 'ResourceId')) { $null = $script:ResultsTable.Columns.Add($c) }
$script:ResultsTable.PrimaryKey = @($script:ResultsTable.Columns['Hostname'])
$script:BindingSrc = New-Object System.Windows.Forms.BindingSource
$script:BindingSrc.DataSource = $script:ResultsTable

function Get-GridRow {
    param([string]$Hostname)
    $found = $script:ResultsTable.Rows.Find($Hostname)
    if ($found) { return $found }
    $row = $script:ResultsTable.NewRow()
    $row['Hostname'] = $Hostname
    $script:ResultsTable.Rows.Add($row)
    return $row
}

function Add-RowDetail {
    param([string]$Hostname, [string]$Text)
    if (-not $script:Details.ContainsKey($Hostname)) { $script:Details[$Hostname] = '' }
    $script:Details[$Hostname] += "[$(Get-Date -Format 'HH:mm:ss')] $Text`r`n`r`n"
}

function Format-JobError {
    param([string]$ErrorText)
    if ($ErrorText -eq $script:CancelText) { return 'Iptal edildi' }
    return "Islem hatasi: $ErrorText"
}

function Get-ShortReason {
    # SADECE ANA PENCERE (New-WorkerScript listesine EKLENMEZ). Uzun / degisken hata metnini Excel'de filtrelenebilir
    # SABIT bir ifadeye cevirir. Uzun metin degismeden detaya (Add-RowDetail / Add-ParserDetail) yazilmaya devam eder.
    # -Text   : asil metin (' | ' ile birlesik notlar olabilir). Bilgi amacli notlar (AD verisi.. / Sadece ping listesinde..
    #           / Script exit code..) diger kolonlarda zaten gorundugu icin atlanir; geriye bir sey kalmazsa '' doner.
    # -Context: sadece eslesmeye katilan ek metin (DNS / Online kolonu); tek basina sonuc URETMEZ.
    # SIRA ONEMLI: yukaridan asagiya ilk eslesen kazanir (-match, buyuk/kucuk harf duyarsiz).
    param([string]$Text, [string[]]$Context = @())
    $parts = @(([string]$Text -split '\s+\|\s+') | ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and $_ -notmatch '^(AD verisi|Sadece ping listesinde|Script exit code)' })
    if ($parts.Count -eq 0) { return '' }
    $all = (@($parts) + @($Context | Where-Object { $_ })) -join ' | '
    $map = @(
        @('^Script calisti',                                                                    'Cikti bos'),   # script calisti ama cikti bos/null: hata DEGIL
        @('No such host',                                                                       'No such host is known'),
        @('IP baska cihazda|baska cihazdan geliyor',                                            'IP baska cihazda'),
        @('kimlik dogrulanamadi|kimligi dogrulanamadi|cihaz dogrulanamadi|makine oldugu dogrulanamadi|Kerberos', 'Kimlik dogrulanamadi'),
        @('CcmExec servisi yok',                                                                'CcmExec servisi yok'),
        @('SCCM ajani bulunmuyor|root\\ccm|namespace',                                          'SCCM ajani yok'),
        @('CcmExec servisi',                                                                    'CcmExec calismiyor'),
        @('RPC sunucusu|RPC server',                                                            'RPC erisilemiyor'),
        @('Erisim engellendi|Access is denied|Access denied',                                   'Erisim engellendi'),
        @('yanit gelmedi|zaman asimi|timed out|timeout',                                        'Script yanit vermedi'),
        @('Iptal edildi',                                                                       'Iptal edildi'),
        @('Cevrimdisi|ping basarisiz|DNS cozumlenemedi',                                        'Cevrimdisi')
    )
    foreach ($m in $map) { if ($all -match $m[0]) { return $m[1] } }
    return 'Islem hatasi'
}

function Get-ShortDns {
    # DNS kolonu (Panel 1 + 2): IP listesi / joker cevap / hata metni yerine SABIT ifade. Uzun hali detayda.
    param([string]$Value)
    if (-not $Value -or $Value -eq $script:PendingText -or $Value -eq 'Iptal edildi' -or $Value -eq 'Alinamadi') { return $Value }
    if ($Value -match 'No such host') { return 'No such host is known' }
    if ($Value -like 'DNS hatasi*')   { return 'DNS hatasi' }
    if ($Value -like 'Islem hatasi*') { return 'Islem hatasi' }
    return 'Cozumlendi'
}

function Get-ShortOnline {
    # Online kolonu (Panel 1 + 2): IP / cihaz adi olmadan SABIT ifade. Uzun hali detayda.
    param([string]$Value)
    if ($Value -like 'Hayir (IP baska cihazda*') { return 'Hayir (IP baska cihazda)' }
    if ($Value -like 'Hayir (ping var*')        { return 'Hayir (dogrulanamadi)' }
    if ($Value -like 'Hayir*')                  { return 'Hayir' }
    return $Value
}

function Get-ShortActionResult {
    # Istemci eylemi sonucu (Panel 1 Islem Sonucu / Panel 2 CcmExec_Durumu "Onarim:"): SABIT kisa ifade.
    # Basarisizsa Get-ShortReason; basariliysa aciklama kismi atilir ("Onarim baslatildi - ccmrepair ..." -> "Onarim baslatildi").
    param([string]$Result, [string]$State)
    $t = ($Result -replace '\s+', ' ').Trim()
    if ($State -eq 'fail' -or $t -like 'Islem hatasi*' -or $t -eq 'Iptal edildi') { return (Get-ShortReason -Text $t) }
    $t = $t -replace '\s*\(cihaz kimligi dogrulanamadi\)', ''
    return (($t -split ' - ')[0]).Trim()
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

function Get-RowState {
    # Satir durumu TEK YERDE: pending / fail / warn / ok (ozet cubugu, renk, filtre, tekrar dene).
    param($Row)
    $v = @{}; foreach ($c in $script:ResultsTable.Columns) { $v[$c.ColumnName] = [string]$Row[$c.ColumnName] }
    foreach ($k in $v.Keys) { if ($v[$k] -eq $script:PendingText) { return 'pending' } }
    if ($v.Online_mi -like 'Hayir*' -or $v.DNS_Durumu -in 'No such host is known', 'DNS hatasi', 'Islem hatasi' -or $v.TS_Durum -eq 'HATA' -or
        $v.Islem_Sonucu -match '^(Hata|Yetki yok|Islem hatasi|Iptal edildi|Alinamadi)' -or $v.Hata_Not -match '^(Islem hatasi|Iptal edildi)') { return 'fail' }
    if ($v.Islem_Sonucu -match '^(Atlandi|Politika yok|SCCM kaydi yok|Direct uye degil|DRY-RUN|Uye degil)' -or $v.TS_Durum -in 'HEDEF_DEGIL', 'DURUM_YOK' -or
        ($v.CcmExec_Durumu -and $v.CcmExec_Durumu -ne 'Running' -and $v.CcmExec_Durumu -notlike 'Kontrol edil*') -or
        ($v.SCCM_Kayit -and $v.SCCM_Kayit -notlike 'Client yuklu (Active)' -and $v.SCCM_Kayit -notlike 'Sorgulanmadi*') -or
        ($v.AD_Durumu -and $v.AD_Durumu -notin 'Enabled', 'Sorgulanamadi', 'Yukleniyor') -or $v.Hata_Not) { return 'warn' }
    foreach ($k in $v.Keys) { if ($k -notin 'Hostname', 'State', 'ResourceId' -and $v[$k]) { return 'ok' } }
    return ''
}

function Set-GridValues {
    # Hucreleri yazar ve satir durumunu (Get-RowState) yeniden hesaplar.
    param([string]$Hostname, [hashtable]$Values)
    $row = Get-GridRow -Hostname $Hostname
    foreach ($k in $Values.Keys) { $row[$k] = [string]$Values[$k] }
    $st = Get-RowState $row
    if ([string]$row['State'] -ne $st) { $row['State'] = $st }
}

function Set-GridPending {
    param([string[]]$Hosts, [string[]]$Columns, [switch]$ClearError)
    foreach ($h in $Hosts) {
        $v = @{}; foreach ($c in $Columns) { $v[$c] = $script:PendingText }
        if ($ClearError) { $v['Hata_Not'] = '' }
        Set-GridValues -Hostname $h -Values $v
    }
}

function Resolve-HostPending {
    # Iptal / hata: hala 'Bekleniyor..' kalan hucreleri doldurur ve durumu yeniden hesaplar.
    param([string[]]$Hosts, [string]$ErrorText)
    foreach ($h in $Hosts) { Resolve-PendingCells -Hostname $h -ErrorText $ErrorText; Set-GridValues -Hostname $h -Values @{} }
}

function Format-HexCode {
    # Hata kodu -> 0xXXXXXXXX (ondalik / negatif / '0x..' kabul eder). 0 veya bos -> ''.
    param($Value)
    $s = ([string]$Value).Trim()
    if (-not $s) { return '' }
    $n = 0L
    if ($s -match '^0x([0-9a-fA-F]{1,16})$') { $n = [Convert]::ToInt64($Matches[1], 16) }
    elseif (-not [int64]::TryParse($s, [ref]$n)) { return $s }
    if ($n -eq 0) { return '' }
    return ('0x{0:X8}' -f ($n -band 4294967295))
}

function Get-SccmKayitText {
    param($Info)
    if (-not $Info) { return 'Sorgulanamadi' }
    if (-not $Info.Ok) { return 'Sorgulanamadi' }
    if (@($Info.Cands).Count -eq 0) { return 'SCCM kaydi yok' }
    if ($Info.Dev) { if ($Info.Dev.Active -eq 1) { return 'Client yuklu (Active)' } else { return 'Client yuklu (Inactive)' } }
    return 'Kayit var, client yok'
}

function ConvertTo-TsView {
    # Istemciden okunan ham TS durumu (Get-TsClientState) -> gride yazilacak SABIT degerler + detay metni.
    # Oncelik: okunamadi -> DURUM_YOK | bu TS calisiyor -> CALISIYOR | calisma gecmisi -> $script:TsStateMap |
    #          gecmis yok + politika var -> BEKLIYOR | gecmis yok + politika yok -> HEDEF_DEGIL
    param($S)
    $v = [pscustomobject]@{ Durum = ''; Adim = ''; Kod = ''; Zaman = ''; Detail = '' }
    $raw = ''
    if (-not $S -or -not $S.Ok) { $v.Durum = 'DURUM_YOK' }
    else {
        $h = $S.Hist
        if ($h) { $raw = [string]$h['_State'] }
        if ($S.HistTime) { $v.Zaman = $S.HistTime.ToString('yyyy-MM-dd HH:mm') } elseif ($h) { $v.Zaman = [string]$h['_RunStartTime'] }
        if ($S.RunningThis) { $v.Durum = 'CALISIYOR' }
        elseif ($h) {
            foreach ($m in $script:TsStateMap) { if ($raw -match $m[0]) { $v.Durum = $m[1]; break } }
            if (-not $v.Durum) { $v.Durum = $(if (Format-HexCode $h['SuccessOrFailureCode']) { 'HATA' } else { 'DURUM_YOK' }) }
        }
        elseif ($S.Policy) { $v.Durum = 'BEKLIYOR' }
        else { $v.Durum = 'HEDEF_DEGIL' }
        if ($v.Durum -eq 'HATA' -and $h) { $v.Kod = Format-HexCode $h['SuccessOrFailureCode'] }
        $st = $S.Step
        if ($st -and $v.Durum -eq 'CALISIYOR' -and $st.Last) { $v.Adim = [string](Get-WmiProp $st.Last 'ActionName') }
        elseif ($st -and $v.Durum -eq 'HATA') {
            $f = if ($st.Fail) { $st.Fail } else { $st.Last }
            if ($f) { $v.Adim = [string](Get-WmiProp $f 'ActionName'); if (-not $v.Kod) { $v.Kod = Format-HexCode (Get-WmiProp $f 'ExitCode') } }
        }
        if (-not $v.Adim -and $v.Durum -in 'CALISIYOR', 'HATA') {   # neden SABIT ifadeyle (ayrinti detayda)
            $v.Adim = if (-not $st) { 'SCCM baglantisi yok' }
                      elseif ($st.Error -eq 'SCCM kaydi yok') { 'SCCM kaydi yok' }
                      elseif ($st.Error) { 'Adim sorgulanamadi' }
                      else { 'Adim henuz SCCM''de yok' }
        }
    }
    $d = "TS durumu (istemci): $($v.Durum)"
    if ($S -and $S.Ok) {
        $d += "  | Politika: $(if ($S.Policy) { 'var' } else { 'yok' })  | Calisma gecmisi: $($S.HistCount) kayit"
        if ($S.Hist) { $d += "`r`n  En yeni: _State=$raw  _RunStartTime=$($S.Hist['_RunStartTime'])  SuccessOrFailureCode=$($S.Hist['SuccessOrFailureCode']) ($(Format-HexCode $S.Hist['SuccessOrFailureCode']))  SuccessOrFailureReason=$($S.Hist['SuccessOrFailureReason'])" }
    }
    if ($S -and $S.Step) {
        $st = $S.Step
        $d += "`r`n  Adim bilgisi (SMS Provider, istemci her adim BITINCE bildirir): $($st.Count) kayit$(if ($st.Error) { " - $($st.Error)" })"
        foreach ($p in @(@('Son bildirilen adim', $st.Last), @('Basarisiz adim', $st.Fail))) {
            $x = $p[1]; if (-not $x) { continue }
            $o = [string](Get-WmiProp $x 'ActionOutput'); if ($o.Length -gt 1500) { $o = $o.Substring(0, 1500) + ' ...' }
            $d += "`r`n  $($p[0]): Step=$(Get-WmiProp $x 'Step') '$(Get-WmiProp $x 'ActionName')' Grup='$(Get-WmiProp $x 'GroupName')' ExitCode=$(Get-WmiProp $x 'ExitCode') ($(Format-HexCode (Get-WmiProp $x 'ExitCode'))) Zaman=$(if ($x._T) { $x._T.ToString('yyyy-MM-dd HH:mm:ss') })" +
                  "`r`n    Mesaj: $(Get-WmiProp $x 'LastStatusMessageID') $(Get-WmiProp $x 'LastStatusMessageIDName')" + $(if ($o) { "`r`n    Cikti: $o" } else { '' })
        }
    }
    elseif ($v.Durum -in 'CALISIYOR', 'HATA') { $d += "`r`n  Adim bilgisi: SCCM'e bagli degil ('Baglan') - adim adi SMS Provider'dan okunur." }
    if ($S -and $S.Step -and @($S.Step.AdvIds).Count) { $d += "`r`n  Adim sorgusundaki deployment(lar): $($S.Step.AdvIds -join ', ')" }
    if ($S -and $S.Detail) { $d += "`r`n$($S.Detail)" }
    $v.Detail = $d
    return $v
}
#endregion

#region Paralel calistirici
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

function Start-Snapshot {
    # SMS Provider TOPLU sorgusu (arka planda, tek is): kayit + Boundary Group(s) / OS Build + koleksiyon uyelikleri.
    # Bitince Apply-Snapshot; -Next verilmisse ardindan cihaz isleri baslar.
    param([string[]]$Hosts, [string]$Title, $Next)
    $ps = [powershell]::Create()
    $ps.RunspacePool = (Get-SharedPool)
    [void]$ps.AddScript($script:SnapshotWorker).AddParameters(@{
        Names = [string[]]$Hosts; SmsProvider = $script:SmsServer; SiteCode = $script:SiteCode
        CollectionIds = [string[]]@($script:Collections | ForEach-Object { $_.Id }) })
    [void]$script:Prefetch.Add([pscustomobject]@{ Panel = 1; Hosts = $Hosts; Next = $Next; Title = $Title; Job = [pscustomobject]@{ PS = $ps; Async = $ps.BeginInvoke(); Key = 'sccm'; Done = $false } })
    Set-UiBusy -Panel 1 -Busy $true
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue
    $lblStatus.Text = "$Title`: SMS Provider toplu sorgusu ($(@($Hosts).Count) cihaz)..."
    if (-not $script:RunTimer.Enabled) { $script:RunTimer.Start() }
}

function Update-Prefetch {
    foreach ($pf in @($script:Prefetch)) {
        if (-not $pf.Job.Async.IsCompleted) { continue }
        [void]$script:Prefetch.Remove($pf)
        $r = Complete-Job $pf.Job
        $snap = @($r.Out | Where-Object { $_ -and $_.PSObject.Properties['Snapshot'] })[0]
        if ($snap) { Apply-Snapshot -Snap $snap -Hosts $pf.Hosts }
        else {
            foreach ($h in $pf.Hosts) {
                Add-RowDetail -Hostname $h -Text "SMS Provider sorgusu: $(Format-JobError $r.Err)"
                $v = @{ SCCM_Kayit = 'Sorgulanamadi'; Uye_Koleksiyonlar = ''; Boundary_Groups = ''; OS_Build = ''; Hata_Not = (Get-ShortReason -Text (Format-JobError $r.Err)) }
                Set-GridValues -Hostname $h -Values $v
            }
        }
        if ($pf.Next) { $a = $pf.Next.RunArgs; Start-ParallelRun @a -Items $pf.Next.Items }
        else {
            Set-UiBusy -Panel 1 -Busy $false
            $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
            $lblStatus.Text = "$($pf.Title) tamamlandi. $(@($pf.Hosts).Count) cihaz."
        }
    }
}

function Format-Membership {
    # Tanimli koleksiyonlardan direct uye olunanlarin KISA adlari (; ile). Uye yoksa '-'.
    param($Coll, $Rid)
    if (-not $Coll -or $null -eq $Rid -or "$Rid" -eq '') { return '' }
    $n = @(foreach ($c in $script:Collections) { $m = $Coll[$c.Id]; if ($m -and $m.Ok -and $m.Rids.Contains([uint32]$Rid)) { $c.Short } })
    if ($n.Count -eq 0) { return '-' }
    return ($n -join '; ')
}

function Apply-Snapshot {
    param($Snap, [string[]]$Hosts)
    $collErr = @(if ($Snap.Coll) { foreach ($c in $script:Collections) { $m = $Snap.Coll[$c.Id]; if ($m -and -not $m.Ok) { "$($c.Short): $($m.Error)" } } })
    foreach ($h in $Hosts) {
        $k = (($h -split '\.')[0]).Trim()
        $info = $Snap.Devices[$k]
        $ref = if ($info -and $info.Dev) { $info.Dev } elseif ($info -and @($info.Cands).Count -gt 0) { $info.Cands[0] } else { $null }
        $rid = if ($ref) { [string]$ref.ResourceId } else { '' }
        $v = @{
            SCCM_Kayit = (Get-SccmKayitText $info); ResourceId = $rid; Uye_Koleksiyonlar = $(if ($rid) { Format-Membership $Snap.Coll $rid } else { '' })
            Boundary_Groups = $(if ($info) { [string]$info.BoundaryGroups } else { '' }); OS_Build = $(if ($info) { [string]$info.OsBuild } else { '' })
        }
        $d = "SCCM: $($v.SCCM_Kayit) (ResourceID=$rid)$(if ($info -and -not $info.Ok) { " - $($info.Error)" })  | Boundary Group(s): $($v.Boundary_Groups)  | OS Build: $($v.OS_Build)"
        if ($Snap.Coll) { $d += "`r`nUye koleksiyonlar (direct): $($v.Uye_Koleksiyonlar)" + $(if ($collErr.Count) { "  | Okunamayan: $($collErr -join '; ')" } else { '' }) }
        Set-GridValues -Hostname $h -Values $v
        Add-RowDetail -Hostname $h -Text $d
    }
}

function Stop-Run {
    # IPTAL: bekleyen / calisan isler durdurulur; tamamlananlar korunur. Gonderilmis TS tetikleme / koleksiyon degisikligi geri ALINAMAZ.
    $runs = @($script:ActiveRuns)
    $pfs  = @($script:Prefetch)
    if ($runs.Count -eq 0 -and $pfs.Count -eq 0) { return }
    $left = 0
    foreach ($pf in $pfs) { $left += @($pf.Hosts).Count }
    foreach ($st in $runs) { $left += @($st.Jobs | Where-Object { -not $_.Done }).Count }
    $c = [System.Windows.Forms.MessageBox]::Show("$left is henuz tamamlanmadi. Iptal edilsin mi?`n`nTamamlanan sonuclar tabloda kalir. Gonderilmis TS tetikleme / koleksiyon degisikligi geri alinamaz.",
        'Iptal', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($c -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    $n = 0
    foreach ($pf in $pfs) {
        [void]$script:Prefetch.Remove($pf)
        try { [void]$pf.Job.PS.BeginStop($null, $null) } catch { }
        [void]$script:CancelledJobs.Add($pf.Job)
        Resolve-HostPending -Hosts $pf.Hosts -ErrorText $script:CancelText
        $n += @($pf.Hosts).Count
    }
    foreach ($st in $runs) {
        foreach ($job in $st.Jobs) {
            if ($job.Done) { continue }
            $job.Done = $true
            if ($job.Async.IsCompleted) { $r = Complete-Job $job; Invoke-JobResult -Run $st -Job $job -Output $r.Out -ErrorText $r.Err -Label $lblStatus; continue }
            try { [void]$job.PS.BeginStop($null, $null) } catch { }
            [void]$script:CancelledJobs.Add($job)
            Invoke-JobResult -Run $st -Job $job -Output @() -ErrorText $script:CancelText -Label $lblStatus
            $n++
        }
        [void]$script:ActiveRuns.Remove($st)
        Write-ActionLog -Hostname '-' -Action "$($st.Title) - IPTAL" -State 'warn' -Text "$n is iptal edildi"
    }
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkOrange
    $lblStatus.Text = "Iptal edildi: $n is durduruldu, tamamlanan sonuclar tabloda."
    Set-UiBusy -Panel 1 -Busy $false
}

$script:RunTimer.Add_Tick({
    foreach ($z in @($script:CancelledJobs)) { if ($z.Async.IsCompleted) { $null = Complete-Job $z; [void]$script:CancelledJobs.Remove($z) } }
    if ($script:Prefetch.Count -gt 0) { Update-Prefetch }
    foreach ($st in @($script:ActiveRuns)) {
        $pending = $false
        foreach ($job in $st.Jobs) {
            if ($job.Done) { continue }
            if (-not $job.Async.IsCompleted) { $pending = $true; continue }
            $job.Done = $true
            $r = Complete-Job $job
            Invoke-JobResult -Run $st -Job $job -Output $r.Out -ErrorText $r.Err -Label $lblStatus
        }
        $txt = "$($st.Title): $($st.Done) / $($st.Total) tamamlandi (ayni anda en fazla $($script:MaxParallel) istek)"
        if ($lblStatus.Text -ne $txt -and $pending) { $lblStatus.Text = $txt }
        if (-not $pending) {
            [void]$script:ActiveRuns.Remove($st)
            if ($st.Unit -eq 'makine') { $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen; $lblStatus.Text = "$($st.Title) tamamlandi. $($st.Total) $($st.Unit) islendi." }
            Set-UiBusy -Panel 1 -Busy $false
        }
    }
    if ($script:ActiveRuns.Count -eq 0 -and $script:CancelledJobs.Count -eq 0 -and $script:Prefetch.Count -eq 0) { $script:RunTimer.Stop() }
})
#endregion

#region Sonuc isleyiciler (UI thread'i)

function Update-FromCheck {
    param($Key, $Output, $ErrorText)
    $r = @($Output | Where-Object { $_ -and $_.PSObject.Properties['DnsStatus'] })[0]
    if (-not $r) {
        $msg = Format-JobError $ErrorText
        Resolve-HostPending -Hosts @($Key) -ErrorText $ErrorText
        Set-GridValues -Hostname $Key -Values @{ Hata_Not = (Get-ShortReason -Text $msg) }
        Add-RowDetail -Hostname $Key -Text "Kontrol: $msg"
        return
    }
    Set-AdInfo $r
    $note = Get-ShortReason -Text $r.ErrorMessage -Context @($r.DnsStatus, $r.OnlineText)
    $v = @{
        DNS_Durumu = (Get-ShortDns $r.DnsStatus); Online_mi = (Get-ShortOnline $r.OnlineText); AD_Durumu = $r.AdStatus; AD_LastLogon = $r.LastLogon
        CcmExec_Durumu = $r.CcmExec; OS_Build_Anlik = $(if ($r.OsLive) { $r.OsLive.Build } else { '' })
    }
    $d = "Kontrol: DNS=$($r.DnsStatus); Online=$($r.OnlineText); AD=$($r.AdStatus); LastLogon=$($r.LastLogon); CcmExec=$($r.CcmExec)" + $(if ($r.ErrorMessage) { "`r`nNot: $($r.ErrorMessage)" } else { '' })
    if ($r.OsLive) { $d += "`r`n$($r.OsLive.Detail)" }
    if ($r.Ts) {
        $tv = ConvertTo-TsView -S $r.Ts
        $v['TS_Durum'] = $tv.Durum; $v['TS_Adim'] = $tv.Adim; $v['TS_HataKodu'] = $tv.Kod; $v['TS_SonGuncelleme'] = $tv.Zaman
        if (-not $note -and $r.IsOnline -and -not $r.Ts.Ok) { $note = Get-ShortReason -Text $r.Ts.Error }   # cevrimici ama TS okunamadi (yetki / WMI)
        $d += "`r`nTS ($($script:CheckTsName)):`r`n$($tv.Detail)"
    }
    $v['Hata_Not'] = $note
    Set-GridValues -Hostname $Key -Values $v
    Add-RowDetail -Hostname $Key -Text $d
}

function Update-FromTrigger {
    param($Key, $Output, $ErrorText)
    $ts = [string]$script:TriggerTitle
    $r = @($Output | Where-Object { $_ -and $_.PSObject.Properties['Code'] })[0]
    if (-not $r) {
        $msg = Format-JobError $ErrorText
        Resolve-HostPending -Hosts @($Key) -ErrorText $ErrorText
        Set-GridValues -Hostname $Key -Values @{ Son_Islem = "TS tetikle: $ts"; Islem_Sonucu = $(if ($ErrorText -eq $script:CancelText) { 'Iptal edildi' } else { 'Hata (Islem hatasi)' }) }
        Add-RowDetail -Hostname $Key -Text "TS tetikle ($ts): $msg"
        Write-ActionLog -Hostname $Key -Action "TS tetikle [$ts]" -State 'fail' -Text $msg
        return
    }
    $short = Get-ShortReason -Text $r.Reason -Context @($r.DnsStatus, $r.OnlineText)
    $res = switch ($r.Code) {
        'Tetiklendi'   { 'Tetiklendi' }
        'DryRun'       { 'DRY-RUN: tetiklenecek' }
        'Politika yok' { 'Politika yok' }
        'Atlandi'      { "Atlandi ($short)" }
        default        { "Hata ($short)" }
    }
    $v = @{ Son_Islem = "TS tetikle: $ts"; Islem_Sonucu = $res }
    if ($r.DnsStatus) { $v['DNS_Durumu'] = Get-ShortDns $r.DnsStatus }
    if ($r.OnlineText) { $v['Online_mi'] = Get-ShortOnline $r.OnlineText }
    if ($r.CcmExec) { $v['CcmExec_Durumu'] = $r.CcmExec }
    Set-GridValues -Hostname $Key -Values $v
    Add-RowDetail -Hostname $Key -Text ("TS tetikle ($ts): $($r.Code)$(if ($r.Reason) { " - $($r.Reason)" })`r`nDNS=$($r.DnsStatus); Online=$($r.OnlineText); CcmExec=$($r.CcmExec)" + $(if ($r.Detail) { "`r`n$($r.Detail)" } else { '' }))
    Write-ActionLog -Hostname $Key -Action "TS tetikle [$ts]" -State $(if ($r.Code -eq 'Tetiklendi') { 'ok' } elseif ($r.Code -eq 'Hata') { 'fail' } else { 'warn' }) -Text "$res $($r.Reason)"
}

function Update-FromAction {
    param($Key, $Output, $ErrorText)
    $r = @($Output | Where-Object { $_ -and $_.PSObject.Properties['Action'] })[0]
    if (-not $r) { $r = [pscustomobject]@{ Hostname = $Key; Action = 'Istemci eylemi'; State = 'fail'; Result = (Format-JobError $ErrorText); Detail = '' } }
    $res = if ($ErrorText -eq $script:CancelText) { 'Iptal edildi' }
           elseif ($r.State -eq 'fail' -and "$($r.Result)" -match '^(Cevrimdisi|DNS:)') { "Atlandi ($(Get-ShortReason -Text $r.Result))" }
           elseif ($r.State -eq 'fail') { "Hata ($(Get-ShortReason -Text $r.Result))" }
           else { Get-ShortActionResult -Result $r.Result -State $r.State }
    Set-GridValues -Hostname $Key -Values @{ Son_Islem = [string]$r.Action; Islem_Sonucu = $res }
    Add-RowDetail -Hostname $Key -Text ("$($r.Action): $($r.Result)" + $(if ($r.Detail) { "`r`n$($r.Detail)" } else { '' }))
    Write-ActionLog -Hostname $Key -Action "Istemci eylemi [$($r.Action)]" -State $r.State -Text $r.Result
}

function Update-FromDetail {
    # Tek cihaz bilgi isleri (TS istemci detayi / TS adim gecmisi): satir detayina eklenir ve pencerede gosterilir.
    param($Key, $Output, $ErrorText)
    $t = @($Output | Where-Object { $_ -is [string] }) -join "`r`n"
    if (-not $t) { $t = Format-JobError $ErrorText }
    $h = [string]$script:DetailHost; $ttl = [string]$script:DetailTitle
    Add-RowDetail -Hostname $h -Text "${ttl}:`r`n$t"
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
    $lblStatus.Text = "$ttl alindi: $h"
    $script:ShowAfterRun = [pscustomobject]@{ Title = "${ttl}: $h"; Text = $t }
}

function Update-FromCollection {
    param($Key, $Output, $ErrorText)
    $op = $script:CollOps[$Key]
    if (-not $op) { return }
    $script:CollOps.Remove($Key)
    $r = @($Output | Where-Object { $_ -and $_.PSObject.Properties['CollOp'] })[0]
    $verb = if ($op.Mode -eq 'Add') { 'Koleksiyona ekle' } else { 'Koleksiyondan cikar' }
    if (-not $r) {
        $msg = Format-JobError $ErrorText
        foreach ($h in $op.Hosts) {
            Set-GridValues -Hostname $h -Values @{ Son_Islem = "${verb}: $($op.Short)"; Islem_Sonucu = $(if ($ErrorText -eq $script:CancelText) { 'Iptal edildi' } else { 'Islem hatasi' }) }
            Add-RowDetail -Hostname $h -Text "$verb ($($op.Short) [$($op.Id)]): $msg$(if ($ErrorText -eq $script:CancelText) { ' (iptal edildi; SMS Provider tarafinda islem tamamlanmis olabilir - Kontrol Et ile uyelige bakin)' })"
        }
        Write-ActionLog -Hostname '-' -Action "$verb [$($op.Id)]" -State 'fail' -Text $msg
        $lblStatus.ForeColor = [System.Drawing.Color]::Firebrick; $lblStatus.Text = "${verb}: $msg"
        return
    }
    foreach ($x in $r.Results) {
        $v = @{ Son_Islem = "${verb}: $($op.Short)"; Islem_Sonucu = $x.Code }
        if ($null -ne $x.Rid) { $v['ResourceId'] = [string]$x.Rid; if ($r.Coll) { $v['Uye_Koleksiyonlar'] = Format-Membership $r.Coll $x.Rid } }
        $di = if ($r.Devices) { $r.Devices[(($x.Host -split '\.')[0]).Trim()] } else { $null }
        if ($di -and $di.Ok) { $v['SCCM_Kayit'] = Get-SccmKayitText $di; $v['Boundary_Groups'] = [string]$di.BoundaryGroups; $v['OS_Build'] = [string]$di.OsBuild }
        Set-GridValues -Hostname $x.Host -Values $v
        Add-RowDetail -Hostname $x.Host -Text "$verb ($($op.Short) [$($op.Id)]): $($x.Code) - $($x.Detail)"
        Write-ActionLog -Hostname $x.Host -Action "$verb [$($op.Id) $($op.Short)]" -State $(if ($x.Code -in 'Eklendi', 'Cikarildi') { 'ok' } elseif ($x.Code -in 'Yetki yok', 'Islem hatasi') { 'fail' } else { 'warn' }) -Text $x.Code
    }
    $phr = [ordered]@{ 'Eklendi' = 'eklendi'; 'Cikarildi' = 'cikarildi'; 'Zaten uye' = 'zaten uyeydi'; 'Uye degil' = 'uye degildi'; 'SCCM kaydi yok' = 'SCCM kaydi yok'
        'Direct uye degil, cikarilamaz' = 'direct uye degil (cikarilamaz)'; 'DRY-RUN: eklenecek' = 'eklenecek (DRY-RUN)'; 'DRY-RUN: cikarilacak' = 'cikarilacak (DRY-RUN)'
        'Yetki yok' = 'yetki yok'; 'Islem hatasi' = 'islem hatasi' }
    $parts = @(foreach ($k in $phr.Keys) { $n = @($r.Results | Where-Object { $_.Code -eq $k }).Count; if ($n) { "$n $($phr[$k])" } })
    $sum = "$($op.Short): " + ($parts -join ', ') + $(if ($r.Refreshed) { ' | koleksiyon yenileme istendi' } else { '' })
    Write-ActionLog -Hostname '-' -Action "$verb [$($op.Id) $($op.Short)] OZET" -State $(if ($r.Ok) { 'ok' } else { 'fail' }) -Text $sum
    $lblStatus.ForeColor = $(if ($r.Ok) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::Firebrick })
    $lblStatus.Text = "${verb} tamamlandi. $sum"
    $errs = @($r.Results | Where-Object { $_.Code -in 'Yetki yok', 'Islem hatasi' } | ForEach-Object { "$($_.Host): $($_.Detail)" } | Select-Object -First 3)
    $script:ShowAfterRun = [pscustomobject]@{ Title = $verb; Text = "$sum$(if ($r.Error) { "`r`n`r`nHata: $($r.Error)" })$(if ($errs.Count) { "`r`n`r`nIlk hatalar (tamami: satira cift tik):`r`n" + ($errs -join "`r`n") })"; Box = $true }
}
#endregion

#region AD verisi
function Start-AdLoad {
    # AD verisi ARKA PLANDA (ayri runspace; ana ekran kilitlenmez), sirayla:
    #   $data = Get-ADComputer -Filter * -SearchBase $sb -Server $domain ...   -> nesneler OLDUGU GIBI ana pencereye doner
    # Bitince $script:AdData = @(); $script:AdData += ... (Update-AdLoad). Bir OU hata verirse onceki verisi korunur.
    $src = @($script:AdCacheSources)
    if ($src.Count -eq 0) { Update-AdCacheLabel; return }
    if ($script:AdLoadJob) { return }
    $ps = [powershell]::Create()
    [void]$ps.AddScript({
        param($Sources)
        Import-Module -Name ActiveDirectory -ErrorAction Stop -Verbose:$false
        for ($i = 0; $i -lt $Sources.Count; $i++) {
            $d = $Sources[$i].Domain; $sb = $Sources[$i].OU
            [System.AppDomain]::CurrentDomain.SetData('ClientTool.AdLoadStatus', "$($i + 1) / $($Sources.Count)  $d")
            $t = [System.Diagnostics.Stopwatch]::StartNew()
            try {
                $data = @(Get-ADComputer -Filter * -SearchBase $sb -Server $d -ResultSetSize $null -ResultPageSize 1000 -Properties Enabled, LastLogonTimestamp -ErrorAction Stop)
                [pscustomobject]@{ Domain = $d; OU = $sb; Ok = $true; Data = $data; Seconds = [int]$t.Elapsed.TotalSeconds; Error = '' }
            }
            catch { [pscustomobject]@{ Domain = $d; OU = $sb; Ok = $false; Data = @(); Seconds = [int]$t.Elapsed.TotalSeconds; Error = $_.Exception.Message } }
        }
    }).AddParameter('Sources', [object[]]@($src | ForEach-Object { [pscustomobject]@{ Domain = [string]$_.Domain; OU = [string]$_.OU } }))
    $script:AdLoadJob = [pscustomobject]@{ PS = $ps; Async = $ps.BeginInvoke(); Key = 'ad'; Done = $false; Start = (Get-Date) }
    $btnAdCache.Enabled = $false
    $lblAdCache.ForeColor = [System.Drawing.Color]::DarkBlue
    $lblAdCache.Text = 'AD verisi yukleniyor...'
}

function Update-AdLoad {
    # SummaryTimer'dan (0,5 sn): yukleme surerken etiket; bitince $AdData + ozet + acik tablolardaki AD kolonlari yenilenir.
    $j = $script:AdLoadJob
    if (-not $j) { return }
    if (-not $j.Async.IsCompleted) {
        $lblAdCache.Text = "AD yukleniyor: $([System.AppDomain]::CurrentDomain.GetData('ClientTool.AdLoadStatus'))  ($([int]((Get-Date) - $j.Start).TotalSeconds) sn)"
        return
    }
    $r = Complete-Job $j
    $script:AdLoadJob = $null
    $btnAdCache.Enabled = $true
    $res = @($r.Out | Where-Object { $_ -and $_.PSObject.Properties['Data'] })
    $old = $script:AdData
    $new = @(); $done = @(); $errs = @()
    foreach ($x in $res) {
        if ($x.Ok) { $new += $x.Data; $done += "{0}: {1:N0} ({2} sn)" -f $x.Domain, @($x.Data).Count, $x.Seconds }
        else {
            $errs += "$($x.Domain) [$($x.OU)]: $($x.Error)"; $done += "$($x.Domain): HATA"
            if ($old) { $new += $old.Where({ ([string]$_.DistinguishedName).EndsWith($x.OU, [System.StringComparison]::OrdinalIgnoreCase) }) }
        }
    }
    if ($r.Err -and $res.Count -eq 0) { $errs += $r.Err }   # or. AD modulu yuklenemedi
    if ($res.Count -gt 0 -and @($res | Where-Object { $_.Ok }).Count -gt 0) { $script:AdData = $new; $script:AdDataTime = Get-Date }
    Update-AdCacheLabel
    Update-AdColumns
    $msg = "AD verisi: {0}/{1} OU basarili, {2:N0} makine ({3} sn) | {4}" -f @($res | Where-Object { $_.Ok }).Count, @($script:AdCacheSources).Count, @($script:AdData).Count, [int]((Get-Date) - $j.Start).TotalSeconds, ($done -join ' | ')
    $lblStatus.ForeColor = $(if ($errs.Count) { [System.Drawing.Color]::DarkOrange } else { [System.Drawing.Color]::DarkGreen })
    $lblStatus.Text = $msg
    Write-ActionLog -Hostname '-' -Action 'AD verisi yukleme' -State $(if (-not @($res | Where-Object { $_.Ok }).Count) { 'fail' } elseif ($errs.Count) { 'warn' } else { 'ok' }) -Text $msg
    if ($errs.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show("Bazi OU'lar alinamadi (onceki verileri korundu):`n`n" + ($errs -join "`n`n"), 'AD Verisi', 'OK', 'Warning') | Out-Null
    }
}

function Set-AdInfo {
    # Worker sonucuna AD kolonlarini ana pencerede $AdData'dan ekler (AdStatus / LastLogon, not; istenirse State).
    param($Obj, [string]$NoteProp = 'ErrorMessage', [switch]$Restate)
    if (-not $Obj -or -not $Obj.PSObject.Properties['AdStatus']) { return }
    $a = Get-AdComputerInfo -ComputerName ([string]$Obj.Hostname)
    $Obj.AdStatus  = $a.Status
    $Obj.LastLogon = $a.LastLogon
    if ($a.Detail -and $Obj.PSObject.Properties[$NoteProp]) { $Obj.$NoteProp = (@($a.Detail, $Obj.$NoteProp) | Where-Object { $_ }) -join ' | ' }
    if ($Restate -and $Obj.State -eq 'ok' -and $a.Status -ne 'Enabled') { $Obj.State = 'warn' }
}

function Update-AdCacheLabel {
    # Bellekteki AD verisinin durumu kontrol grubundaki etikete.
    if ($null -ne $script:AdData) {
        $lblAdCache.Text      = "AD verisi: {0:N0} makine | yuklendi {1}" -f $script:AdData.Count, $script:AdDataTime.ToString('dd.MM HH:mm')
        $lblAdCache.ForeColor = [System.Drawing.Color]::DimGray
    }
    elseif (@($script:AdCacheSources).Count -eq 0) {
        $lblAdCache.Text      = "Domain / OU tanimli degil - 'AD Verisi...' ile ekleyin"
        $lblAdCache.ForeColor = [System.Drawing.Color]::Firebrick
    }
    else {
        $lblAdCache.Text      = 'AD verisi yok - yuklenmedi'
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
    # "Kaydet ve AD Verisini Yukle" ayrica AD'den okuyup bellekteki AD verisini yeniler.
    $dlg = New-Ctl Form 'AD Verisi - Domain / OU Eslemesi' 0 0 860 450 @{ StartPosition = 'CenterParent'; MinimumSize = (New-Object System.Drawing.Size(700, 360)); ShowInTaskbar = $false }
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

    $info = if ($null -ne $script:AdData) { "Bellekteki AD verisi: $($script:AdData.Count) makine, yuklendi $($script:AdDataTime.ToString('yyyy-MM-dd HH:mm')). Dosyaya yazilmaz; program her acilista yukler." } else { "Bellekte AD verisi yok. 'Kaydet ve AD Verisini Yukle' ile yukleyin (program acilisinda da otomatik yuklenir)." }
    $lblInfo    = New-Ctl Label $info 12 332 820 34 @{ Anchor = 'Bottom,Left,Right'; ForeColor = [System.Drawing.Color]::DimGray }
    $btnSave    = New-Ctl Button 'Kaydet' 12 374 110 30 @{ Anchor = 'Bottom,Left' }
    $btnSaveUpd = New-Ctl Button 'Kaydet ve AD Verisini Yukle' 130 374 210 30 @{ Anchor = 'Bottom,Left' }
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
    if ($doUpdate) { Start-AdLoad }   # pencere kapandiktan sonra, arka planda
}

function Update-AdColumns {
    foreach ($row in @($script:ResultsTable.Rows)) {
        $h = [string]$row['Hostname']; if (-not $h -or [string]$row['AD_Durumu'] -eq $script:PendingText -or -not [string]$row['AD_Durumu']) { continue }
        $a = Get-AdComputerInfo -ComputerName $h
        Set-GridValues -Hostname $h -Values @{ AD_Durumu = $a.Status; AD_LastLogon = $a.LastLogon }
    }
}
#endregion

#region SCCM baglantisi / TS ve koleksiyon tanimlari
function Get-CmSiteCode {
    param([string]$Server)
    $locs = @(Get-WmiObject -ComputerName $Server -Namespace 'root\SMS' -Class 'SMS_ProviderLocation' -ErrorAction Stop)
    $loc = $locs | Where-Object { $_.ProviderForLocalSite } | Select-Object -First 1
    if (-not $loc) { $loc = $locs | Select-Object -First 1 }
    if (-not $loc) { throw 'SMS_ProviderLocation bulunamadi (SMS Provider adini kontrol edin).' }
    return [string]$loc.SiteCode
}

function Start-Connect {
    # SMS Provider'a arka planda baglanir (site kodu); UI donmaz. Sonuc SummaryTimer'da (Update-Connect).
    param([switch]$Quiet)
    $srv = $txtSms.Text.Trim()
    if ($srv -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]*$') { if (-not $Quiet) { [System.Windows.Forms.MessageBox]::Show('Gecerli bir SMS Provider adi girin.', 'Uyari') | Out-Null }; return }
    if ($script:ConnectJob) { return }
    $ps = [powershell]::Create()
    [void]$ps.AddScript([scriptblock]::Create("param(`$Server)`nfunction Get-CmSiteCode {`n$((Get-Item function:Get-CmSiteCode).ScriptBlock)`n}`nGet-CmSiteCode -Server `$Server")).AddParameter('Server', $srv)
    $script:ConnectJob = [pscustomobject]@{ PS = $ps; Async = $ps.BeginInvoke(); Server = $srv; Quiet = [bool]$Quiet }
    $btnConnect.Enabled = $false
    $lblSite.Text = 'Baglaniliyor...'
}

function Update-Connect {
    $j = $script:ConnectJob
    if (-not $j -or -not $j.Async.IsCompleted) { return }
    $script:ConnectJob = $null
    $r = Complete-Job $j
    $btnConnect.Enabled = $true
    $site = [string]@($r.Out)[0]
    if ($site -and -not $r.Err) {
        $script:SmsServer = $j.Server; $script:SiteCode = $site
        $lblSite.Text = "Site: $site"; $lblSite.ForeColor = [System.Drawing.Color]::DarkGreen
        Save-Settings
        Write-ActionLog -Hostname '-' -Action 'SMS Provider baglanti' -State 'ok' -Text "$($j.Server) site $site"
    }
    else {
        $script:SmsServer = ''; $script:SiteCode = ''
        $lblSite.Text = 'Baglanilamadi'; $lblSite.ForeColor = [System.Drawing.Color]::Firebrick
        if (-not $j.Quiet) { [System.Windows.Forms.MessageBox]::Show("SMS Provider'a baglanilamadi ($($j.Server)):`n$($r.Err)", 'Hata') | Out-Null }
    }
    Update-ButtonStates
}

function Test-Connected {
    if ($script:SmsServer -and $script:SiteCode) { return $true }
    [System.Windows.Forms.MessageBox]::Show("Once SMS Provider'a baglanin ('Baglan').", 'Uyari') | Out-Null
    return $false
}

function Invoke-CmQuery {
    # UI thread'inde kisa SMS Provider sorgusu (sadece tanim pencerelerinde dogrulama icin).
    param([string]$Query)
    return @(Get-WmiObject -ComputerName $script:SmsServer -Namespace "root\SMS\site_$($script:SiteCode)" -Query $Query -ErrorAction Stop)
}

function Test-TsDefinition {
    # TS tanimini SMS Provider'dan dogrular. Doner: { Ok; Error; PackageId; PkgName; Text }
    param([string]$PackageId, [string]$DeploymentId)
    $res = [pscustomobject]@{ Ok = $false; Error = ''; PackageId = $PackageId; PkgName = ''; Text = '' }
    try {
        if ($DeploymentId) {
            $a = @(Invoke-CmQuery "SELECT AdvertisementID, PackageID, CollectionID FROM SMS_Advertisement WHERE AdvertisementID = '$(ConvertTo-WqlString $DeploymentId)'")
            if ($a.Count -eq 0) { $res.Error = "Deployment bulunamadi: $DeploymentId"; return $res }
            if ($PackageId -and $PackageId -ne [string]$a[0].PackageID) { $res.Error = "Deployment $DeploymentId baska pakete ait: $($a[0].PackageID)"; return $res }
            $res.PackageId = [string]$a[0].PackageID
        }
        $p = @(Invoke-CmQuery "SELECT PackageID, Name FROM SMS_TaskSequencePackage WHERE PackageID = '$(ConvertTo-WqlString $res.PackageId)'")
        if ($p.Count -eq 0) { $res.Error = "Task Sequence paketi bulunamadi: $($res.PackageId)"; return $res }
        $res.PkgName = [string]$p[0].Name
        $w = if ($DeploymentId) { "AdvertisementID = '$(ConvertTo-WqlString $DeploymentId)'" } else { "PackageID = '$(ConvertTo-WqlString $res.PackageId)'" }
        $deps = @(Invoke-CmQuery "SELECT AdvertisementID, CollectionID FROM SMS_Advertisement WHERE $w")
        $dl = @(foreach ($x in $deps) {
            $cn = ''; try { $cn = [string](@(Invoke-CmQuery "SELECT Name FROM SMS_Collection WHERE CollectionID = '$($x.CollectionID)'")[0].Name) } catch { }
            "$($x.AdvertisementID) -> $($x.CollectionID) ($cn)"
        })
        $res.Text = "TS: $($res.PkgName) [$($res.PackageId)]`r`nDeployment(lar): $(if ($dl.Count) { $dl -join '; ' } else { '(deployment yok)' })"
        $res.Ok = $true
    }
    catch { $res.Error = $(if (Test-AccessError $_.Exception.Message) { 'Yetki yok' } else { $_.Exception.Message }) }
    return $res
}

function Test-CollectionDefinition {
    # Koleksiyonu ID veya ad ile bulur; Device Collection (CollectionType 2) olmali. Doner: { Ok; Error; Id; Name; Members }
    param([string]$Text)
    $res = [pscustomobject]@{ Ok = $false; Error = ''; Id = ''; Name = ''; Members = '' }
    try {
        $q = "SELECT CollectionID, Name, CollectionType, MemberCount FROM SMS_Collection WHERE "
        $c = @(Invoke-CmQuery ($q + "CollectionID = '$(ConvertTo-WqlString $Text)'"))
        if ($c.Count -eq 0) { $c = @(Invoke-CmQuery ($q + "Name = '$(ConvertTo-WqlString $Text)'")) }
        if ($c.Count -eq 0) { $res.Error = "Koleksiyon bulunamadi: $Text"; return $res }
        if ($c.Count -gt 1) { $res.Error = "Ayni adla birden fazla koleksiyon var, CollectionID girin: $Text"; return $res }
        if ([int]$c[0].CollectionType -ne 2) { $res.Error = "Device Collection degil (CollectionType=$($c[0].CollectionType)): $Text"; return $res }
        $res.Id = [string]$c[0].CollectionID; $res.Name = [string]$c[0].Name; $res.Members = [string]$c[0].MemberCount; $res.Ok = $true
    }
    catch { $res.Error = $(if (Test-AccessError $_.Exception.Message) { 'Yetki yok' } else { $_.Exception.Message }) }
    return $res
}

function Show-DefDialog {
    # Ortak tanim penceresi (TS veya koleksiyon). Kaydet: her satir SMS Provider'dan dogrulanir; hepsi gecerliyse kaydedilir.
    param([ValidateSet('Ts', 'Coll')][string]$Kind)
    if (-not (Test-Connected)) { return }
    $isTs = ($Kind -eq 'Ts')
    $dlg = New-Ctl Form $(if ($isTs) { 'Task Sequence Tanimlari' } else { 'Koleksiyonlar' }) 0 0 900 460 @{ StartPosition = 'CenterParent'; MinimumSize = (New-Object System.Drawing.Size(700, 360)); ShowInTaskbar = $false }
    $info = if ($isTs) { "Gorunen Ad + TS Package ID ve/veya Deployment ID (biri bos olabilir). Kaydet: SCCM'den dogrulanir, bulunan ad gosterilir. Satir silmek: satir basligini secip Delete." }
            else { "Kisa Ad (gridde gorunur) + Collection ID veya tam ad. Kaydet: SCCM'den dogrulanir (Device Collection olmali, uye sayisi gosterilir). Satir silmek: satir basligini secip Delete." }
    $lbl = New-Ctl Label $info 12 10 860 34 @{ Anchor = 'Top,Left,Right' }
    $grid = New-Ctl DataGridView '' 12 48 860 300 @{
        Anchor = 'Top,Bottom,Left,Right'; AllowUserToAddRows = $true; AllowUserToDeleteRows = $true; AllowUserToResizeRows = $false
        RowHeadersVisible = $true; SelectionMode = 'RowHeaderSelect'; BackgroundColor = [System.Drawing.Color]::White; AutoSizeColumnsMode = 'Fill'
    }
    $cols = if ($isTs) { @(@('Gorunen Ad', 25, $false), @('TS Package ID', 15, $false), @('Deployment ID', 15, $false), @('SCCM (dogrulama)', 45, $true)) }
            else { @(@('Kisa Ad', 20, $false), @('Collection ID veya Ad', 25, $false), @('SCCM (dogrulama)', 55, $true)) }
    foreach ($c in $cols) { $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $col.HeaderText = $c[0]; $col.FillWeight = $c[1]; $col.ReadOnly = $c[2]; [void]$grid.Columns.Add($col) }
    if ($isTs) { foreach ($t in $script:TsDefs) { [void]$grid.Rows.Add($t.Name, $t.PackageId, $t.DeploymentId, "$($t.PkgName) [$($t.PackageId)]") } }
    else { foreach ($c in $script:Collections) { [void]$grid.Rows.Add($c.Short, $c.Id, "$($c.Name) [$($c.Id)]") } }
    $btnSave  = New-Ctl Button 'Dogrula ve Kaydet' 12 360 160 30 @{ Anchor = 'Bottom,Left' }
    $btnClose = New-Ctl Button 'Kapat' 772 360 100 30 @{ Anchor = 'Bottom,Right'; DialogResult = [System.Windows.Forms.DialogResult]::Cancel }
    $script:DefDlg = $dlg; $script:DefGrid = $grid; $script:DefKind = $Kind
    $btnSave.Add_Click({ if (Save-DefDialog) { $script:DefDlg.DialogResult = [System.Windows.Forms.DialogResult]::OK } })
    $dlg.CancelButton = $btnClose
    $dlg.Controls.AddRange(@($lbl, $grid, $btnSave, $btnClose))
    [void]$dlg.ShowDialog($form)
    $dlg.Dispose()
    Update-DefLists
}

function Save-DefDialog {
    $g = $script:DefGrid; $isTs = ($script:DefKind -eq 'Ts')
    $null = $g.EndEdit()
    $list = @(); $errs = @(); $msgs = @()
    $seen = New-NameMap
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        foreach ($row in $g.Rows) {
            if ($row.IsNewRow) { continue }
            $a = ([string]$row.Cells[0].Value).Trim(); $b = ([string]$row.Cells[1].Value).Trim()
            $c3 = if ($isTs) { ([string]$row.Cells[2].Value).Trim() } else { '' }
            if (-not $a -and -not $b -and -not $c3) { continue }
            $n = $row.Index + 1
            if ($isTs) {
                if (-not $a) { $errs += "Satir ${n}: Gorunen Ad bos."; continue }
                if (-not $b -and -not $c3) { $errs += "Satir ${n}: Package ID veya Deployment ID girin."; continue }
                if ($seen.ContainsKey($a)) { $errs += "Satir ${n}: ayni ad iki kez."; continue }
                $v = Test-TsDefinition -PackageId $b.ToUpper() -DeploymentId $c3.ToUpper()
                if (-not $v.Ok) { $errs += "Satir ${n} ($a): $($v.Error)"; $row.Cells[3].Value = "HATA: $($v.Error)"; continue }
                $row.Cells[1].Value = $v.PackageId; $row.Cells[3].Value = "$($v.PkgName) [$($v.PackageId)]"
                $seen[$a] = 1
                $list += [pscustomobject]@{ Name = $a; PackageId = $v.PackageId; DeploymentId = $c3.ToUpper(); PkgName = $v.PkgName }
                $msgs += "${a}: $($v.Text -replace "`r`n", ' | ')"
            }
            else {
                if (-not $b) { $errs += "Satir ${n}: Collection ID veya ad girin."; continue }
                $v = Test-CollectionDefinition -Text $b
                if (-not $v.Ok) { $errs += "Satir ${n} ($b): $($v.Error)"; $row.Cells[2].Value = "HATA: $($v.Error)"; continue }
                if ($seen.ContainsKey($v.Id)) { $errs += "Satir ${n}: ayni koleksiyon iki kez ($($v.Id))."; continue }
                $seen[$v.Id] = 1
                $short = $(if ($a) { $a } else { $v.Name }) -replace ';', ','
                $row.Cells[0].Value = $short; $row.Cells[1].Value = $v.Id; $row.Cells[2].Value = "$($v.Name) [$($v.Id)] - $($v.Members) uye"
                $list += [pscustomobject]@{ Id = $v.Id; Name = $v.Name; Short = $short }
                $msgs += "${short}: $($v.Name) [$($v.Id)], $($v.Members) uye"
            }
        }
    }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    if ($errs.Count -gt 0) { [System.Windows.Forms.MessageBox]::Show(($errs -join "`n"), 'Dogrulama hatasi - kaydedilmedi', 'OK', 'Warning') | Out-Null; return $false }
    if ($isTs) { $script:TsDefs = @($list) } else { $script:Collections = @($list) }
    Save-Settings
    Write-ActionLog -Hostname '-' -Action $(if ($isTs) { 'TS tanimlari kaydedildi' } else { 'Koleksiyon listesi kaydedildi' }) -State 'ok' -Text ($msgs -join ' || ')
    [System.Windows.Forms.MessageBox]::Show("Dogrulandi ve kaydedildi:`n`n" + $(if ($msgs.Count) { $msgs -join "`n" } else { '(liste bos)' }), 'Bilgi') | Out-Null
    return $true
}

function Update-DefLists {
    # TS / koleksiyon ComboBox'lari tanimlardan yeniden doldurulur (secim korunur).
    $sel = [string]$cboTs.SelectedItem; if (-not $sel) { $sel = $script:ActiveTsName }
    $cboTs.Items.Clear(); foreach ($t in $script:TsDefs) { [void]$cboTs.Items.Add($t.Name) }
    if ($sel -and $cboTs.Items.Contains($sel)) { $cboTs.SelectedItem = $sel } elseif ($cboTs.Items.Count -gt 0) { $cboTs.SelectedIndex = 0 }
    $selC = [string]$cboColl.SelectedItem
    $cboColl.Items.Clear(); foreach ($c in $script:Collections) { [void]$cboColl.Items.Add($c.Short) }
    if ($selC -and $cboColl.Items.Contains($selC)) { $cboColl.SelectedItem = $selC } elseif ($cboColl.Items.Count -gt 0) { $cboColl.SelectedIndex = 0 }
    Update-TsInfo
    Update-ButtonStates
}

function Get-ActiveTs {
    $n = [string]$cboTs.SelectedItem
    return @($script:TsDefs | Where-Object { $_.Name -eq $n })[0]
}

function Update-TsInfo {
    $t = Get-ActiveTs
    $lblTsInfo.Text = if ($t) { "$($t.PkgName)  |  Package: $($t.PackageId)  |  Deployment: $(if ($t.DeploymentId) { $t.DeploymentId } else { '(paketin tum deployment''lari)' })" } else { "TS tanimli degil - 'TS Tanimlari...' ile ekleyin" }
}
#endregion

#region Islemler (Kontrol / TS / Koleksiyon)

function Test-WriteAllowed {
    # Yazma islemi on kontrolu: ust sinir + ONAY penceresi. DRY-RUN bilgisi onayda gosterilir.
    param([string]$Op, [string]$Target, [string[]]$Hosts, [string]$Extra = '')
    $max = [int]$nudMax.Value
    if (@($Hosts).Count -gt $max) {
        [System.Windows.Forms.MessageBox]::Show("Tek seferde en fazla $max cihaz (ayar: 'Tek seferde en fazla'). Secilen: $(@($Hosts).Count).`nIslem yapilmadi.", 'Ust sinir asildi', 'OK', 'Warning') | Out-Null
        return $false
    }
    $dry = if ($chkDryRun.Checked) { "`n`nDRY-RUN ACIK: hicbir sey yazilmayacak, sadece ne yapilacagi gosterilecek." } else { "`n`nDIKKAT: bu islem GERCEKTEN uygulanacak." }
    $c = [System.Windows.Forms.MessageBox]::Show("Islem: $Op`nHedef: $Target`nCihaz sayisi: $(@($Hosts).Count)`n`n$(Format-HostPreview -Hosts $Hosts)$Extra$dry`n`nDevam edilsin mi?",
        'Onay', [System.Windows.Forms.MessageBoxButtons]::YesNo, $(if ($chkDryRun.Checked) { [System.Windows.Forms.MessageBoxIcon]::Question } else { [System.Windows.Forms.MessageBoxIcon]::Warning }), [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    return ($c -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Start-Check {
    # Kontrol: SCCM'e bagliysa once SMS Provider toplu sorgusu (kayit / Boundary / OS Build / koleksiyon uyelikleri),
    # sonra cihaz basina DNS / Online / CcmExec (+ aktif TS seciliyse TS durumu ISTEMCIDEN).
    param([string[]]$Hosts, [switch]$Reset)
    if ($Reset) { $script:ResultsTable.Rows.Clear(); $script:Details.Clear() }
    $t = Get-ActiveTs
    $pkg = if ($t) { [string]$t.PackageId } else { '' }
    $cols = @('DNS_Durumu', 'Online_mi', 'AD_Durumu', 'AD_LastLogon', 'CcmExec_Durumu', 'OS_Build_Anlik')
    if ($pkg) { $cols += 'TS_Durum', 'TS_Adim', 'TS_HataKodu', 'TS_SonGuncelleme' }
    $sccm = ($script:SmsServer -and $script:SiteCode)
    if ($sccm) { $cols += 'SCCM_Kayit', 'Uye_Koleksiyonlar', 'Boundary_Groups', 'OS_Build' }
    Set-GridPending -Hosts $Hosts -Columns $cols -ClearError
    if ($Reset) { try { $dgv.ClearSelection(); $dgv.CurrentCell = $null } catch { } }
    $script:CheckTsName = if ($t) { $t.Name } else { '' }
    $items = @(foreach ($h in $Hosts) { [pscustomobject]@{ Key = $h; Params = @{ ComputerName = $h; PackageId = $pkg; SmsProvider = $script:SmsServer; SiteCode = $script:SiteCode; DeploymentId = $(if ($t) { [string]$t.DeploymentId } else { '' }) } } })
    $run = @{ Title = 'Kontrol'; Worker = $script:CheckWorker; OnResultFn = 'Update-FromCheck'; Panel = 1 }
    if ($sccm) { Start-Snapshot -Hosts $Hosts -Title 'Kontrol' -Next ([pscustomobject]@{ Items = $items; RunArgs = $run }) }
    else {
        foreach ($h in $Hosts) { Set-GridValues -Hostname $h -Values @{ SCCM_Kayit = 'Sorgulanmadi (SCCM baglantisi yok)' } }
        Start-ParallelRun @run -Items $items
    }
}

function Start-TsStatus {
    # TS durumu ISTEMCIDEN (SMS Provider'a sorulmaz): DNS / Online / CcmExec + kayit defteri / politika / calisan TS.
    param([string[]]$Hosts)
    $t = Get-ActiveTs
    if (-not $t) { [System.Windows.Forms.MessageBox]::Show("Once aktif TS secin ('TS Tanimlari...').", 'Uyari') | Out-Null; return }
    Set-GridPending -Hosts $Hosts -Columns @('DNS_Durumu', 'Online_mi', 'CcmExec_Durumu', 'OS_Build_Anlik', 'TS_Durum', 'TS_Adim', 'TS_HataKodu', 'TS_SonGuncelleme') -ClearError
    $script:CheckTsName = $t.Name
    $items = @(foreach ($h in $Hosts) { [pscustomobject]@{ Key = $h; Params = @{ ComputerName = $h; PackageId = [string]$t.PackageId; SmsProvider = $script:SmsServer; SiteCode = $script:SiteCode; DeploymentId = [string]$t.DeploymentId } } })
    Start-ParallelRun -Title "TS durumu ($($t.Name))" -Worker $script:CheckWorker -Items $items -OnResultFn 'Update-FromCheck' -Panel 1
}

function Start-TsTrigger {
    param([string[]]$Hosts)
    $t = Get-ActiveTs
    if (-not $t) { [System.Windows.Forms.MessageBox]::Show("Once aktif TS secin ('TS Tanimlari...').", 'Uyari') | Out-Null; return }
    if (-not (Test-WriteAllowed -Op 'TS Tetikle' -Target "$($t.Name) - $($t.PkgName) [$($t.PackageId)]" -Hosts $Hosts -Extra "`n`nOnce Online + kimlik + CcmExec kontrol edilir; uymayan cihaz atlanir. TS istemcide yoksa Machine Policy istenir ve en fazla $($script:PolicyWaitSec) sn beklenir.")) { return }
    $dry = [bool]$chkDryRun.Checked
    $script:TriggerTitle = $t.Name
    Write-ActionLog -Hostname '-' -Action "TS tetikle [$($t.Name)] BASLADI" -State 'ok' -Text "$(@($Hosts).Count) cihaz, PackageID=$($t.PackageId), DRY-RUN=$dry"
    Set-GridPending -Hosts $Hosts -Columns @('Son_Islem', 'Islem_Sonucu')
    $items = @(foreach ($h in $Hosts) { [pscustomobject]@{ Key = $h; Params = @{ ComputerName = $h; PackageId = $t.PackageId; DryRun = $dry; PolicyWaitSec = $script:PolicyWaitSec } } })
    Start-ParallelRun -Title "TS tetikle ($($t.Name))$(if ($dry) { ' [DRY-RUN]' })" -Worker $script:TriggerWorker -Items $items -OnResultFn 'Update-FromTrigger' -Panel 1
}

function Start-CollectionOp {
    param([ValidateSet('Add', 'Remove')][string]$Mode, [string[]]$Hosts, [string]$CollShort)
    if (-not (Test-Connected)) { return }
    if (-not $CollShort) { $CollShort = [string]$cboColl.SelectedItem }
    $c = @($script:Collections | Where-Object { $_.Short -eq $CollShort })[0]
    if (-not $c) { [System.Windows.Forms.MessageBox]::Show("Once koleksiyon secin ('Koleksiyonlar...').", 'Uyari') | Out-Null; return }
    $verb = if ($Mode -eq 'Add') { 'Koleksiyona Ekle' } else { 'Koleksiyondan Cikar' }
    $extra = if ($chkRefresh.Checked) { "`n`nKoleksiyon yenileme (RequestRefresh) istenecek." } else { '' }
    if (-not (Test-WriteAllowed -Op $verb -Target "$($c.Short) - $($c.Name) [$($c.Id)]" -Hosts $Hosts -Extra $extra)) { return }
    $key = "coll-$([guid]::NewGuid().ToString('N').Substring(0, 8))"
    $script:CollOps[$key] = [pscustomobject]@{ Mode = $Mode; Id = $c.Id; Short = $c.Short; Hosts = @($Hosts) }
    Write-ActionLog -Hostname '-' -Action "$verb [$($c.Id) $($c.Short)] BASLADI" -State 'ok' -Text "$(@($Hosts).Count) cihaz, DRY-RUN=$([bool]$chkDryRun.Checked), RequestRefresh=$([bool]$chkRefresh.Checked)"
    Set-GridPending -Hosts $Hosts -Columns @('Son_Islem', 'Islem_Sonucu')
    $p = @{ SmsProvider = $script:SmsServer; SiteCode = $script:SiteCode; CollectionId = $c.Id; Mode = $Mode; Names = [string[]]$Hosts
            DryRun = [bool]$chkDryRun.Checked; Refresh = [bool]$chkRefresh.Checked; AllCollectionIds = [string[]]@($script:Collections | ForEach-Object { $_.Id }) }
    Start-ParallelRun -Title "$verb ($($c.Short))" -Worker $script:CollWorker -Items @([pscustomobject]@{ Key = $key; Params = $p }) -OnResultFn 'Update-FromCollection' -Panel 1 -Unit 'islem'
}

function Start-ClientAction {
    # client_sccm_tool Panel 1 ile ayni: SMS_Client.TriggerSchedule (WMI / DCOM) - ya da onarim / CcmExec yeniden baslatma.
    param([string[]]$Hosts, [string]$ActionName)
    if (-not $ActionName) { $ActionName = [string]$cboAction.SelectedItem }
    $ids = $script:ClientActions[$ActionName]
    if (-not $ids) { return }
    if (@($Hosts).Count -gt [int]$nudMax.Value) {
        [System.Windows.Forms.MessageBox]::Show("Tek seferde en fazla $([int]$nudMax.Value) cihaz. Secilen: $(@($Hosts).Count). Islem yapilmadi.", 'Ust sinir asildi', 'OK', 'Warning') | Out-Null
        return
    }
    if ($ids[0] -like '@*') {   # onarim / yeniden baslatma: onay ister
        $c = [System.Windows.Forms.MessageBox]::Show("Islem: $ActionName`nCihaz sayisi: $(@($Hosts).Count)`n`n$(Format-HostPreview -Hosts $Hosts)`n`nDevam edilsin mi?", 'Onay',
            [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning, [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($c -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }
    Set-GridPending -Hosts $Hosts -Columns @('Son_Islem', 'Islem_Sonucu')
    $items = @(foreach ($h in $Hosts) { [pscustomobject]@{ Key = $h; Params = @{ ComputerName = $h; ActionName = $ActionName; ScheduleIds = [string[]]$ids } } })
    Start-ParallelRun -Title "Istemci eylemi ($ActionName)" -Worker $script:ActionWorker -Items $items -OnResultFn 'Update-FromAction' -Panel 1
}

function Start-TsStepHistory {
    param([string[]]$Hosts)
    if (-not (Test-Connected)) { return }
    $t = Get-ActiveTs
    if (-not $t) { [System.Windows.Forms.MessageBox]::Show("Once aktif TS secin ('TS Tanimlari...').", 'Uyari') | Out-Null; return }
    if (@($Hosts).Count -ne 1) { [System.Windows.Forms.MessageBox]::Show('Tek bir satir secin.', 'Bilgi') | Out-Null; return }
    $script:DetailHost = $Hosts[0]; $script:DetailTitle = "TS adim gecmisi - $($t.Name) (SCCM)"
    Start-ParallelRun -Title "TS adim gecmisi ($($Hosts[0]))" -Worker $script:StepHistWorker -Items @([pscustomobject]@{ Key = $Hosts[0]; Params = @{
        ComputerName = $Hosts[0]; SmsProvider = $script:SmsServer; SiteCode = $script:SiteCode; PackageId = [string]$t.PackageId; DeploymentId = [string]$t.DeploymentId } }) -OnResultFn 'Update-FromDetail' -Panel 1 -Unit 'islem'
}

function Start-TsClientDetail {
    param([string[]]$Hosts)
    $t = Get-ActiveTs
    if (-not $t) { [System.Windows.Forms.MessageBox]::Show("Once aktif TS secin ('TS Tanimlari...').", 'Uyari') | Out-Null; return }
    if (@($Hosts).Count -ne 1) { [System.Windows.Forms.MessageBox]::Show('Tek bir satir secin.', 'Bilgi') | Out-Null; return }
    $script:DetailHost = $Hosts[0]; $script:DetailTitle = 'TS istemci detayi'
    Start-ParallelRun -Title "TS istemci detayi ($($Hosts[0]))" -Worker $script:DetailWorker -Items @([pscustomobject]@{ Key = $Hosts[0]; Params = @{ ComputerName = $Hosts[0]; PackageId = $t.PackageId } }) -OnResultFn 'Update-FromDetail' -Panel 1 -Unit 'islem'
}

function Get-OpHosts {
    # Hedef cihazlar: gridde secili satirlar, yoksa soldaki liste.
    $t = Get-TargetHosts -Grid $dgv
    if (-not $t) { return @() }
    return @($t.Hosts)
}
#endregion

#region Tekrar dene / ozet / otomatik yenileme
function Start-Retry {
    param([int]$Panel = 1)
    if (Test-PanelBusy -Panel 1) { return }
    $op = $script:LastOp[1]
    if (-not $op) { [System.Windows.Forms.MessageBox]::Show('Tekrar denenecek bir islem yok.', 'Bilgi') | Out-Null; return }
    $bad = New-NameMap
    foreach ($row in $script:ResultsTable.Rows) { if ([string]$row['State'] -eq 'fail') { $bad[[string]$row['Hostname']] = 1 } }
    $items = @($op.Items | Where-Object { $bad.ContainsKey([string]$_.Key) })
    if ($items.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Son islemde ($($op.RunArgs.Title)) hata alan cihaz yok.", 'Bilgi') | Out-Null; return }
    $hosts = @($items | ForEach-Object { [string]$_.Key })
    if ($op.RunArgs.OnResultFn -eq 'Update-FromCheck') { Start-Check -Hosts $hosts; return }
    $c = [System.Windows.Forms.MessageBox]::Show("Islem: $($op.RunArgs.Title)`n`nHata alan $($hosts.Count) cihaz:`n$(Format-HostPreview -Hosts $hosts)`n`nDIKKAT: islem bu cihazlara TEKRAR gonderilecek. Devam?",
        'Basarisizlari Tekrar Dene', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($c -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    Write-ActionLog -Hostname '-' -Action "$($op.RunArgs.Title) - TEKRAR" -State 'ok' -Text "$($hosts.Count) cihaz"
    Set-GridPending -Hosts $hosts -Columns @('Son_Islem', 'Islem_Sonucu')
    $a = $op.RunArgs.Clone(); $a.Title = "$($a.Title) - tekrar"
    Start-ParallelRun @a -Items @($items | ForEach-Object { [pscustomobject]@{ Key = $_.Key; Params = $_.Params.Clone() } })
}

function Update-SummaryNow {
    param([int]$Panel = 1)
    $n = @{ ok = 0; warn = 0; fail = 0; pending = 0 }
    foreach ($row in $script:ResultsTable.Rows) { $v = [string]$row['State']; if ($n.ContainsKey($v)) { $n[$v]++ } }
    $b = $script:SumButtons[1]
    $b.ok.Text = "Tamam: $($n.ok)"; $b.warn.Text = "Uyari: $($n.warn)"; $b.fail.Text = "Hata: $($n.fail)"; $b.pending.Text = "Bekleyen: $($n.pending)"
    $b.all.Text = "Tumu: $($script:ResultsTable.Rows.Count)"
    $script:RetryButtons[1].Enabled = ($n.fail -gt 0 -and $script:LastOp[1] -and -not (Test-PanelBusy -Panel 1))
    $script:SumDirty[1] = $false
}

function Update-Filter {
    param([int]$Panel = 1)
    Set-GridFilter -Source $script:BindingSrc -ColMap $script:ColMap1 -ColumnHeader ([string]$cboF1.SelectedItem) -Text $txtF1.Text -StateColumn 'State' -State $script:StateFilter[1]
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

function Reset-AutoRefresh {
    param([int]$Panel)
    $a = $script:Auto[$Panel]
    if ($a.Check.Checked) { $a.Next = (Get-Date).AddMinutes([int]$a.Minutes.Value); $a.Label.Text = "Sonraki: $($a.Next.ToString('HH:mm'))" }
    else { $a.Next = [datetime]::MaxValue; $a.Label.Text = '' }
}

function Invoke-AutoRefresh {
    # Tablodaki tum cihazlar icin Kontrol (aktif TS seciliyse TS durumu da) - satirlar yerinde guncellenir. Yazma YAPMAZ.
    $a = $script:Auto[1]
    if (-not $a -or -not $a.Check.Checked -or (Get-Date) -lt $a.Next) { return }
    if (Test-PanelBusy -Panel 1) { $a.Label.Text = 'Sonraki: mesgul, bekleniyor'; return }
    $hosts = @(foreach ($row in $script:ResultsTable.Rows) { $h = [string]$row['Hostname']; if ($h -match $script:HostRegex) { $h } })
    Reset-AutoRefresh -Panel 1
    if ($hosts.Count -eq 0) { $a.Label.Text += ' (tablo bos)'; return }
    Write-ActionLog -Hostname '-' -Action 'Otomatik yenileme' -State 'ok' -Text "$($hosts.Count) cihaz"
    Start-Check -Hosts $hosts
}
#endregion

#region Diyaloglar / disa aktarma / filtre / log
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
    $files = @(Get-ChildItem -LiteralPath $dir -Filter 'TsTool_*.log' -ErrorAction SilentlyContinue | Sort-Object Name -Descending | ForEach-Object { $_.Name })
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
#endregion

#region Ana form
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

$form = New-Ctl Form 'Task Sequence Kontrol Araci' 0 0 1360 900 @{ StartPosition = 'CenterScreen'; MinimumSize = (New-Object System.Drawing.Size(1150, 700)) }

$lblComputers = New-Ctl Label 'Bilgisayar Isimleri (her satira bir tane):' 10 10
$txtComputers = New-Ctl TextBox '' 10 30 200 800 @{ Multiline = $true; ScrollBars = 'Vertical'; Anchor = 'Top,Bottom,Left' }

# 1) Kontrol
$gbCheck     = New-Ctl GroupBox '1) Baglanti / Kontrol' 220 8 520 150
$txtSms      = New-Ctl TextBox '' 100 20 180 23
$btnConnect  = New-Ctl Button 'Baglan' 288 18 80 27
$lblSite     = New-Ctl Label 'Site: -' 378 24
$btnCheck    = New-Ctl Button 'Kontrol Et' 10 54 110 30
$btnAdCache  = New-Ctl Button 'AD Verisi...' 130 54 110 30
$lblAdCache  = New-Ctl Label 'Onbellek: -' 248 62 262 18 $gray
$lblPingOnly = New-Ctl Label 'Sadece ping:' 10 102
$txtPingOnly = New-Ctl TextBox '' 85 98 420 23
$txtPingOnly.Add_TextChanged({ $env:CLIENTTOOL_PINGONLY = $txtPingOnly.Text })   # worker'lar da okur
$gbCheck.Controls.AddRange(@((New-Ctl Label 'SMS Provider:' 10 24), $txtSms, $btnConnect, $lblSite, $btnCheck, $btnAdCache, $lblAdCache, $lblPingOnly, $txtPingOnly))

# 2) Task Sequence
$gbTs         = New-Ctl GroupBox '2) Task Sequence' 750 8 580 150
$cboTs        = New-Ctl ComboBox '' 70 20 360 23 @{ DropDownStyle = 'DropDownList' }
$btnTsDefs    = New-Ctl Button 'TS Tanimlari...' 440 18 130 27
$btnTsStatus  = New-Ctl Button 'TS Durumu Sorgula' 10 54 160 30
$btnTsTrigger = New-Ctl Button 'TS Tetikle' 180 54 130 30
$lblTsInfo    = New-Ctl Label '' 10 94 560 48 @{ ForeColor = [System.Drawing.Color]::DimGray }
$gbTs.Controls.AddRange(@((New-Ctl Label 'Aktif TS:' 10 24), $cboTs, $btnTsDefs, $btnTsStatus, $btnTsTrigger, $lblTsInfo))

# 3) Koleksiyon
$gbColl        = New-Ctl GroupBox '3) Koleksiyon (Direct uyelik)' 220 164 520 90
$cboColl       = New-Ctl ComboBox '' 85 22 250 23 @{ DropDownStyle = 'DropDownList' }
$btnCollAdd    = New-Ctl Button 'Ekle' 343 20 75 27
$btnCollRemove = New-Ctl Button 'Cikar' 424 20 75 27
$btnColls      = New-Ctl Button 'Koleksiyonlar...' 10 54 130 27
$chkRefresh    = New-Ctl CheckBox 'Koleksiyonu yenile (RequestRefresh)' 150 58
$gbColl.Controls.AddRange(@((New-Ctl Label 'Koleksiyon:' 10 26), $cboColl, $btnCollAdd, $btnCollRemove, $btnColls, $chkRefresh))

# 4) Guvenlik
$gbSafe    = New-Ctl GroupBox '4) Guvenlik / Kayit' 750 164 580 90
$chkDryRun = New-Ctl CheckBox 'DRY-RUN (hicbir sey yazma, sadece ne yapilacagini goster)' 10 24
$nudMax    = New-Ctl NumericUpDown '' 150 54 70 23 @{ Minimum = 1; Maximum = 5000 }
$nudMax.Value = 200
$btnLog    = New-Ctl Button 'Islem Kayitlari...' 440 20 130 27
$gbSafe.Controls.AddRange(@($chkDryRun, (New-Ctl Label 'Tek seferde en fazla:' 10 58), $nudMax, (New-Ctl Label 'cihaz (yazma islemleri)' 226 58 0 0 $gray), $btnLog))

# 5) Istemci eylemi (client_sccm_tool Panel 1 ile ayni)
$gbActions = New-Ctl GroupBox '5) SCCM Istemci Eylemi (cycle)' 220 258 1110 56
$cboAction = New-Ctl ComboBox '' 10 22 400 23 @{ DropDownStyle = 'DropDownList' }
foreach ($k in $script:ClientActions.Keys) { [void]$cboAction.Items.Add($k) }
$cboAction.SelectedIndex = 0
$btnAction = New-Ctl Button 'Eylemi Gonder' 420 19 150 28
$gbActions.Controls.AddRange(@($cboAction, $btnAction, (New-Ctl Label 'Hedef: gridde secili satirlar; yoksa soldaki liste. Sag tik menusunde de var.' 585 26 0 0 $gray)))

$btnCancel1 = New-Ctl Button 'Iptal Et' 220 320 90 24 @{ Enabled = $false }
$lblStatus  = New-Ctl Label '' 318 324 0 0 @{ ForeColor = [System.Drawing.Color]::DarkBlue }

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
    param($Grid, [int]$RowIndex)
    if ($RowIndex -lt 0 -or $RowIndex -ge $Grid.Rows.Count) { return }
    $rv = $Grid.Rows[$RowIndex].DataBoundItem
    if (-not $rv) { return }
    $hn = [string]$rv['Hostname']
    $txt = if ($script:Details.ContainsKey($hn)) { $script:Details[$hn] } else { '(Detay yok)' }
    Show-TextDialog -Title "Detay: $hn" -Text $txt
}

# Sag tik menusu
$script:GridMenu = New-Object System.Windows.Forms.ContextMenuStrip
$script:GridOps = @()
function Add-MenuItem {
    param($Parent, [string]$Text, [scriptblock]$Click, [switch]$Bold, [switch]$Op)
    $it = New-Object System.Windows.Forms.ToolStripMenuItem ($Text)
    if ($Bold) { $it.Font = New-Object System.Drawing.Font($it.Font, [System.Drawing.FontStyle]::Bold) }
    if ($Click) { $it.Add_Click($Click) }
    if ($Parent -is [System.Windows.Forms.ToolStripMenuItem]) { [void]$Parent.DropDownItems.Add($it) } else { [void]$Parent.Items.Add($it) }
    if ($Op) { $script:GridOps += $it }
    return $it
}
$null = Add-MenuItem $script:GridMenu 'Kontrol Et (secili)' { $h = @(Get-SelectedHosts -Grid $dgv); if ($h.Count) { Start-Check -Hosts $h } } -Bold -Op
$null = Add-MenuItem $script:GridMenu 'TS Durumu Sorgula (secili)' { $h = @(Get-SelectedHosts -Grid $dgv); if ($h.Count) { Start-TsStatus -Hosts $h } } -Bold -Op
$null = Add-MenuItem $script:GridMenu 'TS Tetikle (secili)' { $h = @(Get-SelectedHosts -Grid $dgv); if ($h.Count) { Start-TsTrigger -Hosts $h } } -Bold -Op
$script:MenuCollAdd = Add-MenuItem $script:GridMenu 'Koleksiyona Ekle' $null -Bold -Op
$script:MenuCollRem = Add-MenuItem $script:GridMenu 'Koleksiyondan Cikar' $null -Bold -Op
$script:MenuAction = Add-MenuItem $script:GridMenu 'Istemci Eylemi Gonder (secili)' $null -Bold -Op
foreach ($k in $script:ClientActions.Keys) {
    $x = Add-MenuItem $script:MenuAction $k { $h = @(Get-SelectedHosts -Grid $dgv); if ($h.Count) { Start-ClientAction -Hosts $h -ActionName ([string]$this.Text) } }
}
$null = Add-MenuItem $script:GridMenu 'TS adim gecmisi (SCCM, secili satir, sadece okur)' { Start-TsStepHistory -Hosts @(Get-SelectedHosts -Grid $dgv) } -Op
$null = Add-MenuItem $script:GridMenu 'TS istemci detayi (secili satir, sadece okur)' { Start-TsClientDetail -Hosts @(Get-SelectedHosts -Grid $dgv) } -Op
[void]$script:GridMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$null = Add-MenuItem $script:GridMenu 'Kopyala (Ctrl+C)' { Copy-GridData -Grid $script:GridMenu.SourceControl -Mode 'Cells' }
$null = Add-MenuItem $script:GridMenu 'Secili satirlari basliklariyla kopyala (Ctrl+Shift+C)' { Copy-GridData -Grid $script:GridMenu.SourceControl -Mode 'Rows' }
$null = Add-MenuItem $script:GridMenu 'Bu sutunu kopyala (gorunen tum satirlar)' { Copy-GridData -Grid $script:GridMenu.SourceControl -Mode 'Column' }
$null = Add-MenuItem $script:GridMenu 'Hostname listesini kopyala' { Copy-GridData -Grid $script:GridMenu.SourceControl -Mode 'Hosts' }
[void]$script:GridMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
$script:MenuDetail = Add-MenuItem $script:GridMenu 'Satir detayi (cift tik)' { if ($dgv.CurrentCell) { Show-RowDetail -Grid $dgv -RowIndex $dgv.CurrentCell.RowIndex } }
$script:GridMenu.Add_Opening({
    $busy = Test-PanelBusy -Panel 1
    $main = ($script:GridMenu.SourceControl -eq $dgv)   # log penceresi tablosu da ayni menuyu kullanir
    foreach ($it in $script:GridOps) { $it.Enabled = -not $busy; $it.Visible = $main }
    $script:MenuDetail.Visible = $main
    foreach ($m in @(@($script:MenuCollAdd, 'Add'), @($script:MenuCollRem, 'Remove'))) {
        $m[0].DropDownItems.Clear()
        if (@($script:Collections).Count -eq 0) { $x = $m[0].DropDownItems.Add('(koleksiyon tanimli degil)'); $x.Enabled = $false; continue }
        foreach ($c in $script:Collections) {
            $x = New-Object System.Windows.Forms.ToolStripMenuItem ("$($c.Short)  [$($c.Id)]")
            $x.Tag = "$($m[1])|$($c.Short)"
            $x.Add_Click({ $p = ([string]$this.Tag).Split('|', 2); $h = @(Get-SelectedHosts -Grid $dgv); if ($h.Count) { Start-CollectionOp -Mode $p[0] -Hosts $h -CollShort $p[1] } })
            [void]$m[0].DropDownItems.Add($x)
        }
    }
})

# Sonuc tablosu
$dgv = New-ResultGrid
$dgv.AutoSizeColumnsMode = 'Fill'
$dgv.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($script:GridCols | ForEach-Object { New-GridColumn $_[0] $_[1] $_[2] }))
$dgv.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@((New-GridColumn 'State' 'State' 5 $false), (New-GridColumn 'ResourceId' 'ResourceId' 5 $false)))
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
$dgv.add_CellDoubleClick({ param($s, $e) if ($e.RowIndex -ge 0) { Show-RowDetail -Grid $s -RowIndex $e.RowIndex } })
$dgv.add_CellMouseDown({
    param($s, $e)
    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Right -or $e.RowIndex -lt 0 -or $e.ColumnIndex -lt 0) { return }
    $cell = $s.Rows[$e.RowIndex].Cells[$e.ColumnIndex]
    if (-not $cell.Selected) { $s.ClearSelection(); $cell.Selected = $true }
    $s.CurrentCell = $cell
})
$dgv.Add_MouseDown({ param($s, $e) if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left -and $s.HitTest($e.X, $e.Y).Type -eq [System.Windows.Forms.DataGridViewHitTestType]::None) { Clear-GridFocus -Grid $s } })
$dgv.Add_KeyDown({
    param($s, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { Clear-GridFocus -Grid $s; $e.Handled = $true }
    elseif ($e.Control -and $e.Shift -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) { Copy-GridData -Grid $s -Mode 'Rows'; $e.Handled = $true }
})

# Arac cubugu: filtre + Excel + CSV + ozet
$pnlTop1     = New-Ctl Panel '' 0 0 0 0 @{ Dock = 'Top'; Height = 66 }
$cboF1       = New-Ctl ComboBox '' 50 7 150 23 @{ DropDownStyle = 'DropDownList' }
$txtF1       = New-Ctl TextBox '' 206 7 200 23
$btnClearF1  = New-Ctl Button 'Temizle' 412 5 65 27
$btnExcel1   = New-Ctl Button "Excel'e Aktar" 490 5 110 27
$btnCsv1     = New-Ctl Button 'CSV' 606 5 70 27
$btnUnfocus1 = New-Ctl Button 'Focus Kaybet' 688 5 124 27
$lblCount1   = New-Ctl Label '' 822 11 0 0 $gray
$pnlTop1.Controls.AddRange(@((New-Ctl Label 'Filtre:' 6 11), $cboF1, $txtF1, $btnClearF1, $btnExcel1, $btnCsv1, $btnUnfocus1, $lblCount1))
New-SummaryBar -Panel 1 -Parent $pnlTop1

$pnlGrid = New-Ctl Panel '' 220 350 1110 483 @{ Anchor = 'Top,Bottom,Left,Right'; BorderStyle = 'FixedSingle' }
$pnlGrid.Controls.AddRange(@($dgv, $pnlTop1))   # Fill once, Top sonra

$script:ColMap1 = Get-ColumnMap -Grid $dgv
[void]$cboF1.Items.Add($script:AllColsLabel)
foreach ($k in $script:ColMap1.Keys) { [void]$cboF1.Items.Add($k) }
$cboF1.SelectedIndex = 0

$form.Controls.AddRange(@($lblComputers, $txtComputers, $gbCheck, $gbTs, $gbColl, $gbSafe, $gbActions, $btnCancel1, $lblStatus, $pnlGrid))

$tip = New-Object System.Windows.Forms.ToolTip
$tip.AutoPopDelay = 20000
$tip.SetToolTip($btnConnect, "SMS Provider'a baglanir (site kodu otomatik bulunur). Ad kaydedilir; sonraki acilista arka planda otomatik baglanir.")
$tip.SetToolTip($btnCheck, "Hedef: soldaki liste. DNS / Online (ping + cihaz adi dogrulama) / AD / SCCM kaydi / Boundary Group(s) / OS Build / CcmExec / OS Build (anlik, istemci kayit defteri) / koleksiyon uyelikleri; aktif TS seciliyse TS durumu da (istemciden). Yazma YAPMAZ.")
$tip.SetToolTip($btnAdCache, "Domain / OU eslemesi. OU'lardaki bilgisayar hesaplari AD'den cekilip BELLEKTE tutulur; acilista otomatik yuklenir.")
$tip.SetToolTip($txtPingOnly, "AD / WMI yetkisi olmayan cihazlar (desen: * ve ?, virgul ile). Eslesen cihaz ping varsa cevrimici sayilir; kimlik / servis kontrolu yapilmaz, TS tetiklenmez.")
$tip.SetToolTip($btnTsDefs, 'TS tanimlarini ekle / duzenle / sil. Package ID ve/veya Deployment ID; kayitta SCCM''den dogrulanir.')
$tip.SetToolTip($btnTsStatus, "Secili TS'in durumunu ISTEMCIDEN okur (WMI): kayit defteri Execution History\System\<PackageID> + TS politikasi + calisan TS. Ham bilgi: satira cift tik.")
$tip.SetToolTip($btnAction, 'client_sccm_tool Panel 1 ile ayni: secili istemci eylemini (TriggerSchedule) gonderir. Onarim / CcmExec yeniden baslatma onay ister.')
$tip.SetToolTip($btnTsTrigger, "Secili TS'i istemcide baslatir (Software Center 'Yukle' ile ayni: ClientSDK ExecuteProgram). Online + kimlik + CcmExec Running sart; TS istemcide yoksa Machine Policy istenir. ONAY ister.")
$tip.SetToolTip($btnCollAdd, 'Secili koleksiyona DIRECT uyelik kurali ekler. Zaten uye olan atlanir. ONAY ister.')
$tip.SetToolTip($btnCollRemove, 'Secili koleksiyondan DIRECT uyelik kuralini siler. Uye olmayan atlanir; sorgu kuraliyla gelen cikarilamaz. ONAY ister.')
$tip.SetToolTip($btnColls, 'Koleksiyon listesini duzenle (ID veya ad; Device Collection olmali). Liste sag tik menusunde de gorunur.')
$tip.SetToolTip($chkRefresh, 'Isaretliyse ekle / cikar sonrasi koleksiyon icin RequestRefresh istenir (varsayilan KAPALI).')
$tip.SetToolTip($chkDryRun, 'Isaretliyse TS tetikleme ve koleksiyon ekle / cikar HICBIR SEY YAZMAZ; sadece ne yapilacagi gride yazilir.')
$tip.SetToolTip($nudMax, 'Yazma islemlerinde tek seferde en fazla cihaz. Asilirsa islem yapilmaz.')
$tip.SetToolTip($btnLog, 'TsTool_yyyyMM.log denetim kaydi (kim, ne zaman, hangi cihaz, hangi islem, sonuc).')
$tip.SetToolTip($btnCancel1, 'Bekleyen / calisan isleri durdurur. Gonderilmis TS tetikleme / koleksiyon degisikligi geri alinamaz.')
$tip.SetToolTip($btnUnfocus1, 'Tablodaki secimi kaldirir; hedef tekrar soldaki liste olur (Esc / bos alana tik da ayni).')
$tip.SetToolTip($script:RetryButtons[1], 'SON islemi, sonucu Hata (kirmizi) olan cihazlar icin yeniden calistirir (yazma isleminde tekrar onay ister).')
$tip.SetToolTip($script:Auto[1].Check, 'Secilen araliklarla tablodaki tum cihazlar icin Kontrol (aktif TS seciliyse TS durumu dahil) tekrarlanir. Yazma YAPMAZ.')
#endregion

#region UI durum / ayarlar
function Update-ButtonStates {
    $busy = Test-PanelBusy -Panel 1
    $hasTs = ($cboTs.SelectedIndex -ge 0)
    $hasC  = ($cboColl.SelectedIndex -ge 0)
    $btnTsStatus.Enabled  = ($hasTs -and -not $busy)
    $btnTsTrigger.Enabled = ($hasTs -and -not $busy)
    $btnCollAdd.Enabled    = ($hasC -and -not $busy)
    $btnCollRemove.Enabled = ($hasC -and -not $busy)
}

function Set-UiBusy {
    param([bool]$Busy, [int]$Panel = 1)
    foreach ($b in @($btnCheck, $btnTsStatus, $btnTsTrigger, $btnCollAdd, $btnCollRemove, $btnTsDefs, $btnColls, $btnAdCache, $btnAction)) { $b.Enabled = -not $Busy }
    $btnCancel1.Enabled = $Busy
    if ($Busy) { $script:RetryButtons[1].Enabled = $false }
    $script:SumDirty[1] = $true
    if (-not $Busy) { Update-ButtonStates }
}

function Import-Settings {
    try {
        $f = Get-SettingsFile
        if (-not (Test-Path $f)) { return }
        $s = Get-Content -Path $f -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($s.SmsServer) { $txtSms.Text = [string]$s.SmsServer }
        if ($null -ne $s.PingOnly) { $txtPingOnly.Text = [string]$s.PingOnly }
        if ($s.MaxBatch -and [int]$s.MaxBatch -ge 1 -and [int]$s.MaxBatch -le 5000) { $nudMax.Value = [int]$s.MaxBatch }
        if ($s.ActiveTs) { $script:ActiveTsName = [string]$s.ActiveTs }
        if ($null -ne $s.AdCacheSources) { $script:AdCacheSources = @($s.AdCacheSources | Where-Object { $_ -and $_.Domain -and $_.OU } | ForEach-Object { [pscustomobject]@{ Domain = [string]$_.Domain; OU = [string]$_.OU } }) }
        if ($null -ne $s.TsDefs) { $script:TsDefs = @($s.TsDefs | Where-Object { $_ -and $_.Name -and $_.PackageId } | ForEach-Object { [pscustomobject]@{ Name = [string]$_.Name; PackageId = [string]$_.PackageId; DeploymentId = [string]$_.DeploymentId; PkgName = [string]$_.PkgName } }) }
        if ($null -ne $s.Collections) { $script:Collections = @($s.Collections | Where-Object { $_ -and $_.Id } | ForEach-Object { [pscustomobject]@{ Id = [string]$_.Id; Name = [string]$_.Name; Short = [string]$_.Short } }) }
    }
    catch { }
}

function Save-Settings {
    try {
        $obj = [pscustomobject]@{
            SmsServer      = $txtSms.Text.Trim()
            PingOnly       = $txtPingOnly.Text.Trim()
            MaxBatch       = [int]$nudMax.Value
            ActiveTs       = [string]$cboTs.SelectedItem
            AdCacheSources = @($script:AdCacheSources | ForEach-Object { [pscustomobject]@{ Domain = [string]$_.Domain; OU = [string]$_.OU } })
            TsDefs         = @($script:TsDefs | ForEach-Object { [pscustomobject]@{ Name = $_.Name; PackageId = $_.PackageId; DeploymentId = $_.DeploymentId; PkgName = $_.PkgName } })
            Collections    = @($script:Collections | ForEach-Object { [pscustomobject]@{ Id = $_.Id; Name = $_.Name; Short = $_.Short } })
        }
        Set-Content -Path (Get-SettingsFile) -Value ($obj | ConvertTo-Json -Depth 4) -Encoding UTF8
    }
    catch { }
}
#endregion

#region Olaylar
$txtF1.Add_TextChanged({ Update-Filter })
$cboF1.Add_SelectedIndexChanged({ Update-Filter })
$btnClearF1.Add_Click({ $txtF1.Text = ''; $cboF1.SelectedIndex = 0 })
$script:BindingSrc.Add_ListChanged({ $lblCount1.Text = "Satir: $($script:BindingSrc.Count) / $($script:ResultsTable.Rows.Count)" })

$script:ResultsTable.Add_RowChanged({ $script:SumDirty[1] = $true })
$script:ResultsTable.Add_RowDeleted({ $script:SumDirty[1] = $true })
$script:ResultsTable.Add_TableCleared({ $script:SumDirty[1] = $true })
$script:SummaryTimer = New-Object System.Windows.Forms.Timer
$script:SummaryTimer.Interval = 500
$script:SummaryTimer.Add_Tick({
    if ($script:SumDirty[1]) { Update-SummaryNow }
    Update-AdLoad
    Update-Connect
    if ($script:ShowAfterRun -and -not (Test-PanelBusy -Panel 1)) {   # is bittikten sonra gosterilecek pencere
        $x = $script:ShowAfterRun; $script:ShowAfterRun = $null
        if ($x.Box) { [System.Windows.Forms.MessageBox]::Show($x.Text, $x.Title) | Out-Null } else { Show-TextDialog -Title $x.Title -Text $x.Text }
    }
})
$script:SummaryTimer.Start()
$script:AutoTimer = New-Object System.Windows.Forms.Timer
$script:AutoTimer.Interval = 15000
$script:AutoTimer.Add_Tick({ Invoke-AutoRefresh })
$script:AutoTimer.Start()

$btnConnect.Add_Click({ Start-Connect })
$btnCheck.Add_Click({ $h = @(Get-ComputerList); if ($h.Count) { Start-Check -Hosts $h -Reset } })
$btnAdCache.Add_Click({ Show-AdCacheDialog })
$btnTsDefs.Add_Click({ Show-DefDialog -Kind Ts })
$btnColls.Add_Click({ Show-DefDialog -Kind Coll })
$cboTs.Add_SelectedIndexChanged({ Update-TsInfo; Update-ButtonStates })
$cboColl.Add_SelectedIndexChanged({ Update-ButtonStates })
$btnTsStatus.Add_Click({ $h = @(Get-OpHosts); if ($h.Count) { Start-TsStatus -Hosts $h } })
$btnTsTrigger.Add_Click({ $h = @(Get-OpHosts); if ($h.Count) { Start-TsTrigger -Hosts $h } })
$btnCollAdd.Add_Click({ $h = @(Get-OpHosts); if ($h.Count) { Start-CollectionOp -Mode Add -Hosts $h } })
$btnCollRemove.Add_Click({ $h = @(Get-OpHosts); if ($h.Count) { Start-CollectionOp -Mode Remove -Hosts $h } })
$btnCancel1.Add_Click({ Stop-Run })
$btnAction.Add_Click({ $h = @(Get-OpHosts); if ($h.Count) { Start-ClientAction -Hosts $h } })
$btnLog.Add_Click({ Show-LogViewer })
$btnExcel1.Add_Click({
    if ($script:ResultsTable.Rows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Aktarilacak sonuc yok.', 'Bilgi') | Out-Null; return }
    Export-ToExcelAndOpen -Data (Get-GridExportData -Grid $dgv -Source $script:BindingSrc) -DefaultFileName (Get-ExportStamp)
})
$btnCsv1.Add_Click({
    if ($script:ResultsTable.Rows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Aktarilacak sonuc yok.', 'Bilgi') | Out-Null; return }
    Export-GridToCsv -Data (Get-GridExportData -Grid $dgv -Source $script:BindingSrc) -DefaultFileName (Get-ExportStamp)
})

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

$form.Add_FormClosing({
    Save-Settings
    if ($script:RunTimer.Enabled) { $script:RunTimer.Stop() }
    $script:SummaryTimer.Stop(); $script:AutoTimer.Stop()
    if ($script:AdLoadJob) { try { [void]$script:AdLoadJob.PS.BeginStop($null, $null) } catch { } }
    if ($script:ConnectJob) { try { [void]$script:ConnectJob.PS.BeginStop($null, $null) } catch { } }
    foreach ($pf in @($script:Prefetch)) { try { [void]$pf.Job.PS.BeginStop($null, $null) } catch { } }
    foreach ($st in @($script:ActiveRuns)) { foreach ($j in $st.Jobs) { if (-not $j.Done) { try { [void]$j.PS.BeginStop($null, $null) } catch { } } } }
    if ($script:SharedPool) { try { [void]$script:SharedPool.BeginClose($null, $null) } catch { } }
})
#endregion

#region Baslangic
Import-Settings
Update-AdCacheLabel
Update-DefLists
$form.Add_Shown({
    Start-AdLoad
    if ($txtSms.Text.Trim()) { Start-Connect -Quiet }   # kayitli SMS Provider'a arka planda baglan
})
[void]$form.ShowDialog()
#endregion
