#Requires -Version 5.1
<#
.SYNOPSIS
    Uzak makinelerde servis durumu kontrol / duzeltme araci (WinForms GUI)

.ACIKLAMA
    - Sol textbox: Kontrol edilecek bilgisayar isimleri (her satirda bir tane)
    - Ust textbox: Kontrol edilecek servis isimleri (virgulle ayrilmis)
    - Opsiyonel: Dosya yolu / Registry yolu / Registry Value adi kontrolu
    - "Kontrol Et"  -> tum makinelerdeki anlik durumu tabloya basar
    - "Durumu Duzelt" -> calismayan (stopped) servisleri baslatir, ONCEKI ve
      SONRAKI durumu hafizada tutar, ayri bir dialogda degisiklik raporu
      gosterir. Bu rapor CSV veya Excel'e aktarilabilir (Excel aktarilinca
      otomatik acilir) ve "Tab-separated" olarak panoya kopyalanabilir.

.PARALELLIK (ONEMLI)
    Kontrol ve duzeltme islemleri RunspacePool ile PARALEL calisir ancak
    havuz kapasitesi SABIT OLARAK 10 ile sinirlandirilmistir. Bu, ayni anda
    hicbir zaman 10'dan fazla makineye istek gitmeyecegi anlamina gelir
    (SOC ekiplerinin anlamsiz/agir bir trafik patlamasi gormemesi icin
    bilincli olarak konulmus bir esik degerdir). Islemler arka planda
    (UI donmadan) devam eder; ilerleme durum etiketinde gosterilir.

.GECMIS KAYITLAR
    Kayit otomatik yapilmaz. Kullanici "Kaydet" butonuna basar, acilan dialogda
    kayda bir isim verir. Kayda SADECE textbox verileri yazilir (servis isimleri,
    dosya yolu, registry yolu, registry value adi); bilgisayar isimleri ve sonuc
    tablosu kaydedilmez. Tum kayitlar C:\Users\<kullanici>\Desktop\ServiceStatus\
    status.json dosyasinda tek bir JSON dizisi olarak tutulur. En son secilen/
    kaydedilen kaydin kimligi ayni klasordeki lastselected.txt dosyasinda saklanir
    ve program ilk acildiginda o kayit otomatik yuklenir. Combobox son 30 kaydi listeler.

.ONEMLI - REMOTING YONTEMI
    Bu script Invoke-Command / PSSession / WinRM KULLANMAZ. Sadece klasik
    RPC / SMB tabanli remoting kullanilir:

      - Servis durumu/baslatma : Get-Service -ComputerName ... | Start-Service
      - Son kapanma            : Get-WinEvent -ComputerName
      - Dosya kontrolu         : Get-Item -Path \\Bilgisayar\C$\Klasor\dosya.txt
      - Registry kontrolu      : reg.exe query \\Bilgisayar\HKLM\...  (sadece HKLM/HKU)
      - Registry value okuma   : reg.exe query \\Bilgisayar\HKLM\... /v DegerAdi

    Hedef makinelerde: kontrolu yapan kullanici admin olmali, Admin$ paylasimi
    acik olmali (dosya kontrolu icin), Remote Registry servisi calisiyor olmali
    (registry kontrolu icin).
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==================== GLOBAL DEGISKENLER ====================
$script:LastResults           = @()    # Son kontrol sonuclari (ServiceDetail dahil)
$script:FixLog                = @()    # Duzeltme sonrasi degisiklik kayitlari (rapor icin)
$script:SuppressHistorySelect = $false
$script:HistoryItems          = @()

# ==================== YARDIMCI FONKSIYONLAR ====================
# NOT: Asagidaki fonksiyonlar hem ana pencere tarafindan, hem de paralel
# calisan runspace'ler tarafindan (kaynak kodu ${function:Isim} ile alinarak)
# kullanilir. Bu yuzden disaridan hicbir degiskene referans vermezler.

function Convert-ToUNCPath {
    # Get-Item/Test-Path'in -ComputerName parametresi olmadigi icin
    # yerel bir yolu (C:\Klasor\dosya.txt) idari paylasima ceviriyoruz:
    # \\Bilgisayar\C$\Klasor\dosya.txt
    param([string]$ComputerName, [string]$LocalPath)

    if ([string]::IsNullOrWhiteSpace($LocalPath)) { return $null }

    if ($LocalPath -match '^([A-Za-z]):\\(.*)$') {
        $drive = $Matches[1]
        $rest  = $Matches[2]
        return "\\$ComputerName\$drive`$\$rest"
    }

    # Zaten UNC ise oldugu gibi kullan
    return $LocalPath
}

function Get-RemoteRegistryValue {
    # Belirli bir registry degerinin (value) verisini uzaktan okur:
    # reg.exe query \\Bilgisayar\HKLM\...\Anahtar /v DegerAdi
    param([string]$ComputerName, [string]$RegPath, [string]$ValueName)

    if ([string]::IsNullOrWhiteSpace($RegPath) -or [string]::IsNullOrWhiteSpace($ValueName)) {
        return 'N/A'
    }

    $clean = $RegPath -replace '^(HKLM|HKEY_LOCAL_MACHINE):?\\?', 'HKLM\'
    $clean = $clean -replace '^(HKU|HKEY_USERS):?\\?', 'HKU\'
    $clean = $clean.TrimStart('\')

    if ($clean -notmatch '^(HKLM|HKU)\\') {
        return 'Desteklenmiyor (sadece HKLM/HKU)'
    }

    $uncRegPath = "\\$ComputerName\$clean"
    $output = & reg.exe query "$uncRegPath" /v "$ValueName" 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $output) {
        return 'Bulunamadi'
    }

    # Ornek cikti satiri:  "    DegerAdi    REG_SZ    DegerVerisi"
    $line = $output | Where-Object { $_ -match [regex]::Escape($ValueName) } | Select-Object -First 1
    if (-not $line) { return 'Bulunamadi' }

    if ($line -match '^\s*.+?\s+(REG_[A-Z_]+)\s+(.*)$') {
        return $Matches[2].Trim()
    }
    return $line.Trim()
}

function Get-RemoteStatus {
    param(
        [string]$ComputerName,
        [string[]]$ServiceNames,
        [string]$FilePath,
        [string]$RegPath,
        [string]$RegValueName
    )

    $row = [pscustomobject]@{
        ComputerName    = $ComputerName
        Hostname        = $ComputerName
        Running         = ''
        NotRunning      = ''
        LastShutdown    = 'N/A'
        IsOnline        = $false
        FilePathControl = 'N/A'
        RegistryValue   = 'N/A'
        ServiceDetail   = @()   # Name/Status listesi -> Duzelt asamasinda kullanilir
        ErrorMessage    = ''
    }

    $online = Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue
    $row.IsOnline = [bool]$online

    if (-not $online) {
        $row.ErrorMessage = 'Makine yanit vermiyor (ping basarisiz)'
        return $row
    }

    # ---- Servis durumu: Get-Service -ComputerName (RPC/SCM), tek tek kontrol
    #      edilir ki hangi servisin bulunamadigi (NotFound) ayirt edilebilsin ----
    $svcDetail = foreach ($s in $ServiceNames) {
        $s = $s.Trim()
        if (-not $s) { continue }
        try {
            $svc = Get-Service -ComputerName $ComputerName -Name $s -ErrorAction Stop
            [pscustomobject]@{ Name = $s; Status = $svc.Status.ToString() }
        }
        catch {
            [pscustomobject]@{ Name = $s; Status = 'NotFound' }
        }
    }
    $row.ServiceDetail = $svcDetail
    $row.Running    = ($svcDetail | Where-Object { $_.Status -eq 'Running' } | Select-Object -ExpandProperty Name) -join ', '
    $row.NotRunning = ($svcDetail | Where-Object { $_.Status -ne 'Running' } | ForEach-Object { "$($_.Name)($($_.Status))" }) -join ', '

    # ---- Son kapanma zamani: Get-WinEvent -ComputerName (RPC / Event Log) ----
    try {
        $evt = Get-WinEvent -ComputerName $ComputerName -FilterHashtable @{ LogName = 'System'; Id = 1074, 6006, 6008 } -MaxEvents 1 -ErrorAction Stop
        $row.LastShutdown = $evt.TimeCreated
    }
    catch {
        $row.LastShutdown = 'N/A'
    }

    # ---- Dosya kontrolu: UNC yol uzerinde Get-Item ----
    if (-not [string]::IsNullOrWhiteSpace($FilePath)) {
        $uncPath = Convert-ToUNCPath -ComputerName $ComputerName -LocalPath $FilePath
        try {
            $null = Get-Item -Path $uncPath -ErrorAction Stop
            $row.FilePathControl = $true
        }
        catch {
            $row.FilePathControl = $false
        }
    }

    # ---- Registry value (deger) okuma: reg.exe remote query /v ----
    if (-not [string]::IsNullOrWhiteSpace($RegPath) -and -not [string]::IsNullOrWhiteSpace($RegValueName)) {
        try {
            $row.RegistryValue = Get-RemoteRegistryValue -ComputerName $ComputerName -RegPath $RegPath -ValueName $RegValueName
        }
        catch {
            $row.RegistryValue = "Hata: $($_.Exception.Message)"
        }
    }

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
            $results += [pscustomobject]@{
                Hostname = $ComputerName; ServiceName = $s.Name
                OncekiDurum = $s.Status; SonrakiDurum = $s.Status
                Aciklama = 'Zaten calisiyordu, islem yapilmadi.'
            }
            continue
        }

        if ($s.Status -eq 'NotFound') {
            $results += [pscustomobject]@{
                Hostname = $ComputerName; ServiceName = $s.Name
                OncekiDurum = $s.Status; SonrakiDurum = $s.Status
                Aciklama = 'Servis bulunamadi, islem yapilmadi.'
            }
            continue
        }

        try {
            Get-Service -ComputerName $ComputerName -Name $s.Name | Start-Service -ErrorAction Stop
            Start-Sleep -Milliseconds 800

            $newStatus = (Get-Service -ComputerName $ComputerName -Name $s.Name).Status.ToString()

            $desc = if ($newStatus -eq 'Running') {
                "$($s.Status) durumundaydi, Running yapildi."
            } else {
                "Baslatildi ama durum hala $newStatus."
            }

            $results += [pscustomobject]@{
                Hostname = $ComputerName; ServiceName = $s.Name
                OncekiDurum = $s.Status; SonrakiDurum = $newStatus
                Aciklama = $desc
            }
        }
        catch {
            $results += [pscustomobject]@{
                Hostname = $ComputerName; ServiceName = $s.Name
                OncekiDurum = $s.Status; SonrakiDurum = 'Bilinmiyor'
                Aciklama = "Hata: $($_.Exception.Message)"
            }
        }
    }

    return $results
}

function Get-HostSummary {
    # Bir hostun tum servis duzeltme sonuclarini tek satirlik bir Status
    # metnine birlestirir. "Zaten calisiyordu" bilgisi onemsizdir ve raporda
    # gosterilmez - sadece duzeltilen, hala durdurulmus kalan ve bulunamayan
    # servisler raporlanir. Eger verilen tum servisler zaten calisiyorsa
    # tek cumle ile "Tum servisler calisiyor." denir.
    param([pscustomobject[]]$HostResults)

    $fixed        = $HostResults | Where-Object { $_.OncekiDurum -ne 'Running' -and $_.OncekiDurum -ne 'NotFound' -and $_.SonrakiDurum -eq 'Running' }
    $stillStopped = $HostResults | Where-Object { $_.OncekiDurum -ne 'Running' -and $_.OncekiDurum -ne 'NotFound' -and $_.SonrakiDurum -ne 'Running' }
    $notfound     = $HostResults | Where-Object { $_.OncekiDurum -eq 'NotFound' }

    $parts = @()
    if ($fixed)        { $parts += "$(($fixed        | Select-Object -ExpandProperty ServiceName) -join ',') stopped durumda running yapildi" }
    if ($stillStopped) { $parts += "$(($stillStopped | Select-Object -ExpandProperty ServiceName) -join ',') stopped durumda kaldi" }
    if ($notfound)     { $parts += "$(($notfound     | Select-Object -ExpandProperty ServiceName) -join ',') bulunamadi" }

    if (-not $parts) { return 'Tum servisler calisiyor.' }
    return (($parts -join ', ') + '.')
}

function Export-ToExcelAndOpen {
    # Verilen pscustomobject dizisini gercek bir .xlsx dosyasina yazar ve
    # (Excel COM Automation ile) dosyayi kullanicinin ekraninda otomatik acar.
    param(
        [object[]]$Data,
        [string]$DefaultFileName = "Rapor_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    )

    if (-not $Data -or @($Data).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Aktarilacak veri yok.', 'Bilgi') | Out-Null
        return
    }

    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = 'Excel Dosyasi (*.xlsx)|*.xlsx'
    $sfd.FileName = "$DefaultFileName.xlsx"
    if ($sfd.ShowDialog() -ne 'OK') { return }

    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Bu makinede Microsoft Excel bulunamadi, Excel'e aktarma yapilamiyor.`nGerekirse CSV secenegini kullanabilirsiniz.`n`nHata: $($_.Exception.Message)",
            'Hata') | Out-Null
        return
    }

    try {
        $excel.Visible = $false
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)

        $props = $Data[0].PSObject.Properties.Name
        for ($c = 0; $c -lt $props.Count; $c++) {
            $ws.Cells.Item(1, $c + 1) = $props[$c]
        }

        $r = 2
        foreach ($item in $Data) {
            for ($c = 0; $c -lt $props.Count; $c++) {
                $val = $item.($props[$c])
                $ws.Cells.Item($r, $c + 1) = [string]$val
            }
            $r++
        }

        $headerRange = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $props.Count))
        $headerRange.Font.Bold = $true
        $null = $ws.Columns.AutoFit()

        $wb.SaveAs($sfd.FileName)
        $excel.Visible = $true
        # Not: Excel penceresi kasitli olarak acik birakiliyor (kullanici
        # istegi geregi), bu yuzden COM referanslari burada serbest
        # birakilmiyor / Quit() cagirilmiyor.
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Excel'e aktarma sirasinda hata olustu: $($_.Exception.Message)", 'Hata') | Out-Null
        try { $excel.Quit() } catch { }
    }
}

function Show-FixResultDialog {
    param([pscustomobject[]]$Data)

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Duzeltme Sonuc Raporu'
    $dlg.Size = New-Object System.Drawing.Size(900,500)
    $dlg.StartPosition = 'CenterParent'

    $dgvLog = New-Object System.Windows.Forms.DataGridView
    $dgvLog.Location = New-Object System.Drawing.Point(10,10)
    $dgvLog.Size = New-Object System.Drawing.Size(860,400)
    $dgvLog.Anchor = 'Top,Bottom,Left,Right'
    $dgvLog.ReadOnly = $true
    $dgvLog.AllowUserToAddRows = $false
    $dgvLog.AutoSizeColumnsMode = 'Fill'
    $dgvLog.RowHeadersVisible = $false
    $dgvLog.SelectionMode = 'CellSelect'
    $dgvLog.MultiSelect = $true
    $dgvLog.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'

    $null = $dgvLog.Columns.Add('Hostname','Hostname')
    $null = $dgvLog.Columns.Add('Status','Status')

    foreach ($d in $Data) {
        $null = $dgvLog.Rows.Add($d.Hostname, $d.Status)
    }

    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Panoya Kopyala (Tablo)'
    $btnCopy.Location = New-Object System.Drawing.Point(10,420)
    $btnCopy.Size = New-Object System.Drawing.Size(180,30)
    $btnCopy.Anchor = 'Bottom,Left'

    $btnExportCsv = New-Object System.Windows.Forms.Button
    $btnExportCsv.Text = 'CSV Olarak Disa Aktar'
    $btnExportCsv.Location = New-Object System.Drawing.Point(200,420)
    $btnExportCsv.Size = New-Object System.Drawing.Size(180,30)
    $btnExportCsv.Anchor = 'Bottom,Left'

    $btnExportExcel = New-Object System.Windows.Forms.Button
    $btnExportExcel.Text = "Excel'e Aktar (Ac)"
    $btnExportExcel.Location = New-Object System.Drawing.Point(390,420)
    $btnExportExcel.Size = New-Object System.Drawing.Size(180,30)
    $btnExportExcel.Anchor = 'Bottom,Left'

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = 'Kapat'
    $btnClose.Location = New-Object System.Drawing.Point(770,420)
    $btnClose.Size = New-Object System.Drawing.Size(100,30)
    $btnClose.Anchor = 'Bottom,Right'

    $btnCopy.Add_Click({
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.AppendLine("Hostname`tStatus")
        foreach ($d in $Data) {
            [void]$sb.AppendLine("$($d.Hostname)`t$($d.Status)")
        }
        [System.Windows.Forms.Clipboard]::SetText($sb.ToString())
        [System.Windows.Forms.MessageBox]::Show('Tablo panoya kopyalandi. Word / Outlook / Excel gibi programlara yapistirabilirsiniz.','Bilgi') | Out-Null
    }.GetNewClosure())

    $btnExportCsv.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = 'CSV Dosyasi (*.csv)|*.csv'
        $sfd.FileName = "ServisDuzeltmeRaporu_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($sfd.ShowDialog() -eq 'OK') {
            $Data | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Disa aktarildi: $($sfd.FileName)",'Bilgi') | Out-Null
        }
    }.GetNewClosure())

    $btnExportExcel.Add_Click({
        Export-ToExcelAndOpen -Data $Data -DefaultFileName "ServisDuzeltmeRaporu_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    }.GetNewClosure())

    $btnClose.Add_Click({ $dlg.Close() })

    $dlg.Controls.AddRange(@($dgvLog,$btnCopy,$btnExportCsv,$btnExportExcel,$btnClose))
    $dlg.ShowDialog() | Out-Null
}

# ==================== GECMIS KAYIT (HISTORY) FONKSIYONLARI ====================

function Get-HistoryDir {
    # ONEMLI: Buradaki yol bilerek sabit ("C:\Users\<kullanici>\Desktop") olarak
    # kuruluyor. [Environment]::GetFolderPath('Desktop') veya $env:USERPROFILE
    # gibi yontemler OneDrive klasor yonlendirmesi aktifse gercek fiziksel
    # masaustu yerine OneDrive altindaki senkronize klasoru dondurebilir.
    # Bu yuzden $env:USERNAME ile kullanici adi alinip fiziksel yol elle
    # olusturuluyor; OneDrive devre disi/yoksa bile calisir.
    $dir = "C:\Users\$env:USERNAME\Desktop\ServiceStatus"
    if (-not (Test-Path $dir)) {
        $null = New-Item -ItemType Directory -Path $dir -Force
    }
    return $dir
}

function Get-HistoryFile {
    return (Join-Path (Get-HistoryDir) 'status.json')
}

function Get-LastSelectedFile {
    return (Join-Path (Get-HistoryDir) 'lastselected.txt')
}

function Get-LastSelectedKey {
    # En son secilen/kaydedilen kaydin Timestamp degerini dondurur (yoksa $null).
    $f = Get-LastSelectedFile
    if (-not (Test-Path $f)) { return $null }
    try {
        $v = Get-Content -Path $f -Raw -Encoding UTF8
        if (-not [string]::IsNullOrWhiteSpace($v)) { return $v.Trim() }
    }
    catch { }
    return $null
}

function Set-LastSelectedKey {
    param([string]$Key)
    if ([string]::IsNullOrWhiteSpace($Key)) { return }
    try { Set-Content -Path (Get-LastSelectedFile) -Value $Key -Encoding UTF8 } catch { }
}

function ConvertTo-JsonArraySafe {
    # NOT: PowerShell 5.1'de ConvertTo-Json, pipeline'a tek elemanli bir dizi
    # verildiginde diziyi "acip" tek bir JSON nesnesi (kose parantezsiz) uretir;
    # 2+ elemanli dizilerde ise duzgun bir JSON dizisi ([...]) uretir. Tum
    # kayitlari TEK bir dosyada tuttugumuz icin (0 veya 1 kayit olsa bile)
    # her zaman gecerli bir JSON dizisi elde etmemiz gerekiyor - bu fonksiyon
    # eleman sayisina bakmaksizin bunu garanti eder.
    param([object[]]$InputObject, [int]$Depth = 8)

    $arr = @($InputObject)
    if ($arr.Count -eq 0) { return '[]' }
    if ($arr.Count -eq 1) { return '[' + ($arr[0] | ConvertTo-Json -Depth $Depth) + ']' }
    return ($arr | ConvertTo-Json -Depth $Depth)
}

function New-HistoryRecord {
    # Tek bir kaydi (sadece textbox alanlari) olusturur.
    param($Name, $Timestamp, $Services, $FilePath, $RegPath, $RegValueName)
    return [pscustomobject]@{
        Name         = "$Name"
        Timestamp    = "$Timestamp"
        Services     = @($Services)
        FilePath     = "$FilePath"
        RegPath      = "$RegPath"
        RegValueName = "$RegValueName"
    }
}

function Expand-JsonArray {
    # ConvertFrom-Json, PowerShell surumune gore JSON dizisini ya kayit kayit ya
    # da TEK bir dizi nesnesi olarak dondurur. Hangisi olursa olsun, ic ice
    # dizileri acip kayitlari tek tek dondurur.
    param($Item)
    if ($null -eq $Item) { return }
    if ($Item -is [System.Array]) {
        foreach ($x in $Item) { Expand-JsonArray -Item $x }
    }
    else {
        $Item
    }
}

function Read-HistoryRecords {
    # status.json'u okur ve HER ZAMAN duz bir kayit listesi (eskiden yeniye)
    # dondurur. Eski kayitlardaki Results alani (bilgisayar isimleri) atilir.
    #
    # ONEMLI: PowerShell 5.1'de ConvertFrom-Json dizinin tamamini TEK nesne
    # olarak verebilir (@( ) bile bunu acmaz); bu yuzden Expand-JsonArray ile
    # kayitlar tek tek acilir ve foreach ile dolasilir. Pipeline'a verilmez.
    #
    # ONARIM: Onceki hatali surum tum kayitlari tek nesneye cokertmis olabilir
    # (Name/Timestamp/FilePath... alanlari dizi olarak). Boyle bir nesne
    # bulunursa indeks indeks tekrar ayri kayitlara acilir. Cokertme sirasinda
    # kayda ozel Services listeleri birbirine karistigi icin, acilan kayitlara
    # bu birlesik listenin tekrarsiz hali yazilir.
    $file = Get-HistoryFile
    if (-not (Test-Path $file)) { return @() }

    try {
        $raw = Get-Content -Path $file -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
        $parsed = @(Expand-JsonArray -Item (ConvertFrom-Json -InputObject $raw))
    }
    catch {
        return @()
    }

    $out = New-Object System.Collections.ArrayList
    foreach ($rec in $parsed) {
        if ($null -eq $rec) { continue }

        if ($rec.Timestamp -is [System.Array]) {
            $ts    = @($rec.Timestamp)
            $names = @($rec.Name)
            $fps   = @($rec.FilePath)
            $rps   = @($rec.RegPath)
            $rvs   = @($rec.RegValueName)
            $svcs  = @($rec.Services | Where-Object { $_ } | Select-Object -Unique)
            for ($i = 0; $i -lt $ts.Count; $i++) {
                [void]$out.Add((New-HistoryRecord `
                    -Name          $(if ($i -lt $names.Count) { $names[$i] } else { '' }) `
                    -Timestamp     $ts[$i] `
                    -Services      $svcs `
                    -FilePath      $(if ($i -lt $fps.Count) { $fps[$i] } else { '' }) `
                    -RegPath       $(if ($i -lt $rps.Count) { $rps[$i] } else { '' }) `
                    -RegValueName  $(if ($i -lt $rvs.Count) { $rvs[$i] } else { '' })))
            }
        }
        else {
            [void]$out.Add((New-HistoryRecord -Name $rec.Name -Timestamp $rec.Timestamp `
                -Services $rec.Services -FilePath $rec.FilePath -RegPath $rec.RegPath `
                -RegValueName $rec.RegValueName))
        }
    }

    return $out.ToArray()
}

function Save-RunHistory {
    # NOT: Kayda sadece textbox verileri yazilir (servisler, dosya yolu, registry
    # yolu, registry value adi). Bilgisayar isimleri ve sonuc tablosu SAKLANMAZ.
    # Tum kayitlar tek bir status.json dosyasinda bir dizi olarak tutulur.
    # Donus degeri: yeni kaydin Timestamp'i (kayit kimligi olarak kullanilir).
    param(
        [string]$Name,
        [string[]]$Services,
        [string]$FilePath,
        [string]$RegPath,
        [string]$RegValueName
    )

    $file = Get-HistoryFile

    $existing = @(Read-HistoryRecords)

    $newRecord = New-HistoryRecord -Name $Name -Timestamp (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') `
        -Services $Services -FilePath $FilePath -RegPath $RegPath -RegValueName $RegValueName

    $all = @($existing) + $newRecord

    # Depoyu buyumesin diye son 30 kayit tutulur, eskiler otomatik silinir.
    if (@($all).Count -gt 30) {
        $all = @($all | Select-Object -Last 30)
    }

    $json = ConvertTo-JsonArraySafe -InputObject $all -Depth 8
    Set-Content -Path $file -Value $json -Encoding UTF8

    return $newRecord.Timestamp
}

function Get-HistoryList {
    # En yeni kayit listenin basinda olacak sekilde donduruyoruz.
    $all = @(Read-HistoryRecords)

    for ($i = $all.Count - 1; $i -ge 0; $i--) {
        $data = $all[$i]
        $name = if ([string]::IsNullOrWhiteSpace($data.Name)) { $data.Timestamp } else { $data.Name }
        [pscustomobject]@{
            Data    = $data
            Display = "$name  ($($data.Timestamp))"
        }
    }
}

function Show-NameInputDialog {
    # Basit bir "isim gir" dialogu - kayda verilecek ismi kullanicidan alir.
    param(
        [string]$Title = 'Kaydi Isimlendir',
        [string]$Prompt = 'Bu kayit icin bir isim girin:'
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $Title
    $dlg.Size = New-Object System.Drawing.Size(380,150)
    $dlg.StartPosition = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $Prompt
    $lbl.Location = New-Object System.Drawing.Point(10,15)
    $lbl.AutoSize = $true

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(10,40)
    $txt.Size = New-Object System.Drawing.Size(340,23)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'Kaydet'
    $btnOk.Location = New-Object System.Drawing.Point(190,75)
    $btnOk.Size = New-Object System.Drawing.Size(75,28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Iptal'
    $btnCancel.Location = New-Object System.Drawing.Point(275,75)
    $btnCancel.Size = New-Object System.Drawing.Size(75,28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel
    $dlg.Controls.AddRange(@($lbl,$txt,$btnOk,$btnCancel))
    $dlg.Add_Shown({ $txt.Focus() })

    $result = $dlg.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($txt.Text)) {
        return $txt.Text.Trim()
    }
    return $null
}

# ==================== PARALEL CALISTIRMA ICIN SCRIPTBLOCK'LAR ====================
# RunspacePool icindeki her bir runspace ayri bir oturumdur; bu yuzden ihtiyac
# duyulan fonksiyonlarin kaynak kodu ${function:Isim} ile aliniyor ve
# calistirilacak scriptin basina gomuluyor. Bu sayede fonksiyon tanimlari
# TEK BIR YERDE (yukarida) tutulur, kopya kod olusmaz.

$script:CheckFuncDefs = @"
function Convert-ToUNCPath { $(${function:Convert-ToUNCPath}) }
function Get-RemoteRegistryValue { $(${function:Get-RemoteRegistryValue}) }
function Get-RemoteStatus { $(${function:Get-RemoteStatus}) }
"@

$script:CheckScriptBlock = [scriptblock]::Create(@"
param(`$ComputerName, `$ServiceNames, `$FilePath, `$RegPath, `$RegValueName)
$script:CheckFuncDefs
Get-RemoteStatus -ComputerName `$ComputerName -ServiceNames `$ServiceNames -FilePath `$FilePath -RegPath `$RegPath -RegValueName `$RegValueName
"@)

$script:FixFuncDefs = @"
function Start-RemoteFix { $(${function:Start-RemoteFix}) }
"@

$script:FixScriptBlock = [scriptblock]::Create(@"
param(`$ComputerName, `$ServiceDetail)
$script:FixFuncDefs
Start-RemoteFix -ComputerName `$ComputerName -ServiceDetail `$ServiceDetail
"@)

# ==================== SONUC TABLOSU (FILTRELENEBILIR / DUZENLI VERI) ====================

$script:ResultsTable = New-Object System.Data.DataTable
$null = $script:ResultsTable.Columns.Add('ComputerName')
$null = $script:ResultsTable.Columns.Add('Hostname')
$null = $script:ResultsTable.Columns.Add('Running')
$null = $script:ResultsTable.Columns.Add('NotRunning')
$null = $script:ResultsTable.Columns.Add('LastShutdown')
$null = $script:ResultsTable.Columns.Add('IsOnline')
$null = $script:ResultsTable.Columns.Add('FilePathControl')
$null = $script:ResultsTable.Columns.Add('RegistryValue')
$null = $script:ResultsTable.Columns.Add('ErrorMessage')

$script:BindingSrc = New-Object System.Windows.Forms.BindingSource
$script:BindingSrc.DataSource = $script:ResultsTable

# ==================== ANA FORM ====================

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Servis Durumu Kontrol ve Duzeltme Araci'
$form.Size = New-Object System.Drawing.Size(1260,800)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(1050,650)

$lblComputers = New-Object System.Windows.Forms.Label
$lblComputers.Text = 'Bilgisayar Isimleri (her satira bir tane):'
$lblComputers.Location = New-Object System.Drawing.Point(10,10)
$lblComputers.AutoSize = $true

$txtComputers = New-Object System.Windows.Forms.TextBox
$txtComputers.Multiline = $true
$txtComputers.ScrollBars = 'Vertical'
$txtComputers.Location = New-Object System.Drawing.Point(10,30)
$txtComputers.Size = New-Object System.Drawing.Size(200,700)
$txtComputers.Anchor = 'Top,Bottom,Left'

$lblServices = New-Object System.Windows.Forms.Label
$lblServices.Text = 'Servisler (virgulle ayirin):'
$lblServices.Location = New-Object System.Drawing.Point(220,10)
$lblServices.AutoSize = $true

$txtServices = New-Object System.Windows.Forms.TextBox
$txtServices.Location = New-Object System.Drawing.Point(220,30)
$txtServices.Size = New-Object System.Drawing.Size(400,23)
$txtServices.Anchor = 'Top,Left'
$txtServices.Text = 'WiaRpc, WinRM'

$lblFilePath = New-Object System.Windows.Forms.Label
$lblFilePath.Text = 'Dosya Yolu (opsiyonel, orn: C:\Path\file.txt):'
$lblFilePath.Location = New-Object System.Drawing.Point(220,60)
$lblFilePath.AutoSize = $true

$txtFilePath = New-Object System.Windows.Forms.TextBox
$txtFilePath.Location = New-Object System.Drawing.Point(220,80)
$txtFilePath.Size = New-Object System.Drawing.Size(400,23)
$txtFilePath.Anchor = 'Top,Left'

$lblRegPath = New-Object System.Windows.Forms.Label
$lblRegPath.Text = 'Registry Yolu (opsiyonel, orn: HKLM:\Software\X):'
$lblRegPath.Location = New-Object System.Drawing.Point(220,110)
$lblRegPath.AutoSize = $true

$txtRegPath = New-Object System.Windows.Forms.TextBox
$txtRegPath.Location = New-Object System.Drawing.Point(220,130)
$txtRegPath.Size = New-Object System.Drawing.Size(400,23)
$txtRegPath.Anchor = 'Top,Left'

$lblRegValueName = New-Object System.Windows.Forms.Label
$lblRegValueName.Text = 'Registry Value Adi (opsiyonel, orn: DisplayVersion):'
$lblRegValueName.Location = New-Object System.Drawing.Point(220,160)
$lblRegValueName.AutoSize = $true

$txtRegValueName = New-Object System.Windows.Forms.TextBox
$txtRegValueName.Location = New-Object System.Drawing.Point(220,180)
$txtRegValueName.Size = New-Object System.Drawing.Size(400,23)
$txtRegValueName.Anchor = 'Top,Left'

$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = 'Tabloyu Filtrele (Bilgisayar / Hostname):'
$lblFilter.Location = New-Object System.Drawing.Point(220,210)
$lblFilter.AutoSize = $true

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(220,230)
$txtFilter.Size = New-Object System.Drawing.Size(400,23)
$txtFilter.Anchor = 'Top,Left'

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = ''
$lblStatus.Location = New-Object System.Drawing.Point(220,260)
$lblStatus.AutoSize = $true
$lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue

$btnCheck = New-Object System.Windows.Forms.Button
$btnCheck.Text = 'Kontrol Et'
$btnCheck.Location = New-Object System.Drawing.Point(650,30)
$btnCheck.Size = New-Object System.Drawing.Size(150,30)
$btnCheck.Anchor = 'Top,Left'

$btnFix = New-Object System.Windows.Forms.Button
$btnFix.Text = 'Durumu Duzelt'
$btnFix.Location = New-Object System.Drawing.Point(650,70)
$btnFix.Size = New-Object System.Drawing.Size(150,30)
$btnFix.Anchor = 'Top,Left'
$btnFix.Enabled = $false

$lblHistory = New-Object System.Windows.Forms.Label
$lblHistory.Text = 'Kaydedilmis Kayitlar:'
$lblHistory.Location = New-Object System.Drawing.Point(650,110)
$lblHistory.AutoSize = $true

$cboHistory = New-Object System.Windows.Forms.ComboBox
$cboHistory.Location = New-Object System.Drawing.Point(650,130)
$cboHistory.Size = New-Object System.Drawing.Size(220,23)
$cboHistory.DropDownStyle = 'DropDownList'
$cboHistory.Anchor = 'Top,Left'

$btnSaveHistory = New-Object System.Windows.Forms.Button
$btnSaveHistory.Text = 'Kaydet'
$btnSaveHistory.Location = New-Object System.Drawing.Point(880,130)
$btnSaveHistory.Size = New-Object System.Drawing.Size(70,23)
$btnSaveHistory.Anchor = 'Top,Left'

$btnExportMain = New-Object System.Windows.Forms.Button
$btnExportMain.Text = "Tabloyu Excel'e Aktar (Ac)"
$btnExportMain.Location = New-Object System.Drawing.Point(650,165)
$btnExportMain.Size = New-Object System.Drawing.Size(300,30)
$btnExportMain.Anchor = 'Top,Left'

$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location = New-Object System.Drawing.Point(220,285)
$dgv.Size = New-Object System.Drawing.Size(1010,445)
$dgv.Anchor = 'Top,Bottom,Left,Right'
$dgv.ReadOnly = $true
$dgv.AllowUserToAddRows = $false
$dgv.AutoSizeColumnsMode = 'Fill'
$dgv.RowHeadersVisible = $false
# Hucreler tek tek (veya coklu) secilebilsin ve Ctrl+C ile kopyalanabilsin diye:
$dgv.SelectionMode = 'CellSelect'
$dgv.MultiSelect = $true
$dgv.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'

# NOT: Kolonlar burada elle tanimlaniyor (AutoGenerateColumns = $false).
# Onceden DataSource atanip hemen ardindan $dgv.Columns['Ad'].HeaderText=...
# yapiliyordu; DataGridView'un henuz bir Handle'i olusmadigi icin (form
# daha ShowDialog ile gosterilmeden once) AutoGenerateColumns kolonlari
# senkron olarak yaratmiyordu ve $dgv.Columns['...'] $null donuyordu -
# bu da "'HeaderText' cannot be found on this object" hatasina yol aciyordu.
# Elle kolon tanimlamak hem bu hatayi tamamen ortadan kaldirir hem de
# hangi kolonun gosterilecegini/gizlenecegini acikca kontrol etmemizi saglar.
$dgv.AutoGenerateColumns = $false

# ComputerName, Hostname ile ayni deger oldugu icin tabloda ayrica
# gosterilmiyor (gizli tutuluyor) ama Duzelt/Filtre islemleri onu
# kullandigi icin veri kaynaginda kaliyor.
$colComputerName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colComputerName.Name = 'ComputerName'
$colComputerName.DataPropertyName = 'ComputerName'
$colComputerName.HeaderText = 'Bilgisayar'
$colComputerName.Visible = $false

$colHostname = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colHostname.Name = 'Hostname'
$colHostname.DataPropertyName = 'Hostname'
$colHostname.HeaderText = 'Hostname'

$colRunning = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colRunning.Name = 'Running'
$colRunning.DataPropertyName = 'Running'
$colRunning.HeaderText = 'Calisan Servisler'

$colNotRunning = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colNotRunning.Name = 'NotRunning'
$colNotRunning.DataPropertyName = 'NotRunning'
$colNotRunning.HeaderText = 'Calismayan Servisler'

$colLastShutdown = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colLastShutdown.Name = 'LastShutdown'
$colLastShutdown.DataPropertyName = 'LastShutdown'
$colLastShutdown.HeaderText = 'Son Kapanma Zamani'

$colIsOnline = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colIsOnline.Name = 'IsOnline'
$colIsOnline.DataPropertyName = 'IsOnline'
$colIsOnline.HeaderText = 'Online mi'

$colFilePathControl = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colFilePathControl.Name = 'FilePathControl'
$colFilePathControl.DataPropertyName = 'FilePathControl'
$colFilePathControl.HeaderText = 'Dosya Kontrolu'

# Registry icin sadece degerin kendisi gosteriliyor (ayri bir
# "Registry Kontrolu" kolonu kaldirildi - gereksizdi).
$colRegistryValue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colRegistryValue.Name = 'RegistryValue'
$colRegistryValue.DataPropertyName = 'RegistryValue'
$colRegistryValue.HeaderText = 'Registry Value'

$colErrorMessage = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colErrorMessage.Name = 'ErrorMessage'
$colErrorMessage.DataPropertyName = 'ErrorMessage'
$colErrorMessage.HeaderText = 'Hata'

# NOT: PowerShell'de @(...) her zaman System.Object[] uretir; AddRange ise
# kesin olarak DataGridViewColumn[] bekledigi icin PowerShell bazen bunu
# otomatik donusturemeyip "Cannot convert ... System.Object[] ..." hatasi
# verir. Diziyi acikca [System.Windows.Forms.DataGridViewColumn[]] olarak
# tip donusumune tabi tutmak bu hatayi ortadan kaldirir.
$dgv.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@(
    $colHostname, $colRunning, $colNotRunning, $colLastShutdown,
    $colIsOnline, $colFilePathControl, $colRegistryValue,
    $colErrorMessage, $colComputerName
))
$dgv.DataSource = $script:BindingSrc

$dgv.add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }
    $rowView = $dgv.Rows[$e.RowIndex].DataBoundItem
    if (-not $rowView) { return }
    $isOnline   = "$($rowView['IsOnline'])"
    $notRunning = $rowView['NotRunning']
    if ($isOnline -eq 'False') {
        $e.CellStyle.BackColor = [System.Drawing.Color]::LightCoral
    }
    elseif ($notRunning) {
        $e.CellStyle.BackColor = [System.Drawing.Color]::LightYellow
    }
    else {
        $e.CellStyle.BackColor = [System.Drawing.Color]::LightGreen
    }
})

$form.Controls.AddRange(@(
    $lblComputers,$txtComputers,$lblServices,$txtServices,
    $lblFilePath,$txtFilePath,$lblRegPath,$txtRegPath,
    $lblRegValueName,$txtRegValueName,$lblFilter,$txtFilter,
    $btnCheck,$btnFix,$lblHistory,$cboHistory,$btnSaveHistory,$btnExportMain,
    $lblStatus,$dgv
))

# ==================== TABLO FILTRELEME ====================

$txtFilter.Add_TextChanged({
    $q = $txtFilter.Text.Trim().Replace("'", "''")
    if ([string]::IsNullOrWhiteSpace($q)) {
        $script:BindingSrc.RemoveFilter()
    }
    else {
        $script:BindingSrc.Filter = "ComputerName LIKE '%$q%' OR Hostname LIKE '%$q%'"
    }
})

# ==================== GECMIS KAYIT COMBOBOX ====================

function Update-HistoryCombo {
    # -SelectKey verilmezse en son secilen kayit (lastselected.txt) secilir;
    # o da yoksa/bulunamazsa listedeki en yeni kayit secilir. Secilen kayit
    # textboxlara yuklenir.
    param([string]$SelectKey)

    $script:SuppressHistorySelect = $true
    try {
        $script:HistoryItems = @(Get-HistoryList)
        $cboHistory.Items.Clear()
        foreach ($h in $script:HistoryItems) { [void]$cboHistory.Items.Add($h.Display) }

        if (-not $SelectKey) { $SelectKey = Get-LastSelectedKey }

        $idx = -1
        if ($SelectKey) {
            for ($i = 0; $i -lt @($script:HistoryItems).Count; $i++) {
                if ("$($script:HistoryItems[$i].Data.Timestamp)" -eq $SelectKey) { $idx = $i; break }
            }
        }
        if ($idx -lt 0 -and $cboHistory.Items.Count -gt 0) { $idx = 0 }

        if ($idx -ge 0) {
            $cboHistory.SelectedIndex = $idx
            Load-HistoryItem -Index $idx
        }
    }
    finally {
        # Hata olsa bile bayrak mutlaka sifirlanir; aksi halde combo secimi
        # bir daha textboxlari guncellemez.
        $script:SuppressHistorySelect = $false
    }
}

function Load-HistoryItem {
    # Kayittan sadece textbox verileri yuklenir (bilgisayar listesi ve sonuc
    # tablosu kayitta tutulmadigi icin degistirilmez).
    param([int]$Index)
    if ($Index -lt 0 -or $Index -ge @($script:HistoryItems).Count) { return }
    $data = $script:HistoryItems[$Index].Data
    if (-not $data) { return }

    $txtServices.Text     = (@($data.Services) -join ', ')
    $txtFilePath.Text     = "$($data.FilePath)"
    $txtRegPath.Text      = "$($data.RegPath)"
    $txtRegValueName.Text = "$($data.RegValueName)"

    $loadedName = if ([string]::IsNullOrWhiteSpace($data.Name)) { $data.Timestamp } else { $data.Name }
    $lblStatus.Text = "Kayittan yuklendi: $loadedName"
}

function Invoke-HistorySelection {
    # Kullanici combo'dan bir kayit sectiginde cagrilir: kaydi textboxlara
    # yukler ve "en son secilen" olarak isaretler. WinForms olay yoneticilerinde
    # olusan hatalar sessizce yutuldugu icin hata burada yakalanip gosterilir.
    param([int]$Index)
    try {
        Load-HistoryItem -Index $Index
        if ($Index -ge 0 -and $Index -lt @($script:HistoryItems).Count) {
            Set-LastSelectedKey -Key "$($script:HistoryItems[$Index].Data.Timestamp)"
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Kayit yuklenemedi: $($_.Exception.Message)",'Hata') | Out-Null
    }
}

# SelectionChangeCommitted: sadece KULLANICI secim yaptiginda tetiklenir (kod ile
# yapilan degisikliklerde tetiklenmez) ve bayraga bagimli degildir.
$cboHistory.Add_SelectionChangeCommitted({
    Invoke-HistorySelection -Index $cboHistory.SelectedIndex
})

# Klavye ile gezinme gibi durumlar icin yedek olarak SelectedIndexChanged de dinlenir.
$cboHistory.Add_SelectedIndexChanged({
    if ($script:SuppressHistorySelect) { return }
    Invoke-HistorySelection -Index $cboHistory.SelectedIndex
})

# ==================== TEXTBOX VERILERINI ISIMLE KAYDET ====================
# Otomatik kayit yok: kullanici "Kaydet"e basip bir isim verdiginde, o anki
# textbox verileri (servisler, dosya yolu, registry yolu, value adi) kaydedilir.

$btnSaveHistory.Add_Click({
    $servicesArr = @($txtServices.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($servicesArr.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Kaydedilecek bir veri yok. En az bir servis ismi girin.','Bilgi') | Out-Null
        return
    }

    $name = Show-NameInputDialog
    if (-not $name) { return }

    $key = Save-RunHistory -Name $name -Services $servicesArr -FilePath $txtFilePath.Text `
        -RegPath $txtRegPath.Text -RegValueName $txtRegValueName.Text
    Set-LastSelectedKey -Key $key
    Update-HistoryCombo -SelectKey $key
})

# ==================== TABLOYU EXCEL'E AKTAR ====================

$btnExportMain.Add_Click({
    if ($script:ResultsTable.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Once kontrol calistirin.','Bilgi') | Out-Null
        return
    }

    # Not: Aktif filtre varsa sadece filtrelenmis/gorunen satirlar aktarilir.
    $exportData = foreach ($drv in $script:BindingSrc) {
        $h = [ordered]@{}
        foreach ($col in $script:ResultsTable.Columns) { $h[$col.ColumnName] = $drv[$col.ColumnName] }
        [pscustomobject]$h
    }

    Export-ToExcelAndOpen -Data $exportData -DefaultFileName "ServisKontrolRaporu_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
})

# ==================== KONTROL ET (PARALEL, EN FAZLA 10 ES ZAMANLI) ====================

$script:CheckTimer = New-Object System.Windows.Forms.Timer
$script:CheckTimer.Interval = 200
$script:CheckJobs  = @()
$script:CheckPool  = $null
$script:CheckTotal = 0
$script:CheckDone  = 0

$btnCheck.Add_Click({
    $computers = $txtComputers.Text -split "`r`n|`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $services  = $txtServices.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }

    if (-not $computers -or @($computers).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('En az bir bilgisayar ismi girin.','Uyari') | Out-Null
        return
    }
    if (-not $services -or @($services).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('En az bir servis ismi girin.','Uyari') | Out-Null
        return
    }

    $btnCheck.Enabled = $false
    $btnFix.Enabled = $false
    $script:ResultsTable.Rows.Clear()
    $script:LastResults = @()
    $script:CheckTotal = @($computers).Count
    $script:CheckDone  = 0

    # Havuz kapasitesi SABIT 10: ayni anda en fazla 10 makineye istek gider.
    $script:CheckPool = [runspacefactory]::CreateRunspacePool(1, 10)
    $script:CheckPool.Open()

    $fp = $txtFilePath.Text
    $rp = $txtRegPath.Text
    $rv = $txtRegValueName.Text

    $script:CheckJobs = foreach ($c in $computers) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $script:CheckPool
        [void]$ps.AddScript($script:CheckScriptBlock).AddParameters(@{
            ComputerName = $c
            ServiceNames = $services
            FilePath     = $fp
            RegPath      = $rp
            RegValueName = $rv
        })
        [pscustomobject]@{ PS = $ps; Async = $ps.BeginInvoke(); Computer = $c; Done = $false }
    }

    $lblStatus.Text = "0 / $($script:CheckTotal) makine tamamlandi (ayni anda en fazla 10 istek)"
    $script:CheckTimer.Start()
})

$script:CheckTimer.Add_Tick({
    $stillRunning = $false

    foreach ($job in $script:CheckJobs) {
        if ($job.Done) { continue }

        if ($job.Async.IsCompleted) {
            $job.Done = $true
            $r = $null
            try { $r = $job.PS.EndInvoke($job.Async) } catch { $r = $null }
            $job.PS.Dispose()

            if ($r) {
                $script:LastResults += $r
                $null = $script:ResultsTable.Rows.Add(
                    $r.ComputerName, $r.Hostname, $r.Running, $r.NotRunning,
                    $r.LastShutdown, $r.IsOnline, $r.FilePathControl,
                    $r.RegistryValue, $r.ErrorMessage
                )
            }

            $script:CheckDone++
            $lblStatus.Text = "$($script:CheckDone) / $($script:CheckTotal) makine tamamlandi (ayni anda en fazla 10 istek)"
        }
        else {
            $stillRunning = $true
        }
    }

    if (-not $stillRunning) {
        $script:CheckTimer.Stop()
        $script:CheckPool.Close()
        $script:CheckPool.Dispose()

        $lblStatus.Text = "Tamamlandi. $($script:CheckTotal) makine kontrol edildi."
        $btnCheck.Enabled = $true
        $btnFix.Enabled = (@($script:LastResults | Where-Object { $_.IsOnline })).Count -gt 0
        # NOT: Otomatik gecmis kaydi kaldirildi - istenirse "Kaydet" ile elle kaydedilir.
    }
})

# ==================== DURUMU DUZELT (PARALEL, EN FAZLA 10 ES ZAMANLI) ====================

$script:FixTimer = New-Object System.Windows.Forms.Timer
$script:FixTimer.Interval = 200
$script:FixJobs  = @()
$script:FixPool  = $null
$script:FixTotal = 0
$script:FixDone  = 0

$btnFix.Add_Click({
    $targets = $script:LastResults | Where-Object {
        $_.IsOnline -and $_.ServiceDetail -and (@($_.ServiceDetail | Where-Object { $_.Status -ne 'Running' -and $_.Status -ne 'NotFound' })).Count -gt 0
    }

    if (-not $targets -or @($targets).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Duzeltilecek bir servis bulunamadi (hepsi zaten calisiyor).','Bilgi') | Out-Null
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "$(@($targets).Count) makinede durdurulmus servisler bulundu. Duzeltme islemini baslatmak istiyor musunuz? (Ayni anda en fazla 10 makine islenecek)",
        'Onay', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $btnFix.Enabled = $false
    $btnCheck.Enabled = $false
    $script:FixLog = @()
    $script:FixTotal = @($targets).Count
    $script:FixDone  = 0

    # Havuz kapasitesi SABIT 10: duzeltme de ayni esik degerine tabidir.
    $script:FixPool = [runspacefactory]::CreateRunspacePool(1, 10)
    $script:FixPool.Open()

    $script:FixJobs = foreach ($t in $targets) {
        $ps = [powershell]::Create()
        $ps.RunspacePool = $script:FixPool
        [void]$ps.AddScript($script:FixScriptBlock).AddParameters(@{
            ComputerName  = $t.ComputerName
            ServiceDetail = $t.ServiceDetail
        })
        [pscustomobject]@{ PS = $ps; Async = $ps.BeginInvoke(); Computer = $t.ComputerName; Done = $false }
    }

    $lblStatus.Text = "0 / $($script:FixTotal) makine duzeltildi (ayni anda en fazla 10 istek)"
    $script:FixTimer.Start()
})

$script:FixTimer.Add_Tick({
    $stillRunning = $false

    foreach ($job in $script:FixJobs) {
        if ($job.Done) { continue }

        if ($job.Async.IsCompleted) {
            $job.Done = $true
            $fixResult = $null
            try { $fixResult = $job.PS.EndInvoke($job.Async) } catch { $fixResult = $null }
            $job.PS.Dispose()

            if ($fixResult) {
                $script:FixLog += $fixResult
                $newRunning    = ($fixResult | Where-Object { $_.SonrakiDurum -eq 'Running' } | Select-Object -ExpandProperty ServiceName) -join ', '
                $newNotRunning = ($fixResult | Where-Object { $_.SonrakiDurum -ne 'Running' } | ForEach-Object { "$($_.ServiceName)($($_.SonrakiDurum))" }) -join ', '

                foreach ($row in $script:ResultsTable.Rows) {
                    if ($row['ComputerName'] -eq $job.Computer) {
                        $row['Running']    = $newRunning
                        $row['NotRunning'] = $newNotRunning
                    }
                }
            }

            $script:FixDone++
            $lblStatus.Text = "$($script:FixDone) / $($script:FixTotal) makine duzeltildi (ayni anda en fazla 10 istek)"
        }
        else {
            $stillRunning = $true
        }
    }

    if (-not $stillRunning) {
        $script:FixTimer.Stop()
        $script:FixPool.Close()
        $script:FixPool.Dispose()

        $lblStatus.Text = 'Duzeltme islemi tamamlandi.'
        $btnFix.Enabled = $true
        $btnCheck.Enabled = $true

        $hostSummaries = $script:FixLog | Group-Object Hostname | ForEach-Object {
            [pscustomobject]@{ Hostname = $_.Name; Status = Get-HostSummary -HostResults $_.Group }
        }
        # NOT: Otomatik gecmis kaydi kaldirildi - istenirse "Kaydet" ile elle kaydedilir.

        Show-FixResultDialog -Data $hostSummaries
    }
})

# ==================== FORM KAPANIRKEN TEMIZLIK ====================

$form.Add_FormClosing({
    if ($script:CheckTimer.Enabled) { $script:CheckTimer.Stop() }
    if ($script:FixTimer.Enabled) { $script:FixTimer.Stop() }
    if ($script:CheckPool) { try { $script:CheckPool.Close(); $script:CheckPool.Dispose() } catch { } }
    if ($script:FixPool) { try { $script:FixPool.Close(); $script:FixPool.Dispose() } catch { } }
})

# ==================== BASLANGICTA EN SON SECILEN KAYDI YUKLE ====================

Update-HistoryCombo

[void]$form.ShowDialog()