#Requires -Version 5.1
<#
.SYNOPSIS
    Uzak makinelerde servis durumu kontrol / duzeltme araci (WinForms GUI)

.ACIKLAMA
    - Sol textbox: Kontrol edilecek bilgisayar isimleri (her satirda bir tane)
    - Ust textbox: Kontrol edilecek servis isimleri (virgulle ayrilmis)
    - Opsiyonel: Dosya yolu / Registry yolu kontrolu (servisten bagimsiz, genel kontrol)
    - "Kontrol Et"  -> tum makinelerdeki anlik durumu tabloya basar
    - "Durumu Duzelt" -> calismayan (stopped) servisleri baslatir, ONCEKI ve SONRAKI
      durumu hafizada tutar, ayri bir dialogda degisiklik raporu gosterir.
      Bu rapor CSV'e aktarilabilir ve "Tab-separated" olarak panoya kopyalanip
      Word / Outlook / Excel gibi yerlere tablo seklinde yapistirilabilir.

.ONEMLI - REMOTING YONTEMI
    Bu script Invoke-Command / PSSession / WinRM KULLANMAZ. Sadece klasik
    RPC / SMB tabanli remoting kullanilir:

      - Servis durumu/baslatma : Get-Service -ComputerName ... | Start-Service
      - Son kapanma            : Get-WinEvent -ComputerName
      - Dosya kontrolu         : Get-Item -Path \\Bilgisayar\C$\Klasor\dosya.txt (UNC/idari paylasim)
      - Registry kontrolu      : reg.exe query \\Bilgisayar\HKLM\...  (sadece HKLM/HKU)

    Hedef makinelerde: kontrolu yapan kullanici admin olmali, Admin$ paylasimi
    acik olmali (dosya kontrolu icin), Remote Registry servisi calisiyor olmali
    (registry kontrolu icin).
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ==================== GLOBAL DEGISKENLER ====================
$script:LastResults = @()   # Son kontrol sonuclari (ServiceDetail dahil) - Duzelt asamasinda "onceki durum" olarak kullanilir
$script:FixLog       = @()   # Duzeltme sonrasi degisiklik kayitlari (rapor icin)

# ==================== YARDIMCI FONKSIYONLAR ====================

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

function Test-RemoteRegistryPath {
    # Get-Item ile uzak registry kontrolu yoktur; reg.exe'nin yerlesik
    # remote sorgu ozelligi kullanilir: reg.exe query \\Bilgisayar\HKLM\...
    # Sadece HKLM ve HKU uzaktan sorgulanabilir (HKCU desteklenmez).
    param([string]$ComputerName, [string]$RegPath)

    if ([string]::IsNullOrWhiteSpace($RegPath)) { return 'N/A' }

    $clean = $RegPath -replace '^(HKLM|HKEY_LOCAL_MACHINE):?\\?', 'HKLM\'
    $clean = $clean -replace '^(HKU|HKEY_USERS):?\\?', 'HKU\'
    $clean = $clean.TrimStart('\')

    if ($clean -notmatch '^(HKLM|HKU)\\') {
        return 'Desteklenmiyor (sadece HKLM/HKU)'
    }

    $uncRegPath = "\\$ComputerName\$clean"
    $null = & reg.exe query "$uncRegPath" 2>$null
    return ($LASTEXITCODE -eq 0)
}

function Get-RemoteStatus {
    param(
        [string]$ComputerName,
        [string[]]$ServiceNames,
        [string]$FilePath,
        [string]$RegPath
    )

    $row = [pscustomobject]@{
        ComputerName    = $ComputerName
        Hostname        = $ComputerName
        Running         = ''
        NotRunning      = ''
        LastShutdown    = 'N/A'
        IsOnline        = $false
        FilePathControl = 'N/A'
        RegistryControl = 'N/A'
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

    # ---- Registry kontrolu: reg.exe remote query ----
    if (-not [string]::IsNullOrWhiteSpace($RegPath)) {
        try {
            $row.RegistryControl = Test-RemoteRegistryPath -ComputerName $ComputerName -RegPath $RegPath
        }
        catch {
            $row.RegistryControl = "Hata: $($_.Exception.Message)"
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

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = 'CSV Olarak Disa Aktar'
    $btnExport.Location = New-Object System.Drawing.Point(200,420)
    $btnExport.Size = New-Object System.Drawing.Size(180,30)
    $btnExport.Anchor = 'Bottom,Left'

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

    $btnExport.Add_Click({
        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = 'CSV Dosyasi (*.csv)|*.csv'
        $sfd.FileName = "ServisDuzeltmeRaporu_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($sfd.ShowDialog() -eq 'OK') {
            $Data | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Disa aktarildi: $($sfd.FileName)",'Bilgi') | Out-Null
        }
    }.GetNewClosure())

    $btnClose.Add_Click({ $dlg.Close() })

    $dlg.Controls.AddRange(@($dgvLog,$btnCopy,$btnExport,$btnClose))
    $dlg.ShowDialog() | Out-Null
}

# ==================== ANA FORM ====================

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Servis Durumu Kontrol ve Duzeltme Araci'
$form.Size = New-Object System.Drawing.Size(1200,700)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(950,550)

$lblComputers = New-Object System.Windows.Forms.Label
$lblComputers.Text = 'Bilgisayar Isimleri (her satira bir tane):'
$lblComputers.Location = New-Object System.Drawing.Point(10,10)
$lblComputers.AutoSize = $true

$txtComputers = New-Object System.Windows.Forms.TextBox
$txtComputers.Multiline = $true
$txtComputers.ScrollBars = 'Vertical'
$txtComputers.Location = New-Object System.Drawing.Point(10,30)
$txtComputers.Size = New-Object System.Drawing.Size(200,560)
$txtComputers.Anchor = 'Top,Bottom,Left'

$lblServices = New-Object System.Windows.Forms.Label
$lblServices.Text = 'search services: (virgulle ayirin)'
$lblServices.Location = New-Object System.Drawing.Point(220,10)
$lblServices.AutoSize = $true

$txtServices = New-Object System.Windows.Forms.TextBox
$txtServices.Location = New-Object System.Drawing.Point(220,30)
$txtServices.Size = New-Object System.Drawing.Size(400,23)
$txtServices.Anchor = 'Top,Left,Right'
$txtServices.Text = 'WiaRpc, WinRM'

$lblFilePath = New-Object System.Windows.Forms.Label
$lblFilePath.Text = 'Dosya Yolu (opsiyonel, orn: C:\Path\file.txt):'
$lblFilePath.Location = New-Object System.Drawing.Point(220,60)
$lblFilePath.AutoSize = $true

$txtFilePath = New-Object System.Windows.Forms.TextBox
$txtFilePath.Location = New-Object System.Drawing.Point(220,80)
$txtFilePath.Size = New-Object System.Drawing.Size(400,23)
$txtFilePath.Anchor = 'Top,Left,Right'

$lblRegPath = New-Object System.Windows.Forms.Label
$lblRegPath.Text = 'Registry Yolu (opsiyonel, orn: HKLM:\Software\X):'
$lblRegPath.Location = New-Object System.Drawing.Point(220,110)
$lblRegPath.AutoSize = $true

$txtRegPath = New-Object System.Windows.Forms.TextBox
$txtRegPath.Location = New-Object System.Drawing.Point(220,130)
$txtRegPath.Size = New-Object System.Drawing.Size(400,23)
$txtRegPath.Anchor = 'Top,Left,Right'

$btnCheck = New-Object System.Windows.Forms.Button
$btnCheck.Text = 'Kontrol Et'
$btnCheck.Location = New-Object System.Drawing.Point(640,30)
$btnCheck.Size = New-Object System.Drawing.Size(130,30)
$btnCheck.Anchor = 'Top,Left'

$btnFix = New-Object System.Windows.Forms.Button
$btnFix.Text = 'Durumu Duzelt'
$btnFix.Location = New-Object System.Drawing.Point(640,70)
$btnFix.Size = New-Object System.Drawing.Size(130,30)
$btnFix.Anchor = 'Top,Left'
$btnFix.Enabled = $false

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = ''
$lblStatus.Location = New-Object System.Drawing.Point(220,160)
$lblStatus.AutoSize = $true
$lblStatus.ForeColor = [System.Drawing.Color]::DarkBlue

$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Location = New-Object System.Drawing.Point(220,190)
$dgv.Size = New-Object System.Drawing.Size(950,400)
$dgv.Anchor = 'Top,Bottom,Left,Right'
$dgv.ReadOnly = $true
$dgv.AllowUserToAddRows = $false
$dgv.AutoSizeColumnsMode = 'Fill'
$dgv.SelectionMode = 'FullRowSelect'
$dgv.RowHeadersVisible = $false

$null = $dgv.Columns.Add('ComputerName','Bilgisayar')
$null = $dgv.Columns.Add('Hostname','Hostname')
$null = $dgv.Columns.Add('Running','Calisan Servisler')
$null = $dgv.Columns.Add('NotRunning','Calismayan Servisler')
$null = $dgv.Columns.Add('LastShutdown','Son Kapanma Zamani')
$null = $dgv.Columns.Add('IsOnline','Online mi')
$null = $dgv.Columns.Add('FilePathControl','Dosya Kontrolu')
$null = $dgv.Columns.Add('RegistryControl','Registry Kontrolu')
$null = $dgv.Columns.Add('ErrorMessage','Hata')

$form.Controls.AddRange(@(
    $lblComputers,$txtComputers,$lblServices,$txtServices,
    $lblFilePath,$txtFilePath,$lblRegPath,$txtRegPath,
    $btnCheck,$btnFix,$lblStatus,$dgv
))

# ==================== EVENT HANDLERS ====================

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
    $dgv.Rows.Clear()
    $script:LastResults = @()

    $i = 0
    foreach ($c in $computers) {
        $i++
        $lblStatus.Text = "Kontrol ediliyor: $c ($i/$(@($computers).Count))"
        [System.Windows.Forms.Application]::DoEvents()

        $r = Get-RemoteStatus -ComputerName $c -ServiceNames $services -FilePath $txtFilePath.Text -RegPath $txtRegPath.Text
        $script:LastResults += $r

        $rowIdx = $dgv.Rows.Add(
            $r.ComputerName, $r.Hostname, $r.Running, $r.NotRunning,
            $r.LastShutdown, $r.IsOnline, $r.FilePathControl, $r.RegistryControl, $r.ErrorMessage
        )

        if (-not $r.IsOnline) {
            $dgv.Rows[$rowIdx].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightCoral
        } elseif ($r.NotRunning) {
            $dgv.Rows[$rowIdx].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
        } else {
            $dgv.Rows[$rowIdx].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
        }
    }

    $lblStatus.Text = "Tamamlandi. $(@($computers).Count) makine kontrol edildi."
    $btnCheck.Enabled = $true
    $btnFix.Enabled = (@($script:LastResults | Where-Object { $_.IsOnline })).Count -gt 0
})

$btnFix.Add_Click({
    $targets = $script:LastResults | Where-Object {
        $_.IsOnline -and $_.ServiceDetail -and (@($_.ServiceDetail | Where-Object { $_.Status -ne 'Running' -and $_.Status -ne 'NotFound' })).Count -gt 0
    }

    if (-not $targets -or @($targets).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Duzeltilecek bir servis bulunamadi (hepsi zaten calisiyor).','Bilgi') | Out-Null
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "$(@($targets).Count) makinede durdurulmus servisler bulundu. Duzeltme islemini baslatmak istiyor musunuz?",
        'Onay', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $btnFix.Enabled = $false
    $btnCheck.Enabled = $false
    $script:FixLog = @()

    $i = 0
    foreach ($t in $script:LastResults) {
        if (-not $t.IsOnline -or -not $t.ServiceDetail) { continue }
        $i++
        $lblStatus.Text = "Duzeltiliyor: $($t.ComputerName) ($i)"
        [System.Windows.Forms.Application]::DoEvents()

        $fixResult = Start-RemoteFix -ComputerName $t.ComputerName -ServiceDetail $t.ServiceDetail
        $script:FixLog += $fixResult

        $newRunning    = ($fixResult | Where-Object { $_.SonrakiDurum -eq 'Running' } | Select-Object -ExpandProperty ServiceName) -join ', '
        $newNotRunning = ($fixResult | Where-Object { $_.SonrakiDurum -ne 'Running' } | ForEach-Object { "$($_.ServiceName)($($_.SonrakiDurum))" }) -join ', '

        foreach ($row in $dgv.Rows) {
            if ($row.Cells['ComputerName'].Value -eq $t.ComputerName) {
                $row.Cells['Running'].Value = $newRunning
                $row.Cells['NotRunning'].Value = $newNotRunning
                if (-not $newNotRunning) { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen }
            }
        }
    }

    $lblStatus.Text = 'Duzeltme islemi tamamlandi.'
    $btnFix.Enabled = $true
    $btnCheck.Enabled = $true

    $hostSummaries = $script:FixLog | Group-Object Hostname | ForEach-Object {
        [pscustomobject]@{ Hostname = $_.Name; Status = Get-HostSummary -HostResults $_.Group }
    }

    Show-FixResultDialog -Data $hostSummaries
})

[void]$form.ShowDialog()