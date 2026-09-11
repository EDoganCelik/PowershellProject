<#
.SYNOPSIS
    SCCM Run Script UI'dan kopyalanan sonuc metnini yapistirip CSV dosyasina
    donusturen basit bir Windows Forms araci.

.DESCRIPTION
    - Metin kutusuna, SCCM konsolunda "Run Script" sonrasi ekranda gorunen ve
      kopyalanan cikti yapistirilir (her satir bir cihazin sonucu, ';' ile
      ayrilmis degerler).
    - Belirtilen klasore, dosya adi "<gun-saat>-updates.csv" formatinda
      (orn. "11-14-updates.csv") baslik satiri eklenerek yazilir.
    - Dosya olusturulduktan sonra otomatik olarak varsayilan uygulamada
      (genelde Excel) acilir.

.NOTES
    Bu betik SCCM Run Script ile degil, teknisyenin kendi bilgisayarinda
    calistirilmak icin tasarlanmistir (grafik arayuz gerektirir).
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---- Birlesik CSV basligi ----
# Ilk 12 sutun: onceden hazirlanmis diger betigin ciktisi
# Son 9 sutun : client-update-check.ps1 ciktisi (Hostname haric, tekrar etmesin diye)
$CsvHeader = "Hostname;UpdateName;EvalState;PercentComplete;ErrorCode;LastShutdownTime;IsBlueScreen;CBS_NTSTATUS_FROM_WIN32;CBS_STATUS_SXS_FILE_HASH_MISMATCH;UD_CERTIFICATE_ERROR_0x80010002;UD_DOWNLOAD_ERROR_0x80d02002;BuildVersion;SistemDrive_FreeGB;SistemDrive_FreePercent;Wuauserv_Durum;BITS_Durum;CryptSvc_Durum;TrustedInstaller_Durum;SoftwareDistribution_BoyutMB;DISM_CheckHealth;DTS_DP_IndirmeHatalari"
$ExpectedColumnCount = ($CsvHeader -split ";").Count

$DefaultFolder = Join-Path $env:USERPROFILE "Desktop\SCCM-Raporlari"

# ---- Form ----
$form = New-Object System.Windows.Forms.Form
$form.Text = "SCCM Run Script Sonucu -> CSV"
$form.Size = New-Object System.Drawing.Size(760, 560)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = $form.Size

$lblPaste = New-Object System.Windows.Forms.Label
$lblPaste.Text = "SCCM Run Script ciktisini yapistirin (her satir bir cihaz, ';' ile ayrilmis):"
$lblPaste.Location = New-Object System.Drawing.Point(10, 10)
$lblPaste.Size = New-Object System.Drawing.Size(720, 20)
$form.Controls.Add($lblPaste)

$txtPaste = New-Object System.Windows.Forms.TextBox
$txtPaste.Multiline = $true
$txtPaste.ScrollBars = "Vertical"
$txtPaste.AcceptsReturn = $true
$txtPaste.Location = New-Object System.Drawing.Point(10, 35)
$txtPaste.Size = New-Object System.Drawing.Size(720, 350)
$txtPaste.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($txtPaste)

$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Text = "Kayit klasoru:"
$lblPath.Location = New-Object System.Drawing.Point(10, 398)
$lblPath.Size = New-Object System.Drawing.Size(90, 20)
$form.Controls.Add($lblPath)

$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Text = $DefaultFolder
$txtPath.Location = New-Object System.Drawing.Point(105, 395)
$txtPath.Size = New-Object System.Drawing.Size(500, 20)
$form.Controls.Add($txtPath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Gozat..."
$btnBrowse.Location = New-Object System.Drawing.Point(615, 393)
$btnBrowse.Size = New-Object System.Drawing.Size(115, 25)
$form.Controls.Add($btnBrowse)

$btnBrowse.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    if (Test-Path $txtPath.Text) { $folderDialog.SelectedPath = $txtPath.Text }
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtPath.Text = $folderDialog.SelectedPath
    }
})

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = ""
$lblStatus.Location = New-Object System.Drawing.Point(10, 425)
$lblStatus.Size = New-Object System.Drawing.Size(720, 40)
$form.Controls.Add($lblStatus)

$btnCreate = New-Object System.Windows.Forms.Button
$btnCreate.Text = "CSV Olustur ve Ac"
$btnCreate.Location = New-Object System.Drawing.Point(10, 470)
$btnCreate.Size = New-Object System.Drawing.Size(200, 35)
$btnCreate.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnCreate)

$btnCreate.Add_Click({
    $lblStatus.ForeColor = [System.Drawing.Color]::DarkRed
    $pastedText = $txtPaste.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($pastedText)) {
        $lblStatus.Text = "Once metin kutusuna SCCM sonucunu yapistirin."
        return
    }

    $lines = $pastedText -split "`r`n|`n|`r" | Where-Object { $_.Trim() -ne "" }
    if ($lines.Count -eq 0) {
        $lblStatus.Text = "Gecerli bir satir bulunamadi."
        return
    }

    $folder = $txtPath.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($folder)) {
        $lblStatus.Text = "Kayit klasoru bos olamaz."
        return
    }
    if (-not (Test-Path $folder)) {
        try {
            New-Item -Path $folder -ItemType Directory -Force | Out-Null
        } catch {
            $lblStatus.Text = "Klasor olusturulamadi: $($_.Exception.Message)"
            return
        }
    }

    # Kolon sayisi uyarisi (bilgilendirme amacli, olusturmayi engellemez)
    $mismatchCount = 0
    foreach ($line in $lines) {
        if (($line -split ";").Count -ne $ExpectedColumnCount) { $mismatchCount++ }
    }

    $timestamp = Get-Date -Format "dd-HH"
    $fileName  = "$timestamp-updates.csv"
    $fullPath  = Join-Path $folder $fileName

    try {
        $csvContent = @($CsvHeader) + $lines
        $csvContent | Out-File -FilePath $fullPath -Encoding UTF8 -Force
    } catch {
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkRed
        $lblStatus.Text = "CSV yazilamadi: $($_.Exception.Message)"
        return
    }

    if ($mismatchCount -gt 0) {
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkOrange
        $lblStatus.Text = "CSV olusturuldu ($fullPath) ancak $mismatchCount satirda kolon sayisi ($ExpectedColumnCount beklenirken) farkli. Basliklarla hizasinin dogru olup olmadigini kontrol et."
    } else {
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        $lblStatus.Text = "CSV basariyla olusturuldu: $fullPath"
    }

    try {
        Invoke-Item -Path $fullPath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("CSV olusturuldu ancak otomatik acilamadi:`n$fullPath", "Bilgi") | Out-Null
    }
})

[void]$form.ShowDialog()