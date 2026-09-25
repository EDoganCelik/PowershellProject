<#
.SYNOPSIS
    SCCM'de boş Software Update Group ve boş Software Update Deployment Package oluşturur (UI ile).

.DESCRIPTION
    - Boş bir Software Update Group oluşturur (ad + description).
    - Boş bir Software Update Deployment Package oluşturur (ad + description + kaynak yolu).
    - Kaynak klasör yoksa oluşturur.
    - Her başarılı çalıştırmayı C:\Users\<kullanıcı>\Desktop\update-state\state.json dosyasına kaydeder
      (update-state klasörü yoksa otomatik oluşturulur).
      Yol OneDrive/bulut ile senkronize görünüyorsa buluta yazmaz; %LOCALAPPDATA%\update-state'e düşer.
    - Script tekrar açıldığında önceki çalıştırmayı gösterir ve alanları o verilerle doldurur.
    - "Ayı İlerlet" butonu, alanlardaki yyyy-MM / yyyy_MM / yyyy.MM / yyyyMM kalıplarını +1 ay ilerletir.

.NOTES
    Configuration Manager Console'un kurulu olduğu bir makinede çalıştırın.
    Kullanıcının SCCM'de Software Update Group ve Deployment Package oluşturma yetkisi olmalıdır.
    Site sunucusunun bilgisayar hesabının, kaynak yolu (UNC) üzerinde yazma yetkisi olmalıdır.
#>

#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ------------------------------------------------------------------
# Durum (state) dosyası
# ------------------------------------------------------------------
$script:HistoryMax = 24

# Yolun (veya üst klasörlerinin) OneDrive / bulut senkronizasyonuna ait olup olmadığını kontrol eder.
function Test-CloudSyncedPath {
    param([string]$Path)
    if (-not $Path) { return $false }

    # 1) OneDrive kök klasörleri (ortam değişkenleri)
    $target = $Path.TrimEnd('\') + '\'
    $roots = @($env:OneDrive, $env:OneDriveCommercial, $env:OneDriveConsumer) | Where-Object { $_ }
    foreach ($r in $roots) {
        if ($target.StartsWith(($r.TrimEnd('\') + '\'), [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }

    # 2) Yolun kendisi veya mevcut üst klasörleri: yönlendirme (junction/symlink) hedefi
    #    OneDrive/SharePoint mi, ya da bulut "yer tutucu" (Files On-Demand) özniteliği var mı?
    $cur = $Path
    while ($cur) {
        if (Test-Path -LiteralPath $cur) {
            $item = Get-Item -LiteralPath $cur -Force -ErrorAction SilentlyContinue
            if ($item) {
                if ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                    $t = @($item.Target) -join ' '
                    if ($t -match 'OneDrive|SharePoint') { return $true }
                }
                $attr = [int]$item.Attributes
                if (($attr -band 0x400000) -or ($attr -band 0x40000)) { return $true }   # RecallOnDataAccess / RecallOnOpen
            }
        }
        $parent = Split-Path -Path $cur -Parent
        if (-not $parent -or $parent -eq $cur) { break }
        $cur = $parent
    }

    # 3) Yol içinde OneDrive / SharePoint klasör adı geçiyor mu?
    if ($Path -match '\\(OneDrive|SharePoint)[^\\]*(\\|$)') { return $true }
    return $false
}

# Öncelikli konum: Desktop\update-state. Bulut senkronizasyonlu ise yerel LOCALAPPDATA'ya düşer.
$script:StateEnabled = $true
$script:StateWarning = $null
$script:StateDir     = Join-Path $env:USERPROFILE 'Desktop\update-state'
if (Test-CloudSyncedPath $script:StateDir) {
    $script:StateWarning = "Masaüstü yolu bulut (OneDrive) ile senkronize görünüyor: $($script:StateDir). Bulut'a yazılmaması için yerel klasör kullanılacak."
    $script:StateDir = Join-Path $env:LOCALAPPDATA 'update-state'
    if (Test-CloudSyncedPath $script:StateDir) {
        $script:StateEnabled = $false
        $script:StateWarning += ' Yerel klasör de bulutla senkronize göründüğü için durum KAYDEDİLMEYECEK.'
    } else {
        $script:StateWarning += " Yeni konum: $($script:StateDir)"
    }
}
$script:StateFile = Join-Path $script:StateDir 'state.json'

function Get-State {
    if (-not $script:StateEnabled) { return $null }
    if (Test-Path -LiteralPath $script:StateFile) {
        try {
            return (Get-Content -LiteralPath $script:StateFile -Raw -Encoding UTF8 | ConvertFrom-Json)
        } catch {
            return $null
        }
    }
    return $null
}

function Save-Entry {
    param([hashtable]$Entry)

    if (-not $script:StateEnabled) { return $false }

    $history = @()
    $state = Get-State
    if ($state -and $state.History) { $history = @($state.History) }
    $history += $Entry
    if ($history.Count -gt $script:HistoryMax) {
        $history = $history[($history.Count - $script:HistoryMax)..($history.Count - 1)]
    }

    $obj = [ordered]@{
        LastRun = $Entry
        History = $history
    }

    if (-not (Test-Path -LiteralPath $script:StateDir)) {
        New-Item -ItemType Directory -Path $script:StateDir -Force | Out-Null
    }
    $obj | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $script:StateFile -Encoding UTF8
    return $true
}

function Format-Entry {
    param($e)
    if (-not $e) { return 'Önceki çalıştırma kaydı bulunamadı.' }
    $lines = @(
        "Tarih                : $($e.Timestamp)"
        "Site Kodu            : $($e.SiteCode)"
        "Provider Sunucu      : $($e.ProviderServer)"
        "Update Group Adı     : $($e.SugName)"
        "Update Group Açıkl.  : $($e.SugDescription)"
        "Update Group CI_ID   : $($e.SugId)"
        "Package Adı          : $($e.PkgName)"
        "Package Açıklama     : $($e.PkgDescription)"
        "Ana Klasör Yolu     : $($e.PkgBasePath)"
        "Package Kaynak Yolu  : $($e.PkgPath)"
        "Package ID           : $($e.PkgId)"
    )
    return ($lines -join [Environment]::NewLine)
}

# ------------------------------------------------------------------
# UI yardımcıları
# ------------------------------------------------------------------
function New-UiLabel {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 130)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size($W, 20)
    $Parent.Controls.Add($l)
    return $l
}

function New-UiTextBox {
    param($Parent, [int]$X, [int]$Y, [int]$W, [int]$H = 23, [bool]$Multiline = $false)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($X, $Y)
    $t.Size = New-Object System.Drawing.Size($W, $H)
    if ($Multiline) {
        $t.Multiline = $true
        $t.ScrollBars = 'Vertical'
    }
    $Parent.Controls.Add($t)
    return $t
}

function New-UiGroup {
    param($Parent, [string]$Text, [int]$Y, [int]$H)
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = $Text
    $g.Location = New-Object System.Drawing.Point(12, $Y)
    $g.Size = New-Object System.Drawing.Size(740, $H)
    $Parent.Controls.Add($g)
    return $g
}

# ------------------------------------------------------------------
# Form
# ------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'SCCM - Software Update Group & Deployment Package Oluşturucu'
$form.Size = New-Object System.Drawing.Size(780, 815)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# --- Bağlantı ---
$gConn = New-UiGroup $form 'SCCM Bağlantısı' 10 62
[void](New-UiLabel $gConn 'Site Kodu:' 12 27 70)
$txtSite = New-UiTextBox $gConn 85 24 80
[void](New-UiLabel $gConn 'SMS Provider Sunucusu:' 200 27 140)
$txtServer = New-UiTextBox $gConn 345 24 380

# --- Software Update Group ---
$gSug = New-UiGroup $form 'Software Update Group (boş oluşturulur)' 78 118
[void](New-UiLabel $gSug 'Grup Adı:' 12 27)
$txtSugName = New-UiTextBox $gSug 145 24 580
[void](New-UiLabel $gSug 'Description:' 12 57)
$txtSugDesc = New-UiTextBox $gSug 145 54 580 52 $true

# --- Deployment Package ---
$gPkg = New-UiGroup $form 'Software Update Deployment Package (boş oluşturulur)' 204 200
[void](New-UiLabel $gPkg 'Package Adı:' 12 27)
$txtPkgName = New-UiTextBox $gPkg 145 24 580
[void](New-UiLabel $gPkg 'Ana Klasör Yolu (UNC):' 12 57 135)
$txtPkgPath = New-UiTextBox $gPkg 145 54 580
[void](New-UiLabel $gPkg 'Description:' 12 87)
$txtPkgDesc = New-UiTextBox $gPkg 145 84 580 52 $true
[void](New-UiLabel $gPkg 'Oluşacak Yol:' 12 145)
$lblFullPath = New-UiLabel $gPkg '' 145 145 580
$lblFullPath.ForeColor = [System.Drawing.Color]::DarkBlue
$lblHint = New-UiLabel $gPkg 'Not: Package adı, ana klasör yolunun sonuna otomatik eklenir; klasör yoksa oluşturulur. Description en fazla 127 karakter.' 12 168 720
$lblHint.ForeColor = [System.Drawing.Color]::DimGray

# --- Önceki çalıştırma ---
$gPrev = New-UiGroup $form 'Önceki Çalıştırma' 412 150
$txtPrev = New-Object System.Windows.Forms.TextBox
$txtPrev.Location = New-Object System.Drawing.Point(10, 20)
$txtPrev.Size = New-Object System.Drawing.Size(720, 120)
$txtPrev.Multiline = $true
$txtPrev.ReadOnly = $true
$txtPrev.ScrollBars = 'Both'
$txtPrev.WordWrap = $false
$txtPrev.Font = New-Object System.Drawing.Font('Consolas', 9)
$gPrev.Controls.Add($txtPrev)

# --- Butonlar ---
function New-UiButton {
    param([string]$Text, [int]$X, [int]$W)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Location = New-Object System.Drawing.Point($X, 572)
    $b.Size = New-Object System.Drawing.Size($W, 32)
    $form.Controls.Add($b)
    return $b
}
$btnCreate  = New-UiButton 'Oluştur' 12 130
$btnLoad    = New-UiButton 'Önceki Veriyi Yükle' 152 160
$btnAdvance = New-UiButton 'Ayı İlerlet (+1 ay)' 322 160
$btnClear   = New-UiButton 'Alanları Temizle' 492 130
$btnClose   = New-UiButton 'Kapat' 632 120

# --- Log ---
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = 'İşlem Günlüğü:'
$lblLog.Location = New-Object System.Drawing.Point(14, 614)
$lblLog.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(12, 636)
$txtLog.Size = New-Object System.Drawing.Size(740, 120)
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtLog)

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    $txtLog.AppendText($line + [Environment]::NewLine)
    [System.Windows.Forms.Application]::DoEvents()
}

# ------------------------------------------------------------------
# Yardımcı işlevler
# ------------------------------------------------------------------
function Set-FormFromEntry {
    param($e)
    if (-not $e) { return }
    $txtSite.Text    = [string]$e.SiteCode
    $txtServer.Text  = [string]$e.ProviderServer
    $txtSugName.Text = [string]$e.SugName
    $txtSugDesc.Text = [string]$e.SugDescription
    $txtPkgName.Text = [string]$e.PkgName
    if ($e.PkgBasePath) { $txtPkgPath.Text = [string]$e.PkgBasePath }
    else                { $txtPkgPath.Text = [string]$e.PkgPath }   # eski kayıt biçimi
    $txtPkgDesc.Text = [string]$e.PkgDescription
}

function Step-MonthInText {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    $pattern = '(?<!\d)(20\d{2})([-_. ]?)(0[1-9]|1[0-2])(?!\d)'
    $evaluator = [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        $d = ([datetime]::new([int]$m.Groups[1].Value, [int]$m.Groups[3].Value, 1)).AddMonths(1)
        '{0:0000}{1}{2:00}' -f $d.Year, $m.Groups[2].Value, $d.Month
    }
    return [regex]::Replace($Text, $pattern, $evaluator)
}

function Import-CMModule {
    if (Get-Module -Name ConfigurationManager) { return }
    if (-not $env:SMS_ADMIN_UI_PATH) {
        throw 'SMS_ADMIN_UI_PATH bulunamadı. Configuration Manager Console bu makinede kurulu olmalı.'
    }
    $consoleBin = $env:SMS_ADMIN_UI_PATH
    $modulePath = Join-Path (Split-Path $consoleBin -Parent) 'ConfigurationManager.psd1'
    if (-not (Test-Path -LiteralPath $modulePath)) {
        throw "ConfigurationManager modülü bulunamadı: $modulePath"
    }
    Import-Module $modulePath -ErrorAction Stop
}

# ------------------------------------------------------------------
# Ana işlem
# ------------------------------------------------------------------
function Invoke-Create {
    $v = [ordered]@{
        Site    = $txtSite.Text.Trim().ToUpper()
        Server  = $txtServer.Text.Trim()
        SugName = $txtSugName.Text.Trim()
        SugDesc = $txtSugDesc.Text.Trim()
        PkgName = $txtPkgName.Text.Trim()
        PkgBase = $txtPkgPath.Text.Trim().TrimEnd('\')
        PkgDesc = $txtPkgDesc.Text.Trim()
    }

    # Doğrulama
    $errs = @()
    if ($v.Site -notmatch '^[A-Z0-9]{3}$')        { $errs += 'Site kodu 3 karakter (harf/rakam) olmalı.' }
    if (-not $v.Server)                            { $errs += 'SMS Provider sunucusu boş olamaz.' }
    if (-not $v.SugName)                           { $errs += 'Software Update Group adı boş olamaz.' }
    if (-not $v.PkgName)                           { $errs += 'Deployment Package adı boş olamaz.' }
    if ($v.PkgBase -notmatch '^\\\\[^\\]+\\[^\\]+') { $errs += 'Ana klasör yolu UNC biçiminde olmalı (\\sunucu\paylaşım\klasör).' }
    if ($v.PkgName -and ($v.PkgName.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $v.PkgName -match '[. ]$')) {
        $errs += 'Package adı klasör adı olarak da kullanılacağı için şu karakterleri içeremez: \ / : * ? " < > |  (ve nokta/boşlukla bitemez).'
    }
    if ($v.PkgDesc.Length -gt 127)                 { $errs += "Package description en fazla 127 karakter olabilir (şu an: $($v.PkgDesc.Length))." }
    if ($errs.Count -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(($errs -join [Environment]::NewLine), 'Eksik / hatalı bilgi',
            'OK', 'Warning') | Out-Null
        return
    }

    # Package adı, ana klasör yoluna eklenir
    $v['PkgPath'] = $v.PkgBase + '\' + $v.PkgName

    $btnCreate.Enabled = $false
    $pushed = $false
    $sugId = $null
    $pkgId = $null
    $anySuccess = $false

    try {
        Write-Log 'Configuration Manager modülü yükleniyor...'
        Import-CMModule

        if (-not (Get-PSDrive -Name $v.Site -PSProvider CMSite -ErrorAction SilentlyContinue)) {
            Write-Log "CMSite sürücüsü oluşturuluyor: $($v.Site) -> $($v.Server)"
            New-PSDrive -Name $v.Site -PSProvider CMSite -Root $v.Server -Description 'SCCM Site' -ErrorAction Stop | Out-Null
        }
        Push-Location "$($v.Site):"
        $pushed = $true
        Write-Log "Site'a bağlanıldı: $($v.Site)"

        # 1) Software Update Group
        Write-Log "Software Update Group kontrol ediliyor: '$($v.SugName)'"
        $existingSug = Get-CMSoftwareUpdateGroup -Name $v.SugName -ErrorAction SilentlyContinue
        if ($existingSug) {
            $sugId = $existingSug.CI_ID
            Write-Log "Bu isimde bir Update Group zaten var (CI_ID: $sugId). Oluşturma atlandı." 'WARN'
        } else {
            $sugParams = @{ Name = $v.SugName; ErrorAction = 'Stop' }
            if ($v.SugDesc) { $sugParams['Description'] = $v.SugDesc }
            $sug = New-CMSoftwareUpdateGroup @sugParams
            $sugId = $sug.CI_ID
            $anySuccess = $true
            Write-Log "Software Update Group oluşturuldu (CI_ID: $sugId)." 'OK'
        }

        # 2) Deployment Package
        Write-Log "Deployment Package kontrol ediliyor: '$($v.PkgName)'"
        $existingPkg = Get-CMSoftwareUpdateDeploymentPackage -Name $v.PkgName -ErrorAction SilentlyContinue
        if ($existingPkg) {
            $pkgId = $existingPkg.PackageID
            Write-Log "Bu isimde bir Deployment Package zaten var (PackageID: $pkgId). Oluşturma atlandı." 'WARN'
        } else {
            # Klasör kontrolü (UNC yol; CMSite sürücüsünden bağımsız çalışır)
            if ([System.IO.Directory]::Exists($v.PkgPath)) {
                $hasItems = @([System.IO.Directory]::EnumerateFileSystemEntries($v.PkgPath)).Count -gt 0
                Write-Log "Klasör zaten mevcut: $($v.PkgPath)"
                if ($hasItems) {
                    Write-Log 'Klasör boş değil. Başka bir paket bu yolu kullanıyorsa SCCM hata verebilir.' 'WARN'
                }
            } else {
                Write-Log "Klasör bulunamadı, oluşturuluyor: $($v.PkgPath)"
                [void][System.IO.Directory]::CreateDirectory($v.PkgPath)
                Write-Log 'Klasör oluşturuldu.' 'OK'
            }

            $pkgParams = @{ Name = $v.PkgName; Path = $v.PkgPath; ErrorAction = 'Stop' }
            if ($v.PkgDesc) { $pkgParams['Description'] = $v.PkgDesc }
            $pkg = New-CMSoftwareUpdateDeploymentPackage @pkgParams
            $pkgId = $pkg.PackageID
            $anySuccess = $true
            Write-Log "Deployment Package oluşturuldu (PackageID: $pkgId)." 'OK'
        }

        # Durumu kaydet
        $entry = [ordered]@{
            Timestamp      = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            SiteCode       = $v.Site
            ProviderServer = $v.Server
            SugName        = $v.SugName
            SugDescription = $v.SugDesc
            SugId          = $sugId
            PkgName        = $v.PkgName
            PkgDescription = $v.PkgDesc
            PkgBasePath    = $v.PkgBase
            PkgPath        = $v.PkgPath
            PkgId          = $pkgId
            CreatedAnything = $anySuccess
        }
        $saved = Save-Entry -Entry $entry
        $txtPrev.Text = Format-Entry ([pscustomobject]$entry)
        if ($saved) { Write-Log "Bilgiler kaydedildi: $script:StateFile" }
        else        { Write-Log 'Bulut senkronizasyonu riski nedeniyle durum dosyası kaydedilmedi.' 'WARN' }

        [System.Windows.Forms.MessageBox]::Show('İşlem tamamlandı.', 'Bilgi', 'OK', 'Information') | Out-Null
    }
    catch {
        Write-Log $_.Exception.Message 'ERROR'
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Hata', 'OK', 'Error') | Out-Null
    }
    finally {
        if ($pushed) { Pop-Location }
        $btnCreate.Enabled = $true
    }
}

# ------------------------------------------------------------------
# Olaylar
# ------------------------------------------------------------------
$btnCreate.Add_Click({ Invoke-Create })

$btnLoad.Add_Click({
    $s = Get-State
    if ($s -and $s.LastRun) {
        Set-FormFromEntry $s.LastRun
        Write-Log 'Önceki çalıştırma verileri forma yüklendi.'
    } else {
        [System.Windows.Forms.MessageBox]::Show('Kayıtlı önceki çalıştırma bulunamadı.', 'Bilgi', 'OK', 'Information') | Out-Null
    }
})

$btnAdvance.Add_Click({
    $txtSugName.Text = Step-MonthInText $txtSugName.Text
    $txtSugDesc.Text = Step-MonthInText $txtSugDesc.Text
    $txtPkgName.Text = Step-MonthInText $txtPkgName.Text
    $txtPkgPath.Text = Step-MonthInText $txtPkgPath.Text
    $txtPkgDesc.Text = Step-MonthInText $txtPkgDesc.Text
    Write-Log 'Ay bilgileri +1 ay ilerletildi (yyyy-MM benzeri kalıplar).'
})

$btnClear.Add_Click({
    foreach ($t in @($txtSugName, $txtSugDesc, $txtPkgName, $txtPkgPath, $txtPkgDesc)) { $t.Clear() }
})

$btnClose.Add_Click({ $form.Close() })

function Update-PathPreview {
    $base = $txtPkgPath.Text.Trim().TrimEnd('\')
    $name = $txtPkgName.Text.Trim()
    if ($base -and $name) { $lblFullPath.Text = $base + '\' + $name }
    else                  { $lblFullPath.Text = '(ana klasör yolu ve package adı girildiğinde görünür)' }
}
$txtPkgPath.Add_TextChanged({ Update-PathPreview })
$txtPkgName.Add_TextChanged({ Update-PathPreview })

# ------------------------------------------------------------------
# Açılışta önceki çalıştırmayı getir
# ------------------------------------------------------------------
if ($script:StateWarning) { Write-Log $script:StateWarning 'WARN' }
$state = Get-State
if ($state -and $state.LastRun) {
    $txtPrev.Text = Format-Entry $state.LastRun
    Set-FormFromEntry $state.LastRun
    Write-Log "Önceki çalıştırma yüklendi ($($state.LastRun.Timestamp))."
} else {
    $txtPrev.Text = Format-Entry $null
}

Update-PathPreview
[void]$form.ShowDialog()