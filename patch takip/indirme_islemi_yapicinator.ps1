<#
.SYNOPSIS
    SCCM'de "Required" (istemcilerde eksik) olan güncellemeleri listeler ve Software Update Group (SUG)
    üyeliklerini düzenler. patch_takip.ps1 (superseded aracı) ile aynı mantıkta çalışır.

.DESCRIPTION
    Listelenen güncellemeler şu filtreyle gelir (SMS_SoftwareUpdate):
        Vendor      = Microsoft  (metin kutusundan değiştirilebilir)
        Required    > 1          (Required = NumMissing; sayı kutusundan değiştirilebilir)
        Superseded  = Hayır      (sabit)
    Ek seçenekler:
        - Expired olanları hariç tut (varsayılan: açık)
        - Başlık / KB / kategori içinde arama
        - Yalnızca hiçbir SUG'a üye olmayanları göster
    Office 365 / Microsoft 365 Apps güncellemeleri de bu filtrenin içindedir (ürün filtresi yoktur);
    "Kategori" sütununda ürün adı görünür. SUP'ta "Office 365 Client" ürünü seçili ve senkronize olmalıdır.

    Sağ tarafta ortamdaki TÜM Software Update Group'lar listelenir (SCCM'deki "Edit Membership" gibi).
    Seçili güncellemeler için kutular şöyle çalışır:
        işaretli   : seçili güncellemelerin hepsi grubun üyesi (işaretlerseniz eksikler EKLENİR)
        boş        : hiçbiri üye değil (boşaltırsanız üye olanlar ÇIKARILIR)
        dolu kare  : bir kısmı üye; bu durumda değişiklik yapılmaz
    Bir SUG en fazla 1000 güncelleme içerebilir; sınır aşılacaksa işlem başlamadan uyarılır.

    ALT KISIMDAKİ "İndirme" bölümü, soldaki listeden seçilen güncellemelerin içeriğini
    ORTAMDA VAR OLAN bir Deployment Package'a indirir (yeni paket oluşturmaz). Konsoldaki
    "Download" sihirbazının PowerShell karşılığı olan Save-CMSoftwareUpdate cmdlet'i kullanılır.
    İndirme konumu sabittir: İnternet (Microsoft Update) — -Location parametresi hiç
    verilmez, çünkü bu cmdlet'te -Location boş bırakıldığında içerik internetten indirilir.
    Diller, "Windows Update" ve "Office 365 Client Update" için ayrı ayrı (English/Turkish)
    işaretlenir; ancak Save-CMSoftwareUpdate tek bir dil listesini seçilen TÜM güncellemelere
    birlikte uygular (ürüne göre ayrı dil seti göndermez) — bu SCCM'in cmdlet sınırlamasıdır.

    "Simülasyon" kutusu işaretliyken ne SUG üyeliği değişikliği ne de indirme yapılır;
    her iki işlem için de sadece ne yapılacağı günlüğe yazılır.

.NOTES
    Configuration Manager Console'un kurulu olduğu bir makinede çalıştırın.
    Okuma işlemleri SMS Provider WMI (root\SMS\site_XXX) üzerinden yapılır (DCOM/RPC erişimi gerekir).
    SUG üyeliği değişiklikleri ConfigurationManager modülü ile yapılır; SUG düzenleme yetkisi gerekir.
    Site kodu / SMS Provider alanları, varsa state.json'un son kaydından otomatik doldurulur (sadece okunur).
    İndirme (download) işlemi ConfigurationManager modülündeki Save-CMSoftwareUpdate cmdlet'i ile yapılır;
    var olan bir Deployment Package'a yazma/indirme yetkisi gerekir. Bu araç YENİ Deployment Package
    OLUŞTURMAZ; yalnızca ortamda zaten var olan bir paketi hedef olarak seçer.

    SUG ÜYELİK EKLEME/ÇIKARMA MEKANİZMASI:
    Add-CMSoftwareUpdateToGroup cmdlet'i, içeriği henüz indirilmemiş bir güncellemeyi
    DEPLOY EDİLMİŞ bir SUG'a eklerken SMS Provider seviyesinde reddedilir:
        "Non-downloaded software update <update-name> can't be added to a software
         update group that is deployed."
    Configuration Manager konsolundaki "Edit Membership" ise SUG'un üye listesine
    (Updates dizisi) doğrudan yazdığı için bu kısıtlamaya takılmaz. Bu script bu
    yüzden önce Set-CMSoftwareUpdateGroup (-AddSoftwareUpdate/-RemoveSoftwareUpdate)
    kullanır; o da aynı hatayı verirse konsolun yaptığının birebir aynısı olan
    doğrudan bir WMI yazımına (SMS_AuthorizationList.Updates) düşer. Böylece,
    içerik indirilmemiş güncellemeler de -tıpkı konsolda olduğu gibi- deploy
    edilmiş bir SUG'a eklenebilir. Content indirme/dağıtım bu scriptin kapsamı
    dışındadır; devam eden deployment'ın client'lara başarıyla ulaşması için
    içeriğin ayrıca (konsoldan veya ADR ile) bir Deployment Package'a indirilip
    distribution point'lere dağıtılması gerekir.
#>

#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ------------------------------------------------------------------
# Genel durum
# ------------------------------------------------------------------
$script:Groups       = @()      # ortamdaki tüm SUG'lar
$script:GroupById    = @{}      # SUG CI_ID (string) -> SUG
$script:Updates      = @{}      # güncelleme CI_ID (string) -> kayıt
$script:Members      = @{}      # güncelleme CI_ID (string) -> HashSet[string] (üye olduğu SUG CI_ID'leri)
$script:OrigCheck    = @{}      # SUG CI_ID (string) -> kutunun ilk hesaplanan durumu
$script:Suspend      = $false   # UI olaylarını geçici olarak susturmak için
$script:SugLimit     = 1000     # bir SUG'daki en fazla güncelleme sayısı
$script:Packages     = @()      # ortamdaki tüm Deployment Package'lar (indirme hedefi olabilecek)
$script:PackageById  = @{}      # Package ID (string) -> paket kaydı
$script:PkgDisplayToId = @{}    # kombo kutusundaki görünen metin -> Package ID
$script:DownloadJobs   = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
                                 # arkaplanda çalışan/çalışmış indirme işleri (her biri senkronize bir hashtable)

# ------------------------------------------------------------------
# state.json'dan yalnızca site kodu / provider okuma (SADECE okuma)
# ------------------------------------------------------------------
function Get-StateFileCandidates {
    $paths = @(Join-Path $env:USERPROFILE 'Desktop\update-state\state.json')
    if ($env:LOCALAPPDATA) { $paths += (Join-Path $env:LOCALAPPDATA 'update-state\state.json') }
    try {
        $desk = [Environment]::GetFolderPath('Desktop')
        if ($desk) { $paths += (Join-Path $desk 'update-state\state.json') }
    } catch { }
    $unique = @()
    foreach ($p in $paths) { if ($p -and ($unique -notcontains $p)) { $unique += $p } }
    return $unique
}

function Find-StateFile {
    $found = @()
    foreach ($p in (Get-StateFileCandidates)) {
        if (Test-Path -LiteralPath $p) { $found += (Get-Item -LiteralPath $p) }
    }
    if ($found.Count -eq 0) { return $null }
    return ($found | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}

# ------------------------------------------------------------------
# Saf (UI'dan bağımsız) mantık fonksiyonları
# ------------------------------------------------------------------
function ConvertTo-CheckState {
    param($Value)
    if ($Value -is [System.Windows.Forms.CheckState]) { return $Value }
    if ($Value -is [bool] -and $Value) { return [System.Windows.Forms.CheckState]::Checked }
    return [System.Windows.Forms.CheckState]::Unchecked
}

# Kutuların ilk durumu ile şimdiki durumunu karşılaştırıp hangi SUG'a hangi güncellemelerin
# ekleneceğini / çıkarılacağını hesaplar. "Dolu kare" (Indeterminate) = dokunma.
function Get-MembershipPlan {
    param(
        [string[]]$SelectedIds,
        [hashtable]$Members,
        [hashtable]$Orig,
        [hashtable]$Current
    )
    $plan = @()
    foreach ($gid in @($Current.Keys)) {
        $curState  = $Current[$gid]
        $origState = $Orig[$gid]
        if ($curState -eq $origState) { continue }
        if ($curState -eq [System.Windows.Forms.CheckState]::Indeterminate) { continue }

        $add = @()
        $rem = @()
        foreach ($uid in $SelectedIds) {
            $isMember = $Members.ContainsKey($uid) -and $Members[$uid].Contains($gid)
            if ($curState -eq [System.Windows.Forms.CheckState]::Checked) {
                if (-not $isMember) { $add += $uid }
            } else {
                if ($isMember) { $rem += $uid }
            }
        }
        if ($add.Count -gt 0 -or $rem.Count -gt 0) {
            $plan += [pscustomobject]@{
                GroupId = [string]$gid
                Add     = [string[]]$add
                Remove  = [string[]]$rem
            }
        }
    }
    return $plan
}

function ConvertTo-DateText {
    param($Value)
    if (-not $Value) { return '' }
    try {
        return ([System.Management.ManagementDateTimeConverter]::ToDateTime([string]$Value)).ToString('yyyy-MM-dd')
    } catch {
        return [string]$Value
    }
}

# ------------------------------------------------------------------
# UI yardımcıları
# ------------------------------------------------------------------
function New-UiLabel {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 130, [int]$H = 20)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.Location = New-Object System.Drawing.Point($X, $Y)
    $l.Size = New-Object System.Drawing.Size($W, $H)
    $Parent.Controls.Add($l)
    return $l
}

function New-UiTextBox {
    param($Parent, [int]$X, [int]$Y, [int]$W, [int]$H = 23)
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location = New-Object System.Drawing.Point($X, $Y)
    $t.Size = New-Object System.Drawing.Size($W, $H)
    $Parent.Controls.Add($t)
    return $t
}

function New-UiGroup {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H)
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = $Text
    $g.Location = New-Object System.Drawing.Point($X, $Y)
    $g.Size = New-Object System.Drawing.Size($W, $H)
    $Parent.Controls.Add($g)
    return $g
}

function New-UiButton {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H = 28)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.Location = New-Object System.Drawing.Point($X, $Y)
    $b.Size = New-Object System.Drawing.Size($W, $H)
    $Parent.Controls.Add($b)
    return $b
}

function New-UiGrid {
    param($Parent, [int]$X, [int]$Y, [int]$W, [int]$H)
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Location = New-Object System.Drawing.Point($X, $Y)
    $g.Size = New-Object System.Drawing.Size($W, $H)
    $g.Anchor = 'Top,Bottom,Left,Right'
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.AllowUserToResizeRows = $false
    $g.RowHeadersVisible = $false
    $g.BackgroundColor = [System.Drawing.Color]::White
    $g.SelectionMode = 'FullRowSelect'
    $Parent.Controls.Add($g)
    return $g
}

function Add-UiTextColumn {
    param($Grid, [string]$Name, [string]$Header, [int]$Width, [bool]$Fill = $false, [float]$Weight = 100)
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = $Name
    $c.HeaderText = $Header
    $c.ReadOnly = $true
    if ($Fill) {
        $c.AutoSizeMode = 'Fill'
        $c.MinimumWidth = $Width
        $c.FillWeight = $Weight
    } else {
        $c.Width = $Width
    }
    [void]$Grid.Columns.Add($c)
}

# Çift tıklayınca sütunu içeriğine göre genişletir, tekrar çift tıklayınca eski genişliğe döndürür.
# (Başlığa veya herhangi bir hücreye çift tıklanabilir.)
function Enable-ColumnDoubleClickExpand {
    param($Grid)
    $toggle = {
        param($sender, $e)
        if ($e.ColumnIndex -lt 0) { return }
        $col = $sender.Columns[$e.ColumnIndex]
        if ($col.Tag -is [hashtable] -and $col.Tag.Expanded) {
            $col.Width = [int]$col.Tag.Orig
            $col.Tag = $null
        } else {
            $orig = $col.Width
            $sender.AutoResizeColumn($e.ColumnIndex, [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells)
            $col.Tag = @{ Orig = $orig; Expanded = $true }
            # Genişleyen sütun kısmen ekran dışında kaldıysa görünür alana kaydır
            try {
                $rect = $sender.GetColumnDisplayRectangle($e.ColumnIndex, $false)
                if ($rect.Width -lt $col.Width -and $col.Frozen -eq $false) {
                    $sender.FirstDisplayedScrollingColumnIndex = $e.ColumnIndex
                }
            } catch { }
        }
    }
    $Grid.Add_ColumnHeaderMouseDoubleClick($toggle)
    $Grid.Add_CellDoubleClick($toggle)
}

# ------------------------------------------------------------------
# Form
# ------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'SCCM - Required Güncellemeler & SUG Üyelik Yönetimi'
$form.ClientSize = New-Object System.Drawing.Size(1280, 1214)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 1114)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# --- Bağlantı ve Filtre ---
$gTop = New-UiGroup $form 'Bağlantı ve Filtre' 10 8 1260 124
$gTop.Anchor = 'Top,Left,Right'

[void](New-UiLabel $gTop 'Site Kodu:' 12 27 70)
$txtSite = New-UiTextBox $gTop 85 24 70
[void](New-UiLabel $gTop 'SMS Provider:' 170 27 85)
$txtServer = New-UiTextBox $gTop 258 24 230

[void](New-UiLabel $gTop 'Vendor:' 12 62 70)
$txtVendor = New-UiTextBox $gTop 85 58 120
$txtVendor.Text = 'Microsoft'
[void](New-UiLabel $gTop 'Required >' 225 62 70)
$nudReq = New-Object System.Windows.Forms.NumericUpDown
$nudReq.Location = New-Object System.Drawing.Point(298, 58)
$nudReq.Size = New-Object System.Drawing.Size(80, 23)
$nudReq.Minimum = 0
$nudReq.Maximum = 1000000
$nudReq.Value = 0
$gTop.Controls.Add($nudReq)
[void](New-UiLabel $gTop 'Ara (başlık/KB/kategori):' 400 62 150)
$txtTitle = New-UiTextBox $gTop 555 58 250
[void](New-UiLabel $gTop 'Superseded: Hayır (sabit)' 830 62 180)

$chkExpired = New-Object System.Windows.Forms.CheckBox
$chkExpired.Text = 'Expired olanları hariç tut'
$chkExpired.Location = New-Object System.Drawing.Point(12, 92)
$chkExpired.Size = New-Object System.Drawing.Size(200, 22)
$chkExpired.Checked = $true
$gTop.Controls.Add($chkExpired)
$chkNoSug = New-Object System.Windows.Forms.CheckBox
$chkNoSug.Text = 'Yalnızca hiçbir SUG''a üye olmayanlar'
$chkNoSug.Location = New-Object System.Drawing.Point(225, 92)
$chkNoSug.Size = New-Object System.Drawing.Size(300, 22)
$gTop.Controls.Add($chkNoSug)
$btnList = New-UiButton $gTop 'Listele' 1100 86 149 30
$btnList.Anchor = 'Top,Right'

# --- Required güncellemeler ---
$gUpd = New-UiGroup $form 'Required güncellemeler (filtreye uyan)' 10 140 800 500
$gUpd.Anchor = 'Top,Bottom,Left,Right'
$gridUpd = New-UiGrid $gUpd 10 22 780 430
$gridUpd.MultiSelect = $true
$gridUpd.ReadOnly = $true
$gridUpd.AutoSizeColumnsMode = 'None'          # Fill kapalı: sütunlar sabit genişlikte, toplam sığmazsa X scroll çıkar
$gridUpd.ScrollBars = 'Both'
$gridUpd.AllowUserToResizeColumns = $true
Add-UiTextColumn $gridUpd 'uKB' 'KB' 80
Add-UiTextColumn $gridUpd 'uTitle' 'Başlık' 360
Add-UiTextColumn $gridUpd 'uReq' 'Required' 70
Add-UiTextColumn $gridUpd 'uPosted' 'Yayın' 85
Add-UiTextColumn $gridUpd 'uCat' 'Kategori' 150
Add-UiTextColumn $gridUpd 'uGroups' 'Üye olduğu SUG''lar' 260
Add-UiTextColumn $gridUpd 'uId' 'CI_ID' 75
Enable-ColumnDoubleClickExpand $gridUpd

$btnSelAll = New-UiButton $gUpd 'Tümünü Seç' 10 462 100 26
$btnSelAll.Anchor = 'Bottom,Left'
$btnSelNone = New-UiButton $gUpd 'Seçimi Temizle' 115 462 110 26
$btnSelNone.Anchor = 'Bottom,Left'
$lblCount = New-UiLabel $gUpd '' 235 466 545
$lblCount.Anchor = 'Bottom,Left'

# --- Tüm SUG'lar ---
$gGrp = New-UiGroup $form 'Software Update Grupları (ortamdaki tümü)' 820 140 450 500
$gGrp.Anchor = 'Top,Bottom,Right'   # sabit genişlik (450), pencere büyüyünce sağa kayar
[void](New-UiLabel $gGrp 'Grup ara:' 10 26 60)
$txtFilter = New-UiTextBox $gGrp 72 23 368
$txtFilter.Anchor = 'Top,Left,Right'
$gridGrp = New-UiGrid $gGrp 10 52 430 400
$gridGrp.MultiSelect = $false
$colChk = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colChk.Name = 'gChk'
$colChk.HeaderText = ''
$colChk.Width = 36
$colChk.ThreeState = $true
[void]$gridGrp.Columns.Add($colChk)
Add-UiTextColumn $gridGrp 'gName' 'Grup adı' 150 $true 100
Add-UiTextColumn $gridGrp 'gSel' 'Seçili' 55
Add-UiTextColumn $gridGrp 'gTotal' 'Toplam' 60
$lblLegend = New-UiLabel $gGrp 'İşaretli: ekle  |  Boş: çıkar  |  Dolu kare: dokunma.  Güncelleme seçimi değişirse kutular yeniden hesaplanır. SUG başına en fazla 1000 güncelleme.' 10 458 430 36
$lblLegend.Anchor = 'Bottom,Left,Right'
$lblLegend.ForeColor = [System.Drawing.Color]::DimGray

# --- İndirme (var olan Deployment Package'a) ---
$gDl = New-UiGroup $form 'İndirme (Seçili güncellemeleri var olan bir Deployment Package''a indir)' 10 646 1260 140
$gDl.Anchor = 'Bottom,Left,Right'

[void](New-UiLabel $gDl 'Deployment Package:' 12 26 130)
$cmbPkg = New-Object System.Windows.Forms.ComboBox
$cmbPkg.Location = New-Object System.Drawing.Point(158, 23)
$cmbPkg.Size = New-Object System.Drawing.Size(430, 23)
$cmbPkg.DropDownStyle = 'DropDown'
$cmbPkg.Anchor = 'Top,Left,Right'
$gDl.Controls.Add($cmbPkg)
$btnPkgRefresh = New-UiButton $gDl 'Paketleri Yenile' 600 22 140 25
$btnPkgRefresh.Anchor = 'Top,Right'
$lblPkgCount = New-UiLabel $gDl '' 750 27 490
$lblPkgCount.Anchor = 'Top,Right'
$lblPkgCount.ForeColor = [System.Drawing.Color]::DimGray

[void](New-UiLabel $gDl 'Güncelleme Dilleri:' 12 62 130)

$gLangWin = New-UiGroup $gDl 'Windows Update' 158 50 220 50
$chkLangWinEn = New-Object System.Windows.Forms.CheckBox
$chkLangWinEn.Text = 'English'
$chkLangWinEn.Location = New-Object System.Drawing.Point(10, 20)
$chkLangWinEn.Size = New-Object System.Drawing.Size(90, 22)
$chkLangWinEn.Checked = $true
$gLangWin.Controls.Add($chkLangWinEn)
$chkLangWinTr = New-Object System.Windows.Forms.CheckBox
$chkLangWinTr.Text = 'Turkish'
$chkLangWinTr.Location = New-Object System.Drawing.Point(112, 20)
$chkLangWinTr.Size = New-Object System.Drawing.Size(90, 22)
$chkLangWinTr.Checked = $true
$gLangWin.Controls.Add($chkLangWinTr)

$gLangOff = New-UiGroup $gDl 'Office 365 Client Update' 390 50 260 50
$chkLangOffEn = New-Object System.Windows.Forms.CheckBox
$chkLangOffEn.Text = 'English'
$chkLangOffEn.Location = New-Object System.Drawing.Point(10, 20)
$chkLangOffEn.Size = New-Object System.Drawing.Size(90, 22)
$chkLangOffEn.Checked = $true
$gLangOff.Controls.Add($chkLangOffEn)
$chkLangOffTr = New-Object System.Windows.Forms.CheckBox
$chkLangOffTr.Text = 'Turkish'
$chkLangOffTr.Location = New-Object System.Drawing.Point(112, 20)
$chkLangOffTr.Size = New-Object System.Drawing.Size(90, 22)
$chkLangOffTr.Checked = $true
$gLangOff.Controls.Add($chkLangOffTr)

$lblDlNote = New-UiLabel $gDl 'İndirme konumu: İnternet (Microsoft Update) - sabit. Diller, seçilen tüm güncellemelere birlikte uygulanır (SCCM cmdlet sınırlaması). İndirme arkaplanda çalışır; bu düğmeye farklı paket/seçimle tekrar basarak aynı anda birden fazla indirme başlatabilirsiniz - aşağıdaki "Aktif İndirmeler" listesinden takip edin.' 660 56 590 46
$lblDlNote.ForeColor = [System.Drawing.Color]::DimGray
$lblDlNote.Anchor = 'Top,Right'

$btnDownload = New-UiButton $gDl 'Seçili Güncellemeleri İndir (Arkaplanda)' 950 106 300 30
$btnDownload.Anchor = 'Top,Right'

#region --- Aktif İndirmeler (arkaplan indirme ilerlemesi) ---
$gDlStatus = New-UiGroup $form 'Aktif İndirmeler (arkaplanda çalışır - birden fazla indirme aynı anda tetiklenebilir)' 10 792 1260 160
$gDlStatus.Anchor = 'Top,Bottom,Left,Right'
$gridDl = New-UiGrid $gDlStatus 10 22 1240 100
$gridDl.MultiSelect = $false
$gridDl.ReadOnly = $false
$gridDl.EditMode = 'EditProgrammatically'
$gridDl.Anchor = 'Top,Bottom,Left,Right'
Add-UiTextColumn $gridDl 'dPkg' 'Deployment Package' 220 $true 30
Add-UiTextColumn $gridDl 'dProg' 'İlerleme (indi/toplam)' 130
Add-UiTextColumn $gridDl 'dRemain' 'Kalan' 60
Add-UiTextColumn $gridDl 'dPct' 'Yüzde' 60
Add-UiTextColumn $gridDl 'dCurrent' 'Şu an indirilen' 220 $true 40
Add-UiTextColumn $gridDl 'dStatus' 'Durum' 110
Add-UiTextColumn $gridDl 'dSure' 'Süre' 70
$colDlCancel = New-Object System.Windows.Forms.DataGridViewButtonColumn
$colDlCancel.Name = 'dCancel'
$colDlCancel.HeaderText = ''
$colDlCancel.Width = 70
$colDlCancel.UseColumnTextForButtonValue = $false
[void]$gridDl.Columns.Add($colDlCancel)

$btnDlClear = New-UiButton $gDlStatus 'Tamamlananları Listeden Temizle' 10 128 240 26
$btnDlClear.Anchor = 'Bottom,Left'
$lblDlStatusNote = New-UiLabel $gDlStatus 'Her indirme kendi arkaplan thread''inde (runspace) çalışır; farklı paket/seçimle "Seçili Güncellemeleri İndir" tekrar tekrar basılarak aynı anda birden fazla indirme başlatılabilir. "İptal", devam eden öğe bittikten sonra işi durdurur.' 260 130 990 26
$lblDlStatusNote.Anchor = 'Bottom,Left,Right'
$lblDlStatusNote.ForeColor = [System.Drawing.Color]::DimGray
#endregion

# Arkaplan indirme işlerinin ilerlemesini periyodik olarak grid'e yansıtan zamanlayıcı.
# Yalnızca aktif iş varken çalışır (Start-DownloadJob başlatır, Update-DownloadGrid durdurur).
$tmrProgress = New-Object System.Windows.Forms.Timer
$tmrProgress.Interval = 600
$tmrProgress.Add_Tick({ Update-DownloadGrid })
# Not: System.Windows.Forms.Timer bir Control değildir; $form.Controls.Add ile eklenmez.
# $script: kapsamındaki değişken olarak yaşamaya devam eder, Start()/Stop() yeterlidir.

# --- İşlemler ---
$gAct = New-UiGroup $form 'İşlemler' 10 958 1260 62
$gAct.Anchor = 'Bottom,Left,Right'
$chkSim = New-Object System.Windows.Forms.CheckBox
$chkSim.Text = 'Simülasyon (SUG üyeliği / indirme yapma, sadece günlüğe yaz)'
$chkSim.Location = New-Object System.Drawing.Point(14, 26)
$chkSim.Size = New-Object System.Drawing.Size(400, 22)
$chkSim.Checked = $true
$gAct.Controls.Add($chkSim)
$btnApply = New-UiButton $gAct 'SUG Üyeliklerini Uygula' 430 20 210 32
$btnClose = New-UiButton $gAct 'Kapat' 1140 20 105 32
$btnClose.Anchor = 'Top,Right'

# --- Günlük ---
$lblLog = New-UiLabel $form 'İşlem Günlüğü:' 14 1026 200
$lblLog.Anchor = 'Bottom,Left'
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 1048)
$txtLog.Size = New-Object System.Drawing.Size(1260, 148)
$txtLog.Anchor = 'Bottom,Left,Right'
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

function Set-Busy {
    param([bool]$On)
    foreach ($c in @($btnList, $btnApply, $btnPkgRefresh)) { $c.Enabled = (-not $On) }
    if ($On) { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor }
    else     { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    [System.Windows.Forms.Application]::DoEvents()
}

function Show-Message {
    param([string]$Text, [string]$Title = 'Bilgi', [string]$Icon = 'Information')
    [void][System.Windows.Forms.MessageBox]::Show($Text, $Title, 'OK', $Icon)
}

# state.json varsa yalnızca Site kodu ve SMS Provider alanlarını doldurur.
function Import-ConnectionDefaults {
    $path = Find-StateFile
    if (-not $path) { return }
    try {
        $state = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
        $raw = $state.LastRun
        if (-not $raw -and $state.History) { $raw = @($state.History)[-1] }
        if ($raw) {
            if ($raw.SiteCode)       { $txtSite.Text   = [string]$raw.SiteCode }
            if ($raw.ProviderServer) { $txtServer.Text = [string]$raw.ProviderServer }
            Write-Log "Bağlantı bilgileri state.json'dan dolduruldu: $path"
        }
    } catch {
        Write-Log "state.json okunamadı (yoksayıldı): $($_.Exception.Message)" 'WARN'
    }
}

# ------------------------------------------------------------------
# SCCM erişimi
# ------------------------------------------------------------------
function Get-CMModulePath {
    if (-not $env:SMS_ADMIN_UI_PATH) {
        throw 'SMS_ADMIN_UI_PATH bulunamadı. Configuration Manager Console bu makinede kurulu olmalı.'
    }
    $consoleBin = $env:SMS_ADMIN_UI_PATH
    $modulePath = Join-Path (Split-Path $consoleBin -Parent) 'ConfigurationManager.psd1'
    if (-not (Test-Path -LiteralPath $modulePath)) {
        throw "ConfigurationManager modülü bulunamadı: $modulePath"
    }
    return $modulePath
}

function Import-CMModule {
    if (Get-Module -Name ConfigurationManager) { return }
    Import-Module (Get-CMModulePath) -ErrorAction Stop
}

function Get-Connection {
    $site = $txtSite.Text.Trim().ToUpper()
    $server = $txtServer.Text.Trim()
    if ($site -notmatch '^[A-Z0-9]{3}$') { throw 'Site kodu 3 karakter (harf/rakam) olmalı.' }
    if (-not $server) { throw 'SMS Provider sunucusu boş olamaz.' }
    return [pscustomobject]@{ Site = $site; Server = $server }
}

function Invoke-CMWql {
    param([string]$Query)
    $c = Get-Connection
    return @(Get-WmiObject -ComputerName $c.Server -Namespace "root\SMS\site_$($c.Site)" -Query $Query -ErrorAction Stop)
}

function Enter-CMSite {
    $c = Get-Connection
    Import-CMModule
    if (-not (Get-PSDrive -Name $c.Site -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        Write-Log "CMSite sürücüsü oluşturuluyor: $($c.Site) -> $($c.Server)"
        New-PSDrive -Name $c.Site -PSProvider CMSite -Root $c.Server -Description 'SCCM Site' -ErrorAction Stop | Out-Null
    }
    Push-Location "$($c.Site):"
}

# ------------------------------------------------------------------
# SUG üyeliği - düşük seviye uygulama
# ------------------------------------------------------------------
# Add-CMSoftwareUpdateToGroup cmdlet'i, deploy edilmiş bir SUG'a içeriği
# (content) henüz indirilmemiş güncelleme eklenmesini SMS Provider seviyesinde
# reddeder ("non-downloaded software update ... can't be added to a software
# update group that is deployed"). Konsoldaki "Edit Membership" ise SUG'un
# üye listesine (Updates dizisi) doğrudan yazarak aynı kısıtlamaya takılmaz.
# Bu yüzden burada önce Set-CMSoftwareUpdateGroup (-AddSoftwareUpdate /
# -RemoveSoftwareUpdate) deneniyor; o da aynı hatayı verirse konsolun
# yaptığına birebir eşdeğer bir WMI yazımına (SMS_AuthorizationList.Updates)
# düşülüyor. İçerik indirme burada YAPILMAZ; sadece üyelik değiştirilir.
function Add-UpdatesToGroupWmi {
    param([string]$GroupId, [string[]]$UpdateIds)
    $c = Get-Connection
    $sugWmi = Get-WmiObject -ComputerName $c.Server -Namespace "root\SMS\site_$($c.Site)" `
        -Class SMS_AuthorizationList -Filter "CI_ID = $GroupId" -ErrorAction Stop
    if (-not $sugWmi) { throw "WMI: SUG bulunamadı (CI_ID=$GroupId)" }
    $current = @($sugWmi.Updates | ForEach-Object { [int]$_ })
    $toAdd   = @($UpdateIds | ForEach-Object { [int]$_ } | Where-Object { $current -notcontains $_ })
    if ($toAdd.Count -eq 0) { return }
    $sugWmi.Updates = [int[]]($current + $toAdd)
    [void]$sugWmi.Put()
}

function Remove-UpdatesFromGroupWmi {
    param([string]$GroupId, [string[]]$UpdateIds)
    $c = Get-Connection
    $sugWmi = Get-WmiObject -ComputerName $c.Server -Namespace "root\SMS\site_$($c.Site)" `
        -Class SMS_AuthorizationList -Filter "CI_ID = $GroupId" -ErrorAction Stop
    if (-not $sugWmi) { throw "WMI: SUG bulunamadı (CI_ID=$GroupId)" }
    $remove  = @($UpdateIds | ForEach-Object { [int]$_ })
    $current = @($sugWmi.Updates | ForEach-Object { [int]$_ } | Where-Object { $remove -notcontains $_ })
    $sugWmi.Updates = [int[]]$current
    [void]$sugWmi.Put()
}

# Bir SUG'a güncelleme ekler: önce resmi cmdlet (Set-CMSoftwareUpdateGroup),
# başarısız olursa (deploy edilmiş grup + indirilmemiş içerik hatası) doğrudan
# WMI'a düşer. $sug: Get-CMSoftwareUpdateGroup çıktısı. $UpdateIds: CI_ID dizisi.
function Add-UpdatesToGroup {
    param($Sug, [string]$GroupId, [string]$GroupName, [string[]]$UpdateIds)
    try {
        $objs = @($UpdateIds | ForEach-Object { Get-CMSoftwareUpdate -Id $_ -Fast -ErrorAction Stop })
        Set-CMSoftwareUpdateGroup -InputObject $Sug -AddSoftwareUpdate $objs -ErrorAction Stop
    } catch {
        Write-Log "'$GroupName': Set-CMSoftwareUpdateGroup ile eklenemedi ($($_.Exception.Message)); doğrudan WMI ile deneniyor." 'WARN'
        Add-UpdatesToGroupWmi -GroupId $GroupId -UpdateIds $UpdateIds
    }
}

function Remove-UpdatesFromGroup {
    param($Sug, [string]$GroupId, [string]$GroupName, [string[]]$UpdateIds)
    try {
        $objs = @($UpdateIds | ForEach-Object { Get-CMSoftwareUpdate -Id $_ -Fast -ErrorAction Stop })
        Set-CMSoftwareUpdateGroup -InputObject $Sug -RemoveSoftwareUpdate $objs -ErrorAction Stop
    } catch {
        Write-Log "'$GroupName': Set-CMSoftwareUpdateGroup ile çıkarılamadı ($($_.Exception.Message)); doğrudan WMI ile deneniyor." 'WARN'
        Remove-UpdatesFromGroupWmi -GroupId $GroupId -UpdateIds $UpdateIds
    }
}

# ------------------------------------------------------------------
# Veri yükleme
# ------------------------------------------------------------------
function Import-Groups {
    $rows = Invoke-CMWql 'SELECT * FROM SMS_AuthorizationList'
    $script:Groups = @()
    $script:GroupById = @{}
    foreach ($r in ($rows | Sort-Object { [string]$_.LocalizedDisplayName })) {
        $g = [pscustomobject]@{
            Id    = [string]$r.CI_ID
            Name  = [string]$r.LocalizedDisplayName
            Total = [int]$r.NumberOfUpdates
        }
        $script:Groups += $g
        $script:GroupById[$g.Id] = $g
    }
}

# Ortamda VAR OLAN Deployment Package'ları okur (indirme hedefi olarak seçilecekler).
# Bu araç yeni Deployment Package oluşturmaz; yalnızca burada listelenenlerden seçim yapılır.
function Import-Packages {
    $rows = Invoke-CMWql 'SELECT PackageID, Name, PkgSourcePath FROM SMS_SoftwareUpdatesPackage'
    $script:Packages = @()
    $script:PackageById = @{}
    foreach ($r in ($rows | Sort-Object { [string]$_.Name })) {
        $p = [pscustomobject]@{
            Id         = [string]$r.PackageID
            Name       = [string]$r.Name
            SourcePath = [string]$r.PkgSourcePath
        }
        $script:Packages += $p
        $script:PackageById[$p.Id] = $p
    }
}

# Filtreye uyan güncellemeleri okur: Superseded = 0, Required (NumMissing) > N, Vendor = metin kutusu.
# Vendor filtresi sunucuda uygulanır: "Company" kategorisinin ID'si bulunur, ana sorguya SMS_CIAllCategories alt sorgusu olarak eklenir.
function Import-RequiredUpdates {
    $minReq = [int]$nudReq.Value
    $vendor = $txtVendor.Text.Trim()
    $search = $txtTitle.Text.Trim()

    # Vendor, SMS_SoftwareUpdate üzerinde bir alan değil; güncellemenin "Company" kategorisidir.
    # WQL dizi alanlarını (LocalizedCategoryInstanceNames) filtreleyemediği için önce vendor'ın
    # kategori ID'sini buluyoruz (tek küçük sorgu), sonra ana sorguda SMS_CIAllCategories alt sorgusuyla
    # filtrelemeyi SUNUCUDA yapıyoruz. Böylece PowerShell tarafında hiçbir güncelleme elenmiyor.
    $vendorIds = @()
    $vendorClause = ''
    if ($vendor) {
        $vEsc = $vendor.Replace("'", "\'")
        $vRows = @(Invoke-CMWql "SELECT CategoryInstanceID, LocalizedCategoryInstanceName FROM SMS_UpdateCategoryInstance WHERE CategoryTypeName = 'Company' AND LocalizedCategoryInstanceName = '$vEsc'")
        if ($vRows.Count -eq 0) {
            $all = @(Invoke-CMWql "SELECT LocalizedCategoryInstanceName FROM SMS_UpdateCategoryInstance WHERE CategoryTypeName = 'Company'" |
                     ForEach-Object { [string]$_.LocalizedCategoryInstanceName } | Sort-Object)
            Write-Log "Vendor '$vendor' bulunamadı. Mevcut vendor'lar: $($all -join ', ')" 'WARN'
            return
        }
        $vendorIds = @($vRows | ForEach-Object { [int]$_.CategoryInstanceID })
        $cond = ($vendorIds | ForEach-Object { "SMS_CIAllCategories.CategoryInstanceID = $_" }) -join ' OR '
        $vendorClause = " AND SMS_SoftwareUpdate.CI_ID IN (SELECT SMS_CIAllCategories.CI_ID FROM SMS_CIAllCategories WHERE $cond)"
    }

    $base = "SELECT CI_ID, ArticleID, LocalizedDisplayName, DatePosted, IsSuperseded, IsExpired, NumMissing, LocalizedCategoryInstanceNames FROM SMS_SoftwareUpdate WHERE IsSuperseded = 0 AND NumMissing > $minReq"
    if ($chkExpired.Checked) { $base += ' AND IsExpired = 0' }

    Write-Log "Güncellemeler okunuyor (Vendor = '$vendor', Superseded = 0, Required > $minReq)..."
    try {
        $rows = @(Invoke-CMWql ($base + $vendorClause))
    }
    catch {
        if (-not $vendorClause) { throw }
        # Yalnızca SMS Provider alt sorguyu reddederse: vendor'a ait CI_ID kümesi tek sorguyla alınır.
        Write-Log "Alt sorgu reddedildi ($($_.Exception.Message)); vendor CI_ID listesi ile devam ediliyor." 'WARN'
        $set = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($vid in $vendorIds) {
            foreach ($c in @(Invoke-CMWql "SELECT CI_ID FROM SMS_CIAllCategories WHERE CategoryInstanceID = $vid")) { [void]$set.Add([string]$c.CI_ID) }
        }
        $rows = @(Invoke-CMWql $base | Where-Object { $set.Contains([string]$_.CI_ID) })
    }
    Write-Log "Sunucu $($rows.Count) kayıt döndürdü."
    if ($rows.Count -eq 0 -and $minReq -gt 0) {
        Write-Log "Not: 'Required > $minReq' koşulu Required = $minReq olan güncellemeleri hariç tutar. Konsoldaki güncellemelerde Required = $minReq ise sayıyı $($minReq - 1) yapın." 'WARN'
    }

    $kept = 0
    foreach ($r in $rows) {
        $names = @($r.LocalizedCategoryInstanceNames | Where-Object { $_ })

        $art = [string]$r.ArticleID
        if ($art -match '^\d+$') { $kb = "KB$art" } else { $kb = $art }
        $cat = (@($names | Where-Object { $_ -ne $vendor }) -join '; ')
        $title = [string]$r.LocalizedDisplayName

        if ($search) {
            $hay = "$kb $title $cat"
            if ($hay.IndexOf($search, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        }

        $key = [string]$r.CI_ID
        $script:Updates[$key] = [pscustomobject]@{
            Id       = $key
            KB       = $kb
            Title    = $title
            Posted   = ConvertTo-DateText $r.DatePosted
            Required = [int]$r.NumMissing
            Category = $cat
        }
        $kept++
    }

    Write-Log "Filtre sonrası $kept güncelleme kaldı."
}

# Listelenen güncellemelerin hangi SUG'lara üye olduğunu okur (SMS_CIRelation, RelationType=1).
# Performans: SUG başına ayrı bir WMI sorgusu atmak yerine (N+1 sorgu problemi;
# ortamda SUG sayısı kadar ayrı round-trip demektir), TÜM RelationType=1
# ilişkileri TEK sorguda çekip bellekte SUG ID'lerine göre filtreliyoruz.
# Böylece kaç SUG olursa olsun tek bir WMI çağrısı yeterli oluyor.
function Import-Memberships {
    $script:Members = @{}
    foreach ($k in @($script:Updates.Keys)) {
        $script:Members[$k] = New-Object 'System.Collections.Generic.HashSet[string]'
    }
    if ($script:Groups.Count -eq 0) { return }

    $groupIds = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($g in $script:Groups) { if ($g.Id -match '^\d+$') { [void]$groupIds.Add($g.Id) } }
    if ($groupIds.Count -eq 0) { return }

    $rows = Invoke-CMWql 'SELECT FromCIID, ToCIID FROM SMS_CIRelation WHERE RelationType = 1'
    foreach ($r in $rows) {
        $from = [string]$r.FromCIID
        if (-not $groupIds.Contains($from)) { continue }   # SUG'a ait olmayan ilişkileri ele
        $to = [string]$r.ToCIID
        if ($script:Members.ContainsKey($to)) { [void]$script:Members[$to].Add($from) }
    }
}

function Get-MemberNames {
    param([string]$UpdateId)
    if (-not $script:Members.ContainsKey($UpdateId)) { return '' }
    $names = @()
    foreach ($gid in $script:Members[$UpdateId]) {
        if ($script:GroupById.ContainsKey($gid)) { $names += $script:GroupById[$gid].Name }
    }
    return (($names | Sort-Object) -join '; ')
}

# ------------------------------------------------------------------
# Grid doldurma
# ------------------------------------------------------------------
function Show-Groups {
    $prev = $script:Suspend
    $script:Suspend = $true
    try {
        $gridGrp.CurrentCell = $null
        $gridGrp.Rows.Clear()
        foreach ($g in $script:Groups) {
            $i = $gridGrp.Rows.Add()
            $row = $gridGrp.Rows[$i]
            $row.Tag = $g.Id
            $row.Cells['gChk'].Value = [System.Windows.Forms.CheckState]::Unchecked
            $row.Cells['gName'].Value = $g.Name
            $row.Cells['gSel'].Value = ''
            $row.Cells['gTotal'].Value = $g.Total
        }
    } finally {
        $script:Suspend = $prev
    }
}

# Kombo kutusunu var olan Deployment Package'larla doldurur; yazarken otomatik tamamlama içindir.
function Show-Packages {
    $prevText = $cmbPkg.Text
    $cmbPkg.Items.Clear()
    $script:PkgDisplayToId = @{}
    $ac = New-Object System.Windows.Forms.AutoCompleteStringCollection
    foreach ($p in $script:Packages) {
        $disp = '{0}  [{1}]' -f $p.Name, $p.Id
        [void]$cmbPkg.Items.Add($disp)
        $script:PkgDisplayToId[$disp] = $p.Id
        [void]$ac.Add($disp)
    }
    $cmbPkg.AutoCompleteMode = 'SuggestAppend'
    $cmbPkg.AutoCompleteSource = 'CustomSource'
    $cmbPkg.AutoCompleteCustomSource = $ac
    if ($prevText -and $cmbPkg.Items.Contains($prevText)) { $cmbPkg.Text = $prevText }
    $lblPkgCount.Text = "$($script:Packages.Count) Deployment Package bulundu."
}

# Kombo kutusundaki metinden seçili paketin ID'sini bulur (listeden seçilmiş ya da adı birebir yazılmış olabilir).
function Get-SelectedPackageId {
    $t = $cmbPkg.Text.Trim()
    if (-not $t) { return $null }
    if ($script:PkgDisplayToId.ContainsKey($t)) { return $script:PkgDisplayToId[$t] }
    $match = @($script:Packages | Where-Object { $_.Name -eq $t })
    if ($match.Count -eq 1) { return $match[0].Id }
    return $null
}

# İşaretli dil kutularından (Windows Update / Office 365 Client Update) birleşik dil listesini üretir.
# Save-CMSoftwareUpdate tek bir dil listesini tüm seçili güncellemelere birlikte uyguladığı için
# burada yalnızca "English" / "Turkish" birleşimi hesaplanır; ürüne göre ayrı gönderilemez.
function Get-SelectedLanguages {
    $langs = New-Object 'System.Collections.Generic.List[string]'
    if ($chkLangWinEn.Checked -or $chkLangOffEn.Checked) { [void]$langs.Add('English') }
    if ($chkLangWinTr.Checked -or $chkLangOffTr.Checked) { [void]$langs.Add('Turkish') }
    return @($langs)
}

function Show-Updates {
    $prev = $script:Suspend
    $script:Suspend = $true
    try {
        $gridUpd.Rows.Clear()
        $sorted = $script:Updates.Values | Sort-Object @{ Expression = 'Required'; Descending = $true }, KB, Title
        foreach ($u in $sorted) {
            $i = $gridUpd.Rows.Add()
            $row = $gridUpd.Rows[$i]
            $row.Tag = $u.Id
            $row.Cells['uKB'].Value = $u.KB
            $row.Cells['uTitle'].Value = $u.Title
            $row.Cells['uReq'].Value = $u.Required
            $row.Cells['uPosted'].Value = $u.Posted
            $row.Cells['uCat'].Value = $u.Category
            $row.Cells['uGroups'].Value = Get-MemberNames $u.Id
            $row.Cells['uId'].Value = $u.Id
        }
        $gridUpd.ClearSelection()
    } finally {
        $script:Suspend = $prev
    }
    Update-GroupChecks
}

function Update-UpdateRowGroups {
    foreach ($row in $gridUpd.Rows) {
        $row.Cells['uGroups'].Value = Get-MemberNames ([string]$row.Tag)
    }
}

function Get-SelectedUpdateIds {
    $ids = @()
    foreach ($row in $gridUpd.SelectedRows) { $ids += [string]$row.Tag }
    return $ids
}

# Seçili güncellemelere göre SUG kutularını (tam / boş / kısmi) yeniden hesaplar.
function Update-GroupChecks {
    $sel = @(Get-SelectedUpdateIds)
    $n = $sel.Count
    $prev = $script:Suspend
    $script:Suspend = $true
    try {
        foreach ($row in $gridGrp.Rows) {
            $gid = [string]$row.Tag
            $m = 0
            foreach ($uid in $sel) {
                if ($script:Members.ContainsKey($uid) -and $script:Members[$uid].Contains($gid)) { $m++ }
            }
            if ($n -eq 0 -or $m -eq 0) { $st = [System.Windows.Forms.CheckState]::Unchecked }
            elseif ($m -eq $n)         { $st = [System.Windows.Forms.CheckState]::Checked }
            else                       { $st = [System.Windows.Forms.CheckState]::Indeterminate }

            $script:OrigCheck[$gid] = $st
            $row.Cells['gChk'].Value = $st
            if ($n -gt 0) { $row.Cells['gSel'].Value = "$m/$n" } else { $row.Cells['gSel'].Value = '' }
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Empty
        }
    } finally {
        $script:Suspend = $prev
    }
    if ($n -gt 0) { $lblCount.Text = "$($gridUpd.Rows.Count) güncelleme listelendi, $n seçili." }
    else          { $lblCount.Text = "$($gridUpd.Rows.Count) güncelleme listelendi." }
}

function Set-GroupFilter {
    $t = $txtFilter.Text.Trim()
    $gridGrp.CurrentCell = $null
    foreach ($row in $gridGrp.Rows) {
        $name = [string]$row.Cells['gName'].Value
        $visible = (-not $t) -or ($name.IndexOf($t, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
        if ($row.Visible -ne $visible) { $row.Visible = $visible }
    }
}

# ------------------------------------------------------------------
# Ana işlemler
# ------------------------------------------------------------------
function Invoke-List {
    Set-Busy $true
    try {
        [void](Get-Connection)

        Write-Log 'SUG listesi okunuyor...'
        Import-Groups
        Write-Log "$($script:Groups.Count) Software Update Group bulundu."
        Show-Groups

        $script:Updates = @{}
        Import-RequiredUpdates

        Write-Log 'SUG üyelikleri okunuyor...'
        Import-Memberships

        Write-Log 'Deployment Package listesi okunuyor...'
        Import-Packages
        Write-Log "$($script:Packages.Count) Deployment Package bulundu."
        Show-Packages

        if ($chkNoSug.Checked) {
            $before = $script:Updates.Count
            foreach ($k in @($script:Updates.Keys)) {
                if ($script:Members.ContainsKey($k) -and $script:Members[$k].Count -gt 0) {
                    $script:Updates.Remove($k)
                    $script:Members.Remove($k)
                }
            }
            Write-Log "Zaten bir SUG'a üye olan $($before - $script:Updates.Count) güncelleme listeden çıkarıldı."
        }

        Show-Updates
        Set-GroupFilter
        Write-Log "Listeleme tamamlandı: $($script:Updates.Count) benzersiz güncelleme." 'OK'
    }
    catch {
        Write-Log $_.Exception.Message 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    }
    finally {
        Set-Busy $false
    }
}

function Invoke-ApplyMembership {
    $sel = @(Get-SelectedUpdateIds)
    if ($sel.Count -eq 0) { Show-Message 'Önce soldaki listeden en az bir güncelleme seçin.' 'Uyarı' 'Warning'; return }

    $current = @{}
    foreach ($row in $gridGrp.Rows) { $current[[string]$row.Tag] = ConvertTo-CheckState $row.Cells['gChk'].Value }

    $plan = @(Get-MembershipPlan -SelectedIds $sel -Members $script:Members -Orig $script:OrigCheck -Current $current)
    if ($plan.Count -eq 0) { Show-Message 'Grup kutularında değişiklik yok.' 'Bilgi' 'Information'; return }
    $plan = @($plan | Sort-Object { $script:GroupById[$_.GroupId].Name })

    # SUG başına 1000 güncelleme sınırı kontrolü
    $over = @()
    foreach ($p in $plan) {
        $grp = $script:GroupById[$p.GroupId]
        $newTotal = $grp.Total + $p.Add.Count - $p.Remove.Count
        if ($newTotal -gt $script:SugLimit) { $over += ('- {0}: {1} güncelleme olurdu' -f $grp.Name, $newTotal) }
    }
    if ($over.Count -gt 0) {
        Show-Message ("Aşağıdaki SUG'lar $($script:SugLimit) güncelleme sınırını aşacağı için işlem yapılmadı:`n`n" + ($over -join "`n")) 'Uyarı' 'Warning'
        return
    }

    $simulate = $chkSim.Checked
    $lines = @()
    foreach ($p in $plan) {
        $grp = $script:GroupById[$p.GroupId]
        $lines += ('- {0}:  +{1} ekle,  -{2} çıkar  (toplam: {3} -> {4})' -f $grp.Name, $p.Add.Count, $p.Remove.Count, $grp.Total, ($grp.Total + $p.Add.Count - $p.Remove.Count))
    }
    if ($simulate) { $head = "SİMÜLASYON - hiçbir değişiklik yapılmayacak.`n`n" } else { $head = "" }
    $msg = "{0}{1} güncelleme için aşağıdaki SUG değişiklikleri uygulanacak:`n`n{2}`n`nDevam edilsin mi?" -f $head, $sel.Count, ($lines -join "`n")
    if ([System.Windows.Forms.MessageBox]::Show($msg, 'SUG üyeliklerini uygula', 'YesNo', 'Question') -ne 'Yes') { return }

    Set-Busy $true
    $pushed = $false
    $failed = 0
    try {
        if (-not $simulate) {
            if (-not (Get-Command -Name Set-CMSoftwareUpdateGroup -ErrorAction SilentlyContinue)) { Import-CMModule }
            Enter-CMSite
            $pushed = $true
        }
        foreach ($p in $plan) {
            $gname = $script:GroupById[$p.GroupId].Name
            if ($simulate) {
                Write-Log "[SİMÜLASYON] '$gname': +$($p.Add.Count) eklenecek, -$($p.Remove.Count) çıkarılacak." 'WARN'
                continue
            }
            try {
                $sug = Get-CMSoftwareUpdateGroup -Id $p.GroupId -ErrorAction Stop
                if (-not $sug) { throw "SUG bulunamadı: $gname" }
                if ($p.Add.Count -gt 0) {
                    Add-UpdatesToGroup -Sug $sug -GroupId $p.GroupId -GroupName $gname -UpdateIds $p.Add
                    Write-Log "'$gname': $($p.Add.Count) güncelleme eklendi." 'OK'
                }
                if ($p.Remove.Count -gt 0) {
                    Remove-UpdatesFromGroup -Sug $sug -GroupId $p.GroupId -GroupName $gname -UpdateIds $p.Remove
                    Write-Log "'$gname': $($p.Remove.Count) güncelleme çıkarıldı." 'OK'
                }
            } catch {
                $failed++
                Write-Log "'$gname' güncellenemedi: $($_.Exception.Message)" 'ERROR'
            }
        }

        if (-not $simulate) {
            Write-Log 'Üyelikler yeniden okunuyor...'
            Import-Groups
            Import-Memberships
            Show-Groups
            Update-UpdateRowGroups
            Update-GroupChecks
            Set-GroupFilter
        }
        if ($simulate) { Show-Message 'Simülasyon tamamlandı. Ayrıntılar günlükte.' }
        elseif ($failed -gt 0) { Show-Message "$failed grup güncellenemedi. Ayrıntılar günlükte." 'Uyarı' 'Warning' }
        else { Show-Message 'SUG üyelikleri güncellendi.' }
    }
    catch {
        Write-Log $_.Exception.Message 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    }
    finally {
        if ($pushed) { Pop-Location }
        Set-Busy $false
    }
}

# ------------------------------------------------------------------
# Arkaplan indirme (ayrı thread/runspace) - ilerleme takibi
# ------------------------------------------------------------------
# Her indirme "iş"i kendi Runspace'inde çalışır; bu sayede:
#   - UI donmaz (Save-CMSoftwareUpdate ana thread'i bloklamaz),
#   - Birden fazla indirme aynı anda (farklı paket/seçim ile) tetiklenebilir,
#   - Her güncelleme TEK TEK indirilir (SCCM/Save-CMSoftwareUpdate toplu ilerleme
#     bilgisi vermediği için), böylece "X / Y indirildi, Z kaldı" bilgisi elde edilir.
# İş durumu, ana thread'in Timer ile periyodik okuduğu senkronize (thread-safe)
# bir hashtable üzerinden paylaşılır; arkaplan thread'i doğrudan UI kontrollerine
# dokunmaz (WinForms cross-thread erişim kısıtlaması bu şekilde güvenle aşılır).
$script:DownloadWorker = {
    param(
        [string]  $Site,
        [string]  $Server,
        [string]  $ModulePath,
        [string[]]$UpdateIds,
        [hashtable]$Titles,
        [string]  $PackageName,
        [string[]]$Languages,
        [bool]    $Simulate,
        $Job
    )
    try {
        $Job.Status = 'Çalışıyor'
        if (-not $Simulate) {
            Import-Module -Name $ModulePath -ErrorAction Stop
            if (-not (Get-PSDrive -Name $Site -PSProvider CMSite -ErrorAction SilentlyContinue)) {
                New-PSDrive -Name $Site -PSProvider CMSite -Root $Server -Description 'SCCM Site' -ErrorAction Stop | Out-Null
            }
            Push-Location "$($Site):"
        }
        foreach ($id in $UpdateIds) {
            if ($Job.Cancel) { break }
            $Job.CurrentTitle = if ($Titles.ContainsKey($id)) { $Titles[$id] } else { $id }
            if ($Simulate) {
                Start-Sleep -Milliseconds 150
            } else {
                Save-CMSoftwareUpdate -SoftwareUpdateId $id -DeploymentPackageName $PackageName -SoftwareUpdateLanguage $Languages -ErrorAction Stop
            }
            $Job.Done = [int]$Job.Done + 1
        }
        $Job.Status = if ($Job.Cancel) { 'İptal Edildi' } else { 'Tamamlandı' }
    }
    catch {
        $Job.Status = 'Hata'
        $Job.Error  = $_.Exception.Message
    }
    finally {
        if (-not $Simulate) { try { Pop-Location -ErrorAction SilentlyContinue } catch { } }
        $Job.EndTime = Get-Date
    }
}

# Yeni bir indirme işini kuyruğa alır ve HEMEN döner (bloklamaz); iş kendi
# runspace'inde arkaplanda ilerler, ilerleme $tmrProgress ile grid'e yansır.
function Start-DownloadJob {
    param(
        [string[]]$UpdateIds,
        [string]  $PackageName,
        [string[]]$Languages,
        [bool]    $Simulate
    )
    $c = Get-Connection
    $modulePath = if ($Simulate) { '' } else { Get-CMModulePath }

    $titles = @{}
    foreach ($id in $UpdateIds) {
        $u = $script:Updates[$id]
        $titles[$id] = if ($u) { "$($u.KB) $($u.Title)" } else { $id }
    }

    $job = [hashtable]::Synchronized(@{
        Id           = [guid]::NewGuid().ToString()
        Package      = $PackageName
        Total        = $UpdateIds.Count
        Done         = 0
        CurrentTitle = ''
        Status       = 'Bekliyor'
        Error        = ''
        Simulate     = $Simulate
        StartTime    = Get-Date
        EndTime      = $null
        Cancel       = $false
        PS           = $null
        Runspace     = $null
        Handle       = $null
        Cleaned      = $false
    })

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ThreadOptions = 'ReuseThread'
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($script:DownloadWorker).
        AddArgument($c.Site).
        AddArgument($c.Server).
        AddArgument($modulePath).
        AddArgument($UpdateIds).
        AddArgument($titles).
        AddArgument($PackageName).
        AddArgument($Languages).
        AddArgument($Simulate).
        AddArgument($job)

    $job.PS       = $ps
    $job.Runspace = $rs
    $job.Handle   = $ps.BeginInvoke()

    [void]$script:DownloadJobs.Add($job)
    Write-Log "'$PackageName' için arkaplanda indirme kuyruğa alındı: $($UpdateIds.Count) güncelleme, diller: $($Languages -join ', ')."
    $tmrProgress.Start()
}

# Timer tick: aktif işlerin durumunu grid'e yansıtır; biten işleri temizler (EndInvoke/Dispose).
function Update-DownloadGrid {
    $rows = $gridDl.Rows
    $prevSel = @($rows | Where-Object { $_.Selected } | ForEach-Object { [string]$_.Tag })
    $gridDl.SuspendLayout()
    try {
        $rows.Clear()
        $jobsSnapshot = @($script:DownloadJobs.ToArray())
        foreach ($job in ($jobsSnapshot | Sort-Object StartTime -Descending)) {
            $i = $rows.Add()
            $row = $rows[$i]
            $row.Tag = $job.Id
            $row.Cells['dPkg'].Value = $job.Package
            $row.Cells['dProg'].Value = "$($job.Done) / $($job.Total)"
            $remain = [Math]::Max(0, [int]$job.Total - [int]$job.Done)
            $row.Cells['dRemain'].Value = $remain
            $pct = if ($job.Total -gt 0) { [Math]::Round((100.0 * $job.Done / $job.Total)) } else { 0 }
            $row.Cells['dPct'].Value = "%$pct"
            $row.Cells['dCurrent'].Value = $job.CurrentTitle
            $statusTxt = $job.Status
            if ($job.Simulate -and $statusTxt -notin @('Hata')) { $statusTxt = "$statusTxt (sim.)" }
            $row.Cells['dStatus'].Value = $statusTxt
            $endRef = if ($job.EndTime) { $job.EndTime } else { Get-Date }
            $row.Cells['dSure'].Value = '{0:mm\:ss}' -f ($endRef - $job.StartTime)
            $row.Cells['dCancel'].Value = if ($job.Status -in @('Bekliyor', 'Çalışıyor')) { 'İptal' } else { '' }
            if ($job.Status -eq 'Hata') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose }
            elseif ($job.Status -eq 'Tamamlandı') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Honeydew }
            if ($prevSel -contains $job.Id) { $row.Selected = $true }
        }
    } finally {
        $gridDl.ResumeLayout()
    }

    # Bitmiş (Handle.IsCompleted) işleri bir kez EndInvoke edip runspace/PS kaynaklarını serbest bırak.
    foreach ($job in @($script:DownloadJobs.ToArray())) {
        if (-not $job.Cleaned -and $job.Handle -and $job.Handle.IsCompleted) {
            try {
                if ($job.PS) { $job.PS.EndInvoke($job.Handle) }
            } catch {
                if (-not $job.Error) { $job.Error = $_.Exception.Message }
                if ($job.Status -notin @('Hata', 'Tamamlandı', 'İptal Edildi')) { $job.Status = 'Hata' }
            } finally {
                try { if ($job.PS) { $job.PS.Dispose() } } catch { }
                try { if ($job.Runspace) { $job.Runspace.Close(); $job.Runspace.Dispose() } } catch { }
                $job.Cleaned = $true
                if ($job.Status -eq 'Hata' -and $job.Error) {
                    Write-Log "'$($job.Package)' indirmesi hata ile bitti: $($job.Error)" 'ERROR'
                } elseif ($job.Status -eq 'Tamamlandı') {
                    Write-Log "'$($job.Package)': indirme tamamlandı ($($job.Done)/$($job.Total) güncelleme)." 'OK'
                } elseif ($job.Status -eq 'İptal Edildi') {
                    Write-Log "'$($job.Package)': indirme iptal edildi ($($job.Done)/$($job.Total) tamamlanmıştı)." 'WARN'
                }
            }
        }
    }

    # Aktif iş kalmadıysa timer'ı durdur (gereksiz yere UI thread'ini meşgul etmesin).
    $anyActive = @($script:DownloadJobs.ToArray() | Where-Object { $_.Status -in @('Bekliyor', 'Çalışıyor') }).Count -gt 0
    if (-not $anyActive) { $tmrProgress.Stop() }
}

# Seçili güncellemelerin içeriğini VAR OLAN bir Deployment Package'a indirir.
# Save-CMSoftwareUpdate -Location verilmeden çağrılır: bu durumda içerik SCCM tarafından
# doğrudan İnternet'ten (Microsoft Update) indirilir; hiçbir yeni paket oluşturulmaz.
# İndirme arkaplanda (ayrı bir runspace/thread'de) çalışır; bu fonksiyon HEMEN döner,
# bu yüzden pencere yanıt vermeye devam eder ve düğmeye tekrar basılarak (farklı paket/
# seçimle) aynı anda birden fazla indirme başlatılabilir.
function Invoke-Download {
    $sel = @(Get-SelectedUpdateIds)
    if ($sel.Count -eq 0) { Show-Message 'Önce soldaki listeden en az bir güncelleme seçin.' 'Uyarı' 'Warning'; return }

    $pkgId = Get-SelectedPackageId
    if (-not $pkgId) { Show-Message "Var olan bir Deployment Package seçin (kombodan seçin ya da adını birebir yazın).`n`nGerekirse önce ""Paketleri Yenile"" düğmesine basın." 'Uyarı' 'Warning'; return }
    $pkgName = $script:PackageById[$pkgId].Name

    $langs = @(Get-SelectedLanguages)
    if ($langs.Count -eq 0) { Show-Message 'En az bir dil seçin (English / Turkish).' 'Uyarı' 'Warning'; return }

    $simulate = $chkSim.Checked
    $head = if ($simulate) { "SİMÜLASYON - hiçbir indirme başlatılmayacak.`n`n" } else { "" }
    $msg = "{0}{1} güncelleme, '{2}' Deployment Package'ına İNTERNETTEN (Microsoft Update) indirilecek.`nDiller: {3}`n`nİndirme ARKAPLANDA çalışacak; pencere kullanılabilir kalacak ve aşağıdaki ""Aktif İndirmeler"" listesinden ilerleme takip edilebilir. Devam edilsin mi?" -f $head, $sel.Count, $pkgName, ($langs -join ', ')
    if ([System.Windows.Forms.MessageBox]::Show($msg, 'İçerik indir', 'YesNo', 'Question') -ne 'Yes') { return }

    try {
        if (-not $simulate -and -not (Get-Command -Name Save-CMSoftwareUpdate -ErrorAction SilentlyContinue)) {
            # Yalnızca modül yolunun doğrulanması için; asıl Import-Module işi arkaplan runspace'inde yapılır.
            [void](Get-CMModulePath)
        }
        Start-DownloadJob -UpdateIds $sel -PackageName $pkgName -Languages $langs -Simulate $simulate
    }
    catch {
        Write-Log "İndirme başlatılamadı: $($_.Exception.Message)" 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    }
}

# ------------------------------------------------------------------
# Olaylar
# ------------------------------------------------------------------
$btnList.Add_Click({ Invoke-List })
$btnApply.Add_Click({ Invoke-ApplyMembership })
$btnDownload.Add_Click({ Invoke-Download })
$btnPkgRefresh.Add_Click({
    Set-Busy $true
    try {
        [void](Get-Connection)
        Import-Packages
        Show-Packages
        Write-Log "$($script:Packages.Count) Deployment Package bulundu."
    } catch {
        Write-Log $_.Exception.Message 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    } finally {
        Set-Busy $false
    }
})
$btnClose.Add_Click({ $form.Close() })
$btnSelAll.Add_Click({ $gridUpd.SelectAll() })
$btnSelNone.Add_Click({ $gridUpd.ClearSelection() })

$txtFilter.Add_TextChanged({ Set-GroupFilter })

$gridUpd.Add_SelectionChanged({
    if (-not $script:Suspend) { Update-GroupChecks }
})

# 3 durumlu kutuda tıklamanın hemen işlenmesi için
$gridGrp.Add_CurrentCellDirtyStateChanged({
    if ($gridGrp.IsCurrentCellDirty) {
        [void]$gridGrp.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})
$gridGrp.Add_CellValueChanged({
    param($sender, $e)
    if ($script:Suspend -or $e.RowIndex -lt 0 -or $e.ColumnIndex -ne 0) { return }
    $row = $gridGrp.Rows[$e.RowIndex]
    $gid = [string]$row.Tag
    $cur = ConvertTo-CheckState $row.Cells['gChk'].Value
    if ($cur -ne $script:OrigCheck[$gid]) { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGoldenrodYellow }
    else                                  { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Empty }
})
$gridGrp.Add_DataError({ param($sender, $e) $e.ThrowException = $false })
$gridUpd.Add_DataError({ param($sender, $e) $e.ThrowException = $false })

# "İptal" düğmesine tıklanan satırın işine Cancel=true işaretler; iş, o an
# indirmekte olduğu öğe bittikten sonra döngüden çıkar (sert kesme yapılmaz).
$gridDl.Add_CellContentClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $gridDl.Columns[$e.ColumnIndex].Name -ne 'dCancel') { return }
    $jobId = [string]$gridDl.Rows[$e.RowIndex].Tag
    $job = @($script:DownloadJobs.ToArray()) | Where-Object { $_.Id -eq $jobId } | Select-Object -First 1
    if ($job -and $job.Status -in @('Bekliyor', 'Çalışıyor')) {
        $job.Cancel = $true
        Write-Log "'$($job.Package)' indirmesi iptal ediliyor (devam eden öğe bittikten sonra duracak)..." 'WARN'
    }
})

$btnDlClear.Add_Click({
    $keep = New-Object System.Collections.ArrayList
    foreach ($job in @($script:DownloadJobs.ToArray())) {
        if ($job.Status -in @('Bekliyor', 'Çalışıyor')) { [void]$keep.Add($job) }
    }
    $script:DownloadJobs.Clear()
    foreach ($job in $keep) { [void]$script:DownloadJobs.Add($job) }
    Update-DownloadGrid
})

# Pencere kapatılırken arkaplanda kalan indirme işlerini iptal edip kaynakları serbest bırak.
$form.Add_FormClosing({
    $active = @($script:DownloadJobs.ToArray() | Where-Object { $_.Status -in @('Bekliyor', 'Çalışıyor') })
    if ($active.Count -eq 0) { return }
    foreach ($job in $active) { $job.Cancel = $true }
    $deadline = (Get-Date).AddSeconds(5)
    while ((Get-Date) -lt $deadline -and (@($active | Where-Object { -not $_.Handle.IsCompleted }).Count -gt 0)) {
        Start-Sleep -Milliseconds 100
    }
    foreach ($job in $active) {
        try { if ($job.PS) { $job.PS.Stop(); $job.PS.Dispose() } } catch { }
        try { if ($job.Runspace) { $job.Runspace.Close(); $job.Runspace.Dispose() } } catch { }
    }
})

# ------------------------------------------------------------------
# Açılış
# ------------------------------------------------------------------
Write-Log 'Hazır. Filtreyi kontrol edip "Listele" düğmesine basın. Varsayılan olarak SİMÜLASYON açıktır.'
Import-ConnectionDefaults
$lblCount.Text = 'Henüz liste yüklenmedi.'

[void]$form.ShowDialog()