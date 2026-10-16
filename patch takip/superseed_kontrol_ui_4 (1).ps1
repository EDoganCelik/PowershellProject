<#
.SYNOPSIS
    SCCM'de superseded (yerini daha yeni bir güncellemeye bırakmış) güncellemeleri listeler,
    Software Update Group (SUG) üyeliklerini düzenler ve istenirse Deployment Package içeriğinden siler (UI ile).

.DESCRIPTION
    - create-sccm-update-object.ps1 tarafından oluşturulan state.json dosyasını OKUR (yazmaz).
      Aranan konumlar: Desktop\update-state\state.json ve %LOCALAPPDATA%\update-state\state.json
      (ikisi de varsa en yenisi kullanılır; "Gözat..." ile başka bir dosya da seçilebilir).
    - State kaydında (SugId / PkgId) bulunan Deployment Package içeriğindeki ve SUG üyesi olan
      superseded güncellemeleri listeler. İstenirse "Expired" olanlar da listelenir.
    - "Kaynak" seçimi ile state.json'da olmayan, ortamda daha önce (bu araç kullanılmadan) oluşturulmuş
      Deployment Package'lar ve Software Update Group'lar da taranabilir:
          State kaydı / Ortamdaki bir Package / Ortamdaki bir SUG / Ortamdaki TÜM Package'lar + TÜM SUG'lar
    - "Ortamdaki TÜM Superseded + Deploy edilmiş güncellemeler" kaynağı: state.json'a veya tek tek Package/SUG
      gezmeye gerek kalmadan, ortamdaki tüm superseded (istenirse Expired) güncellemeleri TEK sorguyla okur ve
      yalnızca deploy edilmiş bir SUG'a (SMS_UpdateGroupAssignment) üye olanları listeler.
    - "Ortamdaki TÜM Package'larda indirilmiş + Superseded + Deploy edilmiş güncellemeler" kaynağı: ortamdaki tüm
      superseded güncellemeleri tek sorguyla okur; içeriği herhangi bir Software Update Deployment Package'a indirilmiş
      (SMS_PackageToContent) VE deploy edilmiş bir SUG'a (SMS_UpdateGroupAssignment) üye olanları, bulundukları
      Package ID'leriyle listeler. "Seçilileri Package'tan Sil" bu modda da kullanılabilir.
    - "Ortamdaki TÜM Package'larda indirilmiş + Superseded güncellemeler" kaynağı: yukarıdakinin DEPLOY şartı olmayan
      hâli. Ortamdaki tüm superseded güncellemeleri tek sorguyla okur; içeriği herhangi bir Software Update Deployment
      Package'a indirilmiş (SMS_PackageToContent) olanları, deploy edilmiş olup olmadıklarına bakmadan listeler.
      "Seçilileri Package'tan Sil" bu modda da kullanılabilir.
    - "Vendor" kutusu: yalnızca "Company" kategorisi bu değere eşit olan güncellemeler listelenir (varsayılan: Microsoft).
      Boş bırakılırsa vendor filtresi uygulanmaz. Tüm kaynak modları için geçerlidir.
    - Sağ tarafta ortamdaki TÜM Software Update Group'lar listelenir (SCCM'deki "Edit Membership" penceresi gibi).
      Seçili güncellemeler için kutular şöyle çalışır:
          işaretli   : seçili güncellemelerin hepsi grubun üyesi (işaretlerseniz eksikler EKLENİR)
          boş        : hiçbiri üye değil (boşaltırsanız üye olanlar ÇIKARILIR)
          dolu kare  : bir kısmı üye; bu durumda değişiklik yapılmaz
    - "Seçilileri Package'tan Sil": seçili güncellemelerin içeriğini state'teki Deployment Package'tan siler
      (SMS_SoftwareUpdatesPackage.RemoveContent). Başka bir güncelleme ile paylaşılan içerik silinmez.
      Silme sonrası sonuç SMS_PackageToContent'ten DOĞRULANIR; içerik hâlâ package'taysa başarılı sayılmaz.
    - Sağdaki "Deployment Package'lar" sekmesinde işaretlenen (Hariç) package'lar silme işleminden MUAF tutulur:
      bir güncelleme 5 package'ta ise ve 1'i işaretliyse, içerik yalnızca diğer 4'ünden silinir.
    - "Simülasyon" kutusu işaretliyken hiçbir değişiklik yapılmaz, sadece ne yapılacağı günlüğe yazılır.

.NOTES
    Configuration Manager Console'un kurulu olduğu bir makinede çalıştırın.
    Okuma işlemleri SMS Provider WMI (root\SMS\site_XXX) üzerinden yapılır (DCOM/RPC erişimi gerekir).
    SUG üyeliği değişiklikleri ConfigurationManager modülü ile yapılır; SUG düzenleme yetkisi gerekir.
    Package'tan silme için Software Update Package düzenleme yetkisi gerekir.
    Önerilen sıra: önce SUG üyeliklerinden çıkarın, sonra Package'tan silin.
#>

#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ------------------------------------------------------------------
# Genel durum
# ------------------------------------------------------------------
$script:StateFile    = $null
$script:StateEntries = @()      # state.json'dan gelen kayıtlar (combobox ile aynı sırada)
$script:Groups       = @()      # ortamdaki tüm SUG'lar
$script:GroupById    = @{}      # SUG CI_ID (string) -> SUG
$script:Updates      = @{}      # güncelleme CI_ID (string) -> kayıt
$script:Members      = @{}      # güncelleme CI_ID (string) -> HashSet[string] (üye olduğu SUG CI_ID'leri)
$script:OrigCheck    = @{}      # SUG CI_ID (string) -> kutunun ilk hesaplanan durumu
$script:Suspend      = $false   # UI olaylarını geçici olarak susturmak için
$script:EnvPackages  = @()      # ortamdaki tüm Software Update Deployment Package'lar (Id, Name)
$script:DeployedIds  = @{}      # deploy edilmiş SUG (ve varsa doğrudan güncelleme) CI_ID'leri (string -> $true)
$script:ExcludedPkgs = @{}      # silmeden HARİÇ tutulacak PackageID'ler (string -> bool)
$script:PkgNames     = @{}      # PackageID (string) -> package adı
$script:VendorFilter = ''       # Listele sırasında uygulanan vendor (Company kategorisi); boşsa filtre yok
$script:VendorSkipped = 0       # vendor filtresiyle elenen kayıt sayacı
$script:VendorNoInfo  = 0       # kategori bilgisi (UniqueIDs/Names) boş gelen kayıt sayacı
$script:VendorUniqueIds = $null # vendor'a karşılık gelen Company kategorisinin CategoryInstance_UniqueID'leri (küçük harf)
$script:VendorNames     = $null # vendor'a karşılık gelen Company kategori adları (küçük harf)
$script:VendorCiIds     = $null # Company kategorisine bağlı güncelleme CI_ID'leri (SMS_CIAllCategories'ten)
$script:VendorSample    = @()   # eşleşmeyen kayıtlardan teşhis örnekleri

# ------------------------------------------------------------------
# state.json okuma (SADECE okuma)
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

function ConvertTo-StateEntry {
    param($Raw)
    if (-not $Raw) { return $null }
    return [pscustomobject]@{
        Timestamp = [string]$Raw.Timestamp
        Site      = [string]$Raw.SiteCode
        Server    = [string]$Raw.ProviderServer
        SugId     = [string]$Raw.SugId
        SugName   = [string]$Raw.SugName
        PkgId     = [string]$Raw.PkgId
        PkgName   = [string]$Raw.PkgName
    }
}

function Get-EntryLabel {
    param($E)
    return ('{0}  |  SUG: {1}  |  PKG: {2} ({3})' -f $E.Timestamp, $E.SugName, $E.PkgName, $E.PkgId)
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
        # NOT: PowerShell degisken adlari buyuk/kucuk harf duyarsizdir. Eskiden $orig = $Orig[$gid]
        # yazilmisti; bu, [hashtable] tipli $Orig parametresinin uzerine yazmaya calisip hata veriyordu.
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

# Package içindeki (CI_ID, ContentID) ikililerinden, seçili güncellemelere ait ve
# başka güncelleme tarafından KULLANILMAYAN ContentID'leri bulur.
function Get-ContentRemovalSet {
    param(
        [object[]]$Pairs,
        [string[]]$SelectedIds
    )
    $sel = @{}
    foreach ($s in $SelectedIds) { $sel[[string]$s] = $true }

    $refs = @{}
    foreach ($p in $Pairs) {
        $c  = [string]$p.ContentID
        $ci = [string]$p.CI_ID
        if (-not $refs.ContainsKey($c)) { $refs[$c] = New-Object 'System.Collections.Generic.List[string]' }
        $refs[$c].Add($ci)
    }

    $remove = @()
    $shared = @()
    foreach ($c in @($refs.Keys)) {
        $mine = $false
        $others = $false
        foreach ($ci in $refs[$c]) {
            if ($sel.ContainsKey($ci)) { $mine = $true } else { $others = $true }
        }
        if ($mine -and -not $others) { $remove += [uint32]$c }
        elseif ($mine -and $others)  { $shared += [uint32]$c }
    }
    return [pscustomobject]@{
        Remove = [uint32[]]$remove
        Shared = [uint32[]]$shared
    }
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
$form.Text = 'SCCM - Superseded Güncellemeler & SUG Üyelik Yönetimi'
$form.ClientSize = New-Object System.Drawing.Size(1280, 892)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 792)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# --- Bağlantı ve State ---
$gTop = New-UiGroup $form 'Bağlantı ve State' 10 8 1260 124
$gTop.Anchor = 'Top,Left,Right'
[void](New-UiLabel $gTop 'Site Kodu:' 12 27 70)
$txtSite = New-UiTextBox $gTop 85 24 70
[void](New-UiLabel $gTop 'SMS Provider:' 170 27 85)
$txtServer = New-UiTextBox $gTop 258 24 230
[void](New-UiLabel $gTop 'State dosyası:' 505 27 90)
$txtStateFile = New-UiTextBox $gTop 598 24 480
$txtStateFile.ReadOnly = $true
$txtStateFile.Anchor = 'Top,Left,Right'
$btnBrowse = New-UiButton $gTop 'Gözat...' 1085 22 75 26
$btnBrowse.Anchor = 'Top,Right'
$btnReload = New-UiButton $gTop 'Yeniden Oku' 1165 22 84 26
$btnReload.Anchor = 'Top,Right'

[void](New-UiLabel $gTop 'State kaydı:' 12 62 70)
$cmbState = New-Object System.Windows.Forms.ComboBox
$cmbState.DropDownStyle = 'DropDownList'
$cmbState.Location = New-Object System.Drawing.Point(85, 58)
$cmbState.Size = New-Object System.Drawing.Size(560, 24)
$cmbState.Anchor = 'Top,Left,Right'
$gTop.Controls.Add($cmbState)
$lblVendor = New-UiLabel $gTop 'Vendor:' 660 62 50
$lblVendor.Anchor = 'Top,Right'
$txtVendor = New-UiTextBox $gTop 712 58 160
$txtVendor.Text = 'Microsoft'
$txtVendor.Anchor = 'Top,Right'
$chkExpired = New-Object System.Windows.Forms.CheckBox
$chkExpired.Text = 'Expired olanları da listele'
$chkExpired.Location = New-Object System.Drawing.Point(890, 60)
$chkExpired.Size = New-Object System.Drawing.Size(200, 22)
$chkExpired.Anchor = 'Top,Right'
$gTop.Controls.Add($chkExpired)
$btnList = New-UiButton $gTop 'Listele' 1100 55 149 30
$btnList.Anchor = 'Top,Right'

[void](New-UiLabel $gTop 'Kaynak:' 12 96 70)
$cmbSource = New-Object System.Windows.Forms.ComboBox
$cmbSource.DropDownStyle = 'DropDownList'
$cmbSource.Location = New-Object System.Drawing.Point(85, 92)
$cmbSource.Size = New-Object System.Drawing.Size(330, 24)
[void]$cmbSource.Items.Add('State kaydı (yukarıdaki seçim)')
[void]$cmbSource.Items.Add('Ortamdaki bir Deployment Package')
[void]$cmbSource.Items.Add('Ortamdaki bir Software Update Group')
[void]$cmbSource.Items.Add('Ortamdaki TÜM Package''lar + TÜM SUG''lar')
[void]$cmbSource.Items.Add('Ortamdaki TÜM Superseded + Deploy edilmiş güncellemeler')
[void]$cmbSource.Items.Add('Ortamdaki TÜM Package''larda indirilmiş + Superseded + Deploy edilmiş güncellemeler')
[void]$cmbSource.Items.Add('Ortamdaki TÜM Package''larda indirilmiş + Superseded güncellemeler (deploy şartı yok)')
$cmbSource.DropDownWidth = 520   # uzun seçenek adları açılır listede tam görünsün
$cmbSource.SelectedIndex = 0
$gTop.Controls.Add($cmbSource)
[void](New-UiLabel $gTop 'Yapı:' 430 96 40)
$cmbEnv = New-Object System.Windows.Forms.ComboBox
$cmbEnv.DropDownStyle = 'DropDownList'
$cmbEnv.Location = New-Object System.Drawing.Point(472, 92)
$cmbEnv.Size = New-Object System.Drawing.Size(613, 24)
$cmbEnv.Anchor = 'Top,Left,Right'
$cmbEnv.Enabled = $false
$gTop.Controls.Add($cmbEnv)
$btnEnv = New-UiButton $gTop 'Ortamı Oku' 1100 90 149 28
$btnEnv.Anchor = 'Top,Right'

# --- Superseded güncellemeler ---
$gUpd = New-UiGroup $form 'Superseded güncellemeler (seçili kapsam)' 10 140 800 500
$gUpd.Anchor = 'Top,Bottom,Left,Right'
$gridUpd = New-UiGrid $gUpd 10 22 780 430
$gridUpd.MultiSelect = $true
$gridUpd.ReadOnly = $true
$gridUpd.AutoSizeColumnsMode = 'None'          # Fill kapalı: sütunlar sabit genişlikte, toplam sığmazsa X scroll çıkar
$gridUpd.ScrollBars = 'Both'
$gridUpd.AllowUserToResizeColumns = $true
Add-UiTextColumn $gridUpd 'uKB' 'KB' 80
Add-UiTextColumn $gridUpd 'uTitle' 'Başlık' 360
Add-UiTextColumn $gridUpd 'uPosted' 'Yayın' 85
Add-UiTextColumn $gridUpd 'uStatus' 'Durum' 150
Add-UiTextColumn $gridUpd 'uPkg' 'Package' 85
Add-UiTextColumn $gridUpd 'uGroups' 'Üye olduğu SUG''lar' 260
Add-UiTextColumn $gridUpd 'uId' 'CI_ID' 75
Enable-ColumnDoubleClickExpand $gridUpd

$btnSelAll = New-UiButton $gUpd 'Tümünü Seç' 10 462 100 26
$btnSelAll.Anchor = 'Bottom,Left'
$btnSelNone = New-UiButton $gUpd 'Seçimi Temizle' 115 462 110 26
$btnSelNone.Anchor = 'Bottom,Left'
$lblCount = New-UiLabel $gUpd '' 235 466 545
$lblCount.Anchor = 'Bottom,Left'

# --- Sağ panel: SUG'lar ve Deployment Package'lar (sekmeli) ---
$gGrp = New-UiGroup $form 'Gruplar ve Package''lar (ortamdaki tümü)' 820 140 450 500
$gGrp.Anchor = 'Top,Bottom,Right'   # sabit genişlik (450), pencere büyüyünce sağa kayar
$tabRight = New-Object System.Windows.Forms.TabControl
$tabRight.Location = New-Object System.Drawing.Point(6, 18)
$tabRight.Size = New-Object System.Drawing.Size(438, 476)
$tabRight.Anchor = 'Top,Bottom,Left,Right'
$gGrp.Controls.Add($tabRight)
$tabSug = New-Object System.Windows.Forms.TabPage
$tabSug.Text = 'Software Update Grupları'
$tabPkg = New-Object System.Windows.Forms.TabPage
$tabPkg.Text = 'Deployment Package''lar (silmeden hariç tut)'
[void]$tabRight.TabPages.Add($tabSug)
[void]$tabRight.TabPages.Add($tabPkg)

# -- Sekme 1: SUG'lar --
[void](New-UiLabel $tabSug 'Grup ara:' 6 10 60)
$txtFilter = New-UiTextBox $tabSug 70 7 356
$txtFilter.Anchor = 'Top,Left,Right'
$gridGrp = New-UiGrid $tabSug 6 36 424 360
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
Add-UiTextColumn $gridGrp 'gState' ([string][char]0x2605) 30
$lblLegend = New-UiLabel $tabSug ('İşaretli: ekle  |  Boş: çıkar  |  Dolu kare: dokunma.  ' + [char]0x2605 + ' = state.json''daki SUG. Güncelleme seçimi değişirse kutular yeniden hesaplanır.') 6 400 424 40
$lblLegend.Anchor = 'Bottom,Left,Right'
$lblLegend.ForeColor = [System.Drawing.Color]::DimGray

# -- Sekme 2: Deployment Package'lar (silmeden hariç tut) --
[void](New-UiLabel $tabPkg 'Package ara:' 6 10 72)
$txtPkgFilter = New-UiTextBox $tabPkg 82 7 344
$txtPkgFilter.Anchor = 'Top,Left,Right'
$gridPkg = New-UiGrid $tabPkg 6 36 424 330
$gridPkg.MultiSelect = $false
$colPkgChk = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colPkgChk.Name = 'pChk'
$colPkgChk.HeaderText = 'Hariç'
$colPkgChk.Width = 45
$colPkgChk.ThreeState = $false
[void]$gridPkg.Columns.Add($colPkgChk)
Add-UiTextColumn $gridPkg 'pName' 'Package adı' 140 $true 100
Add-UiTextColumn $gridPkg 'pId' 'Package ID' 80
Add-UiTextColumn $gridPkg 'pSel' 'Seçili' 50
Add-UiTextColumn $gridPkg 'pTotal' 'Listede' 55
$lblPkgLegend = New-UiLabel $tabPkg ('İşaretli (Hariç): içerik bu package''ta KALIR, "Seçilileri Package''tan Sil" işleminde atlanır.  Boş: seçili güncellemelerin içeriği bu package''tan silinir.  ' + 'Seçili = seçili güncellemelerden kaçı bu package''ta, Listede = listelenen güncellemelerden kaçı bu package''ta.') 6 370 424 70
$lblPkgLegend.Anchor = 'Bottom,Left,Right'
$lblPkgLegend.ForeColor = [System.Drawing.Color]::DimGray

# --- İşlemler ---
$gAct = New-UiGroup $form 'İşlemler' 10 646 1260 62
$gAct.Anchor = 'Bottom,Left,Right'
$chkSim = New-Object System.Windows.Forms.CheckBox
$chkSim.Text = 'Simülasyon (değişiklik yapma, sadece günlüğe yaz)'
$chkSim.Location = New-Object System.Drawing.Point(14, 26)
$chkSim.Size = New-Object System.Drawing.Size(400, 22)
$chkSim.Checked = $true
$gAct.Controls.Add($chkSim)
$btnApply = New-UiButton $gAct 'SUG Üyeliklerini Uygula' 430 20 210 32
$btnRemove = New-UiButton $gAct 'Seçilileri Package''tan Sil' 650 20 210 32
$btnClose = New-UiButton $gAct 'Kapat' 1140 20 105 32
$btnClose.Anchor = 'Top,Right'

# --- Günlük ---
$lblLog = New-UiLabel $form 'İşlem Günlüğü:' 14 714 200
$lblLog.Anchor = 'Bottom,Left'
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 736)
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
    foreach ($c in @($btnList, $btnApply, $btnRemove, $btnReload, $btnBrowse, $btnEnv)) { $c.Enabled = (-not $On) }
    if ($On) { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor }
    else     { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    [System.Windows.Forms.Application]::DoEvents()
}

function Show-Message {
    param([string]$Text, [string]$Title = 'Bilgi', [string]$Icon = 'Information')
    [void][System.Windows.Forms.MessageBox]::Show($Text, $Title, 'OK', $Icon)
}

# ------------------------------------------------------------------
# state.json -> UI
# ------------------------------------------------------------------
function Import-StateFile {
    param([string]$Path)

    $script:StateFile = $Path
    $script:StateEntries = @()
    $cmbState.Items.Clear()

    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) {
        $txtStateFile.Text = '(state.json bulunamadı - "Gözat..." ile seçin)'
        Write-Log 'state.json bulunamadı.' 'WARN'
        return
    }
    $txtStateFile.Text = $Path

    try {
        $state = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Log "state.json okunamadı: $($_.Exception.Message)" 'ERROR'
        return
    }

    $raws = @()
    if ($state.LastRun) { $raws += $state.LastRun }
    if ($state.History) {
        $hist = @($state.History)
        [array]::Reverse($hist)          # en yeni başta
        $raws += $hist
    }

    $entries = @()
    $seen = @{}
    foreach ($r in $raws) {
        $e = ConvertTo-StateEntry $r
        if (-not $e) { continue }
        if (-not $e.SugId -and -not $e.PkgId) { continue }
        $key = '{0}|{1}|{2}' -f $e.Timestamp, $e.SugId, $e.PkgId
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $entries += $e
    }
    $script:StateEntries = $entries

    if ($entries.Count -eq 0) {
        Write-Log 'state.json içinde kullanılabilir kayıt (SugId/PkgId) bulunamadı.' 'WARN'
        return
    }

    foreach ($e in $entries) { [void]$cmbState.Items.Add((Get-EntryLabel $e)) }
    if ($entries.Count -gt 1) { [void]$cmbState.Items.Add('--- Tüm state kayıtları (History) ---') }
    $cmbState.SelectedIndex = 0

    Write-Log "State yüklendi: $Path ($($entries.Count) kayıt)"
}

# state.json kaydı ile aynı şekilde (SugId / PkgId ...) bir kapsam kaydı üretir (ortamdan seçilen yapılar için).
function New-ScopeEntry {
    param([string]$PkgId = '', [string]$PkgName = '', [string]$SugId = '', [string]$SugName = '')
    return [pscustomobject]@{
        Timestamp = ''
        Site      = ''
        Server    = ''
        SugId     = $SugId
        SugName   = $SugName
        PkgId     = $PkgId
        PkgName   = $PkgName
    }
}

# Taranacak kapsam: state kaydı, ortamdaki tek bir Package / SUG ya da ortamdaki tümü.
function Get-SelectedScope {
    switch ($cmbSource.SelectedIndex) {
        1 {
            $i = $cmbEnv.SelectedIndex
            if ($i -ge 0 -and $i -lt $script:EnvPackages.Count) {
                $pk = $script:EnvPackages[$i]
                return @(New-ScopeEntry -PkgId $pk.Id -PkgName $pk.Name)
            }
            return @()
        }
        2 {
            $i = $cmbEnv.SelectedIndex
            if ($i -ge 0 -and $i -lt $script:Groups.Count) {
                $gr = $script:Groups[$i]
                return @(New-ScopeEntry -SugId $gr.Id -SugName $gr.Name)
            }
            return @()
        }
        3 {
            $all = @()
            foreach ($pk in $script:EnvPackages) { $all += New-ScopeEntry -PkgId $pk.Id -PkgName $pk.Name }
            foreach ($gr in $script:Groups)      { $all += New-ScopeEntry -SugId $gr.Id -SugName $gr.Name }
            return @($all)
        }
        4 {
            # Ortamdaki tüm superseded + deploy edilmiş güncellemeler: Package/SUG kapsamı gerekmez
            # (Invoke-List bu modu tek sorguyla kendisi çözer).
            return @()
        }
        5 {
            # Ortamdaki tüm Package'larda indirilmiş + superseded + deploy edilmiş güncellemeler: kapsam gerekmez (tek sorguyla çözülür).
            return @()
        }
        6 {
            # Ortamdaki tüm Package'larda indirilmiş + superseded güncellemeler (deploy şartı yok): kapsam gerekmez (tek sorguyla çözülür).
            return @()
        }
        default {
            $idx = $cmbState.SelectedIndex
            if ($idx -lt 0 -or $script:StateEntries.Count -eq 0) { return @() }
            if ($idx -ge $script:StateEntries.Count) { return @($script:StateEntries) }
            return @($script:StateEntries[$idx])
        }
    }
}

# ------------------------------------------------------------------
# SCCM erişimi
# ------------------------------------------------------------------
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

# Ortamdaki tüm Deployment Package'ları ve SUG'ları okur (state.json'dan bağımsız).
function Import-Environment {
    Write-Log "Ortamdaki Deployment Package'lar ve SUG'lar okunuyor..."
    Import-Groups
    $rows = Invoke-CMWql 'SELECT PackageID, Name FROM SMS_SoftwareUpdatesPackage'
    $script:EnvPackages = @()
    foreach ($r in ($rows | Sort-Object { [string]$_.Name })) {
        $script:EnvPackages += [pscustomobject]@{
            Id   = [string]$r.PackageID
            Name = [string]$r.Name
        }
    }
    Update-EnvCombo
    Write-Log "$($script:EnvPackages.Count) Deployment Package, $($script:Groups.Count) SUG bulundu." 'OK'
}

# "Yapı" listesini seçili kaynağa göre doldurur.
function Update-EnvCombo {
    $mode = $cmbSource.SelectedIndex
    $cmbEnv.Items.Clear()
    if ($mode -eq 1) {
        foreach ($pk in $script:EnvPackages) { [void]$cmbEnv.Items.Add(('{0}  ({1})' -f $pk.Name, $pk.Id)) }
    }
    elseif ($mode -eq 2) {
        foreach ($gr in $script:Groups) { [void]$cmbEnv.Items.Add(('{0}  (CI_ID {1})' -f $gr.Name, $gr.Id)) }
    }
    if ($cmbEnv.Items.Count -gt 0) { $cmbEnv.SelectedIndex = 0 }
    $cmbEnv.Enabled = (($mode -eq 1 -or $mode -eq 2) -and $cmbEnv.Items.Count -gt 0)
}

# Vendor'a karşılık gelen Company kategorisini SMS_CategoryInstance'tan çözer (tek sorgu).
# Güncelleme eşleştirmesi önce CategoryInstance_UniqueIDs ile, olmazsa kategori adı ile yapılır.
function Import-VendorCategory {
    param([string]$Vendor)
    $v = $Vendor.Trim()
    $script:VendorUniqueIds = New-Object 'System.Collections.Generic.HashSet[string]'
    $script:VendorNames     = New-Object 'System.Collections.Generic.HashSet[string]'
    [void]$script:VendorNames.Add($v.ToLowerInvariant())

    $rows = @(Invoke-CMWql "SELECT CategoryInstance_UniqueID, LocalizedCategoryInstanceName FROM SMS_CategoryInstance WHERE CategoryTypeName = 'Company'")
    $allNames = @()
    $matched = @()
    foreach ($r in $rows) {
        $nm = ([string]$r.LocalizedCategoryInstanceName).Trim()
        $allNames += $nm
        if ($nm -ieq $v) { $matched += $r }
    }
    if ($matched.Count -eq 0) {
        foreach ($r in $rows) {
            if (([string]$r.LocalizedCategoryInstanceName).IndexOf($v, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $matched += $r }
        }
        if ($matched.Count -gt 0) { Write-Log "'$v' ile birebir eşleşen Company kategorisi yok; adında geçenler kullanılıyor." 'WARN' }
    }
    foreach ($r in $matched) {
        $id = ([string]$r.CategoryInstance_UniqueID).Trim()
        if ($id) { [void]$script:VendorUniqueIds.Add($id.ToLowerInvariant()) }
        [void]$script:VendorNames.Add(([string]$r.LocalizedCategoryInstanceName).Trim().ToLowerInvariant())
    }
    # Güvenilir yol: Company kategorisine bağlı CI_ID'leri doğrudan SMS_CIAllCategories'ten oku
    # (güncelleme satırındaki dizi alanlarının biçimine/doluluğuna bağımlı değil).
    $script:VendorCiIds = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($r in $matched) {
        $rawId = ([string]$r.CategoryInstance_UniqueID).Trim()
        if (-not $rawId) { continue }
        if ($rawId -notmatch '^[A-Za-z0-9:_\-\.]+$') { Write-Log "Geçersiz kategori UniqueID atlandı: '$rawId'" 'WARN'; continue }
        try {
            $ci = Invoke-CMWql "SELECT CI_ID FROM SMS_CIAllCategories WHERE CategoryInstance_UniqueID = '$rawId'"
            foreach ($c in $ci) { [void]$script:VendorCiIds.Add([string]$c.CI_ID) }
            Write-Log "SMS_CIAllCategories: '$rawId' kategorisine bağlı $($ci.Count) CI bulundu."
        } catch {
            Write-Log "SMS_CIAllCategories okunamadı: $($_.Exception.Message)" 'WARN'
        }
    }
    $shown = (@($allNames | Sort-Object -Unique | Select-Object -First 40)) -join '; '
    Write-Log "Ortamdaki Company kategorileri ($($rows.Count)): $shown"
    if ($matched.Count -eq 0) {
        Write-Log "Vendor '$v' için Company kategorisi bulunamadı. Yukarıdaki listeden birebir adı yazın veya vendor kutusunu boş bırakın." 'WARN'
    } else {
        Write-Log "Vendor eşleşmesi: $($matched.Count) Company kategorisi (UniqueID: $($script:VendorUniqueIds.Count))."
    }
}

function Add-UpdateRecord {
    param($Row, [string]$PkgId)
    # Vendor filtresi: güncellemenin Company kategorisi vendor'a karşılık gelmeli.
    # 1) CategoryInstance_UniqueIDs, Import-VendorCategory ile çözülen Company UniqueID'si içeriyor mu?
    # 2) Olmazsa LocalizedCategoryInstanceNames içinde vendor adı (büyük/küçük harf duyarsız) geçiyor mu?
    # Ek sorgu gerekmez; zaten okunan satırdan bellekte kontrol edilir.
    if ($script:VendorFilter) {
        $hit = $false
        $hasInfo = $false
        if ($null -ne $script:VendorCiIds -and $script:VendorCiIds.Contains([string]$Row.CI_ID)) { $hit = $true }
        if (-not $hit) {
        foreach ($u in @($Row.CategoryInstance_UniqueIDs)) {
            if ($null -eq $u) { continue }
            $us = ([string]$u).Trim()
            if (-not $us) { continue }
            $hasInfo = $true
            if ($null -ne $script:VendorUniqueIds -and $script:VendorUniqueIds.Contains($us.ToLowerInvariant())) { $hit = $true; break }
        }
        }
        if (-not $hit) {
            foreach ($n in @($Row.LocalizedCategoryInstanceNames)) {
                if ($null -eq $n) { continue }
                $ns = ([string]$n).Trim()
                if (-not $ns) { continue }
                $hasInfo = $true
                if ($null -ne $script:VendorNames -and $script:VendorNames.Contains($ns.ToLowerInvariant())) { $hit = $true; break }
            }
        }
        if (-not $hit) {
            if (-not $hasInfo) { $script:VendorNoInfo++ }
            if ($script:VendorSample.Count -lt 3) {
                $script:VendorSample += ('CI_ID {0} | Names=[{1}] | UniqueIDs=[{2}]' -f [string]$Row.CI_ID, ((@($Row.LocalizedCategoryInstanceNames)) -join ', '), ((@($Row.CategoryInstance_UniqueIDs)) -join ', '))
            }
            $script:VendorSkipped++
            return
        }
    }
    $key = [string]$Row.CI_ID
    if (-not $script:Updates.ContainsKey($key)) {
        $art = [string]$Row.ArticleID
        if ($art -match '^\d+$') { $kb = "KB$art" } else { $kb = $art }
        $script:Updates[$key] = [pscustomobject]@{
            Id         = $key
            KB         = $kb
            Title      = [string]$Row.LocalizedDisplayName
            Posted     = ConvertTo-DateText $Row.DatePosted
            Superseded = [bool]$Row.IsSuperseded
            Expired    = [bool]$Row.IsExpired
            PkgIds     = (New-Object 'System.Collections.Generic.List[string]')
        }
    }
    if ($PkgId) {
        $rec = $script:Updates[$key]
        if (-not $rec.PkgIds.Contains($PkgId)) { $rec.PkgIds.Add($PkgId) }
    }
}

# Listelenen güncellemelerin hangi SUG'lara üye olduğunu okur (SMS_CIRelation, RelationType=1).
# Performans: güncelleme sayısı kadar (50'lik gruplar halinde) ayrı WMI sorgusu atmak yerine
# (N+1 sorgu problemi; kapsam büyüdükçe artan sayıda round-trip demektir), TÜM RelationType=1
# ilişkileri TEK sorguda çekip bellekte SUG ID'lerine göre filtreliyoruz.
# Böylece kaç güncelleme/SUG olursa olsun tek bir WMI çağrısı yeterli oluyor.
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

# Listelenen (superseded) güncellemelerin içeriğinin hangi Software Update Deployment Package'larda
# bulunduğunu TEK çift sorguyla okur ve $script:Updates[..].PkgIds alanını doldurur.
#   1) SMS_PackageToContent  : superseded güncellemelerin içeriklerinin bulunduğu Package'lar
#   2) SMS_CIToContent       : güncelleme (CI_ID) -> ContentID eşlemesi
# Yalnızca Software Update Deployment Package'lar (SMS_SoftwareUpdatesPackage) dikkate alınır.
function Import-UpdatePackages {
    param([string]$Filter)
    $pkgSet = @{}
    foreach ($r in (Invoke-CMWql 'SELECT PackageID FROM SMS_SoftwareUpdatesPackage')) { $pkgSet[[string]$r.PackageID] = $true }

    $pkgByContent = @{}
    $rows = Invoke-CMWql "SELECT pc.* FROM SMS_PackageToContent AS pc JOIN SMS_CIToContent AS cc ON pc.ContentID = cc.ContentID JOIN SMS_SoftwareUpdate AS su ON su.CI_ID = cc.CI_ID WHERE $Filter"
    foreach ($r in $rows) {
        $pk = [string]$r.PackageID
        if (-not $pkgSet.ContainsKey($pk)) { continue }
        $ct = [string]$r.ContentID
        if (-not $pkgByContent.ContainsKey($ct)) { $pkgByContent[$ct] = New-Object 'System.Collections.Generic.HashSet[string]' }
        [void]$pkgByContent[$ct].Add($pk)
    }

    $rows = Invoke-CMWql "SELECT cc.* FROM SMS_CIToContent AS cc JOIN SMS_SoftwareUpdate AS su ON su.CI_ID = cc.CI_ID WHERE $Filter"
    foreach ($r in $rows) {
        $ci = [string]$r.CI_ID
        $ct = [string]$r.ContentID
        if (-not $script:Updates.ContainsKey($ci)) { continue }
        if (-not $pkgByContent.ContainsKey($ct)) { continue }
        $rec = $script:Updates[$ci]
        foreach ($pk in $pkgByContent[$ct]) {
            if (-not $rec.PkgIds.Contains($pk)) { $rec.PkgIds.Add($pk) }
        }
    }
}

# Deploy edilmiş (SMS_UpdateGroupAssignment) SUG'ların CI_ID'lerini TEK sorguda okur.
# AssignedCIs: deployment'ın hedeflediği CI_ID'ler (SUG deployment'ında SUG'ın CI_ID'si).
function Get-DeployedCIIds {
    $ids = @{}
    $rows = Invoke-CMWql 'SELECT * FROM SMS_UpdateGroupAssignment'
    foreach ($r in $rows) {
        foreach ($ci in @($r.AssignedCIs)) {
            if ($null -ne $ci -and "$ci" -match '^\d+$') { $ids[[string]$ci] = $true }
        }
    }
    return $ids
}

# Güncelleme deploy edilmiş bir SUG'a üyeyse (ya da doğrudan deploy edilmişse) $true döner.
function Test-UpdateDeployed {
    param([string]$UpdateId)
    if ($script:DeployedIds.ContainsKey($UpdateId)) { return $true }
    if ($script:Members.ContainsKey($UpdateId)) {
        foreach ($gid in $script:Members[$UpdateId]) {
            if ($script:DeployedIds.ContainsKey($gid)) { return $true }
        }
    }
    return $false
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
    $stateSug = @{}
    foreach ($e in $script:StateEntries) { if ($e.SugId) { $stateSug[$e.SugId] = $true } }
    $boldFont = New-Object System.Drawing.Font($gridGrp.Font, [System.Drawing.FontStyle]::Bold)
    $star = [string][char]0x2605

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
            if ($stateSug.ContainsKey($g.Id)) {
                $row.Cells['gState'].Value = $star
                $row.DefaultCellStyle.Font = $boldFont
            }
        }
    } finally {
        $script:Suspend = $prev
    }
}

function Show-Updates {
    $prev = $script:Suspend
    $script:Suspend = $true
    try {
        $gridUpd.Rows.Clear()
        foreach ($u in ($script:Updates.Values | Sort-Object KB, Title)) {
            $i = $gridUpd.Rows.Add()
            $row = $gridUpd.Rows[$i]
            $row.Tag = $u.Id
            $parts = @()
            if ($u.Superseded) { $parts += 'Superseded' }
            if ($u.Expired)    { $parts += 'Expired' }
            if ($script:DeployedIds.Count -gt 0 -and (Test-UpdateDeployed $u.Id)) { $parts += 'Deployed' }
            $row.Cells['uKB'].Value = $u.KB
            $row.Cells['uTitle'].Value = $u.Title
            $row.Cells['uPosted'].Value = $u.Posted
            $row.Cells['uStatus'].Value = ($parts -join ' + ')
            $row.Cells['uPkg'].Value = ($u.PkgIds -join ', ')
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

# Ortamdaki Software Update Deployment Package adlarını TEK sorguda okur.
function Import-PackageNames {
    $script:PkgNames = @{}
    foreach ($r in (Invoke-CMWql 'SELECT PackageID, Name FROM SMS_SoftwareUpdatesPackage')) {
        $script:PkgNames[[string]$r.PackageID] = [string]$r.Name
    }
}

# Listelenen güncellemelerin bulunduğu Deployment Package'ları sağdaki "Package'lar" sekmesine doldurur.
# "Hariç" işaretleri ($script:ExcludedPkgs) yeniden listelemede korunur.
function Show-Packages {
    $count = @{}
    foreach ($u in $script:Updates.Values) {
        foreach ($pk in $u.PkgIds) {
            if ($count.ContainsKey($pk)) { $count[$pk]++ } else { $count[$pk] = 1 }
        }
    }
    if ($count.Count -gt 0) {
        try { Import-PackageNames } catch { Write-Log "Package adları okunamadı: $($_.Exception.Message)" 'WARN' }
    }
    $prev = $script:Suspend
    $script:Suspend = $true
    try {
        $gridPkg.CurrentCell = $null
        $gridPkg.Rows.Clear()
        $keys = @($count.Keys | Sort-Object { if ($script:PkgNames.ContainsKey($_)) { $script:PkgNames[$_] } else { $_ } })
        foreach ($pkgKey in $keys) {
            $i = $gridPkg.Rows.Add()
            $row = $gridPkg.Rows[$i]
            $row.Tag = $pkgKey
            $isEx = ($script:ExcludedPkgs.ContainsKey($pkgKey) -and $script:ExcludedPkgs[$pkgKey])
            $nm = ''
            if ($script:PkgNames.ContainsKey($pkgKey)) { $nm = $script:PkgNames[$pkgKey] }
            $row.Cells['pChk'].Value = [bool]$isEx
            $row.Cells['pName'].Value = $nm
            $row.Cells['pId'].Value = $pkgKey
            $row.Cells['pSel'].Value = ''
            $row.Cells['pTotal'].Value = $count[$pkgKey]
            if ($isEx) { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen }
        }
    } finally {
        $script:Suspend = $prev
    }
    Update-PackageCounts
}

# Seçili güncellemelerden kaçının her package'ta bulunduğunu ("Seçili" kolonu) günceller.
function Update-PackageCounts {
    $sel = @(Get-SelectedUpdateIds)
    $n = $sel.Count
    foreach ($row in $gridPkg.Rows) {
        $pkgKey = [string]$row.Tag
        $m = 0
        foreach ($uid in $sel) {
            if ($script:Updates.ContainsKey($uid) -and $script:Updates[$uid].PkgIds.Contains($pkgKey)) { $m++ }
        }
        if ($n -gt 0) { $row.Cells['pSel'].Value = "$m/$n" } else { $row.Cells['pSel'].Value = '' }
    }
}

function Set-PackageFilter {
    $t = $txtPkgFilter.Text.Trim()
    $gridPkg.CurrentCell = $null
    foreach ($row in $gridPkg.Rows) {
        $hay = ([string]$row.Cells['pName'].Value) + ' ' + ([string]$row.Cells['pId'].Value)
        $visible = (-not $t) -or ($hay.IndexOf($t, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
        if ($row.Visible -ne $visible) { $row.Visible = $visible }
    }
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

        $mode = $cmbSource.SelectedIndex
        if ($mode -eq 3) { Import-Environment }
        $scope = @(Get-SelectedScope)
        if ($scope.Count -eq 0 -and $mode -ne 4 -and $mode -ne 5 -and $mode -ne 6) {
            if ($mode -eq 0) { Show-Message 'Kullanılabilir bir state kaydı yok. Önce state.json dosyasını yükleyin.' 'Uyarı' 'Warning' }
            else             { Show-Message "Önce 'Ortamı Oku' düğmesiyle ortamdaki yapıları yükleyin ve listeden birini seçin." 'Uyarı' 'Warning' }
            return
        }
        if ($mode -eq 3) { Write-Log 'Tüm ortam taranıyor; Package/SUG sayısına göre biraz sürebilir...' 'WARN' }

        Write-Log 'SUG listesi okunuyor...'
        Import-Groups
        Write-Log "$($script:Groups.Count) Software Update Group bulundu."
        Show-Groups

        $script:Updates = @{}
        $script:VendorFilter = $txtVendor.Text.Trim()
        $script:VendorSkipped = 0
        $script:VendorNoInfo = 0
        $script:VendorUniqueIds = $null
        $script:VendorNames = $null
        $script:VendorCiIds = $null
        $script:VendorSample = @()
        if ($script:VendorFilter) {
            Write-Log "Vendor filtresi: '$($script:VendorFilter)'"
            try { Import-VendorCategory -Vendor $script:VendorFilter }
            catch { Write-Log "Company kategorisi okunamadı, yalnızca ad eşleşmesi kullanılacak: $($_.Exception.Message)" 'WARN'
                    $script:VendorNames = New-Object 'System.Collections.Generic.HashSet[string]'
                    [void]$script:VendorNames.Add($script:VendorFilter.ToLowerInvariant()) }
        } else { Write-Log 'Vendor filtresi yok (tüm vendor''lar).' }
        if ($chkExpired.Checked) { $flt = '(su.IsSuperseded = 1 OR su.IsExpired = 1)' }
        else                     { $flt = 'su.IsSuperseded = 1' }

        $script:DeployedIds = @{}

        # Ortamdaki TÜM superseded güncellemeler (mod 4, 5 ve 6): Package/SUG gezmeden TEK sorgu.
        # Mod 4: deploy filtresi, üyelikler okunduktan sonra aşağıda bellekte uygulanır.
        # Mod 5: Package filtresi burada uygulanır (içeriği hiçbir Deployment Package'ta olmayanlar elenir);
        #        ardından Mod 4 ile aynı deploy filtresi de uygulanır (indirilmiş + superseded + deploy edilmiş).
        # Mod 6: Mod 5'in yalnızca Package filtresi uygulanır; deploy filtresi YOKTUR (indirilmiş + superseded).
        if ($mode -eq 4 -or $mode -eq 5 -or $mode -eq 6) {
            Write-Log 'Ortamdaki superseded güncellemeler tek sorguyla okunuyor...'
            $rows = Invoke-CMWql "SELECT * FROM SMS_SoftwareUpdate AS su WHERE $flt"
            foreach ($r in $rows) { Add-UpdateRecord -Row $r -PkgId $null }
            Write-Log "$($script:Updates.Count) superseded güncelleme okundu."

            if ($mode -eq 5 -or $mode -eq 6) {
                Write-Log 'Deployment Package içerikleri eşleştiriliyor (indirilmiş olanlar aranıyor)...'
                Import-UpdatePackages -Filter $flt
                $before = $script:Updates.Count
                foreach ($k in @($script:Updates.Keys)) {
                    if ($script:Updates[$k].PkgIds.Count -eq 0) { $script:Updates.Remove($k) }
                }
                Write-Log "Package filtresi: $($script:Updates.Count) güncelleme kaldı ($($before - $script:Updates.Count) hiçbir Package'ta indirilmemiş elendi)."
            }
        }
        else {
            $pkgIds = @($scope | ForEach-Object { $_.PkgId } | Where-Object { $_ } | Select-Object -Unique)
            $sugIds = @($scope | ForEach-Object { $_.SugId } | Where-Object { $_ } | Select-Object -Unique)

            foreach ($pkg in $pkgIds) {
                if ($pkg -notmatch '^[A-Za-z0-9]+$') { Write-Log "Geçersiz PackageID atlandı: '$pkg'" 'WARN'; continue }
                Write-Log "Package içeriği taranıyor: $pkg"
                $q = "SELECT su.* FROM SMS_SoftwareUpdate AS su JOIN SMS_CIToContent AS cc ON su.CI_ID = cc.CI_ID JOIN SMS_PackageToContent AS pc ON pc.ContentID = cc.ContentID WHERE pc.PackageID = '$pkg' AND $flt"
                $rows = Invoke-CMWql $q
                foreach ($r in $rows) { Add-UpdateRecord -Row $r -PkgId $pkg }
                Write-Log "  $($pkg): $($rows.Count) satır (aynı güncellemenin birden fazla içeriği tekrar sayılabilir)."
            }

            foreach ($sug in $sugIds) {
                if ($sug -notmatch '^\d+$') { Write-Log "Geçersiz SUG CI_ID atlandı: '$sug'" 'WARN'; continue }
                $gname = $sug
                if ($script:GroupById.ContainsKey($sug)) { $gname = $script:GroupById[$sug].Name }
                Write-Log "State SUG üyeleri taranıyor: $gname ($sug)"
                $q = "SELECT su.* FROM SMS_SoftwareUpdate AS su JOIN SMS_CIRelation AS cr ON su.CI_ID = cr.ToCIID WHERE cr.FromCIID = $sug AND cr.RelationType = 1 AND $flt"
                $rows = Invoke-CMWql $q
                foreach ($r in $rows) { Add-UpdateRecord -Row $r -PkgId $null }
                Write-Log "  $gname : $($rows.Count) güncelleme."
            }
        }

        if ($script:VendorFilter) {
            Write-Log "Vendor filtresi: $($script:Updates.Count) güncelleme kaldı ($($script:VendorSkipped) kayıt vendor '$($script:VendorFilter)' dışında olduğu için elendi)."
            if ($script:VendorNoInfo -gt 0) {
                Write-Log "$($script:VendorNoInfo) kayıtta kategori bilgisi (CategoryInstance_UniqueIDs / LocalizedCategoryInstanceNames) BOŞ geldi; bu kayıtlar vendor filtresinde elendi. Çok sayıdaysa sorgu bu alanları döndürmüyor demektir." 'WARN'
            }
            if ($script:Updates.Count -eq 0 -and $script:VendorSample.Count -gt 0) {
                Write-Log 'Eşleşmeyen kayıtlardan örnek kategori değerleri (teşhis):' 'WARN'
                foreach ($s in $script:VendorSample) { Write-Log "  $s" 'WARN' }
            }
            if ($script:Updates.Count -eq 0 -and $script:VendorSkipped -gt 0) {
                Write-Log "Vendor '$($script:VendorFilter)' ile eşleşen güncelleme yok. Vendor adı 'Company' kategorisiyle birebir aynı olmalı (örn. Microsoft); tüm vendor'lar için kutuyu boş bırakın." 'WARN'
            }
        }

        Write-Log 'SUG üyelikleri okunuyor...'
        Import-Memberships

        if ($mode -eq 4 -or $mode -eq 5) {
            Write-Log 'Deploy edilmiş SUG''lar okunuyor (SMS_UpdateGroupAssignment)...'
            $script:DeployedIds = Get-DeployedCIIds
            Write-Log "$($script:DeployedIds.Count) deploy hedefi (SUG) bulundu."
            if ($script:DeployedIds.Count -eq 0) {
                Write-Log 'Hiç deployment bulunamadı; liste boş olacak. Yetkinizi ve SMS_UpdateGroupAssignment erişimini kontrol edin.' 'WARN'
            }
            $before = $script:Updates.Count
            foreach ($k in @($script:Updates.Keys)) {
                if (-not (Test-UpdateDeployed $k)) {
                    $script:Updates.Remove($k)
                    $script:Members.Remove($k)
                }
            }
            Write-Log "Deploy filtresi: $($script:Updates.Count) güncelleme kaldı ($($before - $script:Updates.Count) deploy edilmemiş elendi)."
        }

        Show-Updates
        Show-Packages
        Set-GroupFilter
        Set-PackageFilter
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

    $simulate = $chkSim.Checked
    $lines = @()
    foreach ($p in $plan) {
        $lines += ('- {0}:  +{1} ekle,  -{2} çıkar' -f $script:GroupById[$p.GroupId].Name, $p.Add.Count, $p.Remove.Count)
    }
    if ($simulate) { $head = "SİMÜLASYON - hiçbir değişiklik yapılmayacak.`n`n" } else { $head = "" }
    $msg = "{0}{1} güncelleme için aşağıdaki SUG değişiklikleri uygulanacak:`n`n{2}`n`nDevam edilsin mi?" -f $head, $sel.Count, ($lines -join "`n")
    if ([System.Windows.Forms.MessageBox]::Show($msg, 'SUG üyeliklerini uygula', 'YesNo', 'Question') -ne 'Yes') { return }

    Set-Busy $true
    $pushed = $false
    $failed = 0
    try {
        if (-not $simulate) {
            if (-not (Get-Command -Name Add-CMSoftwareUpdateToGroup -ErrorAction SilentlyContinue)) { Import-CMModule }
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
                    Add-CMSoftwareUpdateToGroup -SoftwareUpdateGroup $sug -SoftwareUpdateId $p.Add -ErrorAction Stop
                    Write-Log "'$gname': $($p.Add.Count) güncelleme eklendi." 'OK'
                }
                if ($p.Remove.Count -gt 0) {
                    Remove-CMSoftwareUpdateFromGroup -SoftwareUpdateGroup $sug -SoftwareUpdateId $p.Remove -Force -ErrorAction Stop
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
# Package içeriği silme yardımcıları
# ------------------------------------------------------------------
# Package'ın (SMS_PackageToContent) ŞU AN içerdiği ContentID'leri döndürür (gerçek durum kaynağı).
function Get-PackageContentIdSet {
    param([string]$PkgId)
    $h = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($r in (Invoke-CMWql "SELECT ContentID FROM SMS_PackageToContent WHERE PackageID = '$PkgId'")) {
        [void]$h.Add([string]$r.ContentID)
    }
    return ,$h
}

# RemoveContent'i, parametre adlarını WMI'den okuyarak (adlara/sıraya güvenmeden) çağırır.
function Invoke-PkgRemoveContent {
    param($Pkg, [uint32[]]$Ids)
    $in = $Pkg.GetMethodParameters('RemoveContent')
    $arrProp = $null
    $boolProp = $null
    $desc = @()
    foreach ($pr in $in.Properties) {
        $desc += ('{0}:{1}{2}' -f $pr.Name, $pr.Type, $(if ($pr.IsArray) { '[]' } else { '' }))
        if ($pr.IsArray -and -not $arrProp) { $arrProp = $pr.Name }
        elseif ((-not $pr.IsArray) -and ($pr.Type -eq [System.Management.CimType]::Boolean) -and -not $boolProp) { $boolProp = $pr.Name }
    }
    Write-Log ("RemoveContent parametreleri: " + ($desc -join ', ')) 'INFO'
    if (-not $arrProp) { throw 'RemoveContent için ID dizisi parametresi bulunamadı.' }
    $in.Properties[$arrProp].Value = $Ids
    if ($boolProp) { $in.Properties[$boolProp].Value = $true }
    $out = $Pkg.InvokeMethod('RemoveContent', $in, $null)
    $rv = $null
    if ($out) { foreach ($pr in $out.Properties) { if ($pr.Name -eq 'ReturnValue') { $rv = $pr.Value } } }
    if ($null -ne $rv -and [int]$rv -ne 0) { throw "RemoveContent başarısız (dönüş kodu: $rv)." }
    return $rv
}

# Verilen ContentID'ler package'tan kalkana kadar bekler (en fazla ~Tries*DelaySec sn). Geriye KALAN sayısını döner.
function Wait-ContentRemoved {
    param([string]$PkgId, [uint32[]]$ContentIds, [int]$Tries = 6, [int]$DelaySec = 2)
    $left = 0
    for ($i = 1; $i -le $Tries; $i++) {
        $cur = Get-PackageContentIdSet $PkgId
        $left = @($ContentIds | Where-Object { $cur.Contains([string]$_) }).Count
        if ($left -eq 0) { return 0 }
        if ($i -lt $Tries) {
            Start-Sleep -Seconds $DelaySec
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    return $left
}

# İçeriği siler ve SONUCU DOĞRULAR. Önce ContentID ile dener; içerik hâlâ duruyorsa ve güvenliyse
# güncellemelerin CI_ID'leri ile tekrar dener (RemoveContent'in ID türü sürüme göre farklı yorumlanabilir).
# ID çakışması (bir ContentID'nin başka bir güncellemenin CI_ID'si ile aynı sayı olması) durumunda
# yanlış içeriğin silinmemesi için ilgili ID'ler atlanır.
function Remove-PackageContentVerified {
    param($Pkg, [string]$PkgId, [uint32[]]$ContentIds, [string[]]$CiIds, [object[]]$Pairs)

    $selSet = @{}
    foreach ($s in $CiIds) { $selSet[[string]$s] = $true }
    $ciAll = New-Object 'System.Collections.Generic.HashSet[string]'
    $contentsByCi = @{}
    foreach ($pr in $Pairs) {
        $ciKey = [string]$pr.CI_ID
        $ctKey = [string]$pr.ContentID
        [void]$ciAll.Add($ciKey)
        if (-not $contentsByCi.ContainsKey($ciKey)) { $contentsByCi[$ciKey] = New-Object 'System.Collections.Generic.List[string]' }
        $contentsByCi[$ciKey].Add($ctKey)
    }
    $removeSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($c in $ContentIds) { [void]$removeSet.Add([string]$c) }

    $method = '-'

    # --- 1. deneme: ContentID ile ---
    $targets = @()
    $skipped = 0
    foreach ($c in $ContentIds) {
        $cs = [string]$c
        if ($ciAll.Contains($cs) -and -not $selSet.ContainsKey($cs)) { $skipped++; continue }
        $targets += [uint32]$c
    }
    if ($skipped -gt 0) {
        Write-Log "[$PkgId] $skipped içerik ID'si başka bir güncellemenin CI_ID'si ile çakıştığı için güvenlik amacıyla atlandı." 'WARN'
    }
    $left = $targets.Count
    if ($targets.Count -gt 0) {
        $rv = Invoke-PkgRemoveContent -Pkg $Pkg -Ids ([uint32[]]$targets)
        Write-Log "[$PkgId] RemoveContent (ContentID ile, $($targets.Count) içerik) çağrıldı. Dönüş: [$rv]. Sonuç doğrulanıyor..."
        $left = Wait-ContentRemoved -PkgId $PkgId -ContentIds ([uint32[]]$targets)
        if ($left -eq 0) { $method = 'ContentID' }
    }

    # --- 2. deneme: CI_ID ile (yalnızca 1. deneme etkisiz kaldıysa) ---
    if ($targets.Count -gt 0 -and $left -gt 0) {
        Write-Log "[$PkgId] ContentID ile çağrı hata vermedi ama $left içerik hâlâ package'ta. CI_ID ile deneniyor..." 'WARN'
        $inPkg = Get-PackageContentIdSet $PkgId
        $cand = @()
        $candContents = @()
        foreach ($ciKey in $CiIds) {
            if (-not $contentsByCi.ContainsKey($ciKey)) { continue }
            $mine = $true
            $still = $false
            foreach ($ctKey in $contentsByCi[$ciKey]) {
                if (-not $removeSet.Contains($ctKey)) { $mine = $false }
                if ($inPkg.Contains($ctKey)) { $still = $true }
            }
            if (-not $mine -or -not $still) { continue }     # paylaşılan içerik / zaten silinmiş
            if ($inPkg.Contains($ciKey)) {
                Write-Log "[$PkgId] CI_ID $ciKey, package'taki bir ContentID ile çakıştığı için atlandı." 'WARN'
                continue
            }
            $cand += [uint32]$ciKey
            foreach ($ctKey in $contentsByCi[$ciKey]) { $candContents += [uint32]$ctKey }
        }
        if ($cand.Count -gt 0) {
            $rv = Invoke-PkgRemoveContent -Pkg $Pkg -Ids ([uint32[]]$cand)
            Write-Log "[$PkgId] RemoveContent (CI_ID ile, $($cand.Count) güncelleme) çağrıldı. Dönüş: [$rv]. Sonuç doğrulanıyor..."
            [void](Wait-ContentRemoved -PkgId $PkgId -ContentIds ([uint32[]]@($candContents | Select-Object -Unique)))
            $method = 'CI_ID'
        } else {
            Write-Log "[$PkgId] CI_ID ile denenebilecek güvenli aday bulunamadı." 'WARN'
        }
    }

    # --- Nihai durum: package'ta gerçekte ne kaldı? ---
    $cur = Get-PackageContentIdSet $PkgId
    $remaining = @($ContentIds | Where-Object { $cur.Contains([string]$_) }).Count
    return [pscustomobject]@{ Total = $ContentIds.Count; Remaining = $remaining; Method = $method }
}

# Paylaşılan içeriklerin (seçili olmayan güncellemelerle ortak) "sahiplerini" SMS_SoftwareUpdate'ten çözer.
# Bir sahip, AKTİF (superseded/expired değil) veya deploy edilmiş ise o içeriğin silinmesine ENGEL sayılır.
# Superseded/Expired ve deploy edilmemiş sahipler engel sayılmaz (içerik onlar için de gereksiz).
function Get-SharedContentDetails {
    param([object[]]$Pairs, [string[]]$SelectedIds, [uint32[]]$SharedContentIds)

    $sel = @{}
    foreach ($x in $SelectedIds) { $sel[[string]$x] = $true }
    $shared = @{}
    foreach ($c in $SharedContentIds) { $shared[[string]$c] = $true }

    $othersByContent = @{}
    $allOthers = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($p in $Pairs) {
        $c  = [string]$p.ContentID
        $ci = [string]$p.CI_ID
        if (-not $shared.ContainsKey($c) -or $sel.ContainsKey($ci)) { continue }
        if (-not $othersByContent.ContainsKey($c)) { $othersByContent[$c] = New-Object 'System.Collections.Generic.HashSet[string]' }
        [void]$othersByContent[$c].Add($ci)
        [void]$allOthers.Add($ci)
    }

    $meta = @{}
    $list = @($allOthers | Where-Object { $_ -match '^\d+$' })
    for ($i = 0; $i -lt $list.Count; $i += 40) {
        $last = [Math]::Min($i + 39, $list.Count - 1)
        $inList = ($list[$i..$last]) -join ','
        foreach ($r in (Invoke-CMWql "SELECT * FROM SMS_SoftwareUpdate WHERE CI_ID IN ($inList)")) { $meta[[string]$r.CI_ID] = $r }
    }

    $free = @()
    $blocked = @{}
    $lines = @()
    foreach ($c in @($othersByContent.Keys)) {
        $blockers = @()
        foreach ($ci in $othersByContent[$c]) {
            $r = $meta[$ci]
            if ($null -eq $r) {
                $lines += ("  ContentID {0} <- CI_ID {1}: SMS_SoftwareUpdate'te kayıt yok (silinmiş veya güncelleme değil); engel sayılmadı." -f $c, $ci)
                continue
            }
            $sup = [bool]$r.IsSuperseded
            $exp = [bool]$r.IsExpired
            $dep = [bool]$r.IsDeployed
            $st = @()
            if ($sup) { $st += 'Superseded' }
            if ($exp) { $st += 'Expired' }
            if (-not $sup -and -not $exp) { $st += 'AKTİF' }
            if ($dep) { $st += 'Deploy edilmiş' }
            $block = ($dep -or (-not $sup -and -not $exp))
            $title = [string]$r.LocalizedDisplayName
            $lines += ("  ContentID {0} <- CI_ID {1} | {2} | {3} | {4}" -f $c, $ci, ($st -join ', '), $(if ($block) { 'ENGEL' } else { 'engel değil' }), $title)
            if ($block) { $blockers += ("CI_ID {0}: {1} [{2}]" -f $ci, $title, ($st -join ', ')) }
        }
        if ($blockers.Count -gt 0) { $blocked[$c] = $blockers } else { $free += [uint32]$c }
    }
    return [pscustomobject]@{ Free = [uint32[]]$free; Blocked = $blocked; Lines = $lines }
}

function Invoke-RemoveFromPackage {
    $sel = @(Get-SelectedUpdateIds)
    if ($sel.Count -eq 0) { Show-Message 'Önce soldaki listeden en az bir güncelleme seçin.' 'Uyarı' 'Warning'; return }

    # "Hariç" işaretli package'lar: değer = bu package'tan silinmeyecek seçili güncelleme sayısı
    $excl = @{}
    foreach ($k in @($script:ExcludedPkgs.Keys)) { if ($script:ExcludedPkgs[$k]) { $excl[[string]$k] = 0 } }

    $byPkg = @{}
    $stillInSug = 0
    $noPkg = 0
    foreach ($id in $sel) {
        $u = $script:Updates[$id]
        if (-not $u -or $u.PkgIds.Count -eq 0) { $noPkg++; continue }
        $added = $false
        foreach ($p in $u.PkgIds) {
            if ($excl.ContainsKey($p)) { $excl[$p]++; continue }
            if (-not $byPkg.ContainsKey($p)) { $byPkg[$p] = New-Object 'System.Collections.Generic.List[string]' }
            $byPkg[$p].Add($id)
            $added = $true
        }
        if ($added -and $script:Members.ContainsKey($id) -and $script:Members[$id].Count -gt 0) { $stillInSug++ }
    }
    $exTotal = 0
    foreach ($k in @($excl.Keys)) { $exTotal += $excl[$k] }
    if ($byPkg.Count -eq 0) {
        if ($exTotal -gt 0) {
            Show-Message "Seçili güncellemelerin bulunduğu tüm package'lar 'Hariç' olarak işaretli; silinecek bir şey yok." 'Bilgi' 'Information'
        } else {
            Show-Message "Seçili güncellemelerin hiçbiri state'teki Deployment Package içinde değil (yalnızca SUG üyesi)." 'Uyarı' 'Warning'
        }
        return
    }

    $simulate = $chkSim.Checked
    $lines = @()
    foreach ($p in $byPkg.Keys) { $lines += ('- {0}: {1} güncelleme' -f $p, $byPkg[$p].Count) }
    $msg = ''
    if ($simulate) { $msg += "SİMÜLASYON - hiçbir değişiklik yapılmayacak.`n`n" }
    $msg += "Aşağıdaki Deployment Package'lardan seçili güncellemelerin İÇERİĞİ silinecek:`n`n" + ($lines -join "`n")
    if ($noPkg -gt 0) { $msg += "`n`n$noPkg seçili güncelleme package içinde olmadığı için atlanacak." }
    if ($exTotal -gt 0) {
        $exLines = @()
        foreach ($k in @($excl.Keys)) {
            if ($excl[$k] -gt 0) {
                $nm = $k
                if ($script:PkgNames.ContainsKey($k)) { $nm = "$k ($($script:PkgNames[$k]))" }
                $exLines += ('- {0}: {1} güncelleme' -f $nm, $excl[$k])
            }
        }
        $msg += "`n`nHARİÇ TUTULANLAR (bu package'lardan SİLİNMEYECEK):`n" + ($exLines -join "`n")
    }
    if ($stillInSug -gt 0) {
        $msg += "`n`nUYARI: $stillInSug güncelleme hâlâ en az bir SUG'a üye. Önce SUG üyeliklerinden çıkarmanız önerilir."
    }
    if (-not $simulate) {
        $msg += "`n`nBu işlem geri alınamaz; dağıtım noktalarına güncelleme tetiklenir. Devam edilsin mi?"
    } else {
        $msg += "`n`nDevam edilsin mi?"
    }
    if ([System.Windows.Forms.MessageBox]::Show($msg, "Package'tan sil", 'YesNo', 'Warning', 'Button2') -ne 'Yes') { return }

    Set-Busy $true
    $changed = $false
    $failedPkgs = 0
    try {
        foreach ($pkgId in @($byPkg.Keys)) {
            if ($pkgId -notmatch '^[A-Za-z0-9]+$') { Write-Log "Geçersiz PackageID atlandı: '$pkgId'" 'WARN'; continue }
            $ids = $byPkg[$pkgId].ToArray()

            Write-Log "[$pkgId] içerik eşleşmeleri okunuyor..."
            $pairs = Invoke-CMWql "SELECT cc.CI_ID, cc.ContentID FROM SMS_CIToContent AS cc JOIN SMS_PackageToContent AS pc ON pc.ContentID = cc.ContentID WHERE pc.PackageID = '$pkgId'"
            $set = Get-ContentRemovalSet -Pairs $pairs -SelectedIds $ids
            $totalContents = @($pairs | ForEach-Object { [string]$_.ContentID } | Select-Object -Unique).Count
            Write-Log "[$pkgId] $($ids.Count) seçili güncelleme | package'ta eşleşen içerik: $totalContents | doğrudan silinebilir: $($set.Remove.Count) | paylaşılan: $($set.Shared.Count)"

            if ($set.Shared.Count -gt 0) {
                $det = $null
                try { $det = Get-SharedContentDetails -Pairs $pairs -SelectedIds $ids -SharedContentIds $set.Shared }
                catch { Write-Log "[$pkgId] paylaşan güncellemelerin ayrıntıları okunamadı: $($_.Exception.Message)" 'WARN' }

                if ($null -eq $det) {
                    Write-Log "[$pkgId] $($set.Shared.Count) içerik, seçili olmayan başka güncellemelerle paylaşıldığı için ATLANDI." 'WARN'
                } else {
                    Write-Log "[$pkgId] Paylaşılan içerikler (seçili olmayan güncellemelerle ortak):" 'WARN'
                    foreach ($ln in @($det.Lines | Select-Object -First 30)) { Write-Log $ln 'WARN' }
                    if ($det.Lines.Count -gt 30) { Write-Log "  ... ve $($det.Lines.Count - 30) satır daha." 'WARN' }

                    if ($det.Free.Count -gt 0) {
                        Write-Log "[$pkgId] $($det.Free.Count) paylaşılan içeriğin diğer sahipleri superseded/expired ve deploy edilmemiş; silinecekler arasına eklendi." 'INFO'
                        $set.Remove = [uint32[]](@($set.Remove) + @($det.Free))
                    }
                    $blockedKeys = @($det.Blocked.Keys)
                    if ($blockedKeys.Count -gt 0) {
                        Write-Log "[$pkgId] $($blockedKeys.Count) içerik AKTİF veya deploy edilmiş güncellemelerle paylaşıldığı için atlandı." 'WARN'
                        if (-not $simulate) {
                            $bl = @()
                            foreach ($k in $blockedKeys) { $bl += $det.Blocked[$k] }
                            $bl = @($bl | Select-Object -Unique)
                            $txt = "[$pkgId] $($blockedKeys.Count) içerik, aşağıdaki aktif/deploy edilmiş güncellemelerle paylaşılıyor:`n`n" +
                                   (($bl | Select-Object -First 8) -join "`n")
                            if ($bl.Count -gt 8) { $txt += "`n... ve $($bl.Count - 8) güncelleme daha." }
                            $txt += "`n`nBu içeriği silerseniz bu güncellemeler bu package'ta içeriksiz kalır. Yine de silinsin mi?"
                            if ([System.Windows.Forms.MessageBox]::Show($txt, 'Paylaşılan içerik', 'YesNo', 'Warning', 'Button2') -eq 'Yes') {
                                $set.Remove = [uint32[]](@($set.Remove) + @($blockedKeys | ForEach-Object { [uint32]$_ }))
                                Write-Log "[$pkgId] Kullanıcı onayıyla paylaşılan $($blockedKeys.Count) içerik de silinecek." 'WARN'
                            }
                        }
                    }
                }
            }
            if ($set.Remove.Count -eq 0) {
                if ($set.Shared.Count -gt 0) {
                    Write-Log "[$pkgId] seçili güncellemelerin tüm içeriği başka güncellemelerle paylaşıldığı için bu package atlandı." 'WARN'
                } else {
                    Write-Log "[$pkgId] silinecek içerik bulunamadı (seçili güncellemelerin bu package'ta içerik eşleşmesi yok)." 'WARN'
                }
                continue
            }
            if ($simulate) {
                Write-Log "[SİMÜLASYON] [$pkgId] $($ids.Count) güncellemeye ait $($set.Remove.Count) içerik silinecekti." 'WARN'
                continue
            }

            $pkg = Get-WmiObject -ComputerName (Get-Connection).Server -Namespace "root\SMS\site_$((Get-Connection).Site)" `
                -Class SMS_SoftwareUpdatesPackage -Filter "PackageID='$pkgId'" -ErrorAction Stop
            if (-not $pkg) { throw "Deployment Package bulunamadı: $pkgId" }
            if ($pkg -is [array]) { $pkg = $pkg[0] }

            Write-Log "[$pkgId] $($set.Remove.Count) içerik siliniyor..."
            $res = Remove-PackageContentVerified -Pkg $pkg -PkgId $pkgId -ContentIds $set.Remove -CiIds $ids -Pairs $pairs
            if (($res.Total - $res.Remaining) -gt 0) { $changed = $true }
            if ($res.Remaining -eq 0) {
                Write-Log "[$pkgId] $($res.Total) içerik package'tan silindi ve DOĞRULANDI (yöntem: $($res.Method))." 'OK'
            } else {
                $failedPkgs++
                Write-Log "[$pkgId] $($res.Remaining)/$($res.Total) içerik HÂLÂ package'ta duruyor. Dağıtım Yöneticisi işlemi sürdürüyor olabilir; birkaç dakika sonra 'Listele' ile tekrar kontrol edin. Kalıyorsa SMS Provider sunucusunda smsprov.log ve site sunucusunda distmgr.log kayıtlarına bakın." 'ERROR'
            }
        }
        if ($simulate) { Show-Message 'Simülasyon tamamlandı. Ayrıntılar günlükte.' }
        elseif ($failedPkgs -gt 0) { Show-Message "$failedPkgs package'ta içerik silinemedi (hâlâ duruyor). Ayrıntılar günlükte." 'Uyarı' 'Warning' }
        else { Show-Message "Silme işlemi tamamlandı ve doğrulandı. Liste yenileniyor." }
    }
    catch {
        Write-Log $_.Exception.Message 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    }
    finally {
        Set-Busy $false
    }

    if ($changed) { Invoke-List }
}

# ------------------------------------------------------------------
# Olaylar
# ------------------------------------------------------------------
$btnList.Add_Click({ Invoke-List })
$btnApply.Add_Click({ Invoke-ApplyMembership })
$btnRemove.Add_Click({ Invoke-RemoveFromPackage })
$btnClose.Add_Click({ $form.Close() })
$btnSelAll.Add_Click({ $gridUpd.SelectAll() })
$btnSelNone.Add_Click({ $gridUpd.ClearSelection() })

$btnReload.Add_Click({ Import-StateFile (Find-StateFile) })
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'state.json|state.json|JSON dosyaları (*.json)|*.json|Tüm dosyalar (*.*)|*.*'
    if ($dlg.ShowDialog() -eq 'OK') { Import-StateFile $dlg.FileName }
})

$cmbState.Add_SelectedIndexChanged({
    $i = $cmbState.SelectedIndex
    if ($i -ge 0 -and $i -lt $script:StateEntries.Count) {
        $e = $script:StateEntries[$i]
        if ($e.Site)   { $txtSite.Text = $e.Site }
        if ($e.Server) { $txtServer.Text = $e.Server }
    }
})

$cmbSource.Add_SelectedIndexChanged({
    $cmbState.Enabled = ($cmbSource.SelectedIndex -eq 0)
    $cmbEnv.Enabled = ($cmbSource.SelectedIndex -eq 1 -or $cmbSource.SelectedIndex -eq 2)
    $btnEnv.Enabled = ($cmbSource.SelectedIndex -eq 1 -or $cmbSource.SelectedIndex -eq 2 -or $cmbSource.SelectedIndex -eq 3)
    Update-EnvCombo
})
$btnEnv.Add_Click({
    Set-Busy $true
    try {
        [void](Get-Connection)
        Import-Environment
        if ($cmbSource.SelectedIndex -eq 0) {
            Write-Log "Listelemek için 'Kaynak' olarak Package veya SUG seçin." 'WARN'
        }
    }
    catch {
        Write-Log $_.Exception.Message 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    }
    finally {
        Set-Busy $false
    }
})

$txtFilter.Add_TextChanged({ Set-GroupFilter })

$gridUpd.Add_SelectionChanged({
    if (-not $script:Suspend) { Update-GroupChecks; Update-PackageCounts }
})

$txtPkgFilter.Add_TextChanged({ Set-PackageFilter })

# Package "Hariç" kutusu: tıklama anında işlensin ve işaret $script:ExcludedPkgs içinde saklansın
$gridPkg.Add_CurrentCellDirtyStateChanged({
    if ($gridPkg.IsCurrentCellDirty) {
        [void]$gridPkg.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})
$gridPkg.Add_CellValueChanged({
    param($sender, $e)
    if ($script:Suspend -or $e.RowIndex -lt 0 -or $e.ColumnIndex -ne 0) { return }
    $row = $gridPkg.Rows[$e.RowIndex]
    $pkgKey = [string]$row.Tag
    $isEx = ($row.Cells['pChk'].Value -eq $true)
    $script:ExcludedPkgs[$pkgKey] = $isEx
    if ($isEx) { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen }
    else       { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Empty }
})
$gridPkg.Add_DataError({ param($sender, $e) $e.ThrowException = $false })

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

# ------------------------------------------------------------------
# Açılış
# ------------------------------------------------------------------
Write-Log 'Hazır. State kaydını seçip "Listele" düğmesine basın. Varsayılan olarak SİMÜLASYON açıktır.'
Import-StateFile (Find-StateFile)
$lblCount.Text = 'Henüz liste yüklenmedi.'

[void]$form.ShowDialog()