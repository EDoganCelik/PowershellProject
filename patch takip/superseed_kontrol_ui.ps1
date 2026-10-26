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
    - Arka plan işleri: "SUG Üyeliklerini Uygula" ve "Seçilileri Package'tan Sil" işlemleri arka plan
      Runspace'lerinde çalışır; ekran kilitlenmez. İlerleme $script:PackageJobs (senkronize hashtable) ve
      bir System.Windows.Forms.Timer ile izlenir.
    - Kısmi kilitleme: yalnızca işleme alınan güncellemeler (gri satır, "KİLİTLİ" durumu), package'lar ve
      SUG'lar kilitlenir. Kilitli satırlar yeni bir işlemin seçimine dahil edilmez; diğer güncellemelerle
      çalışmaya devam edilebilir. Arka planda iş varken yalnızca "Listele" / "Ortamı Oku" / state okuma kapalıdır.
    - İşlem sonrası tüm liste WMI'dan yeniden çekilmez: ilgili satırlar, Package ve SUG sayıları bellekte
      güncellenir; liste filtresine artık uymayan güncellemeler ekrandan kaldırılır.
    - "Listele" ve "Ortamı Oku" da arka planda çalışır; listeleme sürerken pencere donmaz.
    - Güncelleme listesi VirtualMode'dadır: hücreler yalnızca ekranda göründükçe hesaplanır (çok büyük listelerde hızlı).
    - Denetim kaydı: geri alınamaz işlemler (Package'tan silme, SUG değişikliği) %LOCALAPPDATA%\update-state\audit.log
      dosyasına JSON satırları olarak yazılır. Program açılırken önceki oturumdan yarım kalmış işlemler varsa uyarılır.

.NOTES
    Windows PowerShell 5.1 ile çalıştırın (PowerShell 7 desteklenmez).
    TÜM SCCM işlemleri SMS Provider WMI (root\SMS\site_XXX) üzerinden, DCOM ile yapılır (PacketPrivacy: şifreli).
    CIM/WinRM, PowerShell Remoting (Invoke-Command / Enter-PSSession) ve ConfigurationManager modülü kullanılmaz;
    bu yüzden Configuration Manager Console kurulumu gerekmez.
    SUG üyeliği değişiklikleri SMS_AuthorizationList.Updates + Put() ile yapılır; SUG düzenleme yetkisi gerekir.
    Package'tan silme için Software Update Package düzenleme yetkisi gerekir.
    Önerilen sıra: önce SUG üyeliklerinden çıkarın, sonra Package'tan silin.
#>

#Requires -Version 5.1
#Requires -PSEdition Desktop

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Management
[System.Windows.Forms.Application]::EnableVisualStyles()

# ------------------------------------------------------------------
# Genel durum
# ------------------------------------------------------------------
$script:StateFile    = $null
$script:StateEntries = @()      # state.json'dan gelen kayıtlar (combobox ile aynı sırada)
$script:Groups       = New-Object 'System.Collections.Generic.List[object]'   # ortamdaki tüm SUG'lar
$script:GroupById    = @{}      # SUG CI_ID (string) -> SUG
$script:Updates      = @{}      # güncelleme CI_ID (string) -> kayıt
$script:Members      = @{}      # güncelleme CI_ID (string) -> HashSet[string] (üye olduğu SUG CI_ID'leri)
$script:OrigCheck    = @{}      # SUG CI_ID (string) -> kutunun ilk hesaplanan durumu
$script:Suspend      = $false   # UI olaylarını geçici olarak susturmak için
$script:EnvPackages  = New-Object 'System.Collections.Generic.List[object]'   # ortamdaki tüm Software Update Deployment Package'lar (Id, Name) - önbellek
$script:DeployedIds  = @{}      # deploy edilmiş SUG (ve varsa doğrudan güncelleme) CI_ID'leri (string -> $true)
$script:ExcludedPkgs = @{}      # silmeden HARİÇ tutulacak PackageID'ler (string -> bool)
$script:PkgNames     = @{}      # PackageID (string) -> package adı (Import-PackageCatalog önbelleği)
$script:VendorFilter = ''       # Listele sırasında uygulanan vendor (Company kategorisi); boşsa filtre yok
$script:VendorSkipped = 0       # vendor filtresiyle elenen kayıt sayacı
$script:VendorNoInfo  = 0       # kategori bilgisi (UniqueIDs/Names) boş gelen kayıt sayacı
$script:VendorUniqueIds = $null # vendor'a karşılık gelen Company kategorisinin CategoryInstance_UniqueID'leri (küçük harf)
$script:VendorNames     = $null # vendor'a karşılık gelen Company kategori adları (küçük harf)
$script:VendorCiIds     = $null # Company kategorisine bağlı güncelleme CI_ID'leri (SMS_CIAllCategories'ten)
$script:VendorSample    = New-Object 'System.Collections.Generic.List[string]'   # eşleşmeyen kayıtlardan teşhis örnekleri

# Arka plan işleri (Runspace) ve kısmi kilitleme
$script:PackageJobs  = [hashtable]::Synchronized(@{})   # iş Id (int) -> iş kaydı (Package silme ve SUG işleri)
$script:JobSeq       = 0
$script:LogQueue     = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'   # arka plan -> UI günlük satırları
$script:LockedPkgs   = @{}      # PackageID -> kilitleyen iş Id
$script:ViewIds      = New-Object 'System.Collections.Generic.List[string]'   # VirtualMode: satır i -> güncelleme CI_ID
$script:ViewIndex    = @{}      # güncelleme CI_ID -> satır indeksi
$script:RowCache     = @{}      # güncelleme CI_ID -> hücre değerleri (çizildikçe hesaplanır)
$script:SortCol      = -1       # güncelleme listesinin sıralama sütunu (-1: KB + Başlık)
$script:SortAsc      = $true
$script:GrpRowById   = @{}      # SUG CI_ID -> DataGridViewRow
$script:PkgRowById   = @{}      # PackageID -> DataGridViewRow
$script:ListMode     = -1       # son "Listele"de kullanılan kaynak modu (yerel filtre için)
$script:ListScopeSugIds = New-Object 'System.Collections.Generic.HashSet[string]'   # son listelemenin kapsamındaki SUG'lar
$script:ListJob      = $null    # arka planda süren Listele / Ortamı Oku işi (aynı anda en fazla bir tane)
$script:InTick       = $false   # zamanlayıcı olayı içinde miyiz (yeniden girişi önlemek için)
$script:WorkerState  = $null    # arka plan runspace'leri için InitialSessionState (önbellek)

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
    $plan = New-Object 'System.Collections.Generic.List[object]'
    foreach ($gid in @($Current.Keys)) {
        # NOT: PowerShell degisken adlari buyuk/kucuk harf duyarsizdir. Eskiden $orig = $Orig[$gid]
        # yazilmisti; bu, [hashtable] tipli $Orig parametresinin uzerine yazmaya calisip hata veriyordu.
        $curState  = $Current[$gid]
        $origState = $Orig[$gid]
        if ($curState -eq $origState) { continue }
        if ($curState -eq [System.Windows.Forms.CheckState]::Indeterminate) { continue }

        $add = New-Object 'System.Collections.Generic.List[string]'
        $rem = New-Object 'System.Collections.Generic.List[string]'
        foreach ($uid in $SelectedIds) {
            $isMember = $Members.ContainsKey($uid) -and $Members[$uid].Contains($gid)
            if ($curState -eq [System.Windows.Forms.CheckState]::Checked) {
                if (-not $isMember) { $add.Add($uid) }
            } else {
                if ($isMember) { $rem.Add($uid) }
            }
        }
        if ($add.Count -gt 0 -or $rem.Count -gt 0) {
            $plan.Add([pscustomobject]@{
                GroupId = [string]$gid
                Add     = [string[]]$add.ToArray()
                Remove  = [string[]]$rem.ToArray()
            })
        }
    }
    return $plan.ToArray()
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

    $remove = New-Object 'System.Collections.Generic.List[uint32]'
    $shared = New-Object 'System.Collections.Generic.List[uint32]'
    foreach ($c in @($refs.Keys)) {
        $mine = $false
        $others = $false
        foreach ($ci in $refs[$c]) {
            if ($sel.ContainsKey($ci)) { $mine = $true } else { $others = $true }
        }
        if ($mine -and -not $others) { $remove.Add([uint32]$c) }
        elseif ($mine -and $others)  { $shared.Add([uint32]$c) }
    }
    return [pscustomobject]@{
        Remove = [uint32[]]$remove.ToArray()
        Shared = [uint32[]]$shared.ToArray()
    }
}

# ------------------------------------------------------------------
# Merkezi girdi doğrulama: WQL'e giren her kimlik bu fonksiyonlardan geçer (WQL enjeksiyonuna karşı).
# Test-* : bool döner (uyarı verip atlamak için)   Assert-* : geçersizse hata fırlatır.
# ------------------------------------------------------------------
function Test-PackageId { param([string]$Id) return ($Id -match '^[A-Za-z0-9]{1,16}$') }
function Test-CiId      { param([string]$Id) return ($Id -match '^\d{1,10}$') }
function Test-CategoryUniqueId { param([string]$Id) return ($Id -match '^[A-Za-z0-9:_\-\.]{1,256}$') }

function Assert-PackageId {
    param([string]$Id)
    if (-not (Test-PackageId $Id)) { throw "Geçersiz PackageID: '$Id'" }
}

function Assert-CiId {
    param([string]$Id)
    if (-not (Test-CiId $Id)) { throw "Geçersiz CI_ID: '$Id'" }
}

# SMS_SoftwareUpdate sorgularında yalnızca kullanılan sütunlar çekilir (SELECT * yerine; ağ trafiğini azaltır).
function Get-UpdateSelectList {
    param([string]$Alias = '')
    $cols = 'CI_ID', 'ArticleID', 'LocalizedDisplayName', 'DatePosted', 'IsSuperseded', 'IsExpired',
            'LocalizedCategoryInstanceNames', 'CategoryInstance_UniqueIDs'
    if ($Alias) { return (($cols | ForEach-Object { "$Alias.$_" }) -join ', ') }
    return ($cols -join ', ')
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
            # VirtualMode grid'de AllCells tüm satırların değerini isterdi (çok yavaş); görünen hücrelere göre genişlet.
            if ($sender.VirtualMode) { $mode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::DisplayedCells }
            else                     { $mode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells }
            $sender.AutoResizeColumn($e.ColumnIndex, $mode)
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
$form.ClientSize = New-Object System.Drawing.Size(1280, 892)    # tasarım boyutu; açılışta ekrana sığdırılır (Set-InitialWindowSize)
$form.StartPosition = 'Manual'
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
$gUpd.Anchor = 'Top,Left'           # konum/boyut Update-MainLayout ile hesaplanır
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
# VirtualMode: satırlar grid'e eklenmez, hücre değerleri CellValueNeeded ile istendikçe verilir.
$gridUpd.VirtualMode = $true
foreach ($col in $gridUpd.Columns) { $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic }

$btnSelAll = New-UiButton $gUpd 'Tümünü Seç' 10 462 100 26
$btnSelAll.Anchor = 'Bottom,Left'
$btnSelNone = New-UiButton $gUpd 'Seçimi Temizle' 115 462 110 26
$btnSelNone.Anchor = 'Bottom,Left'
$lblCount = New-UiLabel $gUpd '' 235 466 545
$lblCount.Anchor = 'Bottom,Left'

# --- Sağ panel: SUG'lar ve Deployment Package'lar (sekmeli) ---
$gGrp = New-UiGroup $form 'Gruplar ve Package''lar (ortamdaki tümü)' 820 140 450 500
$gGrp.Anchor = 'Top,Left'           # konum/boyut Update-MainLayout ile hesaplanır (genişlik pencereye göre 380-560)
$tabRight = New-Object System.Windows.Forms.TabControl
$tabRight.Location = New-Object System.Drawing.Point(6, 18)
$tabRight.Size = New-Object System.Drawing.Size(438, 476)
$tabRight.Anchor = 'Top,Left'          # konum/boyut Update-RightPanelLayout ile hesaplanır
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
$txtFilter.Anchor = 'Top,Left'           # Update-RightPanelLayout ile hesaplanır
$gridGrp = New-UiGrid $tabSug 6 36 424 360
$gridGrp.MultiSelect = $false
$gridGrp.Anchor = 'Top,Left'           # Update-RightPanelLayout ile hesaplanır
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
$lblLegend.Anchor = 'Top,Left'           # Update-RightPanelLayout ile hesaplanır
$lblLegend.ForeColor = [System.Drawing.Color]::DimGray

# -- Sekme 2: Deployment Package'lar (silmeden hariç tut) --
[void](New-UiLabel $tabPkg 'Package ara:' 6 10 72)
$txtPkgFilter = New-UiTextBox $tabPkg 82 7 344
$txtPkgFilter.Anchor = 'Top,Left'        # Update-RightPanelLayout ile hesaplanır
$gridPkg = New-UiGrid $tabPkg 6 36 424 330
$gridPkg.MultiSelect = $false
$gridPkg.Anchor = 'Top,Left'           # Update-RightPanelLayout ile hesaplanır
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
$lblPkgLegend.Anchor = 'Top,Left'        # Update-RightPanelLayout ile hesaplanır
$lblPkgLegend.ForeColor = [System.Drawing.Color]::DimGray

# --- İşlemler ---
$gAct = New-UiGroup $form 'İşlemler' 10 646 1260 62
$gAct.Anchor = 'Top,Left'           # konum/boyut Update-MainLayout ile hesaplanır
$chkSim = New-Object System.Windows.Forms.CheckBox
$chkSim.Text = 'Simülasyon (değişiklik yapma, sadece günlüğe yaz)'
$chkSim.Location = New-Object System.Drawing.Point(14, 26)
$chkSim.Size = New-Object System.Drawing.Size(335, 22)
$chkSim.Checked = $true
$gAct.Controls.Add($chkSim)
$btnApply = New-UiButton $gAct 'SUG Üyeliklerini Uygula' 360 20 210 32
$btnRemove = New-UiButton $gAct 'Seçilileri Package''tan Sil' 580 20 210 32
$btnClose = New-UiButton $gAct 'Kapat' 1140 20 105 32
$btnClose.Anchor = 'Top,Right'

# --- Günlük ---
$lblLog = New-UiLabel $form 'İşlem Günlüğü:' 14 714 120
# Arka plan iş durumu, günlük başlığının yanında gösterilir (dar pencerede de görünür kalır).
$lblJobs = New-UiLabel $form '' 144 714 1000
$lblJobs.ForeColor = [System.Drawing.Color]::DarkOrange
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 736)
$txtLog.Size = New-Object System.Drawing.Size(1260, 148)
$txtLog.Anchor = 'Top,Left'         # konum/boyut Update-MainLayout ile hesaplanır
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtLog)

# ------------------------------------------------------------------
# Pencere yerleşimi
# ------------------------------------------------------------------
# Ana bölümler sabit tasarım boyutuna göre değil, pencerenin o anki boyutuna göre yerleştirilir.
# Böylece küçük ya da ölçeklendirilmiş (%125-%150) ekranlarda sağdaki "Gruplar ve Package'lar" paneli ekran dışına taşmaz.
# İç kontroller (tablolar, kutular) kendi Anchor ayarlarıyla bu bölümlere göre büyüyüp küçülür.
$script:LayoutMinClient = New-Object System.Drawing.Size(980, 640)   # tabloların okunabilir kaldığı en küçük iç alan
$script:LayoutMidMinH   = 300    # orta bölüm (güncelleme + SUG/Package tabloları) için en az yükseklik

function Update-MainLayout {
    $W = $form.ClientSize.Width
    $H = $form.ClientSize.Height
    if ($W -le 0 -or $H -le 0) { return }       # simge durumundayken
    $m = 10
    $form.SuspendLayout()
    try {
        $gTop.SetBounds($m, 8, $W - 2 * $m, 124)
        $midY = 140
        # Günlük kutusu pencere yüksekliğiyle orantılı (70-220 px); alan darsa önce günlük küçülür, tablolar korunur.
        $logH = [Math]::Max(70, [Math]::Min(220, [int]($H * 0.17)))
        $fixedBelow = 6 + 62 + 6 + 22 + 8       # İşlemler kutusu + günlük başlığı + kenar boşlukları
        $spare = $H - $midY - $fixedBelow - $logH - $script:LayoutMidMinH
        if ($spare -lt 0) { $logH = [Math]::Max(60, $logH + $spare) }
        $logY = $H - 8 - $logH
        $txtLog.SetBounds($m, $logY, $W - 2 * $m, $logH)
        $lblLog.SetBounds($m + 4, $logY - 22, 120, 20)
        $lblJobs.SetBounds($m + 130, $logY - 22, [Math]::Max(50, $W - 2 * $m - 134), 20)
        $actY = $logY - 22 - 6 - 62
        $gAct.SetBounds($m, $actY, $W - 2 * $m, 62)
        $midH = [Math]::Max(120, $actY - 6 - $midY)
        $rightW = [Math]::Max(380, [Math]::Min(560, [int]($W * 0.36)))
        $leftW  = [Math]::Max(300, $W - 3 * $m - $rightW)
        $gUpd.SetBounds($m, $midY, $leftW, $midH)
        $gGrp.SetBounds($m + $leftW + $m, $midY, $rightW, $midH)
        Update-RightPanelLayout
    } finally {
        $form.ResumeLayout()
    }
}

# Sağ paneldeki (SUG / Package sekmeleri) iç kontrolleri sekme sayfasının GERÇEK iç boyutuna göre yerleştirir.
# Neden Anchor kullanılmıyor: TabPage içindeki kontroller pencere açılmadan, sekme sayfası henüz gerçek boyutuna
# gelmeden eklendiği için Anchor mesafeleri yanlış kaydediliyordu; sayfa sonradan yeniden boyutlanınca tablo ve
# açıklama yazıları sayfanın sağ/alt kenarından taşıp kesiliyordu (panelin bir kısmı görünmüyordu).
$script:InRightLayout = $false

function Set-TabPageLayout {
    param($Filter, $Grid, $Legend, [int]$W, [int]$H)
    $pad = 6
    $cw = [Math]::Max(120, $W - 2 * $pad)
    # Açıklama yazısının yüksekliği, panel genişliğine göre satır sayısından hesaplanır (yazı kesilmesin).
    $sz = [System.Windows.Forms.TextRenderer]::MeasureText($Legend.Text, $Legend.Font, (New-Object System.Drawing.Size($cw, 0)), [System.Windows.Forms.TextFormatFlags]::WordBreak)
    $lh = [Math]::Min($sz.Height + 8, [Math]::Max(24, [int]($H * 0.45)))
    $top = $Grid.Top
    $Filter.SetBounds($Filter.Left, $Filter.Top, [Math]::Max(60, $W - $pad - $Filter.Left), $Filter.Height)
    $Grid.SetBounds($pad, $top, $cw, [Math]::Max(60, $H - $top - $lh - 8))
    $Legend.SetBounds($pad, [Math]::Max($top + 60, $H - $lh - 4), $cw, $lh)
}

function Update-RightPanelLayout {
    if ($script:InRightLayout) { return }       # SetBounds'in tetiklediği olaylarla yeniden girişi engelle
    $script:InRightLayout = $true
    try {
        $tabRight.SetBounds(6, 18, [Math]::Max(120, $gGrp.ClientSize.Width - 12), [Math]::Max(120, $gGrp.ClientSize.Height - 24))
        if ($tabRight.IsHandleCreated) {
            $r = $tabRight.DisplayRectangle         # sekme başlığı çıkarılmış gerçek sayfa alanı
            $w = $r.Width; $h = $r.Height
        } else {
            $w = $tabRight.ClientSize.Width - 8     # pencere açılmadan önce tahmini değer; açılınca yeniden hesaplanır
            $h = $tabRight.ClientSize.Height - 30
        }
        if ($w -lt 120 -or $h -lt 120) { return }
        Set-TabPageLayout $txtFilter    $gridGrp $lblLegend    $w $h
        Set-TabPageLayout $txtPkgFilter $gridPkg $lblPkgLegend $w $h
    } finally {
        $script:InRightLayout = $false
    }
}

# Açılışta pencereyi bulunduğu ekranın çalışma alanına sığdırır ve ortalar; en küçük boyutu da ekrana göre sınırlar.
function Set-InitialWindowSize {
    $wa = [System.Windows.Forms.Screen]::FromPoint([System.Windows.Forms.Cursor]::Position).WorkingArea
    $fw = $form.Width - $form.ClientSize.Width          # çerçeve + başlık çubuğu
    $fh = $form.Height - $form.ClientSize.Height
    $minW = [Math]::Min($script:LayoutMinClient.Width + $fw, $wa.Width)
    $minH = [Math]::Min($script:LayoutMinClient.Height + $fh, $wa.Height)
    $form.MinimumSize = New-Object System.Drawing.Size($minW, $minH)
    $w = [Math]::Min(1280 + $fw, $wa.Width)
    $h = [Math]::Min(892 + $fh, $wa.Height)
    $x = $wa.X + [int](($wa.Width - $w) / 2)
    $y = $wa.Y + [int](($wa.Height - $h) / 2)
    $form.Bounds = New-Object System.Drawing.Rectangle($x, $y, $w, $h)
    # Ekran tasarlanan en küçük boyuttan da küçükse tam ekran aç (hiçbir bölüm dışarıda kalmasın).
    if ($wa.Width -lt ($script:LayoutMinClient.Width + $fw) -or $wa.Height -lt ($script:LayoutMinClient.Height + $fh)) {
        $form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized
    }
    Update-MainLayout
}

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    # Uzun işlemler arka planda olduğundan DoEvents gerekmez (yeniden giriş riski de yok).
    Add-LogText ($line + [Environment]::NewLine)
}

function Add-LogText {
    param([string]$Text)
    if (-not $Text) { return }
    # Günlük kutusu sınırsız büyümesin: çok uzun metinde AppendText yavaşlar ve bellek şişer.
    if ($txtLog.TextLength -gt 1500000) {
        $txtLog.Text = $txtLog.Text.Substring($txtLog.TextLength - 750000)
    }
    $txtLog.AppendText($Text)
}

# Arka plan işlerinin kuyruğa bıraktığı günlük satırlarını UI thread'inde tek seferde ekler.
function Receive-JobLog {
    $sb = New-Object System.Text.StringBuilder
    $line = $null
    while ($script:LogQueue.TryDequeue([ref]$line)) { [void]$sb.AppendLine($line) }
    if ($sb.Length -gt 0) { Add-LogText $sb.ToString() }
}

# Düğmeleri arka plan işlerine göre ayarlar.
#  - Listele / Ortamı Oku sürerken: veri yeniden kurulacağı için SUG/Package işlemleri de kapalıdır.
#  - SUG / Package işi sürerken: listeyi baştan kuran işlemler (Listele / Ortamı Oku / state okuma) kapalıdır;
#    SUG ve Package işlemleri kilitli olmayan güncellemeler için kullanılmaya devam eder.
function Update-ControlStates {
    $jobs = $script:PackageJobs.Count
    $listing = ($null -ne $script:ListJob)
    $mode = $cmbSource.SelectedIndex
    foreach ($c in @($btnList, $btnReload, $btnBrowse)) { $c.Enabled = (-not $listing -and $jobs -eq 0) }
    $btnEnv.Enabled    = (-not $listing -and $jobs -eq 0 -and $mode -ge 1 -and $mode -le 3)
    $btnApply.Enabled  = (-not $listing)
    $btnRemove.Enabled = (-not $listing)
    if ($listing -or $jobs -gt 0) { $form.Cursor = [System.Windows.Forms.Cursors]::AppStarting }
    else                          { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    $parts = New-Object 'System.Collections.Generic.List[string]'
    if ($listing)    { $parts.Add($script:ListJob.Label + ' sürüyor...') }
    if ($jobs -gt 0) { $parts.Add("$jobs arka plan işlemi sürüyor...") }
    $lblJobs.Text = ($parts -join '  ')
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

    $raws = New-Object 'System.Collections.Generic.List[object]'
    if ($state.LastRun) { $raws.Add($state.LastRun) }
    if ($state.History) {
        $hist = @($state.History)
        [array]::Reverse($hist)          # en yeni başta
        foreach ($h in $hist) { $raws.Add($h) }
    }

    $entries = New-Object 'System.Collections.Generic.List[object]'
    $seen = @{}
    foreach ($r in $raws) {
        $e = ConvertTo-StateEntry $r
        if (-not $e) { continue }
        if (-not $e.SugId -and -not $e.PkgId) { continue }
        $key = '{0}|{1}|{2}' -f $e.Timestamp, $e.SugId, $e.PkgId
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $entries.Add($e)
    }
    $script:StateEntries = $entries.ToArray()

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
            $all = New-Object 'System.Collections.Generic.List[object]'
            foreach ($pk in $script:EnvPackages) { $all.Add((New-ScopeEntry -PkgId $pk.Id -PkgName $pk.Name)) }
            foreach ($gr in $script:Groups)      { $all.Add((New-ScopeEntry -SugId $gr.Id -SugName $gr.Name)) }
            return $all.ToArray()
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
# SCCM erişimi (YALNIZCA WMI / DCOM)
# ------------------------------------------------------------------
# Mimari kısıt: her işlem SMS Provider WMI (root\SMS\site_XXX) üzerinden yapılır.
# CIM (WinRM), PowerShell Remoting (Invoke-Command / Enter-PSSession) ve ConfigurationManager modülü KULLANILMAZ.
function Get-Connection {
    $site = $txtSite.Text.Trim().ToUpper()
    $server = $txtServer.Text.Trim()
    if ($site -notmatch '^[A-Z0-9]{3}$') { throw 'Site kodu 3 karakter (harf/rakam) olmalı.' }
    if (-not $server) { throw 'SMS Provider sunucusu boş olamaz.' }
    if ($server -notmatch '^[A-Za-z0-9][A-Za-z0-9\.\-]{0,253}$') { throw "Geçersiz SMS Provider sunucu adı: '$server'" }
    return [pscustomobject]@{ Site = $site; Server = $server }
}

# Tüm Get-WmiObject çağrıları için ortak bağlantı parametreleri.
# PacketPrivacy: DCOM trafiği imzalanır ve ŞİFRELENİR. Impersonate: sunucu kullanıcının kimliğiyle işlem yapar.
function Get-WmiConnectionArgs {
    $c = Get-Connection
    return @{
        ComputerName   = $c.Server
        Namespace      = "root\SMS\site_$($c.Site)"
        Authentication = [System.Management.AuthenticationLevel]::PacketPrivacy
        Impersonation  = [System.Management.ImpersonationLevel]::Impersonate
        ErrorAction    = 'Stop'
    }
}

function Invoke-CMWql {
    param([string]$Query)
    $wa = Get-WmiConnectionArgs
    return @(Get-WmiObject @wa -Query $Query)
}

# Belirli bir sınıftan filtreyle nesne okur (metot çağrısı / Put için tam WMI nesnesi gerektiğinde).
function Get-CMWmiInstance {
    param([string]$Class, [string]$Filter)
    $wa = Get-WmiConnectionArgs
    return @(Get-WmiObject @wa -Class $Class -Filter $Filter)
}

# SUG (SMS_AuthorizationList) üyeliğini doğrudan WMI ile değiştirir ve sonucu tekrar okuyarak DOĞRULAR.
# (ConfigurationManager modülündeki Add/Remove-CMSoftwareUpdate*Group cmdlet'lerinin yaptığı işlemin WMI karşılığı:
#  'Updates' dizisi okunur, değiştirilir ve Put() ile yazılır.)
function Set-SugMembershipWmi {
    param([string]$GroupId, [string[]]$Add = @(), [string[]]$Remove = @())
    Assert-CiId $GroupId
    foreach ($x in (@($Add) + @($Remove))) { if ($x) { Assert-CiId $x } }

    # Aynı SUG'a eşzamanlı yazım koruması: 'Updates' dizisi oku -> değiştir -> yaz (Put) şeklinde güncellendiği için
    # iki iş aynı anda yazarsa biri diğerinin değişikliğini ezebilir. SUG başına isimli mutex ile bu kısa bölüm sıraya alınır
    # (aynı oturumdaki tüm runspace'ler ve aracın diğer açık pencereleri için geçerli). UI'da SUG kilitlenmez.
    $mtx = New-Object System.Threading.Mutex($false, "Local\SccmSupersededToolSug_$GroupId")
    $owned = $false
    try {
        try { $owned = $mtx.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $owned = $true }
        if (-not $owned) {
            Write-Log "SUG (CI_ID $GroupId) şu an başka bir iş tarafından yazılıyor; sıra bekleniyor..."
            try { $owned = $mtx.WaitOne(180000) } catch [System.Threading.AbandonedMutexException] { $owned = $true }
            if (-not $owned) { throw "SUG (CI_ID $GroupId) 3 dakika içinde yazılabilir olmadı (başka bir iş bırakmadı)." }
        }
        return (Set-SugMembershipWmiCore -GroupId $GroupId -Add $Add -Remove $Remove)
    } finally {
        if ($owned) { $mtx.ReleaseMutex() }
        $mtx.Dispose()
    }
}

# Set-SugMembershipWmi'nin mutex altında çalışan asıl kısmı (doğrudan çağırmayın).
function Set-SugMembershipWmiCore {
    param([string]$GroupId, [string[]]$Add = @(), [string[]]$Remove = @())
    $sug = @(Get-CMWmiInstance -Class 'SMS_AuthorizationList' -Filter "CI_ID = $GroupId")
    if ($sug.Count -eq 0) { throw "SUG bulunamadı (CI_ID $GroupId)." }
    $sug = $sug[0]
    $sug.Get()     # 'Updates' lazy özellik; sorgu sonucunda gelmez, nesne tam olarak yüklenmeli

    $seen = New-Object 'System.Collections.Generic.HashSet[uint32]'
    $remSet = New-Object 'System.Collections.Generic.HashSet[uint32]'
    foreach ($r in @($Remove)) { if ($r) { [void]$remSet.Add([uint32]$r) } }
    $new = New-Object 'System.Collections.Generic.List[uint32]'
    foreach ($u in @($sug.Updates)) {
        if ($null -eq $u) { continue }
        $uu = [uint32]$u
        if (-not $seen.Add($uu)) { continue }
        if (-not $remSet.Contains($uu)) { $new.Add($uu) }
    }
    foreach ($a in @($Add)) {
        if (-not $a) { continue }
        $ua = [uint32]$a
        if ($seen.Add($ua)) { $new.Add($ua) }
    }
    $sug.Updates = [uint32[]]$new.ToArray()
    [void]$sug.Put()

    # Doğrulama: SUG'ı yeniden oku
    $verify = @(Get-CMWmiInstance -Class 'SMS_AuthorizationList' -Filter "CI_ID = $GroupId")[0]
    $verify.Get()
    $after = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($u in @($verify.Updates)) { if ($null -ne $u) { [void]$after.Add([string]$u) } }
    $added      = New-Object 'System.Collections.Generic.List[string]'
    $notAdded   = New-Object 'System.Collections.Generic.List[string]'
    $removed    = New-Object 'System.Collections.Generic.List[string]'
    $notRemoved = New-Object 'System.Collections.Generic.List[string]'
    foreach ($a in @($Add))    { if ($a) { if ($after.Contains([string]$a)) { $added.Add([string]$a) } else { $notAdded.Add([string]$a) } } }
    foreach ($r in @($Remove)) { if ($r) { if ($after.Contains([string]$r)) { $notRemoved.Add([string]$r) } else { $removed.Add([string]$r) } } }
    return [pscustomobject]@{
        Added      = [string[]]$added.ToArray()
        Removed    = [string[]]$removed.ToArray()
        NotAdded   = [string[]]$notAdded.ToArray()
        NotRemoved = [string[]]$notRemoved.ToArray()
        Total      = $after.Count
    }
}

# ------------------------------------------------------------------
# Denetim kaydı (audit) - geri alınamaz işlemler dosyaya JSON satırı olarak yazılır
# ------------------------------------------------------------------
# Dosya: %LOCALAPPDATA%\update-state\audit.log  (5 MB'ı geçince audit.log.1'e döndürülür)
# Hem UI thread'inden hem arka plan runspace'lerinden çağrılır; eşzamanlı yazım isimli Mutex ile korunur.
# Audit yazılamazsa işlem DURDURULMAZ, yalnızca günlüğe uyarı düşülür.
function Get-AuditFilePath {
    $base = $env:LOCALAPPDATA
    if (-not $base) { $base = [Environment]::GetFolderPath('LocalApplicationData') }
    return (Join-Path $base 'update-state\audit.log')
}

function Write-Audit {
    param([string]$Action, [hashtable]$Data = @{})
    try {
        $rec = [ordered]@{
            Time    = (Get-Date).ToString('o')
            User    = ('{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME)
            Machine = $env:COMPUTERNAME
            Action  = $Action
        }
        foreach ($k in @($Data.Keys)) { $rec[$k] = $Data[$k] }
        $line = (New-Object psobject -Property $rec) | ConvertTo-Json -Compress -Depth 4
        $path = Get-AuditFilePath
        $mtx = New-Object System.Threading.Mutex($false, 'Local\SccmSupersededToolAudit')
        $owned = $false
        try {
            try { $owned = $mtx.WaitOne(5000) } catch [System.Threading.AbandonedMutexException] { $owned = $true }
            $dir = Split-Path -Path $path -Parent
            if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
            if ((Test-Path -LiteralPath $path) -and ((Get-Item -LiteralPath $path).Length -gt 5MB)) {
                $old = "$path.1"
                if (Test-Path -LiteralPath $old) { Remove-Item -LiteralPath $old -Force }
                Move-Item -LiteralPath $path -Destination $old
            }
            [System.IO.File]::AppendAllText($path, $line + "`r`n", (New-Object System.Text.UTF8Encoding($false)))
        } finally {
            if ($owned) { $mtx.ReleaseMutex() }
            $mtx.Dispose()
        }
    } catch {
        Write-Log "Denetim kaydı yazılamadı ($Action): $($_.Exception.Message)" 'WARN'
    }
}

# ------------------------------------------------------------------
# Veri yükleme (Listele işinde ARKA PLAN runspace'inde çalışır; UI kontrollerine dokunmaz)
# ------------------------------------------------------------------
function Import-Groups {
    $rows = Invoke-CMWql 'SELECT CI_ID, LocalizedDisplayName, NumberOfUpdates FROM SMS_AuthorizationList'
    $script:Groups = New-Object 'System.Collections.Generic.List[object]'
    $script:GroupById = @{}
    foreach ($r in ($rows | Sort-Object { [string]$_.LocalizedDisplayName })) {
        $g = [pscustomobject]@{
            Id    = [string]$r.CI_ID
            Name  = [string]$r.LocalizedDisplayName
            Total = [int]$r.NumberOfUpdates
        }
        $script:Groups.Add($g)
        $script:GroupById[$g.Id] = $g
    }
}

# Ortamdaki Software Update Deployment Package'ları TEK sorguda okur ve önbelleğe alır.
# ($script:EnvPackages: adına göre sıralı liste, $script:PkgNames: PackageID -> ad)
# Listeleme boyunca SMS_SoftwareUpdatesPackage bir daha sorgulanmaz.
function Import-PackageCatalog {
    $rows = Invoke-CMWql 'SELECT PackageID, Name FROM SMS_SoftwareUpdatesPackage'
    $script:EnvPackages = New-Object 'System.Collections.Generic.List[object]'
    $script:PkgNames = @{}
    foreach ($r in ($rows | Sort-Object { [string]$_.Name })) {
        $p = [pscustomobject]@{
            Id   = [string]$r.PackageID
            Name = [string]$r.Name
        }
        $script:EnvPackages.Add($p)
        $script:PkgNames[$p.Id] = $p.Name
    }
}

# Ortamdaki tüm Deployment Package'ları ve SUG'ları okur (state.json'dan bağımsız). "Yapı" listesi UI'da doldurulur.
function Import-Environment {
    Write-Log "Ortamdaki Deployment Package'lar ve SUG'lar okunuyor..."
    Import-Groups
    Import-PackageCatalog
    Write-Log "$($script:EnvPackages.Count) Deployment Package, $($script:Groups.Count) SUG bulundu." 'OK'
}

# "Yapı" listesini seçili kaynağa göre doldurur (UI).
function Update-EnvCombo {
    $mode = $cmbSource.SelectedIndex
    $cmbEnv.BeginUpdate()
    try {
        $cmbEnv.Items.Clear()
        if ($mode -eq 1) {
            foreach ($pk in $script:EnvPackages) { [void]$cmbEnv.Items.Add(('{0}  ({1})' -f $pk.Name, $pk.Id)) }
        }
        elseif ($mode -eq 2) {
            foreach ($gr in $script:Groups) { [void]$cmbEnv.Items.Add(('{0}  (CI_ID {1})' -f $gr.Name, $gr.Id)) }
        }
    } finally {
        $cmbEnv.EndUpdate()
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
    $allNames = New-Object 'System.Collections.Generic.List[string]'
    $matched  = New-Object 'System.Collections.Generic.List[object]'
    foreach ($r in $rows) {
        $nm = ([string]$r.LocalizedCategoryInstanceName).Trim()
        $allNames.Add($nm)
        if ($nm -ieq $v) { $matched.Add($r) }
    }
    if ($matched.Count -eq 0) {
        foreach ($r in $rows) {
            if (([string]$r.LocalizedCategoryInstanceName).IndexOf($v, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { $matched.Add($r) }
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
        if (-not (Test-CategoryUniqueId $rawId)) { Write-Log "Geçersiz kategori UniqueID atlandı: '$rawId'" 'WARN'; continue }
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
                $script:VendorSample.Add(('CI_ID {0} | Names=[{1}] | UniqueIDs=[{2}]' -f [string]$Row.CI_ID, ((@($Row.LocalizedCategoryInstanceNames)) -join ', '), ((@($Row.CategoryInstance_UniqueIDs)) -join ', ')))
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
            IsLocked   = $false     # arka plan işine alındı mı (kısmi kilitleme)
            Status     = ''         # kilitliyken gösterilecek işlem metni
            LockJob    = 0          # kilitleyen iş Id
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
    foreach ($g in $script:Groups) { if (Test-CiId $g.Id) { [void]$groupIds.Add($g.Id) } }
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
# Yalnızca Software Update Deployment Package'lar (önbellekteki $script:PkgNames) dikkate alınır.
function Import-UpdatePackages {
    param([string]$Filter)
    $pkgSet = $script:PkgNames     # Import-PackageCatalog önbelleği (ek sorgu yok)

    $pkgByContent = @{}
    $rows = Invoke-CMWql "SELECT pc.PackageID, pc.ContentID FROM SMS_PackageToContent AS pc JOIN SMS_CIToContent AS cc ON pc.ContentID = cc.ContentID JOIN SMS_SoftwareUpdate AS su ON su.CI_ID = cc.CI_ID WHERE $Filter"
    foreach ($r in $rows) {
        $pk = [string]$r.PackageID
        if (-not $pkgSet.ContainsKey($pk)) { continue }
        $ct = [string]$r.ContentID
        if (-not $pkgByContent.ContainsKey($ct)) { $pkgByContent[$ct] = New-Object 'System.Collections.Generic.HashSet[string]' }
        [void]$pkgByContent[$ct].Add($pk)
    }

    $rows = Invoke-CMWql "SELECT cc.CI_ID, cc.ContentID FROM SMS_CIToContent AS cc JOIN SMS_SoftwareUpdate AS su ON su.CI_ID = cc.CI_ID WHERE $Filter"
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
    $rows = Invoke-CMWql 'SELECT AssignedCIs FROM SMS_UpdateGroupAssignment'
    foreach ($r in $rows) {
        foreach ($ci in @($r.AssignedCIs)) {
            if ($null -ne $ci -and (Test-CiId "$ci")) { $ids[[string]$ci] = $true }
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
    $names = New-Object 'System.Collections.Generic.List[string]'
    foreach ($gid in $script:Members[$UpdateId]) {
        if ($script:GroupById.ContainsKey($gid)) { $names.Add($script:GroupById[$gid].Name) }
    }
    $names.Sort([System.StringComparer]::CurrentCultureIgnoreCase)
    return ($names -join '; ')
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
    $gridGrp.SuspendLayout()
    try {
        $gridGrp.CurrentCell = $null
        $gridGrp.Rows.Clear()
        $script:GrpRowById = @{}
        # Performans: satırlar tek tek Rows.Add() + hücre ataması yerine önce bellekte kurulup AddRange ile eklenir.
        $newRows = New-Object 'System.Collections.Generic.List[System.Windows.Forms.DataGridViewRow]'
        foreach ($g in $script:Groups) {
            $row = New-Object System.Windows.Forms.DataGridViewRow
            $row.CreateCells($gridGrp)
            $row.Tag = $g.Id
            $mark = ''
            if ($stateSug.ContainsKey($g.Id)) {
                $mark = $star
                $row.DefaultCellStyle.Font = $boldFont
            }
            [void]$row.SetValues([object[]]@([System.Windows.Forms.CheckState]::Unchecked, $g.Name, '', $g.Total, $mark))
            $newRows.Add($row)
        }
        if ($newRows.Count -gt 0) { $gridGrp.Rows.AddRange($newRows.ToArray()) }
        foreach ($row in $gridGrp.Rows) {
            $script:GrpRowById[[string]$row.Tag] = $row
            Set-GroupRowStyle $row
        }
    } finally {
        $gridGrp.ResumeLayout()
        $script:Suspend = $prev
    }
}

# SUG satırının görünümü: değiştirilmiş (sarı) / normal.
# SUG'lar KİLİTLENMEZ: aynı SUG üzerinde aynı anda birden fazla iş çalışabilir (yalnızca güncellemeler kilitlenir).
# Aynı SUG'a eşzamanlı WMI yazımı Set-SugMembershipWmi içindeki SUG başına mutex ile sıraya alınır.
function Set-GroupRowStyle {
    param($Row)
    $gid = [string]$Row.Tag
    $chk = $Row.Cells['gChk']
    $cur = ConvertTo-CheckState $chk.Value
    if ($script:OrigCheck.ContainsKey($gid) -and $cur -ne $script:OrigCheck[$gid]) {
        $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
    } else {
        $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Empty
    }
}

# Güncelleme satırının hücre değerleri (sütun sırası: uKB, uTitle, uPosted, uStatus, uPkg, uGroups, uId).
function Get-UpdateRowValues {
    param($U)
    $parts = New-Object 'System.Collections.Generic.List[string]'
    if ($U.Superseded) { $parts.Add('Superseded') }
    if ($U.Expired)    { $parts.Add('Expired') }
    if ($script:DeployedIds.Count -gt 0 -and (Test-UpdateDeployed $U.Id)) { $parts.Add('Deployed') }
    $status = ($parts -join ' + ')
    if ($U.IsLocked) { $status = ('[KİLİTLİ: {0}] {1}' -f $U.Status, $status) }
    return [object[]]@($U.KB, $U.Title, $U.Posted, $status, ($U.PkgIds -join ', '), (Get-MemberNames $U.Id), $U.Id)
}

# ------------------------------------------------------------------
# Güncelleme grid'i VirtualMode'da çalışır:
#   $script:ViewIds   : ekrandaki sıraya göre CI_ID listesi (satır i = ViewIds[i])
#   $script:ViewIndex : CI_ID -> satır indeksi
#   $script:RowCache  : CI_ID -> hücre değerleri (yalnızca ekranda çizilen satırlar için, ilk ihtiyaçta hesaplanır)
# Hücreler CellValueNeeded ile istendikçe doldurulur; on binlerce satırda bile grid'e satır nesnesi eklenmez.
# ------------------------------------------------------------------
function Get-CachedRowValues {
    param([string]$Id)
    $v = $script:RowCache[$Id]
    if ($null -eq $v) {
        $u = $script:Updates[$Id]
        if (-not $u) { return $null }
        $v = Get-UpdateRowValues $u
        $script:RowCache[$Id] = $v
    }
    return ,$v
}

function Update-ViewIndex {
    $script:ViewIndex = @{}
    for ($i = 0; $i -lt $script:ViewIds.Count; $i++) { $script:ViewIndex[$script:ViewIds[$i]] = $i }
}

# Verilen CI_ID'leri (görünürlerse) seçili yapar. Seçim olayları susturulmuş olmalıdır.
function Select-UpdateIds {
    param([string[]]$Ids)
    foreach ($id in @($Ids)) {
        if ($id -and $script:ViewIndex.ContainsKey($id)) { $gridUpd.Rows[$script:ViewIndex[$id]].Selected = $true }
    }
}

# ViewIds listesini yeniden kurar (sıralama/silme sonrası); seçim CI_ID bazında korunur.
function Set-UpdateView {
    param([System.Collections.Generic.List[string]]$Ids)
    $selIds = @(Get-SelectedUpdateIds -IncludeLocked)
    $prev = $script:Suspend
    $script:Suspend = $true
    try {
        $gridUpd.ClearSelection()
        $script:ViewIds = $Ids
        Update-ViewIndex
        if ($gridUpd.RowCount -ne $Ids.Count) { $gridUpd.RowCount = $Ids.Count }
        Select-UpdateIds $selIds
        $gridUpd.Invalidate()
    } finally {
        $script:Suspend = $prev
    }
}

# Sıralama: varsayılan KB + Başlık; başlığa tıklanınca o sütuna göre (tekrar tıklayınca ters).
function Get-SortedUpdateIds {
    $vals = New-Object 'System.Collections.Generic.List[object]'
    $col = $script:SortCol
    foreach ($u in $script:Updates.Values) {
        if ($col -lt 0) { $key = $u.KB; $key2 = $u.Title }
        else {
            $rv = Get-CachedRowValues $u.Id
            $key = $rv[$col]
            if ($col -eq 6) { $key = [int64]$u.Id }     # CI_ID sayısal sıralansın
            $key2 = $u.Title
        }
        $vals.Add([pscustomobject]@{ Id = $u.Id; K = $key; K2 = $key2 })
    }
    $desc = ($script:SortCol -ge 0 -and -not $script:SortAsc)
    $ids = New-Object 'System.Collections.Generic.List[string]'
    foreach ($x in ($vals | Sort-Object -Property @{ Expression = 'K'; Descending = $desc }, @{ Expression = 'K2'; Descending = $desc })) { $ids.Add($x.Id) }
    return ,$ids
}

function Update-UpdateViewOrder {
    $ids = Get-SortedUpdateIds
    Set-UpdateView -Ids $ids
    for ($i = 0; $i -lt $gridUpd.Columns.Count; $i++) {
        $g = [System.Windows.Forms.SortOrder]::None
        if ($i -eq $script:SortCol) {
            if ($script:SortAsc) { $g = [System.Windows.Forms.SortOrder]::Ascending } else { $g = [System.Windows.Forms.SortOrder]::Descending }
        }
        $gridUpd.Columns[$i].HeaderCell.SortGlyphDirection = $g
    }
}

function Show-Updates {
    $prev = $script:Suspend
    $script:Suspend = $true
    try {
        $gridUpd.ClearSelection()
        $gridUpd.RowCount = 0
        $script:RowCache = @{}
        $script:ViewIds = New-Object 'System.Collections.Generic.List[string]'
        $script:ViewIndex = @{}
    } finally {
        $script:Suspend = $prev
    }
    Update-UpdateViewOrder
    $prev = $script:Suspend
    $script:Suspend = $true
    try { $gridUpd.ClearSelection() } finally { $script:Suspend = $prev }
    Update-GroupChecks
}

# Tek bir güncellemenin önbelleğini düşürür ve satırını yeniden çizdirir (Package / SUG / Durum / kilit görünümü).
function Update-UpdateRow {
    param([string]$Id)
    if (-not $Id) { return }
    [void]$script:RowCache.Remove($Id)
    if ($script:ViewIndex.ContainsKey($Id)) { $gridUpd.InvalidateRow($script:ViewIndex[$Id]) }
}

# Seçili güncellemelerin CI_ID'leri. Kilitli (arka plan işindeki) satırlar varsayılan olarak DAHİL EDİLMEZ.
# VirtualMode: SelectedRows yerine satır durum bayrakları dolaşılır (satırlar paylaşımlı kalır, hızlıdır).
function Get-SelectedUpdateIds {
    param([switch]$IncludeLocked)
    $ids = New-Object 'System.Collections.Generic.List[string]'
    $selFlag = [System.Windows.Forms.DataGridViewElementStates]::Selected
    $i = $gridUpd.Rows.GetFirstRow($selFlag)
    while ($i -ge 0) {
        if ($i -lt $script:ViewIds.Count) {
            $id = $script:ViewIds[$i]
            $skip = $false
            if (-not $IncludeLocked) {
                $u = $script:Updates[$id]
                if ($u -and $u.IsLocked) { $skip = $true }
            }
            if (-not $skip) { $ids.Add($id) }
        }
        $i = $gridUpd.Rows.GetNextRow($i, $selFlag)
    }
    return $ids.ToArray()
}

# Listelenen güncellemelerden kaçının her package'ta bulunduğunu bellekteki veriden sayar.
function Get-PackageListCounts {
    $count = @{}
    foreach ($u in $script:Updates.Values) {
        foreach ($pk in $u.PkgIds) {
            if ($count.ContainsKey($pk)) { $count[$pk]++ } else { $count[$pk] = 1 }
        }
    }
    return $count
}

# Package satırının görünümü: kilitli (gri) / Hariç (yeşil) / normal.
function Set-PkgRowStyle {
    param($Row)
    $pkgKey = [string]$Row.Tag
    $isEx = ($script:ExcludedPkgs.ContainsKey($pkgKey) -and $script:ExcludedPkgs[$pkgKey])
    if ($script:LockedPkgs.ContainsKey($pkgKey)) {
        $Row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray
        $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Gainsboro
    } elseif ($isEx) {
        $Row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Empty
        $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
    } else {
        $Row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Empty
        $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Empty
    }
}

# Listelenen güncellemelerin bulunduğu Deployment Package'ları sağdaki "Package'lar" sekmesine doldurur.
# "Hariç" işaretleri ($script:ExcludedPkgs) yeniden listelemede korunur.
function Show-Packages {
    $count = Get-PackageListCounts     # package adları Listele sırasında önbelleğe alındı ($script:PkgNames)
    $prev = $script:Suspend
    $script:Suspend = $true
    $gridPkg.SuspendLayout()
    try {
        $gridPkg.CurrentCell = $null
        $gridPkg.Rows.Clear()
        $script:PkgRowById = @{}
        $newRows = New-Object 'System.Collections.Generic.List[System.Windows.Forms.DataGridViewRow]'
        $keys = @($count.Keys | Sort-Object { if ($script:PkgNames.ContainsKey($_)) { $script:PkgNames[$_] } else { $_ } })
        foreach ($pkgKey in $keys) {
            $row = New-Object System.Windows.Forms.DataGridViewRow
            $row.CreateCells($gridPkg)
            $row.Tag = $pkgKey
            $isEx = ($script:ExcludedPkgs.ContainsKey($pkgKey) -and $script:ExcludedPkgs[$pkgKey])
            $nm = ''
            if ($script:PkgNames.ContainsKey($pkgKey)) { $nm = $script:PkgNames[$pkgKey] }
            [void]$row.SetValues([object[]]@([bool]$isEx, $nm, $pkgKey, '', $count[$pkgKey]))
            $newRows.Add($row)
        }
        if ($newRows.Count -gt 0) { $gridPkg.Rows.AddRange($newRows.ToArray()) }
        foreach ($row in $gridPkg.Rows) {
            $script:PkgRowById[[string]$row.Tag] = $row
            Set-PkgRowStyle $row
        }
    } finally {
        $gridPkg.ResumeLayout()
        $script:Suspend = $prev
    }
    Update-PackageCounts
}

# Package silme / SUG değişikliği sonrası "Listede" sayılarını bellekteki veriden yeniden hesaplar (WMI sorgusu yok).
# Listede hiç güncellemesi kalmayan (ve kilitli olmayan) package satırı kaldırılır.
function Update-PackageTotals {
    $count = Get-PackageListCounts
    $prev = $script:Suspend
    $script:Suspend = $true
    $gridPkg.SuspendLayout()
    try {
        foreach ($pkgKey in @($script:PkgRowById.Keys)) {
            $row = $script:PkgRowById[$pkgKey]
            $n = 0
            if ($count.ContainsKey($pkgKey)) { $n = [int]$count[$pkgKey] }
            if ($n -eq 0 -and -not $script:LockedPkgs.ContainsKey($pkgKey)) {
                if ($row.Index -ge 0) {
                    if ($gridPkg.CurrentCell -and $gridPkg.CurrentCell.RowIndex -eq $row.Index) { $gridPkg.CurrentCell = $null }
                    $gridPkg.Rows.Remove($row)
                }
                [void]$script:PkgRowById.Remove($pkgKey)
                continue
            }
            if ([string]$row.Cells['pTotal'].Value -ne [string]$n) { $row.Cells['pTotal'].Value = $n }
        }
    } finally {
        $gridPkg.ResumeLayout()
        $script:Suspend = $prev
    }
    Update-PackageCounts
}

# Seçili güncellemelerden kaçının her package'ta bulunduğunu ("Seçili" kolonu) günceller.
# Performans: seçili güncellemelerin PkgIds listeleri BİR kez dolaşılıp package başına sayılır (package × seçim yerine).
function Update-PackageCounts {
    $sel = @(Get-SelectedUpdateIds)
    $n = $sel.Count
    $cnt = @{}
    foreach ($uid in $sel) {
        $u = $script:Updates[$uid]
        if (-not $u) { continue }
        foreach ($pk in $u.PkgIds) { if ($cnt.ContainsKey($pk)) { $cnt[$pk]++ } else { $cnt[$pk] = 1 } }
    }
    foreach ($row in $gridPkg.Rows) {
        $pkgKey = [string]$row.Tag
        $m = 0
        if ($cnt.ContainsKey($pkgKey)) { $m = $cnt[$pkgKey] }
        if ($n -gt 0) { $v = "$m/$n" } else { $v = '' }
        if ([string]$row.Cells['pSel'].Value -ne $v) { $row.Cells['pSel'].Value = $v }
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

# Seçili (kilitli olmayan) güncellemelere göre SUG kutularını (tam / boş / kısmi) yeniden hesaplar.
# Performans: seçili güncellemelerin üyelik kümeleri BİR kez dolaşılıp SUG başına sayaç tutulur
# (eskiden her SUG için tüm seçim dolaşılıyordu: SUG × seçim). Maliyet artık toplam üyelik sayısı kadar.
function Update-GroupChecks {
    $sel = @(Get-SelectedUpdateIds)
    $n = $sel.Count
    $cnt = @{}
    foreach ($uid in $sel) {
        if (-not $script:Members.ContainsKey($uid)) { continue }
        foreach ($gid in $script:Members[$uid]) { if ($cnt.ContainsKey($gid)) { $cnt[$gid]++ } else { $cnt[$gid] = 1 } }
    }
    $prev = $script:Suspend
    $script:Suspend = $true
    $gridGrp.SuspendLayout()
    try {
        foreach ($row in $gridGrp.Rows) {
            $gid = [string]$row.Tag
            $m = 0
            if ($cnt.ContainsKey($gid)) { $m = $cnt[$gid] }
            if ($n -eq 0 -or $m -eq 0) { $st = [System.Windows.Forms.CheckState]::Unchecked }
            elseif ($m -eq $n)         { $st = [System.Windows.Forms.CheckState]::Checked }
            else                       { $st = [System.Windows.Forms.CheckState]::Indeterminate }

            $script:OrigCheck[$gid] = $st
            if ($row.Cells['gChk'].Value -ne $st) { $row.Cells['gChk'].Value = $st }
            if ($n -gt 0) { $v = "$m/$n" } else { $v = '' }
            if ([string]$row.Cells['gSel'].Value -ne $v) { $row.Cells['gSel'].Value = $v }
            Set-GroupRowStyle $row
        }
    } finally {
        $gridGrp.ResumeLayout()
        $script:Suspend = $prev
    }
    Update-CountLabel -Selected $n
}

function Update-CountLabel {
    param([int]$Selected = -1)
    if ($Selected -lt 0) { $Selected = @(Get-SelectedUpdateIds).Count }
    $locked = 0
    foreach ($u in $script:Updates.Values) { if ($u.IsLocked) { $locked++ } }
    $txt = "$($gridUpd.RowCount) güncelleme listelendi"
    if ($Selected -gt 0) { $txt += ", $Selected seçili" }
    if ($locked -gt 0)   { $txt += ", $locked kilitli (işlemde)" }
    $lblCount.Text = $txt + '.'
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
    if ($script:ListJob) { return }
    if ($script:PackageJobs.Count -gt 0) {
        Show-Message 'Arka planda süren işlemler var. Liste, işlemler bitince yenilenebilir (kilitli satırların durumu korunmalı).' 'Uyarı' 'Warning'
        return
    }
    try { $conn = Get-Connection }
    catch { Write-Log $_.Exception.Message 'ERROR'; Show-Message $_.Exception.Message 'Hata' 'Error'; return }

    # Kapsam UI thread'inde (combobox'lardan) belirlenir; arka plan işine yalnızca kimlikler aktarılır.
    $mode = $cmbSource.SelectedIndex
    $pkgIds = @()
    $sugIds = @()
    if ($mode -le 2) {
        $scope = @(Get-SelectedScope)
        if ($scope.Count -eq 0) {
            if ($mode -eq 0) { Show-Message 'Kullanılabilir bir state kaydı yok. Önce state.json dosyasını yükleyin.' 'Uyarı' 'Warning' }
            else             { Show-Message "Önce 'Ortamı Oku' düğmesiyle ortamdaki yapıları yükleyin ve listeden birini seçin." 'Uyarı' 'Warning' }
            return
        }
        $pkgIds = [string[]]@($scope | ForEach-Object { $_.PkgId } | Where-Object { $_ } | Select-Object -Unique)
        $sugIds = [string[]]@($scope | ForEach-Object { $_.SugId } | Where-Object { $_ } | Select-Object -Unique)
    }
    $data = @{
        Mode    = $mode
        PkgIds  = $pkgIds
        SugIds  = $sugIds
        Vendor  = $txtVendor.Text.Trim()
        Expired = [bool]$chkExpired.Checked
    }
    Start-BackgroundJob -Kind 'List' -Label 'Listeleme' -Script $script:ListWorkerScript -Data $data -Conn $conn -Simulate $false
}

# "Ortamı Oku": ortamdaki Package ve SUG'ları arka planda okur.
function Invoke-EnvRead {
    if ($script:ListJob) { return }
    try { $conn = Get-Connection }
    catch { Write-Log $_.Exception.Message 'ERROR'; Show-Message $_.Exception.Message 'Hata' 'Error'; return }
    Start-BackgroundJob -Kind 'List' -Label 'Ortam okuma' -Script $script:ListWorkerScript -Data @{ EnvOnly = $true } -Conn $conn -Simulate $false
}

# Listele / Ortamı Oku işi bitince sonucu UI'a uygular (UI thread'i).
function Complete-ListJob {
    param($Job)
    try { [void]$Job.PS.EndInvoke($Job.Handle) }
    catch { if (-not $Job.Sync.Error) { $Job.Sync.Error = $_.Exception.Message } }
    try { foreach ($er in $Job.PS.Streams.Error) { Write-Log "$($Job.Label): $er" 'ERROR' } } catch { }
    Receive-JobLog
    try { $Job.PS.Dispose() } catch { }
    try { $Job.Runspace.Close(); $Job.Runspace.Dispose() } catch { }
    $script:ListJob = $null

    $r = $Job.Sync.Result
    if ($Job.Sync.Error -or $null -eq $r) {
        $msg = [string]$Job.Sync.Error
        if (-not $msg) { $msg = "$($Job.Label) tamamlanamadı." }
        if ($Job.Sync.Cancel) { Write-Log "$($Job.Label) iptal edildi." 'WARN' }
        else { Show-Message $msg 'Hata' 'Error' }
        Update-ControlStates
        return
    }

    $script:Groups      = $r.Groups
    $script:GroupById   = $r.GroupById
    $script:EnvPackages = $r.EnvPackages
    $script:PkgNames    = $r.PkgNames
    $dur = [int]((Get-Date) - $Job.Started).TotalSeconds

    if ($r.EnvOnly) {
        Update-EnvCombo
        if ($cmbSource.SelectedIndex -eq 0) { Write-Log "Listelemek için 'Kaynak' olarak Package veya SUG seçin." 'WARN' }
    } else {
        $script:Updates         = $r.Updates
        $script:Members         = $r.Members
        $script:DeployedIds     = $r.DeployedIds
        $script:ListMode        = [int]$r.Mode
        $script:ListScopeSugIds = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($s in @($r.ScopeSugIds)) { if ($s) { [void]$script:ListScopeSugIds.Add([string]$s) } }
        $script:LockedPkgs   = @{}
        if ($script:ListMode -eq 3) { Update-EnvCombo }
        Show-Groups
        Show-Updates
        Show-Packages
        Set-GroupFilter
        Set-PackageFilter
        Write-Log "Listeleme tamamlandı: $($script:Updates.Count) benzersiz güncelleme ($dur sn)." 'OK'
    }
    Update-ControlStates
}

function Invoke-ApplyMembership {
    $selAll = @(Get-SelectedUpdateIds -IncludeLocked)
    $sel = @(Get-SelectedUpdateIds)
    if ($sel.Count -eq 0) {
        if ($selAll.Count -gt 0) { Show-Message 'Seçili güncellemelerin tamamı şu an başka bir işlemde (kilitli). İşlem bitince tekrar deneyin.' 'Uyarı' 'Warning' }
        else                     { Show-Message 'Önce soldaki listeden en az bir güncelleme seçin.' 'Uyarı' 'Warning' }
        return
    }
    if ($selAll.Count -gt $sel.Count) {
        Write-Log "$($selAll.Count - $sel.Count) kilitli (işlemdeki) güncelleme seçimden otomatik çıkarıldı." 'WARN'
    }

    $current = @{}
    foreach ($row in $gridGrp.Rows) { $current[[string]$row.Tag] = ConvertTo-CheckState $row.Cells['gChk'].Value }

    $plan = @(Get-MembershipPlan -SelectedIds $sel -Members $script:Members -Orig $script:OrigCheck -Current $current)
    # SUG'lar kilitlenmez: başka bir iş aynı SUG üzerinde çalışıyor olsa da bu plan uygulanabilir
    # (işlemdeki güncellemeler kilitli olduğundan seçime giremez; aynı SUG'a yazım worker'da sıraya alınır).
    if ($plan.Count -eq 0) { Show-Message 'Grup kutularında değişiklik yok.' 'Bilgi' 'Information'; return }
    $plan = @($plan | Sort-Object { $script:GroupById[$_.GroupId].Name })

    $simulate = $chkSim.Checked
    $lines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($p in $plan) {
        $lines.Add(('- {0}:  +{1} ekle,  -{2} çıkar' -f $script:GroupById[$p.GroupId].Name, $p.Add.Count, $p.Remove.Count))
    }
    if ($simulate) { $head = "SİMÜLASYON - hiçbir değişiklik yapılmayacak.`n`n" } else { $head = "" }
    $msg = "{0}{1} güncelleme için aşağıdaki SUG değişiklikleri uygulanacak:`n`n{2}`n`nİşlem arka planda çalışır; bu sırada diğer güncellemelerle çalışmaya devam edebilirsiniz.`n`nDevam edilsin mi?" -f $head, $sel.Count, ($lines -join "`n")
    if ([System.Windows.Forms.MessageBox]::Show($msg, 'SUG üyeliklerini uygula', 'YesNo', 'Question') -ne 'Yes') { return }

    try { $conn = Get-Connection }
    catch { Show-Message $_.Exception.Message 'Hata' 'Error'; return }

    # Yalnızca planda gerçekten eklenecek / çıkarılacak güncellemeler ve ilgili SUG'lar kilitlenir.
    $jobUpd = New-Object 'System.Collections.Generic.HashSet[string]'
    $planList = New-Object 'System.Collections.Generic.List[object]'
    foreach ($p in $plan) {
        foreach ($x in $p.Add)    { [void]$jobUpd.Add([string]$x) }
        foreach ($x in $p.Remove) { [void]$jobUpd.Add([string]$x) }
        $planList.Add([pscustomobject]@{
            GroupId   = [string]$p.GroupId
            GroupName = [string]$script:GroupById[$p.GroupId].Name
            Add       = [string[]]$p.Add
            Remove    = [string[]]$p.Remove
        })
    }
    $planData = $planList.ToArray()
    $label = "SUG üyelikleri ($($plan.Count) grup, $($jobUpd.Count) güncelleme)"
    if ($simulate) { $label = "[SİMÜLASYON] $label"; $status = 'SUG simülasyonu' } else { $status = 'SUG güncelleniyor' }

    Start-BackgroundJob -Kind 'Sug' -Label $label -Script $script:SugWorkerScript -Data @{ Plan = $planData } `
        -Conn $conn -Simulate $simulate -UpdateIds ([string[]]@($jobUpd)) -GroupIds ([string[]]@($planData | ForEach-Object { $_.GroupId })) `
        -StatusText $status
}

# ------------------------------------------------------------------
# Package içeriği silme yardımcıları (arka plan runspace'inde çalışır)
# ------------------------------------------------------------------
# Package'ın (SMS_PackageToContent) ŞU AN içerdiği ContentID'leri döndürür (gerçek durum kaynağı).
function Get-PackageContentIdSet {
    param([string]$PkgId)
    Assert-PackageId $PkgId
    $h = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($r in (Invoke-CMWql "SELECT ContentID FROM SMS_PackageToContent WHERE PackageID = '$PkgId'")) {
        [void]$h.Add([string]$r.ContentID)
    }
    return ,$h
}

# Package'ta, verilen package'ın tüm (CI_ID, ContentID) eşleşmelerini okur.
function Get-PackageContentPairs {
    param([string]$PkgId)
    Assert-PackageId $PkgId
    return @(Invoke-CMWql "SELECT cc.CI_ID, cc.ContentID FROM SMS_CIToContent AS cc JOIN SMS_PackageToContent AS pc ON pc.ContentID = cc.ContentID WHERE pc.PackageID = '$PkgId'")
}

# RemoveContent'i, parametre adlarını WMI'den okuyarak (adlara/sıraya güvenmeden) çağırır.
function Invoke-PkgRemoveContent {
    param($Pkg, [uint32[]]$Ids)
    $in = $Pkg.GetMethodParameters('RemoveContent')
    $arrProp = $null
    $boolProp = $null
    $desc = New-Object 'System.Collections.Generic.List[string]'
    foreach ($pr in $in.Properties) {
        $desc.Add(('{0}:{1}{2}' -f $pr.Name, $pr.Type, $(if ($pr.IsArray) { '[]' } else { '' })))
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

# Verilen ContentID'ler package'tan kalkana kadar bekler; geriye KALAN sayısını döner.
# Artan aralıklı (backoff) yoklama: 0,5 sn ile başlar, her seferinde iki katına çıkar (en fazla MaxDelayMs),
# toplam bekleme MaxWaitSec'i geçince durur. Hızlı silinen içerikte erken döner, yavaşta sunucuyu daha az sorgular.
function Wait-ContentRemoved {
    param([string]$PkgId, [uint32[]]$ContentIds, [int]$MaxWaitSec = 20, [int]$FirstDelayMs = 500, [int]$MaxDelayMs = 5000)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $delay = $FirstDelayMs
    while ($true) {
        $cur = Get-PackageContentIdSet $PkgId
        $left = @($ContentIds | Where-Object { $cur.Contains([string]$_) }).Count
        if ($left -eq 0) { return 0 }
        $remainMs = [int]($MaxWaitSec * 1000 - $sw.Elapsed.TotalMilliseconds)
        if ($remainMs -le 0) { return $left }
        if ($global:JobSync -and $global:JobSync.Cancel) { return $left }
        # Arka plan runspace'inde çalışır: bekleme UI thread'ini bloke etmez. Son bekleme süre sınırına kırpılır.
        Start-Sleep -Milliseconds ([Math]::Min($delay, $remainMs))
        $delay = [Math]::Min($delay * 2, $MaxDelayMs)
    }
}

# İçeriği siler ve SONUCU DOĞRULAR. Önce ContentID ile dener; içerik hâlâ duruyorsa ve güvenliyse
# güncellemelerin CI_ID'leri ile tekrar dener (RemoveContent'in ID türü sürüme göre farklı yorumlanabilir).
# ID çakışması (bir ContentID'nin başka bir güncellemenin CI_ID'si ile aynı sayı olması) durumunda
# yanlış içeriğin silinmemesi için ilgili ID'ler atlanır.
function Remove-PackageContentVerified {
    param($Pkg, [string]$PkgId, [uint32[]]$ContentIds, [string[]]$CiIds, [object[]]$Pairs)
    Assert-PackageId $PkgId

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
    $targets = New-Object 'System.Collections.Generic.List[uint32]'
    $skipped = 0
    foreach ($c in $ContentIds) {
        $cs = [string]$c
        if ($ciAll.Contains($cs) -and -not $selSet.ContainsKey($cs)) { $skipped++; continue }
        $targets.Add([uint32]$c)
    }
    if ($skipped -gt 0) {
        Write-Log "[$PkgId] $skipped içerik ID'si başka bir güncellemenin CI_ID'si ile çakıştığı için güvenlik amacıyla atlandı." 'WARN'
    }
    $left = $targets.Count
    if ($targets.Count -gt 0) {
        $rv = Invoke-PkgRemoveContent -Pkg $Pkg -Ids $targets.ToArray()
        Write-Log "[$PkgId] RemoveContent (ContentID ile, $($targets.Count) içerik) çağrıldı. Dönüş: [$rv]. Sonuç doğrulanıyor..."
        $left = Wait-ContentRemoved -PkgId $PkgId -ContentIds $targets.ToArray()
        if ($left -eq 0) { $method = 'ContentID' }
    }

    # --- 2. deneme: CI_ID ile (yalnızca 1. deneme etkisiz kaldıysa) ---
    if ($targets.Count -gt 0 -and $left -gt 0) {
        Write-Log "[$PkgId] ContentID ile çağrı hata vermedi ama $left içerik hâlâ package'ta. CI_ID ile deneniyor..." 'WARN'
        $inPkg = Get-PackageContentIdSet $PkgId
        $cand = New-Object 'System.Collections.Generic.List[uint32]'
        $candContents = New-Object 'System.Collections.Generic.HashSet[uint32]'
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
            $cand.Add([uint32]$ciKey)
            foreach ($ctKey in $contentsByCi[$ciKey]) { [void]$candContents.Add([uint32]$ctKey) }
        }
        if ($cand.Count -gt 0) {
            $rv = Invoke-PkgRemoveContent -Pkg $Pkg -Ids $cand.ToArray()
            Write-Log "[$PkgId] RemoveContent (CI_ID ile, $($cand.Count) güncelleme) çağrıldı. Dönüş: [$rv]. Sonuç doğrulanıyor..."
            [void](Wait-ContentRemoved -PkgId $PkgId -ContentIds ([uint32[]]@($candContents)))
            $method = 'CI_ID'
        } else {
            Write-Log "[$PkgId] CI_ID ile denenebilecek güvenli aday bulunamadı." 'WARN'
        }
    }

    # --- Nihai durum: package'ta gerçekte ne kaldı? ---
    $cur = Get-PackageContentIdSet $PkgId
    $remaining = @($ContentIds | Where-Object { $cur.Contains([string]$_) }).Count
    # Current: package'ın son içerik kümesi (hangi güncellemelerin package'tan tamamen çıktığını hesaplamak için)
    return [pscustomobject]@{ Total = $ContentIds.Count; Remaining = $remaining; Method = $method; Current = $cur }
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
    $list = @($allOthers | Where-Object { Test-CiId $_ })
    for ($i = 0; $i -lt $list.Count; $i += 40) {
        $last = [Math]::Min($i + 39, $list.Count - 1)
        $inList = ($list[$i..$last]) -join ','
        foreach ($r in (Invoke-CMWql "SELECT CI_ID, LocalizedDisplayName, IsSuperseded, IsExpired, IsDeployed FROM SMS_SoftwareUpdate WHERE CI_ID IN ($inList)")) { $meta[[string]$r.CI_ID] = $r }
    }

    $free = New-Object 'System.Collections.Generic.List[uint32]'
    $blocked = @{}
    $lines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($c in @($othersByContent.Keys)) {
        $blockers = New-Object 'System.Collections.Generic.List[string]'
        foreach ($ci in $othersByContent[$c]) {
            $r = $meta[$ci]
            if ($null -eq $r) {
                $lines.Add(("  ContentID {0} <- CI_ID {1}: SMS_SoftwareUpdate'te kayıt yok (silinmiş veya güncelleme değil); engel sayılmadı." -f $c, $ci))
                continue
            }
            $sup = [bool]$r.IsSuperseded
            $exp = [bool]$r.IsExpired
            $dep = [bool]$r.IsDeployed
            $st = New-Object 'System.Collections.Generic.List[string]'
            if ($sup) { $st.Add('Superseded') }
            if ($exp) { $st.Add('Expired') }
            if (-not $sup -and -not $exp) { $st.Add('AKTİF') }
            if ($dep) { $st.Add('Deploy edilmiş') }
            $block = ($dep -or (-not $sup -and -not $exp))
            $title = [string]$r.LocalizedDisplayName
            $lines.Add(("  ContentID {0} <- CI_ID {1} | {2} | {3} | {4}" -f $c, $ci, ($st -join ', '), $(if ($block) { 'ENGEL' } else { 'engel değil' }), $title))
            if ($block) { $blockers.Add(("CI_ID {0}: {1} [{2}]" -f $ci, $title, ($st -join ', '))) }
        }
        if ($blockers.Count -gt 0) { $blocked[$c] = $blockers.ToArray() } else { $free.Add([uint32]$c) }
    }
    return [pscustomobject]@{ Free = [uint32[]]$free.ToArray(); Blocked = $blocked; Lines = $lines.ToArray() }
}

function Invoke-RemoveFromPackage {
    $selAll = @(Get-SelectedUpdateIds -IncludeLocked)
    $sel = @(Get-SelectedUpdateIds)
    if ($sel.Count -eq 0) {
        if ($selAll.Count -gt 0) { Show-Message 'Seçili güncellemelerin tamamı şu an başka bir işlemde (kilitli). İşlem bitince tekrar deneyin.' 'Uyarı' 'Warning' }
        else                     { Show-Message 'Önce soldaki listeden en az bir güncelleme seçin.' 'Uyarı' 'Warning' }
        return
    }
    if ($selAll.Count -gt $sel.Count) {
        Write-Log "$($selAll.Count - $sel.Count) kilitli (işlemdeki) güncelleme seçimden otomatik çıkarıldı." 'WARN'
    }

    # "Hariç" işaretli package'lar: değer = bu package'tan silinmeyecek seçili güncelleme sayısı
    $excl = @{}
    foreach ($k in @($script:ExcludedPkgs.Keys)) { if ($script:ExcludedPkgs[$k]) { $excl[[string]$k] = 0 } }
    # Başka bir arka plan işinin üzerinde çalıştığı package'lar: değer = atlanan seçili güncelleme sayısı
    $lockedSkip = @{}

    $byPkg = @{}
    $stillInSug = 0
    $noPkg = 0
    foreach ($id in $sel) {
        $u = $script:Updates[$id]
        if (-not $u -or $u.PkgIds.Count -eq 0) { $noPkg++; continue }
        $added = $false
        foreach ($p in $u.PkgIds) {
            if ($excl.ContainsKey($p)) { $excl[$p]++; continue }
            if ($script:LockedPkgs.ContainsKey($p)) {
                if ($lockedSkip.ContainsKey($p)) { $lockedSkip[$p]++ } else { $lockedSkip[$p] = 1 }
                continue
            }
            if ($p -notmatch '^[A-Za-z0-9]+$') { Write-Log "Geçersiz PackageID atlandı: '$p'" 'WARN'; continue }
            if (-not $byPkg.ContainsKey($p)) { $byPkg[$p] = New-Object 'System.Collections.Generic.List[string]' }
            $byPkg[$p].Add($id)
            $added = $true
        }
        if ($added -and $script:Members.ContainsKey($id) -and $script:Members[$id].Count -gt 0) { $stillInSug++ }
    }
    $exTotal = 0
    foreach ($k in @($excl.Keys)) { $exTotal += $excl[$k] }
    if ($byPkg.Count -eq 0) {
        if ($lockedSkip.Count -gt 0) {
            Show-Message "Seçili güncellemelerin bulunduğu package'lar şu an başka bir işlemde (kilitli) veya 'Hariç' olarak işaretli. İşlem bitince tekrar deneyin." 'Bilgi' 'Information'
        } elseif ($exTotal -gt 0) {
            Show-Message "Seçili güncellemelerin bulunduğu tüm package'lar 'Hariç' olarak işaretli; silinecek bir şey yok." 'Bilgi' 'Information'
        } else {
            Show-Message "Seçili güncellemelerin hiçbiri state'teki Deployment Package içinde değil (yalnızca SUG üyesi)." 'Uyarı' 'Warning'
        }
        return
    }

    $simulate = $chkSim.Checked
    $lines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($p in $byPkg.Keys) { $lines.Add(('- {0}: {1} güncelleme' -f $p, $byPkg[$p].Count)) }
    $msg = ''
    if ($simulate) { $msg += "SİMÜLASYON - hiçbir değişiklik yapılmayacak.`n`n" }
    $msg += "Aşağıdaki Deployment Package'lardan seçili güncellemelerin İÇERİĞİ silinecek:`n`n" + ($lines -join "`n")
    if ($noPkg -gt 0) { $msg += "`n`n$noPkg seçili güncelleme package içinde olmadığı için atlanacak." }
    if ($exTotal -gt 0) {
        $exLines = New-Object 'System.Collections.Generic.List[string]'
        foreach ($k in @($excl.Keys)) {
            if ($excl[$k] -gt 0) {
                $nm = $k
                if ($script:PkgNames.ContainsKey($k)) { $nm = "$k ($($script:PkgNames[$k]))" }
                $exLines.Add(('- {0}: {1} güncelleme' -f $nm, $excl[$k]))
            }
        }
        $msg += "`n`nHARİÇ TUTULANLAR (bu package'lardan SİLİNMEYECEK):`n" + ($exLines -join "`n")
    }
    if ($lockedSkip.Count -gt 0) {
        $lkLines = New-Object 'System.Collections.Generic.List[string]'
        foreach ($k in @($lockedSkip.Keys)) { $lkLines.Add(('- {0}: {1} güncelleme (iş #{2} çalışıyor)' -f $k, $lockedSkip[$k], $script:LockedPkgs[$k])) }
        $msg += "`n`nKİLİTLİ PACKAGE'LAR (başka bir işlem sürdüğü için bu sefer ATLANACAK):`n" + ($lkLines -join "`n")
    }
    if ($stillInSug -gt 0) {
        $msg += "`n`nUYARI: $stillInSug güncelleme hâlâ en az bir SUG'a üye. Önce SUG üyeliklerinden çıkarmanız önerilir."
    }
    $msg += "`n`nİşlem arka planda çalışır; bu sırada diğer güncellemelerle çalışmaya devam edebilirsiniz."
    if (-not $simulate) {
        $msg += "`n`nBu işlem geri alınamaz; dağıtım noktalarına güncelleme tetiklenir. Devam edilsin mi?"
    } else {
        $msg += "`n`nDevam edilsin mi?"
    }
    if ([System.Windows.Forms.MessageBox]::Show($msg, "Package'tan sil", 'YesNo', 'Warning', 'Button2') -ne 'Yes') { return }

    try { $conn = Get-Connection }
    catch { Show-Message $_.Exception.Message 'Hata' 'Error'; return }

    # Yalnızca bu işe giren güncellemeler ve package'lar kilitlenir (kısmi kilitleme).
    $jobUpd = New-Object 'System.Collections.Generic.HashSet[string]'
    $byPkgData = @{}
    foreach ($k in @($byPkg.Keys)) {
        $byPkgData[$k] = [string[]]$byPkg[$k].ToArray()
        foreach ($x in $byPkg[$k]) { [void]$jobUpd.Add($x) }
    }
    $label = "Package'tan silme ($($byPkg.Count) package, $($jobUpd.Count) güncelleme)"
    if ($simulate) { $label = "[SİMÜLASYON] $label"; $status = 'Silme simülasyonu' } else { $status = 'Package''tan siliniyor' }

    Start-BackgroundJob -Kind 'Pkg' -Label $label -Script $script:PkgWorkerScript -Data @{ ByPkg = $byPkgData } `
        -Conn $conn -Simulate $simulate -UpdateIds ([string[]]@($jobUpd)) -PkgIds ([string[]]@($byPkg.Keys)) `
        -StatusText $status
}

# ------------------------------------------------------------------
# Arka plan işleri (Runspace) - Listele, Package silme ve SUG üyelik işlemleri
# ------------------------------------------------------------------
# Mimari:
#   * Her iş kendi Runspace'inde çalışır; UI thread'i bloke olmaz.
#   * UI ile iş arasındaki tüm veri [hashtable]::Synchronized ($job.Sync) ve ConcurrentQueue'lar üzerinden akar:
#       Sync.Log     : günlük satırları (paylaşılan $script:LogQueue)
#       Sync.Events  : sonuç olayları (PkgResult / SugResult) -> UI bellekteki veriyi YEREL olarak günceller
#       Sync.Result  : Listele / Ortamı Oku işinin sonucu (UI, Complete-ListJob ile uygular)
#       Sync.Question/Answer : işin kullanıcıya soru sorması (ör. paylaşılan içerik onayı)
#       Sync.Cancel  : pencere kapatılırken işe iptal bildirimi
#       Sync.RunId   : denetim kaydındaki (audit) iş kimliği (GUID)
#   * UI tarafı $script:PackageJobs ve $script:ListJob'ı bir System.Windows.Forms.Timer ile periyodik okur.
#   * Arka plan kodu ASLA Windows Forms kontrollerine dokunmaz. Tüm SCCM erişimi WMI (Get-WmiObject) iledir.

# Arka plan runspace'ine aktarılan (UI'dan bağımsız) fonksiyonlar.
$script:WorkerSharedFunctions = @(
    # WMI erişimi, doğrulama, audit
    'Get-WmiConnectionArgs', 'Invoke-CMWql', 'Get-CMWmiInstance',
    'Test-PackageId', 'Test-CiId', 'Test-CategoryUniqueId', 'Assert-PackageId', 'Assert-CiId', 'Get-UpdateSelectList',
    'Get-AuditFilePath', 'Write-Audit',
    # Package silme
    'Get-PackageContentIdSet', 'Get-PackageContentPairs', 'Invoke-PkgRemoveContent', 'Wait-ContentRemoved',
    'Remove-PackageContentVerified', 'Get-ContentRemovalSet', 'Get-SharedContentDetails',
    # SUG üyeliği
    'Set-SugMembershipWmi', 'Set-SugMembershipWmiCore',
    # Listeleme
    'Import-Groups', 'Import-PackageCatalog', 'Import-Environment', 'Import-VendorCategory', 'Add-UpdateRecord',
    'ConvertTo-DateText', 'Import-Memberships', 'Import-UpdatePackages', 'Get-DeployedCIIds', 'Test-UpdateDeployed'
)

# Runspace içinde UI fonksiyonlarının yerini alan yardımcılar (aynı adla tanımlanır).
$script:WorkerHelpers = @{
    'Write-Log' = {
        param([string]$Message, [string]$Level = 'INFO')
        $global:JobSync.Log.Enqueue(('[{0}] [{1}] [{2}] {3}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $global:JobSync.LogTag, $Message))
    }
    'Get-Connection' = {
        # UI kontrollerini okumak yerine iş başlarken alınan bağlantı bilgisinin kopyası kullanılır.
        return $global:JobSync.Conn
    }
    'Send-JobEvent' = {
        param([hashtable]$JobEvent)
        $global:JobSync.Events.Enqueue($JobEvent)
    }
    'Request-UserConfirm' = {
        # Soruyu UI'a bırakır ve cevap gelene kadar bekler (MessageBox UI thread'inde zamanlayıcı ile gösterilir).
        param([string]$Text, [string]$Title)
        $global:JobSync.Answer = $null
        $global:JobSync.Question = @{ Text = $Text; Title = $Title }
        while ($null -eq $global:JobSync.Answer) {
            if ($global:JobSync.Cancel) { return $false }
            Start-Sleep -Milliseconds 200
        }
        $a = [bool]$global:JobSync.Answer
        $global:JobSync.Answer = $null
        return $a
    }
}

# Listele / Ortamı Oku işi (arka planda). Veri fonksiyonları runspace'in kendi $script: kapsamında çalışır;
# sonuç Sync.Result ile UI'a verilir ve UI bunları kendi $script: değişkenlerine atar (Complete-ListJob).
$script:ListWorkerScript = {
    param($Sync)
    $global:JobSync = $Sync
    try {
        $script:Updates     = @{}
        $script:Members     = @{}
        $script:DeployedIds = @{}

        if ($Sync.EnvOnly) {
            Import-Environment
            $Sync.Result = @{
                EnvOnly = $true; Groups = $script:Groups; GroupById = $script:GroupById
                EnvPackages = $script:EnvPackages; PkgNames = $script:PkgNames
            }
            return
        }

        $mode = [int]$Sync.Mode
        $pkgIds = @($Sync.PkgIds)
        $sugIds = @($Sync.SugIds)
        if ($mode -eq 3) {
            Write-Log 'Tüm ortam taranıyor; Package/SUG sayısına göre biraz sürebilir...' 'WARN'
            Import-Environment
            $pkgIds = @($script:EnvPackages | ForEach-Object { $_.Id })
            $sugIds = @($script:Groups | ForEach-Object { $_.Id })
        } else {
            Write-Log 'SUG listesi okunuyor...'
            Import-Groups
            Write-Log "$($script:Groups.Count) Software Update Group bulundu."
            Import-PackageCatalog      # Package adları + Software Update Package kümesi (tek sorgu, önbellek)
        }

        $script:VendorFilter    = [string]$Sync.Vendor
        $script:VendorSkipped   = 0
        $script:VendorNoInfo    = 0
        $script:VendorUniqueIds = $null
        $script:VendorNames     = $null
        $script:VendorCiIds     = $null
        $script:VendorSample    = New-Object 'System.Collections.Generic.List[string]'
        if ($script:VendorFilter) {
            Write-Log "Vendor filtresi: '$($script:VendorFilter)'"
            try { Import-VendorCategory -Vendor $script:VendorFilter }
            catch {
                Write-Log "Company kategorisi okunamadı, yalnızca ad eşleşmesi kullanılacak: $($_.Exception.Message)" 'WARN'
                $script:VendorNames = New-Object 'System.Collections.Generic.HashSet[string]'
                [void]$script:VendorNames.Add($script:VendorFilter.ToLowerInvariant())
            }
        } else { Write-Log 'Vendor filtresi yok (tüm vendor''lar).' }
        if ($Sync.Expired) { $flt = '(su.IsSuperseded = 1 OR su.IsExpired = 1)' }
        else               { $flt = 'su.IsSuperseded = 1' }
        $cols = Get-UpdateSelectList 'su'

        # Ortamdaki TÜM superseded güncellemeler (mod 4, 5 ve 6): Package/SUG gezmeden TEK sorgu.
        # Mod 4: deploy filtresi, üyelikler okunduktan sonra aşağıda bellekte uygulanır.
        # Mod 5: Package filtresi burada uygulanır (içeriği hiçbir Deployment Package'ta olmayanlar elenir);
        #        ardından Mod 4 ile aynı deploy filtresi de uygulanır (indirilmiş + superseded + deploy edilmiş).
        # Mod 6: Mod 5'in yalnızca Package filtresi uygulanır; deploy filtresi YOKTUR (indirilmiş + superseded).
        $scopeSug = New-Object 'System.Collections.Generic.List[string]'
        if ($mode -eq 4 -or $mode -eq 5 -or $mode -eq 6) {
            Write-Log 'Ortamdaki superseded güncellemeler tek sorguyla okunuyor...'
            $rows = Invoke-CMWql "SELECT $cols FROM SMS_SoftwareUpdate AS su WHERE $flt"
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
            foreach ($pkg in $pkgIds) {
                if ($Sync.Cancel) { throw 'İptal edildi.' }
                if (-not (Test-PackageId $pkg)) { Write-Log "Geçersiz PackageID atlandı: '$pkg'" 'WARN'; continue }
                Write-Log "Package içeriği taranıyor: $pkg"
                $q = "SELECT $cols FROM SMS_SoftwareUpdate AS su JOIN SMS_CIToContent AS cc ON su.CI_ID = cc.CI_ID JOIN SMS_PackageToContent AS pc ON pc.ContentID = cc.ContentID WHERE pc.PackageID = '$pkg' AND $flt"
                $rows = Invoke-CMWql $q
                foreach ($r in $rows) { Add-UpdateRecord -Row $r -PkgId $pkg }
                Write-Log "  $($pkg): $($rows.Count) satır (aynı güncellemenin birden fazla içeriği tekrar sayılabilir)."
            }

            foreach ($sug in $sugIds) {
                if ($Sync.Cancel) { throw 'İptal edildi.' }
                if (-not (Test-CiId $sug)) { Write-Log "Geçersiz SUG CI_ID atlandı: '$sug'" 'WARN'; continue }
                $scopeSug.Add([string]$sug)
                $gname = $sug
                if ($script:GroupById.ContainsKey($sug)) { $gname = $script:GroupById[$sug].Name }
                Write-Log "SUG üyeleri taranıyor: $gname ($sug)"
                $q = "SELECT $cols FROM SMS_SoftwareUpdate AS su JOIN SMS_CIRelation AS cr ON su.CI_ID = cr.ToCIID WHERE cr.FromCIID = $sug AND cr.RelationType = 1 AND $flt"
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

        $Sync.Result = @{
            EnvOnly     = $false
            Mode        = $mode
            Groups      = $script:Groups
            GroupById   = $script:GroupById
            EnvPackages = $script:EnvPackages
            PkgNames    = $script:PkgNames
            Updates     = $script:Updates
            Members     = $script:Members
            DeployedIds = $script:DeployedIds
            ScopeSugIds = [string[]]$scopeSug.ToArray()
        }
    } catch {
        $Sync.Error = $_.Exception.Message
        Write-Log $_.Exception.Message 'ERROR'
    }
}

# Package içeriği silme işi (arka planda). Her package için bir PkgResult olayı üretir.
$script:PkgWorkerScript = {
    param($Sync)
    $global:JobSync = $Sync
    $simulate = [bool]$Sync.Simulate
    $runId = [string]$Sync.RunId
    $failed = 0
    try {
        foreach ($pkgId in @($Sync.ByPkg.Keys)) {
            if ($Sync.Cancel) { Write-Log 'İşlem iptal edildi; kalan package''lar atlandı.' 'WARN'; break }
            $ids = [string[]]$Sync.ByPkg[$pkgId]
            $ev = @{ Type = 'PkgResult'; PkgId = [string]$pkgId; RemovedIds = [string[]]@(); Ok = $true }
            $auditStarted = $false
            try {
                Assert-PackageId $pkgId

                Write-Log "[$pkgId] içerik eşleşmeleri okunuyor..."
                $pairs = Get-PackageContentPairs $pkgId
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
                                $bl = New-Object 'System.Collections.Generic.List[string]'
                                foreach ($k in $blockedKeys) { foreach ($b in $det.Blocked[$k]) { if (-not $bl.Contains($b)) { $bl.Add($b) } } }
                                $txt = "[$pkgId] $($blockedKeys.Count) içerik, aşağıdaki aktif/deploy edilmiş güncellemelerle paylaşılıyor:`n`n" +
                                       (($bl | Select-Object -First 8) -join "`n")
                                if ($bl.Count -gt 8) { $txt += "`n... ve $($bl.Count - 8) güncelleme daha." }
                                $txt += "`n`nBu içeriği silerseniz bu güncellemeler bu package'ta içeriksiz kalır. Yine de silinsin mi?"
                                if (Request-UserConfirm -Text $txt -Title 'Paylaşılan içerik') {
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

                $pkg = @(Get-CMWmiInstance -Class 'SMS_SoftwareUpdatesPackage' -Filter "PackageID = '$pkgId'")
                if ($pkg.Count -eq 0) { throw "Deployment Package bulunamadı: $pkgId" }
                $pkg = $pkg[0]

                Write-Audit 'PkgRemoveStart' @{ RunId = $runId; JobId = $Sync.JobId; PackageId = $pkgId; UpdateIds = $ids; ContentIds = [string[]]@($set.Remove) }
                $auditStarted = $true
                Write-Log "[$pkgId] $($set.Remove.Count) içerik siliniyor..."
                $res = Remove-PackageContentVerified -Pkg $pkg -PkgId $pkgId -ContentIds $set.Remove -CiIds $ids -Pairs $pairs

                # Hangi seçili güncellemelerin bu package'taki TÜM içeriği kalktı? (UI'daki yerel güncelleme için)
                $byCi = @{}
                foreach ($pr in $pairs) {
                    $ciKey = [string]$pr.CI_ID
                    if (-not $byCi.ContainsKey($ciKey)) { $byCi[$ciKey] = New-Object 'System.Collections.Generic.List[string]' }
                    $byCi[$ciKey].Add([string]$pr.ContentID)
                }
                $gone = New-Object 'System.Collections.Generic.List[string]'
                foreach ($uid in $ids) {
                    $allGone = $true
                    if ($byCi.ContainsKey($uid)) {
                        foreach ($ct in $byCi[$uid]) { if ($res.Current.Contains($ct)) { $allGone = $false; break } }
                    }
                    if ($allGone) { $gone.Add($uid) }
                }
                $ev.RemovedIds = [string[]]$gone.ToArray()

                if ($res.Remaining -eq 0) { $result = 'OK' } else { $result = 'PARTIAL' }
                Write-Audit 'PkgRemoveResult' @{ RunId = $runId; JobId = $Sync.JobId; PackageId = $pkgId; Result = $result; Total = $res.Total; Remaining = $res.Remaining; Method = $res.Method; RemovedUpdateIds = $ev.RemovedIds }
                $auditStarted = $false

                if ($res.Remaining -eq 0) {
                    Write-Log "[$pkgId] $($res.Total) içerik package'tan silindi ve DOĞRULANDI (yöntem: $($res.Method))." 'OK'
                } else {
                    $failed++
                    $ev.Ok = $false
                    Write-Log "[$pkgId] $($res.Remaining)/$($res.Total) içerik HÂLÂ package'ta duruyor. Dağıtım Yöneticisi işlemi sürdürüyor olabilir; birkaç dakika sonra 'Listele' ile tekrar kontrol edin. Kalıyorsa SMS Provider sunucusunda smsprov.log ve site sunucusunda distmgr.log kayıtlarına bakın." 'ERROR'
                }
            } catch {
                $failed++
                $ev.Ok = $false
                Write-Log "[$pkgId] $($_.Exception.Message)" 'ERROR'
                if ($auditStarted) { Write-Audit 'PkgRemoveResult' @{ RunId = $runId; JobId = $Sync.JobId; PackageId = $pkgId; Result = 'ERROR'; Error = $_.Exception.Message } }
            } finally {
                Send-JobEvent $ev
            }
        }
    } catch {
        $failed++
        Write-Log $_.Exception.Message 'ERROR'
    } finally {
        $Sync.FailedCount = $failed
    }
}

# SUG üyelik işi (arka planda, doğrudan WMI). Her SUG için bir SugResult olayı üretir.
# ConfigurationManager modülü yüklenmez; SMS_AuthorizationList.Updates okunup Put() ile yazılır ve doğrulanır.
$script:SugWorkerScript = {
    param($Sync)
    $global:JobSync = $Sync
    $simulate = [bool]$Sync.Simulate
    $runId = [string]$Sync.RunId
    $failed = 0
    try {
        foreach ($p in @($Sync.Plan)) {
            if ($Sync.Cancel) { Write-Log 'İşlem iptal edildi; kalan SUG''lar atlandı.' 'WARN'; break }
            $gname = $p.GroupName
            $ev = @{ Type = 'SugResult'; GroupId = [string]$p.GroupId; Added = [string[]]@(); Removed = [string[]]@(); Ok = $true }
            $auditStarted = $false
            try {
                if ($simulate) {
                    Write-Log "[SİMÜLASYON] '$gname': +$($p.Add.Count) eklenecek, -$($p.Remove.Count) çıkarılacak." 'WARN'
                    continue
                }
                Write-Audit 'SugChangeStart' @{ RunId = $runId; JobId = $Sync.JobId; GroupId = $p.GroupId; GroupName = $gname; Add = $p.Add; Remove = $p.Remove }
                $auditStarted = $true
                $res = Set-SugMembershipWmi -GroupId $p.GroupId -Add $p.Add -Remove $p.Remove
                $ev.Added = $res.Added
                $ev.Removed = $res.Removed
                if ($res.Added.Count -gt 0)   { Write-Log "'$gname': $($res.Added.Count) güncelleme eklendi (doğrulandı)." 'OK' }
                if ($res.Removed.Count -gt 0) { Write-Log "'$gname': $($res.Removed.Count) güncelleme çıkarıldı (doğrulandı)." 'OK' }
                $result = 'OK'
                if ($res.NotAdded.Count -gt 0 -or $res.NotRemoved.Count -gt 0) {
                    $failed++
                    $ev.Ok = $false
                    $result = 'PARTIAL'
                    Write-Log "'$gname': değişiklik doğrulanamadı - eklenemeyen: $($res.NotAdded.Count), çıkarılamayan: $($res.NotRemoved.Count)." 'ERROR'
                }
                Write-Audit 'SugChangeResult' @{ RunId = $runId; JobId = $Sync.JobId; GroupId = $p.GroupId; GroupName = $gname; Result = $result; Added = $res.Added; Removed = $res.Removed; NotAdded = $res.NotAdded; NotRemoved = $res.NotRemoved }
                $auditStarted = $false
            } catch {
                $failed++
                $ev.Ok = $false
                Write-Log "'$gname' güncellenemedi: $($_.Exception.Message)" 'ERROR'
                if ($auditStarted) { Write-Audit 'SugChangeResult' @{ RunId = $runId; JobId = $Sync.JobId; GroupId = $p.GroupId; GroupName = $gname; Result = 'ERROR'; Error = $_.Exception.Message } }
            } finally {
                Send-JobEvent $ev
            }
        }
    } catch {
        $failed++
        Write-Log $_.Exception.Message 'ERROR'
    } finally {
        $Sync.FailedCount = $failed
    }
}

# Arka plan runspace'leri için oturum durumu (fonksiyon tanımlarıyla birlikte); bir kez kurulup önbelleğe alınır.
function Get-WorkerSessionState {
    if ($script:WorkerState) { return $script:WorkerState }
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    foreach ($fn in $script:WorkerSharedFunctions) {
        $def = (Get-Command -Name $fn -CommandType Function -ErrorAction Stop).Definition
        $iss.Commands.Add((New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $fn, $def))
    }
    foreach ($k in @($script:WorkerHelpers.Keys)) {
        $iss.Commands.Add((New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList $k, $script:WorkerHelpers[$k].ToString()))
    }
    $script:WorkerState = $iss
    return $iss
}

# Kind: 'Pkg' / 'Sug' (kilitli, paralel çalışabilir) veya 'List' (Listele / Ortamı Oku; tek, kilitsiz).
function Start-BackgroundJob {
    param(
        [string]$Kind,
        [string]$Label,
        [scriptblock]$Script,
        [hashtable]$Data,
        $Conn,
        [bool]$Simulate,
        [string[]]$UpdateIds = @(),
        [string[]]$PkgIds = @(),
        [string[]]$GroupIds = @(),
        [string]$StatusText = 'İşlemde'
    )
    $isList = ($Kind -eq 'List')
    $script:JobSeq++
    $id = $script:JobSeq

    $sync = [hashtable]::Synchronized(@{})
    $sync.JobId       = $id
    $sync.RunId       = [guid]::NewGuid().ToString()
    if ($isList) { $sync.LogTag = $Label } else { $sync.LogTag = "İş #$id" }
    $sync.Conn        = [pscustomobject]@{ Site = $Conn.Site; Server = $Conn.Server }
    $sync.Simulate    = $Simulate
    $sync.Log         = $script:LogQueue
    $sync.Events      = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
    $sync.Question    = $null
    $sync.Answer      = $null
    $sync.Cancel      = $false
    $sync.FailedCount = 0
    $sync.Result      = $null
    $sync.Error       = $null
    foreach ($k in @($Data.Keys)) { $sync[$k] = $Data[$k] }

    $rs = $null
    $ps = $null
    try {
        $rs = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace((Get-WorkerSessionState))
        $rs.ApartmentState = [System.Threading.ApartmentState]::STA
        $rs.ThreadOptions  = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $rs.Open()
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.Runspace = $rs
        [void]$ps.AddScript($Script.ToString()).AddArgument($sync)

        $job = [pscustomobject]@{
            Id       = $id
            Kind     = $Kind
            Label    = $Label
            Simulate = $Simulate
            Sync     = $sync
            PS       = $ps
            Runspace = $rs
            Handle   = $null
            Started  = Get-Date
        }
        if ($isList) {
            $script:ListJob = $job
        } else {
            $script:PackageJobs[$id] = $job
            Lock-JobItems -JobId $id -UpdateIds $UpdateIds -PkgIds $PkgIds -StatusText $StatusText
        }
        $job.Handle = $ps.BeginInvoke()
    } catch {
        $err = $_.Exception.Message
        if ($isList) { $script:ListJob = $null }
        elseif ($script:PackageJobs.ContainsKey($id)) {
            $script:PackageJobs.Remove($id)
            Unlock-JobItems -JobId $id
        }
        if ($ps) { try { $ps.Dispose() } catch { } }
        if ($rs) { try { $rs.Dispose() } catch { } }
        Write-Log "Arka plan işi başlatılamadı: $err" 'ERROR'
        Show-Message "Arka plan işi başlatılamadı: $err" 'Hata' 'Error'
        Update-ControlStates
        return
    }
    if ($isList) {
        Write-Log "$Label başlatıldı (arka planda)."
    } else {
        Write-Log "İş #$id başlatıldı (arka planda): $Label"
        if (-not $Simulate) {
            Write-Audit 'JobStart' @{
                RunId = $sync.RunId; JobId = $id; Kind = $Kind; Label = $Label
                Server = $sync.Conn.Server; Site = $sync.Conn.Site
                Packages = $PkgIds; Groups = $GroupIds; UpdateIds = $UpdateIds
            }
        }
    }
    Update-ControlStates
    if (-not $jobTimer.Enabled) { $jobTimer.Start() }
}

# Yalnızca işe alınan güncellemeleri (ve Package silme işinde package'ları) kilitler; kilitli satırlar seçimden çıkarılır.
# SUG'lar kilitlenmez (ortamda gruplar aynı anda kullanılabilmeli); GroupIds yalnızca audit kaydı içindir.
function Lock-JobItems {
    param([int]$JobId, [string[]]$UpdateIds, [string[]]$PkgIds, [string]$StatusText)
    $prev = $script:Suspend
    $script:Suspend = $true
    try {
        foreach ($uid in @($UpdateIds)) {
            if (-not $uid) { continue }
            $u = $script:Updates[$uid]
            if (-not $u) { continue }
            $u.IsLocked = $true
            $u.Status   = $StatusText
            $u.LockJob  = $JobId
            if ($script:ViewIndex.ContainsKey($uid)) { $gridUpd.Rows[$script:ViewIndex[$uid]].Selected = $false }
            Update-UpdateRow $uid
        }
    } finally {
        $script:Suspend = $prev
    }
    foreach ($pk in @($PkgIds)) {
        if (-not $pk) { continue }
        $script:LockedPkgs[$pk] = $JobId
        if ($script:PkgRowById.ContainsKey($pk)) { Set-PkgRowStyle $script:PkgRowById[$pk] }
    }
    # Seçim değişti (kilitli satırlar çıkarıldı): SUG kutularını ve "Seçili" sayılarını kalan seçime göre hesapla.
    Update-GroupChecks
    Update-PackageCounts
}

function Unlock-JobItems {
    param([int]$JobId)
    foreach ($u in @($script:Updates.Values)) {
        if ($u.IsLocked -and $u.LockJob -eq $JobId) {
            $u.IsLocked = $false
            $u.Status   = ''
            $u.LockJob  = 0
            Update-UpdateRow $u.Id
        }
    }
    foreach ($pk in @($script:LockedPkgs.Keys)) {
        if ($script:LockedPkgs[$pk] -ne $JobId) { continue }
        $script:LockedPkgs.Remove($pk)
        if ($script:PkgRowById.ContainsKey($pk)) { Set-PkgRowStyle $script:PkgRowById[$pk] }
    }
    # Kilit kalkınca listede güncellemesi kalmamış package satırları da temizlenir.
    Update-PackageTotals
    Update-CountLabel
}

# Güncelleme, son "Listele"deki kaynak/filtre koşullarına HÂLÂ uyuyor mu? (WMI sorgusu yapmadan, bellekten)
function Test-UpdateInListFilter {
    param([string]$Id)
    $u = $script:Updates[$Id]
    if (-not $u) { return $false }
    $hasPkg = ($u.PkgIds.Count -gt 0)
    switch ($script:ListMode) {
        4 { return (Test-UpdateDeployed $Id) }                     # superseded + deploy edilmiş
        5 { return ($hasPkg -and (Test-UpdateDeployed $Id)) }      # indirilmiş + superseded + deploy edilmiş
        6 { return $hasPkg }                                       # indirilmiş + superseded
        default {
            # State / tek Package / tek SUG / TÜM Package + SUG: kapsamdaki bir package'ta VEYA kapsamdaki bir SUG'da olmalı.
            # (Bu modlarda PkgIds yalnızca kapsamdaki package'ları içerir.)
            if ($hasPkg) { return $true }
            if ($script:Members.ContainsKey($Id)) {
                foreach ($g in $script:ListScopeSugIds) { if ($script:Members[$Id].Contains($g)) { return $true } }
            }
            return $false
        }
    }
}

# İş sonuçlarından etkilenen satırları YEREL olarak günceller (Invoke-List çağrılmaz).
# Filtreye artık uymayan güncellemeler listeden kaldırılır; Package sayıları bellekten yeniden hesaplanır.
function Sync-UpdateViews {
    param([string[]]$UpdateIds)
    $removedSet = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($uid in @($UpdateIds | Select-Object -Unique)) {
        if (-not $uid -or -not $script:Updates.ContainsKey($uid)) { continue }
        if (Test-UpdateInListFilter $uid) { Update-UpdateRow $uid; continue }
        [void]$script:Updates.Remove($uid)
        [void]$script:Members.Remove($uid)
        [void]$script:RowCache.Remove($uid)
        [void]$removedSet.Add($uid)
    }
    if ($removedSet.Count -gt 0) {
        $newIds = New-Object 'System.Collections.Generic.List[string]'
        foreach ($id in $script:ViewIds) { if (-not $removedSet.Contains($id)) { $newIds.Add($id) } }
        Set-UpdateView -Ids $newIds
        Write-Log "$($removedSet.Count) güncelleme artık liste kaynağının/filtresinin koşullarına uymadığı için listeden kaldırıldı (yerel güncelleme)."
    }
    Update-PackageTotals
    Update-CountLabel
}

# İşin ürettiği sonuç olaylarını okuyup bellekteki veriye uygular.
function Receive-JobEvents {
    param($Job)
    $ev = $null
    $affected = New-Object 'System.Collections.Generic.List[string]'
    while ($Job.Sync.Events.TryDequeue([ref]$ev)) {
        if ($Job.Simulate) { continue }     # simülasyonda veri değişmez
        if ($ev.Type -eq 'PkgResult') {
            $pk = [string]$ev.PkgId
            $n = 0
            foreach ($uid in @($ev.RemovedIds)) {
                if (-not $uid) { continue }
                $u = $script:Updates[$uid]
                if ($u -and $u.PkgIds.Remove($pk)) { $affected.Add($uid); $n++ }
            }
            if ($n -gt 0) { Write-Log "[$pk] $n güncellemenin bu package'taki içeriği kalktı; satırlar yerel olarak güncellendi." }
        }
        elseif ($ev.Type -eq 'SugResult') {
            $gid = [string]$ev.GroupId
            $added = @($ev.Added | Where-Object { $_ })
            $removedIds = @($ev.Removed | Where-Object { $_ })
            foreach ($uid in $added) {
                if (-not $script:Members.ContainsKey($uid)) { $script:Members[$uid] = New-Object 'System.Collections.Generic.HashSet[string]' }
                if ($script:Members[$uid].Add($gid)) { $affected.Add($uid) }
            }
            foreach ($uid in $removedIds) {
                if ($script:Members.ContainsKey($uid) -and $script:Members[$uid].Remove($gid)) { $affected.Add($uid) }
            }
            if ($script:GroupById.ContainsKey($gid) -and ($added.Count -gt 0 -or $removedIds.Count -gt 0)) {
                $g = $script:GroupById[$gid]
                $g.Total = [Math]::Max(0, [int]$g.Total + $added.Count - $removedIds.Count)
                if ($script:GrpRowById.ContainsKey($gid)) {
                    $prev = $script:Suspend
                    $script:Suspend = $true
                    try { $script:GrpRowById[$gid].Cells['gTotal'].Value = $g.Total } finally { $script:Suspend = $prev }
                }
            }
        }
    }
    if ($affected.Count -gt 0) { Sync-UpdateViews -UpdateIds $affected.ToArray() }
}

function Complete-BackgroundJob {
    param($Job)
    Receive-JobEvents $Job
    try { [void]$Job.PS.EndInvoke($Job.Handle) }
    catch {
        Write-Log "İş #$($Job.Id) beklenmeyen hatayla sonlandı: $($_.Exception.Message)" 'ERROR'
        $Job.Sync.FailedCount = [int]$Job.Sync.FailedCount + 1
    }
    try { foreach ($er in $Job.PS.Streams.Error) { Write-Log "İş #$($Job.Id): $er" 'ERROR' } } catch { }
    Receive-JobLog
    try { $Job.PS.Dispose() } catch { }
    try { $Job.Runspace.Close(); $Job.Runspace.Dispose() } catch { }

    $script:PackageJobs.Remove($Job.Id)
    Unlock-JobItems -JobId $Job.Id

    $dur = [int]((Get-Date) - $Job.Started).TotalSeconds
    $failed = [int]$Job.Sync.FailedCount
    if (-not $Job.Simulate) {
        Write-Audit 'JobEnd' @{ RunId = $Job.Sync.RunId; JobId = $Job.Id; Kind = $Job.Kind; Failed = $failed; DurationSec = $dur }
    }
    if ($Job.Simulate) {
        Write-Log "İş #$($Job.Id) tamamlandı ($dur sn): $($Job.Label). Ayrıntılar yukarıda." 'OK'
    } elseif ($failed -gt 0) {
        Write-Log "İş #$($Job.Id) tamamlandı ($dur sn) ancak $failed hata var: $($Job.Label)" 'WARN'
        if ($Job.Kind -eq 'Pkg') { $txt = "İş #$($Job.Id): $failed package'ta içerik silinemedi (hâlâ duruyor). Ayrıntılar günlükte." }
        else                     { $txt = "İş #$($Job.Id): $failed grup güncellenemedi. Ayrıntılar günlükte." }
        Show-Message $txt 'Uyarı' 'Warning'
    } else {
        Write-Log "İş #$($Job.Id) başarıyla tamamlandı ve doğrulandı ($dur sn): $($Job.Label)" 'OK'
    }
    Update-ControlStates
}

# ------------------------------------------------------------------
# Yarım kalmış işlerin tespiti (audit kaydından)
# ------------------------------------------------------------------
# JobStart kaydı olup JobEnd / JobAcknowledged kaydı olmayan işler "yarım kalmış" sayılır
# (pencere iptalle kapatıldı: JobAborted, ya da PowerShell beklenmedik şekilde kapandı).
# Performans: önce yalnızca "Job*" satırları ayrıştırılır; adım satırları yalnızca açık kalan RunId'ler için okunur.
function Get-UnfinishedAuditRuns {
    $path = Get-AuditFilePath
    $lines = New-Object 'System.Collections.Generic.List[string]'
    foreach ($f in @("$path.1", $path)) {
        if (-not (Test-Path -LiteralPath $f)) { continue }
        try { $lines.AddRange([System.IO.File]::ReadAllLines($f)) }
        catch { Write-Log "Denetim kaydı okunamadı ($f): $($_.Exception.Message)" 'WARN' }
    }
    $starts  = New-Object 'System.Collections.Generic.List[object]'
    $closed  = New-Object 'System.Collections.Generic.HashSet[string]'
    $aborted = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($ln in $lines) {
        if ($ln.IndexOf('"Action":"Job', [System.StringComparison]::Ordinal) -lt 0) { continue }
        try { $r = $ln | ConvertFrom-Json } catch { continue }
        switch ([string]$r.Action) {
            'JobStart'        { $starts.Add($r) }
            'JobEnd'          { [void]$closed.Add([string]$r.RunId) }
            'JobAborted'      { [void]$aborted.Add([string]$r.RunId) }
            'JobAcknowledged' { foreach ($x in @($r.RunIds)) { [void]$closed.Add([string]$x) } }
        }
    }
    $open = @($starts | Where-Object { -not $closed.Contains([string]$_.RunId) })
    $result = New-Object 'System.Collections.Generic.List[object]'
    foreach ($s in $open) {
        $rid = [string]$s.RunId
        $steps = New-Object 'System.Collections.Generic.List[object]'
        foreach ($ln in $lines) {
            if ($ln.IndexOf($rid, [System.StringComparison]::Ordinal) -lt 0) { continue }
            if ($ln.IndexOf('"Action":"Job', [System.StringComparison]::Ordinal) -ge 0) { continue }
            try { $steps.Add(($ln | ConvertFrom-Json)) } catch { }
        }
        $result.Add([pscustomobject]@{ Start = $s; Aborted = $aborted.Contains($rid); Steps = $steps.ToArray() })
    }
    return $result.ToArray()
}

function Show-UnfinishedJobs {
    $runs = @()
    try { $runs = @(Get-UnfinishedAuditRuns) } catch { Write-Log "Yarım kalmış işler kontrol edilemedi: $($_.Exception.Message)" 'WARN'; return }
    if ($runs.Count -eq 0) { return }
    Write-Log "Önceki oturumlardan YARIM KALMIŞ $($runs.Count) işlem bulundu (denetim kaydı: $(Get-AuditFilePath)):" 'WARN'
    $summary = New-Object 'System.Collections.Generic.List[string]'
    foreach ($run in $runs) {
        $s = $run.Start
        if ($run.Aborted) { $how = 'pencere kapatılarak iptal edildi' } else { $how = 'program beklenmedik şekilde sonlandı' }
        $head = '{0} | {1} | {2} | {3} -> {4}' -f $s.Time, $s.User, $s.Label, $s.Server, $how
        Write-Log "  $head" 'WARN'
        $summary.Add("- $head")
        # Başlayıp sonucu yazılamamış adımlar = kesintinin tam olarak denk geldiği package / SUG
        $done = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($st in $run.Steps) {
            if ($st.Action -eq 'PkgRemoveResult') { [void]$done.Add('P:' + $st.PackageId); Write-Log ("      [{0}] {1}: {2}" -f $st.PackageId, 'silme sonucu', $st.Result) 'WARN' }
            if ($st.Action -eq 'SugChangeResult') { [void]$done.Add('S:' + $st.GroupId);   Write-Log ("      [{0}] {1}: {2}" -f $st.GroupName, 'SUG sonucu', $st.Result) 'WARN' }
        }
        foreach ($st in $run.Steps) {
            if ($st.Action -eq 'PkgRemoveStart' -and -not $done.Contains('P:' + $st.PackageId)) {
                $t = "      [$($st.PackageId)] içerik silme BAŞLADI, sonucu kaydedilmedi ($(@($st.ContentIds).Count) içerik) - package'ı kontrol edin."
                Write-Log $t 'ERROR'; $summary.Add($t.Trim())
            }
            if ($st.Action -eq 'SugChangeStart' -and -not $done.Contains('S:' + $st.GroupId)) {
                $t = "      [$($st.GroupName)] SUG değişikliği BAŞLADI, sonucu kaydedilmedi - SUG üyeliklerini kontrol edin."
                Write-Log $t 'ERROR'; $summary.Add($t.Trim())
            }
        }
        $notStarted = @($s.Packages).Count + @($s.Groups).Count - @($run.Steps | Where-Object { $_.Action -like '*Start' }).Count
        if ($notStarted -gt 0) { Write-Log "      $notStarted package/SUG adımına hiç başlanmamıştı." 'WARN' }
    }
    Write-Audit 'JobAcknowledged' @{ RunIds = [string[]]@($runs | ForEach-Object { [string]$_.Start.RunId }) }
    $txt = "Önceki oturumlardan yarım kalmış $($runs.Count) işlem bulundu:`n`n" + (($summary | Select-Object -First 12) -join "`n")
    if ($summary.Count -gt 12) { $txt += "`n..." }
    $txt += "`n`nİlgili package / SUG'ları 'Listele' ile kontrol edin. Ayrıntılar günlükte. (Bu uyarı bir daha gösterilmeyecek.)"
    Show-Message $txt 'Yarım kalmış işlemler' 'Warning'
}

# Arka plan işlerini periyodik okuyan zamanlayıcı (UI thread'inde çalışır).
$jobTimer = New-Object System.Windows.Forms.Timer
$jobTimer.Interval = 300
$jobTimer.Add_Tick({
    if ($script:InTick) { return }
    $script:InTick = $true
    try {
        Receive-JobLog
        foreach ($job in @($script:PackageJobs.Values)) {
            if ($null -eq $job.Handle) { continue }
            Receive-JobEvents $job
            $q = $job.Sync.Question
            if ($null -ne $q) {
                $job.Sync.Question = $null
                Receive-JobLog
                $ans = [System.Windows.Forms.MessageBox]::Show([string]$q.Text, ('İş #{0} - {1}' -f $job.Id, $q.Title), 'YesNo', 'Warning', 'Button2')
                $job.Sync.Answer = ($ans -eq 'Yes')
            }
            if ($job.Handle.IsCompleted) { Complete-BackgroundJob $job }
        }
        $lj = $script:ListJob
        if ($lj -and $lj.Handle -and $lj.Handle.IsCompleted) { Complete-ListJob $lj }
        Receive-JobLog
    } catch {
        try { Write-Log "Arka plan izleme hatası: $($_.Exception.Message)" 'ERROR' } catch { }
    } finally {
        $script:InTick = $false
        if ($script:PackageJobs.Count -eq 0 -and $null -eq $script:ListJob -and $script:LogQueue.IsEmpty) { $jobTimer.Stop() }
    }
})

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
    Update-EnvCombo
    Update-ControlStates
})
$btnEnv.Add_Click({ Invoke-EnvRead })

$txtFilter.Add_TextChanged({ Set-GroupFilter })

$gridUpd.Add_SelectionChanged({
    if (-not $script:Suspend) { Update-GroupChecks; Update-PackageCounts }
})

# VirtualMode: hücre değeri istendikçe önbellekten verilir.
$gridUpd.Add_CellValueNeeded({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.RowIndex -ge $script:ViewIds.Count) { return }
    $vals = Get-CachedRowValues $script:ViewIds[$e.RowIndex]
    if ($null -ne $vals) { $e.Value = $vals[$e.ColumnIndex] }
})
# Kilitli (işlemdeki) satırlar soluk gri çizilir.
$gridUpd.Add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $e.RowIndex -ge $script:ViewIds.Count) { return }
    $u = $script:Updates[$script:ViewIds[$e.RowIndex]]
    if ($u -and $u.IsLocked) {
        $e.CellStyle.ForeColor = [System.Drawing.Color]::Gray
        $e.CellStyle.BackColor = [System.Drawing.Color]::Gainsboro
        $e.CellStyle.SelectionBackColor = [System.Drawing.Color]::DarkGray
    }
})
# VirtualMode'da otomatik sıralama yoktur: başlığa tıklayınca bellekte sıralanır (tekrar tıklayınca ters).
$gridUpd.Add_ColumnHeaderMouseClick({
    param($sender, $e)
    if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
    if ($script:SortCol -eq $e.ColumnIndex) { $script:SortAsc = -not $script:SortAsc }
    else { $script:SortCol = $e.ColumnIndex; $script:SortAsc = $true }
    Update-UpdateViewOrder
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
    Set-PkgRowStyle $row
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
    Set-GroupRowStyle $row
})
$gridGrp.Add_DataError({ param($sender, $e) $e.ThrowException = $false })
$gridUpd.Add_DataError({ param($sender, $e) $e.ThrowException = $false })

# Arka planda iş varken pencere kapatılırsa kullanıcıya sor; onaylanırsa işlere iptal bildirilir.
$form.Add_FormClosing({
    param($sender, $e)
    if ($script:ListJob) {
        # Listeleme yalnızca okuma yapar; sormadan durdurulur.
        $script:ListJob.Sync.Cancel = $true
        try { [void]$script:ListJob.PS.BeginStop($null, $null) } catch { }
    }
    if ($script:PackageJobs.Count -eq 0) { return }
    $txt = "$($script:PackageJobs.Count) arka plan işlemi hâlâ sürüyor. Kapatırsanız işlemler yarıda kesilebilir " +
           "(ör. bir package'ın içeriği kısmen silinmiş kalabilir).`n`nYine de kapatılsın mı?"
    if ([System.Windows.Forms.MessageBox]::Show($txt, 'Kapat', 'YesNo', 'Warning', 'Button2') -ne 'Yes') { $e.Cancel = $true; return }
    $jobTimer.Stop()
    foreach ($job in @($script:PackageJobs.Values)) {
        $job.Sync.Cancel = $true
        if (-not $job.Simulate) { Write-Audit 'JobAborted' @{ RunId = $job.Sync.RunId; JobId = $job.Id; Kind = $job.Kind; Label = $job.Label } }
        try { [void]$job.PS.BeginStop($null, $null) } catch { }
    }
})

# Pencere açıldıktan sonra: önceki oturumlardan yarım kalmış işlem var mı? (audit kaydından)
$form.Add_Shown({ Show-UnfinishedJobs })

# Pencere boyutu değişince ana bölümleri yeniden yerleştir.
$form.Add_Resize({ Update-MainLayout })

# Sağ paneldeki sekme sayfaları gerçek boyutlarına pencere açıldıktan sonra kavuşur; o anda iç yerleşimi yenile.
$form.Add_Shown({ Update-RightPanelLayout })
$tabRight.Add_HandleCreated({ Update-RightPanelLayout })
$tabRight.Add_SelectedIndexChanged({ Update-RightPanelLayout })
$tabSug.Add_SizeChanged({ Update-RightPanelLayout })
$tabPkg.Add_SizeChanged({ Update-RightPanelLayout })

# ------------------------------------------------------------------
# Açılış
# ------------------------------------------------------------------
Write-Log 'Hazır. State kaydını seçip "Listele" düğmesine basın. Varsayılan olarak SİMÜLASYON açıktır.'
Import-StateFile (Find-StateFile)
$lblCount.Text = 'Henüz liste yüklenmedi.'
Update-ControlStates
Set-InitialWindowSize

[void]$form.ShowDialog()