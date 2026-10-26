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
    her iki işlem için de sadece ne yapılacağı günlüğe yazılır. Kutu her açılışta AÇIK başlar.

    GÜNLÜK, DENETİM KAYDI VE AYARLAR:
        %LOCALAPPDATA%\indirmeyapici\logs\yyyy-MM-dd.log : ekrandaki tüm günlük + AUDIT satırları
            (kullanıcı, makine, site, SUG, eklenen/çıkarılan/atlanan CI_ID'ler, yöntem, sonuç).
            90 günden eski dosyalar açılışta silinir.
        %LOCALAPPDATA%\indirmeyapici\settings.json : filtreler, diller, son paket, pencere boyutu.
    Ekran günlüğü seviyeye göre süzülebilir ve "Günlüğü Kaydet..." ile dışa aktarılabilir.

.NOTES
    Yalnızca Windows PowerShell 5.1. Listeleme ve SUG üyeliği için tüm SCCM erişimi SMS Provider WMI
    (root\SMS\site_XXX, DCOM/RPC) üzerinden yapılır; CIM/WinRM, uzak oturum ve ConfigurationManager
    modülü KULLANILMAZ. SUG düzenleme yetkisi gerekir. Tüm WMI bağlantıları PacketPrivacy (şifreli)
    kimlik doğrulama düzeyiyle ve zaman aşımlarıyla kurulur.
    YALNIZCA "İndirme" bölümü Configuration Manager Console gerektirir (Save-CMSoftwareUpdate'in WMI
    karşılığı yoktur); Console kurulu değilse listeleme ve üyelik işlemleri yine çalışır.
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
    (Updates dizisi) doğrudan yazdığı için bu kısıtlamaya takılmaz. Bu script de
    konsolun yaptığının birebir aynısını yapar: her SUG için SMS_AuthorizationList
    nesnesini tam (lazy Updates dahil) okur, ekleme ve çıkarmaları bellekte birleştirir,
    yeni Updates (uint32) dizisini TEK bir Put() ile yazar ve SUG'u yeniden okuyarak doğrular.
    ConfigurationManager modülüne (Add-/Remove-/Set-CMSoftwareUpdate*) yedek olarak DÜŞÜLMEZ.
    Böylece içerik indirilmemiş güncellemeler de -tıpkı konsolda olduğu gibi- deploy edilmiş
    bir SUG'a eklenebilir.
    Content indirme/dağıtım bu mekanizmanın kapsamı dışındadır; devam eden deployment'ın
    client'lara başarıyla ulaşması için içeriğin ayrıca (alttaki "İndirme" bölümünden,
    konsoldan veya ADR ile) bir Deployment Package'a indirilip DP'lere dağıtılması gerekir.

    ARKAPLAN (NON-BLOCKING) ÜYELİK DEĞİŞİMİ VE KİLİTLEME:
    "SUG Üyeliklerini Uygula" artık ekranı dondurmaz. Her uygulama ayrı bir PowerShell
    Runspace'inde (thread) çalışan bir "üyelik işi" başlatır ve hemen döner.
        - YALNIZCA işteki (üyeliği değişen) güncellemeler iş bitene kadar kilitlenir: satırlar gri
          görünür, "Üye olduğu SUG'lar" hücresinde [İşlemde] yazar ve yeni bir talebe giremezler.
        - SUG'lar ve Deployment Package'lar KİLİTLENMEZ: aynı SUG/paket üzerinde aynı anda birden
          fazla iş başlatılabilir. Aynı SUG'a yazımlar SUG başına isimli bir mutex
          ("Local\SccmSupersededToolSug_<CI_ID>", superseed_kontrol_ui.ps1 ile ortak) ile, aynı pakete
          indirmeler de paket başına bir mutex ile arkaplanda sıraya girer; UI etkilenmez.
        - İş bitince liste baştan YÜKLENMEZ: yalnızca ilgili güncellemelerin
          "Üye olduğu SUG'lar" hücresi ve ilgili SUG'ların Seçili/Toplam/kutu durumu tazelenir.
        - Yazmadan hemen önce eklenecek güncellemeler sunucuda kontrol edilir (Expired /
          Superseded / bulunamayanlar atlanır); yazdıktan sonra SUG yeniden okunup DOĞRULANIR.
        - Deploy edilmiş SUG'dan çıkarma ve içeriği indirilmemiş güncellemeyi deploy edilmiş
          SUG'a ekleme için onay penceresinde ayrı uyarı gösterilir (varsayılan düğme "Hayır").
        - "Son Üyelik İşlemini Geri Al" son işin gerçekten yaptığı değişiklikleri tersine çevirir
          (en fazla 20 adım; simülasyon kutusuna uyar).
    "Listele" ve "Paketleri Yenile" de arkaplanda çalışır. Üyelik ve indirme işleri ayrı runspace
    havuzlarında, her biri için en fazla 10 eşzamanlı iş olacak şekilde çalışır; fazlası sıraya girer.
    İndirmede bir güncellemenin hatası işi durdurmaz; iş "Kısmen Tamamlandı" olarak bitebilir.
#>

#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Management
[System.Windows.Forms.Application]::EnableVisualStyles()


# ------------------------------------------------------------------
# Genel durum
# ------------------------------------------------------------------
$script:Groups       = @()      # ortamdaki tüm SUG'lar (Id, Name, Total, IsDeployed)
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
$script:MembershipJobs = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
                                 # arkaplanda çalışan/çalışmış SUG üyelik işleri (her biri senkronize bir hashtable)
$script:ListJob        = $null  # o an çalışan arkaplan listeleme işi (tek)
$script:Pools          = @{}    # 'Membership' / 'Download' / 'List' -> RunspacePool
$script:MaxParallel    = 10     # üyelik ve indirme için ayrı ayrı en fazla eşzamanlı iş
$script:LockedUpdates  = New-Object 'System.Collections.Generic.HashSet[string]'   # işlemdeki güncelleme CI_ID'leri
$script:UpdRowById     = @{}    # güncelleme CI_ID -> gridUpd satırı (hedefli yenileme için)
$script:GrpRowById     = @{}    # SUG CI_ID -> gridGrp satırı (hedefli yenileme için)
$script:UndoStack      = New-Object System.Collections.ArrayList   # geri alınabilir üyelik işleri (en yeni en sonda)
$script:MemTickBusy    = $false # üyelik zamanlayıcısı yeniden girişim (re-entrancy) koruması
$script:ListBusy       = $false # listeleme sürerken üyelik sonuçlarının uygulanmasını erteler
$script:LastMembershipFail = 0  # son biten üyelik işindeki hata sayısı (sekme başlığı için)
$script:LockBackColor  = [System.Drawing.Color]::Gainsboro
$script:LockForeColor  = [System.Drawing.Color]::DimGray

# Kalıcı dosyalar: %LOCALAPPDATA%\indirmeyapici\ (logs\yyyy-MM-dd.log, settings.json)
$script:AppDir = if ($env:LOCALAPPDATA) { Join-Path $env:LOCALAPPDATA 'indirmeyapici' } else { Join-Path $env:TEMP 'indirmeyapici' }
$script:LogDir           = Join-Path $script:AppDir 'logs'
$script:SettingsPath     = Join-Path $script:AppDir 'settings.json'
$script:LogRetentionDays = 90
$script:LogEntries       = New-Object 'System.Collections.Generic.List[object]'   # ekran günlüğü (filtre için)
$script:LogReady         = $false
$script:FileLogOk        = $true
$script:FileLogError     = ''
$script:SavedClientSize  = $null
$script:Utf8Bom          = New-Object System.Text.UTF8Encoding($true)
$script:Who              = '{0}\{1}@{2}' -f $env:USERDOMAIN, $env:USERNAME, $env:COMPUTERNAME

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
# Fare tekerleği ile form kaydırma
# Fare, kendi kaydırma çubuğu olan bir nesnenin (grid, günlük kutusu, açık liste) üzerinde değilse tekerlek formu kaydırır.
# ------------------------------------------------------------------
if (-not ('FormWheelScrollFilter' -as [type])) {
    Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public sealed class FormWheelScrollFilter : IMessageFilter
{
    [DllImport("user32.dll")]
    private static extern IntPtr WindowFromPoint(Point p);

    private readonly ScrollableControl _host;
    public FormWheelScrollFilter(ScrollableControl host) { _host = host; }

    public bool PreFilterMessage(ref Message m)
    {
        if (m.Msg != 0x020A) return false;                      // WM_MOUSEWHEEL
        if (!_host.IsHandleCreated || !_host.VerticalScroll.Visible) return false;

        Control c = Control.FromHandle(WindowFromPoint(Cursor.Position));
        if (c == null) return false;                            // örn. açık combobox listesi: dokunma
        bool inside = false;
        for (Control p = c; p != null; p = p.Parent) { if (p == _host) { inside = true; break; } }
        if (!inside) return false;
        if (Form.ActiveForm != _host) return false;

        DataGridView g = c as DataGridView;
        if (g != null)
        {
            foreach (Control k in g.Controls)
            {
                VScrollBar v = k as VScrollBar;
                if (v != null && v.Visible) return false;       // grid kendi kaydırmasını kullansın
            }
        }
        TextBoxBase tb = c as TextBoxBase;
        if (tb != null && tb.Multiline && tb.Lines.Length * tb.Font.Height > tb.ClientSize.Height) return false;
        ComboBox cb = c as ComboBox;
        if (cb != null && cb.DroppedDown) return false;

        int delta = unchecked((short)((m.WParam.ToInt64() >> 16) & 0xFFFF));
        int dy = (int)(delta / 120.0 * SystemInformation.MouseWheelScrollLines * 16);
        Point pos = _host.AutoScrollPosition;                   // değerler negatif gelir
        _host.AutoScrollPosition = new Point(-pos.X, -pos.Y - dy);
        return true;
    }
}
"@
}


# ------------------------------------------------------------------
# Form
# ------------------------------------------------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'SCCM - Required Güncellemeler & SUG Üyelik Yönetimi'
# Yerleşim 1280x1314 olarak tasarlandı. Formun gösterilebilir ASGARİ yerleşim alanı 1100x1144'tür:
# pencere (ekran) bundan küçükse nesneler sıkışmaz, form kaydırma çubukları çıkar (fare tekerleği ile de kaydırılır).
# Dikeyde YALNIZCA üstteki "Required güncellemeler" ve "Software Update Grupları" kutuları esner
# (en az $script:UpdGroupMinH yüksekliğe kadar). Altındaki tüm bölümler (İndirme, Arkaplan İşleri,
# İşlemler, Günlük) alta sabitlenmiştir ve yükseklikleri DEĞİŞMEZ; böylece Arkaplan İşleri tablosu
# küçük ekranlarda ezilmez.
$script:LayoutMinW = 1100
$script:LayoutMinH = 1144
$script:DesignW    = 1280
$script:DesignH    = 1314
$script:UpdGroupMinH = 330      # üst kutuların (güncellemeler / SUG'lar) asgari yüksekliği
$script:JobsPanelH   = 260      # Arkaplan İşleri paneli (sabit) yüksekliği
$form.ClientSize = New-Object System.Drawing.Size($script:DesignW, $script:DesignH)   # kontroller bu tasarım boyutuna göre yerleştirilir
$form.AutoScroll = $true
$form.AutoScrollMinSize = New-Object System.Drawing.Size($script:LayoutMinW, $script:LayoutMinH)
$form.MinimumSize = New-Object System.Drawing.Size(640, 420)         # pencere küçültülebilir; sığmayan kısım kaydırılır
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
$gUpd.MinimumSize = New-Object System.Drawing.Size(0, $script:UpdGroupMinH)
$gridUpd = New-UiGrid $gUpd 10 22 780 430
$gridUpd.MultiSelect = $true
$gridUpd.ReadOnly = $true
$gridUpd.AutoSizeColumnsMode = 'None'          # Fill kapalı: sütunlar sabit genişlikte, toplam sığmazsa X scroll çıkar
$gridUpd.ScrollBars = 'Both'
$gridUpd.AllowUserToResizeColumns = $true
# Sütun sırası Show-Updates içindeki Rows.Add değer sırasıyla AYNI olmalıdır.
Add-UiTextColumn $gridUpd 'uKB' 'KB' 80
Add-UiTextColumn $gridUpd 'uTitle' 'Başlık' 340
Add-UiTextColumn $gridUpd 'uReq' 'Required' 70
Add-UiTextColumn $gridUpd 'uContent' 'İçerik' 55
Add-UiTextColumn $gridUpd 'uPosted' 'Yayın' 85
Add-UiTextColumn $gridUpd 'uCat' 'Kategori' 150
Add-UiTextColumn $gridUpd 'uGroups' 'Üye olduğu SUG''lar' 260
Add-UiTextColumn $gridUpd 'uId' 'CI_ID' 75
$gridUpd.Columns['uContent'].ToolTipText = 'Var: içerik bir Deployment Package''a indirilmiş (IsContentProvisioned). Yok: henüz indirilmemiş.'
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
$gGrp.MinimumSize = New-Object System.Drawing.Size(0, $script:UpdGroupMinH)
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
# Sütun sırası Show-Groups içindeki Rows.Add değer sırasıyla AYNI olmalıdır.
Add-UiTextColumn $gridGrp 'gName' 'Grup adı' 130 $true 100
Add-UiTextColumn $gridGrp 'gDep' 'Deploy' 52
Add-UiTextColumn $gridGrp 'gSel' 'Seçili' 55
Add-UiTextColumn $gridGrp 'gTotal' 'Toplam' 58
$gridGrp.Columns['gDep'].ToolTipText = 'Evet: SUG en az bir koleksiyona deploy edilmiş (IsDeployed). Bu SUG''lardaki değişiklikler istemcileri doğrudan etkiler.'
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

$lblDlNote = New-UiLabel $gDl 'İndirme konumu: İnternet (Microsoft Update) - sabit. Diller, seçilen tüm güncellemelere birlikte uygulanır (SCCM cmdlet sınırlaması). İndirme arkaplanda çalışır; bu düğmeye farklı paket/seçimle tekrar basarak aynı anda birden fazla indirme başlatabilirsiniz - aşağıdaki "Arkaplan İşleri > İndirmeler" sekmesinden takip edin.' 660 56 590 46
$lblDlNote.ForeColor = [System.Drawing.Color]::DimGray
$lblDlNote.Anchor = 'Top,Right'

$btnDownload = New-UiButton $gDl 'Seçili Güncellemeleri İndir (Arkaplanda)' 950 106 300 30
$btnDownload.Anchor = 'Top,Right'

#region --- Arkaplan İşleri (sekmeli: İndirmeler / SUG Üyelik İşleri) ---
$gJobs = New-UiGroup $form 'Arkaplan İşleri (her tür için en fazla 10 iş aynı anda çalışır; fazlası "Bekliyor" olarak sıraya girer)' 10 792 1260 $script:JobsPanelH
$gJobs.Anchor = 'Bottom,Left,Right'     # alta sabit, yüksekliği değişmez (pencere küçülünce ezilmez)
$gJobs.MinimumSize = New-Object System.Drawing.Size(0, $script:JobsPanelH)
$tabJobs = New-Object System.Windows.Forms.TabControl
$tabJobs.Location = New-Object System.Drawing.Point(10, 20)
$tabJobs.Size = New-Object System.Drawing.Size(1240, 202)
$tabJobs.MinimumSize = New-Object System.Drawing.Size(0, 202)   # sekme başlığı + ~7 satırlık tablo
$tabJobs.Anchor = 'Top,Bottom,Left,Right'
$gJobs.Controls.Add($tabJobs)
$tpDl = New-Object System.Windows.Forms.TabPage
$tpDl.Text = 'İndirmeler'
$tpMem = New-Object System.Windows.Forms.TabPage
$tpMem.Text = 'SUG Üyelik İşleri'
[void]$tabJobs.TabPages.Add($tpDl)
[void]$tabJobs.TabPages.Add($tpMem)

$gridDl = New-UiGrid $tpDl 0 0 1230 110
$gridDl.Dock = 'Fill'
$gridDl.MultiSelect = $false
$gridDl.ReadOnly = $false
$gridDl.EditMode = 'EditProgrammatically'
Add-UiTextColumn $gridDl 'dPkg' 'Deployment Package' 200 $true 30
Add-UiTextColumn $gridDl 'dProg' 'İlerleme (işlenen/toplam)' 140
Add-UiTextColumn $gridDl 'dRemain' 'Kalan' 55
Add-UiTextColumn $gridDl 'dFail' 'Hatalı' 55
Add-UiTextColumn $gridDl 'dPct' 'Yüzde' 55
Add-UiTextColumn $gridDl 'dCurrent' 'Şu an indirilen' 220 $true 40
Add-UiTextColumn $gridDl 'dStatus' 'Durum' 130
Add-UiTextColumn $gridDl 'dSure' 'Süre' 60
$colDlCancel = New-Object System.Windows.Forms.DataGridViewButtonColumn
$colDlCancel.Name = 'dCancel'
$colDlCancel.HeaderText = ''
$colDlCancel.Width = 70
$colDlCancel.UseColumnTextForButtonValue = $false
[void]$gridDl.Columns.Add($colDlCancel)

$gridMem = New-UiGrid $tpMem 0 0 1230 110
$gridMem.Dock = 'Fill'
$gridMem.MultiSelect = $false
$gridMem.ReadOnly = $false
$gridMem.EditMode = 'EditProgrammatically'
Add-UiTextColumn $gridMem 'mStart' 'Başlangıç' 70
Add-UiTextColumn $gridMem 'mKind' 'Tür' 110
Add-UiTextColumn $gridMem 'mGroups' 'SUG''lar' 260 $true 40
Add-UiTextColumn $gridMem 'mProg' 'İlerleme (SUG)' 100
Add-UiTextColumn $gridMem 'mUpd' 'Kilitli gnc.' 75
Add-UiTextColumn $gridMem 'mCurrent' 'Şu an işlenen' 200 $true 30
Add-UiTextColumn $gridMem 'mStatus' 'Durum' 130
Add-UiTextColumn $gridMem 'mSure' 'Süre' 60
$colMemCancel = New-Object System.Windows.Forms.DataGridViewButtonColumn
$colMemCancel.Name = 'mCancel'
$colMemCancel.HeaderText = ''
$colMemCancel.Width = 70
$colMemCancel.UseColumnTextForButtonValue = $false
[void]$gridMem.Columns.Add($colMemCancel)

$btnJobsClear = New-UiButton $gJobs 'Tamamlananları Listeden Temizle' 10 228 240 26
$btnJobsClear.Anchor = 'Bottom,Left'
$lblJobsNote = New-UiLabel $gJobs 'Her iş kendi arkaplan thread''inde (runspace havuzu) çalışır. "İptal": indirmede devam eden öğe, üyelikte yazılmakta olan SUG bittikten sonra işi durdurur (sert kesme yapılmaz).' 260 232 990 22
$lblJobsNote.Anchor = 'Bottom,Left,Right'
$lblJobsNote.ForeColor = [System.Drawing.Color]::DimGray
#endregion

# Zamanlayıcılar (System.Windows.Forms.Timer bir Control değildir; $form.Controls.Add ile eklenmez,
# $script: kapsamındaki değişken olarak yaşar, Start()/Stop() yeterlidir). Hepsi yalnızca iş varken çalışır.
$tmrProgress = New-Object System.Windows.Forms.Timer      # indirme işleri
$tmrProgress.Interval = 600
$tmrProgress.Add_Tick({ Update-DownloadGrid })
$tmrMembership = New-Object System.Windows.Forms.Timer    # SUG üyelik işleri
$tmrMembership.Interval = 500
$tmrMembership.Add_Tick({ Update-MembershipJobs })
$tmrList = New-Object System.Windows.Forms.Timer          # arkaplan listeleme / paket yenileme
$tmrList.Interval = 300
$tmrList.Add_Tick({ Update-ListJob })

# --- İşlemler ---
$gAct = New-UiGroup $form 'İşlemler' 10 1058 1260 62
$gAct.Anchor = 'Bottom,Left,Right'
$chkSim = New-Object System.Windows.Forms.CheckBox
$chkSim.Text = 'Simülasyon (SUG üyeliği / indirme yapma, sadece günlüğe yaz)'
$chkSim.Location = New-Object System.Drawing.Point(14, 26)
$chkSim.Size = New-Object System.Drawing.Size(400, 22)
$chkSim.Checked = $true
$gAct.Controls.Add($chkSim)
$btnApply = New-UiButton $gAct 'SUG Üyeliklerini Uygula (Arkaplanda)' 430 20 260 32
$btnUndo = New-UiButton $gAct 'Son Üyelik İşlemini Geri Al' 700 20 260 32
$btnUndo.Enabled = $false
$btnClose = New-UiButton $gAct 'Kapat' 1140 20 105 32
$btnClose.Anchor = 'Top,Right'

# --- Günlük ---
$lblLog = New-UiLabel $form 'İşlem Günlüğü:' 14 1130 200
$lblLog.Anchor = 'Bottom,Left'
$lblLogLevel = New-UiLabel $form 'Göster:' 880 1130 50
$lblLogLevel.Anchor = 'Bottom,Right'
$cmbLogLevel = New-Object System.Windows.Forms.ComboBox
$cmbLogLevel.DropDownStyle = 'DropDownList'
$cmbLogLevel.Location = New-Object System.Drawing.Point(930, 1126)
$cmbLogLevel.Size = New-Object System.Drawing.Size(180, 23)
$cmbLogLevel.Anchor = 'Bottom,Right'
[void]$cmbLogLevel.Items.AddRange([object[]]@('Tümü', 'Yalnızca uyarı ve hatalar', 'Yalnızca hatalar'))
$cmbLogLevel.SelectedIndex = 0
$form.Controls.Add($cmbLogLevel)
$btnLogSave = New-UiButton $form 'Günlüğü Kaydet...' 1120 1125 150 25
$btnLogSave.Anchor = 'Bottom,Right'
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(10, 1154)
$txtLog.Size = New-Object System.Drawing.Size(1260, 142)
$txtLog.Anchor = 'Bottom,Left,Right'
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtLog)

# ------------------------------------------------------------------
# Günlük (ekran + kalıcı dosya) ve denetim (audit) kaydı
# ------------------------------------------------------------------
# Ekran günlüğü seviye filtresine göre gösterilir; TÜM satırlar ayrıca gün bazlı dosyaya yazılır:
#   %LOCALAPPDATA%\indirmeyapici\logs\yyyy-MM-dd.log
# AUDIT satırları yalnızca dosyaya yazılır (CI_ID listeleri ekranı boğmasın).
# Dosya satırı: "yyyy-MM-dd HH:mm:ss [SEVİYE] [ALAN\kullanıcı@MAKİNE] mesaj"
function Initialize-AppDirs {
    try {
        if (-not (Test-Path -LiteralPath $script:LogDir)) { New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null }
        $limit = (Get-Date).AddDays(-$script:LogRetentionDays)
        Get-ChildItem -LiteralPath $script:LogDir -Filter '*.log' -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $limit } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    } catch {
        $script:FileLogOk = $false
        $script:FileLogError = $_.Exception.Message
    }
}

function Write-LogFile {
    param([string]$Level, [string]$Message, [datetime]$Time = (Get-Date))
    if (-not $script:FileLogOk) { return }
    try {
        $path = Join-Path $script:LogDir ($Time.ToString('yyyy-MM-dd') + '.log')
        $line = '{0} [{1}] [{2}] {3}' -f $Time.ToString('yyyy-MM-dd HH:mm:ss'), $Level, $script:Who, $Message
        [System.IO.File]::AppendAllText($path, $line + "`r`n", $script:Utf8Bom)
    } catch {
        # Dosyaya yazılamıyorsa bir kez ekrana uyarı verilir, uygulama çalışmaya devam eder
        $script:FileLogOk = $false
        $script:FileLogError = $_.Exception.Message
        $warn = '[{0}] [WARN] Günlük dosyasına yazılamıyor, bu oturumda kalıcı kayıt tutulmayacak: {1}' -f (Get-Date -Format 'HH:mm:ss'), $script:FileLogError
        $script:LogEntries.Add(@{ Level = 'WARN'; Line = $warn })
        $txtLog.AppendText($warn + [Environment]::NewLine)
    }
}

function Write-Audit {
    param([string]$Message)
    Write-LogFile -Level 'AUDIT' -Message $Message
}

function Test-LogLevelVisible {
    param([string]$Level)
    switch ($cmbLogLevel.SelectedIndex) {
        1       { return ($Level -eq 'WARN' -or $Level -eq 'ERROR') }
        2       { return ($Level -eq 'ERROR') }
        default { return $true }
    }
}

# Not: Uzun işlemler artık arkaplanda yapıldığı için burada DoEvents() ÇAĞRILMAZ
# (yeniden girişim / re-entrancy riskini önler).
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO', $Time = $null)
    if (-not $Time) { $Time = Get-Date }
    $t = [datetime]$Time
    $line = '[{0}] [{1}] {2}' -f $t.ToString('HH:mm:ss'), $Level, $Message
    $script:LogEntries.Add(@{ Level = $Level; Line = $line })
    if ($script:LogEntries.Count -gt 20000) { $script:LogEntries.RemoveRange(0, 5000) }   # bellek sınırı
    if (Test-LogLevelVisible $Level) { $txtLog.AppendText($line + [Environment]::NewLine) }
    Write-LogFile -Level $Level -Message $Message -Time $t
}

# Seviye filtresi değişince ekrandaki günlüğü bellekteki kayıtlardan yeniden kurar.
function Update-LogView {
    $sb = New-Object System.Text.StringBuilder
    foreach ($e in $script:LogEntries) {
        if (Test-LogLevelVisible $e.Level) { [void]$sb.AppendLine($e.Line) }
    }
    $txtLog.Text = $sb.ToString()
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.ScrollToCaret()
}

function Save-LogView {
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'Metin dosyası (*.txt)|*.txt|Tüm dosyalar (*.*)|*.*'
    $dlg.FileName = 'indirmeyapici_gunluk_{0}.txt' -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    try {
        if ($dlg.ShowDialog($form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
        [System.IO.File]::WriteAllText($dlg.FileName, $txtLog.Text, $script:Utf8Bom)
        Write-Log "Günlük kaydedildi: $($dlg.FileName)" 'OK'
    } catch {
        Write-Log "Günlük kaydedilemedi: $($_.Exception.Message)" 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    } finally {
        $dlg.Dispose()
    }
}

function Set-Busy {
    param([bool]$On)
    foreach ($c in @($btnList, $btnApply, $btnPkgRefresh)) { $c.Enabled = (-not $On) }
    if ($On) { $form.Cursor = [System.Windows.Forms.Cursors]::AppStarting }    # arkaplan işi: form kullanılabilir kalır
    else     { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    Update-UndoButton
}

function Show-Message {
    param([string]$Text, [string]$Title = 'Bilgi', [string]$Icon = 'Information')
    [void][System.Windows.Forms.MessageBox]::Show($Text, $Title, 'OK', $Icon)
}

# ------------------------------------------------------------------
# Girdi doğrulama
# ------------------------------------------------------------------
function Test-SiteCode {
    param([string]$Value)
    return ($Value -match '^[A-Za-z0-9]{3}$')
}

# SMS Provider: NetBIOS adı, FQDN veya IPv4 kabul edilir; boşluk, ters bölü, tırnak vb. içeremez.
function Test-ServerName {
    param([string]$Value)
    if (-not $Value -or $Value.Length -gt 253) { return $false }
    if ($Value -match '^\d{1,3}(\.\d{1,3}){3}$') {
        $ip = $null
        return [System.Net.IPAddress]::TryParse($Value, [ref]$ip)
    }
    return ($Value -match '^[A-Za-z0-9]([A-Za-z0-9\-]{0,61}[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9\-]{0,61}[A-Za-z0-9])?)*$')
}

# ------------------------------------------------------------------
# Bağlantı bilgisi ve kullanıcı ayarları
# ------------------------------------------------------------------
# state.json varsa yalnızca Site kodu ve SMS Provider alanlarını doldurur (SADECE okunur, doğrulanarak).
function Import-ConnectionDefaults {
    $path = Find-StateFile
    if (-not $path) { return }
    try {
        $state = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
        $raw = $state.LastRun
        if (-not $raw -and $state.History) { $raw = @($state.History)[-1] }
        if ($raw) {
            $filled = $false
            if ($raw.SiteCode) {
                if (Test-SiteCode ([string]$raw.SiteCode)) { $txtSite.Text = ([string]$raw.SiteCode).ToUpper(); $filled = $true }
                else { Write-Log "state.json'daki site kodu geçersiz, yoksayıldı: '$($raw.SiteCode)'" 'WARN' }
            }
            if ($raw.ProviderServer) {
                if (Test-ServerName ([string]$raw.ProviderServer)) { $txtServer.Text = [string]$raw.ProviderServer; $filled = $true }
                else { Write-Log "state.json'daki SMS Provider adı geçersiz, yoksayıldı: '$($raw.ProviderServer)'" 'WARN' }
            }
            if ($filled) { Write-Log "Bağlantı bilgileri state.json'dan dolduruldu: $path" }
        }
    } catch {
        Write-Log "state.json okunamadı (yoksayıldı): $($_.Exception.Message)" 'WARN'
    }
}

# Kullanıcı ayarları: %LOCALAPPDATA%\indirmeyapici\settings.json
# Simülasyon kutusu BİLEREK saklanmaz: her açılışta AÇIK başlar.
# Site/Provider yalnızca state.json'dan gelmediyse (alan boşsa) buradan doldurulur.
function Import-Settings {
    if (-not (Test-Path -LiteralPath $script:SettingsPath)) { return }
    try {
        $s = Get-Content -LiteralPath $script:SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Log "Ayar dosyası okunamadı (varsayılanlar kullanılıyor): $($_.Exception.Message)" 'WARN'
        return
    }
    try {
        if (-not $txtSite.Text   -and $s.SiteCode       -and (Test-SiteCode   ([string]$s.SiteCode)))       { $txtSite.Text   = ([string]$s.SiteCode).ToUpper() }
        if (-not $txtServer.Text -and $s.ProviderServer -and (Test-ServerName ([string]$s.ProviderServer))) { $txtServer.Text = [string]$s.ProviderServer }
        if ($null -ne $s.Vendor)         { $txtVendor.Text = [string]$s.Vendor }
        if ($null -ne $s.Search)         { $txtTitle.Text  = [string]$s.Search }
        if ($null -ne $s.MinRequired) {
            $v = [decimal]$s.MinRequired
            if ($v -ge $nudReq.Minimum -and $v -le $nudReq.Maximum) { $nudReq.Value = $v }
        }
        if ($null -ne $s.ExcludeExpired) { $chkExpired.Checked   = [bool]$s.ExcludeExpired }
        if ($null -ne $s.NoSugOnly)      { $chkNoSug.Checked     = [bool]$s.NoSugOnly }
        if ($null -ne $s.LangWinEn)      { $chkLangWinEn.Checked = [bool]$s.LangWinEn }
        if ($null -ne $s.LangWinTr)      { $chkLangWinTr.Checked = [bool]$s.LangWinTr }
        if ($null -ne $s.LangOffEn)      { $chkLangOffEn.Checked = [bool]$s.LangOffEn }
        if ($null -ne $s.LangOffTr)      { $chkLangOffTr.Checked = [bool]$s.LangOffTr }
        if ($s.LastPackage)              { $cmbPkg.Text = [string]$s.LastPackage }
        if ($null -ne $s.LogLevel -and [int]$s.LogLevel -ge 0 -and [int]$s.LogLevel -lt $cmbLogLevel.Items.Count) { $cmbLogLevel.SelectedIndex = [int]$s.LogLevel }
        if ($s.WindowWidth -and $s.WindowHeight) {
            $script:SavedClientSize = New-Object System.Drawing.Size([int]$s.WindowWidth, [int]$s.WindowHeight)
        }
        Write-Log "Kullanıcı ayarları yüklendi: $($script:SettingsPath)"
    } catch {
        Write-Log "Ayarların bir kısmı uygulanamadı: $($_.Exception.Message)" 'WARN'
    }
}

function Save-Settings {
    try {
        if (-not (Test-Path -LiteralPath $script:AppDir)) { New-Item -ItemType Directory -Path $script:AppDir -Force | Out-Null }
        $o = [ordered]@{
            SiteCode       = $txtSite.Text.Trim().ToUpper()
            ProviderServer = $txtServer.Text.Trim()
            Vendor         = $txtVendor.Text.Trim()
            Search         = $txtTitle.Text.Trim()
            MinRequired    = [int]$nudReq.Value
            ExcludeExpired = $chkExpired.Checked
            NoSugOnly      = $chkNoSug.Checked
            LangWinEn      = $chkLangWinEn.Checked
            LangWinTr      = $chkLangWinTr.Checked
            LangOffEn      = $chkLangOffEn.Checked
            LangOffTr      = $chkLangOffTr.Checked
            LastPackage    = $cmbPkg.Text.Trim()
            LogLevel       = $cmbLogLevel.SelectedIndex
        }
        if ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal) {
            $o.WindowWidth  = $form.ClientSize.Width
            $o.WindowHeight = $form.ClientSize.Height
        } elseif ($script:SavedClientSize) {
            $o.WindowWidth  = $script:SavedClientSize.Width
            $o.WindowHeight = $script:SavedClientSize.Height
        }
        $json = $o | ConvertTo-Json
        [System.IO.File]::WriteAllText($script:SettingsPath, $json, $script:Utf8Bom)
    } catch {
        Write-LogFile -Level 'WARN' -Message "Ayarlar kaydedilemedi: $($_.Exception.Message)"
    }
}

# ------------------------------------------------------------------
# SCCM erişimi (yalnızca doğrulama ve modül yolu; WMI işleri arkaplan thread'lerinde yapılır)
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

function Get-Connection {
    $site = $txtSite.Text.Trim().ToUpper()
    $server = $txtServer.Text.Trim()
    if (-not (Test-SiteCode $site)) { throw 'Site kodu 3 karakter (harf/rakam) olmalı.' }
    if (-not $server) { throw 'SMS Provider sunucusu boş olamaz.' }
    if (-not (Test-ServerName $server)) { throw "SMS Provider adı geçersiz: '$server'. NetBIOS adı, FQDN veya IPv4 adresi girin." }
    return [pscustomobject]@{ Site = $site; Server = $server }
}

# Üyelik ve indirme işleri için runspace havuzları (her biri en fazla $script:MaxParallel eşzamanlı iş).
# Havuz dolunca yeni işler "Bekliyor" durumunda sıraya girer.
function Get-JobPool {
    param([ValidateSet('Membership', 'Download', 'List')][string]$Name)
    if (-not $script:Pools.ContainsKey($Name)) {
        $max = if ($Name -eq 'List') { 2 } else { $script:MaxParallel }
        $p = [runspacefactory]::CreateRunspacePool(1, $max)
        $p.ThreadOptions = 'ReuseThread'
        $p.Open()
        $script:Pools[$Name] = $p
    }
    return $script:Pools[$Name]
}

# Tek bir arkaplan işini havuzda başlatır. $Arguments sırası worker'ın param() sırasıyla aynı olmalıdır.
function Start-PooledScript {
    param([string]$PoolName, [scriptblock]$Script, [object[]]$Arguments, $Job)
    $ps = [powershell]::Create()
    $ps.RunspacePool = Get-JobPool $PoolName
    [void]$ps.AddScript($Script, $true)      # useLocalScope: havuzdaki runspace'te değişken artığı kalmasın
    foreach ($a in $Arguments) { [void]$ps.AddArgument($a) }
    [void]$ps.AddArgument($Job)
    $Job.PS     = $ps
    $Job.Handle = $ps.BeginInvoke()
}

# Biten işin PowerShell nesnesini sonlandırıp serbest bırakır (havuzdaki runspace yeniden kullanılır).
function Close-PooledScript {
    param($Job)
    try {
        if ($Job.PS) { [void]$Job.PS.EndInvoke($Job.Handle) }
    } catch {
        if (-not $Job.Error) { $Job.Error = $_.Exception.Message }
        if ($Job.Status -in @('Bekliyor', 'Çalışıyor')) { $Job.Status = 'Hata' }
    } finally {
        try { if ($Job.PS) { $Job.PS.Dispose() } } catch { }
        $Job.PS = $null
        $Job.Cleaned = $true
    }
}

# İş satırı yardımcıları: grid'de işi Tag (= iş Id) ile bulur, yoksa ekler (yerinde güncelleme; her tick'te
# tablo silinip yeniden kurulmaz, böylece İptal düğmesine tıklama ve seçim kaybolmaz).
function Get-JobRow {
    param($Grid, [string]$JobId)
    foreach ($r in $Grid.Rows) { if ([string]$r.Tag -eq $JobId) { return $r } }
    $i = $Grid.Rows.Add()
    $row = $Grid.Rows[$i]
    $row.Tag = $JobId
    return $row
}

function Remove-StaleJobRows {
    param($Grid, [string[]]$KeepIds)
    for ($i = $Grid.Rows.Count - 1; $i -ge 0; $i--) {
        if ($KeepIds -notcontains [string]$Grid.Rows[$i].Tag) { $Grid.Rows.RemoveAt($i) }
    }
}

function Format-Elapsed {
    param($Job)
    $endRef = if ($Job.EndTime) { $Job.EndTime } else { Get-Date }
    return ('{0:mm\:ss}' -f ($endRef - $Job.StartTime))
}

function Update-JobTabTitles {
    $dlActive  = @($script:DownloadJobs.ToArray()   | Where-Object { $_.Status -in @('Bekliyor', 'Çalışıyor') }).Count
    $memActive = @($script:MembershipJobs.ToArray() | Where-Object { -not $_.Cleaned }).Count
    $tpDl.Text = if ($dlActive -gt 0) { "İndirmeler ($dlActive aktif)" } else { 'İndirmeler' }
    if ($memActive -gt 0) {
        $tpMem.Text = "SUG Üyelik İşleri ($memActive aktif · kilitli: $($script:LockedUpdates.Count) güncelleme)"
    } elseif ($script:LastMembershipFail -gt 0) {
        $tpMem.Text = "SUG Üyelik İşleri (son işte $($script:LastMembershipFail) hata!)"
    } else {
        $tpMem.Text = 'SUG Üyelik İşleri'
    }
}

# ------------------------------------------------------------------
# WMI yardımcıları (arkaplan worker'larında dot-source edilir)
# ------------------------------------------------------------------
# System.Management doğrudan kullanılır; böylece:
#   - Bağlantı, sorgu, okuma ve yazma için ZAMAN AŞIMI verilebilir (takılan iş kilitleri sonsuza dek tutmaz),
#   - Bağlantı PacketPrivacy (şifreli + imzalı) kimlik doğrulama düzeyiyle kurulur,
#   - Bir iş boyunca tek bağlantı (ManagementScope) yeniden kullanılır.
# Metin olarak tutulur ve her worker'a argüman olarak verilir (runspace'ler ana oturumun fonksiyonlarını görmez).
$script:WmiHelperSource = @'
function New-SmsScope {
    param([string]$Server, [string]$Site, [int]$TimeoutSec = 30)
    $co = New-Object System.Management.ConnectionOptions
    $co.Timeout          = [TimeSpan]::FromSeconds($TimeoutSec)
    $co.Impersonation    = [System.Management.ImpersonationLevel]::Impersonate
    $co.Authentication   = [System.Management.AuthenticationLevel]::PacketPrivacy
    $co.EnablePrivileges = $true
    $scope = New-Object System.Management.ManagementScope("\\$Server\root\SMS\site_$Site", $co)
    $scope.Connect()
    return $scope
}

function Invoke-SmsQuery {
    param($Scope, [string]$Query, [int]$TimeoutSec = 300)
    $opt = New-Object System.Management.EnumerationOptions
    $opt.Timeout           = [TimeSpan]::FromSeconds($TimeoutSec)
    $opt.ReturnImmediately = $true
    $opt.Rewindable        = $false
    $searcher = New-Object System.Management.ManagementObjectSearcher($Scope, (New-Object System.Management.ObjectQuery($Query)), $opt)
    $list = New-Object System.Collections.ArrayList
    try {
        foreach ($o in $searcher.Get()) { [void]$list.Add($o) }
    } finally {
        $searcher.Dispose()
    }
    return $list.ToArray()     # çağıranlar @(...) ile sarar
}

# Nesneyi yol (path) ile TAM okur; "lazy" özellikler (örn. SMS_AuthorizationList.Updates) de dolu gelir.
function Get-SmsObject {
    param($Scope, [string]$Path, [int]$TimeoutSec = 60)
    $go = New-Object System.Management.ObjectGetOptions
    $go.Timeout = [TimeSpan]::FromSeconds($TimeoutSec)
    $o = New-Object System.Management.ManagementObject($Scope, (New-Object System.Management.ManagementPath($Path)), $go)
    $o.Get()
    return $o
}

function Save-SmsObject {
    param($Object, [int]$TimeoutSec = 120)
    $po = New-Object System.Management.PutOptions
    $po.Timeout = [TimeSpan]::FromSeconds($TimeoutSec)
    $po.Type    = [System.Management.PutType]::UpdateOnly      # yalnızca var olan nesneyi güncelle, asla yeni oluşturma
    [void]$Object.Put($po)
}
'@

# ------------------------------------------------------------------
# Arkaplan listeleme / paket yenileme
# ------------------------------------------------------------------
# "Listele" ve "Paketleri Yenile" WMI sorgularını UI thread'i yerine bu worker'da çalıştırır.
# Sonuçlar $Job.Result içine konur; UI, Update-ListJob ile hazır olduğunda grid'leri tek seferde doldurur.
$script:ListWorker = {
    param(
        [string]   $Site,
        [string]   $Server,
        [string]   $HelperSource,
        [hashtable]$Opt,
        $Job
    )

    function Add-JobLog {
        param([string]$Message, [string]$Level = 'INFO')
        [void]$Job.Log.Add(@{ Time = (Get-Date); Level = $Level; Message = $Message })
    }

    function ConvertTo-DateText {
        param($Value)
        if (-not $Value) { return '' }
        try { return ([System.Management.ManagementDateTimeConverter]::ToDateTime([string]$Value)).ToString('yyyy-MM-dd') }
        catch { return [string]$Value }
    }

    . ([scriptblock]::Create($HelperSource))

    try {
        $Job.Status = 'Çalışıyor'
        $scope = New-SmsScope -Server $Server -Site $Site
        $result = $Job.Result

        if ($Opt.Mode -eq 'All') {
            # --- SUG'lar ---
            Add-JobLog 'SUG listesi okunuyor...'
            $groups = New-Object System.Collections.ArrayList
            foreach ($r in (@(Invoke-SmsQuery -Scope $scope -Query 'SELECT * FROM SMS_AuthorizationList') | Sort-Object { [string]$_.LocalizedDisplayName })) {
                [void]$groups.Add([pscustomobject]@{
                    Id         = [string]$r.CI_ID
                    Name       = [string]$r.LocalizedDisplayName
                    Total      = [int]$r.NumberOfUpdates
                    IsDeployed = ($r.IsDeployed -eq $true)
                })
            }
            Add-JobLog "$($groups.Count) Software Update Group bulundu."

            # --- Required güncellemeler (Superseded = 0, Required > N, Vendor) ---
            # Vendor, SMS_SoftwareUpdate üzerinde bir alan değil; güncellemenin "Company" kategorisidir.
            # WQL dizi alanlarını filtreleyemediği için önce vendor'ın kategori ID'si bulunur, sonra ana
            # sorguda SMS_CIAllCategories alt sorgusuyla filtreleme SUNUCUDA yapılır.
            $vendor = [string]$Opt.Vendor
            $search = [string]$Opt.Search
            $minReq = [int]$Opt.MinReq
            $updates = @{}
            $vendorOk = $true
            $vendorIds = @()
            $vendorClause = ''
            if ($vendor) {
                $vEsc = $vendor.Replace('\', '\\').Replace("'", "\'")
                $vRows = @(Invoke-SmsQuery -Scope $scope -Query "SELECT CategoryInstanceID, LocalizedCategoryInstanceName FROM SMS_UpdateCategoryInstance WHERE CategoryTypeName = 'Company' AND LocalizedCategoryInstanceName = '$vEsc'")
                if ($vRows.Count -eq 0) {
                    $all = @(Invoke-SmsQuery -Scope $scope -Query "SELECT LocalizedCategoryInstanceName FROM SMS_UpdateCategoryInstance WHERE CategoryTypeName = 'Company'" |
                             ForEach-Object { [string]$_.LocalizedCategoryInstanceName } | Sort-Object)
                    Add-JobLog "Vendor '$vendor' bulunamadı. Mevcut vendor'lar: $($all -join ', ')" 'WARN'
                    $vendorOk = $false
                } else {
                    $vendorIds = @($vRows | ForEach-Object { [int]$_.CategoryInstanceID })
                    $cond = ($vendorIds | ForEach-Object { "SMS_CIAllCategories.CategoryInstanceID = $_" }) -join ' OR '
                    $vendorClause = " AND SMS_SoftwareUpdate.CI_ID IN (SELECT SMS_CIAllCategories.CI_ID FROM SMS_CIAllCategories WHERE $cond)"
                }
            }

            if ($vendorOk) {
                $base = "SELECT CI_ID, ArticleID, LocalizedDisplayName, DatePosted, IsSuperseded, IsExpired, IsDeployed, IsContentProvisioned, NumMissing, LocalizedCategoryInstanceNames FROM SMS_SoftwareUpdate WHERE IsSuperseded = 0 AND NumMissing > $minReq"
                if ($Opt.ExcludeExpired) { $base += ' AND IsExpired = 0' }
                Add-JobLog "Güncellemeler okunuyor (Vendor = '$vendor', Superseded = 0, Required > $minReq)..."
                try {
                    $rows = @(Invoke-SmsQuery -Scope $scope -Query ($base + $vendorClause))
                } catch {
                    if (-not $vendorClause) { throw }
                    # Yalnızca SMS Provider alt sorguyu reddederse: vendor'a ait CI_ID kümesi ayrı sorguyla alınır.
                    Add-JobLog "Alt sorgu reddedildi ($($_.Exception.Message)); vendor CI_ID listesi ile devam ediliyor." 'WARN'
                    $set = New-Object 'System.Collections.Generic.HashSet[string]'
                    foreach ($vid in $vendorIds) {
                        foreach ($c in @(Invoke-SmsQuery -Scope $scope -Query "SELECT CI_ID FROM SMS_CIAllCategories WHERE CategoryInstanceID = $vid")) { [void]$set.Add([string]$c.CI_ID) }
                    }
                    $rows = @(Invoke-SmsQuery -Scope $scope -Query $base | Where-Object { $set.Contains([string]$_.CI_ID) })
                }
                Add-JobLog "Sunucu $($rows.Count) kayıt döndürdü."
                if ($rows.Count -eq 0 -and $minReq -gt 0) {
                    Add-JobLog "Not: 'Required > $minReq' koşulu Required = $minReq olan güncellemeleri hariç tutar. Konsoldaki güncellemelerde Required = $minReq ise sayıyı $($minReq - 1) yapın." 'WARN'
                }

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
                    $updates[$key] = [pscustomobject]@{
                        Id                 = $key
                        KB                 = $kb
                        Title              = $title
                        Posted             = ConvertTo-DateText $r.DatePosted
                        Required           = [int]$r.NumMissing
                        Category           = $cat
                        ContentProvisioned = ($r.IsContentProvisioned -eq $true)
                    }
                }
                Add-JobLog "Filtre sonrası $($updates.Count) güncelleme kaldı."
            }

            # --- Üyelikler (SMS_CIRelation, RelationType = 1) ---
            # SUG başına ayrı sorgu yerine (N+1 problemi) tüm ilişkiler TEK sorguda çekilip bellekte süzülür.
            Add-JobLog 'SUG üyelikleri okunuyor...'
            $members = @{}
            foreach ($k in @($updates.Keys)) { $members[$k] = New-Object 'System.Collections.Generic.HashSet[string]' }
            $groupIds = New-Object 'System.Collections.Generic.HashSet[string]'
            foreach ($g in $groups) { if ($g.Id -match '^\d+$') { [void]$groupIds.Add($g.Id) } }
            if ($groupIds.Count -gt 0 -and $updates.Count -gt 0) {
                foreach ($r in @(Invoke-SmsQuery -Scope $scope -Query 'SELECT FromCIID, ToCIID FROM SMS_CIRelation WHERE RelationType = 1' -TimeoutSec 600)) {
                    $from = [string]$r.FromCIID
                    if (-not $groupIds.Contains($from)) { continue }
                    $to = [string]$r.ToCIID
                    if ($members.ContainsKey($to)) { [void]$members[$to].Add($from) }
                }
            }

            if ($Opt.NoSugOnly) {
                $before = $updates.Count
                foreach ($k in @($updates.Keys)) {
                    if ($members.ContainsKey($k) -and $members[$k].Count -gt 0) { $updates.Remove($k); $members.Remove($k) }
                }
                Add-JobLog "Zaten bir SUG'a üye olan $($before - $updates.Count) güncelleme listeden çıkarıldı."
            }

            $result.Groups  = $groups.ToArray()
            $result.Updates = $updates
            $result.Members = $members
        }

        # --- Deployment Package'lar (her iki modda) ---
        Add-JobLog 'Deployment Package listesi okunuyor...'
        $pkgs = New-Object System.Collections.ArrayList
        foreach ($r in (@(Invoke-SmsQuery -Scope $scope -Query 'SELECT PackageID, Name, PkgSourcePath FROM SMS_SoftwareUpdatesPackage') | Sort-Object { [string]$_.Name })) {
            [void]$pkgs.Add([pscustomobject]@{ Id = [string]$r.PackageID; Name = [string]$r.Name; SourcePath = [string]$r.PkgSourcePath })
        }
        $result.Packages = $pkgs.ToArray()
        Add-JobLog "$($pkgs.Count) Deployment Package bulundu."
        $Job.Status = 'Tamamlandı'
    }
    catch {
        $Job.Status = 'Hata'
        $Job.Error  = $_.Exception.Message
    }
    finally {
        $Job.EndTime = Get-Date
    }
}

# Mode: 'All' (Listele) veya 'Packages' (yalnızca paket listesi). Aynı anda tek listeleme işi çalışır.
function Start-ListJob {
    param([ValidateSet('All', 'Packages')][string]$Mode)
    if ($script:ListJob) { return }
    $c = Get-Connection
    $opt = @{
        Mode           = $Mode
        Vendor         = $txtVendor.Text.Trim()
        Search         = $txtTitle.Text.Trim()
        MinReq         = [int]$nudReq.Value
        ExcludeExpired = $chkExpired.Checked
        NoSugOnly      = $chkNoSug.Checked
    }
    $job = [hashtable]::Synchronized(@{
        Id = [guid]::NewGuid().ToString(); Mode = $Mode; Status = 'Bekliyor'; Error = ''
        StartTime = Get-Date; EndTime = $null; Cleaned = $false; PS = $null; Handle = $null
        Result = [hashtable]::Synchronized(@{})
        Log    = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
        LogPos = 0
    })
    $script:ListBusy = $true      # bu sürede biten üyelik işlerinin sonuçları bekletilir (yarım listeye uygulanmasın)
    Set-Busy $true
    try {
        Start-PooledScript -PoolName 'List' -Script $script:ListWorker -Arguments @($c.Site, $c.Server, $script:WmiHelperSource, $opt) -Job $job
    } catch {
        $script:ListBusy = $false
        Set-Busy $false
        throw
    }
    $script:ListJob = $job
    $tmrList.Start()
}

function Update-ListJob {
    $job = $script:ListJob
    if (-not $job) { $tmrList.Stop(); return }
    $completed = $job.Handle -and $job.Handle.IsCompleted
    $cnt = $job.Log.Count
    while ($job.LogPos -lt $cnt) {
        $l = $job.Log[$job.LogPos]
        $job.LogPos = [int]$job.LogPos + 1
        Write-Log -Message $l.Message -Level $l.Level -Time $l.Time
    }
    if (-not $completed) { return }

    $tmrList.Stop()
    Close-PooledScript $job
    $script:ListJob = $null
    try {
        if ($job.Status -ne 'Tamamlandı') {
            $msg = if ($job.Error) { $job.Error } else { 'Bilinmeyen hata.' }
            Write-Log "Listeleme başarısız: $msg" 'ERROR'
            Show-Message $msg 'Hata' 'Error'
            return
        }
        $r = $job.Result
        if ($job.Mode -eq 'All') {
            $script:Groups = @($r.Groups)
            $script:GroupById = @{}
            foreach ($g in $script:Groups) { $script:GroupById[$g.Id] = $g }
            $script:Updates = $r.Updates
            $script:Members = $r.Members
            Show-Groups
        }
        $script:Packages = @($r.Packages)
        $script:PackageById = @{}
        foreach ($p in $script:Packages) { $script:PackageById[$p.Id] = $p }
        Show-Packages
        if ($job.Mode -eq 'All') {
            Show-Updates
            Set-GroupFilter
            Write-Log "Listeleme tamamlandı: $($script:Updates.Count) benzersiz güncelleme, $($script:Groups.Count) SUG ($(Format-Elapsed $job))." 'OK'
        }
    } catch {
        Write-Log "Liste ekrana yansıtılamadı: $($_.Exception.Message)" 'ERROR'
    } finally {
        $script:ListBusy = $false
        Set-Busy $false
        # Listeleme sırasında biten üyelik işleri varsa şimdi uygulansın
        if (@($script:MembershipJobs.ToArray() | Where-Object { -not $_.Cleaned }).Count -gt 0) { $tmrMembership.Start() }
    }
}

function Invoke-List {
    try {
        $active = @($script:MembershipJobs.ToArray() | Where-Object { -not $_.Cleaned }).Count
        if ($active -gt 0) {
            Write-Log "Not: $active üyelik işi arkaplanda sürüyor; ilgili satırlar kilitli kalacak ve işler bitince tazelenecek." 'WARN'
        }
        Start-ListJob -Mode 'All'
    } catch {
        Write-Log $_.Exception.Message 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    }
}

# ------------------------------------------------------------------
# Grid doldurma ve hedefli yenileme
# ------------------------------------------------------------------
function Get-MemberNames {
    param([string]$UpdateId)
    if (-not $script:Members.ContainsKey($UpdateId)) { return '' }
    $names = @()
    foreach ($gid in $script:Members[$UpdateId]) {
        if ($script:GroupById.ContainsKey($gid)) { $names += $script:GroupById[$gid].Name }
    }
    return (($names | Sort-Object) -join '; ')
}

function Show-Groups {
    $prev = $script:Suspend
    $script:Suspend = $true
    $gridGrp.SuspendLayout()
    try {
        $gridGrp.CurrentCell = $null
        $gridGrp.Rows.Clear()
        $script:GrpRowById = @{}
        $unchecked = [System.Windows.Forms.CheckState]::Unchecked
        foreach ($g in $script:Groups) {
            $dep = if ($g.IsDeployed) { 'Evet' } else { '' }
            # Değerler tek çağrıda verilir (hücre hücre atamaya göre daha az çizim/olay yükü).
            # Sıra: gChk, gName, gDep, gSel, gTotal
            $i = $gridGrp.Rows.Add([object[]]@($unchecked, $g.Name, $dep, '', $g.Total))
            $row = $gridGrp.Rows[$i]
            $row.Tag = $g.Id
            if ($g.IsDeployed) { $row.Cells['gDep'].Style.ForeColor = [System.Drawing.Color]::DarkRed }
            $script:GrpRowById[$g.Id] = $row
        }
    } finally {
        $gridGrp.ResumeLayout()
        $script:Suspend = $prev
    }
}

# SUG satırının görünümünü varsayılana döndürür ("bekleyen değişiklik" sarısını temizler).
# SUG'lar KİLİTLENMEZ: aynı SUG üzerinde aynı anda birden fazla üyelik işi başlatılabilir; aynı SUG'a
# yazımlar arkaplanda SUG başına mutex ile sıraya girer.
function Set-GroupRowStyle {
    param($Row)
    $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Empty
}

# Güncelleme satırının "İçerik" ve "Üye olduğu SUG'lar" hücrelerini ve kilit görünümünü tazeler.
function Set-UpdateRowView {
    param($Row)
    $id = [string]$Row.Tag
    $names = Get-MemberNames $id
    if ($script:Updates.ContainsKey($id)) {
        $u = $script:Updates[$id]
        $Row.Cells['uContent'].Value = if ($u.ContentProvisioned) { 'Var' } else { 'Yok' }
        $Row.Cells['uContent'].Style.ForeColor = if ($u.ContentProvisioned) { [System.Drawing.Color]::DarkGreen } else { [System.Drawing.Color]::DarkOrange }
    }
    if ($script:LockedUpdates.Contains($id)) {
        $Row.Cells['uGroups'].Value = "[İşlemde] $names"
        $Row.DefaultCellStyle.BackColor = $script:LockBackColor
        $Row.DefaultCellStyle.ForeColor = $script:LockForeColor
        $Row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::SlateGray
    } else {
        $Row.Cells['uGroups'].Value = $names
        $Row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Empty
        $Row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Empty
        $Row.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::Empty
    }
}

# Yalnızca verilen güncelleme ID'lerinin grid satırlarını tazeler (tüm listeyi yeniden çizmez).
function Update-UpdateRows {
    param([string[]]$UpdateIds)
    if (-not $UpdateIds -or $UpdateIds.Count -eq 0) { return }
    $prev = $script:Suspend
    $script:Suspend = $true
    $gridUpd.SuspendLayout()
    try {
        foreach ($uid in $UpdateIds) {
            if ($script:UpdRowById.ContainsKey($uid)) { Set-UpdateRowView $script:UpdRowById[$uid] }
        }
    } finally {
        $gridUpd.ResumeLayout()
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
    if ($prevText) { $cmbPkg.Text = $prevText }
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
    return $langs.ToArray()
}

function Show-Updates {
    $prev = $script:Suspend
    $script:Suspend = $true
    $gridUpd.SuspendLayout()
    try {
        $gridUpd.Rows.Clear()
        $script:UpdRowById = @{}
        $sorted = $script:Updates.Values | Sort-Object @{ Expression = 'Required'; Descending = $true }, KB, Title
        foreach ($u in $sorted) {
            # Sütun sırası: uKB, uTitle, uReq, uContent, uPosted, uCat, uGroups, uId
            $i = $gridUpd.Rows.Add([object[]]@($u.KB, $u.Title, $u.Required, '', $u.Posted, $u.Category, '', $u.Id))
            $row = $gridUpd.Rows[$i]
            $row.Tag = $u.Id
            $script:UpdRowById[$u.Id] = $row
            Set-UpdateRowView $row
        }
        $gridUpd.ClearSelection()
    } finally {
        $gridUpd.ResumeLayout()
        $script:Suspend = $prev
    }
    Update-GroupChecks
}

function Get-SelectedUpdateIds {
    $ids = New-Object 'System.Collections.Generic.List[string]'
    foreach ($row in $gridUpd.SelectedRows) { $ids.Add([string]$row.Tag) }
    return $ids.ToArray()     # çağıranlar @(...) ile sarar (0/1 öğede de dizi olur)
}

# Seçili güncellemelerin her SUG'a kaç tanesinin üye olduğunu tek geçişte sayar.
# (SUG sayısı x seçim sayısı kadar döngü yerine yalnızca seçili güncellemelerin üyelikleri dolaşılır.)
function Get-SelectionGroupCounts {
    param([string[]]$SelectedIds)
    $counts = @{}
    foreach ($uid in $SelectedIds) {
        if (-not $script:Members.ContainsKey($uid)) { continue }
        foreach ($gid in $script:Members[$uid]) {
            if ($counts.ContainsKey($gid)) { $counts[$gid]++ } else { $counts[$gid] = 1 }
        }
    }
    return $counts
}

# Tek bir SUG satırının kutu durumunu / Seçili / Toplam hücrelerini ve kilit görünümünü ayarlar.
function Set-GroupRowState {
    param($Row, [hashtable]$Counts, [int]$N)
    $gid = [string]$Row.Tag
    $m = 0
    if ($Counts.ContainsKey($gid)) { $m = [int]$Counts[$gid] }
    if ($N -eq 0 -or $m -eq 0) { $st = [System.Windows.Forms.CheckState]::Unchecked }
    elseif ($m -eq $N)         { $st = [System.Windows.Forms.CheckState]::Checked }
    else                       { $st = [System.Windows.Forms.CheckState]::Indeterminate }

    $script:OrigCheck[$gid] = $st
    $Row.Cells['gChk'].Value = $st
    if ($N -gt 0) { $Row.Cells['gSel'].Value = "$m/$N" } else { $Row.Cells['gSel'].Value = '' }
    if ($script:GroupById.ContainsKey($gid)) { $Row.Cells['gTotal'].Value = $script:GroupById[$gid].Total }
    Set-GroupRowStyle $Row
}

function Update-CountLabel {
    param([int]$N)
    $txt = "$($gridUpd.Rows.Count) güncelleme listelendi"
    if ($N -gt 0) { $txt += ", $N seçili" }
    if ($script:LockedUpdates.Count -gt 0) { $txt += ", $($script:LockedUpdates.Count) işlemde (kilitli)" }
    $lblCount.Text = $txt + '.'
}

# Seçili güncellemelere göre SUG kutularını (tam / boş / kısmi) yeniden hesaplar.
function Update-GroupChecks {
    $sel = @(Get-SelectedUpdateIds)
    $n = $sel.Count
    $counts = Get-SelectionGroupCounts -SelectedIds $sel
    $prev = $script:Suspend
    $script:Suspend = $true
    $gridGrp.SuspendLayout()
    try {
        foreach ($row in $gridGrp.Rows) { Set-GroupRowState -Row $row -Counts $counts -N $n }
    } finally {
        $gridGrp.ResumeLayout()
        $script:Suspend = $prev
    }
    Update-CountLabel $n
}

# Yalnızca verilen SUG'ların satırlarını (kutu, Seçili, Toplam, kilit) mevcut seçime göre tazeler.
function Update-GroupRows {
    param([string[]]$GroupIds)
    if (-not $GroupIds -or $GroupIds.Count -eq 0) { return }
    $sel = @(Get-SelectedUpdateIds)
    $n = $sel.Count
    $counts = Get-SelectionGroupCounts -SelectedIds $sel
    $prev = $script:Suspend
    $script:Suspend = $true
    $gridGrp.SuspendLayout()
    try {
        foreach ($gid in $GroupIds) {
            if ($script:GrpRowById.ContainsKey($gid)) { Set-GroupRowState -Row $script:GrpRowById[$gid] -Counts $counts -N $n }
        }
    } finally {
        $gridGrp.ResumeLayout()
        $script:Suspend = $prev
    }
    Update-CountLabel $n
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
# SUG üyelik talebi (Uygula / Geri Al) - doğrulama ve onay
# ------------------------------------------------------------------
# Plandaki deploy edilmiş SUG'lar için kullanıcıya gösterilecek risk uyarılarını üretir.
function Get-DeployWarnings {
    param([object[]]$Plan)
    $w = New-Object 'System.Collections.Generic.List[string]'
    foreach ($p in $Plan) {
        if (-not $script:GroupById.ContainsKey($p.GroupId)) { continue }
        $g = $script:GroupById[$p.GroupId]
        if (-not $g.IsDeployed) { continue }
        if (@($p.Remove).Count -gt 0) {
            $w.Add("- '$($g.Name)' DEPLOY EDİLMİŞ: $(@($p.Remove).Count) güncelleme ÇIKARILACAK. Bu güncellemelerin deployment'ı istemcilerden geri çekilir (henüz kurmamış istemciler kurmaz).")
        }
        $nc = @(@($p.Add) | Where-Object { $script:Updates.ContainsKey($_) -and -not $script:Updates[$_].ContentProvisioned }).Count
        if ($nc -gt 0) {
            $w.Add("- '$($g.Name)' DEPLOY EDİLMİŞ: içeriği İNDİRİLMEMİŞ $nc güncelleme eklenecek. İçerik bir pakete indirilip DP'lere dağıtılana kadar istemciler bu güncellemeleri kuramaz.")
        }
    }
    return $w.ToArray()
}

# Plan için ortak kontroller + onay penceresi. Onaylanırsa $true döner.
function Confirm-MembershipPlan {
    param([object[]]$Plan, [int]$UpdateCount, [bool]$Simulate, [string]$Title, [string]$Verb)

    # Savunma amaçlı: kilitli (arkaplanda işlemde olan) güncellemeler plana giremez. SUG'lar kilitlenmez.
    $lockedUpd = @($Plan | ForEach-Object { @($_.Add) + @($_.Remove) } | Where-Object { $_ -and $script:LockedUpdates.Contains($_) } | Select-Object -Unique)
    if ($lockedUpd.Count -gt 0) {
        Show-Message "Bu talepteki $($lockedUpd.Count) güncelleme şu anda arkaplanda işlemde (kilitli); çakışmayı önlemek için talep gönderilmedi.`n`nİlgili işler bitince tekrar deneyin." 'Uyarı' 'Warning'
        return $false
    }

    # SUG başına 1000 güncelleme sınırı (worker yazmadan önce gerçek sayıyla tekrar kontrol eder)
    $over = @()
    foreach ($p in $Plan) {
        if (-not $script:GroupById.ContainsKey($p.GroupId)) { continue }
        $grp = $script:GroupById[$p.GroupId]
        $newTotal = $grp.Total + @($p.Add).Count - @($p.Remove).Count
        if ($newTotal -gt $script:SugLimit) { $over += ('- {0}: {1} güncelleme olurdu' -f $grp.Name, $newTotal) }
    }
    if ($over.Count -gt 0) {
        Show-Message ("Aşağıdaki SUG'lar $($script:SugLimit) güncelleme sınırını aşacağı için işlem yapılmadı:`n`n" + ($over -join "`n")) 'Uyarı' 'Warning'
        return $false
    }

    $lines = @()
    foreach ($p in $Plan) {
        $total = if ($script:GroupById.ContainsKey($p.GroupId)) { $script:GroupById[$p.GroupId].Total } else { '?' }
        $after = if ($total -is [int]) { $total + @($p.Add).Count - @($p.Remove).Count } else { '?' }
        $lines += ('- {0}:  +{1} ekle,  -{2} çıkar  (toplam: {3} -> {4})' -f (Get-PlanGroupName $p), @($p.Add).Count, @($p.Remove).Count, $total, $after)
    }
    $warnings = @(Get-DeployWarnings -Plan $Plan)

    $head = if ($Simulate) { "SİMÜLASYON - hiçbir değişiklik yapılmayacak.`n`n" } else { '' }
    $msg = "{0}{1} güncelleme için aşağıdaki SUG değişiklikleri ARKAPLANDA {2}:`n`n{3}" -f $head, $UpdateCount, $Verb, ($lines -join "`n")
    if ($warnings.Count -gt 0) {
        $msg += "`n`n!!! DİKKAT - DEPLOY EDİLMİŞ SUG'LAR !!!`n" + ($warnings -join "`n")
    }
    $msg += "`n`nİşlem sürerken ilgili güncellemeler kilitli kalır; SUG'lar kilitlenmez (aynı SUG'a yazımlar arkaplanda sıraya girer). Devam edilsin mi?"

    if ($warnings.Count -gt 0 -and -not $Simulate) {
        # Riskli işlemde varsayılan düğme "Hayır"
        $ans = [System.Windows.Forms.MessageBox]::Show($msg, $Title, 'YesNo', 'Warning', 'Button2')
    } else {
        $ans = [System.Windows.Forms.MessageBox]::Show($msg, $Title, 'YesNo', 'Question')
    }
    return ($ans -eq 'Yes')
}

function Get-PlanGroupName {
    param($PlanItem)
    if ($PlanItem.PSObject.Properties['GroupName'] -and $PlanItem.GroupName) { return [string]$PlanItem.GroupName }
    if ($script:GroupById.ContainsKey($PlanItem.GroupId)) { return $script:GroupById[$PlanItem.GroupId].Name }
    return "CI_ID=$($PlanItem.GroupId)"
}

function Invoke-ApplyMembership {
    $selAll = @(Get-SelectedUpdateIds)
    if ($selAll.Count -eq 0) { Show-Message 'Önce soldaki listeden en az bir güncelleme seçin.' 'Uyarı' 'Warning'; return }

    # Kilitli (arkaplanda işlemde olan) güncellemeler bu talebe dahil edilemez.
    $sel = @($selAll | Where-Object { -not $script:LockedUpdates.Contains($_) })
    $lockedSel = $selAll.Count - $sel.Count
    if ($sel.Count -eq 0) {
        Show-Message 'Seçili güncellemelerin tamamı şu anda arkaplanda işlemde (kilitli). İşlem bitince tekrar deneyin.' 'Uyarı' 'Warning'
        return
    }
    if ($lockedSel -gt 0) {
        $q = "Seçili $($selAll.Count) güncellemeden $lockedSel tanesi şu anda arkaplanda işlemde (kilitli).`n`nBu $lockedSel güncelleme HARİÇ tutularak kalan $($sel.Count) güncelleme ile devam edilsin mi?"
        if ([System.Windows.Forms.MessageBox]::Show($q, 'Kilitli güncellemeler', 'YesNo', 'Question') -ne 'Yes') { return }
    }

    $current = @{}
    foreach ($row in $gridGrp.Rows) { $current[[string]$row.Tag] = ConvertTo-CheckState $row.Cells['gChk'].Value }

    $plan = @(Get-MembershipPlan -SelectedIds $sel -Members $script:Members -Orig $script:OrigCheck -Current $current)
    if ($plan.Count -eq 0) { Show-Message 'Grup kutularında (kilitli olmayan güncellemeler için) değişiklik yok.' 'Bilgi' 'Information'; return }
    $plan = @($plan | Sort-Object { $script:GroupById[$_.GroupId].Name })

    $simulate = $chkSim.Checked
    if (-not (Confirm-MembershipPlan -Plan $plan -UpdateCount $sel.Count -Simulate $simulate -Title 'SUG üyeliklerini uygula' -Verb 'uygulanacak')) { return }

    try {
        Start-MembershipJob -Plan $plan -Simulate $simulate -Kind 'Uygulama'
    }
    catch {
        Write-Log "Üyelik işi başlatılamadı: $($_.Exception.Message)" 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    }
}

# ------------------------------------------------------------------
# Geri alma (Undo)
# ------------------------------------------------------------------
# Her başarılı (simülasyon olmayan) "Uygulama" işinin SUG bazında GERÇEKTEN eklenen/çıkarılan
# ID'leri $script:UndoStack'e konur (en yeni en sonda). Geri alma, ters planı yeni bir arkaplan
# işi olarak başlatır (Ekle <-> Çıkar). Geri alma işlerinin kendisi yığına eklenmez.
function Update-UndoButton {
    $n = $script:UndoStack.Count
    $btnUndo.Enabled = ($n -gt 0) -and (-not $script:ListBusy)
    $btnUndo.Text = if ($n -gt 0) { "Son Üyelik İşlemini Geri Al ($n)" } else { 'Son Üyelik İşlemini Geri Al' }
}

function Invoke-UndoMembership {
    if ($script:UndoStack.Count -eq 0) { return }
    $rec = $script:UndoStack[$script:UndoStack.Count - 1]

    $plan = @()
    $ids = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($it in $rec.Items) {
        $plan += [pscustomobject]@{
            GroupId   = [string]$it.GroupId
            GroupName = [string]$it.GroupName
            Add       = [string[]]@($it.Removed)     # çıkarılanlar geri eklenir
            Remove    = [string[]]@($it.Added)       # eklenenler geri çıkarılır
        }
        foreach ($u in @($it.Added) + @($it.Removed)) { if ($u) { [void]$ids.Add([string]$u) } }
    }
    if ($plan.Count -eq 0) { $script:UndoStack.RemoveAt($script:UndoStack.Count - 1); Update-UndoButton; return }

    $simulate = $chkSim.Checked
    $title = "Geri al: {0} tarihli üyelik işlemi" -f $rec.Time.ToString('HH:mm:ss')
    if (-not (Confirm-MembershipPlan -Plan $plan -UpdateCount $ids.Count -Simulate $simulate -Title $title -Verb 'GERİ ALINACAK')) { return }

    try {
        Start-MembershipJob -Plan $plan -Simulate $simulate -Kind 'Geri Alma'
        if (-not $simulate) { $script:UndoStack.RemoveAt($script:UndoStack.Count - 1) }   # simülasyonda kayıt korunur
    } catch {
        Write-Log "Geri alma işi başlatılamadı: $($_.Exception.Message)" 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    }
    Update-UndoButton
}

# ------------------------------------------------------------------
# Arkaplan SUG üyelik değişimi (runspace havuzu) - saf WMI, SUG başına mutex, doğrulama
# ------------------------------------------------------------------
# Her "Uygula"/"Geri Al" talebi havuzdaki bir runspace'te çalışan bir üyelik işidir:
#   - UI thread'i bloklanmaz; form yanıt vermeye devam eder.
#   - UI'da YALNIZCA işteki güncellemeler ($script:LockedUpdates) kilitlenir. SUG'lar kilitlenmez:
#     aynı SUG üzerinde aynı anda birden fazla iş başlatılabilir. İki iş aynı güncellemeyi
#     değiştiremez (güncelleme kilidi), aynı SUG'a yazımları ise aşağıdaki mutex ile sıraya girer.
#   - SCCM erişimi YALNIZCA WMI (DCOM, PacketPrivacy) ile yapılır; ConfigurationManager modülü,
#     CIM/WinRM veya uzak oturum KULLANILMAZ (Console kurulumu ve modül yükleme süresi gerekmez).
#   - Yazmadan hemen önce eklenecek güncellemelerin güncel durumu sunucudan sorgulanır
#     (Expired / Superseded / bulunamayan güncellemeler eklenmez).
#   - Her SUG için: [SUG mutex al] -> SMS_AuthorizationList'i TAM oku (lazy Updates) -> ekle/çıkar ->
#     TEK Put() -> yeniden oku ve DOĞRULA -> [mutex bırak]. Eksik kalanlar NotAdded/NotRemoved olarak raporlanır.
#   - SUG mutex'i "Local\SccmSupersededToolSug_<CI_ID>" adını kullanır: superseed_kontrol_ui.ps1 ile AYNI
#     ad olduğu için iki araç aynı oturumda aynı SUG'a aynı anda yazamaz (bu araçtaki tüm runspace'ler ve
#     açık pencereler de dahil). Mutex SCCM Console'dan yapılan eşzamanlı değişikliklere karşı koruma
#     SAĞLAMAZ; onlara karşı yalnızca Put() sonrası doğrulama vardır.
#   - Arkaplan thread'i UI kontrollerine dokunmaz; günlük satırları ve sonuçlar senkronize hashtable
#     üzerinden UI'ya "olay" olarak aktarılır (Level = 'AUDIT' satırları yalnızca denetim dosyasına yazılır).
$script:MembershipWorker = {
    param(
        [string]  $Site,
        [string]  $Server,
        [string]  $HelperSource,
        [object[]]$Plan,
        [int]     $SugLimit,
        [bool]    $Simulate,
        $Job
    )

    # UI'ya dokunmayan günlük: satırlar işin kuyruğuna yazılır, UI zamanlayıcısı aktarır.
    function Add-JobLog {
        param([string]$Message, [string]$Level = 'INFO')
        [void]$Job.Log.Add(@{ Time = (Get-Date); Level = $Level; Message = $Message })
    }

    function Assert-CiId {
        param([string]$Id)
        if ($Id -notmatch '^\d{1,10}$' -or [uint64]$Id -gt [uint32]::MaxValue) { throw "Geçersiz CI_ID: '$Id'" }
    }

    . ([scriptblock]::Create($HelperSource))

    # SUG'un güncel üye listesini (lazy Updates dahil) okur.
    function Read-SugMembers {
        param($Scope, [string]$GroupId)
        try {
            $o = Get-SmsObject -Scope $Scope -Path "SMS_AuthorizationList.CI_ID=$GroupId"
        } catch {
            throw "WMI: SUG okunamadı (CI_ID=$GroupId): $($_.Exception.Message)"
        }
        $set  = New-Object 'System.Collections.Generic.HashSet[uint32]'
        $list = New-Object 'System.Collections.Generic.List[uint32]'
        foreach ($u in @($o.Updates)) { if ($null -ne $u -and $set.Add([uint32]$u)) { $list.Add([uint32]$u) } }
        return @{ Object = $o; Set = $set; List = $list }
    }

    # İç fonksiyon: oku -> değiştir -> Put -> doğrula. YALNIZCA SUG mutex'i alınmışken çağrılır.
    function Set-SugMembershipWmiCore {
        param($Scope, [string]$GroupId, [string[]]$Add, [string[]]$Remove)
        $r = Read-SugMembers -Scope $Scope -GroupId $GroupId
        $pre = [System.Collections.Generic.HashSet[uint32]]::new($r.Set)
        Add-JobLog ("SugChangeStart iş={0} SUG={1} önceki_üye_sayısı={2} eklenecek=[{3}] çıkarılacak=[{4}]" -f $Job.Id, $GroupId, $pre.Count, (@($Add) -join ','), (@($Remove) -join ',')) 'AUDIT'

        $remSet = New-Object 'System.Collections.Generic.HashSet[uint32]'
        foreach ($u in @($Remove)) { if ($u) { [void]$remSet.Add([uint32]$u) } }
        $seen = New-Object 'System.Collections.Generic.HashSet[uint32]'
        $new  = New-Object 'System.Collections.Generic.List[uint32]'
        $removed = 0
        foreach ($u in $r.List) { if ($remSet.Contains($u)) { $removed++ } else { $new.Add($u); [void]$seen.Add($u) } }
        $added = 0
        foreach ($u in @($Add)) {
            if (-not $u) { continue }
            $i = [uint32]$u
            if (-not $remSet.Contains($i) -and $seen.Add($i)) { $new.Add($i); $added++ }
        }
        if ($new.Count -gt $SugLimit) { throw "SUG $SugLimit güncelleme sınırını aşardı ($($new.Count))." }
        $expected = [System.Collections.Generic.HashSet[uint32]]::new($new)
        if ($added -gt 0 -or $removed -gt 0) {
            $r.Object.Updates = [uint32[]]$new.ToArray()
            Save-SmsObject -Object $r.Object
        }

        # Put() sonrası doğrulama (her zaman yapılır). Okuma başarısız olursa yazım geri alınmaz;
        # sonuç "doğrulanmadı" olarak raporlanır.
        $post = $null; $verifyErr = ''
        try { $post = (Read-SugMembers -Scope $Scope -GroupId $GroupId).Set } catch { $verifyErr = $_.Exception.Message }
        return @{ Pre = $pre; Expected = $expected; Post = $post; VerifyError = $verifyErr; Added = $added; Removed = $removed }
    }

    # Dış fonksiyon: doğrulama + SUG başına isimli mutex. Bekleme iptal edilebilir (en fazla 3 dk).
    function Set-SugMembershipWmi {
        param($Scope, [string]$GroupId, [string]$GroupName, [string[]]$Add, [string[]]$Remove)
        Assert-CiId $GroupId
        foreach ($x in (@($Add) + @($Remove))) { if ($x) { Assert-CiId $x } }

        $mtx = New-Object System.Threading.Mutex($false, "Local\SccmSupersededToolSug_$GroupId")
        $owned = $false
        try {
            try { $owned = $mtx.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $owned = $true }
            if (-not $owned) {
                Add-JobLog "'$GroupName' (CI_ID $GroupId) başka bir iş tarafından yazılıyor; sıra bekleniyor..."
                $Job.Current = "$GroupName (sırada bekliyor)"
                $deadline = (Get-Date).AddMinutes(3)
                while (-not $owned) {
                    if ($Job.Cancel) { throw 'İptal edildi (SUG sırası beklenirken).' }
                    if ((Get-Date) -gt $deadline) { throw "SUG (CI_ID $GroupId) 3 dakika içinde yazılabilir olmadı." }
                    try { $owned = $mtx.WaitOne(1000) } catch [System.Threading.AbandonedMutexException] { $owned = $true }
                }
                $Job.Current = $GroupName
            }
            return (Set-SugMembershipWmiCore -Scope $Scope -GroupId $GroupId -Add $Add -Remove $Remove)
        } finally {
            if ($owned) { $mtx.ReleaseMutex() }
            $mtx.Dispose()
        }
    }

    $scope = $null
    try {
        $Job.Status = 'Çalışıyor'

        # --- Girdi doğrulama: tüm ID'ler yalnızca rakam olmalı (WMI yol/sorgu metnine giriyorlar) ---
        foreach ($item in $Plan) {
            Assert-CiId ([string]$item.GroupId)
            foreach ($id in @($item.AddIds) + @($item.RemoveIds)) { if ($id) { Assert-CiId ([string]$id) } }
        }

        # --- Yazma anında güncel durum kontrolü (yalnızca eklenecekler için) ---
        $excluded = @{}     # CI_ID -> sebep
        if (-not $Simulate) {
            $scope = New-SmsScope -Server $Server -Site $Site
            $allAdds = New-Object 'System.Collections.Generic.HashSet[string]'
            foreach ($item in $Plan) { foreach ($id in @($item.AddIds)) { if ($id) { [void]$allAdds.Add([string]$id) } } }
            if ($allAdds.Count -gt 0) {
                try {
                    $ids = @($allAdds)
                    $found = @{}
                    for ($i = 0; $i -lt $ids.Count; $i += 50) {
                        $chunk = $ids[$i..([Math]::Min($i + 49, $ids.Count - 1))]
                        $cond = ($chunk | ForEach-Object { "CI_ID = $_" }) -join ' OR '
                        foreach ($row in @(Invoke-SmsQuery -Scope $scope -Query "SELECT CI_ID, IsExpired, IsSuperseded FROM SMS_SoftwareUpdate WHERE $cond")) {
                            $found[[string]$row.CI_ID] = $row
                        }
                    }
                    foreach ($id in $ids) {
                        if (-not $found.ContainsKey($id))          { $excluded[$id] = 'sunucuda bulunamadı' }
                        elseif ($found[$id].IsExpired -eq $true)    { $excluded[$id] = 'Expired' }
                        elseif ($found[$id].IsSuperseded -eq $true) { $excluded[$id] = 'Superseded' }
                    }
                    if ($excluded.Count -gt 0) {
                        $txt = ($excluded.Keys | ForEach-Object { "$_ ($($excluded[$_]))" }) -join ', '
                        Add-JobLog "Güncel durum kontrolü: $($excluded.Count) güncelleme EKLENMEYECEK: $txt" 'WARN'
                    }
                } catch {
                    Add-JobLog "Güncel durum kontrolü yapılamadı ($($_.Exception.Message)); kontrol atlanarak devam ediliyor." 'WARN'
                }
            }
        }

        foreach ($item in $Plan) {
            $res = @{
                GroupId       = [string]$item.GroupId
                GroupName     = [string]$item.GroupName
                AddIds        = [string[]]@($item.AddIds | Where-Object { $_ -and -not $excluded.ContainsKey([string]$_) })
                RemoveIds     = [string[]]@($item.RemoveIds | Where-Object { $_ })
                SkippedIds    = [string[]]@($item.AddIds | Where-Object { $_ -and $excluded.ContainsKey([string]$_) })
                Ok            = $false
                AddDone       = $false
                RemoveDone    = $false
                Skipped       = $false
                NewTotal      = -1
                Added         = 0
                Removed       = 0
                ActualAdded   = [string[]]@()
                ActualRemoved = [string[]]@()
                NotAdded      = [string[]]@()
                NotRemoved    = [string[]]@()
                PostMember    = $null       # CI_ID -> bool (yazma sonrası sunucudaki gerçek üyelik)
                Verified      = $false
                Method        = ''
                Error         = ''
            }
            if ($Job.Cancel) { $res.Skipped = $true; [void]$Job.Results.Add($res); continue }
            $Job.Current = $res.GroupName
            try {
                if ($Simulate) {
                    Start-Sleep -Milliseconds 300
                    Add-JobLog "[SİMÜLASYON] '$($res.GroupName)': +$($res.AddIds.Count) eklenecek, -$($res.RemoveIds.Count) çıkarılacak." 'WARN'
                    $res.Method = 'Simülasyon'
                    $res.AddDone = $true; $res.RemoveDone = $true
                    $res.Ok = $true
                } elseif ($res.AddIds.Count -eq 0 -and $res.RemoveIds.Count -eq 0) {
                    $res.Method = 'Yok (eklenecek uygun güncelleme kalmadı)'
                    $res.AddDone = $true; $res.RemoveDone = $true
                    $res.Ok = $true
                } else {
                    $w = Set-SugMembershipWmi -Scope $scope -GroupId $res.GroupId -GroupName $res.GroupName -Add $res.AddIds -Remove $res.RemoveIds
                    $res.Method = 'WMI'
                    $res.AddDone = $true; $res.RemoveDone = $true
                    $res.Added = $w.Added; $res.Removed = $w.Removed
                    # DİKKAT: HashSet'ler if-ifadesinin çıktısı olarak atanmaz / if ($x) ile test edilmez (boş küme $false).
                    $pre = $w.Pre; $post = $w.Post; $expected = $w.Expected
                    if ($null -eq $post) {
                        # Doğrulama okuması yapılamadı: beklenen durum esas alınır, sonuç "doğrulanmadı" olarak işaretlenir
                        Add-JobLog "'$($res.GroupName)': yazma sonrası doğrulama okuması yapılamadı ($($w.VerifyError)); sonuç doğrulanmadı." 'WARN'
                        $res.NewTotal      = $expected.Count
                        $res.ActualAdded   = [string[]]@($res.AddIds    | Where-Object { -not $pre.Contains([uint32]$_) -and $expected.Contains([uint32]$_) })
                        $res.ActualRemoved = [string[]]@($res.RemoveIds | Where-Object { $pre.Contains([uint32]$_) -and -not $expected.Contains([uint32]$_) })
                        $res.Ok = $true
                        [void]$Job.Results.Add($res)
                        $Job.Done = [int]$Job.Done + 1
                        continue
                    }
                    $res.NewTotal = $post.Count
                    $res.Verified = $true
                    $pm = @{}
                    foreach ($id in $res.AddIds)    { $pm[$id] = $post.Contains([uint32]$id) }
                    foreach ($id in $res.RemoveIds) { $pm[$id] = $post.Contains([uint32]$id) }
                    $res.PostMember    = $pm
                    $res.ActualAdded   = [string[]]@($res.AddIds    | Where-Object { -not $pre.Contains([uint32]$_) -and $post.Contains([uint32]$_) })
                    $res.ActualRemoved = [string[]]@($res.RemoveIds | Where-Object { $pre.Contains([uint32]$_) -and -not $post.Contains([uint32]$_) })
                    $res.NotAdded      = [string[]]@($res.AddIds    | Where-Object { -not $post.Contains([uint32]$_) })
                    $res.NotRemoved    = [string[]]@($res.RemoveIds | Where-Object { $post.Contains([uint32]$_) })

                    $extra = 0; $gone = 0
                    foreach ($u in $post)     { if (-not $expected.Contains($u)) { $extra++ } }
                    foreach ($u in $expected) { if (-not $post.Contains($u))     { $gone++ } }
                    if ($extra -gt 0 -or $gone -gt 0) {
                        Add-JobLog "'$($res.GroupName)': doğrulamada bu iş dışında bir değişiklik görüldü (+$extra / -$gone güncelleme). SCCM Console'dan aynı anda düzenleme yapılıyor olabilir." 'WARN'
                    }
                    $res.Ok = ($res.NotAdded.Count -eq 0 -and $res.NotRemoved.Count -eq 0)
                    if (-not $res.Ok) {
                        $res.Error = "Doğrulama başarısız: eklenemeyen [$($res.NotAdded -join ',')], çıkarılamayan [$($res.NotRemoved -join ',')]."
                    }
                }
            } catch {
                $res.Error = $_.Exception.Message
            }
            [void]$Job.Results.Add($res)
            $Job.Done = [int]$Job.Done + 1
        }
        $Job.Status = if ($Job.Cancel) { 'İptal Edildi' } else { 'Tamamlandı' }
    }
    catch {
        $Job.Status = 'Hata'
        $Job.Error  = $_.Exception.Message
    }
    finally {
        $Job.Current = ''
        $Job.EndTime = Get-Date
    }
}

# Yeni bir üyelik işini başlatır, YALNIZCA ilgili güncellemeleri kilitler ve HEMEN döner.
# SUG ID'leri işe yalnızca denetim kaydı ve iş bitince hedefli yenileme için aktarılır.
# $Plan öğeleri: GroupId, Add (string[]), Remove (string[]), isteğe bağlı GroupName.
function Start-MembershipJob {
    param([object[]]$Plan, [bool]$Simulate, [string]$Kind = 'Uygulama')

    $c = Get-Connection     # iş başında alınan Site/Server kopyası worker'a verilir (worker UI kutularını okumaz)

    $workerPlan = @()
    $updIds = New-Object 'System.Collections.Generic.HashSet[string]'
    $grpIds = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($p in $Plan) {
        $workerPlan += @{
            GroupId   = [string]$p.GroupId
            GroupName = Get-PlanGroupName $p
            AddIds    = [string[]]@($p.Add)
            RemoveIds = [string[]]@($p.Remove)
        }
        [void]$grpIds.Add([string]$p.GroupId)
        foreach ($u in @($p.Add))    { if ($u) { [void]$updIds.Add([string]$u) } }
        foreach ($u in @($p.Remove)) { if ($u) { [void]$updIds.Add([string]$u) } }
    }

    $job = [hashtable]::Synchronized(@{
        Id         = [guid]::NewGuid().ToString()
        Kind       = $Kind
        UpdateIds  = [string[]]@($updIds)
        GroupIds   = [string[]]@($grpIds)
        GroupNames = (@($workerPlan | ForEach-Object { $_.GroupName }) -join ', ')
        Total      = $workerPlan.Count
        Done       = 0
        Current    = ''
        Status     = 'Bekliyor'
        Error      = ''
        Simulate   = $Simulate
        StartTime  = Get-Date
        EndTime    = $null
        Cancel     = $false
        Results    = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
        Log        = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
        LogPos     = 0
        Fail       = 0
        PS         = $null
        Handle     = $null
        Cleaned    = $false
    })

    # Yalnızca güncellemeleri kilitle ve etkilenen satırları görsel olarak işaretle
    foreach ($u in $job.UpdateIds) { [void]$script:LockedUpdates.Add($u) }
    $prev = $script:Suspend
    $script:Suspend = $true
    try {
        foreach ($g in $job.GroupIds) {
            if (-not $script:GrpRowById.ContainsKey($g)) { continue }
            $row = $script:GrpRowById[$g]
            # Kullanıcının seçtiği hedef durum kutuda görünür kalır; "bekleyen değişiklik" sayılmasın
            $script:OrigCheck[$g] = ConvertTo-CheckState $row.Cells['gChk'].Value
            Set-GroupRowStyle $row
        }
    } finally {
        $script:Suspend = $prev
    }
    Update-UpdateRows -UpdateIds $job.UpdateIds
    Update-CountLabel (@(Get-SelectedUpdateIds).Count)

    try {
        Start-PooledScript -PoolName 'Membership' -Script $script:MembershipWorker `
            -Arguments @($c.Site, $c.Server, $script:WmiHelperSource, [object[]]$workerPlan, $script:SugLimit, $Simulate) -Job $job
    } catch {
        # Başlatılamadıysa kilitleri geri al
        foreach ($u in $job.UpdateIds) { [void]$script:LockedUpdates.Remove($u) }
        Update-UpdateRows -UpdateIds $job.UpdateIds
        Update-GroupRows -GroupIds $job.GroupIds
        throw
    }

    [void]$script:MembershipJobs.Add($job)
    $simTxt = if ($Simulate) { ' [SİMÜLASYON]' } else { '' }
    Write-Log "Üyelik işi ($Kind) arkaplanda başlatıldı$($simTxt): $($job.GroupIds.Count) SUG, $($job.UpdateIds.Count) güncelleme kilitlendi."
    foreach ($w in $workerPlan) {
        Write-Audit ("Üyelik işi kuyruğa alındı{0} tür={1} iş={2} site={3} SUG='{4}'({5}) eklenecek=[{6}] çıkarılacak=[{7}]" -f $simTxt, $Kind, $job.Id, $c.Site, $w.GroupName, $w.GroupId, ($w.AddIds -join ','), ($w.RemoveIds -join ','))
    }
    $tabJobs.SelectedTab = $tpMem
    Update-MembershipGrid
    $tmrMembership.Start()
}

# Biten bir üyelik işinin sonuçlarını bellekteki $script:Members'a uygular, kilitleri kaldırır,
# geri alma kaydını oluşturur ve yalnızca ilgili güncelleme/SUG satırlarını tazeler.
# Günlük satırlarını (hashtable dizisi) döndürür.
function Complete-MembershipJob {
    param($Job)
    $logs = New-Object System.Collections.ArrayList
    $ok = 0; $fail = 0; $skip = 0
    $undoItems = New-Object System.Collections.ArrayList
    $site = $txtSite.Text.Trim().ToUpper()
    $simTxt = if ($Job.Simulate) { ' [SİMÜLASYON]' } else { '' }

    foreach ($r in @($Job.Results.ToArray())) {
        $gid = [string]$r.GroupId
        if ($r.Skipped) {
            $skip++
            [void]$logs.Add(@{ Level = 'WARN'; Message = "'$($r.GroupName)': iş iptal edildiği için atlandı." })
            Write-Audit ("SugChangeResult{0} iş={1} SUG='{2}'({3}) sonuç=ATLANDI (iptal)" -f $simTxt, $Job.Id, $r.GroupName, $gid)
            continue
        }
        if (-not $Job.Simulate) {
            if ($null -ne $r.PostMember) {
                # Sunucudaki gerçek (doğrulanmış) durumu yansıt
                foreach ($uid in @($r.PostMember.Keys)) {
                    if (-not $script:Updates.ContainsKey($uid)) { continue }
                    if (-not $script:Members.ContainsKey($uid)) { $script:Members[$uid] = New-Object 'System.Collections.Generic.HashSet[string]' }
                    if ($r.PostMember[$uid]) { [void]$script:Members[$uid].Add($gid) } else { [void]$script:Members[$uid].Remove($gid) }
                }
            } else {
                if ($r.AddDone) {
                    foreach ($uid in $r.AddIds) {
                        if (-not $script:Updates.ContainsKey($uid)) { continue }
                        if (-not $script:Members.ContainsKey($uid)) { $script:Members[$uid] = New-Object 'System.Collections.Generic.HashSet[string]' }
                        [void]$script:Members[$uid].Add($gid)
                    }
                }
                if ($r.RemoveDone) {
                    foreach ($uid in $r.RemoveIds) {
                        if ($script:Members.ContainsKey($uid)) { [void]$script:Members[$uid].Remove($gid) }
                    }
                }
            }
            if ($script:GroupById.ContainsKey($gid)) {
                $g = $script:GroupById[$gid]
                if ([int]$r.NewTotal -ge 0) { $g.Total = [int]$r.NewTotal }
                else {
                    $delta = 0
                    if ($r.AddDone)    { $delta += @($r.ActualAdded).Count }
                    if ($r.RemoveDone) { $delta -= @($r.ActualRemoved).Count }
                    $g.Total = [Math]::Max(0, $g.Total + $delta)
                }
            }
            if ($Job.Kind -eq 'Uygulama' -and (@($r.ActualAdded).Count -gt 0 -or @($r.ActualRemoved).Count -gt 0)) {
                [void]$undoItems.Add(@{ GroupId = $gid; GroupName = $r.GroupName; Added = [string[]]@($r.ActualAdded); Removed = [string[]]@($r.ActualRemoved) })
            }
        }
        $skipTxt = if (@($r.SkippedIds).Count -gt 0) { ", $(@($r.SkippedIds).Count) uygun olmadığı için atlandı" } else { '' }
        $verTxt  = if ($r.Verified) { 'doğrulandı' } elseif ($Job.Simulate) { '-' } else { 'DOĞRULANMADI' }
        if ($r.Ok) {
            $ok++
            if (-not $Job.Simulate) {
                $tot = if ($script:GroupById.ContainsKey($gid)) { $script:GroupById[$gid].Total } else { $r.NewTotal }
                [void]$logs.Add(@{ Level = 'OK'; Message = "'$($r.GroupName)': +$($r.Added) eklendi, -$($r.Removed) çıkarıldı$skipTxt (yöntem: $($r.Method), $verTxt, yeni toplam: $tot)." })
            }
        } else {
            $fail++
            [void]$logs.Add(@{ Level = 'ERROR'; Message = "'$($r.GroupName)' güncellenemedi: $($r.Error)" })
        }
        $sonuc = if ($r.Ok) { 'BAŞARILI' } else { "HATA: $($r.Error)" }
        Write-Audit ("SugChangeResult{0} tür={1} iş={2} site={3} SUG='{4}'({5}) yöntem={6} doğrulama={7} eklenen=[{8}] çıkarılan=[{9}] eklenemeyen=[{10}] çıkarılamayan=[{11}] atlanan=[{12}] yeni_toplam={13} sonuç={14}" -f `
            $simTxt, $Job.Kind, $Job.Id, $site, $r.GroupName, $gid, $r.Method, $verTxt, (@($r.ActualAdded) -join ','), (@($r.ActualRemoved) -join ','), (@($r.NotAdded) -join ','), (@($r.NotRemoved) -join ','), (@($r.SkippedIds) -join ','), $r.NewTotal, $sonuc)
    }

    if ($undoItems.Count -gt 0) {
        [void]$script:UndoStack.Add(@{ Time = (Get-Date); JobId = $Job.Id; Items = $undoItems.ToArray() })
        while ($script:UndoStack.Count -gt 20) { $script:UndoStack.RemoveAt(0) }      # en fazla 20 adım
    }

    # Kilitleri kaldır ve yalnızca etkilenen satırları tazele (tam liste yeniden yüklenmez)
    foreach ($u in $Job.UpdateIds) { [void]$script:LockedUpdates.Remove($u) }
    Update-UpdateRows -UpdateIds $Job.UpdateIds
    Update-GroupRows  -GroupIds  $Job.GroupIds

    if ($Job.Status -eq 'Hata' -and $Job.Error) {
        $fail++
        [void]$logs.Add(@{ Level = 'ERROR'; Message = "Üyelik işi hata ile bitti: $($Job.Error)" })
        Write-Audit ("Üyelik işi HATA{0} iş={1} hata={2}" -f $simTxt, $Job.Id, $Job.Error)
    }
    $Job.Fail = $fail
    $level = if ($fail -gt 0) { 'WARN' } else { 'OK' }
    [void]$logs.Add(@{ Level = $level; Message = "$($simTxt.Trim()) Üyelik işi ($($Job.Kind)) bitti ($(Format-Elapsed $Job)): $ok SUG başarılı, $fail hatalı, $skip atlandı. Kilitler kaldırıldı.".Trim() })
    $script:LastMembershipFail = $fail
    Update-UndoButton
    return $logs.ToArray()
}

# "SUG Üyelik İşleri" sekmesindeki listeyi yerinde günceller.
function Update-MembershipGrid {
    $jobs = @($script:MembershipJobs.ToArray())
    $gridMem.SuspendLayout()
    try {
        Remove-StaleJobRows -Grid $gridMem -KeepIds @($jobs | ForEach-Object { $_.Id })
        foreach ($job in ($jobs | Sort-Object StartTime -Descending)) {
            $row = Get-JobRow -Grid $gridMem -JobId $job.Id
            $kind = $job.Kind
            if ($job.Simulate) { $kind += ' (sim.)' }
            $status = $job.Status
            if ($job.Cleaned -and $job.Fail -gt 0 -and $status -eq 'Tamamlandı') { $status = "Tamamlandı ($($job.Fail) hata)" }
            $row.Cells['mStart'].Value   = $job.StartTime.ToString('HH:mm:ss')
            $row.Cells['mKind'].Value    = $kind
            $row.Cells['mGroups'].Value  = $job.GroupNames
            $row.Cells['mProg'].Value    = "$($job.Done) / $($job.Total)"
            $row.Cells['mUpd'].Value     = $job.UpdateIds.Count
            $row.Cells['mCurrent'].Value = $job.Current
            $row.Cells['mStatus'].Value  = $status
            $row.Cells['mSure'].Value    = Format-Elapsed $job
            $row.Cells['mCancel'].Value  = if (-not $job.Cleaned -and -not $job.Cancel) { 'İptal' } else { '' }
            $back = [System.Drawing.Color]::Empty
            if ($job.Cleaned) {
                if ($job.Status -eq 'Hata' -or $job.Fail -gt 0)  { $back = [System.Drawing.Color]::MistyRose }
                elseif ($job.Status -eq 'Tamamlandı')            { $back = [System.Drawing.Color]::Honeydew }
                else                                             { $back = [System.Drawing.Color]::LightYellow }
            }
            $row.DefaultCellStyle.BackColor = $back
        }
    } finally {
        $gridMem.ResumeLayout()
    }
    Update-JobTabTitles
}

# Timer tick: arkaplan üyelik işlerinin günlüklerini aktarır, biten işleri sonlandırıp
# sonuçlarını hedefli olarak ekrana yansıtır (tam liste yenilemesi yapılmaz).
function Update-MembershipJobs {
    if ($script:MemTickBusy) { return }
    $script:MemTickBusy = $true
    $logs = New-Object System.Collections.ArrayList
    try {
        foreach ($job in @($script:MembershipJobs.ToArray())) {
            if ($job.Cleaned) { continue }
            # Tamamlanma durumu günlük okunmadan ÖNCE alınır; böylece işin son günlük satırları da kaçmaz.
            $completed = $job.Handle -and $job.Handle.IsCompleted

            $cnt = $job.Log.Count
            while ($job.LogPos -lt $cnt) {
                [void]$logs.Add($job.Log[$job.LogPos])
                $job.LogPos = [int]$job.LogPos + 1
            }

            # Listeleme sürerken sonuçlar bekletilir; listeleme bitince Update-ListJob zamanlayıcıyı yeniden başlatır.
            if (-not $completed -or $script:ListBusy) { continue }
            Close-PooledScript $job
            foreach ($l in @(Complete-MembershipJob $job)) { [void]$logs.Add($l) }
        }
        Update-MembershipGrid
        $active = @($script:MembershipJobs.ToArray() | Where-Object { -not $_.Cleaned }).Count
        if ($active -eq 0 -or ($script:ListBusy -and $active -eq @($script:MembershipJobs.ToArray() | Where-Object { -not $_.Cleaned -and $_.Handle.IsCompleted }).Count)) {
            $tmrMembership.Stop()
        }
    }
    catch {
        [void]$logs.Add(@{ Level = 'ERROR'; Message = "Üyelik işi takibinde hata: $($_.Exception.Message)" })
    }
    finally {
        $script:MemTickBusy = $false
    }
    foreach ($l in $logs) {
        # Worker'dan gelen 'AUDIT' satırları (SugChangeStart) yalnızca denetim dosyasına yazılır
        if ($l.Level -eq 'AUDIT') { Write-Audit $l.Message } else { Write-Log -Message $l.Message -Level $l.Level -Time $l.Time }
    }
}

# ------------------------------------------------------------------
# Arkaplan indirme (runspace havuzu) - ilerleme takibi
# ------------------------------------------------------------------
# Her indirme "iş"i havuzdaki bir runspace'te çalışır; bu sayede:
#   - UI donmaz (Save-CMSoftwareUpdate ana thread'i bloklamaz),
#   - Birden fazla indirme aynı anda (farklı paket/seçim ile) tetiklenebilir (en fazla 10; fazlası sırada bekler),
#   - Her güncelleme TEK TEK indirilir (SCCM/Save-CMSoftwareUpdate toplu ilerleme bilgisi vermediği için),
#     böylece "X / Y işlendi, Z kaldı" bilgisi elde edilir,
#   - Bir güncellemenin hatası işi DURDURMAZ: hata kaydedilir, sıradaki güncellemeyle devam edilir;
#     iş sonunda "Tamamlandı" / "Kısmen Tamamlandı" / "Hata" durumlarından biriyle biter,
#   - Deployment Package UI'da KİLİTLENMEZ: aynı pakete aynı anda birden fazla indirme başlatılabilir.
#     Aynı pakete yazan Save-CMSoftwareUpdate çağrıları paket başına isimli bir mutex
#     ("Local\IndirmeYapicinatorPkg_<PackageID>") ile güncelleme bazında sıraya girer; böylece iki iş
#     aynı paketin kaynağını/sürümünü aynı anda güncellemez. Bekleme iptal edilebilir.
# İş durumu, ana thread'in Timer ile periyodik okuduğu senkronize (thread-safe) bir hashtable üzerinden
# paylaşılır; arkaplan thread'i doğrudan UI kontrollerine dokunmaz.
$script:DownloadWorker = {
    param(
        [string]  $Site,
        [string]  $Server,
        [string]  $ModulePath,
        [string[]]$UpdateIds,
        [hashtable]$Titles,
        [string]  $PackageName,
        [string]  $PackageId,
        [string[]]$Languages,
        [bool]    $Simulate,
        $Job
    )
    $pushed = $false
    try {
        $Job.Status = 'Çalışıyor'
        foreach ($id in $UpdateIds) {
            if ([string]$id -notmatch '^\d{1,10}$') { throw "Geçersiz güncelleme ID: '$id'" }
        }
        if ($PackageId -notmatch '^[A-Za-z0-9]{1,16}$') { throw "Geçersiz Package ID: '$PackageId'" }
        $mtxName = "Local\IndirmeYapicinatorPkg_$PackageId"
        if (-not $Simulate) {
            Import-Module -Name $ModulePath -ErrorAction Stop
            if (-not (Get-PSDrive -Name $Site -PSProvider CMSite -ErrorAction SilentlyContinue)) {
                New-PSDrive -Name $Site -PSProvider CMSite -Root $Server -Description 'SCCM Site' -ErrorAction Stop | Out-Null
            }
            Push-Location "$($Site):"
            $pushed = $true
        }
        foreach ($id in $UpdateIds) {
            if ($Job.Cancel) { break }
            $title = if ($Titles.ContainsKey($id)) { $Titles[$id] } else { $id }
            $Job.CurrentTitle = $title
            try {
                if ($Simulate) {
                    Start-Sleep -Milliseconds 150
                } else {
                    # Aynı pakete yazımı sıraya al (yalnızca bu güncellemenin indirmesi süresince tutulur)
                    $mtx = New-Object System.Threading.Mutex($false, $mtxName)
                    $owned = $false
                    try {
                        try { $owned = $mtx.WaitOne(0) } catch [System.Threading.AbandonedMutexException] { $owned = $true }
                        if (-not $owned) {
                            $Job.CurrentTitle = "(sırada: paket başka bir indirmede) $title"
                            while (-not $owned -and -not $Job.Cancel) {
                                try { $owned = $mtx.WaitOne(1000) } catch [System.Threading.AbandonedMutexException] { $owned = $true }
                            }
                            if (-not $owned) { break }       # sıra beklenirken iptal edildi
                            $Job.CurrentTitle = $title
                        }
                        Save-CMSoftwareUpdate -SoftwareUpdateId $id -DeploymentPackageName $PackageName -SoftwareUpdateLanguage $Languages -ErrorAction Stop
                    } finally {
                        if ($owned) { $mtx.ReleaseMutex() }
                        $mtx.Dispose()
                    }
                }
                [void]$Job.Succeeded.Add([string]$id)
            } catch {
                [void]$Job.Failed.Add(@{ Id = [string]$id; Title = $Job.CurrentTitle; Error = $_.Exception.Message })
            }
            $Job.Done = [int]$Job.Done + 1
        }
        $nOk = $Job.Succeeded.Count
        $nFail = $Job.Failed.Count
        if ($Job.Cancel)       { $Job.Status = 'İptal Edildi' }
        elseif ($nFail -eq 0)  { $Job.Status = 'Tamamlandı' }
        elseif ($nOk -eq 0)    { $Job.Status = 'Hata'; $Job.Error = "Hiçbir güncelleme indirilemedi ($nFail hata)." }
        else                   { $Job.Status = 'Kısmen Tamamlandı' }
    }
    catch {
        $Job.Status = 'Hata'
        $Job.Error  = $_.Exception.Message
    }
    finally {
        if ($pushed) { try { Pop-Location -ErrorAction SilentlyContinue } catch { } }
        $Job.CurrentTitle = ''
        $Job.EndTime = Get-Date
    }
}

# Yeni bir indirme işini kuyruğa alır ve HEMEN döner (bloklamaz).
function Start-DownloadJob {
    param(
        [string[]]$UpdateIds,
        [string]  $PackageName,
        [string]  $PackageId,
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
        PackageId    = $PackageId
        UpdateIds    = $UpdateIds
        Total        = $UpdateIds.Count
        Done         = 0
        CurrentTitle = ''
        Status       = 'Bekliyor'
        Error        = ''
        Simulate     = $Simulate
        StartTime    = Get-Date
        EndTime      = $null
        Cancel       = $false
        Succeeded    = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
        Failed       = [System.Collections.ArrayList]::Synchronized((New-Object System.Collections.ArrayList))
        PS           = $null
        Handle       = $null
        Cleaned      = $false
    })

    Start-PooledScript -PoolName 'Download' -Script $script:DownloadWorker `
        -Arguments @($c.Site, $c.Server, $modulePath, [string[]]$UpdateIds, $titles, $PackageName, $PackageId, [string[]]$Languages, $Simulate) -Job $job

    [void]$script:DownloadJobs.Add($job)
    $simTxt = if ($Simulate) { ' [SİMÜLASYON]' } else { '' }
    Write-Log "'$PackageName' için arkaplanda indirme kuyruğa alındı$($simTxt): $($UpdateIds.Count) güncelleme, diller: $($Languages -join ', ')."
    Write-Audit ("İndirme başladı{0} iş={1} site={2} paket='{3}'({6}) diller={4} güncellemeler=[{5}]" -f $simTxt, $job.Id, $c.Site, $PackageName, ($Languages -join ','), ($UpdateIds -join ','), $PackageId)
    $tabJobs.SelectedTab = $tpDl
    $tmrProgress.Start()
}

# Biten bir indirme işini günlüğe/denetim kaydına yazar ve indirilen güncellemelerin
# "İçerik" hücresini hedefli olarak "Var" yapar.
function Complete-DownloadJob {
    param($Job)
    $simTxt = if ($Job.Simulate) { ' [SİMÜLASYON]' } else { '' }
    $okIds = @($Job.Succeeded.ToArray())
    $fails = @($Job.Failed.ToArray())
    foreach ($f in $fails) {
        Write-Log "'$($Job.Package)': $($f.Title) indirilemedi: $($f.Error)" 'ERROR'
    }
    switch ($Job.Status) {
        'Tamamlandı'        { Write-Log "'$($Job.Package)'$($simTxt): indirme tamamlandı ($($okIds.Count)/$($Job.Total) güncelleme, $(Format-Elapsed $Job))." 'OK' }
        'Kısmen Tamamlandı' { Write-Log "'$($Job.Package)'$($simTxt): indirme kısmen tamamlandı ($($okIds.Count) başarılı, $($fails.Count) hatalı / $($Job.Total))." 'WARN' }
        'İptal Edildi'      { Write-Log "'$($Job.Package)'$($simTxt): indirme iptal edildi ($($Job.Done)/$($Job.Total) işlenmişti)." 'WARN' }
        default             { Write-Log "'$($Job.Package)'$($simTxt): indirme hata ile bitti: $($Job.Error)" 'ERROR' }
    }
    Write-Audit ("İndirme bitti{0} iş={1} paket='{2}' durum={3} başarılı=[{4}] hatalı=[{5}] hata={6}" -f $simTxt, $Job.Id, $Job.Package, $Job.Status, ($okIds -join ','), (($fails | ForEach-Object { $_.Id }) -join ','), $Job.Error)

    if (-not $Job.Simulate -and $okIds.Count -gt 0) {
        foreach ($id in $okIds) { if ($script:Updates.ContainsKey($id)) { $script:Updates[$id].ContentProvisioned = $true } }
        Update-UpdateRows -UpdateIds $okIds
    }
}

# Timer tick: aktif işlerin durumunu grid'e (yerinde) yansıtır; biten işleri sonlandırır.
function Update-DownloadGrid {
    # Önce biten işleri sonlandır (EndInvoke/Dispose) ve sonuçlarını işle
    foreach ($job in @($script:DownloadJobs.ToArray())) {
        if (-not $job.Cleaned -and $job.Handle -and $job.Handle.IsCompleted) {
            Close-PooledScript $job
            Complete-DownloadJob $job
        }
    }

    $jobs = @($script:DownloadJobs.ToArray())
    $gridDl.SuspendLayout()
    try {
        Remove-StaleJobRows -Grid $gridDl -KeepIds @($jobs | ForEach-Object { $_.Id })
        foreach ($job in ($jobs | Sort-Object StartTime -Descending)) {
            $row = Get-JobRow -Grid $gridDl -JobId $job.Id
            $nFail = $job.Failed.Count
            $row.Cells['dPkg'].Value = $job.Package
            $row.Cells['dProg'].Value = "$($job.Done) / $($job.Total)"
            $row.Cells['dRemain'].Value = [Math]::Max(0, [int]$job.Total - [int]$job.Done)
            $row.Cells['dFail'].Value = if ($nFail -gt 0) { $nFail } else { '' }
            $pct = if ($job.Total -gt 0) { [Math]::Round((100.0 * $job.Done / $job.Total)) } else { 0 }
            $row.Cells['dPct'].Value = "%$pct"
            $row.Cells['dCurrent'].Value = $job.CurrentTitle
            $statusTxt = $job.Status
            if ($job.Simulate -and $statusTxt -ne 'Hata') { $statusTxt = "$statusTxt (sim.)" }
            $row.Cells['dStatus'].Value = $statusTxt
            $row.Cells['dSure'].Value = Format-Elapsed $job
            $row.Cells['dCancel'].Value = if ($job.Status -in @('Bekliyor', 'Çalışıyor') -and -not $job.Cancel) { 'İptal' } else { '' }
            $back = [System.Drawing.Color]::Empty
            switch ($job.Status) {
                'Hata'              { $back = [System.Drawing.Color]::MistyRose }
                'Kısmen Tamamlandı' { $back = [System.Drawing.Color]::LightYellow }
                'İptal Edildi'      { $back = [System.Drawing.Color]::LightYellow }
                'Tamamlandı'        { $back = [System.Drawing.Color]::Honeydew }
            }
            $row.DefaultCellStyle.BackColor = $back
        }
    } finally {
        $gridDl.ResumeLayout()
    }
    Update-JobTabTitles

    # Aktif iş kalmadıysa timer'ı durdur (gereksiz yere UI thread'ini meşgul etmesin).
    $anyActive = @($script:DownloadJobs.ToArray() | Where-Object { -not $_.Cleaned }).Count -gt 0
    if (-not $anyActive) { $tmrProgress.Stop() }
}

# Seçili güncellemelerin içeriğini VAR OLAN bir Deployment Package'a indirir.
# Save-CMSoftwareUpdate -Location verilmeden çağrılır: bu durumda içerik SCCM tarafından
# doğrudan İnternet'ten (Microsoft Update) indirilir; hiçbir yeni paket oluşturulmaz.
function Invoke-Download {
    $sel = @(Get-SelectedUpdateIds)
    if ($sel.Count -eq 0) { Show-Message 'Önce soldaki listeden en az bir güncelleme seçin.' 'Uyarı' 'Warning'; return }

    $pkgId = Get-SelectedPackageId
    if (-not $pkgId) { Show-Message "Var olan bir Deployment Package seçin (kombodan seçin ya da adını birebir yazın).`n`nGerekirse önce ""Paketleri Yenile"" düğmesine basın." 'Uyarı' 'Warning'; return }
    $pkgName = $script:PackageById[$pkgId].Name

    $langs = @(Get-SelectedLanguages)
    if ($langs.Count -eq 0) { Show-Message 'En az bir dil seçin (English / Turkish).' 'Uyarı' 'Warning'; return }

    $simulate = $chkSim.Checked
    $already = @($sel | Where-Object { $script:Updates.ContainsKey($_) -and $script:Updates[$_].ContentProvisioned }).Count
    $head = if ($simulate) { "SİMÜLASYON - hiçbir indirme başlatılmayacak.`n`n" } else { "" }
    $note = if ($already -gt 0) { "`n(Seçimden $already güncellemenin içeriği zaten indirilmiş görünüyor; SCCM bunları yeniden indirmez/atlar.)" } else { '' }
    $msg = "{0}{1} güncelleme, '{2}' Deployment Package'ına İNTERNETTEN (Microsoft Update) indirilecek.`nDiller: {3}{4}`n`nİndirme ARKAPLANDA çalışacak; pencere kullanılabilir kalacak ve ""Arkaplan İşleri > İndirmeler"" sekmesinden ilerleme takip edilebilir. Devam edilsin mi?" -f $head, $sel.Count, $pkgName, ($langs -join ', '), $note
    if ([System.Windows.Forms.MessageBox]::Show($msg, 'İçerik indir', 'YesNo', 'Question') -ne 'Yes') { return }

    try {
        Start-DownloadJob -UpdateIds $sel -PackageName $pkgName -PackageId $pkgId -Languages $langs -Simulate $simulate
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
$btnUndo.Add_Click({ Invoke-UndoMembership })
$btnDownload.Add_Click({ Invoke-Download })
$btnPkgRefresh.Add_Click({
    try {
        Start-ListJob -Mode 'Packages'
    } catch {
        Write-Log $_.Exception.Message 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    }
})
$btnClose.Add_Click({ $form.Close() })
$btnSelAll.Add_Click({ $gridUpd.SelectAll() })
$btnSelNone.Add_Click({ $gridUpd.ClearSelection() })
$btnLogSave.Add_Click({ Save-LogView })
$cmbLogLevel.Add_SelectedIndexChanged({ if ($script:LogReady) { Update-LogView } })

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
$gridDl.Add_DataError({ param($sender, $e) $e.ThrowException = $false })
$gridMem.Add_DataError({ param($sender, $e) $e.ThrowException = $false })

# "İptal" düğmesine tıklanan satırın işine Cancel=true işaretler; iş, o an işlediği
# öğe (indirmede güncelleme, üyelikte SUG) bittikten sonra durur (sert kesme yapılmaz).
$gridDl.Add_CellContentClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $gridDl.Columns[$e.ColumnIndex].Name -ne 'dCancel') { return }
    $jobId = [string]$gridDl.Rows[$e.RowIndex].Tag
    $job = @($script:DownloadJobs.ToArray()) | Where-Object { $_.Id -eq $jobId } | Select-Object -First 1
    if ($job -and $job.Status -in @('Bekliyor', 'Çalışıyor') -and -not $job.Cancel) {
        $job.Cancel = $true
        Write-Log "'$($job.Package)' indirmesi iptal ediliyor (devam eden öğe bittikten sonra duracak)..." 'WARN'
        Write-Audit "İndirme iptal istendi iş=$($job.Id) paket='$($job.Package)'"
    }
})
$gridMem.Add_CellContentClick({
    param($sender, $e)
    if ($e.RowIndex -lt 0 -or $gridMem.Columns[$e.ColumnIndex].Name -ne 'mCancel') { return }
    $jobId = [string]$gridMem.Rows[$e.RowIndex].Tag
    $job = @($script:MembershipJobs.ToArray()) | Where-Object { $_.Id -eq $jobId } | Select-Object -First 1
    if ($job -and -not $job.Cleaned -and -not $job.Cancel) {
        $job.Cancel = $true
        Write-Log "Üyelik işi iptal ediliyor (yazılmakta olan SUG bittikten sonra kalan SUG'lar atlanacak): $($job.GroupNames)" 'WARN'
        Write-Audit "Üyelik işi iptal istendi iş=$($job.Id)"
        Update-MembershipGrid
    }
})

# Biten işleri her iki listeden de temizler (çalışanlar kalır).
$btnJobsClear.Add_Click({
    foreach ($list in @($script:DownloadJobs, $script:MembershipJobs)) {
        $keep = @($list.ToArray() | Where-Object { -not $_.Cleaned })
        $list.Clear()
        foreach ($job in $keep) { [void]$list.Add($job) }
    }
    Update-DownloadGrid
    Update-MembershipGrid
})

# Pencere kapatılırken: devam eden üyelik işleri varsa kullanıcıya sorulur (SUG yazımı yarım kalmasın);
# kapatılacaksa üyelik ve indirme işleri iptal edilip kaynaklar serbest bırakılır.
$form.Add_FormClosing({
    param($sender, $e)
    $memActive = @($script:MembershipJobs.ToArray() | Where-Object { -not $_.Cleaned -and -not ($_.Handle -and $_.Handle.IsCompleted) })
    if ($memActive.Count -gt 0) {
        $q = "$($memActive.Count) SUG üyelik işi hâlâ arkaplanda çalışıyor.`n`nŞimdi kapatırsanız, o an yazılmakta olan SUG tamamlanır ancak sıradaki SUG'lar işlenmez.`n`nYine de kapatılsın mı?"
        if ([System.Windows.Forms.MessageBox]::Show($q, 'Üyelik işleri sürüyor', 'YesNo', 'Warning') -ne 'Yes') {
            $e.Cancel = $true
            return
        }
        $tmrMembership.Stop()
        foreach ($job in $memActive) { $job.Cancel = $true }
        # Yazılmakta olan SUG'un Put() çağrısının bitmesi için kısa süre beklenir (sert kesme yapılmaz)
        $deadline = (Get-Date).AddSeconds(15)
        while ((Get-Date) -lt $deadline -and (@($memActive | Where-Object { -not $_.Handle.IsCompleted }).Count -gt 0)) {
            Start-Sleep -Milliseconds 100
        }
        foreach ($job in $memActive) {
            Write-Audit "Pencere kapatıldı: üyelik işi iptal edildi iş=$($job.Id) işlenen=$($job.Done)/$($job.Total) SUG"
            try { if ($job.PS) { $job.PS.Stop(); $job.PS.Dispose() } } catch { }
        }
    }

    $active = @($script:DownloadJobs.ToArray() | Where-Object { -not $_.Cleaned -and $_.Status -in @('Bekliyor', 'Çalışıyor') })
    if ($active.Count -gt 0) {
        foreach ($job in $active) { $job.Cancel = $true }
        $deadline = (Get-Date).AddSeconds(5)
        while ((Get-Date) -lt $deadline -and (@($active | Where-Object { -not $_.Handle.IsCompleted }).Count -gt 0)) {
            Start-Sleep -Milliseconds 100
        }
        foreach ($job in $active) {
            Write-Audit "Pencere kapatıldı: indirme iptal edildi iş=$($job.Id) paket='$($job.Package)' işlenen=$($job.Done)/$($job.Total)"
            try { if ($job.PS) { $job.PS.Stop(); $job.PS.Dispose() } } catch { }
        }
    }
    if ($script:ListJob) { try { $script:ListJob.PS.Stop(); $script:ListJob.PS.Dispose() } catch { } }
})

$form.Add_FormClosed({
    Save-Settings
    foreach ($p in @($script:Pools.Values)) { try { $p.Close(); $p.Dispose() } catch { } }
    Write-LogFile -Level 'INFO' -Message 'Oturum kapandı.'
    [System.Windows.Forms.Application]::RemoveMessageFilter($script:WheelFilter)
})

# ------------------------------------------------------------------
# Açılış
# ------------------------------------------------------------------
Initialize-AppDirs
$script:LogReady = $true
Write-LogFile -Level 'INFO' -Message "Oturum başladı. Betik: $PSCommandPath"
Write-Log 'Hazır. Filtreyi kontrol edip "Listele" düğmesine basın. Varsayılan olarak SİMÜLASYON açıktır.'
if ($script:FileLogOk) { Write-Log "Kalıcı günlük ve denetim kaydı: $($script:LogDir)" }
else { Write-Log "Günlük klasörü oluşturulamadı, kalıcı kayıt tutulmayacak: $($script:FileLogError)" 'WARN' }
Import-ConnectionDefaults
Import-Settings
$lblCount.Text = 'Henüz liste yüklenmedi.'
Update-UndoButton
Update-JobTabTitles

# Başlangıç boyutu: kayıtlı boyut (yoksa tasarım boyutu) ekrana sığmıyorsa çalışma alanına göre kırpılır.
# Yerleşim asgari alanın altına inerse form kaydırma çubukları gösterir; tümünü sığdırmaya çalışmayız.
$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$wantW = $script:DesignW; $wantH = $script:DesignH
if ($script:SavedClientSize) { $wantW = [Math]::Max(640, $script:SavedClientSize.Width); $wantH = [Math]::Max(420, $script:SavedClientSize.Height) }
$capW = [Math]::Min($wantW, $wa.Width - 20)
$capH = [Math]::Min($wantH, $wa.Height - 60)    # başlık çubuğu + kenarlık payı
$form.ClientSize = New-Object System.Drawing.Size($capW, $capH)

$script:WheelFilter = New-Object FormWheelScrollFilter($form)
[System.Windows.Forms.Application]::AddMessageFilter($script:WheelFilter)

[void]$form.ShowDialog()
