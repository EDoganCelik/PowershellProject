<#
.SYNOPSIS
    DeployControl - Bir Software Update Group'u (SUG), seçilen cihaz koleksiyonlarına "Required" olarak deploy eder (UI ile).

.DESCRIPTION
    - create-sccm-update-object.ps1 tarafından oluşturulan state.json dosyasını OKUR (yazmaz).
      Aranan konumlar: Desktop\update-state\state.json ve %LOCALAPPDATA%\update-state\state.json
      (ikisi de varsa en yenisi kullanılır; "Gözat..." ile başka bir dosya da seçilebilir).
    - Site kodu ve SMS Provider YALNIZCA state.json'dan alınır; betikte sabit (hard-coded) ortam bilgisi yoktur.
    - Software Update Group listesi ORTAMDAN (SMS_AuthorizationList) okunur: form açılınca otomatik yüklenir,
      "Ortamdaki SUG'ları Yükle" ile yenilenir. state.json'daki son kayıt (LastRun) SUG'ı listenin başında ★ ile
      ve varsayılan olarak SEÇİLİ gelir; isterseniz ortamdaki herhangi bir SUG'ı seçebilirsiniz.
    - Koleksiyonlar Collection ID'lerine göre listelenir:
          * ID'leri kutuya yapıştırın (boşluk / virgül / noktalı virgül / satır sonu ile ayrılmış) veya dosyadan alın.
          * "Tümünü Çek" ortamdaki tüm cihaz koleksiyonlarını çeker.
      Listeye alınan koleksiyonlar CACHE'lenir (%LOCALAPPDATA%\DeployControl\collection-cache-<SITE>.json);
      bir sonraki açılışta cache'den hemen gelir, SCCM'e tekrar sormaya gerek kalmaz.
    - Cache'deki her koleksiyon için deploy bilgisi gösterilir: seçili SUG bu koleksiyona deploy edilmiş mi, hangi tarihte,
      hangi şekilde (Required/Available, deadline, kullanıcı deneyimi, restart davranışı, indirme ayarı...).
      Alttaki tabloda seçili koleksiyonun TÜM deploy geçmişi (diğer SUG'lar dahil) görünür.
      "Deploy Bilgisini Yenile" bu bilgiyi SCCM'den (SMS_UpdateGroupAssignment) tek sorguyla tazeler.
    - Deploy ayarları varsayılan olarak Program.cs (SccmSugDeploy) ile AYNIDIR; sağdaki sekmelerden değiştirilebilir:
          Required, ASAP başlangıç + ASAP deadline, tüm state mesajları, Software Center ve bildirimler gizli,
          kurulum bakım penceresi dışında, restart bastırma (workstation + server), Wake-on-LAN,
          DPLocality = 524368 (yerel DP + uzak DP + ölçümlü bağlantı), deployment adı: "<SUG Adı> to <Koleksiyon Adı>" (Program.cs'teki "<SUG Adı> <CollectionID>" yerine).
    - "Simülasyon" kutusu işaretliyken hiçbir deployment OLUŞTURULMAZ; yalnızca ne yapılacağı günlüğe yazılır.
    - Günlükler hem ekrandaki kutuya hem Program.cs ile aynı biçimde SABİT olarak
      %USERPROFILE%\Desktop\update-state\yyyy-MM-dd-DeployControl.log dosyasına yazılır;
      klasör yoksa oluşturulur. Günlük klasörü state.json'a BAĞLI DEĞİLDİR.

.NOTES
    SMS Provider'a WMI (root\sms\site_XXX) ile oturum açmış kullanıcının kimliğiyle bağlanır (Program.cs ile aynı: Impersonate / PacketPrivacy).
    Deployment oluşturmak için SCCM'de ilgili koleksiyon üzerinde "Deploy Software Updates" yetkisi gerekir.
    Aynı adlı bir deployment zaten varsa o koleksiyon ATLANIR (Program.cs ile aynı davranış).
#>

#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Management
[System.Windows.Forms.Application]::EnableVisualStyles()

# ------------------------------------------------------------------
# AYARLAR
# Site kodu / SMS Provider burada TUTULMAZ: yalnızca state.json'dan okunur.
# ------------------------------------------------------------------
$script:LogFolderPath     = Join-Path $env:USERPROFILE 'Desktop\update-state'   # %USERPROFILE%\Desktop\update-state (yoksa oluşturulur)
$script:LogWriteWarned    = $false
$script:AppName           = 'DeployControl'          # yyyy-MM-dd-<AppName>.log
if ($env:LOCALAPPDATA) { $script:CacheFolder = Join-Path $env:LOCALAPPDATA 'DeployControl' }
else                   { $script:CacheFolder = Join-Path $env:USERPROFILE 'DeployControl' }

# ------------------------------------------------------------------
# Genel durum
# ------------------------------------------------------------------
$script:StateFile    = $null
$script:StateEntries = @()      # state.json kayıtları (combobox ile aynı sırada)
$script:Groups       = @()      # yalnızca deploy geçmişindeki SUG adlarını göstermek için (liste seçimi için KULLANILMAZ)
$script:GroupById    = @{}      # SUG CI_ID (string) -> SUG
$script:SugItems     = @()      # combobox'taki SUG listesi (state SUG'ları başta)
$script:Cols         = New-Object 'System.Collections.Generic.Dictionary[string,object]' -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase)
$script:Checked      = @{}      # CollectionID -> deploy için işaretli mi
$script:AssignmentNames = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([System.StringComparer]::OrdinalIgnoreCase)
$script:CacheSite    = ''       # cache'i yüklenen site kodu
$script:Suspend      = $false
$script:Scope        = $null
$script:ScopeKey     = ''

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
    return ('{0}  |  SUG: {1} (CI_ID {2})  |  PKG: {3}' -f $E.Timestamp, $E.SugName, $E.SugId, $E.PkgName)
}

# ------------------------------------------------------------------
# Küçük yardımcılar
# ------------------------------------------------------------------
function Format-Dt {
    param($Value)
    if (-not $Value) { return '' }
    try { return ([datetime]$Value).ToString('yyyy-MM-dd HH:mm') } catch { return [string]$Value }
}

function ConvertFrom-DmtfSafe {
    param($Value)
    if (-not $Value) { return $null }
    try { return ([System.Management.ManagementDateTimeConverter]::ToDateTime([string]$Value)).ToString('s') }
    catch { return $null }
}

function Get-Col {
    param([string]$Id)
    $c = $null
    if ($Id -and $script:Cols.TryGetValue($Id, [ref]$c)) { return $c }
    return $null
}

# Deploy şeklini tek satırda özetler (Required/Available, deadline, kullanıcı deneyimi, restart, indirme...).
function Format-DeployShape {
    param($D)
    if (-not $D) { return '' }
    $p = @()
    $t = [string]$D.Type
    if ($t -eq 'Required' -and $D.Deadline) { $p += ('Required (deadline {0})' -f (Format-Dt $D.Deadline)) }
    elseif ($t) { $p += $t }
    if ($D.UserUI -eq $true) {
        if ($D.NotifyUser -eq $true) { $p += 'SC: tüm bildirimler' } else { $p += 'SC: yalnız restart bildirimi' }
    } else { $p += 'SC ve bildirimler gizli' }
    $sr = 0
    try { $sr = [int]$D.SuppressReboot } catch { }
    $rs = @()
    if ($sr -band 1) { $rs += 'WS' }
    if ($sr -band 2) { $rs += 'SRV' }
    if ($rs.Count -gt 0) { $p += ('Restart bastırılır: ' + ($rs -join '+')) } else { $p += 'Restart bastırılmaz' }
    if ($D.OverrideSW -eq $true)    { $p += 'Kurulum MW dışı' }
    if ($D.RebootOutside -eq $true) { $p += 'Restart MW dışı' }
    if ($D.Wol -eq $true)           { $p += 'WoL' }
    $dp = 0
    try { $dp = [int64]$D.DpLocality } catch { }
    $dl = @()
    if ($dp -band 16)     { $dl += 'yerel DP' }
    if ($dp -band 64)     { $dl += 'uzak DP' }
    if ($dp -band 524288) { $dl += 'ölçümlü' }
    if ($dp -band 262144) { $dl += 'WU/MU' }
    if ($dl.Count -gt 0) { $p += ('İndirme: ' + ($dl -join '+')) }
    if ($D.Enabled -eq $false) { $p += 'DEVRE DIŞI' }
    return ($p -join ' | ')
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

function New-UiCheck {
    param($Parent, [string]$Text, [int]$X, [int]$Y, [int]$W = 370, [bool]$Checked = $false)
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $Text
    $c.Location = New-Object System.Drawing.Point($X, $Y)
    $c.Size = New-Object System.Drawing.Size($W, 22)
    $c.Checked = $Checked
    $Parent.Controls.Add($c)
    return $c
}

function New-UiCombo {
    param($Parent, [int]$X, [int]$Y, [int]$W, [string[]]$Items, [int]$Index = 0)
    $c = New-Object System.Windows.Forms.ComboBox
    $c.DropDownStyle = 'DropDownList'
    $c.Location = New-Object System.Drawing.Point($X, $Y)
    $c.Size = New-Object System.Drawing.Size($W, 24)
    foreach ($i in $Items) { [void]$c.Items.Add($i) }
    if ($Items.Count -gt 0) { $c.SelectedIndex = $Index }
    $Parent.Controls.Add($c)
    return $c
}

function New-UiDate {
    param($Parent, [int]$X, [int]$Y, [int]$W)
    $d = New-Object System.Windows.Forms.DateTimePicker
    $d.Format = 'Custom'
    $d.CustomFormat = 'yyyy-MM-dd HH:mm'
    $d.Location = New-Object System.Drawing.Point($X, $Y)
    $d.Size = New-Object System.Drawing.Size($W, 23)
    $Parent.Controls.Add($d)
    return $d
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
    $g.AutoSizeColumnsMode = 'None'          # sütunlar sabit genişlikte; toplam sığmazsa X scroll çıkar
    $g.ScrollBars = 'Both'
    $g.AllowUserToResizeColumns = $true
    $Parent.Controls.Add($g)
    return $g
}

function Add-UiTextColumn {
    param($Grid, [string]$Name, [string]$Header, [int]$Width)
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = $Name
    $c.HeaderText = $Header
    $c.ReadOnly = $true
    $c.Width = $Width
    [void]$Grid.Columns.Add($c)
}

# Çift tıklayınca sütunu içeriğine göre genişletir, tekrar çift tıklayınca eski genişliğe döndürür.
# (Başlığa veya herhangi bir hücreye çift tıklanabilir; onay kutusu sütunları hariç.)
function Enable-ColumnDoubleClickExpand {
    param($Grid)
    $toggle = {
        param($sender, $e)
        if ($e.ColumnIndex -lt 0) { return }
        $col = $sender.Columns[$e.ColumnIndex]
        if ($col -is [System.Windows.Forms.DataGridViewCheckBoxColumn]) { return }
        if ($col.Tag -is [hashtable] -and $col.Tag.Expanded) {
            $col.Width = [int]$col.Tag.Orig
            $col.Tag = $null
        } else {
            $orig = $col.Width
            $sender.AutoResizeColumn($e.ColumnIndex, [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells)
            $col.Tag = @{ Orig = $orig; Expanded = $true }
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
# Yerleşim 1280x940 olarak tasarlandı. Formun gösterilebilir ASGARİ yerleşim alanı 1100x860'tır:
# pencere (ekran) bundan küçükse nesneler sıkışmaz, form kaydırma çubukları çıkar (fare tekerleği ile de kaydırılır).
# Asgari değerin üstünde kalan alanda nesneler (anchor'lar sayesinde) genişleyip uzamaya devam eder.
# ------------------------------------------------------------------
$script:LayoutMinW = 1100
$script:LayoutMinH = 860
$form = New-Object System.Windows.Forms.Form
$form.Text = 'SCCM - SUG Deploy Kontrol (DeployControl)'
$form.ClientSize = New-Object System.Drawing.Size(1280, 940)        # kontroller bu tasarım boyutuna göre yerleştirilir
$form.AutoScroll = $true
$form.AutoScrollMinSize = New-Object System.Drawing.Size($script:LayoutMinW, $script:LayoutMinH)
$form.MinimumSize = New-Object System.Drawing.Size(640, 420)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$tip = New-Object System.Windows.Forms.ToolTip

# --- Bağlantı, State ve SUG ---
$gTop = New-UiGroup $form 'Bağlantı, State ve Deploy SUG' 10 8 1260 128
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
$cmbState.Size = New-Object System.Drawing.Size(1065, 24)
$cmbState.Anchor = 'Top,Left,Right'
$gTop.Controls.Add($cmbState)

[void](New-UiLabel $gTop 'Deploy SUG:' 12 96 72)
$cmbSug = New-Object System.Windows.Forms.ComboBox
$cmbSug.DropDownStyle = 'DropDownList'
$cmbSug.Location = New-Object System.Drawing.Point(85, 92)
$cmbSug.Size = New-Object System.Drawing.Size(700, 24)
$cmbSug.Anchor = 'Top,Left,Right'
$gTop.Controls.Add($cmbSug)
$btnLoadSug = New-UiButton $gTop 'Ortamdaki SUG''ları Yükle' 792 90 200 28
$btnLoadSug.Anchor = 'Top,Right'
$btnSugReset = New-UiButton $gTop 'State SUG''a Dön' 998 90 150 28
$btnSugReset.Anchor = 'Top,Right'
$tip.SetToolTip($cmbSug, '★ işaretli SUG state.json''daki son kayıttır ve varsayılan olarak seçilidir. Ortamdaki herhangi bir SUG''ı seçebilirsiniz.')
$tip.SetToolTip($btnLoadSug, 'Ortamdaki tüm Software Update Group''ları (SMS_AuthorizationList) okuyup listeye ekler.')

# --- Koleksiyonlar ---
$gCol = New-UiGroup $form 'Koleksiyonlar (cache''li liste)' 10 144 820 400
$gCol.Anchor = 'Top,Bottom,Left,Right'
[void](New-UiLabel $gCol 'Collection ID(ler):' 12 25 105)
$txtIds = New-UiTextBox $gCol 118 22 380
$txtIds.Anchor = 'Top,Left,Right'
$btnAddIds = New-UiButton $gCol 'Listeye Ekle' 505 20 95 27
$btnAddIds.Anchor = 'Top,Right'
$btnFile = New-UiButton $gCol 'Dosyadan Al...' 605 20 100 27
$btnFile.Anchor = 'Top,Right'
$btnAll = New-UiButton $gCol 'Tümünü Çek' 710 20 100 27
$btnAll.Anchor = 'Top,Right'
$tip.SetToolTip($txtIds, 'Collection ID''leri boşluk, virgül, noktalı virgül veya satır sonu ile ayırarak yapıştırın. Örn: EDC00012, EDC00034')
$tip.SetToolTip($btnAll, 'Ortamdaki TÜM cihaz koleksiyonlarını çekip cache''e alır.')

[void](New-UiLabel $gCol 'Ara:' 12 57 35)
$txtColFilter = New-UiTextBox $gCol 48 54 200
$chkOnlyNew = New-UiCheck $gCol 'Bu SUG deploy edilmemişler' 258 54 190
$btnRefresh = New-UiButton $gCol 'Deploy Bilgisini Yenile' 555 51 155 27
$btnRefresh.Anchor = 'Top,Right'
$btnRemoveCache = New-UiButton $gCol 'Cache''den Çıkar' 715 51 95 27
$btnRemoveCache.Anchor = 'Top,Right'
$tip.SetToolTip($btnRefresh, 'Cache''deki tüm koleksiyonların deploy bilgisini SCCM''den (SMS_UpdateGroupAssignment) tek sorguyla yeniler.')
$tip.SetToolTip($btnRemoveCache, 'Seçili satırları yalnızca cache''den çıkarır; SCCM''de hiçbir şey silinmez.')

$gridCol = New-UiGrid $gCol 10 84 800 276
$gridCol.ReadOnly = $false
$colChk = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colChk.Name = 'cChk'
$colChk.HeaderText = 'Deploy'
$colChk.Width = 55
[void]$gridCol.Columns.Add($colChk)
Add-UiTextColumn $gridCol 'cId'      'Collection ID'            90
Add-UiTextColumn $gridCol 'cName'    'Koleksiyon Adı'           230
Add-UiTextColumn $gridCol 'cMem'     'Üye'                      55
Add-UiTextColumn $gridCol 'cType'    'Tür'                      65
Add-UiTextColumn $gridCol 'cSug'     'Bu SUG deploy?'           100
Add-UiTextColumn $gridCol 'cDate'    'Deploy Tarihi (bu SUG)'   135
Add-UiTextColumn $gridCol 'cShape'   'Deploy Şekli (bu SUG)'    420
Add-UiTextColumn $gridCol 'cTotal'   'Toplam Deploy'            90
Add-UiTextColumn $gridCol 'cLast'    'Son Deploy (tüm SUG''lar)' 280
Add-UiTextColumn $gridCol 'cCached'  'Cache Zamanı'             125
Enable-ColumnDoubleClickExpand $gridCol

$btnMarkAll = New-UiButton $gCol 'Görünenleri İşaretle' 10 366 150 26
$btnMarkAll.Anchor = 'Bottom,Left'
$btnUnmark = New-UiButton $gCol 'İşaretleri Kaldır' 165 366 125 26
$btnUnmark.Anchor = 'Bottom,Left'
$lblColCount = New-UiLabel $gCol '' 300 371 505
$lblColCount.Anchor = 'Bottom,Left,Right'

# --- Seçili koleksiyonun deploy geçmişi ---
$gDet = New-UiGroup $form 'Seçili koleksiyonun deploy geçmişi' 10 552 820 170
$gDet.Anchor = 'Bottom,Left,Right'
$gridDet = New-UiGrid $gDet 10 20 800 140
$gridDet.ReadOnly = $true
Add-UiTextColumn $gridDet 'dName'    'Deployment Adı'  230
Add-UiTextColumn $gridDet 'dSug'     'SUG'             200
Add-UiTextColumn $gridDet 'dCreated' 'Oluşturulma'     120
Add-UiTextColumn $gridDet 'dStart'   'Başlangıç'       120
Add-UiTextColumn $gridDet 'dDead'    'Deadline'        120
Add-UiTextColumn $gridDet 'dType'    'Tür'              80
Add-UiTextColumn $gridDet 'dShape'   'Deploy Şekli'    430
Add-UiTextColumn $gridDet 'dEnabled' 'Etkin'            50
Add-UiTextColumn $gridDet 'dId'      'AssignmentID'     95
Enable-ColumnDoubleClickExpand $gridDet

# --- Deploy ayarları (Program.cs varsayılanları) ---
$gSet = New-UiGroup $form 'Deploy Ayarları (varsayılanlar Program.cs ile aynı)' 840 144 430 400
$gSet.Anchor = 'Top,Right'
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(8, 20)
$tabs.Size = New-Object System.Drawing.Size(414, 340)
$gSet.Controls.Add($tabs)
$tp1 = New-Object System.Windows.Forms.TabPage; $tp1.Text = 'Zamanlama'; $tp1.UseVisualStyleBackColor = $true
$tp2 = New-Object System.Windows.Forms.TabPage; $tp2.Text = 'Kullanıcı && Restart'; $tp2.UseVisualStyleBackColor = $true
$tp3 = New-Object System.Windows.Forms.TabPage; $tp3.Text = 'İndirme && Diğer'; $tp3.UseVisualStyleBackColor = $true
$tabs.TabPages.Add($tp1); $tabs.TabPages.Add($tp2); $tabs.TabPages.Add($tp3)

# Sekme 1: ad, tür, zamanlama, mesaj düzeyi
[void](New-UiLabel $tp1 'Deployment adı:' 8 13 105)
$txtNameFmt = New-UiTextBox $tp1 118 10 275
[void](New-UiLabel $tp1 'Açıklama:' 8 41 105)
$txtDesc = New-UiTextBox $tp1 118 38 275
[void](New-UiLabel $tp1 'Deployment türü:' 8 69 105)
$cmbType = New-UiCombo $tp1 118 66 200 @('Required (varsayılan)', 'Available (deadline yok)')
[void](New-UiLabel $tp1 'Kullanılabilir zaman:' 8 99 108)
$cmbStart = New-UiCombo $tp1 118 96 110 @('ASAP (şimdi)', 'Özel tarih')
$dtStart = New-UiDate $tp1 234 96 160
[void](New-UiLabel $tp1 'Kurulum deadline:' 8 129 108)
$cmbDead = New-UiCombo $tp1 118 126 110 @('ASAP (başlangıçla aynı)', 'Özel tarih')
$dtDead = New-UiDate $tp1 234 126 160
$chkGmt = New-UiCheck $tp1 'UTC (GMT) saatlerini kullan (kapalı: istemci yerel saati)' 8 158 385
$chkRandom = New-UiCheck $tp1 'Rastgeleleştirme (Randomization) etkin' 8 184 385
[void](New-UiLabel $tp1 'State mesaj düzeyi:' 8 214 108)
$cmbVerb = New-UiCombo $tp1 118 211 275 @('Tüm mesajlar (10)', 'Başarı ve hatalar (5)', 'Yalnızca hatalar (1)')
$chkEnabled = New-UiCheck $tp1 'Deployment etkin (Enabled)' 8 242 385
$tip.SetToolTip($txtNameFmt, 'Kullanılabilir değişkenler: {SUG} {CollectionID} {CollectionName}. Varsayılan: "{SUG} to {CollectionName}". Aynı ad zaten varsa o koleksiyon atlanır.')

# Sekme 2: kullanıcı deneyimi, deadline davranışı, restart
[void](New-UiLabel $tp2 'Kullanıcı bildirimleri:' 8 12 380)
$cmbUx = New-UiCombo $tp2 8 34 385 @('Software Center ve tüm bildirimleri gizle', 'Software Center''da göster, yalnızca restart bildirimleri', 'Software Center''da göster, tüm bildirimleri göster')
$gbDead = New-UiGroup $tp2 'Deadline davranışı' 8 68 385 72
$chkOverrideSW = New-UiCheck $gbDead 'Yazılım kurulumu (bakım penceresi dışında)' 10 20 365 $true
$chkRebootOut = New-UiCheck $gbDead 'Sistem yeniden başlatma (bakım penceresi dışında)' 10 44 365 $false
$gbRst = New-UiGroup $tp2 'Cihaz yeniden başlatma davranışı' 8 146 385 72
$chkSupWs = New-UiCheck $gbRst 'İş istasyonlarında yeniden başlatmayı bastır' 10 20 365 $true
$chkSupSrv = New-UiCheck $gbRst 'Sunucularda yeniden başlatmayı bastır' 10 44 365 $false
$chkWriteFilter = New-UiCheck $tp2 'Write filter: değişiklikleri deadline/bakım penceresinde commit et' 8 226 390
$chkPostReboot = New-UiCheck $tp2 'Restart gerekirse, restart sonrası değerlendirme çalıştır' 8 250 390
$chkSoft = New-UiCheck $tp2 'Kullanıcı tercihine göre deadline erteleme (SoftDeadline)' 8 274 390
$tip.SetToolTip($chkWriteFilter, 'PersistOnWriteFilterDevices - bazı SMS Provider sürümlerinde yoktur; yoksa atlanır ve günlüğe yazılır.')
$tip.SetToolTip($chkPostReboot, 'RequirePostRebootFullScan - bazı SMS Provider sürümlerinde yoktur; yoksa atlanır ve günlüğe yazılır.')
$tip.SetToolTip($chkSoft, 'SoftDeadlineEnabled - bazı SMS Provider sürümlerinde yoktur; yoksa atlanır ve günlüğe yazılır.')

# Sekme 3: indirme (DPLocality) ve diğer
$gbDp = New-UiGroup $tp3 'İndirme ayarları (DPLocality)' 8 8 385 124
$chkDpLocal = New-UiCheck $gbDp 'Yerel DP''den indir (16)' 10 20 365 $true
$chkDpRemote = New-UiCheck $gbDp 'Komşu/uzak DP''den indir (64)' 10 44 365 $true
$chkMetered = New-UiCheck $gbDp 'Ölçümlü bağlantıda deadline sonrası indir (524288)' 10 68 365 $true
$chkWumu = New-UiCheck $gbDp 'Microsoft Update / İnternet''ten indir (262144)' 10 92 365 $false
$lblDp = New-UiLabel $tp3 '' 8 138 385
$chkWol = New-UiCheck $tp3 'Wake-on-LAN (required deployment için)' 8 164 385 $true
$chkSub = New-UiCheck $tp3 'Alt hedeflere uygula (ApplyToSubTargets)' 8 190 385 $false
$chkMomDis = New-UiCheck $tp3 'Operations Manager uyarılarını devre dışı bırak' 8 216 385 $false
$chkMomFail = New-UiCheck $tp3 'Başarısızlıkta Operations Manager uyarısı üret' 8 242 385 $false
$tip.SetToolTip($chkWumu, 'Program.cs bu bayrağı KASITLI olarak set etmez: istemciler Microsoft''tan/internetten indirme yapmaz.')

$btnDefaults = New-UiButton $gSet 'Ayarları Varsayılana Döndür (Program.cs)' 10 364 410 28
$btnDefaults.Anchor = 'Bottom,Left,Right'

# --- Deploy eylemi ---
$gAct = New-UiGroup $form 'Deploy' 840 552 430 170
$gAct.Anchor = 'Bottom,Right'
$chkSim = New-UiCheck $gAct 'Simülasyon (hiçbir deployment oluşturma)' 12 22 330 $true
$btnClose = New-UiButton $gAct 'Kapat' 350 20 70 26
[void](New-UiLabel $gAct 'Ad örneği:' 12 52 62)
$txtPreview = New-UiTextBox $gAct 76 49 344
$txtPreview.ReadOnly = $true
$lblSummary = New-UiLabel $gAct '' 12 78 408 48
$btnDeploy = New-UiButton $gAct 'Seçili SUG''u İşaretli Koleksiyonlara Deploy Et' 12 130 408 32
$btnDeploy.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

# --- Günlük ---
$gLog = New-UiGroup $form 'Günlük' 10 730 1260 200
$gLog.Anchor = 'Bottom,Left,Right'
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ReadOnly = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.Location = New-Object System.Drawing.Point(10, 20)
$txtLog.Size = New-Object System.Drawing.Size(1240, 138)
$txtLog.Anchor = 'Top,Bottom,Left,Right'
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 9)
$gLog.Controls.Add($txtLog)
$lblLogPathCap = New-UiLabel $gLog 'Günlük dosyası:' 10 172 95
$lblLogPathCap.Anchor = 'Bottom,Left'
$txtLogPath = New-UiTextBox $gLog 108 168 1020
$txtLogPath.ReadOnly = $true
$txtLogPath.Anchor = 'Bottom,Left,Right'
$btnOpenLog = New-UiButton $gLog 'Klasörü Aç' 1135 166 115 26
$btnOpenLog.Anchor = 'Bottom,Right'

# ------------------------------------------------------------------
# Günlük, meşgul durumu, mesaj
# ------------------------------------------------------------------
# Günlük dosyası: %USERPROFILE%\Desktop\update-state\yyyy-MM-dd-<AppName>.log (state.json'dan bağımsız, sabit).
function Get-LogFilePath {
    return (Join-Path $script:LogFolderPath ('{0:yyyy-MM-dd}-{1}.log' -f (Get-Date), $script:AppName))
}

function Update-LogPathUi {
    $txtLogPath.Text = (Get-LogFilePath)
}

function Write-FileLog {
    param([string]$Level, [string]$Message)
    $file = $null
    try {
        $file = Get-LogFilePath
        if (-not (Test-Path -LiteralPath $script:LogFolderPath)) { [void](New-Item -ItemType Directory -Path $script:LogFolderPath -Force) }
        $line = '[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}' -f (Get-Date), $Level.ToUpperInvariant(), $Message
        [System.IO.File]::AppendAllText($file, $line + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        # Günlük hatası ana işlemi durdurmaz; ancak sessiz kalmasın: ekranda bir kez göster
        if (-not $script:LogWriteWarned -and $txtLog) {
            $script:LogWriteWarned = $true
            $txtLog.AppendText(('[LOG-HATA] Günlük dosyasına yazılamadı: {0}  ({1})' -f $file, $_.Exception.Message) + [Environment]::NewLine)
        }
    }
}

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $lv = $Level.ToUpperInvariant()
    if ($lv -eq 'WARN') { $lv = 'WARNING' }
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $lv, $Message
    $txtLog.AppendText($line + [Environment]::NewLine)
    Write-FileLog $lv $Message
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-Busy {
    param([bool]$On)
    foreach ($c in @($btnAddIds, $btnFile, $btnAll, $btnRefresh, $btnRemoveCache, $btnDeploy, $btnLoadSug, $btnSugReset, $btnBrowse, $btnReload)) { $c.Enabled = (-not $On) }
    if ($On) { $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor }
    else     { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    [System.Windows.Forms.Application]::DoEvents()
}

# Hata günlüğü için ayrıntı: ileti + istisna türü + betik satırı (hatanın nereden geldiğini bulmak için).
function Get-ErrorDetail {
    param($Err)
    $m = [string]$Err.Exception.Message
    $t = [string]$Err.Exception.GetType().FullName
    $l = ''
    if ($Err.InvocationInfo) {
        $l = (' | satır {0}: {1}' -f $Err.InvocationInfo.ScriptLineNumber, ([string]$Err.InvocationInfo.Line).Trim())
    }
    return ('{0}  [{1}{2}]' -f $m, $t, $l)
}

function Show-Message {
    param([string]$Text, [string]$Title = 'Bilgi', [string]$Icon = 'Information')
    [void][System.Windows.Forms.MessageBox]::Show($Text, $Title, 'OK', $Icon)
}

# ------------------------------------------------------------------
# SMS Provider bağlantısı (Program.cs ile aynı: Impersonate / PacketPrivacy / EnablePrivileges)
# ------------------------------------------------------------------
function Get-Connection {
    $site = $txtSite.Text.Trim().ToUpper()
    $server = $txtServer.Text.Trim()
    if ($site -notmatch '^[A-Z0-9]{3}$') { throw 'Site kodu boş/geçersiz. state.json yüklenmeli ("Gözat..." / "Yeniden Oku") veya site kodu (3 karakter) girilmeli.' }
    if (-not $server) { throw 'SMS Provider boş. state.json yüklenmeli ("Gözat..." / "Yeniden Oku") veya sunucu adı girilmeli.' }
    return [pscustomobject]@{ Site = $site; Server = $server }
}

function Get-SmsScope {
    $c = Get-Connection
    $key = '{0}|{1}' -f $c.Server.ToLowerInvariant(), $c.Site
    if ($script:Scope -and $script:ScopeKey -eq $key -and $script:Scope.IsConnected) { return $script:Scope }
    $opt = New-Object System.Management.ConnectionOptions
    $opt.Impersonation = [System.Management.ImpersonationLevel]::Impersonate
    $opt.Authentication = [System.Management.AuthenticationLevel]::PacketPrivacy
    $opt.EnablePrivileges = $true
    $path = '\\{0}\root\sms\site_{1}' -f $c.Server, $c.Site
    try {
        $scope = New-Object System.Management.ManagementScope($path, $opt)
        $scope.Connect()
    } catch {
        Write-Log "SMS Provider'a bağlanılamadı. Path='$path'. Hata: $($_.Exception.Message)" 'ERROR'
        throw
    }
    Write-Log "SMS Provider'a bağlanıldı. Server='$($c.Server)', Namespace='root\sms\site_$($c.Site)'"
    $script:Scope = $scope
    $script:ScopeKey = $key
    return $scope
}

# Okuma sorguları (SELECT): diğer betiklerde de kullanılan Get-WmiObject yolu.
# (Deployment OLUŞTURMA Program.cs gibi ManagementScope ile yapılır: bkz. New-SugDeployment)
function Invoke-SmsQuery {
    param([string]$Query)
    Write-FileLog 'DEBUG' ('WQL: ' + $Query)
    $c = Get-Connection
    return @(Get-WmiObject -ComputerName $c.Server -Namespace ('root\SMS\site_{0}' -f $c.Site) -Query $Query -ErrorAction Stop)
}

function Import-Groups {
    $rows = Invoke-SmsQuery 'SELECT * FROM SMS_AuthorizationList'
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

# ------------------------------------------------------------------
# SUG seçimi (state.json SUG'ı varsayılan / ★)
# ------------------------------------------------------------------
function Get-SelectedSug {
    $i = $cmbSug.SelectedIndex
    if ($i -ge 0 -and $i -lt $script:SugItems.Count) { return $script:SugItems[$i] }
    return $null
}

function Get-SelectedSugId {
    $s = Get-SelectedSug
    if ($s) { return [string]$s.Id }
    return ''
}

# SUG listesi: önce state.json SUG'ları (ilk kayıt = LastRun = ★ varsayılan), ardından (yüklendiyse) ortamdaki diğer tüm SUG'lar.
function Update-SugCombo {
    param([string]$SelectId)
    $prev = $SelectId
    if (-not $prev) { $cur = Get-SelectedSug; if ($cur) { $prev = [string]$cur.Id } }

    $items = @()
    $seen = @{}
    foreach ($e in $script:StateEntries) {
        if (-not $e.SugId -or $seen.ContainsKey($e.SugId)) { continue }
        $seen[$e.SugId] = $true
        $name = [string]$e.SugName
        $total = $null
        $g = $script:GroupById[$e.SugId]
        if ($g) { $name = $g.Name; $total = $g.Total }     # ortam yüklendiyse güncel ad/sayı
        $items += [pscustomobject]@{ Id = [string]$e.SugId; Name = $name; Total = $total; Kind = $(if ($items.Count -eq 0) { 'Latest' } else { 'State' }) }
    }
    foreach ($g in $script:Groups) {
        if ($seen.ContainsKey($g.Id)) { continue }
        $seen[$g.Id] = $true
        $items += [pscustomobject]@{ Id = $g.Id; Name = $g.Name; Total = $g.Total; Kind = 'Env' }
    }
    $script:SugItems = $items

    $script:Suspend = $true
    try {
        $cmbSug.Items.Clear()
        foreach ($it in $items) {
            $tot = ''
            if ($null -ne $it.Total) { $tot = ', ' + $it.Total + ' güncelleme' }
            switch ($it.Kind) {
                'Latest' { [void]$cmbSug.Items.Add(('★ {0}  (CI_ID {1}{2}) - state.json (varsayılan)' -f $it.Name, $it.Id, $tot)) }
                'State'  { [void]$cmbSug.Items.Add(('{0}  (CI_ID {1}{2}) - state geçmişi' -f $it.Name, $it.Id, $tot)) }
                default  { [void]$cmbSug.Items.Add(('{0}  (CI_ID {1}{2})' -f $it.Name, $it.Id, $tot)) }
            }
        }
    } finally { $script:Suspend = $false }

    $idx = 0
    if ($prev) {
        for ($i = 0; $i -lt $items.Count; $i++) { if ($items[$i].Id -eq $prev) { $idx = $i; break } }
    }
    if ($items.Count -gt 0) { $cmbSug.SelectedIndex = $idx }
    else {
        Show-Collections
        Update-Preview
    }
}

function Select-SugById {
    param([string]$Id)
    if (-not $Id) { return }
    for ($i = 0; $i -lt $script:SugItems.Count; $i++) {
        if ($script:SugItems[$i].Id -eq $Id) { $cmbSug.SelectedIndex = $i; return }
    }
}

# Ortamdaki tüm SUG'ları okuyup listeye ekler. -Quiet: hata olursa açılır pencere gösterme (yalnızca günlük).
function Invoke-LoadSugs {
    param([switch]$Quiet)
    Set-Busy $true
    try {
        [void](Get-Connection)
        Write-Log 'Ortamdaki Software Update Group''lar okunuyor...'
        Import-Groups
        Update-SugCombo
        Write-Log "Ortamdan $($script:Groups.Count) Software Update Group yüklendi (★ state.json'daki varsayılan SUG listenin başında)."
    }
    catch {
        Write-Log (Get-ErrorDetail $_) 'ERROR'
        if (-not $Quiet) { Show-Message $_.Exception.Message 'Hata' 'Error' }
    }
    finally { Set-Busy $false }
}

# ------------------------------------------------------------------
# state.json -> UI
# ------------------------------------------------------------------
function Import-StateFile {
    param([string]$Path)

    $script:StateFile = $Path
    $script:StateEntries = @()
    $cmbState.Items.Clear()
    Update-LogPathUi

    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) {
        $txtStateFile.Text = '(state.json bulunamadı - "Gözat..." ile seçin)'
        Write-Log 'state.json bulunamadı. Site/SMS Provider yalnızca state.json''dan alınır: "Gözat..." ile dosyayı seçin (veya alanları elle doldurun).' 'WARN'
        Update-SugCombo
        return
    }
    $txtStateFile.Text = $Path

    try {
        $state = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Log "state.json okunamadı: $($_.Exception.Message)" 'ERROR'
        Update-SugCombo
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
        if (-not $e.SugId) { continue }          # bu araç için SUG gerekli
        $key = '{0}|{1}|{2}' -f $e.Timestamp, $e.SugId, $e.PkgId
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        $entries += $e
    }
    $script:StateEntries = $entries

    if ($entries.Count -eq 0) {
        Write-Log 'state.json içinde SUG (SugId) içeren kayıt bulunamadı.' 'WARN'
        Update-SugCombo
        return
    }

    foreach ($e in $entries) { [void]$cmbState.Items.Add((Get-EntryLabel $e)) }
    Write-Log "State yüklendi: $Path ($($entries.Count) kayıt). Varsayılan SUG: '$($entries[0].SugName)' (CI_ID $($entries[0].SugId))"
    Update-SugCombo $entries[0].SugId
    $cmbState.SelectedIndex = 0
}

# ------------------------------------------------------------------
# Koleksiyon cache'i
# ------------------------------------------------------------------
function Get-CachePath {
    $site = $txtSite.Text.Trim().ToUpper()
    if ($site -notmatch '^[A-Z0-9]{3}$') { return $null }
    return (Join-Path $script:CacheFolder ('collection-cache-{0}.json' -f $site))
}

function Import-CollectionCache {
    $script:Cols.Clear()
    $script:Checked = @{}
    $script:CacheSite = $txtSite.Text.Trim().ToUpper()
    $p = Get-CachePath
    if (-not $p -or -not (Test-Path -LiteralPath $p)) {
        Write-Log "Koleksiyon cache'i yok ($script:CacheSite). ID'leri ekleyerek oluşturabilirsiniz."
        return
    }
    try {
        $j = Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Write-Log "Koleksiyon cache'i okunamadı: $($_.Exception.Message)" 'WARN'
        return
    }
    foreach ($c in @($j.Collections)) {
        if (-not $c -or -not $c.CollectionID) { continue }
        $deps = @()
        if ($c.Deployments) {
            foreach ($d in @($c.Deployments)) {
                if (-not $d) { continue }
                # PowerShell 7 ISO tarihleri otomatik DateTime'a çevirebilir; metne geri normalize et
                foreach ($f in @('Created', 'Start', 'Deadline')) {
                    if ($d.$f -is [datetime]) { $d.$f = $d.$f.ToString('s') }
                }
                $deps += $d
            }
        }
        $rec = [pscustomobject]@{
            CollectionID         = [string]$c.CollectionID
            Name                 = [string]$c.Name
            MemberCount          = [int]$c.MemberCount
            CollectionType       = [int]$c.CollectionType
            LimitTo              = [string]$c.LimitTo
            Comment              = [string]$c.Comment
            Cached               = [string]$c.Cached
            DeploymentsRefreshed = [string]$c.DeploymentsRefreshed
            Deployments          = $deps
        }
        $script:Cols[$rec.CollectionID] = $rec
    }
    Write-Log "Koleksiyon cache'i yüklendi: $($script:Cols.Count) koleksiyon ($p)"
}

function Save-CollectionCache {
    $p = Get-CachePath
    if (-not $p) { return }
    try {
        if (-not (Test-Path -LiteralPath $script:CacheFolder)) { [void](New-Item -ItemType Directory -Path $script:CacheFolder -Force) }
        $obj = [pscustomobject]@{
            Version     = 1
            Site        = $txtSite.Text.Trim().ToUpper()
            Server      = $txtServer.Text.Trim()
            Updated     = (Get-Date).ToString('s')
            Collections = @($script:Cols.Values | Sort-Object -Property Name)
        }
        $json = ConvertTo-Json -InputObject $obj -Depth 6
        $tmp = $p + '.tmp'
        Set-Content -LiteralPath $tmp -Value $json -Encoding UTF8
        Move-Item -LiteralPath $tmp -Destination $p -Force
    } catch {
        Write-Log "Koleksiyon cache'i yazılamadı: $($_.Exception.Message)" 'WARN'
    }
}

# ------------------------------------------------------------------
# Deploy bilgisi (SMS_UpdateGroupAssignment)
# ------------------------------------------------------------------
function ConvertTo-DeploymentRecord {
    param($R)
    $cis = @()
    foreach ($x in @($R.AssignedCIs)) { if ($null -ne $x) { $cis += [string]$x } }
    if ($cis.Count -eq 0 -and $null -ne $R.AssignedUpdateGroup) { $cis += [string]$R.AssignedUpdateGroup }
    $names = @()
    foreach ($ci in $cis) {
        $g = $script:GroupById[$ci]
        if ($g) { $names += $g.Name } else { $names += "CI_ID $ci" }
    }
    $dead = ConvertFrom-DmtfSafe $R.EnforcementDeadline
    $dct = 1
    if ($null -ne $R.DesiredConfigType) { try { $dct = [int]$R.DesiredConfigType } catch { } }
    $type = 'Required'
    if ($dct -eq 2) { $type = 'NotAllowed' } elseif (-not $dead) { $type = 'Available' }
    $sr = 0; try { $sr = [int]$R.SuppressReboot } catch { }
    $dp = 0; try { $dp = [int64]$R.DPLocality } catch { }
    return [pscustomobject]@{
        AssignmentID  = [string]$R.AssignmentID
        Name          = [string]$R.AssignmentName
        SugCiId       = ($cis -join ',')
        SugName       = ($names -join '; ')
        Created       = (ConvertFrom-DmtfSafe $R.CreationTime)
        Start         = (ConvertFrom-DmtfSafe $R.StartTime)
        Deadline      = $dead
        Type          = $type
        Enabled       = [bool]$R.Enabled
        NotifyUser    = [bool]$R.NotifyUser
        UserUI        = [bool]$R.UserUIExperience
        SuppressReboot = $sr
        OverrideSW    = [bool]$R.OverrideServiceWindows
        RebootOutside = [bool]$R.RebootOutsideOfServiceWindows
        Wol           = [bool]$R.WoLEnabled
        DpLocality    = $dp
    }
}

# Tüm deployment'ları TEK sorguyla okur; cache'deki her koleksiyona dağıtır ve ad çakışması kontrolü için adları saklar.
function Update-AllDeployments {
    Write-Log 'Deploy bilgileri okunuyor (SMS_UpdateGroupAssignment)...'
    if ($script:Groups.Count -eq 0) {
        try { Import-Groups } catch { Write-Log "SUG adları okunamadı: $($_.Exception.Message)" 'WARN' }
    }
    $rows = @(Invoke-SmsQuery 'SELECT * FROM SMS_UpdateGroupAssignment')
    $by = @{}
    $script:AssignmentNames.Clear()
    foreach ($r in $rows) {
        $nm = [string]$r.AssignmentName
        if ($nm) { [void]$script:AssignmentNames.Add($nm) }
        $tc = ([string]$r.TargetCollectionID).ToUpper()
        if (-not $tc) { continue }
        if (-not $by.ContainsKey($tc)) { $by[$tc] = @() }
        $by[$tc] += (ConvertTo-DeploymentRecord $r)
    }
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $n = 0
    foreach ($c in @($script:Cols.Values)) {
        $k = $c.CollectionID.ToUpper()
        if ($by.ContainsKey($k)) { $c.Deployments = @($by[$k] | Sort-Object -Property { [string]$_.Created } -Descending) }
        else { $c.Deployments = @() }
        $c.DeploymentsRefreshed = $stamp
        $n++
    }
    Save-CollectionCache
    Write-Log "$($rows.Count) deployment okundu; cache'deki $n koleksiyonun deploy bilgisi güncellendi."
}

# Seçili SUG'a ait, bu koleksiyondaki en yeni deployment (yoksa $null).
function Get-SugDeployment {
    param($Col, [string]$SugId)
    if (-not $Col -or -not $SugId) { return $null }
    $m = @($Col.Deployments | Where-Object { $_ -and (@(([string]$_.SugCiId) -split ',') -contains $SugId) } |
           Sort-Object -Property { [string]$_.Created } -Descending)
    if ($m.Count -gt 0) { return $m[0] }
    return $null
}

# ------------------------------------------------------------------
# Koleksiyon listesi (grid)
# ------------------------------------------------------------------
function Show-Collections {
    $script:Suspend = $true
    try {
        $gridCol.Rows.Clear()
        $sugId = Get-SelectedSugId
        $q = $txtColFilter.Text.Trim()
        $onlyNew = $chkOnlyNew.Checked
        $shown = 0
        foreach ($c in @($script:Cols.Values | Sort-Object -Property Name)) {
            if ($q) {
                $hitId = ($c.CollectionID.IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
                $hitNm = ($c.Name.IndexOf($q, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
                if (-not $hitId -and -not $hitNm) { continue }
            }
            $known = [bool]$c.DeploymentsRefreshed
            $m = Get-SugDeployment $c $sugId
            if ($onlyNew -and $m) { continue }

            $typeText = 'Cihaz'
            if ($c.CollectionType -eq 1) { $typeText = 'Kullanıcı' }
            elseif ($c.CollectionType -ne 2) { $typeText = [string]$c.CollectionType }

            $status = '?'
            $date = ''
            $shape = ''
            if ($m) {
                $status = 'Evet'
                $date = Format-Dt $m.Created
                $shape = Format-DeployShape $m
            } elseif ($known) {
                $status = 'Hayır'
            }
            $deps = @($c.Deployments | Where-Object { $_ })
            $last = ''
            if ($deps.Count -gt 0) {
                $d0 = $deps[0]
                $last = ('{0} - {1}' -f (Format-Dt $d0.Created), $d0.SugName)
            }
            $chk = ($script:Checked.ContainsKey($c.CollectionID) -and $script:Checked[$c.CollectionID])
            $vals = [object[]]@($chk, $c.CollectionID, $c.Name, $c.MemberCount, $typeText, $status, $date, $shape, $deps.Count, $last, $c.Cached)
            $idx = $gridCol.Rows.Add($vals)
            $row = $gridCol.Rows[$idx]
            $row.Tag = $c.CollectionID
            if ($m) { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen }
            if ($c.CollectionType -ne 2) { $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Gray }
            $shown++
        }
    } finally { $script:Suspend = $false }
    $nChecked = 0
    foreach ($k in @($script:Checked.Keys)) { if ($script:Checked[$k]) { $nChecked++ } }
    $lblColCount.Text = ('Cache: {0} koleksiyon  |  Görünen: {1}  |  İşaretli (deploy): {2}  |  Yeşil: seçili SUG zaten deploy edilmiş' -f $script:Cols.Count, $shown, $nChecked)
    Show-Details
}

function Show-Details {
    $gridDet.Rows.Clear()
    $row = $gridCol.CurrentRow
    if (-not $row -or -not $row.Tag) { $gDet.Text = 'Seçili koleksiyonun deploy geçmişi'; return }
    $c = Get-Col ([string]$row.Tag)
    if (-not $c) { return }
    $ref = 'deploy bilgisi henüz yenilenmedi'
    if ($c.DeploymentsRefreshed) { $ref = 'son yenileme: ' + $c.DeploymentsRefreshed }
    $gDet.Text = ('Seçili koleksiyonun deploy geçmişi: {0} ({1})  -  {2}' -f $c.Name, $c.CollectionID, $ref)
    $sugId = Get-SelectedSugId
    foreach ($d in @($c.Deployments | Where-Object { $_ })) {
        $vals = [object[]]@($d.Name, $d.SugName, (Format-Dt $d.Created), (Format-Dt $d.Start), (Format-Dt $d.Deadline), $d.Type, (Format-DeployShape $d), $d.Enabled, $d.AssignmentID)
        $idx = $gridDet.Rows.Add($vals)
        if ($sugId -and (@(([string]$d.SugCiId) -split ',') -contains $sugId)) {
            $gridDet.Rows[$idx].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGreen
        }
    }
}

# ------------------------------------------------------------------
# Koleksiyon ekleme (ID ile / dosyadan / tümü)
# ------------------------------------------------------------------
function ConvertTo-CollectionRecord {
    param($R, $Existing)
    $deps = @()
    $ref = ''
    if ($Existing) { $deps = @($Existing.Deployments); $ref = [string]$Existing.DeploymentsRefreshed }
    $mc = 0; try { $mc = [int]$R.MemberCount } catch { }
    $ct = 0; try { $ct = [int]$R.CollectionType } catch { }
    return [pscustomobject]@{
        CollectionID         = ([string]$R.CollectionID).ToUpper()
        Name                 = [string]$R.Name
        MemberCount          = $mc
        CollectionType       = $ct
        LimitTo              = [string]$R.LimitToCollectionName
        Comment              = [string]$R.Comment
        Cached               = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        DeploymentsRefreshed = $ref
        Deployments          = $deps
    }
}

function Split-CollectionIds {
    param([string]$Text)
    $valid = @()
    $bad = @()
    foreach ($t in ($Text -split '[\s,;]+')) {
        $s = $t.Trim().ToUpper()
        if (-not $s) { continue }
        if ($s -match '^[A-Z0-9]{8}$') { if ($valid -notcontains $s) { $valid += $s } }
        else { $bad += $s }
    }
    return [pscustomobject]@{ Valid = $valid; Bad = $bad }
}

function Add-CollectionsById {
    param([string[]]$Ids)
    [void](Get-Connection)
    $found = @{}
    for ($i = 0; $i -lt $Ids.Count; $i += 40) {
        $last = [Math]::Min($i + 39, $Ids.Count - 1)
        $inList = (@($Ids[$i..$last]) | ForEach-Object { "'$_'" }) -join ','
        foreach ($r in @(Invoke-SmsQuery "SELECT * FROM SMS_Collection WHERE CollectionID IN ($inList)")) {
            $found[([string]$r.CollectionID).ToUpper()] = $r
        }
    }
    $added = 0
    foreach ($id in $Ids) {
        if (-not $found.ContainsKey($id)) { Write-Log "'$id' ID'sine sahip bir koleksiyon bulunamadı." 'WARN'; continue }
        $r = $found[$id]
        $rec = ConvertTo-CollectionRecord $r (Get-Col $id)
        $script:Cols[$id] = $rec
        if ($rec.CollectionType -ne 2) {
            Write-Log "Koleksiyon eklendi ancak CİHAZ koleksiyonu değil: $id '$($rec.Name)' (software update deployment için cihaz koleksiyonu gerekir)." 'WARN'
        } else {
            Write-Log "Koleksiyon bulundu. ID='$id', Ad='$($rec.Name)', Üye=$($rec.MemberCount)"
        }
        $added++
    }
    return $added
}

# ------------------------------------------------------------------
# Deploy ayarları (UI <-> ayar tablosu). Varsayılanlar Program.cs ile aynıdır.
# ------------------------------------------------------------------
function Get-DefaultSettings {
    $d = Get-Date
    $d = $d.AddMinutes(-$d.Minute).AddSeconds(-$d.Second).AddHours(1)
    return @{
        NameFormat = '{SUG} to {CollectionName}'  # varsayılan: "<SUG Adı> to <Koleksiyon Adı>"
        Description = ''
        DeployType = 0                            # 0 = Required
        StartMode = 0                             # 0 = ASAP
        StartCustom = $d
        DeadlineMode = 0                          # 0 = ASAP
        DeadlineCustom = $d.AddDays(1)
        UseGmt = $false                           # UseGMTTimes = false
        Random = $true                            # RandomizationEnabled = true
        Verbosity = 0                             # 0 -> StateMessageVerbosity = 10 (tüm mesajlar)
        Enabled = $true
        Ux = 0                                    # SC ve bildirimler gizli (NotifyUser = false, UserUIExperience = false)
        OverrideSW = $true                        # OverrideServiceWindows = true
        RebootOutside = $false                    # RebootOutsideOfServiceWindows = false
        SupWs = $true                             # SuppressReboot bit 1 (workstation)
        SupSrv = $true                            # SuppressReboot bit 2 (server)  -> 3
        WriteFilter = $true                       # PersistOnWriteFilterDevices = true (isteğe bağlı property)
        PostReboot = $true                        # RequirePostRebootFullScan = true (isteğe bağlı property)
        SoftDeadline = $false                     # SoftDeadlineEnabled = false (isteğe bağlı property)
        DpLocal = $true                           # 16
        DpRemote = $true                          # 64
        Metered = $true                           # 524288
        Wumu = $false                             # 262144 (Program.cs set etmez)
        Wol = $true                               # WoLEnabled = true
        SubTargets = $false                       # ApplyToSubTargets = false
        MomDis = $false                           # DisableMomAlerts = false
        MomFail = $false                          # RaiseMomAlertsOnFailure = false
    }
}

function Set-DeploySettingsUi {
    param([hashtable]$S)
    $txtNameFmt.Text = $S.NameFormat
    $txtDesc.Text = $S.Description
    $cmbType.SelectedIndex = $S.DeployType
    $cmbStart.SelectedIndex = $S.StartMode
    $dtStart.Value = $S.StartCustom
    $cmbDead.SelectedIndex = $S.DeadlineMode
    $dtDead.Value = $S.DeadlineCustom
    $chkGmt.Checked = $S.UseGmt
    $chkRandom.Checked = $S.Random
    $cmbVerb.SelectedIndex = $S.Verbosity
    $chkEnabled.Checked = $S.Enabled
    $cmbUx.SelectedIndex = $S.Ux
    $chkOverrideSW.Checked = $S.OverrideSW
    $chkRebootOut.Checked = $S.RebootOutside
    $chkSupWs.Checked = $S.SupWs
    $chkSupSrv.Checked = $S.SupSrv
    $chkWriteFilter.Checked = $S.WriteFilter
    $chkPostReboot.Checked = $S.PostReboot
    $chkSoft.Checked = $S.SoftDeadline
    $chkDpLocal.Checked = $S.DpLocal
    $chkDpRemote.Checked = $S.DpRemote
    $chkMetered.Checked = $S.Metered
    $chkWumu.Checked = $S.Wumu
    $chkWol.Checked = $S.Wol
    $chkSub.Checked = $S.SubTargets
    $chkMomDis.Checked = $S.MomDis
    $chkMomFail.Checked = $S.MomFail
    Update-SettingsState
}

function Get-DpLocalityValue {
    $v = 0
    if ($chkDpLocal.Checked)  { $v += 16 }
    if ($chkDpRemote.Checked) { $v += 64 }
    if ($chkMetered.Checked)  { $v += 524288 }
    if ($chkWumu.Checked)     { $v += 262144 }
    return $v
}

# Seçeneklerin birbirine bağlı etkin/pasif durumları ve DPLocality etiketi
function Update-SettingsState {
    $dtStart.Enabled = ($cmbStart.SelectedIndex -eq 1)
    $isReq = ($cmbType.SelectedIndex -eq 0)
    $cmbDead.Enabled = $isReq
    $dtDead.Enabled = ($isReq -and $cmbDead.SelectedIndex -eq 1)
    $lblDp.Text = ('DPLocality değeri: {0}   (Program.cs varsayılanı: 524368)' -f (Get-DpLocalityValue))
}

function Get-DeploySettings {
    $S = @{}
    $S.NameFormat = $txtNameFmt.Text.Trim()
    if (-not $S.NameFormat) { throw 'Deployment adı şablonu boş olamaz.' }
    if ($S.NameFormat -notmatch '\{CollectionID\}|\{CollectionName\}') {
        throw 'Deployment adı şablonunda {CollectionID} (veya {CollectionName}) bulunmalı; aksi halde her koleksiyon için aynı ad üretilir.'
    }
    $S.Description = $txtDesc.Text
    $S.Required = ($cmbType.SelectedIndex -eq 0)
    $S.StartAsap = ($cmbStart.SelectedIndex -eq 0)
    $S.StartCustom = $dtStart.Value
    $S.DeadlineAsap = ($cmbDead.SelectedIndex -eq 0)
    $S.DeadlineCustom = $dtDead.Value
    if ($S.Required -and -not $S.StartAsap -and -not $S.DeadlineAsap -and $S.DeadlineCustom -lt $S.StartCustom) {
        throw 'Deadline, başlangıç zamanından önce olamaz.'
    }
    $S.UseGmt = $chkGmt.Checked
    $S.Random = $chkRandom.Checked
    $S.Verbosity = @(10, 5, 1)[[Math]::Max(0, $cmbVerb.SelectedIndex)]
    $S.Enabled = $chkEnabled.Checked
    $S.Ux = $cmbUx.SelectedIndex
    $S.OverrideSW = $chkOverrideSW.Checked
    $S.RebootOutside = $chkRebootOut.Checked
    $S.SuppressReboot = [uint32](($(if ($chkSupWs.Checked) { 1 } else { 0 })) + ($(if ($chkSupSrv.Checked) { 2 } else { 0 })))
    $S.WriteFilter = $chkWriteFilter.Checked
    $S.PostReboot = $chkPostReboot.Checked
    $S.SoftDeadline = $chkSoft.Checked
    $S.DpLocality = [uint32](Get-DpLocalityValue)
    $S.Wol = $chkWol.Checked
    $S.SubTargets = $chkSub.Checked
    $S.MomDis = $chkMomDis.Checked
    $S.MomFail = $chkMomFail.Checked
    return $S
}

function Get-DeploymentName {
    param([string]$Format, [string]$SugName, $Col)
    $n = $Format.Replace('{SUG}', $SugName).Replace('{CollectionID}', [string]$Col.CollectionID).Replace('{CollectionName}', [string]$Col.Name)
    return $n.Trim()
}

function Update-Preview {
    $sug = Get-SelectedSug
    $sugName = '<SUG>'
    if ($sug) { $sugName = $sug.Name }
    $first = $null
    $n = 0
    foreach ($k in @($script:Checked.Keys)) {
        if ($script:Checked[$k]) {
            $n++
            if (-not $first) { $first = Get-Col $k }
        }
    }
    if (-not $first) { $first = [pscustomobject]@{ CollectionID = 'COLID001'; Name = '<Koleksiyon>' } }
    $fmt = $txtNameFmt.Text.Trim()
    if (-not $fmt) { $fmt = '{SUG} to {CollectionName}' }
    $txtPreview.Text = Get-DeploymentName $fmt $sugName $first
    $typeText = 'Required'
    if ($cmbType.SelectedIndex -eq 1) { $typeText = 'Available' }
    $sugText = '(seçilmedi)'
    if ($sug) { $sugText = ('{0} (CI_ID {1})' -f $sug.Name, $sug.Id) }
    $lblSummary.Text = ("SUG: {0}`nİşaretli koleksiyon: {1}  |  Tür: {2}  |  DPLocality: {3}" -f $sugText, $n, $typeText, (Get-DpLocalityValue))
}

# ------------------------------------------------------------------
# Deployment oluşturma (Program.cs CreateUpdateGroupDeployment ile aynı alanlar)
# ------------------------------------------------------------------
# Deployment özelliği kaydı (ad, değer, isteğe bağlı mı).
# NOT: Bu PowerShell sürümünde List[object] üzerinde @(...) "Argument types do not match" hatası verdiği için
# burada ve diğer yerlerde generic List yerine düz dizi (+=) kullanılır.
function New-AsgProp {
    param([string]$Name, $Value, [bool]$Optional)
    return [pscustomobject]@{ Name = $Name; Value = $Value; Optional = $Optional }
}

function Get-AssignmentProperties {
    param([int]$SugCiId, [string]$CollectionId, [string]$Name, [hashtable]$S)

    $props = @()

    $now = Get-Date
    $start = $now
    if (-not $S.StartAsap) { $start = $S.StartCustom }
    $deadline = $start
    if (-not $S.DeadlineAsap) { $deadline = $S.DeadlineCustom }
    if ($S.UseGmt) { $start = $start.ToUniversalTime(); $deadline = $deadline.ToUniversalTime() }

    # ---------- Genel / Hedef ----------
    $props += New-AsgProp 'AssignmentName' $Name $false
    $props += New-AsgProp 'AssignmentDescription' ([string]$S.Description) $false
    $props += New-AsgProp 'TargetCollectionID' $CollectionId $false
    $props += New-AsgProp 'AssignedCIs' ([int[]]@($SugCiId)) $false
    $props += New-AsgProp 'AssignedUpdateGroup' ([int]$SugCiId) $false
    $props += New-AsgProp 'ApplyToSubTargets' ([bool]$S.SubTargets) $false

    # ---------- Deployment türü ----------
    $props += New-AsgProp 'DesiredConfigType' ([int]1) $false      # 1 = Required
    $props += New-AsgProp 'AssignmentType' ([int]5) $false         # Update Group deployment türü için tipik değer
    $props += New-AsgProp 'AssignmentAction' ([int]2) $false       # Install

    # ---------- Zamanlama ----------
    $props += New-AsgProp 'StartTime' ([System.Management.ManagementDateTimeConverter]::ToDmtfDateTime($start)) $false
    if ($S.Required) {
        $props += New-AsgProp 'EnforcementDeadline' ([System.Management.ManagementDateTimeConverter]::ToDmtfDateTime($deadline)) $false
    }
    $props += New-AsgProp 'UseGMTTimes' ([bool]$S.UseGmt) $false
    $props += New-AsgProp 'RandomizationEnabled' ([bool]$S.Random) $false
    $props += New-AsgProp 'SoftDeadlineEnabled' ([bool]$S.SoftDeadline) $true

    # ---------- Alert / state message ----------
    $props += New-AsgProp 'StateMessageVerbosity' ([int]$S.Verbosity) $false
    $props += New-AsgProp 'LimitStateMessageVerbosity' $false $false
    $props += New-AsgProp 'StateMessagePriority' ([int]1) $false

    # ---------- Kullanıcı deneyimi ----------
    # 0: SC ve bildirimler gizli | 1: SC'de göster, yalnızca restart bildirimi | 2: SC'de göster, tüm bildirimler
    $notify = ($S.Ux -eq 2)
    $userUi = ($S.Ux -ge 1)
    $props += New-AsgProp 'NotifyUser' $notify $false
    $props += New-AsgProp 'UserUIExperience' $userUi $false

    # ---------- Deadline davranışı / restart ----------
    $props += New-AsgProp 'OverrideServiceWindows' ([bool]$S.OverrideSW) $false
    $props += New-AsgProp 'RebootOutsideOfServiceWindows' ([bool]$S.RebootOutside) $false
    $props += New-AsgProp 'SuppressReboot' ([uint32]$S.SuppressReboot) $false
    $props += New-AsgProp 'PersistOnWriteFilterDevices' ([bool]$S.WriteFilter) $true
    $props += New-AsgProp 'RequirePostRebootFullScan' ([bool]$S.PostReboot) $true

    # ---------- Wake on LAN / uyarılar ----------
    $props += New-AsgProp 'WoLEnabled' ([bool]$S.Wol) $false
    $props += New-AsgProp 'DisableMomAlerts' ([bool]$S.MomDis) $false
    $props += New-AsgProp 'RaiseMomAlertsOnFailure' ([bool]$S.MomFail) $false

    # ---------- İndirme ayarları ----------
    $props += New-AsgProp 'DPLocality' ([uint32]$S.DpLocality) $false

    # ---------- Diğer varsayılan / zorunlu alanlar ----------
    $props += New-AsgProp 'Enabled' ([bool]$S.Enabled) $false
    $props += New-AsgProp 'LogComplianceToWinEvent' $false $false
    $props += New-AsgProp 'SendDetailedNonComplianceStatus' $false $false
    $props += New-AsgProp 'NonComplianceCriticality' ([int]1) $false
    $props += New-AsgProp 'LocaleID' ([uint32]1033) $false

    return @($props)
}

function Set-AssignmentProp {
    param($Obj, [string]$Name, $Value, [bool]$Optional)
    try {
        $Obj.Properties[$Name].Value = $Value
    } catch {
        if ($Optional) {
            Write-Log "'$Name' property'si bu SMS Provider'da bulunamadığı için atlandı. Hata: $($_.Exception.Message)" 'WARNING'
        } else {
            throw "'$Name' atanamadı: $($_.Exception.Message)"
        }
    }
}

function New-SugDeployment {
    param([int]$SugCiId, [string]$CollectionId, [string]$Name, [hashtable]$S)
    $scope = Get-SmsScope
    $props = Get-AssignmentProperties -SugCiId $SugCiId -CollectionId $CollectionId -Name $Name -S $S
    $cls = New-Object System.Management.ManagementClass($scope, (New-Object System.Management.ManagementPath('SMS_UpdateGroupAssignment')), $null)
    $asg = $cls.CreateInstance()
    try {
        foreach ($p in $props) { Set-AssignmentProp -Obj $asg -Name $p.Name -Value $p.Value -Optional $p.Optional }
        $path = $asg.Put()
        $id = ''
        if ($path -and ([string]$path.RelativePath) -match 'AssignmentID=(\d+)') { $id = $Matches[1] }
        return $id
    } finally {
        $asg.Dispose()
        $cls.Dispose()
    }
}

# ------------------------------------------------------------------
# Deploy akışı
# ------------------------------------------------------------------
function Invoke-Deploy {
    $simulate = $chkSim.Checked
    $ids = @($script:Checked.Keys | Where-Object { $script:Checked[$_] -and (Get-Col $_) } | Sort-Object)
    if ($ids.Count -eq 0) { Show-Message 'Önce listeden en az bir koleksiyonu işaretleyin (Deploy sütunu).' 'Uyarı' 'Warning'; return }
    $sugItem = Get-SelectedSug
    if (-not $sugItem) { Show-Message 'Deploy edilecek Software Update Group seçilmedi.' 'Uyarı' 'Warning'; return }

    Set-Busy $true
    try {
        $S = Get-DeploySettings
        [void](Get-Connection)

        # --- SUG doğrula (Program.cs: GetUpdateGroupCiId) ---
        $sugId = [string]$sugItem.Id
        if ($sugId -notmatch '^\d+$') { throw "Geçersiz SUG CI_ID: '$sugId'" }
        $sugRows = @(Invoke-SmsQuery "SELECT * FROM SMS_AuthorizationList WHERE CI_ID = $sugId")
        if ($sugRows.Count -eq 0) { throw "CI_ID $sugId ile bir Software Update Group bulunamadı (silinmiş olabilir)." }
        $sugName = [string]$sugRows[0].LocalizedDisplayName
        $sugTotal = 0; try { $sugTotal = [int]$sugRows[0].NumberOfUpdates } catch { }
        Write-Log "Software Update Group bulundu. Ad='$sugName', CI_ID=$sugId, Güncelleme=$sugTotal"
        if ($sugItem.Name -and $sugItem.Name -ne $sugName) {
            Write-Log "Listedeki SUG adı ('$($sugItem.Name)') SCCM'deki ad ('$sugName') ile farklı; SCCM'deki ad kullanılacak." 'WARN'
        }

        # --- Koleksiyonları canlı doğrula (Program.cs: GetCollectionName) ---
        $live = @{}
        for ($i = 0; $i -lt $ids.Count; $i += 40) {
            $last = [Math]::Min($i + 39, $ids.Count - 1)
            $inList = (@($ids[$i..$last]) | ForEach-Object { "'$_'" }) -join ','
            foreach ($r in @(Invoke-SmsQuery "SELECT CollectionID, Name, CollectionType, MemberCount FROM SMS_Collection WHERE CollectionID IN ($inList)")) {
                $live[([string]$r.CollectionID).ToUpper()] = $r
            }
        }

        # --- Güncel deploy bilgisi: ad çakışması ve "bu SUG zaten deploy edilmiş mi" kontrolü ---
        Update-AllDeployments
        Show-Collections

        $plan = @()
        foreach ($id in $ids) {
            $col = Get-Col $id
            $item = [pscustomobject]@{ Id = $id; Name = $col.Name; DeployName = ''; Status = 'NEW'; Note = '' }
            if (-not $live.ContainsKey($id)) {
                $item.Status = 'SKIP'; $item.Note = 'SCCM''de bulunamadı (silinmiş olabilir)'
            } else {
                $lr = $live[$id]
                $col.Name = [string]$lr.Name
                $item.Name = $col.Name
                try { $col.CollectionType = [int]$lr.CollectionType } catch { }
                try { $col.MemberCount = [int]$lr.MemberCount } catch { }
                $item.DeployName = Get-DeploymentName $S.NameFormat $sugName $col
                if ($col.CollectionType -ne 2) {
                    $item.Status = 'SKIP'; $item.Note = 'cihaz koleksiyonu değil'
                } elseif (-not $item.DeployName) {
                    $item.Status = 'SKIP'; $item.Note = 'deployment adı boş'
                } elseif ($script:AssignmentNames.Contains($item.DeployName)) {
                    $item.Status = 'SKIP'; $item.Note = 'aynı adlı deployment zaten var'
                } else {
                    $prior = Get-SugDeployment $col $sugId
                    if ($prior) { $item.Note = ('DİKKAT: bu SUG bu koleksiyona zaten deploy edilmiş ({0}, "{1}")' -f (Format-Dt $prior.Created), $prior.Name) }
                }
            }
            $plan += $item
        }
        $todo = @($plan | Where-Object { $_.Status -eq 'NEW' })
        $skip = @($plan | Where-Object { $_.Status -ne 'NEW' })
        foreach ($sk in $skip) { Write-Log "[$($sk.Id)] atlanacak: $($sk.Note)" 'WARN' }

        if ($todo.Count -eq 0) {
            Show-Message "Deploy edilecek koleksiyon kalmadı (hepsi atlandı). Ayrıntılar günlükte." 'Bilgi' 'Information'
            return
        }

        # --- Onay ---
        $lines = @()
        foreach ($t in ($todo | Select-Object -First 15)) {
            $ln = '- {0} ({1}) -> "{2}"' -f $t.Name, $t.Id, $t.DeployName
            if ($t.Note) { $ln += "`n    " + $t.Note }
            $lines += $ln
        }
        if ($todo.Count -gt 15) { $lines += ('... ve {0} koleksiyon daha.' -f ($todo.Count - 15)) }
        $typeText = 'Required'
        if (-not $S.Required) { $typeText = 'Available (deadline yok)' }
        $msg = ''
        if ($simulate) { $msg += "SİMÜLASYON - hiçbir deployment oluşturulmayacak.`n`n" }
        $msg += ('SUG: {0} (CI_ID {1}, {2} güncelleme)' -f $sugName, $sugId, $sugTotal) + "`n"
        $msg += ('Tür: {0}  |  DPLocality: {1}' -f $typeText, $S.DpLocality) + "`n`n"
        $msg += ('{0} koleksiyona deployment oluşturulacak:' -f $todo.Count) + "`n`n" + ($lines -join "`n")
        if ($skip.Count -gt 0) { $msg += ("`n`nATLANACAK: {0} koleksiyon (ayrıntı günlükte)." -f $skip.Count) }
        if ($sugTotal -eq 0) { $msg += "`n`nUYARI: Bu SUG'da hiç güncelleme yok." }
        if (-not $simulate) { $msg += "`n`nİstemcilere güncellemeler ZORUNLU olarak dağıtılır. Devam edilsin mi?" } else { $msg += "`n`nDevam edilsin mi?" }
        if ([System.Windows.Forms.MessageBox]::Show($msg, 'Deploy onayı', 'YesNo', 'Warning', 'Button2') -ne 'Yes') {
            Write-Log 'Kullanıcı deploy işlemini iptal etti.' 'WARN'
            return
        }

        # --- Oluştur ---
        $created = 0; $failed = 0
        Write-Log ("İşlem başladı. SUG='{0}', Koleksiyon={1}, Simülasyon={2}" -f $sugName, $todo.Count, $simulate)
        foreach ($t in $todo) {
            try {
                if ($simulate) {
                    Write-Log "[SİMÜLASYON] [$($t.Id)] '$($t.DeployName)' oluşturulacaktı. Ayarlar:" 'WARN'
                    foreach ($p in (Get-AssignmentProperties -SugCiId ([int]$sugId) -CollectionId $t.Id -Name $t.DeployName -S $S)) {
                        $opt = ''
                        if ($p.Optional) { $opt = '  (isteğe bağlı)' }
                        $val = $p.Value
                        if ($val -is [array]) { $val = ($val -join ',') }
                        Write-Log ("      {0} = {1}{2}" -f $p.Name, $val, $opt) 'INFO'
                    }
                    $created++
                    continue
                }
                $newId = New-SugDeployment -SugCiId ([int]$sugId) -CollectionId $t.Id -Name $t.DeployName -S $S
                [void]$script:AssignmentNames.Add($t.DeployName)
                Write-Log "[$($t.Id)] Deployment oluşturuldu. AssignmentName='$($t.DeployName)', AssignmentID=$newId" 'OK'
                $created++
            } catch {
                $failed++
                Write-Log "[$($t.Id)] Deployment oluşturulamadı: $(Get-ErrorDetail $_)" 'ERROR'
            }
        }

        if (-not $simulate -and $created -gt 0) {
            try { Update-AllDeployments } catch { Write-Log "Deploy bilgisi yenilenemedi: $($_.Exception.Message)" 'WARN' }
        }
        Save-CollectionCache
        Show-Collections

        if ($simulate) {
            Show-Message ("Simülasyon tamamlandı: {0} koleksiyon için deployment oluşturulacaktı, {1} atlandı. Ayrıntılar günlükte." -f $created, $skip.Count)
        } elseif ($failed -gt 0) {
            Show-Message ("Oluşturulan: {0}  |  Hatalı: {1}  |  Atlanan: {2}. Ayrıntılar günlükte." -f $created, $failed, $skip.Count) 'Uyarı' 'Warning'
        } else {
            Show-Message ("{0} deployment oluşturuldu, {1} atlandı." -f $created, $skip.Count)
        }
    }
    catch {
        Write-Log (Get-ErrorDetail $_) 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    }
    finally {
        Set-Busy $false
    }
}

# ------------------------------------------------------------------
# Olaylar
# ------------------------------------------------------------------
$btnClose.Add_Click({ $form.Close() })
$btnDeploy.Add_Click({ Invoke-Deploy })

$btnReload.Add_Click({ Import-StateFile (Find-StateFile) })
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'state.json|state.json|JSON dosyaları (*.json)|*.json|Tüm dosyalar (*.*)|*.*'
    if ($dlg.ShowDialog() -eq 'OK') { Import-StateFile $dlg.FileName }
})

# State kaydı seçilince: site/sunucu ve o kaydın SUG'ı varsayılan olarak seçilir
$cmbState.Add_SelectedIndexChanged({
    $i = $cmbState.SelectedIndex
    if ($i -ge 0 -and $i -lt $script:StateEntries.Count) {
        $e = $script:StateEntries[$i]
        if ($e.Site)   { $txtSite.Text = $e.Site }
        if ($e.Server) { $txtServer.Text = $e.Server }
        if ($txtSite.Text.Trim().ToUpper() -ne $script:CacheSite) { Import-CollectionCache }
        Select-SugById $e.SugId
        Show-Collections
        Update-Preview
    }
})

$txtSite.Add_Leave({
    if ($txtSite.Text.Trim().ToUpper() -ne $script:CacheSite) {
        Import-CollectionCache
        Show-Collections
        Update-Preview
    }
})

$cmbSug.Add_SelectedIndexChanged({
    if ($script:Suspend) { return }
    Show-Collections
    Update-Preview
})

$btnSugReset.Add_Click({
    $i = $cmbState.SelectedIndex
    if ($i -ge 0 -and $i -lt $script:StateEntries.Count) { Select-SugById $script:StateEntries[$i].SugId }
    elseif ($script:StateEntries.Count -gt 0) { Select-SugById $script:StateEntries[0].SugId }
    else { Write-Log 'Dönülecek bir state SUG''ı yok.' 'WARN' }
})

$btnLoadSug.Add_Click({ Invoke-LoadSugs })

$btnOpenLog.Add_Click({
    $f = Get-LogFilePath
    try {
        if (Test-Path -LiteralPath $f) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('/select,"{0}"' -f $f) }
        elseif (Test-Path -LiteralPath $script:LogFolderPath) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $script:LogFolderPath) }
        else { Show-Message ('Günlük klasörü henüz oluşmadı: {0}' -f $script:LogFolderPath) 'Bilgi' 'Information' }
    } catch { Show-Message $_.Exception.Message 'Hata' 'Error' }
})

$btnAddIds.Add_Click({
    $p = Split-CollectionIds $txtIds.Text
    if ($p.Bad.Count -gt 0) { Write-Log ("Geçersiz Collection ID yoksayıldı: {0} (8 karakter, harf/rakam olmalı)" -f ($p.Bad -join ', ')) 'WARN' }
    if ($p.Valid.Count -eq 0) { Show-Message 'Geçerli bir Collection ID girilmedi.' 'Uyarı' 'Warning'; return }
    Set-Busy $true
    try {
        $n = Add-CollectionsById $p.Valid
        if ($n -gt 0) {
            Save-CollectionCache
            try { Update-AllDeployments } catch { Write-Log "Deploy bilgisi okunamadı: $($_.Exception.Message)" 'WARN' }
            $txtIds.Text = ''
        }
        Show-Collections
        Update-Preview
        Write-Log "$n koleksiyon listeye eklendi / güncellendi."
    }
    catch {
        Write-Log (Get-ErrorDetail $_) 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    }
    finally { Set-Busy $false }
})
$txtIds.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq 'Return') { $e.SuppressKeyPress = $true; $btnAddIds.PerformClick() }
})

$btnFile.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Metin/CSV (*.txt;*.csv)|*.txt;*.csv|Tüm dosyalar (*.*)|*.*'
    if ($dlg.ShowDialog() -ne 'OK') { return }
    try {
        $text = Get-Content -LiteralPath $dlg.FileName -Raw -ErrorAction Stop
        $txtIds.Text = (($text -split '[\r\n]+' | Where-Object { $_.Trim() }) -join ', ')
        $p = Split-CollectionIds $txtIds.Text
        Write-Log ("Dosyadan {0} geçerli Collection ID okundu ({1} geçersiz değer yoksayılacak). Eklemek için 'Listeye Ekle'ye basın." -f $p.Valid.Count, $p.Bad.Count)
    } catch {
        Write-Log "Dosya okunamadı: $($_.Exception.Message)" 'ERROR'
    }
})

$btnAll.Add_Click({
    Set-Busy $true
    try {
        [void](Get-Connection)
        Write-Log 'Ortamdaki tüm cihaz koleksiyonları okunuyor...'
        $rows = @(Invoke-SmsQuery 'SELECT * FROM SMS_Collection WHERE CollectionType = 2')
        foreach ($r in $rows) {
            $id = ([string]$r.CollectionID).ToUpper()
            $script:Cols[$id] = ConvertTo-CollectionRecord $r (Get-Col $id)
        }
        Save-CollectionCache
        Write-Log "$($rows.Count) cihaz koleksiyonu cache'e alındı."
        try { Update-AllDeployments } catch { Write-Log "Deploy bilgisi okunamadı: $($_.Exception.Message)" 'WARN' }
        Show-Collections
    }
    catch {
        Write-Log (Get-ErrorDetail $_) 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    }
    finally { Set-Busy $false }
})

$btnRefresh.Add_Click({
    if ($script:Cols.Count -eq 0) { Show-Message 'Cache boş; önce koleksiyon ekleyin.' 'Bilgi' 'Information'; return }
    Set-Busy $true
    try {
        [void](Get-Connection)
        Update-AllDeployments
        Show-Collections
    }
    catch {
        Write-Log (Get-ErrorDetail $_) 'ERROR'
        Show-Message $_.Exception.Message 'Hata' 'Error'
    }
    finally { Set-Busy $false }
})

$btnRemoveCache.Add_Click({
    $sel = @()
    foreach ($r in $gridCol.SelectedRows) { if ($r.Tag) { $sel += [string]$r.Tag } }
    if ($sel.Count -eq 0) { Show-Message 'Önce cache''den çıkarılacak satırları seçin.' 'Bilgi' 'Information'; return }
    $ans = [System.Windows.Forms.MessageBox]::Show(("{0} koleksiyon cache'den çıkarılacak (SCCM'de hiçbir şey silinmez). Devam edilsin mi?" -f $sel.Count), "Cache'den çıkar", 'YesNo', 'Question', 'Button2')
    if ($ans -ne 'Yes') { return }
    foreach ($id in $sel) {
        [void]$script:Cols.Remove($id)
        $script:Checked.Remove($id)
    }
    Save-CollectionCache
    Show-Collections
    Update-Preview
    Write-Log "$($sel.Count) koleksiyon cache'den çıkarıldı."
})

$txtColFilter.Add_TextChanged({ Show-Collections })
$chkOnlyNew.Add_CheckedChanged({ Show-Collections })

$btnMarkAll.Add_Click({
    foreach ($r in $gridCol.Rows) { if ($r.Tag) { $script:Checked[[string]$r.Tag] = $true } }
    Show-Collections
    Update-Preview
})
$btnUnmark.Add_Click({
    $script:Checked = @{}
    Show-Collections
    Update-Preview
})

# Onay kutusunda tıklamanın hemen işlenmesi + işaretin saklanması
$gridCol.Add_CurrentCellDirtyStateChanged({
    if ($gridCol.IsCurrentCellDirty) {
        [void]$gridCol.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})
$gridCol.Add_CellValueChanged({
    param($sender, $e)
    if ($script:Suspend -or $e.RowIndex -lt 0 -or $e.ColumnIndex -ne 0) { return }
    $row = $gridCol.Rows[$e.RowIndex]
    if (-not $row.Tag) { return }
    $script:Checked[[string]$row.Tag] = ($row.Cells['cChk'].Value -eq $true)
    $n = 0
    foreach ($k in @($script:Checked.Keys)) { if ($script:Checked[$k]) { $n++ } }
    $lblColCount.Text = ('Cache: {0} koleksiyon  |  İşaretli (deploy): {1}  |  Yeşil: seçili SUG zaten deploy edilmiş' -f $script:Cols.Count, $n)
    Update-Preview
})
$gridCol.Add_SelectionChanged({ if (-not $script:Suspend) { Show-Details } })
$gridCol.Add_DataError({ param($sender, $e) $e.ThrowException = $false })
$gridDet.Add_DataError({ param($sender, $e) $e.ThrowException = $false })

# Ayarlar
$btnDefaults.Add_Click({ Set-DeploySettingsUi (Get-DefaultSettings); Update-Preview; Write-Log 'Deploy ayarları Program.cs varsayılanlarına döndürüldü.' })
$cmbType.Add_SelectedIndexChanged({ Update-SettingsState; Update-Preview })
$cmbStart.Add_SelectedIndexChanged({ Update-SettingsState })
$cmbDead.Add_SelectedIndexChanged({ Update-SettingsState })
foreach ($c in @($chkDpLocal, $chkDpRemote, $chkMetered, $chkWumu)) { $c.Add_CheckedChanged({ Update-SettingsState; Update-Preview }) }
$txtNameFmt.Add_TextChanged({ Update-Preview })

# ------------------------------------------------------------------
# Açılış
# ------------------------------------------------------------------
Set-DeploySettingsUi (Get-DefaultSettings)

Write-Log 'Hazır. Deploy SUG''ı kontrol edin, koleksiyonları ekleyip işaretleyin. Varsayılan olarak SİMÜLASYON açıktır.'
Import-StateFile (Find-StateFile)
Update-LogPathUi
Write-Log "Günlük dosyası: $(Get-LogFilePath)"
if ($txtSite.Text.Trim().ToUpper() -ne $script:CacheSite) { Import-CollectionCache }
Show-Collections
Update-Preview

# Başlangıç boyutu: tasarım boyutu ekrana sığmıyorsa çalışma alanına göre kırpılır.
# Yerleşim asgari alanın (1100x860) altına inerse form kaydırma çubukları gösterir; tümünü sığdırmaya çalışmayız.
$wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$capW = [Math]::Min(1280, $wa.Width - 20)
$capH = [Math]::Min(940, $wa.Height - 60)    # başlık çubuğu + kenarlık payı
$form.ClientSize = New-Object System.Drawing.Size($capW, $capH)

$script:WheelFilter = New-Object FormWheelScrollFilter($form)
[System.Windows.Forms.Application]::AddMessageFilter($script:WheelFilter)
$form.Add_FormClosed({ [System.Windows.Forms.Application]::RemoveMessageFilter($script:WheelFilter) })

# Form göründükten sonra ortamdaki SUG'lar otomatik yüklenir (bağlantı bilgisi state.json'dan geldiyse).
$form.Add_Shown({
    if ($txtSite.Text.Trim() -and $txtServer.Text.Trim()) { Invoke-LoadSugs -Quiet }
    else { Write-Log 'Site/SMS Provider boş olduğu için ortamdaki SUG''lar yüklenmedi. state.json yükleyin veya alanları doldurup "Ortamdaki SUG''ları Yükle"ye basın.' 'WARN' }
})

[void]$form.ShowDialog()
