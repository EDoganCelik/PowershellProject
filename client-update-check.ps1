<#
.SYNOPSIS
    SCCM Run Script ile calistirmak uzere, bir istemcinin kumulatif Windows
    guncellemesini neden yukleyemedigini tespit etmeye yonelik tanilama betigi.

.PARAMETER KBArticleID
    (Opsiyonel) Sorun yasanan KB numarasi (orn: KB5034441). Verilirse bu
    guncellemenin sistemde yuklu olup olmadigi da kontrol edilir. SCCM
    Run Script ekraninda parametre olarak tanimlayip calisma anda girebilirsin.

.NOTES
    - Cikti tek satirda ";" ile ayrilmis olarak Write-Output ile basilir.
    - SCCM Run Script varsayilan zaman asimi kisadir (surum bagimli, genelde
      60 sn - birkac dk arasi). DISM /CheckHealth hizli calisir ama yine de
      betigi olustururken timeout suresini rahat tutman (orn. 5 dk) onerilir.
    - Betik SYSTEM baglaminda calisir, bu yuzden HKLM ve servis sorgulari
      icin ek yetki gerekmez.
#>




function Get-SafeValue {
    param($ScriptBlock, $Default = "N/A")
    try {
        $result = & $ScriptBlock
        if ($null -eq $result -or $result -eq "") { return $Default }
        return $result
    } catch {
        return $Default
    }
}




function Get-ServiceStatus {
    param($Name)
    Get-SafeValue { (Get-Service -Name $Name -ErrorAction Stop).Status }
}

$Hostname = $env:COMPUTERNAME

# 2) Sistem surucusunde bos alan (CU kurulumu icin genelde birkac GB gerekir)
$SysDrive = $env:SystemDrive
$DiskInfo = Get-SafeValue {
    $d = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$SysDrive'"
    "{0:N2}" -f ($d.FreeSpace / 1GB)
}
$DiskPercent = Get-SafeValue {
    $d = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$SysDrive'"
    "{0:N1}" -f (($d.FreeSpace / $d.Size) * 100)
}

# 4) Kritik servislerin durumu
$WuauservStatus        = Get-ServiceStatus "wuauserv"
$BITSStatus             = Get-ServiceStatus "BITS"
$CryptSvcStatus         = Get-ServiceStatus "CryptSvc"
$TrustedInstallerStatus = Get-ServiceStatus "TrustedInstaller"

# 8) SoftwareDistribution klasoru boyutu (asiri sismesi sorun isareti olabilir)
$SoftDistSize = Get-SafeValue {
    $path = "$env:SystemRoot\SoftwareDistribution"
    if (Test-Path $path) {
        $size = (Get-ChildItem $path -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        "{0:N1}" -f ($size / 1MB)
    } else { "0" }
}



# 10) DISM ile bilesen deposu saglik kontrolu (hizli mod: CheckHealth)
$DismCheckHealth = Get-SafeValue {
    $result = & dism.exe /Online /Cleanup-Image /CheckHealth 2>&1
    if ($result -match "No component store corruption detected") { "Saglikli" }
    elseif ($result -match "repairable") { "Onarilabilir_Bozulma" }
    else { "Belirsiz" }
} "Hata"


# 16) DataTransferService*.log icinde 4xx/5xx (40*/50*) hata kodu var mi?
#     Rotasyon sonucu olusan tum dosyalar (DataTransferService.log, DataTransferService1.log,
#     DataTransferService.lo_ vb.) taranir. Bulunan tum farkli kodlar tek alanda listelenir.
$DTSDownloadErrors = "Indirme hatasi bulunamadi"
try {
    $dtsLogFiles = Get-ChildItem -Path "$env:SystemRoot\CCM\Logs" -Filter "DataTransferService*.log" -File -ErrorAction SilentlyContinue

    $dtsCodes = @()
    foreach ($dtsLogFile in $dtsLogFiles) {
        $dtsLines = Get-Content -Path $dtsLogFile.FullName -ErrorAction SilentlyContinue
        $dtsCodes += ($dtsLines | Select-String -Pattern '\b([45]\d{2})\b' -AllMatches |
            ForEach-Object { $_.Matches } | ForEach-Object { $_.Groups[1].Value })
    }

    $dtsUniqueCodes = $dtsCodes | Sort-Object -Unique
    if ($dtsUniqueCodes.Count -gt 0) {
        $DTSDownloadErrors = $dtsUniqueCodes -join ","
    }
} catch {
    $DTSDownloadErrors = "Hata"
}

# ---- Ciktiyi olustur ----
$values = @(
    $Hostname,
    $DiskInfo,
    $DiskPercent,
    $WuauservStatus,
    $BITSStatus,
    $CryptSvcStatus,
    $TrustedInstallerStatus,
    $SoftDistSize,
    $DismCheckHealth,
    $DTSDownloadErrors
)

$output = $values -join ";"
Write-Output $output