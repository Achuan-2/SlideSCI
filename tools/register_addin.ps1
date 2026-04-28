
# 手动注册 SlideSCICompat 插件到 PowerPoint 和 WPS
# 运行前请确认 D:\SlideSCI_WPS_PowerPoint_Compat\SlideSCICompat.vsto 文件存在

$installDir = "D:\SlideSCI_WPS_PowerPoint_Compat"
$vstoFile = Join-Path $installDir "SlideSCICompat.vsto"
$addInKey = "SlideSCICompat"

if (-not (Test-Path $vstoFile)) {
    Write-Error "VSTO 文件不存在: $vstoFile"
    exit 1
}

# 构造 file:/// URL（斜杠格式）
$manifestUrl = "file:///" + $vstoFile.Replace("\", "/") + "|vstolocal"

Write-Host "Manifest URL: $manifestUrl"

# 1. 注册到 PowerPoint Addins（让 PowerPoint 和 WPS 都能识别）
$pptAddinKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$addInKey"
if (-not (Test-Path $pptAddinKey)) {
    New-Item -Path $pptAddinKey -Force | Out-Null
}
Set-ItemProperty -Path $pptAddinKey -Name "Description"  -Value "SlideSCI WPS PowerPoint Compat"
Set-ItemProperty -Path $pptAddinKey -Name "FriendlyName" -Value "SlideSCI"
Set-ItemProperty -Path $pptAddinKey -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $pptAddinKey -Name "Manifest"     -Value $manifestUrl
Write-Host "✅ PowerPoint Addins 注册完成"

# 2. 注册到 WPS AddinsWL（WPS 需要额外的白名单注册）
$wpsKey = "HKCU:\Software\Kingsoft\Office\WPP\AddinsWL"
if (-not (Test-Path $wpsKey)) {
    New-Item -Path $wpsKey -Force | Out-Null
}
Set-ItemProperty -Path $wpsKey -Name $addInKey -Value ""
Write-Host "✅ WPS AddinsWL 注册完成"

# 3. VSTO 安全信任注册（从 .vsto 文件中提取公钥）
$vstoContent = Get-Content $vstoFile -Raw
$startTag = "<RSAKeyValue>"
$endTag   = "</RSAKeyValue>"
$startIdx = $vstoContent.IndexOf($startTag)
$endIdx   = $vstoContent.IndexOf($endTag)

if ($startIdx -ge 0 -and $endIdx -gt $startIdx) {
    $publicKey = $vstoContent.Substring($startIdx, $endIdx - $startIdx + $endTag.Length)
    $solutionId  = "{EDE5B327-B8B0-4044-9237-768C42B63E3E}"
    $inclusionId = "{80B4B921-FA89-4AAE-8146-62F13CCC93E4}"
    $metaRoot    = "HKCU:\Software\Microsoft\VSTO\SolutionMetadata"
    $solutionKey = "$metaRoot\$solutionId"
    $inclusionKey = "HKCU:\Software\Microsoft\VSTO\Security\Inclusion\$inclusionId"
    $manifestUrlNoLocal = $manifestUrl.Replace("|vstolocal", "")

    # Inclusion (trust)
    if (-not (Test-Path $inclusionKey)) { New-Item -Path $inclusionKey -Force | Out-Null }
    Set-ItemProperty -Path $inclusionKey -Name "Url"       -Value $manifestUrlNoLocal
    Set-ItemProperty -Path $inclusionKey -Name "PublicKey" -Value $publicKey
    Write-Host "✅ VSTO Security Inclusion 注册完成"

    # SolutionMetadata
    if (-not (Test-Path $metaRoot)) { New-Item -Path $metaRoot -Force | Out-Null }
    Set-ItemProperty -Path $metaRoot -Name $manifestUrlNoLocal -Value $solutionId

    if (-not (Test-Path $solutionKey)) { New-Item -Path $solutionKey -Force | Out-Null }
    Set-ItemProperty -Path $solutionKey -Name "addInName"          -Value $addInKey
    Set-ItemProperty -Path $solutionKey -Name "officeApplication"  -Value "PowerPoint"
    Set-ItemProperty -Path $solutionKey -Name "friendlyName"       -Value "SlideSCI"
    Set-ItemProperty -Path $solutionKey -Name "description"        -Value $addInKey
    Set-ItemProperty -Path $solutionKey -Name "loadBehavior"       -Value 3 -Type DWord
    Set-ItemProperty -Path $solutionKey -Name "compatibleFrameworks" -Value '<compatibleFrameworks xmlns="urn:schemas-microsoft-com:clickonce.v2"><framework targetVersion="4.7.2" profile="Full" supportedRuntime="4.0.30319" /></compatibleFrameworks>'
    Write-Host "✅ VSTO SolutionMetadata 注册完成"
} else {
    Write-Warning "无法从 .vsto 文件中提取公钥，跳过 VSTO 信任注册"
}

Write-Host ""
Write-Host "====================================="
Write-Host "注册完成！请重新启动 WPS 或 PowerPoint"
Write-Host "====================================="
