param(
    [string]$Configuration = "Release",
    [string]$Platform = "AnyCPU",
    [string]$Version = "",
    [string]$BuildToolsPath = "D:\Tools\BuildTools",
    [string]$PublishDir = "",
    [string]$DistDir = "",
    [switch]$SkipInstaller
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$projectPath = Join-Path $repoRoot "SlideSCI\SlideSCI.csproj"

if (-not $Version) {
    [xml]$project = Get-Content $projectPath
    $Version = $project.Project.PropertyGroup |
        Where-Object { $_.ApplicationVersion } |
        Select-Object -First 1 -ExpandProperty ApplicationVersion
}

if (-not $Version) {
    $Version = "1.0.0.0"
}

if (-not $PublishDir) {
    $PublishDir = Join-Path $repoRoot "artifacts\publish"
}

if (-not $DistDir) {
    $DistDir = Join-Path $repoRoot "artifacts\dist"
}

function Find-MSBuild {
    param([string]$PreferredBuildToolsPath)

    $candidates = @()

    if ($PreferredBuildToolsPath) {
        $candidates += Join-Path $PreferredBuildToolsPath "MSBuild\Current\Bin\MSBuild.exe"
        $candidates += Join-Path $PreferredBuildToolsPath "MSBuild\Current\Bin\amd64\MSBuild.exe"
    }

    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $installationPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
        if ($installationPath) {
            $candidates += Join-Path $installationPath "MSBuild\Current\Bin\MSBuild.exe"
            $candidates += Join-Path $installationPath "MSBuild\Current\Bin\amd64\MSBuild.exe"
        }
    }

    $msbuild = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $msbuild) {
        throw "MSBuild.exe was not found. Install Visual Studio Build Tools first."
    }

    return $msbuild
}

function Assert-VstoTargets {
    param(
        [string]$MSBuildPath,
        [string]$VSToolsPath
    )

    $target = $null

    if ($VSToolsPath) {
        $candidate = Join-Path $VSToolsPath "OfficeTools\Microsoft.VisualStudio.Tools.Office.targets"
        if (Test-Path $candidate) {
            $target = Get-Item $candidate
        }
    }

    if (-not $target) {
        $msbuildRoot = Split-Path (Split-Path (Split-Path $MSBuildPath -Parent) -Parent) -Parent
        $target = Get-ChildItem $msbuildRoot -Recurse -Filter "Microsoft.VisualStudio.Tools.Office.targets" -ErrorAction SilentlyContinue |
            Select-Object -First 1
    }

    if (-not $target) {
        throw @"
VSTO OfficeTools targets were not found.

Install the Office/SharePoint build tools workload, then rerun this script:

  & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\setup.exe" modify --installPath "$BuildToolsPath" --add Microsoft.VisualStudio.Workload.OfficeBuildTools --includeRecommended --passive --norestart --wait

If Visual Studio Installer does not modify the existing Build Tools instance, open it manually and add:
  Office/SharePoint build tools
"@
    }

    return Split-Path (Split-Path $target.FullName -Parent) -Parent
}

function Find-InnoCompiler {
    $candidates = @(
        "D:\Tools\InnoSetup\ISCC.exe",
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
        "${env:LOCALAPPDATA}\Programs\Inno Setup 6\ISCC.exe"
    )

    $command = Get-Command "ISCC.exe" -ErrorAction SilentlyContinue
    if ($command) {
        $candidates = @($command.Source) + $candidates
    }

    return $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
}

function Find-VstoUtilitiesDirectory {
    $candidates = @(
        "${env:ProgramFiles}\Microsoft Office\root\Office16\ADDINS\EduWorks Data Streamer Add-In\Microsoft.Office.Tools.Common.v4.0.Utilities.dll",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\ADDINS\EduWorks Data Streamer Add-In\Microsoft.Office.Tools.Common.v4.0.Utilities.dll"
    )

    $dll = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $dll) {
        $dll = Get-ChildItem "${env:ProgramFiles}", "${env:ProgramFiles(x86)}" `
            -Recurse `
            -Filter "Microsoft.Office.Tools.Common.v4.0.Utilities.dll" `
            -ErrorAction SilentlyContinue |
            Select-Object -First 1 -ExpandProperty FullName
    }

    if (-not $dll) {
        throw "Microsoft.Office.Tools.Common.v4.0.Utilities.dll was not found. Install the VSTO runtime or Office developer tools."
    }

    return Split-Path $dll -Parent
}

function Ensure-OfficeTargetsPackage {
    $packageId = "MSBuild.Microsoft.VisualStudio.Tools.Office.targets"
    $packageVersion = "15.0.1"
    $packageRoot = Join-Path $repoRoot "packages\$packageId.$packageVersion"
    $vstoToolsPath = Join-Path $packageRoot "tools\VSToolsPath"
    $target = Join-Path $vstoToolsPath "OfficeTools\Microsoft.VisualStudio.Tools.Office.targets"

    if (Test-Path $target) {
        return $vstoToolsPath
    }

    Write-Host "Downloading $packageId $packageVersion..."
    $tmp = Join-Path $env:TEMP "$packageId.$packageVersion.nupkg"
    Invoke-WebRequest -Uri "https://www.nuget.org/api/v2/package/$packageId/$packageVersion" -OutFile $tmp

    if (Test-Path $packageRoot) {
        Remove-Item $packageRoot -Recurse -Force
    }

    New-Item -ItemType Directory -Force -Path $packageRoot | Out-Null
    Expand-Archive -Path $tmp -DestinationPath $packageRoot -Force

    if (-not (Test-Path $target)) {
        throw "Downloaded $packageId, but $target was not found."
    }

    return $vstoToolsPath
}

function Ensure-Net472ReferenceAssemblies {
    $packageId = "Microsoft.NETFramework.ReferenceAssemblies.net472"
    $packageVersion = "1.0.3"
    $packageRoot = Join-Path $repoRoot "packages\$packageId.$packageVersion"
    $referenceRoot = Join-Path $packageRoot "build"
    $mscorlib = Join-Path $referenceRoot ".NETFramework\v4.7.2\mscorlib.dll"

    if (Test-Path $mscorlib) {
        return $referenceRoot
    }

    Write-Host "Downloading $packageId $packageVersion..."
    $tmp = Join-Path $env:TEMP "$packageId.$packageVersion.nupkg"
    Invoke-WebRequest -Uri "https://www.nuget.org/api/v2/package/$packageId/$packageVersion" -OutFile $tmp

    if (Test-Path $packageRoot) {
        Remove-Item $packageRoot -Recurse -Force
    }

    New-Item -ItemType Directory -Force -Path $packageRoot | Out-Null
    Expand-Archive -Path $tmp -DestinationPath $packageRoot -Force

    if (-not (Test-Path $mscorlib)) {
        throw "Downloaded $packageId, but $mscorlib was not found."
    }

    return $referenceRoot
}

function Ensure-ClickOnceCertificate {
    $subject = "CN=SlideSCI ClickOnce Temporary"
    $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
        Where-Object { $_.Subject -eq $subject -and $_.NotAfter -gt (Get-Date).AddDays(30) } |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1

    if (-not $cert) {
        Write-Host "Creating temporary ClickOnce signing certificate..."
        $cert = New-SelfSignedCertificate `
            -Type CodeSigningCert `
            -Subject $subject `
            -CertStoreLocation Cert:\CurrentUser\My `
            -KeyExportPolicy Exportable `
            -NotAfter (Get-Date).AddYears(5)
    }

    return $cert.Thumbprint
}

$msbuild = Find-MSBuild -PreferredBuildToolsPath $BuildToolsPath
$vstoToolsPath = Ensure-OfficeTargetsPackage
$vstoToolsPath = Assert-VstoTargets -MSBuildPath $msbuild -VSToolsPath $vstoToolsPath
$targetFrameworkRootPath = Ensure-Net472ReferenceAssemblies
$vstoUtilitiesDirectory = Find-VstoUtilitiesDirectory
$clickOnceCertificateThumbprint = Ensure-ClickOnceCertificate

New-Item -ItemType Directory -Force -Path $PublishDir | Out-Null
New-Item -ItemType Directory -Force -Path $DistDir | Out-Null

$defaultAppPublish = Join-Path $repoRoot "SlideSCI\bin\$Configuration\app.publish"
if (Test-Path $PublishDir) {
    Remove-Item (Join-Path $PublishDir "*") -Recurse -Force -ErrorAction SilentlyContinue
}

if (Test-Path $defaultAppPublish) {
    Remove-Item $defaultAppPublish -Recurse -Force
}

$publishDirWithSlash = [System.IO.Path]::GetFullPath($PublishDir)
if (-not $publishDirWithSlash.EndsWith("\")) {
    $publishDirWithSlash += "\"
}

& $msbuild $projectPath `
    /t:Restore,Build,Publish `
    /p:Configuration=$Configuration `
    "/p:Platform=$Platform" `
    /p:VSToolsPath=$vstoToolsPath `
    /p:TargetFrameworkRootPath=$targetFrameworkRootPath `
    /p:ReferencePath=$vstoUtilitiesDirectory `
    /p:SignAssembly=false `
    /p:AssemblyOriginatorKeyFile= `
    /p:SignManifests=true `
    /p:ManifestKeyFile= `
    /p:ManifestCertificateThumbprint=$clickOnceCertificateThumbprint `
    /p:BootstrapperEnabled=false `
    /p:PublishUrl=$publishDirWithSlash `
    /p:ApplicationVersion=$Version `
    /v:m

if ($LASTEXITCODE -ne 0) {
    throw "MSBuild failed with exit code $LASTEXITCODE."
}

if ((Test-Path $defaultAppPublish) -and ((Get-ChildItem $PublishDir -Force -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0)) {
    Copy-Item (Join-Path $defaultAppPublish "*") -Destination $PublishDir -Recurse -Force
}

if ($SkipInstaller) {
    Write-Host "Published VSTO files to $PublishDir"
    exit 0
}

$iscc = Find-InnoCompiler
if (-not $iscc) {
    Write-Warning "Inno Setup compiler (ISCC.exe) was not found. VSTO files were published to $PublishDir, but no single-file installer was created."
    Write-Warning "Install Inno Setup 6, then rerun this script."
    exit 0
}

$env:SLIDESCI_PUBLISH_DIR = $publishDirWithSlash
$env:SLIDESCI_DIST_DIR = [System.IO.Path]::GetFullPath($DistDir)
$env:SLIDESCI_VERSION = $Version

$innoScript = Join-Path $repoRoot "installer\SlideSCI.iss"
& $iscc $innoScript

if ($LASTEXITCODE -ne 0) {
    throw "Inno Setup failed with exit code $LASTEXITCODE."
}

Write-Host "Release artifacts:"
Get-ChildItem $DistDir -File | Select-Object FullName, Length
