param(
    [Parameter(Mandatory = $false)]
    [string]$Version,

    [Parameter(Mandatory = $false)]
    [string]$Channel = "win",

    [Parameter(Mandatory = $false)]
    [string]$UpdateSource,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Auto", "Github", "Simple")]
    [string]$UpdateSourceKind = "Auto",

    [Parameter(Mandatory = $false)]
    [string]$RuntimeIdentifier = "win-x64",

    [Parameter(Mandatory = $false)]
    [string]$Configuration = "Release",

    [Parameter(Mandatory = $false)]
    [switch]$FrameworkDependent,

    [Parameter(Mandatory = $false)]
    [switch]$Prerelease,

    [Parameter(Mandatory = $false)]
    [string]$VstoSigningPfxPath,

    [Parameter(Mandatory = $false)]
    [string]$VstoSigningPfxPassword,

    [Parameter(Mandatory = $false)]
    [string]$VstoSigningThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$VelopackSignParams,

    [Parameter(Mandatory = $false)]
    [string]$VelopackAzureTrustedSignFile,

    [Parameter(Mandatory = $false)]
    [string]$ExpectedPublisherThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$ExpectedPublisherSubject,

    [Parameter(Mandatory = $false)]
    [switch]$RequireVelopackSigning
)

$ErrorActionPreference = "Stop"

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$vstoProject = Join-Path $root "EXLSXS\EXLSXS.csproj"
$hostProject = Join-Path $root "EXLSXS.Host\EXLSXS.Host.csproj"
$artifactsDir = Join-Path $root "artifacts"
$vstoPublishDir = Join-Path $artifactsDir "vsto-publish"
$hostPublishDir = Join-Path $artifactsDir "host-publish"
$stagingDir = Join-Path $artifactsDir "velopack-staging"
$releaseDir = Join-Path $artifactsDir "velopack-releases"

function Assert-InWorkspace {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fullRoot = [System.IO.Path]::GetFullPath($root)
    if (-not $fullPath.StartsWith($fullRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to modify a path outside the workspace: $fullPath"
    }
}

function Reset-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)

    Assert-InWorkspace -Path $Path
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
}

function Get-ProjectValue {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $content = Get-Content -LiteralPath $Path -Raw
    $pattern = "<$Name(?:\s+[^>]*)?>([^<]+)</$Name>"
    if ($content -match $pattern) {
        return $matches[1].Trim()
    }

    return $null
}

function Get-ValueOrEnvironment {
    param(
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$Value,
        [Parameter(Mandatory = $true)][string]$EnvironmentName
    )

    if (-not [string]::IsNullOrWhiteSpace($Value)) {
        return $Value
    }

    $environmentValue = [System.Environment]::GetEnvironmentVariable($EnvironmentName)
    if ([string]::IsNullOrWhiteSpace($environmentValue)) {
        return $null
    }

    return $environmentValue
}

function Normalize-Thumbprint {
    param([Parameter(Mandatory = $false)][AllowEmptyString()][string]$Thumbprint)

    if ([string]::IsNullOrWhiteSpace($Thumbprint)) {
        return $null
    }

    return (($Thumbprint.ToCharArray() | Where-Object { [System.Uri]::IsHexDigit($_) }) -join "").ToUpperInvariant()
}

function Convert-ToPackVersion {
    param([Parameter(Mandatory = $true)][string]$RawVersion)

    $versionCore = ($RawVersion -split '[+-]')[0]
    $suffix = $RawVersion.Substring($versionCore.Length)
    $parts = $versionCore.Split('.')
    if ($parts.Length -eq 4 -and $parts[3] -eq "0") {
        return (($parts[0..2] -join ".") + $suffix)
    }

    return $RawVersion
}

function Convert-ToAssemblyVersion {
    param([Parameter(Mandatory = $true)][string]$PackVersion)

    $versionCore = ($PackVersion -split '[+-]')[0]
    $parts = [System.Collections.Generic.List[string]]::new()
    foreach ($part in $versionCore.Split('.')) {
        $parts.Add($part)
    }
    while ($parts.Count -lt 4) {
        $parts.Add("0")
    }

    return ($parts[0..3] -join ".")
}

function Get-MSBuildPath {
    $vsWhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path -LiteralPath $vsWhere) {
        $found = & $vsWhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
        if (-not [string]::IsNullOrWhiteSpace($found)) {
            return $found
        }
    }

    return "MSBuild.exe"
}

function Get-GitHubRemoteSource {
    $remoteUrl = git -C $root remote get-url origin 2>$null
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($remoteUrl)) {
        return $null
    }

    $remoteUrl = $remoteUrl.Trim()
    if ($remoteUrl -match '^git@github\.com:(?<owner>[^/]+)/(?<repo>.+?)(\.git)?$') {
        return "https://github.com/$($matches.owner)/$($matches.repo -replace '\.git$', '')"
    }

    if ($remoteUrl -match '^https://github\.com/(?<owner>[^/]+)/(?<repo>.+?)(\.git)?$') {
        return "https://github.com/$($matches.owner)/$($matches.repo -replace '\.git$', '')"
    }

    return $null
}

function Ensure-VelopackCli {
    $desiredVersion = "0.0.1369-g1d5c984"
    $toolsPath = Join-Path $env:USERPROFILE ".dotnet\tools"
    if ((Test-Path -LiteralPath $toolsPath) -and ($env:PATH -notlike "*$toolsPath*")) {
        $env:PATH = "$env:PATH;$toolsPath"
    }

    $command = Get-Command vpk -ErrorAction SilentlyContinue
    if ($command -eq $null) {
        dotnet tool install --global vpk --version $desiredVersion
        return
    }

    $installedVersion = $null
    foreach ($line in (dotnet tool list --global)) {
        if ($line -match '^\s*vpk\s+([^\s]+)\s+') {
            $installedVersion = $matches[1]
            break
        }
    }

    if ($installedVersion -ne $desiredVersion) {
        dotnet tool update --global vpk --version $desiredVersion
    }
}

function Import-VstoSigningCertificate {
    param(
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$PfxPath,
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$PfxPassword,
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$PfxBase64
    )

    if ([string]::IsNullOrWhiteSpace($PfxPath) -and -not [string]::IsNullOrWhiteSpace($PfxBase64)) {
        $signingDir = Join-Path $artifactsDir "signing"
        New-Item -ItemType Directory -Path $signingDir -Force | Out-Null
        $PfxPath = Join-Path $signingDir "vsto-signing.pfx"
        [System.IO.File]::WriteAllBytes($PfxPath, [System.Convert]::FromBase64String($PfxBase64))
    }

    if ([string]::IsNullOrWhiteSpace($PfxPath)) {
        return $null
    }

    if (-not (Test-Path -LiteralPath $PfxPath)) {
        throw "VSTO signing PFX was not found: $PfxPath"
    }

    if ([string]::IsNullOrWhiteSpace($PfxPassword)) {
        throw "VSTO signing PFX password is required when a PFX is supplied."
    }

    $securePassword = ConvertTo-SecureString -String $PfxPassword -AsPlainText -Force
    $certificate = Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation Cert:\CurrentUser\My -Password $securePassword
    if ($certificate -eq $null) {
        throw "VSTO signing certificate could not be imported."
    }

    if ($certificate -is [array]) {
        $certificate = $certificate | Select-Object -First 1
    }

    return $certificate.Thumbprint
}

function Assert-CertificateInCurrentUserStore {
    param([Parameter(Mandatory = $true)][string]$Thumbprint)

    $normalized = Normalize-Thumbprint -Thumbprint $Thumbprint
    $certificate = Get-ChildItem Cert:\CurrentUser\My | Where-Object {
        (Normalize-Thumbprint -Thumbprint $_.Thumbprint) -eq $normalized
    } | Select-Object -First 1

    if ($certificate -eq $null) {
        throw "VSTO manifest signing certificate was not found in Cert:\CurrentUser\My: $normalized"
    }

    return $normalized
}

if ([string]::IsNullOrWhiteSpace($Version)) {
    $Version = Get-ProjectValue -Path $vstoProject -Name "ApplicationVersion"
    if ([string]::IsNullOrWhiteSpace($Version)) {
        throw "ApplicationVersion was not found in $vstoProject"
    }
}

$Version = Convert-ToPackVersion -RawVersion $Version
$assemblyVersion = Convert-ToAssemblyVersion -PackVersion $Version

if ([string]::IsNullOrWhiteSpace($UpdateSource) -and -not [string]::IsNullOrWhiteSpace($env:GITHUB_REPOSITORY)) {
    $UpdateSource = "https://github.com/$env:GITHUB_REPOSITORY"
    $UpdateSourceKind = "Github"
}
elseif ([string]::IsNullOrWhiteSpace($UpdateSource)) {
    $remoteSource = Get-GitHubRemoteSource
    if (-not [string]::IsNullOrWhiteSpace($remoteSource)) {
        $UpdateSource = $remoteSource
        $UpdateSourceKind = "Github"
    }
}

Write-Host "Publishing EXLSXS VSTO and host for Velopack" -ForegroundColor Cyan
Write-Host "  Version: $Version"
Write-Host "  AssemblyVersion: $assemblyVersion"
Write-Host "  Channel: $Channel"
Write-Host "  RuntimeIdentifier: $RuntimeIdentifier"
Write-Host "  UpdateSource: $UpdateSource"

Reset-Directory -Path $artifactsDir
$VstoSigningPfxPath = Get-ValueOrEnvironment -Value $VstoSigningPfxPath -EnvironmentName "EXLSXS_VSTO_SIGNING_PFX_PATH"
$VstoSigningPfxPassword = Get-ValueOrEnvironment -Value $VstoSigningPfxPassword -EnvironmentName "EXLSXS_VSTO_SIGNING_PFX_PASSWORD"
$VstoSigningThumbprint = Get-ValueOrEnvironment -Value $VstoSigningThumbprint -EnvironmentName "EXLSXS_VSTO_SIGNING_THUMBPRINT"
$VelopackSignParams = Get-ValueOrEnvironment -Value $VelopackSignParams -EnvironmentName "EXLSXS_VELOPACK_SIGN_PARAMS"
$VelopackAzureTrustedSignFile = Get-ValueOrEnvironment -Value $VelopackAzureTrustedSignFile -EnvironmentName "EXLSXS_VELOPACK_AZURE_TRUSTED_SIGN_FILE"
$ExpectedPublisherThumbprint = Get-ValueOrEnvironment -Value $ExpectedPublisherThumbprint -EnvironmentName "EXLSXS_EXPECTED_PUBLISHER_THUMBPRINT"
$ExpectedPublisherSubject = Get-ValueOrEnvironment -Value $ExpectedPublisherSubject -EnvironmentName "EXLSXS_EXPECTED_PUBLISHER_SUBJECT"
$hasExplicitExpectedPublisherThumbprint = -not [string]::IsNullOrWhiteSpace($ExpectedPublisherThumbprint)

$importedThumbprint = Import-VstoSigningCertificate `
    -PfxPath $VstoSigningPfxPath `
    -PfxPassword $VstoSigningPfxPassword `
    -PfxBase64 ([System.Environment]::GetEnvironmentVariable("EXLSXS_VSTO_SIGNING_PFX_BASE64"))

if ([string]::IsNullOrWhiteSpace($VstoSigningThumbprint)) {
    $VstoSigningThumbprint = $importedThumbprint
}

if ([string]::IsNullOrWhiteSpace($VstoSigningThumbprint)) {
    $VstoSigningThumbprint = Get-ProjectValue -Path $vstoProject -Name "ManifestCertificateThumbprint"
}

$VstoSigningThumbprint = Assert-CertificateInCurrentUserStore -Thumbprint $VstoSigningThumbprint

if ($RequireVelopackSigning -and -not $hasExplicitExpectedPublisherThumbprint) {
    throw "EXLSXS_EXPECTED_PUBLISHER_THUMBPRINT is required when Velopack signing is required."
}

if ([string]::IsNullOrWhiteSpace($ExpectedPublisherThumbprint)) {
    $ExpectedPublisherThumbprint = $VstoSigningThumbprint
}

$ExpectedPublisherThumbprint = Normalize-Thumbprint -Thumbprint $ExpectedPublisherThumbprint

if ($RequireVelopackSigning -and [string]::IsNullOrWhiteSpace($VelopackSignParams) -and [string]::IsNullOrWhiteSpace($VelopackAzureTrustedSignFile)) {
    throw "Velopack signing is required, but neither EXLSXS_VELOPACK_SIGN_PARAMS nor EXLSXS_VELOPACK_AZURE_TRUSTED_SIGN_FILE was supplied."
}

New-Item -ItemType Directory -Path $vstoPublishDir -Force | Out-Null

$msbuild = Get-MSBuildPath
$vstoPublishDirWithSlash = $vstoPublishDir.TrimEnd('\') + '\'
$msbuildArgs = @(
    $vstoProject,
    "/t:Restore,PublishOnly",
    "/p:Configuration=$Configuration",
    "/p:Platform=AnyCPU",
    "/p:PublishDir=$vstoPublishDirWithSlash",
    "/p:PublishUrl=$vstoPublishDirWithSlash",
    "/p:ManifestCertificateThumbprint=$VstoSigningThumbprint"
)
& $msbuild @msbuildArgs

if ($LASTEXITCODE -ne 0) {
    throw "VSTO publish failed with exit code $LASTEXITCODE"
}

Reset-Directory -Path $hostPublishDir
$selfContained = if ($FrameworkDependent) { "false" } else { "true" }
dotnet publish $hostProject `
    -c $Configuration `
    -r $RuntimeIdentifier `
    --self-contained $selfContained `
    -o $hostPublishDir `
    /p:Version=$Version `
    /p:AssemblyVersion=$assemblyVersion `
    /p:FileVersion=$assemblyVersion

if ($LASTEXITCODE -ne 0) {
    throw "Host publish failed with exit code $LASTEXITCODE"
}

Reset-Directory -Path $stagingDir
Copy-Item -Path (Join-Path $hostPublishDir "*") -Destination $stagingDir -Recurse -Force

$vstoStagingDir = Join-Path $stagingDir "vsto"
New-Item -ItemType Directory -Path $vstoStagingDir -Force | Out-Null
Copy-Item -Path (Join-Path $vstoPublishDir "*") -Destination $vstoStagingDir -Recurse -Force

$settings = [ordered]@{
    Update = [ordered]@{
        Source = if ($UpdateSource -eq $null) { "" } else { $UpdateSource }
        SourceKind = $UpdateSourceKind
        Channel = $Channel
        AccessToken = ""
        Prerelease = [bool]$Prerelease
        ExpectedPublisherThumbprint = if ($ExpectedPublisherThumbprint -eq $null) { "" } else { $ExpectedPublisherThumbprint }
        ExpectedPublisherSubject = if ($ExpectedPublisherSubject -eq $null) { "" } else { $ExpectedPublisherSubject }
    }
}
$settingsJson = $settings | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText((Join-Path $stagingDir "appsettings.json"), $settingsJson, [System.Text.Encoding]::UTF8)

Ensure-VelopackCli
Reset-Directory -Path $releaseDir

$vpkArgs = @(
    "pack",
    "--packId", "EXLSXS",
    "--packVersion", $Version,
    "--packTitle", "EXLSXS",
    "--packAuthors", "EXLSXS",
    "--packDir", $stagingDir,
    "--mainExe", "EXLSXS.Host.exe",
    "--outputDir", $releaseDir,
    "--channel", $Channel,
    "--shortcuts", "None"
)

if (-not [string]::IsNullOrWhiteSpace($VelopackSignParams)) {
    $vpkArgs += @("--signParams", $VelopackSignParams)
}

if (-not [string]::IsNullOrWhiteSpace($VelopackAzureTrustedSignFile)) {
    $vpkArgs += @("--azureTrustedSignFile", $VelopackAzureTrustedSignFile)
}

vpk @vpkArgs

if ($LASTEXITCODE -ne 0) {
    throw "vpk pack failed with exit code $LASTEXITCODE"
}

Write-Host "Velopack artifacts created:" -ForegroundColor Green
Get-ChildItem -LiteralPath $releaseDir | ForEach-Object {
    Write-Host "  $($_.FullName)"
}
