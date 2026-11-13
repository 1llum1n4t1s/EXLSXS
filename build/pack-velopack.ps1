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

    # PFX パスワードは平文 [string] で受け取らない (プロセス引数・PSReadLine 履歴・transcription に
    # 平文で残るため)。SecureString パラメータか、CI secret 由来の環境変数経由でのみ受け取る。
    [Parameter(Mandatory = $false)]
    [securestring]$VstoSigningPfxPassword,

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
    [switch]$RequireVelopackSigning,

    # Sigilune (Certum SimplySign のクラウド鍵で署名する CLI、1.1.0 以降) の sigilune.exe。
    # 指定すると Authenticode 署名と ClickOnce マニフェスト署名をすべて Sigilune で行い、
    # 証明書ストアの証明書・PFX・signtool の引数は使わない (GitHub Actions でのリリース用)。
    [Parameter(Mandatory = $false)]
    [string]$SigiluneExe,

    # vpk の出力先。既定はリポジトリの artifacts\velopack-releases (前回までの配布物を差分の基準に使う)。
    [Parameter(Mandatory = $false)]
    [string]$ReleaseDirectory,

    # 出力先に前回の配布物が無いとき (GitHub Actions の新しいランナーなど)、配信元から最新版を取得して
    # 差分パッケージ (delta) とリリース一覧の基準にする。
    [Parameter(Mandatory = $false)]
    [switch]$DownloadPreviousRelease
)




# 検証済みの作業対象だけを標準APIで削除する。
function Remove-LocalWorkItem {
    param([string[]]$LiteralPath, [string]$AllowedRoot, [switch]$IgnoreMissing, [switch]$AllowLeafJunction)
    if ([string]::IsNullOrWhiteSpace($AllowedRoot) -or -not [IO.Path]::IsPathRooted($AllowedRoot)) { throw 'Cleanup root must be absolute.' }
    $boundary = [IO.Path]::GetFullPath($AllowedRoot).TrimEnd('\','/') + [IO.Path]::DirectorySeparatorChar
    foreach ($target in $LiteralPath) {
        if ([string]::IsNullOrWhiteSpace($target) -or -not [IO.Path]::IsPathRooted($target)) { throw 'Cleanup target must be absolute.' }
        $resolved = [IO.Path]::GetFullPath($target)
        if (-not $resolved.StartsWith($boundary, [StringComparison]::OrdinalIgnoreCase)) { throw 'Cleanup target is outside the allowed root.' }
        $ancestor = $resolved
        while ($ancestor) {
            if (Test-Path -LiteralPath $ancestor) {
                $item = Get-Item -LiteralPath $ancestor -Force -ErrorAction Stop
                if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -and -not ($AllowLeafJunction -and $ancestor -eq $resolved -and $item.LinkType -eq 'Junction')) { throw 'Cleanup path contains a reparse point.' }
            }
            $ancestor = Split-Path $ancestor -Parent
        }
        if (-not (Test-Path -LiteralPath $resolved)) { if ($IgnoreMissing) { continue }; throw "Cleanup target is missing: $resolved" }
        $item = Get-Item -LiteralPath $resolved -Force -ErrorAction Stop
        $leafLink = [bool]($item.Attributes -band [IO.FileAttributes]::ReparsePoint)
        if ($item.PSIsContainer -and -not $leafLink) {
            if (@(Get-ChildItem -LiteralPath $resolved -Recurse -Force -ErrorAction Stop | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }).Count) { throw 'Cleanup directory contains a reparse point.' }
        }
        Remove-Item -LiteralPath $resolved -Recurse:($item.PSIsContainer -and -not $leafLink) -Force -ErrorAction Stop
        if (Test-Path -LiteralPath $resolved) { throw "Cleanup target remains: $resolved" }
    }
}

$ErrorActionPreference = "Stop"

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$vstoProject = Join-Path $root "EXLSXS\EXLSXS.csproj"
$hostProject = Join-Path $root "EXLSXS.Host\EXLSXS.Host.csproj"
# vpk CLI は Velopack ライブラリ (PackageReference) と同じ、動作確認済みの安定版へ固定する。
# 無固定取得は署名・パッケージングを担う最特権ツールが NuGet 侵害や予期せぬ最新版で
# 差し替わるサプライチェーン穴になる。更新時はライブラリとこの値を同じ batch で上げ、
# pack→配信→クライアント更新まで検証してから確定する。
$vpkVersion = "1.2.161"
# 中間出力は実行ごとの外部Temp、配布物は従来の正式保存先に置く。
$taskId = [Guid]::NewGuid().ToString('N')
$scratchParent = [IO.Path]::GetFullPath((Join-Path ([IO.Path]::GetTempPath()) 'Codex\EXLSXS'))
$artifactsDir = Join-Path $scratchParent $taskId
$vstoPublishDir = Join-Path $artifactsDir 'vsto-publish'
$hostPublishDir = Join-Path $artifactsDir 'host-publish'
$stagingDir = Join-Path $artifactsDir 'velopack-staging'
$vstoBuildDir = Join-Path $artifactsDir 'vsto-build'
$vstoObjDir = Join-Path $artifactsDir 'vsto-obj'
$releaseDir = if ([string]::IsNullOrWhiteSpace($ReleaseDirectory)) { Join-Path $root 'artifacts\velopack-releases' } else { [IO.Path]::GetFullPath($ReleaseDirectory) }

function Assert-NoReparseAncestor {
    param([Parameter(Mandatory = $true)][string]$Path)
    $ancestor = [IO.Path]::GetFullPath($Path)
    while ($ancestor) {
        if (Test-Path -LiteralPath $ancestor) {
            $item = Get-Item -LiteralPath $ancestor -Force -ErrorAction Stop
            if (-not $item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                throw "Output path has a non-directory or reparse ancestor: $ancestor"
            }
        }
        $ancestor = Split-Path -Parent $ancestor
    }
}

function Assert-ExternalOutputRoot {
    param([Parameter(Mandatory = $true)][string]$Path)
    $absolute = [IO.Path]::GetFullPath($Path).TrimEnd([char[]]'\/')
    $forbiddenRoots = @($root, (Join-Path ([Environment]::GetFolderPath('UserProfile')) '.codex'))
    if (-not [string]::IsNullOrWhiteSpace($env:CODEX_HOME)) { $forbiddenRoots += [IO.Path]::GetFullPath($env:CODEX_HOME) }
    foreach ($forbiddenRoot in $forbiddenRoots) {
        $forbidden = [IO.Path]::GetFullPath($forbiddenRoot).TrimEnd([char[]]'\/')
        if ($absolute.Equals($forbidden, [StringComparison]::OrdinalIgnoreCase) -or $absolute.StartsWith($forbidden + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw "Intermediate output and cleanup ledger must be outside the repository and Codex user directory: $absolute"
        }
    }
    Assert-NoReparseAncestor -Path $absolute
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

# VSTO のビルドターゲットは $(VisualStudioVersion) の Office ツールを読み込む。csproj の既定は 18.0 (VS 2026) なので、
# VS 2022 しか無い環境 (GitHub の windows-2022 ランナーなど) でも動くよう、実際に入っている版を渡す。
function Get-VisualStudioVersion {
    $vsWhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path -LiteralPath $vsWhere)) { return $null }
    $installed = & $vsWhere -latest -requires Microsoft.Component.MSBuild -property installationVersion | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($installed)) { return $null }
    return $installed.Split('.')[0] + '.0'
}

function Get-SignToolPath {
    $kitsRoot = Join-Path ${env:ProgramFiles(x86)} "Windows Kits\10\bin"
    if (Test-Path -LiteralPath $kitsRoot) {
        $signtool = Get-ChildItem -LiteralPath $kitsRoot -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^10\.' } |
            Sort-Object Name -Descending |
            ForEach-Object { Join-Path $_.FullName "x64\signtool.exe" } |
            Where-Object { Test-Path -LiteralPath $_ } |
            Select-Object -First 1
        if ($signtool) { return $signtool }
    }

    $command = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($command -ne $null) { return $command.Source }

    throw "signtool.exe was not found. Install the Windows 10/11 SDK to enable Authenticode signing."
}

function Invoke-SignTool {
    param(
        [Parameter(Mandatory = $true)][string]$SignParams,
        [Parameter(Mandatory = $true)][string]$FilePath
    )

    $signtool = Get-SignToolPath
    # VelopackSignParams (例: /n "..." /fd SHA256 /td SHA256 /tr http://time.certum.pl) を
    # そのまま signtool sign の引数へ展開する。クォート保持のため単一引数文字列で渡す。
    $argumentList = "sign $SignParams `"$FilePath`""
    $process = Start-Process -FilePath $signtool -ArgumentList $argumentList -NoNewWindow -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        throw "signtool sign failed for '$FilePath' (exit $($process.ExitCode))."
    }
}

# AGENTS.md の「Velopack ライブラリと vpk CLI は同じ安定版に揃える」を機械的に守る。
# Dependabot はホスト csproj の Velopack だけを上げ、$vpkVersion は追随しないため、
# 片方だけ上がった状態で長いビルドを始める前に止める。
function Assert-VelopackVersionsMatch {
    $hostXml = [xml](Get-Content -LiteralPath $hostProject -Raw)
    $libraryVersions = @($hostXml.SelectNodes('//PackageReference[@Include="Velopack"]') | ForEach-Object { $_.GetAttribute('Version') })
    if ($libraryVersions.Count -ne 1 -or [string]::IsNullOrWhiteSpace($libraryVersions[0])) {
        throw "Velopack PackageReference version could not be read from $hostProject."
    }
    if ($libraryVersions[0] -ne $vpkVersion) {
        throw "Velopack library ($($libraryVersions[0]), $hostProject) and vpk CLI (`$vpkVersion = $vpkVersion, build/pack-velopack.ps1) must be updated together."
    }
}

function Ensure-VelopackCli {
    $toolsPath = Join-Path $env:USERPROFILE ".dotnet\tools"
    if ((Test-Path -LiteralPath $toolsPath) -and ($env:PATH -notlike "*$toolsPath*")) {
        $env:PATH = "$env:PATH;$toolsPath"
    }

    # 固定版はスクリプト冒頭の $vpkVersion (Assert-VelopackVersionsMatch で照合済み) を使う。
    $installedVersion = $null
    $command = Get-Command vpk -ErrorAction SilentlyContinue
    if ($command -ne $null) {
        $listLine = (dotnet tool list --global 2>$null) | Where-Object { $_ -match '^\s*vpk\s' } | Select-Object -First 1
        if ($listLine -and ($listLine -match '^\s*vpk\s+(\S+)')) {
            $installedVersion = $matches[1]
        }
    }

    if ($installedVersion -eq $vpkVersion) {
        return
    }

    if ($command -ne $null) {
        dotnet tool update --global vpk --version $vpkVersion
    }
    else {
        dotnet tool install --global vpk --version $vpkVersion
    }
    if ($LASTEXITCODE -ne 0) {
        throw "vpk CLI ($vpkVersion) のインストール/更新に失敗しました (exit $LASTEXITCODE)"
    }
}

function Import-VstoSigningCertificate {
    param(
        [Parameter(Mandatory = $false)][AllowEmptyString()][string]$PfxPath,
        [Parameter(Mandatory = $false)][securestring]$PfxPassword,
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

    if ($null -eq $PfxPassword) {
        throw "VSTO signing PFX password is required when a PFX is supplied."
    }

    $certificate = Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation Cert:\CurrentUser\My -Password $PfxPassword
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

$scratchCreated = $false
$temporaryManifestCertificate = $null
$packFailure = $null
$cleanupFailure = $null
$registrationCleanupFailure = $null
try {
    Assert-VelopackVersionsMatch
    Assert-ExternalOutputRoot -Path $artifactsDir
    Assert-NoReparseAncestor -Path $releaseDir
    if (Test-Path -LiteralPath $artifactsDir) { throw "Intermediate output already exists: $artifactsDir" }
    New-Item -ItemType Directory -Path $artifactsDir -ErrorAction Stop | Out-Null
    $scratchCreated = $true
    Write-Host "  Intermediate output: $artifactsDir"
if ([string]::IsNullOrWhiteSpace($Version)) {
    # バージョンの単一ソースはリポジトリ直下の Directory.Build.props の <Version>。
    $directoryBuildProps = Join-Path $root "Directory.Build.props"
    $Version = Get-ProjectValue -Path $directoryBuildProps -Name "Version"
    if ([string]::IsNullOrWhiteSpace($Version)) {
        throw "Version was not found in $directoryBuildProps"
    }
}

$Version = Convert-ToPackVersion -RawVersion $Version
$assemblyVersion = Convert-ToAssemblyVersion -PackVersion $Version

if ([string]::IsNullOrWhiteSpace($UpdateSource)) {
    # ホスト側のパッケージ信頼契約と同じ既定値にする。GitHub は配信元として許可されていない。
    $UpdateSource = "https://exlsxs.kagayoi.com"
    $UpdateSourceKind = "Simple"
}

Write-Host "Publishing EXLSXS VSTO and host for Velopack" -ForegroundColor Cyan
Write-Host "  Version: $Version"
Write-Host "  AssemblyVersion: $assemblyVersion"
Write-Host "  Channel: $Channel"
Write-Host "  RuntimeIdentifier: $RuntimeIdentifier"
Write-Host "  UpdateSource: $UpdateSource"

# 実行専用の空のTempを使用するため、既存artifactsは削除しない。
$VstoSigningPfxPath = Get-ValueOrEnvironment -Value $VstoSigningPfxPath -EnvironmentName "EXLSXS_VSTO_SIGNING_PFX_PATH"
# PFX パスワードはパラメータ未指定時のみ環境変数 (CI secret マスク前提) から読み、即 SecureString 化する。
if ($null -eq $VstoSigningPfxPassword) {
    $pfxPasswordFromEnv = [System.Environment]::GetEnvironmentVariable("EXLSXS_VSTO_SIGNING_PFX_PASSWORD")
    if (-not [string]::IsNullOrWhiteSpace($pfxPasswordFromEnv)) {
        $VstoSigningPfxPassword = ConvertTo-SecureString -String $pfxPasswordFromEnv -AsPlainText -Force
        $pfxPasswordFromEnv = $null
    }
}
$VstoSigningThumbprint = Get-ValueOrEnvironment -Value $VstoSigningThumbprint -EnvironmentName "EXLSXS_VSTO_SIGNING_THUMBPRINT"
$VelopackSignParams = Get-ValueOrEnvironment -Value $VelopackSignParams -EnvironmentName "EXLSXS_VELOPACK_SIGN_PARAMS"
$VelopackAzureTrustedSignFile = Get-ValueOrEnvironment -Value $VelopackAzureTrustedSignFile -EnvironmentName "EXLSXS_VELOPACK_AZURE_TRUSTED_SIGN_FILE"
$ExpectedPublisherThumbprint = Get-ValueOrEnvironment -Value $ExpectedPublisherThumbprint -EnvironmentName "EXLSXS_EXPECTED_PUBLISHER_THUMBPRINT"
$ExpectedPublisherSubject = Get-ValueOrEnvironment -Value $ExpectedPublisherSubject -EnvironmentName "EXLSXS_EXPECTED_PUBLISHER_SUBJECT"
$SigiluneExe = Get-ValueOrEnvironment -Value $SigiluneExe -EnvironmentName "EXLSXS_SIGILUNE_EXE"
$useSigilune = -not [string]::IsNullOrWhiteSpace($SigiluneExe)
$hasExplicitExpectedPublisherThumbprint = -not [string]::IsNullOrWhiteSpace($ExpectedPublisherThumbprint)

if ($useSigilune) {
    if (-not [string]::IsNullOrWhiteSpace($VelopackSignParams) -or -not [string]::IsNullOrWhiteSpace($VelopackAzureTrustedSignFile)) {
        throw "Sigilune signing cannot be combined with EXLSXS_VELOPACK_SIGN_PARAMS or EXLSXS_VELOPACK_AZURE_TRUSTED_SIGN_FILE."
    }
    . (Join-Path $root 'scripts\Sigilune.Common.ps1')
    # ビルド前にログインして、資格情報の不備を早く検出する (失敗しても自動再試行しない)。
    Connect-Sigilune -SigiluneExe $SigiluneExe
    $sigiluneCertificate = Get-SigiluneCertificate -SigiluneExe $SigiluneExe
    if ($hasExplicitExpectedPublisherThumbprint -and (Normalize-Thumbprint -Thumbprint $ExpectedPublisherThumbprint) -ne $sigiluneCertificate.Thumbprint) {
        throw "The Sigilune signing certificate ($($sigiluneCertificate.Thumbprint)) does not match the expected publisher thumbprint."
    }
    if ([string]::IsNullOrWhiteSpace($ExpectedPublisherThumbprint)) { $ExpectedPublisherThumbprint = $sigiluneCertificate.Thumbprint }
    if ([string]::IsNullOrWhiteSpace($ExpectedPublisherSubject)) { $ExpectedPublisherSubject = $sigiluneCertificate.Subject }
    Write-Host "  Signing: Sigilune ($($sigiluneCertificate.Subject), $($sigiluneCertificate.Thumbprint))"

    # MSBuild の VSTO publish は証明書ストアの証明書でマニフェストを署名する前提なので、この実行だけの
    # 使い捨て自己署名証明書でビルドし、ステージング後に Sigilune がクラウド鍵で署名し直す
    # (既存署名の除去と publicKeyToken・publisherIdentity の置換は sign-manifest が行う)。
    $temporaryManifestCertificate = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=EXLSXS temporary build certificate" `
        -CertStoreLocation Cert:\CurrentUser\My -HashAlgorithm SHA256 -NotAfter (Get-Date).AddDays(1)
    $VstoSigningThumbprint = Normalize-Thumbprint -Thumbprint $temporaryManifestCertificate.Thumbprint
}
else {
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

if ($RequireVelopackSigning -and [string]::IsNullOrWhiteSpace($VelopackSignParams) -and [string]::IsNullOrWhiteSpace($VelopackAzureTrustedSignFile)) {
    throw "Velopack signing is required, but neither EXLSXS_VELOPACK_SIGN_PARAMS nor EXLSXS_VELOPACK_AZURE_TRUSTED_SIGN_FILE was supplied."
}
}

$ExpectedPublisherThumbprint = Normalize-Thumbprint -Thumbprint $ExpectedPublisherThumbprint

# publish 成果物は毎回クリーンに作り直す。前回の .deploy 等の残骸が残ると、
# 増分 publish で混ざってマニフェスト記載の論理名と物理ファイル名がずれる原因になる。

New-Item -ItemType Directory -Path $vstoPublishDir -Force | Out-Null

$msbuild = Get-MSBuildPath
$visualStudioVersion = Get-VisualStudioVersion
$visualStudioVersionArgs = @(if ($visualStudioVersion) { "/p:VisualStudioVersion=$visualStudioVersion" })
$vstoPublishDirWithSlash = $vstoPublishDir.TrimEnd('\') + '\'

# VSTO の Rebuild は PrepareForRun → RegisterOfficeAddin で HKCU の Excel アドイン登録を
# 出力先 ($vstoBuildDir = 実行ごとの一時フォルダー) の .vsto へ書き換えてしまう。一時フォルダーは
# パック後に削除されるため、そのままだと開発機の Excel 起動時に「EXLSXS.vsto のダウンロードに
# 失敗しました」となる。ビルド前の登録状態を退避し、MSBuild 実行後に必ず元へ戻す。
$addInRegistryPath = 'HKCU:\Software\Microsoft\Office\Excel\Addins\EXLSXS'

function Save-AddInRegistration {
    if (-not (Test-Path -LiteralPath $addInRegistryPath)) { return $null }
    $key = Get-Item -LiteralPath $addInRegistryPath
    $values = @{}
    foreach ($name in $key.GetValueNames()) {
        $values[$name] = @{ Value = $key.GetValue($name, $null, 'DoNotExpandEnvironmentNames'); Kind = $key.GetValueKind($name) }
    }
    return $values
}

function Restore-AddInRegistration {
    param([AllowNull()][hashtable]$Snapshot)
    if ($null -eq $Snapshot) {
        # ビルド前に登録が無かった場合は、MSBuild が作ったキーを取り除く（開発機を汚さない）。
        if (Test-Path -LiteralPath $addInRegistryPath) {
            Remove-Item -LiteralPath $addInRegistryPath -Recurse -Force -ErrorAction Stop
        }
        return
    }
    if (-not (Test-Path -LiteralPath $addInRegistryPath)) {
        New-Item -Path $addInRegistryPath -Force | Out-Null
    }
    $key = Get-Item -LiteralPath $addInRegistryPath
    # MSBuild が追加した値を消し、退避した値を種類ごと書き戻す。
    foreach ($name in $key.GetValueNames()) {
        if (-not $Snapshot.ContainsKey($name)) {
            Remove-ItemProperty -LiteralPath $addInRegistryPath -Name $name -ErrorAction Stop
        }
    }
    foreach ($name in $Snapshot.Keys) {
        $entry = $Snapshot[$name]
        New-ItemProperty -LiteralPath $addInRegistryPath -Name $name -Value $entry.Value -PropertyType $entry.Kind -Force | Out-Null
    }
}

# VisualStudioForApplicationsBuild ターゲットは SetInclusionListEntry で、ビルド出力の .vsto を
# HKCU の VSTO 信頼リスト (Security\Inclusion\<GUID>) に無条件で追加する。pack は実行ごとに
# 使い捨て証明書と一時フォルダーを使うため、そのままでは存在しないパスの信頼項目が毎回溜まる。
# ビルド前に無かったサブキーのうち、Url がこの実行の .vsto に完全一致するものだけを消す
# (開発用の bin\Debug やインストール版の項目、ビルド前からある項目には触れない)。
# 同じビルドで SolutionMetadata にも「.vsto の URL = {GUID}」の値と {GUID} サブキーが作られる
# (実機の pack 前後比較で確認) ため、同じ条件で一緒に消す。
$vstoInclusionRegistryPath = 'HKCU:\Software\Microsoft\VSTO\Security\Inclusion'
$vstoSolutionMetadataRegistryPath = 'HKCU:\Software\Microsoft\VSTO\SolutionMetadata'

function Test-VstoSolutionMetadataExists {
    param([Parameter(Mandatory = $true)][string]$ManifestUrl)
    if (-not (Test-Path -LiteralPath $vstoSolutionMetadataRegistryPath)) { return $false }
    return @((Get-Item -LiteralPath $vstoSolutionMetadataRegistryPath).GetValueNames()) -contains $ManifestUrl
}

function Remove-BuildVstoSolutionMetadata {
    param([Parameter(Mandatory = $true)][string]$ManifestUrl)
    if (-not (Test-Path -LiteralPath $vstoSolutionMetadataRegistryPath)) { return }
    $metadataKey = Get-Item -LiteralPath $vstoSolutionMetadataRegistryPath
    $valueName = @($metadataKey.GetValueNames()) | Where-Object { $_.Equals($ManifestUrl, [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1
    if ($null -eq $valueName) { return }
    $solutionId = $metadataKey.GetValue($valueName)
    Remove-ItemProperty -LiteralPath $vstoSolutionMetadataRegistryPath -Name $valueName -ErrorAction Stop
    # GUID 形式の値だけを対象にし、想定外の値でサブキーを取り違えない。
    if ($solutionId -is [string] -and $solutionId -match '^\{[0-9A-Fa-f-]{36}\}$') {
        $solutionKeyPath = Join-Path $vstoSolutionMetadataRegistryPath $solutionId
        if (Test-Path -LiteralPath $solutionKeyPath) {
            Remove-Item -LiteralPath $solutionKeyPath -Recurse -Force -ErrorAction Stop
        }
    }
}

function Get-VstoInclusionKeyNames {
    if (-not (Test-Path -LiteralPath $vstoInclusionRegistryPath)) { return @() }
    return @(Get-ChildItem -LiteralPath $vstoInclusionRegistryPath | ForEach-Object { $_.PSChildName })
}

function Remove-BuildVstoInclusionEntries {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$ExistingKeyNames,
        [Parameter(Mandatory = $true)][string]$ManifestUrl
    )
    if (-not (Test-Path -LiteralPath $vstoInclusionRegistryPath)) { return }
    foreach ($subKey in @(Get-ChildItem -LiteralPath $vstoInclusionRegistryPath)) {
        if ($ExistingKeyNames -contains $subKey.PSChildName) { continue }
        $url = $subKey.GetValue('Url')
        if ($url -is [string] -and $url.Equals($ManifestUrl, [StringComparison]::OrdinalIgnoreCase)) {
            Remove-Item -LiteralPath $subKey.PSPath -Recurse -Force -ErrorAction Stop
        }
    }
}

$addInRegistrationSnapshot = Save-AddInRegistration
# 関数から返した空配列は呼び出し側で $null になる (まっさらな CI ランナー等で Inclusion が空のとき)。
# $null は [string[]] の必須引数に束縛できないため、@() で包んで空配列のまま渡す。
$vstoInclusionKeyNamesBeforeBuild = @(Get-VstoInclusionKeyNames)
$vstoBuildManifestUrl = ([Uri](Join-Path $vstoBuildDir 'EXLSXS.vsto')).AbsoluteUri
# 実行ごとに新しい一時フォルダーなので通常は false。ビルド前からあれば触れない。
$vstoSolutionMetadataExistedBeforeBuild = Test-VstoSolutionMetadataExists -ManifestUrl $vstoBuildManifestUrl
try {

# 1. VSTO アセンブリをビルドする (publish はまだ行わない)。
$buildArgs = @(
    $vstoProject,
    "/t:Restore,Rebuild",
    "/p:Configuration=$Configuration",
    "/p:Platform=AnyCPU",
    "/p:OutputPath=$vstoBuildDir\",
    "/p:BaseIntermediateOutputPath=$vstoObjDir\",
    "/p:MSBuildProjectExtensionsPath=$vstoObjDir\",
    "/p:ManifestCertificateThumbprint=$VstoSigningThumbprint"
) + $visualStudioVersionArgs
& $msbuild @buildArgs
if ($LASTEXITCODE -ne 0) {
    throw "VSTO build failed with exit code $LASTEXITCODE"
}

# 2. publish 前に VSTO アセンブリ本体を Authenticode 署名する。
#    vpk は vsto\ サブフォルダを署名対象から外すため、ここで署名しておかないと
#    更新パッケージ内の EXLSXS.dll(.deploy) が未署名になりクライアントの信頼検証が失敗する。
#    次の工程で署名後の DLL から ClickOnce マニフェストを再生成する。
$vstoAssembly = Join-Path $vstoBuildDir 'EXLSXS.dll'
if ($useSigilune) {
    Write-Host "Authenticode-signing VSTO assembly with Sigilune: $vstoAssembly"
    Invoke-SigiluneSign -SigiluneExe $SigiluneExe -Path $vstoAssembly -CertificateSubjectName $ExpectedPublisherSubject
}
elseif (-not [string]::IsNullOrWhiteSpace($VelopackSignParams)) {
    if (-not (Test-Path -LiteralPath $vstoAssembly)) {
        throw "VSTO assembly was not found for signing: $vstoAssembly"
    }
    Write-Host "Authenticode-signing VSTO assembly: $vstoAssembly"
    Invoke-SignTool -SignParams $VelopackSignParams -FilePath $vstoAssembly
}

# 3. DLL を再ビルドせず、署名後の内容でマニフェストを再生成して publish する。
#    PublishOnly 単独では Rebuild 時点の未署名 DLL のハッシュが残る。
$publishArgs = @(
    $vstoProject,
    "/t:ResolveReferences,VisualStudioForApplicationsBuild,PublishOnly",
    "/p:Configuration=$Configuration",
    "/p:Platform=AnyCPU",
    "/p:PublishDir=$vstoPublishDirWithSlash",
    "/p:PublishUrl=$vstoPublishDirWithSlash",
    "/p:OutputPath=$vstoBuildDir\",
    "/p:BaseIntermediateOutputPath=$vstoObjDir\",
    "/p:MSBuildProjectExtensionsPath=$vstoObjDir\",
    # vstolocal でその場ロードするため .deploy 拡張子を付けない。csproj の PropertyGroup 値は
    # VSTO の publish ターゲットに上書きされるので、ターゲットから変更不可なグローバルプロパティ
    # (コマンドライン /p:) として渡す必要がある。これが無いと EXLSXS.dll.deploy が出力され、
    # vstolocal ロードが論理名 EXLSXS.dll を見つけられず FileNotFound でアドインが読み込めない。
    "/p:MapFileExtensions=false",
    "/p:ManifestCertificateThumbprint=$VstoSigningThumbprint"
) + $visualStudioVersionArgs
& $msbuild @publishArgs
if ($LASTEXITCODE -ne 0) {
    throw "VSTO publish failed with exit code $LASTEXITCODE"
}

}
finally {
    # 成功・失敗どちらでも、一時フォルダーを指す登録・信頼項目を残さない。
    # ここで throw するとビルド失敗の原因を上書きしてしまうため、失敗は記録して最後にまとめて報告する。
    try {
        Restore-AddInRegistration -Snapshot $addInRegistrationSnapshot
    }
    catch {
        $registrationCleanupFailure = $_
        Write-Warning "Could not restore the Excel add-in registration ($addInRegistryPath): $($_.Exception.Message)"
    }
    try {
        Remove-BuildVstoInclusionEntries -ExistingKeyNames $vstoInclusionKeyNamesBeforeBuild -ManifestUrl $vstoBuildManifestUrl
    }
    catch {
        if ($null -eq $registrationCleanupFailure) { $registrationCleanupFailure = $_ }
        Write-Warning "Could not remove the build's VSTO inclusion entry ($vstoBuildManifestUrl): $($_.Exception.Message)"
    }
    if (-not $vstoSolutionMetadataExistedBeforeBuild) {
        try {
            Remove-BuildVstoSolutionMetadata -ManifestUrl $vstoBuildManifestUrl
        }
        catch {
            if ($null -eq $registrationCleanupFailure) { $registrationCleanupFailure = $_ }
            Write-Warning "Could not remove the build's VSTO solution metadata ($vstoBuildManifestUrl): $($_.Exception.Message)"
        }
    }
}

New-Item -ItemType Directory -Path $hostPublishDir -ErrorAction Stop | Out-Null
$selfContained = if ($FrameworkDependent) { "false" } else { "true" }
dotnet publish $hostProject `
    --artifacts-path (Join-Path $artifactsDir 'host-build') `
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

New-Item -ItemType Directory -Path $stagingDir -ErrorAction Stop | Out-Null
Copy-Item -Path (Join-Path $hostPublishDir "*") -Destination $stagingDir -Recurse -Force

$vstoStagingDir = Join-Path $stagingDir "vsto"
New-Item -ItemType Directory -Path $vstoStagingDir -Force | Out-Null

# vstolocal 登録は「.vsto と同じフォルダ」を AppBase にするため、staging は ClickOnce の
# ネスト構成 (Application Files\<name>_<ver>\) ではなく、.vsto / .dll.manifest / EXLSXS.dll /
# 依存 DLL を全部同一フォルダに並べるフラット構成にする (動作実績のある Data Streamer と同じ)。
# ネストのまま親 .vsto を登録すると Assembly.Load("EXLSXS") がサブフォルダを探せず
# FileNotFoundException でアドインが読み込めない (2026-06 実機で確認)。
#   - publish 内側フォルダ: 署名済み EXLSXS.dll + フラット参照の EXLSXS.dll.manifest + 依存 DLL
#   - 同じビルド出力の EXLSXS.vsto: codebase がフラット参照 (publish ルートの .vsto はネスト参照なので使わない)
#   - setup.exe: 前提条件ブートストラッパー (publish ルートから)
$vstoInnerDir = Get-ChildItem -Path (Join-Path $vstoPublishDir "Application Files") -Directory | Select-Object -First 1
if ($null -eq $vstoInnerDir) {
    throw "VSTO publish output did not contain an 'Application Files' folder: $vstoPublishDir"
}
Copy-Item -Path (Join-Path $vstoInnerDir.FullName "*") -Destination $vstoStagingDir -Force
$flatDeploymentManifest = Join-Path $vstoBuildDir 'EXLSXS.vsto'
if (-not (Test-Path -LiteralPath $flatDeploymentManifest)) {
    throw "Flat deployment manifest was not found: $flatDeploymentManifest"
}
Copy-Item -Path $flatDeploymentManifest -Destination (Join-Path $vstoStagingDir "EXLSXS.vsto") -Force
Copy-Item -Path (Join-Path $vstoPublishDir "setup.exe") -Destination $vstoStagingDir -Force

$applicationManifest = Join-Path $vstoStagingDir 'EXLSXS.dll.manifest'
if ($useSigilune) {
    # 使い捨て証明書の署名を Certum のクラウド鍵の署名に置き換える。.vsto 内のアプリケーション マニフェスト参照
    # (サイズ・ダイジェスト・publicKeyToken) も sign-manifest が更新するので、下の照合はその結果を検査する。
    Write-Host "Signing VSTO manifests with Sigilune"
    Invoke-SigiluneManifestSign -SigiluneExe $SigiluneExe -ApplicationManifest $applicationManifest `
        -DeploymentManifest (Join-Path $vstoStagingDir 'EXLSXS.vsto') -CertificateSubjectName $ExpectedPublisherSubject
}

# 別ビルドの配置マニフェストが混ざると署名済みでも VSTO が読み込めないため、梱包前に照合する。
[xml]$deploymentXml = Get-Content -LiteralPath (Join-Path $vstoStagingDir 'EXLSXS.vsto') -Raw
[xml]$applicationXml = Get-Content -LiteralPath $applicationManifest -Raw
$rootIdentityPath = "/*[local-name()='assembly']/*[local-name()='assemblyIdentity']"
$deploymentIdentity = $deploymentXml.SelectSingleNode($rootIdentityPath)
$applicationIdentity = $applicationXml.SelectSingleNode($rootIdentityPath)
$applicationReference = $deploymentXml.SelectSingleNode("/*[local-name()='assembly']/*[local-name()='dependency']/*[local-name()='dependentAssembly']")
$referencedIdentity = $applicationReference.SelectSingleNode("*[local-name()='assemblyIdentity']")
$payloadVersion = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $vstoStagingDir 'EXLSXS.dll')).Version.ToString()
if ($deploymentIdentity.GetAttribute('version') -ne $assemblyVersion -or $applicationIdentity.GetAttribute('version') -ne $assemblyVersion -or $payloadVersion -ne $assemblyVersion) {
    throw 'VSTO manifest versions do not match the release version.'
}
foreach ($attribute in @('name', 'version', 'publicKeyToken', 'language', 'processorArchitecture', 'type')) {
    if ($referencedIdentity.GetAttribute($attribute) -cne $applicationIdentity.GetAttribute($attribute)) {
        throw "VSTO deployment and application manifest identities differ: $attribute"
    }
}
$digest = $applicationReference.SelectSingleNode("*[local-name()='hash']/*[local-name()='DigestValue']").InnerText
$digestAlgorithm = $applicationReference.SelectSingleNode("*[local-name()='hash']/*[local-name()='DigestMethod']").GetAttribute('Algorithm')
$algorithm = switch -Regex ($digestAlgorithm) { 'sha256$' { 'SHA256'; break }; 'sha1$' { 'SHA1'; break }; default { throw "Unsupported manifest digest algorithm: $digestAlgorithm" } }
$manifestHash = Get-FileHash -LiteralPath $applicationManifest -Algorithm $algorithm
$actualDigest = [Convert]::ToBase64String([Convert]::FromHexString($manifestHash.Hash))
if ($digest -cne $actualDigest) { throw 'VSTO application manifest digest does not match its deployment reference.' }

$assemblyReference = $applicationXml.SelectSingleNode("//*[local-name()='dependentAssembly' and @codebase='EXLSXS.dll']")
$assemblyPath = Join-Path $vstoStagingDir 'EXLSXS.dll'
if ($null -eq $assemblyReference -or [long]$assemblyReference.GetAttribute('size') -ne (Get-Item -LiteralPath $assemblyPath).Length) {
    throw 'VSTO assembly size does not match its application manifest reference.'
}
$assemblyDigestAlgorithm = $assemblyReference.SelectSingleNode("*[local-name()='hash']/*[local-name()='DigestMethod']").GetAttribute('Algorithm')
$algorithm = switch -Regex ($assemblyDigestAlgorithm) { 'sha256$' { 'SHA256'; break }; 'sha1$' { 'SHA1'; break }; default { throw "Unsupported assembly digest algorithm: $assemblyDigestAlgorithm" } }
$assemblyHash = Get-FileHash -LiteralPath $assemblyPath -Algorithm $algorithm
$actualAssemblyDigest = [Convert]::ToBase64String([Convert]::FromHexString($assemblyHash.Hash))
if ($actualAssemblyDigest -cne $assemblyReference.SelectSingleNode("*[local-name()='hash']/*[local-name()='DigestValue']").InnerText) {
    throw 'VSTO assembly digest does not match its application manifest reference.'
}

# 前提条件ブートストラッパー (vsto\setup.exe) も Authenticode 署名する。
# vpk は vsto\ サブフォルダを署名対象から外すため、ここで署名しないと setup.exe が未署名のまま残り、
# クライアントが起動時 (PrerequisiteChecker) に署名者を検証できず、差し替えられた setup.exe が
# UAC 昇格で実行される余地を残す。本証明書で署名しておけば起動時検証を通過する。
$stagedSetupExe = Join-Path $vstoStagingDir "setup.exe"
if ($useSigilune) {
    Write-Host "Authenticode-signing bootstrapper with Sigilune: $stagedSetupExe"
    # VSTO publish は setup.exe もマニフェスト用の証明書 (ここでは使い捨て証明書) で署名する。
    # その署名だけを置き換え、別の署名者の署名なら想定外として止める。
    $existingSignature = Get-AuthenticodeSignature -LiteralPath $stagedSetupExe
    $replaceSignature = $false
    if ($null -ne $existingSignature.SignerCertificate) {
        if ($existingSignature.SignerCertificate.Thumbprint -ne $temporaryManifestCertificate.Thumbprint) {
            throw "Bootstrapper setup.exe has an unexpected signature: $($existingSignature.SignerCertificate.Subject)"
        }
        $replaceSignature = $true
    }
    Invoke-SigiluneSign -SigiluneExe $SigiluneExe -Path $stagedSetupExe -CertificateSubjectName $ExpectedPublisherSubject -Replace:$replaceSignature
}
elseif (-not [string]::IsNullOrWhiteSpace($VelopackSignParams)) {
    if (-not (Test-Path -LiteralPath $stagedSetupExe)) {
        throw "Bootstrapper setup.exe was not found for signing: $stagedSetupExe"
    }
    Write-Host "Authenticode-signing bootstrapper: $stagedSetupExe"
    Invoke-SignTool -SignParams $VelopackSignParams -FilePath $stagedSetupExe
}

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
# 既存の正式配布物は保持し、vpkに同じ正式出力先を渡す。
New-Item -ItemType Directory -Path $releaseDir -Force | Out-Null
if ($DownloadPreviousRelease -and -not (Test-Path -LiteralPath (Join-Path $releaseDir "releases.$Channel.json"))) {
    if ($UpdateSourceKind -eq 'Github') { throw '-DownloadPreviousRelease supports only an HTTP update source.' }
    Write-Host "Downloading the latest release from $UpdateSource as the delta base"
    vpk download http --url $UpdateSource --channel $Channel --outputDir $releaseDir --allowEmptyChannel
    if ($LASTEXITCODE -ne 0) { throw "vpk download failed with exit code $LASTEXITCODE" }
}

# Setup.exe / Update.exe に付けるアイコン。Host 本体は csproj の ApplicationIcon で
# 同じ app.ico を持つが、Velopack の Setup.exe には vpk pack の --icon を渡さないと付かない。
$packIcon = Join-Path $root "icon\app.ico"
if (-not (Test-Path -LiteralPath $packIcon)) {
    throw "Package icon was not found: $packIcon"
}

$vpkArgs = @(
    "pack",
    "--packId", "EXLSXS",
    "--packVersion", $Version,
    "--packTitle", "EXLSXS",
    "--packAuthors", "EXLSXS",
    "--packDir", $stagingDir,
    "--mainExe", "EXLSXS.Host.exe",
    "--icon", $packIcon,
    "--outputDir", $releaseDir,
    "--channel", $Channel,
    "--shortcuts", "None"
)

# v1.0.11 は公開後に同版の内容が差し替わったため、ベースを一意に決められない。
# v1.0.12 への更新だけは全量配信にし、どちらの v1.0.11 からも更新できるようにする。
if ($Version -eq '1.0.12') {
    $vpkArgs += @('--delta', 'None')
}

if ($useSigilune) {
    # vpk はファイルごとに別プロセスで署名テンプレートを実行し、vsto\ の Microsoft 署名済み DLL も渡してくる。
    # --skip-signed で有効な署名のあるファイルは残し (無効な署名なら失敗)、それ以外を署名する。
    # Sigilune はログインを共有し、署名後に有効性・署名者・タイムスタンプを検証して失敗なら非0で終わる。
    $vpkArgs += @("--signTemplate", "`"$SigiluneExe`" sign --skip-signed {{file}}")
}
elseif (-not [string]::IsNullOrWhiteSpace($VelopackSignParams)) {
    $vpkArgs += @("--signParams", $VelopackSignParams)
}

if (-not [string]::IsNullOrWhiteSpace($VelopackAzureTrustedSignFile)) {
    $vpkArgs += @("--azureTrustedSignFile", $VelopackAzureTrustedSignFile)
}

vpk @vpkArgs

if ($LASTEXITCODE -ne 0) {
    throw "vpk pack failed with exit code $LASTEXITCODE"
}

}
catch {
    $packFailure = $_
}
finally {
    # 使い捨てのビルド用証明書は秘密鍵ごと証明書ストアから取り除く。
    if ($null -ne $temporaryManifestCertificate) {
        try {
            Remove-Item -LiteralPath "Cert:\CurrentUser\My\$($temporaryManifestCertificate.Thumbprint)" -DeleteKey -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not remove the temporary build certificate $($temporaryManifestCertificate.Thumbprint): $($_.Exception.Message)"
        }
    }
    if ($scratchCreated -and (Test-Path -LiteralPath $artifactsDir)) {
        try {
            Remove-LocalWorkItem -LiteralPath $artifactsDir -AllowedRoot $scratchParent
        }
        catch {
            $cleanupFailure = $_
        }
    }
}
if ($null -ne $cleanupFailure) {
    $packMessage = if ($null -ne $packFailure) { " Packaging also failed: $($packFailure.Exception.Message)" } else { '' }
    throw "Intermediate cleanup failed; inspect the retained source. Source: $artifactsDir; Reason: $($cleanupFailure.Exception.Message).$packMessage"
}
if ($null -ne $packFailure) { throw $packFailure }
if ($null -ne $registrationCleanupFailure) {
    throw "Packaging succeeded, but the developer machine's VSTO registration could not be restored: $($registrationCleanupFailure.Exception.Message)"
}
Write-Host "Velopack artifacts created:" -ForegroundColor Green
Get-ChildItem -LiteralPath $releaseDir | ForEach-Object {
    Write-Host "  $($_.FullName)"
}
