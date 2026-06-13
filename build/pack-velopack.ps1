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
    $toolsPath = Join-Path $env:USERPROFILE ".dotnet\tools"
    if ((Test-Path -LiteralPath $toolsPath) -and ($env:PATH -notlike "*$toolsPath*")) {
        $env:PATH = "$env:PATH;$toolsPath"
    }

    # vpk CLI は動作確認済みの固定バージョンを使う。無固定取得 (dotnet tool install --global vpk) だと
    # 署名・パッケージングを担う最特権ツールが NuGet 侵害や予期せぬ最新版で差し替わるサプライチェーン穴になる。
    # Velopack ライブラリ (PackageReference) は prerelease 0.0.x に固定されているが、vpk CLI はそれとは
    # 別管理の安定版でよい (vpk は後方互換で旧ランタイム向けパッケージも pack できる)。
    # 更新時はこの値を手動で上げ、pack→配信→クライアント更新まで検証してから確定する。
    $vpkVersion = "1.2.0"

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

# publish 成果物は毎回クリーンに作り直す。前回の .deploy 等の残骸が残ると、
# 増分 publish で混ざってマニフェスト記載の論理名と物理ファイル名がずれる原因になる。
if (Test-Path $vstoPublishDir) {
    Remove-Item -Path $vstoPublishDir -Recurse -Force
}
New-Item -ItemType Directory -Path $vstoPublishDir -Force | Out-Null

$msbuild = Get-MSBuildPath
$vstoPublishDirWithSlash = $vstoPublishDir.TrimEnd('\') + '\'

# 1. VSTO アセンブリをビルドする (publish はまだ行わない)。
$buildArgs = @(
    $vstoProject,
    "/t:Restore,Rebuild",
    "/p:Configuration=$Configuration",
    "/p:Platform=AnyCPU",
    "/p:ManifestCertificateThumbprint=$VstoSigningThumbprint"
)
& $msbuild @buildArgs
if ($LASTEXITCODE -ne 0) {
    throw "VSTO build failed with exit code $LASTEXITCODE"
}

# 2. publish 前に VSTO アセンブリ本体を Authenticode 署名する。
#    vpk は vsto\ サブフォルダを署名対象から外すため、ここで署名しておかないと
#    更新パッケージ内の EXLSXS.dll(.deploy) が未署名になりクライアントの信頼検証が失敗する。
#    署名済み DLL から .deploy と ClickOnce application manifest が生成されるので、
#    Authenticode 署名と manifest ハッシュの両方が整合する。
$vstoAssembly = Join-Path (Split-Path -Parent $vstoProject) "bin\$Configuration\EXLSXS.dll"
if (-not [string]::IsNullOrWhiteSpace($VelopackSignParams)) {
    if (-not (Test-Path -LiteralPath $vstoAssembly)) {
        throw "VSTO assembly was not found for signing: $vstoAssembly"
    }
    Write-Host "Authenticode-signing VSTO assembly: $vstoAssembly"
    Invoke-SignTool -SignParams $VelopackSignParams -FilePath $vstoAssembly
}

# 3. 署名済み DLL を publish する (PublishOnly は Build を再実行せず bin の成果物を使う)。
$publishArgs = @(
    $vstoProject,
    "/t:PublishOnly",
    "/p:Configuration=$Configuration",
    "/p:Platform=AnyCPU",
    "/p:PublishDir=$vstoPublishDirWithSlash",
    "/p:PublishUrl=$vstoPublishDirWithSlash",
    # vstolocal でその場ロードするため .deploy 拡張子を付けない。csproj の PropertyGroup 値は
    # VSTO の publish ターゲットに上書きされるので、ターゲットから変更不可なグローバルプロパティ
    # (コマンドライン /p:) として渡す必要がある。これが無いと EXLSXS.dll.deploy が出力され、
    # vstolocal ロードが論理名 EXLSXS.dll を見つけられず FileNotFound でアドインが読み込めない。
    "/p:MapFileExtensions=false",
    "/p:ManifestCertificateThumbprint=$VstoSigningThumbprint"
)
& $msbuild @publishArgs
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

# vstolocal 登録は「.vsto と同じフォルダ」を AppBase にするため、staging は ClickOnce の
# ネスト構成 (Application Files\<name>_<ver>\) ではなく、.vsto / .dll.manifest / EXLSXS.dll /
# 依存 DLL を全部同一フォルダに並べるフラット構成にする (動作実績のある Data Streamer と同じ)。
# ネストのまま親 .vsto を登録すると Assembly.Load("EXLSXS") がサブフォルダを探せず
# FileNotFoundException でアドインが読み込めない (2026-06 実機で確認)。
#   - publish 内側フォルダ: 署名済み EXLSXS.dll + フラット参照の EXLSXS.dll.manifest + 依存 DLL
#   - bin\Release の EXLSXS.vsto: codebase がフラット参照 (publish ルートの .vsto はネスト参照なので使わない)
#   - setup.exe: 前提条件ブートストラッパー (publish ルートから)
$vstoInnerDir = Get-ChildItem -Path (Join-Path $vstoPublishDir "Application Files") -Directory | Select-Object -First 1
if ($null -eq $vstoInnerDir) {
    throw "VSTO publish output did not contain an 'Application Files' folder: $vstoPublishDir"
}
Copy-Item -Path (Join-Path $vstoInnerDir.FullName "*") -Destination $vstoStagingDir -Force
$flatDeploymentManifest = Join-Path (Split-Path -Parent $vstoProject) "bin\$Configuration\EXLSXS.vsto"
if (-not (Test-Path -LiteralPath $flatDeploymentManifest)) {
    throw "Flat deployment manifest was not found: $flatDeploymentManifest"
}
Copy-Item -Path $flatDeploymentManifest -Destination (Join-Path $vstoStagingDir "EXLSXS.vsto") -Force
Copy-Item -Path (Join-Path $vstoPublishDir "setup.exe") -Destination $vstoStagingDir -Force

# 前提条件ブートストラッパー (vsto\setup.exe) も Authenticode 署名する。
# vpk は vsto\ サブフォルダを署名対象から外すため、ここで署名しないと setup.exe が未署名のまま残り、
# クライアントが起動時 (PrerequisiteChecker) に署名者を検証できず、差し替えられた setup.exe が
# UAC 昇格で実行される余地を残す。本証明書で署名しておけば起動時検証を通過する。
$stagedSetupExe = Join-Path $vstoStagingDir "setup.exe"
if (-not [string]::IsNullOrWhiteSpace($VelopackSignParams)) {
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
Reset-Directory -Path $releaseDir

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
