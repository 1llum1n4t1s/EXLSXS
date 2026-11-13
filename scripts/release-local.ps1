# release-local.ps1 — ローカル署名付き Velopack リリース (EXLSXS / VSTO アドイン)
#
# 署名は Sigilune (Certum SimplySign のクラウド鍵で署名する CLI、1.1.0 以降) で行う。
# Authenticode 署名に加え、VSTO の ClickOnce マニフェスト (.dll.manifest / .vsto) も同じ鍵で署名する。
# 通常のリリースは GitHub Actions の .github/workflows/release.yml (手動起動・Environment "release") が
# 本スクリプトを実行する。ローカル実行も同じ手順で動き、検証や再開に使う。
#
# EXLSXS は「VSTO アドイン (.NET Fw 4.8.1, Library) + Velopack ホスト」の 2 段構成で、
# ビルド/パッケージングは build/pack-velopack.ps1 が一括 (VSTO publish + ClickOnce 署名 →
# host publish → staging → vpk pack + Authenticode 署名) で担う。本スクリプトはその上に
# プリフライト → pack 呼び出し → 署名検証 → R2 アップロード → キャッシュパージ → 旧版掃除 を被せる。
#
# 前提:
#   - Sigilune を使えること (scripts/Sigilune.Common.ps1 参照)。
#     ローカルは `sigilune setup` で Certum アカウントを登録済み、CI は SIGILUNE_EMAIL / SIGILUNE_TOTP_URI を渡す
#   - Directory.Build.props の <Version> がリリースしたいバージョンであること (/vava 済み)
#   - 環境変数 CLOUDFLARE_API_TOKEN (CI)、なければ D:\Dev\Secret\secrets.json の cloudflare.api_token (ローカル)
#
# 使い方:
#   pwsh scripts/release-local.ps1                # フルリリース (build + sign + upload + cleanup)
#   pwsh scripts/release-local.ps1 -SkipUpload    # ビルド + 署名のみ (アップロードしない動作確認用)
#   pwsh scripts/release-local.ps1 -SkipBuild -ExpectedPackageSha256 <SHA256>
#                                                # 実機検証済みの同一パッケージを再ビルドせず配信

[CmdletBinding()]
param(
    [switch]$SkipUpload,
    [switch]$SkipBuild,
    [string]$ExpectedPackageSha256
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

# R2 への書込みは同梱の PowerShell/.NET 実装で行う。
. (Join-Path $PSScriptRoot 'R2Distribution.Common.ps1')

function Send-ReleaseR2Object {
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)
    # バージョン付き成果物は不変、固定 URL とメタデータは毎回再検証する。
    $immutable = $File.Extension -eq '.nupkg' -or $File.Name -match '\d+\.\d+\.\d+'
    $cacheControl = if ($immutable) { 'public, max-age=31536000, immutable' } else { 'no-cache, must-revalidate' }
    $contentType = switch ($File.Extension.ToLowerInvariant()) {
        '.json' { 'application/json'; break }
        '.exe' { 'application/vnd.microsoft.portable-executable'; break }
        '.zip' { 'application/zip'; break }
        default { if ($File.Name -like 'RELEASES*') { 'text/plain; charset=utf-8' } else { 'application/octet-stream' } }
    }
    $null = Publish-R2Object -AccountId $AccountId -Bucket $Bucket -FilePath $File.FullName `
        -ApiToken $CloudflareApiToken -ContentType $contentType -CacheControl $cacheControl -Immutable:$immutable
}

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ---- 定数 ----
$Bucket          = 'exlsxs-updates'
$BaseUrl         = 'https://exlsxs.kagayoi.com'
$DistributionBaseUrls = @($BaseUrl, 'https://exlsxs.nephilim.jp')
$AccountId       = '10901bfadbf1005164774a7350082985'
$Channel         = 'win'
$SecretsPath     = 'D:\Dev\Secret\secrets.json'
$CertSubjectName = 'Open Source Developer Yuichiro Shinozaki'
# 署名証明書 thumbprint (ClickOnce manifest 署名 + Authenticode 署名者の期待値)。
# クライアントは更新パッケージの署名者をこの値で照合するため、Sigilune の証明書と一致しなければ止める。
$SigningThumbprint = '6285702C9AF1FCFE3D9FE815B7F7F625508130C0'

# 更新パッケージ内で EXLSXS 発行者の Authenticode 署名を要求する必須ファイル。
# Host.exe/Host.dll/EXLSXS.dll はクライアントの UpdatePackageTrustVerifier が適用前に検証する。
# vsto/setup.exe は前提ブートストラッパーで、PrerequisiteChecker が昇格起動の直前に署名者を検証する
# (差し替えた setup.exe の UAC 昇格実行を防ぐ)。いずれも本証明書で署名済みであることを出荷前に保証する。
$RequiredSignedEntries = @(
    [pscustomobject]@{ Label = 'EXLSXS host'; Suffixes = @('EXLSXS.Host.exe') },
    [pscustomobject]@{ Label = 'EXLSXS host assembly'; Suffixes = @('EXLSXS.Host.dll') },
    [pscustomobject]@{ Label = 'EXLSXS VSTO assembly'; Suffixes = @('EXLSXS.dll.deploy', 'EXLSXS.dll') },
    [pscustomobject]@{ Label = 'Velopack updater'; Suffixes = @('Squirrel.exe') },
    [pscustomobject]@{ Label = 'VSTO prerequisite bootstrapper'; Suffixes = @('setup.exe') }
)

$RepoRoot   = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot
$PackScript = Join-Path $RepoRoot 'build\pack-velopack.ps1'
$ReleaseDir = Join-Path $RepoRoot 'artifacts\velopack-releases'

function Invoke-Native {
    param([string]$Description, [scriptblock]$Block)
    & $Block
    if ($LASTEXITCODE -ne 0) { throw "$Description が失敗しました (exit $LASTEXITCODE)" }
}

# ---- 0. プリフライト ----
Write-Host '== プリフライト ==' -ForegroundColor Cyan

# Git Bash (MSYS) 経由起動で括弧入り環境変数が落ちると vswhere.exe 解決が壊れるため補完
if (-not ${env:ProgramFiles(x86)}) { ${env:ProgramFiles(x86)} = 'C:\Program Files (x86)' }
$vsInstallerDir = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer'
if ($env:PATH -notlike "*$vsInstallerDir*") { $env:PATH = "$env:PATH;$vsInstallerDir" }
# vpk (dotnet tool) は net8.0〜net10.0 向けしか同梱しないため、.NET 11 だけの環境でも動くようロールフォワードする
$env:DOTNET_ROLL_FORWARD = 'Major'

# バージョンは Directory.Build.props の <Version> から取得 (表示・確認用)。
# csproj の ApplicationVersion は "$(Version).0" の導出式になったため、生テキストでは読めない。
$versionNode = ([xml](Get-Content 'Directory.Build.props' -Raw)).SelectSingleNode("//*[local-name()='Version']")
$version = if ($versionNode) { $versionNode.InnerText.Trim() } else { $null }
if (-not $version) { throw 'Directory.Build.props から <Version> を取得できませんでした' }
Write-Host "バージョン: $version"

# 署名の準備 (Sigilune の場所とクラウドへのログインを、ビルド前に確認して早く落とす)。
# 再ビルドしない -SkipBuild では署名しないためログインしない。
. (Join-Path $PSScriptRoot 'Sigilune.Common.ps1')
if (-not $SkipBuild) {
    $Sigilune = Resolve-SigiluneExe
    Connect-Sigilune -SigiluneExe $Sigilune
    $sigiluneCertificate = Get-SigiluneCertificate -SigiluneExe $Sigilune
    if ($sigiluneCertificate.Thumbprint -ne $SigningThumbprint -or $sigiluneCertificate.Subject -ne $CertSubjectName) {
        throw "Sigilune の署名証明書 ($($sigiluneCertificate.Subject), $($sigiluneCertificate.Thumbprint)) がクライアントの期待値と一致しません。証明書を更新したときは SigningThumbprint とクライアントの信頼設定を先に更新してください。"
    }
    Write-Host "署名証明書: $($sigiluneCertificate.Subject) ($($sigiluneCertificate.Thumbprint))"
}

# Cloudflare トークン (アップロード時のみ必要)。CI は環境変数、ローカルは secrets.json から読む。
# 値はスクリプト変数で持ち、環境変数には書き戻さない。pack-velopack.ps1 の子プロセス
# (MSBuild・NuGet 復元で動くパッケージのターゲット・vpk・dotnet) に R2 書き込みトークンを継承させないため。
$CloudflareApiToken = $null
if (-not $SkipUpload) {
    $CloudflareApiToken = $env:CLOUDFLARE_API_TOKEN
    if (-not $CloudflareApiToken) {
        $secrets = Get-Content $SecretsPath -Raw | ConvertFrom-Json
        if (-not $secrets.cloudflare.api_token) { throw "secrets.json に cloudflare.api_token が見つかりません" }
        $CloudflareApiToken = $secrets.cloudflare.api_token
    }
}
# CI ではワークフローがステップの環境変数として渡してくる。読み終えたら (アップロードしない場合も)
# このプロセスの環境変数から外し、以降の子プロセス (Sigilune・MSBuild・vpk・dotnet など) へ継承させない。
# 以降の R2 / Cloudflare API 呼び出しはすべて $CloudflareApiToken を明示して使う。
Remove-Item -LiteralPath Env:\CLOUDFLARE_API_TOKEN -ErrorAction SilentlyContinue

# 配信中の最新版がローカルの出力先に無いままアップロードすると、その版を含まない releases.*.json で
# 配信先を上書きしてしまう (例: CI で公開した版がローカルに無い)。pack-velopack.ps1 は出力先に
# manifest が既にあると配信元から取得しないため、ビルド前に照合して止める。
# 出力先に manifest が無い場合は pack が配信元の最新版を取得するので照合しない (CI と同じ経路)。
function Get-ReleaseManifestLatestFullAsset {
    param([Parameter(Mandatory)][AllowNull()][object]$Manifest)
    if ($null -eq $Manifest -or -not $Manifest.PSObject.Properties['Assets']) { return $null }
    return @($Manifest.Assets) |
        Where-Object { $_.PSObject.Properties['Type'] -and $_.Type -eq 'Full' -and $_.PSObject.Properties['FileName'] -and $_.FileName } |
        Sort-Object { [version]$_.Version } -Descending |
        Select-Object -First 1
}

$localManifestPath = Join-Path $ReleaseDir "releases.$Channel.json"
if (-not $SkipUpload -and (Test-Path -LiteralPath $localManifestPath)) {
    $remoteManifest = $null
    $remoteManifestMissing = $false
    try {
        $remoteManifest = Invoke-RestMethod -Uri "$BaseUrl/releases.$Channel.json" -Headers @{ 'Cache-Control' = 'no-cache' } -TimeoutSec 30
    }
    catch [Microsoft.PowerShell.Commands.HttpResponseException] {
        # 初回リリースなど、配信先に manifest が無い (404) 場合だけ照合を省く。それ以外の失敗は止める。
        if ($_.Exception.Response.StatusCode -ne [System.Net.HttpStatusCode]::NotFound) { throw }
        $remoteManifestMissing = $true
    }
    $remoteLatest = Get-ReleaseManifestLatestFullAsset -Manifest $remoteManifest
    if (-not $remoteManifestMissing -and $null -eq $remoteLatest) {
        # 応答は得たのに最新 Full 版を特定できない (ゲートウェイの HTML など) ときに、黙って照合を省かない。
        throw "配信中の $BaseUrl/releases.$Channel.json から最新の Full 版を特定できませんでした。応答を確認してから再実行してください。"
    }
    if ($null -ne $remoteLatest) {
        $localFileNames = @((Get-Content -LiteralPath $localManifestPath -Raw | ConvertFrom-Json).Assets | ForEach-Object { $_.FileName })
        if ($localFileNames -notcontains $remoteLatest.FileName) {
            throw "配信中の最新版 $($remoteLatest.FileName) がローカルの $localManifestPath にありません。このまま配信すると manifest が巻き戻るため中止します。artifacts\velopack-releases を退避して空にすると、pack が配信元の最新版を取得して差分の基準にします。"
        }
        Write-Host "配信中の最新版 $($remoteLatest.FileName) がローカルの manifest にあることを確認しました。"
    }
}

# ---- 1. ビルド + 署名付き Velopack パッケージング (pack-velopack.ps1 に委譲) ----
# 配信元を R2 (SimpleWebSource) に向け、ClickOnce/Authenticode 双方を本証明書で署名する。
Write-Host '== ビルド + 署名 (pack-velopack.ps1) ==' -ForegroundColor Cyan
if ($SkipBuild) {
    if ($ExpectedPackageSha256 -notmatch '^[0-9a-fA-F]{64}$') {
        throw '-SkipBuild には実機検証済みパッケージの -ExpectedPackageSha256 が必要です。'
    }
    $verifiedPackage = Join-Path $ReleaseDir "EXLSXS-$version-full.nupkg"
    if ((Get-FileHash -LiteralPath $verifiedPackage -Algorithm SHA256).Hash -ne $ExpectedPackageSha256) {
        throw '実機検証済みパッケージと配信予定パッケージの SHA-256 が一致しません。'
    }
    Write-Host '検証済みパッケージを使用し、ビルドと再署名を省略します。'
}
else {
# 出力先に前回の配布物が無いとき (CI) は配信元の最新版を差分の基準として取得する。
# ただし -SkipUpload の試行では取得しない (公開済みと同じ版のまま試すと vpk が同版を拒否するため)。
$downloadArgs = @(if (-not $SkipUpload) { '-DownloadPreviousRelease' })
Invoke-Native 'pack-velopack.ps1' {
    pwsh -NoProfile -ExecutionPolicy Bypass -File $PackScript @downloadArgs `
        -Channel $Channel `
        -UpdateSource $BaseUrl `
        -UpdateSourceKind Simple `
        -SigiluneExe $Sigilune `
        -ExpectedPublisherThumbprint $SigningThumbprint `
        -ExpectedPublisherSubject $CertSubjectName `
        -RequireVelopackSigning
}
}

if (-not (Test-Path $ReleaseDir)) { throw "Velopack 出力が見つかりません: $ReleaseDir" }

# ---- 2. 署名検証 ----
# (a) Setup.exe など出力直下の exe が本証明書で Valid 署名されているか
# (b) nupkg 内の全 PE が Authenticode 署名され、EXLSXS 発行者または Microsoft の署名か
# (c) Host.exe/.dll, VSTO dll が EXLSXS 発行者の署名か
#     = クライアントの UpdatePackageTrustVerifier が受理する状態かをリリース前に保証する
Write-Host '== 署名検証 ==' -ForegroundColor Cyan
foreach ($exe in Get-ChildItem $ReleaseDir -Filter '*.exe') {
    $sig = Get-AuthenticodeSignature $exe.FullName
    if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Subject -notlike "CN=$CertSubjectName*") {
        throw "署名検証失敗: $($exe.Name) → $($sig.Status)"
    }
    Write-Host "  ✅ $($exe.Name): Valid"
}

Add-Type -AssemblyName System.IO.Compression.FileSystem
$fullNupkg = Get-Item -LiteralPath (Join-Path $ReleaseDir "EXLSXS-$version-full.nupkg") -ErrorAction SilentlyContinue
if (-not $fullNupkg) { throw "full nupkg が見つかりません: $ReleaseDir" }
$verificationTempRoot = [IO.Path]::GetFullPath($env:TEMP)
$verifyDir = Join-Path $verificationTempRoot "exlsxs-relverify-$([Guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $verifyDir -Force | Out-Null
try {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($fullNupkg.FullName)
    try {
        $portableEntries = @($zip.Entries | Where-Object {
            $name = $_.FullName -replace '\.deploy$',''
            [IO.Path]::GetExtension($name) -in '.exe', '.dll'
        })
        if ($portableEntries.Count -eq 0) { throw 'nupkg 内に PE ファイルが見つかりません' }

        # 必須ファイルも同じ PE エントリーなので、検証済みの署名者をエントリー単位で保持する。
        $verifiedThumbprints = [System.Collections.Generic.Dictionary[System.IO.Compression.ZipArchiveEntry,string]]::new()
        foreach ($entry in $portableEntries) {
            $dest = Join-Path $verifyDir "$([Guid]::NewGuid().ToString('N'))-$([IO.Path]::GetFileName($entry.FullName) -replace '\.deploy$','')"
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dest, $true)
            $sig = Get-AuthenticodeSignature $dest
            if ($sig.Status -ne 'Valid') {
                throw "nupkg 内 '$($entry.FullName)' の Authenticode 署名が無効 ($($sig.Status))。"
            }

            $simpleName = $sig.SignerCertificate.GetNameInfo([Security.Cryptography.X509Certificates.X509NameType]::SimpleName, $false)
            $subjectParts = $sig.SignerCertificate.SubjectName.Decode(
                [Security.Cryptography.X509Certificates.X500DistinguishedNameFlags]::UseNewLines) -split "`r?`n"
            $isExpectedPublisher = $simpleName -eq $CertSubjectName
            $isMicrosoftPublisher = $subjectParts.Trim() -contains 'O=Microsoft Corporation'
            if (-not $isExpectedPublisher -and -not $isMicrosoftPublisher) {
                throw "nupkg 内 '$($entry.FullName)' の署名者が許可対象外です: $($sig.SignerCertificate.Subject)"
            }
            $verifiedThumbprints.Add($entry, ($sig.SignerCertificate.Thumbprint -replace '[^0-9A-Fa-f]','').ToUpperInvariant())
        }

        foreach ($required in $RequiredSignedEntries) {
            $entry = $zip.Entries | Where-Object {
                $entryName = $_.FullName.Replace('\','/')
                @($required.Suffixes | Where-Object { $entryName.EndsWith("/$_") -or $entryName -eq $_ }).Count -gt 0
            } | Select-Object -First 1
            if (-not $entry) { throw "nupkg 内に必須ファイル '$($required.Label)' が見つかりません" }
            $thumb = $verifiedThumbprints[$entry]
            if ($thumb -ne $SigningThumbprint) {
                throw "nupkg 内 '$($entry.FullName)' の署名者 thumbprint 不一致。期待 $SigningThumbprint, 実際 $thumb"
            }
            Write-Host "  ✅ nupkg: $($entry.FullName) → Valid / thumbprint 一致"
        }

        $settingsEntries = @($zip.Entries | Where-Object { $_.FullName.Replace('\','/') -eq 'lib/app/appsettings.json' })
        if ($settingsEntries.Count -ne 1) {
            throw "nupkg 内の lib/app/appsettings.json は1件必須です (実際: $($settingsEntries.Count)件)"
        }
        $reader = [IO.StreamReader]::new($settingsEntries[0].Open())
        try { $packagedSettings = $reader.ReadToEnd() | ConvertFrom-Json }
        finally { $reader.Dispose() }
        $packagedThumbprint = ($packagedSettings.Update.ExpectedPublisherThumbprint -replace '[^0-9A-Fa-f]','').ToUpperInvariant()
        if ($packagedThumbprint -ne $SigningThumbprint -or $packagedSettings.Update.ExpectedPublisherSubject -ne $CertSubjectName) {
            throw 'nupkg 内の発行者信頼設定がリリース用証明書と一致しません'
        }
    } finally { $zip.Dispose() }
} finally { Remove-LocalWorkItem -LiteralPath $verifyDir -AllowedRoot $verificationTempRoot -IgnoreMissing }

if ($SkipUpload) {
    Write-Host "`n✅ -SkipUpload 指定のためここで終了。成果物: $ReleaseDir" -ForegroundColor Green
    Get-ChildItem $ReleaseDir | Format-Table Name, @{n='Size(MB)'; e={[math]::Round($_.Length/1MB,2)}}
    return
}

# ---- 3. R2 アップロード ----
# releases.{channel}.json (manifest) は同 channel の旧版を上書き、*.nupkg は put のみ。
# manifest を最後にアップロードしてフィード整合性の時間窓を防ぐ。
Write-Host '== R2 アップロード ==' -ForegroundColor Cyan
$files = Get-ChildItem $ReleaseDir -File | Sort-Object @{ Expression = { if ($_.Name -like 'releases.*.json') { 3 } elseif ($_.Name -like 'RELEASES*') { 2 } elseif ($_.Name -like 'assets.*.json') { 1 } else { 0 } } }, Name
$uploaded = 0
foreach ($f in $files) {
    Write-Host "  ↑ $($f.Name)"
    Send-ReleaseR2Object -File $f
    $uploaded++
}
Write-Host "✅ R2 アップロード完了: $uploaded ファイル"

# ---- 3.5 Cloudflare エッジキャッシュのパージ ----
# 固定名ファイル (Setup.exe / Portable.zip / RELEASES / releases.*.json / assets.*.json) は
# 毎リリースで中身が変わるのに URL が不変。CDN エッジが旧版を Cache-Control の max-age 分保持するため、
# パージしないと新規ダウンロード・自動更新が旧バージョンを掴む。アップロード直後に該当 URL をパージして
# キャッシュ削除を要求する。バージョン付き nupkg は URL が一意 (旧キャッシュなし) のためパージ不要。
Write-Host '== Cloudflare キャッシュパージ ==' -ForegroundColor Cyan
$cfHeaders = @{ Authorization = "Bearer $CloudflareApiToken" }
$fixedFiles = @($files | Where-Object { $_.Name -notlike '*.nupkg' })
foreach ($distributionBaseUrl in $DistributionBaseUrls) {
    $zoneName = ([uri]$distributionBaseUrl).Host -replace '^[^.]+\.', ''   # <sub>.example.com → example.com (apex)
    $zoneResp = Invoke-RestMethod -Uri "https://api.cloudflare.com/client/v4/zones?name=$zoneName" -Headers $cfHeaders -TimeoutSec 30
    if (-not $zoneResp.success -or @($zoneResp.result).Count -eq 0) { throw "Cloudflare zone '$zoneName' の取得に失敗しました" }
    $zoneId = $zoneResp.result[0].id
    $purgeUrls = @($fixedFiles | ForEach-Object { "$distributionBaseUrl/$($_.Name)" })
    if ($purgeUrls.Count -gt 0) {
        $purgeBody = "{`"files`":$(ConvertTo-Json -InputObject $purgeUrls -AsArray -Compress)}"
        $purgeResp = Invoke-RestMethod -Method Post -Uri "https://api.cloudflare.com/client/v4/zones/$zoneId/purge_cache" `
            -Headers $cfHeaders -ContentType 'application/json' -Body $purgeBody -TimeoutSec 30
        if (-not $purgeResp.success) { throw "Cloudflare キャッシュパージに失敗しました ($zoneName): $($purgeResp.errors | ConvertTo-Json -Compress)" }
        Write-Host "  ✅ $zoneName をパージ: $($purgeUrls.Count) URL"
        $purgeUrls | ForEach-Object { Write-Host "     $_" }
    } else {
        Write-Host "  $zoneName のパージ対象なし"
    }
}

# ---- 4. 旧バージョン nupkg のクリーンアップ (世代保持戦略) ----
# 最新 manifest が指す nupkg に加え、R2 上の nupkg を新しい順に $KeepGenerations 世代保持する。
# 壊れた版を配信したとき 1 つ前の安定版へロールバックできる戻り道を残すため、全削除はしない。
# 固定名ファイル (Setup.exe / RELEASES* / assets.*.json / releases.*.json) は対象外で安全。
Write-Host '== 旧 nupkg クリーンアップ ==' -ForegroundColor Cyan
$KeepGenerations = 2
$keep = @{}
$manifests = Get-ChildItem $ReleaseDir -Filter 'releases.*.json'
if (-not $manifests) { throw 'artifacts に releases.*.json が見つかりません' }
foreach ($m in $manifests) {
    foreach ($asset in (Get-Content $m.FullName -Raw | ConvertFrom-Json).Assets) {
        if ($asset.FileName) { $keep[$asset.FileName] = $true }
    }
}
Write-Host "  保持対象 nupkg: $($keep.Count) 件"

$api = "https://api.cloudflare.com/client/v4/accounts/$AccountId/r2/buckets/$Bucket"
$headers = @{ Authorization = "Bearer $CloudflareApiToken" }

$allKeys = [System.Collections.Generic.List[string]]::new()
$cursor = ''
while ($true) {
    $uri = "$api/objects?per_page=1000" + $(if ($cursor) { "&cursor=$cursor" })
    $resp = Invoke-RestMethod -Uri $uri -Headers $headers -TimeoutSec 30
    foreach ($obj in $resp.result) { $allKeys.Add($obj.key) }
    $info = $resp.PSObject.Properties['result_info']
    if (-not $info -or -not $info.Value) { break }
    $truncated = $info.Value.PSObject.Properties['is_truncated']
    if (-not $truncated -or -not $truncated.Value) { break }
    $cursorProp = $info.Value.PSObject.Properties['cursor']
    $cursor = if ($cursorProp) { $cursorProp.Value } else { '' }
    if (-not $cursor) { break }
}

# R2 上のバージョン付きファイルを世代ごとに分類し、新しい順に $KeepGenerations 世代を保持する。
# '*.nupkg' に限らず、バージョン文字列を含むキーはすべて世代の対象にする (取り残しを防ぐため)。
# 固定ファイル名 (Setup.exe / Portable.zip / RELEASES* / releases.*.json) は
# バージョン文字列を含まないので対象外 = 安全。
$versionPattern = '(\d+\.\d+\.\d+)'
$versionedKeys = $allKeys | Where-Object { [regex]::IsMatch($_, $versionPattern) }
$versionsOnR2 = $versionedKeys |
    ForEach-Object { [regex]::Match($_, $versionPattern).Groups[1].Value } |
    Select-Object -Unique |
    Sort-Object { [version]$_ } -Descending
$keepVersions = @{}
foreach ($v in ($versionsOnR2 | Select-Object -First $KeepGenerations)) { $keepVersions[$v] = $true }
if ($keepVersions.Count -gt 0) { Write-Host "  保持世代: $($keepVersions.Keys -join ', ')" }

$toDelete = $versionedKeys | Where-Object {
    # 最新 manifest が指すものは必ず保持 (消すと自動更新が壊れる)
    if ($keep.ContainsKey($_)) { $false }
    else { -not $keepVersions.ContainsKey([regex]::Match($_, $versionPattern).Groups[1].Value) }
}
if (-not $toDelete) {
    Write-Host '  ✅ 削除対象なし'
} else {
    $deleted = 0; $failed = 0
    foreach ($key in $toDelete) {
        $encoded = [uri]::EscapeDataString($key)
        try {
            Invoke-RestMethod -Method Delete -Uri "$api/objects/$encoded" -Headers $headers -TimeoutSec 30 | Out-Null
            Write-Host "  🗑️  $key"
            $deleted++
        } catch {
            Write-Warning "  削除失敗: $key — $($_.Exception.Message)"
            $failed++
        }
    }
    Write-Host "  🧹 クリーンアップ: $deleted 削除 / $failed 失敗"
    if ($failed -gt 0 -and $deleted -eq 0) { throw '旧 nupkg の削除がすべて失敗しました。API token の権限を確認してください。' }
}

Write-Host "`n🎉 リリース完了: v$version → $BaseUrl" -ForegroundColor Green
