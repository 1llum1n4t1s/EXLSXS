# release-local.ps1 — ローカル署名付き Velopack リリース (EXLSXS / VSTO アドイン)
#
# SimplySign (Certum クラウド署名) は Desktop 接続 + スマホトークンが必要で
# GitHub Actions からは署名できないため、リリースは本スクリプトでローカル実行する。
# 旧 CI リリース (.github/workflows/velopack-release.yml) はこのスクリプトに置換済み。
#
# EXLSXS は「VSTO アドイン (.NET Fw 4.8.1, Library) + Velopack ホスト」の 2 段構成で、
# ビルド/パッケージングは build/pack-velopack.ps1 が一括 (VSTO publish + ClickOnce 署名 →
# host publish → staging → vpk pack + Authenticode 署名) で担う。本スクリプトはその上に
# プリフライト → pack 呼び出し → 署名検証 → R2 アップロード → 配信確認 → 旧版掃除 を被せる。
#
# 前提:
#   - SimplySign Desktop が接続済み (証明書が CurrentUser\My に見えていること)
#   - Directory.Build.props の <Version> がリリースしたいバージョンであること (/vava 済み)
#   - C:\Users\IMT\dev\Secret\secrets.json に cloudflare.api_token があること
#
# 使い方:
#   pwsh scripts/release-local.ps1                # フルリリース (build + sign + upload + cleanup)
#   pwsh scripts/release-local.ps1 -SkipUpload    # ビルド + 署名のみ (アップロードしない動作確認用)

[CmdletBinding()]
param(
    [switch]$SkipUpload
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ---- 定数 ----
$Bucket          = 'exlsxs-updates'
$BaseUrl         = 'https://exlsxs.kagayoi.com'
$AccountId       = '10901bfadbf1005164774a7350082985'
$Channel         = 'win'
$SecretsPath     = 'C:\Users\IMT\dev\Secret\secrets.json'
$CertSubjectName = 'Open Source Developer Yuichiro Shinozaki'
# 署名証明書 thumbprint (ClickOnce manifest 署名 + Authenticode 署名者の期待値)
$SigningThumbprint = '6285702C9AF1FCFE3D9FE815B7F7F625508130C0'
# /n (Subject 名) で選択: 証明書の年次更新で thumbprint が変わっても署名自体は動く
$SignParams      = "/n `"$CertSubjectName`" /fd SHA256 /td SHA256 /tr http://time.certum.pl"
$WranglerVersion = '4.127.1'  # サプライチェーン対策でバージョン固定

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
# vpk (dotnet tool) は .NET 9 ランタイム要求だがローカルは 10 のみ → ロールフォワード
$env:DOTNET_ROLL_FORWARD = 'Major'

# バージョンは Directory.Build.props の <Version> から取得 (表示・確認用)。
# csproj の ApplicationVersion は "$(Version).0" の導出式になったため、生テキストでは読めない。
$versionNode = ([xml](Get-Content 'Directory.Build.props' -Raw)).SelectSingleNode("//*[local-name()='Version']")
$version = if ($versionNode) { $versionNode.InnerText.Trim() } else { $null }
if (-not $version) { throw 'Directory.Build.props から <Version> を取得できませんでした' }
Write-Host "バージョン: $version"

# SimplySign 接続確認 (証明書が見えなければ署名できないので最初に落とす)
$cert = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object { $_.Subject -like "CN=$CertSubjectName*" -and $_.NotAfter -gt (Get-Date) }
if (-not $cert) {
    throw "署名証明書 (CN=$CertSubjectName) が見つかりません。SimplySign Desktop を起動してトークンでログインしてください。"
}
Write-Host "署名証明書: $($cert.Subject) (期限 $($cert.NotAfter.ToString('yyyy-MM-dd')))"

# Cloudflare トークン (アップロード時のみ必要)
if (-not $SkipUpload) {
    $secrets = Get-Content $SecretsPath -Raw | ConvertFrom-Json
    if (-not $secrets.cloudflare.api_token) { throw "secrets.json に cloudflare.api_token が見つかりません" }
    $env:CLOUDFLARE_API_TOKEN  = $secrets.cloudflare.api_token
    $env:CLOUDFLARE_ACCOUNT_ID = $AccountId
}

# ---- 1. ビルド + 署名付き Velopack パッケージング (pack-velopack.ps1 に委譲) ----
# 配信元を R2 (SimpleWebSource) に向け、ClickOnce/Authenticode 双方を本証明書で署名する。
Write-Host '== ビルド + 署名 (pack-velopack.ps1) ==' -ForegroundColor Cyan
Invoke-Native 'pack-velopack.ps1' {
    pwsh -NoProfile -ExecutionPolicy Bypass -File $PackScript `
        -Channel $Channel `
        -UpdateSource $BaseUrl `
        -UpdateSourceKind Simple `
        -VstoSigningThumbprint $SigningThumbprint `
        -VelopackSignParams $SignParams `
        -ExpectedPublisherThumbprint $SigningThumbprint `
        -ExpectedPublisherSubject $CertSubjectName `
        -RequireVelopackSigning
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
$fullNupkg = Get-ChildItem $ReleaseDir -Filter '*-full.nupkg' | Select-Object -First 1
if (-not $fullNupkg) { throw "full nupkg が見つかりません: $ReleaseDir" }
$verifyDir = Join-Path $env:TEMP "exlsxs-relverify-$([Guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $verifyDir -Force | Out-Null
try {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($fullNupkg.FullName)
    try {
        $portableEntries = @($zip.Entries | Where-Object {
            $name = $_.FullName -replace '\.deploy$',''
            [IO.Path]::GetExtension($name) -in '.exe', '.dll'
        })
        if ($portableEntries.Count -eq 0) { throw 'nupkg 内に PE ファイルが見つかりません' }

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
        }

        foreach ($required in $RequiredSignedEntries) {
            $entry = $zip.Entries | Where-Object {
                $entryName = $_.FullName.Replace('\','/')
                @($required.Suffixes | Where-Object { $entryName.EndsWith("/$_") -or $entryName -eq $_ }).Count -gt 0
            } | Select-Object -First 1
            if (-not $entry) { throw "nupkg 内に必須ファイル '$($required.Label)' が見つかりません" }
            $dest = Join-Path $verifyDir ([IO.Path]::GetFileName($entry.FullName) -replace '\.deploy$','')
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dest, $true)
            $sig = Get-AuthenticodeSignature $dest
            if ($sig.Status -ne 'Valid') {
                throw "nupkg 内 '$($entry.FullName)' の Authenticode 署名が無効 ($($sig.Status))。クライアントの更新検証が失敗します。"
            }
            $thumb = ($sig.SignerCertificate.Thumbprint -replace '[^0-9A-Fa-f]','').ToUpperInvariant()
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
} finally { Remove-Item $verifyDir -Recurse -Force -ErrorAction SilentlyContinue }

if ($SkipUpload) {
    Write-Host "`n✅ -SkipUpload 指定のためここで終了。成果物: $ReleaseDir" -ForegroundColor Green
    Get-ChildItem $ReleaseDir | Format-Table Name, @{n='Size(MB)'; e={[math]::Round($_.Length/1MB,2)}}
    return
}

# ---- 3. R2 アップロード ----
# releases.{channel}.json (manifest) は同 channel の旧版を上書き、*.nupkg は put のみ。
# manifest を最後にアップロードしてフィード整合性の時間窓を防ぐ。
Write-Host '== R2 アップロード ==' -ForegroundColor Cyan
$files = Get-ChildItem $ReleaseDir -File | Sort-Object { $_.Name -like 'releases.*.json' }  # manifest を末尾へ
$uploaded = 0
foreach ($f in $files) {
    Write-Host "  ↑ $($f.Name)"
    Invoke-Native "R2 put ($($f.Name))" {
        pnpm dlx "wrangler@$WranglerVersion" r2 object put "$Bucket/$($f.Name)" --file $f.FullName --remote
    }
    $uploaded++
}
Write-Host "✅ R2 アップロード完了: $uploaded ファイル"

# ---- 3.5 Cloudflare エッジキャッシュのパージ ----
# 固定名ファイル (Setup.exe / Portable.zip / RELEASES / releases.*.json / assets.*.json) は
# 毎リリースで中身が変わるのに URL が不変。CDN エッジが旧版を Cache-Control の max-age 分保持するため、
# パージしないと新規ダウンロード・自動更新が旧バージョンを掴む。アップロード直後に該当 URL をパージして
# 伝播を確定する。バージョン付き nupkg は URL が一意 (旧キャッシュなし) のためパージ不要。
Write-Host '== Cloudflare キャッシュパージ ==' -ForegroundColor Cyan
$cfHeaders = @{ Authorization = "Bearer $($env:CLOUDFLARE_API_TOKEN)" }
$zoneName = ([uri]$BaseUrl).Host -replace '^[^.]+\.', ''   # <sub>.kagayoi.com → kagayoi.com (apex)
$zoneResp = Invoke-RestMethod -Uri "https://api.cloudflare.com/client/v4/zones?name=$zoneName" -Headers $cfHeaders -TimeoutSec 30
if (-not $zoneResp.success -or @($zoneResp.result).Count -eq 0) { throw "Cloudflare zone '$zoneName' の取得に失敗しました" }
$zoneId = $zoneResp.result[0].id
$purgeUrls = @($files | Where-Object { $_.Name -notlike '*.nupkg' } | ForEach-Object { "$BaseUrl/$($_.Name)" })
if ($purgeUrls.Count -gt 0) {
    $purgeBody = "{`"files`":$(ConvertTo-Json -InputObject $purgeUrls -AsArray -Compress)}"
    $purgeResp = Invoke-RestMethod -Method Post -Uri "https://api.cloudflare.com/client/v4/zones/$zoneId/purge_cache" `
        -Headers $cfHeaders -ContentType 'application/json' -Body $purgeBody -TimeoutSec 30
    if (-not $purgeResp.success) { throw "Cloudflare キャッシュパージに失敗しました: $($purgeResp.errors | ConvertTo-Json -Compress)" }
    Write-Host "  ✅ パージ: $($purgeUrls.Count) URL"
    $purgeUrls | ForEach-Object { Write-Host "     $_" }
} else {
    Write-Host '  パージ対象なし'
}

# ---- 4. 配信確認 (manifest 完全一致リトライ) ----
# 単純な HTTP 200 だと CDN/edge が古い manifest を返している間に cleanup が走り、
# 旧 manifest を掴んだクライアントが直後に消える nupkg を取りに行く race がある。
# ローカル manifest と完全一致するまでリトライしてから cleanup へ進む。
Write-Host '== 配信確認 (manifest 伝播待機) ==' -ForegroundColor Cyan
$url = "$BaseUrl/releases.$Channel.json"
$localManifest = Get-Content (Join-Path $ReleaseDir "releases.$Channel.json") -Raw |
    ConvertFrom-Json | ConvertTo-Json -Depth 100 -Compress
$maxAttempts = 18
$matched = $false
for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
    $resp = Invoke-WebRequest -Uri "${url}?_=$([Guid]::NewGuid().ToString('N'))" `
        -Headers @{ 'Cache-Control' = 'no-cache' } -TimeoutSec 30
    $raw = $resp.Content
    if ($raw -is [byte[]]) { $raw = [System.Text.Encoding]::UTF8.GetString($raw) }
    $remoteManifest = $raw | ConvertFrom-Json | ConvertTo-Json -Depth 100 -Compress
    if ($localManifest -eq $remoteManifest) {
        Write-Host "  ✅ $url がローカル manifest と一致 (attempt $attempt)"
        $matched = $true
        break
    }
    Write-Host "  ⚠️ remote manifest がまだ古い (attempt $attempt / $maxAttempts)、5 秒待機..."
    Start-Sleep -Seconds 5
}
if (-not $matched) {
    throw "remote manifest が $($maxAttempts * 5) 秒以内にローカルと一致しませんでした。race 回避のため cleanup を中止します: $url"
}

# ---- 5. 旧バージョン nupkg のクリーンアップ (世代保持戦略) ----
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
$headers = @{ Authorization = "Bearer $($env:CLOUDFLARE_API_TOKEN)" }

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
# ⚠️ 旧実装は '*.nupkg' だけを対象にしていたため、バージョン付きの配布物 (zip / deb / rpm /
#    AppImage 等) が R2 に永久に溜まっていた (Ferry で 351 個 7.2GB = 含有枠 10GB の 72%)。
#    固定ファイル名 (Setup.exe / Portable.zip / RELEASES* / releases.*.json) は
#    バージョン文字列を含まないので対象外 = 安全。
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
