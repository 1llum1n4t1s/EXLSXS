# Sigilune.Common.ps1 — Certum SimplySign のクラウド署名を Sigilune で行う共通関数。
#
# Sigilune (https://github.com/1llum1n4t1s/Sigilune) は signtool の /dlib として
# クラウド鍵で署名するため、SimplySign Desktop の接続やワンタイムコードの手入力が要らない。
# 事前に `sigilune setup` で Certum アカウントと TOTP 登録情報を資格情報マネージャーへ登録しておく。
#
# GitHub Actions では `uses: 1llum1n4t1s/Sigilune@v1` が PATH へ導入し、資格情報は
# Environment "release" の Secrets を環境変数 SIGILUNE_EMAIL / SIGILUNE_TOTP_URI で渡す。
#
# 実行ファイルの場所は 環境変数 SIGILUNE_EXE → PATH 上の sigilune.exe → %USERPROFILE%\.local\Sigilune\bin\sigilune.exe の順。
# 同じフォルダーに sigilune_dlib.dll が必要。

# Sigilune の実行ファイルを解決する（見つからなければ導入手順を示して停止する）。
function Resolve-SigiluneExe {
    $onPath = Get-Command 'sigilune.exe' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    $candidate = if ($env:SIGILUNE_EXE) { $env:SIGILUNE_EXE }
        elseif ($onPath) { $onPath.Source }
        else { Join-Path ([Environment]::GetFolderPath('UserProfile')) '.local\Sigilune\bin\sigilune.exe' }
    if (-not [IO.Path]::IsPathFullyQualified($candidate) -or -not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
        throw "Sigilune が見つかりません: $candidate`nhttps://github.com/1llum1n4t1s/Sigilune を導入し、sigilune setup で Certum アカウントを登録してください（場所は SIGILUNE_EXE で指定可）。"
    }
    if (-not (Test-Path -LiteralPath (Join-Path (Split-Path -Parent $candidate) 'sigilune_dlib.dll') -PathType Leaf)) {
        throw "sigilune_dlib.dll が sigilune.exe と同じフォルダーにありません: $candidate"
    }
    return $candidate
}

# Certum クラウドのセッションを用意する（キャッシュが有効ならログインしない）。署名前に呼び、失敗を早く検出する。
function Connect-Sigilune {
    param([Parameter(Mandatory)][string]$SigiluneExe)
    & $SigiluneExe login
    if ($LASTEXITCODE -ne 0) { throw "Sigilune のログインに失敗しました (exit $LASTEXITCODE)。sigilune status で登録を確認してください。自動再試行はしません。" }
}

# ファイルへ署名し、Sigilune の検証（有効・署名者一致・タイムスタンプ）に加えて署名者の CN を照合する。
function Invoke-SigiluneSign {
    param(
        [Parameter(Mandatory)][string]$SigiluneExe,
        [Parameter(Mandatory)][string[]]$Path,
        [Parameter(Mandatory)][string]$CertificateSubjectName,
        # 既存の署名を置き換える (ビルド時の使い捨て証明書の署名など、置き換えてよいと確かめた場合だけ)。
        [switch]$Replace
    )
    $replaceArgs = @(if ($Replace) { '--replace' })
    & $SigiluneExe sign @replaceArgs @Path
    if ($LASTEXITCODE -ne 0) { throw "Sigilune の署名に失敗しました (exit $LASTEXITCODE): $($Path -join ', ')" }
    foreach ($file in $Path) {
        $signature = Get-AuthenticodeSignature -LiteralPath $file
        if ($signature.Status -ne 'Valid' -or $signature.SignerCertificate.Subject -notlike "CN=$CertificateSubjectName*" -or -not $signature.TimeStamperCertificate) {
            throw "署名検証に失敗しました: $file → $($signature.Status)"
        }
    }
}

# 現在のセッションの署名証明書（CN と拇印）を返す。発行者信頼設定の期待値に使う。
function Get-SigiluneCertificate {
    param([Parameter(Mandatory)][string]$SigiluneExe)
    $json = & $SigiluneExe cert --json
    if ($LASTEXITCODE -ne 0) { throw "Sigilune の証明書取得に失敗しました (exit $LASTEXITCODE)" }
    $certificate = ($json -join "`n") | ConvertFrom-Json
    return [pscustomobject]@{
        Subject = [string]$certificate.subject
        Thumbprint = ([string]$certificate.thumbprint -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
    }
}

# ClickOnce / VSTO マニフェストへ署名する（Sigilune 1.1.0 以降の sign-manifest）。
# アプリケーション マニフェストへ先に署名し、配置マニフェスト側の参照（サイズ・ダイジェスト・公開鍵トークン）を
# 更新してから配置マニフェストへ署名する。Sigilune が署名後に ClickOnce の検証（署名・タイムスタンプ）を行い、
# ここでは署名者の CN も照合する。
function Invoke-SigiluneManifestSign {
    param(
        [Parameter(Mandatory)][string]$SigiluneExe,
        [Parameter(Mandatory)][string]$ApplicationManifest,
        [Parameter(Mandatory)][string[]]$DeploymentManifest,
        [Parameter(Mandatory)][string]$CertificateSubjectName
    )
    $tool = if ($env:SIGILUNE_MANIFEST_TOOL) { $env:SIGILUNE_MANIFEST_TOOL } else { Join-Path (Split-Path -Parent $SigiluneExe) 'sigilune-manifest.exe' }
    if (-not (Test-Path -LiteralPath $tool -PathType Leaf)) {
        throw "sigilune-manifest.exe が見つかりません: $tool`nマニフェスト署名には Sigilune 1.1.0 以降が必要です。"
    }
    & $SigiluneExe sign-manifest --application $ApplicationManifest @DeploymentManifest
    if ($LASTEXITCODE -ne 0) { throw "マニフェストの署名に失敗しました (exit $LASTEXITCODE)" }
    & $SigiluneExe verify-manifest --signer $CertificateSubjectName $ApplicationManifest @DeploymentManifest
    if ($LASTEXITCODE -ne 0) { throw "マニフェストの署名者が $CertificateSubjectName と一致しません" }
}
