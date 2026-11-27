param([switch]$Initialize, [string]$PayloadPath, [string]$OutputPath, [string]$KeyDirectory='', [string]$PublicKeyPath='')
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Security
$privatePath=Join-Path $PSScriptRoot '.release-secrets\ncc-receipt-signing.dpapi'
$publicPath=Join-Path $PSScriptRoot 'update_public_key.json'
if ($KeyDirectory) { $privatePath=Join-Path ([IO.Path]::GetFullPath($KeyDirectory)) 'ncc-receipt-signing.dpapi' }
if ($PublicKeyPath) { $publicPath=[IO.Path]::GetFullPath($PublicKeyPath) }
$utf8=New-Object System.Text.UTF8Encoding($false)
$rsa=New-Object System.Security.Cryptography.RSACryptoServiceProvider(3072)
$rsa.PersistKeyInCsp=$false
try {
    if ($Initialize) {
        if ((Test-Path -LiteralPath $privatePath) -or (Test-Path -LiteralPath $publicPath)) { throw '签名密钥已存在，已保留原密钥。' }
        $bytes=$utf8.GetBytes($rsa.ToXmlString($true))
        $encrypted=[Security.Cryptography.ProtectedData]::Protect($bytes,$null,[Security.Cryptography.DataProtectionScope]::CurrentUser)
        New-Item -ItemType Directory -Path (Split-Path -Parent $privatePath) -Force | Out-Null
        [IO.File]::WriteAllBytes($privatePath,$encrypted)
        $parameters=$rsa.ExportParameters($false)
        $digest=[Security.Cryptography.SHA256]::Create()
        try { $keyId=([BitConverter]::ToString($digest.ComputeHash($parameters.Modulus))).Replace('-','').ToLowerInvariant() } finally { $digest.Dispose() }
        $public=@{key_id=$keyId;modulus=[Convert]::ToBase64String($parameters.Modulus);exponent=65537} | ConvertTo-Json
        [IO.File]::WriteAllText($publicPath,$public,$utf8)
        Write-Output '更新签名密钥已创建；私钥由当前 Windows 用户加密，未写入程序或发布包。'
    } else {
        if (-not $PayloadPath -or -not $OutputPath) { throw '请指定待签名清单与输出路径。' }
        $decrypted=[Security.Cryptography.ProtectedData]::Unprotect([IO.File]::ReadAllBytes($privatePath),$null,[Security.Cryptography.DataProtectionScope]::CurrentUser)
        $rsa.FromXmlString($utf8.GetString($decrypted))
        $public=Get-Content -LiteralPath $publicPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([Convert]::ToBase64String($rsa.ExportParameters($false).Modulus) -ne $public.modulus) { throw '公钥与私钥不匹配。' }
        $payload=[IO.File]::ReadAllBytes([IO.Path]::GetFullPath($PayloadPath))
        $sha=[Security.Cryptography.SHA256]::Create()
        try { $signature=$rsa.SignData($payload,$sha) } finally { $sha.Dispose() }
        $envelope=@{key_id=$public.key_id;payload=[Convert]::ToBase64String($payload);signature=[Convert]::ToBase64String($signature)} | ConvertTo-Json
        [IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputPath),$envelope,$utf8)
        Write-Output '清单已签名。'
    }
} finally { $rsa.Dispose() }
