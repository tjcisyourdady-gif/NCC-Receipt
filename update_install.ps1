param([Parameter(Mandatory=$true)][string]$PlanPath,[Parameter(Mandatory=$true)][string]$PlanSha256,[switch]$NoRestart,[string]$ReadyPath='')
$ErrorActionPreference='Stop'
function File-Hash([string]$Path) {
    $stream=[IO.File]::OpenRead($Path)
    $digest=[Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($digest.ComputeHash($stream))).Replace('-','').ToLowerInvariant() }
    finally { $stream.Dispose(); $digest.Dispose() }
}
if ((File-Hash $PlanPath) -ne $PlanSha256) { throw '更新计划校验失败。' }
$plan=Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
$allowedExes=@('收款结算单填写工具V7-正常版.exe','收款结算单填写工具V7-可修改时间版-卡顿使用.exe')
if ($plan.app_id -ne 'ncc-receipt' -or ($plan.exe_name -notin $allowedExes)) { throw '应用身份不符。' }
$appDir=[IO.Path]::GetFullPath($plan.app_dir)
$updateRoot=Join-Path $appDir '.updates'
function In-Root([string]$Path,[string]$Root) {
    $full=[IO.Path]::GetFullPath($Path)
    if (-not $full.StartsWith($Root.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)) { throw '路径越界。' }
    return $full
}
$stage=In-Root $plan.stage_dir $updateRoot
$backup=In-Root $plan.backup_dir $updateRoot
$null=In-Root $PlanPath $updateRoot
$lockPath=In-Root (Join-Path $updateRoot '.ncc_receipt_update.lock') $updateRoot
$exe=Join-Path $appDir $plan.exe_name
$lock=$null
$applied=New-Object System.Collections.Generic.List[object]
$restart=$false
$locked=$false
try {
    $waitIds=@([int]$plan.pid)
    if ($plan.parent_pid) {
        $parent=Get-Process -Id $plan.parent_pid -ErrorAction SilentlyContinue
        if ($parent -and $parent.Path -eq $exe) { $waitIds+=([int]$plan.parent_pid) }
    }
    foreach ($candidateName in $allowedExes) {
        $candidatePath=Join-Path $appDir $candidateName
        $others=@(Get-Process -Name ([IO.Path]::GetFileNameWithoutExtension($candidatePath)) -ErrorAction SilentlyContinue | Where-Object { $_.Path -eq $candidatePath -and $_.Id -notin $waitIds })
        if ($others.Count -gt 0) { throw '同目录还有另一个收款结算单程序正在运行，请先关闭后重试安装。' }
    }
    if ($ReadyPath) {
        $readyFile=In-Root $ReadyPath $updateRoot
        @{app_id='ncc-receipt';plan_sha256=$PlanSha256;helper_pid=$PID;status='waiting_for_exit'} | ConvertTo-Json | Set-Content -LiteralPath $readyFile -Encoding UTF8
    }
    Write-Output '安装助手已就绪，等待旧程序退出。'
    $deadline=(Get-Date).AddSeconds(120)
    while (@($waitIds | ForEach-Object { Get-Process -Id $_ -ErrorAction SilentlyContinue }).Count -gt 0) {
        if ((Get-Date) -gt $deadline) { throw '程序未退出，未修改文件。' }
        Start-Sleep -Milliseconds 250
    }
    $lock=[IO.File]::Open($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::ReadWrite)
    $lock.Lock(0,1)
    $locked=$true
    Write-Output '开始备份并安装更新。'
    if ($lock.Length -eq 0) { $lock.WriteByte(0); $lock.Flush() }
    New-Item -ItemType Directory -Path $backup -Force | Out-Null
    $prepared=New-Object System.Collections.Generic.List[object]
    $seen=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $allowedRootFiles=@(
        '收款结算单填写工具V7-正常版.exe',
        '收款结算单填写工具V7-可修改时间版-卡顿使用.exe',
        'V7双版本及GUI功能说明.txt',
        '运行调试（选择版本并保留日志）.bat',
        'README.md'
    )
    foreach ($file in $plan.files) {
        $relative=[string]$file.path
        if ($relative -match '(^|/)(\.|\.\.)(/|$)|[:"<>|?*]' -or $relative.StartsWith('/') -or $relative.Contains('\')) { throw '非法相对路径。' }
        $parts=$relative.Split('/')
        $allowed=($parts.Count -eq 1 -and $parts[0] -in $allowedRootFiles) -or ($parts.Count -gt 1 -and $parts[0] -eq 'ms-playwright（勿动）')
        if (-not $allowed -or -not $seen.Add($relative)) { throw '更新文件不在白名单或重复。' }
        foreach ($part in $parts) { if (-not $part -or $part.EndsWith('.') -or $part.EndsWith(' ') -or $part -match '^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\.|$)|[\x00-\x1f]') { throw '非法文件名。' } }
        $source=In-Root (Join-Path $stage $relative) $stage
        $target=In-Root (Join-Path $appDir $relative) $appDir
        $saved=In-Root (Join-Path $backup $relative) $backup
        $cursor=$target
        while ($cursor -ne $appDir) {
            if ((Test-Path -LiteralPath $cursor) -and ((Get-Item -LiteralPath $cursor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw '目标路径包含联接。' }
            $cursor=Split-Path -Parent $cursor
        }
        if ((File-Hash $source) -ne $file.sha256) { throw '暂存文件校验失败。' }
        $existed=Test-Path -LiteralPath $target -PathType Leaf
        if ($existed) {
            New-Item -ItemType Directory -Path (Split-Path -Parent $saved) -Force | Out-Null
            [IO.File]::Copy($target,$saved,$false)
        }
        $prepared.Add([pscustomobject]@{Source=$source;Target=$target;Saved=$saved;Existed=$existed;Hash=$file.sha256})
    }
    foreach ($file in $prepared) {
        New-Item -ItemType Directory -Path (Split-Path -Parent $file.Target) -Force | Out-Null
        $temporary=$file.Target+'.update-'+[guid]::NewGuid().ToString('N')
        [IO.File]::Copy($file.Source,$temporary,$false)
        if ($file.Existed) { [IO.File]::Replace($temporary,$file.Target,[System.Management.Automation.Language.NullString]::Value) }
        else { [IO.File]::Move($temporary,$file.Target) }
        $applied.Add($file)
    }
    foreach ($file in $prepared) {
        if ((File-Hash $file.Target) -ne $file.Hash) { throw '安装后校验失败。' }
    }
    @{status='success';version=$plan.version;time=(Get-Date).ToString('s')} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $updateRoot 'last-result.json') -Encoding UTF8
    $restart=$true
} catch {
    $failureMessage=$_.Exception.Message
    $rollback=$true
    for ($i=$applied.Count-1;$i -ge 0;$i--) {
        $file=$applied[$i]
        try {
            if ($file.Existed) { [IO.File]::Copy($file.Saved,$file.Target,$true) }
            elseif (Test-Path -LiteralPath $file.Target) { Remove-Item -LiteralPath $file.Target -Force }
        } catch { $rollback=$false }
    }
    Write-Output ('安装失败：'+$failureMessage)
    @{status='failed';rollback=$rollback;applied_count=$applied.Count;plan_sha256=$PlanSha256;message=$failureMessage;time=(Get-Date).ToString('s')} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $updateRoot 'last-result.json') -Encoding UTF8
    $restart=$rollback -and $locked
} finally {
    if ($lock) { try { $lock.Unlock(0,1) } catch {}; $lock.Dispose() }
}
if ($restart -and -not $NoRestart -and (Test-Path -LiteralPath $exe)) {
    foreach ($name in @((Get-ChildItem Env: | Where-Object { $_.Name -eq '_MEIPASS2' -or $_.Name -like '_PYI_*' -or $_.Name -in @('TCL_LIBRARY','TK_LIBRARY') } | Select-Object -ExpandProperty Name))) {
        [Environment]::SetEnvironmentVariable($name,$null,'Process')
    }
    $env:PYINSTALLER_RESET_ENVIRONMENT='1'
    try {
        $restarted=Start-Process -FilePath $exe -WorkingDirectory $appDir -WindowStyle Normal -PassThru
        $outcome=Get-Content -LiteralPath (Join-Path $updateRoot 'last-result.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $outcome | Add-Member -MemberType NoteProperty -Name restarted_pid -Value $restarted.Id -Force
        $outcome | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $updateRoot 'last-result.json') -Encoding UTF8
    } catch {
        @{status='restart_failed';version=$plan.version;message=$_.Exception.Message;time=(Get-Date).ToString('s')} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $updateRoot 'last-result.json') -Encoding UTF8
        throw
    }
}
