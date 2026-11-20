#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('Install', 'Remove')]
    [string]$Mode = 'Install'
)

$ErrorActionPreference = 'Stop'
$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$sid = $identity.User.Value
$taskName = "NPU-KMS-Monitor-$sid"
$taskPath = '\'
$directory = Join-Path ([Environment]::GetFolderPath('ProgramFiles')) 'NPU-KMS-Monitor'
$target = Join-Path $directory "NPU-KMS-Monitor-$sid.ps1"
$source = Join-Path $PSScriptRoot 'NPU-KMS-Monitor.ps1'
$powerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$administrators = [Security.Principal.SecurityIdentifier]::new('S-1-5-32-544')
$system = [Security.Principal.SecurityIdentifier]::new('S-1-5-18')
$users = [Security.Principal.SecurityIdentifier]::new('S-1-5-32-545')

function Assert-RegularPath([string]$Path) {
    if ((Get-Item -LiteralPath $Path -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
        throw "部署路徑包含重新解析點：$Path"
    }
}

function Assert-InstallDirectory {
    Assert-RegularPath $directory
    $acl = Get-Acl -LiteralPath $directory
    if (-not $acl.AreAccessRulesProtected -or $acl.GetOwner([Security.Principal.SecurityIdentifier]) -ne $administrators) {
        throw "部署目錄的擁有者或繼承權限不符合預期：$directory"
    }
    $rules = $acl.GetAccessRules($true, $true, [Security.Principal.SecurityIdentifier])
    foreach ($rule in $rules) {
        if ($rule.AccessControlType -ne 'Allow' -or $rule.IdentityReference -in @($administrators, $system)) { continue }
        $extraRights = $rule.FileSystemRights -band (-bnot ([Security.AccessControl.FileSystemRights]::ReadAndExecute -bor [Security.AccessControl.FileSystemRights]::Synchronize))
        if ($rule.IdentityReference -ne $users -or $extraRights) {
            throw "部署目錄具有未預期的寫入或存取權限：$directory"
        }
    }
}

function Initialize-InstallDirectory {
    if (-not (Test-Path -LiteralPath $directory)) {
        $acl = [Security.AccessControl.DirectorySecurity]::new()
        $acl.SetOwner($administrators)
        $acl.SetAccessRuleProtection($true, $false)
        $inheritance = [Security.AccessControl.InheritanceFlags]'ContainerInherit, ObjectInherit'
        foreach ($entry in @(
            @{ Principal = $administrators; Rights = 'FullControl' }
            @{ Principal = $system; Rights = 'FullControl' }
            @{ Principal = $users; Rights = 'ReadAndExecute' }
        )) {
            $rule = [Security.AccessControl.FileSystemAccessRule]::new($entry.Principal, $entry.Rights, $inheritance, 'None', 'Allow')
            $acl.AddAccessRule($rule)
        }
        $null = [IO.Directory]::CreateDirectory($directory, $acl)
    }
    Assert-InstallDirectory
}

function Install-Monitor {
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "找不到 $source" }
    if ((Get-ExecutionPolicy) -eq 'Restricted') {
        throw '此電腦的 PowerShell 執行原則禁止指令碼。請由資訊組檢查部署原則。'
    }
    $policy = Get-ExecutionPolicy -List
    $enforcedPolicy = $policy | Where-Object { $_.Scope -in @('MachinePolicy','UserPolicy') -and $_.ExecutionPolicy -ne 'Undefined' } | Select-Object -First 1
    if ($enforcedPolicy.ExecutionPolicy -eq 'AllSigned' -and (Get-AuthenticodeSignature -LiteralPath $source).Status -ne 'Valid') {
        throw '這台電腦要求簽署指令碼。請使用資訊組簽署的版本。'
    }
    if ((Get-Item -LiteralPath $source -Stream Zone.Identifier -ErrorAction SilentlyContinue) -and
        (Get-AuthenticodeSignature -LiteralPath $source).Status -ne 'Valid') {
        throw '監控腳本仍帶有網際網路下載標記。確認壓縮檔的發行來源後，請依使用說明解除封鎖並重新解壓縮。'
    }
    $tokens = $null
    $parseErrors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count -ne 0) { throw "監控腳本無法載入：$($parseErrors[0].Message)" }

    $actionArguments = '-NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy RemoteSigned -File "{0}" -Mode Check' -f $target
    $action = New-ScheduledTaskAction -Execute $powerShell -Argument $actionArguments
    $trigger = New-ScheduledTaskTrigger -Daily -At ([datetime]::Today.AddHours(9))
    $principal = New-ScheduledTaskPrincipal -UserId $sid -LogonType Interactive -RunLevel Limited
    $settings = @{
        StartWhenAvailable = $true
        AllowStartIfOnBatteries = $true
        DontStopIfGoingOnBatteries = $true
        MultipleInstances = 'IgnoreNew'
        ExecutionTimeLimit = New-TimeSpan -Minutes 15
    }
    $taskSettings = New-ScheduledTaskSettingsSet @settings

    Initialize-InstallDirectory
    $temporary = Join-Path $directory ('deploy-' + [guid]::NewGuid() + '.ps1')
    try {
        [IO.File]::WriteAllBytes($temporary, [IO.File]::ReadAllBytes($source))
        $fileAcl = Get-Acl -LiteralPath $temporary
        $fileAcl.SetOwner($administrators)
        Set-Acl -LiteralPath $temporary -AclObject $fileAcl
        if (Test-Path -LiteralPath $target) {
            Assert-RegularPath $target
            $targetAcl = Get-Acl -LiteralPath $target
            if ($targetAcl.GetOwner([Security.Principal.SecurityIdentifier]) -ne $administrators -or $targetAcl.AreAccessRulesProtected) {
                throw "原監控腳本的擁有者或繼承權限不符合預期：$target"
            }
            if ($targetAcl.GetAccessRules($true, $false, [Security.Principal.SecurityIdentifier]).Count -gt 0) {
                throw "原監控腳本具有未預期的個別存取權限：$target"
            }
            [IO.File]::Replace($temporary, $target, [NullString]::Value)
        } else {
            [IO.File]::Move($temporary, $target)
        }
        $null = Register-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Action $action -Trigger $trigger -Principal $principal -Settings $taskSettings -Force
    } finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary }
    }
    Write-Host '已為目前帳號安裝每日 09:00 Office 授權提醒。'
    Write-Host '要關閉提醒，請再次執行 office-kms.bat，輸入 R，再選「13 移除到期提醒」。'
}

function Remove-Monitor {
    $existing = Get-ScheduledTask -TaskPath $taskPath | Where-Object TaskName -eq $taskName
    if ($existing) {
        if ($existing.State -eq 'Running') { Stop-ScheduledTask -InputObject $existing }
        Unregister-ScheduledTask -InputObject $existing -Confirm:$false
    }
    if (Test-Path -LiteralPath $target) {
        Assert-InstallDirectory
        Assert-RegularPath $target
        Remove-Item -LiteralPath $target
    }
    Write-Host '已移除目前帳號的 Office 授權提醒。'
}

try {
    $currentUser = [Security.Principal.WindowsPrincipal]::new($identity)
    if (-not $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw '請使用要收到提醒的帳號，以系統管理員身分開啟本程式。'
    }
    Import-Module (Join-Path $PSHOME 'Modules\ScheduledTasks\ScheduledTasks.psd1')
    $lock = [Threading.Mutex]::new($false, "Global\$taskName")
    try {
        try { $acquired = $lock.WaitOne(0) } catch [Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { throw '同一帳號已有安裝或移除正在進行。' }
        try {
            if ($Mode -eq 'Install') { Install-Monitor } else { Remove-Monitor }
        } finally {
            $lock.ReleaseMutex()
        }
    } finally {
        $lock.Dispose()
    }
    exit 0
} catch {
    Write-Host ('提醒設定失敗。錯誤碼 0x{0:X8}。' -f $_.Exception.HResult) -ForegroundColor Yellow
    if ($_.Exception.Message -match '^[\u3400-\u9FFF]') { Write-Host $_.Exception.Message -ForegroundColor Yellow }
    exit 1
}
