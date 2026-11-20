#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('Menu', 'Check', 'Status', 'Preview', 'Renew')]
    [string]$Mode = 'Check',

    [ValidatePattern('^[0-9a-fA-F,-]+$')]
    [string]$LicenseIds
)

$ErrorActionPreference = 'Stop'
$reminderDays = 7
$campusHosts = @('210.70.253.10', '192.168.168.1')
$kmsPort = 1688
$officeYears = [ordered]@{ '15' = 2013; '16' = 2016; '19' = 2019; '21' = 2021; '24' = 2024 }
$newLine = [Environment]::NewLine
$title = '澎科大 Office 授權提醒'
$vpnHelpUrl = 'https://www.npu.edu.tw/sub/faq/Details.aspx?Parser=26,32,311,,,,2145'
$networkHelp = '學校免費提供的「NPU」無線網路無法授權 Office。只連上「NPU」或人在校外時，請先連上學校 VPN。'
$manual = "續期需要連上可授權的校內網路。$networkHelp$newLine"
$manual += "學校 VPN 下載與教學$newLine$vpnHelpUrl$newLine"
$manual += '要切換授權設定時，請以系統管理員身分執行 office-kms.bat，按 Enter 開始，再選擇目前使用的網路。'
$windowsPowerShell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$programFiles = [Environment]::GetFolderPath('ProgramFiles')
$installedPath = Join-Path $programFiles "NPU-KMS-Monitor\NPU-KMS-Monitor-$sid.ps1"

function Write-MonitorLog([string]$Message) {
    try {
        $directory = Join-Path $env:LOCALAPPDATA 'NPU-KMS-Monitor'
        $null = New-Item -ItemType Directory -Path $directory -Force
        $log = Join-Path $directory 'monitor.log'
        if ((Test-Path -LiteralPath $log) -and (Get-Item -LiteralPath $log).Length -ge 262144) {
            Move-Item -LiteralPath $log -Destination "$log.previous" -Force
        }
        Add-Content -LiteralPath $log -Encoding UTF8 -Value ('{0:yyyy-MM-dd HH:mm:ss} {1}' -f (Get-Date), $Message)
    } catch {
        Write-Warning "無法寫入提醒紀錄。$($_.Exception.Message)"
    }
}

function Show-Notice([string]$Message, [switch]$AskToRenew) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $form = [Windows.Forms.Form]::new()
    $timer = [Windows.Forms.Timer]::new()
    try {
        $form.Text = $title
        $form.AutoScaleMode = 'Dpi'
        $form.StartPosition = 'CenterScreen'
        $form.ClientSize = [Drawing.Size]::new(580, 250)
        $form.MinimizeBox = $false
        $form.MaximizeBox = $false
        $form.FormBorderStyle = 'FixedDialog'
        $form.TopMost = $true

        $text = [Windows.Forms.TextBox]::new()
        $text.Multiline = $true
        $text.ReadOnly = $true
        $text.ScrollBars = 'Vertical'
        $text.Dock = 'Fill'
        $text.Text = $Message
        $form.Controls.Add($text)

        $buttons = [Windows.Forms.FlowLayoutPanel]::new()
        $buttons.Dock = 'Bottom'
        $buttons.Height = 48
        $buttons.FlowDirection = 'RightToLeft'
        $buttons.Padding = [Windows.Forms.Padding]::new(8)
        $form.Controls.Add($buttons)

        $later = [Windows.Forms.Button]::new()
        $later.Text = '確定'
        $later.AutoSize = $true
        $later.DialogResult = 'Cancel'
        $buttons.Controls.Add($later)
        $form.CancelButton = $later
        $form.AcceptButton = $later
        if ($AskToRenew) {
            $later.Text = '稍後提醒'
            $renew = [Windows.Forms.Button]::new()
            $renew.Text = '立即續期'
            $renew.AutoSize = $true
            $renew.DialogResult = 'Yes'
            $buttons.Controls.Add($renew)
        }
        $form.ActiveControl = $later
        $timer.Interval = 120000
        $timer.Add_Tick({ $form.DialogResult = [Windows.Forms.DialogResult]::Cancel }.GetNewClosure())
        $timer.Start()
        $form.ShowDialog() -eq [Windows.Forms.DialogResult]::Yes
    } finally {
        $timer.Dispose()
        $form.Dispose()
    }
}

function Get-OfficeKmsLicenses {
    $filter = "ApplicationID = '0ff1ce15-a989-479d-af46-f275c6370663' AND PartialProductKey IS NOT NULL"
    Get-CimInstance -Namespace 'root/cimv2' -ClassName SoftwareLicensingProduct -Filter $filter |
        Where-Object { $_.Description -match '\bVOLUME_KMSCLIENT\b' -and $_.ID -and $_.PartialProductKey -and -not $_.LicenseIsAddon }
}

function Test-NeedsRenewal($License) {
    $License.LicenseStatus -ne 1 -or
        ($null -ne $License.GracePeriodRemaining -and $License.GracePeriodRemaining -le ($reminderDays * 1440))
}

function Test-CampusEndpoint($License) {
    $License.KeyManagementServiceMachine -in $campusHosts -and
        $License.KeyManagementServicePort -in @($null, 0, $kmsPort)
}

function Test-RenewedLicense($Before, $After) {
    $After -and $After.ID -eq $Before.ID -and $After.LicenseStatus -eq 1 -and
        $After.GracePeriodRemaining -gt 0 -and
        ($Before.LicenseStatus -ne 1 -or $After.GracePeriodRemaining -gt $Before.GracePeriodRemaining)
}

function Get-OfficeYear($License) {
    if ($License.Name -match '2013|2016|2019|2021|2024') { return [int]$Matches[0] }
    if ($License.Name -match '\bOffice\s*(15|16|19|21|24)(?=\s|,|[A-Za-z]|$)') {
        return $officeYears[$Matches[1]]
    }
}

function Get-ProductName($License) {
    $year = Get-OfficeYear $License
    if (-not $year) { return "Office 授權（$($License.ID)）" }
    if ($License.Name -match 'Visio') { $product = "Visio $year" }
    elseif ($License.Name -match 'Project') { $product = "Project $year" }
    else { $product = "Office $year" }

    if ($License.Name -match 'ProPlus') { return "$product 專業增強版" }
    if ($License.Name -match 'Standard|Std') { return "$product 標準版" }
    if ($License.Name -match 'Professional|ProVL|Pro[0-9]{4}') { return "$product 專業版" }
    return $product
}

function Get-LicenseSummary($License) {
    $name = Get-ProductName $License
    if ($License.LicenseStatus -ne 1) { return "$name 需要啟用" }
    if ($null -eq $License.GracePeriodRemaining) { return "$name 有效，剩餘時間無法讀取" }
    $days = [math]::Ceiling($License.GracePeriodRemaining / 1440)
    return "$name 有效，剩餘約 $days 天"
}

function Invoke-LicenseMethod($License, [string]$Method, [hashtable]$Arguments) {
    $parameters = @{ InputObject = $License; MethodName = $Method; OperationTimeoutSec = 60 }
    if ($Arguments) { $parameters.Arguments = $Arguments }
    $result = Invoke-CimMethod @parameters
    if ($null -eq $result.ReturnValue) { throw 'Windows 授權服務未回報結果。' }
    if ($result.ReturnValue -ne 0) {
        $errorCode = [BitConverter]::ToInt32([BitConverter]::GetBytes([uint32]$result.ReturnValue), 0)
        throw [Runtime.InteropServices.COMException]::new('Windows 授權服務回報錯誤。', $errorCode)
    }
}

function Invoke-OfficeRenewal {
    if ([string]::IsNullOrWhiteSpace($LicenseIds)) { throw '請由 Office 授權提醒視窗啟動續期。' }

    $requested = @($LicenseIds.Split(',') | ForEach-Object { ([guid]$_).ToString() } | Select-Object -Unique)
    $installed = @(Get-OfficeKmsLicenses)
    $failed = $false
    foreach ($id in $requested) {
        $license = $installed | Where-Object ID -eq $id | Select-Object -First 1
        if (-not $license -or -not (Test-CampusEndpoint $license)) {
            $failed = $true
            continue
        }
        if (-not (Test-NeedsRenewal $license)) { continue }
        try {
            Invoke-LicenseMethod $license 'Activate'
        } catch {
            [Console]::Error.WriteLine($_.Exception.Message)
            $failed = $true
        }
    }
    if ($failed) { return 1 }
    return 0
}

function Invoke-CampusActivation($License, [string]$CampusHost) {
    $name = Get-ProductName $License
    $previousHost = $License.KeyManagementServiceMachine
    $previousPort = $License.KeyManagementServicePort
    $restoreHost = $false
    $restorePort = $false
    $current = $null
    $stage = '設定校方授權主機'
    try {
        if ($previousHost -ne $CampusHost) {
            $restoreHost = $true
            Invoke-LicenseMethod $License 'SetKeyManagementServiceMachine' @{ MachineName = $CampusHost }
        }
        if ($previousPort -notin @($null, 0, $kmsPort)) {
            $stage = '設定授權連接埠'
            $restorePort = $true
            Invoke-LicenseMethod $License 'SetKeyManagementServicePort' @{ PortNumber = [uint32]$kmsPort }
        }

        $activationError = $null
        try {
            Invoke-LicenseMethod $License 'Activate'
        } catch {
            # The provider can report an error to PowerShell after KMS has renewed the license.
            # Re-read the product before deciding whether the operation failed.
            $activationError = $_
        }
        $stage = '核對 Office 授權狀態'
        for ($attempt = 0; $attempt -lt 6; $attempt++) {
            $current = Get-OfficeKmsLicenses | Where-Object ID -eq $License.ID | Select-Object -First 1
            if ($current -and $current.LicenseStatus -eq 1 -and $current.GracePeriodRemaining -gt 0 -and
                $current.KeyManagementServiceMachine -eq $CampusHost -and (Test-CampusEndpoint $current) -and
                (-not $activationError -or (Test-RenewedLicense $License $current))) {
                Write-Host ('授權成功。' + (Get-LicenseSummary $current))
                return $true
            }
            if ($attempt -lt 5) { Start-Sleep -Seconds 2 }
        }
        if ($activationError) {
            $stage = '送出 Office 授權請求'
            throw $activationError
        }
        throw 'Office 的授權狀態尚未更新。'
    } catch {
        if ($current -and $current.LicenseStatus -eq 1 -and $current.GracePeriodRemaining -gt 0) {
            Write-Host (Get-LicenseSummary $current)
            Write-Host '這次未能確認效期延長。若只要開啟到期提醒，回首頁輸入 R，再選 14。' -ForegroundColor Yellow
        } else {
            Write-Host "未能確認 $name 授權成功。請確認所選網路可連上校方授權服務。" -ForegroundColor Yellow
        }
        $detail = $_.Exception.Message
        $reportedCode = [regex]::Match("$($_.ErrorDetails.Message) $detail", '\b0x[0-9a-fA-F]{8}\b')
        $code = if ($reportedCode.Success) { $reportedCode.Value } else { '0x{0:X8}' -f $_.Exception.HResult }
        Write-Host "未能完成的步驟：$stage（$code）。" -ForegroundColor Yellow
        if ($detail -match '^[\u3400-\u9FFF]') { Write-Host $detail -ForegroundColor Yellow }
    }

    if ($restoreHost) {
        try {
            if ([string]::IsNullOrEmpty($previousHost)) {
                Invoke-LicenseMethod $License 'ClearKeyManagementServiceMachine'
            } else {
                Invoke-LicenseMethod $License 'SetKeyManagementServiceMachine' @{ MachineName = $previousHost }
            }
        } catch {
            Write-Host '無法還原原授權主機，請聯絡圖書資訊館資訊組。' -ForegroundColor Yellow
        }
    }
    if ($restorePort) {
        try {
            Invoke-LicenseMethod $License 'SetKeyManagementServicePort' @{ PortNumber = [uint32]$previousPort }
        } catch {
            Write-Host '無法還原原連接埠，請聯絡圖書資訊館資訊組。' -ForegroundColor Yellow
        }
    }
    return $false
}

function Write-LicenseStatus {
    $licenses = @(Get-OfficeKmsLicenses)
    if ($licenses.Count -eq 0) { Write-Host '本機沒有可讀取的 Office KMS 大量授權。' }
    foreach ($license in $licenses) {
        Write-Host (Get-LicenseSummary $license)
        $hostName = $license.KeyManagementServiceMachine
        if (-not $hostName) { $hostName = '尚未指定' }
        Write-Host "授權主機：$hostName"
        $needed = if (Test-NeedsRenewal $license) { '是' } else { '否' }
        Write-Host "需要續期：$needed"
        Write-Host ('授權狀態代碼：0x{0:X8}' -f [uint32]$license.LicenseStatusReason)
        Write-Host ''
    }
}

function Invoke-ReminderSetup([string]$Mode) {
    $installer = Join-Path $PSScriptRoot 'Install-NPU-KMS-Monitor.ps1'
    if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) {
        Write-Host '找不到提醒安裝程式。請將完整壓縮檔解壓縮到同一資料夾。' -ForegroundColor Yellow
        return
    }
    & $installer -Mode $Mode
    if ($LASTEXITCODE -ne 0) {
        Write-Host '未能變更提醒設定。請確認 Windows 工作排程器可用，或聯絡圖書資訊館資訊組。' -ForegroundColor Yellow
    }
}

function Start-GuidedActivation {
    $licenses = @(Get-OfficeKmsLicenses)
    $years = @($licenses | ForEach-Object { Get-OfficeYear $_ } | Sort-Object -Unique)
    if ($years.Count -eq 0) {
        Write-Host '找不到可認證的 Office。請先安裝學校提供的 Office 大量授權版，再重新執行本批次檔。' -ForegroundColor Yellow
        return
    }

    $year = $years[0]
    if ($years.Count -gt 1) {
        Write-Host "找到這些 Office 版本：$($years -join '、')"
        do {
            $entry = ([string](Read-Host '輸入你要授權的年份，或直接按 Enter 返回')).Trim()
            if (-not $entry) { return }
            if ($entry -match '^[0-9]{4}$' -and [int]$entry -in $years) {
                $year = [int]$entry
                break
            }
            Write-Host "請輸入上方列出的年份：$($years -join '、')。" -ForegroundColor Yellow
        } while ($true)
    }

    $selected = @($licenses | Where-Object { (Get-OfficeYear $_) -eq $year })
    $productNames = ($selected | ForEach-Object { Get-ProductName $_ } | Select-Object -Unique) -join '、'
    Write-Host ''
    Write-Host "準備認證 $productNames。"
    Write-Host 'Office 必須連上可授權的校內網路。你現在用哪一種網路？'
    Write-Host '  1  已連上可用於 Office 授權的校內網路'
    Write-Host '  2  已連上學校 VPN'
    Write-Host $networkHelp -ForegroundColor Yellow
    Write-Host '學校 VPN 下載與教學'
    Write-Host $vpnHelpUrl
    do {
        $networkChoice = ([string](Read-Host '輸入 1 或 2 開始授權，直接按 Enter 返回')).Trim()
        if (-not $networkChoice) { return }
        if ($networkChoice -in @('1', '2')) { break }
        Write-Host '請輸入 1（校內授權網路）或 2（學校 VPN）。' -ForegroundColor Yellow
    } while ($true)

    $campusHost = $campusHosts[[int]$networkChoice - 1]
    $network = if ($networkChoice -eq '2') { '學校 VPN' } else { '校內授權網路' }
    Write-Host "正在用 $network 認證 $productNames，請稍候。"
    $succeeded = 0
    foreach ($license in $selected) {
        if (Invoke-CampusActivation $license $campusHost) { $succeeded++ }
    }
    if ($succeeded -eq 0) { return }
    if ($succeeded -lt $selected.Count) { Write-Host '有部分產品的授權仍待處理。' -ForegroundColor Yellow }

    Write-Host ''
    Write-Host '要在 Office 快到期時收到通知嗎？'
    Write-Host '開啟後，Windows 每天 09:00 檢查，剩餘 7 天內或已到期時跳出提醒。'
    Write-Host '以後要關閉時，重新執行本批次檔，輸入 R，再選 13。'
    do {
        $answer = ([string](Read-Host '輸入 1 開啟提醒，或直接按 Enter 略過')).Trim()
        if (-not $answer) {
            Write-Host '本次略過提醒設定。'
            return
        }
        if ($answer -eq '1') {
            Invoke-ReminderSetup 'Install'
            return
        }
        Write-Host '請輸入 1 開啟提醒，或直接按 Enter 略過。' -ForegroundColor Yellow
    } while ($true)
}

function Show-ReminderMenu {
    while ($true) {
        Write-Host ''
        Write-Host '查詢授權與管理提醒'
        Write-Host '  11  查詢授權狀態    查看 Office 是否已授權、還剩幾天'
        Write-Host '  12  預覽到期提醒    看看日後會跳出的視窗'
        Write-Host '  13  移除到期提醒    關閉目前帳號的提醒'
        Write-Host '  14  開啟到期提醒    先前略過的人可在這裡補開'
        $choice = ([string](Read-Host '輸入功能編號，或直接按 Enter 返回首頁')).Trim()
        switch ($choice) {
            '11' { Write-LicenseStatus }
            '12' { $null = Show-Notice '這是 Office 到期提醒的預覽。兩個按鈕都只會關閉預覽。' -AskToRenew }
            '13' { Invoke-ReminderSetup 'Remove' }
            '14' { Invoke-ReminderSetup 'Install' }
            '' { return }
            default { Write-Host '請輸入 11、12、13 或 14，或直接按 Enter 返回首頁。' -ForegroundColor Yellow }
        }
    }
}

function Show-ActivationHome {
    while ($true) {
        Write-Host '國立澎湖科技大學　Office 校園授權'
        Write-Host '這個程式會替已安裝的校方 Office 完成授權。授權後可選擇開啟到期提醒。'
        Write-Host ('在開始之前請確保您已經連上NPU校內網路，或者依照學校網路"{0}" 使用FortiClient連上校內提供授權的專用VPN。' -f $vpnHelpUrl)
        Write-Host ''
        Write-Host '直接按 Enter 開始授權，程式會先辨識你安裝的 Office。' -ForegroundColor Cyan
        Write-Host '也可以直接輸入 R 查詢授權或管理提醒。輸入 Q 結束。'
        $choice = ([string](Read-Host '下一步')).Trim()
        try {
            switch ($choice) {
                '' { Start-GuidedActivation }
                'R' { Show-ReminderMenu }
                'Q' { return }
                default { Write-Host '請直接按 Enter 開始授權，或輸入 R 或 Q。' -ForegroundColor Yellow }
            }
        } catch {
            Write-Host '無法完成操作。請確認 Windows 授權與工作排程服務可用，或聯絡圖書資訊館資訊組。' -ForegroundColor Yellow
            Write-Host ('錯誤碼 0x{0:X8}。' -f $_.Exception.HResult) -ForegroundColor Yellow
        }
        Write-Host ''
    }
}

function Show-Reminder {
    $licenses = @(Get-OfficeKmsLicenses)
    if ($licenses.Count -eq 0) {
        Write-MonitorLog '找不到已安裝的 Office KMS 用戶端授權。'
        return
    }
    $due = @($licenses | Where-Object { Test-NeedsRenewal $_ })
    Write-MonitorLog "讀取到 $($licenses.Count) 個 Office KMS 授權，$($due.Count) 個需要處理。"
    if ($due.Count -eq 0) { return }

    $summaries = ($due | ForEach-Object { Get-LicenseSummary $_ }) -join $newLine
    $eligible = @($due | Where-Object { Test-CampusEndpoint $_ })
    if ($eligible.Count -eq 0) {
        Write-MonitorLog '需要處理的授權沒有設定核對過的校方 KMS 主機。'
        $null = Show-Notice "$summaries$newLine$newLine$manual"
        return
    }
    $message = "$summaries$newLine$newLine$manual$newLine$newLine"
    $message += '要沿用已設定的校方主機自動續期嗎？按「立即續期」後，請確認 Windows 的系統管理員提示。'
    if (-not (Show-Notice $message -AskToRenew)) {
        Write-MonitorLog '使用者選擇稍後續期，或提醒已逾時。'
        return
    }
    if ($PSCommandPath -ne $installedPath) { throw '請先安裝提醒程式，再從已安裝的排程續期。' }

    $ids = ($eligible | ForEach-Object { ([guid]$_.ID).ToString() }) -join ','
    $arguments = '-NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy RemoteSigned -File "{0}" -Mode Renew -LicenseIds "{1}"' -f $installedPath, $ids
    try {
        $process = Start-Process -FilePath $windowsPowerShell -ArgumentList $arguments -Verb RunAs -Wait -PassThru
        try { Write-MonitorLog "續期程序回傳碼 $($process.ExitCode)。" } finally { $process.Dispose() }
    } catch {
        Write-MonitorLog "無法啟動管理員續期程序。$($_.Exception.Message)"
        $null = Show-Notice "續期程序未啟動。$newLine$newLine$manual"
        return
    }

    # Verify by product ID. An unrelated license dropping out cannot count as renewal.
    $results = @{}
    for ($attempt = 0; $attempt -lt 6; $attempt++) {
        $current = @(Get-OfficeKmsLicenses)
        foreach ($before in $eligible) {
            $after = $current | Where-Object ID -eq $before.ID | Select-Object -First 1
            if (Test-RenewedLicense $before $after) {
                $results[$before.ID] = Get-LicenseSummary $after
            }
        }
        if ($results.Count -eq $eligible.Count -or $attempt -eq 5) { break }
        Start-Sleep -Seconds 2
    }
    $lines = foreach ($before in $due) {
        if ($results.ContainsKey($before.ID)) { '已確認續期。' + $results[$before.ID] }
        else { '待處理。' + (Get-ProductName $before) }
    }
    Write-MonitorLog "已確認 $($results.Count) 個續期，待處理 $($due.Count - $results.Count) 個。"
    $notice = $lines -join $newLine
    if ($results.Count -lt $due.Count) { $notice += "$newLine$newLine$manual" }
    $null = Show-Notice $notice
}

try {
    if ($Mode -ne 'Preview') {
        Import-Module (Join-Path $PSHOME 'Modules\CimCmdlets\CimCmdlets.psd1')
    }
    switch ($Mode) {
        'Menu' {
            $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
            $principal = [Security.Principal.WindowsPrincipal]::new($identity)
            if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                Write-Host '請在 office-kms.bat 上按右鍵，選擇「以系統管理員身分執行」。'
                $null = Read-Host '按 Enter 結束'
                exit 1
            }
            Show-ActivationHome
        }
        'Check' { Show-Reminder }
        'Status' { Write-LicenseStatus }
        'Preview' { $null = Show-Notice '這是 Office 到期提醒的預覽。兩個按鈕都只會關閉預覽。' -AskToRenew }
        'Renew' {
            $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
            $principal = [Security.Principal.WindowsPrincipal]::new($identity)
            if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                throw 'Office 續期需要系統管理員權限。'
            }
            exit (Invoke-OfficeRenewal)
        }
    }
} catch {
    if ($Mode -eq 'Check') { Write-MonitorLog "授權檢查失敗。$($_.Exception.Message)" }
    [Console]::Error.WriteLine('無法完成 Office 授權操作。請聯絡圖書資訊館資訊組。')
    if ($Mode -eq 'Menu') { $null = Read-Host '按 Enter 結束' }
    exit 1
}
