# =====================================================
# 图形版 智能复制校验工具 (Release 1.1)
# 功能：弹出文件夹选择窗口，robocopy 复制 + 多线程 SHA256 校验 + 自动修复
# 兼容：PowerShell 5.1 / 无需管理员权限 / BAT 启动
# =====================================================

Add-Type -AssemblyName System.Windows.Forms

# ==================== 全局变量 ====================
$Script:LogDir        = Join-Path $env:USERPROFILE "Desktop"
$Script:LogFile       = Join-Path $Script:LogDir ("Program_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$Script:FailedLogFile = Join-Path $Script:LogDir ("FailedFiles_{0}.txt" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$Script:ReportFile    = Join-Path $Script:LogDir ("{0}_{1}.txt" -f '复制报告', (Get-Date -Format 'yyyyMMdd_HHmmss'))
$Script:RoboLogFile   = Join-Path $Script:LogDir ("robocopy_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

$Script:TotalCount    = 0
$Script:SkipCount     = 0
$Script:FixCount      = 0
$Script:ErrorCount    = 0
$Script:LongPathSkippedCount = 0
$Script:TotalSize     = [long]0
$Script:ErrorFiles    = [System.Collections.ArrayList]::new()
$Script:LongPathSkippedFiles = [System.Collections.ArrayList]::new()
$Script:StartTime     = $null
$Script:CopyPhaseStart= $null
$Script:CopyPhaseEnd  = $null
$Script:VerifyStart   = $null
$Script:VerifyEnd     = $null
$Script:RoboExitCode  = 0

# 日志缓冲区（减少磁盘 I/O 频率）
$Script:LogBuffer     = [System.Collections.ArrayList]::new()
$Script:LogFlushThreshold = 20

# ==================== Write-Log ====================
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","DEBUG")][string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $line = "[$ts] [$Level] $Message"

    # 缓冲写入，减少磁盘 I/O
    $null = $Script:LogBuffer.Add($line)
    if ($Script:LogBuffer.Count -ge $Script:LogFlushThreshold) {
        Flush-LogBuffer
    }

    switch ($Level) {
        "ERROR" { Write-Host $line -ForegroundColor Red }
        "WARN"  { Write-Host $line -ForegroundColor Yellow }
        default { Write-Host $line -ForegroundColor DarkGray }
    }
}

function Flush-LogBuffer {
    if ($Script:LogBuffer.Count -gt 0) {
        $content = $Script:LogBuffer -join "`r`n"
        try {
            Add-Content -Path $Script:LogFile -Value $content -Encoding UTF8 -ErrorAction SilentlyContinue
        }
        catch {
            # 日志写入失败不应中断主流程
        }
        $Script:LogBuffer.Clear()
    }
}

# ==================== Normalize-LongPath ====================
function Normalize-LongPath {
    param([Parameter(Mandatory)][string]$Path)
    # FIXED: 同时去除末尾反斜杠、空格和制表符
    $p = $Path.TrimEnd('\', ' ', "`t")
    # 已经是长路径格式
    if ($p.StartsWith("\\?\")) { return $p }
    # UNC: \\server\share → \\?\UNC\server\share
    if ($p.StartsWith("\\")) {
        return "\\?\UNC\$($p.Substring(2))"
    }
    # 本地: C:\... → \\?\C:\...
    return "\\?\$p"
}

# ==================== Select-Source ====================
function Select-Source {
    param([Parameter(Mandatory)][string]$Description)
    return Select-Folder -Description $Description -ShowNewFolder $false
}
function Select-Folder {
    param(
        [Parameter(Mandatory)][string]$Description,
        [bool]$ShowNewFolder = $false
    )
    $dialog = $null
    try {
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $Description
        $dialog.ShowNewFolderButton = $ShowNewFolder
        $dialog.RootFolder = [System.Environment+SpecialFolder]::Desktop
        $result = $dialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.SelectedPath
        }
        return $null
    }
    finally {
        if ($null -ne $dialog) { $dialog.Dispose() }
    }
}

# ==================== Write-ConsoleProgress ====================
function Write-ConsoleProgress {
    param(
        [string]$Activity,
        [int]$Current,
        [int]$Total,
        [string]$Status = "",
        [datetime]$StartTime = [datetime]::MinValue
    )
    $pct = if ($Total -gt 0) { [math]::Min(100, [math]::Round($Current / $Total * 100)) } else { 0 }
    $barLen = 30
    $filled = [math]::Round($barLen * $pct / 100)
    $empty = $barLen - $filled
    $bar = ("|" * $filled) + ("." * $empty)

    $etaStr = ""
    if ($StartTime -ne [datetime]::MinValue -and $Current -gt 0 -and $pct -gt 0) {
        $elapsed = (Get-Date) - $StartTime
        if ($pct -lt 100) {
            $remaining = [math]::Round(($elapsed.TotalSeconds / $Current) * ($Total - $Current))
            $min = [math]::Floor($remaining / 60)
            $sec = $remaining % 60
            $etaStr = " 剩余 ${min}m${sec}s"
        }
        else {
            $totalSec = [math]::Round($elapsed.TotalSeconds)
            $min = [math]::Floor($totalSec / 60)
            $sec = $totalSec % 60
            $etaStr = " 总耗时 ${min}m${sec}s"
        }
    }

    $line = "`r$Activity [$bar] $pct%  ($Current/$Total)$etaStr $Status                              "
    Write-Host $line -NoNewline -ForegroundColor Cyan
}

# ==================== Start-Robocopy ====================
function Start-Robocopy {
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$Target,
        [string]$LogFile,
        [string]$SingleFile = "",
        [int]$TotalFiles = 0
    )

    # 科研数据最优参数：
    # /E         递归子目录（含空目录）
    # /COPY:DAT  复制 数据+属性+时间（最安全，无权限要求）
    # /DCOPY:T   复制目录时间戳
    # /R:1       失败重试 1 次（避免无限卡死）
    # /W:3       重试间隔 3 秒
    # /MT:8      8 线程并行
    # /NP        不显示进度百分比（避免日志膨胀）
    # /XO        排除旧文件（避免覆盖更新的文件）
    # /FFT       允许 2 秒时间戳误差（FAT32/NTFS 兼容）
    # /NDL       不记录目录名（避免长路径/特殊字符问题）
    # /NJS       不输出作业名
    # FIXED: 移除 /NFL，改为读取 stdout 追踪复制进度
    $argList = [System.Collections.ArrayList]@()
    $null = $argList.Add("`"$Source`"")
    $null = $argList.Add("`"$Target`"")
    $null = $argList.Add("/E")
    $null = $argList.Add("/COPY:DAT")
    $null = $argList.Add("/DCOPY:T")
    $null = $argList.Add("/R:1")
    $null = $argList.Add("/W:3")
    $null = $argList.Add("/MT:8")
    $null = $argList.Add("/NP")
    $null = $argList.Add("/XO")
    $null = $argList.Add("/FFT")
    $null = $argList.Add("/NDL")   # 不记录目录名（避免长路径问题）
    $null = $argList.Add("/NJS")   # 不输出作业名

    if ($SingleFile -ne "") {
        $null = $argList.Add("/IF")
        $null = $argList.Add("`"$SingleFile`"")
    }

    Write-Log "Robocopy 参数: $($argList -join ' ')" "DEBUG"

    $proc = $null
    try {
        # 使用 .NET Process 类直接启动，避免 PowerShell Start-Process 的参数转义问题
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = "robocopy.exe"
        $startInfo.Arguments = $argList -join ' '
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        # FIXED: 不重定向 stdout/stderr，让 robocopy 直接输出到控制台，避免缓冲区满导致进程阻塞
        $startInfo.RedirectStandardOutput = $false
        $startInfo.RedirectStandardError = $false

        $proc = [System.Diagnostics.Process]::Start($startInfo)

        # IMPROVED: 每 3 秒统计目标目录文件数 + 排空 stderr 防止缓冲区满
        $targetForCount = $Target
        $totalForCount = $TotalFiles
        $lastDone = 0
        $timeoutMs = 300000
        $elapsedMs = [int]0
        $copyStartTime = Get-Date

        # IMPROVED: 立即显示初始进度条，不等 3 秒
        if ($totalForCount -gt 0) {
            try {
                $lastDone = [System.IO.Directory]::GetFiles($targetForCount, "*", [System.IO.SearchOption]::AllDirectories).Count
                Write-ConsoleProgress -Activity "复制中" -Current $lastDone -Total $totalForCount -StartTime $copyStartTime
            } catch {}
        }

        while (-not $proc.HasExited -and $elapsedMs -lt $timeoutMs) {
            Start-Sleep -Milliseconds 3000
            $elapsedMs += 3000
            if ($totalForCount -gt 0) {
                try {
                    $done = [System.IO.Directory]::GetFiles($targetForCount, "*", [System.IO.SearchOption]::AllDirectories).Count
                    if ($done -ne $lastDone) {
                        $lastDone = $done
                        Write-ConsoleProgress -Activity "复制中" -Current $done -Total $totalForCount -StartTime $copyStartTime
                    }
                } catch {}
            }
        }

        if (-not $proc.HasExited) {
            try { $proc.Kill() } catch {}
            Write-Log "Robocopy 超时（300秒），已强制终止" "WARN"
            return 8
        }
        $proc.WaitForExit()
        return $proc.ExitCode
    }
    finally {
        Write-Host ""
        if ($null -ne $proc) { $proc.Dispose() }
    }
}

# ==================== Get-SHA256（支持长路径，大缓冲区加速） ====================
function Get-SHA256 {
    param([Parameter(Mandatory)][string]$FilePath)

    $stream = $null
    try {
        $np = Normalize-LongPath -Path $FilePath
        $stream = [System.IO.File]::Open($np, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $buffer = [byte[]]::new(131072)
        try {
            do {
                $read = $stream.Read($buffer, 0, $buffer.Length)
                if ($read -gt 0) { $sha.TransformBlock($buffer, 0, $read, $null, 0) | Out-Null }
            } while ($read -gt 0)
            $sha.TransformFinalBlock([byte[]]::new(0), 0, 0) | Out-Null
            return ([System.BitConverter]::ToString($sha.Hash) -replace "-", "")
        }
        finally {
            if ($null -ne $sha) { $sha.Dispose() }
        }
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

# ==================== Get-FileSize（支持长路径） ====================
function Get-FileSize {
    param([Parameter(Mandatory)][string]$FilePath)
    try {
        $np = Normalize-LongPath -Path $FilePath
        $info = New-Object System.IO.FileInfo($np)
        return $info.Length
    }
    catch {
        return -1
    }
}

# ==================== Test-HasEnoughSpace ====================
function Test-HasEnoughSpace {
    param(
        [Parameter(Mandatory)][string]$TargetPath,
        [long]$RequiredBytes
    )
    try {
        $drive = [System.IO.Path]::GetPathRoot($TargetPath)
        if (-not $drive) { return $true }
        $driveInfo = Get-PSDrive -Name ($drive.TrimEnd('\').TrimEnd(':')) -ErrorAction SilentlyContinue
        if ($null -ne $driveInfo -and $driveInfo.Free -lt $RequiredBytes) { return $false }
        return $true
    } catch { return $true }
}

# ==================== Test-FileReadable（支持长路径） ====================
function Test-FileReadable {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [int]$MaxRetries = 2,
        [int]$RetryDelayMs = 500
    )
    $attempt = 0
    while ($attempt -lt $MaxRetries) {
        $stream = $null
        try {
            $np = Normalize-LongPath -Path $FilePath
            $stream = [System.IO.File]::Open(
                $np,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::ReadWrite
            )
            return $true
        }
        catch [System.IO.IOException] {
            $attempt++
            if ($attempt -lt $MaxRetries) {
                Start-Sleep -Milliseconds $RetryDelayMs
            }
        }
        catch [System.UnauthorizedAccessException] {
            # 权限问题，不应重试
            return $false
        }
        catch {
            # 其他异常，不重试
            return $false
        }
        finally {
            if ($null -ne $stream) { $stream.Dispose() }
        }
    }
    return $false
}

# ==================== Remove-CorruptedTarget ====================
function Remove-CorruptedTarget {
    param([Parameter(Mandatory)][string]$FilePath)
    try {
        # FIXED: 使用 -LiteralPath 支持长路径和特殊字符
        if (Test-Path -LiteralPath $FilePath -ErrorAction SilentlyContinue) {
            $item = Get-Item -LiteralPath $FilePath -Force -ErrorAction SilentlyContinue
            if ($null -ne $item) {
                $item.Attributes = [System.IO.FileAttributes]::Normal
            }
            Remove-Item -LiteralPath $FilePath -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 200
        }
    }
    catch {
        Write-Log "清理损坏目标文件失败: $FilePath - $($_.Exception.Message)" "WARN"
    }
}

# ==================== Repair-File ====================
function Repair-File {
    param(
        [Parameter(Mandatory)][string]$SourceFile,
        [Parameter(Mandatory)][string]$TargetFile,
        [bool]$CleanupFirst = $false
    )

    # 如果需要，先清理损坏的目标文件
    if ($CleanupFirst) {
        Remove-CorruptedTarget -FilePath $TargetFile
    }

    # 确保目标目录存在
    $targetDir = Split-Path $TargetFile -Parent
    if (-not (Test-Path -LiteralPath $targetDir -ErrorAction SilentlyContinue)) {
        try {
            New-Item -Path $targetDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        catch {
            throw [System.IO.DirectoryNotFoundException]::new("创建或访问目标目录失败: $targetDir - $($_.Exception.Message)")
        }
    }

    $sourceDir  = Split-Path $SourceFile -Parent
    $sourceBase = Split-Path $SourceFile -Leaf

    Write-Log "robocopy 单文件复制: $sourceBase" "INFO"
    $code = Start-Robocopy -Source $sourceDir -Target $targetDir -LogFile $Script:RoboLogFile -SingleFile $sourceBase

    # robocopy 返回码语义：
    # 0 = 无操作（文件已存在且相同）
    # 1 = 复制了一个文件
    # 2 = 复制了额外的文件
    # 3 = 复制了文件 + 额外文件
    # 4-7 = 不同组合的 复制/额外/不匹配/跳过
    # >= 8 = 错误（致命错误/文件被锁定/内存不足等）
    if ($code -ge 0 -and $code -lt 8) {
        # 等待磁盘缓存刷新（防止 Hash 时数据尚未落盘）
        Start-Sleep -Milliseconds 500

        # 轮询文件大小稳定性（最多等 3 秒）
        $stableCount = 0
        $lastSize = -1
        $pollCount = 0
        while ($pollCount -lt 6) {
            $currentSize = Get-FileSize -FilePath $TargetFile
            if ($currentSize -ge 0 -and $currentSize -eq $lastSize) {
                $stableCount++
                if ($stableCount -ge 2) { break }
            }
            else {
                $stableCount = 0
                $lastSize = $currentSize
            }
            Start-Sleep -Milliseconds 500
            $pollCount++
        }

        Write-Log "robocopy 复制成功 (返回码: $code): $sourceBase" "INFO"
        return $true
    }
    else {
        throw [System.IO.IOException]::new("Robocopy 复制失败，退出码: $code ($sourceBase)")
    }
}

# ==================== Invoke-WithRetry ====================
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)][ScriptBlock]$ScriptBlock,
        [string]$ErrorMessage = "",
        [int]$MaxRetries = 5
    )
    $backoffSeconds = @(5, 10, 20, 30, 60)
    $attempt = 0

    while ($attempt -lt $MaxRetries) {
        try {
            $result = & $ScriptBlock
            return @{ Success = $true; Result = $result }
        }
        catch {
            $attempt++
            $ex = $_.Exception
            $errMsg = $ex.Message
            $isRetryable = $false

            # 优先按异常类型判断，其次按消息内容判断
            # 注意：必须先检查子类再检查父类（-is 匹配完整继承链）
            if ($ex -is [System.IO.PathTooLongException]) {
                # 长路径不可重试
                $isRetryable = $false
            }
            elseif ($ex -is [System.IO.FileNotFoundException]) {
                # 文件不存在，可能暂时掉盘，可重试
                $isRetryable = $true
            }
            elseif ($ex -is [System.IO.DirectoryNotFoundException]) {
                $isRetryable = $true
            }
            elseif ($ex -is [System.IO.IOException]) {
                $isRetryable = $true
            }
            elseif ($ex -is [System.UnauthorizedAccessException]) {
                $isRetryable = $true
            }
            elseif ($ex -is [System.ComponentModel.Win32Exception]) {
                $isRetryable = $true
            }
            else {
                # 降级到消息匹配（覆盖所有 USB 常见错误）
                $diskPatterns = @(
                    "I/O device error",
                    "I/O",
                    "device",
                    "不存在的设备",
                    "设备未就绪",
                    "设备不存在",
                    "device not exist",
                    "not ready",
                    "not connected",
                    "not connected",
                    "CRC",
                    "Data Error",
                    "Device not ready",
                    "Semaphore timeout",
                    "找不到路径",
                    "网络路径找不到",
                    "参数错误",
                    "Network unavailable",
                    "Disk structure corrupted",
                    "The semaphore timeout",
                    "incorrect parameter",
                    "path not found",
                    "network path was not found",
                    "access is denied",
                    "unauthorized",
                    "path too long",
                    "file name too long",
                    "共享冲突",
                    "Sharing violation",
                    "being used by another process"
                )
                foreach ($pat in $diskPatterns) {
                    if ($errMsg -match [regex]::Escape($pat)) {
                        $isRetryable = $true
                        break
                    }
                }
            }

            if ($isRetryable -and $attempt -lt $MaxRetries) {
                $wait = if ($attempt -le $backoffSeconds.Count) { $backoffSeconds[$attempt - 1] } else { $backoffSeconds[-1] }
                Write-Log "可重试错误 (尝试 $attempt / $MaxRetries): [$($ex.GetType().Name)] $errMsg" "WARN"
                Write-Log "等待 $wait 秒后重试..." "INFO"
                Start-Sleep -Seconds $wait

                # 检查盘符是否恢复
                if ($null -ne $Script:TargetDriveRoot -and $Script:TargetDriveRoot -ne "") {
                    if (-not (Test-Path -LiteralPath $Script:TargetDriveRoot -ErrorAction SilentlyContinue)) {
                        Write-Log "盘符 $($Script:TargetDriveRoot) 未就绪，额外等待 10 秒..." "WARN"
                        Start-Sleep -Seconds 10
                    }
                }
            }
            else {
                # 不可重试或已达最大重试次数
                $fullMsg = if ($ErrorMessage -ne "") { "$ErrorMessage - [$($ex.GetType().Name)] $errMsg" } else { "[$($ex.GetType().Name)] $errMsg" }
                Write-Log $fullMsg "ERROR"
                return @{ Success = $false; Result = $null; Error = $fullMsg }
            }
        }
    }
    $finalMsg = if ($ErrorMessage -ne "") { "$ErrorMessage - 重试 $MaxRetries 次后仍失败" } else { "重试 $MaxRetries 次后仍失败" }
    Write-Log $finalMsg "ERROR"
    return @{ Success = $false; Result = $null; Error = $finalMsg }
}

# ==================== Write-Report ====================
function Write-Report {
    param(
        [string]$SourcePath,
        [string]$TargetPath,
        [bool]$DoCopy,
        [int]$Total,
        [int]$Skip,
        [int]$Fix,
        [int]$ErrCount,
        [int]$LongPathSkipped,
        [long]$TotalSize,
        [long]$FileCount
    )
    $endTime = Get-Date
    $totalDuration  = if ($Script:CopyPhaseStart) { ($endTime - $Script:CopyPhaseStart).TotalSeconds } else { 0 }
    $copyDuration   = if ($Script:CopyPhaseStart -and $Script:CopyPhaseEnd) { ($Script:CopyPhaseEnd - $Script:CopyPhaseStart).TotalSeconds } else { 0 }
    $verifyDuration = if ($Script:VerifyStart -and $Script:VerifyEnd) { ($Script:VerifyEnd - $Script:VerifyStart).TotalSeconds } else { 0 }
    $speedMB = if ($copyDuration -gt 0) { [math]::Round(($TotalSize / 1MB) / $copyDuration, 2) } else { 0 }

    $lines = [System.Collections.ArrayList]@()
    $null = $lines.Add("=====================================================")
    $null = $lines.Add("  智能复制校验报告")
    $null = $lines.Add("=====================================================")
    $null = $lines.Add("生成时间:       {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    $null = $lines.Add("源路径:         $SourcePath")
    $null = $lines.Add("目标路径:       $TargetPath")
    $null = $lines.Add("Robocopy 日志:  $($Script:RoboLogFile)")
    $null = $lines.Add("程序日志:       $($Script:LogFile)")
    $null = $lines.Add("失败日志:       $($Script:FailedLogFile)")
    $modeText = if ($DoCopy) { "复制+校验+自动修复" } else { "校验+自动修复" }
    $null = $lines.Add("操作模式:       $modeText")
    $null = $lines.Add("")
    $null = $lines.Add("--- 时间统计 ---")
    if ($Script:CopyPhaseStart) {
        $null = $lines.Add("开始时间:       {0}" -f $Script:CopyPhaseStart.ToString('yyyy-MM-dd HH:mm:ss'))
    }
    $null = $lines.Add("结束时间:       {0}" -f $endTime.ToString('yyyy-MM-dd HH:mm:ss'))
    $null = $lines.Add("总耗时:         {0} 秒" -f [math]::Round($totalDuration, 1))
    $null = $lines.Add("复制阶段耗时:   {0} 秒" -f [math]::Round($copyDuration, 1))
    $null = $lines.Add("校验阶段耗时:   {0} 秒" -f [math]::Round($verifyDuration, 1))
    $null = $lines.Add("")
    $null = $lines.Add("--- 数据统计 ---")
    $null = $lines.Add("总文件数:       $FileCount")
    $null = $lines.Add("总数据量:       {0} GB" -f [math]::Round($TotalSize / 1GB, 2))
    $null = $lines.Add("平均复制速度:   $speedMB MB/s")
    $null = $lines.Add("")
    $null = $lines.Add("--- 校验统计 ---")
    $null = $lines.Add("跳过(一致):     $Skip")
    $null = $lines.Add("修复(重复制):   $Fix")
    $null = $lines.Add("失败:           $ErrCount")
    if ($LongPathSkipped -gt 0) {
        $null = $lines.Add("长路径跳过:     $LongPathSkipped")
    }
    $null = $lines.Add("")
    $null = $lines.Add("Robocopy 返回码: $($Script:RoboExitCode)")
    if ($Script:ErrorFiles.Count -gt 0 -or $Script:LongPathSkippedFiles.Count -gt 0) {
        $null = $lines.Add("")
        $null = $lines.Add("--- 失败文件列表 ---")
        if ($Script:LongPathSkippedFiles.Count -gt 0) {
            $null = $lines.Add("")
            $null = $lines.Add("--- 长路径跳过 (未参与 Hash) ---")
            foreach ($lp in $Script:LongPathSkippedFiles) {
                $null = $lines.Add("  $lp")
            }
        }
        foreach ($ef in $Script:ErrorFiles) {
            $null = $lines.Add("  $ef")
        }
    }
    $null = $lines.Add("")
    $null = $lines.Add("=====================================================")

    $reportText = $lines -join "`r`n"
    try {
        $reportText | Out-File -FilePath $Script:ReportFile -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    catch { }
    Write-Host ""
    Write-Host $reportText -ForegroundColor Cyan
    Write-Log "报告已保存至: $($Script:ReportFile)" "INFO"
}

# ==================== Test-DiskSpace ====================
function Test-DiskSpace {
    param(
        [Parameter(Mandatory)][string]$TargetPath,
        [Parameter(Mandatory)][long]$RequiredBytes
    )
    try {
        $driveLetter = [System.IO.Path]::GetPathRoot($TargetPath)
        if ([string]::IsNullOrEmpty($driveLetter)) {
            return @{ OK = $true; FreeGB = 0; RequiredGB = 0; Warning = "无法获取盘符信息" }
        }

        # UNC 路径检测（\\server\share）
        if ($driveLetter.StartsWith("\\")) {
            # UNC 路径无法通过 WMI 获取磁盘空间，输出警告但不阻止
            return @{ OK = $true; FreeGB = 0; RequiredGB = 0; Warning = "UNC 路径无法验证磁盘空间，请手动确认目标有足够空间" }
        }

        # 本地磁盘检测
        $drive = Get-WmiObject -Class Win32_LogicalDisk -Filter ("DeviceID='{0}'" -f $driveLetter.TrimEnd('\')) -ErrorAction SilentlyContinue
        if ($null -eq $drive) {
            return @{ OK = $true; FreeGB = 0; RequiredGB = 0; Warning = "无法获取磁盘信息" }
        }
        $freeBytes = [long]$drive.FreeSpace
        $requiredGB = [math]::Round($RequiredBytes / 1GB, 2)
        $freeGB = [math]::Round($freeBytes / 1GB, 2)
        $ok = $freeBytes -ge $RequiredBytes
        return @{ OK = $ok; FreeGB = $freeGB; RequiredGB = $requiredGB; Warning = "" }
    }
    catch {
        return @{ OK = $true; FreeGB = 0; RequiredGB = 0; Warning = "磁盘空间检测异常: $($_.Exception.Message)" }
    }
}

# ==================== Test-PathValid ====================
function Test-PathValid {
    param([Parameter(Mandatory)][string]$FilePath)
    try {
        # 检查非法文件名字符（GetFileName 可能对含非法字符的路径抛出 ArgumentException）
        $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
        $fileName = $null
        try {
            $fileName = [System.IO.Path]::GetFileName($FilePath)
        }
        catch {
            return @{ Valid = $false; Reason = "路径含非法字符: $FilePath" }
        }
        foreach ($ch in $invalidChars) {
            if ($fileName.Contains($ch)) {
                return @{ Valid = $false; Reason = "文件名含非法字符 '$ch': $fileName" }
            }
        }
        # 检查保留文件名
        $reserved = @("CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9")
        $nameWithoutExt = $null
        try {
            $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        }
        catch {
            return @{ Valid = $false; Reason = "文件名解析失败: $fileName" }
        }
        if ($reserved -contains $nameWithoutExt.ToUpper()) {
            return @{ Valid = $false; Reason = "文件名是 Windows 保留名: $fileName" }
        }
        return @{ Valid = $true; Reason = "" }
    }
    catch {
        return @{ Valid = $false; Reason = "路径验证异常: $($_.Exception.Message)" }
    }
}

# ==================== Get-FileEnumerator（BFS 逐层遍历，异常目录自动跳过，支持长路径） ====================
function Get-FileEnumerator {
    param([Parameter(Mandatory)][string]$RootPath)
    $queue = [System.Collections.Queue]::new()
    $null = $queue.Enqueue($RootPath)
    while ($queue.Count -gt 0) {
        $dir = $queue.Dequeue()
        # 获取子目录（异常则跳过该子树）
        $subDirs = $null
        try {
            $subDirs = [System.IO.Directory]::GetDirectories($dir)
        }
        catch [System.IO.PathTooLongException] {
            # 长路径：尝试 Normalize 后重试
            try {
                $np = Normalize-LongPath -Path $dir
                $subDirs = [System.IO.Directory]::GetDirectories($np)
            }
            catch {
                Write-Log "跳过长路径目录（Normalize 后仍失败）: $dir" "WARN"
                $Script:LongPathSkippedCount++
                $null = $Script:LongPathSkippedFiles.Add("$dir (目录遍历失败)")
                continue
            }
        }
        catch [System.UnauthorizedAccessException] {
            Write-Log "跳过无权访问目录: $dir" "WARN"
            continue
        }
        catch [System.IO.DirectoryNotFoundException] {
            Write-Log "跳过不存在目录: $dir" "WARN"
            continue
        }
        catch {
            Write-Log "跳过异常目录: $dir - $($_.Exception.Message)" "WARN"
            continue
        }
        if ($null -ne $subDirs) {
            foreach ($d in $subDirs) { $null = $queue.Enqueue($d) }
        }
        # 获取当前目录的文件（异常则跳过该目录的文件）
        $files = $null
        try {
            $files = [System.IO.Directory]::GetFiles($dir)
        }
        catch [System.IO.PathTooLongException] {
            # 长路径：尝试 Normalize 后重试
            try {
                $np = Normalize-LongPath -Path $dir
                $files = [System.IO.Directory]::GetFiles($np)
            }
            catch {
                Write-Log "跳过长路径目录（文件 Normalize 后仍失败）: $dir" "WARN"
                $Script:LongPathSkippedCount++
                $null = $Script:LongPathSkippedFiles.Add("$dir (文件遍历失败)")
                continue
            }
        }
        catch [System.UnauthorizedAccessException] {
            Write-Log "跳过无权访问目录（文件）: $dir" "WARN"
            continue
        }
        catch [System.IO.DirectoryNotFoundException] {
            Write-Log "跳过不存在目录（文件）: $dir" "WARN"
            continue
        }
        catch {
            Write-Log "跳过异常目录（文件）: $dir - $($_.Exception.Message)" "WARN"
            continue
        }
        if ($null -ne $files) {
            foreach ($f in $files) { Write-Output $f }
        }
    }
}

# ==================== Compute-DirectorySize（真正 Streaming） ====================
function Compute-DirectorySize {
    param([Parameter(Mandatory)][string]$Path)
    $totalSize = [long]0
    $fileCount = [long]0
    $skippedDirs = [System.Collections.ArrayList]::new()
    try {
        Get-FileEnumerator -RootPath $Path | ForEach-Object {
            try {
                $np = Normalize-LongPath -Path $_
                $info = New-Object System.IO.FileInfo($np)
                $totalSize += $info.Length
                $fileCount++
            }
            catch {
                # 跳过无法访问的文件（权限/长路径等）
            }
        }
    }
    catch {
        $null = $skippedDirs.Add($Path)
        Write-Log "枚举目录失败（已跳过统计）: $Path - $($_.Exception.Message)" "WARN"
    }
    return @{ Size = $totalSize; Count = $fileCount; SkippedDirs = $skippedDirs }
}

# ==================== 主流程 ====================
$Script:TargetDriveRoot = ""
try {
    $Script:StartTime = Get-Date
    Write-Log "程序启动" "INFO"

    # ---------- 1. 选择源文件夹 ----------
    $SourcePath = Select-Source -Description "请选择【源】文件夹（要复制的数据所在）"
    if ([string]::IsNullOrEmpty($SourcePath)) {
        Write-Host "未选择源文件夹，退出。" -ForegroundColor Red
        Read-Host "按任意键退出"
        exit
    }

    # ---------- 2. 选择目标文件夹 ----------
    $TargetPath = Select-Folder -Description "请选择【目标】文件夹（数据要复制到的位置）" -ShowNewFolder $true
    if ([string]::IsNullOrEmpty($TargetPath)) {
        Write-Host "未选择目标文件夹，退出。" -ForegroundColor Red
        Read-Host "按任意键退出"
        exit
    }

    # 路径规范化
    # FIXED: 同时去除末尾空格和反斜杠，避免路径含空格导致 robocopy 参数异常
    $SourcePath = $SourcePath.TrimEnd('\', ' ', "`t")
    $TargetPath = $TargetPath.TrimEnd('\', ' ', "`t")

    # 修复盘符路径：F: → F:\（Windows 中 F: 指当前目录，F:\ 才是根目录）
    if ($SourcePath -match '^[A-Za-z]:$') { $SourcePath = $SourcePath + '\' }

    # 保留源文件夹结构：目标 = 目标\源文件夹名
    $SourceFolderName = Split-Path $SourcePath -Leaf
    $TargetPath = Join-Path $TargetPath $SourceFolderName

    # 提取目标根路径（用于掉盘检测）
    # 支持: C:\  D:\  \\Server\Share  USB  映射盘
    $Script:TargetDriveRoot = [System.IO.Path]::GetPathRoot($TargetPath)

    Write-Host "源路径:   $SourcePath" -ForegroundColor Cyan
    Write-Host "目标路径: $TargetPath" -ForegroundColor Cyan
    Write-Log "源路径: $SourcePath" "INFO"
    Write-Log "目标路径: $TargetPath" "INFO"

    if ($SourcePath -eq $TargetPath) {
        Write-Host "错误：源和目标不能相同！" -ForegroundColor Red
        Read-Host "按任意键退出"
        exit
    }

    # 检查源路径是否存在
    if (-not (Test-Path -LiteralPath $SourcePath -ErrorAction SilentlyContinue)) {
        Write-Host "错误：源路径不存在！" -ForegroundColor Red
        Read-Host "按任意键退出"
        exit
    }

    # 检查目标路径是否存在
    if (-not (Test-Path -LiteralPath $TargetPath -ErrorAction SilentlyContinue)) {
        try {
            New-Item -Path $TargetPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Host "错误：无法创建目标路径！" -ForegroundColor Red
            Read-Host "按任意键退出"
            exit
        }
    }

    # 检查嵌套路径（防止无限递归复制）
    $srcLower = if ($SourcePath.EndsWith('\')) { $SourcePath.ToLower() } else { $SourcePath.ToLower() + '\' }
    $tgtLower = if ($TargetPath.EndsWith('\')) { $TargetPath.ToLower() } else { $TargetPath.ToLower() + '\' }
    if ($srcLower.StartsWith($tgtLower) -or $tgtLower.StartsWith($srcLower)) {
        Write-Host "错误：源和目标不能存在嵌套关系（会导致无限递归）！" -ForegroundColor Red
        Write-Host "  源: $SourcePath" -ForegroundColor Yellow
        Write-Host "  目标: $TargetPath" -ForegroundColor Yellow
        Read-Host "按任意键退出"
        exit
    }

    # ---------- 3. 操作模式（始终：复制 + 完整校验 + 自动修复）----------
    $doCopy = $true
    Write-Log "操作模式: 复制+校验+自动修复" "INFO"

    # ---------- 4. 磁盘空间检查（仅复制模式） ----------
    if ($doCopy) {
        Write-Log "计算源数据总大小..." "INFO"
        $sizeCheck = Compute-DirectorySize -Path $SourcePath
        $diskCheck = Test-DiskSpace -TargetPath $TargetPath -RequiredBytes $sizeCheck.Size
        if (-not $diskCheck.OK) {
            Write-Host "`n警告：目标磁盘空间不足！" -ForegroundColor Yellow
            Write-Host "  源数据大小: $($diskCheck.RequiredGB) GB" -ForegroundColor Yellow
            Write-Host "  目标可用:   $($diskCheck.FreeGB) GB" -ForegroundColor Yellow
            Write-Host "  差额:       $([math]::Round($diskCheck.RequiredGB - $diskCheck.FreeGB, 2)) GB" -ForegroundColor Yellow
            Write-Log "磁盘空间不足: 需要 $($diskCheck.RequiredGB) GB, 可用 $($diskCheck.FreeGB) GB" "WARN"
            $answer = Read-Host "空间不足，是否仍要继续复制+校验？(Y/N)"
            if ($answer -eq 'Y' -or $answer -eq 'y') {
                Write-Log "用户选择继续执行（空间不足）" "INFO"
            }
            else {
                Read-Host "按任意键退出"
                exit
            }
        }
        if ($diskCheck.Warning -ne "") {
            Write-Host "`n警告：$($diskCheck.Warning)" -ForegroundColor Yellow
            Write-Log $diskCheck.Warning "WARN"
        }
        Write-Log "磁盘空间充足: 需要 $($diskCheck.RequiredGB) GB, 可用 $($diskCheck.FreeGB) GB" "INFO"
    }

    # ---------- 5. 复制阶段（robocopy） ----------
    if ($doCopy) {
        Write-Host "`n===== 开始复制阶段（robocopy）=====" -ForegroundColor Yellow
        Write-Log "===== 开始复制阶段 =====" "INFO"
        $Script:CopyPhaseStart = Get-Date

        Write-Host "正在统计源文件总数..." -ForegroundColor Gray -NoNewline
        $Script:VerifyTotalFiles = 0
        Get-FileEnumerator -RootPath $SourcePath | ForEach-Object { $Script:VerifyTotalFiles++ }
        Write-Host "`r源文件总数: $($Script:VerifyTotalFiles) 个                    " -ForegroundColor Cyan

        Write-Host "正在复制文件到: $TargetPath" -ForegroundColor Gray

        $Script:RoboExitCode = Start-Robocopy -Source $SourcePath -Target $TargetPath -LogFile $Script:RoboLogFile -TotalFiles $Script:VerifyTotalFiles

        $Script:CopyPhaseEnd = Get-Date
        $copyDuration = ($Script:CopyPhaseEnd - $Script:CopyPhaseStart).TotalSeconds

        Write-Host ""

        # robocopy 返回码语义检查
        if ($Script:RoboExitCode -ge 8) {
            Write-Host "复制中出现错误（返回码 $($Script:RoboExitCode)），请查看日志。" -ForegroundColor Yellow
            Write-Log "Robocopy 返回码 $($Script:RoboExitCode) (>=8 表示有错误)" "WARN"
        }
        elseif ($Script:RoboExitCode -eq 0) {
            Write-Host "复制阶段完成（无需操作，目标已是最新，耗时 $([math]::Round($copyDuration,1)) 秒）。" -ForegroundColor Green
            Write-Log "Robocopy 返回码 0（无需操作），耗时 $([math]::Round($copyDuration,1)) 秒" "INFO"
        }
        else {
            Write-Host "复制阶段完成（返回码 $($Script:RoboExitCode)，耗时 $([math]::Round($copyDuration,1)) 秒）。" -ForegroundColor Green
            Write-Log "Robocopy 返回码 $($Script:RoboExitCode)，耗时 $([math]::Round($copyDuration,1)) 秒" "INFO"
        }

        # 等待磁盘缓存刷新
        Write-Log "等待 3 秒使 Windows 写缓存刷新..." "INFO"
        Start-Sleep -Seconds 3
    }
    else {
        # dead code path, kept for safety
    }

    # ---------- 6. 校验阶段（多线程并行 SHA256 + 顺序修复） ----------
    Write-Host "`n===== 开始校验阶段（多线程 SHA256 比对 + 自动修复）=====" -ForegroundColor Yellow
    Write-Log "===== 开始校验阶段 =====" "INFO"
    $Script:VerifyStart = Get-Date

    if ($Script:VerifyTotalFiles -eq 0) {
        Write-Host "正在统计源文件总数..." -ForegroundColor Gray -NoNewline
        Get-FileEnumerator -RootPath $SourcePath | ForEach-Object { $Script:VerifyTotalFiles++ }
        Write-Host "`r源文件总数: $($Script:VerifyTotalFiles) 个                    " -ForegroundColor Cyan
    }

    $Script:TotalCount    = 0
    $Script:SkipCount     = 0
    $Script:FixCount      = 0
    $Script:ErrorCount    = 0
    $Script:LongPathSkippedCount = 0
    $Script:ErrorFiles    = [System.Collections.ArrayList]::new()
    $Script:LongPathSkippedFiles = [System.Collections.ArrayList]::new()

    Write-Host "正在收集文件列表..." -ForegroundColor Gray -NoNewline
    $allFiles = [System.Collections.ArrayList]::new()
    $fileCount = [int]0
    Get-FileEnumerator -RootPath $SourcePath | ForEach-Object {
        $null = $allFiles.Add($_)
        $fileCount++
        if ($fileCount % 100 -eq 0) {
            Write-Host "`r正在收集文件列表... $fileCount 个" -NoNewline -ForegroundColor Gray
        }
    }
    Write-Host "`r正在收集文件列表... $fileCount 个                    " -ForegroundColor Cyan
    $totalFiles = $allFiles.Count
    Write-Host "共 $totalFiles 个文件，开始多线程 SHA256 校验..." -ForegroundColor Cyan

    # --- 阶段一：多线程并行计算所有 SHA256（Runspace Pool） ---
    $hashResults = [System.Collections.ArrayList]::new()
    $cpuCores = [Environment]::ProcessorCount
    $maxRunspaces = [math]::Min($cpuCores, 16)
    Write-Host "使用 $maxRunspaces 个线程并行计算 SHA256（CPU 核心: $cpuCores）" -ForegroundColor Cyan

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Runspace Pool：独立会话，避免 PS5.1 委托兼容问题
    $runspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $maxRunspaces)
    $runspacePool.Open()

    $hashScriptBlock = {
        param($FilePath, $SrcPath, $TgtPath)
        Add-Type -AssemblyName System.IO.FileSystem
        try {
            $np = [System.IO.Path]::Combine("\\?\", $FilePath.TrimStart('\').TrimEnd('\'))
            if ($np -notmatch '^\\\\\?\\') { $np = "\\?\$FilePath" }
            $fi = New-Object System.IO.FileInfo($np)
            $srcSize = $fi.Length
        }
        catch {
            return [pscustomobject]@{ FilePath=$FilePath; RelPath=""; SrcHash=$null; TgtHash=$null; SrcSize=0; Status="FileInfoError"; ErrorMsg=$_.Exception.Message }
        }
        if ($FilePath.Length -le $SrcPath.Length) {
            return [pscustomobject]@{ FilePath=$FilePath; RelPath=""; SrcHash=$null; TgtHash=$null; SrcSize=0; Status="Skip"; ErrorMsg="" }
        }
        $relPath = $FilePath.Substring($SrcPath.Length + 1)
        $targetFile = Join-Path $TgtPath $relPath

        try {
            $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
            $fileName = [System.IO.Path]::GetFileName($targetFile)
            foreach ($ch in $invalidChars) { if ($fileName.Contains($ch)) { throw "invalid char" } }
            $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
            $reserved = @("CON","PRN","AUX","NUL","COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9","LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9")
            if ($reserved -contains $nameWithoutExt.ToUpper()) { throw "reserved name" }
        }
        catch {
            return [pscustomobject]@{ FilePath=$FilePath; RelPath=$relPath; SrcHash=$null; TgtHash=$null; SrcSize=$srcSize; Status="PathInvalid"; ErrorMsg=$_.Exception.Message }
        }

        $targetExists = $false
        try { $targetExists = [System.IO.File]::Exists($targetFile) } catch {}
        if (-not $targetExists) {
            return [pscustomobject]@{ FilePath=$FilePath; RelPath=$relPath; SrcHash=$null; TgtHash=$null; SrcSize=$srcSize; Status="TargetMissing"; ErrorMsg="" }
        }

        try {
            $tnp = [System.IO.Path]::Combine("\\?\", $targetFile.TrimStart('\').TrimEnd('\'))
            if ($tnp -notmatch '^\\\\\?\\') { $tnp = "\\?\$targetFile" }
            $fs = [System.IO.File]::Open($tnp, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $fs.Dispose()
        }
        catch {
            return [pscustomobject]@{ FilePath=$FilePath; RelPath=$relPath; SrcHash=$null; TgtHash=$null; SrcSize=$srcSize; Status="TargetUnreadable"; ErrorMsg="" }
        }

        try {
            $tnp2 = [System.IO.Path]::Combine("\\?\", $targetFile.TrimStart('\').TrimEnd('\'))
            if ($tnp2 -notmatch '^\\\\\?\\') { $tnp2 = "\\?\$targetFile" }
            $tgtSize = (New-Object System.IO.FileInfo($tnp2)).Length
        }
        catch { $tgtSize = -1 }
        if ($tgtSize -ge 0 -and $srcSize -ne $tgtSize) {
            return [pscustomobject]@{ FilePath=$FilePath; RelPath=$relPath; SrcHash=$null; TgtHash=$null; SrcSize=$srcSize; TgtSize=$tgtSize; Status="SizeMismatch"; ErrorMsg="" }
        }

        # 计算 SHA256（源 + 目标）
        try {
            $sha = [System.Security.Cryptography.SHA256]::Create()
            $buffer = [byte[]]::new(131072)

            $np = "\\?\$FilePath"
            $stream = [System.IO.File]::Open($np, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            try {
                do { $read = $stream.Read($buffer, 0, $buffer.Length); if ($read -gt 0) { $sha.TransformBlock($buffer, 0, $read, $null, 0) | Out-Null } } while ($read -gt 0)
                $sha.TransformFinalBlock([byte[]]::new(0), 0, 0) | Out-Null
                $srcHash = ([System.BitConverter]::ToString($sha.Hash) -replace "-", "")
            }
            finally { $stream.Dispose(); $sha.Dispose() }

            $sha2 = [System.Security.Cryptography.SHA256]::Create()
            $np2 = "\\?\$targetFile"
            $stream2 = [System.IO.File]::Open($np2, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            try {
                do { $read = $stream2.Read($buffer, 0, $buffer.Length); if ($read -gt 0) { $sha2.TransformBlock($buffer, 0, $read, $null, 0) | Out-Null } } while ($read -gt 0)
                $sha2.TransformFinalBlock([byte[]]::new(0), 0, 0) | Out-Null
                $tgtHash = ([System.BitConverter]::ToString($sha2.Hash) -replace "-", "")
            }
            finally { $stream2.Dispose(); $sha2.Dispose() }
        }
        catch {
            return [pscustomobject]@{ FilePath=$FilePath; RelPath=$relPath; SrcHash=$null; TgtHash=$null; SrcSize=$srcSize; Status="HashError"; ErrorMsg=$_.Exception.Message }
        }

        $status = if ($srcHash -eq $tgtHash) { "Match" } else { "HashMismatch" }
        return [pscustomobject]@{ FilePath=$FilePath; RelPath=$relPath; SrcHash=$srcHash; TgtHash=$tgtHash; SrcSize=$srcSize; Status=$status; ErrorMsg="" }
    }

    $runspaces = [System.Collections.ArrayList]::new()
    $processedCount = [int]0

    foreach ($filePath in $allFiles) {
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.RunspacePool = $runspacePool
        $ps.AddScript($hashScriptBlock).AddArgument($filePath).AddArgument($SourcePath).AddArgument($TargetPath) | Out-Null
        $handle = $ps.BeginInvoke()
        $null = $runspaces.Add(@{ PS = $ps; Handle = $handle })
    }

    # 等待所有 runspace 完成（显示进度）
    $collectedCount = [int]0
    $shaStartTime = Get-Date
    foreach ($rs in $runspaces) {
        $result = $rs.PS.EndInvoke($rs.Handle)
        if ($null -ne $result) {
            foreach ($item in $result) {
                if ($item.Status -ne "Skip") {
                    $null = $hashResults.Add($item)
                }
            }
        }
        $rs.PS.Dispose()
        $collectedCount++
        if ($collectedCount % 20 -eq 0 -or $collectedCount -eq $runspaces.Count) {
            Write-ConsoleProgress -Activity "SHA256 校验中" -Current $collectedCount -Total $runspaces.Count -StartTime $shaStartTime
        }
    }
    $runspacePool.Close()
    $runspacePool.Dispose()

    $stopwatch.Stop()
    Write-Host "`rSHA256 计算完成: $totalFiles 个文件，耗时 $([math]::Round($stopwatch.Elapsed.TotalSeconds, 1)) 秒                      " -ForegroundColor Green

    # --- 阶段二：顺序处理修复（robocopy 不宜过多并发） ---
    if ($doCopy) {
        Write-Host "开始处理修复..." -ForegroundColor Gray
    }
    else {
        Write-Host "跳过复制模式，仅记录不一致文件（不执行修复）..." -ForegroundColor Yellow
    }
    $repairCount = 0
    $repairTotal = if ($doCopy) { ($hashResults | Where-Object { $_.Status -ne "Match" }).Count } else { 0 }
    $repairStartTime = Get-Date

    foreach ($r in $hashResults) {
        $Script:TotalCount++
        $filePath = $r.FilePath
        $relPath  = $r.RelPath

        switch ($r.Status) {
            "FileInfoError" {
                $Script:ErrorCount++
                $null = $Script:ErrorFiles.Add("$filePath  |  无法获取文件信息  |  $($r.ErrorMsg)")
            }
            "PathInvalid" {
                $Script:ErrorCount++
                $null = $Script:ErrorFiles.Add("$relPath  |  路径无效  |  $($r.ErrorMsg)")
                Write-Log "路径无效，跳过: $($r.ErrorMsg)" "WARN"
            }
            "TargetMissing" {
                if ($doCopy) {
                    $repairCount++
                    $targetFile = Join-Path $TargetPath $relPath
                    if (-not (Test-HasEnoughSpace -TargetPath $TargetPath -RequiredBytes $r.SrcSize)) {
                        $Script:ErrorCount++
                        $null = $Script:ErrorFiles.Add("$relPath  |  目标缺失  |  跳过修复（磁盘空间不足，需要 $([math]::Round($r.SrcSize/1MB,1)) MB）")
                        Write-Log "磁盘空间不足，跳过: $relPath (需要 $([math]::Round($r.SrcSize/1MB,1)) MB)" "WARN"
                        break
                    }
                    Write-Log "目标不存在，复制: $relPath" "INFO"
                    Write-ConsoleProgress -Activity "修复中" -Current $repairCount -Total $repairTotal -Status $relPath -StartTime $repairStartTime
                    $repairResult = Invoke-WithRetry -ScriptBlock {
                        Repair-File -SourceFile $filePath -TargetFile $targetFile
                    } -ErrorMessage "复制失败: $relPath"
                    if ($repairResult.Success -and $repairResult.Result -eq $true) {
                        try { $srcHash = Get-SHA256 -FilePath $filePath; $tgtHash = Get-SHA256 -FilePath $targetFile }
                        catch { $Script:ErrorCount++; $null = $Script:ErrorFiles.Add("$relPath  |  SHA256计算失败  |  $($_.Exception.Message)"); continue }
                        if ($srcHash -eq $tgtHash) { $Script:FixCount++ }
                        else { $Script:ErrorCount++; $null = $Script:ErrorFiles.Add("$relPath  |  SHA256不一致  |  源=$srcHash 目标=$tgtHash  |  复制后校验失败") }
                    }
                    else { $Script:ErrorCount++; $null = $Script:ErrorFiles.Add("$relPath  |  复制失败  |  $($repairResult.Error)") }
                }
                else {
                    $Script:ErrorCount++
                    $null = $Script:ErrorFiles.Add("$relPath  |  目标缺失  |  跳过修复（空间不足）")
                    Write-Log "目标不存在（跳过修复）: $relPath" "WARN"
                }
            }
            "TargetUnreadable" {
                if ($doCopy) {
                    $repairCount++
                    $targetFile = Join-Path $TargetPath $relPath
                    if (-not (Test-HasEnoughSpace -TargetPath $TargetPath -RequiredBytes $r.SrcSize)) {
                        $Script:ErrorCount++
                        $null = $Script:ErrorFiles.Add("$relPath  |  目标不可读  |  跳过修复（磁盘空间不足，需要 $([math]::Round($r.SrcSize/1MB,1)) MB）")
                        Write-Log "磁盘空间不足，跳过: $relPath (需要 $([math]::Round($r.SrcSize/1MB,1)) MB)" "WARN"
                        break
                    }
                    Write-Log "文件不可读，清理后重新复制: $relPath" "INFO"
                    Write-ConsoleProgress -Activity "修复中" -Current $repairCount -Total $repairTotal -Status $relPath -StartTime $repairStartTime
                    $repairResult = Invoke-WithRetry -ScriptBlock {
                        Remove-CorruptedTarget -FilePath $targetFile
                        Repair-File -SourceFile $filePath -TargetFile $targetFile
                    } -ErrorMessage "重新复制失败: $relPath"
                    if ($repairResult.Success -and $repairResult.Result -eq $true) {
                        try { $srcHash = Get-SHA256 -FilePath $filePath; $tgtHash = Get-SHA256 -FilePath $targetFile }
                        catch { $Script:ErrorCount++; $null = $Script:ErrorFiles.Add("$relPath  |  SHA256计算失败  |  $($_.Exception.Message)"); continue }
                        if ($srcHash -eq $tgtHash) { $Script:FixCount++ }
                        else { $Script:ErrorCount++; $null = $Script:ErrorFiles.Add("$relPath  |  SHA256不一致  |  源=$srcHash 目标=$tgtHash  |  不可读重新复制后校验失败") }
                    }
                    else { $Script:ErrorCount++; $null = $Script:ErrorFiles.Add("$relPath  |  重新复制失败  |  $($repairResult.Error)") }
                }
                else {
                    $Script:ErrorCount++
                    $null = $Script:ErrorFiles.Add("$relPath  |  目标不可读  |  跳过修复（空间不足）")
                    Write-Log "目标不可读（跳过修复）: $relPath" "WARN"
                }
            }
            "SizeMismatch" {
                if ($doCopy) {
                    $repairCount++
                    $targetFile = Join-Path $TargetPath $relPath
                    if (-not (Test-HasEnoughSpace -TargetPath $TargetPath -RequiredBytes $r.SrcSize)) {
                        $Script:ErrorCount++
                        $null = $Script:ErrorFiles.Add("$relPath  |  大小不同 ($($r.SrcSize) vs $($r.TgtSize))  |  跳过修复（磁盘空间不足，需要 $([math]::Round($r.SrcSize/1MB,1)) MB）")
                        Write-Log "磁盘空间不足，跳过: $relPath (需要 $([math]::Round($r.SrcSize/1MB,1)) MB)" "WARN"
                        break
                    }
                    Write-Log "文件大小不同 ($($r.SrcSize) vs $($r.TgtSize))，重新复制: $relPath" "WARN"
                    Write-ConsoleProgress -Activity "修复中" -Current $repairCount -Total $repairTotal -Status $relPath -StartTime $repairStartTime
                    $repairResult = Invoke-WithRetry -ScriptBlock {
                        Repair-File -SourceFile $filePath -TargetFile $targetFile -CleanupFirst $true
                    } -ErrorMessage "重新复制失败: $relPath"
                    if ($repairResult.Success -and $repairResult.Result -eq $true) {
                        try { $srcHash = Get-SHA256 -FilePath $filePath; $tgtHash = Get-SHA256 -FilePath $targetFile }
                        catch { $Script:ErrorCount++; $null = $Script:ErrorFiles.Add("$relPath  |  SHA256计算失败  |  $($_.Exception.Message)"); continue }
                        if ($srcHash -eq $tgtHash) { $Script:FixCount++ }
                        else { $Script:ErrorCount++; $null = $Script:ErrorFiles.Add("$relPath  |  SHA256不一致  |  源=$srcHash 目标=$tgtHash  |  大小不同重新复制后校验失败") }
                    }
                    else { $Script:ErrorCount++; $null = $Script:ErrorFiles.Add("$relPath  |  重新复制失败  |  $($repairResult.Error)") }
                }
                else {
                    $Script:ErrorCount++
                    $null = $Script:ErrorFiles.Add("$relPath  |  大小不同 ($($r.SrcSize) vs $($r.TgtSize))  |  跳过修复（空间不足）")
                    Write-Log "大小不同（跳过修复）: $relPath  源=$($r.SrcSize) 目标=$($r.TgtSize)" "WARN"
                }
            }
            "HashError" {
                $Script:ErrorCount++
                $null = $Script:ErrorFiles.Add("$relPath  |  SHA256计算失败  |  $($r.ErrorMsg)")
                Write-Log "SHA256 计算失败: $relPath - $($r.ErrorMsg)" "ERROR"
            }
            "HashMismatch" {
                if ($doCopy) {
                    $repairCount++
                    $targetFile = Join-Path $TargetPath $relPath
                    Write-Log "SHA256 不一致，重新复制: $relPath  源=$($r.SrcHash) 目标=$($r.TgtHash)" "WARN"
                    Write-ConsoleProgress -Activity "修复中" -Current $repairCount -Total $repairTotal -Status $relPath -StartTime $repairStartTime
                    $repairResult = Invoke-WithRetry -ScriptBlock {
                        Repair-File -SourceFile $filePath -TargetFile $targetFile -CleanupFirst $true
                    } -ErrorMessage "重新复制失败: $relPath"
                    if ($repairResult.Success -and $repairResult.Result -eq $true) {
                        try { $srcHash2 = Get-SHA256 -FilePath $filePath; $tgtHash2 = Get-SHA256 -FilePath $targetFile }
                        catch { $Script:ErrorCount++; $null = $Script:ErrorFiles.Add("$relPath  |  最终SHA256计算失败  |  $($_.Exception.Message)"); continue }
                        if ($srcHash2 -eq $tgtHash2) { $Script:FixCount++ }
                        else { $Script:ErrorCount++; $null = $Script:ErrorFiles.Add("$relPath  |  SHA256不一致  |  源=$srcHash2 目标=$tgtHash2  |  重新复制后校验失败") }
                    }
                    else { $Script:ErrorCount++; $null = $Script:ErrorFiles.Add("$relPath  |  重新复制失败  |  $($repairResult.Error)") }
                }
                else {
                    $Script:ErrorCount++
                    $null = $Script:ErrorFiles.Add("$relPath  |  SHA256不一致  |  源=$($r.SrcHash) 目标=$($r.TgtHash)  |  跳过修复（空间不足）")
                    Write-Log "SHA256 不一致（跳过修复）: $relPath  源=$($r.SrcHash) 目标=$($r.TgtHash)" "WARN"
                }
            }
            "Match" {
                $Script:SkipCount++
            }
        }
    }

    # 统计一致性检查
    $expectedTotal = $Script:SkipCount + $Script:FixCount + $Script:ErrorCount + $Script:LongPathSkippedCount
    if ($expectedTotal -ne $Script:TotalCount) {
        $mismatchMsg = "Statistics mismatch: Total={0}, Skip={1}, Fix={2}, Error={3}, LongPathSkipped={4}, Sum={5}" -f $Script:TotalCount, $Script:SkipCount, $Script:FixCount, $Script:ErrorCount, $Script:LongPathSkippedCount, $expectedTotal
        Write-Log $mismatchMsg "WARN"
        Write-Host "警告：统计数量不一致，请查看日志" -ForegroundColor Yellow
    }

    # 刷新剩余日志
    Flush-LogBuffer

    $Script:VerifyEnd = Get-Date
    $verifyDuration = ($Script:VerifyEnd - $Script:VerifyStart).TotalSeconds
    Write-Log "校验阶段完成，耗时 $([math]::Round($verifyDuration,1)) 秒" "INFO"
    Write-Host "`n校验阶段完成：" -ForegroundColor Cyan
    Write-Host "  总文件:       $($Script:TotalCount)" -ForegroundColor White
    Write-Host "  跳过(一致):   $($Script:SkipCount)" -ForegroundColor Green
    Write-Host "  修复(重复制): $($Script:FixCount)" -ForegroundColor Yellow
    Write-Host "  失败:         $($Script:ErrorCount)" -ForegroundColor Red
    if ($Script:LongPathSkippedCount -gt 0) {
        Write-Host "  长路径跳过:   $($Script:LongPathSkippedCount)" -ForegroundColor Magenta
    }

    Write-Host ""

    # ---------- 7. 生成报告 ----------
    # 使用 Compute-DirectorySize（真正 Streaming）
    $sizeResult = Compute-DirectorySize -Path $SourcePath
    $Script:TotalSize = $sizeResult.Size

    Flush-LogBuffer

    Write-Report -SourcePath $SourcePath -TargetPath $TargetPath -DoCopy $doCopy `
        -Total $Script:TotalCount -Skip $Script:SkipCount -Fix $Script:FixCount `
        -ErrCount $Script:ErrorCount -LongPathSkipped $Script:LongPathSkippedCount `
        -TotalSize $sizeResult.Size -FileCount $sizeResult.Count

    # 写入失败日志
    if ($Script:ErrorFiles.Count -gt 0 -or $Script:LongPathSkippedFiles.Count -gt 0) {
        $failContent = [System.Collections.ArrayList]@()
        $null = $failContent.Add("失败文件列表 - {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
        $null = $failContent.Add("=" * 60)
        if ($Script:LongPathSkippedFiles.Count -gt 0) {
            $null = $failContent.Add("")
            $null = $failContent.Add("--- 长路径跳过 (未参与 Hash) ---")
            foreach ($lp in $Script:LongPathSkippedFiles) {
                $null = $failContent.Add("  $lp")
            }
        }
        foreach ($ef in $Script:ErrorFiles) {
            $null = $failContent.Add("  $ef")
        }
        try {
            ($failContent -join "`r`n") | Out-File -FilePath $Script:FailedLogFile -Encoding UTF8 -ErrorAction SilentlyContinue
        }
        catch { }
        Write-Log "失败日志已保存至: $($Script:FailedLogFile)" "INFO"
    }

    # 最终结果
    if ($Script:ErrorCount -gt 0) {
        Write-Host "`n有 $($Script:ErrorCount) 个文件失败，请检查硬盘连接和散热后重新运行。" -ForegroundColor Red
    }
    else {
        Write-Host "`n所有文件校验通过，数据完整！" -ForegroundColor Green
    }

    # 自动打开报告
    if ($Script:ReportFile -and (Test-Path -LiteralPath $Script:ReportFile -ErrorAction SilentlyContinue)) {
        Write-Host "正在打开报告..." -ForegroundColor Gray
        Start-Process -FilePath $Script:ReportFile -ErrorAction SilentlyContinue
    }

    Write-Log "程序结束" "INFO"
    Flush-LogBuffer
}
catch {
    Write-Log "程序异常终止: $($_.Exception.Message)" "ERROR"
    Write-Log "堆栈: $($_.ScriptStackTrace)" "ERROR"
    Write-Host "程序异常终止: $($_.Exception.Message)" -ForegroundColor Red
    Flush-LogBuffer
}
finally {
    Flush-LogBuffer
    Write-Host "`n按任意键退出..." -ForegroundColor Gray
    Read-Host
}
