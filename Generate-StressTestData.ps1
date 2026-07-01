# =====================================================
# 黑盒压测数据生成脚本 (Generate-StressTestData.ps1)
# 用途：生成极端边缘场景的测试数据，验证《图形复制校验工具》的健壮性
# 兼容：PowerShell 5.1 / 无需管理员权限
# =====================================================

param(
    [string]$TestRoot = ""
)

# 若未指定目录，默认在 TEMP 下创建
if ([string]::IsNullOrEmpty($TestRoot)) {
    $TestRoot = Join-Path $env:TEMP ("StressTest_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  黑盒压测数据生成器" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "目标目录: $TestRoot" -ForegroundColor Yellow

# 创建根目录
if (-not (Test-Path -LiteralPath $TestRoot)) {
    New-Item -Path $TestRoot -ItemType Directory -Force | Out-Null
}

# =====================================================
# 01_特殊字符与Emoji
# =====================================================
Write-Host ""
Write-Host "[1/4] 生成特殊字符与 Emoji 文件..." -ForegroundColor Green

$dir01 = Join-Path $TestRoot "01_特殊字符与Emoji"
New-Item -Path $dir01 -ItemType Directory -Force | Out-Null

$specialNames = @(
    "`u{1F4C1}test.txt",                  # 文件夹 Emoji
    "`u{1F680}data.csv",                   # 火箭 Emoji
    "`u{1F4A8}log.txt",                    # 喷气 Emoji
    "`u{2764}heart.bin",                   # 爱心 Emoji
    "`u{1F30D}world.dat",                  # 地球 Emoji
    "`u{1F48E}crystal.raw",                # 宝石 Emoji
    "тест_кириллица.txt",                   # 西里尔字母
    "тест_данные.csv",                      # 西里尔字母
    "ทดสอบ_ข้อมูล.txt",                    # 泰文
    "ข้อมูล_ทดสอบ.dat",                    # 泰文
    "%test%.txt",                           # 百分号
    "$data$.csv",                           # 美元符号
    "(spec).txt",                           # 圆括号
    "[tags].dat",                           # 方括号
    "#hash.txt",                            # 井号
    "at@sign.txt",                          # @ 符号
    "tilde~file.txt",                       # 波浪号
    "ampersand&file.txt",                   # & 符号
    "space file.txt",                       # 空格
    "tab	file.txt",                        # 制表符
    "plus+minus-.txt",                      # 加减号
    "equal=sign.txt",                       # 等号
    "curly{braces}.txt",                    # 花括号
    "angle<bracket.txt",                    # 尖括号
    "pipe|symbol.txt",                      # 竖线
    "semicolon;colon.txt",                  # 分号冒号
    "quote'double.txt",                     # 引号
    "comma,period.txt"                      # 逗号句号
)

foreach ($name in $specialNames) {
    try {
        $filePath = Join-Path $dir01 $name
        # 写入随机内容（10-100 字节）
        $rng = [System.Security.Cryptography.RNGCryptoServiceProvider]::new()
        $buf = [byte[]]::new((Get-Random -Minimum 10 -Maximum 101))
        $rng.GetBytes($buf)
        $rng.Dispose()
        [System.IO.File]::WriteAllBytes($filePath, $buf)
        Write-Host "  OK: $name" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "  SKIP: $name ($($_.Exception.Message))" -ForegroundColor DarkYellow
    }
}

# =====================================================
# 02_极端长路径
# =====================================================
Write-Host ""
Write-Host "[2/4] 生成极端长路径（280-320 字符）..." -ForegroundColor Green

$dir02 = Join-Path $TestRoot "02_极端长路径"
New-Item -Path $dir02 -ItemType Directory -Force | Out-Null

# 递归创建 15 层子目录，每层名字足够长
$currentDir = $dir02
for ($i = 1; $i -le 15; $i++) {
    # 每层目录名 12-16 字符，累积使总路径超过 260
    $segment = "level_{0}_{1}" -f $i.ToString("00"), ("x" * 10)
    $currentDir = Join-Path $currentDir $segment
    try {
        New-Item -Path $currentDir -ItemType Directory -Force | Out-Null
    }
    catch {
        Write-Host "  无法创建目录 $currentDir : $($_.Exception.Message)" -ForegroundColor DarkYellow
        break
    }
}

# 在最深层创建文件
$longPathFileNames = @(
    "deep_file_A.txt",
    "deep_file_B_with_long_name_to_increase_path_length_even_further.txt",
    "deep_file_C_final.dat"
)

foreach ($fname in $longPathFileNames) {
    try {
        $fpath = Join-Path $currentDir $fname
        $len = $fpath.Length
        $content = "Long path test file - path length: $len chars"
        [System.IO.File]::WriteAllText($fpath, $content, [System.Text.Encoding]::UTF8)
        Write-Host "  OK: ...$([System.IO.Path]::Combine('...', [System.IO.Path]::GetFileName($currentDir), $fname))  (总长 $len 字符)" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "  SKIP: $fname ($($_.Exception.Message))" -ForegroundColor DarkYellow
    }
}

# 再创建一个刚好 260 字符的文件（边界测试）
try {
    $boundaryDir = Join-Path $dir02 ("boundary_" + ("b" * 30))
    New-Item -Path $boundaryDir -ItemType Directory -Force | Out-Null
    # FIXED: 简化边界文件创建，避免复杂的 padding 计算错误
    # 创建一个路径长度接近 260 字符的文件
    $boundaryFile = Join-Path $boundaryDir "exactly_at_boundary.txt"
    $currentLen = $boundaryFile.Length
    if ($currentLen -lt 255) {
        # 通过子目录名填充到接近 260
        $padLen = 255 - $currentLen
        if ($padLen -gt 0) {
            $padding = "p" * $padLen
            $boundaryDir2 = Join-Path $boundaryDir $padding
            New-Item -Path $boundaryDir2 -ItemType Directory -Force | Out-Null
            $boundaryFile = Join-Path $boundaryDir2 "boundary.txt"
        }
    }
    [System.IO.File]::WriteAllText($boundaryFile, "boundary test", [System.Text.Encoding]::UTF8)
    Write-Host "  OK: boundary file (总长 $($boundaryFile.Length) 字符)" -ForegroundColor DarkGray
}
catch {
    Write-Host "  SKIP: boundary file ($($_.Exception.Message))" -ForegroundColor DarkYellow
}

# =====================================================
# 03_零字节与大文件混合
# =====================================================
Write-Host ""
Write-Host "[3/4] 生成零字节与大文件混合..." -ForegroundColor Green

$dir03 = Join-Path $TestRoot "03_零字节与大文件混合"
New-Item -Path $dir03 -ItemType Directory -Force | Out-Null

# 20 个零字节文件
for ($i = 1; $i -le 20; $i++) {
    $fpath = Join-Path $dir03 ("empty_{0}.txt" -f $i.ToString("00"))
    [System.IO.File]::WriteAllBytes($fpath, [byte[]]::new(0))
}
Write-Host "  OK: 20 个零字节文件" -ForegroundColor DarkGray

# 5 个有随机内容的正常文件（1KB - 100KB）
for ($i = 1; $i -le 5; $i++) {
    $sizeKB = Get-Random -Minimum 1 -Maximum 101
    $fpath = Join-Path $dir03 ("data_{0}_{1}KB.bin" -f $i.ToString("00"), $sizeKB)
    $rng = [System.Security.Cryptography.RNGCryptoServiceProvider]::new()
    $buf = [byte[]]::new($sizeKB * 1024)
    $rng.GetBytes($buf)
    $rng.Dispose()
    [System.IO.File]::WriteAllBytes($fpath, $buf)
}
Write-Host "  OK: 5 个随机内容文件（1-100 KB）" -ForegroundColor DarkGray

# 1 个超大文件（10MB，用于测试流式 SHA256）
try {
    $bigPath = Join-Path $dir03 "large_10MB.bin"
    $rng = [System.Security.Cryptography.RNGCryptoServiceProvider]::new()
    $bigBuf = [byte[]]::new(10 * 1024 * 1024)
    $rng.GetBytes($bigBuf)
    $rng.Dispose()
    [System.IO.File]::WriteAllBytes($bigPath, $bigBuf)
    Write-Host "  OK: large_10MB.bin" -ForegroundColor DarkGray
}
catch {
    Write-Host "  SKIP: large_10MB.bin ($($_.Exception.Message))" -ForegroundColor DarkYellow
}

# =====================================================
# 04_Windows 保留名伪装
# =====================================================
Write-Host ""
Write-Host "[4/4] 生成 Windows 保留名伪装文件..." -ForegroundColor Green

$dir04 = Join-Path $TestRoot "04_Windows保留名伪装"
New-Item -Path $dir04 -ItemType Directory -Force | Out-Null

# 保留名测试（放在子目录下尝试创建）
$reservedTests = @(
    @{ Dir = "CON";    File = "CON.txt";    Content = "reserved name test: CON" },
    @{ Dir = "PRN";    File = "PRN.log";    Content = "reserved name test: PRN" },
    @{ Dir = "AUX";    File = "AUX.dat";    Content = "reserved name test: AUX" },
    @{ Dir = "NUL";    File = "NUL.bin";    Content = "reserved name test: NUL" },
    @{ Dir = "COM1";   File = "COM1.txt";   Content = "reserved name test: COM1" },
    @{ Dir = "LPT1";   File = "LPT1.txt";   Content = "reserved name test: LPT1" },
    @{ Dir = "con";    File = "con.txt";    Content = "lowercase reserved: con" },
    @{ Dir = "aux";    File = "aux.log";    Content = "lowercase reserved: aux" }
)

foreach ($test in $reservedTests) {
    try {
        $testDir = Join-Path $dir04 $test.Dir
        New-Item -Path $testDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null
        $testFile = Join-Path $testDir $test.File
        [System.IO.File]::WriteAllText($testFile, $test.Content, [System.Text.Encoding]::UTF8)
        Write-Host "  OK: $($test.Dir)\$($test.File)" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "  BLOCKED: $($test.Dir)\$($test.File) ($($_.Exception.Message))" -ForegroundColor DarkYellow
    }
}

# 额外：直接在保留名目录下的普通文件
try {
    $extraDir = Join-Path $dir04 "NUL"
    $extraFile = Join-Path $extraDir "normal_file_under_NUL_dir.txt"
    [System.IO.File]::WriteAllText($extraFile, "This file is under a NUL-named directory", [System.Text.Encoding]::UTF8)
    Write-Host "  OK: NUL\normal_file_under_NUL_dir.txt" -ForegroundColor DarkGray
}
catch {
    Write-Host "  SKIP: NUL subfile ($($_.Exception.Message))" -ForegroundColor DarkYellow
}

# =====================================================
# 生成摘要
# =====================================================
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  测试数据生成完毕！" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# 统计
$totalFiles = 0
$totalDirs = 0
try {
    Get-ChildItem -Path $TestRoot -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object { $totalFiles++ }
    Get-ChildItem -Path $TestRoot -Recurse -Directory -ErrorAction SilentlyContinue | ForEach-Object { $totalDirs++ }
}
catch { }

Write-Host "  目录: $TestRoot" -ForegroundColor White
Write-Host "  子目录数: $totalDirs" -ForegroundColor White
Write-Host "  文件数:   $totalFiles" -ForegroundColor White
Write-Host ""
Write-Host "请将此目录作为【源文件夹】运行《图形复制校验工具》" -ForegroundColor Yellow
Write-Host "观察以下关键指标：" -ForegroundColor Yellow
Write-Host "  - 特殊字符文件是否全部正确复制和校验" -ForegroundColor DarkGray
Write-Host "  - 长路径文件是否出现在 LongPathSkipped 统计中" -ForegroundColor DarkGray
Write-Host "  - 零字节文件的 SHA256 是否快速通过" -ForegroundColor DarkGray
Write-Host "  - 保留名文件是否被 Test-PathValid 拦截" -ForegroundColor DarkGray
Write-Host "  - 统计一致性: Total = Skip + Fix + Error + LongPathSkipped" -ForegroundColor DarkGray
