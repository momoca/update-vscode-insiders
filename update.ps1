#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

# ================================
# 0. 基础配置
# ================================
$BaseDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$ZipFile   = Join-Path $BaseDir 'vscode-insider.zip'
$PartFile  = Join-Path $BaseDir 'vscode-insider.zip.part'
$TempDir   = Join-Path $BaseDir 'update_temp'
$Ignore    = '072586267e'
$Url       = 'https://update.code.visualstudio.com/latest/win32-x64-archive/insider'
$UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) PowerShell-Updater/1.0'

# 尽量启用 TLS 1.2（Windows PowerShell 5.1 常见需要）
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {}

function Write-Step($msg) {
    Write-Host "`n=== $msg ==="
}

function Get-RedirectLocation {
    param([Parameter(Mandatory)][string]$Uri)

    try {
        $req = [System.Net.HttpWebRequest]::Create($Uri)
        $req.Method = 'HEAD'
        $req.AllowAutoRedirect = $false
        $req.UserAgent = $UserAgent
        $resp = $req.GetResponse()
        $location = $resp.Headers['Location']
        $resp.Close()
        return $location
    }
    catch [System.Net.WebException] {
        if ($_.Exception.Response) {
            $resp = $_.Exception.Response
            $location = $resp.Headers['Location']
            $resp.Close()
            if ($location) { return $location }
        }
        throw
    }
}

function Get-RemoteFileLength {
    param([Parameter(Mandatory)][string]$Uri)

    try {
        $req = [System.Net.HttpWebRequest]::Create($Uri)
        $req.Method = 'HEAD'
        $req.AllowAutoRedirect = $true
        $req.UserAgent = $UserAgent
        $resp = $req.GetResponse()
        $len = $resp.ContentLength
        $resp.Close()
        if ($len -gt 0) { return [int64]$len }
    }
    catch {
        return $null
    }
    return $null
}

function Download-FileRobust {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string]$PartDestination,
        [int]$MaxAttempts = 5
    )

    $lastError = $null

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Write-Host "下载尝试 $attempt/$MaxAttempts"

            if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue) {
                Import-Module BitsTransfer -ErrorAction SilentlyContinue | Out-Null
                if (Test-Path $PartDestination) {
                    Remove-Item $PartDestination -Force -ErrorAction SilentlyContinue
                }
                Start-BitsTransfer -Source $Uri -Destination $PartDestination -DisplayName 'VSCodeInsiderUpdate' -Description 'Downloading VS Code Insider zip' -ErrorAction Stop
            }
            else {
                $iwrArgs = @{
                    Uri               = $Uri
                    OutFile           = $PartDestination
                    UserAgent         = $UserAgent
                    MaximumRetryCount = 3
                    RetryIntervalSec  = 5
                    ErrorAction       = 'Stop'
                }

                if ((Test-Path $PartDestination) -and $PSVersionTable.PSVersion.Major -ge 6) {
                    $iwrArgs['Resume'] = $true
                }

                Invoke-WebRequest @iwrArgs
            }

            if (!(Test-Path $PartDestination)) {
                throw '下载结束但未生成文件。'
            }

            $size = (Get-Item $PartDestination).Length
            if ($size -lt 10MB) {
                throw "下载文件过小，疑似未完整下载。当前大小: $size 字节"
            }

            if (Test-Path $Destination) {
                Remove-Item $Destination -Force -ErrorAction SilentlyContinue
            }
            Move-Item $PartDestination $Destination -Force
            return
        }
        catch {
            $lastError = $_
            Write-Warning "下载失败：$($_.Exception.Message)"
            if ($attempt -lt $MaxAttempts) {
                $sleepSec = [Math]::Min(30, 3 * $attempt)
                Write-Host "$sleepSec 秒后重试..."
                Start-Sleep -Seconds $sleepSec
            }
        }
    }

    throw "下载最终失败：$($lastError.Exception.Message)"
}
function Wait-AndExit {
    param(
        [int]$Code = 0,
        [string]$Prompt = '按任意键退出...'
    )

    Write-Host ""
    Write-Host $Prompt

    try {
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
    catch {
        Read-Host '按回车退出' | Out-Null
    }

    exit $Code
}
try {
    Write-Host "工作目录: $BaseDir"

    Write-Step '检查远程版本'
    $Location = Get-RedirectLocation -Uri $Url
    if (-not $Location) {
        throw '无法获取重定向地址。'
    }

    $RemoteSha = ([regex]::Match($Location, 'insider/([0-9a-f]{40})/')).Groups[1].Value
    if (-not $RemoteSha) {
        throw "无法从重定向地址解析远程版本：$Location"
    }

    $RemoteVersion = $RemoteSha.Substring(0, 10)
    $LocalRemotePath = Join-Path $BaseDir $RemoteVersion

    Write-Host "远程版本目录: $RemoteVersion"
    Write-Host "下载地址: $Location"

    if (Test-Path $LocalRemotePath) {
        Write-Host "本地已存在最新版本目录，无需更新。"
        Wait-AndExit -Code 0
    }

    Write-Step '下载最新 ZIP'
    $remoteLength = Get-RemoteFileLength -Uri $Location
    if ($remoteLength) {
        Write-Host "远程文件大小: $remoteLength 字节"
    }

    if (Test-Path $ZipFile) {
        Remove-Item $ZipFile -Force -ErrorAction SilentlyContinue
    }

    Download-FileRobust -Uri $Location -Destination $ZipFile -PartDestination $PartFile -MaxAttempts 5

    $localLength = (Get-Item $ZipFile).Length
    Write-Host "本地文件大小: $localLength 字节"

    if ($remoteLength -and $localLength -ne $remoteLength) {
        throw "下载大小不一致。远程: $remoteLength，本地: $localLength"
    }

    Write-Step '解压到临时目录'
    if (Test-Path $TempDir) {
        Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    Expand-Archive $ZipFile -DestinationPath $TempDir -Force

    $NewVersion = Get-ChildItem $TempDir -Directory |
        Where-Object { $_.Name -match '^[0-9a-f]{10}$' -and $_.Name -ne $Ignore } |
        Select-Object -ExpandProperty Name -First 1

    if (-not $NewVersion) {
        throw '解压成功，但未找到新的版本目录。'
    }

    Write-Host "更新版本目录: $NewVersion"

    Write-Step '识别当前旧版本目录'
    $CurrentVersion = Get-ChildItem $BaseDir -Directory |
        Where-Object { $_.Name -match '^[0-9a-f]{10}$' -and $_.Name -ne $Ignore -and $_.Name -ne $NewVersion } |
        Select-Object -ExpandProperty Name -First 1

    if ($CurrentVersion) {
        Write-Host "当前旧版本目录: $CurrentVersion"
    } else {
        Write-Host '未检测到旧版本目录。'
    }

    Write-Step '覆盖更新（保留 data）'
    Copy-Item -Path (Join-Path $TempDir '*') -Destination $BaseDir -Recurse -Force -Exclude 'data'

    Write-Step '清理旧版本'
    if ($CurrentVersion) {
        $PathToDelete = Join-Path $BaseDir $CurrentVersion
        if (Test-Path $PathToDelete) {
            Remove-Item $PathToDelete -Recurse -Force
            Write-Host "已删除旧版本目录: $CurrentVersion"
        }
    }

    Write-Step '清理临时文件'
    if (Test-Path $TempDir) {
        Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path $ZipFile) {
        Remove-Item $ZipFile -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path $PartFile) {
        Remove-Item $PartFile -Force -ErrorAction SilentlyContinue
    }

    Write-Host "`n更新完成，当前版本: $NewVersion"
    Wait-AndExit -Code 0
}
catch {
    Write-Error $_
    Write-Host "`n更新失败，请检查网络、代理、证书链或目标服务器状态。"
    Wait-AndExit -Code 1
}
