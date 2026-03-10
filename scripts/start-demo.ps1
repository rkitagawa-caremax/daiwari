$ErrorActionPreference = 'Stop'

$hostIp = '127.0.0.1'
$basePort = 5173
$maxPort = 5200
$urlPath = '/?mode=local'
$logFile = Join-Path $PSScriptRoot '..\.demo-server.log'

function Get-HttpStatusCode {
  param(
    [string]$Url
  )
  try {
    $res = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 2
    return [int]$res.StatusCode
  } catch {
    return 0
  }
}

function Is-PortListening {
  param(
    [int]$Port
  )
  $lines = netstat -ano | Select-String ":$Port\s+.*LISTENING"
  return $null -ne $lines -and $lines.Count -gt 0
}

function Is-ViteServer {
  param(
    [int]$Port
  )
  $url = "http://$hostIp`:$Port/"
  try {
    $res = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 2
    return $res.Content -like '*@vite/client*'
  } catch {
    return $false
  }
}

function Get-AvailablePort {
  for ($port = $basePort; $port -le $maxPort; $port++) {
    if (-not (Is-PortListening -Port $port)) {
      return $port
    }
    if (Is-ViteServer -Port $port) {
      return $port
    }
  }
  throw "No available port found between $basePort and $maxPort."
}

$port = Get-AvailablePort
$baseUrl = "http://$hostIp`:$port"
$targetUrl = "$baseUrl$urlPath"

if (-not (Is-ViteServer -Port $port)) {
  $projectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
  if (Test-Path $logFile) {
    Remove-Item $logFile -Force
  }

  $command = "npm run dev -- --host $hostIp --port $port --strictPort *> `"$logFile`""
  Start-Process -FilePath powershell.exe -ArgumentList '-NoProfile', '-Command', $command -WorkingDirectory $projectRoot -WindowStyle Hidden | Out-Null

  $started = $false
  for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Milliseconds 500
    if (Get-HttpStatusCode -Url $baseUrl) {
      $started = $true
      break
    }
  }

  if (-not $started) {
    if (Test-Path $logFile) {
      Get-Content $logFile | Select-Object -Last 80
    }
    throw "Failed to start demo server. See .demo-server.log for details."
  }
}

Start-Process $targetUrl
Write-Output "Demo opened: $targetUrl"
