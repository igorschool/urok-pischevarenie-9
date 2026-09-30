$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

function Get-ContentType([string]$path) {
  switch ([IO.Path]::GetExtension($path).ToLowerInvariant()) {
    ".html" { "text/html; charset=utf-8" }
    ".js" { "text/javascript; charset=utf-8" }
    ".css" { "text/css; charset=utf-8" }
    ".json" { "application/json; charset=utf-8" }
    ".wasm" { "application/wasm" }
    ".pdf" { "application/pdf" }
    ".png" { "image/png" }
    ".jpg" { "image/jpeg" }
    ".jpeg" { "image/jpeg" }
    ".gif" { "image/gif" }
    ".webp" { "image/webp" }
    ".mp3" { "audio/mpeg" }
    ".pptx" { "application/vnd.openxmlformats-officedocument.presentationml.presentation" }
    ".txt" { "text/plain; charset=utf-8" }
    default { "application/octet-stream" }
  }
}

$port = 8770
$listener = New-Object System.Net.HttpListener
$started = $false
foreach ($tryPort in 8770, 8771, 8772) {
  $listener = New-Object System.Net.HttpListener
  $listener.Prefixes.Add("http://127.0.0.1:$tryPort/")
  try {
    $listener.Start()
    $port = $tryPort
    $started = $true
    break
  } catch {
    $listener.Close()
  }
}
if (-not $started) {
  Write-Host "Не удалось открыть порт 8770, 8771 или 8772."
  Read-Host "Enter"
  exit 1
}

$url = "http://127.0.0.1:$port/урок.html"
Write-Host "Урок открыт: $url"
Write-Host "Окно не закрывайте до конца урока."
Start-Process $url

while ($listener.IsListening) {
  $context = $listener.GetContext()
  $request = $context.Request
  $response = $context.Response
  try {
    $rel = [Uri]::UnescapeDataString($request.Url.AbsolutePath.TrimStart("/"))
    if ([string]::IsNullOrWhiteSpace($rel)) { $rel = "index.html" }
    $rel = $rel -replace "/", [IO.Path]::DirectorySeparatorChar
    $full = [IO.Path]::GetFullPath((Join-Path $root $rel))
    if (-not $full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase) -or -not (Test-Path -LiteralPath $full -PathType Leaf)) {
      $response.StatusCode = 404
      $bytes = [Text.Encoding]::UTF8.GetBytes("Не найдено")
      $response.OutputStream.Write($bytes, 0, $bytes.Length)
    } else {
      $bytes = [IO.File]::ReadAllBytes($full)
      $response.ContentType = Get-ContentType $full
      $response.ContentLength64 = $bytes.Length
      $response.OutputStream.Write($bytes, 0, $bytes.Length)
    }
  } catch {
    $response.StatusCode = 500
  } finally {
    $response.OutputStream.Close()
  }
}
