<#
  serve-board.ps1

  Serves docs/ on localhost so the board can be opened in a browser exactly as
  it will appear when deployed - without deploying anything.

  Why a server is needed at all: index.html loads its data with
  fetch('data/schedule.json'), and browsers refuse that over file://. Opening
  docs/index.html by double-clicking gives a board with every panel blank, which
  looks like a data problem but is not one.

  No Node or Python required - this uses the .NET HTTP listener that ships with
  Windows PowerShell.

  Usage:
      powershell -ExecutionPolicy Bypass -File tools/serve-board.ps1
  then open http://localhost:8099/ and press Ctrl+C here when finished.

  Nothing here writes to the repo. To refresh the data first, run
  ./extract-schedule.ps1 - that rebuilds docs/data/schedule.json locally and
  still touches nothing remote.
#>
param(
    [string]$Root = (Join-Path $PSScriptRoot '..\docs'),
    [int]$Port = 8099
)

$ErrorActionPreference = 'Stop'
$Root = (Resolve-Path $Root).Path

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
try {
    $listener.Start()
} catch {
    throw "Could not listen on port $Port. Another copy may already be running, or try -Port 8100."
}

Write-Host ""
Write-Host "  Serving $Root"
Write-Host "  Open   http://localhost:$Port/"
Write-Host "  Ctrl+C to stop"
Write-Host ""

$types = @{
    '.html' = 'text/html; charset=utf-8'
    '.js'   = 'application/javascript; charset=utf-8'
    '.css'  = 'text/css; charset=utf-8'
    '.json' = 'application/json; charset=utf-8'
    '.txt'  = 'text/plain; charset=utf-8'
    '.svg'  = 'image/svg+xml'
    '.png'  = 'image/png'
    '.jpg'  = 'image/jpeg'
    '.woff2' = 'font/woff2'
}

try {
    while ($listener.IsListening) {
        $ctx = $listener.GetContext()
        $rel = [Uri]::UnescapeDataString($ctx.Request.Url.AbsolutePath).TrimStart('/')
        if ($rel -eq '') { $rel = 'index.html' }
        $path = Join-Path $Root $rel

        # Refuse anything that escapes the served directory.
        $full = [System.IO.Path]::GetFullPath($path)
        if (-not $full.StartsWith($Root, [StringComparison]::OrdinalIgnoreCase)) {
            $ctx.Response.StatusCode = 403
            Write-Host ("403 {0}" -f $rel)
            $ctx.Response.Close()
            continue
        }

        try {
            if (Test-Path $full -PathType Leaf) {
                $ext = [System.IO.Path]::GetExtension($full).ToLower()
                $ct = 'application/octet-stream'
                if ($types.ContainsKey($ext)) { $ct = $types[$ext] }
                $bytes = [System.IO.File]::ReadAllBytes($full)
                $ctx.Response.ContentType = $ct
                # The board fetches with cache: 'no-store', but be explicit so a
                # re-run of the extractor is picked up on a plain reload too.
                $ctx.Response.Headers.Add('Cache-Control', 'no-store')
                $ctx.Response.ContentLength64 = $bytes.Length
                $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
                Write-Host ("200 {0}" -f $rel)
            } else {
                $ctx.Response.StatusCode = 404
                Write-Host ("404 {0}" -f $rel)
            }
        } catch {
            $ctx.Response.StatusCode = 500
            Write-Host ("500 {0}: {1}" -f $rel, $_.Exception.Message)
        }
        $ctx.Response.Close()
    }
} finally {
    $listener.Stop()
    $listener.Close()
    Write-Host "Stopped."
}
