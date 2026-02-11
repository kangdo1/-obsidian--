$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
Get-ChildItem -Path $dir -Filter "*.ps1" | Where-Object { $_.Name -ne "_fix-encoding.ps1" } | ForEach-Object {
    $content = [System.IO.File]::ReadAllText($_.FullName, [System.Text.UTF8Encoding]::new($false))
    [System.IO.File]::WriteAllText($_.FullName, $content, [System.Text.UTF8Encoding]::new($true))
    Write-Host "BOM added: $($_.Name)"
}
