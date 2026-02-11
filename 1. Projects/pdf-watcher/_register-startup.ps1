$ws = New-Object -ComObject WScript.Shell
$startupDir = [System.IO.Path]::Combine($env:APPDATA, 'Microsoft\Windows\Start Menu\Programs\Startup')
$shortcut = $ws.CreateShortcut("$startupDir\PdfWatcher.lnk")
$shortcut.TargetPath = 'powershell.exe'
$shortcut.Arguments = '-ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Users\user\OneDrive - 강도원\문서\Obsidian_KDW\KDW\1. Projects\pdf-watcher\Watch-Pdf.ps1"'
$shortcut.WindowStyle = 7
$shortcut.Save()
Write-Host "시작 프로그램에 PdfWatcher 등록 완료"
