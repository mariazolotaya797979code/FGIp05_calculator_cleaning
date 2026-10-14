$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$exe = Join-Path $root "Kalkulyator.exe"
$icon = Join-Path $root "icon.ico"
$desktop = [Environment]::GetFolderPath("Desktop")
$lnkPath = Join-Path $desktop "Калькулятор клининга.lnk"
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($lnkPath)
$shortcut.WorkingDirectory = $root
$shortcut.WindowStyle = 1
$shortcut.Description = "Калькулятор стоимости клининга"
if (Test-Path $icon) {
    $shortcut.IconLocation = "$icon,0"
}
if (Test-Path $exe) {
    $shortcut.TargetPath = $exe
} else {
    $pythonw = Get-Command pythonw.exe -ErrorAction SilentlyContinue
    if ($null -eq $pythonw) {
        throw "Не найден Kalkulyator.exe. Для запуска из исходников нужен установленный Python."
    }
    $shortcut.TargetPath = $pythonw.Source
    $shortcut.Arguments = '"' + (Join-Path $root "calculator_app.py") + '"'
    if (-not (Test-Path $icon)) {
        $shortcut.IconLocation = "$($pythonw.Source),0"
    }
}
$shortcut.Save()
Write-Host "На рабочем столе создана иконка: Калькулятор клининга"
