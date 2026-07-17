# Masaustune "UEVCB Raporu" kisayolu olusturur (kurulum.bat cagirir).
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$ws = New-Object -ComObject WScript.Shell
$lnk = $ws.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) 'UEVCB Raporu.lnk'))
if (Test-Path 'C:\Windows\pyw.exe') {
    # pyw.exe: konsol penceresi acilmadan dogrudan arayuz
    $lnk.TargetPath = 'C:\Windows\pyw.exe'
    $lnk.Arguments = '"' + (Join-Path $here 'uevcb_rapor.py') + '"'
} else {
    $lnk.TargetPath = (Join-Path $here 'arayuz.bat')
}
$lnk.WorkingDirectory = $here
$lnk.IconLocation = 'C:\Windows\System32\imageres.dll,165'
$lnk.Description = 'EPIAS UEVCB Raporu'
$lnk.Save()
Write-Host 'Masaustune "UEVCB Raporu" kisayolu olusturuldu.'
