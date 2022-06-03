Write-Host "You might need to run as Administrator and loose the ExecutionPolicy option when running the script." -ForegroundColor Yellow
Write-Host ""
Write-Host "Get the current ExcutionPolicy value in case you want to put it back after the installation."
Write-Host "Get-ExecutionPolicy" -ForegroundColor Green
Get-ExecutionPolicy
Write-Host ""
Write-Host "Only this session will be affected."
Write-Host "Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process" -ForegroundColor Green
Write-Host ""
Write-Host "CurrentUser Profile"
Write-Host "Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser" -ForegroundColor Green

#By Using PowerShell Core we can use the same code accross all the platforms.
if ($PSVersionTable.PSVersion.Major -lt 6) {
    Write-Host "You need to install PowerShell Core to run this script" -ForegroundColor Red
    Write-Host "Please visit this page to install PowerShell Core:"
    Write-Host "https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell?view=powershell-7.2" -ForegroundColor Green
    exit 1
}

try {  
    Install-Module ImportExcel -Scope CurrentUser
    Install-Module Microsoft.Graph -Scope CurrentUser
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red

    if ($isWindows) {
        Write-Host "Specific to Windows:" -ForegroundColor Red
        Write-Host "run this command to delete the PowerShell LockDown Key"
        Write-Host 'reg delete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v _PSLOCKDOWN /f' -ForegroundColor Green
    }    
}