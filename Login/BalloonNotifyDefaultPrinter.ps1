
function Show-BalloonTip {            
[cmdletbinding()]            
param(            
 [parameter(Mandatory=$true)]            
 [string]$Title,            
 [ValidateSet("Info","Warning","Error")]             
 [string]$MessageType = "Info",            
 [parameter(Mandatory=$true)]            
 [string]$Message,            
 [string]$Duration=100000            
)            
 
[system.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null            
$balloon = New-Object System.Windows.Forms.NotifyIcon            
$path = Get-Process -id $pid | Select-Object -ExpandProperty Path            
$icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\support\printer-ink.ico")            
$balloon.Icon = $icon            
$balloon.BalloonTipIcon = $MessageType            
$balloon.BalloonTipText = $Message            
$balloon.BalloonTipTitle = $Title            
$balloon.Visible = $true            
$balloon.ShowBalloonTip($Duration)  
}

$defaultprt = Get-WmiObject -Query "SELECT ShareName FROM Win32_Printer WHERE Default = $true" | Select-Object ShareName

# The nature of running a script a logon is that some resources
# are not available instantly. Here we loop for at most 20secs
# looking for the default printer object.
[int]$cntr = 0
while (!$defaultprt) {
    Start-Sleep -Seconds 5
    $defaultprt = Get-WmiObject -Query "SELECT ShareName FROM Win32_Printer WHERE Default = $true" | Select-Object ShareName    
    #"Loop Count: " + $cntr + " - " + "Detected Default Printer: " + $defaultprt.ShareName >> "C:\Support\BalloonNotifyDefaultPrinter.txt"
    $cntr = $cntr + 1
    If ($cntr -eq 5) {break}
}

Show-BalloonTip -Title "Default Printer" -MessageType "info" -Message ("Your default printer is currently set to " + $defaultprt.ShareName) -Duration 5000