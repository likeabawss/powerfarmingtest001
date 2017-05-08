Function Log-Start{
  <#
  .SYNOPSIS
    Creates log file
  .DESCRIPTION
    Creates log file with path and name that is passed. Checks if log file exists, and if it does deletes it and creates a new one.
    Once created, writes initial logging data
  .PARAMETER LogPath
    Mandatory. Path of where log is to be created. Example: C:\Windows\Temp
  .PARAMETER LogName
    Mandatory. Name of log file to be created. Example: Test_Script.log
      
  .PARAMETER ScriptVersion
    Mandatory. Version of the running script which will be written in the log. Example: 1.5
  .INPUTS
    Parameters above
  .OUTPUTS
    Log file created
  .NOTES
    Version:        1.0
    Author:         Luca Sturlese
    Creation Date:  10/05/12
    Purpose/Change: Initial function development
    Version:        1.1
    Author:         Luca Sturlese
    Creation Date:  19/05/12
    Purpose/Change: Added debug mode support
  .EXAMPLE
    Log-Start -LogPath "C:\Windows\Temp" -LogName "Test_Script.log" -ScriptVersion "1.5"
  #>
    
  [CmdletBinding()]
  
  Param ([Parameter(Mandatory=$false)][string]$LogPath, [Parameter(Mandatory=$false)][string]$LogName, [Parameter(Mandatory=$false)][string]$ScriptVersion)
  
  Process{
    #$sFullPath = $LogPath + "\" + $LogName
        
    #If($LogPath -and $LogName)
    #{
    #    $LogName = $MyInvocation.ScriptName.ToString().Replace(".ps1",".txt")
    #    Write-Host $LogName
    #}
    
    $sFullPath = $MyInvocation.ScriptName.ToString().Replace(".ps1",".txt")
    #

    #Check if file exists and delete if it does
    #If((Test-Path -Path $sFullPath)){
    #  Remove-Item -Path $sFullPath -Force
    #}
    
    #Create file and start logging
    If(-Not(Test-Path -Path $sFullPath)){
        New-Item ($sFullPath) -ItemType File | Out-Null
    }
    
    Add-Content -Path $sFullPath -Value "***************************************************************************************************"
    Add-Content -Path $sFullPath -Value "-->>SCRIPT: $($MyInvocation.ScriptName.ToString())"
    Add-Content -Path $sFullPath -Value "-->>STARTED processing at [$([DateTime]::Now)]."
    Add-Content -Path $sFullPath -Value "***************************************************************************************************"
    #Add-Content -Path $sFullPath -Value ""
    #Add-Content -Path $sFullPath -Value "Running script version [$ScriptVersion]."
    #Add-Content -Path $sFullPath -Value ""
    #Add-Content -Path $sFullPath -Value "***************************************************************************************************"
    #Add-Content -Path $sFullPath -Value ""
  
    #Write to screen for debug mode
    Write-Debug "***************************************************************************************************"
    Write-Debug "Started processing at [$([DateTime]::Now)]."
    Write-Debug "***************************************************************************************************"
    #Write-Debug ""
    Write-Debug "Running script version [$ScriptVersion]."
    #Write-Debug ""
    Write-Debug "***************************************************************************************************"
    Write-Debug ""
  }
}

Function Log-Write{
  <#
  .SYNOPSIS
    Writes to a log file
  .DESCRIPTION
    Appends a new line to the end of the specified log file
  
  .PARAMETER LogPath
    Mandatory. Full path of the log file you want to write to. Example: C:\Windows\Temp\Test_Script.log
  
  .PARAMETER LineValue
    Mandatory. The string that you want to write to the log
      
  .INPUTS
    Parameters above
  .OUTPUTS
    None
  .NOTES
    Version:        1.0
    Author:         Luca Sturlese
    Creation Date:  10/05/12
    Purpose/Change: Initial function development
  
    Version:        1.1
    Author:         Luca Sturlese
    Creation Date:  19/05/12
    Purpose/Change: Added debug mode support
  .EXAMPLE
    Log-Write -LogPath "C:\Windows\Temp\Test_Script.log" -LineValue "This is a new line which I am appending to the end of the log file."
  #>
  
  [CmdletBinding()]
  
  Param ([Parameter(Mandatory=$false)][string]$LogPath, [Parameter(Mandatory=$false)][string]$LineValue)
  
  Process{
    $LogPath = $MyInvocation.ScriptName.ToString().Replace(".ps1",".txt")
    Add-Content -Path $LogPath -Value $LineValue   

    #Write to screen for debug mode
    Write-Debug $LineValue
  }
}

Function Log-Error{
  <#
  .SYNOPSIS
    Writes an error to a log file
  .DESCRIPTION
    Writes the passed error to a new line at the end of the specified log file
  
  .PARAMETER LogPath
    Mandatory. Full path of the log file you want to write to. Example: C:\Windows\Temp\Test_Script.log
  
  .PARAMETER ErrorDesc
    Mandatory. The description of the error you want to pass (use $_.Exception)
  
  .PARAMETER ExitGracefully
    Mandatory. Boolean. If set to True, runs Log-Finish and then exits script
  .INPUTS
    Parameters above
  .OUTPUTS
    None
  .NOTES
    Version:        1.0
    Author:         Luca Sturlese
    Creation Date:  10/05/12
    Purpose/Change: Initial function development
    
    Version:        1.1
    Author:         Luca Sturlese
    Creation Date:  19/05/12
    Purpose/Change: Added debug mode support. Added -ExitGracefully parameter functionality
  .EXAMPLE
    Log-Error -LogPath "C:\Windows\Temp\Test_Script.log" -ErrorDesc $_.Exception -ExitGracefully $True
  #>
  
  [CmdletBinding()]
  
  Param ([Parameter(Mandatory=$true)][string]$LogPath, [Parameter(Mandatory=$true)][string]$ErrorDesc, [Parameter(Mandatory=$true)][boolean]$ExitGracefully)
  
  Process{
    Add-Content -Path $LogPath -Value "Error: An error has occurred [$ErrorDesc]."
  
    #Write to screen for debug mode
    Write-Debug "Error: An error has occurred [$ErrorDesc]."
    
    #If $ExitGracefully = True then run Log-Finish and exit script
    If ($ExitGracefully -eq $True){
      Log-Finish -LogPath $LogPath
      Break
    }
  }
}

Function Log-Finish{
  <#
  .SYNOPSIS
    Write closing logging data & exit
  .DESCRIPTION
    Writes finishing logging data to specified log and then exits the calling script
  
  .PARAMETER LogPath
    Mandatory. Full path of the log file you want to write finishing data to. Example: C:\Windows\Temp\Test_Script.log
  .PARAMETER NoExit
    Optional. If this is set to True, then the function will not exit the calling script, so that further execution can occur
  
  .INPUTS
    Parameters above
  .OUTPUTS
    None
  .NOTES
    Version:        1.0
    Author:         Luca Sturlese
    Creation Date:  10/05/12
    Purpose/Change: Initial function development
    
    Version:        1.1
    Author:         Luca Sturlese
    Creation Date:  19/05/12
    Purpose/Change: Added debug mode support
  
    Version:        1.2
    Author:         Luca Sturlese
    Creation Date:  01/08/12
    Purpose/Change: Added option to not exit calling script if required (via optional parameter)
  .EXAMPLE
    Log-Finish -LogPath "C:\Windows\Temp\Test_Script.log"
.EXAMPLE
    Log-Finish -LogPath "C:\Windows\Temp\Test_Script.log" -NoExit $True
  #>
  
  [CmdletBinding()]
  
  Param ([Parameter(Mandatory=$false)][string]$LogPath, [Parameter(Mandatory=$false)][string]$NoExit)
  
  Process{

    $LogPath = $MyInvocation.ScriptName.ToString().Replace(".ps1",".txt")

    Add-Content -Path $LogPath -Value ""
    Add-Content -Path $LogPath -Value "***************************************************************************************************"
    Add-Content -Path $LogPath -Value "<<--SCRIPT: $($MyInvocation.ScriptName.ToString())"
    Add-Content -Path $LogPath -Value "<<--ENDED processing at [$([DateTime]::Now)]."
    Add-Content -Path $LogPath -Value "***************************************************************************************************"
    Add-Content -Path $LogPath -Value ""
  
    #Write to screen for debug mode
    Write-Debug ""
    Write-Debug "***************************************************************************************************"
    Write-Debug "<<--ENDED processing at [$([DateTime]::Now)]."
    Write-Debug "***************************************************************************************************"
  
    #Exit calling script if NoExit has not been specified or is set to False
    #If(!($NoExit) -or ($NoExit -eq $False)){
    #  Exit
    #}    
  }
}

Function Log-Email{
  <#
  .SYNOPSIS
    Emails log file to list of recipients
  .DESCRIPTION
    Emails the contents of the specified log file to a list of recipients
  
  .PARAMETER LogPath
    Mandatory. Full path of the log file you want to email. Example: C:\Windows\Temp\Test_Script.log
  
  .PARAMETER EmailFrom
    Mandatory. The email addresses of who you want to send the email from. Example: "admin@9to5IT.com"
  .PARAMETER EmailTo
    Mandatory. The email addresses of where to send the email to. Seperate multiple emails by ",". Example: "admin@9to5IT.com, test@test.com"
  
  .PARAMETER EmailSubject
    Mandatory. The subject of the email you want to send. Example: "Cool Script - [" + (Get-Date).ToShortDateString() + "]"
  .INPUTS
    Parameters above
  .OUTPUTS
    Email sent to the list of addresses specified
  .NOTES
    Version:        1.0
    Author:         Luca Sturlese
    Creation Date:  05.10.12
    Purpose/Change: Initial function development
  .EXAMPLE
    Log-Email -LogPath "C:\Windows\Temp\Test_Script.log" -EmailFrom "admin@9to5IT.com" -EmailTo "admin@9to5IT.com, test@test.com" -EmailSubject "Cool Script - [" + (Get-Date).ToShortDateString() + "]"
  #>
  
  [CmdletBinding()]
  
  Param ([Parameter(Mandatory=$true)][string]$LogPath, [Parameter(Mandatory=$true)][string]$EmailFrom, [Parameter(Mandatory=$true)][string]$EmailTo, [Parameter(Mandatory=$true)][string]$EmailSubject)
  
  Process{
    Try{
      $sBody = (Get-Content $LogPath | out-string)
      
      #Create SMTP object and send email
      $sSmtpServer = "smtp.yourserver"
      $oSmtp = new-object Net.Mail.SmtpClient($sSmtpServer)
      $oSmtp.Send($EmailFrom, $EmailTo, $EmailSubject, $sBody)
      Exit 0
    }
    
    Catch{
      Exit 1
    } 
  }
}

# SIG # Begin signature block
# MIIHvgYJKoZIhvcNAQcCoIIHrzCCB6sCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQURfhKAhGW5Gss2tiJxAeUyo8Y
# OYugggWNMIIFiTCCBHGgAwIBAgIKT4GyCAAAAAACBjANBgkqhkiG9w0BAQUFADBt
# MRIwEAYKCZImiZPyLGQBGRYCbnoxEjAQBgoJkiaJk/IsZAEZFgJjbzEcMBoGCgmS
# JomT8ixkARkWDHBvd2VyZmFybWluZzElMCMGA1UEAxMccG93ZXJmYXJtaW5nLVBG
# TlotU1JWLTAyOC1DQTAeFw0xNjA4MjYwMzU5MTFaFw0xNzA4MjYwMzU5MTFaMIGN
# MRIwEAYKCZImiZPyLGQBGRYCbnoxEjAQBgoJkiaJk/IsZAEZFgJjbzEcMBoGCgmS
# JomT8ixkARkWDHBvd2VyZmFybWluZzEeMBwGA1UECxMVTmV3IFplYWxhbmQgV2hv
# bGVzYWxlMQswCQYDVQQLEwJJVDEYMBYGA1UEAxMPTWljaGFlbCBCYXJyZXR0MIGf
# MA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQCV/ir3HELuaq1D9LnarvAGUM7D9tei
# ZEp/I89cvIsLb8lptQFlcXugvz8JJWPEgGVHTw9ocA8OCBC1CQJrQBibmD3HHx8w
# 0lskhbwF+7ydZ7oR12omgmVn6OrYNOUp8nj4dC1KsAr/hVgMl1kBIXLWTS9WI/0R
# P9itG2uRkCiSjQIDAQABo4ICjDCCAogwJQYJKwYBBAGCNxQCBBgeFgBDAG8AZABl
# AFMAaQBnAG4AaQBuAGcwEwYDVR0lBAwwCgYIKwYBBQUHAwMwCwYDVR0PBAQDAgeA
# MB0GA1UdDgQWBBT9SDwzNzMSZFQ6nNLV6oxwQZ50fDAfBgNVHSMEGDAWgBTDKX3i
# xQYwhJKqfUODbUNBL5OTgTCB6QYDVR0fBIHhMIHeMIHboIHYoIHVhoHSbGRhcDov
# Ly9DTj1wb3dlcmZhcm1pbmctUEZOWi1TUlYtMDI4LUNBLENOPVBGTlotU1JWLTAy
# OCxDTj1DRFAsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2VydmljZXMs
# Q049Q29uZmlndXJhdGlvbixEQz1wb3dlcmZhcm1pbmcsREM9Y28sREM9bno/Y2Vy
# dGlmaWNhdGVSZXZvY2F0aW9uTGlzdD9iYXNlP29iamVjdENsYXNzPWNSTERpc3Ry
# aWJ1dGlvblBvaW50MIHYBggrBgEFBQcBAQSByzCByDCBxQYIKwYBBQUHMAKGgbhs
# ZGFwOi8vL0NOPXBvd2VyZmFybWluZy1QRk5aLVNSVi0wMjgtQ0EsQ049QUlBLENO
# PVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENOPVNlcnZpY2VzLENOPUNvbmZpZ3Vy
# YXRpb24sREM9cG93ZXJmYXJtaW5nLERDPWNvLERDPW56P2NBQ2VydGlmaWNhdGU/
# YmFzZT9vYmplY3RDbGFzcz1jZXJ0aWZpY2F0aW9uQXV0aG9yaXR5MDYGA1UdEQQv
# MC2gKwYKKwYBBAGCNxQCA6AdDBttYmFycmV0dEBwb3dlcmZhcm1pbmcuY28ubnow
# DQYJKoZIhvcNAQEFBQADggEBAA2yYue4h05YY1ps9J3QhL+0UW9McEqiJVSOZq5a
# 1eXxFsCsDDAgaUtY/m+NAv7mYOshgnCBs2mTZvMn6mqn97L96z37VYZnUOpJwg/X
# 1Laul3S7tFBqJxHO1Z3xWrMNvzPkCivOExbG/sZbmip6jHb9KlWgnqjQ1WRNI75x
# aJT8bpunM7iHycaUwtoUxfNX5qG6vJ4PdwFUsMlMQyUh+siLYE1v2PxRXZReZmEu
# XPQ6rQ0KXsTKZ7jdAfiyNSah8rIgucuvlXRdtDh99n7KiTlaX2vH/XQ3fChY1QXW
# kuETGEuAapQ800YxGIt/K97ZI+AsmTKaG7CYkvHnTeE/kNIxggGbMIIBlwIBATB7
# MG0xEjAQBgoJkiaJk/IsZAEZFgJuejESMBAGCgmSJomT8ixkARkWAmNvMRwwGgYK
# CZImiZPyLGQBGRYMcG93ZXJmYXJtaW5nMSUwIwYDVQQDExxwb3dlcmZhcm1pbmct
# UEZOWi1TUlYtMDI4LUNBAgpPgbIIAAAAAAIGMAkGBSsOAwIaBQCgeDAYBgorBgEE
# AYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwG
# CisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBSwxB/r
# eSfWv3QcRtsv3NAD80w/0jANBgkqhkiG9w0BAQEFAASBgBzKKQWjziCHdmlXyhj9
# 9xNyiKWq0Cj2hnKc9w3k51gsgzVHHg24P+sLrZHWDLbvWj2658YO2+QHn6O4gR4t
# DgecRKv7+Fvg2CNmnE/in1iIzR77OmaKBKQsLXz+jF7Qy6062WeIa4PR/2bQSe1h
# t8TzSyqD5XeoD0sDusp8Hssp
# SIG # End signature block
