###########################################################################" 
# 
# NAME: Set-OutlookSignature.ps1 
# 
# AUTHOR: Jan Egil Ring 
# Modifications by Darren Kattan 
# 
# COMMENT: Script to create an Outlook signature based on user information from Active Directory. 
# Adjust the variables in the "Custom variables"-section 
# Create an Outlook-signature from Microsoft Word (logo, fonts etc) and copy this signature to \\domain\NETLOGON\sig_files\$CompanyName\$CompanyName.docx 
#     This script supports the following keywords: 
#     DisplayName 
#     Title 
#     Email 
# 
#    See the following blog-posts for more information:  
#    http://blog.powershell.no/2010/01/09/outlook-signature-based-on-user-information-from-active-directory 
#    http://www.immense.net/deploying-unified-email-signature-template-outlook 
# 
# Tested on Office 2003, 2007 and 2010 
# 
# You have a royalty-free right to use, modify, reproduce, and 
# distribute this script file in any way you find useful, provided that 
# you agree that the creator, owner above has no warranty, obligations, 
# or liability for such use. 
# 
# VERSION HISTORY: 
# 1.0 09.01.2010 – Initial release 
# 1.1 11.09.2010 – Modified by Darren Kattan 
#    - Removed bookmarks. Now uses simple find and replace for DisplayName, Title, and Email. 
#    - Email address is generated as a link 
#    - Signature is generated from a single .docx file 
#    - Removed version numbers for script to run. Script runs at boot up when it sees a change in the "Date Modified" property of your signature template. 
# 
# 
###########################################################################" 
 
#Custom variables 
$CompanyName = 'Power Farming New Zealand' 
$DomainName = 'power' 
 
$SigSource = "\\$DomainName\netlogon\svn-netlogon\SignatureFiles\$CompanyName" 
$ForceSignatureNew = '0' #When the signature are forced the signature are enforced as default signature for new messages the next time the script runs. 0 = no force, 1 = force 
$ForceSignatureReplyForward = '0' #When the signature are forced the signature are enforced as default signature for reply/forward messages the next time the script runs. 0 = no force, 1 = force 
 
#Environment variables 
$AppData=(Get-Item env:appdata).value 
$SigPath = '\Microsoft\Signatures' 
$LocalSignaturePath = $AppData+$SigPath 
$RemoteSignaturePathFull = $SigSource+'\'+$CompanyName+'.docx' 
 
#Get Active Directory information for current user 
$UserName = $env:username 
$Filter = "(&(objectCategory=User)(samAccountName=$UserName))" 
$Searcher = New-Object System.DirectoryServices.DirectorySearcher 
$Searcher.Filter = $Filter 
$ADUserPath = $Searcher.FindOne() 
$ADUser = $ADUserPath.GetDirectoryEntry() 
$ADDisplayName = $ADUser.DisplayName 
$ADEmailAddress = $ADUser.mail 
$ADTitle = $ADUser.title 
$ADTelePhoneNumber = $ADUser.TelephoneNumber 
$ADMobileNumber = $ADUser.Mobile 
$ADFaxNumber = $ADUser.FacsimileTelephoneNumber 
$ADLastChanged = $ADUser.whenChanged
 
#Setting registry information for the current user 
$CompanyRegPath = "HKCU:\Software\"+$CompanyName 
 
if (Test-Path $CompanyRegPath) 
{} 
else 
{New-Item -path "HKCU:\Software" -name $CompanyName} 
 
if (Test-Path $CompanyRegPath'\Outlook Signature Settings') 
{} 
else 
{New-Item -path $CompanyRegPath -name "Outlook Signature Settings"} 
 
$SigVersion = (gci $RemoteSignaturePathFull).LastWriteTime #When was the last time the signature was written 
$ForcedSignatureNew = (Get-ItemProperty $CompanyRegPath'\Outlook Signature Settings').ForcedSignatureNew 
$ForcedSignatureReplyForward = (Get-ItemProperty $CompanyRegPath'\Outlook Signature Settings').ForcedSignatureReplyForward 
$SignatureVersion = (Get-ItemProperty $CompanyRegPath'\Outlook Signature Settings').SignatureVersion 
Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name SignatureSourceFiles -Value $SigSource 
$SignatureSourceFiles = (Get-ItemProperty $CompanyRegPath'\Outlook Signature Settings').SignatureSourceFiles 
 
#Forcing signature for new messages if enabled 
if ($ForcedSignatureNew -eq '1') 
{ 
#Set company signature as default for New messages 
$MSWord = New-Object -com word.application 
$EmailOptions = $MSWord.EmailOptions 
$EmailSignature = $EmailOptions.EmailSignature 
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries 
$EmailSignature.NewMessageSignature=$CompanyName 
$MSWord.Quit() 
} 
 
#Forcing signature for reply/forward messages if enabled 
if ($ForcedSignatureReplyForward -eq '1') 
{ 
#Set company signature as default for Reply/Forward messages 
$MSWord = New-Object -com word.application 
$EmailOptions = $MSWord.EmailOptions 
$EmailSignature = $EmailOptions.EmailSignature 
$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries 
$EmailSignature.ReplyMessageSignature=$CompanyName 
$MSWord.Quit() 
} 
 
#Copying signature sourcefiles and creating signature if signature-version are different from local version 
if ($SignatureVersion -ge $SigVersion -and $SignatureVersion -ge $ADLastChanged){} 
else 
{ 
#Copy signature templates from domain to local Signature-folder 
Copy-Item "$SignatureSourceFiles\*" $LocalSignaturePath -Recurse -Force 
 
$ReplaceAll = 2 
$FindContinue = 1 
$MatchCase = $False 
$MatchWholeWord = $True 
$MatchWildcards = $False 
$MatchSoundsLike = $False 
$MatchAllWordForms = $False 
$Forward = $True 
$Wrap = $FindContinue 
$Format = $False 
 
#Insert variables from Active Directory to rtf signature-file 
$MSWord = New-Object -com word.application 
$fullPath = $LocalSignaturePath+'\'+$CompanyName+'.docx' 
$MSWord.Documents.Open($fullPath) 
 
$FindText = "DisplayName" 
$ReplaceText = $ADDisplayName.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 
 
$FindText = "Title" 
$ReplaceText = $ADTitle.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 
 
$MSWord.Selection.Find.Execute("Email") 
$MSWord.ActiveDocument.Hyperlinks.Add($MSWord.Selection.Range, "mailto:"+$ADEmailAddress.ToString(), $missing, $missing, $ADEmailAddress.ToString()) 
$MSWord.Selection.Range.Font.Name = "Arial"
$MSWord.Selection.Range.Font.Size = 8

$FindText = "TelephoneNumber" 
$ReplaceText = $ADTelePhoneNumber.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 

$FindText = "MobileNumber" 
$ReplaceText = $ADMobileNumber.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 

$FindText = "FaxNumber" 
$ReplaceText = $ADFaxNumber.ToString() 
$MSWord.Selection.Find.Execute($FindText, $MatchCase, $MatchWholeWord,    $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap,    $Format, $ReplaceText, $ReplaceAll    ) 

 
$MSWord.ActiveDocument.Save() 
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatHTML"); 
[ref]$BrowserLevel = "microsoft.office.interop.word.WdBrowserLevel" -as [type] 
 
$MSWord.ActiveDocument.WebOptions.OrganizeInFolder = $true 
$MSWord.ActiveDocument.WebOptions.UseLongFileNames = $true 
$MSWord.ActiveDocument.WebOptions.BrowserLevel = $BrowserLevel::wdBrowserLevelMicrosoftInternetExplorer6 
$path = $LocalSignaturePath+'\'+$CompanyName+".htm" 
$MSWord.ActiveDocument.saveas([ref]$path, [ref]$saveFormat) 
 
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatRTF"); 
$path = $LocalSignaturePath+'\'+$CompanyName+".rtf" 
$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$saveFormat) 
 
$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatText"); 
$path = $LocalSignaturePath+'\'+$CompanyName+".rtf" 
$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$saveFormat) 
 
$path = $LocalSignaturePath+'\'+$CompanyName+".txt" 
$MSWord.ActiveDocument.SaveAs([ref] $path, [ref]$SaveFormat::wdFormatText) 
$MSWord.ActiveDocument.Close() 
 
$MSWord.Quit() 
 
} 
 
#Stamp registry-values for Outlook Signature Settings if they doesn`t match the initial script variables. Note that these will apply after the second script run when changes are made in the "Custom variables"-section. 
if ($ForcedSignatureNew -eq $ForceSignatureNew){} 
else 
{Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name ForcedSignatureNew -Value $ForceSignatureNew} 
 
if ($ForcedSignatureReplyForward -eq $ForceSignatureReplyForward){} 
else 
{Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name ForcedSignatureReplyForward -Value $ForceSignatureReplyForward} 
 
if ($SignatureVersion -ge $SigVersion -and $SignatureVersion -ge $ADLastChanged){} 
else 
{
    if($ADLastChanged -gt $SigVersion)
    {
        Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name SignatureVersion -Value $ADLastChanged
    }
    else
    {
        Set-ItemProperty $CompanyRegPath'\Outlook Signature Settings' -name SignatureVersion -Value $SigVersion
    }
}
# SIG # Begin signature block
# MIIEMwYJKoZIhvcNAQcCoIIEJDCCBCACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUhZamOmYbCiL8FLC8mlkH0K1A
# Xu6gggI9MIICOTCCAaagAwIBAgIQtNQJjOdtc6JNbgB9PO78DTAJBgUrDgMCHQUA
# MCwxKjAoBgNVBAMTIVBvd2VyU2hlbGwgTG9jYWwgQ2VydGlmaWNhdGUgUm9vdDAe
# Fw0xMTA1MDYwNTEyNTVaFw0zOTEyMzEyMzU5NTlaMBoxGDAWBgNVBAMTD1Bvd2Vy
# U2hlbGwgVXNlcjCBnzANBgkqhkiG9w0BAQEFAAOBjQAwgYkCgYEAv9JxpKUrk3bF
# 1bGMbCZ1yydHQH7F7Nrsqkj4c44W+ZQEjjoE2jlDZPwF4vZiyCkiGNUq9vFquIaO
# m4zgG+8fT2u+c9GIFUGnsEu41mGBlCEnq2fQ7izZ5KupWkvODzZqUKbF4OIak/84
# t3vIvjONKkVWTLW675W/YefM/GbYWncCAwEAAaN2MHQwEwYDVR0lBAwwCgYIKwYB
# BQUHAwMwXQYDVR0BBFYwVIAQlQYFX3dimX4F2Go5VoUCmaEuMCwxKjAoBgNVBAMT
# IVBvd2VyU2hlbGwgTG9jYWwgQ2VydGlmaWNhdGUgUm9vdIIQMlb92ZPxEIRBhO+x
# rg1EvTAJBgUrDgMCHQUAA4GBAD9TcynLBu6IKjw3HNnsTlnMZIjcgQyCq9me+/oo
# TpZxEDG/ea6n2G7GnnQnooW1erYTs3zkRhf3+N7VYmEIegWluEkHErx2BNTVpnfl
# 63Q66vtlpYhmsuIFQkpc/uM2+ns83Leq+IjKM7heUZN1PjDUqTCszOY9XGHW+bAO
# uCmFMYIBYDCCAVwCAQEwQDAsMSowKAYDVQQDEyFQb3dlclNoZWxsIExvY2FsIENl
# cnRpZmljYXRlIFJvb3QCELTUCYznbXOiTW4AfTzu/A0wCQYFKw4DAhoFAKB4MBgG
# CisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
# AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYE
# FBYuHFy0GoJSyAGCiZSnn/iWHDxeMA0GCSqGSIb3DQEBAQUABIGAVpBviu9fBb+0
# z+gRBie0uGHhYxVb2c7A2xcLoE5YQa8xDHOjFbV29j/ZE6lI/EQzuEQ37SL1De1b
# BLtvYXYfIOibTA3CoL2VSV84vpugDR141VKrBw9aJ11bQvgYqSxhaxXwmOrTY3pN
# Fw4ELsq1yQOmud7MTrSQwfFMLc7K3l4=
# SIG # End signature block
