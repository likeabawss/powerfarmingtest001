function Import-509Certificate {    
  
   param([String]$certPath,[String]$certRootStore,[String]$certStore)    
  
   $pfx = new-object System.Security.Cryptography.X509Certificates.X509Certificate2    
   $pfx.import($certPath)    
  
   $store = new-object System.Security.Cryptography.X509Certificates.X509Store($certStore,$certRootStore)   
   $store.open("MaxAllowed")    
   $store.add($pfx)    
   $store.close()    
}   

Import-509Certificate "\\power\netlogon\svn-netlogon\PS_Powerfarming.cer" "CurrentUser" "My"
# SIG # Begin signature block
# MIIEMwYJKoZIhvcNAQcCoIIEJDCCBCACAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUkk3/mKK/jYR5/XLBUY0naqhf
# KOqgggI9MIICOTCCAaagAwIBAgIQtNQJjOdtc6JNbgB9PO78DTAJBgUrDgMCHQUA
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
# FG54Hw7fuxJP7vsUzXgWYV9xbUX3MA0GCSqGSIb3DQEBAQUABIGAkmsl6F1STaUN
# snVes7HXQhG9P+mzDGfnXJil5mqHIevMSlYIxcw6x9nmPeWvssUcvRiBbNKtnUFT
# zWiS4SY7RRS2Kiw3XskqcDWbaVRzHc4wIwzfb+zCSCggAZH+L39FOnnNYvisa72R
# uH9WqzMklt/RMvFIve1cY1acNerWyX8=
# SIG # End signature block
