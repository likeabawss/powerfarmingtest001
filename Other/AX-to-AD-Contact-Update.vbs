
Call UpdateUserProperties("CN=mbtest,OU=IT,OU=New Zealand Wholesale,DC=powerfarming,DC=co,DC=nz", "testcompany", "testdept", 1234, "testgivenName", 1234, "test@test.com", 1234, 1234, 1234, 1234, 1234, "sn", 1234, "title", "www.test.com")

Sub UpdateUserProperties(dn, company, department, facsimileTelephoneNumber, _
							givenName, homePhone, mail, mobile, otherFacsimileTelephoneNumber, _ 
							otherHomePhone, otherMobile, otherTelephone, sn, telephoneNumber, title, url)

	Const ADS_PROPERTY_UPDATE = 2 
	Set objUser = GetObject("LDAP://" & dn)
	    '("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com") 
	 
	 
		objUser.Put "company", company
		objUser.Put "department", department
		objUser.Put "facsimileTelephoneNumber", facsimileTelephoneNumber
		objUser.Put "givenName", givenName
	 	objUser.Put "homePhone", homePhone
	 	objUser.Put "mail", mail
	 	objUser.Put "mobile", mobile
		objUser.PutEx ADS_PROPERTY_UPDATE, _
		    "otherFacsimileTelephoneNumber", Array(otherFacsimileTelephoneNumber)
		objUser.PutEx ADS_PROPERTY_UPDATE, _
		    "otherHomePhone", Array(otherHomePhone)		    
		objUser.PutEx ADS_PROPERTY_UPDATE, _
		    "otherMobile", Array(otherMobile)		    		    
		objUser.PutEx ADS_PROPERTY_UPDATE, _
		    "otherTelephone", Array(otherTelephone)		    
		objUser.Put "sn", sn
		objUser.Put "telephoneNumber", telephoneNumber
		objUser.Put "title", title
		objUser.Put "url", url
	 
	objUser.SetInfo

End Sub