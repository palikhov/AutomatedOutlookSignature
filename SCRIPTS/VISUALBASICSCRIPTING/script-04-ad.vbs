On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strName = objUser.FullName
strTitle = objUser.Title
strDepartment = objUser.Department
strCompany = objUser.Company
strDirectPhone = objUser.telephoneNumber
strEmail = objUser.mail
strAddress = objUser.streetAddress
strMobile = objUser.mobile
strSwitchPhone = objUser.otherTelephone
strSkype = objUser.ipPhone

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
objSelection.TypeText strName
objSelection.TypeParagraph()
objSelection.TypeText strTitle
objSelection.TypeParagraph()
objSelection.TypeText strCompany
objSelection.TypeParagraph()
objSelection.TypeText strAddress
objSelection.TypeParagraph()
objSelection.Font.Bold = True
objSelection.TypeText "T " 
objSelection.Font.Bold = False
objSelection.TypeText strSwitchPhone & " "
If Not IsEmpty(strDirectPhone) Then
	objSelection.Font.Bold = True
	objSelection.TypeText "D "
	objSelection.Font.Bold = False
	objSelection.TypeText strDirectPhone 
End If
objSelection.TypeParagraph()
If Not IsEmpty(strMobile) Then
	objSelection.Font.Bold = True
	objSelection.TypeText "M "
	objSelection.Font.Bold = False
	objSelection.TypeText strMobile & " "
End If
If Not IsEmpty(strSkype) Then
	objSelection.Font.Bold = True
	objSelection.TypeText "Skype "
	objSelection.Font.Bold = False
	objSelection.TypeText strSkype
End If
objSelection.TypeParagraph()
objSelection.Font.Bold = True
objSelection.TypeText "E "
objSelection.Font.Bold = False
objLink = objSelection.Hyperlinks.Add(objSelection.Range,"mailto:" & strEmail,,strEmail,strEmail)
objSelection.TypeParagraph()
objSelection.Font.Bold = True
objSelection.TypeText "W "
objSelection.Font.Bold = False
objLink = objSelection.Hyperlinks.Add(objSelection.Range,"http://www.company.com",,"Company Home Page","www.Company.com")
objSelection.TypeParagraph()
objSelection.TypeParagraph()
objSelection.Font.Bold = True
objSelection.TypeText "Tag Line goes Here"
Set objSelection = objDoc.Range()
objSelection.Font.Name = "Arial"
objSelection.Font.Size = 10

objSignatureEntries.Add strName, objSelection
objSignatureObject.NewMessageSignature = strName
objSignatureObject.ReplyMessageSignature = strName

objDoc.Saved = True
objWord.Quit
