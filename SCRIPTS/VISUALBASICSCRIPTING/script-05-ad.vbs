strTEmplate = "\\server\documents\sig1.dotx"

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add(strTemplate)

Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
With objDoc.Bookmarks
	.Item("FullName").Range.Text = objUser.FullName
	.Item("Title").Range.Text = objUser.Title
	.Item("streetaddress").Range.Text = objUser.streetaddress
	.Item("telephoneNumber").Range.Text = objUser.telephoneNumber
	If objUser.mobile.Length <> 0  Then
		.Item("Mobile").Range.Text = objUser.mobile
	Else
		.Item("Mobile").Range.Delete()
	End If
	.Item("facsimileTelephoneNumber").Range.Text = objUser.facsimileTelephoneNumber
	.Item("mail").Range.Text = objUser.mail
End With

Set selection = objDoc.Range()
With objWord.EmailOptions.EmailSignature
	.EmailSignatureEntries.Add "Celtrino AD Signature", selection
	.NewMessageSignature = "Celtrino AD Signature"
	.ReplyMessageSignature = "Celtrino AD Signature"
End With

objDoc.Saved = True
objWord.Quit