'here are numerous errors and mistakes in you script.  Don't ever use "On Error Resume Next" unless you plane to check each startemnt for an error.  Once that was removed most of the errors became obvious.

'Here is how to generate a signature the easy way.

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
'This loads a template from the network that has the format, images and links you want.  It then just assigns the bookmarks by name and quits.

'This version should work easily in a logon script.   You should create a way to prevent it from running every time the user logs on.  The easy way to do that is to check the time on the template file and save it in the registry. If the time is newer then regenerate the signature.

'If you want I can give you a link to the template that I used for testing.  I created it from your original code nut even the Word doc created had numerous issues and would not work well in various word versions.

'\_(ツ)_/