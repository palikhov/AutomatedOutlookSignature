On Error Resume Next

'Setting up the script to work with the file system.
Set WshShell = WScript.CreateObject("WScript.Shell")
Set FileSysObj = CreateObject("Scripting.FileSystemObject")

'Connecting to Active Directory to get user’s data.
Set objSysInfo = CreateObject("ADSystemInfo")
Set UserObj = GetObject("LDAP://" & objSysInfo.UserName)
strAppData = WshShell.ExpandEnvironmentStrings("%APPDATA%")
SigFolder = StrAppData & "\Microsoft\Signatures\"
SigFile = SigFolder & UserObj.sAMAccountName & ".htm"

'Setting placeholders for the signature.
strUserName = UserObj.sAMAccountName
strFullName = UserObj.displayname
strTitle = UserObj.title
strMobile = UserObj.mobile
strEmail = UserObj.mail
strCompany = UserObj.company
strOfficePhone = UserObj.telephoneNumber

'Setting global placeholders for the signature. Those values will be identical for all users - make sure to replace them with the right values!
strCompanyLogo = """https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/simplephoto-with-logo/logo.png"
strCompanyAddress1 = "16 Freedom St, Deer Hill"
strCompanyAddress2 = "58-500 Poland"
strWebsite = """https://www.my-company.com"
strFacebook = """https://www.facebook.com/"
strTwitter = """https://www.twitter.com/"
strYouTube = """https://www.youtube.com/"
strLinkedIn = """https://www.linkedin.com/"
strInstagram = """https://www.instagram.com/"
strPinterest = """https://www.pinterest.com/"

'Creating HTM signature file for the user's profile, if the file with such a name is found, it will be overwritten.
Set CreateSigFile = FileSysObj.CreateTextFile (SigFile, True, True)

'Signature’s HTML code.
CreateSigFile.WriteLine "<!DOCTYPE HTML PUBLIC " & """-//W3C//DTD HTML 4.0 Transitional//EN" & """>"
CreateSigFile.WriteLine "<HTML><HEAD><TITLE>Email Signature</TITLE>"
CreateSigFile.WriteLine "<META content=" & """text/html; charset=utf-8" & """ http-equiv=" & """Content-Type" & """>"
CreateSigFile.WriteLine "</HEAD>"
CreateSigFile.WriteLine "<BODY style=" & """font-size: 10pt; font-family: Arial, sans-serif;" & """>"
CreateSigFile.Writeline "<table style=" & """width: 420px; font-size: 10pt; font-family: Arial, sans-serif;" & """ cellpadding=" & """0" & """ cellspacing=" & """0" & """>"
CreateSigFile.Writeline "<tbody>"
CreateSigFile.Writeline "<tr>"
CreateSigFile.Writeline "<td width=" & """130" & """ style=" & """font-size: 10pt; font-family: Arial, sans-serif; border-right: 1px solid; border-right-color: #008080; width: 130px; padding-right: 10px; vertical-align: top;" & """ valign=" & """top" & """ rowspan=" & """6" & """> <a href=" & strWebsite & """ target=" & """_blank" & """><img border=" & """0" & """ alt=" & """Logo" & """ width=" & """110" & """ style=" & """width:110px; height:auto; border:0;" & """ src=" & strCompanyLogo & """></a>"
CreateSigFile.Writeline "</td>"
CreateSigFile.Writeline "<td>"
CreateSigFile.Writeline "<table cellpadding="& """0"& """ cellspacing="& """0"& """>"
CreateSigFile.Writeline "<tbody>"
CreateSigFile.Writeline "<tr>"
CreateSigFile.Writeline "<td style="& """font-size: 10pt; color:#0079ac; font-family: Arial, sans-serif; width: 305px; padding-bottom: 5px; padding-left: 10px; vertical-align: top; line-height:25px;"& """ valign="& """top"& """>"
CreateSigFile.Writeline "<strong><span style="& """font-size: 14pt; font-family: Arial, sans-serif; color:#008080;"& """>" & strFullName & "<br></span></strong>"
CreateSigFile.Writeline "<span style="& """font-family: Arial, sans-serif; font-size:10pt; color:#545454;"& """>" & strTitle & "</span>"
CreateSigFile.Writeline "<span style="& """font-family: Arial, sans-serif; font-size:10pt; color:#545454;"& """> | </span>"
CreateSigFile.Writeline "<span style="& """font-family: Arial, sans-serif; font-size:10pt; color:#545454;"& """>" & strCompany & "</span>"
CreateSigFile.Writeline "</td>"
CreateSigFile.Writeline "</tr>"
CreateSigFile.Writeline "<tr>"
CreateSigFile.Writeline "<td style="& """font-size: 10pt; color:#444444; font-family: Arial, sans-serif; padding-bottom: 5px; padding-top: 5px; padding-left: 10px; vertical-align: top;"& """ valign="& """top"& """>"
CreateSigFile.Writeline "<span><span style="& """color: #008080;"& """><strong>m:</strong></span><span style="& """font-size: 10pt; font-family: Arial, sans-serif; color:#545454;"& """>" & strMobile & "<br></span></span>"
CreateSigFile.Writeline "<span><span style="& """color: #008080;"& """><strong>p:</strong></span><span style="& """font-size: 10pt; font-family: Arial, sans-serif; color:#545454;"& """>" & strOfficePhone & "<br></span></span>"
CreateSigFile.Writeline "<span><span style="& """color: #008080;"& """><strong>e:</strong></span><span style="& """font-size: 10pt; font-family: Arial, sans-serif; color:#545454;"& """>" & strEmail& "</span></span>"
CreateSigFile.Writeline "</td>"
CreateSigFile.Writeline "</tr>"
CreateSigFile.Writeline "<tr>"
CreateSigFile.Writeline "<td style="& """font-size: 10pt; font-family: Arial, sans-serif; padding-bottom: 5px; padding-top: 5px; padding-left: 10px; vertical-align: top; color: #0079ac;"& """ valign="& """top"& """>"
CreateSigFile.Writeline "<span style="& """font-size: 10pt; font-family: Arial, sans-serif; color: #008080;"& """>" & strCompanyAddress1 & "<span><br></span></span>"
CreateSigFile.Writeline "<span style="& """font-size: 10pt; font-family: Arial, sans-serif; color: #008080;"& """>" & strCompanyAddress2 &"</span>"
CreateSigFile.Writeline "</td>"
CreateSigFile.Writeline "</tr>"
CreateSigFile.Writeline "<tr>"
CreateSigFile.Writeline "<td style="& """font-size: 10pt; font-family: Arial, sans-serif; padding-bottom: 5px; padding-top: 5px; padding-left: 10px; vertical-align: top; color: #0079ac;"& """ valign="& """top"& """>"
CreateSigFile.Writeline "<a href="& """https://www.my-company.com"& """ target="& """_blank"& """ rel="& """noopener"& """ style="& """text-decoration:none;"& """><span style="& """font-size: 10pt; font-family: Arial, sans-serif; color: #008080;"& """>www.my-company.com</span></a>"
CreateSigFile.Writeline "</td>"
CreateSigFile.Writeline "</tr>"
CreateSigFile.Writeline "<tr>"
CreateSigFile.Writeline "<td style="& """font-size: 10pt; font-family: Arial, sans-serif; padding-bottom: 5px; padding-top: 5px; padding-left: 10px; vertical-align: top;"& """ valign="& """top"& """>"
CreateSigFile.Writeline "<span><a href="& strFacebook & """ target="& """_blank"& """ rel="& """noopener"& """><img border="& """0"& """ width="& """21"& """ alt="& """facebook icon"& """ style="& """border:0; height:21px; width:21px"& """ src="& """https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/simplephoto-with-logo/fb.png"& """></a>&nbsp;</span><span><a href="& strTwitter & """ target="& """_blank"& """ rel="& """noopener"& """><img border="& """0"& """ width="& """21"& """ alt="& """twitter icon"& """ style="& """border:0; height:21px; width:21px"& """ src="& """https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/simplephoto-with-logo/tt.png"& """></a>&nbsp;</span><span><a href="& strYouTube & """ target="& """_blank"& """ rel="& """noopener"& """><img border="& """0"& """ width="& """21"& """ alt="& """youtube icon"& """ style="& """border:0; height:21px; width:21px"& """ src="& """https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/simplephoto-with-logo/yt.png"& """></a>&nbsp;</span><span><a href="& strLinkedIn & """ target="& """_blank"& """ rel="& """noopener"& """><img border="& """0"& """ width="& """21"& """ alt="& """linkedin icon"&
""" style="& """border:0; height:21px; width:21px"& """ src="& """https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/simplephoto-with-logo/ln.png"& """></a>&nbsp;</span><span><a href="& strInstagram & """ target="& """_blank"& """ rel="& """noopener"& """><img border="& """0"& """ width="& """21"& """ alt="& """instagram icon"& """ style="& """border:0; height:21px; width:21px"& """ src="& """https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/simplephoto-with-logo/it.png"& """></a>&nbsp;</span><span><a href="& strPinterest & """ target="& """_blank"& """ rel="& """noopener"& """><img border="& """0"& """ width="& """21"& """ alt="& """pinterest icon"& """ style="& """border:0; height:21px; width:21px"& """ src="& """https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/simplephoto-with-logo/pt.png"& """></a></span>"
CreateSigFile.Writeline "</td>"
CreateSigFile.Writeline "</tr>"
CreateSigFile.Close

'Applying the signature in Outlook’s settings.
Set objWord = CreateObject("Word.Application")
Set objSignatureObjects = objWord.EmailOptions.EmailSignature

'Setting the signature as default for new messages.
objSignatureObjects.NewMessageSignature = strUserName & "1"

'Setting the signature as default for replies & forwards.
objSignatureObjects.ReplyMessageSignature = strUserName & "1"
objWord.Quit
