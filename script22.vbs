On Error Resume Next

'Setting up the script to work with the file system.
Set WshShell = WScript.CreateObject("WScript.Shell")
Set FileSysObj = CreateObject("Scripting.FileSystemObject")
Set objSysInfo = CreateObject("ADSystemInfo")
Set UserObj = GetObject("LDAP://" & objSysInfo.UserName)
strAppData = WshShell.ExpandEnvironmentStrings("%APPDATA%")
SigFolder = StrAppData & "\Microsoft\Signatures\"
SigFile = SigFolder & UserObj.sAMAccountName & "2" & ".htm"

'Setting placeholders for the signature. They will be automatically replaced with data from Active Directory.
strUserName = UserObj.sAMAccountName
strFullName = UserObj.displayname
strTitle = UserObj.title
strMobile = UserObj.mobile
strEmail = UserObj.mail
strCompany = UserObj.company
strOfficePhone = UserObj.telephoneNumber

'Setting global placeholders for the signature. Those values will be identical for all users - make sure to replace them with the right values!
strCompanyLogo = "https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/hexagonal-logo/logo.png"
strBanner = "https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/hexagonal-logo/banner.png"
strBannerLinkingTo = "https://www.codetwo.com/email-signatures/"
strCompanyAddress = "16 Freedom St, Deer Hill 58-500 Poland"
strWebsite = "www.my-company.com"
strFacebook = "https://www.facebook.com/"
strTwitter = "https://www.twitter.com/"
strYouTube = ""
strLinkedIn = "https://www.linkedin.com/"
strInstagram = "https://www.instagram.com/"
strPinterest = ""

'Creating HTM signature file for the user's profile.
Set CreateSigFile = FileSysObj.CreateTextFile (SigFile, True, True)

'Signature’s HTML code
CreateSigFile.WriteLine "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>"
CreateSigFile.WriteLine "<HTML><HEAD><TITLE>Email Signature</TITLE>"
CreateSigFile.WriteLine "<META content='text/html; charset=utf-8' http-equiv='Content-Type'>"
CreateSigFile.WriteLine "</HEAD>"
CreateSigFile.WriteLine "<BODY style='font-size: 10pt; font-family: Arial, sans-serif;'>"
CreateSigFile.WriteLine "<table width='480' style='font-size: 11pt; font-family: Arial, sans-serif;' cellpadding='0' cellspacing='0' border='0'>"
CreateSigFile.WriteLine "<tbody>"
CreateSigFile.WriteLine "<tr>"
CreateSigFile.WriteLine "<td width='160' style='font-size: 10pt; font-family: Arial, sans-serif; width: 160px; vertical-align: top;' valign='top'> <a href='https://www.my-company.com/' target='_blank'><img border='0' alt='Logo' width='125' style='width:125px; height:auto; border:0;' src='" & strCompanyLogo & "'></a>"
CreateSigFile.WriteLine "</td>"
CreateSigFile.WriteLine "<td valign='top' width='270' style='width:270px; vertical-align: top; line-height:11px; border-right:2px solid #29abe1'><table cellpadding='0' cellspacing='0' border='0' width='270'><tbody> <tr> <td style='font-size:12pt; height:14px; line-height:14px'><strong style='font-family: Arial, sans-serif;font-size: 12pt;color:#29abe1;'>" & strFullName & "</strong></td> </tr> <tr> <td style='font-size:9pt; height:14px; line-height:14px'> <span style='font-family: Arial, sans-serif; font-size:9pt; color:#000000;'>" & strTitle & "</span> <span style='font-family: Arial, sans-serif; font-size:9pt; color:#000000;'> |" & strCompany & "</span> </td> </tr> <tr> <td style='height:14px; line-height:14px'>&nbsp;</td> </tr> <tr> <td style='font-size:9pt; height:14px; line-height:14px'> <span style='font-family: Arial, sans-serif;color:#000000;FONT-SIZE: 9pt'><strong>M</strong> " & strMobile & "</span> <span style='font-family: Arial, sans-serif;color:#000000;FONT-SIZE: 9pt'> | <strong>P</strong> " & strOfficePhone & "</span> </td> </tr> <tr> <td style='font-size:9pt; height:12px; line-height:12px'> <span style='font-family: Arial, sans-serif;color:#000000;FONT-SIZE: 9pt'><strong>E</strong> " & strEmail & "</span> </td> </tr> <tr> <td style='height:14px; line-height:14px'>&nbsp;</td> </tr> <tr> <td style='font-size:9pt; height:12px; line-height:12px'> <span style='font-family: Arial, sans-serif;color:#000000;FONT-SIZE: 9pt'>" & strCompanyAddress & "</span> </td> </tr> <tr> <td style='font-size:9pt; height:12px; line-height:12px'> <span><a href='https://" & strWebsite & "' target='_blank' rel='noopener' style=' text-decoration:none;'><strong style='color:#29abe1; font-family:Arial, sans-serif; font-size:9pt'>" & strWebsite & "</strong></a></span> </td> </tr> </tbody><tbody> </tbody></table>"
CreateSigFile.WriteLine "</td>"
CreateSigFile.WriteLine "<td style='vertical-align: top; padding-left:10px' valign='top' width='35'> <table cellpadding='0' cellspacing='0' border='0' width='25'> <tbody>"
If strFacebook <> "" Then
CreateSigFile.WriteLine "<tr> <td width='25' height='30' valign='top' style='vertical-align: top;'><a href='" & strFacebook & "' target='_blank' rel='noopener'><img border='0' width='26' height='25' alt='facebook icon' style='border:0; height:25px; width:26px;' src='https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/hexagonal-logo/fb.png'></a></td></tr>"
End If
If strTwitter <> "" Then
CreateSigFile.WriteLine "<tr> <td width='25' height='30' valign='top' style='vertical-align: top;'><a href='" & strTwitter & "' target='_blank' rel='noopener'><img border='0' width='26' height='25' alt='twitter icon' style='border:0; height:25px; width:26px;' src='https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/hexagonal-logo/tt.png'></a></td></tr>"
End If
If strYouTube <> "" Then
CreateSigFile.WriteLine "<tr> <td width='25' height='30' valign='top' style='vertical-align: top;'><a href='"& strYouTube &"' target='_blank' rel='noopene
r'><img border='0' width='26' height='25' alt='youtube icon' style='border:0; height:25px; width:26px' src='https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/hexagonal-logo/yt.png'></a></td> </tr>"
End If
CreateSigFile.WriteLine "</tbody></table>"
CreateSigFile.WriteLine "</td>"
CreateSigFile.WriteLine "<td style='vertical-align: top;' valign='top' width='25'>"
CreateSigFile.WriteLine "<table cellpadding='0' cellspacing='0' border='0' width='25'> <tbody>"
If strLinkedIn <> "" Then
CreateSigFile.WriteLine "<tr><td style='height:12px; font-size:1px' height='12'>&nbsp;</td></tr> <tr> <td width='25' height='30' valign='top' style='vertical-align: top;'><a href='" & strLinkedIn & "' target='_blank' rel='noopener'><img border='0' width='26' height='25' alt='linkedin icon' style='border:0; height:25px; width:26px;' src='https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/hexagonal-logo/ln.png'></a></td> </tr>"
End If
If strInstagram <> "" Then
CreateSigFile.WriteLine "<tr> <td width='25' height='30' valign='top' style='vertical-align: top;'><a href='" & strInstagram & "' target='_blank' rel='noopener'><img border='0' width='26' height='25' alt='instagram icon' style='border:0; height:25px; width:26px;' src='https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/hexagonal-logo/it.png'></a></td> </tr>"
End If
If strPinterest <> "" Then
CreateSigFile.WriteLine "<tr> <td width='25' height='30' valign='top' style='vertical-align: top;'><a href='" & strPinterest & "' target='_blank' rel='noopener'><img border='0' width='26' height='25' alt='pinterest icon' style='border:0; height:25px; width:26px' src='https://codetwocdn.azureedge.net/images/mail-signatures/generator-dm/hexagonal-logo/pt.png'></a></td> </tr>"
End If
CreateSigFile.WriteLine "</tbody></table>"
CreateSigFile.WriteLine "</td>"
CreateSigFile.WriteLine "</tr>"
CreateSigFile.WriteLine "<tr><td colspan='4' style='padding-top:15px;'> <a href='" & strBannerLinkingTo & "' target='_blank' rel='noopener'><img border='0' alt='Banner' width='479' style='max-width:479px; height:auto; border:0;' src='" & strBanner & "'></a> </td>"
CreateSigFile.WriteLine "</tr>"
CreateSigFile.WriteLine "<tr><td colspan='4' style='padding-top:15px; line-height:14px; font-size: 7.5pt; color: #808080; font-family: Arial, sans-serif;'>The content of this email is confidential and intended for the recipient specified in message only. It is strictly forbidden to share any part of this message with any third party, without a written consent of the sender. If you received this message by mistake, please reply to this message and follow with its deletion, so that we can ensure such a mistake does not occur in the future.</td>"
CreateSigFil
CreateSigFile.WriteLine "</tbody>"
CreateSigFile.WriteLine "</table>"
CreateSigFile.WriteLine "</BODY>"
CreateSigFile.WriteLine "</HTML>"
CreateSigFile.Close
Set objWord = CreateObject("Word.Application")
Set objSignatureObjects = objWord.EmailOptions.EmailSignature
objSignatureObjects.NewMessageSignature = strUserName & "2"
objSignatureObjects.ReplyMessageSignature = strUserName & "2"
objWord.Quit
