Set objUser = CreateObject(“WScript.Network”)
userName = objUser.UserName
domainName = objUser.UserDomain

FUNCTION GetUserDN(BYVAL UN, BYVAL DN)
Set ObjTrans = CreateObject(“NameTranslate”)
objTrans.init 1, DN
objTrans.set 3, DN & “\” & UN
strUserDN = objTrans.Get(1)
GetUserDN = strUserDN
END FUNCTION

Set objLDAPUser = GetObject(“LDAP://” & GetUserDN(userName,domainName))

‘Getting prepared to write the files
Dim objFSO, objWsh, appDataPath, pathToCopyTo, plainTextFile, plainTextFilePath, htmlFile, htmlFilePath
Set objFSO = CreateObject(“Scripting.FileSystemObject”)
Set objWsh = CreateObject(“WScript.Shell”)
appDataPath = objWsh.ExpandEnvironmentStrings(“%APPDATA%”)
pathToCopyTo = appDataPath & “\Microsoft\Signatures\”

‘Plain text signature – signature.txt
plainTextFilePath = pathToCopyTo & “Signature.txt”
Set plainTextFile = objFSO.CreateTextFile(plainTextFilePath, TRUE)
plainTextFile.WriteLine(objLDAPUser.DisplayName)
plainTextFile.WriteLine(objLDAPUser.Description)
plainTextFile.WriteLine(objLDAPUser.telephoneNumber)
plainTextFile.WriteLine(“________________________________”)
plainTextFile.Close

‘HTML signature – signature.htm
htmlFilePath = pathToCopyTo & “signature.htm”
Set htmlFile = objFSO.CreateTextFile(htmlFilePath, TRUE)
htmlfile.WriteLine(“<body>”)
htmlfile.WriteLine(“<font face=””Arial”” COLOR=””#424242″” size=””2″”>”)
htmlfile.WriteLine(“<div><strong>” & objLDAPUser.DisplayName & “</strong></div>”)
htmlfile.WriteLine(“<div><strong>” & objLDAPUser.description & “</strong></div>”)
htmlfile.WriteLine(“<html>”)
htmlfile.WriteLine(“<font face=””Arial”” COLOR=””#424242″” size=””2″”>”)
htmlfile.WriteLine(“<div><strong> Email: <a href=mailto:” & objLDAPUser.mail &”><span style=’color:#008C98′>” & objLDAPUser.mail & “</span></a></trong></div>”)
htmlfile.WriteLine(“<div><strong>Tel: ” & objLDAPUser.telephoneNumber & “</strong></div>”)
htmlfile.WriteLine(“<HR WIDTH=””100%”” COLOR=””#088A85″” SIZE=””1″”>”)
htmlfile.WriteLine(“</body>”)
htmlfile.WriteLine(“</html>”)

‘RTF signature – Copies over pre-made RTF signature
Set fso = CreateObject(“Scripting.FileSystemObject”)
fso.CopyFile “\\fileserver\share\signature.rtf”, appDataPath & “\Microsoft\Signatures\”,TRUE