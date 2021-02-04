Public Sub CreateMessageSignature()
Dim objMsg As MailItem
Dim strBuffer As String
enviro = CStr(Environ("appdata"))
Debug.Print enviro
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Edit the signature file name on the following line
strSigFilePath = enviro & "\Microsoft\Signatures\"
Debug.Print strSigFilePath
Set objSignatureFile = objFSO.OpenTextFile(strSigFilePath & "Palikhov eng.htm")
strBuffer = objSignatureFile.ReadAll
objSignatureFile.Close

Set objMsg = Application.CreateItem(olMailItem)
With objMsg
  .Subject = "Subject goes here"
  .HTMLBody = "<p>Something here.</p><p> </p>" & strBuffer
  .Display
End With
  
End Sub
