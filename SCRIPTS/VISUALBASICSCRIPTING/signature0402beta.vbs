On Error Resume Next
		Function orange(objSelection)
			With objSelection
				With .Font
					.Name = "Arial"
					.Size = 10
					.Bold = False
					.Color = RGB(31,73,125)
					.Italic = False
					.Underline = False
		
				End With
			End With
	
		End Function
	
		Function black(objSelection)
			With objSelection
				With .Font
						.Name = "Arial"
						.Size = 10
						.Bold = True
						.Color = RGB(0,0,0)
						.Italic = False
						.Underline = False
				End With
			End With
			
		End Function
Function red(objSelection)
			With objSelection
				With .Font
						.Name = "Arial"
						.Size = 7.5
						.Bold = True
						.Color = RGB(204,0,0)
						.Italic = False
						.Underline = False
				End With
			End With
			
		End Function
	Set objSysInfo = CreateObject("ADSystemInfo")
		strUser = objSysInfo.UserName
	Set objUser = GetObject("LDAP://" & strUser)

	With objUser
  
		strName = .FullName
		strTitle = .Title
strDepartment = .Department 
		stradr = .streetAddress
		strpostal = .postalCode
		strl = .l
		strco = .co
		strcomp = .company
		strhome = .homePhone
		strIpPhone = .ipPhone
		strMobile = .Mobile
		strPhone = .TelephoneNumber
		strMail = .mail
	

	End With


Set objword = CreateObject("Word.Application")
	With objword
		Set objDoc = .Documents.Add()
		Set objSelection = .Selection
		Set objEmailOptions = .EmailOptions
	End With

Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
With objSelection



		.ParagraphFormat.Alignment = wdAlignParagraphLeft
		.ParagraphFormat.SpaceAfter = 0
		.ParagraphFormat.SpaceBefore = 0
		
		
		'NAME
		 With .Font
				.Name = "Arial"
				.Size = 10
				.Bold = True
				.Color = RGB(31,73,125)
				.Italic = False
				.Underline = False
		End With
		.TypeText "З повагою, "&strName 
		.TypeText Chr(11) & Chr(6)
		
		
		'JOB TITLE
		
		orange(objSelection)
		.TypeText strTitle & Chr(32) & Chr(11)
		 
		orange(objSelection)
		.TypeText strDepartment & Chr(32) & Chr(11)
		
		'main
		
		
		
		'mobile
		
		If (strMobile = Empty) Then
			.TypeText  Chr(32)
		Else
			orange(objSelection)
			.TypeText ("Моб.тел: ")
			'Mobilephone
			orange(objSelection)
			.TypeText   Chr(32)& strMobile & Chr(11) 
		End If
		'General Phone
		If (strhome <> Empty) Then
			orange(objSelection)
			.TypeText ("Телефон: ") & Chr(32)& Chr(32) 
			orange(objSelection)
			.TypeText  strhome & Chr(32)& Chr(32)
		
		End If
		
		'Direct Phone
		If (strIpPhone <> Empty) Then
			orange(objSelection)
			.TypeText ("Direct: ") & Chr(32)& Chr(32) 
			orange(objSelection)
			.TypeText  strIpPhone 
		End If
		'Email
		orange(objSelection)
		.TypeText ("Email: ")
		orange(objSelection)
		.TypeText  strMail & Chr(32) & Chr(11) & Chr(6) 
		.TypeText ("www.ideabank.ua")& Chr(11) & Chr(11)
		
'.InlineShapes.AddPicture "https://github.com/palikhov/AutomatedOutlookSignature/raw/master/logo.jpg", True, True		
.InlineShapes.AddPicture "https://github.com/palikhov/AutomatedOutlookSignature/raw/master/logo.jpg", True, True
		.TypeText Chr(11) & Chr(11)
red(objSelection)
	.TypeText  "Інформація, яку містить це повідомлення, включаючи будь-які вкладення, є конфіденційною. Якщо Ви не є адресатом, прохання не копіювати, не використовувати і не розголошувати цю інформацію. Попередьте відправника, відповівши на це повідомлення, а потім видаліть отриманий лист з Вашої системи. "	
	'SOCIAL BANNER
	
		'.TypeText   Chr(32)& Chr(32) & Chr(11)	& Chr(11) & Chr(11)
		
		'.InlineShapes.AddPicture "C:\Users\APalikhov\Pictures\icon.png", True, True
		'.TypeText Chr (9)
		'.InlineShapes.AddPicture "C:\Users\APalikhov\Pictures\icon.png", True, True
		'.TypeText Chr (9)
		'.InlineShapes.AddPicture "C:\Users\APalikhov\Pictures\icon.png", True, True
		'.TypeText Chr (9)
		'.InlineShapes.AddPicture "C:\Users\APalikhov\Pictures\icon.png", True, True
		'.TypeText Chr (9)
		
End With

	'objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(1),"https://www.linkedin.com/in/avpalikhov"
	'objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(2),"mailto:palikhov@outlook.com"
	'objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(3),"https://palikhov.wordpress.com"
	'objDoc.Hyperlinks.Add objDoc.InlineShapes.Item(4),"https://palikhov.wordpress.com"
	
Set objSelection = objDoc.Range()

objSignatureEntries.Add "AD Signature", objSelection
objSignatureObject.NewMessageSignature = "AD Signature"
objSignatureObject.ReplyMessageSignature = "AD Signature"
objDoc.Saved = True
objword.Quit