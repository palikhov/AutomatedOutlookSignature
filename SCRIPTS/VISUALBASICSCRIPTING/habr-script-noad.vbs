'запрос информации о компании
    Dim objIE
    Dim Brand 
    Set objIE = CreateObject( "InternetExplorer.Application" )
    objIE.Navigate "about:blank"
    objIE.Document.Title = "Запрос для подписи"
    objIE.ToolBar        = False
    objIE.Resizable      = False
    objIE.StatusBar      = False
    objIE.Width          = 600
    objIE.Height         = 250
    Do While objIE.Busy
        WScript.Sleep 200
    Loop
	objIE.Document.Body.InnerHTML = "<DIV align=""Left""><P>"&_
		"<datalist id=""rw""><option>Управляющий</option><option>Помощник</option></datalist>"&_
		"<datalist id=""obj""><option>Объект1</option><option>Объект2</option><option>Объект3</option></datalist>"&_
		"<input type='radio' name='RadioOption' value='1'>U
"&_
		"<input type='radio' name='RadioOption' value='2'>M
"&_
		"<input type='radio' name='RadioOption' value='3'>C<br>"&_
		"<input type='radio' name='RadioOption' value='4'>L<br>"&_
		"<input List='rw' name='Dol' >Должность<br>"&_
		"<input List='obj' name='objt' >Обьект<br>"&_
		"<input type='text' name='FIO' >Фамилия Имя<br>"&_
		"<input type='tel' name='tel' >Мобильный Телефон в формате +7(***)***-**-**<br>"&_
		"<input id='OK' type='hidden' value='0' name='OK'>"&_
		"<input type='submit' value='OK' onClick='VBScript:OK.Value=1'>"
    objIE.Visible = True
    Do While objIE.Document.All.OK.Value = 0
        WScript.Sleep 200
    Loop
    If objIE.Document.All.RadioOption(0).checked=true then Brand ="U"
    If objIE.Document.All.RadioOption(1).checked=true then Brand="M"
    If objIE.Document.All.RadioOption(2).checked=true then Brand="c"
    If objIE.Document.All.RadioOption(3).checked=true then Brand="L"
    
    If objIE.Document.All.RadioOption(0).checked=true then strCompany="U"
    If objIE.Document.All.RadioOption(1).checked=true then strCompany="M"
    If objIE.Document.All.RadioOption(2).checked=true then strCompany="C"
    If objIE.Document.All.RadioOption(3).checked=true then strCompany="L"

    If objIE.Document.All.RadioOption(3).checked=true then dolj= objIE.Document.All.Dol.Value+" кофейней" else dolj= objIE.Document.All.Dol.Value+" столовой"

    objtj= objIE.Document.All.objt.Value
    strMobile = objIE.Document.All.tel.Value
    strName = objIE.Document.All.FIO.Value
    objIE.Quit

    On Error Resume Next
strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)
    For Each objItem in colItems
        login_full =objItem.UserName
    Next
    Set objItem = Nothing: Set colItems = Nothing: Set objWMIService = Nothing
login_find = "\"
login_pos = InStr(1,login_full,login_find)
login_len = len(login_full)
login = right(login_full,login_len-login_pos)

folder_find = "C:\Users\"&login&"\AppData\Local\Microsoft\Outlook"

Set objShellApp = CreateObject("Shell.Application")
Set objFolder = objShellApp.NameSpace(folder_find)
Set objFolderItems = objFolder.Items()
objFolderItems.Filter 64+128, "*.ost"
For Each file in objFolderItems
    file_name = file
Next

file_len = len(file_name)
strEmail = left(file_name,file_len-4)

strZpov = "С уважением, "
strTitle = dolj+" "+objtj
strweb = "www.www.ru"
strLogo1 = "\\cabinet\Секретариат\Логотипы\"&Brand&"_logo_wl.jpg" 'основной логотип
strLogo3 = "\\cabinet\Секретариат\Логотипы\Ins.jpg" ' значек instagram
strLogo2 = "\\cabinet\Секретариат\Логотипы\F.JPG" ' значек facebook
strLogo4 = "\\cabinet\Секретариат\Логотипы\line.png" ' просто горизонтальная линия
strLogo5 = "\\cabinet\Секретариат\Логотипы\Save_wood.jpg" 'подпись о спасении леса

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
Set objRange = objDoc.Range()

'формируем табличку в которую будут подставлены нужные записи в соответствующие блоки.
'большая подпись представляет из себя табличку из 3 строчных блоков, 2 строка разделена на 2 ячейки

objDoc.Tables.Add objRange,3,2
Set objTable = objDoc.Tables(1)

objTable.Rows(1).select  ' строка 1, выделяем 
objSelection.Cells.Merge ' обьеденяем в единую строку во всю ширину таблички

objTable.Cell(1, 1).select ' выделяем 1 строку и задаем ей ширину
objTable.Cell(1, 1).Width = 605

objselection.font.name = "Cambria"
objSelection.Font.Size = "10"
objSelection.Font.Color = RGB(88,89,91)

 ' начинаем наполнять ячейку текстом о сотруднике ( ФИО, Должность, обьект, моб, почта)
' адрес почты делаем кликабельным для быстрой отправки письма mailto:
objSelection.TypeText strZpov & strName & CHR(11)
objSelection.TypeText strTitle & CHR(11)
objSelection.Font.Bold = true
objSelection.TypeText strCompany & CHR(11)
objSelection.Font.Bold = false 
objSelection.TypeText strMobile  & CHR(11)
hyp.Range.Font.name = "Cambria"
hyp.Range.Font.Size = "10"
hyp.Range.Font.Name = "Cambria"
Set hyp = objSelection.Hyperlinks.Add(objSelection.Range, "mailto: " & strEmail,,, strEmail)
hyp.Range.Font.Name = "Cambria"
hyp.Range.Font.Size = "10"

'строка 2, ячейка 1, вставляем в нее логотип компании
objTable.Cell(2, 1).select
objTable.Cell(2, 1).Width = 150
objTable.Cell(2, 1).Text = objSelection.InlineShapes.AddPicture(strLogo1)

' строка 2 ячейка 2, вставляем текст с адресом и данными о компании
objTable.Cell(2, 2).select  
objselection.font.name = "Cambria"
objSelection.Font.Size = "9,5"
objSelection.Font.Color = RGB(88,89,91)
objSelection.TypeText "111111, Москва, Кремль д1 с10, " & CHR(11)

' Ниже кусочек кода, изначально хотел узнавать у пользователя и телефон подразделения
' если он пустой, то вставляется общий телефон, если не пустой, то вставляется указанный номер
'так же если используется добавочный номер, то можно его запрашивать в переменную strintPhone

if (strPhone <> "") then objSelection.TypeText "БЦ «Кремль», " & strPhone else objSelection.TypeText "БЦ «Кремль», +7(111)111-11-11"
if (strIntPhone <> "") then objSelection.TypeText " доб. " & strIntPhone & CHR(11) else objSelection.TypeText CHR(11) 


Set hyp = objSelection.Hyperlinks.Add(objSelection.Range, strWeb,,, strWeb) 
hyp.Range.Font.Name = "Cambria"
hyp.Range.Font.Size = "9,5"
objSelection.TypeText CHR(9)

set p_f = objSelection.InlineShapes.AddPicture(strLogo2)
Set hyp = objSelection.HyperLinks.Add(p_f, "https://www.facebook.com/kremlin/",,,"Image")
hyp.Range.Font.Name = "Cambria"
hyp.Range.Font.Size = "9,5"
objSelection.TypeText " "
set p_i = objSelection.InlineShapes.AddPicture(strLogo3)
Set hyp = objSelection.HyperLinks.Add(p_i, "https://www.instagram.com/kremlin/",,,"Image")
hyp.Range.Font.Name = "Cambria"
hyp.Range.Font.Size = "9,5"




objselection.font.name = "Cambria"
objSelection.Font.Size = "9,5"
objSelection.Font.Color = RGB(88,89,91)

objSelection.TypeText " @kremlin"
objTable.Rows(3).select ' строка 3, обьединяем в единую ячейку и вставляем лого - береги бумагу
objSelection.Cells.Merge

objTable.Cell(3, 1).select
objTable.Cell(3, 1).Width = 605
objTable.Cell(3, 1).Text = objSelection.InlineShapes.AddPicture(strLogo5)


'''
'данный код формирует из документа подпись и подпихивает его в outlook для новых писем
Set objSelection = objDoc.Range()
objSignatureEntries.Add "AD Signature", objSelection
objSignatureObject.NewMessageSignature = "AD Signature"
objDoc.Saved = True
objDoc.Close
objWord.Quit

'''
' Ниже формируется краткая подпись с теми же данными только для писем ответов
' дабы не захламлять переписку огромными подписями в ответах на письма используется
' сокращенный шаблон

Set objWord = CreateObject("Word.Application")
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
Set objRange = objDoc.Range()

objselection.font.name = "Cambria"
objSelection.Font.Size = "10"
objSelection.Font.Color = RGB(88,89,91)

objSelection.TypeText strZpov & strName & CHR(11)
objSelection.TypeText strTitle & CHR(11)
if (strMobile <> "") then objSelection.TypeText strMobile & "   |    "

if (strPhone <> "") then objSelection.TypeText strPhone else objSelection.TypeText "+7(111)111-11-11"
if (strIntPhone <> "") then objSelection.TypeText " доб. " & strIntPhone & CHR(11) else objSelection.TypeText CHR(11) 


'''


Set objSelection = objDoc.Range()
objSignatureEntries.Add "Short_Signature", objSelection
objSignatureObject.ReplyMessageSignature = "Short_Signature"

objDoc.Saved = True
objDoc.Close
objWord.Quit