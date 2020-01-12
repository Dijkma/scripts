Set objSysInfo = CreateObject("ADSystemInfo")

strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)

strName = objUser.FirstName & " " & objUser.lastName
strTitle = objUser.Title
strDepartment = objUser.Department
strCompany = objUser.Company
strPhone = objUser.telephoneNumber
strMobile = objUser.mobile
strEmail = objUser.mail
strLogo = "\\dijkma-dc01\NETLOGON\logo\logo.jpg"

Set objWord = CreateObject("Word.Application")

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

Set objEmailOptions = objWord.EmailOptions
Set objSignatureObject = objEmailOptions.EmailSignature

Set objSignatureEntries = objSignatureObject.EmailSignatureEntries

' Beginning of signature block

objSelection.Font.Color = RGB(89,89,89)
objSelection.Font.Name = "Calibri"
objSelection.Font.Size = 11

objSelection.TypeText "Met vriendelijke groet,"
objSelection.TypeParagraph()
objSelection.TypeText strName & Chr(11) & Chr(11)
objSelection.InlineShapes.AddPicture(strLogo)
objSelection.TypeText Chr(11)
objSelection.TypeText "Dijkma Electronics B.V." & Chr(11)
objSelection.TypeText "Hoofdstraat 58" & Chr(11)
objSelection.TypeText "3781 AH  Voorthuizen" & Chr(11)

if IsEmpty(strPhone) = false Then
objSelection.TypeText "T:            " & strPhone & Chr(11)
End if

if IsEmpty(strMobile) = false Then
objSelection.TypeText "M:           " & strMobile & Chr(11)
End if

objSelection.TypeText "E:            "
objDoc.Hyperlinks.Add objSelection.Range, "mailto:" & strEmail,,,strEmail
objSelection.TypeText Chr(11)
objSelection.TypeText "W:          "
objDoc.Hyperlinks.Add objSelection.Range, "https://www.dijkma.nl",,,"www.dijkma.nl"
objSelection.TypeText Chr(11) & Chr(11)
 
if isMember("CN=winkels") Then
objSelection.TypeText "Nu ook te bereiken via WhatsApp"   & Chr(11)
objSelection.TypeText "Voeg ons nummer toe aan uw contacten"   & Chr(11)
objSelection.TypeText "0342-474025 en start een gesprek met ons."   & Chr(11)
end if
' End of signature block

Set objSelection = objDoc.Range()

objSignatureEntries.Add "AD Signature", objSelection
objSignatureObject.NewMessageSignature = "AD Signature"
objSignatureObject.ReplyMessageSignature = "AD Signature"

objDoc.Saved = True
objWord.Quit

Function IsMember(groupName)
   Set groupListD = CreateObject("Scripting.Dictionary")
   groupListD.CompareMode = 1
   For Each objGroup in objUser.Groups
      groupListD.Add objGroup.Name, "-"
   Next
   IsMember = CBool(groupListD.Exists(groupName))
End Function
