Public Sub SaveQuotes(olItem As MailItem) ' CODE 1

Dim oMail As Outlook.MailItem ' CODE 2
Dim objItem As Object ' CODE 2
Dim dtDate As Date ' CODE 2
Dim sName As String ' CODE 2
Dim SenderName As String ' CODE 2
Dim enviro As String ' CODE 2

Dim olAttach As Attachment ' CODE 1
Dim strFname As String ' CODE 1
Dim strExt As String ' CODE 1
Dim ProjectNumber As String ' CODE 1
Dim Description As String ' CODE 1
Dim FolderName As String ' CODE 1
Dim sFolderName As String, sFolder As String ' CODE 1
Dim strSaveFlder2 As String ' CODE 1
Dim j As Long ' CODE 1
Const strSaveFldr As String = "X:\Shared Files\Quotes\" 'CODE 1 Folder must exist

enviro = CStr(Environ("USERPROFILE")) ' CODE 2
   For Each objItem In ActiveExplorer.Selection ' CODE 2
   If objItem.MessageClass = "IPM.Note" Then ' CODE 2
    Set oMail = objItem ' CODE 2
SenderName = oMail.SenderName ' CODE 2
  sName = oMail.Subject ' CODE 2
  ReplaceCharsForFileName sName, "-" ' CODE 2
 
  dtDate = oMail.ReceivedTime ' CODE 2
  sName = SenderName & " - " & sName & ".msg" ' CODE 2

sFolderName = Format(dtDate, "yyyy.MM.dd") ' CODE 1

ProjectNumber = InputBox("What is the Project Number?") 'CODE 1 Asks the project number

Description = InputBox("What are these attachments?") 'CODE 1 Asks for a description

FolderName = sFolderName & " - " & ProjectNumber & " " & Description ' CODE 1

strSaveFlder2 = strSaveFldr & FolderName & "\" ' CODE 1

MkDir (strSaveFlder2) ' CODE 1

     
  Debug.Print strSaveFlder2 & sName ' CODE 2
  oMail.SaveAs strSaveFlder2 & sName, olMsg ' CODE 2
  End If ' CODE 2
  Next ' CODE 2

On Error GoTo lbl_Exit ' CODE 1
If olItem.Attachments.Count > 0 Then ' CODE 1
For j = 1 To olItem.Attachments.Count ' CODE 1
Set olAttach = olItem.Attachments(j) ' CODE 1
If Not olAttach.FileName Like "image*.*" Then ' CODE 1
strFname = olAttach.FileName ' CODE 1
strExt = Right(strFname, Len(strFname) - InStrRev(strFname, Chr(46))) ' CODE 1
strFname = FileNameUnique(strSaveFldr, strFname, strExt) ' CODE 1
olAttach.SaveAsFile strSaveFlder2 & strFname ' CODE 1
End If ' CODE 1
Next j ' CODE 1
olItem.Save ' CODE 1
End If ' CODE 1
lbl_Exit: ' CODE 1
Set olAttach = Nothing ' CODE 1
Set olItem = Nothing ' CODE 1

Exit Sub
End Sub

Public Sub SaveEmails(olItem As MailItem) ' CODE 1

Dim oMail As Outlook.MailItem ' CODE 2
Dim objItem As Object ' CODE 2
Dim dtDate As Date ' CODE 2
Dim sName As String ' CODE 2
Dim SenderName As String ' CODE 2
Dim enviro As String ' CODE 2

Dim olAttach As Attachment ' CODE 1
Dim strFname As String ' CODE 1
Dim strExt As String ' CODE 1
Dim ProjectNumber As String ' CODE 1
Dim Description As String ' CODE 1
Dim FolderName As String ' CODE 1
Dim sFolderName As String, sFolder As String ' CODE 1
Dim strSaveFlder2 As String ' CODE 1
Dim j As Long ' CODE 1
Const strSaveFldr As String = "X:\Shared Files\Emails\" 'CODE 1 Folder must exist

enviro = CStr(Environ("USERPROFILE")) ' CODE 2
   For Each objItem In ActiveExplorer.Selection ' CODE 2
   If objItem.MessageClass = "IPM.Note" Then ' CODE 2
    Set oMail = objItem ' CODE 2
SenderName = oMail.SenderName ' CODE 2
  sName = oMail.Subject ' CODE 2
  ReplaceCharsForFileName sName, "-" ' CODE 2
 
  dtDate = oMail.ReceivedTime ' CODE 2
  sName = SenderName & " - " & sName & ".msg" ' CODE 2

sFolderName = Format(dtDate, "yyyy.MM.dd") ' CODE 1

ProjectNumber = InputBox("What is the Project Number?") 'CODE 1 Asks the project number

Description = InputBox("What are these attachments?") 'CODE 1 Asks for a description

FolderName = sFolderName & " - " & ProjectNumber & " " & Description ' CODE 1

strSaveFlder2 = strSaveFldr & FolderName & "\" ' CODE 1

MkDir (strSaveFlder2) ' CODE 1

     
  Debug.Print strSaveFlder2 & sName ' CODE 2
  oMail.SaveAs strSaveFlder2 & sName, olMsg ' CODE 2
  End If ' CODE 2
  Next ' CODE 2

On Error GoTo lbl_Exit ' CODE 1
If olItem.Attachments.Count > 0 Then ' CODE 1
For j = 1 To olItem.Attachments.Count ' CODE 1
Set olAttach = olItem.Attachments(j) ' CODE 1
If Not olAttach.FileName Like "image*.*" Then ' CODE 1
strFname = olAttach.FileName ' CODE 1
strExt = Right(strFname, Len(strFname) - InStrRev(strFname, Chr(46))) ' CODE 1
strFname = FileNameUnique(strSaveFldr, strFname, strExt) ' CODE 1
olAttach.SaveAsFile strSaveFlder2 & strFname ' CODE 1
End If ' CODE 1
Next j ' CODE 1
olItem.Save ' CODE 1
End If ' CODE 1
lbl_Exit: ' CODE 1
Set olAttach = Nothing ' CODE 1
Set olItem = Nothing ' CODE 1

Exit Sub
End Sub

'Private Function for the Save Attachment Code
Private Function FileNameUnique(strPath As String, _
strFileName As String, _
strExtension As String) As String
Dim lngF As Long
Dim lngName As Long
lngF = 1
lngName = Len(strFileName) - (Len(strExtension) + 1)
strFileName = Left(strFileName, lngName)
Do While FileExists(strPath & strFileName & Chr(46) & strExtension) = True
strFileName = Left(strFileName, lngName) & "(" & lngF & ")"
lngF = lngF + 1
Loop
FileNameUnique = strFileName & Chr(46) & strExtension
lbl_Exit:
Exit Function
End Function

'Private Function for the Save Attachment Code
Private Function FileExists(filespec) As Boolean
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(filespec) Then
FileExists = True
Else
FileExists = False
End If
lbl_Exit:
Exit Function
End Function

'Private Function for the Save Attachment Code
Sub SaveQuote()
Dim olMsg As MailItem
On Error Resume Next
Set olMsg = ActiveExplorer.Selection.Item(1)
SaveQuotes olMsg
lbl_Exit:
Exit Sub
End Sub

'Private Function for the Save Attachment Code
Sub SaveEmail()
Dim olMsg As MailItem
On Error Resume Next
Set olMsg = ActiveExplorer.Selection.Item(1)
SaveEmails olMsg
lbl_Exit:
Exit Sub
End Sub

'Private Function for the Save Message Code
Private Sub ReplaceCharsForFileName(sName As String, _
  sChr As String _
)
  sName = Replace(sName, "'", sChr)
  sName = Replace(sName, "*", sChr)
  sName = Replace(sName, "/", sChr)
  sName = Replace(sName, "\", sChr)
  sName = Replace(sName, ":", sChr)
  sName = Replace(sName, "?", sChr)
  sName = Replace(sName, Chr(34), sChr)
  sName = Replace(sName, "<", sChr)
  sName = Replace(sName, ">", sChr)
  sName = Replace(sName, "|", sChr)
End Sub










