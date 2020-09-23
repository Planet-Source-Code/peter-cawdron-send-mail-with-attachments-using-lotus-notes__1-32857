<div align="center">

## Send mail with Attachments using Lotus Notes


</div>

### Description

apidude posted this last week, I've added the ability to send attachments as well and resubmitted it. Personally, I've been looking for some code like this for a long, long time... Thanks Apidude....

The idea application for this is to build a bulk email program/Access DB that allows bulk email to be sent with each one personalised or carrying information specific to an individual. ie. Sending out customer statements by email, etc...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Peter Cawdron](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/peter-cawdron.md)
**Level**          |Intermediate
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Access, VBA MS Excel
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/peter-cawdron-send-mail-with-attachments-using-lotus-notes__1-32857/archive/master.zip)





### Source Code

```
'**************************************
' Name: Use Lotus Notes to send email
' Description:Creates a Lotus Notes sess
'   ion and use it to send an email
' By: apidude
'   attachments added by pcawdron
'
' Inputs:strMessage: The message
'strSubject: the subject
'strSendTo: the recipient 's email address
'lngLogo:Specifies the letter head To use (Lotus Notes specific)
'
' Assumes:The Font & Color values for th
'   e NotesRichTextItem class I'm not too su
'   re of because I don't have the DevKit or
'   the headers
'
'This code is copyrighted and has' limited warranties.Please see http://w
'   ww.Planet-Source-Code.com/xq/ASP/txtCode
'   Id.32603/lngWId.1/qx/vb/scripts/ShowCode
'   .htm'for details.'**************************************
Function SendNotesMail(strMessage As String, _
  strSubject As String, _
  strSendTo As String, _
  lngLogo As Long, strAttachment As String)
  On Error GoTo NotesMail_Err
  Dim lnSession As Object
  Dim lnDatabase As Object
  Dim lnDocument As Object
  Dim lnRTStyle As Object
  Dim lRTItem As Object
  Dim lnATTACHMENT As Object
  Dim sMessage As String
  Dim lLogo As Long
  ''start a notes session...
  Set lnSession = CreateObject("Notes.Notessession")
  ''create a new style object to control t
  '   he appearance of the message
  Set lnRTStyle = lnSession.CreateRichTextStyle
  ''get the current database...
  Set lnDatabase = lnSession.GetDatabase("", "")
  lnDatabase.OpenMail
  ''create a new document
  Set lnDocument = lnDatabase.CreateDocument
  ''create a new NotesRichTextItem object
  '   in which we can store,
  ''and format the main message body in Ri
  '   chText format
  Set lnRTItem = lnDocument.CreateRichTextItem("Body")
  If strAttachment <> "" Then
    Set lnATTACHMENT = lnRTItem.EMBEDOBJECT _
    (1454, "", strAttachment, "Sample")
  End If
  sMessage = "Mail sent: " & Date & " " & Time & vbCrLf & vbCrLf & _
  strMessage
  ''format the message
  lnRTStyle.NotesFont = 4 ''Courier
  lnRTStyle.Bold = True
  lnRTStyle.NotesColor = 2 ''red
  Call lnRTItem.AppendStyle(lnRTStyle)
  Call lnRTItem.AppendText(sMessage)
  'Call lnRTItem.AddNewLine(1)
  ''logo values are between 0 and 31
  lLogo = lngLogo
  If lLogo < 0 Or lLogo > 31 Then
    lLogo = 0
  End If
  ''replace some of the fields that we nee
  '   d...
  With lnDocument
    ''who we want to send to...
    ''recipient
    .ReplaceItemValue "SendTo", strSendTo
    ''subject
    .ReplaceItemValue "Subject", strSubject
    ''body - non RichText
    '.ReplaceItemValue "Body", "The body of
    '   the message!"
    ''set the logo! (letter head)
    .ReplaceItemValue "Logo", "StdNotesLtr" & Trim$(Str$(lLogo))
    ''send the message
    .Send False
  End With
  Set lRTItem = Nothing
  Set lnRTStyle = Nothing
  Set lnDocument = Nothing
  Set lnDatabase = Nothing
  Set lnSession = Nothing
  MsgBox "Mail was sent!", vbInformation, _
  strSendTo
  Exit Function
NotesMail_Err:
  MsgBox Err.Description, _
  vbExclamation, _
  "Send mail error! (" & Trim$(Str$(Err)) & ")"
End Function
Function Test_note()
  SendNotesMail "Hello! This is an email message! with an attachment", _
  "Test Lotus Notes Email - Attachment test", _
  "youraddress@work", 0, "C:\autoexec.bat"
End Function
```

