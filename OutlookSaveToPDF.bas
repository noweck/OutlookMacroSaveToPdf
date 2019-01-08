'#####################################
'### Outlook Save To PDF           ###
'### Developed by: Matthiew Morin, ###
'###               noweck          ###
'#####################################
' requires reference to Microsoft Scripting Runtime
' \Windows\System32\Scrrun.dll
' Also requires reference to Microsoft Word Object Library
' Add reference Microsoft Forms 2.0 Object Library (c:\Windows\System32\FM20.dll) IF needed

Option Explicit
#If VBA7 Then
  Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As Long
  Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
  Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As Long
  Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
  Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
  Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As Long
#Else
  Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
  Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
  Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
  Declare Function CloseClipboard Lib "User32" () As Long
  Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
  Declare Function EmptyClipboard Lib "User32" () As Long
  Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
  Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
#End If

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Sub SaveAsMsg()
    Dim fso As FileSystemObject
    Dim strSubject As String
    Dim strSaveName As String
    Dim blnOverwrite As Boolean
    Dim strFolderPath As String
    Dim looper As Integer
    Dim strID As String
    Dim olNS As Outlook.NameSpace
    Dim oMail As Outlook.MailItem
    Dim atmtName As String
    
    Dim MyMail As Outlook.MailItem
    Dim emailSubject As String
    Dim sendEmailAddr As String
    Dim companyDomain As String
    
    Dim bPath As String
    Dim yPath As String
    Dim mPath As String
    Dim dPath As String
    Dim ePath As String
    Dim saveName As String
    Dim pdfSave As String
    
    
    Set MyMail = Application.ActiveExplorer.Selection(1)
    
    strID = MyMail.EntryID
    Set olNS = Application.GetNamespace("MAPI")
    Set oMail = olNS.GetItemFromID(strID)
    
    'Get Sender email domain
    sendEmailAddr = oMail.SenderEmailAddress
    companyDomain = Right(sendEmailAddr, Len(sendEmailAddr) - InStr(sendEmailAddr, "@"))
    
    ' ### USER OPTIONS ###
    blnOverwrite = False ' False = don't overwrite, True = do overwrite
    
    '### THIS IS WHERE SAVE LOCATIONS ARE SET ###
    bPath = "C:\Mails\"                         ' Defines the base path to save the email
    yPath = bPath & Format(Now(), "yyyy") & "\" ' Add year subfolder
    mPath = yPath & Format(Now(), "MMMM") & "\" ' Add month subfolder
    dPath = mPath & Format(Now(), "dd") & "\"   ' Add day subfolder
    
    '### Path Validity ###
    If Dir(bPath, vbDirectory) = vbNullString Then
        MkDir bPath
    End If
    If Dir(yPath, vbDirectory) = vbNullString Then
        MkDir yPath
    End If
    If Dir(mPath, vbDirectory) = vbNullString Then
        MkDir mPath
    End If
    If Dir(dPath, vbDirectory) = vbNullString Then
        MkDir dPath
    End If
    
    
    '### Get Email subject & set name to be saved as ###
    emailSubject = CleanFileName(oMail.Subject)
    
    ePath = dPath & emailSubject & "\" ' Add day subfolder - clear file name
    
    If Dir(ePath, vbDirectory) = vbNullString Then
        MkDir ePath
    End If
    
    ' ### Set email file name ###
    '  saveName = Format(oMail.ReceivedTime, "yyyymmdd") & "_" & emailSubject & ".mht"
        
    saveName = emailSubject & ".mht"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '### If don't overwrite is on then ###
    If blnOverwrite = False Then
        looper = 0
        Do While fso.FileExists(dPath & saveName)
            looper = looper + 1
            saveName = emailSubject & looper & ".mht" '& "_" Format(oMail.ReceivedTime, "yyyymmdd") & "_" &
            pdfSave = emailSubject & looper & ".pdf"  '& "_" Format(oMail.ReceivedTime, "yyyymmdd") & "_" &
            Loop
    Else '### If don't overwrite is off, delete the file ###
        If fso.FileExists(dPath & saveName) Then
            fso.DeleteFile dPath & saveName
        End If
    End If
    oMail.SaveAs ePath & saveName, olMHTML
    pdfSave = ePath & "email_" & emailSubject & ".pdf"  '& "_" Format(oMail.ReceivedTime, "yyyymmdd") &
    
    '### Open Word to convert file to PDF ###
    Dim wrdApp As Word.Application
    Dim wrdDoc As Word.Document
    Set wrdApp = CreateObject("Word.Application")
    
    Set wrdDoc = wrdApp.Documents.Open(FileName:=ePath & saveName, Visible:=True)
    wrdApp.ActiveDocument.ExportAsFixedFormat OutputFileName:= _
                pdfSave, ExportFormat:= _
                wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
                wdExportOptimizeForPrint, Range:=wdExportAllDocument, From:=0, To:=0, _
                Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
                CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
                BitmapMissingFonts:=True, UseISO19005_1:=False
    
    wrdDoc.Close
    wrdApp.Quit
    
    '### Clean up files ###
    With New FileSystemObject
        If .FileExists(ePath & saveName) Then
            .DeleteFile ePath & saveName
        End If
    End With
    
    '### If Mail Attachments: clean file name, save into path ###
    If oMail.Attachments.Count > 0 Then
        Dim atmt As Object
        For Each atmt In oMail.Attachments
            atmtName = CleanFileName(atmt.FileName)
            atmt.SaveAsFile ePath & atmtName ' & "_" Format(oMail.ReceivedTime, "yyyymmdd")
        Next
        Set atmt = Nothing
    End If
    
    Set oMail = Nothing
    Set olNS = Nothing
    Set fso = Nothing
        
    '### Copy emailSubject to clipboard ###
    ClipBoard_SetData emailSubject
    
    '### Open Exlorer on ePath ###
    Shell "explorer " & ePath, vbNormalFocus

End
End Sub

'### Filename corrections ###
Private Function CleanFileName(strText As String) As String
    Dim strStripChars As String
    Dim intLen As Integer
    Dim i As Integer
    strText = Trim(strText)
    strStripChars = "/\[]:=,?" & Chr(34)
    intLen = Len(strStripChars)
    strText = Trim(strText)
    For i = 1 To intLen
    strText = Replace(strText, Mid(strStripChars, i, 1), "")
    Next
    CleanFileName = strText
End Function

'### Copy mail subject to Clipboard ###
Private Function ClipBoard_SetData(sPutToClip As String) As Boolean

    ' www.msdn.microsoft.com/en-us/library/office/ff192913.aspx

    Dim hGlobalMemory As Long
    Dim lpGlobalMemory As Long
    Dim hClipMemory As Long
    Dim X As Long
    
    On Error GoTo ExitWithError_

    ' Allocate moveable global memory
    hGlobalMemory = GlobalAlloc(GHND, Len(sPutToClip) + 1)

    ' Lock the block to get a far pointer to this memory
    lpGlobalMemory = GlobalLock(hGlobalMemory)

    ' Copy the string to this global memory
    lpGlobalMemory = lstrcpy(lpGlobalMemory, sPutToClip)

    ' Unlock the memory
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "Memory location could not be unlocked. Clipboard copy aborted", vbCritical, "API Clipboard Copy"
        GoTo ExitWithError_
    End If

    ' Open the Clipboard to copy data to
    If OpenClipboard(0&) = 0 Then
        MsgBox "Clipboard could not be opened. Copy aborted!", vbCritical, "API Clipboard Copy"
        GoTo ExitWithError_
    End If

    ' Clear the Clipboard
    X = EmptyClipboard()

    ' Copy the data to the Clipboard
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)
    ClipBoard_SetData = True
    
    If CloseClipboard() = 0 Then
        MsgBox "Clipboard could not be closed!", vbCritical, "API Clipboard Copy"
    End If
    Exit Function
ExitWithError_:
    On Error Resume Next
    If Err.Number > 0 Then MsgBox "Clipboard error: " & Err.Description, vbCritical, "API Clipboard Copy"
    ClipBoard_SetData = False

End Function

