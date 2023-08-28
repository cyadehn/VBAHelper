VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_errNumber As Long
Dim m_errSource As String
Dim m_errDescription As String

Public Sub RaiseError(number As Long, source As String, description As String, line As Integer)
    m_errNumber = number
    If m_errSource = "" Then m_errSource = source Else m_errSource = source & " (" & line & ") > " & m_errSource
    m_errDescription = description
End Sub

Public Sub LogError()
    Dim userMessage As String, logSource As String, logDirectory As String
    logSource = UCase(Environ("username")) & "(" & UCase(Environ("computername")) & "), Error #" & m_errNumber & " at " & m_errSource
    userMessage = "Error #" & m_errNumber & " at " & m_errSource & ". " & vbCrLf & vbCrLf & "Error: " & m_errDescription
    logDirectory = Application.ActiveWorkbook.path & "\Manifests\errors\"
    '   Check if directories exist
    checker.CreateInPath (Split(logDirectory, Application.ActiveWorkbook.path & "\")(1))
    '   Log error message
    logger.openFile (logDirectory & "Error Log.txt")
    logger.writeToFile "Error", logSource, m_errDescription
    logger.closeFile
    '   Save copy of workbook
    '   Notify user
    MsgBox (userMessage & vbCrLf & vbCrLf & "Please correct & try again or email cdehner@ups.com with any questions.")
End Sub

Sub SaveErrorData()
    Application.ScreenUpdating = False
    
    Set template = Workbooks(ActiveWorkbook.name)
    Dim loc As String, name As String, path As String, fullpath As String
    loc = "FWSManifests\Error\" & Format(Now(), "MMM-DD-YYYY")
    name = ActiveWorkbook.name
    path = ActiveWorkbook.path
    With checker
        .CreateInPath (loc)
    End With
    '   Save DATA sheet to location
    fullpath = path & "\" & loc & "\error " & Format(Now, "MM-DD HHMM") & ".xlsm"
    template.SaveCopyAs (fullpath)
    '   Email Chris with link to error data
    '   Create Mail With Attached Files
    Dim em As Outlook.MailItem
    Set em = emailer.AddNew
    With em
        .To = "cdehner@ups.com"
        .Recipients.ResolveAll
        .subject = ("Manifest Error " & Format(Now(), "MM-DD HH:MM"))
        .Attachments.Add (fullpath)
        .Display
        .HTMLBody = "Please add any details on the error below:" & vbNewLine & .HTMLBody
    End With
    Application.ScreenUpdating = True
End Sub
