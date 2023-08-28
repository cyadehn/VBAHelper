VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VCSHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'   Code adapted from Jim Harte's answer at the following link:
'   https://stackoverflow.com/questions/49724/programmatically-extract-macro-vba-code-from-word-2007-docs/49796#49796
'   //FEATURE: Import code
Public Sub ImportCode()
    '   Exit if access to VB object model is not allowed
    '   Exit if VBA window is not open
    '   Check if workbook is in a VSC directory
    '   Check and make directories
    '   Import components
End Sub

Sub ExportCode()

    If Not CanAccessVBOM Then Exit Sub ' Exit if access to VB object model is not allowed
    If (ThisWorkbook.VBProject.VBE.ActiveWindow Is Nothing) Then
        Exit Sub ' Exit if VBA window is not open
    End If
    Dim comp As VBComponent
    Dim codeFolder As String

    '   //TODO: use FileHandler to determine if workbook is in a VCS directory (search for .git?)
    codeFolder = "C:\Users\slj9prv\Desktop\DEV\CODE\working\AutoManifest\src"
    On Error Resume Next
    MkDir CombinePaths(codeFolder, "cls")
    MkDir CombinePaths(codeFolder, "bas")
    MkDir CombinePaths(codeFolder, "frm")
    MkDir CombinePaths(codeFolder, "sht")
    On Error GoTo 0
    Dim fName As String

    '   Export components
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case vbext_ct_ClassModule
                fName = CombinePaths(codeFolder, "cls\" & comp.name & ".cls")
                DeleteFile fName
                comp.Export fName
            Case vbext_ct_StdModule
                fName = CombinePaths(codeFolder, "bas\" & comp.name & ".bas")
                DeleteFile fName
                comp.Export fName
            Case vbext_ct_MSForm
                fName = CombinePaths(codeFolder, "frm\" & comp.name & ".frm")
                DeleteFile fName
                comp.Export fName
            Case vbext_ct_Document
                fName = CombinePaths(codeFolder, "sht\" & comp.name & ".cls")
                DeleteFile fName
                comp.Export fName
        End Select
    Next

End Sub

Public Sub BackupWorkbook(sPath As String, sName As String)
    Dim checker As New DirectoryChecker
    Dim wb As Workbook
    Set wb = ThisWorkbook
    checker.CreateInPath (sPath)
    Set checker = Nothing
    '//TODO: delete backups older than a month
    wb.SaveCopyAs (sPath & "/" & sName)
End Sub

Function CanAccessVBOM() As Boolean
    ' Check resgistry to see if we can access the VB object model
    Dim wsh As Object
    Dim str1 As String
    Dim AccessVBOM As Long

    Set wsh = CreateObject("WScript.Shell")
    str1 = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & _
        Application.Version & "\Excel\Security\AccessVBOM"
    On Error Resume Next
    AccessVBOM = wsh.RegRead(str1)
    Set wsh = Nothing
    CanAccessVBOM = (AccessVBOM = 1)
End Function


Sub DeleteFile(fileName As String)
    On Error Resume Next
    Kill fileName
End Sub

Function GetWorkbookPath() As String
    Dim fullName As String
    Dim wrkbookName As String
    Dim pos As Long

    wrkbookName = ThisWorkbook.name
    fullName = ThisWorkbook.fullName

    pos = InStr(1, fullName, wrkbookName, vbTextCompare)

    GetWorkbookPath = Left$(fullName, pos - 1)
End Function

Function CombinePaths(ByVal Path1 As String, ByVal Path2 As String) As String
    If Not EndsWith(Path1, "\") Then
        Path1 = Path1 & "\"
    End If
    CombinePaths = Path1 & Path2
End Function

Function EndsWith(ByVal InString As String, ByVal TestString As String) As Boolean
    EndsWith = (Right$(InString, Len(TestString)) = TestString)
End Function

Function GetUsername() As String
    Set objAD = CreateObject("ADSystemInfo")
    Set objUser = GetObject("LDAP://" & objAD.UserName)
    strDisplayName = objUser.DisplayName
    GetUsername = strDisplayName
End Function

Public Function IsUserAuthorized() As Boolean
    If InStr(UCase(GetUsername), "SLJ9PRV") Then IsUserAuthorized = True Else IsUserAuthorized = False
End Function
