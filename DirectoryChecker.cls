VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DirectoryChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Friend Sub CreateInPath(location As String)
    Dim directoryArr As Variant
    Dim checkDirectory As String, fullDirectory As String
    '   Split path
    Dim sTempPath As String, wbPath As String
    wbPath = ActiveWorkbook.path
    sTempPath = Split(location, wbPath & "\")(1)
    '   Initialize directory with workbook path folder
    fullDirectory = wbPath
    directoryArr = Split(sTempPath, "\")
    Dim directory As Variant
    For Each directory In directoryArr
        fullDirectory = fullDirectory & "\" & directory
        If Len(dir(fullDirectory, vbDirectory)) = 0 Then
            MkDir (fullDirectory)
        End If
    Next
End Sub
