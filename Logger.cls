VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'   logger built with help from https://excelmacromastery.com/vba-error-handling/

Dim m_fileName As String
Dim m_fileNumber As Variant

Public Sub openFile(path As String)
    '   Archive at 20kb of data
    If dir(path) <> "" Then
        If FileLen(path) > 20000 Then
            FileCopy path _
                , Replace(path, ".txt", Format(Now, "ARCHIVED ddmmyyyy hhmmss.txt"))
            Kill path
        End If
    End If
    m_fileName = path
    m_fileNumber = FreeFile
    Open m_fileName For Append As m_fileNumber
End Sub

Public Sub writeToFile(sType As String, sSource As String, sDetails As String, Optional path As String)
    Dim data As String
    data = sType & ", " & sSource & ", " & sDetails
    If Not path = "" Then
        openFile (path)
        printToFile (data)
        closeFile
    Else
        printToFile (data)
    End If
End Sub

Public Sub closeFile()
    Close m_fileNumber
End Sub

Private Sub printToFile(data As String)
    Debug.Print data
    Print #m_fileNumber, data
End Sub


