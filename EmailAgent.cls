VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailAgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_emails As Collection
Dim m_outApp As Outlook.Application
Dim m_activeEmail As Outlook.MailItem

Private Sub Class_Initialize()
    Set m_outApp = GetObject(, "Outlook.Application")
    Set m_emails = New Collection
End Sub

Property Get outApp() As Outlook.Application
    Set outApp = m_outApp
End Property

Property Get emails(i As Integer) As Outlook.MailItem
    If i <= m_emails.Count And i > 0 Then
        Set emails = m_emails(i)
    End If
End Property

Property Get activeEmail() As Outlook.MailItem
    activeEmail = m_activeEmail
End Property

Property Let activeCC(addresses As String)
    With m_activeEmail
        .cc = addresses
        .Recipients.ResolveAll
    End With
End Property

Property Let activeTo(addresses As String)
    With m_activeEmail
        .To = addresses
        .Recipients.ResolveAll
    End With
End Property

Property Let activeSubject(subject As String)
    m_activeEmail.subject = subject
End Property

Public Function AddNew() As Outlook.MailItem
    Dim em As Outlook.MailItem
    '   Set up email in outlook object
    Set em = m_outApp.CreateItem(olMailItem)
    '   Add to collection for later use
    m_emails.Add em
    Set m_activeEmail = em
    Set AddNew = em
End Function

Public Sub SaveDrafts()
    Dim em As Outlook.MailItem
    For Each em In m_emails
        em.Close (olSave)
    Next em
End Sub

Public Sub DeleteAll()
    Dim em As Outlook.MailItem
    For Each em In m_emails
        em.Delete
    Next em
End Sub

Public Sub SendAll()
    Dim em As Outlook.MailItem
    For Each em In m_emails
        em.Send
    Next em
End Sub

Public Sub AttachFile(filePath As String)
    m_activeEmail.Attachments.Add (filePath)
End Sub

Public Function SetActiveEmail(subjectSearchTerm As String) As Integer
    Dim i As Integer, em As Outlook.MailItem
    For i = 1 To m_emails.Count
        Set em = m_emails(i)
        If InStr(em.subject, subjectSearchTerm) Then
            Set m_activeEmail = em
            FindEmail = 0
        Else
            FindEmail = 1
        End If
    Next i
End Function
