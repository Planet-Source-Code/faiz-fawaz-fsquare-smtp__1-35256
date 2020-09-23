VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SMTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"EmailTos"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarSMTPServer As String 'local copy
Private mvarSMTPPort As Long 'local copy
Private mvarMailFrom As String 'local copy
Private mvarEmailTos As EmailTos
Private mvarData As String 'local copy
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent MailSend[(arg1, arg2, ... , argn)]
Public Event MailSend()
Public Event MailErrors(ErrText As String)
Public Event MailStatus(StatText As String)

Public Sub MailSuccess()
    RaiseEvent MailSend
End Sub

Public Sub SendStatus(StatText As String)
    RaiseEvent MailStatus(StatText)
End Sub

Public Sub MailError(ErrText As String)
    RaiseEvent MailErrors(ErrText)
End Sub
Public Sub SendMail()
    Dim ToMails()           As String
    Dim Ctr                 As Long
    Dim xMail               As EmailTo
    
    ReDim ToMails(Me.EmailTos.Count)
    
    Ctr = 0
    
    For Each xMail In Me.EmailTos
        Ctr = Ctr + 1
        ToMails(Ctr) = xMail.ToAddress
    Next
    
    frmWS.SendMail Me, mvarSMTPServer, mvarSMTPPort, mvarMailFrom, ToMails, mvarData
End Sub

Public Property Let Data(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Data = 5
    mvarData = vData
End Property


Public Property Get Data() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Data
    Data = mvarData
End Property





Public Property Get EmailTos() As EmailTos
    If mvarEmailTos Is Nothing Then
        Set mvarEmailTos = New EmailTos
    End If


    Set EmailTos = mvarEmailTos
End Property


Public Property Set EmailTos(vData As EmailTos)
    Set mvarEmailTos = vData
End Property
Private Sub Class_Terminate()
    Set mvarEmailTos = Nothing
End Sub



Public Property Let MailFrom(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MailFrom = 5
    mvarMailFrom = vData
End Property


Public Property Get MailFrom() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MailFrom
    MailFrom = mvarMailFrom
End Property



Public Property Let SMTPPort(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SMTPPort = 5
    mvarSMTPPort = vData
End Property


Public Property Get SMTPPort() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SMTPPort
    SMTPPort = mvarSMTPPort
End Property



Public Property Let SMTPServer(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SMTPServer = 5
    mvarSMTPServer = vData
End Property


Public Property Get SMTPServer() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SMTPServer
    SMTPServer = mvarSMTPServer
End Property



