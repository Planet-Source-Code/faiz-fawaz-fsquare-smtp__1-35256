VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmWS 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1320
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmWS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xFromMail           As String
Dim xToMail()           As String
Dim xDataStr            As String
Dim xRetObj             As SMTP
Public Function SendMail(RetObj As SMTP, Server As String, Port As Long, FromMail As String, ToMail() As String, DataStr As String)
    xFromMail = FromMail
    xToMail = ToMail
    xDataStr = DataStr
    
    Set xRetObj = RetObj
    
    xRetObj.SendStatus "Connecting to server [" & Server & ":" & Port & "]..."
    Winsock1.Connect Server, Port
    
End Function

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim xStr            As String
    Static CurState     As String
    Static CurPos       As Long
    
    Winsock1.GetData xStr
    
    If CurState = "" Then CurState = "1"
    
    If Mid(xStr, 1, 1) = "3" Or Mid(xStr, 1, 1) = "2" Then
        Select Case CurState
            Case "1"
                Winsock1.SendData "HELO FSQUARE" & vbCrLf
                xRetObj.SendStatus "Handshaking with server..."
                CurState = "2"
            Case "2"
                xRetObj.SendStatus "Sending From details..."
                Winsock1.SendData "MAIL FROM: " & xFromMail & vbCrLf
                CurState = "3"
                CurPos = 0
            Case "3"
                CurPos = CurPos + 1
                If CurPos = UBound(xToMail) Then
                    xRetObj.SendStatus "Sending to details (" & CurPos & ")..."
                    CurState = "4"
                End If
                Winsock1.SendData "RCPT TO: " & xToMail(CurPos) & vbCrLf
            Case "4"
                xRetObj.SendStatus "Sending to data..."
                Winsock1.SendData "DATA" & vbCrLf
                Winsock1.SendData xDataStr & vbCrLf & "." & vbCrLf
                CurState = "5"
            Case "5"
                xRetObj.SendStatus "Saying bye to server..."
                Winsock1.SendData "QUIT" & vbCrLf
                CurState = ""
                xRetObj.MailSuccess
        End Select
    Else
        xRetObj.MailError xStr
    End If
End Sub

