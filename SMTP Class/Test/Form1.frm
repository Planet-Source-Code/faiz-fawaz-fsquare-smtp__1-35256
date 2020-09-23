VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   2295
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   2640
      Width           =   6615
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "25"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Mail"
      Height          =   1215
      Left            =   5040
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "To"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "SMTP Server"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents xyz          As FSquareSMTP.SMTP
Attribute xyz.VB_VarHelpID = -1
Private Sub Command1_Click()
    Set xyz = New FSquareSMTP.SMTP
    
    xyz.SMTPServer = Text1
    xyz.SMTPPort = Text2
    xyz.MailFrom = Text3
    xyz.EmailTos.Add Text4
    
    xyz.Data = Text5
    
    xyz.SendMail
End Sub

Private Sub xyz_MailError(ErrText As String)
    MsgBox "OOOps...." & vbCrLf & ErrText
End Sub

Private Sub xyz_MailSend()
    MsgBox "Done"
End Sub

Private Sub xyz_MailStatus(StatText As String)
    Label1 = StatText
End Sub
