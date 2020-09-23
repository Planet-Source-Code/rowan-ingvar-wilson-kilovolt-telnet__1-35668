VERSION 5.00
Begin VB.Form fConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3780
   Icon            =   "Connect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect..."
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdOtherHost 
      Caption         =   "..."
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.ComboBox Port 
      Height          =   315
      ItemData        =   "Connect.frx":0442
      Left            =   1080
      List            =   "Connect.frx":044F
      TabIndex        =   2
      Text            =   "Telnet"
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblHost 
      Caption         =   "Host"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "fConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConnect_Click()
On Error GoTo ErrHandler
' Issue "Connect" command
Select Case Port
    Case "Telnet"
        Port = "23"
    Case "FTP"
        Port = "21"
    Case "SMTP"
        Port = "25"
    Case Else
        
End Select
Main.Connect txtHost.Text, CLng(Port)
DoEvents
Main.Show
Unload Me
Exit Sub
ErrHandler:
MsgBox Err.Description
End Sub
