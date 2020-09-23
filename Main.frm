VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Main 
   Caption         =   "KiloVolt Telnet"
   ClientHeight    =   5880
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000FF00&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   7560
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3480
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   1320
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox Echo 
      Caption         =   "&Echo local characters"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   7335
   End
   Begin VB.TextBox Terminal 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
   Begin VB.Menu mnuConnect 
      Caption         =   "&Connect..."
      Begin VB.Menu mnuConnectHost 
         Caption         =   "&Host..."
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect..."
      End
   End
   Begin VB.Menu mnuPref 
      Caption         =   "&Preferences"
      Begin VB.Menu mnuSetForeground 
         Caption         =   "&Foreground colour"
      End
      Begin VB.Menu mnuSetBackground 
         Caption         =   "&Background colour"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Kilovolt Telnet by Rowan Ingvar Wilson
' leet_llama@hotmail.com - email/mSn

Dim fColor As Long
Dim bColor As Long
Dim connected As Boolean

Function AddToTerminal(What As String)
Dim Char As String
' Step through the characters...
For i = 1 To Len(What)
        ' Get next char
        Char = Mid$(What, i, 1)
        ' Find out if it's within the range of useful characters
        If Asc(Char) >= 0 And Asc(Char) < 128 Then
                Terminal.Text = Terminal.Text & Char
        End If
Next i
Terminal.SelStart = Len(Terminal)
End Function

Private Sub Form_Load()
' Load the colours
fColor = GetSetting(App.Title, "Colours", "Forecolor", vbWhite)
bColor = GetSetting(App.Title, "Colours", "Backcolor", vbBlack)
' Apply them
Terminal.BackColor = bColor
Terminal.ForeColor = fColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
' Clean up
DoEvents
Socket.Close
End

End Sub

Private Sub mnuConnectHost_Click()
' Show the connect form
fConnect.Show
End Sub

Function Connect(host As String, Port As Long)
On Error GoTo ErrHandle
AddToTerminal "Connecting " & host & "..." & vbCrLf
Socket.Connect host, Port

Exit Function
ErrHandle:
AddToTerminal "Error connecting to host - " & Err.Description & vbCrLf
connected = False
End Function

Private Sub mnuDisconnect_Click()
' Close the socket
Socket.Close
mnuConnectHost.Visible = True
connected = False
End Sub

Private Sub mnuSetBackground_Click()
' Set bg colour
CommonDialog.Color = Terminal.BackColor
CommonDialog.ShowColor
bColor = CommonDialog.Color
SaveSetting App.Title, "Colours", "Backcolor", bColor
DoEvents
Terminal.BackColor = bColor
End Sub

Private Sub mnuSetForeground_Click()
' Set fg colour
CommonDialog.Color = Terminal.ForeColor
CommonDialog.ShowColor
fColor = CommonDialog.Color
SaveSetting App.Title, "Colours", "Forecolor", fColor
DoEvents
Terminal.ForeColor = fColor
End Sub

Private Sub Socket_Close()
mnuConnectHost.Visible = True
AddToTerminal "Connection closed by remote host"
connected = False
End Sub

Private Sub Socket_Connect()
AddToTerminal "Connect " & Socket.RemoteHost & "(" & Socket.RemoteHostIP & ")" & vbCrLf
mnuConnectHost.Visible = False
connected = True
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
Dim tData As String
Socket.GetData tData
' Add it to the terminal
AddToTerminal tData

End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
' Log it
AddToTerminal "Socket error: " & Description & vbCrLf
connected = False
End Sub

Private Sub Terminal_KeyPress(KeyAscii As Integer)
If connected Then
If Echo Then
    AddToTerminal Chr$(KeyAscii)
End If
Socket.SendData Chr$(KeyAscii)
DoEvents
End If
End Sub
