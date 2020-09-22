VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   4440
   ClientLeft      =   5520
   ClientTop       =   3645
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Client"
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   3120
      Width           =   3735
      Begin VB.TextBox txtCPort 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Text            =   "1005"
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdClient 
         Caption         =   "Connect"
         Height          =   615
         Left            =   2160
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "IP :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   840
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sck 
      Index           =   0
      Left            =   3120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1005
      LocalPort       =   1005
   End
   Begin MSWinsockLib.Winsock sck2 
      Left            =   3120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Text            =   "Programming Host"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server"
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.ListBox lstUsers 
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3495
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start Server"
         Height          =   615
         Left            =   2160
         TabIndex        =   3
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Text            =   "1005"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Text            =   "127.0.0.1"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Connected Users"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "My IP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Alias:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   4200
      Width           =   735
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClient_Click()
Connect txtServer, txtCPort
End Sub
Private Sub Form_Load()
txtIP = sck(0).LocalIP
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim i As Integer
For i = 1 To sck.UBound - 1
    sck(i).Close
    Unload sck(i)
Next i
End
End Sub
Private Sub sck2_Close()
txtName.Locked = False
End Sub
Private Sub sck2_Connect()
un = txtName
sck2.SendData un & "|*|USER"
txtName.Locked = True
frmChat.Show
End Sub
Private Sub cmdStart_Click()
sck(0).Close
DoEvents
sck(0).LocalPort = txtPort
sck(0).Listen
Connect "localhost", txtPort
While sck2.State <> sckConnected
    DoEvents
Wend
End Sub
Sub Connect(IP As String, Port As Integer)
If sck2.State <> sckClosed Then Exit Sub
un = txtName
sck2.Connect IP, Port
DoEvents
End Sub
Private Sub sck_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
Load sck(AvSocket + 1)
sck(AvSocket).Accept requestID
DoEvents
cmdStart_Click
End Sub
Private Sub sck_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
sck(Index).GetData strData, vbString
DoEvents
Disect strData
End Sub
Sub Disect(strS As String)
On Error Resume Next
strS = strS & "|*|"
un = Left$(strS, InStr(1, strS, "|*|") - 1)
strS = Right$(strS, Len(strS) - Len(un) - 3)
cmd = Left$(strS, InStr(1, strS, "|*|") - 1)
strS = Right$(strS, Len(strS) - Len(cmd) - 3)
txt = Left$(strS, InStr(1, strS, "|*|") - 1)
strS = Right$(strS, Len(strS) - Len(txt) - 3)
sn = Left$(strS, InStr(1, strS, "|*|") - 1)

Select Case cmd
Case Is = "USER"
lstUsers.AddItem un
For i = 1 To sck.UBound - 1
    sck(i).SendData un & "|*|USER"
    DoEvents
Next i
Case Is = "CHAT"
For i = 1 To sck.UBound - 1
    sck(i).SendData un & "|*|CHAT|*|" & txt
    DoEvents
Next i
Case Is = "SHARE"
For i = 1 To sck.UBound - 1
    sck(i).SendData un & "|*|SHARE|*|" & txt & "|*|" & sn
    DoEvents
Next i
Case Is = "CLOSE"
For i = 1 To sck.UBound - 1
    sck(i).SendData un & "|*|CLOSE"
    DoEvents
Next i
For i = 0 To lstUsers.ListCount - 1
    If lstUsers.List(i) = un Then lstUsers.RemoveItem (i): Exit Sub
Next i
End Select
End Sub
Sub Disect2(strS As String)
On Error Resume Next
strS = strS & "|*|"
un = Left$(strS, InStr(1, strS, "|*|") - 1)
strS = Right$(strS, Len(strS) - Len(un) - 3)
cmd = Left$(strS, InStr(1, strS, "|*|") - 1)
strS = Right$(strS, Len(strS) - Len(cmd) - 3)
txt = Left$(strS, InStr(1, strS, "|*|") - 1)
strS = Right$(strS, Len(strS) - Len(txt) - 3)
sn = Left$(strS, InStr(1, strS, "|*|") - 1)

Select Case cmd
Case Is = "USER"
frmChat.txtChat.Text = frmChat.txtChat.Text & "User +" & un & "+ has connected." & vbNewLine
Case Is = "CLOSE"
frmChat.txtChat.Text = frmChat.txtChat.Text & "User +" & un & "+ has disconnected." & vbNewLine
Case Is = "CHAT"
    frmChat.txtChat.Text = frmChat.txtChat.Text & un & ": " & txt & vbNewLine
Case Is = "SHARE"
    frmChat.AddShare sn, txt
End Select
End Sub
Private Sub sck_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox Err.Description, Err.Number
End Sub
Function AvSocket() As Integer
On Error Resume Next
Load sck(1)
For i = 1 To sck.UBound
If sck(i).State <> sckConnected Then
sck(i).Close
AvSocket = i
Exit Function
End If
Next i
End Function
Private Sub sck2_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
sck2.GetData strData, vbString
DoEvents
Disect2 strData
End Sub
