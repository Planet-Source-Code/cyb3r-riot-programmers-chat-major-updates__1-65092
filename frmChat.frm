VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programmers Chat"
   ClientHeight    =   8895
   ClientLeft      =   6870
   ClientTop       =   1425
   ClientWidth     =   9015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   2760
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save File"
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load File"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   3240
      Width           =   1095
   End
   Begin VB.PictureBox picLN 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   5155
      Left            =   3
      ScaleHeight     =   5160
      ScaleWidth      =   450
      TabIndex        =   10
      Top             =   3720
      Width           =   450
   End
   Begin VB.Timer tmrLN 
      Interval        =   25
      Left            =   3360
      Top             =   3240
   End
   Begin RichTextLib.RichTextBox txtSharing 
      Height          =   5175
      Left            =   480
      TabIndex        =   9
      Top             =   3720
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9128
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmChat.frx":0000
   End
   Begin VB.TextBox txtWF 
      Height          =   315
      Left            =   5640
      TabIndex        =   5
      Top             =   3300
      Width           =   1935
   End
   Begin VB.ListBox lstFiles 
      Height          =   2790
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdSendShare 
      Caption         =   "Update"
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtChat 
      Height          =   2295
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   360
      Width           =   6615
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox txtSend 
      Height          =   405
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2760
      Width           =   5775
   End
   Begin VB.Label Label3 
      Caption         =   "Chat Area :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   45
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Workable Files :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   50
      TabIndex        =   7
      Top             =   50
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Working File Name :"
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   3360
      Width           =   1935
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim ShareFiles() As String
Private Sub cmdLoad_Click()
With cd
.Filter = "All Files (*.*)|*.*|"
.InitDir = App.Path
.ShowOpen
If .FileName = "" Then MsgBox "Load Canceled!", , "Canceled": Exit Sub
LoadIt .FileName
End With
End Sub
Private Sub cmdSave_Click()
With cd
.Filter = "All Files (*.*)|*.*|"
.InitDir = App.Path
.ShowSave
If .FileName = "" Then MsgBox "Save Canceled!", , "Canceled": Exit Sub
SaveIt .FileName, ShareFiles(lstFiles.ListIndex)
End With
End Sub
Sub LoadIt(FN As String)
On Error GoTo error
Dim txt As String, txt2 As String
Open FN For Input As #1
Do While Not EOF(1)
    Line Input #1, txt
    txt2 = txt2 & txt & vbNewLine
Loop
txt2 = Left$(txt2, Len(txt2) - 1)
FN = StrReverse$(FN)
FN = Left$(FN, InStr(1, FN, "\") - 1)
FN = StrReverse$(FN)
DoEvents
frmServer.sck2.SendData frmServer.txtName.Text & "|*|SHARE|*|" & txt2 & "|*|" & FN
DoEvents
Close #1
Exit Sub
error:
    x = MsgBox("Error Loading!", vbOKOnly, "Error")
End Sub
Sub SaveIt(FN As String, txt As String)
On Error GoTo error
Open FN For Output As #1
Print #1, txt
Close 1
Exit Sub
error:
    x = MsgBox("Error Saving!", vbOKOnly, "Error")
End Sub
Private Sub cmdSend_Click()
On Error Resume Next
frmServer.sck2.SendData frmServer.txtName.Text & "|*|CHAT|*|" & txtSend
DoEvents
txtSend = ""
End Sub
Private Sub cmdSendShare_Click()
frmServer.sck2.SendData frmServer.txtName.Text & "|*|SHARE|*|" & txtSharing.Text & "|*|" & txtWF.Text
DoEvents
End Sub
Private Sub Form_Load()
ReDim ShareFiles(0)
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
frmServer.sck2.SendData frmServer.txtName.Text & "|*|CLOSE"
DoEvents
frmServer.sck2.Close
frmServer.txtName.Locked = False
End Sub
Private Sub lstFiles_Click()
txtWF = lstFiles.List(lstFiles.ListIndex)
txtSharing.Text = ShareFiles(lstFiles.ListIndex)
End Sub
Private Sub tmrLN_Timer()
Dim i As Long, lc As Long, cl As Long
lc = SendMessageLong(txtSharing.hWnd, &HBA, 0, 0&)
cl = SendMessageLong(txtSharing.hWnd, &HCE, 0, 0&) + 1
txtLN = ""
With picLN
.Cls
hgt = .TextHeight("WOW")
cy = 50
For i = cl To lc
    .CurrentX = 0
    .CurrentY = cy
    picLN.Print i
    cy = cy + hgt
Next i
.Refresh
End With
End Sub
Private Sub txtChat_Change()
txtChat.SelStart = Len(txtChat.Text)
End Sub
Private Sub txtSend_GotFocus()
cmdSend.Default = True
End Sub
Sub AddShare(sn, txt)
Dim i As Integer
For i = 0 To lstFiles.ListCount - 1
    If lstFiles.List(i) = sn Then
    ShareFiles(i) = txt
    Exit Sub
    End If
Next i
lstFiles.AddItem sn
ReDim Preserve ShareFiles(lstFiles.ListCount - 1)
ShareFiles(lstFiles.ListCount - 1) = txt
End Sub

