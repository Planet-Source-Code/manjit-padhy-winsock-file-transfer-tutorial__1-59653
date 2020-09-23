VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer(client) - Manjit"
   ClientHeight    =   5880
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   6405
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   6405
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   5520
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   5400
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   4080
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   3975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5520
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   5160
      Width           =   4935
   End
   Begin VB.Menu mnu_exit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mnu_abt 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim sz
Dim size As Integer


Private Sub Command1_Click()
Dim chat As String
chat = Text1.Text
List1.AddItem (chat)
Winsock2.SendData chat
Text1.Text = ""

End Sub

Private Sub Form_Load()
Winsock1.LocalPort = "165"
Winsock1.Listen
size = 0

Winsock2.LocalPort = "166"
Winsock2.Listen
End Sub

Private Sub mnu_abt_Click()
About.Show

End Sub

Private Sub mnu_exit_Click()
Winsock1.Close
Winsock2.Close
End
End Sub

Private Sub winsock1_ConnectionRequest(ByVal idrequest As Long)
Winsock1.Close
Winsock1.Accept idrequest
End Sub

Private Sub winsock1_dataarrival(ByVal bytestotal As Long)

Dim data As String
Dim data4 As String
Dim data2 As String
Dim data3 As String
Dim data5 As String
Dim data6 As String
Dim data7 As String
Dim data8 As String


Winsock1.GetData data, vbString

data2 = Left(data, 4)
Select Case data2
Case "rqst"  'file request arrives

data3 = Right(data, Len(data) - (4)) 'Get the file name

Dim msg1 As Integer  'Stores user's selection
msg1 = MsgBox(Winsock1.RemoteHostIP & " wants to send you file " & data3 & " accept ? ", vbYesNo, "Manjit")  'msgbox displayed


If msg1 = 6 Then  'if user selects yes
Winsock1.SendData "okay"
cd.FileName = data3
data5 = Split(data3, ".")(1)
data6 = "*." & data5
data7 = "Orignal extension (" & data6 & ") |All Files (*.*)|*.*"

cd.Filter = data7


cd.ShowSave
data4 = cd.FileName

Open data4 For Binary As #1

Else
Winsock1.SendData "deny"
Exit Sub
End If


Case "EnDf"
Label1.Caption = "File revieved.Size of file : " & sz & " Kb"
MsgBox "File recieved!!!", , "Manjit"
size = 0
sz = 0

Close #1
Case Else

size = size + 1
sz = size * 8
Label1.Caption = sz & "Kb Recieved"

Put #1, , data
End Select
End Sub
Private Sub winsock2_connectionrequest(ByVal idrequest As Long)
Winsock2.Close
Winsock2.Accept idrequest
End Sub
Private Sub winsock2_dataarrival(ByVal bytestotal As Long)
Dim cht As String
Winsock2.GetData cht, vbString
List1.AddItem (cht)

End Sub

