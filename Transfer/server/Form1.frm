VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "File Transfer(server) - Manjit"
   ClientHeight    =   6420
   ClientLeft      =   2295
   ClientTop       =   960
   ClientWidth     =   7290
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   7290
   Begin VB.CommandButton Command6 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "About"
      Height          =   735
      Left            =   6360
      TabIndex        =   10
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send"
      Height          =   735
      Left            =   5400
      TabIndex        =   6
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   4920
      Width           =   5055
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   6615
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   6840
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "send"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   5880
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6840
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6840
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "File:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Connect to ip :"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1815
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


Dim fname As String
Dim fnamea As String


Private Const chunk = 8000

Private Sub Command1_Click()
Winsock1.Connect Text3, "165"
Winsock2.Connect Text3, "166"

End Sub

Private Sub Command2_Click()
cd1.ShowOpen
If vbOK Then Text1 = cd1.FileName
End Sub



Private Sub Command3_Click()

'GET FILE NAME
'using getfilename function to get the file name


If Text1.Text = "" Then
MsgBox "Please type the file name!!!", , "Manjit"
Exit Sub
End If
fname = Text1.Text
'checking wether the file exists
If Dir(fname) = "" Then
MsgBox "File Does not exist Exists", , "manjit"
Exit Sub 'exiting sub it file does not exists
End If

fnamea = GetFileName(Text1.Text)
'sending file name of file
Dim temp2 As String
temp2 = "rqst" & fnamea
Winsock1.SendData temp2







End Sub
Private Sub send(fname As String)
Command2.Enabled = False
Command3.Enabled = False
Text1.Enabled = False

Dim data As String
Dim a As Long
Dim data1 As String
Dim data2 As String




Open fname For Binary As #1

Do While Not EOF(1)
data = Input(chunk, #1)
Winsock1.SendData data

DoEvents
Loop

Winsock1.SendData "EnDf"
Close #1
Command2.Enabled = True
Command3.Enabled = True
Text1.Enabled = True

End Sub




Function GetFileName(attach_str As String) As String
    Dim s As Integer
    Dim temp As String
    s = InStr(1, attach_str, "\")
    temp = attach_str
    Do While s > 0
        temp = Mid(temp, s + 1, Len(temp))
        s = InStr(1, temp, "\")
    Loop
    GetFileName = temp
End Function

Private Sub Command4_Click()
List1.AddItem (Text2.Text)

Winsock2.SendData Text2.Text
Text2.Text = ""
End Sub

Private Sub Command5_Click()
About.Show

End Sub

Private Sub Command6_Click()
Winsock1.Close
Winsock2.Close

End Sub

Private Sub mnu_abt_Click()
About.Show

End Sub

Private Sub mnu_exit_Click()
Winsock1.Close
Winsock2.Close
End
End Sub

Private Sub winsock1_dataarrival(ByVal bytestotal As Long)
Dim response As String
Winsock1.GetData response, vbString
Select Case response
Case "okay"
send fname
Case "deny"
MsgBox "Your request to send the file " & fname & " has been denied", , "manjit"
End Select

End Sub

Private Sub winsock2_dataarrival(ByVal bytestotal As Long)
Dim cht As String
Winsock2.GetData cht, vbString
List1.AddItem (cht)
End Sub
