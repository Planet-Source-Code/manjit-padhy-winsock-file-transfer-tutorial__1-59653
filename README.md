<div align="center">

## \[ winsock file transfer tutorial \]

<img src="PIC2005326723571389.JPG">
</div>

### Description

This tutorial is to explain how to send files (of any size) to any ip using winsock.I've assumed that the reader knows only the basics of winsock...so i've explained it in detail and the codes are higly commented...PLEASE VOTE FOR ME!!!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2005-03-24 18:33:56
**By**             |[Manjit Padhy](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/manjit-padhy.md)
**Level**          |Intermediate
**User Rating**    |4.5 (204 globes from 45 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Complete Applications](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/complete-applications__1-27.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[\[\_winsock\_1868443262005\.zip](https://github.com/Planet-Source-Code/manjit-padhy-winsock-file-transfer-tutorial__1-59653/archive/master.zip)





### Source Code

<p>This tutorial shows how to transfer file of any size using winsock control.<br>
<br>
open vb<br>
select standard exe<br>
<br>
press cntrl+t to show the add component window<br>
<br>
select winsock control and microsoft common dialog<br>
<br>
add one winsock control in the project--name it winsock1<br>
<br>
if you want to add chat then add another winsock and name it winsock2<br>
<br>
insert another winsock object if you want to add chat also<br>
<br>
add a microsoft common dialog box --- name it cd<br>
<br>
we will use this winsock1 object to transfer the file and winsock2 for chat<br>
<br>
<br>
<br>
The basic idea :<br>
To send a file of any size to any ip using winsock first we have to open the file in binary mode.<br>
then get chunks of data from it, chunk is a constant which is initialized to 8000, so we get 8000<br>
bytes of data each time and send it using winsock to the client.<br>
for example let "fname" be the string variable containg the file name then :<br>
<br>
<br>
<br>
Private Const chunk = 8000<br>
<br>
dim fname as string 'get the name of the file<br>
<br>
<br>
Open fname For Binary As #1<br>
Do While Not EOF(1)<br>
data = Input(chunk, #1)<br>
winsock1,sendata data<br>
DoEvents<br>
loop<br>
<br>
<br>
<br>
<br>
this will send 8000bytes of data from the file until the file ends.<br>
<br>
<br>
but before sending data from file to client we must send info about the file <br>
like..the name of the file...the extension...etc<br>
<br>
so when send is clicked first check wether a file is there i mean check wether something<br>
is typed in the text box and if yes check wether the file exists<br>
<br>
if both the above conditions are met then get the filename with the extension.<br>
send the file name to the client with "rqst" in front.<br>
for eg. if the name of file is "text.txt" then send "rqsttext.txt" to the client<br>
<br>
the client will then get the file name and display a msgbox with the name of the<br>
file and the user will be given a choice wether to accept the file or not<br>
if he\she selects yes then the client sends "okay" to the server and if he\she selects <br>
no then it sends "deny" to the server..this data i.e. "okay" or '"deny" arrivers on winsock1's<br>
local port the data is then checked using select case if its okay then "send" function<br>
is called with file address as an argument and send button and all buttons and text boxes<br>
associated with send file are disabled.<br>
If the response from client is "deny" then a msgbox is shown on server saying that the<br>
request to send the file .... as been denied..the user can send another request..or <br>
ask the client's user to accept the file using the chat module...<br>
<br>
<br>
<br>
this is called when send is clicked<br>
private sub send_click()<br>
<br>
'GET FILE NAME<br>
'using getfilename function to get the file name<br>
dim fnamea as string<br>
dim fname as string<br>
<br>
if text1.text = "" then<br>
msgbox "Please type the file name!!!", ,"Manjit"<br>
exit sub<br>
end if<br>
fname = text1.text<br>
'checking wether the file exists<br>
If Dir(fname) = "" Then<br>
MsgBox "File Does not exist Exists", ,"manjit"<br>
exit sub 'exiting sub it file does not exists<br>
end if<br>
<br>
fnamea=GetFileName(text1.text)<br>
fname=text2.text<br>
dim temp as string<br>
temp= "rqst" & fnamea<br>
<br>
'SEND<br>
<br>
winsock1.senddata temp 'sending file name of file<br>
<br>
end sub <br>
<br>
now the request is sent to the client<br>
then the server has to wait for the client's response<br>
<br>
this event is called when data arrives on winsock1<br>
<br>
Private Sub winsock1_dataarrival(ByVal bytestotal As Long)<br>
Dim response As String<br>
Winsock1.GetData response, vbString<br>
Select Case response<br>
Case "okay"<br>
send fname 'send function is called with file name as argument<br>
Case "deny"<br>
MsgBox "Your request to send the file " & fname & " has been denied", , "manjit" 'message when request is denied<br>
End Select<br>
End Sub<br>
<br>
The send function which actally sends the file<br>
<br>
<br>
Private Sub send(fname As String)<br>
Command2.Enabled = False<br>
Command3.Enabled = False<br>
Text1.Enabled = False<br>
<br>
Dim data As String<br>
Dim a As Long<br>
Dim data1 As String<br>
Dim data2 As String<br>
<br>
<br>
Open fname For Binary As #1<br>
<br>
Do While Not EOF(1)<br>
data = Input(chunk, #1)<br>
Winsock1.SendData data<br>
DoEvents<br>
Loop<br>
<br>
Winsock1.SendData "EnDf"<br>
Close #1<br>
Command2.Enabled = True<br>
Command3.Enabled = True<br>
Text1.Enabled = True<br>
<br>
End Sub<br>
<br>
<br>
'Other supporting functions: <br>
<br>
Function GetFileName(attach_str As String) As String<br>
 Dim s As Integer<br>
 Dim temp As String<br>
 s = InStr(1, attach_str, "\")<br>
 temp = attach_str<br>
 Do While s > 0<br>
  temp = Mid(temp, s + 1, Len(temp))<br>
  s = InStr(1, temp, "\")<br>
 Loop<br>
 GetFileName = temp<br>
End Function<br>
<br>
<br>
<br>
On the client side : <br>
<br>
set winsock1 to listen to a particular port say : 165<br>
and winsoc2 if you want chat too :166<br>
<br>
<br>
winsock1 is listening to port 165 and winsock2 is listening to port 166<br>
on the client side<br>
<br>
<br>
so when connection request arrives :<br>
<br>
private sub winsock1_connectionrequest(byval idrequest as long)<br>
if winsock1.state <> sckConnected then<br>
winsock1.close<br>
winsock1.accept idrequest<br>
end if<br>
end sub<br>
<br>
and:<br>
<br>
private sub winsock2_connectionrequest(byval idrequest as long)<br>
if winsock2.state <> sckConnected then<br>
winsock2.close<br>
winsock2.accept idrequest<br>
end if<br>
end sub<br>
<br>
DATA ARRIVAL:<br>
<br>
and when data arrives<br>
<br>
<br>
<br>
<br>
Private Sub winsock1_dataarrival(ByVal bytestotal As Long)<br>
<br>
Dim data As String<br>
Dim data4 As String<br>
Dim data2 As String<br>
Dim data3 As String<br>
Dim data5 As String<br>
Dim data6 As String<br>
<br>
Winsock1.GetData data, vbString<br>
<br>
data2 = Left(data, 4)<br>
Select Case data2<br>
Case "rqst" 'file request arrives<br>
<br>
data3 = Right(data, Len(data) - (4)) 'Get the file name<br>
<br>
Dim msg1 As Integer 'Stores user's selection<br>
msg1 = MsgBox(Winsock1.RemoteHost & " wants to send you file " & data3 & " accept ? ", vbYesNo, "Manjit") 'msgbox displayed<br>
<br>
<br>
If msg1 = 6 Then 'if user selects yes<br>
Winsock1.SendData "okay"<br>
cd.FileName = data3<br>
data5 = Split(data3, ".")(1)<br>
data6 = "*." & data5<br>
cd.DefaultExt = "(data6)"<br>
data4 = App.Path & "\" & data3<br>
'MsgBox data5<br>
'cd.ShowSave<br>
<br>
Open data4 For Binary As #1<br>
<br>
Else<br>
Winsock1.SendData "deny"<br>
Exit Sub<br>
End If<br>
<br>
Case "EnDf"<br>
Label1.Caption = "File revieved.Size of file : " & sz & " Kb"<br>
size=0<br>
sz=o<br>
Close #1<br>
Case Else<br>
<br>
size = size + 1<br>
Label1.Caption = size * 8 & "Kb Recieved"<br>
sz = size * 8<br>
Put #1, , data<br>
End Select<br>
End Sub<br>
<br>
This will take care of file transfer now for the chat:<br>
we will be using winsock2 for chat:<br>
<br>
On server side :<br>
<br>
WHEN SEND IS CLICKED<br>
<br>
Private Sub Command1_Click()<br>
Dim chat As String<br>
chat = Text1.Text<br>
List1.AddItem (chat)<br>
Winsock2.SendData chat<br>
<br>
End Sub<br>
<br>
when data arrives :<br>
<br>
Private Sub winsock2_dataarrival(ByVal bytestotal As Long)<br>
Dim cht As String<br>
Winsock2.GetData cht, vbString<br>
List1.AddItem (cht)<br>
<br>
End Sub<br>
<br>
<br>
the same will be on the client side also...if you want a better chat client then visit<br>
my tutorial on planet source code.com :<br>
<br>
http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=59417&lngWId=1<br>
<br>
or mail for the tutorial...<br>
<br>
I've included a copy of this tutorial in the zip file(tutorial.txt)<br>
<br>
Hope you liked it!!!..PLEASE RATE ME!!!!!!!!!<br>
<br>
<br>
<br>
</p>

