VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00B46249&
   Caption         =   "Server Kelompo3"
   ClientHeight    =   1650
   ClientLeft      =   3975
   ClientTop       =   3180
   ClientWidth     =   5280
   FillColor       =   &H00B46249&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1650
   ScaleWidth      =   5280
   Begin VB.TextBox Text1 
      Height          =   1335
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin MSWinsockLib.Winsock kelompok3 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Server
Private Sub Form_Load()
''kelompok3 - Winsock name
Form1.Visible = True
Do
    If kelompok3.State <> sckConnected And kelompok3.State <> sckListening Then 'Is the connection available or do we listen to the port?
        kelompok3.Close 'All the connections are switched off
        kelompok3.Listen ''Listen port 125
    End If
    DoEvents
Loop
End Sub
Private Sub kelompok3_ConnectionRequest(ByVal requestID As Long) 'Request for connection
kelompok3.Close 'Listen close
kelompok3.Accept requestID 'Let's tap a Client with the number of his request.
End Sub

Private Sub kelompok3_DataArrival(ByVal bytesTotal As Long)
Dim Data As String 'Variable Data
kelompok3.GetData Data 'It will contain the received data

Text1.Text = Data
If Data = "END" Then End ''If the text command END is received, you shall finish the work of server
If Data = "NOTEPAD" Then Shell ("notepad.exe") ''To start up the application Notepad on the side of the server
End Sub

