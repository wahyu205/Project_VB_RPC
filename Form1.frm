VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00B46249&
   Caption         =   "Client  Kelompok 3"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00B4DCD9&
      Caption         =   "Connection"
      Height          =   1065
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4215
      Begin VB.TextBox Port 
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Text            =   "125"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox IP 
         Height          =   285
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Connect"
         Height          =   350
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Disconnect"
         Height          =   350
         Left            =   2040
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00B4DCD9&
         Caption         =   "Port"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00B4DCD9&
         Caption         =   "IP"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00B4DCD9&
      Caption         =   "Send Message/Command"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4215
      Begin VB.CommandButton Command5 
         Caption         =   "Close Server"
         Height          =   350
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   3735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Load Notepad"
         Height          =   350
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Send Message"
         Height          =   350
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
      End
   End
   Begin MSWinsockLib.Winsock kelompok3 
      Left            =   120
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() '|Connection| Button
kelompok3.Close ''Close connection
kelompok3.RemoteHost = IP 'Get IP address of PC
kelompok3.RemotePort = Port 'Port number - in Port.Text (TextBox)
kelompok3.Connect ''Set Connection
End Sub

Private Sub Command2_Click() '|Disconnection| Button
Command2.Enabled = False
kelompok3.Close ''Close connection
End Sub

Private Sub Command3_Click()
If kelompok3.State <> sckConnected Then Exit Sub ''Test connecting
kelompok3.SendData "Web site http://www.kelompok3.com"
End Sub

Private Sub Command4_Click()
If kelompok3.State <> sckConnected Then Exit Sub ''Test connecting
kelompok3.SendData "NOTEPAD"
''Shell ("setup.bat")
End Sub

Private Sub Command5_Click()
If kelompok3.State <> sckConnected Then Exit Sub ''Test connecting
kelompok3.SendData "END"
End Sub

Private Sub Form_Load()
IP.Text = kelompok3.LocalIP 'IP address of this PC
End Sub

Private Sub kelompok3_Connect()
Form1.Caption = "Anda Telah Terkoneksi"
End Sub

