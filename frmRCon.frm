VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmRCon 
   Caption         =   "Quake2 Remote Console"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRCon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock udpRCON 
      Left            =   7980
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox txtCommand 
      Height          =   315
      Left            =   930
      TabIndex        =   0
      Top             =   60
      Width           =   6675
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   30
      Width           =   855
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5025
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   450
      Width           =   8475
   End
   Begin VB.Label Label1 
      Caption         =   "Command:"
      Height          =   225
      Left            =   90
      TabIndex        =   3
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "frmRCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSend_Click()
    
    Dim sPassword As String     'Server password
    
    'Close winsock and reset to server
    Me.udpRCON.Close
    Me.udpRCON.RemoteHost = frmGameMonitor.lvServers.ListItems(frmGameMonitor.lvServers.SelectedItem.Index).SubItems(1)
    Me.udpRCON.RemotePort = Val(frmGameMonitor.lvServers.ListItems(frmGameMonitor.lvServers.SelectedItem.Index).SubItems(2))
    Me.udpRCON.LocalPort = 0
    
    'Get password for server
    sPassword = frmGameMonitor.lvServers.ListItems(frmGameMonitor.lvServers.SelectedItem.Index).SubItems(6)
    
    'Send command to server
    If Me.udpRCON.State = 0 Then
        Me.udpRCON.SendData Chr(255) & Chr(255) & Chr(255) & Chr(255) & "rcon " & sPassword & " " & Me.txtCommand.Text
    End If
End Sub

Private Sub Form_Resize()
    Me.txtOutput.Width = Me.Width - 290
    Me.txtOutput.Height = Me.Height - 1200
    Me.cmdSend.Left = Me.Width - Me.cmdSend.Width - 290
    Me.txtCommand.Width = Me.Width - Me.cmdSend.Width - 1300
End Sub

Private Sub udpRCON_DataArrival(ByVal bytesTotal As Long)
    
    Dim sIncoming As String         'Incoming data
    Dim arrIncoming() As String     'Array to hold data pieces
    Dim nCount As Integer           'loop
    
    'Get data and store in string
    Me.udpRCON.GetData sIncoming
    
    'Format data into array
    arrIncoming = Split(sIncoming, vbLf)
    
    'Loop through data and display on screen
    For nCount = 0 To UBound(arrIncoming)
        Me.txtOutput.Text = Me.txtOutput.Text & arrIncoming(nCount) & vbCrLf
    Next nCount
End Sub

