VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Manager"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4170
      TabIndex        =   6
      Top             =   3330
      Width           =   1065
   End
   Begin VB.Frame fraAddServer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   300
      TabIndex        =   9
      Top             =   480
      Width           =   4755
      Begin VB.TextBox txtServer 
         Height          =   315
         Index           =   3
         Left            =   1590
         TabIndex        =   4
         Top             =   1890
         Width           =   2955
      End
      Begin VB.TextBox txtServer 
         Height          =   315
         Index           =   0
         Left            =   1590
         TabIndex        =   0
         Top             =   330
         Width           =   2955
      End
      Begin VB.ComboBox cboGameType 
         Height          =   315
         Left            =   1590
         TabIndex        =   1
         Top             =   720
         Width           =   2985
      End
      Begin VB.TextBox txtServer 
         Height          =   315
         Index           =   2
         Left            =   1590
         TabIndex        =   3
         Top             =   1500
         Width           =   2955
      End
      Begin VB.TextBox txtServer 
         Height          =   315
         Index           =   1
         Left            =   1590
         TabIndex        =   2
         Top             =   1110
         Width           =   2955
      End
      Begin VB.Label lblServer 
         Caption         =   "RCon Password:"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   14
         Top             =   1950
         Width           =   1245
      End
      Begin VB.Label lblServer 
         Caption         =   "Server Name:"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Game Type:"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   795
         Width           =   1455
      End
      Begin VB.Label lblServer 
         Caption         =   "Game Port:"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label lblServer 
         Caption         =   "Internet Address:"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   1185
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3060
      TabIndex        =   5
      Top             =   3330
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   60
      Left            =   150
      TabIndex        =   8
      Top             =   3150
      Width           =   5085
   End
   Begin MSComctlLib.TabStrip tbServers 
      Height          =   2955
      Left            =   150
      TabIndex        =   7
      Top             =   150
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   5212
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Server Properties"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bytMode As Byte          'Byte variable to store if add server (0) or server # if edit

Private Sub cmdCancel_Click()
    'Close the form
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'Call save record routine
    If mfSaveRecord Then
        'Unload form if successfull
        Unload Me
    Else
        'Too many servers error
        MsgBox "Directory is full.  Please remove a server first.", vbOKOnly + vbInformation + vbApplicationModal, "Directory Full"
    End If
End Sub

Private Function mfSaveRecord() As Boolean

    'Record the server into the registry
    Dim nCount As Integer    'Loop through servers
    
    'If bytMode = 0 then new server
    'else index of selected server to edit
    If bytMode = 0 Then
    
        'Search for an emtpy slot
        For nCount = 0 To 19
            'Open each registry entry and see if server name is blank
            If GetSetting("GameMonitor", "Servers", nCount) = "" Then
                'Save new server in empty slot
                Call SaveSetting("GameMonitor", "Servers", nCount, Me.txtServer(0).Text & "|" & Me.cboGameType.Text & "|" & Me.txtServer(1).Text & "|" & Me.txtServer(2).Text & "|" & Me.txtServer(3).Text)
                'Set bytmode to empty slot#
                bytMode = nCount
                mfSaveRecord = True
                Exit Function
            End If
        Next nCount
        'Failed to save record. No empty slots
        mfSaveRecord = False
        
    'Else update proper registry
    Else
        'Save edited server info into it's slot
        Call SaveSetting("GameMonitor", "Servers", bytMode - 1, Me.txtServer(0).Text & "|" & Me.cboGameType.Text & "|" & Me.txtServer(1).Text & "|" & Me.txtServer(2).Text & "|" & Me.txtServer(3).Text)
        mfSaveRecord = True
    End If
   
End Function

Private Sub Form_Load()
    'Load types of games
    '*** ONLY QUAKE2 IS WORKING
    Call mpLoadGameTypes
    
    'If bytMode <> 0 then it's an edit
    'Display server info
    If Not bytMode = 0 Then
        Call mpShowServerInfo
    End If
End Sub

Private Sub mpLoadGameTypes()
    Me.cboGameType.AddItem "Half-Life"
    Me.cboGameType.AddItem "Quake"
    Me.cboGameType.AddItem "Quake2"
    Me.cboGameType.AddItem "Quake3"
    Me.cboGameType.AddItem "Tribes"
    Me.cboGameType.AddItem "Unreal Tournament"
End Sub

Private Sub mpShowServerInfo()
    'Display selected server info for editing
    Me.txtServer(0).Text = frmGameMonitor.lvServers.ListItems(frmGameMonitor.lvServers.SelectedItem.Index).Text
    Me.cboGameType.Text = frmGameMonitor.lvServers.ListItems(frmGameMonitor.lvServers.SelectedItem.Index).SubItems(3)
    Me.txtServer(1).Text = frmGameMonitor.lvServers.ListItems(frmGameMonitor.lvServers.SelectedItem.Index).SubItems(1)
    Me.txtServer(2).Text = frmGameMonitor.lvServers.ListItems(frmGameMonitor.lvServers.SelectedItem.Index).SubItems(2)
    Me.txtServer(3).Text = frmGameMonitor.lvServers.ListItems(frmGameMonitor.lvServers.SelectedItem.Index).SubItems(6)
End Sub
