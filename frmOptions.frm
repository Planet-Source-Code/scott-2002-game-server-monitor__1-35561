VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   90
      TabIndex        =   5
      Top             =   1710
      Width           =   3885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1890
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
      Height          =   1035
      Left            =   240
      TabIndex        =   3
      Top             =   420
      Width           =   3585
      Begin VB.TextBox txtSettings 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1350
         TabIndex        =   0
         Text            =   "30"
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "seconds"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   675
      End
      Begin VB.Label lblServer 
         Caption         =   "Refresh Rate:"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   480
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2940
      TabIndex        =   2
      Top             =   1890
      Width           =   1065
   End
   Begin MSComctlLib.TabStrip tbServers 
      Height          =   1545
      Left            =   90
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   90
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   2725
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Settings"
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
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    'Unload the form
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'Save the setting
    If mfSaveRecord Then
        'If successful, unload form
        Unload Me
    End If
End Sub

Private Function mfSaveRecord() As Boolean
On Error GoTo mfSaveRecord_Error

    'Save the setting to the registry
    Call SaveSetting("GameMonitor", "Settings", "Refresh", Val(Me.txtSettings.Text))
    
'Error Handler
mfSaveRecord_Error:
    Select Case Err.Number
        Case 0
            mfSaveRecord = True
        Case Else
            mfSaveRecord = False
            MsgBox Err.Description, vbOKOnly + vbExclamation, "Game Monitor"
            Err.Clear
    End Select
End Function

Private Sub Form_Load()
    'Get store settings
    Call mpGetSettings
End Sub

Private Sub mpGetSettings()
    Dim nRefresh As Integer 'Refresh rate
    
    'Get the saved refresh rate
    nRefresh = GetSetting("GameMonitor", "Settings", "Refresh", 30)
    
    'Force nRefresh to be between 3 and 60
    If nRefresh < 3 Or nRefresh > 60 Then nRefresh = 30

    Me.txtSettings.Text = nRefresh
End Sub

Private Sub txtSettings_LostFocus()
    'Validate entry.  Keep at or above 3 settings to
    'prevent spamming the server
    If Val(Me.txtSettings.Text) < 3 Or Val(Me.txtSettings.Text) > 60 Then
        MsgBox "The refresh rate must be between 3 and 60.", vbOKOnly + vbInformation, "Invalid Entry"
        'Return focus to textbox
        Me.txtSettings.SelStart = 0
        Me.txtSettings.SelLength = Len(Me.txtSettings.Text)
        Me.txtSettings.SetFocus
    End If
End Sub
