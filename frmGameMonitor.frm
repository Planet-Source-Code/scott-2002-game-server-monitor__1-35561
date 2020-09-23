VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmGameMonitor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game Monitor"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6900
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGameMonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   7095
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   18
            MinWidth        =   18
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8149
            Text            =   "Not Connected"
            TextSave        =   "Not Connected"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1931
            MinWidth        =   1940
            TextSave        =   "6/7/02"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1931
            MinWidth        =   1940
            TextSave        =   "12:15 AM"
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
   Begin MSWinsockLib.Winsock udpQuake2 
      Left            =   6990
      Top             =   1590
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Timer tmrPoll 
      Interval        =   5000
      Left            =   6960
      Top             =   1140
   End
   Begin MSComctlLib.ListView lvServer 
      Height          =   5175
      Left            =   3240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   9128
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Rule"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ListView lvPlayers 
      Height          =   5175
      Left            =   30
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1920
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   9128
      SortKey         =   2
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Player"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Score"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ping"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ScoreSort"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "PingSort"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvServers 
      Height          =   1845
      Left            =   30
      TabIndex        =   3
      Top             =   60
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   3254
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2593
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Address"
         Object.Width           =   2327
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Port"
         Object.Width           =   1191
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Game Type"
         Object.Width           =   1983
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Players"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Map"
         Object.Width           =   1905
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Password"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   -720
      X2              =   7320
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileServers 
         Caption         =   "&Add Game Server"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMonitor 
         Caption         =   "&Start Monitor"
         Index           =   0
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileMonitor 
         Caption         =   "S&top Monitor"
         Index           =   1
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsRCON 
         Caption         =   "&Remote Console"
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmGameMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bReceived As Boolean        ' Track if server has responded to last request

Private Sub cmdRCON_Click()
    'Display the Remote Control console
    frmRCon.Show
End Sub

Private Sub Form_Load()
    
    'Get refresh rate
    Call mpGetSettings
    
    'Refresh the screen
    Call mpLoadServers
    
    'Default to 1st server
    If Not Me.lvServers.ListItems.Count = 0 Then
        Me.lvServers.ListItems(1).Selected = True
    End If

    'Default player sort order by score
    Call mpSortColumn(4, True)

    'Refresh the screen
    Call mpRefreshScreen
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Terminate the application
    End
End Sub

Private Sub lvPlayers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    'Needed local variables
    Static bytLastColumn As Byte    'Store current sort order
    Static bSortOrder(2) As Boolean    'Store last sorted column
    
    'If current column is same as last invert bSortOrder
    bSortOrder(ColumnHeader.Index - 2) = Not bSortOrder(ColumnHeader.Index - 2)
    
    'Select clicked column
    Select Case ColumnHeader.Index
        'Player Name
        Case 2
            Call mpSortColumn(1, bSortOrder(ColumnHeader.Index - 2))
        'Player Score
        Case 3
            Call mpSortColumn(4, bSortOrder(ColumnHeader.Index - 2))
        'Player Ping
        Case 4
            Call mpSortColumn(5, bSortOrder(ColumnHeader.Index - 2))
    End Select
        
    'Remember the last column click
    bytLastColumn = ColumnHeader.Index
End Sub


Private Sub lvServers_Click()
    
    'Clear the listboxes
    Me.lvPlayers.ListItems.Clear
    Me.lvServer.ListItems.Clear
    
    'Refresh the screen
    Call mpRefreshScreen
    
End Sub

Private Sub lvServers_DblClick()

    Dim nCurrentIndex As Integer        'Current selected server
    
    'Set server form to edit mode
    nCurrentIndex = Me.lvServers.SelectedItem.Index
    frmServer.bytMode = nCurrentIndex
    
    'Display the form
    frmServer.Show vbModal, Me
    
    'Reload the servers
    Call mpLoadServers
    
    'Refresh Screen
    If Not Me.lvServers.ListItems.Count = 0 Then
        Me.lvServers.ListItems(nCurrentIndex).Selected = True
        Call mpRequestStatus
    End If
End Sub

Private Sub lvServers_KeyDown(KeyCode As Integer, Shift As Integer)
    'Delete key pressed
    If KeyCode = 46 Then
        'Call delete server routine
        Call mpDeleteServer
    End If
End Sub

Private Sub mnuFile_Click()
    'Display proper start/stop options based on timer enabled
    Me.mnuFileMonitor(0).Enabled = Not Me.tmrPoll.Enabled
    Me.mnuFileMonitor(1).Enabled = Me.tmrPoll.Enabled
End Sub

Private Sub mnuFileMonitor_Click(Index As Integer)
    'Reverse status of timer
    Me.tmrPoll.Enabled = Not Me.tmrPoll.Enabled
    
    'Refresh screen now if enabled
    If Me.tmrPoll.Enabled Then
        Call mpRefreshScreen
    End If
End Sub

Private Sub mnuFileServers_Click()
    'Open server manager form
    frmServer.bytMode = 0
    frmServer.Show vbModal, Me
    'Refresh the screen
    Call mpLoadServers
    
    'Refresh the screen
    Call mpRefreshScreen
    
End Sub

Private Sub mnuTools_Click()
    'Only allow RCON if server is selected
    Me.mnuToolsRCON.Enabled = False
    If Not Me.lvServers.ListItems.Count = 0 Then
        If Not Me.lvServers.SelectedItem.Index = 0 Then
            Me.mnuToolsRCON.Enabled = True
        End If
    End If
End Sub

Private Sub mnuToolsOptions_Click()
    'Display the Options form
    frmOptions.Show vbModal, Me
    
    'Reset the refresh rate
    Call mpGetSettings
    
End Sub

Private Sub mnuToolsRCON_Click()
    'Display the RCON form
    frmRCon.Show
End Sub

Private Sub tmrPoll_Timer()
    'Refresh the screen
    Call mpRefreshScreen
End Sub

Private Sub udpQuake2_DataArrival(ByVal bytesTotal As Long)
    
    'Declare local needed variables
    Dim sIncoming As String         'String of incoming text
    Dim arrIncoming() As String     'Array to hold each line of server text
    Dim nMaxPlayers As Integer      'Total players allowed
    Dim nTotalPlayers As Integer    'Number of players on server
    
    'Make bReceived true so we know server has responded to request
    bReceived = True
    
    'Get data from buffer
    Me.udpQuake2.GetData sIncoming, vbString
    
    'Make sure data is valid.  Should be over 10 characters at least
    If Not Len(Trim(sIncoming)) < 10 Then
        'Call function to convert gobbly-gook of text into array
        'for easier management
        Call mpFormatData(sIncoming, arrIncoming)
        
        If UBound(arrIncoming) > 0 Then
            'Call function to display server stats in listbox
            nMaxPlayers = mfDisplayStatInfo(arrIncoming)
            'Call function to display player stats in listbox
            nTotalPlayers = mfDisplayPlayerInfo(arrIncoming)
        End If
    End If
    
    Me.lvServers.ListItems(Me.lvServers.SelectedItem.Index).SubItems(4) = nTotalPlayers & "/" & nMaxPlayers
    'Show last refresh time on status bar
    Me.sbStatus.Panels(2).Text = "Last Refresh: " & Format(Now, "hh:nn:ss AM/PM")
End Sub

Private Sub mpLoadServers()

    Dim arrServers() As String      'Array to hold server info
    Dim nCount As Integer           'Loop
    
    'Clear server list
    Me.lvServers.ListItems.Clear

    'Get server information
    Call mfGetServers(arrServers)
    
    'Loop through returned array and display information
    For nCount = 0 To 19
        'Make sure it's vaild
        If Not arrServers(nCount, 0) = "" Then
            Me.lvServers.ListItems.Add , , arrServers(nCount, 0)
            Me.lvServers.ListItems(Me.lvServers.ListItems.Count).SubItems(1) = arrServers(nCount, 2)
            Me.lvServers.ListItems(Me.lvServers.ListItems.Count).SubItems(2) = arrServers(nCount, 3)
            Me.lvServers.ListItems(Me.lvServers.ListItems.Count).SubItems(3) = arrServers(nCount, 1)
            Me.lvServers.ListItems(Me.lvServers.ListItems.Count).SubItems(6) = arrServers(nCount, 4)
        End If
    Next nCount

End Sub

Private Sub mpRefreshScreen()

    'Refresh Screen
    If Not Me.lvServers.ListItems.Count = 0 Then
        If Not Me.lvServers.SelectedItem.Index = 0 Then
            Call mpRequestStatus
        End If
    End If
End Sub

Private Sub mpRequestStatus()

    'This sub sends the request for server info
    
    Dim sCommand As String      'Command to send to server
    
    
    'Select gametype and get correct command to request info
    Select Case Me.lvServers.ListItems(Me.lvServers.SelectedItem.Index).SubItems(3)
        Case "Half-Life"
        Case "Quake"
        Case "Quake2"
            sCommand = Chr(255) & Chr(255) & Chr(255) & Chr(255) & "status"
        Case "Quake3"
            sCommand = Chr(255) & Chr(255) & Chr(255) & Chr(255) & "getstatus"
        Case "Tribes"
        Case "Unreal Tournament"
            sCommand = "\info\\players\"
    End Select
    'This function calls command to request data
    Call mpSendCommand(sCommand)
End Sub

Private Sub mpSendCommand(ByVal sData As String)
On Error GoTo msSendCommand_Error
    
    'Setup winsock control for udp message
    Me.udpQuake2.Close
    Me.udpQuake2.RemoteHost = Me.lvServers.ListItems(Me.lvServers.SelectedItem.Index).SubItems(1)
    Me.udpQuake2.RemotePort = Val(Me.lvServers.ListItems(Me.lvServers.SelectedItem.Index).SubItems(2))
    Me.udpQuake2.LocalPort = 0
    
    'Let the socket close
    DoEvents
    
    'If state is "No Connection"
    If Me.udpQuake2.State = 0 Then
        'Send Data
        Me.udpQuake2.SendData sData
    End If

    'if bReceived is false, server hasn't responded since last request
    If bReceived = False Then Me.sbStatus.Panels(2).Text = "No response from server"
    
    'Set bReceived to false to monitor if server has replied
    bReceived = False
    
'Error Handler
msSendCommand_Error:
    Select Case Err.Number
        'No Error
        Case 0
        'Bad IP
        Case 10014
            Me.sbStatus.Panels(2).Text = "Unable to connect to server"
        'Other error
        Case Else
            Me.sbStatus.Panels(2).Text = Err.Description
    End Select
    'Clear error if exists, where done with it
    Err.Clear

End Sub

Private Sub mpFormatData(ByVal sIncoming As String, arrIncoming() As String)

    'Take incoming string of data and convert into array
    'The linefeed is the breaking point for each set of data
    arrIncoming = Split(sIncoming, vbLf)
    
End Sub

Private Function mfDisplayStatInfo(arrIncoming() As String) As Integer
    
    'Declare local needed variables
    Dim arrTemp() As String     'Array to hold sub-breakdown of line data
    Dim nCount As Integer       'Count for looping through arrays
    Dim nPlayers As Integer     'Number of players
    Dim nMaxPlayers As Integer  'Number of allowed players
    
    'Server stat data is stored in arrIncoming(1)
    If Not Len(arrIncoming(1)) = 0 Then
    
        'Data exists.  Split line of data down to individuals pieces
        arrTemp = Split(arrIncoming(1), "\")
        
        'Clear the stat listbox
        Me.lvServer.ListItems.Clear
        'Turn off sorting
        Me.lvServer.Sorted = False
        
        'Display name of current map
        Me.lvServers.ListItems(Me.lvServers.SelectedItem.Index).SubItems(5) = arrTemp(4)
        
        'Loop through the array of line data and populate the stat listbox
        For nCount = LBound(arrTemp) To UBound(arrTemp) - 1 Step 2
            'Add Server rule to listbox
            Me.lvServer.ListItems.Add , , arrTemp(nCount + 1)
            'Add server rule value to listbox
            Me.lvServer.ListItems(Me.lvServer.ListItems.Count).SubItems(1) = arrTemp(nCount + 2)
            
            'Get MaxClients setting
            If Not nCount = 0 Then
                
                If Trim(UCase(arrTemp(nCount - 1))) = "MAXCLIENTS" Or Trim(UCase(arrTemp(nCount - 1))) = "SV_MAXCLIENTS" Then
                    nMaxPlayers = Val(arrTemp(nCount))
                End If
            End If
        Next nCount
        
        'Sort the server stat listbox
        Me.lvServer.SortKey = 0
        Me.lvServer.Sorted = True
        Me.lvServer.SortOrder = lvwAscending
        Me.lvServer.ListItems(1).Selected = False
    End If
    
    'Return allowed number of players
    mfDisplayStatInfo = nMaxPlayers
    
End Function

Private Function mfDisplayPlayerInfo(arrIncoming() As String) As Integer
    
    'Declare local needed variables
    Dim arrTemp() As String     'Array to hold sub-breakdown of line data
    Dim nCount As Integer       'Count for looping through arrays
    
    'Data for players exists in arrIncoming(3+)
    If Not UBound(arrIncoming) < 3 Then
        
        'Data exists.  Clear listbox
        Me.lvPlayers.ListItems.Clear
        'Turn off sorting
        Me.lvPlayers.Sorted = False
        
        'Loop through remaining arrincoming and get player data
        For nCount = LBound(arrIncoming) + 2 To UBound(arrIncoming) - 1
        
            'Split each line into individual pieces
            arrTemp = Split(arrIncoming(nCount), " ")
            
            'Add Player ID (not important)
            Me.lvPlayers.ListItems.Add , , nCount
            'Add Player Name
            Me.lvPlayers.ListItems(Me.lvPlayers.ListItems.Count).SubItems(1) = mfStripQuotes(arrTemp(2))
            'Add Player Score
            Me.lvPlayers.ListItems(Me.lvPlayers.ListItems.Count).SubItems(2) = arrTemp(0)
            'Add Player Ping
            Me.lvPlayers.ListItems(Me.lvPlayers.ListItems.Count).SubItems(3) = arrTemp(1)
            'Add sort order (score)
            Me.lvPlayers.ListItems(Me.lvPlayers.ListItems.Count).SubItems(4) = String(3 - Len(arrTemp(0)), "_") & arrTemp(0)
            'Add sort order (ping)
            Me.lvPlayers.ListItems(Me.lvPlayers.ListItems.Count).SubItems(5) = String(3 - Len(arrTemp(1)), "_") & arrTemp(1)
        Next nCount
        
        'Sort by player score
        Me.lvPlayers.Sorted = True
            
        'Return number of players
        mfDisplayPlayerInfo = nCount - 2
    Else
        'Return number of players
        mfDisplayPlayerInfo = 0
    End If
    
    
End Function

Private Sub mpSortColumn(ByVal bytColumn As Byte, ByVal bSortOrder As Boolean)
    
    'Function sorts column
    
    'Set column to sort by
    Me.lvPlayers.SortKey = bytColumn
    
    'Set sort order - True = Descending, False = Ascending
    If bSortOrder Then
        Me.lvPlayers.SortOrder = lvwDescending
    Else
        Me.lvPlayers.SortOrder = lvwAscending
    End If
    
    Me.lvPlayers.Sorted = True
End Sub

Private Function mfStripQuotes(ByVal sString As String) As String

    'This function removes quotes from string sString
    mfStripQuotes = Replace(sString, Chr(34), "")
    
End Function

Private Sub mpGetSettings()
    
    Dim nRefresh As Integer 'Refresh rate
    
    'Get the saved refresh rate
    nRefresh = GetSetting("GameMonitor", "Settings", "Refresh", 30)
    
    'Force nRefresh to be between 3 and 60
    If nRefresh < 3 Or nRefresh > 60 Then nRefresh = 30
    
    'Set timer refresh rate
    Me.tmrPoll.Interval = nRefresh * 1000
End Sub

Private Sub mpDeleteServer()
    If MsgBox("Are you sure you want to delete this server?", vbYesNo + vbQuestion + vbApplicationModal, "Confirm Delete") = vbYes Then
        'Delete server
        Call SaveSetting("GameMonitor", "Servers", Me.lvServers.SelectedItem.Index - 1, "")
        'Reload servers
        Call mpLoadServers
        'Refresh the screen
        Call mpRefreshScreen
    End If
End Sub
