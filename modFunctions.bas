Attribute VB_Name = "modFunctions"
Option Explicit

Public Function mfGetServers(arrServerInfo() As String) As String
    
    Dim nCount As Integer       'Counter
    Dim sServerInfo As String   'Array to hold server info
    Dim arrTemp() As String     'Temp array
    
    'Clear out the array
    ReDim arrServerInfo(19, 4)
    
    'Loop through the 20 registry slots and get server information
    For nCount = 0 To 19
        'Get settings from registry
        sServerInfo = GetSetting("GameMonitor", "Servers", nCount, "")
        If Not sServerInfo = "" Then
            'Store results into an array to be returned
            arrTemp = Split(sServerInfo, "|")
            arrServerInfo(nCount, 0) = arrTemp(0)
            arrServerInfo(nCount, 1) = arrTemp(1)
            arrServerInfo(nCount, 2) = arrTemp(2)
            arrServerInfo(nCount, 3) = arrTemp(3)
            arrServerInfo(nCount, 4) = arrTemp(4)
        End If
    Next nCount
    
End Function

