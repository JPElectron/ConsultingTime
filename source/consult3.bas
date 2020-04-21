Attribute VB_Name = "modSettings"
Option Explicit

Sub Load_SizeAndPosition(F As Form)

' Retrieve Form size and position info
' from the Window's Registry
' (uses Form IDE assigned info, if not yet saved )
Dim WasMaximized As Boolean
Dim H As Single
Dim W As Single
Dim T As Single
Dim L As Single
    
    On Error Resume Next
    ' Get Form Location on the screen
    L = GetSetting(App.Title, F.Name & "_Settings", "FormLeft", F.Left)
    T = GetSetting(App.Title, F.Name & "_Settings", "FormTop", F.Top)
    
    ' get Form Height/Width info -- move into position based on sizable property
    If F.BorderStyle = vbSizable Then
        W = GetSetting(App.Title, F.Name & "_Settings", "FormWidth", F.Width)
        H = GetSetting(App.Title, F.Name & "_Settings", "FormHeight", F.Height)
        ' move Form to last position and size
        F.Move L, T, W, H
    Else
    ' otherwise just position Form as fixed size
        F.Move L, T
    End If
    
    ' check if Form was Maximized at last exit
    WasMaximized = _
    CBool(GetSetting(App.Title, F.Name & "_Settings", "WasMaximized", False) = "True")
    ' if Form (was) maximized -- set it to that state NOW
    
    If WasMaximized = True Then
        F.WindowState = vbMaximized
    End If
End Sub

' move Form toward visible region of screen if needed
Sub Fix_Position(F As Form)
    If F.Left < 0 Then F.Left = 0
    If F.Left > Abs(Screen.Width - F.Width) Then F.Left = Abs(Screen.Width - F.Width)
    If F.Top < 0 Then F.Top = 0
    If F.Top > Abs(Screen.Height - F.Height) Then F.Top = Abs(Screen.Height - F.Height)
End Sub

Sub Save_SizeAndPosition(F As Form)

' Save Form's size and position info
' in the Window's Registry
    
    '  as long as Form  not maximized or minimized -- then
    If (F.WindowState <> vbMinimized) And _
        (F.WindowState <> vbMaximized) Then
            SaveSetting App.Title, F.Name & "_Settings", "FormLeft", F.Left
            SaveSetting App.Title, F.Name & "_Settings", "FormTop", F.Top
            SaveSetting App.Title, F.Name & "_Settings", "FormWidth", F.Width
            SaveSetting App.Title, F.Name & "_Settings", "FormHeight", F.Height
    End If
        
    ' save Boolean value -- indicates if Form was Maximized
    SaveSetting App.Title, F.Name & "_Settings", _
    "WasMaximized", (F.WindowState = vbMaximized)
End Sub

' A little something extra. You can limit all Forms to a
' (common) minimum height and width.

Sub Limit_MinimumFormSize(F As Form)
'=====================================
' these limits are in Twips, YOU DECIDE how big
Const MINWID = 4000
Const MINHGT = 3000

    If F.WindowState <> vbMinimized And _
        F.WindowState <> vbMaximized Then
        ' if Form is not fixed size
        If F.BorderStyle = vbSizable Then
            ' if below either size threshold, then apply appropriate correction
            If F.Width < MINWID Then F.Width = MINWID
            If F.Height < MINHGT Then F.Height = MINHGT
        End If
    End If
End Sub

