VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.Ocx"
Begin VB.Form Form1 
   Caption         =   "Consulting Time"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2775
   Icon            =   "consult.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   3000
      TabIndex        =   23
      Top             =   6000
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Frame Frame2 
      Caption         =   "Notification"
      Height          =   1575
      Left            =   120
      TabIndex        =   16
      Top             =   7680
      Width           =   2535
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Wave file..."
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Beep"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Text            =   "10"
         Top             =   240
         Width           =   375
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Notify every"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "minutes"
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   6480
      Width           =   2535
      Begin VB.CheckBox Check1 
         Caption         =   "Always on top"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Warn for multiple timers"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Disallow multiple timers"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Custom clients.ini filename:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.ComboBox Client 
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   2535
   End
   Begin VB.ComboBox Client 
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   2535
   End
   Begin VB.ComboBox Client 
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   2535
   End
   Begin VB.ComboBox Client 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   2535
   End
   Begin VB.ComboBox Client 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3600
      Top             =   5400
   End
   Begin VB.TextBox Time 
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Time 
      Height          =   285
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Time 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Time 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox Time 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   2400
      Picture         =   "consult.frx":0442
      Stretch         =   -1  'True
      Top             =   6015
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   600
      Picture         =   "consult.frx":06EC
      Top             =   6000
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "consult.frx":0A2E
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   240
   End
   Begin VB.Image Go 
      Height          =   255
      Index           =   4
      Left            =   1800
      Picture         =   "consult.frx":0CD8
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Stop 
      Height          =   255
      Index           =   4
      Left            =   2280
      Picture         =   "consult.frx":108E
      Top             =   5400
      Width           =   255
   End
   Begin VB.Image Go 
      Height          =   255
      Index           =   3
      Left            =   1800
      Picture         =   "consult.frx":1444
      Top             =   4200
      Width           =   255
   End
   Begin VB.Image Stop 
      Height          =   255
      Index           =   3
      Left            =   2280
      Picture         =   "consult.frx":17FA
      Top             =   4200
      Width           =   255
   End
   Begin VB.Image Go 
      Height          =   255
      Index           =   2
      Left            =   1800
      Picture         =   "consult.frx":1BB0
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image Stop 
      Height          =   255
      Index           =   2
      Left            =   2280
      Picture         =   "consult.frx":1F66
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image Go 
      Height          =   255
      Index           =   1
      Left            =   1800
      Picture         =   "consult.frx":231C
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Stop 
      Height          =   255
      Index           =   1
      Left            =   2280
      Picture         =   "consult.frx":26D2
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Go 
      Height          =   255
      Index           =   0
      Left            =   1800
      Picture         =   "consult.frx":2A88
      Top             =   600
      Width           =   255
   End
   Begin VB.Image Stop 
      Height          =   255
      Index           =   0
      Left            =   2280
      Picture         =   "consult.frx":2E3E
      Top             =   600
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AlwaysOnTop API
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'Open a file with its associated document API
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
    As Long) As Long

Public LastState As Integer

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_MOVE = &HF010&
Private Const SC_RESTORE = &HF120&
Private Const SC_SIZE = &HF000&

Dim TickTock(0 To 4) As elapsedtime

'True or false value referring to whether a give timer is running or not, 5 element array
Dim Timing(0 To 4) As Boolean
Dim LastSave
Dim CurrentPath
Dim ClientPath
Dim NotDT
Dim ClientErr

Private Sub Check1_Click()
On Error GoTo quit
'Always on top
If Check1.Value = 1 Then SetWindowPos hwnd, -1, 0, 0, 0, 0, &H1 Or &H2 Or &H10 Or &H40
'Normal
If Check1.Value = 0 Then SetWindowPos hwnd, -2, 0, 0, 0, 0, &H1 Or &H2 Or &H10 Or &H40

Open "settings.ini" For Output As #4

Print #4, Text1.Text
Print #4, Check1.Value
Print #4, Check2.Value
Print #4, Check3.Value
Print #4, Text2.Text
Print #4, Check4.Value
Print #4, Option1.Value
Print #4, Option2.Value
Print #4, Text3.Text

Close #4

quit:
End Sub

Private Sub Check2_Click()
On Error GoTo quit
Open "settings.ini" For Output As #4
Print #4, Text1.Text
Print #4, Check1.Value
Print #4, Check2.Value
Print #4, Check3.Value
Print #4, Text2.Text
Print #4, Check4.Value
Print #4, Option1.Value
Print #4, Option2.Value
Print #4, Text3.Text

Close #4
quit:
End Sub

Private Sub Check3_Click()
On Error GoTo quit
Open "settings.ini" For Output As #4
Print #4, Text1.Text
Print #4, Check1.Value
Print #4, Check2.Value
Print #4, Check3.Value
Print #4, Text2.Text
Print #4, Check4.Value
Print #4, Option1.Value
Print #4, Option2.Value
Print #4, Text3.Text

Close #4
quit:
End Sub

Private Sub Check4_Click()
On Error GoTo quit
Open "settings.ini" For Output As #4
Print #4, Text1.Text
Print #4, Check1.Value
Print #4, Check2.Value
Print #4, Check3.Value
Print #4, Text2.Text
Print #4, Check4.Value
Print #4, Option1.Value
Print #4, Option2.Value
Print #4, Text3.Text

Close #4
quit:
End Sub

Private Sub Form_Load()
On Error Resume Next
Load Form2
Load Form3
modSettings.Load_SizeAndPosition Me
'CurrentPath = App.Path
Open "settings.ini" For Input As #1
'If Err Then GoTo Skip
Input #1, ClientPath
Text1.Text = ClientPath
Input #1, T
Check1.Value = T
Input #1, T
Check2.Value = T
Input #1, T
Check3.Value = T
Input #1, T
Text2.Text = T
Input #1, T
Check4.Value = T
Input #1, T
Option1.Value = T
Input #1, T
Option2.Value = T
Input #1, T
Text3.Text = T
Close
'Check1_Click
If ClientPath Like "*\" = True And Len(ClientPath) > 3 Then ClientPath = Left(ClientPath, Len(ClientPath) - 1)
Text1.Text = ClientPath

Skip:

Err = ""
Open Text1.Text For Input As #1
If Err Then
ChDir (App.Path)
If Dir$("clients.ini") = "" Then
Open "clients.ini" For Output As #1
Close
End If
Open "clients.ini" For Input As #1
End If
'Read data from client file.
'This will loop until the EOF; End Of File
Do While Not EOF(1)
'Each 'Input' statement reads one line of the file and the moves
'to the next line.  One name is stored per line, so each name is
'in turn read into the 'cname' variable
Input #1, cname
'Loop through all 5 combo boxes and add the 'cname' variable to
'the drop down list.
For i = 0 To 4
Client(i).AddItem cname
Next i
Loop

'Close the file.
Close
skipclient:
'ChDir (CurrentPath)

'Ignore system form colors so that the images used as buttons
'have appropriately colored backgrounds.
Form1.BackColor = &H8000000F

'This loop initalizes each of the 5 timers.
For i = 0 To 4

'Set all elapsed time fields to 0
TickTock(i).Hour = 0
TickTock(i).Minute = 0
TickTock(i).Second = 0
TickTock(i).Dtee = 0

'Put the first digit (hours) into the text box
Time(i).Text = TickTock(i).Hour & ":"

'This block will check if the minutes digit is a single number,
'or a double digit.  If it is a single digit, it will add a
'zero to the beginning for display purposes, and add it to the
'time display box.
temp = Int(TickTock(i).Minute)
If Len(temp) = 1 Then
Time(i).Text = Time(i).Text & "0" & TickTock(i).Minute & ":"
Else
Time(i).Text = Time(i).Text & TickTock(i).Minute & ":"
End If

'Same as above, for the seconds.
temp = Int(TickTock(i).Second)
If Len(temp) = 1 Then
Time(i).Text = Time(i).Text & "0" & TickTock(i).Second
Else
Time(i).Text = Time(i).Text & TickTock(i).Second
End If

'All timers are initially off, 0 = False
Timing(i) = 0
Next i

    If WindowState = vbMinimized Then
        LastState = vbNormal
    Else
        LastState = WindowState
    End If

    AddToTray Me ', mnuTray

    SetTrayTip "Consulting Timer"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
For i = 0 To 4
If Timing(i) = True Then Running = True
Next i

If Running = True Then
    Response = MsgBox("NOTE: There are still timers running." & vbCrLf & vbCrLf & "Do you wish to close?", vbOKCancel, "Close")
End If

If LastSave = "" Then
    Response = MsgBox("Data has NOT been saved, do you wish to close?", vbOKCancel, "Close")
End If

If Response = vbCancel Then Cancel = True

End Sub

Private Sub Form_Resize()
    'SetTrayMenuItems WindowState

    If WindowState <> vbMinimized Then _
        LastState = WindowState
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
Unload Form3
modSettings.Save_SizeAndPosition Me
End Sub

Private Sub Go_Click(Index As Integer)
On Error Resume Next
'This is a temporary variable array that will hold the currently
'entered (or selected) name in the text boxes while they are
'being refreshed, and the replace it when the refresh finishes.
Dim tn(0 To 4)

'Check if a client is selected.
If Client(Index) = "" Then
    'If not, give an error.
    MsgBox "Please enter a client name.", vbOKOnly, "Warning"
Else
    If Check2.Value = 1 Then
        For i = 0 To 4
        If Timing(i) = True Then tOn = True
        Next i
        If tOn = True Then
            Response = MsgBox("Start multiple timers?", vbOKCancel, "Warning")
            If Response = vbCancel Then Exit Sub
        End If
    End If
    
    If Check3.Value = 1 Then
        For i = 0 To 4
        If Timing(i) = True Then tOn = True
        Next i
        If tOn = True Then
            For k = 0 To 4
            Timing(k) = False
            Next k
        End If
End If

RemoveFromTray
Form1.Icon = Form2.Icon
    If WindowState = vbMinimized Then
        LastState = vbNormal
    Else
        LastState = WindowState
    End If

    AddToTray Me

    SetTrayTip "Consulting Time"

'Open the clients file
Err = ""
Open Text1.Text For Input As #1
If Err Then Open "clients.ini" For Input As #1

'Place the selected client's name in a temp. variable
newname = Client(Index).Text

'Assume initially that the name is new; not found in
'the current clients file.
found = False
    'Loop through the entire file, checking if any of the names
    'match the name currently entered.
    Do While Not EOF(1)
    Input #1, cname
    'If a match is found, set Found = True
    If cname = Client(Index).Text Then found = True
    Loop
Close

'If the name already exists in the client file, skip
'the section immediately below, which would add a new
'name to the file.
If found = True Then GoTo relist

'Open the client file for nondestructive write.
Open "clients.ini" For Append As #1

'Add the current name to the list.
Print #1, newname
Close
ChDir (App.Path)

relist:
    
    'Store each selected client in the temp. 'tn' var, then
    'clear the box.  Then reenter the boxes contents from
    'the 'tn' var
    For i = 0 To 4
    tn(i) = Client(i).Text
    Client(i).Clear
    Client(i).Text = tn(i)
    Next i

'Open clients for read.
Open "clients.ini" For Input As #1

    'Loop through each client entry
    Do While Not EOF(1)
    Input #1, cname
    'Loop through all the timer lists
    For i = 0 To 4
    'Add the name to each one
    Client(i).AddItem cname
    Next i
    Loop
    
Close

'Assume the timer is currently OFF: (This eliminates errors from clicking start twice)
If Timing(Index) = False Then

'Find the current hour
markhour = Hour(Now)
'Make sure it's in proper two digit format
If Len(markhour) = 1 Then markhour = "0" & markhour
'Add it's comma separated value to the StopHour set
TickTock(Index).StopHours = TickTock(Index).StopHours & markhour & ","

'Repeat for minutes and seconds
markminute = Minute(Now)
If Len(markminute) = 1 Then markminute = "0" & markminute
TickTock(Index).StopMinutes = TickTock(Index).StopMinutes & markminute & ","
marksecond = Second(Now)
If Len(marksecond) = 1 Then marksecond = "0" & marksecond
TickTock(Index).StopSeconds = TickTock(Index).StopSeconds & marksecond & ","
End If

'Now start the timer.
LastSave = ""
Timing(Index) = True

End If
End Sub

Private Sub Image1_Click()
LastSave = Hour(Now)

'For each element in the array
For i = 0 To 4
'Stop the timers
Stop_Click (i)
Next i

'Set save display box initial directory to the desktop.
CommonDialog1.InitDir = "%userprofile%\desktop"
'Set up the initial filename with the current date.
CommonDialog1.FileName = "Consulting " & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & ".txt"
'Show the save dialog
CommonDialog1.ShowSave

'If they have not selected a directory, quit this procedure.
If CommonDialog1.FileName = "Consulting " & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & ".txt" Then Exit Sub

'If the have, open the chosen file name for write access.
Open CommonDialog1.FileName For Output As #1

    'For each array element.
    For i = 0 To 4
    'Write data only if the field is in use, i.e. a client's name has been
    'enetered.
    If Client(i).Text <> "" Then
    'Cast the client name to a string variable for ease of manipulation.
    Cli = CStr(Client(i).Text)
    'Write it to file.
    Print #1, Cli
    'Skip a line.
    Print #1, ""
    
    'Back up the start/stop mark data
    TickTock(i).SHback = TickTock(i).StopHours
    TickTock(i).SMback = TickTock(i).StopMinutes
    TickTock(i).SSback = TickTock(i).StopSeconds
    
    'This variable is equal to the length of one of the mark variables.
    'When the remaining length = 0 then all the marks have been processed.
    Remaining = Len(TickTock(i).StopHours)
    
        'Repeat until all marks have been processed.
        Do While Remaining > 0
        
        'In this portion, the program will parse two marks out of each of the time
        'variables.  The first mark will be a start time, and the seocnd will be an
        'end time, since they are always written in pairs.  When it reads both marks,
        'It deletes them, and if there are more marks left, it repeats the process.
        
        'The variable 'j' is equal to the length of the portion of the mark string
        'currently until scrutiny.  It's length is increased until the string
        'includes a comma
        j = 0
            'Repeat until the selected protion of the string contains a comma at the end of it
            Do While Left(TickTock(i).StopHours, j) Like "*," = False
            j = j + 1
            Loop
        'Back off one byte so that the comma is ignored.
        j = j - 1
        'Retrieve the mark hour
        markhour1 = Left(TickTock(i).StopHours, j)
        'Reinclude the comma and then remove the lot from the mark string.
        j = j + 1
        TickTock(i).StopHours = Right(TickTock(i).StopHours, Remaining - j)
        
        'Repeat for the minutes start and seconds start marks
        j = 0
            Do While Left(TickTock(i).StopMinutes, j) Like "*," = False
            j = j + 1
            Loop
        j = j - 1
        markminute1 = Left(TickTock(i).StopMinutes, j)
        j = j + 1
        TickTock(i).StopMinutes = Right(TickTock(i).StopMinutes, Remaining - j)
        
        j = 0
            Do While Left(TickTock(i).StopSeconds, j) Like "*," = False
            j = j + 1
            Loop
        j = j - 1
        marksecond1 = Left(TickTock(i).StopSeconds, j)
        j = j + 1
        TickTock(i).StopSeconds = Right(TickTock(i).StopSeconds, Remaining - j)
        
        'Reset the remaining length of marks.  This is neccesary to prevent
        'errors due to differing numbers of digits.
        Remaining = Len(TickTock(i).StopHours)
        
        'As above, parse until the comma, and retrive the stop marks for
        'hours minutes and seconds.
        j = 0
            Do While Left(TickTock(i).StopHours, j) Like "*," = False
            j = j + 1
            Loop
        j = j - 1
        markhour2 = Left(TickTock(i).StopHours, j)
        j = j + 1
        TickTock(i).StopHours = Right(TickTock(i).StopHours, Remaining - j)
        
        j = 0
            Do While Left(TickTock(i).StopMinutes, j) Like "*," = False
            j = j + 1
            Loop
        j = j - 1
        markminute2 = Left(TickTock(i).StopMinutes, j)
        j = j + 1
        TickTock(i).StopMinutes = Right(TickTock(i).StopMinutes, Remaining - j)
        
        j = 0
            Do While Left(TickTock(i).StopSeconds, j) Like "*," = False
            j = j + 1
            Loop
        j = j - 1
        marksecond2 = Left(TickTock(i).StopSeconds, j)
        j = j + 1
        TickTock(i).StopSeconds = Right(TickTock(i).StopSeconds, Remaining - j)
        
        'Calculate the difference in Hour and Minutes from the
        'start mark to the end mark.
        Dhour = markhour2 - markhour1
        Dmin = markminute2 - markminute1
        
        'If the difference in minutes is negative then correctly
        'compensate for the base 60 time system (instead of base
        '100
        If Dmin < 0 Then
        Dhour = Dhour - 1
        Dmin = 60 + Dmin
        End If
        
        'Assume the times are reported in the am
        suffix1 = "am"
        suffix2 = "am"
        'If a mark time is in pm, convert it from 24 hour time,
        'and set the suffix to pm
        If markhour1 > 12 Then
        markhour1 = markhour1 - 12
        suffix1 = "pm"
        End If
        If markhour2 > 12 Then
        markhour2 = markhour2 - 12
        suffix2 = "pm"
        End If
        
        'fix for noon pm
        If markhour1 = 12 Then suffix1 = "pm"
        If markhour2 = 12 Then suffix2 = "pm"
        
        'Write the start and stop marks, as well as the elapsed
        'time for that mark to the file.
        Print #1, Chr(9) & markhour1 & ":" & markminute1 & ":" & marksecond1 & " " & suffix1; " - " & markhour2 & ":" & markminute2 & ":" & marksecond2 & " " & suffix2 & Chr(9) & Dhour & " hr " & Dmin & " min"
        
        'Update the remaining marks.  If there are some left, repeat the above process.
        Remaining = Len(TickTock(i).StopHours)
        Loop
    'Blank line.
    Print #1, ""
    'Print total client time (from timer diaply box).
    Print #1, Chr(9) & Chr(9) & "TOTAL:" & Chr(9) & TickTock(i).Hour & " hr " & TickTock(i).Minute & " min"
    
    'Blank line to space one client from another.
    Print #1, ""
    
    'Restore the mark data from the backup.
    TickTock(i).StopHours = TickTock(i).SHback
    TickTock(i).StopMinutes = TickTock(i).SMback
    TickTock(i).StopSeconds = TickTock(i).SSback
    End If
    
    'Repeat for the next client.
    Next i

Close
ChDir (App.Path)
End Sub

Private Sub Image2_Click()
LastSave = Hour(Now)

'Stop all timers.
For i = 0 To 4
Stop_Click (i)
Next i

    'For each client.
    For i = 0 To 4
    'Only if the client is in use.
    If Client(i).Text <> "" Then
    'Open a calendar file in the program directory.  Its name will
    'always be "<clientnumber>.vcs"
    Open i & ".vcs" For Output As #1
    'Cast client name to string for manipulation.
    Cli = CStr(Client(i).Text)
    'VCS file headers.
    Print #1, "BEGIN:VCALENDAR"
    Print #1, "BEGIN:VEVENT"
    'VCS file Subject / Summary field.  This is simply the
    'client's name.
    Print #1, "SUMMARY;CHARSET=ISO-8859-1;ENCODING=quoted-printable:" & Cli
    'Blank line, for VCS compatibility.
    Print #1, ""
    
    'Back up mark data.
    TickTock(i).SHback = TickTock(i).StopHours
    TickTock(i).SMback = TickTock(i).StopMinutes
    TickTock(i).SSback = TickTock(i).StopSeconds
    
    'The below section behaves just as the parsing routines do in the
    'above procedure for saving data to a file.  Differences are commented.
    
    Remaining = Len(TickTock(i).StopHours)
    'Instead of writing each set of marks right to a file, they are compiled
    'into one string (since the body of the VCS must have no carriage returns)
    'and written all at once.  'Summary' holds this string until written.
    Summary = ""
    'Also, the start and end times of the entire client session become the
    'start and end of the appointment.  Thus the first mark must be recorded.
    'The last mark does not need to be recorded, since it's data will be
    'lingering in the variables when the loop ends.  The 'First' variable will
    'ensure that the first data point is captured and not overwritten.
    First = True
        Do While Remaining > 0
        j = 0
            Do While Left(TickTock(i).StopHours, j) Like "*," = False
            j = j + 1
            Loop
        j = j - 1
        markhour1 = Left(TickTock(i).StopHours, j)
        j = j + 1
        TickTock(i).StopHours = Right(TickTock(i).StopHours, Remaining - j)
        'If this is the first mark, save it.
        If First = True Then firsthour = markhour1
        
        j = 0
            Do While Left(TickTock(i).StopMinutes, j) Like "*," = False
            j = j + 1
            Loop
        j = j - 1
        markminute1 = Left(TickTock(i).StopMinutes, j)
        j = j + 1
        TickTock(i).StopMinutes = Right(TickTock(i).StopMinutes, Remaining - j)
        'If this is the first mark, save it.
        If First = True Then firstminute = markminute1
        
        j = 0
            Do While Left(TickTock(i).StopSeconds, j) Like "*," = False
            j = j + 1
            Loop
        j = j - 1
        marksecond1 = Left(TickTock(i).StopSeconds, j)
        j = j + 1
        TickTock(i).StopSeconds = Right(TickTock(i).StopSeconds, Remaining - j)
        'If this is the first mark, save it.
        If First = True Then
        firstsecond = marksecond1
        'Also, since all the mark's data have been gathered now, set
        ''First' = false
        First = False
        End If
        
        Remaining = Len(TickTock(i).StopHours)
        
        j = 0
            Do While Left(TickTock(i).StopHours, j) Like "*," = False
            j = j + 1
            Loop
        j = j - 1
        markhour2 = Left(TickTock(i).StopHours, j)
        j = j + 1
        TickTock(i).StopHours = Right(TickTock(i).StopHours, Remaining - j)

        
        j = 0
            Do While Left(TickTock(i).StopMinutes, j) Like "*," = False
            j = j + 1
            Loop
        j = j - 1
        markminute2 = Left(TickTock(i).StopMinutes, j)
        j = j + 1
        TickTock(i).StopMinutes = Right(TickTock(i).StopMinutes, Remaining - j)

        
        j = 0
            Do While Left(TickTock(i).StopSeconds, j) Like "*," = False
            j = j + 1
            Loop
        j = j - 1
        marksecond2 = Left(TickTock(i).StopSeconds, j)
        j = j + 1
        TickTock(i).StopSeconds = Right(TickTock(i).StopSeconds, Remaining - j)

        '
        Dhour = markhour2 - markhour1
        Dmin = markminute2 - markminute1
        
        If Dmin < 0 Then
        Dhour = Dhour - 1
        Dmin = 60 + Dmin
        End If
        
        firsthour2 = markhour2
        suffix1 = "am"
        suffix2 = "am"
        If markhour1 > 12 Then
        markhour1 = markhour1 - 12
        suffix1 = "pm"
        End If
        If markhour2 > 12 Then
        markhour2 = markhour2 - 12
        suffix2 = "pm"
        End If
        
        'fix for noon pm
        If markhour1 = 12 Then suffix1 = "pm"
        If markhour2 = 12 Then suffix2 = "pm"
        
        'Compile the Summary based on the all the mark data.
        'This must be one line.  This string is cumulative and
        'will include the marks from each loop so that all are
        'included.
        Summary = Summary & markhour1 & ":" & markminute1 & ":" & marksecond1 & " " & suffix1 & " - " & markhour2 & ":" & markminute2 & ":" & marksecond2 & " " & suffix2 & " --> " & Dhour & " hr " & Dmin & " min.  "
        
        Remaining = Len(TickTock(i).StopHours)
        Loop
        
    'Write the mark summary to the calendar entry body.
    Print #1, "DESCRIPTION;CHARSET=ISO-8859-1;ENCODING=quoted-printable:" & Summary
    
    'VCS requires all 2 digits dates, so ensure that
    'the days and month reported form the OS are in
    '2 digit format.  (years are always reported as 4)
    writeday = Day(Now)
    If Len(writeday) = 1 Then writeday = "0" & writeday
    writemonth = Month(Now)
    If Len(writemonth) = 1 Then writemonth = "0" & writemonth
    
    'Write the start time.  The format is:
    'DSTART:YYYYMMDDTHHMMSS?
    'The central T is a constant, and the ? at the end
    'can be used as a time zone conversion.  Valid values
    'for ? are 'Z' and others (which i dont know).  If
    ' ? is omitted, local time is used.
    Print #1, "DTSTART:" & Year(Now) & writemonth & writeday & "T" & firsthour & firstminute & firstsecond
    'Write the end time, same format as above.
    Print #1, "DTEND:" & Year(Now) & writemonth & writeday & "T" & firsthour2 & markminute2 & marksecond2
    'VCS footers.
    Print #1, "END:VEVENT"
    Print #1, "END:VCALENDAR"

    Close
    
    'Open .VCS with associated app.
    'Params. are "Window handle, action, filename, , , Window state"
    ShellExecute hwnd, "open", i & ".vcs", vbNullString, vbNullString, SW_SHOWNORMAL
    
    'Restore mark data from backup
    TickTock(i).StopHours = TickTock(i).SHback
    TickTock(i).StopMinutes = TickTock(i).SMback
    TickTock(i).StopSeconds = TickTock(i).SSback

    End If
    
    'Repeat for the next client.
    Next i

ChDir (App.Path)
End Sub

Private Sub Image3_Click()
If Form1.Height > 6780 Then
Form1.Height = 6780
Else
Form1.Height = 9780
End If
End Sub

Private Sub Option1_Click()
On Error GoTo quit
Open "settings.ini" For Output As #4
Print #4, Text1.Text
Print #4, Check1.Value
Print #4, Check2.Value
Print #4, Check3.Value
Print #4, Text2.Text
Print #4, Check4.Value
Print #4, Option1.Value
Print #4, Option2.Value
Print #4, Text3.Text

Close #4
quit:
End Sub

Private Sub Option2_Click()
On Error GoTo quit
Open "settings.ini" For Output As #4
Print #4, Text1.Text
Print #4, Check1.Value
Print #4, Check2.Value
Print #4, Check3.Value
Print #4, Text2.Text
Print #4, Check4.Value
Print #4, Option1.Value
Print #4, Option2.Value
Print #4, Text3.Text

Close #4
quit:
End Sub

Private Sub Stop_Click(Index As Integer)

'Assume the timer is ON:  (eliminates double click
'errors, as in the 'Go_Click' procedure)
If Timing(Index) = True Then

'Add CSV mark time values for each of the respective
'measurements, as in 'Go_Click'
markhour = Hour(Now)
If Len(markhour) = 1 Then markhour = "0" & markhour
TickTock(Index).StopHours = TickTock(Index).StopHours & markhour & ","
markminute = Minute(Now)
If Len(markminute) = 1 Then markminute = "0" & markminute
TickTock(Index).StopMinutes = TickTock(Index).StopMinutes & markminute & ","
marksecond = Second(Now)
If Len(marksecond) = 1 Then marksecond = "0" & marksecond
TickTock(Index).StopSeconds = TickTock(Index).StopSeconds & marksecond & ","
End If

'Turn the timer off
Timing(Index) = False
TickTock(Index).Dtee = 0

off = True
For i = 0 To 4
If Timing(i) = True Then off = False
Next i

If off = True Then
RemoveFromTray
Form1.Icon = Form3.Icon
    If WindowState = vbMinimized Then
        LastState = vbNormal
    Else
        LastState = WindowState
    End If

    AddToTray Me ', mnuTray

    SetTrayTip "Consulting Timer"
End If

End Sub


Private Sub Text1_Change()
On Error GoTo quit
Open "settings.ini" For Output As #4
Print #4, Text1.Text
ClientPath = Text1.Text

Print #4, Check1.Value
Print #4, Check2.Value
Print #4, Check3.Value
Print #4, Text2.Text
Print #4, Check4.Value
Print #4, Option1.Value
Print #4, Option2.Value
Print #4, Text3.Text


Close #4
quit:
End Sub

Private Sub Text2_Change()
On Error GoTo quit
Open "settings.ini" For Output As #4
Print #4, Text1.Text
Print #4, Check1.Value
Print #4, Check2.Value
Print #4, Check3.Value
Print #4, Text2.Text
Print #4, Check4.Value
Print #4, Option1.Value
Print #4, Option2.Value
Print #4, Text3.Text

Close #4
quit:
End Sub

Private Sub Text3_Change()
On Error GoTo quit
Open "settings.ini" For Output As #4
Print #4, Text1.Text
Print #4, Check1.Value
Print #4, Check2.Value
Print #4, Check3.Value
Print #4, Text2.Text
Print #4, Check4.Value
Print #4, Option1.Value
Print #4, Option2.Value
Print #4, Text3.Text

Close #4
quit:
End Sub

Private Sub Timer1_Timer()
'For each of the 5 fields:
For i = 0 To 4

'If the timer is on.
If Timing(i) = True Then

'Add a second to the second hand (this is a 1 second timer)
TickTock(i).Second = TickTock(i).Second + 1

'If the seconds hand has reached 60, reset it to zero, and
'add a digit to the minutes hand.
If TickTock(i).Second > 59 Then
TickTock(i).Minute = TickTock(i).Minute + 1
TickTock(i).Second = 0
End If

'If the minutes have reach 60, reset to zero, and add one
'to the hour hand.
If TickTock(i).Minute > 59 Then
TickTock(i).Hour = TickTock(i).Hour + 1
TickTock(i).Minute = 0
End If

'Format and display the time in the timer boxes, just as in
''Form_Load'
Time(i).Text = TickTock(i).Hour & ":"

temp = Int(TickTock(i).Minute)
'Double digit check
If Len(temp) = 1 Then
Time(i).Text = Time(i).Text & "0" & TickTock(i).Minute & ":"
Else
Time(i).Text = Time(i).Text & TickTock(i).Minute & ":"
End If

temp = Int(TickTock(i).Second)
'Double digit check
If Len(temp) = 1 Then
Time(i).Text = Time(i).Text & "0" & TickTock(i).Second
Else
Time(i).Text = Time(i).Text & TickTock(i).Second
End If

If Check4.Value = 1 Then
TickTock(i).Dtee = TickTock(i).Dtee + 1
checkola = CInt(Text2.Text)
If TickTock(i).Dtee = checkola * 60 Then
TickTock(i).Dtee = 0
If Option1.Value = True Then Beep
If Option2.Value = True Then
'play wave in text3.text
MMControl1.FileName = Text3.Text
MMControl1.Command = "close"
MMControl1.Command = "open"
MMControl1.Command = "play"
End If
End If

End If

End If
Next i
End Sub
