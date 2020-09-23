VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Date Changa"
   ClientHeight    =   6015
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7215
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "file"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox lstFiles 
      Height          =   2820
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   1815
   End
   Begin VB.DriveListBox lstDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1815
   End
   Begin VB.DirListBox lstFolders 
      Height          =   2565
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "New Attributes"
      Height          =   3855
      Left            =   2040
      TabIndex        =   22
      Top             =   2040
      Width           =   5055
      Begin VB.Timer tmrTime 
         Interval        =   500
         Left            =   4560
         Top             =   240
      End
      Begin VB.CommandButton cmdSet 
         Caption         =   "Set Time"
         Height          =   855
         Left            =   2040
         Picture         =   "frmMain.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2880
         Width           =   2895
      End
      Begin VB.CheckBox chkAccessed 
         Caption         =   "Accessed"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CheckBox chkModified 
         Caption         =   "Modified"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   3240
         Width           =   975
      End
      Begin VB.CheckBox chkCreated 
         Caption         =   "Created"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   3000
         Width           =   975
      End
      Begin VB.Frame fmeTime 
         Caption         =   "Custom Time"
         Enabled         =   0   'False
         Height          =   1815
         Left            =   2880
         TabIndex        =   23
         Top             =   840
         Width           =   2055
         Begin VB.CommandButton cmdCurrent 
            Caption         =   "Current Time"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1815
         End
         Begin MSComCtl2.UpDown udSec 
            Height          =   285
            Left            =   1695
            TabIndex        =   27
            Top             =   960
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtSec"
            BuddyDispid     =   196619
            OrigLeft        =   1680
            OrigTop         =   1080
            OrigRight       =   1920
            OrigBottom      =   1335
            Max             =   60
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udMin 
            Height          =   285
            Left            =   1695
            TabIndex        =   9
            Top             =   600
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtMin"
            BuddyDispid     =   196620
            OrigLeft        =   1680
            OrigTop         =   720
            OrigRight       =   1920
            OrigBottom      =   975
            Max             =   60
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtSec 
            Height          =   285
            Left            =   960
            TabIndex        =   10
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtMin 
            Height          =   285
            Left            =   960
            TabIndex        =   8
            Text            =   "0"
            Top             =   600
            Width           =   735
         End
         Begin MSComCtl2.UpDown udHour 
            Height          =   285
            Left            =   1695
            TabIndex        =   7
            Top             =   240
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtHour"
            BuddyDispid     =   196621
            OrigLeft        =   1680
            OrigTop         =   360
            OrigRight       =   1920
            OrigBottom      =   615
            Max             =   60
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtHour 
            Height          =   285
            Left            =   960
            TabIndex        =   6
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Second:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   600
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minute:"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   600
            Width           =   525
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hour:"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   390
         End
      End
      Begin VB.CheckBox chkCurrent 
         Caption         =   "Use Current Time"
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin MSComCtl2.MonthView calendar 
         Height          =   2370
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   19202049
         CurrentDate     =   37547
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2880
         TabIndex        =   29
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Label1 
         Caption         =   "Change:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current File Attributes"
      Height          =   1815
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtAccess 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txtModify 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtCreate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Last Accessed:"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Last Modified:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Created on:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Menu About 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'thanx go to mark for his help for stuff
'Visit his Homepage at
'http://www.geocities.com/marskarthik
'http://marskarthik.virtualave.net
'Email: marskarthik@angelfire.com
'and vbcode.com, and blackbeltvb.com for the help on FILETIME and SYSTEMTIME

Private Sub About_Click()
'pretty self explanatory :D
MsgBox "Date Changa v1.9" & vbCrLf & vbCrLf & "Will read and change date's and times of files created on and modified on." & vbCrLf & vbCrLf & "Any comments or suggestion email chris.truscott@bigpond.com" & vbCrLf & vbCrLf & "By Isaac - 4truscis@students.nudgee.com", vbInformation, "About Date Changa v1.2"
End Sub

Private Sub cmdCurrent_Click() ' update the 3 text boxes to current time
txtHour.Text = Hour(Time)
txtMin.Text = Minute(Time)
txtSec.Text = Second(Time)
End Sub

Private Sub chkCurrent_Click() ' enable/disable the current time and time label depending on value
If chkCurrent.Value = 1 Then
lblTime.Visible = True
fmeTime.Enabled = False
Else
lblTime.Visible = False
fmeTime.Enabled = True
End If
End Sub

Private Sub cmdSet_Click()
On Error GoTo errhandle

'dim variables
Dim hFile As Long, rval As Long
Dim buff As OFSTRUCT
Dim ctime As FILETIME, mtime As FILETIME, latime As FILETIME
Dim stime As SYSTEMTIME
Dim filen As String


'attempt to open file already opened for write
filen = CStr(lstFiles.Path & "\" & lstFiles.FileName)
hFile = OpenFile(filen, buff, OF_WRITE)

If hFile Then

    'get original file times
    rval = GetFileTime(hFile, ctime, latime, mtime)

    '-------------
    'create time
    '-------------
    
    If chkCreated.Value = 1 Then
        
        'convert system to file time
        rval = FileTimeToLocalFileTime(ctime, ctime)
        rval = FileTimeToSystemTime(ctime, stime)

        'Change filetimes
        stime.wYear = calendar.Year
        stime.wMonth = calendar.Month
        stime.wDay = calendar.Day
        If chkCurrent.Value = 1 Then
            stime.wHour = Hour(Time)
            stime.wMinute = Minute(Time)
            stime.wSecond = Second(Time)
        Else
            stime.wHour = txtHour.Text
            stime.wMinute = txtMin.Text
            stime.wSecond = txtSec.Text
        End If
        
        'reconvert it back and save it
        rval = SystemTimeToFileTime(stime, ctime)
        rval = LocalFileTimeToFileTime(ctime, ctime)
    
    End If
    
    
    '-------------
    'last write / modified time
    '-------------
    
    If chkModified.Value = 1 Then
        
        'convert system to file time
        rval = FileTimeToLocalFileTime(mtime, mtime)
        rval = FileTimeToSystemTime(mtime, stime)
    
            'Change filetimes
        stime.wYear = calendar.Year
        stime.wMonth = calendar.Month
        stime.wDay = calendar.Day
        If chkCurrent.Value = 1 Then
            stime.wHour = Hour(Time)
            stime.wMinute = Minute(Time)
            stime.wSecond = Second(Time)
        Else
            stime.wHour = txtHour.Text
            stime.wMinute = txtMin.Text
            stime.wSecond = txtSec.Text
        End If
        

        'reconvert it back and save it
        rval = SystemTimeToFileTime(stime, mtime)
        rval = LocalFileTimeToFileTime(mtime, mtime)
        
    End If
    
    '-------------
    'last accessed time
    '-------------
    
    If chkAccessed.Value = 1 Then
    
        'convert system to file time
        rval = FileTimeToLocalFileTime(latime, latime)
        rval = FileTimeToSystemTime(latime, stime)

        'Change filetimes
        stime.wYear = calendar.Year
        stime.wMonth = calendar.Month
        stime.wDay = calendar.Day

        If chkCurrent.Value = 1 Then
            stime.wHour = Hour(Time)
            stime.wMinute = Minute(Time)
            stime.wSecond = Second(Time)
        Else
            stime.wHour = txtHour.Text
            stime.wMinute = txtMin.Text
            stime.wSecond = txtSec.Text
        End If
        

        'reconvert it back and save it
        rval = SystemTimeToFileTime(stime, latime)
        rval = LocalFileTimeToFileTime(latime, latime)
    
    End If


    'and finally write 'em
    rval = SetFileTime(hFile, ctime, latime, mtime)

End If

'close file
rval = CloseHandle(hFile)

'alert and refresh time
MsgBox "File Saved succesfully!", vbOKOnly + vbInformation, "Succesful Edit"
lstFiles_Click

Exit Sub

errhandle: 'handles errors
    MsgBox "Error #" & Err.Number & ", " & Err.Description & ".", vbCritical, "ERROR!"
    Exit Sub
End Sub

Private Sub lstDrive_Change() 'updates folders to current drive
lstFolders.Path = lstDrive.Drive
End Sub

Private Sub lstFiles_Click()
If lstFiles.FileName <> "" Then

'dim variables
Dim hFile As Long, rval As Long
Dim buff As OFSTRUCT
Dim ctime As FILETIME, mtime As FILETIME, latime As FILETIME
Dim stime As SYSTEMTIME
Dim filen As String

txtCreate.Text = ""
txtModify.Text = ""
txtAccess.Text = ""


'attempt to open file already opened for write
filen = CStr(lstFiles.Path & "\" & lstFiles.FileName)
hFile = OpenFile(filen, buff, OF_WRITE)

If hFile Then

    'get th times of the files into ctime,mtime and latime
    rval = GetFileTime(hFile, ctime, latime, mtime)
    
    '---------------------------------
    'file create part
    '---------------------------------

    'convert it to local, then to system time
    rval = FileTimeToLocalFileTime(ctime, ctime)
    rval = FileTimeToSystemTime(ctime, stime)

    'put the name of the day of the week in
    Select Case stime.wDayOfWeek
        Case 0
            txtCreate.Text = txtCreate.Text & "Sunday, "
        Case 1
            txtCreate.Text = txtCreate.Text & "Monday, "
        Case 2
            txtCreate.Text = txtCreate.Text & "Tuesday, "
        Case 3
            txtCreate.Text = txtCreate.Text & "Wednesday, "
        Case 4
            txtCreate.Text = txtCreate.Text & "Thursday, "
        Case 5
            txtCreate.Text = txtCreate.Text & "Friday, "
        Case 6
            txtCreate.Text = txtCreate.Text & "Saturday, "
    End Select
    
    'put the name of the month in
    Select Case stime.wMonth
        Case 1
            txtCreate.Text = txtCreate.Text & "January "
        Case 2
            txtCreate.Text = txtCreate.Text & "February "
        Case 3
            txtCreate.Text = txtCreate.Text & "March "
        Case 4
            txtCreate.Text = txtCreate.Text & "April "
        Case 5
            txtCreate.Text = txtCreate.Text & "May "
        Case 6
            txtCreate.Text = txtCreate.Text & "June "
        Case 7
            txtCreate.Text = txtCreate.Text & "July "
        Case 8
            txtCreate.Text = txtCreate.Text & "August "
        Case 9
            txtCreate.Text = txtCreate.Text & "September "
        Case 10
            txtCreate.Text = txtCreate.Text & "October "
        Case 11
            txtCreate.Text = txtCreate.Text & "November "
        Case 12
            txtCreate.Text = txtCreate.Text & "December "
    End Select
    
    'make sure the date is always 2 digits long
    txtCreate.Text = txtCreate.Text & Format(stime.wDay, "00") & ", "
    
    'put the year in
    txtCreate.Text = txtCreate.Text & stime.wYear & ", at "

    'convert from 24 hour to 12 hour if necessary with am and pm
    If stime.wHour > 12 Then
        txtCreate.Text = txtCreate.Text & (stime.wHour - 12) & ":" & Format(stime.wMinute, "00") & ":" & Format(stime.wSecond, "00") & " PM"
    Else
        txtCreate.Text = txtCreate.Text & stime.wHour & ":" & Format(stime.wMinute, "00") & ":" & Format(stime.wSecond, "00") & " AM"
    End If
    
    '---------------------------------------
    ' File modified part
    '---------------------------------------
    
    'convert to local and system times
    rval = FileTimeToLocalFileTime(mtime, mtime)
    rval = FileTimeToSystemTime(mtime, stime)
    
    'put the name of the day of the week in
    Select Case stime.wDayOfWeek
        Case 0
            txtModify.Text = txtModify.Text & "Sunday, "
        Case 1
            txtModify.Text = txtModify.Text & "Monday, "
        Case 2
            txtModify.Text = txtModify.Text & "Tuesday, "
        Case 3
            txtModify.Text = txtModify.Text & "Wednesday, "
        Case 4
            txtModify.Text = txtModify.Text & "Thursday, "
        Case 5
            txtModify.Text = txtModify.Text & "Friday, "
        Case 6
            txtModify.Text = txtModify.Text & "Saturday, "
    End Select
    
    'put the name of the month in
    Select Case stime.wMonth
        Case 1
            txtModify.Text = txtModify.Text & "January "
        Case 2
            txtModify.Text = txtModify.Text & "February "
        Case 3
            txtModify.Text = txtModify.Text & "March "
        Case 4
            txtModify.Text = txtModify.Text & "April "
        Case 5
            txtModify.Text = txtModify.Text & "May "
        Case 6
            txtModify.Text = txtModify.Text & "June "
        Case 7
            txtModify.Text = txtModify.Text & "July "
        Case 8
            txtModify.Text = txtModify.Text & "August "
        Case 9
            txtModify.Text = txtModify.Text & "September "
        Case 10
            txtModify.Text = txtModify.Text & "October "
        Case 11
            txtModify.Text = txtModify.Text & "November "
        Case 12
            txtModify.Text = txtModify.Text & "December "
    End Select
    
    'make sure the date is always 2 digits long
    txtModify.Text = txtModify.Text & Format(stime.wDay, "00") & ", "
    
    'put the year in
    txtModify.Text = txtModify.Text & stime.wYear & ", at "
    
    'convert from 24 hour to 12 hour if necessary with am and pm
    If stime.wHour > 12 Then
        txtModify.Text = txtModify.Text & (stime.wHour - 12) & ":" & Format(stime.wMinute, "00") & ":" & Format(stime.wSecond, "00") & " PM"
    Else
        txtModify.Text = txtModify.Text & stime.wHour & ":" & Format(stime.wMinute, "00") & ":" & Format(stime.wSecond, "00") & " AM"
    End If
    
    '---------------------------------------
    ' File last accessed part part
    '---------------------------------------
    
    'convert to local and system times
    rval = FileTimeToLocalFileTime(latime, latime)
    rval = FileTimeToSystemTime(latime, stime)
    
    'put the name of the day of the week in
    Select Case stime.wDayOfWeek
        Case 0
            txtAccess.Text = txtAccess.Text & "Sunday, "
        Case 1
            txtAccess.Text = txtAccess.Text & "Monday, "
        Case 2
            txtAccess.Text = txtAccess.Text & "Tuesday, "
        Case 3
            txtAccess.Text = txtAccess.Text & "Wednesday, "
        Case 4
            txtAccess.Text = txtAccess.Text & "Thursday, "
        Case 5
            txtAccess.Text = txtAccess.Text & "Friday, "
        Case 6
            txtAccess.Text = txtAccess.Text & "Saturday, "
    End Select
    
    'put the name of the month in
    Select Case stime.wMonth
        Case 1
            txtAccess.Text = txtAccess.Text & "January "
        Case 2
            txtAccess.Text = txtAccess.Text & "February "
        Case 3
            txtAccess.Text = txtAccess.Text & "March "
        Case 4
            txtAccess.Text = txtAccess.Text & "April "
        Case 5
            txtAccess.Text = txtAccess.Text & "May "
        Case 6
            txtAccess.Text = txtAccess.Text & "June "
        Case 7
            txtAccess.Text = txtAccess.Text & "July "
        Case 8
            txtAccess.Text = txtAccess.Text & "August "
        Case 9
            txtAccess.Text = txtAccess.Text & "September "
        Case 10
            txtAccess.Text = txtAccess.Text & "October "
        Case 11
            txtAccess.Text = txtAccess.Text & "November "
        Case 12
            txtAccess.Text = txtAccess.Text & "December "
    End Select
    
    'make sure the date is always 2 digits long
    txtAccess.Text = txtAccess.Text & Format(stime.wDay, "00") & ", "
    
    'put the year in
    txtAccess.Text = txtAccess.Text & stime.wYear
    
End If
rval = CloseHandle(hFile) 'close file

End If
End Sub

Private Sub lstFolders_Change() 'update current files to current folder
If lstFolders.Path <> "" Then
    lstFiles.Path = lstFolders.Path
End If
End Sub

Private Sub tmrTime_Timer() ' update the current time label
lblTime.Caption = Time
End Sub

Private Sub txtHour_Change() 'keep the hour in check
If txtHour.Text = "" Or IsNumeric(txtHour.Text) = False Then
    txtHour.Text = 0
    Beep
End If
If txtHour.Text > 60 Then
    txtHour.Text = 60
End If
udHour.Value = txtHour.Text
End Sub

Private Sub txtMin_Change() ' keep the minutes in check
If txtMin.Text = "" Or IsNumeric(txtMin.Text) = False Then
    txtMin.Text = 0
    Beep
End If
If txtMin.Text > 60 Then
    txtMin.Text = 60
End If
udMin.Value = txtMin.Text
End Sub

Private Sub txtSec_Change() ' keep the seconds in check
If txtSec.Text = "" Or IsNumeric(txtSec.Text) = False Then
    txtSec.Text = 0
    Beep
End If
If txtSec.Text > 60 Then
    txtSec.Text = 60
End If
udSec.Value = txtSec.Text
End Sub
