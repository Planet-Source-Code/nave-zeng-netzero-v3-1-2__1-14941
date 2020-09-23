VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "The Internet was meant to be free coded by Navarchy"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6315
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Restore Banners"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   500
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Kill Banners"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   20
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Click Here If You Didn't Install To The Default Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "C:\Program Files\NetZero\"
      Top             =   210
      Width           =   4095
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   6015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Please select the directory in which NetZero is installed"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   930
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Directory where NetZero is installed"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    If Me.Height = 3180 Then Exit Sub
    a = MsgBox("Did you install NetZero to the default directory?", vbYesNo + vbQuestion, "")
    If a <> 7 Then Exit Sub
    Dir1.SetFocus
    Me.Height = 3180

End Sub

Private Sub Command2_Click()

    If Dir$(Text1.Text & "chkras.exe") = "" Then
        MsgBox "It appears as though this is an invalid NetZero directory." & vbCrLf & "Please select the correct directory and try again", vbInformation, "Invalid Directory"
        Exit Sub
    End If
    If Dir$(Text1.Text & "chkras.bak") <> "" Then
        MsgBox "It appears as though the NetZero banners have already been killed. If the" & vbCrLf & "banners are not killed then try clicking 'Restore Banners' and then 'Kill Banners'", vbInformation, "Already Killed"
        Exit Sub
    End If
    On Error GoTo errorh
    FileCopy Text1.Text & "chkras.exe", Text1.Text & "chkras.bak"
  Dim FileNumber As Integer
  Dim exeBuffer() As Byte
    exeBuffer = LoadResData(2, "CUSTOM")
    FileNumber = FreeFile
    Open Text1.Text & "chkras.exe" For Binary Access Write As #FileNumber
    Put #FileNumber, , exeBuffer
    Close #FileNumber
    If MsgBox("The NetZero banners have been successfuly killed" & vbCrLf & "Now would you like to exit?", vbExclamation + vbYesNo, "Sweet Success") <> 7 Then Unload Me: End

Exit Sub

errorh:
    MsgBox "There was an error killing the NetZero banners." & vbCrLf & "A possible cause is that NetZero is running." & vbCrLf & "If it is close it and try again.", vbCritical, "Error"

End Sub

Private Sub Command3_Click()
On Error GoTo errorh
    If Dir$(Text1.Text & "chkras.exe") = "" Then
        MsgBox "It appears as though this is an invalid NetZero directory." & vbCrLf & "Please select the correct directory and try again", vbInformation, "Invalid Directory"
        Exit Sub
    End If
    If Dir$(Text1.Text & "chkras.bak") = "" Then
        MsgBox "NetZero cannot be restored. It is possible that this program was not used to kill the NetZero banners or that someone has manually edited the NetZero files.", vbInformation, "Cannot Restore"
        Exit Sub
    End If
    FileCopy Text1.Text & "chkras.bak", Text1.Text & "chkras.exe"
    Kill Text1.Text & "chkras.bak"
    If MsgBox("The NetZero banners have been successfuly restored" & vbCrLf & "Now would you like to exit?", vbExclamation + vbYesNo, "Sweet Success") <> 7 Then Unload Me: End
    Exit Sub
errorh:
MsgBox "There was an error restoring the NetZero banners." & vbCrLf & "A possible cause is that NetZero is running." & vbCrLf & "If it is close it and try again.", vbCritical, "Error"
End Sub

Private Sub Dir1_Change()

    If Right$(Dir1.Path, 1) <> "\" Then a = "\"
    Text1.Text = Dir1.Path & a

End Sub

Private Sub Drive1_Change()

    On Error GoTo Err
    Dir1.Path = Drive1.Drive

Exit Sub

Err:
    Drive1.Drive = Left$(Dir1.Path, 3)
    MsgBox "That device is not ready", vbCritical, "Error"

End Sub

Private Sub Form_Load()

    Me.Move Screen.Width / 2 - Me.Width / 2, Screen.Height / 2 - Me.Height / 2

End Sub

