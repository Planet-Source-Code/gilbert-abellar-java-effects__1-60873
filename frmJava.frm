VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmJava 
   Caption         =   "Javascript Source Code"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10830
   Icon            =   "frmJava.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrClipboard 
      Interval        =   1000
      Left            =   -1440
      Top             =   3480
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   -1800
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create &HTM"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   14
      Top             =   6960
      Width           =   3975
   End
   Begin VB.Timer tmrProcess 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   -2040
      Top             =   3480
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      ScaleHeight     =   225
      ScaleWidth      =   3945
      TabIndex        =   20
      Top             =   6600
      Width           =   3975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Menus:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3360
      TabIndex        =   19
      Top             =   5280
      Width           =   3255
      Begin VB.TextBox txtMen 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtMen 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtMen 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtMen 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Command:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6720
      TabIndex        =   18
      Top             =   5280
      Width           =   3975
      Begin VB.CommandButton cmdClear 
         Caption         =   "Cle&ar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "&Copy To Clipboard"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   720
      Width           =   10575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Java Effects:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   15
      Top             =   5280
      Width           =   3135
      Begin VB.CommandButton cmdTextGen 
         Caption         =   "&Generate"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.VScrollBar VS 
         Enabled         =   0   'False
         Height          =   315
         Left            =   870
         TabIndex        =   6
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txtInt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cmbOptions 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Interval:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Style:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.TextBox txtMes 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   9975
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6720
      TabIndex        =   21
      Top             =   6360
      Width           =   45
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Javascript Source Code:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4380
      TabIndex        =   17
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Text:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "frmJava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbOptions_Click()
    If cmbOptions.ListIndex = 6 Then
        Frame3.Enabled = False
        txtMes.Enabled = False
        txtInt.Enabled = True
        VS.Enabled = True
        cmdTextGen.Enabled = True
        VS.Value = 30
        txtInt = VS.Value
        Exit Sub
    End If
    If cmbOptions.ListIndex = 10 Or cmbOptions.ListIndex = 11 Then
        Frame3.Enabled = True
        txtMes.Enabled = False
        txtInt.Enabled = False
        VS.Enabled = False
        cmdTextGen.Enabled = True
        Exit Sub
    End If
    If cmbOptions.ListIndex = 8 Or cmbOptions.ListIndex = 9 Then
        txtMes.Enabled = False
        txtInt.Enabled = True
        VS.Enabled = True
        cmdTextGen.Enabled = True
        Frame3.Enabled = False
        Exit Sub
    Else
        cmdTextGen.Enabled = True
        Frame3.Enabled = False
        txtMes.Enabled = True
        txtInt.Enabled = False
        VS.Enabled = False
    End If
    If cmbOptions.ListIndex > 5 Then
        cmdTextGen.Enabled = True
        txtMes.Enabled = False
        Frame3.Enabled = False
        txtInt.Enabled = False
        VS.Enabled = False
    Else
        cmdTextGen.Enabled = True
        Frame3.Enabled = False
        txtMes.Enabled = True
        txtInt.Enabled = True
        VS.Enabled = True
    End If
    If cmbOptions.ListIndex = 12 Or cmbOptions.ListIndex = 13 Or _
            cmbOptions.ListIndex = 15 Or cmbOptions.ListIndex = 16 Then
        cmdTextGen.Enabled = True
    End If
    If cmbOptions.ListIndex = 18 Then
        txtMes.Enabled = True
    End If
    If cmbOptions.ListIndex = 17 Then
        txtMes.Enabled = False
        txtInt.Enabled = False
        VS.Enabled = False
        cmdTextGen.Enabled = True
        Frame3.Enabled = True
        Frame3 = "Web Links:"
        Exit Sub
    Else
        Frame3 = "Menus:"
    End If
    If cmbOptions.ListIndex = 14 Then
        cmdTextGen.Enabled = True
        txtInt.Enabled = True
        VS.Enabled = True
        Label3 = "&Expiration Date:"
        VS.Value = 30
        txtInt = VS.Value
    Else
        Label3 = "&Interval:"
    End If
    Select Case cmbOptions.ListIndex
        Case 0, 1, 2, 3, 4, 5
            VS.Value = 50
            txtInt = VS.Value
    End Select
End Sub

Private Sub cmdClear_Click()
    Dim X As Integer
    txtMes = vbNullString
    txtOutput = vbNullString
    VS.Value = VS.Min
    txtInt = VS.Value
    Clipboard.Clear
    cmbOptions.ListIndex = -1
    cmdTextGen.Enabled = False
    txtMes.Enabled = False
    txtInt.Enabled = False
    VS.Enabled = False
    cmdCopy.Enabled = False
    cmdCreate.Enabled = False
    cmdClear.Enabled = False
    Frame3.Enabled = False
    Label3 = "&Interval"
    Frame3 = "Menus:"
    For X = 0 To 3
        txtMen(X) = vbNullString
    Next X
    Form_Load
    On Error Resume Next
    txtMes.SetFocus
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText TextString
    cmdClear.Enabled = True
    On Error Resume Next
    txtMes.SetFocus
End Sub

Private Sub cmdCreate_Click()
    If Not txtOutput = vbNullString Then
        On Error GoTo ErrSave
        CD.Filter = "All Files (*.*)|*.*|Web Pages (*.html;*.htm)|*.htm;*.html|Text Files (*.txt)|*.txt"
        CD.CancelError = True
        CD.FilterIndex = 2
        CD.DialogTitle = "Export Source Code"
        CD.Flags = &H2
        CD.FileName = cmbOptions.List(cmbOptions.ListIndex)
        CD.ShowSave
        If Not CD.FileName = vbNullString Then
            Open CD.FileName For Output As #1
            CreateDocument cmbOptions.ListIndex, txtMes, Val(Trim$(txtInt)), cmbOptions.List(cmbOptions.ListIndex)
            Close #1
        End If
    Else
        MsgBox "There is no codes to create.", vbOKOnly + vbExclamation + vbApplicationModal, "Java Effects"
        Exit Sub
    End If
    Exit Sub

ErrSave:
End Sub

Private Sub cmdTextGen_Click()
    If cmbOptions.ListIndex >= 0 And cmbOptions.ListIndex <= 5 Then
        If txtMes = vbNullString Then
            MsgBox "Please fill-up " & """" & "TEXT" & """" & " to complete the task.", vbOKOnly + vbCritical + vbApplicationModal, "Java Effects"
            Exit Sub
        ElseIf txtInt = vbNullString Then
            MsgBox "Please fill-up " & """" & "INTERVAL" & """" & " to complete the task.", vbOKOnly + vbCritical + vbApplicationModal, "Java Effects"
            Exit Sub
        End If
    ElseIf cmbOptions.ListIndex = 10 Or cmbOptions.ListIndex = 11 Then
        If txtMen(0) = vbNullString Or txtMen(1) = vbNullString Or txtMen(2) = vbNullString Or txtMen(3) = vbNullString Then
            MsgBox "Please fill-up " & """" & "MENUS" & """" & " to complete the task.", vbOKOnly + vbCritical + vbApplicationModal, "Java Effects"
            Exit Sub
        ElseIf txtInt = vbNullString Then
            MsgBox "Please fill-up " & """" & "INTERVAL" & """" & " to complete the task.", vbOKOnly + vbCritical + vbApplicationModal, "Java Effects"
            Exit Sub
        Else
            If InStr(txtMen(0), ".") = 0 Or InStr(txtMen(1), ".") = 0 Or InStr(txtMen(2), ".") = 0 Or InStr(txtMen(3), ".") = 0 Then
                MsgBox "Invalid Web Links.", vbOKOnly + vbApplicationModal + vbCritical, "Java Effects"
                Exit Sub
            Else
                If InStr(txtMen(0), ".") = Len(txtMen(0)) Or InStr(txtMen(1), ".") = Len(txtMen(1)) Or InStr(txtMen(2), ".") = Len(txtMen(2)) Or InStr(txtMen(3), ".") = Len(txtMen(3)) Then
                    MsgBox "Invalid Web Links.", vbOKOnly + vbApplicationModal + vbCritical, "Java Effects"
                    Exit Sub
                End If
            End If
        End If
    ElseIf cmbOptions.ListIndex = 6 Then
        If txtInt = vbNullString Then
            MsgBox "Please fill-up " & """" & "INTERVAL" & """" & " to complete the task.", vbOKOnly + vbCritical + vbApplicationModal, "Java Effects"
            Exit Sub
        End If
    ElseIf cmbOptions.ListIndex = 8 Or cmbOptions.ListIndex = 9 Then
        If txtInt = vbNullString Then
            MsgBox "Please fill-up " & """" & "INTERVAL" & """" & " to complete the task.", vbOKOnly + vbCritical + vbApplicationModal, "Java Effects"
            Exit Sub
        End If
    ElseIf cmbOptions.ListIndex = 14 Then
        If txtInt = vbNullString Then
            MsgBox "Please fill-up " & """" & "EXPIRATION DATE" & """" & " to complete the task.", vbOKOnly + vbCritical + vbApplicationModal, "Java Effects"
            Exit Sub
        End If
    ElseIf cmbOptions.ListIndex = 17 Then
        If txtMen(0) = vbNullString Or txtMen(1) = vbNullString Or txtMen(2) = vbNullString Or txtMen(3) = vbNullString Then
            MsgBox "Please fill-up " & """" & "MENUS" & """" & " to complete the task.", vbOKOnly + vbCritical + vbApplicationModal, "Java Effects"
            Exit Sub
        Else
            If InStr(txtMen(0), ".") = 0 Or InStr(txtMen(1), ".") = 0 Or InStr(txtMen(2), ".") = 0 Or InStr(txtMen(3), ".") = 0 Then
                MsgBox "Invalid Web Links.", vbOKOnly + vbApplicationModal + vbCritical, "Java Effects"
                Exit Sub
            Else
                If InStr(txtMen(0), ".") = Len(txtMen(0)) Or InStr(txtMen(1), ".") = Len(txtMen(1)) Or InStr(txtMen(2), ".") = Len(txtMen(2)) Or InStr(txtMen(3), ".") = Len(txtMen(3)) Then
                    MsgBox "Invalid Web Links.", vbOKOnly + vbApplicationModal + vbCritical, "Java Effects"
                    Exit Sub
                End If
            End If
        End If
    ElseIf cmbOptions.ListIndex = 18 Then
        If txtMes = vbNullString Then
            MsgBox "Please fill-up " & """" & "TEXT" & """" & " to complete the task.", vbOKOnly + vbCritical + vbApplicationModal, "Java Effects"
            Exit Sub
        End If
    End If
    MousePointer = FormMouse(Me, 11)
    Mode = False
    Generating = True
    Call ResetValue(Mode)
    LoopTimes = 0
    txtOutput = vbNullString
    tmrProcess.Enabled = True
    cmdTextGen.Enabled = False
    cmbOptions.Enabled = False
    cmdCopy.Enabled = False
    txtOutput.Enabled = False
    txtMes.Enabled = False
    Frame3.Enabled = False
    txtInt.Enabled = False
    VS.Enabled = False
    cmdClear.Enabled = False
    cmdCreate.Enabled = False
    On Error Resume Next
    txtMes.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
    If KeyCode = vbKeyReturn Then If Not cmdTextGen.Enabled = False Then cmdTextGen_Click
End Sub

Private Sub Form_Load()
    Sleep 1000
    If Not (Not UCase$(GetSetting(Caption, "WindowState", "350")) = vbNullString) = UCase$("false") Then
        WindowState = Val(Trim$(GetSetting(Caption, "WindowState", "350")))
    Else
        WindowState = 0
        SaveSetting Caption, "WindowState", "350", WindowState
    End If
    If ApplicationRunning Then
        MsgBox "Java Effects can only run one at a time!", vbOKOnly + vbExclamation + vbApplicationModal, "Java Effects"
        End
    End If
    If WindowState = 0 Then
        Height = 7845
        Width = 10950
    End If
    cmbOptions.Clear
    With cmbOptions
        .AddItem "From Left"
        .AddItem "From Right"
        .AddItem "Letter By Letter"
        .AddItem "Text Run Out"
        .AddItem "Sneezy Text"
        .AddItem "Scrolling Text"
        .AddItem "Mouse Trail Clock"
        .AddItem "No Right Click (No Alert)"
        .AddItem "Shake Me"
        .AddItem "Shake Screen"
        .AddItem "Movable Menu"
        .AddItem "Portable Menu"
        .AddItem "Background Changer"
        .AddItem "Search The Dictionary"
        .AddItem "Visit Counter"
        .AddItem "Search The Internet"
        .AddItem "IP Grabber"
        .AddItem "Multi Site Search"
        .AddItem "Remove Ads"
    End With
    VS.Min = 1000
    VS.Max = 10
    VS.Value = 10
    txtInt = VS.Value
    If Not ClipboardWatcher Then
        cmdClear.Enabled = False
    Else
        cmdClear.Enabled = True
    End If
    Mode = False
    ResetValue Mode
End Sub

Private Sub Form_Resize()
    If Not WindowState = 1 Then
        SaveSetting Caption, "WindowState", "350", WindowState
        If Height < 7845 Or Width < 10950 Then
            Width = 10950
            Height = 7845
            Exit Sub
        End If
        Label1.Left = 120
        Label1.Top = 120
        txtMes.Top = 120
        txtMes.Left = Label1.Left + Label1.Width + 120
        txtMes.Width = Width - (Label1.Left + Label1.Width + (120 * 3))
        Label4.Top = (120 * 2) + txtMes.Height
        Label4.Left = (Width / 2) - (Label4.Width / 2)
        txtOutput.Left = 120
        txtOutput.Width = Width - (120 * 3)
        txtOutput.Top = ((120 * 3) + txtMes.Height + Label4.Height)
        Frame1.Top = Height - (Frame1.Height + 120) - (120 * 4)
        Frame1.Left = 120
        Frame3.Top = Frame1.Top
        Frame3.Left = Frame1.Left + Frame1.Width + 120
        txtOutput.Height = Frame1.Top - 120 - txtOutput.Top
        Frame2.Top = Frame3.Top
        Frame2.Left = Frame3.Left + Frame3.Width + 120
        lblDisplay.Top = Frame2.Top + Frame2.Height
        lblDisplay.Left = Frame2.Left
        picProgress.Top = lblDisplay.Top + lblDisplay.Height
        picProgress.Left = lblDisplay.Left
        cmdCreate.Top = picProgress.Top + picProgress.Height + 120
        cmdCreate.Left = picProgress.Left
    End If
    If WindowState = 0 Then
        Left = (Screen.Width / 2) - (Width / 2)
        Top = (Screen.Height / 2) - (Height / 2)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Generating = False Then
        Beep
        If MsgBox("Are you sure do you want to quit now?", vbYesNo + vbQuestion + vbApplicationModal + vbDefaultButton2, Caption) = vbYes Then
            Sleep 1000
            SaveSetting Caption, "WindowState", "350", WindowState
            End
        End If
    Else
        MsgBox "Application is buzy generating your request. Please wait until the application has been finished what it's doing.", vbOKOnly + vbCritical + vbApplicationModal, "Java Effects"
    End If
    Cancel = 1
End Sub

Private Sub tmrClipboard_Timer()
    If Not tmrProcess.Enabled = True Then
        If Not ClipboardWatcher Then
            cmdClear.Enabled = False
        Else
            cmdClear.Enabled = True
        End If
    End If
End Sub

Private Sub tmrProcess_Timer()
    Randomize
    If Repeat = False Then
        If Wait = False Then
            If Process >= 80 Then
                If Process >= 90 Then
                    tmrProcess.Interval = (Rnd * 100) + 10
                    MoreThan = True
                Else
                    tmrProcess.Interval = (Rnd * 500) + 10
                End If
            Else
                tmrProcess.Interval = (Rnd * 1000) + 100
            End If
            If MoreThan = False Then
                If Process >= 10 Then
                    Process = Process + (Rnd * 5) + 1
                Else
                    Process = Process + (Rnd * 3) + 1
                End If
            Else
                Process = Process + (Rnd * 3) + 1
            End If
            If Process > 100 Then
                Process = 100
                tmrProcess.Interval = (Rnd * 1000) + 100
                Wait = True
            End If
            lblDisplay = "Generating Script..."
            ProgBar picProgress, CLng(Process), , &H800000
            picProgress.Refresh
        Else
            If HideUnhide = False Then
                HideUnhide = True
            Else
                GoTo EndLoop
            End If
        End If
    Else
        GoTo EndLoop
    End If
    If Wait = True Then GoTo EndLoop
    Exit Sub
    
EndLoop:
    Mode = False
    ResetValue Mode
    LoopTimes = LoopTimes + 1
    Sleep 3000
    ProgBar picProgress, CLng(Process), , &H800000
    picProgress.Refresh
    Sleep 1000
    Call Watcher(Me)
    cmdCopy.Enabled = True
    tmrProcess.Enabled = False
    cmbOptions_Click
    txtOutput.Enabled = True
    cmbOptions.Enabled = True
    cmdCreate.Enabled = True
    MousePointer = FormMouse(Me, 0)
    Generating = False
    lblDisplay = vbNullString
End Sub

Private Sub txtInt_GotFocus()
    With txtInt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtInt_KeyPress(KeyAscii As Integer)
    Key = Chr$(KeyAscii)
    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        If Val(Trim$(txtInt)) < VS.Max Then
            txtInt = VS.Max
            VS.Value = VS.Max
            On Error Resume Next
            With txtInt
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
            Exit Sub
        End If
        If Not Val(Trim$(txtInt)) > VS.Min Then
            VS.Value = Val(Trim$(txtInt))
            On Error Resume Next
            With txtInt
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
            Exit Sub
        Else
            txtInt = VS.Min
            VS.Value = VS.Min
            On Error Resume Next
            With txtInt
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
        End If
    End If
    Select Case Key
        Case "0" To "9"
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtInt_LostFocus()
    If txtInt = vbNullString Or txtInt < VS.Max Then
        txtInt = VS.Max
        VS.Value = VS.Max
    ElseIf txtInt > VS.Min Then
        txtInt = VS.Min
        VS.Value = VS.Min
    End If
End Sub

Private Sub txtMen_GotFocus(Index As Integer)
    With txtMen(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMes_Change()
    If Not txtMes = vbNullString Or Not ClipboardWatcher = False Then
        cmdClear.Enabled = True
    Else
        cmdClear.Enabled = False
    End If
End Sub

Private Sub txtMes_GotFocus()
    With txtMes
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtOutput_GotFocus()
    On Error Resume Next
    If Not txtMes.Enabled = False Then
        txtMes.SetFocus
    Else
        cmbOptions.SetFocus
    End If
End Sub

Private Sub VS_Change()
    txtInt = VS.Value
    On Error Resume Next
    With txtInt
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
