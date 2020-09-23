VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fox - Info   Options"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   Icon            =   "Fox-InfofrmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   15
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Updating Speed   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   7095
      Begin VB.ComboBox PagingFileUpdateSpeed 
         Height          =   315
         Left            =   3240
         TabIndex        =   14
         Text            =   "Normal Speed ( 1000 milliseconds )"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.ComboBox RamUpdateSpeed 
         Height          =   315
         Left            =   3240
         TabIndex        =   12
         Text            =   "Normal Speed ( 1000 milliseconds )"
         Top             =   1200
         Width           =   3375
      End
      Begin VB.ComboBox CpuUpdateSpeed 
         Height          =   315
         Left            =   3240
         TabIndex        =   10
         Text            =   "Normal Speed ( 1000 milliseconds )"
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Paging File Update Speed :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ram Update Speed :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   2190
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cpu Update Speed :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   2115
      End
   End
   Begin VB.PictureBox lblGrey 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5880
      MouseIcon       =   "Fox-InfofrmOptions.frx":0ECA
      MousePointer    =   99  'Custom
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox lblBlue 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5280
      MouseIcon       =   "Fox-InfofrmOptions.frx":101C
      MousePointer    =   99  'Custom
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox lblYellow 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4680
      MouseIcon       =   "Fox-InfofrmOptions.frx":116E
      MousePointer    =   99  'Custom
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox lblRed 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4080
      MouseIcon       =   "Fox-InfofrmOptions.frx":12C0
      MousePointer    =   99  'Custom
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox lblWhite 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3480
      MouseIcon       =   "Fox-InfofrmOptions.frx":1412
      MousePointer    =   99  'Custom
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.PictureBox lblGreen 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2880
      MouseIcon       =   "Fox-InfofrmOptions.frx":1564
      MousePointer    =   99  'Custom
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   2
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPU Progress bar color :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   2595
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2887
      TabIndex        =   0
      Top             =   120
      Width           =   1800
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim TMP
Dim DbFile
Dim PwdString
Private Sub TOPmostPOSITIONonSCREEN_Click()
TMP = SetTopMostWindow(frmOptions.hwnd, True)
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub CpuUpdateSpeed_Click()
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Data\Options.mdb")
    PwdString = "swordfishofvolen"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("tblOptions", dbOpenTable)

Select Case CpuUpdateSpeed.ListIndex
Case 4
rs.Edit
rs("CPUupdatespeed") = 250
rs.Update
Form2.Timer1.Interval = 250
Case 3
rs.Edit
rs("CPUupdatespeed") = 500
rs.Update
Form2.Timer1.Interval = 500
Case 2
rs.Edit
rs("CPUupdatespeed") = 1000
rs.Update
Form2.Timer1.Interval = 1000
Case 1
rs.Edit
rs("CPUupdatespeed") = 2000
rs.Update
Form2.Timer1.Interval = 2000
Case 0
rs.Edit
rs("CPUupdatespeed") = 3000
rs.Update
Form2.Timer1.Interval = 3000
End Select
End Sub

Private Sub Form_Load()
Dim DbFile
Dim PwdString
TOPmostPOSITIONonSCREEN_Click
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
RamUpdateSpeed.AddItem "Very Slow ( 3000 milliseconds )"
RamUpdateSpeed.AddItem "Slow ( 2000 milliseconds )"
RamUpdateSpeed.AddItem "Normal Speed ( 1000 milliseconds )"
RamUpdateSpeed.AddItem "Fast ( 500 milliseconds )"
RamUpdateSpeed.AddItem "Very Fast ( 250 milliseconds )"
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
CpuUpdateSpeed.AddItem "Very Slow ( 3000 milliseconds )"
CpuUpdateSpeed.AddItem "Slow ( 2000 milliseconds )"
CpuUpdateSpeed.AddItem "Normal Speed ( 1000 milliseconds )"
CpuUpdateSpeed.AddItem "Fast ( 500 milliseconds )"
CpuUpdateSpeed.AddItem "Very Fast ( 250 milliseconds )"
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
PagingFileUpdateSpeed.AddItem "Very Slow ( 3000 milliseconds )"
PagingFileUpdateSpeed.AddItem "Slow ( 2000 milliseconds )"
PagingFileUpdateSpeed.AddItem "Normal Speed ( 1000 milliseconds )"
PagingFileUpdateSpeed.AddItem "Fast ( 500 milliseconds )"
PagingFileUpdateSpeed.AddItem "Very Fast ( 250 milliseconds )"
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Data\Options.mdb")
    PwdString = "swordfishofvolen"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("tblOptions", dbOpenTable)
Select Case rs("CPUupdatespeed")
Case 250
CpuUpdateSpeed.ListIndex = 4
Case 500
CpuUpdateSpeed.ListIndex = 3
Case 1000
CpuUpdateSpeed.ListIndex = 2
Case 2000
CpuUpdateSpeed.ListIndex = 1
Case 3000
CpuUpdateSpeed.ListIndex = 0
End Select

Select Case rs("PagingFileUpdateSpeed")
Case 250
PagingFileUpdateSpeed.ListIndex = 4
Case 500
PagingFileUpdateSpeed.ListIndex = 3
Case 1000
PagingFileUpdateSpeed.ListIndex = 2
Case 2000
PagingFileUpdateSpeed.ListIndex = 1
Case 3000
PagingFileUpdateSpeed.ListIndex = 0
End Select

Select Case rs("Ramupdatespeed")
Case 250
RamUpdateSpeed.ListIndex = 4
Case 500
RamUpdateSpeed.ListIndex = 3
Case 1000
RamUpdateSpeed.ListIndex = 2
Case 2000
RamUpdateSpeed.ListIndex = 1
Case 3000
RamUpdateSpeed.ListIndex = 0
End Select
End Sub

Private Sub lblBlue_Click()
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Data\Options.mdb")
    PwdString = "swordfishofvolen"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("tblOptions", dbOpenTable)

rs.Edit
rs("CPUcolor") = lblBlue.BackColor
rs.Update

Form2.CPUBAR.FillColor = lblBlue.BackColor
End Sub

Private Sub lblGreen_Click()
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Data\Options.mdb")
    PwdString = "swordfishofvolen"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("tblOptions", dbOpenTable)

rs.Edit
rs("CPUcolor") = lblGreen.BackColor
rs.Update

Form2.CPUBAR.FillColor = lblGreen.BackColor
End Sub

Private Sub lblGrey_Click()
Set WS = DBEngine.Workspaces(0)
   DbFile = (App.Path & "\Data\Options.mdb")
    PwdString = "swordfishofvolen"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("tblOptions", dbOpenTable)

rs.Edit
rs("CPUcolor") = lblGrey.BackColor
rs.Update

Form2.CPUBAR.FillColor = lblGrey.BackColor
End Sub

Private Sub lblRed_Click()
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Data\Options.mdb")
    PwdString = "swordfishofvolen"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("tblOptions", dbOpenTable)

rs.Edit
rs("CPUcolor") = lblRed.BackColor
rs.Update

Form2.CPUBAR.FillColor = lblRed.BackColor
End Sub

Private Sub lblWhite_Click()
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Data\Options.mdb")
    PwdString = "swordfishofvolen"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("tblOptions", dbOpenTable)

rs.Edit
rs("CPUcolor") = lblWhite.BackColor
rs.Update

Form2.CPUBAR.FillColor = lblWhite.BackColor
End Sub

Private Sub lblYellow_Click()
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Data\Options.mdb")
    PwdString = "swordfishofvolen"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("tblOptions", dbOpenTable)

rs.Edit
rs("CPUcolor") = lblYellow.BackColor
rs.Update

Form2.CPUBAR.FillColor = lblYellow.BackColor
End Sub

Private Sub PagingFileUpdateSpeed_Click()
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Data\Options.mdb")
    PwdString = "swordfishofvolen"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("tblOptions", dbOpenTable)

Select Case PagingFileUpdateSpeed.ListIndex
Case 4
rs.Edit
rs("PagingFileUpdateSpeed") = 250
rs.Update
Form2.Timer3.Interval = 250
Case 3
rs.Edit
rs("PagingFileUpdateSpeed") = 500
rs.Update
Form2.Timer3.Interval = 500
Case 2
rs.Edit
rs("PagingFileUpdateSpeed") = 1000
rs.Update
Form2.Timer3.Interval = 1000
Case 1
rs.Edit
rs("PagingFileUpdateSpeed") = 2000
rs.Update
Form2.Timer3.Interval = 2000
Case 0
rs.Edit
rs("PagingFileUpdateSpeed") = 3000
rs.Update
Form2.Timer3.Interval = 3000
End Select
End Sub

Private Sub RamUpdateSpeed_Click()
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Data\Options.mdb")
    PwdString = "swordfishofvolen"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("tblOptions", dbOpenTable)

Select Case RamUpdateSpeed.ListIndex
Case 4
rs.Edit
rs("RamUpdateSpeed") = 250
rs.Update
Form2.RAMtimer.Interval = 250
Case 3
rs.Edit
rs("RamUpdateSpeed") = 500
rs.Update
Form2.RAMtimer.Interval = 500
Case 2
rs.Edit
rs("RamUpdateSpeed") = 1000
rs.Update
Form2.RAMtimer.Interval = 1000
Case 1
rs.Edit
rs("RamUpdateSpeed") = 2000
rs.Update
Form2.RAMtimer.Interval = 2000
Case 0
rs.Edit
rs("RamUpdateSpeed") = 3000
rs.Update
Form2.RAMtimer.Interval = 3000
End Select
End Sub
