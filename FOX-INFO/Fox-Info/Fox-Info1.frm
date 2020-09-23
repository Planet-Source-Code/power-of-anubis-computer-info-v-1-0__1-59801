VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fox-Info v 1.0"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10095
   Icon            =   "Fox-Info1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   195
      Left            =   9240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   30
      Text            =   "Fox-Info1.frx":08CA
      Top             =   360
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComDlg.CommonDialog LogSave 
      Left            =   8760
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox LISTFORIP 
      Height          =   255
      Left            =   9240
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9480
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   135
      Left            =   9360
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      MouseIcon       =   "Fox-Info1.frx":08D0
      TabCaption(0)   =   "Computer Profile"
      TabPicture(0)   =   "Fox-Info1.frx":08EC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "List1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "System Info"
      TabPicture(1)   =   "Fox-Info1.frx":0908
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "processorlbl"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "wanIPADDRESSlbl"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "WinOSlbl"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label6"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "ramlbl"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "ramfreelbl"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "ScreenResolutionLBL"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label8"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label9"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "numberofprocessorslbl"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label10"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "systemdrivelbl"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label11"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "runningprogramslbl"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label12"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "printerlbl"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label15"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "DXVERSIONlbl"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label13"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "totalpagingfilelbl"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label14"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "freepagingfilelbl"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label16"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "networkipaddresslBl"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "RAMtimer"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Timer3"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).ControlCount=   30
      TabCaption(2)   =   "Drives Space"
      TabPicture(2)   =   "Fox-Info1.frx":0924
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvwHD"
      Tab(2).Control(1)=   "Command2"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Processes"
      TabPicture(3)   =   "Fox-Info1.frx":0940
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LabelProcesses"
      Tab(3).Control(1)=   "Processeslbl"
      Tab(3).Control(2)=   "LvW"
      Tab(3).Control(3)=   "terminateprocess"
      Tab(3).Control(4)=   "Command1"
      Tab(3).Control(5)=   "Timer2"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Bios Info"
      TabPicture(4)   =   "Fox-Info1.frx":095C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "ListView1"
      Tab(4).Control(1)=   "Command3"
      Tab(4).ControlCount=   2
      Begin VB.Timer Timer3 
         Interval        =   1000
         Left            =   -66000
         Top             =   960
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Get Bios Info"
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
         Left            =   -74760
         TabIndex        =   39
         Top             =   5640
         Width           =   2175
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5055
         Left            =   -74760
         TabIndex        =   38
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   8916
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Empty The Recycle Bin"
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
         Left            =   -68280
         TabIndex        =   31
         Top             =   5640
         Width           =   2655
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -65880
         Top             =   360
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh List"
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
         Left            =   -70920
         TabIndex        =   29
         Top             =   5400
         Width           =   2295
      End
      Begin VB.CommandButton terminateprocess 
         Caption         =   "Terminate Selected"
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
         Left            =   -68520
         TabIndex        =   28
         Top             =   5400
         Width           =   2655
      End
      Begin MSComctlLib.ListView LvW 
         Height          =   4575
         Left            =   -74520
         TabIndex        =   19
         Top             =   720
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8070
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvwHD 
         Height          =   4935
         Left            =   -74760
         TabIndex        =   18
         Top             =   600
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   8705
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   12648447
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Timer RAMtimer 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -66000
         Top             =   480
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         ItemData        =   "Fox-Info1.frx":0978
         Left            =   360
         List            =   "Fox-Info1.frx":097A
         TabIndex        =   2
         Top             =   600
         Width           =   8895
      End
      Begin VB.Label networkipaddresslBl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -72360
         TabIndex        =   45
         Top             =   2160
         Width           =   75
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Network Ip Address :"
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
         Left            =   -74640
         TabIndex        =   44
         Top             =   2160
         Width           =   2160
      End
      Begin VB.Label freepagingfilelbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -72720
         TabIndex        =   43
         Top             =   4320
         Width           =   75
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Free Paging File :"
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
         Left            =   -74640
         TabIndex        =   42
         Top             =   4320
         Width           =   1860
      End
      Begin VB.Label totalpagingfilelbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -72600
         TabIndex        =   41
         Top             =   3960
         Width           =   75
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Paging File :"
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
         Left            =   -74640
         TabIndex        =   40
         Top             =   3960
         Width           =   1920
      End
      Begin VB.Label DXVERSIONlbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -72840
         TabIndex        =   37
         Top             =   5400
         Width           =   75
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DirectX Version :"
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
         Left            =   -74640
         TabIndex        =   36
         Top             =   5400
         Width           =   1740
      End
      Begin VB.Label printerlbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -73800
         TabIndex        =   35
         Top             =   5040
         Width           =   75
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Printer :"
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
         Left            =   -74640
         TabIndex        =   34
         Top             =   5040
         Width           =   810
      End
      Begin VB.Label runningprogramslbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -71520
         TabIndex        =   33
         Top             =   4680
         Width           =   75
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of running programs :"
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
         Left            =   -74640
         TabIndex        =   32
         Top             =   4680
         Width           =   3060
      End
      Begin VB.Label systemdrivelbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -73080
         TabIndex        =   27
         Top             =   1440
         Width           =   75
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System Drive :"
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
         Left            =   -74640
         TabIndex        =   26
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label numberofprocessorslbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -72120
         TabIndex        =   25
         Top             =   1080
         Width           =   75
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of processors :"
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
         Left            =   -74640
         TabIndex        =   24
         Top             =   1080
         Width           =   2430
      End
      Begin VB.Label Processeslbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -72600
         TabIndex        =   22
         Top             =   5520
         Width           =   75
      End
      Begin VB.Label LabelProcesses 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Processes :"
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
         Left            =   -74520
         TabIndex        =   21
         Top             =   5520
         Width           =   1845
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Screen Resolution :"
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
         Left            =   -74640
         TabIndex        =   20
         Top             =   2520
         Width           =   2040
      End
      Begin VB.Label ScreenResolutionLBL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -72480
         TabIndex        =   17
         Top             =   2520
         Width           =   75
      End
      Begin VB.Label ramfreelbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -71400
         TabIndex        =   16
         Top             =   3600
         Width           =   75
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Free Physical Memory (RAM) :"
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
         Left            =   -74640
         TabIndex        =   15
         Top             =   3600
         Width           =   3150
      End
      Begin VB.Label ramlbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -71280
         TabIndex        =   14
         Top             =   3240
         Width           =   75
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Physical Memory (RAM) :"
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
         Left            =   -74640
         TabIndex        =   13
         Top             =   3240
         Width           =   3210
      End
      Begin VB.Label WinOSlbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -72480
         TabIndex        =   12
         Top             =   2880
         Width           =   75
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operating System :"
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
         Left            =   -74640
         TabIndex        =   11
         Top             =   2880
         Width           =   1980
      End
      Begin VB.Label wanIPADDRESSlbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fox-Info can't get your Wan Ip Address."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -72720
         MouseIcon       =   "Fox-Info1.frx":097C
         TabIndex        =   9
         Top             =   1800
         Width           =   4065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wan Ip Address :"
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
         Left            =   -74640
         TabIndex        =   8
         Top             =   1800
         Width           =   1785
      End
      Begin VB.Label processorlbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   -73320
         TabIndex        =   7
         Top             =   720
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Processor :"
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
         Left            =   -74640
         TabIndex        =   6
         Top             =   720
         Width           =   1200
      End
   End
   Begin VB.Label lblCheckedHDSPACE 
      Caption         =   "0"
      Height          =   135
      Left            =   9360
      TabIndex        =   23
      Top             =   240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label CPUPercentage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPU USAGE"
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
      Left            =   10080
      TabIndex        =   5
      Top             =   0
      Width           =   1320
   End
   Begin VB.Shape CPUBAR 
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1560
      Top             =   7920
      Width           =   8295
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   375
      Left            =   1560
      Top             =   7920
      Width           =   8295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPU Usage"
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
      TabIndex        =   3
      Top             =   8040
      Width           =   1230
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   0
      Picture         =   "Fox-Info1.frx":0ACE
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fox-Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   1440
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLog1 
         Caption         =   "Make a log file of ""Computer Profile"""
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuLog2 
         Caption         =   "Make a log file of ""System Info"""
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuMinimizeFoxInfo 
         Caption         =   "Minimize Fox-Info"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuAlwaysOnTop 
         Caption         =   "Always On Top"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuSEPARATOR 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions2 
         Caption         =   "Options"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************
'*  _ __  (_) _ _  ___  _ __      _ _  ___  / _|_____  *
'* | '_ ` | || '_\/ _ \| '_ `    / __|/ _ \| |_ _   _| *
'* | | | || || |   (_) | | | |   \__ \ (_) |  _| | |   *
'* |_| |_||_||_|  \___/|_| |_|   |___/\___/|_|   |_|   *
'*                                                     *
'*******************************************************
'
'Fox-Info v 1.0
'Copyright © : 2005
'
'Thank's to http://www.planet-source-code.com form where I used few codes.
'I worked hard to make this program, so please vote.




Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal numBytes As Long)
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Private Declare Function GetTickCount Lib "Kernel32.dll" () As Long
Dim Information As New GetInfomation
Private Type CounterInfo
    hCounter As Long
    strName As String
End Type
Dim db As Database
Dim rs As Recordset
Dim WS As Workspace
Dim RProgramS
Dim PRINTERA
Dim datata
Dim VERSIATANABIOS
Dim DbFile
Dim PwdString
Dim TMP
Dim pdhStatus As PDH_STATUS
Dim hQuery As Long
Dim Counters(0 To 99) As CounterInfo
Dim currentCounterIdx As Long
Dim iPerformanceDetail As PERF_DETAIL
Dim BARVALUE As Integer
Dim PCName As String
Dim Ppp As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type
Dim PROCESSORSINFO
Dim a1
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
'RECYCLE BIN
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4
'--
Dim ListTotal As Integer, i, j, k, L
Dim intCnt As Integer
'---SCREEN RESOLUTION------
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Const ENUM_CURRENT_SETTINGS = &HFFFF - 1
Private Type DEVMODE
    dmDeviceName      As String * 32
    dmSpecVersion     As Integer
    dmDriverVersion   As Integer
    dmSize            As Integer
    dmDriverExtra     As Integer
    dmFields          As Long
    dmOrientation     As Integer
    dmPaperSize       As Integer
    dmPaperLength     As Integer
    dmPaperWidth      As Integer
    dmScale           As Integer
    dmCopies          As Integer
    dmDefaultSource   As Integer
    dmPrintQuality    As Integer
    dmColor           As Integer
    dmDuplex          As Integer
    dmYResolution     As Integer
    dmTTOption        As Integer
    dmCollate         As Integer
    dmFormName        As String * 32
    dmUnusedPadding   As Integer
    dmBitsPerPel      As Integer
    dmPelsWidth       As Long
    dmPelsHeight      As Long
    dmDisplayFlags    As Long
    dmDisplayFrequency As Long
End Type

Private Sub Command2_Click()
On Error Resume Next
Dim EmptyNowInOrder
EmptyNowInOrder = SHEmptyRecycleBin(Form2.hwnd, "", SHERB_NOPROGRESSUI)
ERRORA:

End Sub



Private Sub Command3_Click()
   ListView1.ListItems.Clear
   Call wmiBiosInfo
   Call lvAutosizeControl(ListView1)
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuAlwaysOnTop_Click()
If mnuAlwaysOnTop.Checked = True Then
mnuAlwaysOnTop.Checked = False
TMP = SetTopMostWindow(Form2.hwnd, False)
Else
mnuAlwaysOnTop.Checked = True
TMP = SetTopMostWindow(Form2.hwnd, True)
End If
End Sub
Private Sub Command1_Click()
processeslisting
End Sub

Private Sub TOPmostPOSITIONonSCREEN_Click()
TMP = SetTopMostWindow(Form2.hwnd, True)
End Sub

Private Sub Form_Load()
Dim cpubarcolor
Set WS = DBEngine.Workspaces(0)
    DbFile = (App.Path & "\Data\Options.mdb")
    PwdString = "swordfishofvolen"
Set db = DBEngine.OpenDatabase(DbFile, False, False, ";PWD=" & PwdString)
Set rs = db.OpenRecordset("tblOptions", dbOpenTable)
cpubarcolor = rs("CPUcolor")
CPUBAR.FillColor = cpubarcolor
Timer1.Interval = rs("CPUupdatespeed")
TOPmostPOSITIONonSCREEN_Click
pdhStatus = PdhOpenQuery(0, 1, hQuery)
    If pdhStatus <> ERROR_SUCCESS Then
        MsgBox "Open Query failed"
       Resume Next
    End If
    AddCounter "\Processor(0)\% Processor Time", hQuery
    UpdateValues
TAB1INFOS
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuFile
End If
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuMinimizeFoxInfo_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub mnuOptions2_Click()
frmOptions.Show
End Sub

Private Sub SSTab1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SSTab1.Tab = 1 And Button = 2 Then
PopupMenu mnuFile
End If
End Sub


Private Sub mnuFile_Click()
If processorlbl.Caption = vbNullString Then
mnuLog2.Enabled = False
Else
mnuLog2.Enabled = True
End If
End Sub
Sub TAB1INFOS()
On Error Resume Next
Set Information = New GetInfomation
Dim lngTickCount As Long
Dim UpTimeWin As String
lngTickCount = GetTickCount
UpTimeWin = CStr(Round((lngTickCount / 1000 / 60))) & " Minutes"
Ppp = NameOfTheComputer(PCName)
List1.AddItem "Computer Up-Time          = " & UpTimeWin
List1.AddItem "Computer Name             = " & PCName
List1.AddItem "Current Username          = " & Information.GetUSERNAME
List1.AddItem "Country                   = " & Information.GetCountry
List1.AddItem "Language                  = " & Information.GetLanguage
List1.AddItem "System Drive              = " & Information.GetSystemDrive
List1.AddItem "Win Dir                   = " & Information.GetWinDir
List1.AddItem "Windows Temp Dir          = " & Information.TempDir
List1.AddItem "Win\System Dir            = " & Information.SystemDir
List1.AddItem "----------------"
List1.AddItem "Currency                  = " & Information.GetCurrencySymbol
List1.AddItem "Date Separator            = " & Information.GetDateSeparator
List1.AddItem "Decimal Separator         = " & Information.GetDecimalSeparator
List1.AddItem "Digit Grouping            = " & Information.GetDigitGrouping
List1.AddItem "Leading Zeros For Decimal = " & Information.GetLeadingZerosForDecimal
List1.AddItem "Long Date Format          = " & Information.GetLongDateFormat
List1.AddItem "Long Month 1              = " & Information.GetLongMonthName1
List1.AddItem "Long Month 2              = " & Information.GetLongMonthName2
List1.AddItem "Long Month 3              = " & Information.GetLongMonthName3
List1.AddItem "Long Month 4              = " & Information.GetLongMonthName4
List1.AddItem "Long Month 5              = " & Information.GetLongMonthName5
List1.AddItem "Long Month 6              = " & Information.GetLongMonthName6
List1.AddItem "Long Month 7              = " & Information.GetLongMonthName7
List1.AddItem "Long Month 8              = " & Information.GetLongMonthName8
List1.AddItem "Long Month 9              = " & Information.GetLongMonthName9
List1.AddItem "Long Month 10             = " & Information.GetLongMonthName10
List1.AddItem "Long Month 11             = " & Information.GetLongMonthName11
List1.AddItem "Long Month 12             = " & Information.GetLongMonthName12
List1.AddItem "Long Day 1                = " & Information.GetLongNameDay1
List1.AddItem "Long Day 2                = " & Information.GetLongNameDay2
List1.AddItem "Long Day 3                = " & Information.GetLongNameDay3
List1.AddItem "Long Day 4                = " & Information.GetLongNameDay4
List1.AddItem "Long Day 5                = " & Information.GetLongNameDay5
List1.AddItem "Long Day 6                = " & Information.GetLongNameDay6
List1.AddItem "Long Day 7                = " & Information.GetLongNameDay7
List1.AddItem "Negative Sign             = " & Information.GetNegativeSign
List1.AddItem "Negative Sign Position    = " & Information.GetNegativeSignPosition
List1.AddItem "Number Fractional Digits  = " & Information.GetNumberOfFractionalDigits
List1.AddItem "Positive Sign             = " & Information.GetPositiveSign
List1.AddItem "Positive Sign Position    = " & Information.GetPositiveSignPosition
List1.AddItem "Short Date Format         = " & Information.GetShortDateFormat
List1.AddItem "Short Month 1             = " & Information.GetShortMonthName1
List1.AddItem "Short Month 2             = " & Information.GetShortMonthName2
List1.AddItem "Short Month 3             = " & Information.GetShortMonthName3
List1.AddItem "Short Month 4             = " & Information.GetShortMonthName4
List1.AddItem "Short Month 5             = " & Information.GetShortMonthName5
List1.AddItem "Short Month 6             = " & Information.GetShortMonthName6
List1.AddItem "Short Month 7             = " & Information.GetShortMonthName7
List1.AddItem "Short Month 8             = " & Information.GetShortMonthName8
List1.AddItem "Short Month 9             = " & Information.GetShortMonthName9
List1.AddItem "Short Month 10            = " & Information.GetShortMonthName10
List1.AddItem "Short Month 11            = " & Information.GetShortMonthName11
List1.AddItem "Short Month 12            = " & Information.GetShortMonthName12
List1.AddItem "Short Day 1               = " & Information.GetShortNameDay1
List1.AddItem "Short Day 2               = " & Information.GetShortNameDay2
List1.AddItem "Short Day 3               = " & Information.GetShortNameDay3
List1.AddItem "Short Day 4               = " & Information.GetShortNameDay4
List1.AddItem "Short Day 5               = " & Information.GetShortNameDay5
List1.AddItem "Short Day 6               = " & Information.GetShortNameDay6
List1.AddItem "Short Day 7               = " & Information.GetShortNameDay7
List1.AddItem "Thousand Separator        = " & Information.GetThousandSeparator
List1.AddItem "Time Format               = " & Information.GetTimeFormat
List1.AddItem "Time Separator            = " & Information.GetTimeSeparator
Timer1.Enabled = True 'Physical Memory Timer Update Values
End Sub
Public Sub AddCounter(strCounterName As String, hQuery As Long)
    Dim pdhStatus As PDH_STATUS
    Dim hCounter As Long
    
    pdhStatus = PdhVbAddCounter(hQuery, strCounterName, hCounter)
    Counters(currentCounterIdx).hCounter = hCounter
    Counters(currentCounterIdx).strName = strCounterName
    currentCounterIdx = currentCounterIdx + 1
End Sub

Private Sub UpdateValues()
    Dim dblCounterValue As Double
    Dim pdhStatus As Long
    Dim strInfo As String
    Dim i As Long
        
    PdhCollectQueryData (hQuery)
    
    i = 0  'Only one counter but you can add more

    dblCounterValue = _
            PdhVbGetDoubleCounterValue(Counters(i).hCounter, pdhStatus)
        
        'Some error checking, make sure the query went through
        If (pdhStatus = PDH_CSTATUS_VALID_DATA) _
        Or (pdhStatus = PDH_CSTATUS_NEW_DATA) Then
        PB1.Value = dblCounterValue
        End If
        

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    PdhCloseQuery (hQuery)
End Sub

Private Sub mnuLog1_Click()
Dim PlaceToSave As String

LogSave.DialogTitle = "Make Log File..."
 LogSave.Filter = "Text File (*.txt)|*.txt|"
   LogSave.ShowSave
    PlaceToSave = LogSave.FileName
    

Open PlaceToSave For Output As #1
      List1.ListIndex = 0
      
On Error Resume Next

    For i = 0 To List1.ListCount
       Print #1, List1
       List1.ListIndex = List1.ListIndex + 1
    Next i
Close #1

End Sub

Private Sub mnuLog2_Click()
Dim Prr
Dim Lrr
Dim Crr
Dim Drr
Dim Krr
Dim Orr
Dim Hrr
Dim Nrr
Dim JJJ
Dim KKLL
Dim KKKJJJ
Dim HOHO
Dim WOW
Dim BBBj
Dim PlaceToSave As String
Prr = Label3.Caption & " " & processorlbl.Caption
Lrr = Label9.Caption & " " & numberofprocessorslbl.Caption
Crr = Label10.Caption & " " & systemdrivelbl.Caption
Drr = Label4.Caption & " " & wanIPADDRESSlbl.Caption
Krr = Label8.Caption & " " & ScreenResolutionLBL.Caption
Orr = Label5.Caption & " " & WinOSlbl.Caption
Hrr = Label6.Caption & " " & ramlbl.Caption
Nrr = Label7.Caption & " " & ramfreelbl.Caption
JJJ = Label11.Caption & " " & runningprogramslbl.Caption
KKLL = Label12.Caption & " " & printerlbl.Caption
KKKJJJ = Label15.Caption & " " & DXVERSIONlbl.Caption
HOHO = Label13.Caption & " " & totalpagingfilelbl.Caption
WOW = Label14.Caption & " " & freepagingfilelbl.Caption
BBBj = Label16.Caption & " " & networkipaddresslBl.Caption
Text1.Text = Prr & vbNewLine & Lrr & vbNewLine & Crr & vbNewLine & Drr & vbNewLine & BBBj & vbNewLine & Krr & vbNewLine & Orr & vbNewLine & Hrr & vbNewLine & Nrr & vbNewLine & JJJ & vbNewLine & KKLL & vbNewLine & KKKJJJ & vbNewLine & HOHO & vbNewLine & WOW

LogSave.DialogTitle = "Make Log File..."
 LogSave.Filter = "Text File (*.txt)|*.txt|"
  LogSave.ShowSave
   PlaceToSave = LogSave.FileName
    
On Error GoTo ERRORA

 Open PlaceToSave For Output As #1
  Print #1, Text1.Text
     Close #1
ERRORA:
Exit Sub
End Sub

Private Sub RAMtimer_Timer()
    Call GlobalMemoryStatus(memInfo)
        ramlbl.Caption = memInfo.dwTotalPhys / 1024 & " KB"
        ramfreelbl.Caption = memInfo.dwAvailPhys / 1024 & " KB"

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
Select Case SSTab1.Tab
Case 0
List1.Clear
TAB1INFOS
RAMtimer.Enabled = False
Case 1
On Error Resume Next
networkipaddresslBl.Caption = GetIPAddress()
Dim MS As MEMORYSTATUS
MS.dwLength = Len(MS)
GlobalMemoryStatus MS
totalpagingfilelbl.Caption = Format$(MS.dwTotalPageFile / 1024, "###,###,###,###") & " Kbyte"
freepagingfilelbl.Caption = Format$(MS.dwAvailPageFile / 1024, "###,###,###,###") & " Kbyte"
DXVERSIONlbl.Caption = GetDirectXVersion
PRINTERA = ReadKey("HKEY_CURRENT_USER\Printers\DeviceOld")
RProgramS = ReadKey("HKEY_CURRENT_USER\SessionInformation\ProgramCount")
printerlbl.Caption = PRINTERA
runningprogramslbl.Caption = RProgramS
numberofprocessorslbl.Caption = Environ("Number_Of_Processors")
systemdrivelbl.Caption = Environ("SystemDrive")
RAMtimer.Enabled = True
Dim os As OSVERSIONINFO
Dim m As Long
Dim mv As Long
Dim pd As Long
Dim miv As Long
'--------------------
    Dim curDPS As DEVMODE
    Dim colors As String
    Dim SMR As Long
    
    SMR = EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, curDPS)
    
    If SMR = 0 Then
        ScreenResolutionLBL.Caption = "Error evaluating the current screen resolution!"
    Else
        Select Case curDPS.dmBitsPerPel
            Case 4:      colors = "16 Color"
            Case 8:      colors = "256 Color"
            Case 16:     colors = "High Color"
            Case 24, 32: colors = "True Color"
        End Select
        ScreenResolutionLBL.Caption = Format(curDPS.dmPelsWidth, "@@@@") + " x " + _
                      Format(curDPS.dmPelsHeight, "@@@@") + "  " + _
                      Format(colors, "@@@@@@@@@@@@@  ") + _
                      Format(curDPS.dmDisplayFrequency, "@@@ Hz")
    End If
'--------------------
os.dwOSVersionInfoSize = Len(os)
m = GetVersionEx(os)
mv = os.dwMajorVersion
pd = os.dwPlatformId
miv = os.dwMinorVersion
If pd = 2 Then WinOSlbl.Caption = "Windows NT" & " " & mv & "." & miv
If pd = 1 Then
If miv = 10 Then WinOSlbl.Caption = "Windows 98"
If miv = 0 Then WinOSlbl.Caption = "Windows 95"
If miv = 90 Then WinOSlbl.Caption = "Windows Me"
End If
PROCESSORSINFO = ReadKey("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0\Processornamestring")
processorlbl.Caption = PROCESSORSINFO
wanIPADDRESSlbl.Caption = GetPublicIP()
Case 2
If lblCheckedHDSPACE.Caption = "0" Then
FillListview
lblCheckedHDSPACE.Caption = "1"
Else
End If
Case 3
processeslisting
Case 4
With ListView1
      .ListItems.Clear
      .ColumnHeaders.Clear
      .ColumnHeaders.Add , , "WMI Property"
      .ColumnHeaders.Add , , "Value(s)"
      .View = lvwReport
      .Sorted = False
   End With
   Command3_Click
End Select
End Sub

Sub processeslisting()
'----------
Dim header As ColumnHeader
LvW.View = lvwReport
LvW.ColumnHeaders.Clear
Set header = LvW.ColumnHeaders.Add(, "first", "Processes", LvW.Width / 4 * 3) 'set listview width
Set header = LvW.ColumnHeaders.Add(, "second", "ID", LvW.Width - LvW.Width / 4 * 3)
LvW.Refresh
'----------
Dim ret
Dim TheLoopingProcess
Dim proc As PROCESSENTRY32
Dim snap As Long
Dim exename As String
LvW.ListItems.Clear 'clear listview contents
snap = CreateToolhelpSnapshot(TH32CS_SNAPall, 0) 'get snapshot handle
proc.dwSize = Len(proc)
TheLoopingProcess = ProcessFirst(snap, proc)       'first process and return value
Processeslbl.Caption = -1
i = 0
While TheLoopingProcess <> 0      'next process
exename = proc.szExeFile
ret = LvW.ListItems.Add(, "first" & CStr(i), exename)   'add process name to listview
LvW.ListItems("first" & CStr(i)).SubItems(1) = proc.th32ProcessID   'add process ID to listview
Processeslbl.Caption = Processeslbl.Caption + 1
i = i + 1
TheLoopingProcess = ProcessNext(snap, proc)
Wend
CloseHandle snap
End Sub

Private Sub terminateprocess_Click()
  Dim i As Integer
  Dim Counter As Integer
  Dim lngSuccess As Long
  Dim dblPID As Double
    
    Counter = LvW.ListItems.Count
    For i = 1 To Counter
        With LvW.ListItems.Item(i)
            If .Selected = True Then
                KillProcessById (.SubItems(1))
            End If
        End With
    Next i
processeslisting
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()

Dim CpuValueCreating
    UpdateValues
    PB1.Value = Round(PB1.Value, 0)
    CPUPercentage.Left = Shape1.Left + Shape1.Width / 2 - CPUPercentage.Width / 2
    CPUPercentage.Top = Shape1.Top + Shape1.Height / 2 - CPUPercentage.Height / 2
    CPUPercentage.Caption = PB1.Value & "%"
    BARVALUE = PB1.Value
    CpuValueCreating = CPUpersonalBar(BARVALUE, CPUBAR)
End Sub

Public Function CPUpersonalBar(ProgressBarValue As Integer, ProgressBarName As Shape)
ProgressBarName.Width = ProgressBarValue * 82.95 ' Write here the value for one percent
End Function
Public Function NameOfTheComputer(MachineName As String) As Long
    Dim NameSize As Long
    Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
End Function


Sub FillListview()
    On Error Resume Next
    Dim strDrive As String
    Dim strMessage As String
    Dim rtn
    Dim fs, H, s, bt As String
    Dim Check, Counter
    Dim varible, lWidth
    Dim a, b, C, n, d As String
    lvwHD.ListItems.Clear
    lWidth = lvwHD.Width / 5
    lvwHD.ColumnHeaders.Clear
    lvwHD.ColumnHeaders.Add , , "Name", lWidth
    lvwHD.ColumnHeaders.Add , , "Type", lWidth
    lvwHD.ColumnHeaders.Add , , "Total Size", lWidth
    lvwHD.ColumnHeaders.Add , , "Free Space", lWidth
    lvwHD.ColumnHeaders.Add , , "Used Space", lWidth - 60
    lvwHD.View = lvwReport
    lvwHD.ListItems.Clear
    Check = True: Counter = 65
    For Counter = 65 To 86
        strDrive = Chr(Counter)
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set H = fs.GetDrive(fs.GetDriveName(strDrive + ":\"))
        Select Case GetDriveType(strDrive + ":\")
            Case DRIVE_FIXED
                If H.FreeSpace < 1024 ^ 3 Then
                    bt = " MB"
                    n = H.volumename & " (" & UCase(strDrive & ":") & ")"
                    a = "Local Disk"
                    b = Left(Format(H.TotalSize / 1048576, "#,##.00"), 3) & bt
                    C = Left(Format(H.FreeSpace / 1048576, "#,##.00"), 3) & bt
                ElseIf H.FreeSpace > 1024 ^ 3 Then
                    bt = " GB"
                    n = H.volumename & " (" & UCase(strDrive & ":") & ")"
                    a = "Local Disk"
                    b = Left(Format(H.TotalSize / 1071576, "#,##.00"), 4) & bt
                    C = Left(Format(H.FreeSpace / 1071576, "#,##.00"), 4) & bt
                End If
                k = H.TotalSize - H.FreeSpace
                If k < 1024 ^ 3 Then
                    If k < 1024 ^ 2 Then
                        bt = " KB"
                        k = H.TotalSize - H.FreeSpace
                        d = Left(Format(k / 1024, "#,##.00"), 4) & bt
                    Else
                        bt = " MB"
                        k = H.TotalSize - H.FreeSpace
                        d = Left(Format(k / 1048576, "#,##.00"), 3) & bt
                    End If
                ElseIf k > 1024 ^ 3 Then
                    bt = " GB"
                    k = H.TotalSize - H.FreeSpace
                    d = Left(Format(k / 1071576, "#,##.00"), 4) & bt
                End If
                Set varible = lvwHD.ListItems.Add(, , n)
                varible.SubItems(1) = a
                varible.SubItems(2) = b
                varible.SubItems(3) = C
                varible.SubItems(4) = d
                i = i + H.TotalSize
                j = j + H.FreeSpace
                L = L + k
                rtn = "Hard Drive"
            Case DRIVE_REMOTE
                rtn = "Network Drive"
            Case DRIVE_CDROM
                rtn = "CD-ROM Drive"
            Case DRIVE_RAMDISK
                rtn = "RAM Disk"
            Case Else
                rtn = ""
        End Select
    Next Counter
    Set varible = lvwHD.ListItems.Add(, , "-----------------")
    varible.SubItems(1) = "-----------------"
    varible.SubItems(2) = "-----------------"
    varible.SubItems(3) = "-----------------"
    varible.SubItems(4) = "-----------------"
    Set varible = lvwHD.ListItems.Add(, , "Totals")
    varible.SubItems(1) = lvwHD.ListItems.Count - 2 & " Local Disks"
    varible.SubItems(2) = Left(Format(i / 1071576, "#,##.00"), 5) & " GB"
    varible.SubItems(3) = Left(Format(j / 1071576, "#,##.00"), 5) & " GB"
    varible.SubItems(4) = Left(Format(L / 1071576, "#,##.00"), 5) & " GB"
    lvwHD.FlatScrollBar = False
End Sub


Private Sub Timer2_Timer()
processeslisting
Timer2.Enabled = False
End Sub

  
Function GetDirectXVersion() As String
Dim handle As Long

Dim resString As String
Dim strVersion As String

Dim resBinary() As Byte
  
If RegOpenKeyEx(&H80000002, "SOFTWARE\Microsoft\DirectX", 0, &H20019, handle) Then Exit Function
  
  
ReDim resBinary(1023) As Byte
  
Call RegQueryValueEx(handle, "Version", 0, 0, resBinary(0), 1024)
  
resString = Space$(1023)
CopyMemory ByVal resString, resBinary(0), 1023
  
RegCloseKey handle
  
resString = Left(resString, 12)
  
Select Case resString
    Case "4.02.0095"
        GetDirectXVersion = "1.0"
    Case "4.03.00.1096"
        GetDirectXVersion = "2.0"
    Case "4.04.0068", "4.04.0069"
        GetDirectXVersion = "3.0"
    Case "4.05.00.0155"
        GetDirectXVersion = "5.0"
    Case "4.05.01.1721", "4.05.01.1998"
        GetDirectXVersion = "5.0"
    Case "4.06.02.0436"
        GetDirectXVersion = "6.0"
    Case "4.07.00.0700"
        GetDirectXVersion = "7.0"
    Case "4.07.00.0716"
        GetDirectXVersion = "7.0a"
    Case "4.08.00.0400"
        GetDirectXVersion = "8.0"
    Case "4.08.01.0881", "4.08.01.0810"
        GetDirectXVersion = "8.1"
    Case "4.09.0000.0900"
        GetDirectXVersion = "9.0"
    Case "4.09.0000.0901"
        GetDirectXVersion = "9.0a"
    Case "4.09.0000.0902"
        GetDirectXVersion = "9.0b"
    Case "4.09.00.0904"
        GetDirectXVersion = "9.0c"
End Select
  
End Function

Private Sub lvAutosizeControl(lv As ListView)

   Dim col2adjust As Long

  '/* Size each column based on the maximum of
  '/* EITHER the columnheader text width, or,
  '/* if the items below it are wider, the
  '/* widest list item in the column
   For col2adjust = 0 To lv.ColumnHeaders.Count - 1
   
      Call SendMessage(lv.hwnd, _
                       LVM_SETCOLUMNWIDTH, _
                       col2adjust, _
                       ByVal LVSCW_AUTOSIZE_USEHEADER)

   Next
   
End Sub


Private Sub wmiBiosInfo()
      
   Dim BiosSet As SWbemObjectSet
   Dim bios As SWbemObject
   Dim itmx As ListItem
   Dim cnt As Long
   Dim msg As String
   
   Set BiosSet = GetObject("winmgmts:{impersonationLevel=impersonate}"). _
                                      InstancesOf("Win32_BIOS")
   
   On Local Error Resume Next
   
   For Each bios In BiosSet
   
      Set itmx = ListView1.ListItems.Add(, , "PrimaryBIOS")
      itmx.SubItems(1) = bios.PrimaryBIOS
            
      Set itmx = ListView1.ListItems.Add(, , "Status")
      itmx.SubItems(1) = bios.Status
      
      For cnt = LBound(bios.BIOSVersion) To UBound(bios.BIOSVersion)
         Set itmx = ListView1.ListItems.Add(, , IIf(cnt = 0, "BIOSVersion strings", ""))
         itmx.SubItems(1) = bios.BIOSVersion(cnt)
      Next
      
      Set itmx = ListView1.ListItems.Add(, , "Caption")
      itmx.SubItems(1) = bios.Caption
      
      Set itmx = ListView1.ListItems.Add(, , "Description")
      itmx.SubItems(1) = bios.Description
      
      Set itmx = ListView1.ListItems.Add(, , "Name")
      itmx.SubItems(1) = bios.Name

      Set itmx = ListView1.ListItems.Add(, , "Manufacturer")
      itmx.SubItems(1) = bios.Manufacturer

      Set itmx = ListView1.ListItems.Add(, , "ReleaseDate")
      itmx.SubItems(1) = bios.ReleaseDate

      Set itmx = ListView1.ListItems.Add(, , "SerialNumber")
      itmx.SubItems(1) = bios.SerialNumber

      Set itmx = ListView1.ListItems.Add(, , "SMBIOSBIOSVersion")
      itmx.SubItems(1) = bios.SMBIOSBIOSVersion
      
      Set itmx = ListView1.ListItems.Add(, , "SMBIOSMajorVersion")
      itmx.SubItems(1) = bios.SMBIOSMajorVersion
      
      Set itmx = ListView1.ListItems.Add(, , "SMBIOSMinorVersion")
      itmx.SubItems(1) = bios.SMBIOSMinorVersion

      Set itmx = ListView1.ListItems.Add(, , "SMBIOSPresent")
      itmx.SubItems(1) = bios.SMBIOSPresent
      
      Set itmx = ListView1.ListItems.Add(, , "SoftwareElementID")
      itmx.SubItems(1) = bios.SoftwareElementID
      
      Set itmx = ListView1.ListItems.Add(, , "SoftwareElementState")
      Select Case bios.SoftwareElementState
         Case 0: msg = "deployable"
         Case 1: msg = "installable"
         Case 2: msg = "executable"
         Case 3: msg = "running"
      End Select
      itmx.SubItems(1) = msg
      
      Set itmx = ListView1.ListItems.Add(, , "Version")
      itmx.SubItems(1) = bios.Version

      Set itmx = ListView1.ListItems.Add(, , "InstallableLanguages")
      itmx.SubItems(1) = bios.InstallableLanguages

      Set itmx = ListView1.ListItems.Add(, , "CurrentLanguage")
      itmx.SubItems(1) = bios.CurrentLanguage
        
      For cnt = LBound(bios.ListOfLanguages) To UBound(bios.ListOfLanguages)
      
         Set itmx = ListView1.ListItems.Add(, , IIf(cnt = 0, "ListOfLanguages", ""))
         itmx.SubItems(1) = bios.ListOfLanguages(cnt)
         
      Next cnt

      For cnt = LBound(bios.BiosCharacteristics) To UBound(bios.BiosCharacteristics)
      
         Set itmx = ListView1.ListItems.Add(, , IIf(cnt = 0, "BIOS Characteristics", ""))
      
         Select Case bios.BiosCharacteristics(cnt)
            Case 0: msg = "reserved"
            Case 1: msg = "reserved"
            Case 2: msg = "unknown"
            Case 3: msg = "BIOS characteristics not supported"
            Case 4: msg = "ISA supported"
            Case 5: msg = "MCA supported"
            Case 6: msg = "EISA supported"
            Case 7: msg = "PCI supported"
            Case 8: msg = "PC Card (PCMCIA) supported"
            Case 9: msg = "Plug and Play supported"
            Case 10: msg = "APM is supported"
            Case 11: msg = "BIOS upgradable (Flash)"
            Case 12: msg = "BIOS shadowing allowed"
            Case 13: msg = "VL-VESA supported"
            Case 14: msg = "ESCD support available"
            Case 15: msg = "Boot from CD supported"
            Case 16: msg = "Selectable boot supported"
            Case 17: msg = "BIOS ROM socketed"
            Case 18: msg = "Boot from PC card (PCMCIA) supported"
            Case 19: msg = "EDD (Enhanced Disk Drive) specification supported"
            Case 20: msg = "Int 13h, Japanese Floppy for NEC 9800 1.2mb (3.5, 1k b/s, 360 RPM) supported"
            Case 21: msg = "Int 13h, Japanese Floppy for Toshiba 1.2mb (3.5, 360 RPM) supported"
            Case 22: msg = "Int 13h, 5.25 / 360 KB floppy services supported"
            Case 23: msg = "Int 13h, 5.25 /1.2MB floppy services supported"
            Case 24: msg = "Int 13h 3.5 / 720 KB floppy services supported"
            Case 25: msg = "Int 13h, 3.5 / 2.88 MB floppy services supported"
            Case 26: msg = "Int 5h, print screen service supported"
            Case 27: msg = "Int 9h, 8042 keyboard services supported"
            Case 28: msg = "Int 14h, serial services supported"
            Case 29: msg = "Int 17h, printer services supported"
            Case 30: msg = "Int 10h, CGA/Mono video aervices supported"
            Case 31: msg = "NEC PC-98"
            Case 32: msg = "ACPI supported"
            Case 33: msg = "USB Legacy supported"
            Case 34: msg = "AGP supported"
            Case 35: msg = "I2O boot supported"
            Case 36: msg = "LS-120 boot supported"
            Case 37: msg = "ATAPI ZIP drive boot supported"
            Case 38: msg = "1394 boot supported"
            Case 39: msg = "Smart battery supported"
         End Select
         
         itmx.SubItems(1) = msg
         
      Next  'For cnt
      
   Next  'For Each bios

End Sub

Private Sub Timer3_Timer()
Dim MS As MEMORYSTATUS
MS.dwLength = Len(MS)
GlobalMemoryStatus MS
totalpagingfilelbl.Caption = Format$(MS.dwTotalPageFile / 1024, "###,###,###,###") & " Kbyte"
freepagingfilelbl.Caption = Format$(MS.dwAvailPageFile / 1024, "###,###,###,###") & " Kbyte"

End Sub
