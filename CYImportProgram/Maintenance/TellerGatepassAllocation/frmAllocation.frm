VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAllocation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CCR Allocation "
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab sstAlloc 
      Height          =   11055
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   19500
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Export "
      TabPicture(0)   =   "frmAllocation.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdExit(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdNext(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdDel(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdEdit(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdAdd(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraDetails(0)"
      Tab(0).Control(6)=   "grdAlloc(0)"
      Tab(0).Control(7)=   "cmdRefresh(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "CFS Import "
      TabPicture(1)   =   "frmAllocation.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRefresh(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "grdAlloc(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraDetails(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdAdd(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdEdit(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdDel(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdNext(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdExit(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "CY Import "
      TabPicture(2)   =   "frmAllocation.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "cmdRefresh(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdAdd(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdEdit(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "grdAlloc(2)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "fraDetails(2)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdDel(2)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdNext(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdExit(2)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Empties"
      TabPicture(3)   =   "frmAllocation.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdNext(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdDel(3)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdEdit(3)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdAdd(3)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdExit(3)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "fraDetails(3)"
      Tab(3).Control(6)=   "grdAlloc(3)"
      Tab(3).Control(7)=   "cmdRefresh(3)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "CY Special Service"
      TabPicture(4)   =   "frmAllocation.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdRefresh(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "grdAlloc(4)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "fraDetails(4)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "cmdAdd(4)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "cmdEdit(4)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cmdDel(4)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "cmdNext(4)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "cmdExit(4)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).ControlCount=   8
      Begin VB.CommandButton cmdNext 
         Caption         =   "F11=Next"
         Height          =   735
         Index           =   3
         Left            =   -66360
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "F8=Del"
         Height          =   735
         Index           =   3
         Left            =   -68280
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "F7=Edit"
         Height          =   735
         Index           =   3
         Left            =   -70200
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "F3=Exit"
         Height          =   735
         Index           =   4
         Left            =   -62880
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "F11=Next"
         Height          =   735
         Index           =   4
         Left            =   -66360
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "F8=Del"
         Height          =   735
         Index           =   4
         Left            =   -68280
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "F7=Edit"
         Height          =   735
         Index           =   4
         Left            =   -70200
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "F6=Add"
         Height          =   735
         Index           =   4
         Left            =   -72120
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.Frame fraDetails 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2775
         Index           =   4
         Left            =   -74040
         TabIndex        =   67
         Top             =   720
         Width           =   13095
         Begin VB.TextBox txtTeller 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   4
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtStartCCR 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   4
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtEndCCR 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   4
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Last CCR Issued"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5640
            TabIndex        =   105
            Top             =   720
            Width           =   2760
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Teller ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   104
            Top             =   720
            Width           =   1230
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Last Issue Date/Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5640
            TabIndex        =   103
            Top             =   1320
            Width           =   2715
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   102
            Top             =   1320
            Width           =   1170
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Allocated by"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   101
            Top             =   1920
            Width           =   2715
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   100
            Top             =   1920
            Width           =   1260
         End
         Begin VB.Label lblSysdte 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   4
            Left            =   8640
            TabIndex        =   74
            Top             =   2280
            Width           =   165
         End
         Begin VB.Label lblLastCCR 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   4
            Left            =   8640
            TabIndex        =   73
            Top             =   720
            Width           =   165
         End
         Begin VB.Label lblLastDate 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   4
            Left            =   8640
            TabIndex        =   72
            Top             =   1320
            Width           =   165
         End
         Begin VB.Label lblAllocBy 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   4
            Left            =   8640
            TabIndex        =   71
            Top             =   1920
            Width           =   165
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "F6=Add"
         Height          =   735
         Index           =   3
         Left            =   -72120
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "F3=Exit"
         Height          =   735
         Index           =   3
         Left            =   -62880
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.Frame fraDetails 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2775
         Index           =   3
         Left            =   -74040
         TabIndex        =   52
         Top             =   720
         Width           =   13095
         Begin VB.TextBox txtEndCCR 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   3
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox txtStartCCR 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   3
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtTeller 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   3
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Last CCR Issued"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5640
            TabIndex        =   99
            Top             =   720
            Width           =   2760
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Teller ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   98
            Top             =   720
            Width           =   1230
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Last Issue Date/Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5640
            TabIndex        =   97
            Top             =   1320
            Width           =   2715
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   96
            Top             =   1320
            Width           =   1170
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Allocated by"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   95
            Top             =   1920
            Width           =   2715
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   94
            Top             =   1920
            Width           =   1260
         End
         Begin VB.Label lblAllocBy 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   3
            Left            =   8640
            TabIndex        =   59
            Top             =   1920
            Width           =   165
         End
         Begin VB.Label lblLastDate 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   3
            Left            =   8640
            TabIndex        =   58
            Top             =   1320
            Width           =   165
         End
         Begin VB.Label lblLastCCR 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   3
            Left            =   8640
            TabIndex        =   57
            Top             =   720
            Width           =   165
         End
         Begin VB.Label lblSysdte 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   3
            Left            =   8640
            TabIndex        =   56
            Top             =   2280
            Width           =   165
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "F3=Exit"
         Height          =   735
         Index           =   2
         Left            =   12120
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "F3=Exit"
         Height          =   735
         Index           =   1
         Left            =   -62880
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "F3=Exit"
         Height          =   735
         Index           =   0
         Left            =   -62880
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "F11=Next"
         Height          =   735
         Index           =   2
         Left            =   8640
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "F11=Next"
         Height          =   735
         Index           =   1
         Left            =   -66360
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "F11=Next"
         Height          =   735
         Index           =   0
         Left            =   -66360
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "F8=Del"
         Height          =   735
         Index           =   2
         Left            =   6720
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "F8=Del"
         Height          =   735
         Index           =   1
         Left            =   -68280
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "F7=Edit"
         Height          =   735
         Index           =   1
         Left            =   -70200
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "F6=Add"
         Height          =   735
         Index           =   1
         Left            =   -72120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "F8=Del"
         Height          =   735
         Index           =   0
         Left            =   -68280
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "F7=Edit"
         Height          =   735
         Index           =   0
         Left            =   -70200
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "F6=Add"
         Height          =   735
         Index           =   0
         Left            =   -72120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.Frame fraDetails 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2775
         Index           =   2
         Left            =   960
         TabIndex        =   45
         Top             =   720
         Width           =   13095
         Begin VB.TextBox txtTeller 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   2
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtStartCCR 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   2
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtEndCCR 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   2
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   93
            Top             =   1920
            Width           =   1260
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Allocated by"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   92
            Top             =   1920
            Width           =   2715
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   91
            Top             =   1320
            Width           =   1170
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Last Issue Date/Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5640
            TabIndex        =   90
            Top             =   1320
            Width           =   2715
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Teller ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   89
            Top             =   720
            Width           =   1230
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Last CCR Issued"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   88
            Top             =   720
            Width           =   2760
         End
         Begin VB.Label lblSysdte 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   2
            Left            =   8640
            TabIndex        =   51
            Top             =   2280
            Width           =   165
         End
         Begin VB.Label lblLastCCR 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   2
            Left            =   8640
            TabIndex        =   48
            Top             =   720
            Width           =   165
         End
         Begin VB.Label lblLastDate 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   2
            Left            =   8640
            TabIndex        =   47
            Top             =   1320
            Width           =   165
         End
         Begin VB.Label lblAllocBy 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   2
            Left            =   8640
            TabIndex        =   46
            Top             =   1920
            Width           =   165
         End
      End
      Begin VB.Frame fraDetails 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2775
         Index           =   1
         Left            =   -74040
         TabIndex        =   41
         Top             =   720
         Width           =   13095
         Begin VB.TextBox txtEndCCR 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   1
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox txtStartCCR 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   1
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtTeller 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   1
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            TabIndex        =   87
            Top             =   1920
            Width           =   1260
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Allocated by"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   86
            Top             =   1920
            Width           =   2715
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   85
            Top             =   1320
            Width           =   1170
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Last Issue Date/Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5640
            TabIndex        =   84
            Top             =   1320
            Width           =   2715
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Teller ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   83
            Top             =   720
            Width           =   1230
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Last CCR Issued"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5640
            TabIndex        =   82
            Top             =   720
            Width           =   2760
         End
         Begin VB.Label lblSysdte 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   1
            Left            =   8640
            TabIndex        =   50
            Top             =   2280
            Width           =   165
         End
         Begin VB.Label lblAllocBy 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   1
            Left            =   8640
            TabIndex        =   44
            Top             =   1920
            Width           =   165
         End
         Begin VB.Label lblLastDate 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   1
            Left            =   8640
            TabIndex        =   43
            Top             =   1320
            Width           =   165
         End
         Begin VB.Label lblLastCCR 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   1
            Left            =   8640
            TabIndex        =   42
            Top             =   720
            Width           =   165
         End
      End
      Begin VB.Frame fraDetails 
         BackColor       =   &H80000010&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   2775
         Index           =   0
         Left            =   -74040
         TabIndex        =   11
         Top             =   720
         Width           =   13095
         Begin VB.TextBox txtTeller 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   0
            Left            =   2400
            MaxLength       =   10
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtStartCCR 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   0
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtEndCCR 
            BackColor       =   &H80000010&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000018&
            Height          =   345
            Index           =   0
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label lblSysdte 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   0
            Left            =   8640
            TabIndex        =   49
            Top             =   2280
            Width           =   165
         End
         Begin VB.Label lblLastCCR 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   0
            Left            =   8640
            TabIndex        =   40
            Top             =   720
            Width           =   165
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Last CCR Issued"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5640
            TabIndex        =   39
            Top             =   720
            Width           =   2880
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Teller ID"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   38
            Top             =   720
            Width           =   1230
         End
         Begin VB.Label lblLastDate 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   0
            Left            =   8640
            TabIndex        =   37
            Top             =   1320
            Width           =   165
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Last Issue Date/Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5640
            TabIndex        =   26
            Top             =   1320
            Width           =   2715
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   25
            Top             =   1320
            Width           =   1170
         End
         Begin VB.Label lblAllocBy 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            ForeColor       =   &H80000018&
            Height          =   300
            Index           =   0
            Left            =   8640
            TabIndex        =   24
            Top             =   1920
            Width           =   165
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "Allocated by"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   5640
            TabIndex        =   13
            Top             =   1920
            Width           =   2715
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H80000010&
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   840
            TabIndex        =   12
            Top             =   1920
            Width           =   1260
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdAlloc 
         Height          =   5295
         Index           =   0
         Left            =   -74040
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   3840
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   9340
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorSel    =   -2147483646
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLines       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdAlloc 
         Height          =   5295
         Index           =   1
         Left            =   -74040
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3840
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   9340
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorSel    =   -2147483646
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLines       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdAlloc 
         Height          =   5295
         Index           =   2
         Left            =   960
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3840
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   9340
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorSel    =   -2147483646
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLines       =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdAlloc 
         Height          =   5295
         Index           =   3
         Left            =   -74040
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   3840
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   9340
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorSel    =   -2147483646
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLines       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid grdAlloc 
         Height          =   5295
         Index           =   4
         Left            =   -74040
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   3840
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   9340
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         BackColor       =   -2147483624
         BackColorSel    =   -2147483646
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLines       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "F5=Refresh"
         Height          =   735
         Index           =   0
         Left            =   -74040
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "F5=Refresh"
         Height          =   735
         Index           =   1
         Left            =   -74040
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "F7=Edit"
         Height          =   735
         Index           =   2
         Left            =   4800
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "F6=Add"
         Height          =   735
         Index           =   2
         Left            =   2880
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "F5=Refresh"
         Height          =   735
         Index           =   2
         Left            =   960
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "F5=Refresh"
         Height          =   735
         Index           =   3
         Left            =   -74040
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "F5=Refresh"
         Height          =   735
         Index           =   4
         Left            =   -74040
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   9600
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Index As Integer
Dim EditSw As Boolean
Dim LosFocs As Boolean
Dim VASwch As Boolean   ' exit from the RowColChange event when validating an allocation

Const CCR = 0
Const CFM = 1
Const CYM = 2
Const CYE = 3
Const SPL = 4

Private Sub cmdRefresh_Click(Index As Integer)
  Call FillGrid
End Sub

Private Sub Form_Load()
    Call Main
    EditSw = False
    LosFocs = False
    VASwch = False
    Do Until Index > 4
        Call SetGrid
        Index = Index + 1
    Loop
End Sub

Private Sub Main()
Dim c As clsCCRAllocation

gConnStr = GetConnString
Set c = New clsCCRAllocation
With c
    .ConnectByStr (gConnStr)
    .Userid = zCurrentUser
End With

End Sub

Private Function GetConnString() As String
Dim xCnt As Integer

Open App.Path & "\BILLINGDB.INI" For Binary Access Read As #1

Do While Not EOF(1)
    xCnt = xCnt + 1
    Select Case xCnt
        Case 1
            Line Input #1, GetConnString
    End Select
Loop
    GetConnString = Trim(GetConnString)
    Close #1
End Function

Private Sub SetGrid()
    Dim ColHdgs(7) As String
    Dim ColWide(7) As String
    Dim ColCounter As Integer
        
    ColHdgs(0) = "Teller ID"
    ColHdgs(1) = "  Start"
    ColHdgs(2) = "   End"
    ColHdgs(3) = "Last Issued"
    ColHdgs(4) = "Last Issue Date/Time"
    ColHdgs(5) = "Allocated By"
    ColHdgs(6) = ""
    ColHdgs(7) = "Company Code" 'PRNH - Company Code
    
    ColWide(0) = 1700
    ColWide(1) = 1800
    ColWide(2) = 1800
    ColWide(3) = 2000
    ColWide(4) = 3500
    ColWide(5) = 2220
    ColWide(6) = 0
    ColWide(7) = 2220 'PRNH
       
    With grdAlloc(Index)    ' Set column heading and width
        Do Until ColCounter > 7
            .Col = ColCounter: .Row = 0: .Text = ColHdgs(ColCounter)
            .ColWidth(ColCounter) = ColWide(ColCounter)
            ColCounter = ColCounter + 1
        Loop
    End With
    Call FillGrid
End Sub

Private Sub FillGrid()
    Dim rstCCR As New ADODB.Recordset
    Dim rstCFM As New ADODB.Recordset
    Dim rstCYM As New ADODB.Recordset
    Dim rstCYE As New ADODB.Recordset
    Dim rstSPL As New ADODB.Recordset
   
    If Index > 4 Then
        Index = 0
    End If
    grdAlloc(Index).Rows = 2
    
    Select Case Index
        Case CCR
            With grdAlloc(CCR)
                rstCCR.Open "Select * from CCRAlloc order by strccr desc", gcnnBilling, , , adCmdText
                .Row = 0
                Do Until rstCCR.EOF
                    If .Row > 0 Then
                        .AddItem ("")
                    End If
                    .Row = .Row + 1
                    .Col = 0: .Text = "" & rstCCR.Fields("teller")
                    .Col = 1: .Text = "" & rstCCR.Fields("strccr")
                    .Col = 2: .Text = "" & rstCCR.Fields("endccr")
                    .Col = 3: .Text = "" & rstCCR.Fields("prvccr")
                    .Col = 4: .Text = "" & rstCCR.Fields("prvdte")
                    .Col = 5: .Text = "" & rstCCR.Fields("userid")
                    .Col = 6: .Text = "" & rstCCR.Fields("sysdte")
                    
                    'PRNH - Added CompanyCode
                    .Col = 7: .Text = "" & rstCCR.Fields("CompanyCode")
                    rstCCR.MoveNext
                Loop
                rstCCR.Close
                .Row = 1: .Col = 0: .ColSel = 6
            End With
        Case CFM
            With grdAlloc(CFM)
                rstCFM.Open "Select * from CFMAlloc order by strgps desc", gcnnBilling, , , adCmdText
                .Row = 0
                Do Until rstCFM.EOF
                    If .Row > 0 Then
                        .AddItem ("")
                    End If
                    .Row = .Row + 1
                    .Col = 0: .Text = "" & rstCFM.Fields("teller")
                    .Col = 1: .Text = "" & rstCFM.Fields("strgps")
                    .Col = 2: .Text = "" & rstCFM.Fields("endgps")
                    .Col = 3: .Text = "" & rstCFM.Fields("prvgps")
                    .Col = 4: .Text = "" & rstCFM.Fields("prvdte")
                    .Col = 5: .Text = "" & rstCFM.Fields("userid")
                    .Col = 6: .Text = "" & rstCFM.Fields("sysdte")
                    rstCFM.MoveNext
                Loop
                rstCFM.Close
                .Row = 1: .Col = 0: .ColSel = 6
            End With
        Case CYM
            With grdAlloc(CYM)
                rstCYM.Open "Select * from CYMAlloc order by strgps desc", gcnnBilling, , , adCmdText
                .Row = 0
                Do Until rstCYM.EOF
                    If .Row > 0 Then
                        .AddItem ("")
                    End If
                    .Row = .Row + 1
                    .Col = 0: .Text = "" & rstCYM.Fields("teller")
                    .Col = 1: .Text = "" & rstCYM.Fields("strgps")
                    .Col = 2: .Text = "" & rstCYM.Fields("endgps")
                    .Col = 3: .Text = "" & rstCYM.Fields("prvgps")
                    .Col = 4: .Text = "" & rstCYM.Fields("prvdte")
                    .Col = 5: .Text = "" & rstCYM.Fields("userid")
                    .Col = 6: .Text = "" & rstCYM.Fields("sysdte")
                    
                    'PRNH - Added CompanyCode
                    .Col = 7: .Text = "" & rstCYM.Fields("CompanyCode")
                    rstCYM.MoveNext
                Loop
                rstCYM.Close
                .Row = 1: .Col = 0: .ColSel = 6
            End With
        Case CYE
            With grdAlloc(CYE)
                rstCYE.Open "Select * from CYEAlloc order by strgps desc", gcnnBilling, , , adCmdText
                .Row = 0
                Do Until rstCYE.EOF
                    If .Row > 0 Then
                        .AddItem ("")
                    End If
                    .Row = .Row + 1
                    .Col = 0: .Text = "" & rstCYE.Fields("teller")
                    .Col = 1: .Text = "" & rstCYE.Fields("strgps")
                    .Col = 2: .Text = "" & rstCYE.Fields("endgps")
                    .Col = 3: .Text = "" & rstCYE.Fields("prvgps")
                    .Col = 4: .Text = "" & rstCYE.Fields("prvdte")
                    .Col = 5: .Text = "" & rstCYE.Fields("userid")
                    .Col = 6: .Text = "" & rstCYE.Fields("sysdte")
                    rstCYE.MoveNext
                Loop
                rstCYE.Close
                .Row = 1: .Col = 0: .ColSel = 6
            End With
        Case SPL
            With grdAlloc(SPL)
                rstSPL.Open "Select * from SPLAlloc order by strccr desc", gcnnBilling, , , adCmdText
                .Row = 0
                Do Until rstSPL.EOF
                    If .Row > 0 Then
                        .AddItem ("")
                    End If
                    .Row = .Row + 1
                    .Col = 0: .Text = "" & rstSPL.Fields("teller")
                    .Col = 1: .Text = "" & rstSPL.Fields("strccr")
                    .Col = 2: .Text = "" & rstSPL.Fields("endccr")
                    .Col = 3: .Text = "" & rstSPL.Fields("prvccr")
                    .Col = 4: .Text = "" & rstSPL.Fields("prvdte")
                    .Col = 5: .Text = "" & rstSPL.Fields("userid")
                    .Col = 6: .Text = "" & rstSPL.Fields("sysdte")
                    
                    'PRNH - Added CompanyCode
                    .Col = 7: .Text = "" & rstSPL.Fields("CompanyCode")
                    rstSPL.MoveNext
                Loop
                rstSPL.Close
                .Row = 1: .Col = 0: .ColSel = 6
            End With
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    gcnnBilling.Close
End Sub

Private Sub grdAlloc_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Call cmdExit_Click(Index)
        Case vbKeyF5
            Call cmdRefresh_Click(Index)
        Case vbKeyF6
            Call cmdAdd_Click(Index)
        Case vbKeyF7
            Call cmdEdit_Click(Index)
        Case vbKeyF8
            Call cmdDel_Click(Index)
        Case vbKeyF11
            Call FieldAdvance(KeyCode, sstAlloc, sstAlloc)
    End Select
End Sub

Private Sub grdAlloc_RowColChange(Index As Integer)
    If VASwch = True Then   ' VASwch is true when validating allocation using the grid
        Exit Sub
    End If
    With grdAlloc(Index)
        txtTeller(Index) = .TextMatrix(.Row, 0)
        txtStartCCR(Index) = .TextMatrix(.Row, 1)
        txtEndCCR(Index) = .TextMatrix(.Row, 2)
        lblLastCCR(Index) = .TextMatrix(.Row, 3)
        lblLastDate(Index) = .TextMatrix(.Row, 4)
        lblAllocBy(Index) = .TextMatrix(.Row, 5)
        lblSysdte(Index) = .TextMatrix(.Row, 6)
    End With
End Sub


Private Sub sstAlloc_Click(PreviousTab As Integer)
    SetIndex
End Sub

Private Sub sstAlloc_GotFocus()
    SetIndex
End Sub

Private Sub SetIndex()
On Error Resume Next
    Select Case sstAlloc.Tab
        Case CCR
            Index = CCR
            grdAlloc(CCR).SetFocus
        Case CFM
            Index = CFM
            grdAlloc(CFM).SetFocus
        Case CYM
            Index = CYM
            grdAlloc(CYM).SetFocus
        Case CYE
            Index = CYE
            grdAlloc(CYE).SetFocus
        Case SPL
            Index = SPL
            grdAlloc(SPL).SetFocus
    End Select
End Sub

Private Sub cmdExit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdAdd_Click(Index As Integer)
  'MsgBox Index
  On Error Resume Next
    Call DisableCommands
    Call ClearDetails
    txtTeller(Index).Enabled = True
    txtTeller(Index).SetFocus
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    If txtTeller(Index) = "" Then   ' if teller is blank, then no record exists
        grdAlloc(Index).SetFocus
        Exit Sub
    End If
    EditSw = True
    Call DisableCommands
    lblAllocBy(Index).Caption = UCase(zCurrentUser())
    lblSysdte(Index).Caption = gzGetSysDate
    txtTeller(Index).Enabled = False
    txtStartCCR(Index).SetFocus
End Sub

Private Sub cmdDel_Click(Index As Integer)
    Dim Reply As Integer
    
    If txtTeller(Index) = "" Then   ' if teller is blank, then no record exists
        grdAlloc(Index).SetFocus
        Exit Sub
    End If
    Reply = MsgBox("Delete this allocation?", vbYesNo, "Delete")
    If Reply = vbYes Then
        Call DeleteAllocation
    End If
    Call grdAlloc_RowColChange(Index)
End Sub

Private Sub cmdNext_Click(Index As Integer)
    FieldAdvance vbKeyF11, sstAlloc, sstAlloc
End Sub

Private Sub ClearDetails()
    txtTeller(Index) = ""
    txtStartCCR(Index) = ""
    txtEndCCR(Index) = ""
    'lblLastCCR(Index) = "0"
    lblLastCCR(Index) = ""
    lblLastDate(Index) = ""
    lblAllocBy(Index) = UCase(zCurrentUser())
    lblSysdte(Index) = gzGetSysDate
End Sub

Private Sub DisableCommands()
    If Index > 4 Then Index = 0
    fraDetails(Index).Enabled = True
    cmdRefresh(Index).Enabled = False
    cmdAdd(Index).Enabled = False
    cmdEdit(Index).Enabled = False
    cmdDel(Index).Enabled = False
    cmdNext(Index).Enabled = False
    cmdExit(Index).Enabled = False
    If Index = CCR Then
        sstAlloc.TabEnabled(CYM) = False
        sstAlloc.TabEnabled(CFM) = False
        sstAlloc.TabEnabled(CYE) = False
        sstAlloc.TabEnabled(SPL) = False
    ElseIf Index = CFM Then
        sstAlloc.TabEnabled(CCR) = False
        sstAlloc.TabEnabled(CYM) = False
        sstAlloc.TabEnabled(CYE) = False
        sstAlloc.TabEnabled(SPL) = False
    ElseIf Index = CYM Then
        sstAlloc.TabEnabled(CCR) = False
        sstAlloc.TabEnabled(CFM) = False
        sstAlloc.TabEnabled(CYE) = False
        sstAlloc.TabEnabled(SPL) = False
    ElseIf Index = CYE Then
        sstAlloc.TabEnabled(CCR) = False
        sstAlloc.TabEnabled(CFM) = False
        sstAlloc.TabEnabled(CYM) = False
        sstAlloc.TabEnabled(SPL) = False
    ElseIf Index = SPL Then
        sstAlloc.TabEnabled(CCR) = False
        sstAlloc.TabEnabled(CFM) = False
        sstAlloc.TabEnabled(CYM) = False
        sstAlloc.TabEnabled(CYE) = False
    End If
    grdAlloc(CCR).Enabled = False
    grdAlloc(CFM).Enabled = False
    grdAlloc(CYM).Enabled = False
    grdAlloc(CYE).Enabled = False
    grdAlloc(SPL).Enabled = False
End Sub

Private Sub EnableCommands()
    Dim X As Integer
    
    On Error Resume Next
    cmdRefresh(Index).Enabled = True
    cmdAdd(Index).Enabled = True
    cmdEdit(Index).Enabled = True
    cmdDel(Index).Enabled = True
    cmdNext(Index).Enabled = True
    cmdExit(Index).Enabled = True
    sstAlloc.TabEnabled(CCR) = True
    sstAlloc.TabEnabled(CFM) = True
    sstAlloc.TabEnabled(CYM) = True
    sstAlloc.TabEnabled(CYE) = True
    sstAlloc.TabEnabled(SPL) = True
    txtTeller(Index).Enabled = True
    fraDetails(Index).Enabled = False
    X = 0
    Do Until X > SPL
        With grdAlloc(X)
            .Enabled = True: .Row = 1: .Col = 0: .ColSel = 6
        End With
        X = X + 1
    Loop
    grdAlloc(Index).SetFocus
    Call grdAlloc_RowColChange(Index)
        
End Sub

Private Sub FieldAdvance(pKeyCode As Integer, pPreviousControl As Control, pNextControl As Control)
    Select Case pKeyCode
        Case vbKeyDown
            pNextControl.SetFocus
        Case vbKeyReturn
            pNextControl.SetFocus
        Case vbKeyUp
            If pPreviousControl.Enabled = True Then
                pPreviousControl.SetFocus
            End If
        Case vbKeyEscape
            Call EnableCommands
        Case vbKeyF11
            Select Case sstAlloc.Tab
                Case CCR
                    sstAlloc.Tab = CFM
                    grdAlloc(CFM).SetFocus
                    Index = CFM
                Case CFM
                    sstAlloc.Tab = CYM
                    grdAlloc(CYM).SetFocus
                    Index = CYM
                Case CYM
                    sstAlloc.Tab = CYE
                    grdAlloc(CYE).SetFocus
                    Index = CYE
                Case CYE
                    sstAlloc.Tab = SPL
                    grdAlloc(SPL).SetFocus
                    Index = SPL
                Case SPL
                    sstAlloc.Tab = CCR
                    grdAlloc(CCR).SetFocus
                    Index = CCR
            End Select
        Case Else
    End Select
End Sub

Private Sub txtEndCCR_GotFocus(Index As Integer)
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

'Private Sub txtEndCCR_LostFocus(Index As Integer)
'    Dim Reply As Integer
'
'    If LosFocs = True Then
'        LosFocs = False
'        Exit Sub
'    End If
'    Reply = MsgBox("Save this allocation?", vbYesNo, "Save")
'    If Reply = vbYes Then
'        Call SaveAllocation
'    End If
'    Call EnableCommands
'End Sub

Private Sub txtEndCCR_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
'            LosFocs = False
'            Call txtEndCCR_LostFocus(Index)
              Call SaveAllocation
              EnableCommands
'            LosFocs = True
        Case vbKeyUp
'            LosFocs = True
            txtStartCCR(Index).SetFocus
        Case vbKeyEscape
'            LosFocs = True
            Call EnableCommands
'            LosFocs = True
        Case Else
    End Select
End Sub

Private Sub txtStartCCR_GotFocus(Index As Integer)
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtStartCCR_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtTeller(Index), txtEndCCR(Index))
End Sub

Private Sub txtTeller_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtTeller(Index), txtStartCCR(Index))
End Sub

Private Sub txtTeller_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub SaveAllocation()
    Dim rstCCR As New ADODB.Recordset
    Dim rstCFM As New ADODB.Recordset
    Dim rstCYM As New ADODB.Recordset
    Dim rstCYE As New ADODB.Recordset
    Dim rstSPL As New ADODB.Recordset
    
    If MsgBox("Save this allocation?", vbYesNo, "Save") = vbNo Then
        Exit Sub
    End If
    
    'Check if teller is valid when adding a new allocation
    If EditSw = False Then
        If gzChkUserInfo(CStr(txtTeller(Index).Text)) = False Then
            MsgBox "Please specify a valid teller ID.", vbExclamation, "Allocation Error"
            Exit Sub
        End If
    End If
    
    If txtEndCCR(Index) = "" Then
        txtEndCCR(Index) = 0
    End If
    
    If txtStartCCR(Index) = "" Then
        txtStartCCR(Index) = 0
    End If
    
    If txtTeller(Index) = "" Or CLng(txtStartCCR(Index)) < 0 Or CLng(txtEndCCR(Index)) < 0 Or _
       CLng(txtStartCCR(Index)) > CLng(txtEndCCR(Index)) Then
            MsgBox "Please specify valid entries.", vbExclamation, "Allocation Error"
            EditSw = False
            Exit Sub
    End If
    
    If ValidAllocation = False Then
        EditSw = False
        Exit Sub
    End If
    
    If EditSw = True Then   ' if editing save valid allocation to same record
        Call EditAllocation
        EditSw = False
        Exit Sub
    End If
        
    Select Case Index
        Case CCR
            With rstCCR
                .Open "CCRAlloc", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
                .AddNew
                .Fields("teller") = txtTeller(CCR)
                .Fields("strccr") = txtStartCCR(CCR)
                .Fields("endccr") = txtEndCCR(CCR)
                .Fields("prvccr") = 0
                .Fields("wrkstn") = ""
                .Fields("userid") = lblAllocBy(CCR).Caption
                .Fields("status") = ""
                .Fields("updcde") = ""
                .Fields("sysdte") = gzGetSysDate
                .Update
                .Close
            End With
        Case CFM
            With rstCFM
                .Open "CFMAlloc", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
                .AddNew
                .Fields("teller") = txtTeller(CFM)
                .Fields("strgps") = txtStartCCR(CFM)
                .Fields("endgps") = txtEndCCR(CFM)
                .Fields("prvgps") = 0
                .Fields("wrkstn") = ""
                .Fields("userid") = lblAllocBy(CFM).Caption
                .Fields("status") = ""
                .Fields("updcde") = ""
                .Fields("sysdte") = gzGetSysDate
                .Update
                .Close
            End With
        Case CYM
            With rstCYM
                .Open "CYMAlloc", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
                .AddNew
                .Fields("teller") = txtTeller(CYM)
                .Fields("strgps") = txtStartCCR(CYM)
                .Fields("endgps") = txtEndCCR(CYM)
                .Fields("prvgps") = 0
                .Fields("wrkstn") = ""
                .Fields("userid") = lblAllocBy(CYM).Caption
                .Fields("status") = ""
                .Fields("updcde") = ""
                .Fields("sysdte") = gzGetSysDate
                .Update
                .Close
            End With
        Case CYE
            With rstCYE
                .Open "CYEAlloc", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
                .AddNew
                .Fields("teller") = txtTeller(CYE)
                .Fields("strgps") = txtStartCCR(CYE)
                .Fields("endgps") = txtEndCCR(CYE)
                .Fields("prvgps") = 0
                .Fields("wrkstn") = ""
                .Fields("userid") = lblAllocBy(CYE).Caption
                .Fields("status") = ""
                .Fields("updcde") = ""
                .Fields("sysdte") = gzGetSysDate
                .Update
                .Close
            End With
        Case SPL
            With rstSPL
                .Open "SPLAlloc", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
                .AddNew
                .Fields("teller") = txtTeller(SPL)
                .Fields("strccr") = txtStartCCR(SPL)
                .Fields("endccr") = txtEndCCR(SPL)
                .Fields("prvccr") = 0
                .Fields("wrkstn") = ""
                .Fields("userid") = lblAllocBy(SPL).Caption
                .Fields("status") = ""
                .Fields("updcde") = ""
                .Fields("sysdte") = gzGetSysDate
                .Update
                .Close
            End With
    End Select
    Call FillGrid
    
End Sub

Private Sub EditAllocation()
    Dim rstCCR As New ADODB.Recordset
    Dim rstCFM As New ADODB.Recordset
    Dim rstCYM As New ADODB.Recordset
    Dim rstCYE As New ADODB.Recordset
    Dim rstSPL As New ADODB.Recordset
    
    Select Case Index
        Case CCR
            With rstCCR
                .Open "Select * from CCRAlloc where teller = '" & Trim(txtTeller(Index)) & "'", _
                    gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
                .Fields("strccr") = txtStartCCR(CCR)
                .Fields("endccr") = txtEndCCR(CCR)
                .Fields("userid") = lblAllocBy(CCR).Caption
                .Fields("sysdte") = gzGetSysDate
                .Fields("prvccr") = 0
                .Update
                .Close
            End With
        Case CFM
            With rstCFM
                .Open "Select * from CFMAlloc where teller ='" & Trim(txtTeller(Index)) & "'", _
                    gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
                .Fields("strgps") = txtStartCCR(CFM)
                .Fields("endgps") = txtEndCCR(CFM)
                .Fields("userid") = lblAllocBy(CFM).Caption
                .Fields("sysdte") = gzGetSysDate
                .Fields("prvgps") = 0
                .Update
                .Close
            End With
        Case CYM
            With rstCYM
                .Open "Select * from CYMAlloc where teller = '" & Trim(txtTeller(Index)) & "'", _
                    gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
                .Fields("strgps") = txtStartCCR(CYM)
                .Fields("endgps") = txtEndCCR(CYM)
                .Fields("userid") = lblAllocBy(CYM).Caption
                .Fields("sysdte") = gzGetSysDate
                .Fields("prvgps") = 0
                .Update
                .Close
            End With
        Case CYE
            With rstCYE
                .Open "Select * from CYEAlloc where teller = '" & Trim(txtTeller(Index)) & "'", _
                    gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
                .Fields("strgps") = txtStartCCR(CYE)
                .Fields("endgps") = txtEndCCR(CYE)
                .Fields("userid") = lblAllocBy(CYE).Caption
                .Fields("sysdte") = gzGetSysDate
                .Fields("prvgps") = 0
                .Update
                .Close
            End With
        Case SPL
            With rstSPL
                .Open "Select * from SPLAlloc where teller = '" & Trim(txtTeller(Index)) & "'", _
                    gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
                .Fields("strccr") = txtStartCCR(SPL)
                .Fields("endccr") = txtEndCCR(SPL)
                .Fields("userid") = lblAllocBy(SPL).Caption
                .Fields("sysdte") = gzGetSysDate
                .Fields("prvccr") = 0
                .Update
                .Close
            End With
        End Select
    Call FillGrid
    
End Sub

Private Sub DeleteAllocation()
    Dim rstCCR As New ADODB.Recordset
    Dim rstCFM As New ADODB.Recordset
    Dim rstCYM As New ADODB.Recordset
    Dim rstCYE As New ADODB.Recordset
    Dim rstSPL As New ADODB.Recordset
    Dim noRec As Boolean
    
    noRec = False
    Select Case Index
        Case CCR
            With rstCCR
                .Open "Select * from CCRAlloc where teller = '" & Trim(txtTeller(Index)) & "'", _
                    gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
                .Delete
                .MoveNext
                If .EOF Then
                    noRec = True
                End If
                .Close
            End With
        Case CFM
            With rstCFM
                .Open "Select * from CFMAlloc where teller = '" & Trim(txtTeller(Index)) & "'", _
                    gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
                .Delete
                .MoveNext
                If .EOF Then
                    noRec = True
                End If
                .Close
            End With
        Case CYM
            With rstCYM
                .Open "Select * from CYMAlloc where teller = '" & Trim(txtTeller(Index)) & "'", _
                    gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
                .Delete
                .MoveNext
                If .EOF Then
                    noRec = True
                End If
                .Close
            End With
        Case CYE
            With rstCYE
                .Open "Select * from CYEAlloc where teller = '" & Trim(txtTeller(Index)) & "'", _
                    gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
                .Delete
                .MoveNext
                If .EOF Then
                    noRec = True
                End If
                .Close
            End With
        Case SPL
            With rstSPL
                .Open "Select * from SPLAlloc where teller = '" & Trim(txtTeller(Index)) & "'", _
                    gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
                .Delete
                .MoveNext
                If .EOF Then
                    noRec = True
                End If
                .Close
            End With
        End Select
    If noRec = False Then
        Call FillGrid
    Else                    ' if noRec = True, table is already empty
        grdAlloc(Index).Clear
        grdAlloc(Index).Rows = 2
        grdAlloc(Index).Refresh
        Call SetGrid
    End If
End Sub
' Returns true if allocation does not overlap any CCR from the grid or from detail tables
Private Function ValidAllocation() As Boolean
    Dim TempSccr As Long    ' temp for txtStartCCR(index)
    Dim TempEccr As Long    ' temp for txtEndCCR(index)
    Dim TempTlr As String * 10 ' temp for txtTeller(index)
    Dim TempStart As Long   ' temp for grid's StartCCR
    Dim TempEnd As Long     ' temp for grid's EndCCR
    
    ValidAllocation = True
    VASwch = True  ' Set to true to bypass grdAlloc_RowColChange
    TempSccr = CLng(txtStartCCR(Index))
    TempEccr = CLng(txtEndCCR(Index))
    TempTlr = Trim(txtTeller(Index))
            
    With grdAlloc(Index)
        .Row = 0
        Do Until .Row = .Rows - 1
            .Row = .Row + 1: .Col = 0
            If Trim(.Text) = "" Then   ' exit loop since no data from the grid exists
                Exit Do
            End If
            If EditSw = True And Trim(TempTlr) = Trim(.Text) Then
                EditSw = True
            Else
                If Trim(TempTlr) = Trim(.Text) Then
                    MsgBox "Teller has existing allocation.", vbExclamation, "Allocation Error"
                    ValidAllocation = False
                    Exit Do
                End If
                .Col = 1: TempStart = CLng(.Text)
                .Col = 2: TempEnd = CLng(.Text)
                If (TempSccr >= TempStart And TempSccr <= TempEnd) Or _
                   (TempEccr >= TempStart And TempEccr <= TempEnd) Or _
                   (TempSccr < TempStart And TempEccr > TempEnd) Then
                    MsgBox "The allocation overlaps another allocation.", vbExclamation, "Allocation Error"
                    ValidAllocation = False
                    Exit Do
                End If
            End If
        Loop
    End With
    
    If ValidAllocation = True Then
        Select Case Index
            Case CCR
                If gzChkCCRExists(CLng(txtStartCCR(Index)), CLng(txtEndCCR(Index))) Then
                    MsgBox "The allocation overlaps a CCR issued.", vbExclamation, "Allocation Error"
                    ValidAllocation = False
                End If
            Case CFM
                If gzChkGPSExists(CLng(txtStartCCR(Index)), CLng(txtEndCCR(Index)), 1) Then
                    MsgBox "The allocation overlaps a gatepass issued.", vbExclamation, "Allocation Error"
                    ValidAllocation = False
                End If
            Case CYM
                If gzChkGPSExists(CLng(txtStartCCR(Index)), CLng(txtEndCCR(Index)), 2) Then
                    MsgBox "The allocation overlaps a gatepass issued.", vbExclamation, "Allocation Error"
                    ValidAllocation = False
                End If
            Case CYE
                If gzChkGPSExists(CLng(txtStartCCR(Index)), CLng(txtEndCCR(Index)), 3) Then
                    MsgBox "The allocation overlaps a gatepass issued.", vbExclamation, "Allocation Error"
                    ValidAllocation = False
                End If
            Case SPL
                If gzChkCCRExists(CLng(txtStartCCR(Index)), CLng(txtEndCCR(Index))) Then
                    MsgBox "The allocation overlaps a CCR issued.", vbExclamation, "Allocation Error"
                    ValidAllocation = False
                End If
        End Select
    End If

    VASwch = False
End Function
