VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCYXINQ02 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "( SUBIC Version CYXINQ02 ) CY Export Teller Collection Inquiry"
   ClientHeight    =   10800
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "IBM3270 - 1254"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCYXINQ02.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   10425
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12674
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3757
            Picture         =   "frmCYXINQ02.frx":08CA
            Text            =   "CYXINQ02"
            TextSave        =   "CYXINQ02"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   2619
            TextSave        =   "7/11/00"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   2619
            TextSave        =   "8:42 AM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "F3 - E&xit"
      Height          =   780
      Left            =   10800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   9480
      Width           =   4335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   10575
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   15135
      Begin VB.CommandButton cmdComputation 
         Caption         =   "F9 - Computation Details"
         Height          =   780
         Left            =   10680
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   7800
         Width           =   4335
      End
      Begin VB.Frame Frame5 
         Height          =   1455
         Left            =   10680
         TabIndex        =   83
         Top             =   6000
         Width           =   4335
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   240
            TabIndex        =   85
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackColor       =   &H00808000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CCR STATUS"
            Height          =   375
            Left            =   0
            TabIndex        =   84
            Top             =   0
            Width           =   4335
         End
      End
      Begin VB.CommandButton cmdTab 
         Caption         =   "F11 - Container #"
         Height          =   780
         Left            =   10680
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   8640
         Width           =   4335
      End
      Begin TabDlg.SSTab ST 
         Height          =   5295
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   600
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   9340
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   2
         ShowFocusRect   =   0   'False
         BackColor       =   12632256
         TabCaption(0)   =   "F6 - Cons. Summary"
         TabPicture(0)   =   "frmCYXINQ02.frx":11A4
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "ccrNo"
         Tab(0).Control(1)=   "Frame4"
         Tab(0).Control(2)=   "Frame3"
         Tab(0).Control(3)=   "Label19"
         Tab(0).Control(4)=   "lblRefnum"
         Tab(0).Control(5)=   "lblBrker"
         Tab(0).Control(6)=   "Label16"
         Tab(0).Control(7)=   "lblRemarks"
         Tab(0).Control(8)=   "lblCommodity"
         Tab(0).Control(9)=   "lblBroker"
         Tab(0).Control(10)=   "Label15"
         Tab(0).Control(11)=   "Label14"
         Tab(0).Control(12)=   "Label13"
         Tab(0).Control(13)=   "lblEntry(9)"
         Tab(0).Control(14)=   "lblEntry(8)"
         Tab(0).Control(15)=   "lblEntry(7)"
         Tab(0).Control(16)=   "lblEntry(6)"
         Tab(0).Control(17)=   "lblEntry(5)"
         Tab(0).Control(18)=   "lblEntry(4)"
         Tab(0).Control(19)=   "lblEntry(3)"
         Tab(0).Control(20)=   "lblEntry(2)"
         Tab(0).Control(21)=   "lblEntry(1)"
         Tab(0).Control(22)=   "lblEntry(0)"
         Tab(0).Control(23)=   "lblExporter"
         Tab(0).Control(24)=   "Label12"
         Tab(0).Control(25)=   "Label11"
         Tab(0).Control(26)=   "Label10"
         Tab(0).ControlCount=   27
         TabCaption(1)   =   "F7 - Teller Summary"
         TabPicture(1)   =   "frmCYXINQ02.frx":11C0
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label4"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "gFlx"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Number"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Prefix"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         Begin CYXINQ02.pNumeric ccrNo 
            Height          =   420
            Left            =   -72120
            TabIndex        =   0
            Top             =   480
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   741
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CYXINQ02.pText Prefix 
            Height          =   420
            Left            =   2640
            TabIndex        =   1
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   741
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin CYXINQ02.pNumeric Number 
            Height          =   420
            Left            =   3600
            TabIndex        =   2
            Top             =   120
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   741
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Frame Frame4 
            Height          =   135
            Left            =   -74640
            TabIndex        =   80
            Top             =   3480
            Width           =   14295
         End
         Begin VB.Frame Frame3 
            Height          =   135
            Left            =   -74640
            TabIndex        =   79
            Top             =   2280
            Width           =   14295
         End
         Begin MSFlexGridLib.MSFlexGrid gFlx 
            Height          =   3975
            Left            =   120
            TabIndex        =   3
            Top             =   1200
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   7011
            _Version        =   393216
            Cols            =   38
            FixedCols       =   0
            FocusRect       =   2
            SelectionMode   =   1
            FormatString    =   $"frmCYXINQ02.frx":11DC
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Reference Number"
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   -74640
            TabIndex        =   88
            Top             =   1320
            Width           =   3495
         End
         Begin VB.Label lblRefnum 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Left            =   -71040
            TabIndex        =   61
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label lblBrker 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Left            =   -72120
            TabIndex        =   86
            Top             =   3720
            Width           =   7575
         End
         Begin VB.Label Label16 
            Caption         =   "<Enter> - Retrieve Records"
            BeginProperty Font 
               Name            =   "IBM3270 - 1254"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -69000
            TabIndex        =   81
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label lblRemarks 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Left            =   -72120
            TabIndex        =   78
            Top             =   4680
            Width           =   7575
         End
         Begin VB.Label lblCommodity 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Left            =   -72120
            TabIndex        =   77
            Top             =   4200
            Width           =   7575
         End
         Begin VB.Label lblBroker 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Left            =   -72000
            TabIndex        =   76
            Top             =   6120
            Width           =   7575
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Remarks"
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   -74640
            TabIndex        =   75
            Top             =   4680
            Width           =   2415
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Commodity"
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   -74640
            TabIndex        =   74
            Top             =   4200
            Width           =   2415
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Broker"
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   -74640
            TabIndex        =   73
            Top             =   3720
            Width           =   2415
         End
         Begin VB.Label lblEntry 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Index           =   9
            Left            =   -65880
            TabIndex        =   72
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label lblEntry 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Index           =   8
            Left            =   -67440
            TabIndex        =   71
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label lblEntry 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Index           =   7
            Left            =   -69000
            TabIndex        =   70
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label lblEntry 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Index           =   6
            Left            =   -70560
            TabIndex        =   69
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label lblEntry 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Index           =   5
            Left            =   -72120
            TabIndex        =   68
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label lblEntry 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Index           =   4
            Left            =   -65880
            TabIndex        =   67
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblEntry 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Index           =   3
            Left            =   -67440
            TabIndex        =   66
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblEntry 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Index           =   2
            Left            =   -69000
            TabIndex        =   65
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblEntry 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Index           =   1
            Left            =   -70560
            TabIndex        =   64
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblEntry 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Index           =   0
            Left            =   -72120
            TabIndex        =   63
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label lblExporter 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   420
            Left            =   -71040
            TabIndex        =   62
            Top             =   1800
            Width           =   7335
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Entry Numbers"
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   -74640
            TabIndex        =   60
            Top             =   2520
            Width           =   2415
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Exporter"
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   -74640
            TabIndex        =   59
            Top             =   1800
            Width           =   3495
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H00800080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CCR Number"
            BeginProperty Font 
               Name            =   "IBM3270 - 1254"
               Size            =   16.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   420
            Left            =   -74640
            TabIndex        =   58
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CONTAINER NO."
            Height          =   420
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   2415
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " (1) - View CCR Computation Details"
            BeginProperty Font 
               Name            =   "IBM3270 - 1254"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   11535
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Height          =   4455
         Left            =   120
         TabIndex        =   8
         Top             =   5880
         Width           =   10455
         Begin VB.Label dtetme 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   7080
            TabIndex        =   57
            Top             =   3600
            Width           =   3135
         End
         Begin VB.Label teller 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   4920
            TabIndex        =   56
            Top             =   3600
            Width           =   2175
         End
         Begin VB.Label fe 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   4200
            TabIndex        =   55
            Top             =   3600
            Width           =   735
         End
         Begin VB.Label size 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   3240
            TabIndex        =   54
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label cntnum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   53
            Top             =   3600
            Width           =   3135
         End
         Begin VB.Label dtetme 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   7080
            TabIndex        =   52
            Top             =   3240
            Width           =   3135
         End
         Begin VB.Label teller 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   4920
            TabIndex        =   51
            Top             =   3240
            Width           =   2175
         End
         Begin VB.Label fe 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   4200
            TabIndex        =   50
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label size 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   3240
            TabIndex        =   49
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label cntnum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   48
            Top             =   3240
            Width           =   3135
         End
         Begin VB.Label dtetme 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   7080
            TabIndex        =   47
            Top             =   2880
            Width           =   3135
         End
         Begin VB.Label teller 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   4920
            TabIndex        =   46
            Top             =   2880
            Width           =   2175
         End
         Begin VB.Label fe 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   4200
            TabIndex        =   45
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label size 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   3240
            TabIndex        =   44
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label cntnum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   43
            Top             =   2880
            Width           =   3135
         End
         Begin VB.Label dtetme 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   7080
            TabIndex        =   42
            Top             =   2520
            Width           =   3135
         End
         Begin VB.Label teller 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   4920
            TabIndex        =   41
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label fe 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   4200
            TabIndex        =   40
            Top             =   2520
            Width           =   735
         End
         Begin VB.Label size 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   3240
            TabIndex        =   39
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label cntnum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   38
            Top             =   2520
            Width           =   3135
         End
         Begin VB.Label dtetme 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   7080
            TabIndex        =   37
            Top             =   2160
            Width           =   3135
         End
         Begin VB.Label teller 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   4920
            TabIndex        =   36
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label fe 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   4200
            TabIndex        =   35
            Top             =   2160
            Width           =   735
         End
         Begin VB.Label size 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   3240
            TabIndex        =   34
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label cntnum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   33
            Top             =   2160
            Width           =   3135
         End
         Begin VB.Label dtetme 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   32
            Top             =   1800
            Width           =   3135
         End
         Begin VB.Label teller 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   4920
            TabIndex        =   31
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label fe 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   4200
            TabIndex        =   30
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label size 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   3240
            TabIndex        =   29
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label cntnum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   1800
            Width           =   3135
         End
         Begin VB.Label dtetme 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   7080
            TabIndex        =   27
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label teller 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   4920
            TabIndex        =   26
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label fe 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   4200
            TabIndex        =   25
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label size 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   3240
            TabIndex        =   24
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label cntnum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label dtetme 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   7080
            TabIndex        =   22
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label teller 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   4920
            TabIndex        =   21
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label fe 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   4200
            TabIndex        =   20
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label size 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   3240
            TabIndex        =   19
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label cntnum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Date && Time"
            Height          =   375
            Left            =   7080
            TabIndex        =   17
            Top             =   600
            Width           =   3135
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Teller"
            Height          =   375
            Left            =   4920
            TabIndex        =   16
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "F/E"
            Height          =   375
            Left            =   4200
            TabIndex        =   15
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Size"
            Height          =   375
            Left            =   3240
            TabIndex        =   14
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Container Number"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   3135
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00808000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CCR PARTICULARS"
            BeginProperty Font 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   9
            Top             =   120
            Width           =   10455
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PAYMENT TRANSACTION PARTICULARS INQUIRY"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   15135
      End
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu FileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmCYXINQ02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub GetInformation()
    Dim rs As Recordset
    DE.GetInformation
    Set rs = DE.rsgetInformation
    If rs.RecordCount > 0 Then
        With rs
            SB.Panels(1).Text = .Fields("workstation")
            SB.Panels(2).Text = gUserid
        End With
    End If
    rs.Close
    Set rs = Nothing
End Sub
Private Sub ccrNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call FillCCR
    End If
End Sub
Private Sub cmdComputation_Click()
    MsgBox "Not Yet Installed"
'    frmComputations.Show vbModal
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdTab_Click()
    If ST.Tab = 0 Then
        ST.Tab = 1
        cmdTab.Caption = "F11 - CCR No."
        Call NewContainer
    Else
        ST.Tab = 0
        cmdTab.Caption = "F11 - Container #"
        Call NewCCR
    End If
End Sub
Private Sub FileExit_Click()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            If cmdExit.Enabled Then
                Call cmdExit_Click
            End If
        Case vbKeyF9
            If cmdComputation.Enabled Then
                Call cmdComputation_Click
            End If
        Case vbKeyF11
            If cmdTab.Enabled Then
                Call cmdTab_Click
            End If
    End Select
End Sub
Private Sub Form_Load()
    Call NewContainer
End Sub
Private Sub gFlx_EnterCell()
    Call MoveFromGridToDetails
End Sub
Private Sub gFlx_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        If Chr(KeyAscii) = "1" Then
            gFlx.TextMatrix(gFlx.Row, 0) = Chr(KeyAscii)
        Else
            Beep
            KeyAscii = 0
        End If
    Else
        gFlx.TextMatrix(gFlx.Row, 0) = ""
    End If
End Sub
Private Sub Number_Change()
    If FlxRef > 0 Then
        Call InitializeGrid
        Call ClearParticulars
        Number.SetFocus
    End If
End Sub
Private Sub Number_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(Trim(Prefix.Text)) > 0 And Len(Trim(Number.Value)) > 0 Then
            If FillGrid Then
                Call MoveFromGridToDetails
                gFlx.SetFocus
            Else
                Number.SetFocus
            End If
        End If
    End If
End Sub
Private Sub Prefix_Change()
    If Len(Trim(Prefix.Text)) = 4 Then
        Number.SetFocus
    Else
        If FlxRef > 0 Then
            Call InitializeGrid
            Call ClearParticulars
            Prefix.SetFocus
        End If
    End If
End Sub
Public Function FillGrid() As Boolean
    
    Dim X As Integer
    Dim ContainerNumber As String
    Dim cntPref As String * 4
    Dim cntnum As String * 8
    
    Dim rs As Recordset
    Dim rs2 As Recordset
    
    cntPref = ""
    cntnum = ""
    cntPref = Mid(UCase(Prefix.Text), 1, 4)
    cntnum = Mid(UCase(Number.Value), 1, 8)
    ContainerNumber = cntPref & cntnum
    DE.Container ContainerNumber
    Set rs = DE.rsContainer
        
    If rs.RecordCount > 0 Then
        With rs
            Call InitializeGrid
            Do While Not .EOF
                 FlxRef = FlxRef + 1
                If FlxRef > 1 Then
                    gFlx.AddItem " "
                End If
                gFlx.TextMatrix(FlxRef, 1) = .Fields("exprtr")
                gFlx.TextMatrix(FlxRef, 2) = .Fields("refnum")
                gFlx.TextMatrix(FlxRef, 3) = .Fields("ccrnum")
                gFlx.TextMatrix(FlxRef, 4) = .Fields("sysdttm")
                If .Fields("status") = "CAN" Then
                    gFlx.TextMatrix(FlxRef, 37) = "CANCELLED"
                Else
                    If .Fields("updcde") = "R" Then
                        gFlx.TextMatrix(FlxRef, 37) = "REFUNDED"
                    Else
                        gFlx.TextMatrix(FlxRef, 37) = ""
                    End If
                End If
                DE.CCR .Fields("ccrnum")
                Set rs2 = DE.rsCCR
                If rs2.RecordCount > 0 Then
                    X = 5
                    Do While Not rs2.EOF
                        If X < 36 Then
                            gFlx.TextMatrix(FlxRef, X) = rs2.Fields("cntnum")
                            gFlx.TextMatrix(FlxRef, X + 1) = rs2.Fields("cntsze")
                            gFlx.TextMatrix(FlxRef, X + 2) = rs2.Fields("fulemp")
                            gFlx.TextMatrix(FlxRef, X + 3) = rs2.Fields("userid")
                        End If
                        X = X + 4
                        rs2.MoveNext
                    Loop
                End If
                rs2.Close
                Set rs2 = Nothing
                .MoveNext
            Loop
        End With
        rs.Close
        Set rs = Nothing
        FillGrid = True
    Else
        Beep
        MsgBox "No Records Found !", vbInformation + vbOKOnly, "Search Result"
        FillGrid = False
        rs.Close
        Set rs = Nothing
    End If
End Function
Public Sub MoveFromGridToDetails()
    Dim X As Integer
    With gFlx
        For X = 0 To 7
            cntnum(X) = .TextMatrix(.Row, (4 * (X + 1)) + 1)
            size(X) = .TextMatrix(.Row, (4 * (X + 1)) + 2)
            fe(X) = .TextMatrix(.Row, (4 * (X + 1)) + 3)
            teller(X) = .TextMatrix(.Row, (4 * (X + 1)) + 4)
            If Len(Trim(cntnum(X))) > 0 Then
                dtetme(X) = .TextMatrix(.Row, 4)
            End If
        Next
        lblStatus.Caption = Trim(UCase(.TextMatrix(.Row, 37)))
    End With
End Sub
Private Sub Prefix_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(Trim(Prefix.Text)) > 0 And Len(Trim(Number.Value)) > 0 Then
            If FillGrid Then
                Call MoveFromGridToDetails
                gFlx.SetFocus
            Else
                Prefix.SetFocus
            End If
        End If
    End If
End Sub
Private Sub Prefix_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Public Sub InitializeGrid()
    Dim X As Integer
    With gFlx
        .Clear
        .Rows = 2
        .FormatString = "Opt | Exporter | Reference | CCR Number | Date & Time |5|6|7|8|9|10|11|12|13|14|15|16|17|18|19|20|21|22|23|24|25|26|27|28|29|30|31|32|33|34|35|36|37"
        .Cols = 38
        .ColWidth(0) = 700
        .ColWidth(1) = 3500
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 3000
        For X = 5 To 37
            .ColWidth(X) = 1
        Next
    End With
    FlxRef = 0
End Sub
Public Sub Tab0(D As Boolean)
    ccrNo.TabStop = D
    ccrNo.Enabled = D
End Sub
Public Sub Tab1(D As Boolean)
    Prefix.TabStop = D
    Prefix.Enabled = D
    Number.TabStop = D
    Number.Enabled = D
    gFlx.TabStop = D
    gFlx.Enabled = D
End Sub
Public Sub NewContainer()
    Call InitializeGrid
    Prefix.Text = ""
    Number.Value = ""
    Call ClearParticulars
    ST.Tab = 1
    Tab0 (False)
    Tab1 (True)
    cmdComputation.Enabled = False
    cmdComputation.TabStop = False
    If Prefix.Visible Then
        Prefix.SetFocus
    End If
End Sub
Public Sub NewCCR()
    Dim X As Integer
    lblExporter.Caption = ""
    For X = 0 To 9
        lblEntry(X).Caption = ""
    Next
    lblRefnum.Caption = ""
    lblBroker.Caption = ""
    lblCommodity.Caption = ""
    lblRemarks.Caption = ""
    Tab0 (True)
    Tab1 (False)
    cmdComputation.Enabled = True
    cmdComputation.TabStop = True
    Call ClearParticulars
    ccrNo.SetFocus
End Sub
Public Sub ClearParticulars()
    Dim X As Integer
    For X = 0 To 7
        cntnum(X) = ""
        size(X) = ""
        fe(X) = ""
        teller(X) = ""
        dtetme(X) = ""
    Next
End Sub
Public Sub FillCCR()
    Dim X As Integer
    Dim rs As Recordset
    DE.CCR Trim(ccrNo.Value)
    Set rs = DE.rsCCR
    If rs.RecordCount > 0 Then
        With rs
            lblExporter.Caption = .Fields("exprtr")
            lblBrker.Caption = .Fields("broker")
            lblCommodity.Caption = .Fields("commod")
            lblRemarks.Caption = .Fields("remark")
            lblRefnum.Caption = Format(.Fields("refnum"), "00000000")
            lblEntry(0).Caption = Mid(.Fields("entnum"), 1, 8)
            lblEntry(1).Caption = Mid(.Fields("entnum"), 9, 8)
            lblEntry(2).Caption = Mid(.Fields("entnum"), 17, 8)
            lblEntry(3).Caption = Mid(.Fields("entnum"), 25, 8)
            lblEntry(4).Caption = Mid(.Fields("entnum"), 33, 8)
            lblEntry(5).Caption = Mid(.Fields("entnum"), 41, 8)
            lblEntry(6).Caption = Mid(.Fields("entnum"), 49, 8)
            lblEntry(7).Caption = Mid(.Fields("entnum"), 57, 8)
            lblEntry(8).Caption = Mid(.Fields("entnum"), 65, 8)
            X = 0
            Do While Not .EOF
                If X < 8 Then
                    cntnum(X).Caption = .Fields("cntnum")
                    size(X).Caption = .Fields("cntsze")
                    fe(X).Caption = .Fields("fulemp")
                    teller(X).Caption = .Fields("userid")
                    dtetme(X).Caption = Format(.Fields("sysdttm"), "YYYY-MM-DD hh:nn:ss")
                End If
                .MoveNext
                X = X + 1
            Loop
        End With
    Else
        Beep
        MsgBox "No Records Found !", vbInformation + vbOKOnly, "Search Result"
    End If
    rs.Close
    Set rs = Nothing
End Sub
