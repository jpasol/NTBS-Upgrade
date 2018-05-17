VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCYXINQ01 
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
   Icon            =   "frmCYXINQ01.frx":0000
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
   Begin VB.CommandButton cmdTab 
      Caption         =   "F11 - Teller Inquiry"
      Height          =   495
      Left            =   360
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   4080
      Width           =   3855
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "F5 - Refresh"
      Height          =   780
      Left            =   12000
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8640
      Width           =   3135
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
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
            Picture         =   "frmCYXINQ01.frx":08CA
            Text            =   "CYXINQ01"
            TextSave        =   "CYXINQ01"
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
      Left            =   12000
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9480
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Height          =   10575
      Left            =   120
      TabIndex        =   8
      Top             =   -120
      Width           =   15135
      Begin TabDlg.SSTab ST 
         Height          =   4215
         Left            =   120
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   600
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   7435
         _Version        =   393216
         TabOrientation  =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   2
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "F6 - Cons. Summary"
         TabPicture(0)   =   "frmCYXINQ01.frx":11A4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label14"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblTotalCollection"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label15"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label16"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "ThisDay"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label18"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label17"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lblTotalClearedCollection"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label20"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "cTime(1)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "flxM"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "cTime(0)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).ControlCount=   13
         TabCaption(1)   =   "F7 - Teller Summary"
         TabPicture(1)   =   "frmCYXINQ01.frx":11C0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label27"
         Tab(1).Control(1)=   "Label28"
         Tab(1).Control(2)=   "Label29"
         Tab(1).Control(3)=   "Label30"
         Tab(1).Control(4)=   "iTime(1)"
         Tab(1).Control(5)=   "Frame6"
         Tab(1).Control(6)=   "Frame7"
         Tab(1).Control(7)=   "txtDate"
         Tab(1).Control(8)=   "iTime(0)"
         Tab(1).Control(9)=   "cmdRetrieve"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "txtTeller"
         Tab(1).ControlCount=   11
         Begin CYXINQ02.prvusrctrlTime cTime 
            Height          =   420
            Index           =   0
            Left            =   11760
            TabIndex        =   51
            Top             =   1680
            Width           =   2895
            _ExtentX        =   5953
            _ExtentY        =   741
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   8
         End
         Begin VB.TextBox txtTeller 
            BackColor       =   &H8000000F&
            Height          =   420
            Left            =   -70560
            TabIndex        =   2
            Top             =   1260
            Width           =   4335
         End
         Begin VB.CommandButton cmdRetrieve 
            Caption         =   "F12 - Retrieve Collection"
            Height          =   495
            Left            =   -69960
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   3600
            Width           =   6615
         End
         Begin MSFlexGridLib.MSFlexGrid flxM 
            Height          =   2895
            Left            =   120
            TabIndex        =   0
            Top             =   600
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   5106
            _Version        =   393216
            Cols            =   22
            FixedCols       =   0
            BackColorFixed  =   12113909
            FocusRect       =   2
            SelectionMode   =   1
            FormatString    =   "Opt | Teller | Cash | ADR | Total |5|6|7|8|9|10|11|12|13|14|15|16|17| Status"
         End
         Begin CYXINQ02.prvusrctrlTime iTime 
            Height          =   420
            Index           =   0
            Left            =   -70560
            TabIndex        =   3
            Top             =   1740
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   8
         End
         Begin CYXINQ02.prvusrctrlDate txtDate 
            Height          =   420
            Left            =   -70560
            TabIndex        =   1
            Top             =   780
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   741
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
         Begin VB.Frame Frame7 
            Height          =   135
            Left            =   -74880
            TabIndex        =   44
            Top             =   3360
            Width           =   14655
         End
         Begin VB.Frame Frame6 
            Height          =   135
            Left            =   -74880
            TabIndex        =   40
            Top             =   60
            Width           =   14655
         End
         Begin CYXINQ02.prvusrctrlTime iTime 
            Height          =   420
            Index           =   1
            Left            =   -68040
            TabIndex        =   4
            Top             =   1740
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   741
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   8
         End
         Begin CYXINQ02.prvusrctrlTime cTime 
            Height          =   420
            Index           =   1
            Left            =   11760
            TabIndex        =   52
            Top             =   2640
            Width           =   2895
            _ExtentX        =   5953
            _ExtentY        =   741
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   8
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cleared"
            BeginProperty Font 
               Name            =   "IBM3270 - 1254"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   9120
            TabIndex        =   57
            Top             =   3600
            Width           =   1815
         End
         Begin VB.Label lblTotalClearedCollection 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "IBM3270 - 1254"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   10920
            TabIndex        =   56
            Top             =   3600
            Width           =   3855
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Use <Tab> to Pan through input fields"
            BeginProperty Font 
               Name            =   "IBM3270 - 1254"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   6000
            TabIndex        =   55
            Top             =   120
            Width           =   5655
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackColor       =   &H00808000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "D A T E"
            Height          =   375
            Left            =   11760
            TabIndex        =   54
            Top             =   120
            Width           =   2895
         End
         Begin VB.Label ThisDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   11760
            TabIndex        =   53
            Top             =   555
            Width           =   2895
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "From Time"
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   11760
            TabIndex        =   50
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "To Time"
            ForeColor       =   &H80000008&
            Height          =   420
            Left            =   11760
            TabIndex        =   49
            Top             =   2280
            Width           =   2895
         End
         Begin VB.Label lblTotalCollection 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "IBM3270 - 1254"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   5400
            TabIndex        =   46
            Top             =   3600
            Width           =   3615
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "IBM3270 - 1254"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   4080
            TabIndex        =   45
            Top             =   3600
            Width           =   1335
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "To"
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   -68640
            TabIndex        =   43
            Top             =   1740
            Width           =   495
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Time Range"
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   -74040
            TabIndex        =   42
            Top             =   1740
            Width           =   3375
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Teller"
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   -74040
            TabIndex        =   41
            Top             =   1260
            Width           =   3375
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Date To Process"
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   -74040
            TabIndex        =   39
            Top             =   780
            Width           =   3375
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   " (1) - Notify  (2) - Clear "
            BeginProperty Font 
               Name            =   "IBM3270 - 1254"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   5895
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5655
         Left            =   120
         TabIndex        =   10
         Top             =   4800
         Width           =   11655
         Begin VB.Frame Frame4 
            Height          =   135
            Left            =   120
            TabIndex        =   25
            Top             =   4800
            Width           =   7215
         End
         Begin VB.Frame Frame3 
            Height          =   135
            Left            =   120
            TabIndex        =   24
            Top             =   2880
            Width           =   11415
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "P O S "
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   4200
            Width           =   3015
         End
         Begin VB.Label lblPOS 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3240
            TabIndex        =   60
            Top             =   4200
            Width           =   4095
         End
         Begin VB.Label lblAdr 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3240
            TabIndex        =   59
            Top             =   3840
            Width           =   4095
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A D R "
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   3840
            Width           =   3015
         End
         Begin VB.Label lblSubTotalUG 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   7560
            TabIndex        =   36
            Top             =   2520
            Width           =   4095
         End
         Begin VB.Label lblSubTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3360
            TabIndex        =   35
            Top             =   2520
            Width           =   4095
         End
         Begin VB.Label lblWtxUG 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   7560
            TabIndex        =   34
            Top             =   2160
            Width           =   4095
         End
         Begin VB.Label lblWtx 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3360
            TabIndex        =   33
            Top             =   2160
            Width           =   4095
         End
         Begin VB.Label lblNVatUG 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   7560
            TabIndex        =   32
            Top             =   1800
            Width           =   4095
         End
         Begin VB.Label lblNVat 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3360
            TabIndex        =   31
            Top             =   1800
            Width           =   4095
         End
         Begin VB.Label lblVatUG 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   7560
            TabIndex        =   30
            Top             =   1440
            Width           =   4095
         End
         Begin VB.Label lblVat 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3360
            TabIndex        =   29
            Top             =   1440
            Width           =   4095
         End
         Begin VB.Label lblTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3240
            TabIndex        =   28
            Top             =   5160
            Width           =   4095
         End
         Begin VB.Label lblAmountDue 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3240
            TabIndex        =   27
            Top             =   3480
            Width           =   4095
         End
         Begin VB.Label lblWhf 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3240
            TabIndex        =   26
            Top             =   3120
            Width           =   4095
         End
         Begin VB.Label lblArrUG 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   7560
            TabIndex        =   23
            Top             =   1080
            Width           =   4095
         End
         Begin VB.Label lblArr 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3360
            TabIndex        =   22
            Top             =   1080
            Width           =   4095
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total "
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   5160
            Width           =   3015
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Amount Due "
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   3480
            Width           =   3015
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Wharfage "
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   3120
            Width           =   3015
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sub - Total "
            Height          =   375
            Left            =   240
            TabIndex        =   18
            Top             =   2520
            Width           =   3015
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Witholding Tax "
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   2160
            Width           =   3015
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "No Vat Amount "
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   1800
            Width           =   3015
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vat Charges "
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Amount "
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   1080
            Width           =   3015
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A R R A S T R E  U G"
            Height          =   375
            Left            =   7560
            TabIndex        =   13
            Top             =   600
            Width           =   4095
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "A R R A S T R E"
            Height          =   375
            Left            =   3360
            TabIndex        =   12
            Top             =   600
            Width           =   4095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00808000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Collection Breakdown"
            Height          =   375
            Left            =   0
            TabIndex        =   11
            Top             =   120
            Width           =   11655
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T E L L E R   C O L L E C T I O N   I N Q U I R Y"
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
         TabIndex        =   9
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
Attribute VB_Name = "frmCYXINQ01"
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

    DE.GetUserType gUserid
    Set rs = DE.rsGetUserType
    If rs.RecordCount > 0 Then
        UserType = UCase(Trim(rs.Fields("offcde")))
    Else
        UserType = "T"
    End If
    rs.Close
    Set rs = Nothing

End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdRefresh_Click()
    Call RefreshGrid
End Sub
Private Sub cmdRetrieve_Click()
    Dim TDay As Date
    Dim FromDte As String
    Dim ToDte As String
    TDay = CDate(txtDate.Text)
    FromDte = Year(TDay) & "-" & Month(TDay) & "-" & Day(TDay) & " " & Trim(iTime(0).Text)
    ToDte = Year(TDay) & "-" & Month(TDay) & "-" & Day(TDay) & " " & Trim(iTime(1).Text)
    Call FillBreakdown(FromDte, ToDte, UCase(Trim(txtTeller.Text)), 1, 2)
End Sub
Private Sub cmdTab_Click()
    Select Case ST.Tab
        Case 0
            cmdTab.Width = 4815
            cmdTab.Caption = "F11 - Cons. Teller Inquiry"
            cmdRefresh.Enabled = False
            cmdRefresh.TabStop = False
            Call RefreshTeller
        Case 1
            cmdTab.Width = 3855
            cmdTab.Caption = "F11 - Teller Inquiry"
            cmdRefresh.Enabled = True
            cmdRefresh.TabStop = True
            Call RefreshGrid
    End Select
End Sub
Private Sub cTime_Change(Index As Integer)
    If IsDate(cTime(Index).Text) And Len(Trim(cTime(Index).Text)) = 8 Then
        Call RefreshGrid
    End If
End Sub

Private Sub cTime_LostFocus(Index As Integer)
    If Len(Trim(cTime(Index).Text)) = 0 Then
        If Index = 0 Then
            cTime(0).Text = "00:00:01"
        Else
            cTime(1).Text = "23:59:59"
        End If
    End If
End Sub
Private Sub FileExit_Click()
    Unload Me
End Sub
Private Sub flxM_EnterCell()
    Call FromGridToForm
End Sub
Private Sub flxM_KeyDown(KeyCode As Integer, Shift As Integer)
' * Trap Enter Key
    If KeyCode = vbKeyReturn Then
        Call ReadGrid
    End If
End Sub
Private Sub flxM_KeyPress(KeyAscii As Integer)
' * Trap Alpha - numeric Keys
    If KeyAscii <> 8 Then
        If Chr(KeyAscii) = "1" Or Chr(KeyAscii) = "2" Then
            If Len(Trim(flxM.TextMatrix(flxM.Row, 18))) = 0 Then
                flxM.TextMatrix(flxM.Row, 0) = Chr(KeyAscii)
            Else
                Beep
                KeyAscii = 0
            End If
        Else
            Beep
            KeyAscii = 0
        End If
    Else
        flxM.TextMatrix(flxM.Row, 0) = ""
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Call cmdExit_Click
        Case vbKeyF5
            If cmdRefresh.Enabled Then
                Call cmdRefresh_Click
            End If
        Case vbKeyF11
            If cmdTab.Enabled Then
                Call cmdTab_Click
            End If
        Case vbKeyF12
            If ST.Tab = 1 Then
                Call cmdRetrieve_Click
            End If
    End Select
End Sub
Private Sub Form_Load()
    Call GetInformation
    If UserType = "O" Then
        Call RefreshGrid
    Else
        ST.Tab = 1
        cmdTab.Width = 4815
        cmdTab.Caption = "F11 - Cons. Teller Inquiry"
        cmdRefresh.Enabled = False
        cmdRefresh.TabStop = False
        cmdTab.Enabled = False
        cmdTab.TabStop = False
        Call RefreshTeller
    End If
End Sub
Public Sub InitializeGrid()
    Dim X As Integer
    flxM.Clear
    flxM.Cols = 22
    flxM.Rows = 2
    flxM.FormatString = "Opt | Teller | Cash | ADR | Total |5|6|7|8|9|10|11|12|13|14|15|16|17|Status|19"
    flxM.ColWidth(0) = 700
    flxM.ColWidth(1) = 2100
    flxM.ColWidth(2) = 2100
    flxM.ColWidth(3) = 2100
    flxM.ColWidth(4) = 2700
    For X = 5 To 17
        flxM.ColWidth(X) = 1
    Next
    flxM.ColWidth(18) = 1700
    flxM.ColWidth(19) = 1
End Sub
Private Sub iTime_Change(Index As Integer)
        Call ClearBreakdown
End Sub
Private Sub txtDate_Change()
    Call ClearBreakdown
End Sub
Private Sub txtTeller_Change()
    Call ClearBreakdown
End Sub
Private Sub txtTeller_GotFocus()
    txtTeller.BackColor = &HFFFFFF
    txtTeller.SelStart = 0
    txtTeller.SelLength = Len(txtTeller.Text)
End Sub
Private Sub txtTeller_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        SendKeys "+{Tab}", True
    Else
        If KeyCode = 40 Then
            SendKeys "{Tab}", True
        End If
    End If
End Sub
Private Sub txtTeller_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    Else
        If KeyAscii = 27 Then
            SendKeys "+{Tab}", True
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
End Sub
Private Sub txtTeller_LostFocus()
    txtTeller.BackColor = &H8000000F
End Sub
Public Sub Tab0(D As Boolean)
    flxM.TabStop = D
    flxM.Enabled = D
    cTime(0).TabStop = D
    cTime(0).Enabled = D
    cTime(1).TabStop = D
    cTime(1).Enabled = D
End Sub
Public Sub Tab1(D As Boolean)
    txtDate.TabStop = D
    txtDate.Enabled = D
    txtTeller.TabStop = D
    txtTeller.Enabled = D
    iTime(0).TabStop = D
    iTime(0).Enabled = D
    iTime(1).TabStop = D
    iTime(1).Enabled = D
End Sub
Public Sub ClearBreakdown()
        lblArr.Caption = ".00"
        lblArrUG.Caption = ".00"
        lblVat.Caption = ".00"
        lblVatUG.Caption = ".00"
        lblNVat.Caption = ".00"
        lblNVatUG.Caption = ".00"
        lblWtx.Caption = ".00"
        lblWtxUG.Caption = ".00"
        lblSubTotal.Caption = ".00"
        lblSubTotalUG.Caption = ".00"
        lblWhf.Caption = ".00"
        lblAmountDue.Caption = ".00"
        lblTotal.Caption = ".00"
        lblAmountDue.Caption = ".00"
        lblPOS.Caption = ".00"
End Sub
Public Sub FillGrid()
    Dim TDay As Date
    Dim TotalCollection As Double
    Dim TotalClearedCollection As Double
    Dim rs As Recordset
    Dim FromDte As String
    Dim ToDte As String
    
    If Len(Trim(cTime(0).Text)) = 0 Then
        cTime(0).Text = "00:00:01"
    End If
    If Len(Trim(cTime(1).Text)) = 0 Then
        cTime(1).Text = "23:59:59"
    End If
    
    DE.GetDate TDay
    FromDte = Year(TDay) & "-" & Month(TDay) & "-" & Day(TDay) & " " & Trim(cTime(0).Text)
    ToDte = Year(TDay) & "-" & Month(TDay) & "-" & Day(TDay) & " " & Trim(cTime(1).Text)
    ThisDay.Caption = Format(TDay, "YYYY-MM-DD")
    DE.GetTellerTotals CDate(FromDte), CDate(ToDte)
    
    FlxRef = 0
    Set rs = DE.rsGetTellerTotals
    With rs
        TotalCollection = 0
        Call InitializeGrid
        Do While Not .EOF
            FlxRef = FlxRef + 1
            If FlxRef > 1 Then
                flxM.AddItem " "
            End If
            flxM.TextMatrix(FlxRef, 0) = " "
            flxM.TextMatrix(FlxRef, 1) = .Fields("userid")
            flxM.TextMatrix(FlxRef, 2) = Format(CDbl(.Fields("cash")), "###,###,###.#0")
            flxM.TextMatrix(FlxRef, 3) = Format(CDbl(.Fields("adr")), "###,###,###.#0")
            flxM.TextMatrix(FlxRef, 4) = Format(CDbl(.Fields("collection")), "###,###,###.#0")
            
            Call FillBreakdown(FromDte, ToDte, .Fields("userid") & "", FlxRef, 1)
            flxM.ColAlignment(18) = 0
            
            If .Fields("updcde") = "Y" Then
                flxM.TextMatrix(FlxRef, 18) = "CLEARED"
                TotalClearedCollection = TotalClearedCollection + CDbl(.Fields("collection"))
            Else
                TotalCollection = TotalCollection + CDbl(.Fields("collection"))
                flxM.TextMatrix(FlxRef, 18) = ""
            End If
            .MoveNext
        Loop
    End With
    rs.Close
    Set rs = Nothing

' ** Update Display
    
    lblTotalCollection.Caption = Format(TotalCollection, "###,###,###.#0")
    lblTotalClearedCollection.Caption = Format(TotalClearedCollection, "###,###,###.#0")
    
    Call FromGridToForm

End Sub
Public Sub FillBreakdown(FromD As String, ToD As String, User As String, cRef As Long, Target As Integer)

    Dim ArrAmt As Double
    Dim ArrAmtUG As Double
    Dim ArrVat As Double
    Dim ArrVatUG As Double
    Dim ArrNoVat As Double
    Dim ArrNoVatUG As Double
    Dim ArrWtx As Single
    Dim ArrWtxUG As Single
    Dim SubTotal As Double
    Dim SubTotalUG As Double
    Dim WhfAmt As Double
    Dim AmtDue As Double
    Dim GrandTotal As Double
    Dim rs As Recordset
    Dim Date1 As Date
    Dim Date2 As Date
    Dim AdrAmount As Double
    Dim POSAmt As Double


' ** ADR Amount
    DE.AdrAmount FromD, ToD, User
    Set rs = DE.rsAdrAmount
    If rs.RecordCount > 0 Then
        If Not IsNull(rs.Fields("out")) Then
            AdrAmount = RTDbl(rs.Fields("out"))
        Else
            AdrAmount = 0
        End If
    End If
    rs.Close
    Set rs = Nothing
' ** Arrastre
    DE.ArrastreAmt FromD, ToD, User
    Set rs = DE.rsArrastreAmt
    If rs.RecordCount > 0 Then
        If Not IsNull(rs.Fields("out")) Then
            ArrAmt = RTDbl(rs.Fields("out"))
        Else
            ArrAmt = 0
        End If
    End If
    rs.Close
    Set rs = Nothing
' ** Arrastre Under Guarantee
    DE.ArrastreAmtUG FromD, ToD, User
    Set rs = DE.rsArrastreAmtUG
    If rs.RecordCount > 0 Then
        If Not IsNull(rs.Fields("out")) Then
            ArrAmtUG = RTDbl(rs.Fields("out"))
        Else
            ArrAmtUG = 0
        End If
    End If
    rs.Close
    Set rs = Nothing
' ** Arrastre VAT
    DE.ArrastreVat FromD, ToD, User
    Set rs = DE.rsArrastreVat
    If rs.RecordCount > 0 Then
        If Not IsNull(rs.Fields("out")) Then
            ArrVat = RTDbl(rs.Fields("out"))
        Else
            ArrVat = 0
        End If
    End If
    rs.Close
    Set rs = Nothing
' ** Arrastre VAT Under Guarantee
    DE.ArrastreVatUG FromD, ToD, User
    Set rs = DE.rsArrastreVatUG
    If rs.RecordCount > 0 Then
        If Not IsNull(rs.Fields("out")) Then
            ArrVatUG = RTDbl(rs.Fields("out"))
        Else
            ArrVatUG = 0
        End If
    End If
    rs.Close
    Set rs = Nothing
' ** Arrastre Amount - No Vat
    DE.ArrastreNV FromD, ToD, User
    Set rs = DE.rsArrastreNV
    If rs.RecordCount > 0 Then
        If Not IsNull(rs.Fields("out")) Then
            ArrNoVat = RTDbl(rs.Fields("out"))
        Else
            ArrNoVat = 0
        End If
    End If
    rs.Close
    Set rs = Nothing
' ** Arrastre Amount - No Vat Under Guarantee
    DE.ArrastreNVUG FromD, ToD, User
    Set rs = DE.rsArrastreNVUG
    If rs.RecordCount > 0 Then
        If Not IsNull(rs.Fields("out")) Then
            ArrNoVatUG = RTDbl(rs.Fields("out"))
        Else
            ArrNoVatUG = 0
        End If
    End If
    rs.Close
    Set rs = Nothing
' ** Arrastre WTX
    DE.ArrastreWtx FromD, ToD, User
    Set rs = DE.rsArrastreWtx
    If rs.RecordCount > 0 Then
        If Not IsNull(rs.Fields("out")) Then
            ArrWtx = RTDbl(rs.Fields("out"))
        Else
            ArrWtx = 0
        End If
    End If
    rs.Close
    Set rs = Nothing
' ** Arrastre WTX UG
    DE.ArrastreWtxUG FromD, ToD, User
    Set rs = DE.rsArrastreWtxUG
    If rs.RecordCount > 0 Then
        If Not IsNull(rs.Fields("out")) Then
            ArrWtxUG = RTDbl(rs.Fields("out"))
        Else
            ArrWtxUG = 0
        End If
    End If
    rs.Close
    Set rs = Nothing
' ** Total Wharfage Amount
    DE.Wharfage FromD, ToD, User
    Set rs = DE.rsWharfage
    If rs.RecordCount > 0 Then
        If Not IsNull(rs.Fields("out")) Then
            WhfAmt = RTDbl(rs.Fields("out"))
        Else
            WhfAmt = 0
        End If
    End If
    rs.Close
    Set rs = Nothing
' ** POS Amounts
    DE.TotalPOS FromD, ToD, User
    Set rs = DE.rsTotalPOS
    If rs.RecordCount > 0 Then
        If Not IsNull(rs.Fields("POS")) Then
            POSAmt = RTDbl(rs.Fields("POS"))
        Else
            POSAmt = 0
        End If
    End If
    rs.Close
    Set rs = Nothing


' ** Where to save the results

' ** Target = '1' - Grid
' **              '2' - Form

    SubTotal = ArrAmt + ArrVat + ArrNoVat - ArrWtx
    SubTotalUG = ArrAmtUG + ArrVatUG + ArrNoVatUG - ArrWtxUG
    AmtDue = SubTotal
    GrandTotal = AmtDue + WhfAmt - AdrAmount - POSAmt
    
    Select Case Target
        Case 1
            With flxM
                .TextMatrix(cRef, 5) = Format(ArrAmt, "###,###,###.#0")
                .TextMatrix(cRef, 6) = Format(ArrAmtUG, "###,###,###.#0")
                .TextMatrix(cRef, 7) = Format(ArrVat, "###,###,###.#0")
                .TextMatrix(cRef, 8) = Format(ArrVatUG, "###,###,###.#0")
                .TextMatrix(cRef, 9) = Format(ArrNoVat, "###,###,###.#0")
                .TextMatrix(cRef, 10) = Format(ArrNoVatUG, "###,###,###.#0")
                .TextMatrix(cRef, 11) = Format(ArrWtx, "###,###,###.#0")
                .TextMatrix(cRef, 12) = Format(ArrWtxUG, "###,###,###.#0")
                .TextMatrix(cRef, 13) = Format(SubTotal, "###,###,###.#0")
                .TextMatrix(cRef, 14) = Format(SubTotalUG, "###,###,###.#0")
                .TextMatrix(cRef, 15) = Format(WhfAmt, "###,###,###.#0")
                .TextMatrix(cRef, 16) = Format(AmtDue, "###,###,###.#0")
                .TextMatrix(cRef, 17) = Format(GrandTotal, "###,###,###.#0")
                .TextMatrix(cRef, 19) = Format(AdrAmount, "###,###,###.#0")
                .TextMatrix(cRef, 20) = Format(POSAmt, "###,###,###.#0")
            End With
        Case 2
            lblArr.Caption = Format(ArrAmt, "###,###,###.#0")
            lblArrUG.Caption = Format(ArrAmtUG, "###,###,###.#0")
            lblVat.Caption = Format(ArrVat, "###,###,###.#0")
            lblVatUG.Caption = Format(ArrVatUG, "###,###,###.#0")
            lblNVat.Caption = Format(ArrNoVat, "###,###,###.#0")
            lblNVatUG.Caption = Format(ArrNoVatUG, "###,###,###.#0")
            lblWtx.Caption = Format(ArrWtx, "###,###,###.#0")
            lblWtxUG.Caption = Format(ArrWtxUG, "###,###,###.#0")
            lblSubTotal.Caption = Format(SubTotal, "###,###,###.#0")
            lblSubTotalUG.Caption = Format(SubTotalUG, "###,###,###.#0")
            lblWhf.Caption = Format(WhfAmt, "###,###,###.#0")
            lblAmountDue.Caption = Format(AmtDue, "###,###,###.#0")
            lblTotal.Caption = Format(GrandTotal, "###,###,###.#0")
            lblAdr.Caption = Format(AdrAmount, "###,###,###.#0")
            lblPOS.Caption = Format(POSAmt, "###,###,###.#0")
    End Select
End Sub
Public Function RTSng(Amt As String) As Single
    If Not IsNull(Amt) Then
        If IsNumeric(Amt) Then
            RTSng = CSng(Amt)
        Else
            RTSng = 0
        End If
    Else
        RTSng = 0
    End If
End Function
Public Function RTDbl(Amt As String) As Single
    If Not IsNull(Amt) Then
        If IsNumeric(Amt) Then
            RTDbl = CDbl(Amt)
        Else
            RTDbl = 0
        End If
    Else
        RTDbl = 0
    End If
End Function
Public Sub RefreshGrid()
'    Call InitializeGrid
    Call ClearBreakdown
    ST.Tab = 0
    Tab0 (True)
    Tab1 (False)
    Call FillGrid
    If flxM.Visible Then
        flxM.SetFocus
    End If
End Sub
Public Sub RefreshTeller()
    Call ClearBreakdown
    txtDate.Text = Year(Now) & "-" & Format(Month(Now), "00") & "-" & Format(Day(Now), "00")
    txtTeller.Text = gUserid
    iTime(0).Text = "00:00:01"
    iTime(1).Text = "23:59:59"
    ST.Tab = 1
    Tab0 (False)
    Tab1 (True)
    If txtDate.Visible Then
        txtDate.SetFocus
    End If
End Sub
Public Sub FromGridToForm()
    With flxM
            lblArr.Caption = .TextMatrix(.Row, 5)
            lblArrUG.Caption = .TextMatrix(.Row, 6)
            lblVat.Caption = .TextMatrix(.Row, 7)
            lblVatUG.Caption = .TextMatrix(.Row, 8)
            lblNVat.Caption = .TextMatrix(.Row, 9)
            lblNVatUG.Caption = .TextMatrix(.Row, 10)
            lblWtx.Caption = .TextMatrix(.Row, 11)
            lblWtxUG.Caption = .TextMatrix(.Row, 12)
            lblSubTotal.Caption = .TextMatrix(.Row, 13)
            lblSubTotalUG.Caption = .TextMatrix(.Row, 14)
            lblWhf.Caption = .TextMatrix(.Row, 15)
            lblAmountDue.Caption = .TextMatrix(.Row, 16)
            lblTotal.Caption = .TextMatrix(.Row, 17)
            lblAdr.Caption = .TextMatrix(.Row, 19)
    End With
End Sub
Public Sub ReadGrid()
    
    Dim UserN As String * 10
    Dim X As Integer
    Dim TDay As Date
    Dim FromDte As String
    Dim ToDte As String
    If Len(Trim(cTime(0).Text)) = 0 Then
        cTime(0).Text = "00:00:01"
    End If
    If Len(Trim(cTime(1).Text)) = 0 Then
        cTime(1).Text = "23:59:59"
    End If
    DE.GetDate TDay
    FromDte = Year(TDay) & "-" & Month(TDay) & "-" & Day(TDay) & " " & Trim(cTime(0).Text)
    ToDte = Year(TDay) & "-" & Month(TDay) & "-" & Day(TDay) & " " & Trim(cTime(1).Text)
    With flxM
        For X = 1 To FlxRef
            Select Case .TextMatrix(X, 0)
                Case "1"
                    Shell "Net Send " & Trim(.TextMatrix(X, 1)) & " PLEASE REMIT YOUR COLLECTION NOW ! ", vbMaximizedFocus
                    .TextMatrix(X, 0) = ""
                Case "2"
                    UserN = Trim(.TextMatrix(X, 1))
                    DE.RemitCCRPay FromDte, ToDte, UserN
                    DE.RemitCCRCyx FromDte, ToDte, UserN
                    .TextMatrix(X, 0) = ""
            End Select
        Next
    End With
    Call RefreshGrid
End Sub
