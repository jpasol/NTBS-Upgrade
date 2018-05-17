VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCYCancelCorrect 
   Caption         =   "Cancellation / Correction"
   ClientHeight    =   11010
   ClientLeft      =   270
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab sstMain 
      Height          =   10845
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   19129
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Correct Gatepass"
      TabPicture(0)   =   "frmCYCancelCorrect.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdSaveCorrectGatePass"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "fraDetail"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Cancel Gatepass"
      TabPicture(1)   =   "frmCYCancelCorrect.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "cmdCancelGatepass"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Correct Payment"
      TabPicture(2)   =   "frmCYCancelCorrect.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "fraPayment"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "View"
      TabPicture(3)   =   "frmCYCancelCorrect.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "Frame6"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame6 
         Caption         =   "Gatepass Detail"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   8055
         Left            =   -74760
         TabIndex        =   106
         Top             =   2160
         Width           =   14535
         Begin VB.TextBox txtUserID4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   130
            Top             =   6480
            Width           =   2655
         End
         Begin VB.TextBox txtContainerNo4 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   9240
            TabIndex        =   116
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txtCommodity4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   115
            Top             =   1200
            Width           =   5055
         End
         Begin VB.CheckBox chkCustomsGuard4 
            Caption         =   "Customs Guard?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   3600
            TabIndex        =   114
            Top             =   6000
            Width           =   2655
         End
         Begin VB.TextBox txtRemarks4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   113
            Top             =   5400
            Width           =   5055
         End
         Begin VB.TextBox txtVesselCode4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   7
            TabIndex        =   112
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox txtDeclaredWeight4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   15
            TabIndex        =   111
            Top             =   3600
            Width           =   2655
         End
         Begin VB.TextBox txtBoatNote4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   110
            Top             =   3000
            Width           =   1455
         End
         Begin VB.TextBox txtPDIGNo4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   15
            TabIndex        =   109
            Top             =   2400
            Width           =   2655
         End
         Begin VB.TextBox txtBroker4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   108
            Top             =   4800
            Width           =   5055
         End
         Begin VB.TextBox txtConsignee4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   107
            Top             =   1800
            Width           =   5055
         End
         Begin MSMask.MaskEdBox mskGatepassNo4 
            Height          =   405
            Left            =   3600
            TabIndex        =   117
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "User ID:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   37
            Left            =   1920
            TabIndex        =   129
            Top             =   6480
            Width           =   1575
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Gatepass Number:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   375
            Index           =   36
            Left            =   720
            TabIndex        =   128
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label lblMain 
            Caption         =   "Container No.:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   35
            Left            =   6960
            TabIndex        =   127
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Commodity:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   34
            Left            =   1560
            TabIndex        =   126
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Remarks:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   33
            Left            =   1920
            TabIndex        =   125
            Top             =   5400
            Width           =   1575
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Declared Weight:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   32
            Left            =   840
            TabIndex        =   124
            Top             =   3600
            Width           =   2655
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Vessel Code:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   31
            Left            =   1320
            TabIndex        =   123
            Top             =   4200
            Width           =   2175
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Boat Note:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   30
            Left            =   1800
            TabIndex        =   122
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "PDIG No.:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   29
            Left            =   1560
            TabIndex        =   121
            Top             =   2400
            Width           =   1935
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Broker:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   28
            Left            =   2160
            TabIndex        =   120
            Top             =   4800
            Width           =   1335
         End
         Begin VB.Label lblManifest 
            Alignment       =   1  'Right Justify
            Caption         =   "Consignee:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   119
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Consignee:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   27
            Left            =   1680
            TabIndex        =   118
            Top             =   1800
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   100
         Top             =   600
         Width           =   6135
         Begin VB.CommandButton cmdGetView 
            Caption         =   "Get"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   4440
            TabIndex        =   101
            Top             =   840
            Width           =   1575
         End
         Begin MSMask.MaskEdBox mskReference4 
            Height          =   405
            Left            =   3600
            TabIndex        =   102
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskSequence4 
            Height          =   405
            Left            =   3600
            TabIndex        =   103
            Top             =   840
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##"
            PromptChar      =   " "
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Reference:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   375
            Index           =   26
            Left            =   1680
            TabIndex        =   105
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Sequence:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   375
            Index           =   25
            Left            =   1800
            TabIndex        =   104
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdCancelGatepass 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   -62400
         TabIndex        =   19
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CommandButton cmdSaveCorrectGatePass 
         Caption         =   "F4  Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   -62520
         TabIndex        =   15
         Top             =   9840
         Width           =   2175
      End
      Begin VB.Frame Frame4 
         Caption         =   "Gatepass Detail"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   8055
         Left            =   -74760
         TabIndex        =   69
         Top             =   2400
         Width           =   14535
         Begin VB.TextBox txtContainerNo2 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   9240
            TabIndex        =   20
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txtCommodity2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   22
            Top             =   1200
            Width           =   5055
         End
         Begin VB.CheckBox chkCustomsGuard2 
            Caption         =   "Customs Guard?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   3600
            TabIndex        =   30
            Top             =   6000
            Width           =   2655
         End
         Begin VB.TextBox txtRemarks2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   29
            Top             =   5400
            Width           =   5055
         End
         Begin VB.TextBox txtVesselCode2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   7
            TabIndex        =   27
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox txtDeclaredWeight2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   15
            TabIndex        =   26
            Top             =   3600
            Width           =   2655
         End
         Begin VB.TextBox txtBoatNote2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   25
            Top             =   3000
            Width           =   1455
         End
         Begin VB.TextBox txtPDIGNo2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   15
            TabIndex        =   24
            Top             =   2400
            Width           =   2655
         End
         Begin VB.TextBox txtBroker2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   28
            Top             =   4800
            Width           =   5055
         End
         Begin VB.TextBox txtConsignee2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   23
            Top             =   1800
            Width           =   5055
         End
         Begin MSMask.MaskEdBox mskGatepassNo2 
            Height          =   405
            Left            =   3600
            TabIndex        =   21
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Gatepass Number:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   375
            Index           =   24
            Left            =   720
            TabIndex        =   80
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label lblMain 
            Caption         =   "Container No.:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   23
            Left            =   6960
            TabIndex        =   79
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Commodity:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   22
            Left            =   1560
            TabIndex        =   78
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Remarks:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   21
            Left            =   1920
            TabIndex        =   77
            Top             =   5400
            Width           =   1575
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Declared Weight:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   20
            Left            =   840
            TabIndex        =   76
            Top             =   3600
            Width           =   2655
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Vessel Code:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   19
            Left            =   1320
            TabIndex        =   75
            Top             =   4200
            Width           =   2175
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Boat Note:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   18
            Left            =   1800
            TabIndex        =   74
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "PDIG No.:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   17
            Left            =   1560
            TabIndex        =   73
            Top             =   2400
            Width           =   1935
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Broker:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   16
            Left            =   2160
            TabIndex        =   72
            Top             =   4800
            Width           =   1335
         End
         Begin VB.Label lblManifest 
            Alignment       =   1  'Right Justify
            Caption         =   "Consignee:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   0
            Left            =   5400
            TabIndex        =   71
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Consignee:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   15
            Left            =   1680
            TabIndex        =   70
            Top             =   1800
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   66
         Top             =   840
         Width           =   6135
         Begin VB.CommandButton cmdGetCancelGatepass 
            Caption         =   "Get"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   4440
            TabIndex        =   18
            Top             =   840
            Width           =   1575
         End
         Begin MSMask.MaskEdBox mskReference2 
            Height          =   405
            Left            =   3600
            TabIndex        =   16
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskSequence2 
            Height          =   405
            Left            =   3600
            TabIndex        =   17
            Top             =   840
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##"
            PromptChar      =   " "
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Reference:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   375
            Index           =   14
            Left            =   1680
            TabIndex        =   68
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Sequence:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   375
            Index           =   12
            Left            =   1800
            TabIndex        =   67
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame fraPayment 
         Height          =   9855
         Left            =   600
         TabIndex        =   56
         Top             =   600
         Width           =   13815
         Begin MSMask.MaskEdBox mskADRNum 
            Height          =   375
            Left            =   5400
            TabIndex        =   99
            Top             =   6120
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtBank 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   0
            Left            =   7440
            MaxLength       =   10
            TabIndex        =   85
            Top             =   3480
            Width           =   1815
         End
         Begin VB.TextBox txtBank 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   1
            Left            =   7440
            MaxLength       =   10
            TabIndex        =   84
            Top             =   3960
            Width           =   1815
         End
         Begin VB.TextBox txtBank 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   2
            Left            =   7440
            MaxLength       =   10
            TabIndex        =   83
            Top             =   4440
            Width           =   1815
         End
         Begin VB.TextBox txtBank 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   3
            Left            =   7440
            MaxLength       =   10
            TabIndex        =   82
            Top             =   4920
            Width           =   1815
         End
         Begin VB.TextBox txtBank 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Index           =   4
            Left            =   7440
            MaxLength       =   10
            TabIndex        =   81
            Top             =   5400
            Width           =   1815
         End
         Begin VB.Frame Frame2 
            Height          =   975
            Left            =   240
            TabIndex        =   64
            Top             =   480
            Width           =   6135
            Begin VB.CommandButton cmdGetCorrectPayment 
               Caption         =   "Get"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   4440
               TabIndex        =   32
               Top             =   360
               Width           =   1575
            End
            Begin MSMask.MaskEdBox mskReference3 
               Height          =   405
               Left            =   2400
               TabIndex        =   31
               Top             =   360
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   714
               _Version        =   393216
               MaxLength       =   8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   15
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "########"
               PromptChar      =   " "
            End
            Begin VB.Label lblMain 
               Alignment       =   1  'Right Justify
               Caption         =   "Reference:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000002&
               Height          =   375
               Index           =   13
               Left            =   480
               TabIndex        =   65
               Top             =   360
               Width           =   1815
            End
         End
         Begin VB.CommandButton cmdSaveCorrectPayment 
            Caption         =   "F4 Save"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   11520
            TabIndex        =   40
            Top             =   9240
            Width           =   2175
         End
         Begin VB.TextBox txtCustomerName 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   6840
            MaxLength       =   40
            TabIndex        =   38
            Top             =   7560
            Width           =   6615
         End
         Begin VB.TextBox txtCustomerCode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   2640
            MaxLength       =   6
            TabIndex        =   37
            Top             =   7560
            Width           =   1335
         End
         Begin MSMask.MaskEdBox mskAmountToPay 
            Height          =   405
            Left            =   2640
            TabIndex        =   33
            Top             =   2040
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   714
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCashAmount 
            Height          =   405
            Left            =   2640
            TabIndex        =   34
            Top             =   2880
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskADRAmount 
            Height          =   405
            Left            =   2640
            TabIndex        =   35
            Top             =   6120
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   714
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskChange 
            Height          =   405
            Left            =   2640
            TabIndex        =   36
            Top             =   6840
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   714
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskADRBalance 
            Height          =   405
            Left            =   6840
            TabIndex        =   39
            Top             =   8280
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   714
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCheckAmount 
            Height          =   405
            Index           =   1
            Left            =   2640
            TabIndex        =   86
            Top             =   3960
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCheckAmount 
            Height          =   405
            Index           =   0
            Left            =   2640
            TabIndex        =   87
            Top             =   3480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCheckAmount 
            Height          =   405
            Index           =   2
            Left            =   2640
            TabIndex        =   88
            Top             =   4440
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCheckAmount 
            Height          =   405
            Index           =   3
            Left            =   2640
            TabIndex        =   89
            Top             =   4920
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCheckAmount 
            Height          =   405
            Index           =   4
            Left            =   2640
            TabIndex        =   90
            Top             =   5400
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCheckNo 
            Height          =   405
            Index           =   0
            Left            =   5400
            TabIndex        =   91
            Top             =   3480
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCheckNo 
            Height          =   405
            Index           =   1
            Left            =   5400
            TabIndex        =   92
            Top             =   3960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCheckNo 
            Height          =   405
            Index           =   2
            Left            =   5400
            TabIndex        =   93
            Top             =   4440
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCheckNo 
            Height          =   405
            Index           =   3
            Left            =   5400
            TabIndex        =   94
            Top             =   4920
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskCheckNo 
            Height          =   405
            Index           =   4
            Left            =   5400
            TabIndex        =   95
            Top             =   5400
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label lblManifest 
            Alignment       =   1  'Right Justify
            Caption         =   "Check Amount:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   285
            Index           =   52
            Left            =   360
            TabIndex        =   98
            Top             =   3480
            Width           =   2175
         End
         Begin VB.Label lblManifest 
            Caption         =   "Check No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   30
            Left            =   5400
            TabIndex        =   97
            Top             =   3000
            Width           =   1815
         End
         Begin VB.Label lblManifest 
            Caption         =   "Bank Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   31
            Left            =   7440
            TabIndex        =   96
            Top             =   3000
            Width           =   1815
         End
         Begin VB.Label lblManifest 
            Alignment       =   1  'Right Justify
            Caption         =   "ADR Balance:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   70
            Left            =   4680
            TabIndex        =   63
            Top             =   8280
            Width           =   2055
         End
         Begin VB.Label lblManifest 
            Alignment       =   1  'Right Justify
            Caption         =   "Change:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   69
            Left            =   960
            TabIndex        =   62
            Top             =   6840
            Width           =   1575
         End
         Begin VB.Label lblManifest 
            Alignment       =   1  'Right Justify
            Caption         =   "Customer Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   68
            Left            =   4200
            TabIndex        =   61
            Top             =   7560
            Width           =   2535
         End
         Begin VB.Label lblManifest 
            Alignment       =   1  'Right Justify
            Caption         =   "Customer Code:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   67
            Left            =   120
            TabIndex        =   60
            Top             =   7560
            Width           =   2415
         End
         Begin VB.Label lblManifest 
            Alignment       =   1  'Right Justify
            Caption         =   "ADR Amount:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   53
            Left            =   480
            TabIndex        =   59
            Top             =   6120
            Width           =   2055
         End
         Begin VB.Label lblManifest 
            Alignment       =   1  'Right Justify
            Caption         =   "Cash Amount:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   51
            Left            =   480
            TabIndex        =   58
            Top             =   2880
            Width           =   2055
         End
         Begin VB.Label lblManifest 
            Alignment       =   1  'Right Justify
            Caption         =   "Amount to Pay:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   50
            Left            =   120
            TabIndex        =   57
            Top             =   2040
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1455
         Left            =   -74760
         TabIndex        =   53
         Top             =   840
         Width           =   6135
         Begin VB.CommandButton cmdGetCorrectGatePass 
            Caption         =   "Get"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   4440
            TabIndex        =   3
            Top             =   840
            Width           =   1575
         End
         Begin MSMask.MaskEdBox mskReference 
            Height          =   405
            Left            =   3600
            TabIndex        =   1
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskSequence 
            Height          =   405
            Left            =   3600
            TabIndex        =   2
            Top             =   840
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "##"
            PromptChar      =   " "
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Sequence:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   55
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Reference:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   375
            Index           =   1
            Left            =   1680
            TabIndex        =   54
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame fraDetail 
         Caption         =   "Gatepass Detail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   8055
         Left            =   -74760
         TabIndex        =   41
         Top             =   2400
         Width           =   14535
         Begin VB.TextBox txtConsignee 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1800
            Width           =   5055
         End
         Begin VB.TextBox txtBroker 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   12
            Top             =   4800
            Width           =   5055
         End
         Begin VB.TextBox txtPDIGNo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   15
            TabIndex        =   8
            Top             =   2400
            Width           =   2655
         End
         Begin VB.TextBox txtBoatNote 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   9
            Top             =   3000
            Width           =   1455
         End
         Begin VB.TextBox txtDeclaredWeight 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   15
            TabIndex        =   10
            Top             =   3600
            Width           =   2655
         End
         Begin VB.TextBox txtVesselCode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   7
            TabIndex        =   11
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox txtRemarks 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   13
            Top             =   5400
            Width           =   5055
         End
         Begin VB.CheckBox chkCustomsGuard 
            Caption         =   "Customs Guard?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   3600
            TabIndex        =   14
            Top             =   6000
            Width           =   2655
         End
         Begin VB.TextBox txtCommodity 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   3600
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1200
            Width           =   5055
         End
         Begin VB.TextBox txtContainerNo 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   400
            Left            =   9240
            TabIndex        =   5
            Top             =   600
            Width           =   2895
         End
         Begin MSMask.MaskEdBox mskGatePassNo 
            Height          =   405
            Left            =   3600
            TabIndex        =   4
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   714
            _Version        =   393216
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "########"
            PromptChar      =   " "
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Consignee:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   5
            Left            =   1680
            TabIndex        =   52
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label lblManifest 
            Alignment       =   1  'Right Justify
            Caption         =   "Consignee:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   22
            Left            =   5400
            TabIndex        =   51
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Broker:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   10
            Left            =   2160
            TabIndex        =   50
            Top             =   4800
            Width           =   1335
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "PDIG No.:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   6
            Left            =   1560
            TabIndex        =   49
            Top             =   2400
            Width           =   1935
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Boat Note:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   7
            Left            =   1800
            TabIndex        =   48
            Top             =   3000
            Width           =   1695
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Vessel Code:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   9
            Left            =   1320
            TabIndex        =   47
            Top             =   4200
            Width           =   2175
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Declared Weight:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   8
            Left            =   840
            TabIndex        =   46
            Top             =   3600
            Width           =   2655
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Remarks:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   11
            Left            =   1920
            TabIndex        =   45
            Top             =   5400
            Width           =   1575
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Commodity:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   4
            Left            =   1560
            TabIndex        =   44
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label lblMain 
            Caption         =   "Container No.:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Index           =   3
            Left            =   7080
            TabIndex        =   43
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label lblMain 
            Alignment       =   1  'Right Justify
            Caption         =   "Gatepass Number:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   375
            Index           =   2
            Left            =   720
            TabIndex        =   42
            Top             =   600
            Width           =   2775
         End
      End
   End
End
Attribute VB_Name = "frmCYCancelCorrect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cTabCorrectGatepass As Integer = 0
Const cTabCancelPayment As Integer = 1
Const cTabCorrectPayment As Integer = 2
Const cTabView As Integer = 3

Private Type CYMFields
    refnum As Long
    seqnum As Integer
    gpsnum As Long
    gpstyp As String * 1
    cntnum As String * 12
    enttyp As String * 1
    entnum As Long
    cntsze As Integer
    fulemp As String * 1
    forexm As String * 1
    cntovl As Currency
    cntovw As Currency
    cntovh As Currency
    ovzums As String * 1
    trncde As String * 1
    whfcde As String * 1
    whfrate As Currency
    whfonly As String * 1
    conscde As String * 1
    dclwgt As String * 15
    bilnum As String * 22
    regnum As String * 12
    crodte As Date
    vslcde As String * 7
    silnum As String * 8
    ordsup As String * 8
    boatnt As String * 8
    shplin As String * 7
    prtorg As String * 15
    cnsgne As String * 30
    Broker As String * 30
    brknum As String * 7
    PDIGNo As String * 15
    commod As String * 30
    dgrcls As String * 1
    dgramt As Currency
    pctdsc As Currency
    revton As Currency
    ovzamt As Currency
    stoday As Integer
    freday As Integer
    stosta As String * 1
    stoamt As Currency
    arramt As Currency
    whfamt As Currency
    wghamt As Currency
    rframt As Currency
    stovat As Currency
    arrvat As Currency
    wghvat As Currency
    rfrvat As Currency
    stotax As Currency
    arrtax As Currency
    wghtax As Currency
    rfrtax As Currency
    vatcde As String * 1
    gtycde As String * 1
    cusgrd As String * 1
    plugin As Date
    plugou As Date
    lstdch As Date
    mntdte As Date
    stobeg As Date
    stoend As Date
    remark As String * 30
    ppanum As Long
    status As String * 3
    UserID As String * 10
    sysdte As Date
    updcde As String * 1
    trknam As String * 20
    pltnum As String * 10
    trkchs As String * 35
    gtekpr As String * 10
    outdte As Date
End Type

Dim rstCYMGps As ADODB.Recordset
Dim rstCYMGpsZ As ADODB.Recordset
Dim rstCYMPAY As ADODB.Recordset
Dim rstACOCTN As ADODB.Recordset
Dim CYMField As CYMFields
Dim intResponse As Integer
Dim lngPreviousGatepass As Long
Dim curPreviousADRAmount As Currency
Dim curPreviousCashAmount As Currency
Dim curPreviousCheck1 As Currency
Dim curPreviousCheck2 As Currency
Dim curPreviousCheck3 As Currency
Dim curPreviousCheck4 As Currency
Dim curPreviousCheck5 As Currency
Dim strPreviousCustomerCode As String
Dim lngControlNo As Long
Dim intTabNumber As Integer

Private Sub FieldAdvance(pKeyCode As Integer, pPreviousControl As Control, pNextControl As Control)
    Select Case pKeyCode
        Case vbKeyDown
            If (TypeOf pNextControl Is TextBox) Or (TypeOf pNextControl Is MaskEdBox) Then
                pNextControl.SelStart = 0
                pNextControl.SelLength = pNextControl.MaxLength
            End If
            pNextControl.SetFocus
        Case vbKeyReturn
            If (TypeOf pNextControl Is TextBox) Or (TypeOf pNextControl Is MaskEdBox) Then
                pNextControl.SelStart = 0
                pNextControl.SelLength = pNextControl.MaxLength
            End If
            pNextControl.SetFocus
        Case vbKeyUp
            If (TypeOf pPreviousControl Is TextBox) Or (TypeOf pPreviousControl Is MaskEdBox) Then
                pPreviousControl.SelStart = 0
                pPreviousControl.SelLength = pPreviousControl.MaxLength
            End If
            pPreviousControl.SetFocus
        Case vbKeyF11
            Select Case sstMain.Tab
                Case cTabCorrectGatepass
                    sstMain.Tab = IIf(sstMain.TabEnabled(cTabCorrectGatepass) = True, cTabView, cTabCorrectGatepass)
                Case cTabCancelPayment
                    sstMain.Tab = IIf(sstMain.TabEnabled(cTabCancelPayment - 1) = True, cTabCancelPayment - 1, cTabCancelPayment)
                Case cTabCorrectPayment
                    sstMain.Tab = IIf(sstMain.TabEnabled(cTabCorrectPayment - 1) = True, cTabCorrectPayment - 1, cTabCorrectPayment)
                Case cTabView
                    sstMain.Tab = IIf(sstMain.TabEnabled(cTabView - 1) = True, cTabView - 1, cTabView)
            End Select
        Case vbKeyF12
            Select Case sstMain.Tab
                Case cTabCorrectGatepass
                    sstMain.Tab = IIf(sstMain.TabEnabled(cTabCorrectGatepass + 1) = True, cTabCorrectGatepass + 1, cTabCorrectGatepass)
                Case cTabCancelPayment
                    sstMain.Tab = IIf(sstMain.TabEnabled(cTabCancelPayment + 1) = True, cTabCancelPayment + 1, cTabCancelPayment)
                Case cTabCorrectPayment
                    sstMain.Tab = IIf(sstMain.TabEnabled(cTabCorrectPayment + 1) = True, cTabCorrectPayment + 1, cTabCorrectPayment)
                Case cTabView
                    sstMain.Tab = IIf(sstMain.TabEnabled(cTabView) = True, cTabCorrectGatepass, cTabView)
            End Select
        Case vbKeyF3
            intResponse = MsgBox("Do you really want to Exit?", vbYesNo + vbCritical, "Quit Program")
            If intResponse = vbYes Then
                Unload Me
            End If
        Case vbKeyF4
            Select Case sstMain.Tab
                Case cTabCorrectGatepass
                    Call cmdSaveCorrectGatePass_Click
                Case cTabCorrectPayment
                    Call cmdSaveCorrectPayment_Click
            End Select
    End Select
End Sub

Private Sub chkCustomsGuard_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtRemarks, cmdSaveCorrectGatePass)
End Sub

Private Sub cmdCancelGatepass_Click()
    intResponse = MsgBox("Are you sure you want to cancel this detail?", vbCritical + vbYesNo, "")
    If intResponse = vbYes Then
        intTabNumber = 2
        Call WriteCancelGatepassTab
        Call WriteToLogOrig
        Call WriteToLogUpdated
        Call InitializeCancelGatepassTab
        mskReference2.SetFocus
        intTabNumber = 0
    Else
        mskReference2.SetFocus
    End If
End Sub

Private Sub cmdCancelGatepass_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, cmdGetCancelGatepass, sstMain)
End Sub

Private Sub cmdGetCancelGatepass_Click()
    Call GetInfo2
End Sub

Private Sub cmdGetCorrectGatePass_Click()
    Call GetInfo
End Sub

Private Sub GetInfo()
    Dim blnPPAExist As Boolean
    Dim blnCancelled As Boolean
    
    Set rstCYMGps = New ADODB.Recordset
    rstCYMGps.LockType = adLockOptimistic
    rstCYMGps.CursorType = adOpenStatic
    rstCYMGps.Open "Select * from CYMGps where refnum =" & Val(mskReference) & " and seqnum =" & Val(mskSequence), gcnnBilling, , , adCmdText

    With rstCYMGps
        If .EOF And .BOF Then
            intResponse = MsgBox("Reference/Sequence Number not found.", vbOKOnly + vbInformation, "")
            mskReference.SetFocus
        Else
            blnPPAExist = Val(.Fields("ppanum")) > 0
            blnCancelled = Trim(.Fields("status")) = "CAN"
            If blnPPAExist Then
                intResponse = MsgBox("Cannot edit. PPA OR existing.", vbInformation + vbOKOnly, "")
                mskReference.SetFocus
            ElseIf blnCancelled Then
                intResponse = MsgBox("Cannot edit. Already cancelled.", vbInformation + vbOKOnly, "")
                mskReference.SetFocus
            Else
                Call getCYMFields
                mskGatePassNo = .Fields("gpsnum")
                lngPreviousGatepass = .Fields("gpsnum")
                txtContainerNo = .Fields("cntnum")
                txtCommodity = .Fields("commod")
                txtConsignee = .Fields("cnsgne")
                txtPDIGNo = .Fields("pdigno")
                txtBoatNote = .Fields("boatnt")
                txtDeclaredWeight = .Fields("dclwgt")
                txtVesselCode = .Fields("vslcde")
                txtBroker = .Fields("broker")
                txtRemarks = .Fields("remark")
                chkCustomsGuard.Value = ConvertToNum(.Fields("cusgrd"))
                mskGatePassNo.SetFocus
            End If
        End If
        .Close
    End With
End Sub

Private Sub getCYMFields()
    With rstCYMGps
        CYMField.refnum = .Fields("refnum")
        CYMField.seqnum = .Fields("seqnum")
        CYMField.gpsnum = .Fields("gpsnum")
        CYMField.gpstyp = .Fields("gpstyp")
        CYMField.cntnum = .Fields("cntnum")
        CYMField.enttyp = .Fields("enttyp")
        CYMField.entnum = .Fields("entnum")
        CYMField.cntsze = .Fields("cntsze")
        
        CYMField.fulemp = .Fields("fulemp")
        CYMField.forexm = .Fields("forexm")
        CYMField.cntovl = .Fields("cntovl")
        CYMField.cntovw = .Fields("cntovw")
        CYMField.cntovh = .Fields("cntovh")
        CYMField.ovzums = .Fields("ovzums")
        CYMField.trncde = .Fields("trncde")
        CYMField.whfcde = .Fields("whfcde")
        
        CYMField.whfrate = .Fields("whfrate")
        CYMField.whfonly = .Fields("whfonly")
        CYMField.conscde = .Fields("conscde")
        CYMField.dclwgt = .Fields("dclwgt")
        CYMField.bilnum = .Fields("bilnum")
        CYMField.regnum = .Fields("regnum")
        If IsDate(.Fields("crodte")) Then
            CYMField.crodte = .Fields("crodte")
        End If
        CYMField.vslcde = .Fields("vslcde")
        
        CYMField.silnum = .Fields("silnum")
        CYMField.ordsup = .Fields("ordsup")
        CYMField.boatnt = .Fields("boatnt")
        CYMField.shplin = .Fields("shplin")
        CYMField.prtorg = .Fields("prtorg")
        CYMField.cnsgne = .Fields("cnsgne")
        CYMField.Broker = .Fields("broker")
        CYMField.brknum = .Fields("brknum")
        
        CYMField.PDIGNo = .Fields("pdigno")
        CYMField.commod = .Fields("commod")
        CYMField.dgrcls = .Fields("dgrcls")
        CYMField.dgramt = .Fields("dgramt")
        CYMField.pctdsc = .Fields("pctdsc")
        CYMField.revton = .Fields("revton")
        CYMField.ovzamt = .Fields("ovzamt")
        CYMField.stoday = .Fields("stoday")
    
        CYMField.freday = .Fields("freday")
        CYMField.stosta = .Fields("stosta")
        CYMField.stoamt = .Fields("stoamt")
        CYMField.arramt = .Fields("arramt")
        CYMField.whfamt = .Fields("whfamt")
        CYMField.wghamt = .Fields("wghamt")
        CYMField.rframt = .Fields("rframt")
        CYMField.stovat = .Fields("stovat")
        
        CYMField.arrvat = .Fields("arrvat")
        CYMField.wghvat = .Fields("wghvat")
        CYMField.rfrvat = .Fields("rfrvat")
        CYMField.stotax = .Fields("stotax")
        CYMField.arrtax = .Fields("arrtax")
        CYMField.wghtax = .Fields("wghtax")
        CYMField.rfrtax = .Fields("rfrtax")
        CYMField.vatcde = .Fields("vatcde")
    
        CYMField.gtycde = .Fields("gtycde")
        CYMField.cusgrd = .Fields("cusgrd")
        If IsDate(.Fields("plugin")) Then
            CYMField.plugin = .Fields("plugin")
        End If
        
        If IsDate(.Fields("plugou")) Then
            CYMField.plugou = .Fields("plugou")
        End If
        
        If IsDate(.Fields("lstdch")) Then
            CYMField.lstdch = .Fields("lstdch")
        End If
        
        If IsDate(.Fields("mntdte")) Then
            CYMField.mntdte = .Fields("mntdte")
        End If
        
        If IsDate(.Fields("stobeg")) Then
            CYMField.stobeg = .Fields("stobeg")
        End If
        
        If IsDate(.Fields("stoend")) Then
            CYMField.stoend = .Fields("stoend")
        End If
        
        CYMField.remark = .Fields("remark")
        CYMField.ppanum = .Fields("ppanum")
        CYMField.status = .Fields("status")
        CYMField.UserID = .Fields("userid")
        CYMField.sysdte = .Fields("sysdte")
        
        CYMField.updcde = .Fields("updcde")
        If Not IsNull(.Fields("trknam")) Then
            CYMField.trknam = .Fields("trknam")
        End If
        If Not IsNull(.Fields("pltnum")) Then
            CYMField.pltnum = .Fields("pltnum")
        End If
        If Not IsNull(.Fields("trkchs")) Then
            CYMField.trkchs = .Fields("trkchs")
        End If
        If Not IsNull(.Fields("gtekpr")) Then
            CYMField.gtekpr = .Fields("gtekpr")
        End If
        If IsDate(.Fields("outdte")) Then
            CYMField.outdte = .Fields("outdte")
        End If
    End With
End Sub

Private Sub GetInfo2()
    Dim blnPPAExist As Boolean
    Dim blnCancelled As Boolean
    Dim blnADRPresent As Boolean
    
    Set rstCYMPAY = New ADODB.Recordset
    rstCYMPAY.LockType = adLockOptimistic
    rstCYMPAY.CursorType = adOpenStatic
    rstCYMPAY.Open "Select * from CYMPay where refnum =" & Val(mskReference2), gcnnBilling, , , adCmdText
    
    With rstCYMPAY
        If .EOF And .BOF Then
            intResponse = MsgBox("Reference Number not found.", vbOKOnly + vbInformation, "")
            mskReference2.SetFocus
        Else
            blnADRPresent = .Fields("adramt") > 0
        End If
    End With
    
    Set rstCYMGps = New ADODB.Recordset
    rstCYMGps.LockType = adLockOptimistic
    rstCYMGps.CursorType = adOpenStatic
    rstCYMGps.Open "Select * from CYMGps where refnum =" & Val(mskReference2) & " and seqnum =" & Val(mskSequence2), gcnnBilling, , , adCmdText

    With rstCYMGps
        If .EOF And .BOF Then
            intResponse = MsgBox("Reference/Sequence Number not found.", vbOKOnly + vbInformation, "")
            mskReference2.SetFocus
        Else
            blnPPAExist = Val(.Fields("ppanum")) > 0
            blnCancelled = Trim(.Fields("status")) = "CAN"
            
            If blnADRPresent Then
                intResponse = MsgBox("Cannot cancel. ADR Present.", vbInformation + vbOKOnly, "")
                mskReference2.SetFocus
            End If
            If blnPPAExist Then
                intResponse = MsgBox("Cannot cancel. PPA OR existing.", vbInformation + vbOKOnly, "")
                mskReference2.SetFocus
            ElseIf blnCancelled Then
                intResponse = MsgBox("Cannot cancel. Already cancelled.", vbInformation + vbOKOnly, "")
                mskReference2.SetFocus
            ElseIf blnADRPresent Then
                intResponse = MsgBox("Cannot cancel. ADR Present.", vbInformation + vbOKOnly, "")
                mskReference2.SetFocus
            Else
                Call getCYMFields
                mskGatepassNo2 = .Fields("gpsnum")
                txtContainerNo2 = .Fields("cntnum")
                txtCommodity2 = .Fields("commod")
                txtConsignee2 = .Fields("cnsgne")
                txtPDIGNo2 = .Fields("pdigno")
                txtBoatNote2 = .Fields("boatnt")
                txtDeclaredWeight2 = .Fields("dclwgt")
                txtVesselCode2 = .Fields("vslcde")
                txtBroker2 = .Fields("broker")
                txtRemarks2 = .Fields("remark")
                chkCustomsGuard2.Value = ConvertToNum(.Fields("cusgrd"))
                cmdCancelGatepass.SetFocus
            End If
        End If
        .Close
    End With
End Sub

Private Function ConvertToNum(pString As String)
    ConvertToNum = (IIf(Trim(pString) = "Y", 1, 0))
End Function

Private Function ConvertToChar(pValue As Integer) As String
    ConvertToChar = IIf(pValue = 1, "Y", "N")
End Function

Private Sub cmdGetCorrectGatePass_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, mskSequence, mskGatePassNo)
End Sub

Private Sub cmdGetCorrectPayment_Click()
    Set rstCYMPAY = New ADODB.Recordset
    rstCYMPAY.LockType = adLockOptimistic
    rstCYMPAY.CursorType = adOpenStatic
    rstCYMPAY.Open "Select * from CYMPay where refnum =" & Val(mskReference3), gcnnBilling, , , adCmdText
    
    With rstCYMPAY
        If .EOF And .BOF Then
            intResponse = MsgBox("Reference Number not found.", vbOKOnly + vbInformation, "")
            mskReference3.SetFocus
        Else
            txtCustomerCode = .Fields("cuscde")
            txtCustomerName = .Fields("cusnam")
            
            If IsNumeric(txtCustomerCode) Then
                mskADRBalance = lzGetADRBal(Trim(txtCustomerCode))
            End If
            mskCashAmount = .Fields("cshamt")
            '
            mskCheckAmount(0) = .Fields("chkamt1")
            mskCheckAmount(1) = .Fields("chkamt2")
            mskCheckAmount(2) = .Fields("chkamt3")
            mskCheckAmount(3) = .Fields("chkamt4")
            mskCheckAmount(4) = .Fields("chkamt5")
            '
            mskCheckNo(0) = .Fields("chkno1")
            mskCheckNo(1) = .Fields("chkno2")
            mskCheckNo(2) = .Fields("chkno3")
            mskCheckNo(3) = .Fields("chkno4")
            mskCheckNo(4) = .Fields("chkno5")
            '
            txtBank(0) = .Fields("chkbnk1")
            txtBank(1) = .Fields("chkbnk2")
            txtBank(2) = .Fields("chkbnk3")
            txtBank(3) = .Fields("chkbnk4")
            txtBank(4) = .Fields("chkbnk5")
            '
            mskADRAmount = .Fields("adramt")
            mskADRNum = .Fields("adrnum")
            mskChange = .Fields("chgamt")
            mskAmountToPay = CCur(mskCashAmount) + CCur(mskCheckAmount(0)) + CCur(mskCheckAmount(1)) _
                                + CCur(mskCheckAmount(2)) + CCur(mskCheckAmount(3)) + CCur(mskCheckAmount(4)) _
                                + CCur(mskADRAmount) - CCur(mskChange)
            curPreviousADRAmount = CCur(mskADRAmount)
            curPreviousCashAmount = CCur(mskCashAmount)
            curPreviousCheck1 = CCur(mskCheckAmount(0))
            curPreviousCheck2 = CCur(mskCheckAmount(1))
            curPreviousCheck3 = CCur(mskCheckAmount(2))
            curPreviousCheck4 = CCur(mskCheckAmount(3))
            curPreviousCheck5 = CCur(mskCheckAmount(4))
            strPreviousCustomerCode = .Fields("cuscde")
            mskCashAmount.SetFocus
        End If
        .Close
    End With
End Sub

Private Sub cmdGetView_Click()
    Set rstCYMGps = New ADODB.Recordset
    rstCYMGps.LockType = adLockOptimistic
    rstCYMGps.CursorType = adOpenStatic
    rstCYMGps.Open "Select * from CYMGps where refnum =" & Val(mskReference4) & " and seqnum =" & Val(mskSequence4), gcnnBilling, , , adCmdText

    With rstCYMGps
        If .EOF And .BOF Then
            intResponse = MsgBox("Reference/Sequence Number not found.", vbOKOnly + vbInformation, "")
            mskReference4.SelStart = 0
            mskReference4.SelLength = mskReference4.MaxLength
            mskReference4.SetFocus
        Else
            mskGatepassNo4 = ""
            txtContainerNo4 = ""
            txtCommodity4 = ""
            txtConsignee4 = ""
            txtPDIGNo4 = ""
            txtBoatNote4 = ""
            txtDeclaredWeight4 = ""
            txtVesselCode4 = ""
            txtBroker4 = ""
            txtRemarks4 = ""
            chkCustomsGuard4.Value = 0
            txtUserID4 = ""
            txtUserID4 = ""
            
            mskGatepassNo4 = .Fields("gpsnum")
            txtContainerNo4 = .Fields("cntnum")
            txtCommodity4 = .Fields("commod")
            txtConsignee4 = .Fields("cnsgne")
            txtPDIGNo4 = .Fields("pdigno")
            txtBoatNote4 = .Fields("boatnt")
            txtDeclaredWeight4 = .Fields("dclwgt")
            txtVesselCode4 = .Fields("vslcde")
            txtBroker4 = .Fields("broker")
            txtRemarks4 = .Fields("remark")
            chkCustomsGuard4.Value = ConvertToNum(.Fields("cusgrd"))
            txtUserID4 = .Fields("userid")
            mskSequence4.SetFocus
            mskSequence4.SelStart = 0
            mskSequence4.SelLength = mskSequence4.MaxLength
        End If
    End With
End Sub

Private Sub cmdSaveCorrectGatePass_Click()
    intResponse = MsgBox("Save the following changes?", vbYesNo + vbInformation, "")
    If intResponse = vbYes Then
        If CLng(mskGatePassNo) <> lngPreviousGatepass Then
            If lzChkCYMgpIfExist(mskGatePassNo) Then
                intResponse = MsgBox("Cannot continue. Gatepass already existing.", vbOKOnly + vbExclamation, "")
                mskGatePassNo.SetFocus
            Else
                intTabNumber = 1
                Call WriteCorrectGatepassTab
                Call WriteToLogOrig
                Call WriteToLogUpdated
                Call InitializeCorrectGatepassTab
                intTabNumber = 1
                mskReference.SetFocus
            End If
        Else
            intTabNumber = 1
            Call WriteCorrectGatepassTab
            Call WriteToLogOrig
            Call WriteToLogUpdated
            Call InitializeCorrectGatepassTab
            intTabNumber = 0
            mskReference.SetFocus
        End If
    End If
End Sub

Private Sub WriteToLogOrig()
    Set rstCYMGpsZ = New ADODB.Recordset
    rstCYMGpsZ.LockType = adLockOptimistic
    rstCYMGpsZ.CursorType = adOpenDynamic
    rstCYMGpsZ.Open "CYMgpsZ", gcnnBilling, , , adCmdTable
    With rstCYMGpsZ
        .AddNew
        .Fields("refnum") = CYMField.refnum
        .Fields("seqnum") = CYMField.seqnum
        .Fields("gpsnum") = CYMField.gpsnum
            
        .Fields("gpstyp") = CYMField.gpstyp
        .Fields("cntnum") = CYMField.cntnum
        .Fields("enttyp") = CYMField.enttyp
        .Fields("entnum") = CYMField.entnum
        .Fields("cntsze") = CYMField.cntsze
        
        .Fields("fulemp") = CYMField.fulemp
        .Fields("forexm") = CYMField.forexm
        .Fields("cntovl") = CYMField.cntovl
        .Fields("cntovw") = CYMField.cntovw
        .Fields("cntovh") = CYMField.cntovh
        .Fields("ovzums") = CYMField.ovzums
        .Fields("trncde") = CYMField.trncde
        .Fields("whfcde") = CYMField.whfcde
        
        .Fields("whfrate") = CYMField.whfrate
        .Fields("whfonly") = CYMField.whfonly
        .Fields("conscde") = CYMField.conscde
        .Fields("dclwgt") = CYMField.dclwgt
        .Fields("bilnum") = CYMField.bilnum
        .Fields("regnum") = CYMField.regnum
        If Not IsNull(CYMField.crodte) Then
            .Fields("crodte") = CYMField.crodte
        End If
        
        .Fields("vslcde") = CYMField.vslcde
        
        .Fields("silnum") = CYMField.silnum
        .Fields("ordsup") = CYMField.ordsup
        .Fields("boatnt") = CYMField.boatnt
        .Fields("shplin") = CYMField.shplin
        .Fields("prtorg") = CYMField.prtorg
        
        .Fields("cnsgne") = CYMField.cnsgne
        .Fields("broker") = CYMField.Broker
        .Fields("brknum") = CYMField.brknum
        .Fields("pdigno") = CYMField.PDIGNo
        
        .Fields("commod") = CYMField.commod
        
        .Fields("dgrcls") = CYMField.dgrcls
        .Fields("dgramt") = CYMField.dgramt
        .Fields("pctdsc") = CYMField.pctdsc
        .Fields("revton") = CYMField.revton
        .Fields("ovzamt") = CYMField.ovzamt
        .Fields("stoday") = CYMField.stoday
    
        .Fields("freday") = CYMField.freday
        .Fields("stosta") = CYMField.stosta
        .Fields("stoamt") = CYMField.stoamt
        .Fields("arramt") = CYMField.arramt
        .Fields("whfamt") = CYMField.whfamt
        .Fields("wghamt") = CYMField.wghamt
        .Fields("rframt") = CYMField.rframt
        .Fields("stovat") = CYMField.stovat
        
        .Fields("arrvat") = CYMField.arrvat
        .Fields("wghvat") = CYMField.wghvat
        .Fields("rfrvat") = CYMField.rfrvat
        .Fields("stotax") = CYMField.stotax
        .Fields("arrtax") = CYMField.arrtax
        .Fields("wghtax") = CYMField.wghtax
        .Fields("rfrtax") = CYMField.rfrtax
        .Fields("vatcde") = CYMField.vatcde
    
        .Fields("gtycde") = CYMField.gtycde
        .Fields("cusgrd") = CYMField.cusgrd
        
        If Not IsNull(CYMField.plugin) Then
            .Fields("plugin") = CYMField.plugin
        End If
        
        If Not IsNull(CYMField.plugou) Then
            .Fields("plugou") = CYMField.plugin
        End If
        
        If Not IsNull(CYMField.lstdch) Then
            .Fields("lstdch") = CYMField.lstdch
        End If
        
        If Not IsNull(CYMField.mntdte) Then
            .Fields("mntdte") = CYMField.mntdte
        End If
        
        If Not IsNull(CYMField.stobeg) Then
            .Fields("stobeg") = CYMField.stobeg
        End If
        
        If Not IsNull(CYMField.stoend) Then
            .Fields("stoend") = CYMField.stoend
        End If
        
        .Fields("remark") = CYMField.remark
        .Fields("ppanum") = CYMField.ppanum
        .Fields("status") = CYMField.status
        .Fields("userid") = CYMField.UserID
        .Fields("sysdte") = CYMField.sysdte
        .Fields("updcde") = CYMField.updcde
        If Not IsNull(CYMField.trknam) Then
            .Fields("trknam") = CYMField.trknam
        End If
        
        If Not IsNull(CYMField.pltnum) Then
            .Fields("pltnum") = CYMField.pltnum
        End If
        
        If Not IsNull(CYMField.trkchs) Then
            .Fields("trkchs") = CYMField.trkchs
        End If
        
        If Not IsNull(CYMField.gtekpr) Then
            .Fields("gtekpr") = CYMField.gtekpr
        End If
            
        If Not IsNull(CYMField.outdte) Then
            .Fields("outdte") = CYMField.outdte
        End If
        .Update
    End With
End Sub

Private Sub WriteToLogUpdated()
    With rstCYMGpsZ
        .AddNew
        .Fields("refnum") = CYMField.refnum
        .Fields("seqnum") = CYMField.seqnum
        If intTabNumber = 1 Then
            .Fields("gpsnum") = mskGatePassNo
        Else
            .Fields("gpsnum") = CYMField.gpsnum
        End If
            
        .Fields("gpstyp") = CYMField.gpstyp
        .Fields("cntnum") = CYMField.cntnum
        .Fields("enttyp") = CYMField.enttyp
        .Fields("entnum") = CYMField.entnum
        .Fields("cntsze") = CYMField.cntsze
        
        .Fields("fulemp") = CYMField.fulemp
        .Fields("forexm") = CYMField.forexm
        .Fields("cntovl") = CYMField.cntovl
        .Fields("cntovw") = CYMField.cntovw
        .Fields("cntovh") = CYMField.cntovh
        .Fields("ovzums") = CYMField.ovzums
        .Fields("trncde") = CYMField.trncde
        .Fields("whfcde") = CYMField.whfcde
        
        .Fields("whfrate") = CYMField.whfrate
        .Fields("whfonly") = CYMField.whfonly
        .Fields("conscde") = CYMField.conscde
        If intTabNumber = 1 Then
            .Fields("dclwgt") = txtDeclaredWeight
        Else
            .Fields("dclwgt") = CYMField.dclwgt
        End If
        .Fields("bilnum") = CYMField.bilnum
        .Fields("regnum") = CYMField.regnum
        If Not IsNull(CYMField.crodte) Then
            .Fields("crodte") = CYMField.crodte
        End If
        
        If intTabNumber = 1 Then
            .Fields("vslcde") = txtVesselCode
        Else
            .Fields("vslcde") = CYMField.vslcde
        End If
        
        .Fields("silnum") = CYMField.silnum
        .Fields("ordsup") = CYMField.ordsup
        If intTabNumber = 1 Then
            .Fields("boatnt") = txtBoatNote
        Else
            .Fields("boatnt") = CYMField.boatnt
        End If
        .Fields("shplin") = CYMField.shplin
        .Fields("prtorg") = CYMField.prtorg
        
        If intTabNumber = 1 Then
            .Fields("cnsgne") = txtConsignee
        Else
            .Fields("cnsgne") = CYMField.cnsgne
        End If
        
        If intTabNumber = 1 Then
            .Fields("broker") = txtBroker
        Else
            .Fields("broker") = CYMField.Broker
        End If
        
        .Fields("brknum") = CYMField.brknum
        
        If intTabNumber = 1 Then
            .Fields("pdigno") = txtPDIGNo
        Else
            .Fields("pdigno") = CYMField.PDIGNo
        End If
        
        If intTabNumber = 1 Then
            .Fields("commod") = txtCommodity
        Else
            .Fields("commod") = CYMField.commod
        End If
        
        .Fields("dgrcls") = CYMField.dgrcls
        .Fields("dgramt") = CYMField.dgramt
        .Fields("pctdsc") = CYMField.pctdsc
        .Fields("revton") = CYMField.revton
        .Fields("ovzamt") = CYMField.ovzamt
        .Fields("stoday") = CYMField.stoday
    
        .Fields("freday") = CYMField.freday
        .Fields("stosta") = CYMField.stosta
        .Fields("stoamt") = CYMField.stoamt
        .Fields("arramt") = CYMField.arramt
        .Fields("whfamt") = CYMField.whfamt
        .Fields("wghamt") = CYMField.wghamt
        .Fields("rframt") = CYMField.rframt
        .Fields("stovat") = CYMField.stovat
        
        .Fields("arrvat") = CYMField.arrvat
        .Fields("wghvat") = CYMField.wghvat
        .Fields("rfrvat") = CYMField.rfrvat
        .Fields("stotax") = CYMField.stotax
        .Fields("arrtax") = CYMField.arrtax
        .Fields("wghtax") = CYMField.wghtax
        .Fields("rfrtax") = CYMField.rfrtax
        .Fields("vatcde") = CYMField.vatcde
    
        .Fields("gtycde") = CYMField.gtycde
        If intTabNumber = 1 Then
            .Fields("cusgrd") = ConvertToChar(chkCustomsGuard.Value)
        Else
            .Fields("cusgrd") = CYMField.cusgrd
        End If
        
        If Not IsNull(CYMField.plugin) Then
            .Fields("plugin") = CYMField.plugin
        End If
        
        If Not IsNull(CYMField.plugou) Then
            .Fields("plugou") = CYMField.plugin
        End If
        
        If Not IsNull(CYMField.lstdch) Then
            .Fields("lstdch") = CYMField.lstdch
        End If
        
        If Not IsNull(CYMField.mntdte) Then
            .Fields("mntdte") = CYMField.mntdte
        End If
        
        If Not IsNull(CYMField.stobeg) Then
            .Fields("stobeg") = CYMField.stobeg
        End If
        
        If Not IsNull(CYMField.stoend) Then
            .Fields("stoend") = CYMField.stoend
        End If
        
        If intTabNumber = 1 Then
            .Fields("remark") = txtRemarks
        Else
            .Fields("remark") = CYMField.remark
        End If
        .Fields("ppanum") = CYMField.ppanum
        If intTabNumber = 1 Then
            .Fields("status") = "COR"
        Else
            .Fields("status") = "CAN"
        End If
        .Fields("userid") = gUserID
        .Fields("sysdte") = gzGetSysDate
        
'        .Fields("userid") = CYMField.userid
'        .Fields("sysdte") = CYMField.sysdte
        
        If intTabNumber = 1 Then
            .Fields("updcde") = "U"
        Else
            .Fields("updcde") = CYMField.updcde
        End If
        If Not IsNull(CYMField.trknam) Then
            .Fields("trknam") = CYMField.trknam
        End If
        
        If Not IsNull(CYMField.pltnum) Then
            .Fields("pltnum") = CYMField.pltnum
        End If
        
        If Not IsNull(CYMField.trkchs) Then
            .Fields("trkchs") = CYMField.trkchs
        End If
        
        If Not IsNull(CYMField.gtekpr) Then
            .Fields("gtekpr") = CYMField.gtekpr
        End If
            
        If Not IsNull(CYMField.outdte) Then
            .Fields("outdte") = CYMField.outdte
        End If
        .Update
        .Close
    End With
End Sub

Private Sub WriteCorrectGatepassTab()
    Set rstCYMGps = New ADODB.Recordset
    rstCYMGps.LockType = adLockOptimistic
    rstCYMGps.CursorType = adOpenDynamic
    rstCYMGps.Open "Select * from CYMGps where refnum =" & Val(mskReference) & " and seqnum =" & Val(mskSequence), gcnnBilling, , , adCmdText
    
    With rstCYMGps
        .Fields("gpsnum") = CLng(mskGatePassNo)
        .Fields("cntnum") = Trim(txtContainerNo)
        .Fields("commod") = txtCommodity
        .Fields("cnsgne") = txtConsignee
        .Fields("pdigno") = txtPDIGNo
        .Fields("boatnt") = txtBoatNote
        .Fields("dclwgt") = txtDeclaredWeight
        .Fields("vslcde") = txtVesselCode
        .Fields("broker") = txtBroker
        .Fields("remark") = txtRemarks
        .Fields("cusgrd") = ConvertToChar(chkCustomsGuard.Value)
        .Fields("updcde") = "U"
        .Update
        .Close
    End With
End Sub

Private Sub WriteCancelGatepassTab()
    Dim lngGatepass As Long
    Dim strContainer As String
    Set rstCYMGps = New ADODB.Recordset
    rstCYMGps.LockType = adLockOptimistic
    rstCYMGps.CursorType = adOpenDynamic
    rstCYMGps.Open "Select * from CYMGps where refnum =" & Val(mskReference2) & " and seqnum =" & Val(mskSequence2), gcnnBilling, , , adCmdText
    
    With rstCYMGps
        lngGatepass = .Fields("gpsnum")
        strContainer = Trim(.Fields("cntnum"))
        Call UpdateACOCTN(lngGatepass, strContainer)
        .Fields("status") = "CAN"
        .Update
        .Close
    End With
End Sub

Private Sub UpdateACOCTN(pGatePass As Long, pContainer As String)
    Dim cmdCancel As ADODB.Command
    
    Set cmdCancel = New ADODB.Command
    Set cmdCancel.ActiveConnection = gcnnBilling
    
    'for container
    cmdCancel.CommandText = "update Acoctn set ctn_gpsnum=0 where ctn_ctnnum like " & _
                            "'" & (Left(pContainer, 10) & "_") & "' and ctn_gpsnum = " & pGatePass
    cmdCancel.Execute
End Sub
Private Sub InitializeCorrectGatepassTab()
   'mskReference = ""
    mskSequence = ""
    mskGatePassNo = ""
    txtContainerNo = ""
    txtCommodity = ""
    txtConsignee = ""
    txtPDIGNo = ""
    txtBoatNote = ""
    txtDeclaredWeight = ""
    txtVesselCode = ""
    txtBroker = ""
    txtRemarks = ""
    chkCustomsGuard.Value = 0
End Sub

Private Sub InitializeCancelGatepassTab()
'    mskReference2 = ""
    mskSequence2 = ""
    mskGatepassNo2 = ""
    txtContainerNo2 = ""
    txtCommodity2 = ""
    txtConsignee2 = ""
    txtPDIGNo2 = ""
    txtBoatNote2 = ""
    txtDeclaredWeight2 = ""
    txtVesselCode2 = ""
    txtBroker2 = ""
    txtRemarks2 = ""
    chkCustomsGuard2.Value = 0
End Sub

Private Sub InitializeCorrectPayment()
    mskReference3 = ""
    mskAmountToPay = 0
    mskCashAmount = 0
    mskCheckAmount(0) = 0
    mskCheckAmount(1) = 0
    mskCheckAmount(2) = 0
    mskCheckAmount(3) = 0
    mskCheckAmount(4) = 0
    mskCheckNo(0) = 0
    mskCheckNo(1) = 0
    mskCheckNo(2) = 0
    mskCheckNo(3) = 0
    mskCheckNo(4) = 0
    txtBank(0) = ""
    txtBank(1) = ""
    txtBank(2) = ""
    txtBank(3) = ""
    txtBank(4) = ""
    mskADRAmount = 0
    mskADRNum = ""
    mskChange = 0
    txtCustomerCode = ""
    txtCustomerName = ""
    mskADRBalance = 0
End Sub

Private Sub cmdSaveCorrectPayment_Click()
    Dim curComputedTotal As Currency
    Dim curADRDifference As Currency
    
    If curPreviousADRAmount = mskADRAmount And curPreviousCashAmount = mskCashAmount Then
        'no need
    Else
        'compute first
        If mskChange < 0 Then
            intResponse = MsgBox("not balanced. please fix.", vbExclamation + vbOKOnly, "")
            Exit Sub
        Else
            If (curPreviousADRAmount <> CCur(mskADRAmount)) And (curPreviousADRAmount > 0 Or CCur(mskADRAmount) > 0) Then
                curADRDifference = curPreviousADRAmount + CCur(mskADRBalance) - mskADRAmount
                If curADRDifference < 0 Then
                    intResponse = MsgBox("insuffficient ADR amount. please check.", vbExclamation + vbOKOnly, "")
                    Exit Sub
                Else
                    If (curPreviousADRAmount = 0) And (CCur(mskADRAmount) > 0) Then
                        lngControlNo = lzApplyADR(txtCustomerCode, "CYM", CCur(mskReference3), CCur(mskADRAmount), UCase(zCurrentUser()), "")
                    ElseIf (curPreviousADRAmount > 0) And (CCur(mskADRAmount) = 0) Then
                        lngControlNo = lzVoidADR(strPreviousCustomerCode, mskADRNum, UCase(zCurrentUser()), "")
                    ElseIf (curPreviousADRAmount > 0) And (CCur(mskADRAmount) > 0) Then
                        lngControlNo = lzVoidADR(strPreviousCustomerCode, mskADRNum, UCase(zCurrentUser()), "")
                        lngControlNo = lzApplyADR(txtCustomerCode, "CYM", CCur(mskReference3), CCur(mskADRAmount), UCase(zCurrentUser()), "")
                    End If
                End If
            Else
                txtCustomerCode = ""
                txtCustomerName = ""
                mskADRAmount = 0
            End If
            Call SaveCorrectPayment
            Call InitializeCorrectPayment
        End If
    End If
End Sub

Private Sub SaveCorrectPayment()
    Set rstCYMPAY = New ADODB.Recordset
    rstCYMPAY.LockType = adLockOptimistic
    rstCYMPAY.CursorType = adOpenStatic
    rstCYMPAY.Open "Select * from CYMPay where refnum =" & Val(mskReference3), gcnnBilling, , , adCmdText
    
    With rstCYMPAY
        .Fields("cuscde") = txtCustomerCode
        .Fields("cusnam") = txtCustomerName
        .Fields("cshamt") = CCur(mskCashAmount)
        .Fields("adramt") = CCur(mskADRAmount)
        .Fields("adrnum") = lngControlNo
        .Fields("chgamt") = CCur(mskChange)
        .Fields("chkno1") = "" & mskCheckNo(0)
        .Fields("chkno2") = "" & mskCheckNo(1)
        .Fields("chkno3") = "" & mskCheckNo(2)
        .Fields("chkno4") = "" & mskCheckNo(3)
        .Fields("chkno5") = "" & mskCheckNo(4)
        .Fields("chkamt1") = mskCheckAmount(0)
        .Fields("chkamt2") = mskCheckAmount(1)
        .Fields("chkamt3") = mskCheckAmount(2)
        .Fields("chkamt4") = mskCheckAmount(3)
        .Fields("chkamt5") = mskCheckAmount(4)
        .Fields("chkbnk1") = "" & txtBank(0)
        .Fields("chkbnk2") = "" & txtBank(1)
        .Fields("chkbnk3") = "" & txtBank(2)
        .Fields("chkbnk4") = "" & txtBank(3)
        .Fields("chkbnk5") = "" & txtBank(4)
        .Update
        .Close
    End With
End Sub

Private Sub Form_Activate()
    mskReference.SetFocus
End Sub

Private Sub mskADRAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, mskADRAmount, txtCustomerCode)
End Sub

Private Sub mskADRAmount_LostFocus()
    If Not IsNumeric(mskADRAmount) Then
            mskADRAmount = 0
    Else
        If (CCur(mskADRAmount) > CCur(mskADRBalance)) Or (CCur(mskADRAmount) > CCur(mskAmountToPay)) Then
            mskADRAmount = mskAmountToPay
        End If
    End If
    Call SumPaymentTypes
    If CCur(mskADRAmount) = 0 Then
        txtCustomerCode = ""
        txtCustomerName = ""
        mskADRBalance = 0
        mskADRAmount.Enabled = False
    End If
    mskADRAmount = Format(mskADRAmount, "###,###,##0.00")
End Sub

Private Sub mskCashAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, cmdGetCorrectPayment, mskCheckAmount(0))
End Sub

Private Sub mskCashAmount_LostFocus()
    mskCashAmount = Format(mskCashAmount, "###,###,##0.00")
    Call SumPaymentTypes
End Sub

Private Sub SumPaymentTypes()
    Dim intCheckCtr As Integer
    Dim curTotalCheckAMT As Currency
    If Not IsNumeric(mskCashAmount) Then mskCashAmount = 0
    For intCheckCtr = 0 To 4
        If Not IsNumeric(mskCheckAmount(intCheckCtr)) Then mskCheckAmount(intCheckCtr) = 0
        curTotalCheckAMT = curTotalCheckAMT + CCur(mskCheckAmount(intCheckCtr))
    Next intCheckCtr
    If Not IsNumeric(mskADRAmount) Then mskADRAmount = 0
    If Not IsNumeric(mskAmountToPay) Then mskAmountToPay = 0

    mskChange = CCur(mskADRAmount) + curTotalCheckAMT + CCur(mskCashAmount) - CCur(mskAmountToPay)
End Sub


Private Sub mskCheckAmount_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 0
            Call FieldAdvance(KeyCode, mskCashAmount, mskCheckNo(Index))
        Case 1, 2, 3, 4
            Call FieldAdvance(KeyCode, mskCheckAmount(Index - 1), mskCheckNo(Index))
    End Select
End Sub

Private Sub mskCheckAmount_LostFocus(Index As Integer)
    If Not IsNumeric(mskCheckAmount(Index)) Then mskCheckAmount(Index) = 0
    mskCheckAmount(Index) = Format(mskCheckAmount(Index), "###,###,##0.00")
    Call SumPaymentTypes
End Sub

Private Sub mskCheckNo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, mskCheckAmount(Index), txtBank(Index))
End Sub

Private Sub mskCheckNo_LostFocus(Index As Integer)
    If Not IsNumeric(mskCheckNo(Index)) Then mskCheckNo(Index) = 0
End Sub

Private Sub mskGatePassNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, cmdGetCorrectGatePass, txtCommodity)
End Sub

Private Sub mskReference_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, sstMain, mskSequence)
End Sub

Private Sub mskReference2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, sstMain, mskSequence2)
End Sub

Private Sub mskReference3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, sstMain, cmdGetCorrectPayment)
End Sub

Private Sub mskReference4_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, sstMain, mskSequence4)
End Sub

Private Sub mskSequence_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, mskReference, cmdGetCorrectGatePass)
End Sub

Private Sub mskSequence2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, mskReference2, cmdGetCancelGatepass)
End Sub

Private Sub mskSequence4_KeyDown(KeyCode As Integer, Shift As Integer)
     Call FieldAdvance(KeyCode, mskReference4, cmdGetView)
End Sub

Private Sub sstMain_Click(PreviousTab As Integer)
    Select Case sstMain.Tab
        Case cTabCorrectGatepass
            mskReference.SetFocus
        Case cTabCancelPayment
            mskReference2.SetFocus
        Case cTabCorrectPayment
            mskReference3.SetFocus
        Case cTabView
            mskReference4.SetFocus
    End Select
End Sub

Private Sub sstMain_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyDown Then
        Select Case sstMain.Tab
            Case cTabCorrectGatepass
                mskReference.SetFocus
            Case cTabCancelPayment
                mskReference2.SetFocus
            Case cTabCorrectPayment
                mskReference3.SetFocus
            Case cTabView
                mskReference4.SetFocus
        End Select
    ElseIf KeyCode = vbKeyF11 Or KeyCode = vbKeyF12 Then
        Call FieldAdvance(KeyCode, sstMain, sstMain)
    End If
End Sub

Private Sub txtBank_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 0, 1, 2, 3
            Call FieldAdvance(KeyCode, mskCheckNo(Index), mskCheckAmount(Index + 1))
        Case 4
            Call FieldAdvance(KeyCode, mskCheckNo(Index), txtCustomerCode)
    End Select
End Sub

Private Sub txtBank_LostFocus(Index As Integer)
    txtBank(Index) = "" & txtBank(Index)
End Sub

Private Sub txtBoatNote_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtPDIGNo, txtDeclaredWeight)
End Sub

Private Sub txtBroker_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtVesselCode, txtRemarks)
End Sub

Private Sub txtCommodity_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, mskGatePassNo, txtConsignee)
End Sub

Private Sub txtConsignee_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtCommodity, txtPDIGNo)
End Sub

Private Sub txtCustomerCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        mskADRBalance = lzGetADRBal(txtCustomerCode)
        txtCustomerName = Left(lzGetCustomerName(txtCustomerCode), 30)

        mskADRAmount = mskADRBalance
        If CCur(mskADRAmount) > CCur(mskAmountToPay) Then
            mskADRAmount = mskAmountToPay - (CCur(mskCashAmount) + CCur(mskCheckAmount(0)) + CCur(mskCheckAmount(1)) _
                         + CCur(mskCheckAmount(2)) + CCur(mskCheckAmount(3)) + CCur(mskCheckAmount(4)))
            If CCur(mskADRAmount) > CCur(mskADRBalance) Then
                mskADRAmount = mskADRBalance
            End If
        Else
            mskADRAmount = mskADRBalance
        End If
        Call SumPaymentTypes
        If txtCustomerName <> "" Then
            mskADRAmount.Enabled = True
            Call FieldAdvance(KeyCode, txtCustomerCode, mskADRAmount)
        End If
    Else
        Call FieldAdvance(KeyCode, mskCheckAmount(4), cmdSaveCorrectPayment)
    End If
End Sub

Private Sub txtDeclaredWeight_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtBoatNote, txtVesselCode)
End Sub

Private Sub txtPDIGNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtConsignee, txtBoatNote)
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtBroker, chkCustomsGuard)
End Sub

Private Sub txtVesselCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtDeclaredWeight, txtBroker)
End Sub

Private Function lzApplyADR(ByVal pCusCde As String, _
                            ByVal pREFTYP As String, _
                            ByVal pREFNUM As Long, _
                            ByVal pADRAmt As Currency, _
                            ByVal pUserID As String, _
                            ByVal pRemark As String) As Long

Dim cmdGetCustomer As ADODB.Command
Dim prmGetCustomer As ADODB.Parameter
    
    ' create command
    Set cmdGetCustomer = New ADODB.Command
    With cmdGetCustomer
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_applyadr"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pCusCde
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adChar
        .Parameters(2).Value = pREFTYP
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adNumeric
        .Parameters(3).Value = pREFNUM
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Value = pADRAmt
        .Parameters(4).Direction = adParamInput
        .Parameters(5).Type = adChar
        .Parameters(5).Value = pRemark
        .Parameters(5).Direction = adParamInput
        .Parameters(6).Type = adChar
        .Parameters(6).Value = pUserID
        .Parameters(6).Direction = adParamInput
       
        .Execute
        
        lzApplyADR = .Parameters(0)
        If lzApplyADR > 0 Then
            MsgBox "ADR Control Number:  " & Trim(Str(.Parameters(0))), vbInformation
        Else
            MsgBox "Error on ADR transaction. Please check all values, then retry.", vbQuestion
        End If
        
     End With
    
End Function

Private Function lzVoidADR(ByVal pCusCde As String, _
                           ByVal pREFNUM As Long, _
                           ByVal pUserID As String, _
                           ByVal pReason As String) As Long
Dim cmdGetCustomer As ADODB.Command
Dim prmGetCustomer As ADODB.Parameter
    
    ' create command
    Set cmdGetCustomer = New ADODB.Command
    With cmdGetCustomer
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_voidadrtran"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pCusCde
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adNumeric
        .Parameters(2).Value = pREFNUM
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adChar
        .Parameters(3).Value = pReason
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adChar
        .Parameters(4).Value = pUserID
        .Parameters(4).Direction = adParamInput
       
        .Execute
        
        lzVoidADR = .Parameters(0)
        If lzVoidADR > 0 Then
            MsgBox "ADR Control Number:  " & Trim(Str(.Parameters(0))), vbInformation
        Else
            MsgBox "Error on ADR transaction. Please check all values, then retry.", vbQuestion
        End If
     
     End With
    
End Function

Private Function lzGetADRBal(ByVal pCode As String) As Currency
Dim cmdGetADRBal As ADODB.Command
Dim prmGetADRBal As ADODB.Parameter
    
    ' create command
    Set cmdGetADRBal = New ADODB.Command
    With cmdGetADRBal
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getadrbal"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pCode
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adCurrency
        .Parameters(2).Direction = adParamOutput
        .Execute

        If Not IsNull(.Parameters(2)) Then
            lzGetADRBal = .Parameters(2)
        Else
            lzGetADRBal = 0
        End If
        
    End With

End Function

Private Function lzGetCustomerName(ByVal pCode As String) As String
Dim cmdGetCustomer As ADODB.Command
Dim prmGetCustomer As ADODB.Parameter
    
    ' create command
    Set cmdGetCustomer = New ADODB.Command
    With cmdGetCustomer
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getcustomerinfo"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pCode
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adChar
        .Parameters(2).Direction = adParamOutput
        .Parameters(3).Type = adChar
        .Parameters(3).Direction = adParamOutput
        .Parameters(4).Type = adChar
        .Parameters(4).Direction = adParamOutput
        .Parameters(5).Type = adChar
        .Parameters(5).Direction = adParamOutput
        .Parameters(6).Type = adChar
        .Parameters(6).Direction = adParamOutput
       
        .Execute
        
        lzGetCustomerName = Trim("" & .Parameters(3))
     End With
    
End Function

