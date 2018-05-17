VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCYSCorrection 
   Caption         =   "CY Special Services Voiding / Correction"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "IBM3270 - 1254"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CYSCorrection.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7665
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tabMenu 
      Height          =   7515
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   13256
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   794
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CCR VOIDING && CORRECTION"
      TabPicture(0)   =   "CYSCorrection.frx":014A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(3)=   "cmdVoid"
      Tab(0).Control(4)=   "cmdExit"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "PAYMENT CORRECTION"
      TabPicture(1)   =   "CYSCorrection.frx":0166
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cmdSave"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame Frame4 
         Caption         =   "Reference No."
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   840
         Left            =   225
         TabIndex        =   51
         Top             =   600
         Width           =   3315
         Begin VB.CommandButton cmdPayGet 
            Caption         =   "&Get"
            Height          =   390
            Left            =   1950
            TabIndex        =   9
            Top             =   300
            Width           =   1140
         End
         Begin MSMask.MaskEdBox txtPayRefNo 
            Height          =   390
            Left            =   225
            TabIndex        =   8
            Top             =   300
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   5265
         Left            =   225
         TabIndex        =   39
         Top             =   1425
         Width           =   8865
         Begin MSMask.MaskEdBox txtCshAmt 
            Height          =   390
            Left            =   3225
            TabIndex        =   10
            Top             =   675
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            ForeColor       =   16711680
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkAmt 
            Height          =   390
            Index           =   0
            Left            =   150
            TabIndex        =   11
            Top             =   1575
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkAmt 
            Height          =   390
            Index           =   1
            Left            =   150
            TabIndex        =   14
            Top             =   2025
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkAmt 
            Height          =   390
            Index           =   2
            Left            =   150
            TabIndex        =   17
            Top             =   2475
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkAmt 
            Height          =   390
            Index           =   3
            Left            =   150
            TabIndex        =   20
            Top             =   2925
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkAmt 
            Height          =   390
            Index           =   4
            Left            =   150
            TabIndex        =   23
            Top             =   3375
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkBank 
            Height          =   390
            Index           =   0
            Left            =   6225
            TabIndex        =   13
            Top             =   1575
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkBank 
            Height          =   390
            Index           =   1
            Left            =   6225
            TabIndex        =   16
            Top             =   2025
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkBank 
            Height          =   390
            Index           =   2
            Left            =   6225
            TabIndex        =   19
            Top             =   2475
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkBank 
            Height          =   390
            Index           =   3
            Left            =   6225
            TabIndex        =   22
            Top             =   2925
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkBank 
            Height          =   390
            Index           =   4
            Left            =   6225
            TabIndex        =   25
            Top             =   3375
            Width           =   2490
            _ExtentX        =   4392
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkNo 
            Height          =   390
            Index           =   0
            Left            =   3225
            TabIndex        =   12
            Top             =   1575
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkNo 
            Height          =   390
            Index           =   1
            Left            =   3225
            TabIndex        =   15
            Top             =   2025
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkNo 
            Height          =   390
            Index           =   2
            Left            =   3225
            TabIndex        =   18
            Top             =   2475
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkNo 
            Height          =   390
            Index           =   3
            Left            =   3225
            TabIndex        =   21
            Top             =   2925
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtChkNo 
            Height          =   390
            Index           =   4
            Left            =   3225
            TabIndex        =   24
            Top             =   3375
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtCusCode 
            Height          =   390
            Left            =   2175
            TabIndex        =   27
            Top             =   4725
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   688
            _Version        =   393216
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "######"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtADRAmt 
            Height          =   390
            Left            =   150
            TabIndex        =   26
            Top             =   4725
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   688
            _Version        =   393216
            ForeColor       =   16711680
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtCustomer 
            Height          =   390
            Left            =   3600
            TabIndex        =   28
            Top             =   4725
            Width           =   5115
            _ExtentX        =   9022
            _ExtentY        =   688
            _Version        =   393216
            ForeColor       =   -2147483641
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   40
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CUSTOMER NAME"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   3600
            TabIndex        =   52
            Top             =   4350
            Width           =   5115
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ADR AMOUNT"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   150
            TabIndex        =   50
            Top             =   4350
            Width           =   1965
         End
         Begin VB.Label Label55 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CODE"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   2175
            TabIndex        =   49
            Top             =   4350
            Width           =   1365
         End
         Begin VB.Label lblChkTot 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   150
            TabIndex        =   48
            Top             =   3825
            Width           =   2565
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CHECK NO"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   3225
            TabIndex        =   47
            Top             =   1200
            Width           =   2565
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CHECK BANK"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   6225
            TabIndex        =   46
            Top             =   1200
            Width           =   2490
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CHECK AMT"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   150
            TabIndex        =   45
            Top             =   1200
            Width           =   2565
         End
         Begin VB.Label lblAmtDue 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   150
            TabIndex        =   44
            Top             =   675
            Width           =   2565
         End
         Begin VB.Label lblChange 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   6225
            TabIndex        =   43
            Top             =   675
            Width           =   2490
         End
         Begin VB.Label Label48 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CASH AMT"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   3225
            TabIndex        =   42
            Top             =   300
            Width           =   2565
         End
         Begin VB.Label Label50 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AMOUNT DUE"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   150
            TabIndex        =   41
            Top             =   300
            Width           =   2565
         End
         Begin VB.Label Label45 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CHANGE"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   6225
            TabIndex        =   40
            Top             =   300
            Width           =   2490
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "E&xit"
         Height          =   465
         Left            =   7650
         TabIndex        =   30
         Top             =   6825
         Width           =   1365
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   465
         Left            =   6150
         TabIndex        =   29
         Top             =   6825
         Width           =   1365
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   465
         Left            =   -67350
         TabIndex        =   7
         Top             =   6825
         Width           =   1365
      End
      Begin VB.CommandButton cmdVoid 
         Caption         =   "&Void"
         Enabled         =   0   'False
         Height          =   465
         Left            =   -68850
         TabIndex        =   6
         Top             =   6825
         Width           =   1365
      End
      Begin VB.Frame Frame3 
         Caption         =   "CCR Number"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   840
         Left            =   -70350
         TabIndex        =   35
         Top             =   600
         Width           =   3540
         Begin VB.CommandButton cmdCANSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   390
            Left            =   2025
            TabIndex        =   5
            Top             =   300
            Width           =   1365
         End
         Begin MSMask.MaskEdBox txtCANCCRNo 
            Height          =   390
            Left            =   225
            TabIndex        =   4
            Top             =   300
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "########"
            PromptChar      =   " "
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "CCR Info"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   5115
         Left            =   -74775
         TabIndex        =   33
         Top             =   1575
         Width           =   8790
         Begin MSFlexGridLib.MSFlexGrid grdVoid 
            Height          =   3615
            Left            =   150
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   825
            Width           =   8490
            _ExtentX        =   14975
            _ExtentY        =   6376
            _Version        =   393216
            Rows            =   10
            Cols            =   6
            FocusRect       =   0
            HighLight       =   0
            FillStyle       =   1
            ScrollBars      =   0
            SelectionMode   =   1
            FormatString    =   "^  |  RATE  |^ CONTAINER # |^ SZ |REFERENCE|>   AMOUNT  "
         End
         Begin VB.Label lblVoided 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "VOIDED"
            ForeColor       =   &H8000000E&
            Height          =   390
            Left            =   7425
            TabIndex        =   38
            Top             =   4575
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblCANIssueInfo 
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   150
            TabIndex        =   37
            Top             =   4575
            Width           =   7215
         End
         Begin VB.Label Label4 
            Caption         =   "Customer"
            Height          =   315
            Left            =   150
            TabIndex        =   36
            Top             =   375
            Width           =   1515
         End
         Begin VB.Label lblCANCustomer 
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   1650
            TabIndex        =   34
            Top             =   300
            Width           =   6990
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Reference / Sequence"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   840
         Left            =   -74775
         TabIndex        =   32
         Top             =   600
         Width           =   4140
         Begin VB.CommandButton cmdCANGet 
            Caption         =   "&Get"
            Height          =   390
            Left            =   2775
            TabIndex        =   3
            Top             =   300
            Width           =   1140
         End
         Begin MSMask.MaskEdBox txtCANRefNo 
            Height          =   390
            Left            =   225
            TabIndex        =   1
            Top             =   300
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   8
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtCANSeqNo 
            Height          =   390
            Left            =   1950
            TabIndex        =   2
            Top             =   300
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   3
            Mask            =   "###"
            PromptChar      =   " "
         End
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuMenuNext 
         Caption         =   "Next &Tab"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuF1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuExit 
         Caption         =   "E&xit"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmCYSCorrection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cNullDate = #12:00:00 AM#
Const cCANFormat = "^  |  RATE  |^ CONTAINER # |^ SZ |REFERENCE|>   AMOUNT  "

Dim clsADR As Object
Dim vUserID As String
Dim nOldCCR, nOldADRNo, curOldADRAmt As Long

Private Sub cmdCANGet_Click()
    Call lzGetCCR2Void
End Sub

Private Sub cmdCANGet_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            SendKeys "+{TAB}"
        Case Else
    End Select
End Sub

Private Sub cmdCANSave_Click()
Dim nCCR As Long
    If MsgBox("CCR number " & Str(nOldCCR) & " will be changed to " & _
              Trim(txtCANCCRNo) & ". Continue?", vbYesNo + vbDefaultButton2) = vbYes Then
        nCCR = CLng("0" & Trim(txtCANCCRNo))
        If lzCCRValid(nCCR) Then
            Call lzSaveCCR(nCCR)
        End If
    End If
End Sub

Private Sub cmdCANSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            SendKeys "+{TAB}"
        Case Else
    End Select
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPayGet_Click()
    Call lzGetPay
End Sub

Private Sub cmdSave_Click()
    If MsgBox("About to change payment information. Continue?", vbYesNo + vbDefaultButton2) = vbYes Then
        Call lzSavePay
    End If
End Sub

Private Sub cmdSave_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            txtCshAmt.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            Call lzSavePay
        Case Else
    End Select
End Sub

Private Sub cmdVoid_Click()
    If MsgBox("CCR number " & Str(nOldCCR) & " will be voided. Continue?", vbYesNo + vbDefaultButton2) = vbYes Then
        Call lzVoidCCR
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    ' connect to ADR module
    Set clsADR = CreateObject("ADR.cADR")
    If Not clsADR.IsConnected Then clsADR.Connect

    'initialize
    Call lzInitialize

End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    clsADR.Disconnect
    Set clsADR = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = (MsgBox("Exit CY Special Services Correction/Voiding?", vbYesNo, "Quit") = vbNo)
End Sub

Private Sub lzInitialize()
    vUserID = gzCurrentUser
    nOldCCR = 0
    Call lzSetTab(1)
End Sub

Private Sub mnuMenuExit_Click()
    Unload Me
End Sub

Private Sub mnuMenuNext_Click()
Dim n, t As Integer
    t = tabMenu.Tabs
    n = tabMenu.Tab + 1
    Call lzSetTab(IIf(n = t, 1, n + 1))
    tabMenu.SetFocus
End Sub

Private Sub tabMenu_GotFocus()
    Select Case tabMenu.Tab
        Case 0
            Call lzClearVoid
        Case 1
            Call lzClearPay
        Case Else
    End Select
End Sub

Private Sub txtADRAmt_GotFocus()
    With txtADRAmt
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txtADRAmt_KeyPress(KeyAscii As Integer)
Dim n As Integer
    Select Case KeyAscii
        Case vbKeyEscape
            n = 5
            While n > 0
                If Trim(txtChkBank(n - 1)) <> "" Then
                    txtChkBank(n - 1).SetFocus
                    KeyAscii = 0
                    Exit Sub
                Else
                    n = n - 1
                End If
            Wend
            txtChkAmt(0).SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            If (CCur("0" & Trim(txtADRAmt)) > 0) Then
                If Trim(txtCusCode) <> "" Then
                    Call lzGetADRAmt
                Else
                    cmdSave.Enabled = False
                    txtCusCode.SetFocus
                End If
            Else
                txtCusCode = Space(txtCusCode.MaxLength)
                txtCustomer.SetFocus
            End If
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtADRAmt_LostFocus()
    txtADRAmt.BackColor = vbWindowBackground
    If (CCur("0" & Trim(txtADRAmt)) = 0) Then txtCusCode = Space(txtCusCode.MaxLength)
    Call lzComputePay
End Sub

Private Sub txtCANCCRNo_Change()
Dim bCCRChanged As Boolean
    bCCRChanged = (CLng("0" & Trim(txtCANCCRNo)) <> nOldCCR) And (nOldCCR <> 0)
    cmdCANSave.Enabled = bCCRChanged
    cmdVoid.Enabled = Not bCCRChanged
End Sub

Private Sub txtCANCCRNo_GotFocus()
    With txtCANCCRNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txtCANCCRNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            txtCANCCRNo = nOldCCR
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtCANCCRNo_LostFocus()
    txtCANCCRNo.BackColor = vbWindowBackground
End Sub

Private Sub txtCANCCRNo_Validate(Cancel As Boolean)
    Cancel = (CLng("0" & Trim(txtCANCCRNo)) <= 0)
End Sub

Private Sub txtCANRefNo_Change()
    cmdVoid.Enabled = False
End Sub

Private Sub txtCANRefNo_GotFocus()
    With txtCANRefNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txtCANRefNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtCANRefNo_LostFocus()
    txtCANRefNo.BackColor = vbWindowBackground
    txtCANRefNo = CLng("0" & Trim(txtCANRefNo))
End Sub

Private Sub txtCANRefNo_Validate(Cancel As Boolean)
    Cancel = (CLng("0" & Trim(txtCANRefNo)) <= 0) And (Len(Trim(txtCANRefNo)) > 0)
End Sub

Private Sub txtCANSeqNo_Change()
    cmdVoid.Enabled = False
End Sub

Private Sub txtCANSeqNo_GotFocus()
    With txtCANSeqNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txtCANSeqNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtCANSeqNo_LostFocus()
    txtCANSeqNo.BackColor = vbWindowBackground
    txtCANSeqNo = CLng("0" & Trim(txtCANSeqNo))
End Sub

Private Sub lzClearVoid()
Dim n As Integer
    txtCANRefNo = ""
    txtCANSeqNo = ""
    txtCANCCRNo = ""
    lblCANCustomer = ""
    grdVoid.Clear
    grdVoid.FormatString = cCANFormat
    lblCANIssueInfo = ""
    lblVoided.Visible = False
    cmdVoid = False
    txtCANRefNo.SetFocus
End Sub

Private Sub lzGetCCR2Void()
Dim wait As New CWaitCursor
Dim rst As ADODB.Recordset
Dim sSQL As String
Dim n As Integer
Dim nAmt, nTotal As Currency
Dim bVoided, bADRPaid As Boolean

    wait.SetCursor
    On Error GoTo err_Get
    Set rst = New ADODB.Recordset
    sSQL = "SELECT A.*, B.cusnam, B.adramt FROM CCRDtl as A"
    sSQL = sSQL & " JOIN CCRPay as B on A.refnum = B.refnum"
    sSQL = sSQL & " WHERE A.refnum = " & Trim(txtCANRefNo)
    sSQL = sSQL & " AND A.seqnum = " & Trim(txtCANSeqNo)
    rst.Open sSQL, gcnnBilling, adOpenStatic, adLockReadOnly, adCmdText
    
    With rst
        If Not .EOF Then
            .MoveFirst
            txtCANCCRNo = "" & Trim(!ccrnum): nOldCCR = CLng("0" & Trim(txtCANCCRNo))
            lblCANCustomer = !cusnam
            lblCANIssueInfo = "Issued by " & Trim(!userid) & " on " & Format(!sysdttm, "YYYY/MM/DD hh:mm")
            bVoided = (!Status = "CAN")
            If bVoided Then
                cmdVoid.Enabled = False
                MsgBox "CCR already voided."
            End If
            bADRPaid = (!adramt > 0)
            If bADRPaid Then
                cmdVoid.Enabled = False
                MsgBox "CCR is paid by ADR. Cannot be voided."
            End If
            lblVoided.Visible = bVoided
            cmdVoid.Enabled = Not (bVoided Or bADRPaid)
            n = 0: nAmt = 0: nTotal = 0
            grdVoid.Clear: grdVoid.FormatString = cCANFormat
            While Not .EOF
                n = n + 1
                grdVoid.TextMatrix(n, 0) = n
                grdVoid.TextMatrix(n, 1) = "" & !chargetyp
                grdVoid.TextMatrix(n, 2) = "" & !cntnum
                grdVoid.TextMatrix(n, 3) = "" & !cntsze
                grdVoid.TextMatrix(n, 4) = "" & !docrefno
                nAmt = !amt + !dgramt + !ovzamt + !vatamt - !wtax
                grdVoid.TextMatrix(n, 5) = Format(nAmt, "#,###,##0.00")
                nTotal = nTotal + nAmt
                .MoveNext
            Wend
            With grdVoid
                .Row = .Rows - 1
                .TextMatrix(9, 4) = "TOTAL"
                .Col = 5: .CellBackColor = vbYellow: .Text = Format(nTotal, "#,###,##0.00")
            End With
            txtCANCCRNo.Enabled = Not bVoided
            If Not bVoided Then
                txtCANCCRNo.SetFocus
            Else
                txtCANRefNo.SetFocus
            End If
        End If
    End With
        
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    On Error GoTo 0
    
    Exit Sub

err_Get:
    MsgBox "Error accessing CCR Table/s ...", vbCritical
    On Error Resume Next
    Set rst = Nothing
    On Error GoTo 0
End Sub

Private Sub txtCANSeqNo_Validate(Cancel As Boolean)
    Cancel = (Len(Trim(txtCANSeqNo)) = 0)
End Sub

Private Sub lzSetTab(ByVal pTab As Integer)
Dim n As Integer
    For n = 1 To tabMenu.Tabs
        tabMenu.TabEnabled(n - 1) = False
    Next
    tabMenu.TabEnabled(pTab - 1) = True
    tabMenu.Tab = pTab - 1
End Sub

Private Sub lzVoidCCR()
Dim wait As New CWaitCursor
Dim rst, rstLog As ADODB.Recordset
Dim sSQL As String
Dim n As Integer
Dim dVoidDate As Date

    ' get system date
    dVoidDate = gzGetSysDate
    
    wait.SetCursor
    On Error GoTo err_Get
    
    ' open detail file
    Set rst = New ADODB.Recordset
    sSQL = "SELECT * FROM CCRDtl"
    sSQL = sSQL & " WHERE refnum = " & Trim(txtCANRefNo)
    sSQL = sSQL & " AND seqnum = " & Trim(txtCANSeqNo)
    rst.Open sSQL, gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
    
    ' open log file
    Set rstLog = New ADODB.Recordset
    rstLog.Open "CCRDtlZ", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
    
    If Not rst.EOF Then
        rst.MoveFirst
        While Not rst.EOF
            
            ' log detail
            rstLog.AddNew
            rstLog!refnum = rst!refnum
            rstLog!seqnum = rst!seqnum
            rstLog!itmnum = rst!itmnum
            rstLog!ccrnum = rst!ccrnum
            rstLog!ccrtyp = rst!ccrtyp
            rstLog!chargetyp = rst!chargetyp
            rstLog!descr = rst!descr
            rstLog!docrefno = rst!docrefno
            rstLog!entnum = rst!entnum
            rstLog!regnum = rst!regnum
            rstLog!cntnum = rst!cntnum
            rstLog!cntsze = rst!cntsze
            rstLog!fulemp = rst!fulemp
            rstLog!amt = rst!amt
            rstLog!vatamt = rst!vatamt
            rstLog!wtax = rst!wtax
            rstLog!vatcde = rst!vatcde
            rstLog!stostat = rst!stostat
            rstLog!lngth = rst!lngth
            rstLog!Width = rst!Width
            rstLog!Height = rst!Height
            rstLog!ums = rst!ums
            rstLog!quantity = rst!quantity
            rstLog!dgrcls = rst!dgrcls
            rstLog!dgramt = rst!dgramt
            rstLog!revton = rst!revton
            rstLog!ovzamt = rst!ovzamt
            rstLog!enrfrdttm = rst!enrfrdttm
            rstLog!enstodttm = rst!enstodttm
            rstLog!stordys = rst!stordys
            rstLog!remark = rst!remark
            rstLog!guarntycde = rst!guarntycde
            rstLog!Status = rst!Status
            rstLog!shplin = rst!shplin
            rstLog!vslcde = rst!vslcde
            rstLog!pod = rst!pod
            rstLog!userid = rst!userid
            rstLog!sysdttm = rst!sysdttm
            rstLog!updcde = rst!updcde
            rstLog!outdttm = rst!outdttm
            rstLog.Update
            
            ' tag as cancelled
            rst!Status = "CAN"
            rst!updcde = "C"
            rst.Update
        
            ' log detail
            rstLog.AddNew
            rstLog!refnum = rst!refnum
            rstLog!seqnum = rst!seqnum
            rstLog!itmnum = rst!itmnum
            rstLog!ccrnum = rst!ccrnum
            rstLog!ccrtyp = rst!ccrtyp
            rstLog!chargetyp = rst!chargetyp
            rstLog!descr = rst!descr
            rstLog!docrefno = rst!docrefno
            rstLog!entnum = rst!entnum
            rstLog!regnum = rst!regnum
            rstLog!cntnum = rst!cntnum
            rstLog!cntsze = rst!cntsze
            rstLog!fulemp = rst!fulemp
            rstLog!amt = rst!amt
            rstLog!vatamt = rst!vatamt
            rstLog!wtax = rst!wtax
            rstLog!vatcde = rst!vatcde
            rstLog!stostat = rst!stostat
            rstLog!lngth = rst!lngth
            rstLog!Width = rst!Width
            rstLog!Height = rst!Height
            rstLog!ums = rst!ums
            rstLog!quantity = rst!quantity
            rstLog!dgrcls = rst!dgrcls
            rstLog!dgramt = rst!dgramt
            rstLog!revton = rst!revton
            rstLog!ovzamt = rst!ovzamt
            rstLog!enrfrdttm = rst!enrfrdttm
            rstLog!enstodttm = rst!enstodttm
            rstLog!stordys = rst!stordys
            rstLog!remark = rst!remark
            rstLog!guarntycde = rst!guarntycde
            rstLog!Status = rst!Status
            rstLog!shplin = rst!shplin
            rstLog!vslcde = rst!vslcde
            rstLog!pod = rst!pod
            rstLog!userid = vUserID
            rstLog!sysdttm = dVoidDate
            rstLog!updcde = rst!updcde
            rstLog!outdttm = rst!outdttm
            rstLog.Update
            
            rst.MoveNext
        Wend
    
        cmdVoid.Enabled = False
        lblVoided.Visible = True
        txtCANRefNo.SetFocus
        
    End If
        
    On Error Resume Next
    rst.Close
    rstLog.Close
    Set rst = Nothing
    Set rstLog = Nothing
    On Error GoTo 0
    
    Exit Sub

err_Get:
    MsgBox "Error accessing CCR Table/s ...", vbCritical
    On Error Resume Next
    Set rst = Nothing
    Set rstLog = Nothing
    On Error GoTo 0
End Sub

Private Sub lzSaveCCR(ByVal pNewCCR As Long)
Dim wait As New CWaitCursor
Dim rst, rstLog As ADODB.Recordset
Dim sSQL As String
Dim n As Integer
Dim dLogDate As Date
    
    ' get system date
    dLogDate = gzGetSysDate
    
    wait.SetCursor
    On Error GoTo err_Get
    
    ' open detail file
    Set rst = New ADODB.Recordset
    sSQL = "SELECT * FROM CCRDtl"
    sSQL = sSQL & " WHERE refnum = " & Trim(txtCANRefNo)
    sSQL = sSQL & " AND seqnum = " & Trim(txtCANSeqNo)
    rst.Open sSQL, gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
    
    ' open log file
    Set rstLog = New ADODB.Recordset
    rstLog.Open "CCRDtlZ", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
    
    If Not rst.EOF Then
        rst.MoveFirst
        While Not rst.EOF
            
            ' log detail
            rstLog.AddNew
            rstLog!refnum = rst!refnum
            rstLog!seqnum = rst!seqnum
            rstLog!itmnum = rst!itmnum
            rstLog!ccrnum = rst!ccrnum
            rstLog!ccrtyp = rst!ccrtyp
            rstLog!chargetyp = rst!chargetyp
            rstLog!descr = rst!descr
            rstLog!docrefno = rst!docrefno
            rstLog!entnum = rst!entnum
            rstLog!regnum = rst!regnum
            rstLog!cntnum = rst!cntnum
            rstLog!cntsze = rst!cntsze
            rstLog!fulemp = rst!fulemp
            rstLog!amt = rst!amt
            rstLog!vatamt = rst!vatamt
            rstLog!wtax = rst!wtax
            rstLog!vatcde = rst!vatcde
            rstLog!stostat = rst!stostat
            rstLog!lngth = rst!lngth
            rstLog!Width = rst!Width
            rstLog!Height = rst!Height
            rstLog!ums = rst!ums
            rstLog!quantity = rst!quantity
            rstLog!dgrcls = rst!dgrcls
            rstLog!dgramt = rst!dgramt
            rstLog!revton = rst!revton
            rstLog!ovzamt = rst!ovzamt
            rstLog!enrfrdttm = rst!enrfrdttm
            rstLog!enstodttm = rst!enstodttm
            rstLog!stordys = rst!stordys
            rstLog!remark = rst!remark
            rstLog!guarntycde = rst!guarntycde
            rstLog!Status = rst!Status
            rstLog!shplin = rst!shplin
            rstLog!vslcde = rst!vslcde
            rstLog!pod = rst!pod
            rstLog!userid = rst!userid
            rstLog!sysdttm = rst!sysdttm
            rstLog!updcde = rst!updcde
            rstLog!outdttm = rst!outdttm
            rstLog.Update
            
            ' change CCR number and tag as updated
            rst!ccrnum = pNewCCR
            rst!updcde = "U"
            rst.Update
        
            ' log detail
            rstLog.AddNew
            rstLog!refnum = rst!refnum
            rstLog!seqnum = rst!seqnum
            rstLog!itmnum = rst!itmnum
            rstLog!ccrnum = rst!ccrnum
            rstLog!ccrtyp = rst!ccrtyp
            rstLog!chargetyp = rst!chargetyp
            rstLog!descr = rst!descr
            rstLog!docrefno = rst!docrefno
            rstLog!entnum = rst!entnum
            rstLog!regnum = rst!regnum
            rstLog!cntnum = rst!cntnum
            rstLog!cntsze = rst!cntsze
            rstLog!fulemp = rst!fulemp
            rstLog!amt = rst!amt
            rstLog!vatamt = rst!vatamt
            rstLog!wtax = rst!wtax
            rstLog!vatcde = rst!vatcde
            rstLog!stostat = rst!stostat
            rstLog!lngth = rst!lngth
            rstLog!Width = rst!Width
            rstLog!Height = rst!Height
            rstLog!ums = rst!ums
            rstLog!quantity = rst!quantity
            rstLog!dgrcls = rst!dgrcls
            rstLog!dgramt = rst!dgramt
            rstLog!revton = rst!revton
            rstLog!ovzamt = rst!ovzamt
            rstLog!enrfrdttm = rst!enrfrdttm
            rstLog!enstodttm = rst!enstodttm
            rstLog!stordys = rst!stordys
            rstLog!remark = rst!remark
            rstLog!guarntycde = rst!guarntycde
            rstLog!Status = rst!Status
            rstLog!shplin = rst!shplin
            rstLog!vslcde = rst!vslcde
            rstLog!pod = rst!pod
            rstLog!userid = vUserID
            rstLog!sysdttm = dLogDate
            rstLog!updcde = rst!updcde
            rstLog!outdttm = rst!outdttm
            rstLog.Update
            
            rst.MoveNext
        Wend
    
        txtCANCCRNo.Enabled = False
        cmdCANSave.Enabled = False
        txtCANRefNo.SetFocus
        
    End If
        
    On Error Resume Next
    rst.Close
    rstLog.Close
    Set rst = Nothing
    Set rstLog = Nothing
    On Error GoTo 0
    
    Exit Sub

err_Get:
    MsgBox "Error accessing CCR Table/s ...", vbCritical
    On Error Resume Next
    Set rst = Nothing
    Set rstLog = Nothing
    On Error GoTo 0
End Sub

Private Function lzCCRValid(ByVal pCCRNo As Long) As Boolean
Dim rst As ADODB.Recordset
Dim sSQL As String
    
    On Error GoTo err_Get
    
    ' check if CCR number exists
    Set rst = New ADODB.Recordset
    sSQL = "SELECT TOP 1 ccrnum FROM CCRDtl"
    sSQL = sSQL & " WHERE ccrnum = " & Trim(Str(pCCRNo))
    rst.Open sSQL, gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
        
    lzCCRValid = rst.EOF
    
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    On Error GoTo 0
    
    Exit Function

err_Get:
    MsgBox "Error accessing CCR Table ...", vbCritical
    On Error Resume Next
    Set rst = Nothing
    On Error GoTo 0
End Function

Private Sub lzClearPay()
Dim n As Integer
    lblAmtDue = ""
    txtCshAmt = ""
    For n = 0 To 4
        txtChkAmt(n) = ""
        txtChkNo(n) = Space(txtChkNo(n).MaxLength)
        txtChkBank(n) = Space(txtChkBank(n).MaxLength)
    Next n
    lblChkTot = ""
    txtCusCode = Space(txtCusCode.MaxLength)
    txtADRAmt = ""
    lblChange = ""
    txtCustomer = Space(txtCustomer.MaxLength)
    cmdSave.Enabled = False
    txtPayRefNo.SetFocus
End Sub

Private Sub lzGetPay()
Dim wait As New CWaitCursor
Dim rst As ADODB.Recordset
Dim sSQL As String
Dim n As Integer
Dim nAmtDue, nTotalChk As Currency

    wait.SetCursor
    'On Error GoTo err_Get
    On Error GoTo 0
    Set rst = New ADODB.Recordset
    sSQL = "SELECT * FROM CCRPay"
    sSQL = sSQL & " WHERE refnum = " & Trim(txtPayRefNo)
    sSQL = sSQL & " AND ccrtyp = '2'"
    rst.Open sSQL, gcnnBilling, adOpenStatic, adLockReadOnly, adCmdText
    
    With rst
        If Not .EOF Then
            .MoveFirst
            
            nAmtDue = !cshamt + !chkamt1 + !chkamt2 + !chkamt3 + !chkamt4 + !chkamt5 + !adramt - !chgamt
            lblAmtDue = Format(nAmtDue, "#,###,##0.00")
            txtCshAmt = IIf(!cshamt > 0, Format(!cshamt, "#,###,##0.00"), "")
            lblChange = IIf(!chgamt > 0, Format(!chgamt, "#,###,##0.00"), "")
            
            txtChkAmt(0) = IIf(!chkamt1 > 0, Format(!chkamt1, "#,###,##0.00"), "")
            txtChkAmt(1) = IIf(!chkamt2 > 0, Format(!chkamt2, "#,###,##0.00"), "")
            txtChkAmt(2) = IIf(!chkamt3 > 0, Format(!chkamt3, "#,###,##0.00"), "")
            txtChkAmt(3) = IIf(!chkamt4 > 0, Format(!chkamt4, "#,###,##0.00"), "")
            txtChkAmt(4) = IIf(!chkamt5 > 0, Format(!chkamt5, "#,###,##0.00"), "")
            txtChkNo(0) = !chkno1
            txtChkNo(1) = !chkno2
            txtChkNo(2) = !chkno3
            txtChkNo(3) = !chkno4
            txtChkNo(4) = !chkno5
            txtChkBank(0) = !chkbnk1
            txtChkBank(1) = !chkbnk2
            txtChkBank(2) = !chkbnk3
            txtChkBank(3) = !chkbnk4
            txtChkBank(4) = !chkbnk5
            nTotalChk = !chkamt1 + !chkamt2 + !chkamt3 + !chkamt4 + !chkamt5
            lblChkTot = Format(nTotalChk, "#,###,##0.00")
            
            txtCusCode = !cuscde
            txtADRAmt = IIf(!adramt > 0, Format(!adramt, "#,###,##0.00"), "")
            curOldADRAmt = !adramt
            nOldADRNo = !adrnum
            txtCustomer = !cusnam
            txtCustomer.Enabled = (curOldADRAmt = 0)
            
            txtCshAmt.SetFocus
        
        End If
    End With
        
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    On Error GoTo 0
    
    Exit Sub

err_Get:
    MsgBox "Error accessing CCR Table/s ...", vbCritical
    On Error Resume Next
    Set rst = Nothing
    On Error GoTo 0
End Sub

Private Sub txtChkAmt_GotFocus(Index As Integer)
    With txtChkAmt(Index)
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtChkAmt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            If Index < 4 Then
                If (Trim(txtChkAmt(Index)) & Trim(txtChkAmt(Index + 1)) = "") Then
                    txtADRAmt.SetFocus
                Else
                    SendKeys ("{TAB}")
                End If
            Else
                SendKeys ("{TAB}")
            End If
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtChkAmt_LostFocus(Index As Integer)
    Call lzComputePay
    txtChkAmt(Index).BackColor = vbWindowBackground
    If (Trim(txtChkAmt(Index)) <> "") And (Trim(txtChkNo(Index)) = "") Then txtChkNo(Index).SetFocus
End Sub

Private Sub txtChkAmt_Validate(Index As Integer, Cancel As Boolean)
Dim n As Integer
Dim curTot As Currency
    Cancel = Not IsNumeric("0" & txtChkAmt(Index))
    If Not Cancel Then
        curTot = 0
        For n = 0 To 4
            curTot = curTot + CCur("0" & txtChkAmt(n))
        Next
        lblChkTot = Format(curTot, "##,###,##0.00")
        If Trim(txtChkAmt(Index)) <> "" Then
            txtChkAmt(Index).BackColor = vbWindowBackground
            txtChkNo(Index).SetFocus
        Else
            txtChkNo(Index) = Space(txtChkNo(Index).MaxLength)
            txtChkBank(Index) = Space(txtChkBank(Index).MaxLength)
        End If
    End If
End Sub

Private Sub txtChkBank_GotFocus(Index As Integer)
    With txtChkBank(Index)
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtChkBank_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtChkBank_LostFocus(Index As Integer)
    txtChkBank(Index).BackColor = vbWindowBackground
End Sub

Private Sub txtChkBank_Validate(Index As Integer, Cancel As Boolean)
    Cancel = (Trim(txtChkBank(Index).Text) = "") And _
             (Trim(txtChkAmt(Index).Text) <> "")
    If Cancel Then
        MsgBox "Check bank code required.", vbExclamation
        With txtChkNo(Index)
            .BackColor = vbInfoBackground
            .SelStart = 0
            .SelLength = .MaxLength
        End With
    Else
        If Trim(txtChkAmt(Index)) <> "" Then
            txtChkNo(Index).BackColor = vbWindowBackground
            txtChkBank(Index).SetFocus
        End If
    End If
End Sub

Private Sub txtChkNo_GotFocus(Index As Integer)
    With txtChkNo(Index)
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtChkNo_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtChkNo_LostFocus(Index As Integer)
    txtChkNo(Index).BackColor = vbWindowBackground
    If (Trim(txtChkNo(Index)) <> "") And (Trim(txtChkBank(Index)) = "") Then txtChkBank(Index).SetFocus
End Sub

Private Sub txtChkNo_Validate(Index As Integer, Cancel As Boolean)
    Cancel = (Trim(txtChkNo(Index).Text) = "") And _
             (Trim(txtChkAmt(Index).Text) <> "")
    If Cancel Then
        MsgBox "Check number required.", vbExclamation
        With txtChkNo(Index)
            .BackColor = vbInfoBackground
            .SelStart = 0
            .SelLength = .MaxLength
        End With
    Else
        If Trim(txtChkAmt(Index)) <> "" Then
            txtChkNo(Index).BackColor = vbWindowBackground
            txtChkBank(Index).SetFocus
        End If
    End If
End Sub

Private Sub txtCshAmt_GotFocus()
    With txtCshAmt
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtCshAmt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtCshAmt_LostFocus()
    txtCshAmt.BackColor = vbWindowBackground
    Call lzComputePay
End Sub

Private Sub txtCshAmt_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtCshAmt)
End Sub

Private Sub txtCusCode_GotFocus()
    With txtCusCode
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtCusCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys "+{TAB}"
            KeyAscii = 0
        Case vbKeyReturn
            If Trim(txtCusCode) <> "" Then
                txtADRAmt.Enabled = True
                Call lzGetADRAmt
            Else
                txtADRAmt = ""
                Call lzComputePay
                If cmdSave.Enabled Then
                    cmdSave.SetFocus
                Else
                    txtCshAmt.SetFocus
                End If
            End If
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtCusCode_LostFocus()
    txtCusCode.BackColor = vbWindowBackground
    If Trim(txtCusCode) <> "" Then
        Call lzGetADRAmt
    Else
        txtADRAmt = ""
        Call lzComputePay
    End If
End Sub

Private Sub txtCusCode_Validate(Cancel As Boolean)
Dim s As String
    With txtCusCode
        .Text = Right("000000" & Trim(.Text), 6)
        If .Text = "000000" Then .Text = Space(6)
    End With
End Sub

Private Sub txtCustomer_GotFocus()
    With txtCustomer
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txtCustomer_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtCustomer_LostFocus()
    txtCustomer.BackColor = vbWindowBackground
    txtCustomer = UCase(Trim(txtCustomer))
End Sub

Private Sub txtPayRefNo_Change()
    cmdSave.Enabled = False
End Sub

Private Sub txtPayRefNo_GotFocus()
    With txtPayRefNo
        .BackColor = vbInfoBackground
        .SelStart = 0
        .SelLength = Len(Trim(.Text))
    End With
End Sub

Private Sub txtPayRefNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtPayRefNo_LostFocus()
    txtPayRefNo.BackColor = vbWindowBackground
    txtPayRefNo = CLng("0" & Trim(txtPayRefNo))
End Sub

Private Sub txtPayRefNo_Validate(Cancel As Boolean)
    Cancel = (CLng("0" & Trim(txtPayRefNo)) <= 0) And (Len(Trim(txtPayRefNo)) > 0)
End Sub

Private Sub lzComputePay()
Dim n As Integer
Dim curChkTot, curChange, curCash As Currency
    
    curChkTot = 0
    For n = 0 To 4
        curChkTot = curChkTot + CCur("0" & txtChkAmt(n))
    Next
    lblChkTot = Format(curChkTot, "##,###,##0.00")
    curCash = CCur("0" & txtCshAmt)
    curChange = curCash + CCur("0" & lblChkTot) + CCur("0" & txtADRAmt) - CCur("0" & lblAmtDue)
    If curChange > 0 Then
        If curCash >= curChange Then
            curCash = curCash - curChange
        Else
            curCash = 0
        End If
    ElseIf curChange < 0 Then
        curCash = curCash + Abs(curChange)
        curChange = 0
    End If
    curChange = curCash + CCur("0" & lblChkTot) + CCur("0" & txtADRAmt) - CCur("0" & lblAmtDue)
    txtCshAmt = IIf(curCash > 0, Format(curCash, "##,###,##0.00"), "")
    lblChange = IIf(curChange = 0, "", Format(curChange, "##,###,##0.00"))
    lblChange.ForeColor = IIf(curChange < 0, vbRed, vbBlue)
    cmdSave.Enabled = (curChange >= 0)
    
End Sub

Private Sub lzGetADRAmt()
Dim c As New CWaitCursor
Dim curAmtDue, curADRAmt, curBalance, curCash As Currency
    
    txtCustomer.Enabled = (curOldADRAmt = 0)
    On Error GoTo err_ADR
    c.SetCursor
    curBalance = clsADR.GetADRBal(txtCusCode)
    If curBalance <= 0 Then
        curBalance = 0
        MsgBox "Customer ADR information not found or balance is zero..."
        txtADRAmt.SetFocus
        Exit Sub
    End If
    c.Restore
    On Error GoTo 0
    
    If Trim(clsADR.ErrorCode) <> "" Then
        MsgBox "Customer ADR information not found or with error..."
        txtCusCode.SetFocus
    Else
        txtCustomer = Left(Trim(clsADR.CustomerName), 40)
        txtCustomer.Enabled = False
        On Error Resume Next
        
        curADRAmt = CCur("0" & txtADRAmt)
        curAmtDue = CCur("0" & lblAmtDue)
        If curADRAmt > curAmtDue Then curADRAmt = curAmtDue
        If curBalance < curADRAmt Then curADRAmt = curBalance
        txtADRAmt = Format(curADRAmt, "#,###,##0.00")
        
        Call lzComputePay
       
        If cmdSave.Enabled Then
            cmdSave.SetFocus
        Else
            txtCshAmt.SetFocus
        End If
    End If
    
    Exit Sub

err_ADR:
    MsgBox "Error accessing ADR information...", vbExclamation
    txtADRAmt = ""
    txtCshAmt.SetFocus
End Sub

Private Sub lzSavePay()
Dim wait As New CWaitCursor
Dim rst, rstLog As ADODB.Recordset
Dim sSQL As String
Dim n As Integer
Dim dLogDate As Date
Dim curADRAmt As Currency
Dim vADRNum As Long

    vADRNum = 0
    curADRAmt = CCur("0" & txtADRAmt)
    If curADRAmt <> curOldADRAmt Then
        With clsADR
            If curOldADRAmt > 0 Then
                If .CancelADR(txtCusCode, nOldADRNo, vUserID, "CYSSR Correction") Then
                    If .ApplyADR(txtCusCode, "CCR", CLng("0" & Trim(txtPayRefNo)), curADRAmt, vUserID, "CYSSR Correction") Then
                        vADRNum = CLng("0" & .ControlNo)
                    Else
                        MsgBox "Error accessing customer ADR information...", vbExclamation
                        txtCshAmt.SetFocus
                        Exit Sub
                    End If
                Else
                    MsgBox "Error accessing customer ADR information...", vbExclamation
                    txtCshAmt.SetFocus
                    Exit Sub
                End If
            Else
                If .ApplyADR(txtCusCode, "CCR", CLng("0" & Trim(txtPayRefNo)), curADRAmt, vUserID, "CYSSR Correction") Then
                    vADRNum = CLng("0" & .ControlNo)
                Else
                    MsgBox "Error accessing customer ADR information...", vbExclamation
                    txtCshAmt.SetFocus
                    Exit Sub
                End If
            End If
        End With
    End If
    
    ' get system date
    dLogDate = gzGetSysDate
    
    wait.SetCursor
    On Error GoTo 0
    'On Error GoTo err_Get
    
    ' open detail file
    Set rst = New ADODB.Recordset
    sSQL = "SELECT * FROM CCRPay"
    sSQL = sSQL & " WHERE refnum = " & Trim(txtPayRefNo)
    sSQL = sSQL & " AND ccrtyp = '2'"
    rst.Open sSQL, gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
    
    ' open log file
    Set rstLog = New ADODB.Recordset
    rstLog.Open "CCRPayZ", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
    
    If Not rst.EOF Then
        rst.MoveFirst
        
        ' log detail
        rstLog.AddNew
        rstLog!refnum = rst!refnum
        rstLog!cuscde = rst!cuscde
        rstLog!cusnam = rst!cusnam
        rstLog!cshamt = rst!cshamt
        rstLog!adramt = rst!adramt
        rstLog!adrnum = rst!adrnum
        rstLog!chgamt = rst!chgamt
        rstLog!chkno1 = rst!chkno1
        rstLog!chkno2 = rst!chkno2
        rstLog!chkno3 = rst!chkno3
        rstLog!chkno4 = rst!chkno3
        rstLog!chkno5 = rst!chkno5
        rstLog!chkamt1 = rst!chkamt1
        rstLog!chkamt2 = rst!chkamt2
        rstLog!chkamt3 = rst!chkamt3
        rstLog!chkamt4 = rst!chkamt4
        rstLog!chkamt5 = rst!chkamt5
        rstLog!chkbnk1 = rst!chkbnk1
        rstLog!chkbnk2 = rst!chkbnk2
        rstLog!chkbnk3 = rst!chkbnk3
        rstLog!chkbnk4 = rst!chkbnk4
        rstLog!chkbnk5 = rst!chkbnk5
        rstLog!Status = rst!Status
        rstLog!rectag = rst!rectag
        rstLog!userid = rst!userid
        rstLog!sysdttm = rst!sysdttm
        rstLog!updcde = rst!updcde
        rstLog!ccrtyp = rst!ccrtyp
        rstLog.Update
        
        ' update record
        rst!cuscde = txtCusCode
        rst!cusnam = Left(Trim(txtCustomer), 40)
        rst!cshamt = CCur("0" & txtCshAmt)
        rst!adramt = CCur("0" & txtADRAmt)
        rst!adrnum = vADRNum
        rst!chgamt = CCur("0" & lblChange)
        rst!chkno1 = txtChkNo(0)
        rst!chkno2 = txtChkNo(1)
        rst!chkno3 = txtChkNo(2)
        rst!chkno3 = txtChkNo(3)
        rst!chkno5 = txtChkNo(4)
        rst!chkamt1 = CCur("0" & txtChkAmt(0))
        rst!chkamt2 = CCur("0" & txtChkAmt(1))
        rst!chkamt3 = CCur("0" & txtChkAmt(2))
        rst!chkamt4 = CCur("0" & txtChkAmt(3))
        rst!chkamt5 = CCur("0" & txtChkAmt(4))
        rst!chkbnk1 = txtChkBank(0)
        rst!chkbnk2 = txtChkBank(1)
        rst!chkbnk3 = txtChkBank(2)
        rst!chkbnk4 = txtChkBank(3)
        rst!chkbnk5 = txtChkBank(4)
        rst!updcde = "U"
        rst.Update
    
        ' log detail
        rstLog.AddNew
        rstLog!refnum = rst!refnum
        rstLog!cuscde = rst!cuscde
        rstLog!cusnam = rst!cusnam
        rstLog!cshamt = rst!cshamt
        rstLog!adramt = rst!adramt
        rstLog!adrnum = rst!adrnum
        rstLog!chgamt = rst!chgamt
        rstLog!chkno1 = rst!chkno1
        rstLog!chkno2 = rst!chkno2
        rstLog!chkno3 = rst!chkno3
        rstLog!chkno4 = rst!chkno3
        rstLog!chkno5 = rst!chkno5
        rstLog!chkamt1 = rst!chkamt1
        rstLog!chkamt2 = rst!chkamt2
        rstLog!chkamt3 = rst!chkamt3
        rstLog!chkamt4 = rst!chkamt4
        rstLog!chkamt5 = rst!chkamt5
        rstLog!chkbnk1 = rst!chkbnk1
        rstLog!chkbnk2 = rst!chkbnk2
        rstLog!chkbnk3 = rst!chkbnk3
        rstLog!chkbnk4 = rst!chkbnk4
        rstLog!chkbnk5 = rst!chkbnk5
        rstLog!Status = rst!Status
        rstLog!rectag = rst!rectag
        rstLog!userid = vUserID
        rstLog!sysdttm = dLogDate
        rstLog!updcde = rst!updcde
        rstLog!ccrtyp = rst!ccrtyp
        rstLog.Update
        
        cmdSave.Enabled = False
        txtPayRefNo.SetFocus
        
    End If
        
    On Error Resume Next
    rst.Close
    rstLog.Close
    Set rst = Nothing
    Set rstLog = Nothing
    On Error GoTo 0
    
    Exit Sub

err_Get:
    MsgBox "Error accessing CCR Table/s ...", vbCritical
    On Error Resume Next
    Set rst = Nothing
    Set rstLog = Nothing
    On Error GoTo 0
End Sub
