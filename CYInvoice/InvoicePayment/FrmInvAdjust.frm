VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmInvAdjust 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Invoice Payment Adjsutment"
   ClientHeight    =   8220
   ClientLeft      =   240
   ClientTop       =   525
   ClientWidth     =   14580
   BeginProperty Font 
      Name            =   "IBM3270 - 1254"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   14580
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Frame fraPaydtls 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Payment Details "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4815
      Left            =   5160
      TabIndex        =   24
      Top             =   2400
      Visible         =   0   'False
      Width           =   8655
      Begin VB.Frame fraORAdj 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Invoice Adjustment"
         ForeColor       =   &H00FF0000&
         Height          =   4335
         Left            =   5280
         TabIndex        =   54
         Top             =   360
         Width           =   3255
         Begin VB.TextBox txtAdj 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtAdj 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtAdj 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   2
            Left            =   1320
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtAdj 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox txtAdj 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   1335
            Index           =   4
            Left            =   1320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1695
         End
         Begin VB.TextBox txtAdj 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   5
            Left            =   1440
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   3720
            Width           =   1695
         End
         Begin VB.Label lbladj 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice #"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   66
            Top             =   840
            Width           =   675
         End
         Begin VB.Label lbladj 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "O.R. #"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   65
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lbladj 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adjustment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   64
            Top             =   1440
            Width           =   780
         End
         Begin VB.Label lbladj 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   63
            Top             =   1920
            Width           =   345
         End
         Begin VB.Label lbladj 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   62
            Top             =   2400
            Width           =   630
         End
         Begin VB.Label lbladj 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Adjustment"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   61
            Top             =   3840
            Width           =   1185
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grid_PaymentDetails 
         Height          =   3735
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ForeColor       =   8388608
         ForeColorFixed  =   65535
         BackColorSel    =   16711680
         ForeColorSel    =   16777215
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame fraORType 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OR Details"
         ForeColor       =   &H00FF0000&
         Height          =   4455
         Left            =   5280
         TabIndex        =   33
         Top             =   240
         Width           =   3255
         Begin VB.TextBox txtordtl 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   0
            Left            =   1320
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtordtl 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   1
            Left            =   1320
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtordtl 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   2
            Left            =   1320
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox txtordtl 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   3
            Left            =   1320
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox txtordtl 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   4
            Left            =   1320
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   2280
            Width           =   1695
         End
         Begin VB.TextBox txtordtl 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   5
            Left            =   1320
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1695
         End
         Begin VB.TextBox txtordtl 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   6
            Left            =   1320
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   3120
            Width           =   1695
         End
         Begin VB.TextBox txtordtl 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   7
            Left            =   1320
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   3480
            Width           =   1695
         End
         Begin VB.TextBox txtordtl 
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   8
            Left            =   1320
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   3960
            Width           =   1695
         End
         Begin VB.Label lblcheck1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check #1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   53
            Top             =   840
            Width           =   705
         End
         Begin VB.Label lblcheck1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   52
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label lblcheck1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   3
            Left            =   840
            TabIndex        =   51
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label lblcheck1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check #2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   50
            Top             =   2040
            Width           =   705
         End
         Begin VB.Label lblcheck1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   49
            Top             =   2400
            Width           =   1050
         End
         Begin VB.Label lblcheck1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   6
            Left            =   840
            TabIndex        =   48
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label lblcheck1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cash Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   47
            Top             =   3240
            Width           =   945
         End
         Begin VB.Label lblornum 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "lblOR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   1320
            TabIndex        =   46
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblcheck1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "O.R. #"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   45
            Top             =   360
            Width           =   480
         End
         Begin VB.Label lblcheck1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Available Credit Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   435
            Index           =   8
            Left            =   240
            TabIndex        =   44
            Top             =   3480
            Width           =   960
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblcheck1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OR  Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   43
            Top             =   4080
            Width           =   870
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "< ESC -Close >"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   5
         Left            =   1800
         TabIndex        =   26
         Top             =   4200
         Width           =   1905
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid_InvAdjust 
      Height          =   2655
      Left            =   240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4200
      Width           =   14085
      _ExtentX        =   24844
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
      ForeColor       =   8388608
      ForeColorFixed  =   65535
      BackColorSel    =   8388608
      ForeColorSel    =   65535
      WordWrap        =   -1  'True
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbCustomer 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   11535
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   5
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10680
      TabIndex        =   7
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Frame FraAdjust 
      Caption         =   "Invoice Adjustment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2655
      Left            =   7440
      TabIndex        =   19
      Top             =   840
      Width           =   5985
      Begin VB.PictureBox picTip 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         Picture         =   "FrmInvAdjust.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   2535
         TabIndex        =   23
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtadjust 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Text            =   "txtadjust"
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtremarks 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   1800
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1440
         Width           =   3855
      End
      Begin VB.TextBox txtinv 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Text            =   "txtinv"
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No. "
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   10
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   1620
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   9
         Left            =   240
         TabIndex        =   21
         Top             =   1680
         Width           =   1410
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustment"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Keys"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   360
      TabIndex        =   18
      Top             =   7440
      Width           =   8415
      Begin VB.Label lblf6 
         AutoSize        =   -1  'True
         Caption         =   "F6-Process Adj."
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   6120
         TabIndex        =   32
         Top             =   240
         Width           =   2040
      End
      Begin VB.Label lblf5 
         AutoSize        =   -1  'True
         Caption         =   "F5-Remove Invoice"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   3600
         TabIndex        =   31
         Top             =   240
         Width           =   2310
      End
      Begin VB.Label lblf4 
         AutoSize        =   -1  'True
         Caption         =   "F4-Add Invoice "
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   1560
         TabIndex        =   30
         Top             =   240
         Width           =   2040
      End
      Begin VB.Label lblF3 
         AutoSize        =   -1  'True
         Caption         =   "F3-Close "
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame FraInvDetails 
      Caption         =   "Invoice Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2655
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   6735
      Begin VB.Label lblTaxAmount 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblTaxAmount"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2880
         TabIndex        =   68
         Top             =   1080
         Width           =   3000
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Amount"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   6
         Left            =   1080
         TabIndex        =   67
         Top             =   1200
         Width           =   1365
      End
      Begin VB.Label lblBalance 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblBalance"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   2040
         Width           =   3000
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Amount"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   4
         Left            =   600
         TabIndex        =   16
         Top             =   2160
         Width           =   1905
      End
      Begin VB.Label lblPayment 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblPayment"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   1440
         Width           =   3000
      End
      Begin VB.Label lblInvAmt 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblInvAmt"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2880
         TabIndex        =   14
         Top             =   720
         Width           =   3000
      End
      Begin VB.Label lblInvNo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblInvNo"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Top             =   360
         Width           =   3000
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Payment/s  Made "
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   3
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Amount    "
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   2
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   1965
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Invoice No. "
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2355
      End
   End
   Begin VB.Label lblTotal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lbltotal"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   4200
      TabIndex        =   28
      Top             =   6960
      Width           =   3000
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount For Adjustment"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   360
      TabIndex        =   27
      Top             =   7080
      Width           =   3660
   End
   Begin VB.Label lblgrid 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "List of Invoice/s for Adjustment"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   3720
      Width           =   14055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer "
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   270
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1230
   End
End
Attribute VB_Name = "FrmInvAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstInv_Update As ADODB.Recordset
Dim rsInvDetail_Adj As ADODB.Recordset

'gcnnBilling

Private Sub cmbCustomer_Click()
   Call clear_Frame_invdetails
   Call clear_FrameAdjustment
   txtinv.Enabled = True
   txtadjust.Enabled = True
   txtremarks.Enabled = True
End Sub

Private Sub cmbCustomer_GotFocus()
   SendKeys "%{down}", False
   DoEvents
End Sub

Private Sub cmbCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
     Call key_precedence(KeyCode, 1)
End Sub
Private Sub cmdClose_Click()
      If MsgBox("Are you sure you want to exit? ", vbYesNo + vbQuestion, "Invoice Payment Adjust") = vbYes Then
                 Unload Me
      End If
End Sub

Private Sub cmdProcess_Click()
  If MsgBox("Are all Entries correct? ", vbYesNo + vbQuestion, "Process Invoice Entry for Adjustmetn") = vbYes Then
        Call AdjustmentProcess(grid_InvAdjust, LTrim(Mid(cmbCustomer.Text, 1, 6)))
        Call grid_InvAdjustHeader
        cmdProcess.Enabled = False
        Call FilterRecordset(frmMain.cmbcust.List(frmMain.cmbcust.ListIndex))
        Call List_UnpaidBills
        cmbCustomer.SetFocus
 End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
   Case vbKeyF3
        If MsgBox("Are you sure you want to exit? ", vbYesNo + vbQuestion, "Invoice Payment Adjust") = vbYes Then
                 Unload Me
        End If
 End Select
End Sub

Private Sub Form_Load()
 Call initialize_settings
 Call grid_InvAdjustHeader
 Call list_customer
 txtinv.Enabled = True
 txtadjust.Enabled = True
 txtremarks.Enabled = True
 FraAdjust.Enabled = True
End Sub

Private Sub grid_PaymentDetails_Header()
  Dim col As Integer
  Dim colhdgs(3) As String
  
  colhdgs(0) = "O.R.No"
  colhdgs(1) = "OR Type"
  colhdgs(2) = "Payment Made"
  colhdgs(3) = "Payment Date"
  With grid_PaymentDetails
        .Rows = 2
        .Cols = 4
        .row = 0
        For col = 0 To 3
            .col = col: .Text = colhdgs(col): .CellAlignment = 4
        Next col
        .ColWidth(0) = 900
        .ColWidth(1) = 955
        .ColWidth(2) = 1600
        .ColWidth(3) = 1600
        
  End With
End Sub

Private Sub grid_InvAdjustHeader()
  Dim col As Integer
  Dim colhdgs(9) As String
  colhdgs(0) = "Invoice#"
  colhdgs(1) = "Inv. Amount"
  colhdgs(2) = "VAT"
  colhdgs(3) = "WTax"
  colhdgs(4) = "Total"
  colhdgs(5) = "Date"
  colhdgs(6) = "Payment"
  colhdgs(7) = "Adjustment"
  colhdgs(8) = " Balance"
  colhdgs(9) = "Remarks"
  
  With grid_InvAdjust
    .Clear
    .Rows = 2
    .Cols = 10
    .row = 0
    .ColWidth(0) = 1000
    .ColWidth(1) = 1500
    .ColWidth(2) = 800
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1415
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500
    .ColWidth(9) = 2000

       For col = 0 To 9
            .col = col: .Text = colhdgs(col): .CellAlignment = 4
       Next col
       .Width = 14100: .Height = 2655: .Left = 240: .Top = 4200
  End With
  picTip.Visible = False
  grid_PaymentDetails.Visible = True
  fraPaydtls.Visible = False
  txtadjust.Enabled = False
  txtremarks.Enabled = False
End Sub

Private Sub clear_Frame_invdetails()
    lblInvNo.Caption = ""
    lblInvAmt.Caption = "0"
    lblPayment.Caption = "0"
    lblBalance.Caption = "0"
    lblTaxAmount.Caption = "0"
End Sub
Private Sub clear_FrameAdjustment()
  txtinv.Text = ""
  txtadjust.Text = "" ' : Enabled = False
  txtremarks.Text = "" ' : Enabled = False
End Sub

Private Sub grid_InvAdjust_KeyDown(KeyCode As Integer, Shift As Integer)
        Call grid_keyIn(KeyCode)
End Sub

Private Sub grid_PaymentDetails_Click()
      Call show_ORdetails(13)
End Sub

Private Sub grid_PaymentDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    Call show_ORdetails(KeyCode)
End Sub

Private Sub grid_PaymentDetails_SelChange()
      Call show_ORdetails(13)
End Sub

Private Sub txtadjust_GotFocus()
  picTip.Visible = True
End Sub
Private Sub txtadjust_KeyDown(KeyCode As Integer, Shift As Integer)
    Call key_precedence(KeyCode, 3)
End Sub

Private Sub txtadjust_LostFocus()
  picTip.Visible = False
  txtadjust.Text = Format(txtadjust, "###,###,###.#0")
End Sub


Private Sub clear_ORdetails()
    Dim X As TextBox
    For Each X In txtordtl
        X.Text = ""
    Next X
    lblornum.Caption = ""
End Sub

Private Sub list_customer()
    Dim rst As New ADODB.Recordset
    rst.Open "customer", gcnnBilling, , , adCmdTable
    With rst
      Do While Not .EOF
       cmbCustomer.AddItem !cuscde & "| " & !cusnam
       .MoveNext
      Loop
      cmbCustomer.Text = cmbCustomer.List(0)
    End With
    rst.Close
    Set rst = Nothing
    
        
End Sub

Private Sub txtinv_KeyDown(KeyCode As Integer, Shift As Integer)
    Call key_precedence(KeyCode, 2)
End Sub


Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Call key_precedence(KeyCode, 4)
End Sub
Private Sub show_InvDetails(ByVal cuscde As String, ByVal invno As Long)
     Dim rst As New ADODB.Recordset
     Dim sSql As String
     sSql = "Select cuscde,invnum, invvat,invtax,invamt,totalpay,status from invict where invnum=" & invno _
            & " And cuscde='" & cuscde & "'"
     
     rst.Open sSql, gcnnBilling, , , adCmdText
     With rst
             If Not .EOF Then ' Found invoice no. and show details
                    Call clear_Frame_invdetails
                    lblInvNo.Caption = !invnum
                    lblInvAmt.Caption = Format(!invamt, "###,###,###.#0")
                    lblPayment.Caption = IIf(IsNull(!totalpay), 0, Format(.Fields("totalpay").Value, "###,###,###.#0"))
                    lblBalance.Caption = Format(CCur(!invamt + !invvat - !invtax) - CCur(lblPayment), "(" & "##,###,###.#0" & ")")
                    lblTaxAmount = Format(CCur(!invtax), "##,###,###.#0")
             End If
            .Close
     End With
    Set rst = Nothing
End Sub

Private Sub List_Payment(ByVal Inv As Long)
  Dim rst As New ADODB.Recordset
  Dim row As Integer
  
  rst.Open "Select invpaydtl.ornum,invpayhdr.ortype,payamt,paydate,invpaydtl.remarks  from invpaydtl,invpayhdr where invnum= " & Inv & " AND invpaydtl.ornum=invpayhdr.ornum", gcnnBilling, , , adCmdText
  
 row = 1
  Do While Not rst.EOF
  With grid_PaymentDetails
      If row > 1 Then
         .AddItem ""
       End If
      
        .row = row
        .TextMatrix(row, 0) = rst!ornum: .CellAlignment = 4
        .TextMatrix(row, 1) = rst!ortype: .CellAlignment = 4
        .TextMatrix(row, 2) = Format(rst!payamt, "###,###,###.#0"): .CellAlignment = 7
        .TextMatrix(row, 3) = Format(rst!paydate, "yyyy/mm/dd"): .CellAlignment = 1
        row = row + 1
  End With
  rst.MoveNext
Loop
rst.Close
Set rst = Nothing
        grid_PaymentDetails.col = 0
        grid_PaymentDetails.row = 1
        grid_PaymentDetails.ColSel = 2
        grid_PaymentDetails.RowSel = 1
        grid_PaymentDetails.SelectionMode = flexSelectionByRow
        grid_PaymentDetails.HighLight = flexHighlightAlways
        SendKeys "{~}", False
        DoEvents
End Sub

Private Sub key_precedence(ByVal keyp As KeyCodeConstants, txtselect As Integer)
 Dim msg As String
 
  Select Case txtselect
    Case 1  ' cmbcustomer
            Select Case keyp
                  Case 13 ' enter key
                        Call clear_Frame_invdetails
                        Call clear_FrameAdjustment
                        Call grid_InvAdjustHeader
                        FraAdjust.Enabled = True
                        txtinv.Enabled = True
                        lblTotal.Caption = 0#
                        txtinv.SetFocus
                  Case 27, vbKeyF3 ' escape key press
                        If MsgBox("Are you sure you want to exit? ", vbYesNo + vbQuestion, "Invoice Payment Adjust") = vbYes Then
                            Unload Me
                        End If
                        'FraAdjust.Enabled = False
                 Case vbKeyF6 And grid_InvAdjust.TextMatrix(1, 0) <> ""
                        grid_keyIn (vbKeyF6)
                 
             End Select
    
    Case 2  ' txt Invoice
            Select Case keyp
              Case 13 And Trim(txtinv) <> ""
                 If InList_INVno(LTrim(txtinv.Text)) = False Then
                    Call show_InvDetails(Mid(cmbCustomer.Text, 1, 6), Val(txtinv.Text))
                    If IsValid_Invoice(Mid(cmbCustomer.Text, 1, 6), Val(txtinv.Text), msg) = True Then
                        txtadjust.Enabled = True
                        SendKeys "{tab}", False
                        DoEvents
                    Else
                        MsgBox msg, vbOKOnly + vbInformation
                        txtinv.SetFocus
                        SendKeys "{HOME}", False
                        DoEvents
                        SendKeys "+{END}"
                        DoEvents
                    End If
                Else: MsgBox " Invoice No. Already in the list ", vbOKOnly + vbInformation, "Already Exist"
                        txtinv.SetFocus
                        SendKeys "{HOME}", False
                        DoEvents
                        SendKeys "+{END}"
                        DoEvents
                        
                End If
             Case 27
                    txtinv.Text = ""
                    cmbCustomer.SetFocus
            Case vbKeyF6 And CSng(lblTotal.Caption) > 0
                        grid_keyIn (vbKeyF6)
            Case 27, vbKeyF3 ' escape key press
                   If MsgBox("Are you sure you want to exit? ", vbYesNo + vbQuestion, "Invoice Payment Adjust") = vbYes Then
                            Unload Me
                   End If
            End Select
    
    Case 3  ' txt Adjustment
           Select Case keyp
            Case vbKeyF7 And Trim(txtinv) <> ""
                  If (lblPayment.Caption) <> 0 Then
                    FraAdjust.Enabled = False
                    grid_InvAdjust.Enabled = False
                    cmdClose.Enabled = False
                    cmdProcess.Enabled = False
                    cmbCustomer.Enabled = False
                    Call clear_ORdetails
                    Call grid_PaymentDetails_Header
                    Call List_Payment(LTrim(str(txtinv.Text)))
                    fraPaydtls.Visible = True
                    grid_PaymentDetails.Enabled = True
                    grid_PaymentDetails.SetFocus
                  Else
                    MsgBox "Invoice no. " & txtinv.Text & " Has no Partial Payment.", vbOKOnly + vbInformation, "No Payment to View"
                  End If
            Case 13 And Trim(txtadjust) <> ""
                    If isValid_adjustAmt(LTrim(txtinv.Text), CSng(txtadjust.Text)) = True Then
                       txtadjust.Text = Format(txtadjust, "###,###,###.#0")
                       txtremarks.Enabled = True
                        txtremarks.SetFocus
                    Else
                       MsgBox "Adjustment Amount Should not be greater than the Balance Amount", vbOKOnly + vbInformation, "Check Amount Entry"
                      SendKeys "{HOME}", False
                      DoEvents
                      SendKeys "+{END}", False
                      DoEvents
                    End If
                    
            Case 27
                    txtadjust.Text = ""
                    txtinv.Enabled = True
                    txtinv.SetFocus
                    SendKeys "{HOME}", False
                    DoEvents
                    SendKeys "+{END}", False
                    DoEvents
            End Select
    Case 4  ' Remarks
        Select Case keyp
          Case 13 And txtremarks.Text <> ""
              Call add_invListAdjustmet(LTrim(txtinv.Text))
              lblTotal.Caption = Format(CSng(lblTotal.Caption) + CSng(txtadjust.Text), "###,###,###.#0")
              lblf5.Enabled = True
              lblf6.Enabled = True
              cmdProcess.Enabled = True
                If MsgBox("Another Invoice Entry for Adjustment? ", vbYesNo + vbQuestion, "Invoice Adjustment ") = vbYes Then
                     Call clear_Frame_invdetails
                     Call clear_FrameAdjustment
                     'txtinv.Enabled = True
                     txtinv.SetFocus
                Else
                     Call clear_Frame_invdetails
                     Call clear_FrameAdjustment
                     With grid_InvAdjust
                        .col = 0
                        .SetFocus
                        .row = 1:
                        .ColSel = 9
                        .RowSel = 1
                        .SelectionMode = flexSelectionByRow
                        .HighLight = flexHighlightAlways
                     End With
                     SendKeys "{HOME}", False
                     DoEvents
                     SendKeys "+{END}", False
                     DoEvents
                End If
         Case 27
                txtremarks.Text = ""
                'txtadjust.Enabled = True
                txtadjust.SetFocus
       End Select
 End Select
    
End Sub

Private Sub show_ORdetails(ByVal keyp As Integer)
    Select Case keyp
        Case 114, 27  'F3 = 114
            FraAdjust.Enabled = True
            grid_InvAdjust.Enabled = True
            fraPaydtls.Visible = False
            txtadjust.Enabled = True: txtadjust.SetFocus
            cmbCustomer.Enabled = True
            cmdClose.Enabled = True
        Case Else
           If Trim(grid_PaymentDetails.TextMatrix(grid_PaymentDetails.RowSel, 1)) = "OR" Then
                ' showing details of particular OR
                  fraORType.Visible = True
                  fraORAdj.Visible = False
                  Call clear_ORdetails
                  Call OR_details(grid_PaymentDetails.TextMatrix(grid_PaymentDetails.RowSel, 0))
            Else
                  fraORType.Visible = False
                  fraORAdj.Visible = True
                  Call clear_ORAdjust
                  Call OR_Adjust(grid_PaymentDetails.TextMatrix(grid_PaymentDetails.RowSel, 0), Trim(txtinv.Text))
                  txtAdj(5).Text = Format((Sum_OR_Adjust(Trim(txtinv.Text))), "###,###,###.#0")
            End If
        End Select
End Sub
Private Sub clear_ORAdjust()
    Dim txt As TextBox
    
    For Each txt In txtAdj
        txt.Text = ""
    Next txt
End Sub


Private Function Sum_OR_Adjust(ByVal invadj As Long) As Single
    Dim rst As New ADODB.Recordset
    rst.Open "select sum(payamt) as TAdjust from vuetadjust where invnum=" & invadj, gcnnBilling, , , adCmdText
    If IsNull(rst!TAdjust) Then
            Sum_OR_Adjust = 0#
    Else
        Sum_OR_Adjust = rst!TAdjust
    End If
    rst.Close
    Set rst = Nothing
End Function


Private Sub add_invListAdjustmet(ByVal Inv As Long)
  Dim rst As New ADODB.Recordset
  Static row As Integer
  Dim balance As Single
  Dim invtotal As Single
  Dim invamt As Single
  Dim invvat As Single
  Dim invtax As Single
  Dim invpay As Single
  
  
  rst.Open "invict", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
  rst.MoveFirst
  rst.Find "invnum=" & Inv, , adSearchForward
  
  If Not rst.EOF Then
     With rst
        If IsNull(.Fields("invamt").Value) = False Then
                    invamt = rst.Fields("invamt").Value
        Else: invamt = 0
        End If
        If IsNull(.Fields("invvat").Value) = False Then
                    invvat = rst.Fields("invvat").Value
        Else: invvat = 0
        End If
        If IsNull(.Fields("invtax").Value) = False Then
                    invtax = rst.Fields("invtax").Value
        Else: invtax = 0
        End If
        If IsNull(.Fields("totalpay").Value) = False Then
                    invpay = .Fields("totalpay").Value
        Else: invpay = 0
        End If
        invtotal = invamt + invvat - invtax
        balance = invtotal - (invpay + CSng(txtadjust.Text))
     End With
     
     With grid_InvAdjust
        If .TextMatrix(.Rows - 1, 0) <> "" Then
                 .AddItem ""
        End If
         .TextMatrix(.Rows - 1, 0) = rst.Fields("invnum").Value
         .TextMatrix(.Rows - 1, 1) = Format(invamt, "###,###,###.#0")
         .TextMatrix(.Rows - 1, 2) = Format(invvat, "###,###.#0")
         .TextMatrix(.Rows - 1, 3) = Format(invtax, "###,###.#0")
         .TextMatrix(.Rows - 1, 4) = Format(invtotal, "###,###.#0")
         .TextMatrix(.Rows - 1, 5) = Format(rst.Fields("invdttm").Value, "yyyy/mm/dd")
         .TextMatrix(.Rows - 1, 6) = Format(invpay, "###,###.#0")
         .TextMatrix(.Rows - 1, 7) = Format(txtadjust.Text, "###,###.#0")
         .TextMatrix(.Rows - 1, 8) = Format(balance, "(" & "###,###.#0" & ")")
         .TextMatrix(.Rows - 1, 9) = UCase(txtremarks.Text)
    End With
  End If
 rst.Close
 Set rst = Nothing
    
End Sub

Private Sub grid_keyIn(ByVal keyp As KeyCodeConstants)
  Select Case keyp
     Case vbKeyF3, 27 ' exit
            If MsgBox("Exit to Invoice Payment Adjustment? ", vbYesNo + vbYesNo, "Exit Confirmation") = vbYes Then
                    Unload Me
            End If
            
     Case vbKeyF4 ' add another invoice
                If MsgBox("Another Invoice Entry for Adjustment? ", vbYesNo + vbQuestion, "Invoice Adjustment ") = vbYes Then
                     Call clear_Frame_invdetails
                     Call clear_FrameAdjustment
                     txtinv.Enabled = True
                     txtinv.SetFocus
                Else 'No
                     Call clear_Frame_invdetails
                     Call clear_FrameAdjustment
                     With grid_InvAdjust
                        .col = 0
                        .row = 1
                        .ColSel = 9
                        .RowSel = 1
                        .SelectionMode = flexSelectionByRow
                        .HighLight = flexHighlightAlways
                        .SetFocus
                     End With
                     SendKeys "{HOME}", False
                     DoEvents
                     SendKeys "+{END}"
                     DoEvents
               End If
                
     Case vbKeyF5  ' Remove invoice
            Call Remove_Row(grid_InvAdjust.RowSel)
            
     Case vbKeyF6  ' Process all
            If MsgBox("Are all Entries correct? ", vbYesNo + vbQuestion, "Process Invoice Entry for Adjustmetn") = vbYes Then
              Call AdjustmentProcess(grid_InvAdjust, LTrim(Mid(cmbCustomer.Text, 1, 6)))
              Call grid_InvAdjustHeader
              Call FilterRecordset(frmMain.cmbcust.List(frmMain.cmbcust.ListIndex))
              Call List_UnpaidBills
              cmdProcess.Enabled = False
              cmbCustomer.SetFocus
            End If
    End Select
    
End Sub

Private Sub Remove_Row(row As Integer)
  lblTotal.Caption = Format(CSng(lblTotal.Caption) - CSng(grid_InvAdjust.TextMatrix(grid_InvAdjust.RowSel, 7)), "###,###,###.#0")
  If CSng(lblTotal.Caption) = 0 Then
     lblf5.Enabled = False
     lblf6.Enabled = False
     cmdProcess.Enabled = False
  End If
  
  If grid_InvAdjust.Rows > 2 Then
     grid_InvAdjust.RemoveItem (row)
  Else  ' fixed no of row reached
      grid_InvAdjust.Clear
      Call grid_InvAdjustHeader
      txtinv.Enabled = True
      txtinv.SetFocus
  End If
  If CSng(lblTotal.Caption) > 0 Then
    cmdProcess.Enabled = True
  Else: cmdProcess.Enabled = False
  End If
End Sub
Public Function InList_INVno(Inv As Long) As Boolean
  Dim row As Integer
 If Val(grid_InvAdjust.Rows - 1) >= 1 And (grid_InvAdjust.TextMatrix(1, 0) <> "") Then
    For row = 1 To (grid_InvAdjust.Rows - 1)
        If Val(grid_InvAdjust.TextMatrix(row, 0)) = Inv Then
            InList_INVno = True
            Exit For
        End If
    Next row
End If
End Function

Private Function isValid_adjustAmt(ByVal invnum As Long, ByVal AdjustAmt As Single) As Boolean
   Dim rst As New ADODB.Recordset
   Dim str As String
     str = "Select  invnum,cuscde,cusnam,invamt,isnull(invtax,0),isnull(invvat,0),isnull(totalpay,0),invdttm,status,(invamt + isnull(invvat,0) - isnull(invtax,0)) as invtotal," _
        & " ((invamt+ isnull(invvat,0) - isnull(invtax,0))- isnull(totalPay,0)) as Balance From invict " _
        & " where invnum=" & invnum
     
     rst.Open str, gcnnBilling, , , adCmdText
     If AdjustAmt > rst!balance Then
         isValid_adjustAmt = False
     Else
        isValid_adjustAmt = True
    End If
End Function
Private Function HasPartialPayment(ByVal Inv As Long) As Boolean
    Dim rst As New ADODB.Recordset
    
    rst.Open "invpaydtl", gcnnBilling, adOpenDynamic, , adCmdTable
    
    With rst
       .Find "invnum=" & Inv, , adSearchForward
       If Not .EOF Then
            HasPartialPayment = True
        Else
            HasPartialPayment = False
       End If
    End With
    
    rst.Close
    Set rst = Nothing
End Function
Private Sub initialize_settings()
With lblgrid
    .Height = 495
    .Width = 14055
    .Top = 3720
    .Left = 240
 End With
 With FraInvDetails
    .Height = 2655
    .Width = 6735
    .Top = 840
    .Left = 360
 End With
 
 With FraAdjust
    .Height = 2655
    .Width = 5985
    .Top = 840
    .Left = 7440
 End With
 
 With FrmInvAdjust
        .Enabled = True
        .Width = 14700
        .Height = 8600
        .Top = 1500
        .Left = 1500
End With
cmdProcess.Enabled = False
cmdClose.Enabled = True
lblTotal.Caption = "0"
lblf5.Enabled = False
lblf6.Enabled = False
End Sub

' ===================================Invoice Payment Adjustment======================
'         Procedures and Function Calling
Private Function IsValid_Invoice(ByVal cuscde As String, ByVal Inv As Long, ByRef msg1 As String) As Boolean
  Dim rst As New ADODB.Recordset
  Dim str1 As String
  Dim total As Single
    
    str1 = "Select  invnum,cuscde,cusnam,invamt,invvat,invtax,isnull(invtax,0),isnull(invvat,0),isnull(totalpay,0),invdttm,status,(invamt + isnull(invvat,0) - isnull(invtax,0)) as invtotal," _
        & " ((invamt+ isnull(invvat,0) - isnull(invtax,0))- isnull(totalPay,0)) as Balance,totalpay  From invict " _
        & " Where LTRIM(cuscde)=" & cuscde _
        & " AND invnum= " & Inv
    
    rst.Open str1, gcnnBilling, , , adCmdText
    
    If Not rst.EOF Then
        total = rst!invamt + rst!invvat - rst!invtax
    End If
   
    If rst.EOF Then
           msg1 = "Invoice no. " & Inv & " Not found "
            IsValid_Invoice = False
    
    ElseIf (total - rst!totalpay) <= 0 Then
            msg1 = "Invoice No. " & Inv & " is Fully Paid Already"
            IsValid_Invoice = False
            
    ElseIf UCase(Trim(rst!Status)) = "CAN" Or IsNull(rst!Status) = False Then
            msg1 = "This is a Cancelled Invoice"
            IsValid_Invoice = False
    Else
            IsValid_Invoice = True
    End If
    rst.Close
    Set rst = Nothing
End Function

Private Sub OR_Adjust(ByVal orno As String, ByVal invadj As Long)
  Dim rst As New ADODB.Recordset
  
  rst.Open "select * from vuetadjust where invnum=" & invadj & " AND ornum=" & orno, gcnnBilling, , , adCmdText
  
   With FrmInvAdjust
    .txtAdj(0).Text = rst!ornum
    .txtAdj(1).Text = rst!invnum
    .txtAdj(2).Text = Format(rst!payamt, "###,###,###.#0")
    .txtAdj(3).Text = Format(rst!paydate, "yyyy/mm/dd")
    .txtAdj(4).Text = UCase(rst!Remarks)
   End With
    rst.Close
    Set rst = Nothing
End Sub


Private Sub OR_details(ByVal ornum As String)
  Dim rst As New ADODB.Recordset
  Dim str As String
  
  str = "select * From invpayhdr where ornum=" & ornum
  
  rst.Open str, gcnnBilling, , , adCmdText
  
  With FrmInvAdjust
     .lblornum.Caption = ornum
        If IsNull(rst!CheckNo1) = False Then
            .txtordtl(0).Text = rst!CheckNo1
        End If
        If IsNull(rst!CheckAMT1) = False Then
            .txtordtl(1).Text = Format(rst!CheckAMT1, "###,###,###.#0")
        End If
        If IsNull(rst!CheckBnk1) = False Then
            .txtordtl(2).Text = rst!CheckBnk1
        End If
        If IsNull(rst!CheckNo2) = False Then
            .txtordtl(3).Text = rst!CheckNo2
        End If
        If IsNull(rst!CheckAMT2) = False Then
            .txtordtl(4).Text = Format(rst!CheckAMT2, "###,###,###.#0")
        End If
        If IsNull(rst!CheckBnk2) = False Then
            .txtordtl(5).Text = rst!CheckBnk2
        End If
        If IsNull(rst!CashAMT) = False Then
            .txtordtl(6).Text = Format(rst!CashAMT, "###,###,###.#0")
        End If
        If IsNull(rst!AvailAMT) = False Then
            .txtordtl(7).Text = Format(rst!AvailAMT, "###,###,###.#0")
        End If
        If IsNull(rst!TotalAmt) = False Then
            .txtordtl(8).Text = Format(rst!TotalAmt, "###,###,###.#0")
        End If
       
  End With
  
  rst.Close
  Set rst = Nothing
End Sub

Private Sub AdjustmentProcess(grdList As Object, ByVal cuscde As String)
  Dim rstInv As New ADODB.Recordset
  Dim row As Integer
  Dim pInv As String
  Dim pInvAmt As Currency
  Dim pAdjustAmt As Currency
  Dim RBalance As Currency
  Dim pRemark As String
  Dim pORnum As Long
  Dim paydate As Date
  Dim TotalAmt As Currency
    ' Adjustment amount will be used as the Pay Amount
  paydate = Now()
  pORnum = GetNewORno()
  TotalAmt = 0#
  Screen.MousePointer = vbHourglass
  Set rstInv_Update = New ADODB.Recordset
  Set rsInvDetail_Adj = New ADODB.Recordset
  
  rstInv_Update.Open "Invict", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
  rsInvDetail_Adj.Open "Invpaydtl", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable

  With grdList
     For row = 1 To (.Rows - 1)
            pInv = Val(LTrim(.TextMatrix(row, 0))) ' invum
            pInvAmt = CCur(.TextMatrix(row, 1)) ' Inv Amount alone
            pAdjustAmt = CCur(.TextMatrix(row, 7)) ' Adjustment Amount
            RBalance = Abs(CCur(.TextMatrix(row, 8)))  'Running Balance
            pRemark = UCase(.TextMatrix(row, 9))
            TotalAmt = TotalAmt + pAdjustAmt
            Call Update_Invict(pInv, pAdjustAmt)
            Call Save_InvDetail(pORnum, pInv, pInvAmt, pAdjustAmt, pRemark, paydate, RBalance)
     Next row
       Call Save_InvHeader(pORnum, "ADJ", cuscde, TotalAmt, paydate)
       rstInv_Update.Close
       rsInvDetail_Adj.Close
       Set rstInv_Update = Nothing
       Set rsInvDetail_Adj = Nothing
  End With
  FrmInvAdjust.lblTotal.Caption = "0"
  FrmInvAdjust.cmdProcess.Enabled = False
  Screen.MousePointer = vbDefault
End Sub


Public Sub Update_Invict(ByVal Inv As String, ByVal AdjAmt As Single)
 On Error GoTo err_msg
    With rstInv_Update
        .Find "invnum=" & Inv, , adSearchForward, 1
        If Not .EOF Then  'found
             If IsNull(.Fields("totalpay").Value) Then
                    .Fields("totalpay").Value = AdjAmt
             Else
                .Fields("totalpay").Value = .Fields("totalpay").Value + AdjAmt
            End If
            .Update
        End If
    End With
    Exit Sub
err_msg:
  
  MsgBox "Error occur while processing Invoice Adjustment " & Chr(13) & "Contact MIS for Assistance " & Err.Description _
        , vbOKOnly + vbCritical, "Error"
  
End Sub

Private Sub Save_InvDetail(ByVal ornum As Long, Inv As String, ByVal amt As Single, ByVal AdjustAmt As Single, ByVal Remark As String, ByVal dte As Date, ByVal pBalance@)
On Error GoTo err_msg
   With rsInvDetail_Adj
        .AddNew
        .Fields("ornum") = ornum
        .Fields("invnum") = Inv
        .Fields("payamt") = AdjustAmt
        .Fields("Rbalance") = pBalance
        .Fields("remarks") = "Adjusment-" & Space(2) & UCase(Trim(Remark))
        .Fields("paydate") = dte
        .Fields("invamt") = amt
        .Update
    End With
    Exit Sub
err_msg:
 MsgBox "Error occur while processing Invoice Adjustment(Detail Section) " & Chr(13) & "Contact MIS for Assistance " & Err.Description _
        , vbOKOnly + vbCritical, "Error"
  
End Sub

'OR TYPE  ADJ = Adjustment
Private Sub Save_InvHeader(ByVal ornum As Long, ByVal ortype As String, ByVal cde As String, ByVal AdjustAmt As Single, ByVal dte As Date)
   Dim rst As New ADODB.Recordset
  
  On Error GoTo err_msg
  rst.Open "Invpayhdr", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
  With rst
        .AddNew
        .Fields("ornum") = ornum
        .Fields("cuscde") = cde
        .Fields("ortype") = ortype
        .Fields("totalamt") = AdjustAmt
        .Fields("Availamt") = 0
        .Fields("ordate") = dte
        .Fields("userid") = zCurrentUser
        .Fields("CheckAMT1") = 0
        .Fields("CheckNo1") = 0
        .Fields("CheckBnk1") = ""
        .Fields("CheckAMT2") = 0
        .Fields("CheckNo2") = 0
        .Fields("CheckBnk2") = ""
        .Update
        .Close
  End With
  Set rst = Nothing
  Exit Sub
err_msg:
 MsgBox "Error occur while processing Invoice Adjustment(Header Section) " & Chr(13) & "Contact MIS for Assistance " & Err.Description _
        , vbOKOnly + vbCritical, "Error"
End Sub


Public Function GetNewORno() As Long
Dim cmdGetORNum As New ADODB.Command
Dim prmGetORNum As New ADODB.Parameter

With cmdGetORNum
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_ORControlNo"
        .CommandType = adCmdStoredProc
        .Parameters(0) = adParamReturnValue
        .Execute
        GetNewORno = .Parameters(0)
End With

End Function

