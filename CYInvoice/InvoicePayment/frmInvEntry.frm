VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInvEntry 
   Caption         =   "Invoice Payment"
   ClientHeight    =   11235
   ClientLeft      =   2910
   ClientTop       =   2100
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11235
   ScaleWidth      =   11880
   Begin VB.Frame fraPayment 
      Caption         =   "Payment Entry"
      Height          =   8055
      Left            =   4320
      TabIndex        =   24
      Top             =   1560
      Width           =   6975
      Begin MSComCtl2.UpDown updIncrement 
         Height          =   495
         Left            =   6181
         TabIndex        =   67
         Top             =   6360
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   873
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtPtax"
         BuddyDispid     =   196650
         OrigLeft        =   6360
         OrigTop         =   6360
         OrigRight       =   6555
         OrigBottom      =   6855
         Max             =   100
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkApplyCredit 
         Caption         =   "Credit Amount <F1>"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4320
         Width           =   3135
      End
      Begin VB.CommandButton cmdShowEntry 
         Caption         =   "Invoice Entry<F2>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1560
         Picture         =   "frmInvEntry.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   6960
         Width           =   2055
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Close <F3>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         Picture         =   "frmInvEntry.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6960
         Width           =   2055
      End
      Begin MSMask.MaskEdBox txtCheckNo1 
         Height          =   495
         Left            =   1560
         TabIndex        =   20
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
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
         Mask            =   "9999999999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBank1 
         Height          =   495
         Left            =   1560
         TabIndex        =   21
         Top             =   1680
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
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
         Mask            =   ">??????????"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCheckNo2 
         Height          =   495
         Left            =   4200
         TabIndex        =   6
         Top             =   1200
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
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
         Mask            =   "9999999999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtBank2 
         Height          =   495
         Left            =   4200
         TabIndex        =   7
         Top             =   1680
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
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
         Mask            =   ">??????????"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCash 
         Height          =   495
         Left            =   3480
         TabIndex        =   8
         Top             =   3240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,###,###.##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPtax 
         Height          =   495
         Left            =   3600
         TabIndex        =   9
         Top             =   6360
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###"
         Mask            =   "999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtChkAmount1 
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,###,###.##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtChkAmount2 
         Height          =   495
         Left            =   4200
         TabIndex        =   5
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,###,###.##0"
         PromptChar      =   "_"
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage for WTax "
         Height          =   315
         Index           =   18
         Left            =   315
         TabIndex        =   68
         Top             =   6360
         Width           =   3315
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "With Holding Tax"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   240
         TabIndex        =   66
         Top             =   5880
         Width           =   6495
      End
      Begin VB.Label lblAvailCredit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3480
         TabIndex        =   65
         Top             =   4200
         Width           =   3315
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Apply Credit Amount "
         Height          =   315
         Index           =   14
         Left            =   240
         TabIndex        =   64
         Top             =   3840
         Width           =   6495
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Check Amount"
         Height          =   315
         Index           =   8
         Left            =   360
         TabIndex        =   41
         Top             =   2280
         Width           =   2985
      End
      Begin VB.Label lblTchkAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblTchkAmt"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3480
         TabIndex        =   40
         Top             =   2280
         Width           =   3240
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "TOTAL AMOUNT"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   240
         TabIndex        =   39
         Top             =   4800
         Width           =   6495
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblTotal"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   31
         Top             =   5160
         Width           =   6480
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   315
         Index           =   6
         Left            =   2280
         TabIndex        =   30
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Cash Amount"
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   2880
         Width           =   6495
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
         Height          =   315
         Index           =   5
         Left            =   720
         TabIndex        =   28
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Check#"
         Height          =   315
         Index           =   4
         Left            =   360
         TabIndex        =   27
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   315
         Index           =   3
         Left            =   480
         TabIndex        =   26
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Check "
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Payments"
      Enabled         =   0   'False
      Height          =   3615
      Left            =   5400
      TabIndex        =   32
      Top             =   0
      Width           =   4815
      Begin VB.CheckBox chkCredit 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Used Amount"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   19
         Left            =   240
         TabIndex        =   99
         Top             =   3000
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label lblUsedAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2160
         TabIndex        =   98
         Top             =   3000
         Visible         =   0   'False
         Width           =   2520
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Amount"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   2025
      End
      Begin VB.Label lblTCash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   62
         Top             =   840
         Width           =   2235
      End
      Begin VB.Label lblTChkamount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   61
         Top             =   480
         Width           =   2235
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Amount"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   360
         TabIndex        =   60
         Top             =   480
         Width           =   1875
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Total Amount "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   10
         Left            =   240
         TabIndex        =   57
         Top             =   2040
         Width           =   1680
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Amount"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   17
         Left            =   360
         TabIndex        =   38
         Top             =   840
         Width           =   1875
      End
      Begin VB.Label lblTamount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2160
         TabIndex        =   36
         Top             =   2040
         Width           =   2520
      End
      Begin VB.Label lblCreditAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   35
         Top             =   1200
         Width           =   2235
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         DrawMode        =   5  'Not Copy Pen
         Index           =   4
         X1              =   4560
         X2              =   240
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Avail. Balance "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   13
         Left            =   240
         TabIndex        =   34
         Top             =   2520
         Width           =   1680
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2160
         TabIndex        =   33
         Top             =   2520
         Width           =   2520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   5
         X1              =   240
         X2              =   4560
         Y1              =   1800
         Y2              =   1800
      End
   End
   Begin VB.Frame fraContest 
      Caption         =   "Contest Invoice"
      Height          =   3615
      Left            =   10320
      TabIndex        =   81
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmdContest 
         Caption         =   "Cancel<ESC>"
         Height          =   495
         Index           =   1
         Left            =   2400
         TabIndex        =   91
         Top             =   3000
         Width           =   2175
      End
      Begin VB.CommandButton cmdContest 
         Caption         =   "OK"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   90
         Top             =   3000
         Width           =   2175
      End
      Begin MSMask.MaskEdBox txtPayContest 
         Height          =   495
         Index           =   0
         Left            =   2280
         TabIndex        =   96
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,###,###.##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPayContest 
         Height          =   495
         Index           =   1
         Left            =   2280
         TabIndex        =   97
         Top             =   1560
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,###,###.##0"
         PromptChar      =   "_"
      End
      Begin VB.Label lblContestInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   89
         Top             =   2520
         Width           =   4515
      End
      Begin VB.Label lblContestInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblContestInfo"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   88
         Top             =   720
         Width           =   2235
      End
      Begin VB.Label lblContestInfo 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblContestInfo"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   87
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label lblContest 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   86
         Top             =   2160
         Width           =   4545
      End
      Begin VB.Label lblContest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "W/Tax Amount"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   85
         Top             =   1680
         Width           =   1980
      End
      Begin VB.Label lblContest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Amount"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   480
         TabIndex        =   84
         Top             =   1200
         Width           =   1650
      End
      Begin VB.Label lblContest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " Amount"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   83
         Top             =   720
         Width           =   1635
      End
      Begin VB.Label lblContest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice#"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   82
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame fraOR 
      Height          =   3615
      Left            =   240
      TabIndex        =   53
      Top             =   0
      Width           =   5055
      Begin MSMask.MaskEdBox txtCustomerCode 
         Height          =   495
         Left            =   2520
         TabIndex        =   0
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "999999"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK "
         Height          =   975
         Left            =   480
         TabIndex        =   2
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close <F3>"
         Height          =   975
         Left            =   2520
         TabIndex        =   3
         Top             =   2280
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtORNum 
         Height          =   495
         Left            =   2520
         TabIndex        =   1
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#####"
         Mask            =   "#9999"
         PromptChar      =   "_"
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OR #"
         Height          =   315
         Index           =   9
         Left            =   1320
         TabIndex        =   56
         Top             =   840
         Width           =   660
      End
      Begin VB.Label lblInvEntry 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   2295
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCusName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "lblCusName"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   54
         Top             =   1440
         Width           =   4695
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraInvoice 
      Height          =   7575
      Left            =   240
      TabIndex        =   42
      Top             =   3600
      Width           =   14895
      Begin MSMask.MaskEdBox txtPayment 
         Height          =   375
         Left            =   11880
         TabIndex        =   14
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "##,###,###.##0"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTax 
         Height          =   375
         Left            =   10200
         TabIndex        =   13
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###"
         Mask            =   "99999"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtInvNum 
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483633
         PromptInclude   =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99999"
         PromptChar      =   "_"
      End
      Begin VB.Frame Frame1 
         Caption         =   "Legend"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1215
         Left            =   9840
         TabIndex        =   70
         Top             =   6240
         Width           =   4815
         Begin VB.Label lblLegend 
            BackStyle       =   0  'Transparent
            Caption         =   "Show List of Payments"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   11
            Left            =   840
            TabIndex        =   95
            Top             =   960
            Width           =   3855
         End
         Begin VB.Label lblLegend 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F12-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   10
            Left            =   120
            TabIndex        =   94
            Top             =   960
            Width           =   555
         End
         Begin VB.Label lblLegend 
            BackStyle       =   0  'Transparent
            Caption         =   "Show Contest Invoice Entry"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   9
            Left            =   840
            TabIndex        =   80
            Top             =   720
            Width           =   3855
         End
         Begin VB.Label lblLegend 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F10-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   8
            Left            =   120
            TabIndex        =   79
            Top             =   720
            Width           =   555
         End
         Begin VB.Label lblLegend 
            BackStyle       =   0  'Transparent
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   7
            Left            =   3720
            TabIndex        =   78
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lblLegend 
            BackStyle       =   0  'Transparent
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   6
            Left            =   3720
            TabIndex        =   77
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblLegend 
            BackStyle       =   0  'Transparent
            Caption         =   "Full Payment"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   5
            Left            =   840
            TabIndex        =   76
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblLegend 
            BackStyle       =   0  'Transparent
            Caption         =   "Partial Payment"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Index           =   4
            Left            =   840
            TabIndex        =   75
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label lblLegend 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   3
            Left            =   3240
            TabIndex        =   74
            Top             =   480
            Width           =   285
         End
         Begin VB.Label lblLegend 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   2
            Left            =   3240
            TabIndex        =   73
            Top             =   240
            Width           =   285
         End
         Begin VB.Label lblLegend 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "F  -"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   72
            Top             =   480
            Width           =   555
         End
         Begin VB.Label lblLegend 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P  -"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   71
            Top             =   240
            Width           =   555
         End
      End
      Begin VB.CommandButton cmdCancelInvoice 
         Caption         =   "Close <F3>"
         Height          =   855
         Left            =   7200
         TabIndex        =   19
         Top             =   6480
         Width           =   2175
      End
      Begin VB.CommandButton cmdProcess 
         Caption         =   "Process <F7>"
         Height          =   855
         Left            =   4920
         TabIndex        =   18
         Top             =   6480
         Width           =   2175
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove <F6>"
         Height          =   855
         Left            =   2640
         TabIndex        =   17
         Top             =   6480
         Width           =   2175
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add <F5>"
         Height          =   855
         Left            =   360
         TabIndex        =   16
         Top             =   6480
         Width           =   2175
      End
      Begin MSFlexGridLib.MSFlexGrid grdInvoice 
         Height          =   4575
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   8070
         _Version        =   393216
         Rows            =   1
         Cols            =   11
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   12632256
         BackColorSel    =   65535
         ForeColorSel    =   0
         GridColor       =   8421504
         WordWrap        =   -1  'True
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblTTaxAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11880
         TabIndex        =   93
         Top             =   5640
         Width           =   2640
      End
      Begin VB.Label lblgrid 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Invoice WTax Amount"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   11
         Left            =   13080
         TabIndex        =   92
         Top             =   240
         Width           =   1665
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblgrid 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contes- ted"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   10
         Left            =   12240
         TabIndex        =   69
         Top             =   240
         Width           =   825
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInvEntry 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   315
         Index           =   12
         Left            =   6240
         TabIndex        =   63
         Top             =   5760
         Width           =   825
      End
      Begin VB.Label lblTotalInvAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         TabIndex        =   59
         Top             =   5640
         Width           =   2280
      End
      Begin VB.Label lblAppliedAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9720
         TabIndex        =   58
         Top             =   5640
         Width           =   2160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         DrawMode        =   5  'Not Copy Pen
         Index           =   0
         X1              =   14640
         X2              =   120
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Label lblgrid 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inv Stat"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   9
         Left            =   11400
         TabIndex        =   52
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblgrid 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Index           =   8
         Left            =   9480
         TabIndex        =   51
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label lblgrid 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Balance   (Less W/Tax)"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   7
         Left            =   7560
         TabIndex        =   50
         Top             =   240
         Width           =   1920
      End
      Begin VB.Label lblgrid 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "WTax (%)"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Index           =   6
         Left            =   6840
         TabIndex        =   49
         Top             =   240
         Width           =   705
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblgrid 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Balance"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   5
         Left            =   5040
         TabIndex        =   48
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblgrid 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Partial Payment"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   4
         Left            =   3120
         TabIndex        =   47
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label lblgrid 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Index           =   3
         Left            =   3240
         TabIndex        =   46
         Top             =   240
         Width           =   15
      End
      Begin VB.Label lblgrid 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VAT"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   2
         Left            =   3240
         TabIndex        =   45
         Top             =   240
         Width           =   15
      End
      Begin VB.Label lblgrid 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Index           =   1
         Left            =   1200
         TabIndex        =   44
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label lblgrid 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inv #"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Index           =   0
         Left            =   165
         TabIndex        =   43
         Top             =   240
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000012&
         BorderWidth     =   2
         DrawMode        =   5  'Not Copy Pen
         Index           =   1
         X1              =   14640
         X2              =   120
         Y1              =   6240
         Y2              =   6240
      End
   End
End
Attribute VB_Name = "frmInvEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*-----------------------------------------------------------------------
'*Revision: Allow saving of existing OR with balance by BGR,Feb. 8, 2006
'*-----------------------------------------------------------------------

Option Explicit
Const Fill_upMsg = 1 ' fill up all entries
Const Invalid = 2  ' Invalid Input
Const No_Amount = 3
Const Invalid_Invoice = 4

Const Next_Row = 1 'next row in the grid
Const Prev_Row = 0

Const cInvNum = 0        'Invoice#
Const cInvAmount = 1     ' Invoice Amount
Const cVAT = 2           ' VAT
Const cAmtWVAT = 3       ' Invoice Amt + VAT amt
Const cPPayment = 4      ' Partial Payment Amt
Const cBalance = 5       ' Balance
Const cTax = 6           ' Percentage WTax
Const cBalLessTax = 7    ' Invoice Balance Less Tax (Amount+VAT-TaxAmt)
Const cPayAmount = 8     ' Payment
Const cPayStatus = 9     ' Payment Status (F=Full/P=Partial)
Const cConStat = 10      ' Contested Invoice Status (Y/N)
Const cTaxAmount = 11    ' computed Tax Amount

Dim rsCust_Inv As ADODB.Recordset
Dim grdRows As Integer
Dim tmpPayment  As Currency

Dim blnOR_Valid As Boolean 'Boolean to identify OR with balance
Dim sngAvailAmt As Single 'Store remaining amount for the existing OR
Dim sngUsedAmt As Single 'Used up amount of OR

Private Sub chkApplyCredit_Click()
  If chkApplyCredit.Value = 1 Then  'Check
    lblTotal = FVal(CCur(lblTchkAmt) + CCur(Val(txtCash)) + CCur(lblAvailCredit))
  Else
    lblTotal = FVal(CCur(lblTchkAmt) + CCur(Val(txtCash)))
  End If
End Sub

Private Sub chkApplyCredit_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case 13
     txtCash = FVal(txtCash)
     lblTotal = FVal(CDbl(lblTchkAmt) + txtCash)
     Case vbKeyF2 'Show Invoice Entry
            Call cmdShowEntry_Click
     Case vbKeyF3
            Call cmdCancel_Click
     Case vbKeyF1
          Call chkApplyCredit_Click
    End Select
End Sub

Private Sub chkApplyCredit_LostFocus()
 If chkApplyCredit.Value = 1 And CCur(lblAvailCredit) > 0 Then 'Check
    lblTotal = FVal(CCur(lblTchkAmt) + CCur(txtCash) + CCur(lblAvailCredit))
 Else
    lblTotal = FVal(CCur(lblTchkAmt) + CCur(txtCash))
  End If
  lblCreditAmount = CCur(lblAvailCredit)
  chkCredit.Value = chkApplyCredit.Value
End Sub

Private Sub cmdAdd_Click()
 Dim isValidPayment As Boolean
 Dim Payment@, InvAmount@, RBalance@
 
 Call WriteValuesToGrid
 Call Update_RunningBalance
 RBalance = CCur(lblBalance)
    If RBalance > 0 Then
       If MsgBox("Add Invoice?", vbYesNo + vbQuestion, "Add") = vbYes Then
          With grdInvoice
            Payment = CCur(Val(Format(.TextMatrix(.RowSel, cPayAmount), "#########.#0")))
            InvAmount = CCur(Val(Format(.TextMatrix(.RowSel, cAmtWVAT), "#########.#0")))
                        
            If .Rows = 1 And grdRows = 1 Then
               Call SetTextboxVisible(True)
               grdRows = grdRows + 1
               cmdRemove.Enabled = True
               cmdProcess.Enabled = True
               txtInvNum.SetFocus
               
            ElseIf CCur(lblBalance) > 0 And grdRows > 1 Then
              Call AddRow
              cmdRemove.Enabled = True
              cmdProcess.Enabled = True
              Call TextBox_Position(Next_Row)
            End If
           End With
       End If
    End If
End Sub

Private Sub cmdAdd_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case vbKeyUp
         grdInvoice.SetFocus
         grdInvoice.row = grdInvoice.Rows - 1
         Call SetTextboxVisible(True)
         grdInvoice.SetFocus
    Case vbKeyF3, vbKeyF5, vbKeyF6, vbKeyF7
           Call Grid_ShortCutKeys(KeyCode)
  End Select

End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Close Payment Window?", vbYesNo + vbQuestion, "Close") = vbYes Then
            fraPayment.Visible = False
            fraDetails.Visible = False
            fraOR.Enabled = True
            cmdOk.Enabled = False
            cmdClose.Enabled = True
            txtORNum = "_____"
            txtCustomerCode.SetFocus
    End If
End Sub

Private Sub cmdCancelInvoice_Click()
  If MsgBox("Are you sure you want to cancel all Entries?", vbYesNo + vbCritical, "Confirmation") = vbYes Then
         fraInvoice.Visible = False
         fraContest.Visible = False
         Call Ini_GridTexbox
         lblAppliedAmount = "0.0"
         lblTotalInvAmt = "0.0"
         txtPayment = 0
         txtTax = 0
         txtInvNum = ""
         lblBalance = lblTamount
         fraPayment.Visible = True
         txtChkAmount1.SetFocus
  End If
End Sub

Private Sub cmdCancelInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
           Call Grid_ShortCutKeys(KeyCode)
End Sub

Private Sub cmdClose_Click()
  If MsgBox("Close Invoice Payment Window?", vbYesNo + vbQuestion, "Close") = vbYes Then
     If rsCust_Inv.State = 1 Then 'open
        rsCust_Inv.Close
        Set rsCust_Inv = Nothing
     End If
     Unload Me
  End If
End Sub

Private Sub cmdClose_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If MsgBox("Close Invoice Payment Window?", vbYesNo + vbQuestion, "Close") = vbYes Then
         If rsCust_Inv.State = 1 Then 'open
            rsCust_Inv.Close
            Set rsCust_Inv = Nothing
         End If
         Unload Me
    End If
End If
End Sub

Private Sub cmdContest_Click(Index As Integer)
Dim ContestStatus As String

With grdInvoice
 Select Case Index
   Case 0 'Ok
        If CCur(lblContestInfo(2)) <= CCur(.TextMatrix(.RowSel, cBalance)) Then
            fraInvoice.Enabled = True
            ContestStatus = GetInvoiceContest_Status(Val(.TextMatrix(.RowSel, cInvNum)))
            
            If HasPartialPayment(Val(.TextMatrix(.RowSel, cInvNum))) Then
               ContestStatus = GetInvoiceContest_Status(Val(.TextMatrix(.RowSel, cInvNum)))
            Else
               ContestStatus = .TextMatrix(.RowSel, cConStat)
            End If
            
            txtPayment = CCur(txtPayContest(0))
            .TextMatrix(.RowSel, cConStat) = "Y" ' Contest Invoice Status
            .TextMatrix(.RowSel, cTax) = 0 'Tax %
            .TextMatrix(.RowSel, cBalLessTax) = .TextMatrix(.RowSel, cBalance)
             txtTax = 0
            .TextMatrix(.RowSel, cTaxAmount) = txtPayContest(1) 'Payment
             fraContest.Visible = False
             fraContest.Enabled = False
             fraInvoice.Enabled = True
             txtPayment.SetFocus
        Else
            MsgBox "Contested Amount should not be Greater than the Balance Amount", vbOKOnly + vbCritical, "Error"
            txtPayContest(0).SetFocus
        End If
          
   Case 1 'Cancel
      fraInvoice.Enabled = True
      txtPayment = tmpPayment
      grdInvoice.TextMatrix(grdInvoice.RowSel, cConStat) = GetInvoiceContest_Status(Val(.TextMatrix(.RowSel, cInvNum)))
      .TextMatrix(.RowSel, cTaxAmount) = "0.0"
      txtPayContest(0) = 0
      txtPayContest(1) = 0
      fraContest.Enabled = False
      fraContest.Visible = False
      fraInvoice.Enabled = True
      txtPayment.SetFocus
 End Select
 End With
End Sub

Private Sub cmdOk_Click()
   fraPayment.ZOrder 0
   Call Ini_GridTexbox
   Call InitializePayment
   fraOR.Enabled = False
   fraDetails.Visible = False
   fraPayment.Visible = True
   
   'Retrieve payment header for the invoice if existing------------------
   Call Get_PayHdr
   '---------------------------------------------------------------------
   
   cmdOk.Enabled = False
   cmdClose.Enabled = False
   txtChkAmount1.SetFocus
End Sub

Private Sub Get_PayHdr()
    Dim strSQL As String
    Dim rstInvHdr As New Recordset
    
    strSQL = "SELECT * FROM INVPAYHDR WHERE ORNum=" & txtORNum
    
    rstInvHdr.Open strSQL, gcnnBilling, , , adCmdText
    
    If Not rstInvHdr.EOF Then
        With rstInvHdr
            'Cheque No. 1
            txtChkAmount1 = !CheckAMT1
            txtCheckNo1 = !CheckNo1
            txtBank1 = !CheckBnk1
            'Cheque No.2
            txtChkAmount2 = !CheckAMT2
            txtCheckNo2 = !CheckNo2
            txtBank2 = !CheckBnk2
            'Total cheque amount
            lblTchkAmt = !CheckAMT1 + !CheckAMT2
            'Cash
            txtCash = !CashAMT
            'Total amount
            lblTotal = !TotalAmt
            'Remaining amount
            sngAvailAmt = !AvailAMT
            'Used up amount
            sngUsedAmt = !TotalAmt - !AvailAMT
            lblInvEntry(19).Visible = True
            lblUsedAmt.Visible = True
            lblUsedAmt = sngUsedAmt
            blnOR_Valid = True
        End With
    End If
    
    Set rstInvHdr = Nothing
End Sub

Private Sub cmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    If MsgBox("Close Invoice Payment Window?", vbYesNo + vbQuestion, "Close") = vbYes Then
       If rsCust_Inv.State = 1 Then 'open
          rsCust_Inv.Close
          Set rsCust_Inv = Nothing
       End If
       Unload Me
    End If
End If
End Sub

Private Sub cmdProcess_Click()
 Dim sMsg As String
 If OKtoSave(sMsg) Then
   If MsgBox("Process Payments Now?", vbYesNo + vbQuestion, "Process") = vbYes Then
        Call ProcessPayment
        fraInvoice.Visible = False
        fraPayment.Visible = False
        fraDetails.Visible = False
        fraContest.Visible = False
        fraOR.Enabled = True
        cmdOk.Enabled = False
        cmdClose.Enabled = True
        Call FilterRecordset(frmMain.cmbcust.List(frmMain.cmbcust.ListIndex))
        Call List_UnpaidBills
        txtCustomerCode.SetFocus
    End If
 Else
    MsgBox sMsg, vbOKOnly + vbCritical, "Error"
 End If
End Sub

Private Function OKtoSave(ByRef pMsg As String) As Boolean
 Dim row%, pRow%
 OKtoSave = True
 Call UpdateValues_PayDetails
 With grdInvoice
     For row = 0 To .Rows - 1
       For pRow = row + 1 To .Rows - 1
        If Val(.TextMatrix(row, cInvNum)) = Val(.TextMatrix(pRow, cInvNum)) Then
           pMsg = "Duplicate Invoice Number Found in the list"
           OKtoSave = False
           Exit Function
        End If
        Next pRow
     Next row
 End With
  If CCur(lblAppliedAmount) > CCur(lblTamount) Then
    pMsg = "Total Payment Entry should not be greater than the Total Amount" _
      & Chr(13) & " Please Check you Entry "
      OKtoSave = False
      Exit Function
  End If
 
End Function

Private Sub cmdProcess_KeyDown(KeyCode As Integer, Shift As Integer)
           Call Grid_ShortCutKeys(KeyCode)
End Sub

Private Sub cmdRemove_Click()
 If Trim(grdInvoice.TextMatrix(grdInvoice.RowSel, cInvNum)) <> "" Then
  If MsgBox("Remove Invoice # " & Val(grdInvoice.TextMatrix(grdInvoice.RowSel, cInvNum)), vbYesNo + vbQuestion, "Remove Invoice") = vbYes Then
        Call RemoveRow(grdInvoice.RowSel)
  End If
 End If
End Sub

Private Sub cmdRemove_KeyDown(KeyCode As Integer, Shift As Integer)
     Call Grid_ShortCutKeys(KeyCode)
End Sub

Private Sub cmdShowEntry_Click()
 Dim sSql As String
  If CCur(lblTotal) > 0 Then
    If IsFillUp = True Then
       sSql = "Select invnum,cuscde,cusnam,invnum,invamt,invvat,invtax,ISNULL(totalpay,0) as totalpay , contested, " _
            & " (invamt+invvat-invtax)  as InvAmount,  ((invamt+invvat-invtax)-isnull(totalpay,0)) as Balance" _
            & " From Invict " _
            & " Where cuscde = " & Trim(txtCustomerCode.Text) _
            & " AND (totalpay<>(invamt+invvat-invtax) or totalpay is null) and (status <> 'CAN' or status is NULL)"
        fraPayment.Visible = False
        fraDetails.Visible = True
        fraInvoice.Visible = True

        If rsCust_Inv.State = 0 Then
            rsCust_Inv.Open sSql, gcnnBilling, adOpenKeyset, , adCmdText
        Else
          rsCust_Inv.Requery
        End If
        
        'grdInvoice.Clear
        grdRows = 1
        Call UpdateValues_PayDetails
        Call ClearTextBoxEntry
        fraContest.Visible = False
        cmdRemove.Enabled = False
        cmdProcess.Enabled = False
        cmdAdd.Enabled = True
        lblBalance = sngAvailAmt
        cmdAdd.SetFocus
    Else
      Call ErrorHandler(Fill_upMsg)
      txtChkAmount1.SetFocus
    End If
 Else
   Call ErrorHandler(No_Amount)
  End If
End Sub

Private Sub RemoveRow(ByVal pRow As Long)
  On Error GoTo Err_RemovingRow
With grdInvoice
  If pRow = 0 And grdInvoice.Rows = 1 Then ' 1st row
     .Clear
     grdRows = 1
     cmdRemove.Enabled = False
     cmdProcess.Enabled = False
     Call ClearTextBoxEntry
     Call WriteValuesToGrid
     cmdAdd.SetFocus
     '.SetFocus
     'txtInvNum.SetFocus
     
  
  ElseIf pRow = 0 And grdInvoice.Rows > 1 Then ' 1st row
     Call TextBox_Position(Prev_Row)
     .RemoveItem (pRow)
     Call WriteValuesToTextBox
     txtPayment = grdInvoice.TextMatrix(.RowSel, cPayAmount)
     'txtInvNum.SetFocus
  Else
     Call ClearTextBoxEntry
     Call TextBox_Position(Prev_Row)
     .RemoveItem (pRow)
     Call WriteValuesToTextBox
     txtPayment = grdInvoice.TextMatrix(.RowSel, cPayAmount)
     'txtInvNum.SetFocus
  End If
End With
 lblTotalInvAmt = FVal(GetSumInvoice)
 lblAppliedAmount = FVal(GetSumPaymentEntry)
 lblTTaxAmt = FVal(GetTotalTaxAmount)
 Call Update_RunningBalance
 cmdAdd.Enabled = IIf(CCur(lblBalance) > 0, True, False)
 Exit Sub
 
Err_RemovingRow:
   Resume Next
End Sub


Private Function IsFillUp() As Boolean
  If Trim(txtChkAmount1.Text) <> "_________.__" And CCur(Trim(txtChkAmount1.Text)) >= 0 Then
     If (Trim(txtBank1.Text) <> "" And Len(Trim(txtCheckNo1)) > 0) _
        Or CDbl(txtCash) >= 0 Then
        IsFillUp = True
     End If
  End If
  If Trim(txtChkAmount2.Text) <> "_________.__" And CCur(Trim(txtChkAmount2.Text)) >= 0 Then
    If (Trim(txtBank2.Text) <> "" And Len(Trim(txtCheckNo2)) > 0) _
        Or CDbl(txtCash) >= 0 Then
        IsFillUp = True
     End If
  End If
End Function



Private Sub cmdShowEntry_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Or KeyCode = vbKeyUp Then
    Call FieldAdvance(txtPtax, cmdShowEntry, KeyCode)
 End If
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
'    Case 13
'     txtCash = FVal(Val(txtCash))
'     lblTotal = FVal(CDbl(lblTchkAmt) + txtCash)
     
     Case vbKeyF2 'Show Invoice Entry
          Call cmdShowEntry_Click
          
     Case vbKeyF1 'And chkApplyCredit.Value = 1
          chkApplyCredit.Value = IIf(chkApplyCredit.Value = 1, 0, 1)
          Call chkApplyCredit_Click
    Case vbKeyF12 ' show List of Payments
          frmListORpayment.Show vbModal
    End Select
End Sub

Private Sub Form_Load()
  On Error GoTo Err_conn
    Set rsCust_Inv = New ADODB.Recordset
    Call initializeSetting
    Call InitializePayment
    Exit Sub
Err_conn:
   MsgBox "Unable to establish Connection!" & vbCr & " Contact MIS for Assistance ", _
      vbOKOnly + vbCritical, "Error Connection"
      Unload Me
      Resume Next
End Sub

Private Sub initializeSetting()
    txtCustomerCode.Text = "______"
    txtORNum = "_____"
    lblCusName = ""
    fraPayment.Visible = False
    fraDetails.Visible = False
    fraInvoice.Visible = False
    fraContest.Visible = False
    cmdOk.Enabled = False
    grdInvoice.Rows = 1
    grdInvoice.FixedRows = 0
    fraPayment.Left = Me.Width / 2 * 0.5
    fraPayment.Top = Me.Height / 2 * 0.3
End Sub


Private Sub InitializePayment()
    txtChkAmount1.Text = "00"
    txtChkAmount2.Text = "00"
    txtBank1.Text = "__________"
    txtBank2.Text = "__________"
    txtCheckNo1 = "__________"
    txtCheckNo2 = "__________"
    txtPtax = 1
    lblTchkAmt = "0.0"
    txtCash = "0.00"
    lblTotal = FVal(GetAvailbleCredit(Trim(txtCustomerCode.Text)))
    lblTamount = "0.0"
    lblCreditAmount = "0.0"
    chkCredit.Value = 0 ' Status if Available amount is to be applied
    chkApplyCredit.Value = 0
    lblAppliedAmount = "0.0"
    lblTotalInvAmt = "0.0"
    lblBalance = "0.0"
    txtInvNum = "_____"
    txtTax = 0
    txtPayment = 0#
    lblAvailCredit = FVal(GetAvailbleCredit(Trim(txtCustomerCode.Text)))
    chkApplyCredit.Value = IIf(CCur(lblAvailCredit) > 0, 1, 0)
    chkApplyCredit.Enabled = chkApplyCredit.Value
    lblTTaxAmt = 0
End Sub

Public Function GetAvailbleCredit(ByVal pCusCode As String) As Currency
  Dim rst1 As New ADODB.Recordset
  Dim str As String
  
  str = "Select sum(AvailAmt) as Total_Availamt " _
        & " FROM INVPAYHDR " _
        & " WHERE cuscde ='" & pCusCode & "'" _
        & " AND (availAmt > 0) and  ornum<>" & Val(txtORNum)
        
  rst1.Open str, gcnnBilling, adOpenKeyset, , adCmdText
  
  If IsNull(rst1!Total_Availamt) Then
        GetAvailbleCredit = 0#
  Else
      GetAvailbleCredit = rst1!Total_Availamt
  End If
rst1.Close
Set rst1 = Nothing

End Function

Private Sub UpdateValues_PayDetails() ' in fraDetails Frame
   lblTChkamount = FVal(lblTchkAmt)
   lblTCash = FVal(txtCash)
   lblCreditAmount = FVal(GetCreditAmount(Val(txtCustomerCode.Text)))
   lblAppliedAmount = FVal(GetSumPaymentEntry)
   lblTotalInvAmt = FVal(GetSumInvoice())
   lblTTaxAmt = FVal(GetTotalTaxAmount)
   lblBalance = FVal(CDbl(lblTamount) - GetSumPaymentEntry)
   chkCredit.Value = chkApplyCredit.Value
   If chkApplyCredit.Value = 1 Then
        lblTamount = FVal(CDbl(lblTChkamount) + CCur(lblTCash) + CCur(lblAvailCredit))
   Else
        lblTamount = FVal(CCur(lblTChkamount) + CCur(lblTCash))
   End If
End Sub


Private Sub Ini_GridTexbox()
 Dim gridLabel As Label
 Dim colCtr As Integer
 colCtr = 0
 With grdInvoice
   .Clear
   .Rows = 1
   .Cols = 12
   For Each gridLabel In lblgrid
       .row = 0
       .ColWidth(colCtr) = gridLabel.Width + 5
       Select Case colCtr
         Case 0 ' Invoice #
          .TextMatrix(.RowSel, colCtr) = ""
         Case 10 ' Contest Invoice
          .TextMatrix(.RowSel, colCtr) = "N"
         Case Else
         .TextMatrix(.RowSel, colCtr) = FVal(0)
       End Select
        colCtr = colCtr + 1
   Next gridLabel
 End With
 With grdInvoice
   .row = 0
   .col = cInvNum
     .CellBackColor = &HC0C000
     txtInvNum.Left = .Left + 50
     txtInvNum.Width = .ColWidth(cInvNum)
     txtInvNum.Top = .Top + 10
     txtInvNum.Height = .RowHeight(0)
     txtInvNum.Font = .Font
   .col = 6
     .CellBackColor = &HC0C000
     .TextMatrix(0, cTax) = "0"
     txtTax.Left = .Left + .ColPos(cTax) + 50
     txtTax.Width = .ColWidth(cTax)
     txtTax.Height = .RowHeight(0)
     txtTax.Top = .Top + 10
  .col = cPayAmount
     .CellBackColor = &HC0C000
     .TextMatrix(0, cPayAmount) = "0"
     txtPayment.Left = .Left + .ColPos(cPayAmount) + 60
     txtPayment.Width = .ColWidth(cPayAmount)
     txtPayment.Height = .RowHeight(0)
     txtPayment.Top = .Top + 10
  End With
End Sub

Private Sub grdInvoice_Click()
  With grdInvoice
        .row = .RowSel
         Call SetTextboxVisible(True)
  End With
End Sub

Private Sub grdInvoice_GotFocus()
  grdInvoice.row = grdInvoice.RowSel
  Call SetTextboxVisible(True)
End Sub

Private Sub grdInvoice_KeyDown(KeyCode As Integer, Shift As Integer)
      Call Grid_ShortCutKeys(KeyCode)
End Sub

Private Sub grdInvoice_SelChange()
Dim TopPos As Long 'current Top Column Position
 With grdInvoice
    .row = .RowSel
    .col = cInvNum
    TopPos = .CellTop + .Top + 20
    txtInvNum.Left = .Left + 50
    txtInvNum.Top = TopPos
    .col = cTax
    txtTax.Left = .Left + .ColPos(cTax) + 50
    txtTax.Top = TopPos
    .col = cPayAmount
    txtPayment.Left = .Left + .ColPos(cPayAmount) + 60
    txtPayment.Top = TopPos
    txtInvNum = .TextMatrix(.RowSel, cInvNum)
    txtTax = .TextMatrix(.RowSel, cTax)
    txtPayment = .TextMatrix(.RowSel, cPayAmount)
    Call SetTextboxVisible(True)
End With
End Sub


Private Sub txtBank1_GotFocus()
    txtBank1.BackColor = &H80000005
    txtBank1.SelStart = 0
    txtBank1.SelLength = Len(txtBank1)
End Sub

Private Sub txtBank1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13, vbKeyDown, vbKeyUp
        If CCur(Val(txtChkAmount1.Text)) > 0 And (Len(Trim(txtCheckNo1)) <= 0 _
                Or Len(Trim(txtBank1.Text)) <= 0) Then
                Call ErrorHandler(Fill_upMsg)
                txtBank1.SetFocus
        Else: txtChkAmount2.SetFocus
        End If
   Case Else
        Call FieldAdvance(txtChkAmount1, txtChkAmount2, KeyCode)
   End Select

End Sub

Private Sub txtBank1_LostFocus()
    If CCur(Val(txtChkAmount1.Text)) > 0 And (Len(Trim(txtCheckNo1)) <= 0 _
            And Len(Trim(txtBank1.Text)) <= 0) Then
            Call ErrorHandler(Fill_upMsg)
            txtBank1.SetFocus
    End If
    txtBank1.BackColor = &H8000000F
End Sub

Private Sub txtBank2_GotFocus()
   txtBank2.BackColor = &H80000005
   txtBank2.SelStart = 0
   txtBank2.SelLength = Len(txtBank2)
End Sub

Private Sub txtBank2_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
   Case 13
     If CCur(Val(txtChkAmount2.Text)) > 0 And (Len(Trim(txtCheckNo2)) <= 0 _
            Or Len(Trim(txtBank2.Text)) <= 0) Then
            Call ErrorHandler(Fill_upMsg)
            txtBank2.SetFocus
     Else
        txtCash.SetFocus
     End If
    Case Else
        Call FieldAdvance(txtChkAmount2, txtBank2, KeyCode)
  End Select
End Sub

Private Sub txtBank2_LostFocus()
 If CCur(Val(txtChkAmount2.Text)) > 0 And Len(Trim(txtCheckNo2)) <= 0 _
            And Len(Trim(txtBank2.Text)) <= 0 Then
            Call ErrorHandler(Fill_upMsg)
            txtBank2.SetFocus
    End If
     txtBank2.BackColor = &H8000000F
End Sub

Private Sub txtCash_GotFocus()
  txtCash.BackColor = &H80000005
  txtCash.SelStart = 0
  txtCash.SelLength = Len(txtCash)
End Sub

Private Sub txtCash_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case 13
        txtCash = FVal(txtCash)
        lblTotal = FVal(CDbl(lblTchkAmt) + txtCash)
        txtPtax.SetFocus
    Case Else
        Call FieldAdvance(txtChkAmount1, txtPtax, KeyCode)
   End Select
End Sub

Private Sub txtCash_LostFocus()
  If Trim(txtCash) = "" Then
       txtCash = "0.0"
  Else
    'txtCash = FVal(Val(Format(txtCash, "#########.#0")))
    txtCash = FVal(txtCash)
  End If
    lblTotal = FVal(CCur(lblTchkAmt) + CCur(txtCash))
    txtCash.BackColor = &H8000000F
End Sub

Private Sub txtCheckNo1_GotFocus()
  txtCheckNo1.BackColor = &H80000005
  txtCheckNo1.SelStart = 0
  txtCheckNo1.SelLength = Len(txtCheckNo1)
End Sub

Private Sub txtCheckNo1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13, vbKeyDown
    If CCur(txtChkAmount1.Text) > 0 And (Len(Trim(txtCheckNo1)) <= 0 Or Trim(txtCheckNo1) = "") Then
            Call ErrorHandler(Fill_upMsg)
    Else: txtBank1.SetFocus
    End If
 Case Else
    Call FieldAdvance(txtChkAmount1, txtBank1, KeyCode)
 End Select

End Sub

Private Sub txtCheckNo1_LostFocus()
   txtCheckNo1.BackColor = &H8000000F
End Sub

Private Sub txtCheckNo2_GotFocus()
  txtCheckNo2.BackColor = &H80000005
  txtCheckNo2.SelStart = 0
  txtCheckNo2.SelLength = Len(txtCheckNo2)
End Sub

Private Sub txtCheckNo2_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
    Case 13, vbKeyDown
        If CCur(txtChkAmount2.Text) > 0 And (Len(Trim(txtCheckNo2)) <= 0 Or Trim(txtCheckNo2) = "") Then
                Call ErrorHandler(Fill_upMsg)
        Else: Call FieldAdvance(txtChkAmount2, txtBank2, KeyCode)
        End If
    
    Case Else
         Call FieldAdvance(txtChkAmount2, txtBank2, KeyCode)
    End Select

End Sub




Private Sub txtCheckNo2_LostFocus()
  txtCheckNo2.BackColor = &H8000000F
End Sub

Private Sub txtChkAmount1_GotFocus()
  txtChkAmount1.BackColor = &H80000005
  txtChkAmount1.SelStart = 0
  txtChkAmount1.SelLength = Len(txtChkAmount1)
End Sub

Private Sub txtChkAmount1_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case 13, vbKeyDown
      If Not IsNumeric(txtChkAmount1.Text) Then
        Call ErrorHandler(Invalid)
        txtChkAmount1.SetFocus
        Exit Sub
      ElseIf CCur(txtChkAmount1) <= 0 Then  'no check amt entry
         txtCheckNo1 = ""
         txtBank1 = ""
         txtCash.SetFocus
      Else
         txtCheckNo1.SetFocus
      End If
      lblTchkAmt = FVal(CDbl(txtChkAmount1.Text) + CDbl(txtChkAmount2.Text))
      lblTotal = FVal(CDbl(lblTchkAmt) + CDbl(txtCash))

      
    Case 27 'esc
      If MsgBox("Close Payment Window?", vbYesNo + vbQuestion, "Close") = vbYes Then
            fraPayment.Visible = False
            fraOR.Enabled = True
            txtORNum.SetFocus
       End If
   Case Else
     Call FieldAdvance(txtChkAmount1, txtCheckNo1, KeyCode)
   End Select
End Sub

Private Sub txtChkAmount1_LostFocus()
   txtChkAmount1.BackColor = &H8000000F
End Sub


Private Sub txtChkAmount2_GotFocus()
  txtChkAmount2.BackColor = &H80000005
  txtChkAmount2.SelStart = 0
  txtChkAmount2.SelLength = Len(txtChkAmount2)
End Sub

Private Sub txtChkAmount2_KeyDown(KeyCode As Integer, Shift As Integer)
  Select Case KeyCode
    Case 13, vbKeyDown
      If Not IsNumeric(txtChkAmount2.Text) Then
        Call ErrorHandler(Invalid)
        txtChkAmount2.SetFocus
        Exit Sub
      ElseIf CCur(txtChkAmount2) <= 0 Then  'no check amt entry
         txtCheckNo2 = ""
         txtBank2 = ""
         txtCash.SetFocus
      Else
        txtCheckNo2.SetFocus
      End If
      lblTchkAmt = FVal(CDbl(txtChkAmount1.Text) + CDbl(txtChkAmount2.Text))
      lblTotal = FVal(CDbl(lblTchkAmt) + CDbl(txtCash))
      
    Case Else
       Call FieldAdvance(txtChkAmount1, txtCheckNo2, KeyCode)
      
 End Select
   
End Sub

Private Sub txtChkAmount2_LostFocus()
     txtChkAmount2.BackColor = &H8000000F
End Sub


Private Sub txtCustomerCode_GotFocus()
  txtCustomerCode.BackColor = &H80000005
End Sub

Private Sub txtCustomerCode_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim rsCust As ADODB.Recordset
 Dim rsAccounts As ADODB.Recordset
 Dim sSql As String
 Dim sCust As String
 
Select Case KeyCode
 Case 13 And Len(Trim(txtCustomerCode.Text)) > 0
    sCust = "Select cuscde,cusnam From customer " _
          & " Where cuscde ='" & Trim(txtCustomerCode.Text) & "'"
          
     sSql = "Select cuscde From Invict " _
          & " Where cuscde ='" & Trim(txtCustomerCode.Text) & "'" _
          & " AND (isnull(totalpay,0) <>(invamt+invvat-invtax) or totalpay is null) and (status <> 'CAN'" _
          & " OR Status is NULL) "
     Set rsCust = New ADODB.Recordset
     rsCust.Open sCust, gcnnBilling, adOpenStatic, , adCmdText
     
     If rsCust.EOF Then
        MsgBox "Customer Code does not Exist", vbOKOnly + vbCritical, "Not found"
        txtCustomerCode.SetFocus
                 
     Else 'then check if there is unpaid invoice for that customer
        lblCusName = UCase(rsCust!cusnam)
        Set rsAccounts = New ADODB.Recordset
        rsAccounts.Open sSql, gcnnBilling, , , adCmdText
        If rsAccounts.EOF Then
             MsgBox "Invoice/s already been settled ", vbOKOnly + vbInformation, "Accounts"
             txtCustomerCode.SetFocus
        Else
          chkCredit.Value = 0
          txtORNum.SetFocus
       End If
       
       rsAccounts.Close
       Set rsAccounts = Nothing
       rsCust.Close
       Set rsCust = Nothing
     End If
     
 Case vbKeyF3
    If MsgBox("Close Invoice Payment Window?", vbYesNo + vbQuestion, "Close") = vbYes Then
       If rsCust_Inv.State = 1 Then 'open
          rsCust_Inv.Close
          Set rsCust_Inv = Nothing
       End If
       Unload Me
    End If
   Unload Me
End Select
  
End Sub


Private Sub txtInvNum_Click()
  txtInvNum.SetFocus
End Sub

Private Sub txtCustomerCode_LostFocus()
txtCustomerCode.BackColor = &H8000000F
End Sub

Private Sub txtInvNum_GotFocus()
  txtInvNum.BackColor = &H80000005
  txtInvNum.SelStart = 0
  txtInvNum.SelLength = Len(txtInvNum)
End Sub

Private Sub txtInvNum_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim isFound_Inv As Boolean
 Dim HasDuplicate_inv As Boolean
 Dim tmpInvNum As Long
 Select Case KeyCode
   Case 13 'enter
   
       isFound_Inv = IsFound(Val(txtInvNum))
       grdInvoice.TextMatrix(grdInvoice.RowSel, cInvNum) = Val(txtInvNum)
       HasDuplicate_inv = HasDuplicate(Val(txtInvNum))
       If isFound_Inv = False Then
               MsgBox " Invoice #  " & Val(txtInvNum) & " Not  Found ! ", _
                         vbOKOnly + vbCritical, "Not found"
          txtInvNum.SetFocus
          Exit Sub
       End If
       If HasDuplicate_inv = True Then
          MsgBox "Duplicate Invoice Found ! Invoice # " & txtInvNum, _
                         vbOKOnly + vbCritical, "Ivalid Entry"
          grdInvoice.TextMatrix(grdInvoice.RowSel, cInvNum) = ""
          txtInvNum.SetFocus
          Exit Sub
       End If
       tmpPayment = Val(Format(txtPayment, "#########.#0"))
       Call PlaceValueToGrid(Val(txtInvNum))
       
       With grdInvoice
          If CCur(.TextMatrix(.RowSel, cPPayment)) > 0 Or Trim(.TextMatrix(.RowSel, cConStat)) = "Y" Then 'Has Partial Payment or Contested
            txtTax = Val(.TextMatrix(.RowSel, cTax))
            .TextMatrix(.RowSel, cBalLessTax) = FVal(CCur(.TextMatrix(.RowSel, cBalance)))
            txtPayment = IIf(CCur(Format(.TextMatrix(.RowSel, cBalLessTax), "#########.#0")) > CCur(lblBalance), Format(lblBalance, "#########.#0"), Format(.TextMatrix(.RowSel, cBalLessTax), "#########.#0"))
            lblTotalInvAmt = GetSumInvoice
            DoEvents
            txtPayment.SetFocus
          Else
            txtTax.SetFocus
          End If
      End With
       
  Case vbKeyUp
    If grdInvoice.RowSel > 0 Then
        Call TextBox_Position(Prev_Row)
        Call WriteValuesToTextBox
        txtPayment = Format(grdInvoice.TextMatrix(grdInvoice.RowSel, cPayAmount), "#########.#0")
    End If
 
 Case vbKeyDown
    If grdInvoice.RowSel < grdInvoice.Rows - 1 Then
        Call TextBox_Position(Next_Row)
        Call WriteValuesToTextBox
        txtPayment = Format(grdInvoice.TextMatrix(grdInvoice.RowSel, cPayAmount), "#########.#0")
    End If

  Case Else
         Call Grid_ShortCutKeys(KeyCode)
 End Select
 End Sub
Private Function Get_TaxPercentage(ByVal pInvnum As Long) As Single
  rsCust_Inv.Find "invnum=" & pInvnum, , adSearchForward, 1
  With rsCust_Inv
    If Not .EOF Then
      If !invtax > 0 Then
       Get_TaxPercentage = (!invamt + !invvat) / !invtax
      Else
         Get_TaxPercentage = 0
      End If
    End If
  End With
End Function


Private Sub Grid_ShortCutKeys(ByVal pKeyCode As Long)
   Select Case pKeyCode
    Case vbKeyF3
        Call cmdCancelInvoice_Click
    Case vbKeyF5 'Add Row
         Call cmdAdd_Click
    Case vbKeyF6 'Remove Row
      If grdInvoice.Rows >= 1 And Trim(grdInvoice.TextMatrix(grdInvoice.RowSel, cInvNum)) <> "" And Val(grdInvoice.TextMatrix(grdInvoice.RowSel, cInvNum)) <> 0 Then
        If MsgBox("Remove Invoice # " & Val(grdInvoice.TextMatrix(grdInvoice.RowSel, cInvNum)), vbYesNo + vbQuestion, "Remove Invoice") = vbYes Then
             Call RemoveRow(grdInvoice.RowSel)
        End If
      End If
    Case vbKeyF7 'Process Payments
        Call cmdProcess_Click
        
    Case vbKeyF10 'Show Contest Invoice Entry
       If Val(txtPayment) <= CCur(grdInvoice.TextMatrix(grdInvoice.RowSel, cBalance)) Then
            fraInvoice.Enabled = False
            fraContest.Visible = True
            Call Show_ContestWindow(Val(grdInvoice.TextMatrix(grdInvoice.RowSel, cInvNum)))
            fraContest.Enabled = True
            fraContest.Visible = True
            txtPayContest(0).SetFocus
       End If
       
    Case vbKeyUp
     With grdInvoice
        .SetFocus
        .row = grdInvoice.Rows - 1
        .col = 0
        .Cols = 11
       .SelectionMode = flexSelectionByRow
       .HighLight = flexHighlightAlways
     End With
  End Select
End Sub

Private Sub Show_ContestWindow(ByVal pInvnum As Long)
  With grdInvoice
       lblContestInfo(0) = pInvnum  'Invoice #
       lblContestInfo(1) = .TextMatrix(.RowSel, cAmtWVAT)
       txtPayContest(0) = 0
       txtPayContest(1) = 0
  End With
End Sub


Private Function IsFound(ByVal pInv As Long) As Boolean
   rsCust_Inv.MoveFirst
   rsCust_Inv.Find "invnum=" & pInv, , adSearchForward, 1
   If Not rsCust_Inv.EOF Then
        IsFound = True
    End If
End Function

Private Sub PlaceValueToGrid(ByVal pInv As Long)
   rsCust_Inv.Find "invnum=" & pInv, , adSearchForward, 1
   
   If Not rsCust_Inv.EOF Then
        With grdInvoice
        .TextMatrix(.RowSel, cInvNum) = pInv
        .TextMatrix(.RowSel, cInvAmount) = FVal(rsCust_Inv!invamt)
        .TextMatrix(.RowSel, cVAT) = FVal(rsCust_Inv!invvat)
        .TextMatrix(.RowSel, cAmtWVAT) = FVal(rsCust_Inv!invvat + rsCust_Inv!invamt) 'Total InvAmount w/o tax
        .TextMatrix(.RowSel, cPPayment) = FVal(rsCust_Inv!totalpay) 'Partial Payment
        .TextMatrix(.RowSel, cBalance) = FVal(rsCust_Inv!balance) 'Balance
        .TextMatrix(.RowSel, cTax) = IIf(rsCust_Inv!totalpay > 0, 0, Val(txtPtax))  'Tax
        .TextMatrix(.RowSel, cBalLessTax) = FVal(IIf(rsCust_Inv!balance = 0, rsCust_Inv!balance, (rsCust_Inv!invamt + rsCust_Inv!invvat) - rsCust_Inv!totalpay)) 'Balance less Tax
        If Trim(.TextMatrix(.RowSel, cPayAmount)) = "" Or CCur(.TextMatrix(.RowSel, cPayAmount)) <= 0 Then
             .TextMatrix(.RowSel, cPayAmount) = 0# 'Payamount
        Else
             .TextMatrix(.RowSel, cPayAmount) = .TextMatrix(.RowSel, cPayAmount)
        End If
        .TextMatrix(.RowSel, cPayStatus) = IIf(CCur(.TextMatrix(.RowSel, cBalLessTax)) = CCur(.TextMatrix(.RowSel, cPayAmount)), "F", "P") 'status
        .TextMatrix(.RowSel, cConStat) = rsCust_Inv!Contested  'Contest Invoice Status
        .TextMatrix(.RowSel, cTaxAmount) = "0.0" 'Contest Invoice Tax Amount
        txtTax = Val(.TextMatrix(.RowSel, cTax))
    End With
  End If
End Sub

' Note: If invoice have a partial payment made then....tax should not be included for another payment
Private Function HasPartialPayment(ByVal pInv As Long) As Boolean
  rsCust_Inv.Requery
  rsCust_Inv.Find "invnum=" & pInv, , adSearchForward, 1
  With rsCust_Inv
    If Not .EOF Then
        HasPartialPayment = IIf(!totalpay > 0, True, False)
    End If
  End With
End Function
Private Sub txtInvNum_LostFocus()
  txtInvNum.BackColor = &H8000000F
  tmpPayment = Val(Format(txtPayment, "#########.#0"))
End Sub

Private Sub txtORNum_GotFocus()
 txtORNum.BackColor = &H80000005
 
 lblInvEntry(19).Visible = False
 lblUsedAmt.Visible = False
 lblUsedAmt = ""
 lblBalance = ""
 lblAppliedAmount = ""
 sngAvailAmt = 0
 sngUsedAmt = 0
 blnOR_Valid = False
End Sub

Private Sub txtORNum_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case 13 'enter
      If Trim(txtORNum) <> "_____" Then
            If OReXist(Val(txtORNum)) = True Then
              MsgBox "OR # " & Val(txtORNum) & " Already Exist", vbOKOnly + vbCritical, "Error"
              txtORNum.SetFocus
            Else
              cmdOk.Enabled = True
              DoEvents
              cmdOk.SetFocus
            End If
      Else
        DoEvents
        txtORNum.SetFocus
      End If
    Case vbKeyF3
        If MsgBox("Close Invoice Payment Window?", vbYesNo + vbQuestion, "Close") = vbYes Then
           If rsCust_Inv.State = 1 Then 'open
              rsCust_Inv.Close
              Set rsCust_Inv = Nothing
           End If
          Unload Me
        End If
    End Select
End Sub

Private Function OReXist(ByVal pORnum As Long) As Boolean
  Dim rst As ADODB.Recordset
  Dim sSql As String
  Set rst = New ADODB.Recordset
  'sSql = "Select ornum from invpaydtl where ornum=" & pORnum
  sSql = "Select ornum from invpayhdr where availamt = 0 and ornum=" & pORnum
  rst.Open sSql, gcnnBilling, , , adCmdText
  OReXist = IIf(rst.EOF, False, True)
  rst.Close
  Set rst = Nothing
End Function

'Description : Get the sum of Payment per invoice entry
Private Function GetSumPaymentEntry() As Double
Dim nRow As Integer
Dim Payments As Double
 Payments = 0#
    With grdInvoice
        For nRow = 0 To .Rows - 1
         Payments = Payments + CCur(Val(Format(.TextMatrix(nRow, cPayAmount), "#########.#0")))
        Next nRow
    End With
  GetSumPaymentEntry = Payments + sngUsedAmt
End Function

'Description : Get the sum of list of invoice
Private Function GetSumInvoice() As Double
Dim nRow As Integer
Dim TotalInvAmt As Double
Dim InvAmtLessTax As Double
 TotalInvAmt = 0#
 InvAmtLessTax = 0#
    With grdInvoice
        For nRow = 0 To .Rows - 1
         InvAmtLessTax = Abs(Val(Format(.TextMatrix(nRow, cBalLessTax), "#########.#0")))
         TotalInvAmt = TotalInvAmt + InvAmtLessTax
        Next nRow
    End With
  GetSumInvoice = TotalInvAmt
End Function

Private Function GetTotalTaxAmount() As Double
Dim nRow As Integer
Dim TotalTax As Double
 TotalTax = 0#
    With grdInvoice
        For nRow = 0 To .Rows - 1
             TotalTax = TotalTax + CCur(.TextMatrix(nRow, cTaxAmount))
        Next nRow
    End With
  GetTotalTaxAmount = TotalTax
End Function


'Description: Check if inv# has a duplicate entry in the Grid
Private Function HasDuplicate(ByVal pInvnum As Long) As Boolean
    Dim nRow As Integer
    Dim nDupilicates As Integer
    nDupilicates = 0
    With grdInvoice
        For nRow = 0 To (.Rows - 1)
           If Val(.TextMatrix(nRow, 0)) = pInvnum Then
                nDupilicates = nDupilicates + 1 '1st instance(copy)
                If nDupilicates > 1 Then
                   HasDuplicate = True
                   Exit Function
                End If
           End If
        Next nRow
    End With
End Function
Private Function IsValid_Invoice(ByVal pInv As Long) As Boolean
 Dim TaxAmount As Double
 rsCust_Inv.Find "invnum=" & pInv, , adSearchForward, 1
 If Not rsCust_Inv.EOF Then
       With grdInvoice
         TaxAmount = 0#
        .TextMatrix(.RowSel, cInvNum) = pInv
        .TextMatrix(.RowSel, cInvAmount) = FVal(rsCust_Inv!invamt)
        .TextMatrix(.RowSel, cVAT) = FVal(rsCust_Inv!invvat)
        .TextMatrix(.RowSel, cAmtWVAT) = FVal(rsCust_Inv!invvat + rsCust_Inv!invamt)
        .TextMatrix(.RowSel, cPPayment) = FVal(IIf(rsCust_Inv!totalpay > 0#, (rsCust_Inv!invamt - rsCust_Inv!totalpay), "0"))   'Partial Payment
        .TextMatrix(.RowSel, cBalance) = FVal((rsCust_Inv!totalpay) - (rsCust_Inv!invamt + rsCust_Inv!invvat - rsCust_Inv!invtax)) 'Balance
        .TextMatrix(.RowSel, cTax) = txtTax   ' FVal(IIf(CDbl(.TextMatrix(.RowSel, cTax)) <= 0, "0", .TextMatrix(.RowSel, cTax))) 'Tax
         TaxAmount = (rsCust_Inv!invamt + rsCust_Inv!invvat) * (CDbl(.TextMatrix(.RowSel, cTax)) / 100)
        .TextMatrix(.RowSel, cBalLessTax) = FVal(Abs(CDbl(.TextMatrix(.RowSel, cBalance))) - TaxAmount) 'Balance less Tax
        .TextMatrix(.RowSel, cPayAmount) = FVal(IIf(CDbl(.TextMatrix(.RowSel, cPayAmount)) <= 0, "0", .TextMatrix(.RowSel, cPayAmount))) 'Payment
        .TextMatrix(.RowSel, cPayStatus) = IIf((CDbl(.TextMatrix(.RowSel, cBalance)) - TaxAmount) = CDbl(.TextMatrix(.RowSel, cPayAmount)), "F", "P") 'status
        .TextMatrix(.RowSel, cConStat) = "N" 'Contest Default Status
        .TextMatrix(.RowSel, cTaxAmount) = "0.0"
        IsValid_Invoice = True
        If CDbl(lblBalance) > 0 And txtPayment >= CDbl(lblBalance) Then 'still have availabele amt left
               txtPayment = CDbl(.TextMatrix(.RowSel, cBalLessTax))
               .TextMatrix(.RowSel, cPayAmount) = FVal(txtPayment)
                          
        ElseIf CDbl(lblBalance) < txtPayment Then   'still have availabele amt left
                txtPayment = CDbl(lblBalance)
                .TextMatrix(.RowSel, cPayAmount) = FVal(txtPayment)
                
        ElseIf CDbl(lblBalance) <= 0 Then
          MsgBox "No Available amount!", vbOKOnly + vbCritical, "Run Out of account"
          txtPayment = 0
          txtInvNum.SetFocus
        End If
    End With
 End If
End Function
Private Function FVal(ByVal FormatVal As Currency) As String
  FVal = Format(Val(FormatVal), "###,###,###.#0")
End Function

Private Sub ErrorHandler(ByVal pError As Integer)
 Dim sMsg As String
 Select Case pError
  Case Fill_upMsg
    sMsg = "Please Fill up all necessay Entry/Entries"
  Case Invalid
     sMsg = "Invalid Input"
  Case No_Amount
     sMsg = " Please Enter amount for Invoice Payment! "
  Case Invalid_Invoice
    sMsg = "Invalid Invoice Number "
  End Select
  MsgBox sMsg, vbOKOnly + vbInformation, "Error"
  
End Sub

Private Sub SetTextboxVisible(ByVal pStatus As Boolean)
  txtPayment.Visible = pStatus
  txtInvNum.Visible = pStatus
  txtTax.Visible = pStatus
End Sub

Private Sub Update_RunningBalance()
Dim SumPayment As Currency
Dim RunningBalance As Currency

'If blnOR_Valid = True Then
'    RunningBalance = CCur(sngAvailAmt)
'    blnOR_Valid = False
'Else
    SumPayment = FVal(GetSumPaymentEntry)
    lblAppliedAmount = FVal(SumPayment)
    RunningBalance = CCur(lblTamount) - SumPayment
'End If

With grdInvoice
     If RunningBalance >= 0 Then
            lblBalance = FVal(RunningBalance)
     Else
            lblBalance = "0.0"
     End If
End With
End Sub


Private Sub AddRow()
  Dim RowColor As Long
  Dim colCtr As Integer
    With grdInvoice
        Call ClearTextBoxEntry
           .AddItem ""
        .row = .Rows - 1
       If (.Rows Mod 2) = 0 Then 'toggle row color
             RowColor = vbGreen
       Else
            RowColor = &HC0C0C0
       End If
       .row = .Rows - 1
       
       For colCtr = 0 To .Cols - 1
          .col = colCtr
          .CellBackColor = RowColor
          
          If colCtr = 0 Or colCtr = 6 Or colCtr = 8 Then
                .CellBackColor = &HC0C000
                .TextMatrix(.RowSel, colCtr) = "0"
          Else
                .CellBackColor = RowColor
         End If
       Next colCtr
       .TextMatrix(.RowSel, cInvNum) = ""
       .TextMatrix(.RowSel, cPayStatus) = "P" 'Default Status Status P=Partial ;F=Full
       .TextMatrix(.RowSel, cConStat) = GetInvoiceContest_Status(Val(.TextMatrix(.RowSel, cInvNum)))
    End With
    Call SetTextboxVisible(True)
    txtInvNum.SetFocus
End Sub

Private Sub WriteValuesToGrid()
 With grdInvoice
        .TextMatrix(.RowSel, cInvNum) = Val(txtInvNum)
        .TextMatrix(.RowSel, cTax) = Val(txtTax)
        .TextMatrix(.RowSel, cPayAmount) = FVal(txtPayment)
  End With
End Sub

Private Sub ClearTextBoxEntry()
        txtInvNum = ""
        txtTax = Val(txtPtax)
        txtPayment = 0
End Sub
Private Sub WriteValuesToTextBox()
  With grdInvoice
        txtInvNum = .TextMatrix(.RowSel, cInvNum)
        txtTax = .TextMatrix(.RowSel, cTax)
        txtPayment = IIf(CCur(.TextMatrix(.RowSel, cPayAmount)) > 0, Format(.TextMatrix(.RowSel, cPayAmount), "#########.#0"), 0#)
        DoEvents
        txtInvNum.SetFocus
  End With
End Sub

Private Sub TextBox_Position(ByVal pPosition As Integer)
 Dim TopPos As Long
   With grdInvoice
   Select Case pPosition
      Case Next_Row
            .row = .RowSel
      Case Prev_Row
            .row = IIf(.RowSel = 0, 0, .RowSel - 1)
      End Select
        .col = 0
        TopPos = .CellTop + .Top + 10
        txtInvNum.Left = .Left + 50
        txtInvNum.Top = TopPos
        .col = cTax
        txtTax.Left = .Left + .ColPos(cTax) + 50
        txtTax.Top = TopPos
        .col = cPayAmount
        txtPayment.Left = .Left + .ColPos(cPayAmount) + 60
        txtPayment.Top = TopPos
    End With
End Sub

Private Sub txtORNum_LostFocus()
  txtORNum.BackColor = &H8000000F
End Sub

Private Sub txtPayContest_GotFocus(Index As Integer)
   txtPayContest(Index).BackColor = &H80000005
   txtPayContest(Index).SelStart = 0
   txtPayContest(Index).SelLength = Len(txtPayContest(Index))
End Sub

Private Sub txtPayContest_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim sMsg As String
    If KeyCode = 13 Then
      Select Case Index
        Case 0 'Payamount
           With grdInvoice
              .TextMatrix(.RowSel, cPayAmount) = txtPayContest(0)
              txtPayment = txtPayContest(0)
              If IsPaymentValid(sMsg) = True Then
                   Call WriteValuesToGrid
                  .TextMatrix(.RowSel, cPayStatus) = IIf(CCur(.TextMatrix(.RowSel, cBalLessTax)) = CCur(.TextMatrix(.RowSel, cPayAmount)), "F", "P") 'status
                  txtPayContest(1).SetFocus
              Else
                  .TextMatrix(.RowSel, cPayAmount) = 0
                   MsgBox sMsg, vbOKOnly + vbCritical, "Error"
                   txtPayContest(0).SetFocus
              End If
            End With
            
        Case 1 'Tax amount
            cmdContest(0).SetFocus
      lblContestInfo(2) = Val(txtPayContest(0)) + Val(txtPayContest(1))
      lblContestInfo(2) = FVal(lblContestInfo(2))
       End Select
       
    ElseIf KeyCode = 27 Then 'escape
         Call cmdContest_Click(1) 'Cancel
    End If
End Sub

Private Sub txtPayContest_LostFocus(Index As Integer)
        'txtPayContest(Index).Value = FVal(Val(Format(txtPayContest(Index), "#########.#0")))
         txtPayContest(Index).BackColor = &H8000000F
        lblContestInfo(2).Caption = CCur(txtPayContest(0)) + CCur(txtPayContest(1))
        lblContestInfo(2) = FVal(lblContestInfo(2))
End Sub

Private Sub txtPayment_GotFocus()
   txtPayment.BackColor = &H80000005
   txtPayment.SelStart = 0
   txtPayment.SelLength = Len(txtPayment)
End Sub

Private Sub txtPayment_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim sMsg As String
  Select Case KeyCode
    Case 13
       If IsPaymentValid(sMsg) = True Then
         Call WriteValuesToGrid
         With grdInvoice
            .TextMatrix(.RowSel, cPayStatus) = IIf(CCur(.TextMatrix(.RowSel, cBalLessTax)) = CCur(.TextMatrix(.RowSel, cPayAmount)), "F", "P") 'status
        End With
            If CCur(lblBalance) = 0 Then
              DoEvents
              cmdAdd.Enabled = False
              cmdProcess.SetFocus
            Else
             DoEvents
             cmdAdd.Enabled = True
             cmdAdd.SetFocus
            End If
        
        Else
           MsgBox sMsg, vbOKOnly + vbCritical, "Error"
           txtPayment.SetFocus
        End If
        
    Case vbKeyUp
        If grdInvoice.RowSel > 0 Then
             grdInvoice.row = grdInvoice.RowSel - 1
             Call TextBox_Position(Prev_Row)
             Call WriteValuesToTextBox
             txtPayment = Format(grdInvoice.TextMatrix(grdInvoice.RowSel, cPayAmount), "#########.#0")
        End If
    
    Case vbKeyDown
        If grdInvoice.RowSel < (grdInvoice.Rows - 1) Then
            grdInvoice.row = grdInvoice.RowSel + 1
            Call TextBox_Position(Next_Row)
            Call WriteValuesToTextBox
            txtPayment = Format(grdInvoice.TextMatrix(grdInvoice.RowSel, cPayAmount), "#########.#0")
        End If

  Case vbKeyF3, vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF10
         Call Grid_ShortCutKeys(KeyCode)
  End Select
End Sub

Private Sub txtPayment_LostFocus()
  With grdInvoice
    .TextMatrix(.RowSel, cPayStatus) = IIf(CCur(.TextMatrix(.RowSel, cBalLessTax)) = CCur(.TextMatrix(.RowSel, cPayAmount)), "F", "P") 'status
  End With
  txtPayment.BackColor = &H8000000F
End Sub

Private Sub txtPtax_GotFocus()
  txtPtax.BackColor = &H80000005
  txtPtax.SelStart = 0
  txtPtax.SelLength = Len(txtPtax)
End Sub

Private Sub txtPtax_KeyDown(KeyCode As Integer, Shift As Integer)
   Call FieldAdvance(txtCash, cmdShowEntry, KeyCode)
End Sub

Private Sub txtPtax_LostFocus()
  If Trim(txtPtax) = "" Then
     txtPtax = 0
  End If
  txtPtax.BackColor = &H8000000F
End Sub

Private Sub txtTax_GotFocus()
  txtTax.BackColor = &H80000005
  txtTax.SelStart = 0
  txtTax.SelLength = Len(txtTax)
End Sub

Private Sub txtTax_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim TaxAmount  As Currency
 Dim InvoiceAmt  As Currency 'Balance less tax
 Dim PaymentTobeApplied As Currency
  Select Case KeyCode
  Case 13
     If Val(txtTax) >= 0 And Val(txtTax) <= 100 Then 'valid tax range 1 to 100 %
          With grdInvoice
             If HasPartialPayment(Val(txtInvNum)) Then
                    .TextMatrix(.RowSel, cTax) = 0
                     txtTax = 0
                     InvoiceAmt = Abs(CCur(.TextMatrix(.RowSel, cBalance)))
                     txtPayment.SetFocus
                     
             Else 'No partial Payment
                .TextMatrix(.RowSel, cTax) = Val(txtTax)
                TaxAmount = CCur(.TextMatrix(.RowSel, cBalance)) * (CDbl(.TextMatrix(.RowSel, cTax)) / 100)
                InvoiceAmt = Abs(CCur(.TextMatrix(.RowSel, cBalance))) - Abs(TaxAmount) 'less tax
                .TextMatrix(.RowSel, cTaxAmount) = FVal(CCur(.TextMatrix(.RowSel, cInvAmount)) * (CDbl(.TextMatrix(.RowSel, cTax)) / 100))
                txtPayment.SetFocus
             End If
             
             .TextMatrix(.RowSel, cBalLessTax) = FVal(InvoiceAmt) 'Balance less tax
             lblTotalInvAmt = FVal(GetSumInvoice)
             lblTTaxAmt = FVal(GetTotalTaxAmount)
             
             If CCur(.TextMatrix(.RowSel, cPayAmount)) > 0 Then 'Payment
                txtPayment = CCur(.TextMatrix(.RowSel, cPayAmount))
             Else
                txtPayment = Abs(IIf(CDbl(lblBalance) >= InvoiceAmt, CCur(.TextMatrix(.RowSel, cBalLessTax)), lblBalance))
             End If
          End With
          
      Else
         MsgBox "Tax Value should Range from 1-100 Only", _
                vbOKOnly + vbCritical, "Invalid Input"
         txtTax.SetFocus
      End If
  
  Case vbKeyF3, vbKeyF5, vbKeyF6, vbKeyF7
         Call Grid_ShortCutKeys(KeyCode)

 End Select
End Sub

Private Sub txtTax_LostFocus()
  If HasPartialPayment(Val(txtInvNum)) Then
       grdInvoice.TextMatrix(grdInvoice.RowSel, cTax) = 0
 Else
       grdInvoice.TextMatrix(grdInvoice.RowSel, cTax) = Val(txtTax)
 End If
  lblTotalInvAmt = FVal(GetSumInvoice)
  txtPayment = IIf(txtPayment = "", 0, txtPayment)
  txtTax.BackColor = &H8000000F
End Sub


Private Function IsPaymentValid(ByRef pMsg As String) As Boolean
    Dim tmpOrigRbalance@, Payment@, tmpOrigPayment@
    Dim tmpRBalance@, invBalance@
    
 With grdInvoice
    Payment = CCur(Val(txtPayment))
    invBalance = CCur(.TextMatrix(.RowSel, cBalLessTax))
    If Payment > invBalance Then
       IsPaymentValid = False
       pMsg = "Payment should not be greater than Invoice Balance(less tax)"
       Exit Function
    ElseIf Payment <= 0 Then
       IsPaymentValid = False
       pMsg = " Please enter payment amount for this invoice"
       Exit Function
    End If

    tmpOrigPayment = CCur(.TextMatrix(.RowSel, cPayAmount))
    tmpOrigRbalance = CCur(lblBalance)
    .TextMatrix(.RowSel, cPayAmount) = Val(txtPayment)
    tmpRBalance = GetRunningBalance
    
    If tmpRBalance >= 0 Then 'Has Sufficient avail.amt
      IsPaymentValid = True
      Call Update_RunningBalance
      lblAppliedAmount = FVal(GetSumPaymentEntry)
    ElseIf tmpRBalance < 0 Then
      .TextMatrix(.RowSel, cPayAmount) = FVal(tmpOrigPayment) 'retain old payment value
      lblAppliedAmount = GetSumPaymentEntry
      IsPaymentValid = False
      pMsg = "Insufficient Avalaible amount"
    End If
 End With
End Function

Private Function GetRunningBalance() As Currency
   Dim nRow As Integer
   Dim Payments As Currency
   Payments = 0#
   With grdInvoice
     For nRow = 0 To (.Rows - 1)
        Payments = Payments + CCur(.TextMatrix(nRow, cPayAmount))
     Next nRow
        GetRunningBalance = CCur(lblTamount) - Payments
   End With
End Function

Private Function GetCreditAmount(ByVal pCusCode As String) As Double
  Dim rst As ADODB.Recordset
  Dim sSql As String
  sSql = " Select sum(AvailAmt) as AvailableAmt from invpayhdr " _
    & " WHERE cuscde=" & pCusCode & " and ortype <> 'ADJ'"
  Set rst = New ADODB.Recordset
  rst.Open sSql, gcnnBilling, , , adCmdText
  GetCreditAmount = IIf(IsNull(rst!AvailableAmt), 0#, rst!AvailableAmt)
  rst.Close
  Set rst = Nothing
End Function

Private Function GetSumofRunningBalance(ByVal pInvnum As Long) As Currency
 Dim rs As ADODB.Recordset
 Dim sSql As String
 Set rs = New ADODB.Recordset
 
 sSql = "Select (invamt+invvat-invTax)-isnull(totalpay,0) as RBalance FROM Invict where invnum=" & pInvnum
 rs.Open sSql, gcnnBilling, , , adCmdText
 GetSumofRunningBalance = IIf(rs.EOF Or IsNull(rs!RBalance), 0, rs!RBalance)
 rs.Close
 Set rs = Nothing
End Function

Private Function GetInvoiceContest_Status(ByVal pInvnum As Long) As String
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    
    rst.Open "Select contested,totalpay from invict where invnum=" & pInvnum, gcnnBilling, , , adCmdText
    With rst
      If rst.EOF Then
            GetInvoiceContest_Status = "N"
      Else
         GetInvoiceContest_Status = rst!Contested
      End If
    End With
    rst.Close
    Set rst = Nothing
End Function


' --------------------- Saving and Updating values to Database, NTBS-Billing --------------------------
'
'Tables : INVICT, INVPAYDTL,INVPAYHDR
'
Private Sub ProcessPayment()
    Dim rsDetailPayment As ADODB.Recordset 'Invoice Payment Detail
    Dim invnum As Long
    Dim InvAmount As Currency 'including  VAT
    Dim TaxAmount As Currency
    Dim PayAmount As Currency
    Dim invBalance As Currency
    Dim paydate As Date
    Dim Contested As String * 1
    Dim sSql As String
    Dim nRows As Integer
    Dim PartialAmtApplied As Currency
    Dim CreditAmount As Currency
    Dim CashCheckTotal As Currency
    Dim invPartialPayment As Currency

    
  'On Error GoTo err_msg
    
    Set rsDetailPayment = New ADODB.Recordset
    rsDetailPayment.Open "INVPayDtl", gcnnBilling, adOpenKeyset, adLockOptimistic, adCmdTable
    With grdInvoice
          CreditAmount = GetCreditAmount(Trim(txtCustomerCode.Text))
          CashCheckTotal = CCur(lblTChkamount) + CCur(lblTCash)
           paydate = gzGetSysDate()
           nRows = 0
       While CashCheckTotal > 0 And nRows < .Rows
                invnum = Val(.TextMatrix(nRows, cInvNum))
                InvAmount = CCur(.TextMatrix(nRows, cAmtWVAT)) ' VAT Include
                PayAmount = CCur(.TextMatrix(nRows, cPayAmount))

                invBalance = CCur(.TextMatrix(nRows, cBalLessTax)) 'InvAmt+VAT-Tax
                invPartialPayment = CCur(Val(.TextMatrix(nRows, cPPayment))) 'Partial Payment
                
                If Trim(.TextMatrix(nRows, cConStat)) = "N" Then 'Not a contested Invoice
                    TaxAmount = CCur(.TextMatrix(nRows, cBalance)) - CCur(.TextMatrix(nRows, cBalLessTax))
                    TaxAmount = IIf(invPartialPayment > 0, 0, TaxAmount)  'If has partial, taxamt is set to Zero
                    Contested = "N"
                ElseIf Trim(.TextMatrix(nRows, cConStat)) = "Y" And invPartialPayment <= 0 Then  'Contested, 1st instance
                    TaxAmount = CCur(.TextMatrix(nRows, cTaxAmount))
                    invBalance = Abs(Abs(invBalance) - TaxAmount)
                    Contested = "Y"
                ElseIf Trim(.TextMatrix(nRows, cConStat)) = "Y" And invPartialPayment > 0 Then  'Contested, 2nd instance
                     Contested = "Y"
                    TaxAmount = CCur(.TextMatrix(nRows, cTaxAmount))
                    invBalance = invBalance - TaxAmount
                End If
                
            If CashCheckTotal > 0 And CashCheckTotal >= PayAmount Then 'Apply amount based on Payment Entry
                rsDetailPayment.AddNew
                rsDetailPayment.Fields("ORNum") = Val(txtORNum)
                rsDetailPayment.Fields("InvNum") = invnum
                rsDetailPayment.Fields("InvAmt") = CCur(.TextMatrix(nRows, cInvAmount)) 'excluding VAT and tax
                rsDetailPayment.Fields("PayAmt") = PayAmount
                rsDetailPayment.Fields("RBalance") = Abs(invBalance) - PayAmount
                rsDetailPayment.Fields("PayDate") = paydate
                rsDetailPayment.Fields("remarks") = ""
                rsDetailPayment.Update
                Call Update_InvictTable(invnum, TaxAmount, PayAmount, Contested)
                CashCheckTotal = Abs(CashCheckTotal) - PayAmount
                PartialAmtApplied = 0
                nRows = nRows + 1
                
            ElseIf CashCheckTotal < PayAmount And CashCheckTotal <> 0 Then
                rsDetailPayment.AddNew
                rsDetailPayment.Fields("ORNum") = Val(txtORNum)
                rsDetailPayment.Fields("InvNum") = invnum
                rsDetailPayment.Fields("InvAmt") = CCur(.TextMatrix(nRows, cInvAmount)) 'excluding VAT and tax
                rsDetailPayment.Fields("PayAmt") = CashCheckTotal 'Insufficient Cash & Check Amt
                rsDetailPayment.Fields("RBalance") = Abs(invBalance) - CashCheckTotal
                rsDetailPayment.Fields("PayDate") = paydate
                rsDetailPayment.Fields("remarks") = ""
                rsDetailPayment.Update
                Call Update_InvictTable(invnum, TaxAmount, CashCheckTotal, Contested)
                PayAmount = invPartialPayment + CashCheckTotal
                PartialAmtApplied = CashCheckTotal
                CashCheckTotal = 0
             End If
      Wend
   If blnOR_Valid = True Then
        Call Update_PaymentHeader_2(paydate, CashCheckTotal)
   Else
        Call Update_PaymentHeader(paydate, CashCheckTotal)
   End If
   rsDetailPayment.Close
    If chkApplyCredit.Value = 1 And CreditAmount > 0 And CashCheckTotal = 0 And nRows < .Rows Then
        Call ApplyCreditAmount(nRows, PartialAmtApplied, paydate)
    End If
 End With
   txtCustomerCode.Text = "______"
   txtORNum = "_____"
   Set rsDetailPayment = Nothing
   rsCust_Inv.Close
   MsgBox "All Payment/s Transaction was successfully Saved ", vbOKOnly + vbInformation, "Saved"
   Exit Sub
'err_msg:
'     MsgBox "Error occur while processing Payment" & Chr(13) _
'     & "Contact MIS for Assistance ", vbOKOnly + vbCritical, "Error"
'     If rsDetailPayment.State = adStateOpen Then
'        rsDetailPayment.Close
'        Set rsDetailPayment = Nothing
'     End If
End Sub


'Apply Available Credit amount iif Cash and Check=0 and iif chosen to apply
' This is being apply to the rest of Invoice when Amount Remitted becomes 0
'Note pPartialAmtApplied= insufficient Cash & check Amt for that Payment but has an available amt to be applied
'           for that invoice
'
Private Sub ApplyCreditAmount(ByVal pRow As Long, ByVal pPartialAmtApplied As Currency, ByVal pPayDate As Date)
  Dim rsDetailPayment As ADODB.Recordset 'Invoice Payment Detail
  Dim rsListCreditAmount As ADODB.Recordset 'OR List w/ Avail. Amount
  Dim TotalCreditAmount@, crORNum@, crAvailAmt@ 'Credit amount variables
  Dim Payment@, TaxAmount@, PartialPayment@
  Dim HasPartialPayment_Status As Boolean
  Dim invnum As Long
  Dim sSql As String
  Dim Contested As String
  
  On Error GoTo err_msg
  Set rsDetailPayment = New ADODB.Recordset
  Set rsListCreditAmount = New ADODB.Recordset
  
  rsDetailPayment.Open "INVPayDtl", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
  
  sSql = "Select ORNum,AvailAmt from InvpayHdr where Cuscde='" & Trim(txtCustomerCode.Text) & " '" _
            & " And AvailAmt>0 " & " order by ornum"
  rsListCreditAmount.Open sSql, gcnnBilling, adOpenDynamic, adCmdText
  
  With grdInvoice
     pPartialAmtApplied = CCur(.TextMatrix(pRow, cPayAmount)) - pPartialAmtApplied 'Payamount-partially applied amt
     Payment = pPartialAmtApplied
     TotalCreditAmount = GetCreditAmount(Trim(txtCustomerCode.Text))
     crORNum = rsListCreditAmount!ornum
     crAvailAmt = rsListCreditAmount!AvailAMT
 End With
 
 With rsDetailPayment
    While pRow < grdInvoice.Rows And TotalCreditAmount > 0
        invnum = Val(grdInvoice.TextMatrix(pRow, cInvNum))
        PartialPayment = CCur(grdInvoice.TextMatrix(pRow, cBalance)) - CCur(grdInvoice.TextMatrix(pRow, cBalLessTax))
        TaxAmount = CCur(grdInvoice.TextMatrix(pRow, cBalance)) - CCur(grdInvoice.TextMatrix(pRow, cBalLessTax))
        TaxAmount = IIf(PartialPayment > 0, 0, TaxAmount)  'If has partial, taxamt is set to Zero
        HasPartialPayment_Status = HasPartialPayment(invnum)
        
        If Trim(grdInvoice.TextMatrix(pRow, cConStat)) = "N" Then 'Not a contested Invoice
            TaxAmount = CCur(grdInvoice.TextMatrix(pRow, cBalance)) - CCur(grdInvoice.TextMatrix(pRow, cBalLessTax))
            TaxAmount = IIf(HasPartialPayment(invnum), 0, TaxAmount)  'If has partial, taxamt is set to Zero
            Contested = "N"
            
        ElseIf Trim(grdInvoice.TextMatrix(pRow, cConStat)) = "Y" And HasPartialPayment_Status = False Then 'Contested, 1st instance
            TaxAmount = CCur(grdInvoice.TextMatrix(pRow, cTaxAmount))
            Contested = "Y"
        ElseIf Trim(grdInvoice.TextMatrix(pRow, cConStat)) = "Y" And HasPartialPayment_Status = True Then 'Contested, 2nd instance
            Contested = GetInvoiceContest_Status(invnum)
            TaxAmount = CCur(grdInvoice.TextMatrix(pRow, cTaxAmount))
        End If

        
        
         Do While Payment <> 0 And TotalCreditAmount > 0   'offsetting payment to Zero by Avail Amount Applied
                 If Payment <= crAvailAmt Then 'Has sufficient Avail amt
                    .AddNew
                    .Fields("ORNum") = crORNum
                    .Fields("InvNum") = invnum
                    .Fields("InvAmt") = CCur(grdInvoice.TextMatrix(pRow, cInvAmount)) 'excluding VAT and tax
                    .Fields("PayAmt") = Payment
                    .Fields("RBalance") = Abs(GetSumofRunningBalance(invnum) - (Payment + TaxAmount))
                    .Fields("PayDate") = pPayDate
                    .Fields("remarks") = "Credit Amount Applied"
                    .Update
                    Call Update_InvictTable(invnum, Abs(TaxAmount), Payment, Contested)
                    TotalCreditAmount = TotalCreditAmount - Payment
                    Call Update_AvailAmount_Header(crORNum, Payment)
                    crAvailAmt = crAvailAmt - Payment
                    Payment = 0
                    pRow = pRow + 1
                 Else
                    .AddNew
                    .Fields("ORNum") = crORNum
                    rsDetailPayment.Fields("InvNum") = invnum
                    .Fields("InvAmt") = CCur(grdInvoice.TextMatrix(pRow, cInvAmount)) 'excluding VAT and tax
                    .Fields("PayAmt") = crAvailAmt
                    .Fields("RBalance") = Abs(GetSumofRunningBalance(invnum) - (crAvailAmt + TaxAmount))
                    .Fields("PayDate") = pPayDate
                    .Fields("remarks") = "Credit Amount Applied"
                    .Update
                    Call Update_InvictTable(invnum, Abs(TaxAmount), crAvailAmt, Contested)
                    TotalCreditAmount = TotalCreditAmount - crAvailAmt
                    Payment = Payment - crAvailAmt
                    Call Update_AvailAmount_Header(crORNum, crAvailAmt)
                     crAvailAmt = 0
                 End If
                
                If crAvailAmt = 0 Then
                   rsListCreditAmount.MoveNext  'Next OR w/ Avail Amt
                   If Not rsListCreditAmount.EOF Then
                        crORNum = rsListCreditAmount!ornum
                        crAvailAmt = rsListCreditAmount!AvailAMT
                    Else
                       Exit Do
                    End If
               End If
               
               If Payment = 0 And pRow < (grdInvoice.Rows) Then 'Next Row
                   Payment = CCur(grdInvoice.TextMatrix(pRow, cPayAmount)) 'Next Invoice , Payment
                   Exit Do
               End If
          Loop
     Wend
  End With
  rsDetailPayment.Close
  rsListCreditAmount.Close
  Set rsListCreditAmount = Nothing
  Set rsListCreditAmount = Nothing
  Exit Sub
  
err_msg:
    MsgBox "Error occur while applying Available Credit Amount(Adjustments)", vbOKOnly + vbCritical, "Error"
    If rsDetailPayment.State = adStateOpen Then
        rsDetailPayment.Close
        Set rsDetailPayment = Nothing
    End If
    If rsListCreditAmount.State = adStateOpen Then
        rsListCreditAmount.Close
        Set rsListCreditAmount = Nothing
    End If
End Sub

Private Sub Update_InvictTable(ByVal pInvnum As Long, ByVal pTax As Currency, ByVal pTotalPay As Currency, ByVal pContested$)
  Dim rsUpdate As ADODB.Command 'Update Invict table
  
  Set rsUpdate = New ADODB.Command
  rsUpdate.ActiveConnection = gcnnBilling
  rsUpdate.CommandText = "Update Invict set totalpay=isnull(totalpay,0) +" & pTotalPay _
         & " ,invtax=invtax +" & pTax _
         & " ,Contested='" & pContested$ & "'" _
         & " Where Invnum=" & pInvnum
  rsUpdate.Execute
  Set rsUpdate = Nothing
End Sub

Private Sub Update_PaymentHeader(ByVal pORdate As Date, ByVal pCashCheckTotal As Currency)
  Dim rsInsert As ADODB.Command 'Insert Record to  InvPaydHdr table
  Dim sCmdText As String
  Dim PaymentHeader As OR_Payment
  With PaymentHeader
    .ornum = Val(txtORNum)
    .ortype = "OR"
    .cuscde = Trim(txtCustomerCode.Text)
    .CheckAMT1 = Val(Format(txtChkAmount1.Text, "#########.#0"))
    .CheckNo1 = Val(txtCheckNo1)
    .CheckBnk1 = UCase(txtBank1.Text)
    .CheckAMT2 = Val(Format(txtChkAmount2.Text, "#########.#0"))
    .CheckNo2 = Val(txtCheckNo2)
    .CheckBnk2 = UCase(txtBank2.Text)
    .CashAMT = Val(Format(txtCash, "#########.#0"))
    .TotalAmt = .CashAMT + .CheckAMT1 + .CheckAMT2
    .AvailAMT = pCashCheckTotal
    .ORDate = pORdate
    .Userid = zCurrentUser
    sCmdText = ""
    sCmdText = .ornum & "," _
            & "'" & Trim(.ortype) & "'" & "," _
            & "'" & Trim(.cuscde) & "'" & "," _
            & .CheckAMT1 & "," _
            & "'" & Trim(.CheckNo1) & "'" & "," _
            & "'" & Trim(.CheckBnk1) & "'" & "," _
            & .CheckAMT2 & "," _
            & "'" & Trim(.CheckNo2) & "'" & "," _
            & "'" & Trim(.CheckBnk2) & "'" & "," _
            & .CashAMT & "," _
            & .TotalAmt & "," _
            & .AvailAMT & "," _
            & "'" & Trim(.ORDate) & "'" & "," _
            & "'" & Trim(.Userid) & "'"
  End With
  Set rsInsert = New ADODB.Command
  rsInsert.ActiveConnection = gcnnBilling
  rsInsert.CommandText = "INSERT INTO InvPayHdr Values ( " & sCmdText & ")"
  rsInsert.Execute
  Set rsInsert = Nothing
End Sub

Private Sub Update_PaymentHeader_2(ByVal pORdate As Date, ByVal pCashCheckTotal As Currency)
  Dim rsUpdate As ADODB.Command 'Update Record to  InvPaydHdr table
  Dim sCmdText As String
  Dim PaymentHeader As OR_Payment
  With PaymentHeader
    .ornum = Val(txtORNum)
    .ortype = "OR"
    .cuscde = Trim(txtCustomerCode.Text)
    .CheckAMT1 = Val(Format(txtChkAmount1.Text, "#########.#0"))
    .CheckNo1 = Val(txtCheckNo1)
    .CheckBnk1 = UCase(txtBank1.Text)
    .CheckAMT2 = Val(Format(txtChkAmount2.Text, "#########.#0"))
    .CheckNo2 = Val(txtCheckNo2)
    .CheckBnk2 = UCase(txtBank2.Text)
    .CashAMT = Val(Format(txtCash, "#########.#0"))
    .TotalAmt = .CashAMT + .CheckAMT1 + .CheckAMT2
    .AvailAMT = pCashCheckTotal
    .ORDate = pORdate
    .Userid = zCurrentUser
    sCmdText = ""
    sCmdText = "ORNUM=" & .ornum & "," _
            & "ORTYPE='" & Trim(.ortype) & "'" & "," _
            & "CUSCDE='" & Trim(.cuscde) & "'" & "," _
            & "CheckAMT1=" & .CheckAMT1 & "," _
            & "CheckNo1='" & Trim(.CheckNo1) & "'" & "," _
            & "CheckBnk1='" & Trim(.CheckBnk1) & "'" & "," _
            & "CheckAMT2=" & .CheckAMT2 & "," _
            & "CheckNo2='" & Trim(.CheckNo2) & "'" & "," _
            & "CheckBnk2='" & Trim(.CheckBnk2) & "'" & "," _
            & "CashAMT=" & .CashAMT & "," _
            & "TotalAmt=" & .TotalAmt & "," _
            & "AvailAmt=" & .AvailAMT & "," _
            & "ORDate='" & Trim(.ORDate) & "'" & "," _
            & "Userid='" & Trim(.Userid) & "' Where ORNUM=" & .ornum
  End With
  Set rsUpdate = New ADODB.Command
  rsUpdate.ActiveConnection = gcnnBilling
  rsUpdate.CommandText = "UPDATE InvPayHdr SET " & sCmdText
  rsUpdate.Execute
  Set rsUpdate = Nothing
End Sub

Private Sub Update_AvailAmount_Header(ByVal pORnum As Long, ByVal pAvailAmt As Currency)
  Dim rsUpdate As ADODB.Command 'Update InvPayDtl table
  Set rsUpdate = New ADODB.Command
  rsUpdate.ActiveConnection = gcnnBilling
  rsUpdate.CommandText = "Update InvPayhdr set availamt=isnull(availamt,0)-" & pAvailAmt _
         & " Where ORNum=" & pORnum
  rsUpdate.Execute
  Set rsUpdate = Nothing
End Sub

Private Sub FieldAdvance(ByVal pPrev, pNext As Control, ByVal pKeyCode As KeyCodeConstants)
    Select Case pKeyCode
      Case 13
               If TypeOf pNext Is TextBox Or TypeOf pNext Is MaskEdBox Then
                    pNext.SelStart = 0
                    pNext.SelLength = Len(pNext)
                End If
                pNext.SetFocus
                
      Case vbKeyUp, 27
           If TypeOf pNext Is TextBox Or TypeOf pNext Is MaskEdBox Then
                pPrev.SelStart = 0
                pPrev.SelLength = Len(pPrev)
           End If
           pPrev.SetFocus
       
      Case vbKeyF2 'Show Invoice Entry
               Call cmdShowEntry_Click
               
      Case vbKeyF3
               Call cmdCancel_Click
               
     Case vbKeyF1 And chkApplyCredit.Value = 1
             Call chkApplyCredit_Click
      End Select
End Sub

Private Sub updIncrement_DownClick()
    txtPtax = updIncrement.Value
End Sub

Private Sub updIncrement_UpClick()
  txtPtax = updIncrement.Value
End Sub
