VERSION 5.00
Object = "{215FDF11-012A-11D3-BD4F-00105A64485A}#5.0#0"; "MpUserControls.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl usrctrlCCRde08 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "IBM3270 - 1254"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   11925
   ScaleWidth      =   15360
   Begin VB.CommandButton cmdRefund 
      Caption         =   "F8 - Refund"
      Height          =   600
      Left            =   12240
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   600
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   156
      Top             =   1200
      Width           =   15015
   End
   Begin VB.CommandButton cmdContainer 
      Caption         =   "F4 - Container No"
      Height          =   600
      Left            =   120
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "F5 - Details"
      Height          =   600
      Left            =   3240
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancellation 
      Caption         =   "F6 - Cancellation"
      Height          =   600
      Left            =   6240
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton cmdPayment 
      Caption         =   "F7 - Payment"
      Height          =   600
      Left            =   9360
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   600
      Width           =   2775
   End
   Begin TabDlg.SSTab CorrectionTab 
      Height          =   10455
      Left            =   0
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   1440
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   18441
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "CONTAINER NUMBER"
      TabPicture(0)   =   "usrctrlCCRde08.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame12"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdClose1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancel1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSave1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "DETAILS"
      TabPicture(1)   =   "usrctrlCCRde08.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label39"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame13"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "utxtDetTeller1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "utxtDetDteTme"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdClose2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdCancel2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdSave2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "utxtDetEntnum(0)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "utxtDetEntnum(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "CANCELLATION"
      TabPicture(2)   =   "usrctrlCCRde08.ctx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label40"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame14"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "utxtCnlSeqnum"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "utxtCnlRefnum"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "utxtCnlExporter"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "utxtCnlBroker"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "utxtCnlCommodity"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "utxtCnlVessel"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "utxtCnlCCRNum"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "utxtCnlEntnum10"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "utxtCnlEntnum1"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "utxtCnlEntnum9"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "utxtCnlEntnum8"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "utxtCnlEntnum7"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "utxtCnlEntnum6"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "utxtCnlEntnum5"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "utxtCnlEntnum4"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "utxtCnlEntnum3"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "utxtCnlEntnum2"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "utxtCnlDteTme"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "utxtCnlTeller2"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "cmdClose3"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "cmdCancel3"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "cmdSave3"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).ControlCount=   25
      TabCaption(3)   =   "PAYMENT"
      TabPicture(3)   =   "usrctrlCCRde08.ctx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label37"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame15"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame7"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "utxtPymChqNum(2)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "utxtPymChqNum(1)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "utxtPymCash"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "utxtPymChq(0)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "utxtPymChq(1)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "utxtPymChq(2)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "utxtPymChq(3)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "utxtPymChq(4)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "utxtPymAdr"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "utxtCustNo"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "utxtChange"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "utxtCustName"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "utxtPymReference"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "cmdClose4"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "cmdCancel4"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "cmdSave4"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "utxtPymChqNum(0)"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "utxtPymChqNum(3)"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "utxtPymChqNum(4)"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "utxtPymChqBnk(1)"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "utxtPymChqBnk(2)"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "utxtPymChqBnk(3)"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "utxtPymChqBnk(4)"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).ControlCount=   26
      TabCaption(4)   =   "Refund"
      TabPicture(4)   =   "usrctrlCCRde08.ctx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label65"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame19"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame16"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "utxtRefPrefix"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "utxtRefCntno"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "utxtRefSeq"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "utxtRefRefNum"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "utxtRefExporter"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "utxtRefBroker"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "utxtRefCommodity"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "utxtRefVessel"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "utxtRefCCR"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "utxtRefEntno(0)"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "utxtRefDate"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "utxtRefTeller"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "utxtRefEntno(1)"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "utxtRefEntno(2)"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "utxtRefEntno(3)"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "utxtRefEntno(4)"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "utxtRefEntno(5)"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "utxtRefEntno(6)"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "utxtRefEntno(7)"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "utxtRefEntno(8)"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "utxtRefEntno(9)"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "cmdSave5"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "cmdCancel5"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "cmdClose5"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).ControlCount=   27
      Begin VB.CommandButton cmdClose5 
         Caption         =   "F3 - Close"
         Height          =   615
         Left            =   -65520
         TabIndex        =   182
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancel5 
         Caption         =   "F12 - Cancel"
         Height          =   615
         Left            =   -63000
         TabIndex        =   181
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdSave5 
         Caption         =   "F2 - Continue"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -68040
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin MpUserControls.utxtNumBilling utxtRefEntno 
         Height          =   420
         Index           =   9
         Left            =   -66840
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   6090
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtRefEntno 
         Height          =   420
         Index           =   8
         Left            =   -68400
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   6090
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtRefEntno 
         Height          =   420
         Index           =   7
         Left            =   -69960
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   6090
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtRefEntno 
         Height          =   420
         Index           =   6
         Left            =   -71520
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   6090
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtRefEntno 
         Height          =   420
         Index           =   5
         Left            =   -73080
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   6090
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtRefEntno 
         Height          =   420
         Index           =   4
         Left            =   -66840
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   5490
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtRefEntno 
         Height          =   420
         Index           =   3
         Left            =   -68400
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   5490
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtRefEntno 
         Height          =   420
         Index           =   2
         Left            =   -69960
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   5490
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtRefEntno 
         Height          =   420
         Index           =   1
         Left            =   -71520
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   5490
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin CCRde08.utxtEntry utxtDetEntnum 
         Height          =   420
         Index           =   1
         Left            =   -72360
         TabIndex        =   19
         Top             =   5490
         Width           =   1455
         _ExtentX        =   2566
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
         MaxLength       =   8
         Alignment       =   1
      End
      Begin CCRde08.utxtEntry utxtDetEntnum 
         Height          =   420
         Index           =   0
         Left            =   -73920
         TabIndex        =   18
         Top             =   5490
         Width           =   1455
         _ExtentX        =   2566
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
         MaxLength       =   8
         Alignment       =   1
      End
      Begin MpUserControls.utxtTextBilling utxtPymChqBnk 
         Height          =   420
         Index           =   4
         Left            =   -67200
         TabIndex        =   65
         Top             =   4770
         Width           =   2415
         _ExtentX        =   4260
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
         Alignment       =   1
         MaxLength       =   10
      End
      Begin MpUserControls.utxtTextBilling utxtPymChqBnk 
         Height          =   420
         Index           =   3
         Left            =   -67200
         TabIndex        =   62
         Top             =   4290
         Width           =   2415
         _ExtentX        =   4260
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
         Alignment       =   1
         MaxLength       =   10
      End
      Begin MpUserControls.utxtTextBilling utxtPymChqBnk 
         Height          =   420
         Index           =   2
         Left            =   -67200
         TabIndex        =   59
         Top             =   3810
         Width           =   2415
         _ExtentX        =   4260
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
         Alignment       =   1
         MaxLength       =   10
      End
      Begin MpUserControls.utxtTextBilling utxtPymChqBnk 
         Height          =   420
         Index           =   1
         Left            =   -67200
         TabIndex        =   56
         Top             =   3330
         Width           =   2415
         _ExtentX        =   4260
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
         Alignment       =   1
         MaxLength       =   10
      End
      Begin MpUserControls.utxtTextBilling utxtPymChqNum 
         Height          =   420
         Index           =   4
         Left            =   -69720
         TabIndex        =   64
         Top             =   4770
         Width           =   2415
         _ExtentX        =   4260
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
         Alignment       =   1
         MaxLength       =   10
      End
      Begin MpUserControls.utxtTextBilling utxtPymChqNum 
         Height          =   420
         Index           =   3
         Left            =   -69720
         TabIndex        =   61
         Top             =   4290
         Width           =   2415
         _ExtentX        =   4260
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
         Alignment       =   1
         MaxLength       =   10
      End
      Begin MpUserControls.utxtTextBilling utxtPymChqNum 
         Height          =   420
         Index           =   0
         Left            =   -69720
         TabIndex        =   52
         Top             =   2850
         Width           =   2415
         _ExtentX        =   4260
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
         Alignment       =   1
         MaxLength       =   10
      End
      Begin VB.CommandButton cmdSave4 
         Caption         =   "F2 - Continue"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -68040
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdSave3 
         Caption         =   "F2 - Continue"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -68040
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdSave2 
         Caption         =   "F2 - Continue"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -68040
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdSave1 
         Caption         =   "F2 - Continue"
         Enabled         =   0   'False
         Height          =   615
         Left            =   6960
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancel4 
         Caption         =   "F12 - Cancel"
         Height          =   615
         Left            =   -63000
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancel3 
         Caption         =   "F12 - Cancel"
         Height          =   615
         Left            =   -63000
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancel2 
         Caption         =   "F12 - Cancel"
         Height          =   615
         Left            =   -63000
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdCancel1 
         Caption         =   "F12 - Cancel"
         Height          =   615
         Left            =   12000
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdClose4 
         Caption         =   "F3 - Close"
         Height          =   615
         Left            =   -65520
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdClose3 
         Caption         =   "F3 - Close"
         Height          =   615
         Left            =   -65520
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdClose2 
         Caption         =   "F3 - Close"
         Height          =   615
         Left            =   -65520
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin VB.CommandButton cmdClose1 
         Caption         =   "F3 - Close"
         Height          =   615
         Left            =   9480
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   8040
         Width           =   2415
      End
      Begin MpUserControls.utxtNumBilling utxtPymReference 
         Height          =   420
         Left            =   -72240
         TabIndex        =   49
         Top             =   810
         Width           =   1455
         _ExtentX        =   2566
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
         MaxLength       =   8
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
         DecimalPlaces   =   2
         First           =   -1  'True
         Last            =   -1  'True
      End
      Begin MpUserControls.utxtTextBilling utxtCustName 
         Height          =   420
         Left            =   -69720
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   5760
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   40
      End
      Begin MpUserControls.utxtNumBilling utxtChange 
         Height          =   420
         Left            =   -72240
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   6960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   16
         Maskformat      =   "###,###.00"
         Maskformat      =   "###,###.00"
         DecimalPlaces   =   2
      End
      Begin MpUserControls.utxtNumBilling utxtCustNo 
         Height          =   420
         Left            =   -72240
         TabIndex        =   66
         Top             =   5760
         Width           =   2415
         _ExtentX        =   4260
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
         MaxLength       =   6
         DecimalPlaces   =   2
      End
      Begin MpUserControls.utxtNumBilling utxtPymAdr 
         Height          =   420
         Left            =   -72240
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   6360
         Width           =   2415
         _ExtentX        =   4260
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
         MaxLength       =   18
         Maskformat      =   "###,###.00"
         Maskformat      =   "###,###.00"
         DecimalPlaces   =   2
         Last            =   -1  'True
      End
      Begin MpUserControls.utxtNumBilling utxtPymChq 
         Height          =   420
         Index           =   4
         Left            =   -72240
         TabIndex        =   63
         Top             =   4710
         Width           =   2415
         _ExtentX        =   4260
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
         MaxLength       =   20
         Maskformat      =   "###,###.00"
         Maskformat      =   "###,###.00"
         DecimalPlaces   =   2
      End
      Begin MpUserControls.utxtNumBilling utxtPymChq 
         Height          =   420
         Index           =   3
         Left            =   -72240
         TabIndex        =   60
         Top             =   4230
         Width           =   2415
         _ExtentX        =   4260
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
         MaxLength       =   23
         Maskformat      =   "###,###.00"
         Maskformat      =   "###,###.00"
         DecimalPlaces   =   2
      End
      Begin MpUserControls.utxtNumBilling utxtPymChq 
         Height          =   420
         Index           =   2
         Left            =   -72240
         TabIndex        =   57
         Top             =   3750
         Width           =   2415
         _ExtentX        =   4260
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
         MaxLength       =   24
         Maskformat      =   "###,###.00"
         Maskformat      =   "###,###.00"
         DecimalPlaces   =   2
      End
      Begin MpUserControls.utxtNumBilling utxtPymChq 
         Height          =   420
         Index           =   1
         Left            =   -72240
         TabIndex        =   54
         Top             =   3300
         Width           =   2415
         _ExtentX        =   4260
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
         MaxLength       =   25
         Maskformat      =   "###,###.00"
         Maskformat      =   "###,###.00"
         DecimalPlaces   =   2
      End
      Begin MpUserControls.utxtNumBilling utxtPymChq 
         Height          =   420
         Index           =   0
         Left            =   -72240
         TabIndex        =   51
         Top             =   2820
         Width           =   2415
         _ExtentX        =   4260
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
         MaxLength       =   26
         Maskformat      =   "###,###.00"
         Maskformat      =   "###,###.00"
         DecimalPlaces   =   2
      End
      Begin MpUserControls.utxtNumBilling utxtPymCash 
         Height          =   420
         Left            =   -72240
         TabIndex        =   50
         Top             =   2160
         Width           =   2415
         _ExtentX        =   4260
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
         MaxLength       =   27
         Maskformat      =   "###,###.00"
         Maskformat      =   "###,###.00"
         DecimalPlaces   =   2
         First           =   -1  'True
      End
      Begin MpUserControls.utxtTextBilling utxtCnlTeller2 
         Height          =   420
         Left            =   -70560
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   7350
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
      End
      Begin MpUserControls.utxtTextBilling utxtCnlDteTme 
         Height          =   420
         Left            =   -70560
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   6870
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   20
      End
      Begin MpUserControls.utxtNumBilling utxtCnlEntnum2 
         Height          =   420
         Left            =   -72360
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   5670
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtCnlEntnum3 
         Height          =   420
         Left            =   -70800
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   5670
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtCnlEntnum4 
         Height          =   420
         Left            =   -69240
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   5670
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtCnlEntnum5 
         Height          =   420
         Left            =   -67680
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   5670
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtCnlEntnum6 
         Height          =   420
         Left            =   -73920
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtCnlEntnum7 
         Height          =   420
         Left            =   -72360
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtCnlEntnum8 
         Height          =   420
         Left            =   -70800
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtCnlEntnum9 
         Height          =   420
         Left            =   -69240
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtCnlEntnum1 
         Height          =   420
         Left            =   -73920
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   5700
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtCnlEntnum10 
         Height          =   420
         Left            =   -67680
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtCnlCCRNum 
         Height          =   420
         Left            =   -70680
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtTextBilling utxtCnlVessel 
         Height          =   420
         Left            =   -70680
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2790
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   12
      End
      Begin MpUserControls.utxtTextBilling utxtCnlCommodity 
         Height          =   420
         Left            =   -70680
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3270
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   30
      End
      Begin MpUserControls.utxtTextBilling utxtCnlBroker 
         Height          =   420
         Left            =   -70680
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3750
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   30
      End
      Begin MpUserControls.utxtTextBilling utxtCnlExporter 
         Height          =   420
         Left            =   -70680
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4200
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   30
      End
      Begin MpUserControls.utxtNumBilling utxtCnlRefnum 
         Height          =   420
         Left            =   -70680
         TabIndex        =   30
         Top             =   750
         Width           =   1455
         _ExtentX        =   2566
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
         MaxLength       =   8
         First           =   -1  'True
      End
      Begin MpUserControls.utxtNumBilling utxtCnlSeqnum 
         Height          =   420
         Left            =   -70680
         TabIndex        =   31
         Top             =   1230
         Width           =   735
         _ExtentX        =   1296
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
         MaxLength       =   3
         Last            =   -1  'True
      End
      Begin MpUserControls.utxtTextBilling utxtDetDteTme 
         Height          =   420
         Left            =   -70680
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   6840
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   20
      End
      Begin MpUserControls.utxtTextBilling utxtDetTeller1 
         Height          =   420
         Left            =   -70680
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   7320
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
      End
      Begin VB.Frame Frame6 
         Height          =   7455
         Left            =   -74880
         TabIndex        =   141
         Top             =   480
         Width           =   15015
         Begin VB.Frame Frame11 
            Caption         =   " BOC Permit Numbers "
            Height          =   1455
            Left            =   720
            TabIndex        =   142
            Top             =   4920
            Width           =   8295
         End
         Begin VB.Label Label28 
            Caption         =   "Teller"
            Height          =   255
            Left            =   720
            TabIndex        =   152
            Top             =   7080
            Width           =   2895
         End
         Begin VB.Label Label27 
            Caption         =   "Date/Time Issued"
            Height          =   285
            Left            =   720
            TabIndex        =   151
            Top             =   6600
            Width           =   2895
         End
         Begin VB.Label Label26 
            Caption         =   "Exporter"
            Height          =   405
            Left            =   720
            TabIndex        =   150
            Top             =   3840
            Width           =   2895
         End
         Begin VB.Label Label25 
            Caption         =   "Broker"
            Height          =   405
            Left            =   720
            TabIndex        =   149
            Top             =   3360
            Width           =   2895
         End
         Begin VB.Label Label24 
            Caption         =   "Commodity"
            Height          =   405
            Left            =   720
            TabIndex        =   148
            Top             =   2880
            Width           =   2895
         End
         Begin VB.Label Label23 
            Caption         =   "Vessel"
            Height          =   405
            Left            =   720
            TabIndex        =   147
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label Label22 
            Caption         =   "CCR Number"
            Height          =   405
            Left            =   720
            TabIndex        =   146
            Top             =   1920
            Width           =   2895
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Other Details"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   145
            Top             =   1320
            Width           =   15135
         End
         Begin VB.Label Label20 
            Caption         =   "Sequence Number"
            Height          =   405
            Left            =   720
            TabIndex        =   144
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label Label19 
            Caption         =   "Reference Number"
            Height          =   405
            Left            =   720
            TabIndex        =   143
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame5 
         Height          =   7455
         Left            =   -74880
         TabIndex        =   129
         Top             =   480
         Width           =   15015
         Begin CCRde08.utxtEntry utxtDetEntnum 
            Height          =   420
            Index           =   9
            Left            =   7200
            TabIndex        =   27
            Top             =   5640
            Width           =   1455
            _ExtentX        =   2566
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
            MaxLength       =   8
            Alignment       =   1
         End
         Begin CCRde08.utxtEntry utxtDetEntnum 
            Height          =   420
            Index           =   8
            Left            =   5640
            TabIndex        =   26
            Top             =   5640
            Width           =   1455
            _ExtentX        =   2566
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
            MaxLength       =   8
            Alignment       =   1
         End
         Begin CCRde08.utxtEntry utxtDetEntnum 
            Height          =   420
            Index           =   7
            Left            =   4080
            TabIndex        =   25
            Top             =   5640
            Width           =   1455
            _ExtentX        =   2566
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
            MaxLength       =   8
            Alignment       =   1
         End
         Begin CCRde08.utxtEntry utxtDetEntnum 
            Height          =   420
            Index           =   6
            Left            =   2520
            TabIndex        =   24
            Top             =   5640
            Width           =   1455
            _ExtentX        =   2566
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
            MaxLength       =   8
            Alignment       =   1
         End
         Begin CCRde08.utxtEntry utxtDetEntnum 
            Height          =   420
            Index           =   5
            Left            =   960
            TabIndex        =   23
            Top             =   5640
            Width           =   1455
            _ExtentX        =   2566
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
            MaxLength       =   8
            Alignment       =   1
         End
         Begin CCRde08.utxtEntry utxtDetEntnum 
            Height          =   420
            Index           =   4
            Left            =   7200
            TabIndex        =   22
            Top             =   5040
            Width           =   1455
            _ExtentX        =   2566
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
            MaxLength       =   8
            Alignment       =   1
         End
         Begin CCRde08.utxtEntry utxtDetEntnum 
            Height          =   420
            Index           =   3
            Left            =   5640
            TabIndex        =   21
            Top             =   5040
            Width           =   1455
            _ExtentX        =   2566
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
            MaxLength       =   8
            Alignment       =   1
         End
         Begin CCRde08.utxtEntry utxtDetEntnum 
            Height          =   420
            Index           =   2
            Left            =   4080
            TabIndex        =   20
            Top             =   5040
            Width           =   1455
            _ExtentX        =   2566
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
            MaxLength       =   8
            Alignment       =   1
         End
         Begin VB.Frame Frame9 
            Caption         =   " BOC Permit Numbers "
            Height          =   1575
            Left            =   720
            TabIndex        =   130
            Top             =   4680
            Width           =   8175
         End
         Begin MpUserControls.utxtTextBilling utxtDetExporter 
            Height          =   420
            Left            =   3840
            TabIndex        =   15
            Top             =   3120
            Width           =   5055
            _ExtentX        =   8916
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
            MaxLength       =   30
         End
         Begin MpUserControls.utxtTextBilling utxtDetBroker 
            Height          =   420
            Left            =   3840
            TabIndex        =   14
            Top             =   2640
            Width           =   5055
            _ExtentX        =   8916
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
            MaxLength       =   30
         End
         Begin MpUserControls.utxtTextBilling utxtDetVessel 
            Height          =   420
            Left            =   3840
            TabIndex        =   12
            Top             =   1680
            Width           =   2175
            _ExtentX        =   3836
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
            MaxLength       =   12
         End
         Begin MpUserControls.utxtNumBilling utxtDetCCRNum 
            Height          =   420
            Left            =   3840
            TabIndex        =   11
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
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
            MaxLength       =   8
            Maskformat      =   "00000000"
            Maskformat      =   "00000000"
         End
         Begin MpUserControls.utxtNumBilling utxtDetSeqnum 
            Height          =   420
            Left            =   8760
            TabIndex        =   10
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            MaxLength       =   3
            Last            =   -1  'True
         End
         Begin MpUserControls.utxtNumBilling utxtDetRefnum 
            Height          =   420
            Left            =   3840
            TabIndex        =   9
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
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
            MaxLength       =   8
            First           =   -1  'True
         End
         Begin MpUserControls.utxtTextBilling utxtDetCommodity 
            Height          =   420
            Left            =   3840
            TabIndex        =   13
            Top             =   2160
            Width           =   5055
            _ExtentX        =   8916
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
            MaxLength       =   30
         End
         Begin MpUserControls.utxtTextBilling utxtDetWhfcde 
            Height          =   420
            Left            =   9000
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   3120
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   741
            BackColor       =   -2147483633
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   30
         End
         Begin MpUserControls.utxtTextBilling utxtDetGuarantee 
            Height          =   420
            Left            =   3840
            TabIndex        =   17
            Top             =   3600
            Width           =   375
            _ExtentX        =   661
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
            MaxLength       =   30
         End
         Begin VB.Label Label54 
            Caption         =   "{Y/N}"
            Height          =   285
            Left            =   4440
            TabIndex        =   166
            Top             =   3720
            Width           =   855
         End
         Begin VB.Label Label46 
            Caption         =   "Guarantee"
            Height          =   285
            Left            =   720
            TabIndex        =   165
            Top             =   3720
            Width           =   2895
         End
         Begin VB.Label Label18 
            Caption         =   "Teller"
            Height          =   285
            Left            =   840
            TabIndex        =   140
            Top             =   6960
            Width           =   2895
         End
         Begin VB.Label Label17 
            Caption         =   "Date/Time Issued"
            Height          =   285
            Left            =   840
            TabIndex        =   139
            Top             =   6480
            Width           =   2895
         End
         Begin VB.Label Label16 
            Caption         =   "Exporter"
            Height          =   405
            Left            =   720
            TabIndex        =   138
            Top             =   3240
            Width           =   2895
         End
         Begin VB.Label Label15 
            Caption         =   "Broker"
            Height          =   405
            Left            =   720
            TabIndex        =   137
            Top             =   2760
            Width           =   2895
         End
         Begin VB.Label Label14 
            Caption         =   "Commodity"
            Height          =   405
            Left            =   720
            TabIndex        =   136
            Top             =   2280
            Width           =   2895
         End
         Begin VB.Label Label13 
            Caption         =   "Vessel"
            Height          =   405
            Left            =   720
            TabIndex        =   135
            Top             =   1800
            Width           =   2895
         End
         Begin VB.Label Label12 
            Caption         =   "CCR Number"
            Height          =   405
            Left            =   720
            TabIndex        =   134
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Correct Details"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   -360
            TabIndex        =   133
            Top             =   720
            Width           =   15495
         End
         Begin VB.Label Label10 
            Caption         =   "Sequence Number"
            Height          =   405
            Left            =   6120
            TabIndex        =   132
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label9 
            Caption         =   "Reference Number"
            Height          =   405
            Left            =   960
            TabIndex        =   131
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame Frame3 
         Height          =   7695
         Left            =   120
         TabIndex        =   120
         Top             =   240
         Width           =   15015
         Begin MpUserControls.utxtTextBilling utxtCntNum 
            Height          =   420
            Left            =   4680
            TabIndex        =   3
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
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
         Begin MpUserControls.utxtTextBilling utxtTeller 
            Height          =   420
            Left            =   3600
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   4800
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   741
            BackColor       =   -2147483633
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   10
         End
         Begin MpUserControls.utxtTextBilling utxtTimeIssue 
            Height          =   420
            Left            =   3600
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   4200
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   741
            BackColor       =   -2147483633
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   10
         End
         Begin MpUserControls.utxtTextBilling utxtDateIssue 
            Height          =   420
            Left            =   3600
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   3600
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   741
            BackColor       =   -2147483633
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   10
         End
         Begin MpUserControls.utxtTextBilling utxtNewCntPrefix 
            Height          =   420
            Left            =   3600
            TabIndex        =   4
            Top             =   3000
            Width           =   855
            _ExtentX        =   1508
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
            MaxLength       =   4
         End
         Begin MpUserControls.utxtTextBilling utxtCntPrefix 
            Height          =   420
            Left            =   3720
            TabIndex        =   2
            Top             =   1920
            Width           =   855
            _ExtentX        =   1508
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
            MaxLength       =   4
         End
         Begin MpUserControls.utxtNumBilling utxtSeqnum 
            Height          =   420
            Left            =   3720
            TabIndex        =   1
            Top             =   1320
            Width           =   735
            _ExtentX        =   1296
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
            MaxLength       =   3
            Maskformat      =   "000"
            Maskformat      =   "000"
         End
         Begin MpUserControls.utxtNumBilling utxtRefnum 
            Height          =   420
            Left            =   3720
            TabIndex        =   0
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
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
            MaxLength       =   8
            Maskformat      =   "00000000"
            Maskformat      =   "00000000"
         End
         Begin MpUserControls.utxtTextBilling utxtNewCntNum 
            Height          =   420
            Left            =   4560
            TabIndex        =   5
            Top             =   3000
            Width           =   1455
            _ExtentX        =   2566
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
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            BackColor       =   &H00800080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Container Correction"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   -1200
            TabIndex        =   184
            Top             =   0
            Width           =   17400
         End
         Begin VB.Label Label8 
            Caption         =   "Teller"
            Height          =   405
            Left            =   2400
            TabIndex        =   128
            Top             =   4920
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Time Issued"
            Height          =   405
            Left            =   1560
            TabIndex        =   127
            Top             =   4320
            Width           =   1935
         End
         Begin VB.Label Label6 
            Caption         =   "Date Issued"
            Height          =   405
            Left            =   1560
            TabIndex        =   126
            Top             =   3720
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Container Number"
            Height          =   405
            Left            =   600
            TabIndex        =   125
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Correct Container"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   -360
            TabIndex        =   124
            Top             =   2520
            Width           =   16455
         End
         Begin VB.Label Label3 
            Caption         =   "Container Number"
            Height          =   405
            Left            =   720
            TabIndex        =   123
            Top             =   1920
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Sequence  Number"
            Height          =   405
            Left            =   720
            TabIndex        =   122
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Reference Number"
            Height          =   405
            Left            =   720
            TabIndex        =   121
            Top             =   720
            Width           =   2895
         End
      End
      Begin VB.Frame Frame12 
         Height          =   1095
         Left            =   120
         TabIndex        =   159
         Top             =   7680
         Width           =   15015
      End
      Begin VB.Frame Frame13 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   161
         Top             =   7680
         Width           =   15015
      End
      Begin VB.Frame Frame14 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   162
         Top             =   7680
         Width           =   15015
      End
      Begin MpUserControls.utxtTextBilling utxtPymChqNum 
         Height          =   420
         Index           =   1
         Left            =   -69720
         TabIndex        =   55
         Top             =   3330
         Width           =   2415
         _ExtentX        =   4260
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
         Alignment       =   1
         MaxLength       =   10
      End
      Begin MpUserControls.utxtTextBilling utxtPymChqNum 
         Height          =   420
         Index           =   2
         Left            =   -69720
         TabIndex        =   58
         Top             =   3810
         Width           =   2415
         _ExtentX        =   4260
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
         Alignment       =   1
         MaxLength       =   10
      End
      Begin VB.Frame Frame7 
         Height          =   7455
         Left            =   -74880
         TabIndex        =   111
         Top             =   480
         Width           =   15015
         Begin MpUserControls.utxtNumBilling utxtTotal 
            Height          =   495
            Left            =   7320
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   873
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "IBM3270 - 1254"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MpUserControls.utxtTextBilling utxtPymChqBnk 
            Height          =   420
            Index           =   0
            Left            =   7680
            TabIndex        =   53
            Top             =   2400
            Width           =   2415
            _ExtentX        =   4260
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
            Alignment       =   1
            MaxLength       =   10
         End
         Begin VB.Label Label44 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Total Amount"
            ForeColor       =   &H00FFFFFF&
            Height          =   450
            Left            =   4800
            TabIndex        =   164
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Label42 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Bank"
            ForeColor       =   &H00FFFFFF&
            Height          =   450
            Left            =   7320
            TabIndex        =   158
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label Label41 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Number"
            ForeColor       =   &H00FFFFFF&
            Height          =   450
            Left            =   4800
            TabIndex        =   157
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Label Label36 
            Caption         =   "Reference"
            Height          =   375
            Left            =   600
            TabIndex        =   119
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label35 
            Caption         =   "Change"
            Height          =   405
            Left            =   1080
            TabIndex        =   118
            Top             =   6480
            Width           =   975
         End
         Begin VB.Label Label34 
            Caption         =   "Adr"
            Height          =   375
            Left            =   1560
            TabIndex        =   117
            Top             =   5880
            Width           =   495
         End
         Begin VB.Label Label33 
            Caption         =   "Cheque"
            Height          =   375
            Left            =   1080
            TabIndex        =   116
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label32 
            Caption         =   "Cash"
            Height          =   375
            Left            =   1440
            TabIndex        =   115
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label31 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Customer Name"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   4800
            TabIndex        =   114
            Top             =   4800
            Width           =   6015
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Customer Code"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2280
            TabIndex        =   113
            Top             =   4800
            Width           =   2415
         End
         Begin VB.Label Label29 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Amount"
            ForeColor       =   &H00FFFFFF&
            Height          =   450
            Left            =   2280
            TabIndex        =   112
            Top             =   1080
            Width           =   2415
         End
      End
      Begin MpUserControls.utxtTextBilling utxtRefTeller 
         Height          =   420
         Left            =   -70680
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   7320
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
      End
      Begin MpUserControls.utxtTextBilling utxtRefDate 
         Height          =   420
         Left            =   -70680
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   6840
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   20
      End
      Begin MpUserControls.utxtNumBilling utxtRefEntno 
         Height          =   420
         Index           =   0
         Left            =   -73080
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtNumBilling utxtRefCCR 
         Height          =   420
         Left            =   -70200
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
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
         Maskformat      =   "00000000"
         Maskformat      =   "00000000"
      End
      Begin MpUserControls.utxtTextBilling utxtRefVessel 
         Height          =   420
         Left            =   -70200
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   12
      End
      Begin MpUserControls.utxtTextBilling utxtRefCommodity 
         Height          =   420
         Left            =   -70200
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   3240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   30
      End
      Begin MpUserControls.utxtTextBilling utxtRefBroker 
         Height          =   420
         Left            =   -70200
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   3720
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   30
      End
      Begin MpUserControls.utxtTextBilling utxtRefExporter 
         Height          =   420
         Left            =   -70200
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   4200
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   741
         BackColor       =   -2147483633
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   30
      End
      Begin MpUserControls.utxtNumBilling utxtRefRefNum 
         Height          =   420
         Left            =   -70200
         TabIndex        =   71
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
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
         MaxLength       =   8
         First           =   -1  'True
      End
      Begin MpUserControls.utxtNumBilling utxtRefSeq 
         Height          =   420
         Left            =   -68640
         TabIndex        =   72
         Top             =   720
         Width           =   735
         _ExtentX        =   1296
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
         MaxLength       =   3
      End
      Begin MpUserControls.utxtNumBilling utxtRefCntno 
         Height          =   420
         Left            =   -69240
         TabIndex        =   74
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
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
         MaxLength       =   8
         Last            =   -1  'True
      End
      Begin MpUserControls.utxtTextBilling utxtRefPrefix 
         Height          =   420
         Left            =   -70200
         TabIndex        =   73
         Top             =   1200
         Width           =   855
         _ExtentX        =   1508
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
         MaxLength       =   4
      End
      Begin VB.Frame Frame16 
         Height          =   7485
         Left            =   -74880
         TabIndex        =   167
         Top             =   480
         Width           =   15015
         Begin VB.Frame Frame18 
            Caption         =   "Entry Numbers"
            Height          =   1575
            Left            =   720
            TabIndex        =   168
            Top             =   4680
            Width           =   9015
         End
         Begin VB.Label Label63 
            Caption         =   "Container Number"
            Height          =   405
            Left            =   720
            TabIndex        =   183
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label Label64 
            Caption         =   "Reference / Sequence"
            Height          =   405
            Left            =   720
            TabIndex        =   177
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label62 
            Alignment       =   2  'Center
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Other Details"
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   176
            Top             =   1320
            Width           =   15135
         End
         Begin VB.Label Label61 
            Caption         =   "CCR Number"
            Height          =   405
            Left            =   720
            TabIndex        =   175
            Top             =   1920
            Width           =   2895
         End
         Begin VB.Label Label60 
            Caption         =   "Vessel"
            Height          =   405
            Left            =   720
            TabIndex        =   174
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label Label59 
            Caption         =   "Commodity"
            Height          =   405
            Left            =   720
            TabIndex        =   173
            Top             =   2880
            Width           =   2895
         End
         Begin VB.Label Label58 
            Caption         =   "Broker"
            Height          =   405
            Left            =   720
            TabIndex        =   172
            Top             =   3360
            Width           =   2895
         End
         Begin VB.Label Label57 
            Caption         =   "Exporter"
            Height          =   405
            Left            =   720
            TabIndex        =   171
            Top             =   3840
            Width           =   2895
         End
         Begin VB.Label Label56 
            Caption         =   "Date/Time Issued"
            Height          =   285
            Left            =   720
            TabIndex        =   170
            Top             =   6480
            Width           =   2895
         End
         Begin VB.Label Label55 
            Caption         =   "Teller"
            Height          =   285
            Left            =   720
            TabIndex        =   169
            Top             =   6960
            Width           =   2895
         End
      End
      Begin VB.Frame Frame15 
         Height          =   1065
         Left            =   -74880
         TabIndex        =   163
         Top             =   7680
         Width           =   15015
      End
      Begin VB.Frame Frame19 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   179
         Top             =   7680
         Width           =   15015
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Container Correction"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -75120
         TabIndex        =   185
         Top             =   0
         Width           =   15375
      End
      Begin VB.Label Label65 
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Refund Payment"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -74880
         TabIndex        =   178
         Top             =   120
         Width           =   15015
      End
      Begin VB.Label Label40 
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CCR Cancellation"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -74880
         TabIndex        =   154
         Top             =   120
         Width           =   15015
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Payment Correction"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -74880
         TabIndex        =   153
         Top             =   60
         Width           =   15015
      End
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   155
      Top             =   9240
      Width           =   15015
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "F3 - EXIT"
      Height          =   615
      Left            =   120
      TabIndex        =   109
      Top             =   9480
      Width           =   2775
   End
   Begin VB.Frame Frame4 
      Height          =   135
      Left            =   120
      TabIndex        =   186
      Top             =   10080
      Width           =   15015
   End
   Begin VB.Label Label43 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CY EXPORT - File Maintenance"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   160
      Top             =   120
      Width           =   15015
   End
End
Attribute VB_Name = "usrctrlCCRde08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim cusListing As cCustomer 'Customer Listing
'**********************
'* StandAlone FLR
'Dim CTCSinfo As Object   'Container Info

'Dim ADR As cADR

' ** ADR Variables
Dim numberin As Boolean
Dim sngPreviousADR As Single
Dim AdrBalance As Single

'  ** Variable Fields
Dim lngRefnum As Long
Dim lngSeqnum As Long
Dim lngItmnum As Long
Dim strCntnum As String
Dim lngCcrnum As Long
Dim lngCntSze As Long
Dim strFulemp As String
Dim strDgrcls As String
Dim strVslcde As String
Dim sngWhfamt As Single
Dim sngArramt As Single
Dim sngOvzamt As Single
Dim sngDgramt As Single
Dim sngArrvat As Single
Dim sngArrtax As Single
Dim strVatcde As String
Dim sngCntovzl As Single
Dim sngCntovzw As Single
Dim sngCntovzh As Single
Dim strOvzums As String
Dim sngRevton As Single
Dim strTrncde As String
Dim strWhfcde As String
Dim strGuarntycde As String
Dim sngDolrte As Single
Dim strExprtr As String
Dim strBroker As String
Dim strEntnum As String
Dim strCommod As String
Dim strRemark As String
Dim strTrknam As String
Dim strPltnum As String
Dim strTrkchs As String
Dim strStatus As String
Dim lngOvrCCr As Long
Dim lngPpanum As Long
Dim strUserId As String
Dim strSysdttm As String
Dim strUpdcde As String
Dim strOutdttm As String
Dim strSBMAPermit As String
Dim sngPrevADR As Single

'  ** Variable Details
Dim lngVarCCRnum As Long
Dim strVarVessel As String
Dim lngVarItmnum As Long
Dim strVarCommodity As String
Dim strVarBroker As String
Dim strVarExporter As String
Dim strVarSBMAPermit As String
Dim strVarEntnum1 As String
Dim strVarEntnum2 As String
Dim strVarEntnum3 As String
Dim strVarEntnum4 As String
Dim strVarEntnum5 As String
Dim strVarEntnum6 As String
Dim strVarEntnum7 As String
Dim strVarEntnum8 As String
Dim strVarEntnum9 As String
Dim strVarEntnum10 As String
Dim strVarExempCode As String
Dim strVarGuarantee As String
Dim strVarDteTme As String
Dim strVarTeller As String


'  ** Payment Details

Dim sngPymCheque As Single
Dim sngPymChange As Single
Dim lngPymADRNum As Long
Dim strPymCustCode As String
Dim strPymCustName As String
Dim sngPymTotalAmt As Single
Dim sngPymAmtPay As Single
Dim sngPymCash As Single
Dim sngPymAdr As Single
Dim sngPymChk1 As Single
Dim sngPymChk2 As Single
Dim sngPymChk3 As Single
Dim sngPymChk4 As Single
Dim sngPymChk5 As Single
Dim strPymChkN1 As String
Dim strPymChkN2 As String
Dim strPymChkN3 As String
Dim strPymChkN4 As String
Dim strPymChkN5 As String
Dim strPymChkB1 As String
Dim strPymChkB2 As String
Dim strPymChkB3 As String
Dim strPymChkB4 As String
Dim strPymChkB5 As String

'   ** Previous Payment Variables
Dim lngPrevCCrNum As Long
Dim lngPrevRefnum As Long
Dim lngPrevDate As Long
Dim lngPrevTime As Long
Dim strPrevTeller As String
Dim strPrevTrncde As String
Dim strPrevExporter As String
Dim strPrevWhfcde As String

'   **  Saving Details Variables
Dim lngSveDetRefnum As Long
Dim lngSveDetSequence As Long
Dim lngSveDetCCRNum As Long
Dim strSveDetVessel As String
Dim strSveDetCntnum As String
Dim strSveDetCommod As String
Dim strSveDetBroker As String
Dim strSveDetExporter As String
Dim strSveDetSBMAPermit As String
Dim strSveDetEntnum As String
Dim strSveDetEntnum1 As String * 8
Dim strSveDetEntnum2 As String * 8
Dim strSveDetEntnum3 As String * 8
Dim strSveDetEntnum4 As String * 8
Dim strSveDetEntnum5 As String * 8
Dim strSveDetEntnum6 As String * 8
Dim strSveDetEntnum7 As String * 8
Dim strSveDetEntnum8 As String * 8
Dim strSveDetEntnum9 As String * 8
Dim strSveDetEntnum10 As String * 8
Dim strSveDetExempCode As String
Dim strSveDetguarantee As String
Dim strSveDetTeller As String
Dim lngSveDetDate As Long
Dim lngSveDetTime As Long

Dim blnTabstop As Boolean
Dim blnPPA As Boolean
Dim blnCancel As Boolean
Dim blnOutdttm As Boolean
Dim blnCntPaid As Boolean
Dim blnRefund As Boolean


'Property Variables:
Dim m_ObjectToUnload As Object

Event Closing()

Private Sub cmdCancel1_Click()
    Call cmdContainer_Click
    utxtRefnum.SetFocus
End Sub

Private Sub cmdCancel2_Click()
    Call cmdDetails_Click
    'utxtDetRefnum.SetFocus
End Sub

Private Sub cmdCancel3_Click()
    Call cmdCancellation_Click
    utxtCnlRefnum.SetFocus
End Sub

Private Sub cmdCancel4_Click()
    Call cmdPayment_Click
    utxtPymReference.SetFocus
End Sub

Private Sub cmdCancel5_Click()
    Call cmdRefund_Click
    utxtRefRefNum.SetFocus
End Sub

Private Sub cmdCancellation_Click()
     With CorrectionTab
        .Visible = True
        .TabEnabled(0) = False
        .TabEnabled(1) = False
        .TabEnabled(2) = True
        .TabEnabled(3) = False
        .TabEnabled(4) = False
        .Tab = 2
    End With
'    utxtCnlRefnum.Value = ""
'    utxtCnlSeqnum.Value = ""
    Call ResetTab02(False)
    Call Tab02(True)
    Call CommandTabstop(False)
End Sub

Private Sub cmdClose1_Click()
    Call ResetTab00(False)
    Call CommandTabstop(True)
    Call IniVariable
    CorrectionTab.Visible = False
End Sub

Private Sub cmdClose2_Click()
    Call ResetTab01(False)
    Call CommandTabstop(True)
    Call IniVariable
    CorrectionTab.Visible = False
End Sub

Private Sub cmdClose3_Click()
    Call ResetTab02(False)
    Call IniVariable
    Call CommandTabstop(True)
    CorrectionTab.Visible = False
End Sub

Private Sub cmdClose4_Click()
    Call ResetTab03(False)
    Call CommandTabstop(True)
    Call IniVariable
    CorrectionTab.Visible = False
End Sub

Private Sub cmdClose5_Click()
    Call ResetTab04(False)
    Call CommandTabstop(True)
    Call IniVariable
    CorrectionTab.Visible = False
End Sub

Private Sub cmdContainer_Click()
    With CorrectionTab
        .Visible = True
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .TabEnabled(3) = False
        .TabEnabled(4) = False
        .Tab = 0
    End With
'    utxtRefnum.Value = ""
'    utxtSeqnum.Value = ""
'    utxtCntPrefix.Text = ""
'    utxtCntNum.Value = ""
    Call ResetTab00(False)
    Call CommandTabstop(False)
    cmdSave1.Enabled = False
    utxtRefnum.SetFocus
    utxtRefnum.TabStop = True
    utxtSeqnum.TabStop = True
    utxtCntPrefix.TabStop = True
    utxtCntNum.TabStop = True
End Sub

Private Sub cmdDetails_Click()
With CorrectionTab
        .Visible = True
        .TabEnabled(0) = False
        .TabEnabled(1) = True
        .TabEnabled(2) = False
        .TabEnabled(3) = False
        .TabEnabled(4) = False
        .Tab = 1
    End With
    Call ResetTab01(False)
    Call Tab01(False)
    Call CommandTabstop(False)
'    utxtDetRefnum.Value = ""
'    utxtDetSeqnum.Value = ""
    cmdSave2.Enabled = False
    utxtDetRefnum.Enabled = True
    utxtDetSeqnum.Enabled = True
    utxtDetRefnum.TabStop = True
    utxtDetSeqnum.TabStop = True
    utxtDetRefnum.SetFocus
   
End Sub

Private Sub cmdExit_Click()
Dim strExit1 As String
Dim strExit2 As String

strExit1 = ""
strExit2 = ""
strResponse = False
strExit1 = "Do you want to Exit from the Program ? "
Call SystemMessage(strExit1, strExit2)
If strResponse = True Then
'**********************
'* StandAlone FLR
'    CTCSinfo.DisConnect
    
    RaiseEvent Closing
End If
End Sub


Private Sub cmdPayment_Click()
    With CorrectionTab
        .Visible = True
        .TabEnabled(0) = False
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .TabEnabled(3) = True
        .TabEnabled(4) = True
        .Tab = 3
    End With
    Call ResetTab03(False)
    Call Tab03(False)
    Call CommandTabstop(False)
    cmdSave4.Enabled = False
    utxtPymAdr.TabStop = False
    utxtPymAdr.Enabled = False
    utxtPymReference.TabStop = True
'    utxtPymReference.Value = ""
    utxtPymReference.SetFocus
End Sub

Private Sub cmdRefund_Click()
With CorrectionTab
        .Visible = True
        .TabEnabled(0) = False
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .TabEnabled(3) = False
        .TabEnabled(4) = True
        .Tab = 4
    End With
    Call ResetTab04(False)
    Call Tab04(False)
    Call CommandTabstop(False)
'    utxtRefRefNum.Value = ""
'    utxtRefSeq.Value = ""
'    utxtRefPrefix.Text = ""
'    utxtRefCntno.Value = ""
    cmdSave5.Enabled = False
    utxtRefRefNum.TabStop = True
    utxtRefRefNum.SetFocus
    utxtRefSeq.TabStop = True
End Sub

Private Sub cmdSave1_Click()
Dim strPref As String * 4
Dim strNum As String * 8
Dim strSaveMess1 As String
Dim strSavemess2 As String
Dim strTmpUser As String
Dim strTmpDate As String

strSaveMess1 = ""
strSavemess2 = ""
strPref = ""
strNum = ""
strTmpUser = ""
strTmpDate = ""
strResponse = False

'If CheckContainer Then
    strSaveMess1 = "Do You Want to Continue Processing ?"
    Call SystemMessage(strSaveMess1, strSavemess2)
    If strResponse = True Then
        
        Call WriteToLog(utxtRefnum.Value, utxtSeqnum.Value, strContainer, _
                                strTmpDate, strTmpUser)
        strPref = ""
        strContainer = ""
        strPref = utxtCntPrefix.Text
        strContainer = strPref & utxtCntNum.Text
       
       Call UpdateCCRContainer(utxtRefnum.Value, utxtSeqnum.Value, _
                                strContainer, strTrncde, strExprtr, strWhfcde)
        strTmpUser = gUserid
        strTmpDate = Format(zGetSysDate, "yyyy-mm-dd hh:nn:ss")
        strPref = ""
        strNum = ""
        strContainer = ""
        strPref = utxtNewCntPrefix.Text
        strNum = utxtNewCntNum.Text
        strContainer = strPref & strNum
        Call WriteToLog(utxtRefnum.Value, utxtSeqnum.Value, strContainer, _
                                strTmpDate, strTmpUser)
        strContainer = ""
        Call cmdContainer_Click
'    Else
'        MsgBox "Save1 cmdNo"
    End If
'Else
'    utxtNewCntPrefix.SetFocus
'End If
End Sub

Private Sub cmdSave2_Click()
Dim lngTempCCR As Long
Dim Message1 As String
Dim Message2 As String
Dim rstCCRNum As Recordset
Dim strTmpUser As String
Dim strTmpDate As String
Dim lngUpdRefno As Long
Dim lngUpdSeqno As Long
Dim lngTempRefnum As Long
Dim lngTempSeqnum As Long

lngTempCCR = 0
Message1 = ""
Message2 = ""
strUserId = ""
strTmpUser = ""
strTmpDate = ""
lngUpdRefno = 0
lngUpdSeqno = 0
lngTempRefnum = 0
lngTempSeqnum = 0

If Len(utxtDetCCRNum.Value) <> 0 Then
    
    lngTempCCR = utxtDetCCRNum.Value
    If lngTempCCR <> lngVarCCRnum Then
            
        Select Case ChkCCR(lngTempCCR)
        Case 0, -2
                Message1 = "CCR No. " & lngTempCCR & " is not Allocated."
                Message2 = "Cannot Continue Saving . . ."
                Call ErrorMessage(Message1, Message2)
                utxtDetCCRNum.Value = lngVarCCRnum
                utxtDetCCRNum.SetFocus
                Exit Sub
        Case -1
                Message1 = "CCR No. " & lngTempCCR & " Already Exist"
                Message2 = "Cannot Continue Saving . . ."
                Call ErrorMessage(Message1, Message2)
                utxtDetCCRNum.Value = lngVarCCRnum
                utxtDetCCRNum.SetFocus
                Exit Sub
        End Select
    End If
        strUserId = Trim(utxtDetTeller1.Text)
        strResponse = False
        Message1 = "Do You Want to Continue Processing ?"
        Call SystemMessage(Message1, Message2)
        If strResponse Then
                lngUpdRefno = utxtDetRefnum.Value
                lngUpdSeqno = utxtDetSeqnum.Value
                strContainer = ""
                Call WriteToLog(utxtDetRefnum.Value, utxtDetSeqnum.Value, "", _
                                        strTmpDate, strTmpUser)
                Call UpdateCCRDetails(lngTempCCR, lngVarCCRnum, lngUpdRefno, lngUpdSeqno, strUserId)
                Call UpdateOVRCCR(lngTempCCR, lngVarCCRnum)
                strTmpUser = gUserid
                strTmpDate = Format(zGetSysDate, "yyyy-mm-dd hh:nn:ss")
                Call WriteToLog(utxtDetRefnum.Value, utxtDetSeqnum.Value, "", _
                                        strTmpDate, strTmpUser)
                Call cmdDetails_Click
        Else
                lngTempRefnum = utxtDetRefnum.Value
                lngTempSeqnum = utxtDetSeqnum.Value
                Call ResetTab01(False)
                Call Tab01(True)
                utxtDetRefnum.TabStop = False
                utxtDetSeqnum.TabStop = False
                utxtDetRefnum.SetFocus
                utxtDetRefnum.Value = lngTempRefnum
                utxtDetSeqnum.Value = lngTempSeqnum
                ChkRefnoSeqno utxtDetRefnum.Value, utxtDetSeqnum.Value
                Call MoveDetailsToDetailFields
                lngTempRefnum = 0
                lngTempSeqnum = 0
                cmdSave2.Enabled = True
                'utxtDetCCRNum.SetFocus
        End If
Else
        Message1 = "Cannot continue saving."
        Message2 = "Complete details to Save."
        Call ErrorMessage(Message1, Message2)
        utxtDetCCRNum.SetFocus
End If
End Sub

Private Sub cmdSave3_Click()
Dim strSaveMess3 As String
Dim strSaveMess4 As String
Dim strTmpDate As String
Dim strTmpUser As String
strSaveMess3 = ""
strSaveMess4 = ""
strTmpDate = ""
strTmpUser = ""
strResponse = False
strSaveMess3 = "Do You Want to Continue Cancellation ?"
Call SystemMessage(strSaveMess3, strSaveMess4)
If strResponse = True Then
    Call WriteToLog(utxtCnlRefnum.Value, utxtCnlSeqnum.Value, "", _
                            strTmpDate, strTmpUser)
    Call UpdateCNLCCR(utxtCnlRefnum.Value, utxtCnlSeqnum.Value)
    Call UpdateCNLPayment(utxtCnlRefnum.Value, utxtCnlSeqnum.Value)
    strTmpUser = gUserid
    strTmpDate = Format(zGetSysDate, "yyyy-mm-dd hh:nn:ss")
    Call WriteToLog(utxtCnlRefnum.Value, utxtCnlSeqnum.Value, "", _
                            strTmpDate, strTmpUser)
    Call cmdCancel3_Click
End If
utxtCnlRefnum.SetFocus
End Sub
Private Sub cmdSave4_Click()
Dim strSaveMess4 As String
Dim strSaveMess5 As String

strSaveMess4 = ""
strSaveMess5 = ""
strResponse = False

    If utxtChange.Value > 0 Or utxtChange.Value = 0 Then
        strSaveMess4 = "Do You Want to Continue Processing ?"
        Call SystemMessage(strSaveMess4, strSaveMess5)
        If strResponse = True Then
            Call UpdatePayment(utxtPymReference.Value)
            Call cmdPayment_Click
        Else
            Call cmdPayment_Click
        End If
    Else
        strSaveMess4 = "Cannot Save Update Payment. Total Amount PAID"
        strSaveMess5 = "is LESS THAN AMOUNT to be paid."
        Call ErrorMessage(strSaveMess4, strSaveMess5)
    End If
    
End Sub

Private Sub Tab00(blnTabstop As Boolean)
    utxtRefnum.TabStop = blnTabstop
    utxtSeqnum.TabStop = blnTabstop
    utxtCntPrefix.TabStop = blnTabstop
    utxtCntNum.TabStop = blnTabstop
    utxtNewCntPrefix.TabStop = blnTabstop
    utxtNewCntNum.TabStop = blnTabstop
End Sub

Private Sub Tab01(blnTabstop As Boolean)
Dim intCtr As Integer

    utxtDetRefnum.TabStop = blnTabstop
    utxtDetSeqnum.TabStop = blnTabstop
    utxtDetCCRNum.TabStop = blnTabstop
    utxtDetVessel.TabStop = blnTabstop
    utxtDetCommodity.TabStop = blnTabstop
    utxtDetBroker.TabStop = blnTabstop
    utxtDetExporter.TabStop = blnTabstop
'    utxtDetSBMAPermit.TabStop = blnTabstop
'    utxtDetWhfcde.TabStop = blnTabstop
    utxtDetGuarantee.TabStop = blnTabstop
    For intCtr = 0 To 9
        utxtDetEntnum(intCtr).TabStop = blnTabstop
        utxtDetEntnum(intCtr).Enabled = blnTabstop
    Next
    utxtDetRefnum.Enabled = blnTabstop
    utxtDetSeqnum.Enabled = blnTabstop
    utxtDetCCRNum.Enabled = blnTabstop
    utxtDetVessel.Enabled = blnTabstop
    utxtDetCommodity.Enabled = blnTabstop
'    utxtDetWhfcde.Enabled = blnTabstop
    utxtDetGuarantee.Enabled = blnTabstop
    utxtDetBroker.Enabled = blnTabstop
    utxtDetExporter.Enabled = blnTabstop
End Sub

Private Sub Tab02(blnTabstop As Boolean)
    utxtCnlRefnum.TabStop = blnTabstop
    utxtCnlSeqnum.TabStop = blnTabstop
End Sub
Private Sub Tab03(blnTabstop As Boolean)
    Dim intCtr As Integer
    utxtPymReference.TabStop = blnTabstop
    utxtPymCash.TabStop = blnTabstop
    For intCtr = 0 To 4
        utxtPymChq(intCtr).TabStop = blnTabstop
        utxtPymChqNum(intCtr).TabStop = blnTabstop
        utxtPymChqBnk(intCtr).TabStop = blnTabstop
    Next
    utxtCustNo.TabStop = blnTabstop
    utxtCustName.TabStop = blnTabstop
End Sub
Private Sub Tab04(blnTabstop As Boolean)
Dim intCtr As Integer
    utxtRefRefNum.TabStop = blnTabstop
    utxtRefSeq.TabStop = blnTabstop
End Sub
Private Sub ResetTab00(blnTabstop As Boolean)
    Call Tab00(blnTabstop)
    Call Tab01(blnTabstop)
    Call Tab02(blnTabstop)
    Call Tab03(blnTabstop)
    Call Tab04(blnTabstop)
'    utxtRefnum.Value = ""
'    utxtSeqnum.Value = ""
'    utxtCntPrefix.Text = ""
'    utxtCntNum.Value = ""
    utxtNewCntPrefix.Text = ""
    utxtNewCntNum.Text = ""
    utxtDateIssue.Text = ""
    utxtTimeIssue.Text = ""
    utxtTeller.Text = ""
End Sub
Private Sub ResetTab01(blnTabstop As Boolean)
Dim intCtr As Integer
    Call Tab00(blnTabstop)
    Call Tab01(blnTabstop)
    Call Tab02(blnTabstop)
    Call Tab03(blnTabstop)
    Call Tab04(blnTabstop)
'    utxtDetRefnum.Value = ""
'    utxtDetSeqnum.Value = ""
    utxtDetCCRNum.Value = ""
    utxtDetVessel.Text = ""
    utxtDetCommodity.Text = ""
    utxtDetBroker.Text = ""
    utxtDetExporter.Text = ""
'    utxtDetSBMAPermit.Text = ""
'    utxtDetWhfcde.Text = ""
    utxtDetGuarantee.Text = ""
    For intCtr = 0 To 9
        utxtDetEntnum(intCtr).Value = ""
    Next
        utxtDetDteTme.Text = ""
    utxtDetTeller1.Text = ""
End Sub
Private Sub ResetTab02(blnTabstop As Boolean)
    Call Tab00(blnTabstop)
    Call Tab01(blnTabstop)
    Call Tab02(blnTabstop)
    Call Tab03(blnTabstop)
    Call Tab04(blnTabstop)
'    utxtCnlRefnum.Value = ""
'    utxtCnlSeqnum.Value = ""
    utxtCnlCCRNum.Value = ""
    utxtCnlVessel.Text = ""
    utxtCnlCommodity.Text = ""
    utxtCnlBroker.Text = ""
    utxtCnlExporter.Text = ""
'    utxtCnlSBMAPermit.Text = ""
    utxtCnlEntnum1.Value = ""
    utxtCnlEntnum2.Value = ""
    utxtCnlEntnum3.Value = ""
    utxtCnlEntnum4.Value = ""
    utxtCnlEntnum5.Value = ""
    utxtCnlEntnum6.Value = ""
    utxtCnlEntnum7.Value = ""
    utxtCnlEntnum8.Value = ""
    utxtCnlEntnum9.Value = ""
    utxtCnlEntnum10.Value = ""
    utxtCnlDteTme.Text = ""
    utxtCnlTeller2.Text = ""
End Sub
Private Sub ResetTab03(blnTabstop As Boolean)
Dim intCtr As Integer
intCtr = 0
    Call Tab00(blnTabstop)
    Call Tab01(blnTabstop)
    Call Tab02(blnTabstop)
    Call Tab03(blnTabstop)
    Call Tab04(blnTabstop)
    utxtPymCash.Value = ".00"
    For intCtr = 0 To 4
        utxtPymChq(intCtr).Value = ".00"
        utxtPymChqNum(intCtr).Text = ""
        utxtPymChqBnk(intCtr).Text = ""
    Next
    strPymCustName = ""
    strPymCustCode = ""
    sngPymAdr = 0
    sngPreviousADR = 0
    AdrBalance = 0
    utxtCustNo.Value = "0"
    utxtCustName.Text = ""
    utxtPymAdr.Value = ".00"
    utxtChange.Value = ".00"
    utxtTotal.Value = ".00"
'    utxtPymReference.Value = ""
    sngPymAmtPay = 0
End Sub
Private Sub ResetTab04(blnTabstop As Boolean)

Dim intCtr As Integer
intCtr = 0
    Call Tab00(blnTabstop)
    Call Tab01(blnTabstop)
    Call Tab02(blnTabstop)
    Call Tab03(blnTabstop)
    Call Tab04(blnTabstop)
'utxtRefRefNum.Value = ""
'utxtRefSeq.Value = ""
'utxtRefPrefix.Text = ""
'utxtRefCntno.Value = ""
utxtRefCCR.Value = ""
utxtRefVessel.Text = ""
utxtRefCommodity.Text = ""
utxtRefBroker.Text = ""
utxtRefExporter.Text = ""
For intCtr = 0 To 9
    utxtRefEntno(intCtr).Value = ""
Next
utxtRefDate.Text = ""
utxtRefTeller.Text = ""
End Sub
Private Sub CommandTabstop(blnTabstop)
    cmdExit.TabStop = blnTabstop
    cmdCancellation.TabStop = blnTabstop
    cmdContainer.TabStop = blnTabstop
    cmdDetails.TabStop = blnTabstop
    cmdPayment.TabStop = blnTabstop
    cmdRefund.TabStop = blnTabstop
    
    cmdExit.Enabled = blnTabstop
    cmdCancellation.Enabled = blnTabstop
    cmdContainer.Enabled = blnTabstop
    cmdDetails.Enabled = blnTabstop
    cmdPayment.Enabled = blnTabstop
    cmdRefund.Enabled = blnTabstop
End Sub

Private Sub cmdSave5_Click()
Dim strSaveMess3 As String
Dim strSaveMess4 As String
Dim strPref As String * 4
Dim strNum As String * 8
Dim strTmpDate As String
Dim strTmpUser As String

strSaveMess3 = ""
strSaveMess4 = ""
strTmpDate = ""
strTmpUser = ""
strResponse = False

strSaveMess3 = "Do You Want to Continue the Refund Transaction?"
Call SystemMessage(strSaveMess3, strSaveMess4)
If strResponse = True Then
        strPref = ""
        strNum = ""
        strContainer = ""
        strPref = Trim(utxtRefPrefix.Text)
        strNum = Trim(utxtRefCntno.Value)
        strContainer = strPref & strNum
        Call WriteToLog(utxtRefRefNum.Value, utxtRefSeq.Value, strContainer, _
                                strTmpDate, strTmpUser)
        Call RefundUpdate
        strTmpUser = gUserid
        strTmpDate = Format(zGetSysDate, "yyyy-mm-dd hh:nn:ss")
        strPref = Trim(utxtRefPrefix.Text)
        strNum = Trim(utxtRefCntno.Value)
        strContainer = strPref & strNum
        Call WriteToLog(utxtRefRefNum.Value, utxtRefSeq.Value, strContainer, _
                                strTmpDate, strTmpUser)
        strContainer = ""
        Call cmdCancel5_Click
End If
utxtRefRefNum.SetFocus
End Sub


Private Sub Label49_Click()

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF2
        If CorrectionTab.Visible = True Then
            If CorrectionTab.Tab = 0 And cmdSave1.Enabled = True Then
                Call cmdSave1_Click
            ElseIf CorrectionTab.Tab = 1 And cmdSave2.Enabled = True Then
                Call cmdSave2_Click
            ElseIf CorrectionTab.Tab = 2 And cmdSave3.Enabled = True Then
                Call cmdSave3_Click
            ElseIf CorrectionTab.Tab = 3 And cmdSave4.Enabled = True Then
                Call cmdSave4_Click
            ElseIf CorrectionTab.Tab = 4 And cmdSave5.Enabled = True Then
                Call cmdSave5_Click
            End If
        End If
    Case vbKeyF3
        If CorrectionTab.Visible = True Then
            If CorrectionTab.Tab = 0 Then
                Call cmdClose1_Click
            ElseIf CorrectionTab.Tab = 1 Then
                Call cmdClose2_Click
            ElseIf CorrectionTab.Tab = 2 Then
                Call cmdClose3_Click
            ElseIf CorrectionTab.Tab = 3 Then
                Call cmdClose4_Click
            ElseIf CorrectionTab.Tab = 4 Then
                Call cmdClose5_Click
            End If
        Else
            Call cmdExit_Click
        End If
    Case vbKeyF4
        If CorrectionTab.Visible = False Then
            Call cmdContainer_Click
        End If
    Case vbKeyF5
        If CorrectionTab.Visible = False Then
            Call cmdDetails_Click
        End If
    Case vbKeyF6
        If CorrectionTab.Visible = False Then
            Call cmdCancellation_Click
        End If
    Case vbKeyF7
        If CorrectionTab.Visible = False Then
            Call cmdPayment_Click
        End If
    Case vbKeyF8
        If CorrectionTab.Visible = False Then
            Call cmdRefund_Click
        End If
    Case vbKeyF12
        If CorrectionTab.Visible = True Then
            If CorrectionTab.Tab = 0 Then
                Call cmdCancel1_Click
                utxtRefnum.SetFocus
            ElseIf CorrectionTab.Tab = 1 Then
                Call cmdCancel2_Click
                'utxtDetRefnum.SetFocus
            ElseIf CorrectionTab.Tab = 2 Then
                Call cmdCancel3_Click
                utxtCnlRefnum.SetFocus
            ElseIf CorrectionTab.Tab = 3 Then
                Call cmdCancel4_Click
                utxtPymReference.SetFocus
            ElseIf CorrectionTab.Tab = 4 Then
                Call cmdCancel5_Click
                utxtRefRefNum.SetFocus
            End If
        End If
End Select
End Sub

'Private Sub UserControl_Terminate()
'    ADR.Disconnect
'    Set ADR = Nothing
'End Sub

Private Sub utxtCnlRefnum_Change()
If Len(utxtCnlRefnum.Value) = 0 Then
    cmdSave3.Enabled = False
    Call cmdCancellation_Click
    utxtCnlRefnum.SetFocus
ElseIf cmdSave3.Enabled = True And Len(utxtCnlRefnum.Value) <> 0 Then
    cmdSave3.Enabled = False
    Call cmdCancellation_Click
    utxtCnlRefnum.SetFocus
End If
End Sub

Private Sub utxtCnlSeqnum_Change()
If cmdSave3.Enabled = True And Len(utxtCnlSeqnum.Value) <> 0 Then
    cmdSave3.Enabled = False
    Call cmdCancellation_Click
    utxtCnlRefnum.SetFocus
ElseIf Len(utxtCnlSeqnum.Value) = 0 Then
    cmdSave3.Enabled = False
    Call cmdCancellation_Click
    utxtCnlRefnum.SetFocus
End If
End Sub

Private Sub utxtCnlSeqnum_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Message1 As String
Dim Message2 As String
Message1 = ""
Message2 = ""

If KeyCode = vbKeyReturn Then
    If Len(utxtCnlRefnum.Value) <> 0 And Len(utxtCnlSeqnum.Value) <> 0 Then
        If CheckReference(CLng(utxtCnlRefnum.Value)) Then
            If Not ChkRefnoSeqno(utxtCnlRefnum.Value, utxtCnlSeqnum.Value) Then
                Message1 = "There is NO TRANSACTION that exist."
                Message2 = "Enter Another transaction."
                Call ErrorMessage(Message1, Message2)
            Else
                If blnOutdttm Then
                    Message1 = "Cannot CANCEL a record NOT IN YARD."
                    Message2 = "Enter Another transaction."
                    Call ErrorMessage(Message1, Message2)
                ElseIf blnCancel Then
                    Message1 = "Cannot CANCEL a CANCELLED transaction."
                    Message2 = "Enter Another transaction."
                    Call ErrorMessage(Message1, Message2)
                ElseIf blnPPA Then
                    Message1 = "Cannot CANCEL current transaction. PPA record exist."
                    Message2 = "Enter Another transaction."
                    Call ErrorMessage(Message1, Message2)
                ElseIf Not blnOutdttm And Not blnCancel And Not blnPPA Then
                    Call MoveDetailsToCancelFields
                End If
            End If
        Else
                Message1 = "Cannot CANCEL current transaction. Transaction is"
                Message2 = "paid with ADR."
                Call ErrorMessage(Message1, Message2)
        End If
    End If
utxtCnlRefnum.SetFocus
End If
End Sub

Private Sub utxtCntNum_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Message1 As String
Dim Message2 As String

Message1 = ""
Message2 = ""

    If KeyCode = vbKeyReturn Then
        If Not ChkRefNoSeqNoCntNum(utxtRefnum.Value, utxtSeqnum.Value, utxtCntPrefix.Text, _
                                                                utxtCntNum.Text) Then
            Message1 = "There is NO TRANSACTION that exist."
            Message2 = "Enter Another transaction."
            Call ErrorMessage(Message1, Message2)
            utxtRefnum.SetFocus
        Else
            If blnOutdttm Then
                Message1 = "Cannot CORRECT a record NOT IN YARD."
                Message2 = "Enter Another transaction."
                Call ErrorMessage(Message1, Message2)
            ElseIf blnCancel Then
                Message1 = "Cannot CORRECT a CANCELLED transaction."
                Message2 = "Enter Another transaction."
                Call ErrorMessage(Message1, Message2)
            End If
            If blnOutdttm Or blnCancel Then
                utxtRefnum.SetFocus
            Else
                Call Tab00(False)
                utxtNewCntPrefix.TabStop = True
                utxtNewCntNum.TabStop = True
'                utxtNewCntPrefix.SetFocus
            End If
        End If
    End If
End Sub

Private Function ChkRefNoSeqNoCntNum(lngRefTemp As Long, lngSeqTemp As Long, _
                         strPrefixTemp As String, lngCntnoTemp As String) As Boolean
Dim rstCCR As Recordset
Dim strPrefix As String * 4
Dim lngRefno As Long
Dim lngSeqno As Long

blnPPA = False
blnOutdttm = False
blnCancel = False
blnRefund = False
ChkRefNoSeqNoCntNum = False
lngRefno = 0
lngSeqno = 0
strContainer = ""
strWhfcde = ""
strPrefix = ""
strTrncde = ""
strExprtr = ""

If Len(lngRefTemp) <> 0 Then
    lngRefno = Trim(lngRefTemp)
End If
If Len(lngSeqTemp) <> 0 Then
    lngSeqno = Trim(lngSeqTemp)
End If
strPrefix = strPrefixTemp
'If Len(strPrefixTemp) <> 0 Or Len(lngCntnoTemp) <> 0 Then
If Len(lngCntnoTemp) <> 0 Then
    strContainer = (strPrefix) & Trim(lngCntnoTemp)
End If
If lngRefno <> 0 And lngSeqno <> 0 And strContainer <> "" Then
    
    DE.RtvContainer lngRefno, lngSeqno, strContainer
    Set rstCCR = DE.rsRtvContainer
    
    With rstCCR
        If .RecordCount <> 0 Then
            ChkRefNoSeqNoCntNum = True
            If .Fields("status") <> "CAN" Then
                If .Fields("ppanum") > 0 Then
                    blnPPA = True
                Else
                    blnPPA = False
                End If
                If .Fields("updcde") = "R" Then
                    blnRefund = True
                End If
                If IsNull(.Fields("outdttm")) Then
                    Call ClearVariableDetails
                    Call MoveCCRDetailsToVariables(rstCCR)
                    utxtDateIssue.Text = Format(strSysdttm, "yyyy-mm-dd")
                    utxtTimeIssue.Text = Format(strSysdttm, "Hh:Nn")
                    utxtTeller.Text = strUserId
                Else
                    blnOutdttm = True
                End If
                
            Else
                blnCancel = True
            End If
        End If
    End With
rstCCR.Close
Set rstCCR = Nothing
End If

End Function

Private Function ChkRefnoSeqno(ByVal lngTempReference As Long, ByVal lngTempSequence As Long) As Boolean
Dim rstCCRRefSeq As Recordset
Dim lngRefno As Long
Dim lngSeqno As Long

blnPPA = False
blnCancel = False
blnOutdttm = False
ChkRefnoSeqno = False
lngRefno = 0
lngSeqno = 0
strEntnum = ""

If Len(lngTempReference) <> 0 Then
    lngRefno = Trim(lngTempReference) 'Trim(utxtCnlRefnum.Value)
End If
If Len(lngTempSequence) <> 0 Then
    lngSeqno = Trim(lngTempSequence) 'Trim(utxtCnlSeqnum.Value)
End If
If lngRefno <> 0 And lngSeqno <> 0 Then

    DE.RtvDetails lngRefno, lngSeqno
    Set rstCCRRefSeq = DE.rsRtvDetails
    
    With rstCCRRefSeq
        If .RecordCount <> 0 Then
            ChkRefnoSeqno = True
            If .Fields("status") <> "CAN" Then
                If .Fields("ppanum") = 0 Then
                    If IsNull(.Fields("outdttm")) Then
                        Call ClearScreenVariableDetails
                        Call MoveDetailsToVariables(rstCCRRefSeq)
                        cmdSave3.Enabled = True
                    Else
                        blnOutdttm = True
                    End If
                Else
                    blnPPA = True
                End If
            Else
                blnCancel = True
            End If
        End If
    End With
rstCCRRefSeq.Close
Set rstCCRRefSeq = Nothing
End If

End Function

Private Sub UpdateCNLCCR(lngRefno, lngSeqno)
Dim rstCCRDet As Recordset
Dim lngReference As Long
Dim lngSequence As Long
Dim lngCnlDate As Long
Dim lngCnlTime As Long
Dim strCnlTeller As String
Dim strCnlTrncde As String
Dim strCnlExporter As String
Dim strCnlWhfcde As String
Dim strCntnum As String * 12
Const strStatus = "CAN"

DE.RtvDetails lngRefno, lngSeqno
Set rstCCRDet = DE.rsRtvDetails
With rstCCRDet
    Do Until .EOF
        lngReference = 0
        lngSequence = 0
        lngCnlDate = 0
        lngCnlTime = 0
        strCntnum = ""
        strCnlTeller = ""
        strCnlTrncde = ""
        strCnlExporter = ""
        strCnlWhfcde = ""
        Call ClearVariableDetails
        Call MoveCCRDetailsToVariables(rstCCRDet)
        lngReference = ReturnLong(CStr(lngRefnum))
        lngSequence = ReturnLong(CStr(lngSeqnum))
        strCntnum = Trim(.Fields("cntnum"))
        lngCnlDate = ReturnLong(CStr(Trim(Format(strSysdttm, "yyyymmdd"))))
        lngCnlTime = ReturnLong(CStr(Trim(Format(strSysdttm, "hhnnss"))))
        strCnlTeller = Trim(strUserId)
        strCnlTrncde = Trim(strTrncde)
        strCnlExporter = Trim(strExprtr)
        strCnlWhfcde = Trim(strWhfcde)
        'write to logfile
        DE.UpdateCCR strStatus, lngReference, lngSequence, strCntnum
         '   **  Delete Previous Container in File Expor21

'**********************
'* StandAlone FLR
'        CTCSinfo.DelCYExport strCntnum
'        If PreviousPayment(strCntnum) = True Then
'            CTCSinfo.WriteCYExport strCntnum, lngPrevRefnum, lngPrevDate, lngPrevTime, strPrevTeller, strPrevExporter, strPrevTrncde, strPrevWhfcde
'        End If
'**********************
        
        'write to logfile
        .MoveNext
    Loop
End With
rstCCRDet.Close
Set rstCCRDet = Nothing
End Sub

Private Sub UpdateCCRContainer(lngRefno As Long, lngSeqno As Long, strContNum As String, _
                                                    strTrncde As String, strExporter As String, strWhfcde As String)
Dim strNewContnum As String * 12
Dim strUpdateCode As String * 1
Dim lngCYRefnum As Long
Dim lngCYDate As Long
Dim lngCYTime As Long
Dim strCYTeller As String * 10
Dim strCYTrncde As String
Dim strCYExporter As String
Dim strCYWhfcde As String
Dim strPref1 As String * 4

strNewContnum = ""
strUpdateCode = ""
strUpdateCode = "U"
lngCYRefnum = 0
lngCYDate = 0
lngCYTime = 0
strCYTeller = ""
strCYTrncde = ""
strCYExporter = ""
strCYWhfcde = ""
strPref1 = ""

lngCYRefnum = utxtRefnum.Value
lngCYDate = Format(utxtDateIssue.Text, "yyyymmdd")
lngCYTime = Format(utxtTimeIssue.Text, "HHNNss")
strCYTeller = Trim(utxtTeller.Text)
strCYTrncde = Trim(strTrncde)
strCYExporter = Trim(strExporter)
strCYWhfcde = Trim(strWhfcde)
strPref1 = utxtNewCntPrefix.Text
strNewContnum = strPref1 & utxtNewCntNum.Text
 '   **  Delete Previous Container in File Expor21
'**********************
'* StandAlone FLR
'CTCSinfo.DelCYExport strContNum

If PreviousPayment(strNewContnum) = True Then
    DE.UpdateCntNum strNewContnum, strUpdateCode, lngPrevCCrNum, lngRefno, lngSeqno, strContNum
    '   **  Delete Existing New Container in File Expor21
'**********************
'* StandAlone FLR
'    CTCSinfo.DelCYExport strNewContnum
Else
    DE.UpdateCntNum strNewContnum, strUpdateCode, lngPrevCCrNum, lngRefno, lngSeqno, strContNum
End If
'**********************
'* StandAlone FLR
'If PreviousPayment(strContNum) = True Then
'    CTCSinfo.WriteCYExport strContNum, lngPrevRefnum, lngPrevDate, lngPrevTime, strPrevTeller, strPrevExporter, strPrevTrncde, strPrevWhfcde
'End If
'CTCSinfo.WriteCYExport strNewContnum, lngCYRefnum, lngCYDate, lngCYTime, strCYTeller, strCYExporter, strCYTrncde, strCYWhfcde

End Sub

Private Sub UpdateCCRDetails(lngCCR As Long, lngOldCCR As Long, lngUpdSveRefno As Long, lngUpdSveSeqno As Long, StrUser As String)
Dim rstCCRDetails As Recordset

Const strSveDetUpdte = "U"

'  ** Update Details Variables

lngSveDetSequence = 0
Call ClearDetails
If lngCCR <> lngOldCCR Then
    DE.UpdateCCRAlloc lngCCR, StrUser
End If
 
DE.RtvDetails lngUpdSveRefno, lngUpdSveSeqno
    Set rstCCRDetails = DE.rsRtvDetails
With rstCCRDetails
    Do Until .EOF
        Call ClearDetails
        Call ClearVariableDetails
        Call MoveCCRDetailsToVariables(rstCCRDetails)
        lngSveDetRefnum = ReturnLong(CStr(lngRefnum))
        lngSveDetSequence = ReturnLong(CStr(lngSeqnum))
        lngSveDetCCRNum = ReturnLong(CStr(lngCCR))
        strSveDetCntnum = Trim(strCntnum)
        strSveDetVessel = Trim(utxtDetVessel.Text) & ""
        strSveDetCommod = Trim(utxtDetCommodity.Text) & ""
        strSveDetBroker = Trim(utxtDetBroker.Text) & ""
        strSveDetExporter = Trim(utxtDetExporter.Text) & ""
'        strSveDetSBMAPermit = Trim(utxtDetSBMAPermit.Text) & ""
        lngSveDetDate = ReturnLong(CStr(Format(.Fields("sysdttm"), "YYYYMMDD")))
        lngSveDetTime = ReturnLong(CStr(Format(.Fields("sysdttm"), "HHNNSS")))
'        strSveDetExempCode = Trim(utxtDetWhfcde.Text)
        strSveDetguarantee = Trim(utxtDetGuarantee.Text)
        strSveDetTeller = strUserId
        strSveDetEntnum = ""
        If Len(utxtDetEntnum(0).Value) <> 0 Then
            strSveDetEntnum1 = utxtDetEntnum(0).Value
            strSveDetEntnum = strSveDetEntnum & strSveDetEntnum1
        End If
        If Len(utxtDetEntnum(1).Value) <> 0 Then
            strSveDetEntnum2 = utxtDetEntnum(1).Value
            strSveDetEntnum = strSveDetEntnum & strSveDetEntnum2
        End If
        If Len(utxtDetEntnum(2).Value) <> 0 Then
            strSveDetEntnum3 = utxtDetEntnum(2).Value
            strSveDetEntnum = strSveDetEntnum & strSveDetEntnum3
        End If
        If Len(utxtDetEntnum(3).Value) <> 0 Then
            strSveDetEntnum4 = utxtDetEntnum(3).Value
            strSveDetEntnum = strSveDetEntnum & strSveDetEntnum4
        End If
        If Len(utxtDetEntnum(4).Value) <> 0 Then
            strSveDetEntnum5 = utxtDetEntnum(4).Value
            strSveDetEntnum = strSveDetEntnum & strSveDetEntnum5
        End If
        If Len(utxtDetEntnum(5).Value) <> 0 Then
            strSveDetEntnum6 = utxtDetEntnum(5).Value
            strSveDetEntnum = strSveDetEntnum & strSveDetEntnum6
        End If
        If Len(utxtDetEntnum(6).Value) <> 0 Then
            strSveDetEntnum7 = utxtDetEntnum(6).Value
            strSveDetEntnum = strSveDetEntnum & strSveDetEntnum7
        End If
        If Len(utxtDetEntnum(7).Value) <> 0 Then
            strSveDetEntnum8 = utxtDetEntnum(7).Value
            strSveDetEntnum = strSveDetEntnum & strSveDetEntnum8
        End If
        If Len(utxtDetEntnum(8).Value) <> 0 Then
            strSveDetEntnum9 = utxtDetEntnum(8).Value
            strSveDetEntnum = strSveDetEntnum & strSveDetEntnum9
        End If
        If Len(utxtDetEntnum(9).Value) <> 0 Then
            strSveDetEntnum10 = utxtDetEntnum(9).Value
            strSveDetEntnum = strSveDetEntnum & strSveDetEntnum10
        End If
        
        DE.UpdateDtlCCR lngCCR, strSveDetVessel, strSveDetCommod, strSveDetBroker, strSveDetExporter, _
                           strSveDetEntnum, strSveDetUpdte, strSveDetExempCode, strSveDetguarantee, _
                           lngSveDetRefnum, lngSveDetSequence
'**********************
'* StandAlone FLR
'        CTCSinfo.DelCYExport strSveDetCntnum
'        CTCSinfo.WriteCYExport strSveDetCntnum, lngSveDetRefnum, lngSveDetDate, _
'                        lngSveDetTime, strSveDetTeller, _
'                        strSveDetExporter, strPrevTrncde, strPrevWhfcde
'**********************
        
        'write to logfile
        .MoveNext
    Loop
End With
rstCCRDetails.Close
Set rstCCRDetails = Nothing

End Sub


Private Sub MoveCCRDetailsToVariables(PassRecordset As Recordset)
With PassRecordset
    lngRefnum = ReturnLong(CStr(.Fields("refnum")))
    lngSeqnum = ReturnLong(CStr(.Fields("seqnum")))
    lngItmnum = ReturnLong(CStr(.Fields("itmnum")))
    strCntnum = .Fields("cntnum") & ""
    lngCcrnum = ReturnLong(CStr(.Fields("ccrnum")))
    lngCntSze = ReturnLong(CStr(.Fields("cntsze")))
    strFulemp = .Fields("fulemp") & ""
    strDgrcls = .Fields("dgrcls") & ""
    strVslcde = .Fields("vslcde") & ""
    sngWhfamt = ReturnSingle(CStr(.Fields("whfamt")))
    sngArramt = ReturnSingle(CStr(.Fields("arramt")))
    sngOvzamt = ReturnSingle(CStr(.Fields("ovzamt")))
    sngDgramt = ReturnSingle(CStr(.Fields("dgramt")))
    sngArrvat = ReturnSingle(CStr(.Fields("arrvat")))
    sngArrtax = ReturnSingle(CStr(.Fields("arrtax")))
    strVatcde = .Fields("vatcde") & ""
    sngCntovzl = ReturnSingle(CStr(.Fields("cntovzl")))
    sngCntovzw = ReturnSingle(CStr(.Fields("cntovzw")))
    sngCntovzh = ReturnSingle(CStr(.Fields("cntovzh")))
    strOvzums = .Fields("ovzums") & ""
    sngRevton = ReturnSingle(CStr(.Fields("revton")))
    strTrncde = .Fields("trncde") & ""
    strWhfcde = .Fields("whfcde") & ""
    strGuarntycde = .Fields("guarntycde") & ""
    sngDolrte = ReturnSingle(CStr(.Fields("dolrte")))
    strExprtr = .Fields("exprtr") & ""
    strBroker = .Fields("broker") & ""
    strEntnum = .Fields("entnum") & ""
    strCommod = .Fields("commod") & ""
    strRemark = .Fields("remark") & ""
    strTrknam = .Fields("trknam") & ""
    strPltnum = .Fields("pltnum") & ""
    strTrkchs = .Fields("trkchs") & ""
    strStatus = .Fields("status") & ""
    lngOvrCCr = ReturnLong(CStr(.Fields("ovrccr")))
    lngPpanum = ReturnLong(CStr(.Fields("ppanum")))
    strUserId = .Fields("userid") & ""
    strSysdttm = .Fields("sysdttm") & ""
    strUpdcde = .Fields("updcde") & ""
    strOutdttm = .Fields("outdttm") & ""
End With
End Sub
Private Sub MoveSaveData(tmpRecordset As Recordset, tmpDate As String, tmpUserid As String)
'Dim rstCCRCyxz As Recordset
'Set rstCCRCyxz = New ADODB.Recordset
'    rstCCRCyxz.CursorType = adOpenDynamic
'    rstCCRCyxz.LockType = adLockOptimistic
'    rstCCRCyxz.Open "CCRcyxz", gcnnBilling, , , adCmdTable
'
With tmpRecordset
    .AddNew
    .Fields("refnum") = ReturnLong(CStr(lngRefnum))
    .Fields("seqnum") = ReturnLong(CStr(lngSeqnum))
    .Fields("itmnum") = ReturnLong(CStr(lngItmnum))
    .Fields("cntnum") = strCntnum
    .Fields("ccrnum") = ReturnLong(CStr(lngCcrnum))
    .Fields("cntsze") = ReturnLong(CStr(lngCntSze))
    .Fields("fulemp") = strFulemp
    .Fields("dgrcls") = strDgrcls
    .Fields("vslcde") = strVslcde
    .Fields("whfamt") = ReturnSingle(CStr(sngWhfamt))
    .Fields("arramt") = ReturnSingle(CStr(sngArramt))
    .Fields("ovzamt") = ReturnSingle(CStr(sngOvzamt))
    .Fields("dgramt") = ReturnSingle(CStr(sngDgramt))
    .Fields("arrvat") = ReturnSingle(CStr(sngArrvat))
    .Fields("arrtax") = ReturnSingle(CStr(sngArrtax))
    .Fields("vatcde") = strVatcde
    .Fields("cntovzl") = ReturnSingle(CStr(sngCntovzl))
    .Fields("cntovzw") = ReturnSingle(CStr(sngCntovzw))
    .Fields("cntovzh") = ReturnSingle(CStr(sngCntovzh))
    .Fields("ovzums") = strOvzums
    .Fields("revton") = ReturnSingle(CStr(sngRevton))
    .Fields("trncde") = strTrncde
    .Fields("whfcde") = strWhfcde
    .Fields("guarntycde") = strGuarntycde
    .Fields("dolrte") = ReturnSingle(CStr(sngDolrte))
    .Fields("exprtr") = Trim(strExprtr)
    .Fields("broker") = Trim(strBroker)
    .Fields("entnum") = strEntnum
    .Fields("commod") = strCommod
    .Fields("remark") = strRemark
    .Fields("trknam") = strTrknam
    .Fields("pltnum") = strPltnum
    .Fields("trkchs") = strTrkchs
    .Fields("status") = strStatus
    .Fields("ovrccr") = ReturnLong(CStr(lngOvrCCr))
    .Fields("ppanum") = ReturnLong(CStr(lngPpanum))
    If Len(tmpUserid) > 0 Then
        .Fields("userid") = UCase(Trim(tmpUserid))
        .Fields("sysdttm") = Format(Trim(tmpDate), "YYYY-MM-DD HH:NN:SS")
    Else
        .Fields("userid") = Trim(strUserId)
        .Fields("sysdttm") = Format(strSysdttm, "yyyy-mm-dd HH:NN.SS")
    End If
    .Fields("updcde") = strUpdcde
    If Len(strOutdttm) > 0 Then
        .Fields("outdttm") = Trim(strOutdttm)
    End If
    tmpRecordset.Update
End With
'rstCCRCyxz.Close
'Set rstCCRCyxz = Nothing
End Sub
Private Sub MoveSavePaymnt(tmpRecordset As Recordset, tmpDate As String, tmpUserid As String)
With tmpRecordset
    .AddNew
    If Len(tmpUserid) > 0 Then
        .Fields("userid") = UCase(Trim(tmpUserid))
        .Fields("sysdttm") = Format(Trim(tmpDate), "YYYY-MM-DD HH:NN:SS")
    Else
        .Fields("userid") = Trim(strUserId)
        .Fields("sysdttm") = Format(strSysdttm, "yyyy-mm-dd HH:NN.SS")
    End If
    tmpRecordset.Update
End With
End Sub

Public Function zGetSysDate() As Date
Dim rsDate As Recordset
DE.getdate

' ** Return Recordset
Set rsDate = DE.rsGetDate
    With rsDate
        zGetSysDate = .Fields(0)
        .Close
    End With
Set rsDate = Nothing
End Function
Private Sub ErrorMessage(strMessage1 As String, strMessage2 As String)
    Beep
    frmMessPYX3.lblMessPYX31.Caption = strMessage1
    frmMessPYX3.lblMessPYX32.Caption = strMessage2
    frmMessPYX3.Timer1 = True
    frmMessPYX3.Show (vbModal)
End Sub

Private Sub utxtCntPrefix_Change()
If Len(utxtCntPrefix.Text) = 4 Then
    utxtCntNum.SetFocus
End If
End Sub

Private Sub utxtCustName_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub utxtCustNo_Change()
If Len(Trim(utxtCustNo.Value)) <> 0 Then
    DE.GetADRBal Trim(utxtCustNo.Value), AdrBalance
    DE.getCustomerName Trim(utxtCustNo.Value), strPymCustName
    utxtCustName.Text = strPymCustName
    
        If AdrBalance <> 0 Then
            utxtPymAdr.TabStop = True
            utxtPymAdr.Enabled = True
            utxtPymAdr.SetFocus
            cmdSave4.Enabled = False
        Else
            If ReturnSingle(CStr(utxtPymAdr.Value)) = 0 Then
                utxtPymAdr.Value = ".00"
                utxtPymAdr.TabStop = False
                utxtPymAdr.Enabled = False
            End If
        End If
Else
    AdrBalance = 0
    strPymCustName = ""
    utxtPymAdr.Value = ".00"
    utxtPymAdr.TabStop = False
    utxtPymAdr.Enabled = False
    utxtCustName.Text = ""
    cmdSave4.Enabled = True
End If
End Sub

Private Sub utxtCustNo_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF4
        If cusListing.ShowList Then
'            DE.GetADRBal Trim(utxtCustNo.Value), AdrBalance
            DE.GetADRBal cusListing.Code, AdrBalance
            utxtCustNo.Value = cusListing.Code
            strPymCustName = cusListing.Name
            utxtCustName.Text = strPymCustName
            If AdrBalance <> 0 Then
                utxtPymAdr.TabStop = True
                utxtPymAdr.Enabled = True
            Else
                utxtPymAdr.TabStop = False
                utxtPymAdr.Enabled = False
            End If
        End If
    Case 37 'Left
        If AdrBalance > 0 Then
            utxtPymAdr.SetFocus
        End If
End Select
End Sub


Private Sub utxtDetGuarantee_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    utxtDetEntnum(0).SetFocus
End If
End Sub

Private Sub utxtDetGuarantee_KeyPress(KeyAscii As Integer)
If UCase(Chr(KeyAscii)) <> "Y" And UCase(Chr(KeyAscii)) <> "N" Then
    KeyAscii = 0
End If
End Sub

Private Sub utxtDetSeqnum_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Message1 As String
Dim Message2 As String

Message1 = ""
Message2 = ""
If KeyCode = vbKeyReturn Then
    If Len(utxtDetRefnum.Value) <> 0 And Len(utxtDetSeqnum.Value) <> 0 Then
        If Not ChkRefnoSeqno(utxtDetRefnum.Value, utxtDetSeqnum.Value) Then
            Message1 = "There is NO TRANSACTION that exist."
            Message2 = "Enter Another transaction."
            Call ErrorMessage(Message1, Message2)
            utxtDetRefnum.SetFocus
        Else
            If blnOutdttm Then
                Message1 = "Cannot CORRECT a record NOT IN YARD."
                Message2 = "Enter Another transaction."
                Call ErrorMessage(Message1, Message2)
                utxtDetRefnum.SetFocus
            ElseIf blnCancel Then
                Message1 = "Cannot CORRECT a CANCELLED transaction."
                Message2 = "Enter Another transaction."
                Call ErrorMessage(Message1, Message2)
                utxtDetRefnum.SetFocus
            ElseIf blnPPA Then
                Message1 = "Cannot CORRECT current transaction. PPA record exist."
                Message2 = "Enter Another transaction."
                Call ErrorMessage(Message1, Message2)
                utxtDetRefnum.SetFocus
            ElseIf Not blnOutdttm And Not blnCancel And Not blnPPA Then
                Call MoveDetailsToDetailFields
                Call Tab01(True)
                utxtDetRefnum.TabStop = False
                utxtDetSeqnum.TabStop = False
                utxtDetRefnum.Enabled = False
                utxtDetSeqnum.Enabled = False
'                utxtDetCCRNum.SetFocus
                cmdSave2.Enabled = True
            End If
        End If
    End If

End If
End Sub

Private Sub MoveDetailsToVariables(PassedRcrdSet As Recordset)
Dim strEntnum As String
strEntnum = ""

With PassedRcrdSet
    strEntnum = Trim(.Fields("entnum")) & ""
    lngVarCCRnum = ReturnLong(CStr(.Fields("ccrnum")))
    lngVarItmnum = ReturnLong(CStr(.Fields("itmnum")))
    strVarVessel = .Fields("vslcde") & ""
    strVarCommodity = .Fields("commod") & ""
    strVarBroker = .Fields("broker") & ""
    strVarExporter = .Fields("exprtr") & ""
'    strVarSBMAPermit = .Fields("") & ""
    strVarEntnum1 = Mid(strEntnum, 1, 8)
    strVarEntnum2 = Mid(strEntnum, 9, 8)
    strVarEntnum3 = Mid(strEntnum, 17, 8)
    strVarEntnum4 = Mid(strEntnum, 25, 8)
    strVarEntnum5 = Mid(strEntnum, 33, 8)
    strVarEntnum6 = Mid(strEntnum, 41, 8)
    strVarEntnum7 = Mid(strEntnum, 49, 8)
    strVarEntnum8 = Mid(strEntnum, 57, 8)
    strVarEntnum9 = Mid(strEntnum, 65, 8)
    strVarEntnum10 = Mid(strEntnum, 73, 8)
    strVarExempCode = Trim(.Fields("whfcde"))
    strVarGuarantee = Trim(.Fields("guarntycde"))
    strVarDteTme = Format(.Fields("sysdttm"), "yyyy-mm-dd Hh:Nn")
    strVarTeller = .Fields("userid")
End With
End Sub
Private Sub MoveDetailsToRefundFields()
    utxtRefCCR.Value = ReturnLong(CStr(lngCcrnum))
    utxtRefVessel.Text = Trim(strVslcde)
    utxtRefCommodity.Text = Trim(strCommod)
    utxtRefBroker.Text = Trim(strBroker)
    utxtRefExporter.Text = Trim(strExprtr)
    utxtRefEntno(0).Value = Mid(strEntnum, 1, 8)
    utxtRefEntno(1).Value = Mid(strEntnum, 9, 8)
    utxtRefEntno(2).Value = Mid(strEntnum, 17, 8)
    utxtRefEntno(3).Value = Mid(strEntnum, 25, 8)
    utxtRefEntno(4).Value = Mid(strEntnum, 33, 8)
    utxtRefEntno(5).Value = Mid(strEntnum, 41, 8)
    utxtRefEntno(6).Value = Mid(strEntnum, 49, 8)
    utxtRefEntno(7).Value = Mid(strEntnum, 57, 8)
    utxtRefEntno(8).Value = Mid(strEntnum, 65, 8)
    utxtRefEntno(9).Value = Mid(strEntnum, 73, 8)
    utxtRefDate.Text = strSysdttm
    utxtRefTeller.Text = Trim(strUserId)
End Sub
Private Sub MoveDetailsToCancelFields()
    utxtCnlCCRNum.Value = ReturnLong(CStr(lngVarCCRnum))
    utxtCnlVessel.Text = Trim(strVarVessel)
    utxtCnlCommodity.Text = Trim(strVarCommodity)
    utxtCnlBroker.Text = Trim(strVarBroker)
    utxtCnlExporter.Text = Trim(strVarExporter)
'    utxtCnlSBMAPermit.Text = Trim(strVarSBMAPermit)
    utxtCnlEntnum1.Value = strVarEntnum1
    utxtCnlEntnum2.Value = strVarEntnum2
    utxtCnlEntnum3.Value = strVarEntnum3
    utxtCnlEntnum4.Value = strVarEntnum4
    utxtCnlEntnum5.Value = strVarEntnum5
    utxtCnlEntnum6.Value = strVarEntnum6
    utxtCnlEntnum7.Value = strVarEntnum7
    utxtCnlEntnum8.Value = strVarEntnum8
    utxtCnlEntnum9.Value = strVarEntnum9
    utxtCnlEntnum10.Value = strVarEntnum10
    utxtCnlDteTme.Text = strVarDteTme
    utxtCnlTeller2.Text = Trim(strVarTeller)
End Sub
Private Sub ClearScreenVariableDetails()
    lngVarCCRnum = 0
    lngVarItmnum = 0
    strVarVessel = ""
    strVarCommodity = ""
    strVarBroker = ""
    strVarExporter = ""
    strVarSBMAPermit = ""
    strVarEntnum1 = ""
    strVarEntnum2 = ""
    strVarEntnum3 = ""
    strVarEntnum4 = ""
    strVarEntnum5 = ""
    strVarEntnum6 = ""
    strVarEntnum7 = ""
    strVarEntnum8 = ""
    strVarEntnum9 = ""
    strVarEntnum10 = ""
    strVarDteTme = ""
    strVarTeller = ""
    strVarExempCode = ""
    strVarGuarantee = ""
End Sub
Private Sub MoveDetailsToDetailFields()
    utxtDetRefnum.Value = ReturnLong(CStr(Format(utxtDetRefnum.Value, "00000000")))
    utxtDetSeqnum.Value = ReturnLong(CStr(Format(utxtDetSeqnum.Value, "000")))
    utxtDetCCRNum.Value = ReturnLong(CStr(lngVarCCRnum))
    utxtDetVessel.Text = strVarVessel
    utxtDetCommodity.Text = strVarCommodity
    utxtDetBroker.Text = strVarBroker
    utxtDetExporter.Text = strVarExporter
'    utxtDetSBMAPermit.Text = strVarSBMAPermit
    utxtDetEntnum(0).Value = strVarEntnum1
    utxtDetEntnum(1).Value = strVarEntnum2
    utxtDetEntnum(2).Value = strVarEntnum3
    utxtDetEntnum(3).Value = strVarEntnum4
    utxtDetEntnum(4).Value = strVarEntnum5
    utxtDetEntnum(5).Value = strVarEntnum6
    utxtDetEntnum(6).Value = strVarEntnum7
    utxtDetEntnum(7).Value = strVarEntnum8
    utxtDetEntnum(8).Value = strVarEntnum9
    utxtDetEntnum(9).Value = strVarEntnum10
'    utxtDetWhfcde.Text = strVarExempCode
    utxtDetGuarantee.Text = strVarGuarantee
    utxtDetDteTme.Text = strVarDteTme
    utxtDetTeller1.Text = strVarTeller
End Sub
Private Sub ClearVariableDetails()
    lngRefnum = 0
    lngSeqnum = 0
    lngItmnum = 0
    strCntnum = ""
    lngCcrnum = 0
    lngCntSze = 0
    strFulemp = ""
    strDgrcls = ""
    strVslcde = ""
    sngWhfamt = 0
    sngArramt = 0
    sngOvzamt = 0
    sngDgramt = 0
    sngArrvat = 0
    sngArrtax = 0
    strVatcde = ""
    sngCntovzl = 0
    sngCntovzw = 0
    sngCntovzh = 0
    strOvzums = ""
    sngRevton = 0
    strTrncde = ""
    strWhfcde = ""
    strGuarntycde = ""
    sngDolrte = 0
    strExprtr = ""
    strBroker = ""
    strEntnum = ""
    strCommod = ""
    strRemark = ""
    strTrknam = ""
    strPltnum = ""
    strTrkchs = ""
    strStatus = ""
    lngOvrCCr = 0
    lngPpanum = 0
    strUserId = ""
    strSysdttm = ""
    strUpdcde = ""
    strOutdttm = ""
End Sub
Private Sub utxtDetWhfcde_KeyPress(KeyAscii As Integer)
'If Not UCase(Chr(KeyAscii)) = "0" And _
'        Not UCase(Chr(KeyAscii)) = "1" And _
'        Not UCase(Chr(KeyAscii)) = "2" And _
'        Not UCase(Chr(KeyAscii)) = "3" And _
'        Not UCase(Chr(KeyAscii)) = "4" And _
'        Not UCase(Chr(KeyAscii)) = "5" And _
'            KeyAscii <> 8 And KeyAscii <> 13 Then
'            Beep
'            KeyAscii = 0
'End If
End Sub
Private Sub utxtNewCntNum_Change()
'If Len(utxtNewCntPrefix.Text) = 4 Then
'    utxtNewCntNum.SetFocus
    If Len(utxtNewCntPrefix.Text) <> 0 Or Len(utxtNewCntNum.Text) <> 0 Then
        cmdSave1.Enabled = True
    Else
        cmdSave1.Enabled = False
        utxtNewCntPrefix.SetFocus
    End If
'Else
'    utxtNewCntPrefix.SetFocus
'End If
End Sub
Private Sub utxtNewCntNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        utxtNewCntPrefix.SetFocus
    End If
End Sub
Private Sub utxtNewCntPrefix_Change()
'If Len(utxtNewCntPrefix.Text) = 4 Then
    If Len(utxtNewCntPrefix.Text) <> 0 And Len(utxtNewCntNum.Text) <> 0 Then
        cmdSave1.Enabled = True
    Else
        cmdSave1.Enabled = False
        utxtNewCntPrefix.SetFocus
    End If
'End If
End Sub

Private Sub utxtPymAdr_Change()
Dim strAdrMsg1 As String
Dim strAdrMsg2 As String
Dim intResponse As Integer

If numberin Then
    If IsNumeric(utxtPymAdr.Value) Then
        If CSng(utxtPymAdr.Value) > AdrBalance Then
            utxtPymAdr.Value = Trim(utxtPymAdr.Value)
            utxtCustName.Text = Trim(utxtCustName.Text)
            strAdrMsg1 = utxtPymAdr.Value & " is greater than the Current ADR Running Balance for Customer "
            strAdrMsg2 = utxtCustName.Text
            intResponse = MsgBox(strAdrMsg1 & strAdrMsg2, vbOKOnly, "ADR Amount Error Message")
            utxtPymAdr.Value = sngPreviousADR
            utxtPymAdr.SelStart = 0
            utxtPymAdr.SelLength = utxtPymAdr.Value
            utxtPymAdr.SetFocus
        Else
            If CSng(Trim(utxtPymAdr.Value)) > CSng(Trim(sngPymAmtPay)) Then
                utxtPymAdr.Value = Trim(utxtPymAdr.Value)
                utxtCustName.Text = Trim(utxtCustName.Text)
                strAdrMsg1 = utxtPymAdr.Value & " is greater than the Total Amount to be paid "
                strAdrMsg2 = utxtCustName.Text
                intResponse = MsgBox(strAdrMsg1 & strAdrMsg2, vbOKOnly, "ADR Amount Error Message")
                utxtPymAdr.Value = sngPreviousADR
                utxtPymAdr.SelStart = 0
                utxtPymAdr.SelLength = utxtPymAdr.Value
                utxtPymAdr.SetFocus
            Else
                cmdSave4.Enabled = True
            End If
        End If
        numberin = False
     End If
End If
    utxtChange.Value = EvaluateChange
    utxtChange.Value = Format(utxtChange.Value, "###,###,##0.00")
End Sub
Private Function EvaluateChange() As Single
Dim sngTotalAmt As Single
sngTotalAmt = 0

If IsNumeric(utxtPymCash.Value) Then
    sngTotalAmt = sngTotalAmt + CSng(utxtPymCash.Value)
End If

If IsNumeric(utxtPymChq(0).Value) Then
    sngTotalAmt = sngTotalAmt + CSng(utxtPymChq(0).Value)
End If

If IsNumeric(utxtPymChq(1).Value) Then
    sngTotalAmt = sngTotalAmt + CSng(utxtPymChq(1).Value)
End If

If IsNumeric(utxtPymChq(2).Value) Then
    sngTotalAmt = sngTotalAmt + CSng(utxtPymChq(2).Value)
End If

If IsNumeric(utxtPymChq(3).Value) Then
    sngTotalAmt = sngTotalAmt + CSng(utxtPymChq(3).Value)
End If

If IsNumeric(utxtPymChq(4).Value) Then
    sngTotalAmt = sngTotalAmt + CSng(utxtPymChq(4).Value)
End If

If IsNumeric(utxtPymAdr.Value) Then
    sngTotalAmt = sngTotalAmt + CSng(utxtPymAdr.Value)
End If

EvaluateChange = sngTotalAmt - sngPymAmtPay

End Function

Private Sub utxtPymAdr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 39 Then
        utxtCustNo.SetFocus
Else
    If Asc(KeyCode) > 47 And Asc(KeyCode) < 58 And IsNumeric(utxtPymAdr.Value) Then
        sngPreviousADR = utxtPymAdr.Value
        numberin = True
    Else
        sngPreviousADR = 0
        numberin = False
        utxtPymAdr.Value = ".00"
    End If
End If
End Sub

Private Sub utxtPymChq_Change(Index As Integer)
    utxtChange.Value = EvaluateChange
    utxtChange.Value = Format(utxtChange.Value, "###,###,##0.00")
    If utxtChange.Value < 0 Then
        cmdSave4.Enabled = False
    Else
        cmdSave4.Enabled = True
    End If
End Sub

Private Sub utxtPymChq_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case 37 'Left
        Case 39 'Right
            utxtPymChqNum(Index).SetFocus
    End Select
End Sub

Private Sub utxtPymChqBnk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37 'Left
        utxtPymChqNum(Index).SetFocus
    Case 39 'Right
End Select
End Sub

Private Sub utxtPymChqNum_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37 'Left
        utxtPymChq(Index).SetFocus
    Case 39 'Right
        utxtPymChqBnk(Index).SetFocus
End Select
End Sub

Private Sub utxtPymReference_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Message1 As String
Dim Message2 As String

Message1 = ""
Message2 = ""

If KeyCode = vbKeyReturn Then
    If Len(utxtPymReference.Value) > 0 Then
        If Not ChkPayment Then
            Message1 = "There is NO TRANSACTION that exist."
            Message2 = "Enter Another transaction."
            Call ErrorMessage(Message1, Message2)
            utxtPymReference.SetFocus
        Else
            cmdSave4.Enabled = True
            Call Tab03(True)
            utxtPymReference.TabStop = False
            utxtPymCash.SetFocus
        End If
    End If
End If
End Sub

Private Function ChkPayment() As Boolean
Dim rstPayment As Recordset
Dim lngPayReference As Long

ChkPayment = False
lngPayReference = 0

lngPayReference = utxtPymReference.Value
DE.ChkPayment lngPayReference
Set rstPayment = DE.rsChkPayment
With rstPayment
    If .RecordCount <> 0 Then
        ChkPayment = True
        Call MovePaymentToVariables(rstPayment)
        Call MovePayVarToScreen
    End If
.Close
Set rstPayment = Nothing
End With
End Function

Private Sub MovePaymentToVariables(rstPay As Recordset)
Call ClearPaymentVariables
sngPymAmtPay = 0
sngPymCash = 0
sngPymAdr = 0
sngPymCheque = 0
sngPymChange = 0
sngPymTotalAmt = 0
lngPymADRNum = 0
With rstPay
    sngPymCash = ReturnSingle(CStr(.Fields("cshamt")))
    sngPymAdr = .Fields("adramt")
    sngPymCheque = .Fields("chkamt1") + .Fields("chkamt2") + .Fields("chkamt3") _
                        + .Fields("chkamt4") + .Fields("chkamt5")
    sngPymChange = .Fields("chgamt")
    If IsNumeric(.Fields("cuscde")) Then
        strPymCustCode = .Fields("cuscde")
        lngPymADRNum = .Fields("adrnum")
    Else
        strPymCustCode = "0"
    End If
    strPymCustName = .Fields("cusnam") & ""
    
    sngPymChk1 = Trim(.Fields("chkamt1"))
    sngPymChk2 = Trim(.Fields("chkamt2"))
    sngPymChk3 = Trim(.Fields("chkamt3"))
    sngPymChk4 = Trim(.Fields("chkamt4"))
    sngPymChk5 = Trim(.Fields("chkamt5"))
    strPymChkN1 = Trim(.Fields("chkno1"))
    strPymChkN2 = Trim(.Fields("chkno2"))
    strPymChkN3 = Trim(.Fields("chkno3"))
    strPymChkN4 = Trim(.Fields("chkno4"))
    strPymChkN5 = Trim(.Fields("chkno5"))
    strPymChkB1 = Trim(.Fields("chkbnk1"))
    strPymChkB2 = Trim(.Fields("chkbnk2"))
    strPymChkB3 = Trim(.Fields("chkbnk3"))
    strPymChkB4 = Trim(.Fields("chkbnk4"))
    strPymChkB5 = Trim(.Fields("chkbnk5"))
    
    sngPymAmtPay = sngPymCash + sngPymCheque + sngPymAdr - sngPymChange
End With
utxtTotal.Value = Format(sngPymAmtPay, "###,##0.00")
End Sub

Private Sub MovePayVarToScreen()
    utxtPymCash.Value = Format(sngPymCash, "###,###.00")
    utxtPymAdr.Value = Format(sngPymAdr, "###,###.00")
    utxtCustNo.Value = Trim(strPymCustCode)
    utxtCustName.Text = Trim(strPymCustName)
    utxtPymChq(0).Value = Trim(Format(sngPymChk1, "###,###.00"))
    utxtPymChq(1).Value = Trim(Format(sngPymChk2, "###,###.00"))
    utxtPymChq(2).Value = Trim(Format(sngPymChk3, "###,###.00"))
    utxtPymChq(3).Value = Trim(Format(sngPymChk4, "###,###.00"))
    utxtPymChq(4).Value = Trim(Format(sngPymChk5, "###,###.00"))
    utxtPymChqNum(0).Text = Trim(strPymChkN1)
    utxtPymChqNum(1).Text = Trim(strPymChkN2)
    utxtPymChqNum(2).Text = Trim(strPymChkN3)
    utxtPymChqNum(3).Text = Trim(strPymChkN4)
    utxtPymChqNum(4).Text = Trim(strPymChkN5)
    utxtPymChqBnk(0).Text = Trim(strPymChkB1)
    utxtPymChqBnk(1).Text = Trim(strPymChkB2)
    utxtPymChqBnk(2).Text = Trim(strPymChkB3)
    utxtPymChqBnk(3).Text = Trim(strPymChkB4)
    utxtPymChqBnk(4).Text = Trim(strPymChkB5)
End Sub
Private Sub ClearPaymentSceenVar()
Dim intCtr As Integer
    utxtPymCash.Value = ".00"
For intCtr = 0 To 4
    utxtPymChq(intCtr).Value = ".00"
    utxtPymChqNum(intCtr).Text = ""
    utxtPymChqBnk(intCtr).Text = ""
Next
    utxtTotal.Value = ".00"
    utxtPymAdr.Value = ".00"
    utxtCustNo.Value = "0"
    utxtCustName.Text = ""
    utxtChange.Value = ".00"
    strPymCustName = ""
    AdrBalance = 0
End Sub
Private Sub UpdatePayment(lngRefno As Long)
Dim lngAdrnum As Long
Dim intCtr As Integer
Dim lngAdrChk As Long

lngAdrChk = 0
intCtr = 0
lngAdrnum = 0
utxtCustNo.Value = ReturnLong(CStr((utxtCustNo.Value)))
If utxtCustNo.Value = 0 Then
    utxtCustNo.Value = 0
    utxtCustName.Text = ""
End If
If CLng(utxtCustNo.Value) <> CLng(strPymCustCode) Or CSng(utxtPymAdr.Value) <> sngPymAdr Then
    If CLng(strPymCustCode) = 0 And CLng(utxtCustNo.Value) <> 0 Then
        lngAdrnum = DE.ApplyAdr(utxtCustNo.Value, "CCR", lngRefno, CSng(utxtPymAdr.Value), "", gUserid)
        If lngAdrnum <= 0 Then
            MsgBox "Apply ADR is not successful", vbOKOnly, "ADR Message"
            Exit Sub
        End If
    ElseIf CLng(strPymCustCode) <> 0 And CLng(utxtCustNo.Value) = 0 Then
        lngAdrChk = DE.CancelADr(strPymCustCode, lngPymADRNum, "", gUserid)
        If lngAdrChk < 1 Then
            MsgBox "ADR Cancelation is not successful", vbOKOnly, "ADR Message"
            Exit Sub
        End If
    ElseIf CLng(utxtCustNo.Value) = CLng(strPymCustCode) Then
        If CSng(utxtPymAdr.Value) <> sngPymAdr Then
            lngAdrChk = DE.CancelADr(strPymCustCode, lngPymADRNum, "", gUserid)
            If lngAdrChk < 1 Then
                MsgBox "ADR Cancelation is not successful", vbOKOnly, "ADR Message"
                Exit Sub
            End If
            lngAdrnum = DE.ApplyAdr(utxtCustNo.Value, "CCR", lngRefno, CSng(utxtPymAdr.Value), "", gUserid)
            If lngAdrnum <= 0 Then
                MsgBox "Apply ADR is not successful", vbOKOnly, "ADR Message"
                Exit Sub
            End If
         Else
            lngAdrnum = lngPymADRNum
        End If
    End If
    Else
        lngAdrnum = lngPymADRNum
End If
    utxtPymCash.Value = ReturnSingle(CStr(utxtPymCash.Value))
    For intCtr = 0 To 4
        utxtPymChq(intCtr).Value = ReturnSingle(CStr((utxtPymChq(intCtr).Value)))
    Next
    
DE.CCRPay CSng(utxtPymCash.Value), CSng(utxtPymChq(0).Value), utxtPymChqNum(0).Text, utxtPymChqBnk(0).Text, CSng(utxtPymChq(1).Value), utxtPymChqNum(1).Text, utxtPymChqBnk(1).Text, _
        CSng(utxtPymChq(2).Value), utxtPymChqNum(2).Text, utxtPymChqBnk(2).Text, CSng(utxtPymChq(3).Value), utxtPymChqNum(3).Text, utxtPymChqBnk(3).Text, CSng(utxtPymChq(4).Value), _
        utxtPymChqNum(4).Text, utxtPymChqBnk(4).Text, CSng(utxtChange.Value), utxtCustNo.Value, Trim(utxtCustName.Text), CSng(utxtPymAdr.Value), lngAdrnum, lngRefno
End Sub
Private Sub ClearPaymentVariables()
    sngPymCheque = 0
    sngPymChange = 0
    lngPymADRNum = 0
    sngPymTotalAmt = 0
    sngPymAmtPay = 0
    sngPymCash = 0
    sngPymAdr = 0
    sngPymChk1 = 0
    sngPymChk2 = 0
    sngPymChk3 = 0
    sngPymChk4 = 0
    sngPymChk5 = 0
    strPymChkN1 = ""
    strPymChkN2 = ""
    strPymChkN3 = ""
    strPymChkN4 = ""
    strPymChkN5 = ""
    strPymChkB1 = ""
    strPymChkB2 = ""
    strPymChkB3 = ""
    strPymChkB4 = ""
    strPymChkB5 = ""
    strPymCustCode = "0"
    strPymCustName = ""
End Sub
Private Sub utxtPymCash_Change()
    
    utxtChange.Value = EvaluateChange
    utxtChange.Value = Format(utxtChange.Value, "###,###,##0.00")
    If utxtChange.Value < 0 Then
        cmdSave4.Enabled = False
    Else
        cmdSave4.Enabled = True
    End If
End Sub
Private Sub SystemMessage(strSysMess As String, strSysMess1 As String)
    Beep
    frmMessPYXS3.lblMessPYXS3.Caption = strSysMess
    frmMessPYXS3.lblMessPYXS31.Caption = strSysMess1
    frmMessPYXS3.Timer2 = True
    frmMessPYXS3.Show (vbModal)
End Sub

Private Function PreviousPayment(strTempContainer As String) As Boolean
Dim rstPrevPay As Recordset

PreviousPayment = False
lngPrevCCrNum = 0
lngPrevRefnum = 0
lngPrevDate = 0
lngPrevTime = 0
strPrevTeller = 0
strPrevExporter = ""
strPrevTrncde = ""
strPrevWhfcde = ""

DE.SelectContainer strTempContainer
Set rstPrevPay = DE.rsSelectContainer
With rstPrevPay
    If .RecordCount <> 0 Then
        lngPrevRefnum = .Fields("refnum")
        lngPrevCCrNum = .Fields("ccrnum")
        lngPrevDate = Format(.Fields("sysdttm"), "yyyymmdd")
        lngPrevTime = Format(.Fields("sysdttm"), "HhNnss")
        strPrevTeller = .Fields("userid")
        strPrevTrncde = .Fields("trncde")
        strPrevExporter = .Fields("exprtr")
        strPrevWhfcde = .Fields("whfcde")
        PreviousPayment = True
    End If
    .Close
End With
Set rstPrevPay = Nothing
End Function

Private Function CheckContainer() As Boolean
Dim strValid1 As String
Dim strValid2 As String
Dim lng7thDigit As Long
Dim str6thDigit As String
Dim lngTmp7thDigit As Long
Dim strCntNo As String

strValid1 = ""
strValid2 = ""
lng7thDigit = 0
str6thDigit = ""
strCntNo = ""
lngTmp7thDigit = 0
CheckContainer = True

'   ** Checks Container

If Len(utxtNewCntPrefix.Text) = 4 And Len(utxtNewCntNum.Text) <> 0 Then
'   ** Check Container 7th digit
        If Len(utxtNewCntNum.Text) = 7 Then
            strCntNo = utxtNewCntNum.Text
            str6thDigit = Mid(strCntNo, 1, 6)
            lng7thDigit = CSng(Mid(strCntNo, 1, 6)) Mod 11
            lngTmp7thDigit = CSng(Mid(strCntNo, 7, 1))
            
            If lng7thDigit <> lngTmp7thDigit Then
                strValid1 = "Seventh digit shoul be  " & lng7thDigit & " "
                strValid2 = "Do You Want to Accept Seventh Digit ?"
                strResponse = False
                Call SystemMessage(strValid1, strValid2)
                If strResponse Then
                    utxtNewCntNum.Text = str6thDigit & lng7thDigit
                 End If
            End If
        End If
        If Not ContainerValid(utxtNewCntPrefix.Text & utxtNewCntNum.Text) Then
               strValid1 = "Cannot Update to Container " & utxtNewCntPrefix.Text & utxtNewCntNum.Text & " "
               strValid2 = "Record Exist with the same Transaction"
               strResponse = False
               Call ErrorMessage(strValid1, strValid2)
               CheckContainer = False
           End If
Else
    CheckContainer = False
    utxtNewCntPrefix.SetFocus
End If
End Function

Private Function ContainerValid(ByVal Container As String) As Boolean
Dim CntValid As Recordset
Dim strTempRef As String
strTempRef = ""
ContainerValid = False
DE.SelectContainer Container
Set CntValid = DE.rsSelectContainer
With CntValid
    If .RecordCount > 0 Then
        strTempRef = Format(.Fields("refnum"), "00000000")
        If strTempRef <> utxtRefnum.Value Then
            ContainerValid = True
        End If
    Else
        ContainerValid = True
    End If
.Close
Set CntValid = Nothing
End With
End Function

Private Function ChkCCR(lngChkCCRNo As Long) As Long
' ** retrieve the returned values
ChkCCR = DE.ChkCCRNum(utxtDetTeller1.Text, lngChkCCRNo)
End Function

Private Sub UpdateOVRCCR(lngNewOvrCCr As Long, lngOldOvrCCr As Long)
DE.UpdateOVRCCR lngNewOvrCCr, lngOldOvrCCr
End Sub

Private Sub ClearDetails()
    lngSveDetRefnum = 0
    lngSveDetSequence = 0
    lngSveDetCCRNum = 0
    strSveDetVessel = ""
    strSveDetCntnum = ""
    strSveDetCommod = ""
    strSveDetBroker = ""
    strSveDetExporter = ""
    strSveDetSBMAPermit = ""
    strSveDetEntnum = ""
    strSveDetEntnum1 = ""
    strSveDetEntnum2 = ""
    strSveDetEntnum3 = ""
    strSveDetEntnum4 = ""
    strSveDetEntnum5 = ""
    strSveDetEntnum6 = ""
    strSveDetEntnum7 = ""
    strSveDetEntnum8 = ""
    strSveDetEntnum9 = ""
    strSveDetEntnum10 = ""
    strSveDetExempCode = ""
    strSveDetguarantee = ""
    strSveDetTeller = ""
    lngSveDetDate = 0
    lngSveDetTime = 0
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,0,0,0
Public Property Get ObjectToUnload() As Object
    Set ObjectToUnload = m_ObjectToUnload
End Property

Public Property Set ObjectToUnload(ByVal New_ObjectToUnload As Object)
    Set m_ObjectToUnload = New_ObjectToUnload
    PropertyChanged "ObjectToUnload"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set m_ObjectToUnload = PropBag.ReadProperty("ObjectToUnload", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ObjectToUnload", m_ObjectToUnload, Nothing)
End Sub

Public Sub StartInitialization()
Set cusListing = New cCustomer
    cusListing.FillCustomer
'**********************
'* StandAlone FLR
'    Set CTCSinfo = CreateObject("CTCS.cCTCS")    'Opens detail for CTCS
'    CTCSinfo.Connect
    CorrectionTab.Visible = False
    Load frmMessPYX3
    frmMessPYX3.Timer1 = False
    Load frmMessPYXS3
    frmMessPYXS3.Timer2 = False
    Call IniVariable
End Sub

Private Sub utxtRefCntno_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Message1 As String
Dim Message2 As String

Message1 = ""
Message2 = ""

    If KeyCode = vbKeyReturn Then
        If Not ChkRefNoSeqNoCntNum(utxtRefRefNum.Value, utxtRefSeq.Value, utxtRefPrefix.Text, _
                                utxtRefCntno.Value) Then
            Message1 = "There is NO TRANSACTION that exist."
            Message2 = "Enter Another transaction."
            Call ErrorMessage(Message1, Message2)
            utxtRefRefNum.SetFocus
        Else
            If blnOutdttm Then
                Message1 = "Cannot Refund Transaction NOT IN YARD."
                Message2 = "Enter Another transaction."
                Call ErrorMessage(Message1, Message2)
            ElseIf blnCancel Then
                Message1 = "Cannot Refund a CANCELLED Transaction."
                Message2 = "Enter Another transaction."
                Call ErrorMessage(Message1, Message2)
            ElseIf blnRefund Then
                Message1 = "Cannot Refund a Refunded Transaction."
                Message2 = "Enter Another transaction."
                Call ErrorMessage(Message1, Message2)
'            ElseIf blnPPA Then
'                Message1 = "Cannot Refund Transaction, PPA OR Exist."
'                Message2 = "Enter Another Transaction."
'                Call ErrorMessage(Message1, Message2)
            ElseIf Not blnOutdttm And Not blnCancel Then  'And Not blnPPA Then
                Call MoveDetailsToRefundFields
                cmdSave5.Enabled = True
            End If
         utxtRefRefNum.SetFocus
        End If
    End If
End Sub

Private Sub utxtTotal_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Function CheckReference(lngReference As Long) As Boolean
Dim rstChkReference As Recordset

CheckReference = True
DE.ChkReference lngReference
Set rstChkReference = DE.rsChkReference
    
With rstChkReference
    If .RecordCount > 0 Then
        CheckReference = False
    End If
    .Close
Set rstChkReference = Nothing
End With
End Function

Private Sub UpdateCNLPayment(lngRefno As Long, lngSeqno As Long)
Dim rstDetails As Recordset
Dim rstPayment As Recordset
Dim rstCCrPay As Recordset

Dim sngTotalAmt As Single
Dim sngTotalPay As Single
Dim sngTotalRemain As Single
Dim sngAmtDeduct As Single

sngTotalAmt = 0
sngTotalPay = 0
sngTotalRemain = 0
sngAmtDeduct = 0

Call ClearPaymentVariables

DE.RtvDetails lngRefno, lngSeqno
Set rstDetails = DE.rsRtvDetails
With rstDetails
    Do While Not .EOF
        If .Fields("guarntycde") <> "Y" Then
            sngTotalAmt = .Fields("arramt") + .Fields("ovzamt") + .Fields("dgramt") + .Fields("arrvat") - _
                .Fields("arrtax")
        End If
        If .Fields("whfcde") = 0 Then
            sngTotalAmt = sngTotalAmt + .Fields("whfamt")
        End If
    .MoveNext
    Loop
    sngAmtDeduct = sngTotalAmt
.Close
Set rstDetails = Nothing
End With

DE.ChkPayment lngRefno
Set rstPayment = DE.rsChkPayment
With rstPayment
    sngPymCash = .Fields("cshamt")
    sngPymChk1 = .Fields("chkamt1")
    strPymChkN1 = .Fields("chkno1")
    strPymChkB1 = .Fields("chkbnk1")
    sngPymChk2 = .Fields("chkamt2")
    strPymChkN2 = .Fields("chkno2")
    strPymChkB2 = .Fields("chkbnk2")
    sngPymChk3 = .Fields("chkamt3")
    strPymChkN3 = .Fields("chkno3")
    strPymChkB3 = .Fields("chkbnk3")
    sngPymChk4 = .Fields("chkamt4")
    strPymChkN4 = .Fields("chkno4")
    strPymChkB4 = .Fields("chkbnk4")
    sngPymChk5 = .Fields("chkamt5")
    strPymChkN5 = .Fields("chkno5")
    strPymChkB5 = .Fields("chkbnk5")
    sngTotalPay = .Fields("cshamt") + .Fields("chkamt1") + .Fields("chkamt2") + .Fields("chkamt3") + _
                            .Fields("chkamt4") + .Fields("chkamt5") - .Fields("chgamt")
.Close
Set rstPayment = Nothing
End With

'   Check1 Deduction
If sngTotalAmt > 0 Then
    If sngPymChk1 > 0 Then
        If sngPymChk1 > sngTotalAmt Then
            sngPymChk1 = sngPymChk1 - sngTotalAmt
            sngTotalAmt = 0
        Else
            sngTotalAmt = sngTotalAmt - sngPymChk1
            sngPymChk1 = 0
            strPymChkN1 = ""
            strPymChkB1 = ""
        End If
    End If
End If

'   Check2 Deduction
If sngTotalAmt > 0 Then
    If sngPymChk2 > 0 Then
        If sngPymChk2 > sngTotalAmt Then
            sngPymChk2 = sngPymChk2 - sngTotalAmt
            sngTotalAmt = 0
        Else
            sngTotalAmt = sngTotalAmt - sngPymChk2
            sngPymChk2 = 0
            strPymChkN2 = ""
            strPymChkB2 = ""
        End If
    End If
End If

'   Check3 Deduction
If sngTotalAmt > 0 Then
    If sngPymChk3 > 0 Then
        If sngPymChk3 > sngTotalAmt Then
            sngPymChk3 = sngPymChk3 - sngTotalAmt
            sngTotalAmt = 0
        Else
            sngTotalAmt = sngTotalAmt - sngPymChk3
            sngPymChk3 = 0
            strPymChkN3 = ""
            strPymChkB3 = ""
        End If
    End If
End If

'   Check4 Deduction
If sngTotalAmt > 0 Then
    If sngPymChk4 > 0 Then
        If sngPymChk4 > sngTotalAmt Then
            sngPymChk4 = sngPymChk4 - sngTotalAmt
            sngTotalAmt = 0
        Else
            sngTotalAmt = sngTotalAmt - sngPymChk4
            sngPymChk4 = 0
            strPymChkN4 = ""
            strPymChkB4 = ""
        End If
    End If
End If

'   Check5 Deduction
If sngTotalAmt > 0 Then
    If sngPymChk5 > 0 Then
        If sngPymChk5 > sngTotalAmt Then
            sngPymChk5 = sngPymChk5 - sngTotalAmt
            sngTotalAmt = 0
        Else
            sngTotalAmt = sngTotalAmt - sngPymChk5
            sngPymChk5 = 0
            strPymChkN5 = ""
            strPymChkB5 = ""
        End If
    End If
End If

'   Cash Deduction
If sngTotalAmt > 0 Then
    If sngPymCash > 0 Then
        If sngPymCash > sngTotalAmt Then
            sngPymCash = sngPymCash - sngTotalAmt
            sngTotalAmt = 0
        Else
            sngTotalAmt = sngTotalAmt - sngPymCash
            sngPymCash = 0
        End If
    End If
End If

sngTotalRemain = sngPymCash + sngPymChk1 + sngPymChk2 + sngPymChk3 + sngPymChk4 + sngPymChk5
sngPymChange = sngTotalRemain - (sngTotalPay - sngAmtDeduct)

'Update CCRPay

DE.CCRPay sngPymCash, sngPymChk1, strPymChkN1, strPymChkB1, sngPymChk2, strPymChkN2, strPymChkB2, sngPymChk3, _
        strPymChkN3, strPymChkB3, sngPymChk4, strPymChkN4, strPymChkB4, sngPymChk5, strPymChkN5, strPymChkB5, sngPymChange, _
        strPymCustCode, strPymCustName, sngPymAdr, lngPymADRNum, lngRefno
End Sub
Private Function ReturnSingle(AmountStr As String) As Single
If Not IsNull(AmountStr) Then
    If IsNumeric(AmountStr) Then
        ReturnSingle = CSng(AmountStr)
    Else
        ReturnSingle = 0
    End If
Else
    ReturnSingle = 0
End If
End Function
Private Function ReturnLong(AmountStr As String) As Long
If Not IsNull(AmountStr) Then
    If IsNumeric(AmountStr) Then
        ReturnLong = CLng(AmountStr)
    Else
        ReturnLong = 0
    End If
Else
    ReturnLong = 0
End If
End Function

Private Sub RefundUpdate()
Dim rstRefund As Recordset
DE.RefundUpdate lngRefnum, lngSeqnum, strCntnum
'**********************
'* StandAlone FLR
'CTCSinfo.DelCYExport strCntnum
DE.CyxRefund
Set rstRefund = DE.rsCyxRefund
With rstRefund
    .AddNew
    .Fields("refnum") = ReturnLong(CStr(lngRefnum))
    .Fields("seqnum") = ReturnLong(CStr(lngSeqnum))
    .Fields("itmnum") = ReturnLong(CStr(lngItmnum))
    .Fields("cntnum") = Trim(strCntnum)
    .Fields("ccrnum") = ReturnLong(CStr(lngCcrnum))
    .Fields("cntsze") = ReturnLong(CStr(lngCntSze))
    .Fields("fulemp") = Trim(strFulemp)
    .Fields("broker") = Trim(strBroker)
    .Fields("exprtr") = Trim(strExprtr)
    .Fields("paydte") = Format(strSysdttm, "yyyy-mm-dd hh:nn:ss")
    .Fields("userid") = Trim(gUserid)
    .Fields("sysdte") = zGetSysDate
    .Update
    .Close
End With

Set rstRefund = Nothing
End Sub

Private Sub IniVariable()
    utxtRefnum.Value = ""
    utxtSeqnum.Value = ""
    utxtCntPrefix.Text = ""
    utxtCntNum.Text = ""
    utxtDetRefnum.Value = ""
    utxtDetSeqnum.Value = ""
    utxtCnlRefnum.Value = ""
    utxtCnlSeqnum.Value = ""
    utxtPymReference.Value = ""
    utxtRefRefNum.Value = ""
    utxtRefSeq.Value = ""
    utxtRefPrefix.Text = ""
    utxtRefCntno.Value = ""
End Sub

Private Sub WriteToLog(tmpRefnum As Long, tmpSeqnum As Long, _
                    tmpCntnum As String, tmpDte As String, tmpUser As String)
Dim rstCCRCyxz As Recordset
Dim rstCCRCyx As Recordset

Set rstCCRCyxz = New ADODB.Recordset
rstCCRCyxz.CursorType = adOpenDynamic
rstCCRCyxz.LockType = adLockOptimistic
rstCCRCyxz.Open "CCRcyxz", gcnnBilling, , , adCmdTable

If Len(tmpCntnum) > 0 Then
    DE.RtvContainer tmpRefnum, tmpSeqnum, tmpCntnum
    Set rstCCRCyx = DE.rsRtvContainer
Else
    DE.RtvDetails tmpRefnum, tmpSeqnum
    Set rstCCRCyx = DE.rsRtvDetails
End If
With rstCCRCyx
    Do Until .EOF
        Call ClearVariableDetails
        Call MoveCCRDetailsToVariables(rstCCRCyx)
        Call MoveSaveData(rstCCRCyxz, tmpDte, tmpUser)
        .MoveNext
     Loop
End With
rstCCRCyx.Close
Set rstCCRCyx = Nothing
rstCCRCyxz.Close
Set rstCCRCyxz = Nothing
End Sub
