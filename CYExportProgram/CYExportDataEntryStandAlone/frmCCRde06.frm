VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCCRde06 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ( SBMA SUBIC - CCRDE06 )  CY Export Data Entry"
   ClientHeight    =   11850
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   15525
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCCRde06.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11850
   ScaleWidth      =   15525
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox cMode 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10920
      TabIndex        =   165
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin TabDlg.SSTab sstMain 
      CausesValidation=   0   'False
      Height          =   11415
      Left            =   0
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   20135
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Details"
      TabPicture(0)   =   "frmCCRde06.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtTotDue"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtPpaTotal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtIctsiDue"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblNoCnt"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label50"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label19"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label20"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label21"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "frmCCRde06"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "flexDetails"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdEdit"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdHeader"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdExit"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdAdd"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkNewCCR"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdCancel"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdGrid"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdDelete"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdPayment"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "utxtPref"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "utxtSze"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "utxtFEmp"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "utxtTshipMnt"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "utxtDollar"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "utxtLength"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "utxtWidth"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "utxtHeight"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "utxtUMS"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "utxtNumDangr"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "flexDangerClass"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "flexTshipMnt"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "utxtNo"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Header"
      TabPicture(1)   =   "frmCCRde06.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "utxtSBMAPermit"
      Tab(1).Control(1)=   "utxtWhfExmp"
      Tab(1).Control(2)=   "utxtUGuarantee"
      Tab(1).Control(3)=   "utxtWhfOnly"
      Tab(1).Control(4)=   "utxtWhfExmpt"
      Tab(1).Control(5)=   "utxtVatCode"
      Tab(1).Control(6)=   "utxtVessel"
      Tab(1).Control(7)=   "utxtCommodity"
      Tab(1).Control(8)=   "utxtRemark"
      Tab(1).Control(9)=   "utxtBroker"
      Tab(1).Control(10)=   "utxtExporter"
      Tab(1).Control(11)=   "utxtEntry1(10)"
      Tab(1).Control(12)=   "utxtEntry1(9)"
      Tab(1).Control(13)=   "utxtEntry1(7)"
      Tab(1).Control(14)=   "utxtEntry1(6)"
      Tab(1).Control(15)=   "utxtEntry1(5)"
      Tab(1).Control(16)=   "utxtEntry1(4)"
      Tab(1).Control(17)=   "utxtEntry1(3)"
      Tab(1).Control(18)=   "utxtEntry1(2)"
      Tab(1).Control(19)=   "utxtEntry1(1)"
      Tab(1).Control(20)=   "utxtEntry1(0)"
      Tab(1).Control(21)=   "Frame10"
      Tab(1).Control(22)=   "cmdBack"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Frame5"
      Tab(1).Control(24)=   "Frame3"
      Tab(1).Control(25)=   "Frame7"
      Tab(1).Control(26)=   "Frame16"
      Tab(1).Control(27)=   "Frame2"
      Tab(1).Control(28)=   "Label17"
      Tab(1).ControlCount=   29
      TabCaption(2)   =   "Payment"
      TabPicture(2)   =   "frmCCRde06.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "utxtChqNo(4)"
      Tab(2).Control(1)=   "utxtChqNo(3)"
      Tab(2).Control(2)=   "utxtChqNo(2)"
      Tab(2).Control(3)=   "utxtChqNo(1)"
      Tab(2).Control(4)=   "utxtChqNo(0)"
      Tab(2).Control(5)=   "utxtAdrAmt"
      Tab(2).Control(6)=   "utxtCustName"
      Tab(2).Control(7)=   "utxtCustNo"
      Tab(2).Control(8)=   "utxtChqBnk(4)"
      Tab(2).Control(9)=   "utxtChqBnk(3)"
      Tab(2).Control(10)=   "utxtChqBnk(2)"
      Tab(2).Control(11)=   "utxtChqBnk(1)"
      Tab(2).Control(12)=   "utxtChqBnk(0)"
      Tab(2).Control(13)=   "utxtChq(4)"
      Tab(2).Control(14)=   "utxtChq(3)"
      Tab(2).Control(15)=   "utxtChq(2)"
      Tab(2).Control(16)=   "utxtChq(1)"
      Tab(2).Control(17)=   "utxtChq(0)"
      Tab(2).Control(18)=   "utxtCsh"
      Tab(2).Control(19)=   "utxtCCRNo"
      Tab(2).Control(20)=   "cmdPymCancel"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "cmdPymBack"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "cmdPrint"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Frame6"
      Tab(2).Control(24)=   "Frame8"
      Tab(2).ControlCount=   25
      Begin CCRDE06.utxtTextBilling utxtSBMAPermit 
         Height          =   420
         Left            =   -66840
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   30
      End
      Begin CCRDE06.pText utxtNo 
         Height          =   420
         Left            =   5040
         TabIndex        =   1
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexTshipMnt 
         Height          =   495
         Left            =   4440
         TabIndex        =   83
         Top             =   2400
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   873
         _Version        =   393216
         Rows            =   4
         Cols            =   1
         FixedCols       =   0
         HighLight       =   0
         FormatString    =   "                 Transshipment"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flexDangerClass 
         Height          =   495
         Left            =   4440
         TabIndex        =   84
         Top             =   1920
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   873
         _Version        =   393216
         Rows            =   10
         Cols            =   1
         FixedCols       =   0
         HighLight       =   0
         FormatString    =   "                  Danger Class"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CCRDE06.utxtTextBilling utxtChqNo 
         Height          =   420
         Index           =   4
         Left            =   -69600
         TabIndex        =   55
         Top             =   5640
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CCRDE06.utxtTextBilling utxtChqNo 
         Height          =   420
         Index           =   3
         Left            =   -69600
         TabIndex        =   52
         Top             =   5160
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CCRDE06.utxtTextBilling utxtChqNo 
         Height          =   420
         Index           =   2
         Left            =   -69600
         TabIndex        =   49
         Top             =   4680
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CCRDE06.utxtTextBilling utxtChqNo 
         Height          =   420
         Index           =   1
         Left            =   -69600
         TabIndex        =   46
         Top             =   4200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CCRDE06.utxtTextBilling utxtChqNo 
         Height          =   420
         Index           =   0
         Left            =   -69600
         TabIndex        =   43
         Top             =   3720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CCRDE06.utxtNumBilling utxtWhfExmp 
         Height          =   420
         Left            =   -69720
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   5160
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   741
         BackColor       =   -2147483633
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
         MaxLength       =   1
      End
      Begin CCRDE06.utxtNumBilling utxtAdrAmt 
         Height          =   420
         Left            =   -72240
         TabIndex        =   58
         Top             =   7200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   14
         Alignment       =   1
         Maskformat      =   "###,###,###.00"
         Maskformat      =   "###,###,###.00"
         DecimalPlaces   =   2
         Last            =   -1  'True
      End
      Begin CCRDE06.utxtTextBilling utxtCustName 
         Height          =   420
         Left            =   -69600
         TabIndex        =   59
         Top             =   6600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   741
         BackColor       =   -2147483633
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
         Alignment       =   2
      End
      Begin CCRDE06.utxtNumBilling utxtCustNo 
         Height          =   420
         Left            =   -72240
         TabIndex        =   57
         Top             =   6600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   8
         Alignment       =   1
         Maskformat      =   "########"
         Maskformat      =   "########"
      End
      Begin CCRDE06.utxtTextBilling utxtChqBnk 
         Height          =   420
         Index           =   4
         Left            =   -66720
         TabIndex        =   56
         Top             =   5640
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
      End
      Begin CCRDE06.utxtTextBilling utxtChqBnk 
         Height          =   420
         Index           =   3
         Left            =   -66720
         TabIndex        =   53
         Top             =   5160
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
      End
      Begin CCRDE06.utxtTextBilling utxtChqBnk 
         Height          =   420
         Index           =   2
         Left            =   -66720
         TabIndex        =   50
         Top             =   4680
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
      End
      Begin CCRDE06.utxtTextBilling utxtChqBnk 
         Height          =   420
         Index           =   1
         Left            =   -66720
         TabIndex        =   47
         Top             =   4200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
      End
      Begin CCRDE06.utxtTextBilling utxtChqBnk 
         Height          =   420
         Index           =   0
         Left            =   -66720
         TabIndex        =   44
         Top             =   3720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
      End
      Begin CCRDE06.utxtNumBilling utxtChq 
         Height          =   420
         Index           =   4
         Left            =   -72240
         TabIndex        =   54
         Top             =   5640
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   14
         Alignment       =   1
         Maskformat      =   "###,###,###.00"
         Maskformat      =   "###,###,###.00"
         DecimalPlaces   =   2
      End
      Begin CCRDE06.utxtNumBilling utxtChq 
         Height          =   420
         Index           =   3
         Left            =   -72240
         TabIndex        =   51
         Top             =   5160
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   14
         Alignment       =   1
         Maskformat      =   "###,###,###.00"
         Maskformat      =   "###,###,###.00"
         DecimalPlaces   =   2
      End
      Begin CCRDE06.utxtNumBilling utxtChq 
         Height          =   420
         Index           =   2
         Left            =   -72240
         TabIndex        =   48
         Top             =   4680
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   14
         Alignment       =   1
         Maskformat      =   "###,###,###.00"
         Maskformat      =   "###,###,###.00"
         DecimalPlaces   =   2
      End
      Begin CCRDE06.utxtNumBilling utxtChq 
         Height          =   420
         Index           =   1
         Left            =   -72240
         TabIndex        =   45
         Top             =   4200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   14
         Alignment       =   1
         Maskformat      =   "###,###,###.00"
         Maskformat      =   "###,###,###.00"
         DecimalPlaces   =   2
      End
      Begin CCRDE06.utxtNumBilling utxtChq 
         Height          =   420
         Index           =   0
         Left            =   -72240
         TabIndex        =   42
         Top             =   3720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   14
         Alignment       =   1
         Maskformat      =   "###,###,###.00"
         Maskformat      =   "###,###,###.00"
         DecimalPlaces   =   2
      End
      Begin CCRDE06.utxtNumBilling utxtCsh 
         Height          =   420
         Left            =   -72240
         TabIndex        =   41
         Top             =   3240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   14
         Alignment       =   1
         Maskformat      =   "###,###,###.00"
         Maskformat      =   "###,###,###.00"
         DecimalPlaces   =   2
      End
      Begin CCRDE06.utxtNumBilling utxtCCRNo 
         Height          =   975
         Left            =   -72240
         TabIndex        =   39
         Top             =   1200
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   1720
         BackColor       =   -2147483633
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CCRDE06.utxtTextBilling utxtUGuarantee 
         Height          =   420
         Left            =   -74400
         TabIndex        =   26
         Top             =   6120
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   1
      End
      Begin CCRDE06.utxtTextBilling utxtWhfOnly 
         Height          =   420
         Left            =   -74400
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   5160
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   741
         BackColor       =   -2147483633
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
         MaxLength       =   1
      End
      Begin CCRDE06.utxtTextBilling utxtWhfExmpt 
         Height          =   420
         Left            =   -74400
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   5280
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   741
         BackColor       =   -2147483633
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
         MaxLength       =   1
      End
      Begin CCRDE06.utxtNumBilling utxtVatCode 
         Height          =   420
         Left            =   -74400
         TabIndex        =   22
         Top             =   4080
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CCRDE06.utxtTextBilling utxtVessel 
         Height          =   420
         Left            =   -72000
         TabIndex        =   21
         Top             =   2880
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   12
      End
      Begin CCRDE06.utxtTextBilling utxtCommodity 
         Height          =   420
         Left            =   -72000
         TabIndex        =   18
         Top             =   1920
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   30
      End
      Begin CCRDE06.utxtTextBilling utxtRemark 
         Height          =   420
         Left            =   -72000
         TabIndex        =   20
         Top             =   2400
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   30
      End
      Begin CCRDE06.utxtTextBilling utxtBroker 
         Height          =   420
         Left            =   -72000
         TabIndex        =   17
         Top             =   1440
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   30
      End
      Begin CCRDE06.utxtTextBilling utxtExporter 
         Height          =   420
         Left            =   -72000
         TabIndex        =   16
         Top             =   960
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   30
      End
      Begin CCRDE06.utxtNumBilling utxtNumDangr 
         Height          =   420
         Left            =   3720
         TabIndex        =   4
         Top             =   1920
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   1
      End
      Begin CCRDE06.utxtTextBilling utxtUMS 
         Height          =   420
         Left            =   11760
         TabIndex        =   10
         Top             =   3000
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   1
      End
      Begin CCRDE06.utxtNumBilling utxtHeight 
         Height          =   420
         Left            =   9360
         TabIndex        =   9
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   6
         Maskformat      =   "###.00"
         Maskformat      =   "###.00"
      End
      Begin CCRDE06.utxtNumBilling utxtWidth 
         Height          =   420
         Left            =   6480
         TabIndex        =   8
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   6
         Maskformat      =   "###.00"
         Maskformat      =   "###.00"
      End
      Begin CCRDE06.utxtNumBilling utxtLength 
         Height          =   420
         Left            =   3720
         TabIndex        =   7
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   6
         Maskformat      =   "###.00"
         Maskformat      =   "###.00"
      End
      Begin CCRDE06.utxtNumBilling utxtDollar 
         Height          =   420
         Left            =   6600
         TabIndex        =   6
         Top             =   2400
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   6
         Maskformat      =   "###.00"
         Maskformat      =   "###.00"
         DecimalPlaces   =   2
      End
      Begin CCRDE06.utxtTextBilling utxtTshipMnt 
         Height          =   420
         Left            =   3720
         TabIndex        =   5
         Top             =   2400
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   1
      End
      Begin CCRDE06.utxtTextBilling utxtFEmp 
         Height          =   420
         Left            =   8640
         TabIndex        =   3
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   1
      End
      Begin CCRDE06.utxtNumBilling utxtSze 
         Height          =   420
         Left            =   3720
         TabIndex        =   2
         Top             =   1320
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   20
      End
      Begin CCRDE06.utxtTextBilling utxtPref 
         Height          =   420
         Left            =   3720
         TabIndex        =   0
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   4
      End
      Begin CCRDE06.utxtEntry utxtEntry1 
         Height          =   420
         Index           =   10
         Left            =   -68160
         TabIndex        =   38
         Top             =   8280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin CCRDE06.utxtEntry utxtEntry1 
         Height          =   420
         Index           =   9
         Left            =   -69720
         TabIndex        =   37
         Top             =   8280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin CCRDE06.utxtEntry utxtEntry1 
         Height          =   420
         Index           =   7
         Left            =   -71280
         TabIndex        =   36
         Top             =   8280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin CCRDE06.utxtEntry utxtEntry1 
         Height          =   420
         Index           =   6
         Left            =   -72840
         TabIndex        =   35
         Top             =   8280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin CCRDE06.utxtEntry utxtEntry1 
         Height          =   420
         Index           =   5
         Left            =   -74400
         TabIndex        =   34
         Top             =   8280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin CCRDE06.utxtEntry utxtEntry1 
         Height          =   420
         Index           =   4
         Left            =   -68160
         TabIndex        =   33
         Top             =   7800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin CCRDE06.utxtEntry utxtEntry1 
         Height          =   420
         Index           =   3
         Left            =   -69720
         TabIndex        =   32
         Top             =   7800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin CCRDE06.utxtEntry utxtEntry1 
         Height          =   420
         Index           =   2
         Left            =   -71280
         TabIndex        =   31
         Top             =   7800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin CCRDE06.utxtEntry utxtEntry1 
         Height          =   420
         Index           =   1
         Left            =   -72840
         TabIndex        =   30
         Top             =   7800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin CCRDE06.utxtEntry utxtEntry1 
         Height          =   420
         Index           =   0
         Left            =   -74400
         TabIndex        =   29
         Top             =   7800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin CCRDE06.utxtEntry utxtEntry1 
         Height          =   420
         Index           =   8
         Left            =   -69840
         TabIndex        =   61
         Top             =   5880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame10 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74760
         TabIndex        =   103
         Top             =   5760
         Width           =   14775
         Begin VB.Frame FrameCust 
            BorderStyle     =   0  'None
            Caption         =   "Frame17"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   5160
            TabIndex        =   167
            Top             =   240
            Visible         =   0   'False
            Width           =   9375
            Begin CCRDE06.utxtTextBilling utxtCustName1 
               Height          =   420
               Left            =   2640
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   480
               Width           =   6495
               _ExtentX        =   11456
               _ExtentY        =   741
               BackColor       =   -2147483633
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
               Alignment       =   2
            End
            Begin CCRDE06.utxtNumBilling utxtCustNo1 
               Height          =   420
               Left            =   0
               TabIndex        =   27
               Top             =   480
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   741
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   15
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxLength       =   8
               Alignment       =   1
               Maskformat      =   "########"
               Maskformat      =   "########"
            End
            Begin VB.Label Label65 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Customer Code"
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
               ForeColor       =   &H00800080&
               Height          =   450
               Left            =   0
               TabIndex        =   169
               Top             =   0
               Width           =   2415
            End
            Begin VB.Label Label64 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Customer Name"
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
               ForeColor       =   &H00800080&
               Height          =   450
               Left            =   2640
               TabIndex        =   168
               Top             =   0
               Width           =   6495
            End
         End
         Begin VB.Label Label55 
            Alignment       =   2  'Center
            Caption         =   "{Y/N} Under Guarantee"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   104
            Top             =   360
            Width           =   3615
         End
      End
      Begin VB.CommandButton cmdPymCancel 
         Caption         =   "F12 - Cancel Transaction"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   -64080
         Style           =   1  'Graphical
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   9360
         Width           =   3855
      End
      Begin VB.CommandButton cmdPymBack 
         Caption         =   "F6 - Back"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   -74640
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   8640
         Width           =   2775
      End
      Begin VB.CommandButton cmdPayment 
         Caption         =   "F12 - Payment"
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
         Height          =   690
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   9360
         Width           =   3400
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "F9 - Cont. No"
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
         Height          =   690
         Left            =   4080
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   9360
         Width           =   3400
      End
      Begin VB.CommandButton cmdGrid 
         Caption         =   "F2 - Grid"
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
         Height          =   690
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   9360
         Width           =   3400
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "F6 - Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   9360
         Width           =   3400
      End
      Begin VB.CheckBox chkNewCCR 
         Caption         =   "New CCR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9600
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "F4 - Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   8640
         Width           =   3400
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "F7 - Save and Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   -64080
         Style           =   1  'Graphical
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   8640
         Width           =   3855
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "F7 - Continue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -65880
         Style           =   1  'Graphical
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   9120
         Width           =   2775
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "F3 - Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   8640
         Width           =   3400
      End
      Begin VB.CommandButton cmdHeader 
         Caption         =   "F7 -  Header"
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
         Height          =   690
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   8640
         Width           =   3400
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "F5 - Edit"
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
         Height          =   690
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   8640
         Width           =   3400
      End
      Begin MSFlexGridLib.MSFlexGrid flexDetails 
         Height          =   4095
         Left            =   240
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   3600
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7223
         _Version        =   393216
         Cols            =   46
         FixedCols       =   0
         BackColor       =   12910591
         BackColorFixed  =   8388608
         ForeColorFixed  =   65535
         BackColorSel    =   8388736
         BackColorBkg    =   16777215
         Enabled         =   0   'False
         FocusRect       =   2
         HighLight       =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame frmCCRde06 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   240
         TabIndex        =   107
         Top             =   120
         Width           =   14775
         Begin VB.CheckBox chkWeighing 
            Caption         =   "Weighing"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   12840
            TabIndex        =   170
            Top             =   2880
            Value           =   1  'Checked
            Width           =   1815
         End
         Begin VB.Frame Frame14 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   120
            TabIndex        =   110
            Top             =   960
            Width           =   14535
         End
         Begin VB.Frame Frame13 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   120
            TabIndex        =   109
            Top             =   1560
            Width           =   14535
         End
         Begin VB.Frame Frame12 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   120
            TabIndex        =   108
            Top             =   2640
            Width           =   14535
         End
         Begin VB.Label lblCompCode 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   11520
            TabIndex        =   172
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label66 
            Caption         =   "Company Code:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   9360
            TabIndex        =   171
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label11 
            Caption         =   "Full or Empty Code (F/E)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   4320
            TabIndex        =   123
            Top             =   1200
            Width           =   3975
         End
         Begin VB.Label Label10 
            Caption         =   "UMS       C/I"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   10800
            TabIndex        =   122
            Top             =   2880
            Width           =   2655
         End
         Begin VB.Label Label9 
            Caption         =   "Height :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7920
            TabIndex        =   121
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label8 
            Caption         =   "Width :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5280
            TabIndex        =   120
            Top             =   2880
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Dollar Rate :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4200
            TabIndex        =   119
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label Label6 
            Caption         =   "Length :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2400
            TabIndex        =   118
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Over Size"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   240
            TabIndex        =   117
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Transshipment Code "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   116
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "Danger Class "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   960
            TabIndex        =   115
            Top             =   1800
            Width           =   2415
         End
         Begin VB.Label Label2 
            Caption         =   "Size "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2280
            TabIndex        =   114
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Container Number "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   113
            Top             =   600
            Width           =   3135
         End
         Begin VB.Label Label32 
            BackColor       =   &H00800080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CCR Container Details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   112
            Top             =   120
            Width           =   14775
         End
         Begin VB.Label lblStatus 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   420
            Left            =   6720
            TabIndex        =   111
            Top             =   600
            Width           =   7815
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   106
         Top             =   8400
         Width           =   14895
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   -74880
         TabIndex        =   124
         Top             =   8400
         Width           =   14895
         Begin VB.Label lblSaveMessage 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   540
            Left            =   240
            TabIndex        =   125
            Top             =   1080
            Width           =   8655
         End
      End
      Begin VB.Frame Frame5 
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
         Height          =   855
         Left            =   -74760
         TabIndex        =   140
         Top             =   4800
         Visible         =   0   'False
         Width           =   14775
         Begin VB.Label Label37 
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Wharfage Status"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   150
            Top             =   0
            Width           =   14775
         End
         Begin VB.Label Label45 
            Caption         =   "(1)"
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
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   5760
            TabIndex        =   147
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label46 
            Caption         =   "(4)"
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
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   11280
            TabIndex        =   146
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label47 
            Caption         =   "(3)"
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
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   9120
            TabIndex        =   145
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label48 
            Caption         =   "(2)"
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
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   7440
            TabIndex        =   144
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label49 
            Caption         =   "(5)"
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
            ForeColor       =   &H00800080&
            Height          =   255
            Left            =   5760
            TabIndex        =   143
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label56 
            Alignment       =   2  'Center
            Caption         =   "{Y/N} Wharfage Only"
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
            Height          =   375
            Left            =   1200
            TabIndex        =   142
            Top             =   600
            Width           =   3255
         End
         Begin VB.Label Label57 
            Alignment       =   2  'Center
            Caption         =   "{Y/N} Wharfage Exempt"
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
            Height          =   375
            Left            =   1200
            TabIndex        =   141
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label43 
            Caption         =   "     - BOI     - PEZA    - NAPOCOR    -  Wharfage PAID"
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
            Height          =   255
            Left            =   5520
            TabIndex        =   149
            Top             =   360
            Width           =   9015
         End
         Begin VB.Label Label44 
            Caption         =   "     - Phil. Postal Corp.        "
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
            Height          =   255
            Left            =   5520
            TabIndex        =   148
            Top             =   480
            Width           =   4575
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74760
         TabIndex        =   94
         Top             =   3600
         Width           =   14775
         Begin VB.Label Label63 
            Caption         =   "(6)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   375
            Left            =   6120
            TabIndex        =   166
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label34 
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " VAT Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   102
            Top             =   0
            Width           =   14775
         End
         Begin VB.Label Label39 
            Caption         =   "(2)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   375
            Left            =   2520
            TabIndex        =   100
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label42 
            Caption         =   "(3)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   375
            Left            =   5520
            TabIndex        =   99
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label51 
            Caption         =   "(1)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   375
            Left            =   1080
            TabIndex        =   98
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label41 
            Caption         =   "(4)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   375
            Left            =   1080
            TabIndex        =   96
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label58 
            Caption         =   "(5)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   375
            Left            =   4680
            TabIndex        =   95
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label38 
            Caption         =   "  - 0 Vat        1 - 0 Vat Less 1% W/Tax        -  0 Vat Less 2% W/Tax"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   101
            Top             =   360
            Width           =   13095
         End
         Begin VB.Label Label59 
            Caption         =   "    - 10% Vat  Less 1%  W/Tax        - 6% Vat        - 6% Vat  Less 1%  W/Tax"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1320
            TabIndex        =   97
            Top             =   720
            Width           =   13095
         End
      End
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -74760
         TabIndex        =   105
         Top             =   8880
         Width           =   14775
         Begin VB.CommandButton cmdNewCCR 
            Caption         =   "F8 - Another CCR"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   11760
            Style           =   1  'Graphical
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame16 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74760
         TabIndex        =   92
         Top             =   7080
         Width           =   14775
         Begin VB.Label Label33 
            BackColor       =   &H00800000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "BOC Permit Numbers"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   93
            Top             =   120
            Width           =   14775
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   -74760
         TabIndex        =   86
         Top             =   600
         Width           =   14775
         Begin VB.Label Label13 
            Caption         =   "Broker Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   480
            TabIndex        =   91
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label16 
            Caption         =   "Vessel :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            TabIndex        =   90
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label Label15 
            Caption         =   "Commodity :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            TabIndex        =   89
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   " Exporter Name :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   360
            TabIndex        =   88
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label54 
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            TabIndex        =   87
            Top             =   1800
            Width           =   1695
         End
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8295
         Left            =   -74880
         TabIndex        =   126
         Top             =   120
         Width           =   14895
         Begin VB.Label utxtAmtPay 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   420
            Left            =   2640
            TabIndex        =   40
            Top             =   2160
            Width           =   2415
         End
         Begin VB.Label Label30 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Initial CCR No"
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
            ForeColor       =   &H00800080&
            Height          =   450
            Left            =   2640
            TabIndex        =   139
            Top             =   600
            Width           =   4335
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Customer Name"
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
            ForeColor       =   &H00800080&
            Height          =   450
            Left            =   5280
            TabIndex        =   138
            Top             =   6000
            Width           =   6495
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Amount "
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
            ForeColor       =   &H00800080&
            Height          =   450
            Left            =   2640
            TabIndex        =   137
            Top             =   2640
            Width           =   2415
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            Caption         =   "Amount Due"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   480
            TabIndex        =   136
            Top             =   2160
            Width           =   2055
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            Caption         =   "Cash "
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1440
            TabIndex        =   135
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label26 
            Alignment       =   2  'Center
            Caption         =   "Cheque"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1080
            TabIndex        =   134
            Top             =   3600
            Width           =   1455
         End
         Begin VB.Label Label25 
            Caption         =   "ADR Amount"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   600
            TabIndex        =   133
            Top             =   7080
            Width           =   1815
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            Caption         =   "Change"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1320
            TabIndex        =   132
            Top             =   7560
            Width           =   1215
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Customer Code"
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
            ForeColor       =   &H00800080&
            Height          =   450
            Left            =   2640
            TabIndex        =   131
            Top             =   6000
            Width           =   2415
         End
         Begin VB.Label Label31 
            BackColor       =   &H00800080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " PAYMENT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   130
            Top             =   120
            Width           =   14895
         End
         Begin VB.Label lblWarning 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   16.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1020
            Left            =   7080
            TabIndex        =   129
            Top             =   600
            Width           =   7455
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Number"
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
            ForeColor       =   &H00800080&
            Height          =   450
            Left            =   5280
            TabIndex        =   128
            Top             =   3120
            Width           =   2655
         End
         Begin VB.Label Label52 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Bank"
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
            ForeColor       =   &H00800080&
            Height          =   450
            Left            =   8160
            TabIndex        =   127
            Top             =   3120
            Width           =   2655
         End
         Begin VB.Label utxtChange 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   420
            Left            =   2640
            TabIndex        =   60
            Top             =   7560
            Width           =   2415
         End
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
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
         Height          =   4095
         Left            =   600
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3600
         Width           =   14775
      End
      Begin VB.Label Label21 
         Caption         =   "PPA Total"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10680
         TabIndex        =   155
         Top             =   7380
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "ICTSI Total Dues"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9360
         TabIndex        =   154
         Top             =   6900
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label19 
         Caption         =   "Total Dues"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10440
         TabIndex        =   153
         Top             =   7860
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CCR Header Contents"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   -74760
         TabIndex        =   152
         Top             =   240
         Width           =   14775
      End
      Begin VB.Label Label50 
         Caption         =   "No. of Cnt"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   151
         Top             =   7800
         Width           =   1815
      End
      Begin VB.Label lblNoCnt 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   2160
         TabIndex        =   15
         Top             =   7800
         Width           =   735
      End
      Begin VB.Label txtIctsiDue 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   12480
         TabIndex        =   12
         Top             =   6840
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label txtPpaTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   12480
         TabIndex        =   13
         Top             =   7320
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label txtTotDue 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   12480
         TabIndex        =   14
         Top             =   7800
         Width           =   2535
      End
   End
   Begin VB.Frame Frame15 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   156
      Top             =   6240
      Width           =   15015
      Begin VB.TextBox utxtWorkStn 
         BackColor       =   &H8000000F&
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
         Height          =   420
         Left            =   3240
         TabIndex        =   189
         TabStop         =   0   'False
         Top             =   360
         Width           =   4935
      End
      Begin CCRDE06.utxtTextBilling txtTranMode 
         Height          =   420
         Left            =   3240
         TabIndex        =   158
         Top             =   1800
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin CCRDE06.utxtTextBilling txtSupervisor 
         Height          =   420
         Left            =   3240
         TabIndex        =   157
         Top             =   1320
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   741
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   19
      End
      Begin VB.TextBox txtUserid 
         BackColor       =   &H8000000F&
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
         Height          =   420
         Left            =   3240
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   840
         Width           =   4215
      End
      Begin VB.ListBox cmbPrinter 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   159
         Top             =   2280
         Width           =   9615
      End
      Begin VB.Label Label23 
         Caption         =   "Workstation ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   190
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label62 
         Caption         =   "Supervisor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   163
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label61 
         Caption         =   "Teller"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   162
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label60 
         Caption         =   "Transaction Mode                     F - Foreign / D - Domestic"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   161
         Top             =   1800
         Width           =   9015
      End
      Begin VB.Label Label53 
         Caption         =   "Printer to Use"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   160
         Top             =   2280
         Width           =   2415
      End
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   78
      Top             =   11475
      Width           =   15525
      _ExtentX        =   27384
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
            Object.Width           =   14049
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2990
            Text            =   "CCRDE01"
            TextSave        =   "CCRDE01"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "11/6/2017"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "3:27 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame11 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   80
      Top             =   3000
      Width           =   15015
      Begin VB.Frame Frame17 
         Caption         =   "CCR Last Issued"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   182
         Top             =   1680
         Width           =   13095
         Begin VB.TextBox utxtLastCCRISI 
            BackColor       =   &H8000000F&
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
            Height          =   420
            Left            =   3000
            TabIndex        =   187
            TabStop         =   0   'False
            Top             =   840
            Width           =   3615
         End
         Begin VB.TextBox utxtLastIssuedDteISI 
            BackColor       =   &H8000000F&
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
            Height          =   420
            Left            =   6720
            TabIndex        =   186
            TabStop         =   0   'False
            Top             =   840
            Width           =   5535
         End
         Begin VB.TextBox utxtLastCCR 
            BackColor       =   &H8000000F&
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
            Height          =   420
            Left            =   3000
            TabIndex        =   184
            TabStop         =   0   'False
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox utxtLastIssuedDte 
            BackColor       =   &H8000000F&
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
            Height          =   420
            Left            =   6720
            TabIndex        =   183
            TabStop         =   0   'False
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label Label69 
            Alignment       =   1  'Right Justify
            Caption         =   "ISI"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   188
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "SBITC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   480
            TabIndex        =   185
            Top             =   480
            Width           =   2415
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "CCR Allocation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   173
         Top             =   240
         Width           =   13095
         Begin VB.TextBox utxtStrtCCRISI 
            BackColor       =   &H8000000F&
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
            Height          =   420
            Left            =   3000
            TabIndex        =   179
            TabStop         =   0   'False
            Top             =   840
            Width           =   3615
         End
         Begin VB.TextBox utxtEndCCRISI 
            BackColor       =   &H8000000F&
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
            Height          =   420
            Left            =   6960
            TabIndex        =   178
            TabStop         =   0   'False
            Top             =   840
            Width           =   4455
         End
         Begin VB.TextBox utxtStrtCCR 
            BackColor       =   &H8000000F&
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
            Height          =   420
            Left            =   3000
            TabIndex        =   175
            TabStop         =   0   'False
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox utxtEndCCR 
            BackColor       =   &H8000000F&
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
            Height          =   420
            Left            =   6960
            TabIndex        =   174
            TabStop         =   0   'False
            Top             =   360
            Width           =   4455
         End
         Begin VB.Label Label68 
            Alignment       =   1  'Right Justify
            Caption         =   "ISI"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   181
            Top             =   960
            Width           =   2775
         End
         Begin VB.Label Label67 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            TabIndex        =   180
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "SBITC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   177
            Top             =   480
            Width           =   2775
         End
         Begin VB.Label Label35 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            TabIndex        =   176
            Top             =   480
            Width           =   135
         End
      End
   End
   Begin VB.Frame Frame9 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   79
      Top             =   9240
      Width           =   15015
      Begin VB.CommandButton cmdExit1 
         Caption         =   "F3 - Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdContinue 
         Caption         =   "F7 - Continue"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Label lblAlloc 
      Caption         =   "No Allocation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   75
      Top             =   2520
      Width           =   7575
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu ChangePinter 
         Caption         =   "Change &Printer"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MainExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmCCRde06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim sqlConBilling As String
Dim sqlConNavis As String

Dim CCRList(1 To 100) As CList
Dim gCusnam As String
Dim lRef As Integer
Dim chkUG As String
' ** Reciept Variables
Dim mvarCCRNumber As Long
Dim StrChk1 As String * 30
Dim StrChk2 As String * 30
Dim StrChk3 As String * 30
Dim StrChk4 As String * 30
Dim StrChk5 As String * 30
' ** ADR Variables
Dim blnnumberin As Boolean
Dim sngPreviousADR As Currency
Dim sngAdrBalance As Currency
Dim strCustName As String
' ** Misc Class Declarations
Dim cusListing As cCustomer 'Customer PickList
'Dim CTCSinfo As Object
Dim blnF3KeyPressed As Boolean
Dim lngRow As Long  'row Fill Grid
Dim lngCtrRow As Long
Dim lngNoCnt As Long
'ICTSI Total
Dim sngIctArr As Currency
Dim sngIctWgh As Currency
Dim sngIctVat As Currency
Dim sngIctWtax As Currency
Dim sngIctTot As Currency
Dim sngPpa As Currency
Dim sngGrndTot As Currency
'Details Temporary variables
Dim sngArr As Currency 'Total Arrastre
Dim sngWhf As Currency 'Wharfage Amount
Dim sngBscArr As Currency 'Basic Arrastre
Dim sngBscWhf As Currency  'Basic Wharfage
Dim sngWeighing As Currency
Dim sngWghAmt As Currency
Dim lngItemNum As Long
Dim lngRefnum As Long
Dim sngOvzAmt As Currency 'Oversize Amount
Dim sngVat As Currency 'Vat Amount
Dim sngdangramt As Currency 'Danger Class Amount
Dim sngWtax As Currency
Dim sngOvzLength As Single
Dim sngOvzWidth As Single
Dim sngOvzHeight As Single
Dim sngDollarAmt As Currency
Dim sngRton As Currency
Dim strSBMAPermit As String * 12 ' SBMA Permit Number
Dim strCntNum As String * 12 'Container Number
Dim strCommodity As String  'Commodity
Dim strExporter As String  'Exporter
Dim strBroker As String   'Broker
Dim strVessel As String  'Vessel Code
Dim strUg As String  'Under guarantee Code
Dim strVatCode As String  'Vat Code
Dim strWhfCode As String  'Wharfage exempt or Not
Dim strExmCode As String  'Exemption Code
Dim strTshpCode As String  'Transhipment Code
Dim strFulemp As String  'Full / Empty Code
Dim strFCCR As String  'Forced CCR Code
Dim strWhfOnly As String  'Wharfage Only
Dim strDate As String 'Date Value
'   ** Temporary Entry field
Dim strEntry0 As String * 8
Dim strEntry1 As String * 8
Dim strEntry2 As String * 8
Dim strEntry3 As String * 8
Dim strEntry4 As String * 8
Dim strEntry5 As String * 8
Dim strEntry6 As String * 8
Dim strEntry7 As String * 8
Dim strEntry8 As String * 8
Dim strEntry9 As String * 8
Dim strEntry10 As String * 8
Dim strEntry As String
Dim strRemark As String
Dim intSize As Integer
Dim strDangr As String
Dim strUms As String
Dim lngPrevCCR As Long
Dim lngOvrCCr As Long
Dim strPrevPay As String * 1
Dim lngCCRStart As Long
Dim lngCCREnd As Long
Dim lngCCRLastIssued As Long
Dim dtmCCRLastIssuedDate As String

'PRNH - Company Code
Dim strCompCode As String


'   **  Expor21 Temporary Variables
Dim strExpr21Contnum As String * 12
Dim lngExpr21Refnum As Long
Dim lngExpr21Date As Long
Dim lngExpr21Time As Long
Dim strExpr21Userid As String
Dim strExpr21Trncde As String
Dim strExpr21ExpName As String
Dim strExpr21Whfcde As String
Dim TLength As Single
Dim TWidth As Single
Dim THeight As Single
Dim TUms As String
Dim blnChanging As Boolean
Const defaultUnitMeasurement As String = "I"
Event Closing()
Event OutMain()
Event InMain()

Private Sub ChangePinter_Click()
    frmPrinter.Show vbModal
    Call getUserId
End Sub


Private Sub Form_Load()

Call Main
    txtSupervisor.Text = gSuprvsr
    Call StartInitialization
    Call getUserId
    Call GetSparcsN4Host
    SB.Visible = False
    blnExit = True
    ChangePinter.Enabled = False
End Sub

Private Sub Main()
    Dim mp As clsCCRde06
    'sharon
    Call ReadConfig
    
    Call gzCurrentUser
    gConnStr = sqlConBilling
    
'
'    "Provider=sqloledb" & _
'        ";Data Source=sbitcbilling" & _
'        ";Initial Catalog=BILLING" & _
'        ";Integrated Security=SSPI"
'        ";UID=tosadmin; password=password"


'         gConnStr = "Provider=sqloledb" & _
'        ";Data Source=sbitc-dev" & _
'        ";Initial Catalog=sbitcbilling" & _
'        ";UID=sa_ictsi; password=Ictsi123"
        
       '";Integrated Security=SSPI"
    Set mp = New clsCCRde06
    mp.Userid = gzCurrentUser
    mp.ConnectByStr gConnStr

    Set mp = Nothing
End Sub


Public Function gzCurrentUser() As String
Dim lpUserName As String * 64
    If WNetGetUser("", lpUserName, Len(lpUserName)) Then
        gzCurrentUser = ""
    Else
        gzCurrentUser = Left(lpUserName, InStr(lpUserName, Chr(0)) - 1)
    End If
End Function

Public Sub ReadConfig()
Dim Xcnt As Integer
Open App.Path & "\" & "Conn.cfg" For Binary Access Read As #1

Do While Not EOF(1)
    Xcnt = Xcnt + 1
    Select Case Xcnt
        Case 1
            Line Input #1, sqlConBilling
        Case 2
            Line Input #1, sqlConNavis
    End Select
Loop
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If SB.Visible = False Then
        If blnExit Then
            Cancel = 0
        Else
            MsgBox "Exit"
            Cancel = 1
        End If
    Else
        Cancel = 1
    End If
End Sub
Public Sub getUserId()
    Dim rsUsr As Recordset
        DE.getInformation
        Set rsUsr = DE.rsgetInformation
        strWrkstn = rsUsr.Fields("workstation")
        strPrinter = Printer.DeviceName
        SB.Panels(1) = UCase(gUserID)
        SB.Panels(2) = strWrkstn
        SB.Panels(3) = "Printer Device :" & strPrinter
        rsUsr.Close
        Set rsUsr = Nothing
    Exit Sub
ErrgetUserId:
    Beep
    MsgBox "Contact any MIS Support Staff", vbExclamation + vbCritical, "Error Connection"
End Sub

Private Sub MainExit_Click()
    Call cmdExit1_Click
End Sub
Private Sub cmdNewCCR_Click()
    Call CheckPaymentOk
    Call UpdateHeaderGrid
    Call EnableTotalVisible(True)
    Call ResetValTab1A
    sstMain.Tab = 0
    utxtDollar.Value = 0
    utxtLength.Value = "0"
    utxtWidth.Value = "0"
    utxtHeight.Value = "0"
    utxtUMS.Text = defaultUnitMeasurement
    utxtNumDangr.Value = ""
    utxtTshipMnt.Text = ""
    chkNewCCR.Value = vbChecked
    lngTagNewCCR = lngTagNewCCR + 1
End Sub

Private Sub txtSupervisor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
            
    End If
End Sub
Private Sub txtTranMode_Change()
    If Trim(txtTranMode.Text) = "D" Then
        DomesticMode = True
        cMode.Text = "DOMESTIC Transaction"
    Else
        DomesticMode = False
        cMode.Text = "FOREIGN Transaction"
    End If
End Sub
Private Sub txtTranMode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 100 ' * d
            txtTranMode.Text = "D"
        Case 68 ' * D
            txtTranMode.Text = "D"
        Case 102 ' * f
            txtTranMode.Text = "F"
        Case 70 ' * F
            txtTranMode.Text = "F"
        Case Else
            Beep
    End Select
    KeyAscii = 0
End Sub


Private Sub utxtCommodity_Change()
    Call CheckPaymentOk
End Sub

Private Sub utxtCustNo1_Change()
 If Len(Trim(utxtCustNo1.Value)) <> 0 Then
    DE.getCustomerName Trim(utxtCustNo1.Value), strCustName
    utxtCustName1.Text = strCustName
    utxtCustNo.Value = utxtCustNo1.Value
    utxtCustName.Text = utxtCustName1.Text
 Else
    utxtCustName1.Text = ""
    'utxtCustNo.Value = ""
    'utxtCustName.Text = ""
 End If
End Sub

Private Sub utxtEntry1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        utxtExporter.SetFocus
    Else
        If KeyCode = vbKeyDown Then
            utxtVatCode.SetFocus
        End If
    End If
    Call CheckPaymentOk
End Sub
Private Sub utxtExporter_Change()
Call CheckPaymentOk
End Sub
Private Sub utxtHeight_LostFocus()
    If Len(Trim(utxtHeight.Value)) = 0 Then
        utxtHeight.Value = 0
    End If
End Sub
Private Sub utxtLength_LostFocus()
    If Len(Trim(utxtLength.Value)) = 0 Then
        utxtLength.Value = 0
    End If
End Sub
Private Sub utxtPref_Change()
    If utxtPref.Enabled Then
        If Len(utxtPref.Text) = 4 Then
            utxtNo.SetFocus
        End If
    End If
End Sub
Private Sub utxtTshipMnt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            If utxtTshipMnt.Text = "F" Then
              utxtDollar.Visible = True
              Label7.Visible = True
              utxtDollar.Enabled = True
              utxtDollar.TabStop = True
              utxtDollar.SetFocus
            End If
    End Select
End Sub

Private Sub utxtUGuarantee_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If Not UCase(Chr(KeyAscii)) = "Y" And _
            Not UCase(Chr(KeyAscii)) = "N" Then
                Beep
                KeyAscii = 0
                utxtUGuarantee.SetFocus
        Else
            utxtUGuarantee.Text = UCase(Chr(KeyAscii))
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub cmbPrinter_Click()
    Set Printer = Printers(cmbPrinter.ListIndex)
    PrinterRef = cmbPrinter.ListIndex
End Sub
Private Sub cmdBack_Click()
    Dim strBack1 As String
    Dim strBack2 As String
    Dim strBack3 As String
    strBack1 = ""
    strBack2 = ""
    strBack3 = ""
    strBack2 = "NEXT DATA FOR NEW CCR ?"
    If MsgBox(strBack2, vbDefaultButton2 + vbYesNo + vbQuestion) = vbYes Then
        strResponse = True
    Else
        strResponse = False
    End If
    If strResponse = True Then
        Call cmdNewCCR_Click
    Else
        Call UpdateHeaderGrid
        Call EnableTotalVisible(True)
        sstMain.Tab = 0
    End If
End Sub
Private Sub cmdCancel_Click()
    Call MoveToText
    cmdEdit.Caption = "F5 - Edit"
    cmdAdd.Caption = "F4 - Add"
    cmdDelete.Caption = "F9 - Cont. No"
    cmdCancel.Visible = False
    cmdAdd.Enabled = True
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    Call CheckPaymentOk
    cmdHeader.Enabled = lngRow > 0 And flexDetails.Enabled = False
    cmdGrid.Enabled = lngRow > 0 And flexDetails.Enabled = False
    Call EnableForNewContainer(True)
    flexDetails.ColSel = 0
    flexDetails.HighLight = flexHighlightNever
    utxtPref.SetFocus
End Sub
Private Sub cmdContinue_Click()
    sstMain.Visible = True
    Call MainProcedure
    Call Tab00off(True)
    Call Tab01off(False)
    Call Tab02off(False)
    Call EnableForNewContainer(True)
    SB.Visible = True
    Call getUserId
    blnExit = False
    MainExit.Enabled = False
    ChangePinter.Enabled = True
    sstMain.Tab = 0
    Call GridHeader
    utxtPref.SetFocus
End Sub
Private Sub cmdExit1_Click()
    If MsgBox("Do you want to Exit from the Program ?", vbExclamation + vbYesNo, "Exit Confirmation") = vbYes Then
        Unload Me
    End If
End Sub
Private Sub cmdHeader_Click()
    Call ClearTotals
    Call ClearTagCCR
    sstMain.Tab = 1
    'Call Sparcs_Commodity(strCntNum)
    'Call Sparcs_VesselCode(strCntNum)
End Sub
Private Sub cmdDelete_Click()
    Dim strTagCde As String * 1
    Select Case cmdDelete.Caption
        Case "F9 - Cont. No"
            ' ** reset contents
            Call InitializeDetail
            Call EnableForNewContainer(True)
            cmdEdit.Enabled = False
            cmdDelete.Enabled = False
            cmdCancel.Visible = False
            flexDetails.TabStop = False
            flexDetails.Enabled = False
            Call CheckPaymentOk
        Case "F9 - Delete"
            lngNoCnt = lngNoCnt - 1
            lblNoCnt.Caption = lngNoCnt
            lblNoCnt.Refresh
            If lngRow = 1 Then
                lngRow = 0
                Call InitializeDetail
                cmdEdit.Enabled = False
                cmdGrid.Enabled = False
                cmdDelete.Enabled = False
                cmdHeader.Enabled = lngRow > 0 And flexDetails.Enabled = False
                cmdDelete.Caption = "F9 - Cont. No"
                chkNewCCR.Value = vbChecked
                flexDetails.Visible = False
                flexDetails.Clear
                flexDetails.Refresh
                lngTagNewCCR = 0
                utxtPref.Text = ""
                utxtNo.Text = ""
                Call GridHeader
                Call ClearTotals
                Call Clear_GrandTotal
            Else
                With flexDetails
                    
                    If .Row < lngRow Then
                        If .TextMatrix(.Row, 0) = "*" And .TextMatrix(.Row + 1, 0) <> "*" Then
                            .RemoveItem (.Row)
                            .TextMatrix(.Row, 0) = "*"
                            lngRow = lngRow - 1
                        Else
                            .RemoveItem (.Row)
                            lngRow = lngRow - 1
                        End If
                    Else
                        lngTagNewCCR = .TextMatrix(.Row - 1, 41)
                        .RemoveItem (.Row)
                        lngRow = lngRow - 1
                    End If
                    Call ReSequence_ItemNum
                    Call MoveToText
                    Call ClearTotals
                    Call Clear_GrandTotal
                    Call ComputeTotal
                End With
                cmdGrid.Enabled = True
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
            End If
            cmdAdd.Caption = "F4 - Add"
            cmdDelete.Caption = "F9 - Cont. No"
            Call EnableForNewContainer(True)
            flexDetails.ColSel = 0
            flexDetails.HighLight = flexHighlightNever
            flexDetails.TabStop = False
            flexDetails.Enabled = False
            utxtPref.SetFocus
            cmdHeader.Enabled = lngRow > 0 And flexDetails.Enabled = False
            Call CheckPaymentOk
    End Select
End Sub
Private Sub cmdEdit_Click()
    Select Case cmdEdit.Caption
        Case "F5 - Edit"
                Call MoveToText
                Call EnableForNewContainer(False)
                Call Tab00off(True)
                Call Tab01off(False)
                Call Tab02off(False)
                flexDetails.Enabled = True
                utxtPref.TabStop = False
                utxtNo.TabStop = False
                utxtPref.Enabled = False
                utxtNo.Enabled = False
                chkNewCCR.Enabled = False
                utxtSze.SetFocus
                cmdEdit.Caption = "F5 - Save"
                cmdCancel.Visible = True
                cmdAdd.Enabled = False
                cmdGrid.Enabled = False
                flexDetails.TabStop = False
                flexDetails.Enabled = False
                cmdDelete.Enabled = False
                cmdPayment.Enabled = False
                cmdHeader.Enabled = False
        Case "F5 - Save"
                cmdEdit.Caption = "F5 - Edit"
                cmdAdd.Caption = "F4 - Add"
                strCntNum = (utxtPref.Text) & Trim(utxtNo.Text)
                strCntNum = Trim(strCntNum)
                Call MoveDetailToVariables
                Call ClearTotals
                Call SaveToGrid
                Call UpdatePerDet(flexDetails.Row)
                'Call UpdateGrid
                Call InitializeDetail
                Call ClearVariables
                Call EnableForNewContainer(True)
                ' ** for cmdPayment
                Call CheckPaymentOk
                cmdHeader.Enabled = lngRow > 0 And flexDetails.Enabled = False
                cmdEdit.Enabled = False
                cmdDelete.Enabled = False
                cmdGrid.Enabled = True
                cmdCancel.Visible = False
                flexDetails.TabStop = False
                flexDetails.Enabled = False
                flexDetails.ColSel = 0
                flexDetails.HighLight = flexHighlightNever
                utxtPref.SetFocus
    End Select
End Sub
Private Sub cmdExit_Click()
   If MsgBox("Are you sure you want to Reset the Transaction ?", vbExclamation + vbYesNo, "Exit Confirmation") = vbYes Then
        blnF3KeyPressed = True
        sstMain.Visible = False
        txtSupervisor.TabStop = True
        txtTranMode.TabStop = True
        cmbPrinter.TabStop = True
        cMode.Visible = False
        
        cmbPrinter.ListIndex = PrinterRef
        SB.Visible = False
        blnExit = True
        MainExit.Enabled = True
        ChangePinter.Enabled = False
        Call GetAllocation
    End If
End Sub
Private Sub cmdGrid_Click()
    Call Tab00off(False)
    Call EnableForNewContainer(True)
    cmdEdit.Enabled = True
    cmdDelete.Enabled = True
    cmdAdd.Enabled = True
    cmdGrid.Enabled = False
    cmdHeader.Enabled = lngRow > 0 And flexDetails.Enabled = False
    cmdPayment.Enabled = False
    cmdAdd.Caption = "F4 - Cont. No."
    cmdDelete.Caption = "F9 - Delete"
    utxtPref.Enabled = False
    utxtNo.Enabled = False
    Call MoveToText
    flexDetails.Enabled = True
    flexDetails.TabStop = True
    flexDetails.Row = 1
    flexDetails.Col = 0
    flexDetails.ColSel = 40
    flexDetails.SelectionMode = flexSelectionByRow
    flexDetails.HighLight = flexHighlightAlways
    flexDetails.SetFocus
End Sub
Private Sub cmdAdd_Click()
    Dim ForcedCode As String * 1
    Dim UsrCode As String * 1
    Select Case cmdAdd.Caption
        Case "F4 - Add"
            ForcedCode = ""
            UsrCode = ""
            If Check_Details <> False Then
                If Container_Valid Then
                    lngItemNum = lngItemNum + 1
                    If chkNewCCR.Value = vbChecked Then
                        UsrCode = "*"
                        lngItemNum = 1
                    Else
                        If lngItemNum > 8 Then
                            ForcedCode = "F"
                            UsrCode = "*"
                            lngItemNum = 1
                        Else
                            If lngRow = 0 Then
                                UsrCode = "*"
                                lngItemNum = 1
                            Else
                                ForcedCode = ""
                                UsrCode = ""
                            End If
                        End If
                    End If
                    lngRow = lngRow + 1
                    lngNoCnt = lngNoCnt + 1
                    lblNoCnt.Caption = lngNoCnt
                    lblNoCnt.Refresh
                    Call MoveDetailToVariables
                    Call FillGrid
                    Call ClearTotals
                    Call UpdateGrid
                    Call InitializeDetail
                    Call ClearVariables
                    flexDetails.TextMatrix(lngRow, 35) = ForcedCode
                    flexDetails.TextMatrix(lngRow, 0) = UsrCode
                    If lngRow > 7 Then
                        flexDetails.TopRow = lngRow - 6
                    End If
                    flexDetails.Row = lngRow
                    flexDetails.Col = 0
                    flexDetails.ColSel = 40
                    flexDetails.SelectionMode = flexSelectionByRow
                    flexDetails.HighLight = flexHighlightAlways
                    Call ReSequence_ItemNum
                    cmdAdd.Enabled = True
                    cmdGrid.Enabled = True
                    cmdHeader.Enabled = lngRow > 0 And flexDetails.Enabled = False
                    Call CheckPaymentOk
                    Call EnableForNewContainer(True)
                    utxtPref.SetFocus
                    chkNewCCR.Value = vbUnchecked
                End If
            End If
        Case "F4 - Cont. No."
            cmdAdd.Caption = "F4 - Add"
            cmdDelete.Caption = "F9 - Cont. No."
            cmdEdit.Enabled = False
            cmdGrid.Enabled = True
            cmdDelete.Enabled = False
            cmdCancel.Visible = False
            flexDetails.TabStop = False
            flexDetails.Enabled = False
            flexDetails.ColSel = 0
            flexDetails.HighLight = flexHighlightNever
            cmdHeader.Enabled = lngRow > 0 And flexDetails.Enabled = False
            Call InitializeDetail
            Call EnableForNewContainer(True)
            utxtPref.SetFocus
            Call CheckPaymentOk
    End Select
End Sub
Function Container_Valid() As Boolean
    Dim strChkCont1 As String
    Dim strChkcont2 As String
    strChkCont1 = ""
    strChkcont2 = ""
    Container_Valid = False
    lngCtrRow = 0
    If flexDetails.TextMatrix(lngCtrRow, 2) = "" Then
        Container_Valid = True
        Exit Function
    Else
        Container_Valid = True
    End If
    If lngRow > 0 Then
        Do While Not (lngCtrRow = lngRow)
            lngCtrRow = lngCtrRow + 1
            If strCntNum = (flexDetails.TextMatrix(lngCtrRow, 2)) Then
                strChkCont1 = "Inputted Container Already exist in Detail"
                Call ErrorMessage(strChkCont1, strChkcont2)
                utxtPref.SetFocus
                Container_Valid = False
                Exit Do
            End If
        Loop
    End If
End Function
Private Sub cmdPayment_Click()
    sstMain.Tab = 2
End Sub
Private Sub cmdPrint_Click()
    Dim intError As ADODB.Error
    If CCur(utxtChange.Caption) < 0 Then
        utxtCsh.SetFocus
    Else
        ' ** get control number
        lngRefnum = GetNextControlNumber
        ' ** global refnum variable
        lblSaveMessage.Caption = "Transaction being Saved. Please Wait . . . . . . "
        lblSaveMessage.Refresh
                                      
        If SaveToCCRPay Then
             If SavetoCCRCyx Then
                mvarCCRNumber = 0
                lblSaveMessage.Caption = "Printing . . . Please Wait . . . . . . "
                lblSaveMessage.Refresh
                PrintCCR (lngRefnum)
                lRef = 0
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        lblSaveMessage.Caption = ""
        cmdPrint.Enabled = False
        Call ClearTotals
        Call ResetProcedure
        Call EnableForNewContainer(True)
        sstMain.Tab = 0
        sngAdrBalance = 0
        strCustName = ""
        utxtCCRNo.Value = ""
        chkUG = ""
    End If
    Exit Sub
ErrorSave:
    Beep
    MsgBox "Error : " & intError.Number & " - " & intError.Description & " Transaction Failed", vbExclamation
End Sub
Private Sub cmdPymBack_Click()
    sstMain.Tab = 0
End Sub

Private Sub cmdPymCancel_Click()
    If MsgBox("Are you sure you want to CANCEL THE TRANSACTION?", vbCritical + vbYesNo, "Cancel Transaction") = vbYes Then
        Call ClearTotals
        Call cmdContinue_Click
    End If
End Sub
Private Sub cmdRefresh_Click()
    Call GetAllocation
End Sub
Private Sub flexDetails_RowColChange()
    If cmdGrid.Enabled = False Then
        Call MoveToText
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF2
        If sstMain.Visible = True And sstMain.Tab = 0 And cmdGrid.Enabled = True Then
            Call cmdGrid_Click
        End If
    Case vbKeyF3
        If sstMain.Visible = True And sstMain.Tab = 0 Then
            Call cmdExit_Click
            Exit Sub
        ElseIf sstMain.Visible = False Then
            Call cmdExit1_Click
        End If
    Case vbKeyF4
       If sstMain.Visible = True And sstMain.Tab = 0 And cmdAdd.Enabled = True Then
            Call cmdAdd_Click
        End If
    Case vbKeyF5
        If sstMain.Visible = True And sstMain.Tab = 0 And cmdEdit.Enabled = True Then
            Call cmdEdit_Click
        ElseIf sstMain.Visible = False Then
            Call GetAllocation
        End If
    Case vbKeyF6
        If sstMain.Visible = True And sstMain.Tab = 0 And (cmdCancel.Visible = True And cmdCancel.Enabled = True) Then
            Call cmdCancel_Click
        ElseIf sstMain.Visible = True And sstMain.Tab = 2 And cmdPymBack.Enabled = True Then
            Call cmdPymBack_Click
        End If
    Case vbKeyF7
        If sstMain.Visible = False And cmdContinue.Enabled = True Then
            Call cmdContinue_Click
        ElseIf sstMain.Visible = True And sstMain.Tab = 0 And cmdHeader.Enabled = True Then
                Call cmdHeader_Click
        ElseIf sstMain.Visible = True And sstMain.Tab = 1 And cmdBack.Enabled = True Then
            Call cmdBack_Click
        ElseIf sstMain.Visible = True And sstMain.Tab = 2 And cmdPrint.Enabled = True Then
            Call cmdPrint_Click
        End If
    Case vbKeyF8
        If sstMain.Visible = True And sstMain.Tab = 1 And cmdNewCCR.Enabled = True Then  'And cmdNewCCR.Enabled = True  Then
            Call cmdNewCCR_Click
        End If
    Case vbKeyF9
        If cmdDelete.Enabled = True Then
            Call cmdDelete_Click
            utxtPref.SetFocus
        End If
    Case vbKeyF12
        If sstMain.Visible = True And sstMain.Tab = 0 And cmdPayment.Enabled = True Then
            Call cmdPayment_Click
        ElseIf sstMain.Visible = True And sstMain.Tab = 2 Then
            Call cmdPymCancel_Click
        End If
End Select
End Sub
Private Sub sstMain_Click(PreviousTab As Integer)
    blnF3KeyPressed = False
    Select Case sstMain.Tab
        Case 0
            Call Tab00off(True)
            Call Tab01off(False)
            Call Tab02off(False)
            Call EnableForNewContainer(True)
            Call CheckPaymentOk
            utxtPref.SetFocus
        Case 1
            Call Tab00off(False)
            Call Tab01off(True)
            Call Tab02off(False)
            utxtExporter.SetFocus
        Case 2
            Call Tab00off(False)
            Call Tab01off(False)
            Call Tab02off(True)
            utxtAmtPay.Caption = Format(CCur(txtTotDue.Caption), "###,##0.00")
            utxtCsh.Value = txtTotDue.Caption
            If Not IsNumeric(utxtCCRNo.Value) Then
                'PRNH - OLD
                'lngCCR = GetNextCCRNumber
                lngCCR = GetNextCCRNumber(flexDetails.TextMatrix(1, 45))
            Else
                If Val(utxtCCRNo.Value) <= 0 Then
                    'PRNH - OLD
                    'lngCCR = GetNextCCRNumber
                    lngCCR = GetNextCCRNumber(flexDetails.TextMatrix(1, 45))
                Else
                    'PRNH - OLD
                    'lngCCR = GetNextCCRNumber
                    lngCCR = GetNextCCRNumber(flexDetails.TextMatrix(1, 45))
                End If
            End If
            
            If lngCCR > 0 Then
                utxtCCRNo.Value = lngCCR
            Else
                utxtCCRNo.Value = ""
            End If
            utxtChange.Caption = EvaluateChange
            If sngAdrBalance <> 0 Then
                utxtAdrAmt.TabStop = True
                utxtAdrAmt.Enabled = True
            Else
                utxtAdrAmt.TabStop = False
                'utxtAdrAmt.Enabled = False
            End If
            utxtCsh.SetFocus
            NumberOfCCR = GetNumberOfCCR
            Call CheckAllocationRange
    End Select
End Sub
Private Sub AddDanger_Class()
With flexDangerClass
    .ColAlignment(0) = 0
    .TextMatrix(1, 0) = "(1) Explosives"
    .TextMatrix(2, 0) = "(2) Gases"
    .TextMatrix(3, 0) = "(3) Inflammable Liquid"
    .TextMatrix(4, 0) = "(4) Inflammable Solid"
    .TextMatrix(5, 0) = "(5) Oxidizing Agents / Organic Peroxide"
    .TextMatrix(6, 0) = "(6) Poisonous(toxic) and Infectious Substances"
    .TextMatrix(7, 0) = "(7) Radoiactive Substances"
    .TextMatrix(8, 0) = "(8) Corrosives"
    .TextMatrix(9, 0) = "(9) Miscellaneous Dangerous Substances"
End With
End Sub
Private Sub AddTshipMent_Code()
With flexTshipMnt
    .ColAlignment(0) = 0
    .TextMatrix(1, 0) = "  - Regular Container"
    .TextMatrix(2, 0) = "R - Relay"
    .TextMatrix(3, 0) = "F - Foreign Transshipment"
End With
End Sub
Private Sub MoveRates()
Dim lngCtr As Long
Dim rstCYRte As Recordset
Set rstCYRte = New Recordset
' ** Call to use CY Rate file (CYRate)
DE.CyRate
' ** Return Recordset
Set rstCYRte = DE.rsCyRate
lngCtr = 0
Do Until rstCYRte.EOF
    lngCtr = lngCtr + 1
    RateArr(lngCtr).Rtecode = rstCYRte.Fields("cyr_rtecde")
    If Not IsNull(rstCYRte.Fields("cyr_cntsze")) And (rstCYRte.Fields("cyr_cntsze")) <> "" Then
        RateArr(lngCtr).CntSze = Val(rstCYRte.Fields("cyr_cntsze"))
    Else
        RateArr(lngCtr).CntSze = 0
    End If
    RateArr(lngCtr).RteAmt = rstCYRte.Fields("cyr_rteamt")
    rstCYRte.MoveNext
Loop
rstCYRte.Close
Set rstCYRte = Nothing
End Sub
Private Sub GridHeader()
    With flexDetails
        .Redraw = False
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 300
        .TextMatrix(0, 1) = "CCR #"
        .ColWidth(1) = 1
        .TextMatrix(0, 2) = "Container Number"
        .ColWidth(2) = 2800
        .TextMatrix(0, 3) = "Arrastre"
        .ColWidth(3) = 2000
        .TextMatrix(0, 4) = "Vat"
        .ColWidth(4) = 0 '1500
        .TextMatrix(0, 5) = "Tax "
        .ColWidth(5) = 1400
        .TextMatrix(0, 6) = "Weighing"
        .ColWidth(6) = 1500
        .TextMatrix(0, 7) = "U/G "
        .ColWidth(7) = 600
        .TextMatrix(0, 8) = "Basic Arrastre"
        .ColWidth(8) = 1
        .TextMatrix(0, 9) = "Size"
        .ColWidth(9) = 1
        .TextMatrix(0, 10) = "Full / Empty"
        .ColWidth(10) = 1
        .TextMatrix(0, 11) = "Danger Class"
        .ColWidth(11) = 1
        .TextMatrix(0, 12) = "Danger Amount"
        .ColWidth(12) = 1
        .TextMatrix(0, 13) = "Vessel Code"
        .ColWidth(13) = 1
        .TextMatrix(0, 14) = "Vat Code"
        .ColWidth(14) = 1
        .TextMatrix(0, 15) = "Oversize Length"
        .ColWidth(15) = 1
        .TextMatrix(0, 16) = "Oversize Width"
        .ColWidth(16) = 1
        .TextMatrix(0, 17) = "Oversize Height"
        .ColWidth(17) = 1
        .TextMatrix(0, 18) = "UMS"
        .ColWidth(18) = 1
        .TextMatrix(0, 19) = "Oversize Amt"
        .ColWidth(19) = 1
        .TextMatrix(0, 20) = "Tonnage"
        .ColWidth(20) = 1300
        .TextMatrix(0, 21) = "Transhipment"
        .ColWidth(21) = 1
        .TextMatrix(0, 22) = "Consolidation Code"
        .ColWidth(22) = 1
        .TextMatrix(0, 23) = "Exporter"
        .ColWidth(23) = 3500
        .TextMatrix(0, 24) = "Broker"
        .ColWidth(24) = 1
        .TextMatrix(0, 25) = "Entry"
        .ColWidth(25) = 1
        .TextMatrix(0, 26) = "Commodity"
        .ColWidth(26) = 1
        .TextMatrix(0, 27) = "Status"
        .ColWidth(27) = 1
        .TextMatrix(0, 28) = "Rectag"
        .ColWidth(28) = 1
        .TextMatrix(0, 29) = "Teller"
        .ColWidth(29) = 1
        .TextMatrix(0, 30) = "System Date"
        .ColWidth(30) = 1
        .TextMatrix(0, 31) = "Update Code"
        .ColWidth(31) = 1
        .TextMatrix(0, 32) = "Out Date/Time"
        .ColWidth(32) = 1
        .TextMatrix(0, 33) = "Prev CCR"
        .ColWidth(33) = 1
        .TextMatrix(0, 34) = "Dollar"
        .ColWidth(34) = 1
        .TextMatrix(0, 35) = "Forced CCR"
        .ColWidth(35) = 1
        .TextMatrix(0, 36) = "Item No"
        .ColWidth(36) = 1
        .TextMatrix(0, 37) = "PrevPay"
        .ColWidth(37) = 1
        .TextMatrix(0, 38) = "Remark"
        .ColWidth(38) = 1
        .TextMatrix(0, 39) = "HeaderTag"
        .ColWidth(39) = 1
        .TextMatrix(0, 40) = "WharfageBasic"
        .ColWidth(40) = 1
        .TextMatrix(0, 41) = "New CCR Tag"
        .ColWidth(41) = 1
        .TextMatrix(0, 42) = "Wharfage Only"
        .ColWidth(42) = 1000
        .Redraw = True
    End With
End Sub
Private Sub FillGrid()
    ' ** Move Data per container in grid
    flexDetails.Visible = False
    If lngRow > 1 Then
        flexDetails.AddItem (" ")
    End If
    With flexDetails
        .Redraw = False
        .TextMatrix(lngRow, 2) = strCntNum
        .TextMatrix(lngRow, 9) = intSize
        .TextMatrix(lngRow, 10) = strFulemp
        .TextMatrix(lngRow, 11) = strDangr
        .TextMatrix(lngRow, 15) = sngOvzLength
        .TextMatrix(lngRow, 16) = sngOvzWidth
        .TextMatrix(lngRow, 17) = sngOvzHeight
        .TextMatrix(lngRow, 18) = strUms
        .TextMatrix(lngRow, 21) = strTshpCode
        .TextMatrix(lngRow, 22) = strExmCode 'strWhfCode
        .TextMatrix(lngRow, 34) = sngDollarAmt
        .TextMatrix(lngRow, 33) = lngPrevCCR
        .TextMatrix(lngRow, 37) = strPrevPay
        .TextMatrix(lngRow, 41) = lngTagNewCCR
        .Redraw = True
        .Visible = True
    End With
End Sub
Private Sub SaveToGrid()
    ' ** Update Data to Grid
    With flexDetails
        .Redraw = False
        .TextMatrix(.Row, 2) = strCntNum
        .TextMatrix(.Row, 9) = intSize
        .TextMatrix(.Row, 10) = strFulemp
        .TextMatrix(.Row, 11) = strDangr
        .TextMatrix(.Row, 15) = sngOvzLength
        .TextMatrix(.Row, 16) = sngOvzWidth
        .TextMatrix(.Row, 17) = sngOvzHeight
        .TextMatrix(.Row, 18) = strUms
        .TextMatrix(.Row, 21) = strTshpCode
        .TextMatrix(.Row, 34) = sngDollarAmt
        .Redraw = True
    End With
End Sub
Private Sub Tab00off(blnoff As Boolean)
    ' ** Set TabStop to details in TAB 0
    utxtPref.TabStop = blnoff
    utxtPref.Enabled = blnoff
    utxtNo.TabStop = blnoff
    utxtSze.TabStop = blnoff
    utxtSze.Enabled = blnoff
    utxtFEmp.TabStop = blnoff
    chkNewCCR.TabStop = False
    utxtNumDangr.TabStop = blnoff
    utxtUMS.TabStop = blnoff
    utxtTshipMnt.TabStop = blnoff
    utxtDollar.TabStop = blnoff
    utxtLength.TabStop = blnoff
    utxtWidth.TabStop = blnoff
    utxtHeight.TabStop = blnoff
    flexDetails.TabStop = blnoff
End Sub
Private Sub Tab01off(blnoff As Boolean)
    Dim intCtr As Integer
    ' ** Set TabStop to details in TAB 1
    utxtBroker.TabStop = blnoff
    utxtExporter.TabStop = blnoff
    utxtCommodity.TabStop = blnoff
    utxtRemark.TabStop = blnoff
    utxtVessel.TabStop = blnoff
    For intCtr = 0 To 10
        utxtEntry1(intCtr).TabStop = blnoff
    Next
    utxtVatCode.TabStop = blnoff
    utxtUGuarantee.TabStop = blnoff
End Sub
Private Sub Tab02off(blnoff As Boolean)
Dim X As Integer
' ** Set TabStop to details in TAB 2
        utxtCsh.TabStop = blnoff
        For X = 0 To 4
           utxtChq(X).TabStop = blnoff
           utxtChqNo(X).TabStop = blnoff
           utxtChqBnk(X).TabStop = blnoff
        Next
        utxtAdrAmt.TabStop = False
        utxtCustNo.TabStop = blnoff
        utxtCCRNo.TabStop = blnoff
        utxtCCRNo.Enabled = blnoff
End Sub
Private Sub utxtCCRNo_Change()
        blnChanging = True
        Call CheckAllocationRange
End Sub

Private Sub utxtCCRNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And IsNumeric(utxtCCRNo.Value) Then
        Call CheckAllocationRange
    End If
End Sub

Private Sub utxtChq_Change(Index As Integer)
   utxtChange.Caption = EvaluateChange
   utxtChange.Caption = Format(utxtChange.Caption, "###,###,##0.00")
End Sub

Private Sub utxtChq_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13
            If Index >= 0 And Index < 4 Then
                'If CCur(utxtChq(Index - 1).Value) = 0 Then
                If CCur("0" & utxtChq(Index).Value) = 0 Then
                    'utxtChq(4).SetFocus
                    utxtChqBnk(4).SetFocus
                End If
            End If
        Case 37 'Left
        Case 39 'Right
            utxtChqNo(Index).SetFocus
    End Select
End Sub

Private Sub utxtChqBnk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 37 'Left
            utxtChqNo(Index).SetFocus
        Case 39 'Right
    End Select
End Sub

Private Sub utxtChqNo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 37 'Left
            utxtChq(Index).SetFocus
        Case 39 'Right
            utxtChqBnk(Index).SetFocus
    End Select
End Sub

Private Sub utxtCustName_KeyPress(KeyAscii As Integer)
    Beep
    KeyAscii = 0
End Sub
Private Sub utxtCustNo_Change()
If Len(Trim(utxtCustNo.Value)) <> 0 Then
    DE.getADRBal Trim(utxtCustNo.Value), sngAdrBalance
    DE.getCustomerName Trim(utxtCustNo.Value), strCustName
    utxtCustName.Text = strCustName
    If sngAdrBalance <> 0 Then
        utxtAdrAmt.TabStop = True
        utxtAdrAmt.Enabled = True
        utxtAdrAmt.SetFocus
        cmdPrint.Enabled = True
    Else
        utxtAdrAmt.Value = ".00"
        utxtAdrAmt.TabStop = False
        'utxtAdrAmt.Enabled = False
        cmdPrint.Enabled = False
    End If
Else
    sngAdrBalance = 0
    utxtAdrAmt.Value = ".00"
    strCustName = ""
    utxtAdrAmt.TabStop = False
    'utxtAdrAmt.Enabled = False
    utxtCustName.Text = ""
    cmdPrint.Enabled = True
End If
If Len(utxtCustNo.Value) = 0 Then
    utxtCustName.Text = ""
End If
If chkUG = "Y" Then
    If Len(utxtCustNo.Value) > 0 Then
        cmdPrint.Enabled = True
    Else
        cmdPrint.Enabled = False
    End If
Else
        cmdPrint.Enabled = False
End If
End Sub

Private Sub utxtFEmp_KeyPress(KeyAscii As Integer)
If KeyAscii <> 27 And KeyAscii <> 13 And KeyAscii <> 8 Then
    If Not UCase(Chr(KeyAscii)) = "F" And Not UCase(Chr(KeyAscii)) = "E" Then
        Beep
        KeyAscii = 0
    Else
        If Len(Trim(utxtFEmp.Text)) = 1 Then
            Beep
            KeyAscii = 0
        End If
    End If
End If
End Sub

Private Sub utxtFEmp_LostFocus()
    If Len(Trim(utxtFEmp.Text)) = 0 Then
        utxtFEmp.Text = ""
    End If
End Sub

Private Sub utxtHeight_Change()
    If Len(Trim(utxtHeight.Value)) > 0 Then
        If Len(Trim(utxtUMS.Text)) = 0 Then
            utxtUMS.Text = defaultUnitMeasurement
        End If
    End If
End Sub

Private Sub utxtLength_Change()
    If Len(Trim(utxtLength.Value)) > 0 Then
        If Len(Trim(utxtUMS.Text)) = 0 Then
            utxtUMS.Text = defaultUnitMeasurement
        End If
    End If
End Sub

'Private Sub utxtSze_LostFocus()
   ' 'If Len(Trim(utxtSze.Value)) < 2 And Len(Trim(utxtSze.Value)) > 0 Then
   'If Trim(utxtPref.Text) <> "" Then
   ' If Trim(utxtSze.Value) <> "20" And Trim(utxtSze.Value) <> "40" And Trim(utxtSze.Value) <> "45" Then
   '   'If utxtSze.Enabled Then
   '     Beep
   '     MsgBox "Invalid Container Size, Please re-enter", vbExclamation + vbOKOnly, "Container Size Error"
   '     utxtSze.SelStart = 1
   '     utxtSze.SelLength = Len(Trim(utxtSze.Value))
   '     utxtSze.SetFocus
   '   'End If
   ' End If
   'Else
   '  utxtPref.SetFocus
   'End If
'End Sub

Private Sub utxtTshipMnt_Change()
'utxtWhfExmpt.Enabled = True
'utxtWhfExmpt.TabStop = True
'utxtWhfOnly.Enabled = True
'utxtWhfOnly.TabStop = True
If utxtTshipMnt.Text = "" Then
    utxtDollar.Visible = False
    Label7.Visible = False
    utxtDollar.Enabled = False
    utxtDollar.TabStop = False
    utxtDollar.Value = 0
Else
    Select Case utxtTshipMnt.Text
        Case Chr(vbKeyF)
            utxtDollar.Visible = True
            Label7.Visible = True
            utxtDollar.Enabled = True
            utxtDollar.TabStop = True
            utxtLength.Value = 0
            utxtWidth.Value = 0
            utxtHeight.Value = 0
            utxtUMS.Text = defaultUnitMeasurement
'            utxtWhfExmpt.Enabled = False
'            utxtWhfExmpt.TabStop = False
'            utxtWhfOnly.Enabled = False
'            utxtWhfOnly.TabStop = False
            utxtDollar.SetFocus
        Case Chr(vbKeyR)
            utxtDollar.Visible = False
            Label7.Visible = False
            utxtDollar.Enabled = False
            utxtDollar.TabStop = False
            
            utxtDollar.Value = 0
         Case Else
            Beep
            MsgBox "Transhipment Code Error", vbExclamation + vbOKOnly, "Transhipment Error"
            utxtTshipMnt.SelStart = 0
            utxtTshipMnt.SelLength = Len(Trim(utxtTshipMnt.Text))
            utxtTshipMnt.SetFocus
    End Select
End If
End Sub

Private Sub utxtUGuarantee_LostFocus()
If Trim(utxtUGuarantee.Text) = "Y" Then
  FrameCust.Visible = True
  utxtCustNo1.SetFocus
Else
  If Len(utxtUGuarantee.Text) = 0 Then
      utxtUGuarantee.Text = "N"
  End If
  utxtCustNo1.Value = ""
  utxtCustName1.Text = ""
  FrameCust.Visible = False
  utxtEntry1(0).SetFocus
End If
End Sub

Private Sub utxtUMS_KeyPress(KeyAscii As Integer)
' ** Trapping Invalid Characters
    If UCase(Chr(KeyAscii)) <> "C" And UCase(Chr(KeyAscii)) <> "I" And KeyAscii <> 8 Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub utxtVatCode_KeyPress(KeyAscii As Integer)
' ** Trapping Invalid Characters
If KeyAscii <> 8 And KeyAscii <> 27 And KeyAscii <> 13 Then
    If UCase(Chr(KeyAscii)) <> "1" And _
        UCase(Chr(KeyAscii)) <> "2" And _
        UCase(Chr(KeyAscii)) <> "3" And _
        UCase(Chr(KeyAscii)) <> "4" And _
        UCase(Chr(KeyAscii)) <> "5" And _
        UCase(Chr(KeyAscii)) <> "6" And _
        KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
    Else
        utxtVatCode.Value = UCase(Chr(KeyAscii))
        KeyAscii = 0
    End If
End If
If KeyAscii = 13 Then
    If utxtTshipMnt.Text = "F" Then
       utxtUGuarantee.SetFocus
    End If
End If
End Sub

'Private Sub utxtSze_Change()
'    If Len(Trim(utxtSze.Value)) = 2 Then
'        If Trim(utxtSze.Value) = "20" Or Trim(utxtSze.Value) = "40" Or Trim(utxtSze.Value) = "45" Then
'            Exit Sub
'        Else
'            Beep
'            MsgBox "Invalid Container Size", vbExclamation + vbOKOnly, "Container Size Error"
'            utxtSze.SelStart = 0
'            utxtSze.SelLength = Len(Trim(utxtSze.Value))
'            utxtSze.SetFocus
'        End If
'    End If
'End Sub
Private Sub utxtSze_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And KeyAscii <> 13 Then
        If Len(Trim(utxtSze.Value)) = 2 And utxtSze.SelLength <> 2 Then
            Beep
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub utxtTshipMnt_GotFocus()
    flexTshipMnt.Visible = True
    flexTshipMnt.Height = 1500
    flexTshipMnt.Width = 5510
    flexTshipMnt.ColWidth(0) = 5500
End Sub
Private Sub utxtTshipMnt_LostFocus()
    flexTshipMnt.Visible = False
End Sub

'PRNH-OLD
'Private Sub utxtAdrAmt_Change()
'    Dim strAdrMsg1 As String
'    Dim strAdrMsg2 As String
'    Dim intResponse As Integer
'    strAdrMsg1 = ""
'    strAdrMsg2 = ""
'    intResponse = 0
'    If blnnumberin Then
'        If IsNumeric(utxtAdrAmt.Value) Then
'            If CCur(utxtAdrAmt.Value) > sngAdrBalance Then
'                utxtAdrAmt.Value = Trim(utxtAdrAmt.Value)
'                utxtCustName.Text = Trim(utxtCustName.Text)
'                strAdrMsg1 = utxtAdrAmt.Value & " is greater than the Current ADR Running Balance for Customer "
'                strAdrMsg2 = utxtCustName.Text
'                intResponse = MsgBox(strAdrMsg1 & strAdrMsg2, vbOKOnly, "ADR Amount Error Message")
'                utxtAdrAmt.Value = sngPreviousADR
'                utxtAdrAmt.SelStart = 0
'                utxtAdrAmt.SelLength = utxtAdrAmt.Value
'                utxtAdrAmt.SetFocus
'            Else
'                If CCur(utxtAdrAmt.Value) > CCur(utxtAmtPay.Caption) Then
'                    utxtAdrAmt.Value = Trim(utxtAdrAmt.Value)
'                    utxtCustName.Text = Trim(utxtCustName.Text)
'                    strAdrMsg1 = utxtAdrAmt.Value & " is greater than the Total Amount to be paid "
'                    strAdrMsg2 = utxtCustName.Text
'                    intResponse = MsgBox(strAdrMsg1 & strAdrMsg2, vbOKOnly, "ADR Amount Error Message")
'                    utxtAdrAmt.Value = sngPreviousADR
'                    utxtAdrAmt.SelStart = 0
'                    utxtAdrAmt.SelLength = utxtAdrAmt.Value
'                    utxtAdrAmt.SetFocus
'                Else
'                    cmdPrint.Enabled = True
'                End If
'            End If
'            blnnumberin = False
'        End If
'    End If
'    utxtChange.Caption = EvaluateChange
'    utxtChange.Caption = Format(utxtChange.Caption, "###,###,##0.00")
'End Sub

'PRNH - Removed validations
Private Sub utxtAdrAmt_Change()
    Dim strAdrMsg1 As String
    Dim strAdrMsg2 As String
    Dim intResponse As Integer
    strAdrMsg1 = ""
    strAdrMsg2 = ""
    intResponse = 0
    If blnnumberin Then
        If IsNumeric(utxtAdrAmt.Value) Then
'            If CCur(utxtAdrAmt.Value) > sngAdrBalance Then
'                utxtAdrAmt.Value = Trim(utxtAdrAmt.Value)
'                utxtCustName.Text = Trim(utxtCustName.Text)
'                strAdrMsg1 = utxtAdrAmt.Value & " is greater than the Current ADR Running Balance for Customer "
'                strAdrMsg2 = utxtCustName.Text
'                intResponse = MsgBox(strAdrMsg1 & strAdrMsg2, vbOKOnly, "ADR Amount Error Message")
'                utxtAdrAmt.Value = sngPreviousADR
'                utxtAdrAmt.SelStart = 0
'                utxtAdrAmt.SelLength = utxtAdrAmt.Value
'                utxtAdrAmt.SetFocus
'            Else
                If CCur(utxtAdrAmt.Value) > CCur(utxtAmtPay.Caption) Then
                    utxtAdrAmt.Value = Trim(utxtAdrAmt.Value)
                    utxtCustName.Text = Trim(utxtCustName.Text)
                    strAdrMsg1 = utxtAdrAmt.Value & " is greater than the Total Amount to be paid "
                    strAdrMsg2 = utxtCustName.Text
                    intResponse = MsgBox(strAdrMsg1 & strAdrMsg2, vbOKOnly, "ADR Amount Error Message")
                    utxtAdrAmt.Value = sngPreviousADR
                    utxtAdrAmt.SelStart = 0
                    utxtAdrAmt.SelLength = utxtAdrAmt.Value
                    utxtAdrAmt.SetFocus
                Else
                    cmdPrint.Enabled = True
                End If
            'End If
            blnnumberin = False
        End If
    End If
    utxtChange.Caption = EvaluateChange
    utxtChange.Caption = Format(utxtChange.Caption, "###,###,##0.00")
End Sub

'PRNH-ORIG
'Private Sub utxtAdrAmt_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 39 Then
'        utxtCustNo.SetFocus
'    Else
'        If Asc(KeyCode) > 47 And Asc(KeyCode) < 58 And IsNumeric(utxtAdrAmt.Value) Then
'            sngPreviousADR = utxtAdrAmt.Value
'            blnnumberin = True
'        Else
'            sngPreviousADR = 0
'            blnnumberin = False
'            utxtAdrAmt.Value = ".00"
'        End If
'    End If
'End Sub

'PRNH - Removed validations in ADR

Private Sub utxtAdrAmt_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 39 Then
'        utxtCustNo.SetFocus
'    Else
'        If Asc(KeyCode) > 47 And Asc(KeyCode) < 58 And IsNumeric(utxtAdrAmt.Value) Then
'            'sngPreviousADR = utxtAdrAmt.Value
'            blnnumberin = True
'        Else
'            sngPreviousADR = 0
'            blnnumberin = False
'            utxtAdrAmt.Value = ".00"
'        End If
'    End If
End Sub
Private Sub utxtCsh_Change()
    utxtChange.Caption = EvaluateChange
    utxtChange.Caption = Format(utxtChange.Caption, "###,###,##0.00")
End Sub
Private Sub utxtCustName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        If cusListing.ShowList Then
            utxtCustNo.Value = cusListing.Code
            DE.getADRBal Trim(utxtCustNo.Value), sngAdrBalance
            DE.getCustomerName Trim(utxtCustNo.Value), strCustName
            utxtCustName.Text = strCustName
            If sngAdrBalance <> 0 Then
                utxtAdrAmt.TabStop = True
                utxtAdrAmt.Enabled = True
            Else
                utxtAdrAmt.TabStop = False
                'utxtAdrAmt.Enabled = False
            End If
        End If
    End If
End Sub

Private Sub utxtCustNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF4
            If cusListing.ShowList Then
                utxtCustNo.Value = cusListing.Code
                DE.getADRBal Trim(cusListing.Code), sngAdrBalance
                utxtCustName.Text = Trim(cusListing.Name)
                If sngAdrBalance <> 0 Then
                    utxtAdrAmt.TabStop = True
                    utxtAdrAmt.Enabled = True
                Else
                    utxtAdrAmt.TabStop = False
                    'utxtAdrAmt.Enabled = False
                End If
            End If
        Case 37 'Left
            If sngAdrBalance > 0 Then
                utxtAdrAmt.SetFocus
            End If
        Case vbKeyReturn
            If Len(utxtCustNo.Value) > 0 Then
               cmdPrint.Enabled = True
            End If
    End Select
End Sub

Private Sub utxtCustNo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
            If cusListing.ShowList Then
                utxtCustNo1.Value = cusListing.Code
                utxtCustName1.Text = Trim(cusListing.Name)
            End If
    End If
End Sub

Private Sub utxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strSze As String * 2
    Dim strFEmp As String * 1
    Dim strccrNodte As String
    Dim strTeller As String * 10
    Dim strContNum As String * 12
    Dim strPrefix As String * 4
    Dim strNo As String * 8
    Dim strExprtr As String
    Dim strCont1 As String
    Dim strCont2 As String
    Dim strCont3 As String
    lngPrevCCR = 0
    strSze = ""
    strContNum = ""
    strTeller = ""
    strFEmp = ""
    strExprtr = ""
    strPrefix = ""
    strNo = ""
    strCont1 = ""
    strCont2 = ""
    strCont3 = ""
    If KeyCode = vbKeyReturn Then
        If (utxtPref.Text) = "" And (utxtNo.Text = "") Then
            utxtPref.SetFocus
            Exit Sub
        End If
        strResponse = True
        LSet strPrefix = (utxtPref.Text)
        If Mid(utxtNo.Text, 1, 1) = " " Then
            LSet strNo = utxtNo.Text
            strCntNum = strPrefix & strNo
            strContNum = strCntNum
            strCont1 = "Please Check Inputted Container"
            strCont2 = strContNum
            strCont3 = "Before Verifying!"
            strResponse = True
            Call SystemMessage(strCont1, strCont2, strCont3)
        End If
        If strResponse = True Then
            strPrefix = ""
            strNo = ""
            strCntNum = ""
            lblStatus.Caption = "Verifying Container Number, Please Wait ......."
            lblStatus.Refresh
            LSet strPrefix = (utxtPref.Text)
            LSet strNo = utxtNo.Text
            strCntNum = strPrefix & strNo
            strContNum = strCntNum
            If ValidateCntNum(strSze, strFEmp, lngPrevCCR, strccrNodte, strTeller, strResponse, strPrevPay, strContNum, strExprtr) = True Then
            '=========Navis===============
             ConnectToNavis
'            'GET Last Discharge
'            strGKey = GetGKey("APHU6623529", "QUEUED", "STORAGE")
'            Call Sparcs_LastDisch("APHU6623529", "STORAGE", "", strGKey)
'            mskLastDischargeDate.Text = Format(dStorage, "yyyy-mm-dd")
'            strDate = Format(DateAdd("d", 7, CDate(dStorage)), "yyyy-mm-dd") 'CRODate = Last Discharge + 7 days by default
'            mskCRODate.Text = strDate
'
'            'GET Reefer
'            strGKey = ""
'            strGKey = GetGKey("YMLU5305095", "QUEUED", "REEFER")
'            Call Sparcs_LastDisch("YMLU5305095", "REEFER", "", strGKey)
'            Me.mskPlugIN.Text = Format(dReefer, "yyyy-mm-dd hh:mm:ss")
            
            'Get the container size
            Sparcs_ExpSize (strCntNum)
            
            'Get the freight kind
            Sparcs_FreightKind (strCntNum)
                
            'Get DGCode
            Sparcs_DGCode (strCntNum)
            
            'Get the transshipment code
            'What is the transhipment code equivalent in N4?
            
            'GetOOG
            Sparcs_OOG (strCntNum)
            'added by Navis Project Team 10/27/2009
            
            'PRNH
            CheckContainerInNAVIS (Trim(strContNum))
            
            If lngRow > 0 Then
                If lblCompCode.Caption <> "" Then
                    If flexDetails.TextMatrix(1, 45) <> lblCompCode.Caption Then
                        MsgBox "Company code must be the same with the first container."
ResetCont:
                        Call InitializeDetail
                        Call EnableForNewContainer(True)
                        lblStatus.Caption = " "
                        utxtPref.SetFocus
                        Exit Sub
                    End If
                Else
                    GoTo ResetCont
                End If
            End If
            
            
            
            If utxtSze.Value <> "" Then
                If CInt(utxtHeight.Value) > 0 Or CInt(utxtLength.Value) > 0 Or CInt(utxtWidth.Value > 0) Then
                    If utxtSze.Value = "40" Then
                        If CInt(utxtLength.Value) > 0 Then
                            utxtLength.Value = CInt(utxtLength.Value) / 2.54
                        End If
                        utxtLength.Value = CDbl(utxtLength.Value) + 480
                    ElseIf utxtSze.Value = "20" Then
                        If CInt(utxtLength.Value) > 0 Then
                            utxtLength.Value = CInt(utxtLength.Value) / 2.54
                        End If
                        utxtLength.Value = CDbl(utxtLength.Value) + 240
                    End If
                    If CInt(utxtHeight.Value) > 0 Then
                        utxtHeight.Value = CInt(utxtHeight.Value) / 2.54
                    End If
                    utxtHeight.Value = CDbl(utxtHeight.Value) + 102
                    
                    If CInt(utxtWidth.Value) > 0 Then
                        utxtWidth.Value = CInt(utxtWidth.Value) / 2.54
                    End If
                    utxtWidth.Value = CDbl(utxtWidth.Value) + 96
                End If
            End If

            
            'Get Exporter Name
            Sparcs_Shipper (strCntNum)
            
            'Get Commodity Code
            Sparcs_Commodity (strCntNum)
            
            'Get Vessel Name
            Sparcs_VesselName (strCntNum)
            
            '=============================
                    Call EnableForNewContainer(False)
                    ' ** check case 1
                    If Len(Trim(strSze)) <> 0 Then  'And Len(strFEmp) <> 0 Then
                        utxtSze.Value = Trim(strSze)
                        'utxtFEmp.Text = Trim(strFEmp)
                    Else
                        utxtSze.SetFocus
                    End If
                    ' ** check case 2
                    If strResponse = True Then
                        utxtSze.SetFocus
                        'utxtSze.Value = Trim(strSze)
                        'utxtFEmp.Text = Trim(strFEmp)
                    End If
                    'utxtUMS.Text = defaultUnitMeasurement
            Else
                    If strResponse = False Then
                        utxtPref.SetFocus
                    End If
            End If
        Else
            utxtPref.SetFocus
        End If
    End If
    lblStatus.Caption = " "
End Sub
Private Sub MoveHeaderToVariables()
    ' ** Move Header Details to Variables
    strEntry = ""
    strVessel = ""
    strExporter = ""
    strBroker = ""
    strCommodity = ""
    strSBMAPermit = ""
    strExmCode = ""
    strRemark = ""
    Call ClearEntry
    If Len(utxtEntry1(0).Value) <> 0 Then
        RSet strEntry1 = utxtEntry1(0).Value
        strEntry = (strEntry) & strEntry1
    End If
    If Len(utxtEntry1(1).Value) <> 0 Then
        RSet strEntry2 = utxtEntry1(1).Value
        strEntry = (strEntry) & strEntry2
    End If
    If Len(utxtEntry1(2).Value) <> 0 Then
        RSet strEntry3 = utxtEntry1(2).Value
        strEntry = (strEntry) & strEntry3
    End If
    If Len(utxtEntry1(3).Value) <> 0 Then
        RSet strEntry4 = utxtEntry1(3).Value
        strEntry = (strEntry) & strEntry4
    End If
    If Len(utxtEntry1(4).Value) <> 0 Then
        RSet strEntry5 = utxtEntry1(4).Value
        strEntry = (strEntry) & strEntry5
    End If
    If Len(utxtEntry1(5).Value) <> 0 Then
        RSet strEntry6 = utxtEntry1(5).Value
        strEntry = (strEntry) & strEntry6
    End If
    If Len(utxtEntry1(6).Value) <> 0 Then
        RSet strEntry7 = utxtEntry1(6).Value
        strEntry = (strEntry) & strEntry7
    End If
    If Len(utxtEntry1(7).Value) <> 0 Then
        RSet strEntry8 = utxtEntry1(7).Value
        strEntry = (strEntry) & strEntry8
    End If
    If Len(utxtEntry1(8).Value) <> 0 Then
        RSet strEntry9 = utxtEntry1(8).Value
        strEntry = (strEntry) & strEntry9
    End If
    If Len(utxtEntry1(9).Value) <> 0 Then
        RSet strEntry10 = utxtEntry1(9).Value
        strEntry = (strEntry) & strEntry10
    End If
    If Len(utxtEntry1(10).Value) <> 0 Then
        RSet strEntry0 = utxtEntry1(10).Value
        strEntry = (strEntry) & strEntry0
    End If
    
    strExporter = Trim(utxtExporter.Text)
    strBroker = Trim(utxtBroker.Text)
    strCommodity = Trim(utxtCommodity.Text)
    strSBMAPermit = Trim(utxtSBMAPermit.Text)
    strVessel = Trim(utxtVessel.Text)
    strEntry = Trim(strEntry)
    strRemark = Trim(utxtRemark.Text)
    Select Case utxtUGuarantee.Text
        Case "Y"
            strUg = "Y"
            chkUG = "Y"
        Case "N"
            strUg = "N"
    End Select
End Sub
Private Sub ComputeCharges(lngTempRow, ByVal TempLength As Single, ByVal TempWidth As Single, ByVal TempHeight As Single)
    ' ** Compute Charges for Specific Container Number
    strWhfCode = "N"
    sngArr = Comp_Arrastre(sngdangramt, sngBscArr, sngRton, sngOvzAmt, _
                   intSize, strDangr, TempLength, TempWidth, TempHeight, strUms)
    If chkWeighing.Value = 1 Then
        sngWeighing = ComputeWeighing()
    Else
        sngWeighing = 0
    End If
    'Select Case Trim(flexDetails.TextMatrix(lngTempRow, 21))
    '    Case "F"
    '        Call ComputeTranshipmentWharfage(lngTempRow)
    '    Case Else
    '        Call ComputeRegular_Wharfage
    'End Select
    Call Compute_Vat
End Sub

Private Sub ComputeRegular_Wharfage()
    ' ** Wharfage Computation if not exempted
    sngWhf = 0
    strWhfOnly = "N"
    sngWhf = GetRate("EXWF", intSize)
    sngBscWhf = GetRate("EXWF", intSize)
    strWhfCode = "N"
End Sub

Private Sub ComputeTranshipmentWharfage(lngTempRow)
' ** Wharfage Computation for Transhipment Cargo
Dim sngWhfT As Currency
    
    sngWhfT = 0
    sngWhf = 0
    sngArr = 0
    sngBscArr = 0
    sngBscWhf = 0
    
    sngWhfT = GetRate("EXWT", intSize)
    Select Case intSize
        Case 20
'            sngWhf = sngDollarAmt * sngWhfT
            sngWhf = Trim(flexDetails.TextMatrix(lngTempRow, 34)) * sngWhfT
            sngBscWhf = sngWhf
        Case 40
            sngWhf = Trim(flexDetails.TextMatrix(lngTempRow, 34)) * sngWhfT
        Case 45
            sngWhf = Trim(flexDetails.TextMatrix(lngTempRow, 34)) * sngWhfT
    End Select
    sngBscWhf = sngWhf
End Sub
Private Sub Compute_Vat()
' ** Computes for specific Vat
    sngVat = 0
    sngWtax = 0
    strVatCode = ""
    Select Case Val(utxtVatCode.Value)
        Case 1
            sngVat = 0
            strVatCode = "1"
        Case 2
            sngVat = 0
            sngWtax = sngArr * 0.01
            strVatCode = "2"
        Case 3
            sngVat = 0
            sngWtax = sngArr * 0.02
            strVatCode = "3"
            'sngVat = sngArr * 0.1
            'strVatCode = "3"
        Case 4
            sngVat = sngArr * 0.1
            sngWtax = sngArr * 0.01
            strVatCode = "4"
        Case 5
            sngVat = sngArr * 0.06
            strVatCode = "5"
        Case 6
            sngVat = sngArr * 0.06
            sngWtax = sngArr * 0.01
            strVatCode = "6"
    End Select
End Sub
Private Sub ComputeTotal()
Dim lngCtrRow As Long
lngCtrRow = 0
    Do While Not (lngCtrRow = lngRow)
        lngCtrRow = lngCtrRow + 1
            If Trim(flexDetails.TextMatrix(lngCtrRow, 7)) = "Y" Then
                sngIctArr = 0
                sngIctVat = 0
                sngIctWtax = 0
            Else
                sngIctArr = sngIctArr + Trim(flexDetails.TextMatrix(lngCtrRow, 3))
                sngIctVat = sngIctVat + Trim(flexDetails.TextMatrix(lngCtrRow, 4))
                sngIctWtax = sngIctWtax + Trim(flexDetails.TextMatrix(lngCtrRow, 5))
                sngIctWgh = sngIctWgh + Trim(flexDetails.TextMatrix(lngCtrRow, 6))
            End If
            If Trim(flexDetails.TextMatrix(lngCtrRow, 7)) = "Y" Then
                sngIctTot = sngIctTot + 0
             Else
                sngIctTot = sngIctArr + sngIctWgh + sngIctVat - sngIctWtax
            End If
            sngGrndTot = sngIctTot
    Loop
    txtIctsiDue.Caption = sngIctTot
    txtTotDue.Caption = sngGrndTot
    Call ReformatDisplayTotal
End Sub
Private Sub InitializeDetail()
    utxtPref.Text = ""
    utxtNo.Text = ""
    utxtSze.Value = ""
    utxtFEmp.Text = ""
    utxtNumDangr.Value = ""
    utxtTshipMnt.Text = ""
    utxtLength.Value = 0
    utxtLength.Value = CSng(Format(utxtLength.Value, "###,###.00"))
    utxtWidth.Value = 0
    utxtWidth.Value = CSng(Format(utxtWidth.Value, "###,###.00"))
    utxtHeight.Value = 0
    utxtHeight.Value = CSng(Format(utxtHeight.Value, "###,###.00"))
    utxtUMS.Text = defaultUnitMeasurement
    
    'PRNH
    lblCompCode.Caption = ""
End Sub
Private Sub ClearVariables()
    sngArr = 0
    sngBscArr = 0
    sngWhf = 0
    sngBscWhf = 0
    sngOvzAmt = 0
    sngVat = 0
    sngWtax = 0
    sngOvzLength = 0
    sngOvzWidth = 0
    sngOvzHeight = 0
    sngdangramt = 0
    sngDollarAmt = 0
    sngRton = 0
    intSize = 0
    strEntry = ""
    strRemark = ""
    strDangr = ""
    strCntNum = ""
    strFulemp = ""
    strVessel = ""
    strVatCode = ""
    strUms = ""
    strTshpCode = ""
    strExmCode = ""
    strWhfOnly = ""
    strExporter = ""
    strBroker = ""
    strCommodity = ""
    strSBMAPermit = ""
End Sub
Private Sub ClearTotals()
    sngIctTot = 0
    sngPpa = 0
    sngGrndTot = 0
    sngIctArr = 0
    sngIctVat = 0
    sngIctWtax = 0
    sngIctWgh = 0
End Sub

Public Function zGetSysDate() As Date
Dim rsDate As Recordset
DE.GetDate
' ** Return Recordset
Set rsDate = DE.rsGetDate
    With rsDate
        zGetSysDate = .Fields(0)
        .Close
    End With
Set rsDate = Nothing
End Function

Private Function ValidateCntNum(ByRef strSze As String, ByRef strFEmp As String, _
            ByRef lngPrevCCR As Long, ByRef strccrNodte As String, ByRef strTeller As String, _
            ByRef strResponse As Boolean, ByRef strPrevPay As String, _
            strContNum As String, ByRef strExprtr As String) As Boolean

Dim strValid1 As String
Dim strValid2 As String
Dim strValid3 As String
Dim lng7thDigit As Long
Dim str6thDigit As String
Dim lngTmp7thDigit As Long
Dim strCntNo As String

strValid1 = ""
strValid2 = ""
strValid3 = ""
ValidateCntNum = False
strPrevPay = ""
lng7thDigit = 0
str6thDigit = ""
strCntNo = ""
lngTmp7thDigit = 0
strSze = ""
strFEmp = "F"


'  ** Check Container Number if existing in grid
    If Not Container_Valid Then
        ValidateCntNum = False
        Exit Function
    End If
' ** Check for paid container
    If GetPrevCCRPay(strContNum, lngPrevCCR, lngOvrCCr, strccrNodte, strTeller, strExprtr) = True Then
        ValidateCntNum = True
        strPrevPay = "P"
        strValid1 = "Container No. " & strContNum & " is billed by " & strTeller & _
                                    " on " & strccrNodte
        strValid2 = "with CCR No. " & lngPrevCCR & " Exporter : " & strExprtr
        strValid3 = "DO YOU WANT TO OVERWITE PAYMENT ?"
        If MsgBox(strValid1 & vbCrLf & strValid2 & vbCrLf & vbCrLf & strValid3, vbDefaultButton2 + vbYesNo + vbQuestion) = vbYes Then
            strResponse = True
        Else
            strResponse = False
        End If
'         Call SystemMessage(strValid1, strValid2, strValid3)
        If Not strResponse Then
            ValidateCntNum = False
            strPrevPay = ""
            Exit Function
        Else
            ValidateCntNum = True
            strPrevPay = "P"
        End If
    Else
        ValidateCntNum = True
    End If

'***************************************
'* Get container Information
'* FLR StandAlone
'***************************************
'        With CTCSinfo
'            .GetLatestMove strContNum
'            If .ContSize <> 0 Then
'                strSze = .ContSize
'                If .ContFull = True Then
'                    strFEmp = "F"
'                End If
'                ValidateCntNum = True
'                Exit Function
'            Else
'                strValid1 = "Container Not In Yard "
'                strValid2 = "Please VERIFY with SHIPPING LINE"
'                strResponse = False
'                Call ErrorMessage(strValid1, strValid2)
'                ValidateCntNum = True
'            End If
'        End With
'***************************************

End Function
Private Sub UpdateGrid()
' ** Update Details from Grid
    Dim lngTempRow As Long
    lngTempRow = 0
    Do While Not (lngTempRow = lngRow)
            lngTempRow = lngTempRow + 1
        If lngTempRow = lngRow Then
            With flexDetails
                If Len(Trim(.TextMatrix(lngTempRow, 39))) = 0 Then
                    intSize = Trim(flexDetails.TextMatrix(lngTempRow, 9))
                    strDangr = Trim(flexDetails.TextMatrix(lngTempRow, 11))
                    sngOvzLength = Trim(flexDetails.TextMatrix(lngTempRow, 15))
                    sngOvzWidth = Trim(flexDetails.TextMatrix(lngTempRow, 16))
                    sngOvzHeight = Trim(flexDetails.TextMatrix(lngTempRow, 17))
                    strUms = Trim(flexDetails.TextMatrix(lngTempRow, 18))
                    'Call ConvertSize
                    Call MoveHeaderToVariables
                    Call ComputeCharges(lngTempRow, sngOvzLength, sngOvzWidth, sngOvzHeight)
                    Call MoveToGrid(lngTempRow)
                End If
            End With
        End If
    Loop
    Call ComputeTotal
    Call ClearVariables
End Sub
Private Sub UpdatePerDet(lngTempRow As Long)
' ** Update Details from Grid
    With flexDetails
        If Len(Trim(.TextMatrix(lngTempRow, 39))) = 0 Then
            intSize = Trim(flexDetails.TextMatrix(lngTempRow, 9))
            strDangr = Trim(flexDetails.TextMatrix(lngTempRow, 11))
            sngOvzLength = Trim(flexDetails.TextMatrix(lngTempRow, 15))
            sngOvzWidth = Trim(flexDetails.TextMatrix(lngTempRow, 16))
            sngOvzHeight = Trim(flexDetails.TextMatrix(lngTempRow, 17))
            strUms = Trim(flexDetails.TextMatrix(lngTempRow, 18))
            'Call ConvertSize
            Call MoveHeaderToVariables
            Call ComputeCharges(lngTempRow, sngOvzLength, sngOvzWidth, sngOvzHeight)
            Call MoveToGrid(lngTempRow)
        End If
    End With
    Call ComputeTotal
    Call ClearVariables
End Sub

Private Sub UpdateHeaderGrid()
' ** Update Details from Grid
    Dim lngTempRow As Long
    lngTempRow = 0
    Do While Not (lngTempRow = lngRow)
        lngTempRow = lngTempRow + 1
        With flexDetails
            If Trim(.TextMatrix(lngTempRow, 41)) = "" Then '<> lngTagNewCCR Then
                .TextMatrix(lngTempRow, 41) = lngTagNewCCR
                intSize = Trim(.TextMatrix(lngTempRow, 9))
                strDangr = Trim(.TextMatrix(lngTempRow, 11))
                sngOvzLength = Trim(.TextMatrix(lngTempRow, 15))
                sngOvzWidth = Trim(.TextMatrix(lngTempRow, 16))
                sngOvzHeight = Trim(.TextMatrix(lngTempRow, 17))
                strUms = Trim(.TextMatrix(lngTempRow, 18))
                
                Call MoveHeaderToVariables
                Call ComputeCharges(lngTempRow, sngOvzLength, sngOvzWidth, sngOvzHeight)
                'Added by Navis Project Team 10/26/2009
                Call MoveHeaderToGrid(lngTempRow)
'                Call MoveToGrid(lngTempRow)
            End If
        End With
    Loop
    Call ComputeTotal
    Call ClearVariables
End Sub

'Added by Navis Project Team 10/26/2009
Private Sub MoveHeaderToGrid(lngTempRow As Long)
    flexDetails.Visible = False
    With flexDetails
        .Redraw = False
        
        .TextMatrix(lngTempRow, 23) = strExporter
        .TextMatrix(lngTempRow, 24) = strBroker
        .TextMatrix(lngTempRow, 26) = strCommodity
        .TextMatrix(lngTempRow, 38) = strRemark
        .TextMatrix(lngTempRow, 13) = strVessel
        .TextMatrix(lngTempRow, 14) = strVatCode
        .TextMatrix(lngTempRow, 7) = strUg
        .TextMatrix(lngTempRow, 25) = strEntry
        
        .Visible = True
        .Redraw = True

    End With
End Sub

Private Sub MoveToGrid(lngTempRow As Long)
' ** Move Details from Variables to Grid
    flexDetails.Visible = False
    With flexDetails
        .Redraw = False
        .TextMatrix(lngTempRow, 3) = sngArr
        .TextMatrix(lngTempRow, 3) = CCur(.TextMatrix(lngTempRow, 3))
        .TextMatrix(lngTempRow, 3) = Format(.TextMatrix(lngTempRow, 3), "#,###,###.#0")
        .TextMatrix(lngTempRow, 4) = sngVat
        .TextMatrix(lngTempRow, 4) = CCur(.TextMatrix(lngTempRow, 4))
        .TextMatrix(lngTempRow, 4) = Format(.TextMatrix(lngTempRow, 4), "#,###,###.#0")
        .TextMatrix(lngTempRow, 5) = sngWtax
        .TextMatrix(lngTempRow, 5) = CCur(.TextMatrix(lngTempRow, 5))
        .TextMatrix(lngTempRow, 5) = Format(.TextMatrix(lngTempRow, 5), "#,###,###.#0")
' remove by bien
        .TextMatrix(lngTempRow, 6) = sngWeighing
        .TextMatrix(lngTempRow, 6) = CCur(.TextMatrix(lngTempRow, 6))
        .TextMatrix(lngTempRow, 6) = Format(.TextMatrix(lngTempRow, 6), "#,###,###.#0")
        .TextMatrix(lngTempRow, 7) = strUg
        .TextMatrix(lngTempRow, 8) = sngBscArr
        .TextMatrix(lngTempRow, 8) = Format(.TextMatrix(lngTempRow, 8), "#,###,###.#0")
        .TextMatrix(lngTempRow, 12) = sngdangramt
        .TextMatrix(lngTempRow, 12) = Format(.TextMatrix(lngTempRow, 12), "#,###,###.#0")
        .TextMatrix(lngTempRow, 13) = strVessel
        .TextMatrix(lngTempRow, 14) = strVatCode
        .TextMatrix(lngTempRow, 15) = sngOvzLength
        .TextMatrix(lngTempRow, 16) = sngOvzWidth
        .TextMatrix(lngTempRow, 17) = sngOvzHeight
        .TextMatrix(lngTempRow, 18) = strUms
        .TextMatrix(lngTempRow, 19) = sngOvzAmt
        .TextMatrix(lngTempRow, 20) = sngRton
        .TextMatrix(lngTempRow, 22) = strExmCode
        .TextMatrix(lngTempRow, 23) = strExporter
        .TextMatrix(lngTempRow, 24) = strBroker
        .TextMatrix(lngTempRow, 25) = strEntry
        .TextMatrix(lngTempRow, 26) = strCommodity
        .TextMatrix(lngTempRow, 43) = strSBMAPermit
        .TextMatrix(lngTempRow, 38) = strRemark
        .TextMatrix(lngTempRow, 40) = sngBscWhf
        If strPrevPay = "Y" Then
            .TextMatrix(lngTempRow, 42) = strWhfOnly
        Else
            .TextMatrix(lngTempRow, 42) = "N"
        End If
        
        'PRNH - Company Code
        .TextMatrix(lngTempRow, 45) = strCompCode
        
        .Visible = True
        .Redraw = True

    End With
End Sub

Private Sub ReformatDisplayTotal()
    txtIctsiDue.Caption = Format(txtIctsiDue.Caption, "#,###,##0.00")
    txtPpaTotal.Caption = Format(txtPpaTotal.Caption, "#,###,##0.00")
    txtTotDue.Caption = Format(txtTotDue.Caption, "#,###,##0.00")
    txtIctsiDue.Caption = Format(txtIctsiDue.Caption, "@@@@@@@@@@@@@")
    txtPpaTotal.Caption = Format(txtPpaTotal.Caption, "@@@@@@@@@@@@@")
    txtTotDue.Caption = Format(txtTotDue.Caption, "@@@@@@@@@@@@@")
End Sub

Private Sub MoveDetailToVariables()
' ** Move Details from Form to Variables
    
    intSize = utxtSze.Value
    strFulemp = utxtFEmp.Text
    If utxtNumDangr.Value <> "" Then
        strDangr = utxtNumDangr.Value
    End If
    If Len(utxtLength.Value) <> 0 And Len(utxtWidth.Value) <> 0 And Len(utxtHeight.Value) <> 0 Then
        sngOvzLength = utxtLength.Value
        sngOvzWidth = utxtWidth.Value
        sngOvzHeight = utxtHeight.Value
        strUms = Trim(utxtUMS.Text)
    Else
        sngOvzLength = 0
        sngOvzWidth = 0
        sngOvzHeight = 0
        strUms = ""
    End If
    strTshpCode = utxtTshipMnt.Text
    If utxtTshipMnt.Text = "F" Then
        sngOvzLength = 0
        sngOvzWidth = 0
        sngOvzHeight = 0
        strUms = "I"
    End If
    sngDollarAmt = ReturnCurrency(utxtDollar.Value)
    
    'PRNH - Company Code
    strCompCode = lblCompCode.Caption
End Sub

Private Sub EnableTotalVisible(blnVisible As Boolean)
    Label19.Visible = blnVisible
    Label20.Visible = blnVisible
    Label21.Visible = blnVisible
    txtIctsiDue.Visible = blnVisible
    txtPpaTotal.Visible = blnVisible
    txtTotDue.Visible = blnVisible
End Sub

Private Sub MoveToText()
    Dim strNum As String * 12
    strNum = ""
    utxtPref.Text = ""
    utxtNo.Text = ""
    With flexDetails
        strNum = (.TextMatrix(.Row, 2))
        utxtPref.Text = Mid(strNum, 1, 4)
        utxtNo.Text = Mid(strNum, 5, 8)
        utxtSze.Value = Trim(.TextMatrix(.Row, 9))
        utxtFEmp.Text = Trim(.TextMatrix(.Row, 10))
        utxtDollar.Value = Trim(.TextMatrix(.Row, 34))
        If Trim(.TextMatrix(.Row, 15)) = 0 Then
            utxtLength.Value = ""
        Else
            utxtLength.Value = Trim(.TextMatrix(.Row, 15))
        End If
        If Trim(.TextMatrix(.Row, 16)) = 0 Then
            utxtWidth.Value = ""
        Else
            utxtWidth.Value = Trim(.TextMatrix(.Row, 16))
        End If
        If Trim(.TextMatrix(.Row, 17)) = 0 Then
            utxtHeight.Value = ""
        Else
            utxtHeight.Value = Trim(.TextMatrix(.Row, 17))
        End If
        utxtUMS.Text = Trim(.TextMatrix(.Row, 18))
        If .TextMatrix(.Row, 11) = "" Then
            utxtNumDangr.Value = 0
        Else
            utxtNumDangr.Value = Trim(.TextMatrix(.Row, 11))
        End If
        utxtTshipMnt.Text = Trim(.TextMatrix(.Row, 21))
    End With
    Call InitializeDetail
End Sub
Private Sub MainProcedure()
    Call ResetValTab0
    Call ResetValTab1
    Call ResetValTab2
    Call Clear_PaymentTotals
    lngTagNewCCR = 0
    cmbPrinter.TabStop = False
    txtSupervisor.TabStop = False
    txtTranMode.TabStop = False
    cMode.Visible = True
End Sub
Private Sub ResetProcedure()
    Call ResetValTab0
    Call ResetValTab1A
    Call ResetValTab2
    Call Clear_PaymentTotals
    lngTagNewCCR = 0
    cmbPrinter.TabStop = False
    txtSupervisor.TabStop = False
    txtTranMode.TabStop = False
    cMode.Visible = True
End Sub
Private Sub utxtNumDangr_GotFocus()
    flexDangerClass.Visible = True
    flexDangerClass.Height = 3500
    flexDangerClass.Width = 5510
    flexDangerClass.ColWidth(0) = 5500
End Sub
Private Sub utxtNumDangr_LostFocus()
    flexDangerClass.Visible = False
End Sub

Private Sub utxtPref_GotFocus()
    utxtPref.SelStart = 0
    utxtPref.SelLength = Len(utxtPref.Text)
End Sub

Private Sub utxtUMS_LostFocus()
    If Len(Trim(utxtUMS.Text)) = 0 Then
        utxtUMS.Text = defaultUnitMeasurement
    End If
End Sub

Function Check_Details() As Boolean
    Check_Details = False
    
    If Trim(utxtSze.Value) = "45" Or Trim(utxtSze.Value) = "40" Or Trim(utxtSze.Value) = "20" Then
        Check_Details = True
    Else
        MsgBox "Invalid Size Code", vbCritical + vbOKOnly, "Missing Information"
        utxtSze.SetFocus
        Check_Details = False
        Exit Function
    End If
    
    If Trim(utxtFEmp.Text) = "F" Or Trim(utxtFEmp.Text) = "E" Then
        Check_Details = True
    Else
        MsgBox "Invalid Freight Kind Code", vbCritical + vbOKOnly, "Missing Information"
        utxtFEmp.SetFocus
        Check_Details = False
        Exit Function
    End If
    
    'Check_Details = False
    
    If utxtTshipMnt.Text = "F" Then
        If IsNumeric(utxtDollar.Value) Then
            If CCur(utxtDollar.Value) <= 0 Then
                MsgBox "Invalid Dollar Rate", vbExclamation + vbOKOnly, "Dollar Rate Error"
                utxtDollar.SelStart = 1
                utxtDollar.SelLength = Len(Trim(utxtDollar.Value))
                Check_Details = True
                utxtDollar.SetFocus
            End If
        End If
        If Len(Trim(utxtDollar.Value)) = 0 And UCase(Trim(utxtTshipMnt.Text)) = "F" Then
            MsgBox "Invalid Dollar Rate", vbExclamation + vbOKOnly, "Dollar Rate Error"
            utxtDollar.SelStart = 1
            utxtDollar.SelLength = Len(Trim(utxtDollar.Value))
            Check_Details = True
            utxtDollar.SetFocus
        End If
    End If
    
End Function

Private Sub ReSequence_ItemNum()

Dim lngCtrRow As Long
Dim strAstr As String * 1
Dim strForced As String * 1
Dim lngCtr As Long
lngCtr = 0

For lngCtrRow = 1 To lngRow
    
    lngCtr = lngCtr + 1
    
    If lngCtr > 8 Then
            lngCtr = 1
            strAstr = "*"
            strForced = "F"
    Else
        If Trim(flexDetails.TextMatrix(lngCtrRow, 0) = "*") And Len(Trim(flexDetails.TextMatrix(lngCtrRow, 35))) = 0 Then
            lngCtr = 1
            strAstr = "*"
            strForced = ""
        Else
            strForced = ""
            strAstr = ""
        End If
    End If

    flexDetails.TextMatrix(lngCtrRow, 35) = Trim(strForced)
    flexDetails.TextMatrix(lngCtrRow, 0) = Trim(strAstr)
    flexDetails.TextMatrix(lngCtrRow, 36) = Trim(lngCtr)

Next
End Sub

Private Sub Clear_PaymentTotals()
    Dim X As Integer
    For X = 0 To 4
        utxtChq(X).Value = ".00"
        utxtChqNo(X).Text = ""
        utxtChqBnk(X).Text = ""
    Next
    utxtAmtPay.Caption = ".00"
    utxtCsh.Value = ".00"
    utxtAdrAmt.Value = ".00"
    utxtChange.Caption = ".00"
    utxtCustName.Text = ""
    utxtCustNo.Value = ""
    utxtCustName1.Text = ""
    utxtCustNo1.Value = ""
End Sub

Public Function GetNextControlNumber() As Long
' ** call Control Extraction Number function
Dim lngTempControlNo As Long
Const strTempTyp = "CCR"
lngTempControlNo = 0
DE.GetControlNo strTempTyp, lngTempControlNo
' ** retrieve the returned values
GetNextControlNumber = lngTempControlNo
End Function

'PRNH -With Company Code
Public Function GetNextCCRNumber(ByVal compCode As String) As Long
Dim NxtNo As Long
Dim tp As Recordset
Dim StartCCR As Long
Dim EndCCR As Long
Dim PrvCCR As Long

DE.CCRAllocation UCase(gUserID)
Set tp = DE.rsCCRAllocation

With tp
    If .RecordCount > 0 Then
        .MoveFirst
        While Not .EOF = True
            If .Fields("CompanyCode") = compCode Then
                StartCCR = .Fields("strccr")
                EndCCR = .Fields("endccr")
                PrvCCR = .Fields("prvccr")
                
            End If
            .MoveNext
        Wend
    Else
        StartCCR = 0
        EndCCR = 0
        PrvCCR = 0
    End If
    NxtNo = 0
    If EndCCR = 0 Then
        NxtNo = 0
    Else
        If PrvCCR = EndCCR Then
            NxtNo = -1
        Else
            If PrvCCR = 0 Or PrvCCR > EndCCR Or PrvCCR < StartCCR Then
                NxtNo = StartCCR
            Else
                NxtNo = PrvCCR + 1
            End If
        End If
    End If
    tp.Close
End With

Set tp = Nothing
GetNextCCRNumber = NxtNo

'PRNH - Old
'Public Function GetNextCCRNumber() As Long
'Dim NxtNo As Long
'Dim tp As Recordset
'Dim StartCCR As Long
'Dim EndCCR As Long
'Dim PrvCCR As Long
'
'DE.CCRAllocation UCase(gUserID)
'Set tp = DE.rsCCRAllocation
'
'If tp.RecordCount > 0 Then
'    StartCCR = tp.Fields("strccr")
'    EndCCR = tp.Fields("endccr")
'    PrvCCR = tp.Fields("prvccr")
'Else
'    StartCCR = 0
'    EndCCR = 0
'    PrvCCR = 0
'End If
'NxtNo = 0
'If EndCCR = 0 Then
'    NxtNo = 0
'Else
'    If PrvCCR = EndCCR Then
'        NxtNo = -1
'    Else
'        If PrvCCR = 0 Or PrvCCR > EndCCR Or PrvCCR < StartCCR Then
'            NxtNo = StartCCR
'        Else
'            NxtNo = PrvCCR + 1
'        End If
'    End If
'End If
'tp.Close
'Set tp = Nothing
'GetNextCCRNumber = NxtNo
'
End Function

Public Function GetPrevCCRPay(strContNum As String, ByRef lngPrevCCR As Long, _
                 ByRef lngOvrCCr As Long, ByRef strccrNodte As String, _
                ByRef strTeller As String, ByRef strExprtr As String) As Boolean

Dim PrevCCR As Recordset
strContNum = Trim(strContNum)
' ** Call to Get Previous Payment
DE.getPrevCCRPayment strContNum
' ** Return Recordset
Set PrevCCR = DE.rsgetPrevCCRPayment
GetPrevCCRPay = False
If PrevCCR.RecordCount > 0 Then
    With PrevCCR
        lngPrevCCR = .Fields("ccrnum")
        If Not IsNull(.Fields("ovrccr")) Then
            lngOvrCCr = .Fields("ovrccr")
        Else
            lngOvrCCr = 0
        End If
        strccrNodte = Format(.Fields("sysdttm"), "yyyy-mm-dd Hh:Nn")
        strTeller = .Fields("userid")
        strExprtr = .Fields("exprtr")
    End With
GetPrevCCRPay = True
End If
PrevCCR.Close
Set PrevCCR = Nothing
End Function

Private Function EvaluateChange() As Currency
Dim TotalAmount As Currency
Dim AmtPAy As Currency
Dim lngCtrX As Long
TotalAmount = 0
AmtPAy = 0
lngCtrX = 0

If IsNumeric(utxtCsh.Value) Then
    TotalAmount = TotalAmount + CCur(utxtCsh.Value)
End If
For lngCtrX = 0 To 4
    If IsNumeric(utxtChq(lngCtrX).Value) Then
        TotalAmount = TotalAmount + CCur(utxtChq(lngCtrX).Value)
    End If
Next
If IsNumeric(utxtAdrAmt.Value) Then
    TotalAmount = TotalAmount + CCur(utxtAdrAmt.Value)
End If
AmtPAy = CCur(txtTotDue.Caption)
EvaluateChange = TotalAmount - AmtPAy
If EvaluateChange >= 0 Then
    If chkUG = "Y" Then
        cmdPrint.Enabled = False
    Else
        cmdPrint.Enabled = True
    End If
Else
    cmdPrint.Enabled = False
End If
End Function
Private Sub Clear_GrandTotal()
    txtIctsiDue.Caption = ""
    txtPpaTotal.Caption = ""
    txtTotDue.Caption = ""
End Sub
Private Function SavetoCCRCyx() As Boolean
    Dim X As Integer
    Dim rstCCRCyx As ADODB.Recordset  'Recordset for Export File
    Dim lngTempItem As Long
    Dim lngTempRow As Long
    Dim lngTempSeq As Long
    Dim lngTempCCR As Long
    Dim lngTempPrevCCr As Long
    Dim strTempCntnum As String * 12
    Dim strContNo As String
    Dim curOverSizeAmt As Currency
    Dim curDGAmt As Currency
    Dim strOOGGranted As String
    Dim strDGGranted As String
    Dim N4CurrentCategory As String
    Dim bHasUnitOut As Boolean

    For X = 1 To 100
        CCRList(X).CCRnum = 0
        CCRList(X).Refnum = 0
        CCRList(X).Seqnum = 0
        CCRList(X).Cusnam = ""
    Next
    lRef = 1
    strContNo = ""
    SavetoCCRCyx = False
    lngTempCCR = 0
    lngTempItem = 0
    lngTempRow = 1
    lngTempSeq = 1
    lngTempPrevCCr = 0
    strTempCntnum = ""
    Set rstCCRCyx = New ADODB.Recordset
    rstCCRCyx.CursorType = adOpenDynamic
    rstCCRCyx.LockType = adLockOptimistic
    rstCCRCyx.Open "CCRcyx", gcnnBilling, , , adCmdTable
    Do Until lngCCR > 0
        'PRNH-old
        'DE.GetNextCCR UCase(gUserID), lngCCR
        
        'PRNH - Company Code
        DE.GetNextCCR UCase(gUserID), lngCCR, flexDetails.TextMatrix(lngTempRow, 45)
        Call MsgCheck(lngCCR)
    Loop
        CCRList(lRef).CCRnum = lngCCR
        CCRList(lRef).Refnum = lngRefnum
        CCRList(lRef).Seqnum = lRef
        CCRList(lRef).Cusnam = gCusnam
        If Len(Trim(flexDetails.TextMatrix(lngTempRow, 7))) = 0 Then
            CCRList(lRef).UGCode = ""
        Else
            CCRList(lRef).UGCode = Trim(flexDetails.TextMatrix(lngTempRow, 7))
        End If
    With flexDetails
        Do While Not (lngTempRow > lngRow)
            If Trim(.TextMatrix(lngTempRow, 0)) = "*" And lngTempRow > 1 Then
                ' ** apply previous ccr
                'PRNH - Old
                'ApplyCCR (lngCCR)
                'with Company Code
                Call ApplyCCR(lngCCR, .TextMatrix(lngTempRow, 45))
                
                ' ** now get the next ccr available
                'PRNH-Old
                'lngCCR = GetNextCCRNumber
                'with Company Code
                lngCCR = GetNextCCRNumber(.TextMatrix(lngTempRow, 45))
                
                
                Do Until lngCCR > 0
                    'PRNH-Old
                    'lngCCR = GetNextCCRNumber
                    'With Company Code
                    lngCCR = GetNextCCRNumber(.TextMatrix(lngTempRow, 45))
                    
                    Call MsgCheck(lngCCR)
                Loop
                lngTempSeq = lngTempSeq + 1
                lngTempItem = 0
                lRef = lRef + 1
                CCRList(lRef).CCRnum = lngCCR
                CCRList(lRef).Refnum = lngRefnum
                CCRList(lRef).Seqnum = lRef
                CCRList(lRef).Cusnam = gCusnam
                If Len(Trim(.TextMatrix(lngTempRow, 7))) = 0 Then
                    CCRList(lRef).UGCode = ""
                Else
                    CCRList(lRef).UGCode = Trim(.TextMatrix(lngTempRow, 7))
                End If
            
            End If
            rstCCRCyx.AddNew
            '** Clear EXPOR21 Variables
            Call ClearExpor21
            
            lngTempItem = lngTempItem + 1
            rstCCRCyx.Fields("itmnum") = lngTempItem
            rstCCRCyx.Fields("refnum") = lngRefnum
            rstCCRCyx.Fields("seqnum") = lngTempSeq
            rstCCRCyx.Fields("ccrnum") = lngCCR
            RSet strTempCntnum = (.TextMatrix(lngTempRow, 2))
            strContNo = .TextMatrix(lngTempRow, 2)
            rstCCRCyx.Fields("cntnum") = strTempCntnum
            strExpr21Contnum = strTempCntnum
            rstCCRCyx.Fields("arrvat") = ReturnCurrency(CCur(.TextMatrix(lngTempRow, 4)))
            rstCCRCyx.Fields("arrtax") = ReturnCurrency(.TextMatrix(lngTempRow, 5))
            rstCCRCyx.Fields("wghamt") = ReturnCurrency(.TextMatrix(lngTempRow, 6))
            If Len(Trim(.TextMatrix(lngTempRow, 7))) = 0 Then
                rstCCRCyx.Fields("guarntycde") = ""
            Else
                rstCCRCyx.Fields("guarntycde") = .TextMatrix(lngTempRow, 7)
            End If
            rstCCRCyx.Fields("arramt") = ReturnCurrency(CCur(.TextMatrix(lngTempRow, 8)))
            If Len(Trim(.TextMatrix(lngTempRow, 9))) = 0 Then
                rstCCRCyx.Fields("cntsze") = 0
            Else
                rstCCRCyx.Fields("cntsze") = .TextMatrix(lngTempRow, 9)
            End If
            If Len(Trim(.TextMatrix(lngTempRow, 10))) = 0 Then
                rstCCRCyx.Fields("fulemp") = ""
            Else
                rstCCRCyx.Fields("fulemp") = .TextMatrix(lngTempRow, 10)
            End If
            If Len(Trim(.TextMatrix(lngTempRow, 11))) = 0 Then
                rstCCRCyx.Fields("dgrcls") = ""
            Else
                rstCCRCyx.Fields("dgrcls") = .TextMatrix(lngTempRow, 11)
            End If
            rstCCRCyx.Fields("dgramt") = ReturnCurrency(.TextMatrix(lngTempRow, 12))
            'Added Navis Project Team 09/26/2009
            curDGAmt = ReturnCurrency(.TextMatrix(lngTempRow, 12))
            If Len(Trim(.TextMatrix(lngTempRow, 13))) = 0 Then
                rstCCRCyx.Fields("vslcde") = ""
            Else
                rstCCRCyx.Fields("vslcde") = .TextMatrix(lngTempRow, 13)
            End If
            If Len(Trim(.TextMatrix(lngTempRow, 14))) = 0 Then
                rstCCRCyx.Fields("vatcde") = 1
            Else
                rstCCRCyx.Fields("vatcde") = .TextMatrix(lngTempRow, 14)
            End If
            
            TLength = 0
            TWidth = 0
            THeight = 0
            TUms = ""
            curOverSizeAmt = 0
            If Len(Trim(.TextMatrix(lngTempRow, 18))) = 0 Then
                rstCCRCyx.Fields("ovzums") = ""
                rstCCRCyx.Fields("ovzamt") = 0
                rstCCRCyx.Fields("cntovzl") = 0
                rstCCRCyx.Fields("cntovzw") = 0
                rstCCRCyx.Fields("cntovzh") = 0
                'Added Navis Project Team 09/26/2009
                curOverSizeAmt = 0
            Else
                rstCCRCyx.Fields("ovzums") = .TextMatrix(lngTempRow, 18)
                rstCCRCyx.Fields("ovzamt") = ReturnCurrency(.TextMatrix(lngTempRow, 19))
                'Added Navis Project Team 09/26/2009
                curOverSizeAmt = ReturnCurrency(.TextMatrix(lngTempRow, 19))
                TUms = .TextMatrix(lngTempRow, 18)
                ConvertSize TLength, TWidth, THeight, TUms, ReturnSingle(CSng(.TextMatrix(lngTempRow, 15))), _
                    ReturnSingle(CSng(.TextMatrix(lngTempRow, 16))), ReturnSingle(CSng(.TextMatrix(lngTempRow, 17)))
                
                    rstCCRCyx.Fields("cntovzl") = TLength
                    rstCCRCyx.Fields("cntovzw") = TWidth
                    rstCCRCyx.Fields("cntovzh") = THeight
                
            End If
            rstCCRCyx.Fields("revton") = ReturnCurrency(CCur(.TextMatrix(lngTempRow, 20)))
            If Len(Trim(.TextMatrix(lngTempRow, 21))) = 0 Then
                rstCCRCyx.Fields("trncde") = ""
                strExpr21Trncde = ""
            Else
                rstCCRCyx.Fields("trncde") = .TextMatrix(lngTempRow, 21)
                strExpr21Trncde = .TextMatrix(lngTempRow, 21)
            End If
            If Len(Trim(.TextMatrix(lngTempRow, 22))) = 0 Then
                rstCCRCyx.Fields("whfcde") = 0
                strExpr21Whfcde = 0
            Else
                rstCCRCyx.Fields("whfcde") = .TextMatrix(lngTempRow, 22)
                strExpr21Whfcde = .TextMatrix(lngTempRow, 22)
                If strExpr21Whfcde = "4" And strWhfOnly = "Y" Then
                    If Trim(.TextMatrix(lngTempRow, 37)) = "P" Then
                        DE.UpdatePrevPay strTempCntnum, Trim(.TextMatrix(lngTempRow, 33))
                    End If
                End If
            End If
            If Len(Trim(.TextMatrix(lngTempRow, 23))) = 0 Then
                rstCCRCyx.Fields("broker") = ""
            Else
                rstCCRCyx.Fields("broker") = .TextMatrix(lngTempRow, 24)
            End If
            If Len(Trim(.TextMatrix(lngTempRow, 23))) = 0 Then
                rstCCRCyx.Fields("exprtr") = ""
                strExpr21ExpName = ""
            Else
                rstCCRCyx.Fields("exprtr") = .TextMatrix(lngTempRow, 23)
                strExpr21ExpName = .TextMatrix(lngTempRow, 23)
            End If
            If Len(Trim(.TextMatrix(lngTempRow, 25))) = 0 Then
                rstCCRCyx.Fields("entnum") = ""
            Else
                rstCCRCyx.Fields("entnum") = .TextMatrix(lngTempRow, 25)
            End If
            
            If Len(Trim(.TextMatrix(lngTempRow, 26))) = 0 Then
                rstCCRCyx.Fields("commod") = ""
            Else
                rstCCRCyx.Fields("commod") = .TextMatrix(lngTempRow, 26)
            End If

            If Len(Trim(.TextMatrix(lngTempRow, 33))) = 0 Then
                lngTempPrevCCr = 0
            Else
                lngTempPrevCCr = Trim(.TextMatrix(lngTempRow, 33))
            End If
            rstCCRCyx.Fields("dolrte") = ReturnCurrency(.TextMatrix(lngTempRow, 34))
            If Trim(.TextMatrix(lngTempRow, 37)) = "P" Then
                strTempCntnum = ""
                rstCCRCyx.Fields("ovrccr") = lngTempPrevCCr
                rstCCRCyx.Fields("status") = "OVR"
                lngTempPrevCCr = 0
            Else
                rstCCRCyx.Fields("ovrccr") = 0
            End If
            rstCCRCyx.Fields("remark") = .TextMatrix(lngTempRow, 38)
            rstCCRCyx.Fields("whfamt") = ReturnCurrency(.TextMatrix(lngTempRow, 40))
            rstCCRCyx.Fields("ppanum") = 0
            rstCCRCyx.Fields("status") = ""
            rstCCRCyx.Fields("sysdttm") = Format(strDate, "yyyy-mm-dd Hh:Nn:ss")
            rstCCRCyx.Fields("userid") = UCase(gUserID)
            rstCCRCyx.Fields("trknam") = ""
            rstCCRCyx.Fields("pltnum") = ""
            rstCCRCyx.Fields("trkchs") = ""
            rstCCRCyx.Fields("updcde") = ""
            rstCCRCyx.Fields("supvsr") = Mid(Trim(txtSupervisor.Text), 1, 15)
            
            'PRNH - Conmpany Code
            rstCCRCyx.Fields("CompanyCode") = .TextMatrix(lngTempRow, 45)
            
            lngExpr21Refnum = lngRefnum
            lngExpr21Date = Format(strDate, "yyyymmdd")
            lngExpr21Time = Format(strDate, "HHNnss")
            strExpr21Userid = UCase(gUserID)
            lblSaveMessage.Caption = "Writing to Database. Please Wait . . . . . . "
            lblSaveMessage.Refresh
            rstCCRCyx.Update
            
'            'PRNH - Company code
'            Call ApplyCCR(lngCCR, .TextMatrix(lngTempRow, 45))
            lngTempRow = lngTempRow + 1
            
            'Save to Sparcs
            'Dim gkey As String
            'gkey = GetGKey(strTempCntnum, "QUEUED", "")
            'Call SavePaymentToSparcs(MoveToField(Column.ContainerID, "C"), "STORAGE", "", gkey)
            'Dim n4Updater As BillingInventory.ArgoServices(operatorId, complexId, facilityId, yardId,
            'username, password, uri)
            
            'Dim n4Updater As BillingInventory.ArgoServices
            'n4Updater = New BillingInventory.ArgoServices
            
            'n4Updater = New BillingInventory.ArgoServices "ICTSI", "PH", "SBITC", "SBITC, "strN4UserName", "strN4Password", "http://sbitc-dev:9080/apex/services/argoservice?wsdl"
            
            'If n4Updater.UpdateHoldPermission(strTempCntnum, "BILLING", "GRANT_PERMISSION").CommonResponseMessage.Status = "0" Then
                'Update SBITC Billing IsN4BillingPermissionGranted
            '    MsgBox ("Updated In N4")
            'End If
            
            'Call ReleaseHold(MoveToField(Column.ContainerID, "C"))
            '.TextMatrix(lngTempRow, 2)
            
'            If (ReleaseHold(rstCCRCyx.Fields("cntnum")) = "0") Then
'                UpdateIsN4BillingPermissionGrantedStatus (rstCCRCyx.Fields("cntnum"))
'            End If

            Call GetContainerLastestCategory(strContNo, N4CurrentCategory, bHasUnitOut)
            
            If N4CurrentCategory = "EXPRT" And bHasUnitOut = False Then
                If (ReleaseHold(strContNo) = "0") Then
                    UpdateIsN4BillingPermissionGrantedStatus (strContNo)
                End If
                
                If curOverSizeAmt > 0 Then
                    strOOGGranted = ReleaseOOGHold(strContNo)
                    If strOOGGranted = "0" Then
                        Call GrantOOGPermission(strContNo, lngCCR)
                    End If
                End If
                
                If curDGAmt > 0 Then
                   strDGGranted = ReleaseDGHold(strContNo)
                    If strDGGranted = "0" Then
                        Call GrantDGPermission(strContNo, lngCCR)
                    End If
                End If
            End If


        Loop
        'PRNH - Removed due to Company Code extraction
        'ApplyCCR (lngCCR)
        'PRNH - Company code
        'Call ApplyCCR(lngCCR, .TextMatrix(lngTempRow, 45))
        Call ApplyCCR(lngCCR, .TextMatrix(1, 45))
    End With
    rstCCRCyx.Close
    Set rstCCRCyx = Nothing
    SavetoCCRCyx = True
    Exit Function
ErrorHd:
    SavetoCCRCyx = False
End Function

Public Function SaveToCCRPay() As Boolean
    Dim sngTotalCsh As Currency
    Dim sngAdramt As Currency
    Dim sngChange As Currency
    Dim lngCtrX As Long
    Dim lngAdr As Long
    Dim lngAdrControl As Long
    Dim rstCCRPay As ADODB.Recordset
    Set rstCCRPay = New ADODB.Recordset
    rstCCRPay.CursorType = adOpenDynamic
    rstCCRPay.LockType = adLockOptimistic
    rstCCRPay.Open "CCRpay", gcnnBilling, , , adCmdTable
    gCusnam = ""
    sngTotalCsh = 0
    sngChange = 0
    sngAdramt = 0
    lngCtrX = 0
    lngAdrControl = 0
    If ReturnCurrency(utxtCsh.Value) > 0 Then
        sngTotalCsh = sngTotalCsh + CCur(utxtCsh.Value)
    End If
    If ReturnLong(utxtCustNo.Value) > 0 And ReturnCurrency(utxtAdrAmt.Value) > 0 Then
        sngAdramt = ReturnCurrency(utxtAdrAmt.Value)
    End If
    If Len(utxtChange.Caption) > 0 Then
        sngChange = utxtChange.Caption
    End If
    If ReturnLong(utxtCustNo.Value) > 0 And ReturnCurrency(CStr(sngAdramt)) > 0 Then
            lngAdrControl = DE.ApllyADR(Trim(utxtCustNo.Value), "CCR", lngRefnum, sngAdramt, "", Trim(gUserID))
            If lngAdrControl <= 0 Then
                SaveToCCRPay = False
                Exit Function
            End If
    End If
    With rstCCRPay
        .AddNew
        strDate = zGetSysDate
        strDate = Format(strDate, "yyyy-mm-dd Hh:Nn:ss")
        '.Fields ("")
        If Len(utxtCustNo.Value) > 0 Then
            .Fields("cuscde") = utxtCustNo.Value
            .Fields("cusnam") = utxtCustName.Text
            gCusnam = Trim(utxtCustName.Text)
        Else
            .Fields("cuscde") = 0
            .Fields("cusnam") = ""
            gCusnam = ""
        End If
        .Fields("ftramt") = 0
        .Fields("ftrfee") = 0
        .Fields("ackno") = ""
        
        'PRNH - Removed ADR Validations
        'If ReturnLong(utxtCustNo.Value) > 0 And ReturnCurrency(CStr(utxtAdrAmt.Value)) > 0 Then
            .Fields("adrnum") = 0 'lngAdrControl - No data yet
            .Fields("adramt") = ReturnCurrency(utxtAdrAmt.Value)
        'Else
        '    .Fields("adramt") = 0
        '    .Fields("adrnum") = 0
        'End If
        
        If Len(sngTotalCsh) <> 0 Then
            .Fields("cshamt") = sngTotalCsh
        Else
            .Fields("cshamt") = 0
        End If
        If Len(utxtChq(0).Value) <> 0 Then
            .Fields("chkamt1") = CCur(utxtChq(0).Value)
            .Fields("chkno1") = Trim(utxtChqNo(0).Text)
            .Fields("chkbnk1") = Trim(utxtChqBnk(0).Text)
        Else
            .Fields("chkamt1") = 0
            .Fields("chkno1") = ""
            .Fields("chkbnk1") = ""
        End If
        If Len(utxtChq(1).Value) <> 0 Then
            .Fields("chkamt2") = CCur(utxtChq(1).Value)
            .Fields("chkno2") = Trim(utxtChqNo(1).Text)
            .Fields("chkbnk2") = Trim(utxtChqBnk(1).Text)
        Else
            .Fields("chkamt2") = 0
            .Fields("chkno2") = ""
            .Fields("chkbnk2") = ""
        End If
        If Len(utxtChq(2).Value) <> 0 Then
            .Fields("chkamt3") = CCur(utxtChq(2).Value)
            .Fields("chkno3") = Trim(utxtChqNo(2).Text)
            .Fields("chkbnk3") = Trim(utxtChqBnk(2).Text)
        Else
            .Fields("chkamt3") = 0
            .Fields("chkno3") = ""
            .Fields("chkbnk3") = ""
        End If
        If Len(utxtChq(3).Value) <> 0 Then
            .Fields("chkamt4") = CCur(utxtChq(3).Value)
            .Fields("chkno4") = Trim(utxtChqNo(3).Text)
            .Fields("chkbnk4") = Trim(utxtChqBnk(3).Text)
        Else
            .Fields("chkamt4") = 0
            .Fields("chkno4") = ""
            .Fields("chkbnk4") = ""
        End If
        If Len(utxtChq(4).Value) <> 0 Then
            .Fields("chkamt5") = CCur(utxtChq(4).Value)
            .Fields("chkno5") = Trim(utxtChqNo(4).Text)
            .Fields("chkbnk5") = Trim(utxtChqBnk(4).Text)
        Else
            .Fields("chkamt5") = 0
            .Fields("chkno5") = ""
            .Fields("chkbnk5") = ""
        End If
        If Len(sngChange) <> 0 Then
            .Fields("chgamt") = sngChange
        End If
        .Fields("refnum") = lngRefnum
        .Fields("userid") = UCase(gUserID)
        .Fields("sysdttm") = Format(strDate, "yyyy-mm-dd Hh:Nn:ss")
        .Fields("ccrtyp") = "1"
        .Fields("status") = ""
        .Fields("rectag") = ""
        .Fields("updcde") = ""
        If DomesticMode Then
            .Fields("ccrmod") = "D"
        Else
            .Fields("ccrmod") = "F"
        End If
        .Update
    End With
    rstCCRPay.Close
    Set rstCCRPay = Nothing
    SaveToCCRPay = True
    Exit Function
ErrorHd:
    SaveToCCRPay = False
End Function

Private Sub SaveToExpor21()
'   **  Save Details to File EXPOR21
Dim rstExp21 As Recordset
lngExpr21Refnum = lngRefnum
DE.SelectCnt lngRefnum
Set rstExp21 = DE.rsSelectCnt
With rstExp21
If .RecordCount > 0 Then
    Do Until .EOF
        Call ClearExpor21
        RSet strExpr21Contnum = .Fields("cntnum")
        lngExpr21Refnum = lngRefnum
        strExpr21Trncde = .Fields("trncde")
        strExpr21ExpName = Trim(.Fields("exprtr"))
        strExpr21Whfcde = .Fields("whfcde")
        lngExpr21Date = Format(.Fields("sysdttm"), "yyyymmdd")
        lngExpr21Time = Format(.Fields("sysdttm"), "HHNnss")
        strExpr21Userid = .Fields("userid")
'**************************
'* Stand Alone FLR
'        With CTCSinfo
'            .WriteCYExport strExpr21Contnum, lngExpr21Refnum, lngExpr21Date, _
'            lngExpr21Time, strExpr21Userid, strExpr21ExpName, strExpr21Trncde, strExpr21Whfcde
'        End With
    .MoveNext
    Loop
End If
.Close
End With
Set rstExp21 = Nothing
End Sub

Private Sub ResetValTab0()
    ' ** initialize tab0 values - DETAILS
    flexDetails.Visible = False
    flexDetails.Clear
    flexDetails.Rows = 2
    Call GridHeader
    lngRow = 0
    lngNoCnt = 0
    lblNoCnt.Caption = ""
    txtIctsiDue.Caption = ".00"
    txtPpaTotal.Caption = ".00"
    txtTotDue.Caption = ".00"
    
    Call ReformatDisplayTotal
    
    utxtPref.Text = ""
    utxtNo.Text = ""
    utxtSze.Value = ""
    utxtFEmp.Text = ""
    
    chkNewCCR.Value = vbChecked
    utxtDollar.Value = 0
    utxtLength.Value = "0"
    utxtWidth.Value = "0"
    utxtHeight.Value = "0"
    utxtUMS.Text = defaultUnitMeasurement
    
    utxtNumDangr.Value = ""
    utxtTshipMnt.Text = ""
    flexDangerClass.Visible = False
    flexTshipMnt.Visible = False
    cmdGrid.Enabled = False
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdHeader.Enabled = lngRow > 0 And flexDetails.Enabled = False
    cmdPayment.Enabled = False
    utxtDollar.Visible = False
    Label7.Visible = False
    utxtDollar.Enabled = False
    utxtDollar.TabStop = False
    utxtDollar.Value = 0
    
End Sub
Private Sub ResetValTab1()
    Dim intCtr As Integer
    ' ** initialize tab1 values - Header
    utxtBroker.Text = ""
    utxtExporter.Text = ""
    utxtCommodity.Text = ""
    utxtRemark.Text = ""
    utxtVessel.Text = ""
    For intCtr = 0 To 10
        utxtEntry1(intCtr).Value = ""
    Next
    utxtVatCode.Value = "1"
    utxtUGuarantee.Text = "N"
    FrameCust.Visible = False
    Call Tab01off(False)
    ' ** initialize buttons
    cmdCancel.Visible = False
End Sub
Private Sub ResetValTab1A()
Dim intCtr As Integer
    ' ** initialize tab1 values - Header
    utxtVatCode.Value = "1"
    utxtUGuarantee.Text = "N"
    FrameCust.Visible = False
    Call Tab01off(False)
    ' ** initialize buttons
    cmdCancel.Visible = False
    For intCtr = 0 To 10
        utxtEntry1(intCtr).Value = ""
    Next
End Sub
Private Sub ResetValTab2()
    ' ** initialize tab2 values - Payment
    Call Clear_PaymentTotals
    Call Tab02off(False)
End Sub
Private Function EnableForNewContainer(blnEnable As Boolean)
    Tab00off (Not blnEnable)
    utxtPref.Enabled = blnEnable
    utxtPref.TabStop = blnEnable
    utxtNo.Enabled = blnEnable
    utxtNo.TabStop = blnEnable
    cmdAdd.Enabled = Not blnEnable
    cmdDelete.Enabled = Not blnEnable
    utxtSze.Enabled = Not blnEnable
    utxtFEmp.Enabled = Not blnEnable
    chkNewCCR.Enabled = Not blnEnable
    utxtNumDangr.Enabled = Not blnEnable
    utxtTshipMnt.Enabled = Not blnEnable
    utxtDollar.Enabled = Not blnEnable
    utxtLength.Enabled = Not blnEnable
    utxtWidth.Enabled = Not blnEnable
    utxtHeight.Enabled = Not blnEnable
    utxtUMS.Enabled = Not blnEnable
    utxtSze.TabStop = Not blnEnable
    utxtFEmp.TabStop = Not blnEnable
    chkNewCCR.TabStop = Not blnEnable
    utxtNumDangr.TabStop = Not blnEnable
    utxtTshipMnt.TabStop = Not blnEnable
    utxtDollar.TabStop = Not blnEnable
    utxtLength.TabStop = Not blnEnable
    utxtWidth.TabStop = Not blnEnable
    utxtHeight.TabStop = Not blnEnable
    utxtUMS.TabStop = Not blnEnable
    If blnEnable = True Then
        utxtPref.BorderStyle = 1
        utxtNo.BorderStyle = 1
    Else
        utxtPref.BorderStyle = 0
        utxtNo.BorderStyle = 0
    End If
End Function
Private Sub MsgCheck(lngTempCCR As Long)
Dim strMsg1 As String
Dim strMsg2 As String
strMsg1 = ""
strMsg2 = ""
    Select Case lngTempCCR
    Case -1
        strMsg1 = "There are no CCR  allocated to " & UCase(gUserID)
        strMsg2 = "Please verify to Supervisor for UPDATE."
        Call ErrorMessage(strMsg1, strMsg2)
    Case 0
        strMsg1 = "CCR Number is not in range of allocation to " & UCase(gUserID)
        strMsg2 = "Please verify to Supervisor for UPDATE."
        Call ErrorMessage(strMsg1, strMsg2)
    End Select
End Sub
Private Function ReturnSingle(AmountStr As String) As Single
    If IsNumeric(AmountStr) Then
        ReturnSingle = CSng(AmountStr)
    Else
        ReturnSingle = 0
    End If
End Function
Private Function ReturnLong(AmountStr As String) As Long
    If IsNumeric(AmountStr) Then
        ReturnLong = CLng(AmountStr)
    Else
        ReturnLong = 0
    End If
End Function
Private Function ReturnCurrency(AmountStr As String) As Currency
    If IsNumeric(AmountStr) Then
        ReturnCurrency = CCur(AmountStr)
    Else
        ReturnCurrency = 0
    End If
End Function
Private Sub ErrorMessage(strError1 As String, strError2 As String)
    Beep
    frmMessPYX1.Timer1.Enabled = True
    frmMessPYX1.lblMessPYX11.Caption = strError1
    frmMessPYX1.lblMessPYX12.Caption = strError2
    frmMessPYX1.Show (vbModal)
End Sub
Private Sub SystemMessage(strError3 As String, strError4 As String, strError5 As String)
    Beep
    frmMessPYXS1.lblMessPYXS11.Caption = strError3
    frmMessPYXS1.lblMessPYXS12.Caption = strError4
    frmMessPYXS1.lblMessPYXS13.Caption = strError5
    frmMessPYXS1.Show (vbModal)
End Sub
Public Sub GetAllocation()
    Dim gI As Recordset
    Dim Allocation As Recordset
    cmdContinue.Enabled = True
    utxtWorkStn.Text = ""
    utxtStrtCCR.Text = ""
    utxtEndCCR.Text = ""
    utxtLastCCR.Text = ""
    utxtLastIssuedDte.Text = ""
    DE.getInformation
    Set gI = DE.rsgetInformation
    If gI.RecordCount > 0 Then
        With gI
            utxtWorkStn.Text = .Fields("workstation")
        End With
    Else
        gUserID = "Invalid"
        utxtWorkStn.Text = "Invalid"
        cmdContinue.Enabled = False
    End If
    gI.Close
    Set gI = Nothing
    txtUserid.Text = UCase(gUserID)
    DE.CCRAllocation UCase(gUserID)
    Set Allocation = DE.rsCCRAllocation
    
'    If Allocation.RecordCount > 0 Then
'        With Allocation
'            lngCCRStart = .Fields("strccr")
'            lngCCREnd = .Fields("endccr")
'            If Not IsNull(.Fields("prvccr")) Then
'                If IsNumeric(.Fields("PRVCCR")) Then
'                    lngCCRLastIssued = .Fields("prvccr")
'                End If
'            Else
'                lngCCRLastIssued = 0
'            End If
'            If Not IsNull(.Fields("prvdte")) Then
'                dtmCCRLastIssuedDate = .Fields("prvdte")
'            Else
'                dtmCCRLastIssuedDate = " "
'            End If
'            utxtStrtCCR.Text = lngCCRStart
'            utxtEndCCR.Text = lngCCREnd
'            utxtLastCCR.Text = lngCCRLastIssued
'            utxtLastIssuedDte.Text = Format(dtmCCRLastIssuedDate, "yyyy-mm-dd HH:Nn:ss")
'            If Not IsNumeric(.Fields("prvccr")) Then
'                lblAlloc.Caption = ""
'            Else
'                If .Fields("prvccr") < .Fields("endccr") Then
'                    cmdContinue.Enabled = True
'                    lblAlloc.Caption = ""
'                Else
'                    cmdContinue.Enabled = False
'                    lblAlloc.Caption = "No CCR Allocation for Teller " & UCase(gUserID)
'                End If
'            End If
'        End With
'    Else
'        cmdContinue.Enabled = False
'        lblAlloc.Caption = "No CCR Allocation for Teller " & UCase(gUserID)
'    End If

'PRNH
    If Allocation.RecordCount > 0 Then
        With Allocation
            .MoveFirst
            While Not .EOF
                If Not IsNull(.Fields("prvdte")) Then
                    dtmCCRLastIssuedDate = .Fields("prvdte")
                Else
                    dtmCCRLastIssuedDate = " "
                End If
                
                If Not IsNull(.Fields("prvccr")) Then
                    If IsNumeric(.Fields("PRVCCR")) Then
                        lngCCRLastIssued = .Fields("prvccr")
                    End If
                Else
                    lngCCRLastIssued = 0
                End If
                
                If .Fields("CompanyCode") = "SBITC" Then
                    utxtStrtCCR = .Fields("strccr")
                    utxtEndCCR = .Fields("endccr")
                    utxtLastIssuedDte.Text = Format(dtmCCRLastIssuedDate, "yyyy-mm-dd HH:Nn:ss")
                    utxtLastCCR.Text = lngCCRLastIssued
                ElseIf .Fields("CompanyCode") = "ISI" Then
                    utxtStrtCCRISI = .Fields("strccr")
                    utxtEndCCRISI = .Fields("endccr")
                    utxtLastIssuedDteISI.Text = Format(dtmCCRLastIssuedDate, "yyyy-mm-dd HH:Nn:ss")
                    utxtLastCCRISI.Text = lngCCRLastIssued
                End If
                .MoveNext
            Wend
            
                If Val(utxtLastCCR.Text) < Val(utxtEndCCR.Text) And Val(utxtLastCCRISI.Text) < Val(utxtEndCCRISI.Text) Then
                    cmdContinue.Enabled = True
                    lblAlloc.Caption = ""
                Else
                    cmdContinue.Enabled = False
                    lblAlloc.Caption = "No CCR Allocation for Teller " & UCase(gUserID)
                End If
        End With
    Else
        cmdContinue.Enabled = False
        lblAlloc.Caption = "No CCR Allocation for Teller " & UCase(gUserID)
    End If
    Allocation.Close
    Set Allocation = Nothing
End Sub

Private Sub ClearEntry()
    strEntry0 = ""
    strEntry1 = ""
    strEntry2 = ""
    strEntry3 = ""
    strEntry4 = ""
    strEntry5 = ""
    strEntry6 = ""
    strEntry7 = ""
    strEntry8 = ""
    strEntry9 = ""
    strEntry10 = ""
End Sub
Private Sub utxtVatCode_LostFocus()
    If Len(Trim(utxtVatCode.Value)) = 0 Then
        utxtVatCode.Value = "1"
    End If
End Sub


Private Sub utxtWidth_Change()
    If Len(Trim(utxtWidth.Value)) > 0 Then
        If Len(Trim(utxtUMS.Text)) = 0 Then
            utxtUMS.Text = defaultUnitMeasurement
        End If
    End If
End Sub
Private Sub CheckPaymentOk()
    Dim lngEntry As Long
    lngEntry = 0
    lngEntry = Len(Trim(utxtEntry1(0).Value)) + _
                     Len(Trim(utxtEntry1(1).Value)) + _
                     Len(Trim(utxtEntry1(2).Value)) + _
                     Len(Trim(utxtEntry1(3).Value)) + _
                     Len(Trim(utxtEntry1(4).Value)) + _
                     Len(Trim(utxtEntry1(5).Value)) + _
                     Len(Trim(utxtEntry1(6).Value)) + _
                     Len(Trim(utxtEntry1(7).Value)) + _
                     Len(Trim(utxtEntry1(8).Value)) + _
                     Len(Trim(utxtEntry1(9).Value))
    ' ** Required Fields for Payment
    If Len(Trim(utxtExporter.Text)) <> 0 And _
        Len(Trim(utxtCommodity.Text)) <> 0 And _
        lngRow > 0 And lngEntry > 0 Then
        cmdPayment.Enabled = True
        cmdBack.Enabled = True
        cmdNewCCR.Enabled = True
    Else
    ' ** otherwise
        cmdPayment.Enabled = False
        cmdBack.Enabled = False
        cmdNewCCR.Enabled = False
    End If
End Sub
Private Sub CheckAllocationRange()
    Dim ReturnVal As Long
    Dim ReturnVal2 As Long
    ReturnVal = 0
    If IsNumeric(utxtCCRNo.Value) Then
        ReturnVal = DE.chkValidCCR(UCase(gUserID), Val(utxtCCRNo.Value), flexDetails.TextMatrix(1, 45))
        If Not ReturnVal = Val(utxtCCRNo.Value) Then
            If Not blnChanging Then
                Beep
                MsgBox "Invalid CCR Number"
                utxtCCRNo.Value = lngCCR
                utxtCCRNo.SetFocus
            End If
            lblWarning.Caption = "Invalid CCR Number"
        Else
        ' ** Check allocation
            ReturnVal2 = ReturnVal + NumberOfCCR
            ReturnVal = DE.chkValidCCR(UCase(gUserID), ReturnVal2, flexDetails.TextMatrix(1, 45))
            If Not ReturnVal = ReturnVal2 Then
                ' ** Include a warning
                lblWarning.Caption = "Number of CCR to Print is not in Allocated Range"
                lngCCR = Val(utxtCCRNo.Value)
            Else
                ' ** Disable warning
                lblWarning.Caption = ""
                lngCCR = Val(utxtCCRNo.Value)
            End If
            cmdPrint.Enabled = True
        End If
    Else
        lblWarning.Caption = "No CCR Allocation Available"
    End If
    blnChanging = False
    cmdPrint.Enabled = (Len(Trim(lblWarning.Caption)) = 0 Or Trim(lblWarning.Caption) = "Number of CCR to Print is not in Allocated Range") And Len(Trim(chkUG)) = 0
End Sub
Private Function GetNumberOfCCR() As Integer
    Dim NCCR As Integer
    Dim lngCtrRow As Long
    For lngCtrRow = 1 To lngRow
            If flexDetails.TextMatrix(lngCtrRow, 0) = "*" Then
                NCCR = NCCR + 1
            End If
    Next
    GetNumberOfCCR = NCCR - 1
End Function
Private Sub ClearExpor21()
    strExpr21Contnum = ""
    lngExpr21Refnum = 0
    lngExpr21Date = 0
    lngExpr21Time = 0
    strExpr21Userid = ""
    strExpr21Trncde = "'"
    strExpr21ExpName = ""
    strExpr21Whfcde = ""
End Sub
Public Sub StartInitialization()
    Set cusListing = New cCustomer      'Fills Customer PickList
    cusListing.FillCustomer
    '*********************************
    '* StandAlone FLR
    '    Set CTCSinfo = CreateObject("CTCS.cCTCS")
    On Error Resume Next
    '*********************************
    '* StandAlone FLR
    '    CTCSinfo.Connect
    sstMain.Visible = False
    Call AddDanger_Class
    Call AddTshipMent_Code
    Call MoveRates
    Call Tab00off(False)
    Call Tab01off(False)
    Call Tab02off(False)
    Call ResetValTab0
    Call ResetValTab1
    Call ResetValTab2
    sngAdrBalance = 0
    lngTagNewCCR = 0
    strCustName = ""
    chkUG = ""
    Load frmMessPYX1
    frmMessPYX1.Timer1.Enabled = False
    Load frmMessPYXS1
    Call GetAllocation
    Call LoadPrinter
    blnChanging = False
    txtTranMode.Text = "F"
    DomesticMode = False
    cMode.Text = "Foreign Transaction"
End Sub
Public Sub LoadPrinter()
    Dim Pr As Printer
    Dim ref As Long
    Dim refTouse As Long
    Dim strRef As String * 2
    ref = 0
    For Each Pr In Printers
        strRef = Str(ref + 1)
        cmbPrinter.AddItem strRef & "| " & Pr.DeviceName
        If Pr.DeviceName = Printer.DeviceName Then
            refTouse = ref
        End If
        ref = ref + 1
    Next Pr
    cmbPrinter.ListIndex = refTouse
End Sub
Private Sub utxtWidth_LostFocus()
    If Len(Trim(utxtWidth.Value)) = 0 Then
        utxtWidth.Value = 0
    End If
End Sub
Private Sub ClearTagCCR()
    Dim lngTempRow As Long
    lngTempRow = 0
    With flexDetails
        Do While Not (lngTempRow = lngRow)
             lngTempRow = lngTempRow + 1
             If Trim(.TextMatrix(lngTempRow, 41)) = lngTagNewCCR Then
                 .TextMatrix(lngTempRow, 41) = ""
            End If
        Loop
    End With
End Sub

'PRNH - With Company Code
Public Sub ApplyCCR(CTA As Long, compCode As String)
    'PRNH - Company Code
    Dim NDate As Date
    Dim NxtNo As Long
    Dim tp As Recordset
    Dim StartCCR As Long
    Dim EndCCR As Long
    Dim PrvCCR As Long
    DE.CCRAllocation UCase(gUserID)
    Set tp = DE.rsCCRAllocation
    With tp
        If .RecordCount > 0 Then
            .MoveFirst
            While Not .EOF
                If .Fields("CompanyCode") = compCode Then
                    StartCCR = .Fields("strccr")
                    EndCCR = .Fields("endccr")
                    PrvCCR = .Fields("prvccr")
                    If (CTA <= EndCCR) And (CTA > PrvCCR) And (CTA >= StartCCR) Then
                        .Fields("prvccr") = CTA
                        DE.NowDate NDate
                        .Fields("prvdte") = NDate
                        .Update
                        GoTo break
                    End If
                Else
                    .MoveNext
                End If
            Wend
break:
            .Close
        End If
    End With
    Set tp = Nothing
End Sub

'PRNH - OLD
'Public Sub ApplyCCR(CTA As Long)
'    Dim NDate As Date
'    Dim NxtNo As Long
'    Dim tp As Recordset
'    Dim StartCCR As Long
'    Dim EndCCR As Long
'    Dim PrvCCR As Long
'    DE.CCRAllocation UCase(gUserID)
'    Set tp = DE.rsCCRAllocation
'    If tp.RecordCount > 0 Then
'        StartCCR = tp.Fields("strccr")
'        EndCCR = tp.Fields("endccr")
'        PrvCCR = tp.Fields("prvccr")
'    Else
'        StartCCR = 0
'        EndCCR = 0
'        PrvCCR = 0
'    End If
'    If (CTA <= EndCCR) And (CTA > PrvCCR) And (CTA >= StartCCR) Then
'        tp.Fields("prvccr") = CTA
'        DE.NowDate NDate
'        tp.Fields("prvdte") = NDate
'        tp.Update
'    End If
'    tp.Close
'    Set tp = Nothing
'End Sub

Public Function PrintCCR(pRefnum As Long) As Boolean
    Dim X As Integer
    Dim CRL As Recordset
    Dim ADR As Recordset
    Dim DETTl As Recordset
    DE.getAdrAmt (pRefnum)
    Set ADR = DE.rsgetADRAmt
    sngTempAmt = 0
    ChkAmt1 = 0
    ChkAmt2 = 0
    ChkAmt3 = 0
    ChkAmt4 = 0
    ChkAmt5 = 0
    ChkTotal = 0
    AdrAmt = 0
    CashAmt = 0
    ChkAmount = 0
    AdrAmount = 0
    CashAmount = 0
    If ADR.Fields("adramt") <> 0 Or Not IsNull(ADR.Fields("adramt")) Then
        AdrAmt = ADR.Fields("adramt")
    End If
    If ADR.Fields("chkamt1") <> 0 Or Not IsNull(ADR.Fields("chkamt1")) Then
        ChkAmt1 = ADR.Fields("chkamt1")
        ChkTotal = ChkTotal + ChkAmt1
        StrChk1 = ADR.Fields("chkno1") & " " & Format(ADR.Fields("chkamt1"), "#########.#0") & " " & ADR.Fields("chkbnk1")
    End If
    If ADR.Fields("chkamt2") <> 0 Or Not IsNull(ADR.Fields("chkamt2")) Then
        ChkAmt2 = ADR.Fields("chkamt2")
        ChkTotal = ChkTotal + ChkAmt2
        StrChk2 = ADR.Fields("chkno2") & " " & Format(ADR.Fields("chkamt2"), "#########.#0") & " " & ADR.Fields("chkbnk2")
    End If
    If ADR.Fields("chkamt3") <> 0 Or Not IsNull(ADR.Fields("chkamt3")) Then
        ChkAmt3 = ADR.Fields("chkamt3")
        ChkTotal = ChkTotal + ChkAmt3
        StrChk3 = ADR.Fields("chkno3") & " " & Format(ADR.Fields("chkamt3"), "#########.#0") & " " & ADR.Fields("chkbnk3")
    End If
    If ADR.Fields("chkamt4") <> 0 Or Not IsNull(ADR.Fields("chkamt4")) Then
        ChkAmt4 = ADR.Fields("chkamt4")
        ChkTotal = ChkTotal + ChkAmt4
        StrChk4 = ADR.Fields("chkno4") & " " & Format(ADR.Fields("chkamt4"), "#########.#0") & " " & ADR.Fields("chkbnk4")
    End If
    If ADR.Fields("chkamt5") <> 0 Or Not IsNull(ADR.Fields("chkamt5")) Then
        ChkAmt5 = ADR.Fields("chkamt5")
        ChkTotal = ChkTotal + ChkAmt5
        StrChk5 = ADR.Fields("chkno5") & " " & Format(ADR.Fields("chkamt5"), "#########.#0") & " " & ADR.Fields("chkbnk5")
    End If
    If ADR.Fields("cshamt") <> 0 Or Not IsNull(ADR.Fields("cshamt")) Then
        CashAmt = ADR.Fields("cshamt")
    End If
    ADR.Close
    Set ADR = Nothing
    For X = 1 To lRef
        DetailTl = 0
        DetailAmt = 0
        TotalAmt = 0
        sngTempAmt = 0
        strCshAmt = ".00"
        strChqAmt = ".00"
        strAdrAmt = ".00"
        blnChkno1 = False
        blnChkno2 = False
        blnChkno3 = False
        blnChkno4 = False
        blnChkno5 = False
        DE.getTotal CCRList(X).Refnum, CCRList(X).Seqnum
        Set DETTl = DE.rsgetTotal
        If CCRList(X).UGCode <> "Y" Then ' ** Not Under Guarantee
            DetailTl = DETTl.Fields("TotalAmt")
            DetailAmt = DETTl.Fields("TotalAmt")
            TotalAmt = DETTl.Fields("totalamt")
            '   ** Liquidation of ADR
            If AdrAmt <> 0 Then
                If DetailTl > AdrAmt Then
                    AdrAmount = AdrAmt
                    DetailTl = DetailTl - AdrAmt
                    sngTempAmt = sngTempAmt + AdrAmount
                    AdrAmt = 0
                    strAdrAmt = Format(AdrAmount, "###,###.00")
                Else
                    AdrAmount = DetailTl
                    sngTempAmt = sngTempAmt + AdrAmount
                    AdrAmt = AdrAmt - DetailTl
                    DetailTl = 0
                    DetailAmt = 0
                    strAdrAmt = Format(AdrAmount, "###,###.00")
                End If
            End If
            If sngTempAmt = TotalAmt Then
                GoTo NextCCRTag
            End If
            '   ** Liquidation of CHEQUES
            If ChkTotal <> 0 Then
                DetailAmt = DetailTl
                If DetailTl > ChkTotal Then
                    ChkAmount = ChkTotal
                    DetailTl = DetailTl - ChkTotal
                    sngTempAmt = sngTempAmt + ChkAmount
                    strChqAmt = Format(ChkAmount, "###,###.00")
                    ChkTotal = 0
                    ChkAmount = 0
                Else
                    ChkAmount = DetailTl
                    sngTempAmt = sngTempAmt + ChkAmount
                    DetailTl = ChkTotal - DetailTl
                    ChkTotal = DetailTl
                    strChqAmt = Format(ChkAmount, "###,###.00")
                    DetailTl = 0
                    ChkAmount = 0
                End If
            End If
            '   ** Cheque 1
            If ChkAmt1 <> 0 And DetailAmt <> 0 Then
                If DetailAmt > ChkAmt1 Then
                    ChkAmount1 = ChkAmt1
                    DetailAmt = DetailAmt - ChkAmt1
                    sngTempAmt = sngTempAmt + ChkAmount1
                    blnChkno1 = True
                    ChkAmt1 = 0
                    ChkAmount1 = 0
                Else
                    ChkAmount1 = DetailTl
                    sngTempAmt = sngTempAmt + ChkAmount1
                    DetailAmt = ChkAmt1 - DetailAmt
                    ChkAmt1 = DetailAmt
                    blnChkno1 = True
                    DetailAmt = 0
                    ChkAmount1 = 0
                End If
                If sngTempAmt = TotalAmt Then
                    GoTo NextCCRTag
                End If
            End If
            '   ** Cheque 2
            If ChkAmt2 <> 0 And DetailAmt <> 0 Then
                If DetailAmt > ChkAmt2 Then
                    ChkAmount2 = ChkAmt2
                    DetailAmt = DetailAmt - ChkAmt2
                    sngTempAmt = sngTempAmt + ChkAmount2
                    blnChkno2 = True
                    ChkAmt2 = 0
                    ChkAmount2 = 0
                Else
                    ChkAmount2 = DetailAmt
                    sngTempAmt = sngTempAmt + ChkAmount2
                    DetailAmt = ChkAmt2 - DetailAmt
                    ChkAmt2 = DetailAmt
                    blnChkno2 = True
                    DetailAmt = 0
                    ChkAmount2 = 0
                End If
                If sngTempAmt = TotalAmt Then
                    GoTo NextCCRTag
                End If
            End If
            '   ** Cheque 3
            If ChkAmt3 <> 0 And DetailAmt <> 0 Then
                If DetailAmt > ChkAmt3 Then
                    ChkAmount1 = ChkAmt3
                    DetailAmt = DetailAmt - ChkAmt3
                    sngTempAmt = sngTempAmt + ChkAmount3
                    blnChkno3 = True
                    ChkAmt3 = 0
                    ChkAmount3 = 0
                Else
                    ChkAmount3 = DetailAmt
                    sngTempAmt = sngTempAmt + ChkAmount3
                    DetailAmt = ChkAmt3 - DetailAmt
                    ChkAmt3 = DetailAmt
                    blnChkno3 = True
                    DetailAmt = 0
                    ChkAmount3 = 0
                End If
                If sngTempAmt = TotalAmt Then
                    GoTo NextCCRTag
                End If
            End If
            '   ** Cheque 4
            If ChkAmt4 <> 0 And DetailAmt <> 0 Then
                If DetailAmt > ChkAmt4 Then
                    ChkAmount4 = ChkAmt4
                    DetailAmt = DetailAmt - ChkAmt4
                    sngTempAmt = sngTempAmt + ChkAmount4
                    blnChkno4 = True
                    ChkAmt4 = 0
                    ChkAmount4 = 0
                Else
                    ChkAmount4 = DetailAmt
                    sngTempAmt = sngTempAmt + ChkAmount4
                    DetailAmt = ChkAmt4 - DetailAmt
                    ChkAmt4 = DetailAmt
                    DetailAmt = 0
                    blnChkno4 = True
                    ChkAmount4 = 0
                End If
                If sngTempAmt = TotalAmt Then
                    GoTo NextCCRTag
                End If
            End If
            '   ** Cheque 5
            If ChkAmt5 <> 0 And DetailAmt <> 0 Then
                If DetailAmt > ChkAmt5 Then
                    ChkAmount5 = ChkAmt5
                    DetailAmt = DetailAmt - ChkAmt5
                    sngTempAmt = sngTempAmt + ChkAmount5
                    blnChkno5 = True
                    ChkAmt5 = 0
                    ChkAmount5 = 0
                Else
                    ChkAmount5 = DetailAmt
                    sngTempAmt = sngTempAmt + ChkAmount5
                    DetailAmt = ChkAmt5 - DetailAmt
                    ChkAmt5 = DetailAmt
                    blnChkno5 = True
                    DetailAmt = 0
                    ChkAmount5 = 0
                End If
                If sngTempAmt = TotalAmt Then
                    GoTo NextCCRTag
                End If
            End If
            '   ** Cash Amount
            If CashAmt <> 0 Or (ChkTotal = 0 And AdrAmt = 0) Then
                If DetailTl > CashAmt Then
                    CashAmount = CashAmt
                    strCshAmt = Format(CashAmount, "###,###.00")
                    CashAmt = 0
                Else
                    CashAmount = DetailTl
                    CashAmt = CashAmt - DetailTl
                    strCshAmt = Format(CashAmount, "###,###.00")
                    DetailTl = 0
                End If
            End If
        End If
NextCCRTag:
        If mvarCCRNumber <> 0 Then
            If CCRList(X).CCRnum = mvarCCRNumber Then
                Call OutCCRPC(CCRList(X).Refnum, CCRList(X).Seqnum, _
                    CCRList(X).Cusnam, strAdrAmt, strCshAmt, strChqAmt, _
                    blnChkno1, blnChkno2, blnChkno3, blnChkno4, blnChkno5)
            End If
        Else
            Call OutCCRPC(CCRList(X).Refnum, CCRList(X).Seqnum, _
                CCRList(X).Cusnam & "", strAdrAmt, strCshAmt, strChqAmt, _
                blnChkno1, blnChkno2, blnChkno3, blnChkno4, blnChkno5)
        End If
        DETTl.Close
        Set DETTl = Nothing
    Next
End Function



''--ella's version
'Private Sub OutCCRPC(pRefnum As Long, pSeqnum As Long, pCustomer As String, pAdrAmt As String, pCashAmt As String, _
'                                 pChqAmt As String, pChkno1 As Boolean, pChkno2 As Boolean, pChkno3 As Boolean, pChkno4 As Boolean, _
'                                 pChkno5 As Boolean)
'' *************************
'' ** Printing of receipt **
'' *************************
'Dim ctrCnt As Integer
'Dim tmp1 As String * 30
'Dim tmp2 As String * 30
'Dim tmpString As String
'Dim Word1 As String * 36
'Dim Word2 As String * 36
'Dim Word3 As String * 36
'Dim strEntry As String * 80
'Dim Refn As String * 10
'Dim Seqf As String * 10
'Dim CCRf As String * 10
'Dim DateTime As String
'Dim strExporter As String * 30
'Dim strSize As String * 4
'Dim strCtnnum As String * 12
'Dim strArrastre As String * 12
'Dim strWArrastre As String * 12
'Dim strWharfage As String * 12
'Dim strTArrastre As String * 12
'Dim strWgh As String * 12
'Dim strWeighing As String * 9
'Dim strTWArrastre As String * 12
'Dim strTWharfage As String * 12
'Dim sngArrastre As Currency
'Dim sngWArrastre  As Currency
'Dim sngWharfage As Currency
'Dim sngTArrastre As Currency
'Dim sngTWArrastre As Currency
'Dim sngTWharfage As Currency
'Dim sngWgh As Currency
'Dim sngWeighing As Currency
'Dim sngVat As Currency
'Dim sngWtx As Currency
'Dim vslName As String * 10
'Dim X As Integer
'Dim strRemarks As String * 15
'Dim strRemarks30 As String * 36
'Dim CD As ADODB.Recordset
'Dim WhfRteAmt As Currency
'Dim strRemarkOut As String * 16
'Dim remark1 As String
'Dim remark2  As String
'Dim remark3 As String
'Dim UserName As String
'Dim strValidation As String * 35
'Dim RevTonnage As String * 16
'ctrCnt = 11
'On Error Resume Next
'WhfRteAmt = 0
'Set CD = New ADODB.Recordset
'CD.Open "SELECT * From ccrcyx WHERE refnum = " & Trim(CStr(pRefnum)) & "" _
'        & " AND seqnum = " & Trim(CStr(pSeqnum)) & " order by itmnum", _
'        gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
'If CD.BOF <> True And CD.EOF <> True Then
'    With CD
'        Select Case .Fields("vatcde")
'        Case "1"
'            remark1 = "0Vat"
'        Case "2"
'            remark1 = "1%Less"
'        Case "3"
'            remark1 = "10%Vat"
'        Case "4"
'            remark1 = "1%Less"
'        Case "5"
'            remark1 = "6%Vat"
'        Case "6"
'            remark1 = "1%Less"
'        Case Else
'            remark1 = " "
'        End Select
'        Select Case .Fields("whfcde")
'        Case "1"
'            remark2 = "BOI"
'        Case "2"
'            remark2 = "PEZA"
'        Case "3"
'            remark2 = "Napocor"
'        Case "4"
'            remark2 = "WhfPd."
'        Case "5"
'            remark2 = "PPC"
'        Case "6"
'            remark2 = "ShutOut"
'        Case Else
'            remark2 = " "
'        End Select
'        If .Fields("guarntycde") = "Y" Then
'            remark3 = "U/G"
'        Else
'            remark3 = " "
'        End If
'        UserName = Trim(UCase(.Fields("userid") & "")) & Space(8) & Trim(UCase(.Fields("supvsr") & ""))
'        Refn = Format(.Fields("refnum"), "000000")
'        Seqf = Trim(Format(.Fields("seqnum"), "0000"))
'        CCRf = Format(.Fields("ccrnum"), "000000")
'        DateTime = Format(.Fields("sysdttm"), "     YYYY-MM-DD hh:nn")
'        strExporter = Mid(.Fields("broker"), 1, 15) & "/" & Mid(.Fields("exprtr"), 1, 15)
'        strRemarks = Mid(.Fields("remark") & "", 1, 15)
'        strRemarkOut = Trim(Mid(remark1, 1, 6)) & Trim(Mid(remark2, 1, 7)) & Trim(Mid(remark3, 1, 3))
'        strEntry = .Fields("entnum")
'        strValidation = Trim(Refn) & " " & Trim(Seqf) & " " & Trim(CCRf) & " " & Format(.Fields("sysdttm"), "YY-MM-DD hh:nn")
'        vslName = .Fields("vslcde") & ""
'        Printer.Font = "Courier 12cpi"
'        Printer.FontSize = 10
'        Printer.Print " "
'        'Printer.Print Space(23) & "REF " & Refn & " SEQ " & Seqf & Space(2) & CCRf
'        'Printer.Print " "
'        'Printer.Print " "
'        Printer.Print Space(74) & DateTime
'        Printer.Print " "
'        Printer.Print " "
'        Printer.Print Space(4) & strExporter & Space(13) & vslName & Space(3) & _
'                        Trim(Mid(strEntry, 1, 8) & " " & _
'                        Mid(strEntry, 9, 8) & " " & _
'                        Mid(strEntry, 17, 8) & " " & _
'                        Mid(strEntry, 25, 8) & " " & _
'                        Mid(strEntry, 33, 8))
'        Printer.Print Space(45) & Trim(Mid(strEntry, 41, 8) & " " & _
'                        Mid(strEntry, 49, 8) & " " & _
'                        Mid(strEntry, 57, 8) & " " & _
'                        Mid(strEntry, 65, 8) & " " & _
'                        Mid(strEntry, 73, 8))
'        Printer.Print " "
'        Printer.Print " "
'        Printer.Print " "
'        Printer.Print Space(2) & .Fields("commod")
'        If CSng(.Fields("dolrte")) > 0 Then
'            WhfRteAmt = CSng(.Fields("dolrte"))
'        End If
'        sngVat = 0
'        sngWtx = 0
'        Do While Not .EOF
'            If .Fields("vatcde") <> "1" Then
'                sngVat = sngVat + CSng(.Fields("arrvat"))
'            End If
'            If .Fields("vatcde") = "2" Or .Fields("vatcde") = "4" Or .Fields("vatcde") = "6" Then
'                sngWtx = sngWtx + CSng(.Fields("arrtax"))
'            End If
'            sngArrastre = CSng(.Fields("arramt")) + CSng(.Fields("ovzamt")) + CSng(.Fields("dgramt")) + CSng(.Fields("arrvat")) - CSng(.Fields("arrtax"))
'            sngWArrastre = CSng(.Fields("arramt")) + CSng(.Fields("ovzamt")) + CSng(.Fields("dgramt"))
'            sngWgh = CSng(.Fields("wghamt"))
'            sngTArrastre = sngTArrastre + sngArrastre
'            sngTWArrastre = sngTWArrastre + sngWArrastre
'            sngWeighing = sngWeighing + sngWgh
'            If .Fields("whfcde") = "0" Then
'                sngWharfage = CSng(.Fields("whfamt"))
'            Else
'                sngWharfage = 0
'            End If
'            sngTWharfage = sngTWharfage + sngWharfage
'            If sngArrastre > 0 Then
'                strArrastre = Format(sngArrastre, "###,###,###.#0")
'            Else
'                strArrastre = " "
'            End If
'            If sngWArrastre > 0 Then
'                strWArrastre = Format(sngWArrastre, "###,###,###.#0")
'            Else
'                strWArrastre = " "
'            End If
'            If sngWharfage > 0 Then
'                strWharfage = Format(sngWharfage, "###,###,###.#0")
'            Else
'                strWharfage = " "
'            End If
'
'            If sngWgh > 0 Then
'                strWgh = Format(sngWgh, "##,###.#0")
'            Else
'                strWgh = 0
'            End If
'            strSize = .Fields("cntsze")
'            strCtnnum = .Fields("cntnum")
'            If CSng(.Fields("ovzamt")) > 0 Then
'                RevTonnage = Format(CSng(.Fields("revton")), "###,###,###.#0")
'            Else
'                RevTonnage = ""
'            End If
'            Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & RevTonnage & Space(2) & strArrastre & Space(8) & strWgh & Space(8) & (CDbl(strArrastre) + CDbl(strWgh))
'            'Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & RevTonnage & Space(2) & strWeighing & Space(29) & strWeighing
'            ctrCnt = ctrCnt - 1
'            .MoveNext
'        Loop
'        If ctrCnt > 0 Then
'            For X = 1 To ctrCnt
'                Printer.Print " "
'            Next
'        End If
'        If sngTArrastre > 0 Then
'            strTArrastre = Format(sngTArrastre, "###,###,###.#0")
'        Else
'            strTArrastre = " "
'        End If
'        If sngTWArrastre > 0 Then
'            strTWArrastre = Format(sngTWArrastre, "###,###,###.#0")
'        Else
'            strTWArrastre = " "
'        End If
'        If sngTWharfage > 0 Then
'            strTWharfage = Format(sngTWharfage, "###,###,###.#0")
'        Else
'            strTWharfage = " "
'        End If
'        If sngWeighing > 0 Then
'            strWeighing = Format(sngWeighing, "###,###,###.#0")
'        Else
'            strWeighing = " "
'        End If
'
'        If sngVat > 0 Then
'            If sngWtx > 0 Then
'                tmpString = "VAT INCLUSIVE LESS W/TAX"
'            Else
'                tmpString = "VAT INCLUSIVE"
'            End If
'        Else
'            tmpString = "ZERO RATED VAT"
'        End If
'        'Printer.Print " "  ' 3
'        Printer.Print " "  ' 3
'        Printer.Print " "  ' 3
'        Printer.Print Space(6) & Space(15) & Space(17) & Space(17) & Trim(tmpString) & Space(10) & (CDbl(strTArrastre) + CDbl(strWeighing))
'        tmpString = NumToText(CCur(sngTArrastre))
'        Word1 = Mid(tmpString, 1, 35)
'        If Len(Trim(Mid(tmpString, 35, 1))) <> 0 And Len(Trim(Mid(tmpString, 36, 1))) <> 0 Then
'            Word1 = Trim(Word1) & "-"
'        End If
'        Word2 = Mid(tmpString, 36, 35)
'        If Len(Trim(Mid(tmpString, 71, 1))) <> 0 And Len(Trim(Mid(tmpString, 72, 1))) <> 0 Then
'            Word2 = Trim(Word2) & "-"
'        End If
'        Word3 = Mid(tmpString, 72, 35)
'        Printer.Print " "
'        Printer.Print " "
'        Printer.Print Space(46) & Word1
'        Printer.Print Space(2) & strRemarks & Space(6) & strRemarkOut & Space(7) & Word2
'
'        If DomesticMode Then
'            Printer.Print "  DOMESTIC" & Space(36) & Word3
'        Else
'            Printer.Print "  FOREIGN " & Space(36) & Word3
'        End If
'        Printer.Print " "
'        tmpString = strChqAmt & " CK    " & strCshAmt & " CS"
'        Printer.Print Space(44) & tmpString
'        tmpString = strAdrAmt & " AD"
'        Printer.Print Space(5) & UserName  ' & Space(26) & tmpString
'        If blnChkno1 Then
'            tmp1 = StrChk1
'        Else
'            tmp1 = " "
'        End If
'        If blnChkno2 Then
'            If Len(tmp1) > 0 Then
'                tmp2 = ", " & StrChk2
'            Else
'                tmp2 = " " & StrChk2
'            End If
'        Else
'            tmp2 = " "
'        End If
'        Printer.Print Space(44) & Trim(tmp1) & tmp2
'        If blnChkno3 Then
'            tmp1 = StrChk3
'        Else
'            tmp1 = " "
'        End If
'        If blnChkno4 Then
'            If Len(tmp1) > 0 Then
'                tmp2 = ", " & StrChk4
'            Else
'                tmp2 = " " & StrChk4
'            End If
'        Else
'            tmp2 = " "
'        End If
'        Printer.Print Space(44) & Trim(tmp1) & tmp2
'        If blnChkno5 Then
'            tmp1 = StrChk5
'        Else
'            tmp1 = " "
'        End If
'        Printer.Print Space(44) & tmp1
'        Printer.Print Space(44) & strValidation
'        Printer.Print ""
'        Printer.Print ""
'        Printer.Print Space(5) & "REF " & Refn & " SEQ " & Seqf & Space(2) & CCRf
'        Printer.FontSize = 10
'        Printer.EndDoc
'    End With
'End If
'
'CD.Close
'Set CD = Nothing
'
'End Sub

Private Sub OutCCRPC(pRefnum As Long, pSeqnum As Long, pCustomer As String, pAdrAmt As String, pCashAmt As String, _
                                 pChqAmt As String, pChkno1 As Boolean, pChkno2 As Boolean, pChkno3 As Boolean, pChkno4 As Boolean, _
                                 pChkno5 As Boolean)
' *************************
' ** Printing of receipt **
' *************************
Dim ctrCnt As Integer
Dim tmp1 As String * 30
Dim tmp2 As String * 30
Dim tmpString As String
Dim Word1 As String * 36
Dim Word2 As String * 36
Dim Word3 As String * 36
Dim strEntry As String * 80
Dim Refn As String * 10
Dim Seqf As String * 10
Dim CCRf As String * 10
Dim DateTime As String
Dim strExporter As String * 30
Dim strSize As String * 4
Dim strCtnnum As String * 12
Dim strArrastre As String * 12
Dim strWArrastre As String * 12
Dim strWharfage As String * 12
Dim strTArrastre As String * 12

Dim strWgh As String * 9
Dim strWeighing As String * 9
Dim sngWeighTotal As Currency


Dim strTWArrastre As String * 12
Dim strTWharfage As String * 12
Dim sngArrastre As Currency
Dim sngWArrastre  As Currency

Dim sngWharfage As Currency
Dim sngTArrastre As Currency
Dim sngTWArrastre As Currency
Dim sngTWharfage As Currency
Dim sngVat As Currency
Dim sngWtx As Currency
Dim sngWgh As Currency
Dim vslName As String * 10
Dim X As Integer
Dim strRemarks As String * 15
Dim strRemarks30 As String * 36
Dim CD As ADODB.Recordset
Dim WhfRteAmt As Currency
Dim strRemarkOut As String * 16
Dim remark1 As String
Dim remark2  As String
Dim remark3 As String
Dim UserName As String
Dim strValidation As String * 35
Dim RevTonnage As String * 16
ctrCnt = 11
On Error Resume Next
WhfRteAmt = 0
Set CD = New ADODB.Recordset
CD.Open "SELECT * From ccrcyx WHERE refnum = " & Trim(CStr(pRefnum)) & "" _
        & " AND seqnum = " & Trim(CStr(pSeqnum)) & " order by itmnum", _
        gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
If CD.BOF <> True And CD.EOF <> True Then
    With CD
        Select Case .Fields("vatcde")
        Case "1"
            remark1 = "0Vat"
        Case "2"
            remark1 = "1%Less"
        Case "3"
            remark1 = "10%Vat"
        Case "4"
            remark1 = "1%Less"
        Case "5"
            remark1 = "6%Vat"
        Case "6"
            remark1 = "1%Less"
        Case Else
            remark1 = " "
        End Select
        Select Case .Fields("whfcde")
        Case "1"
            remark2 = "BOI"
        Case "2"
            remark2 = "PEZA"
        Case "3"
            remark2 = "Napocor"
        Case "4"
            remark2 = "WhfPd."
        Case "5"
            remark2 = "PPC"
        Case "6"
            remark2 = "ShutOut"
        Case Else
            remark2 = " "
        End Select
        If .Fields("guarntycde") = "Y" Then
            remark3 = "U/G"
        Else
            remark3 = " "
        End If
        UserName = Trim(UCase(.Fields("userid") & "")) & Space(8) & Trim(UCase(.Fields("supvsr") & ""))
        Refn = Format(.Fields("refnum"), "000000")
        Seqf = Trim(Format(.Fields("seqnum"), "0000"))
        CCRf = Format(.Fields("ccrnum"), "000000")
        DateTime = Format(.Fields("sysdttm"), "     YYYY-MM-DD hh:nn")
        strExporter = Mid(.Fields("broker"), 1, 15) & "/" & Mid(.Fields("exprtr"), 1, 15)
        strRemarks = Mid(.Fields("remark") & "", 1, 15)
        strRemarkOut = Trim(Mid(remark1, 1, 6)) & Trim(Mid(remark2, 1, 7)) & Trim(Mid(remark3, 1, 3))
        strEntry = .Fields("entnum")
        strValidation = Trim(Refn) & " " & Trim(Seqf) & " " & Trim(CCRf) & " " & Format(.Fields("sysdttm"), "YY-MM-DD hh:nn")
        vslName = .Fields("vslcde") & ""
        Printer.Font = "Courier 12cpi"
        Printer.FontSize = 10
'        If Printer.Height > Printer.Width Then
'            Printer.Orientation = vbPRORPortrait
'        Else
'            Printer.Orientation = vbPRORLandscape
'        End If
        
        Printer.Print " "
        'Printer.Print Space(23) & "REF " & Refn & " SEQ " & Seqf & Space(2) & CCRf
        'Printer.Print " "
        'Printer.Print " "
        Printer.Print Space(74) & DateTime
        Printer.Print " "
        Printer.Print " "
        Printer.Print Space(4) & strExporter & Space(13) & vslName & Space(3) & _
                        Trim(Mid(strEntry, 1, 8) & " " & _
                        Mid(strEntry, 9, 8) & " " & _
                        Mid(strEntry, 17, 8) & " " & _
                        Mid(strEntry, 25, 8) & " " & _
                        Mid(strEntry, 33, 8))
        Printer.Print Space(45) & Trim(Mid(strEntry, 41, 8) & " " & _
                        Mid(strEntry, 49, 8) & " " & _
                        Mid(strEntry, 57, 8) & " " & _
                        Mid(strEntry, 65, 8) & " " & _
                        Mid(strEntry, 73, 8))
        Printer.Print " "
        Printer.Print " "
        Printer.Print " "
        Printer.Print Space(2) & .Fields("commod")


        If CSng(.Fields("dolrte")) > 0 Then
            WhfRteAmt = CSng(.Fields("dolrte"))
        End If
        sngVat = 0
        sngWtx = 0
        sngWeighTotal = 0

        Do While Not .EOF
            If .Fields("vatcde") <> "1" Then
                sngVat = sngVat + CSng(.Fields("arrvat"))
            End If
            If .Fields("vatcde") = "2" Or .Fields("vatcde") = "4" Or .Fields("vatcde") = "6" Then
                sngWtx = sngWtx + CSng(.Fields("arrtax"))
            End If
            sngArrastre = CSng(.Fields("arramt")) + CSng(.Fields("ovzamt")) + CSng(.Fields("dgramt")) + CSng(.Fields("arrvat")) - CSng(.Fields("arrtax"))
            sngWArrastre = CSng(.Fields("arramt")) + CSng(.Fields("ovzamt")) + CSng(.Fields("dgramt"))
            sngTArrastre = sngTArrastre + sngArrastre
            sngTWArrastre = sngTWArrastre + sngWArrastre
            sngWgh = CSng(.Fields("wghamt"))
            sngWeighing = sngWeighing + sngWgh
            sngWeighTotal = sngWeighTotal + sngWgh
            If .Fields("whfcde") = "0" Then
                sngWharfage = CSng(.Fields("whfamt"))
            Else
                sngWharfage = 0
            End If
            sngTWharfage = sngTWharfage + sngWharfage
            If sngArrastre > 0 Then
                strArrastre = Format(sngArrastre, "###,###,###.#0")
            Else
                strArrastre = " "
            End If
            If sngWArrastre > 0 Then
                strWArrastre = Format(sngWArrastre, "###,###,###.#0")
            Else
                strWArrastre = " "
            End If
            If sngWharfage > 0 Then
                strWharfage = Format(sngWharfage, "###,###,###.#0")
            Else
                strWharfage = " "
            End If
            If sngWgh > 0 Then
                strWgh = Format(sngWgh, "##,###.#0")
            Else
                strWgh = 0
            End If
            strSize = .Fields("cntsze")
            strCtnnum = .Fields("cntnum")
            If CSng(.Fields("ovzamt")) > 0 Then
                RevTonnage = Format(CSng(.Fields("revton")), "###,###,###.#0")
            Else
                RevTonnage = ""
            End If
            'Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & RevTonnage & Space(2) & strArrastre & Space(29) & strArrastre
            'sharon 05Nov2009 Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & RevTonnage & Space(2) & strArrastre & Space(2) & strWgh & Space(26) & Format(CDbl(strArrastre) + CDbl(strWgh), "###,###,###.#0")
            Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & RevTonnage & Space(2) & strArrastre & Space(2) & strWgh & Space(20) & Format(CDbl(strArrastre) + CDbl(strWgh), "###,###,###.#0")
            ctrCnt = ctrCnt - 1
            .MoveNext
        Loop
        If ctrCnt > 0 Then
            For X = 1 To ctrCnt
                Printer.Print " "
            Next
        End If
        If sngTArrastre > 0 Then
            strTArrastre = Format(sngTArrastre, "###,###,###.#0")
        Else
            strTArrastre = " "
        End If
        If sngTWArrastre > 0 Then
            strTWArrastre = Format(sngTWArrastre, "###,###,###.#0")
        Else
            strTWArrastre = " "
        End If
        If sngTWharfage > 0 Then
            strTWharfage = Format(sngTWharfage, "###,###,###.#0")
        Else
            strTWharfage = " "
        End If
        If sngVat > 0 Then
            If sngWtx > 0 Then
                tmpString = "VAT INCLUSIVE LESS W/TAX"
            Else
                tmpString = "VAT INCLUSIVE"
            End If
        Else
            tmpString = "ZERO RATED VAT"
        End If
        'Printer.Print " "  ' 3
        Printer.Print " "  ' 3
        Printer.Print " "  ' 3
        'Printer.Print Space(6) & Space(15) & Space(17) & Space(17) & Trim(tmpString) & Space(10) & strTArrastre
        Printer.Print Space(6) & Space(15) & Space(17) & Space(17) & Trim(tmpString) & Space(10) & Format(CDbl(strTArrastre) + CDbl(sngWeighTotal), "###,###,###.#0")
        tmpString = NumToText(CCur(sngTArrastre) + sngWeighTotal)
        Word1 = Mid(tmpString, 1, 35)
        If Len(Trim(Mid(tmpString, 35, 1))) <> 0 And Len(Trim(Mid(tmpString, 36, 1))) <> 0 Then
            Word1 = Trim(Word1) & "-"
        End If
        Word2 = Mid(tmpString, 36, 35)
        If Len(Trim(Mid(tmpString, 71, 1))) <> 0 And Len(Trim(Mid(tmpString, 72, 1))) <> 0 Then
            Word2 = Trim(Word2) & "-"
        End If
        Word3 = Mid(tmpString, 72, 35)
        Printer.Print " "
        Printer.Print " "
        Printer.Print Space(46) & Word1
        Printer.Print Space(2) & strRemarks & Space(6) & strRemarkOut & Space(7) & Word2

        If DomesticMode Then
            Printer.Print "  DOMESTIC" & Space(36) & Word3
        Else
            Printer.Print "  FOREIGN " & Space(36) & Word3
        End If
        Printer.Print " "
        tmpString = strChqAmt & " CK    " & strCshAmt & " CS"
        Printer.Print Space(44) & tmpString
        tmpString = strAdrAmt & " AD"
        Printer.Print Space(5) & UserName  ' & Space(26) & tmpString
        If blnChkno1 Then
            tmp1 = StrChk1
        Else
            tmp1 = " "
        End If
        If blnChkno2 Then
            If Len(tmp1) > 0 Then
                tmp2 = ", " & StrChk2
            Else
                tmp2 = " " & StrChk2
            End If
        Else
            tmp2 = " "
        End If
        Printer.Print Space(44) & Trim(tmp1) & tmp2
        If blnChkno3 Then
            tmp1 = StrChk3
        Else
            tmp1 = " "
        End If
        If blnChkno4 Then
            If Len(tmp1) > 0 Then
                tmp2 = ", " & StrChk4
            Else
                tmp2 = " " & StrChk4
            End If
        Else
            tmp2 = " "
        End If
        Printer.Print Space(44) & Trim(tmp1) & tmp2
        If blnChkno5 Then
            tmp1 = StrChk5
        Else
            tmp1 = " "
        End If
        Printer.Print Space(44) & tmp1
        Printer.Print Space(44) & strValidation
        Printer.Print ""
        Printer.Print ""
        Printer.Print Space(5) & "REF " & Refn & " SEQ " & Seqf & Space(2) & CCRf
        Printer.FontSize = 10
        Printer.EndDoc
    End With
End If

CD.Close
Set CD = Nothing

End Sub

'sharon orig Private Sub OutCCRPC(pRefnum As Long, pSeqnum As Long, pCustomer As String, pAdrAmt As String, pCashAmt As String, _
'                                 pChqAmt As String, pChkno1 As Boolean, pChkno2 As Boolean, pChkno3 As Boolean, pChkno4 As Boolean, _
'                                 pChkno5 As Boolean)
'' *************************
'' ** Printing of receipt **
'' *************************
'Dim ctrCnt As Integer
'Dim tmp1 As String * 30
'Dim tmp2 As String * 30
'Dim tmpString As String
'Dim Word1 As String * 36
'Dim Word2 As String * 36
'Dim Word3 As String * 36
'Dim strEntry As String * 80
'Dim Refn As String * 10
'Dim Seqf As String * 10
'Dim CCRf As String * 10
'Dim DateTime As String
'Dim strExporter As String * 30
'Dim strSize As String * 4
'Dim strCtnnum As String * 12
'Dim strArrastre As String * 12
'Dim strWArrastre As String * 12
'Dim strWharfage As String * 12
'Dim strTArrastre As String * 12
'Dim strTWArrastre As String * 12
'Dim strTWharfage As String * 12
'Dim sngArrastre As Currency
'Dim sngWArrastre  As Currency
'Dim sngWharfage As Currency
'Dim sngTArrastre As Currency
'Dim sngTWArrastre As Currency
'Dim sngTWharfage As Currency
'Dim sngVat As Currency
'Dim sngWtx As Currency
'Dim vslName As String * 10
'Dim X As Integer
'Dim strRemarks As String * 15
'Dim strRemarks30 As String * 36
'Dim CD As ADODB.Recordset
'Dim WhfRteAmt As Currency
'Dim strRemarkOut As String * 16
'Dim remark1 As String
'Dim remark2  As String
'Dim remark3 As String
'Dim UserName As String
'Dim strValidation As String * 35
'Dim RevTonnage As String * 16
'ctrCnt = 11
'On Error Resume Next
'WhfRteAmt = 0
'Set CD = New ADODB.Recordset
'CD.Open "SELECT * From ccrcyx WHERE refnum = " & Trim(CStr(pRefnum)) & "" _
'        & " AND seqnum = " & Trim(CStr(pSeqnum)) & " order by itmnum", _
'        gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
'If CD.BOF <> True And CD.EOF <> True Then
'    With CD
'        Select Case .Fields("vatcde")
'        Case "1"
'            remark1 = "0Vat"
'        Case "2"
'            remark1 = "1%Less"
'        Case "3"
'            remark1 = "10%Vat"
'        Case "4"
'            remark1 = "1%Less"
'        Case "5"
'            remark1 = "6%Vat"
'        Case "6"
'            remark1 = "1%Less"
'        Case Else
'            remark1 = " "
'        End Select
'        Select Case .Fields("whfcde")
'        Case "1"
'            remark2 = "BOI"
'        Case "2"
'            remark2 = "PEZA"
'        Case "3"
'            remark2 = "Napocor"
'        Case "4"
'            remark2 = "WhfPd."
'        Case "5"
'            remark2 = "PPC"
'        Case "6"
'            remark2 = "ShutOut"
'        Case Else
'            remark2 = " "
'        End Select
'        If .Fields("guarntycde") = "Y" Then
'            remark3 = "U/G"
'        Else
'            remark3 = " "
'        End If
'        UserName = Trim(UCase(.Fields("userid") & "")) & Space(8) & Trim(UCase(.Fields("supvsr") & ""))
'        Refn = Format(.Fields("refnum"), "000000")
'        Seqf = Trim(Format(.Fields("seqnum"), "0000"))
'        CCRf = Format(.Fields("ccrnum"), "000000")
'        DateTime = Format(.Fields("sysdttm"), "     YYYY-MM-DD hh:nn")
'        strExporter = Mid(.Fields("broker"), 1, 15) & "/" & Mid(.Fields("exprtr"), 1, 15)
'        strRemarks = Mid(.Fields("remark") & "", 1, 15)
'        strRemarkOut = Trim(Mid(remark1, 1, 6)) & Trim(Mid(remark2, 1, 7)) & Trim(Mid(remark3, 1, 3))
'        strEntry = .Fields("entnum")
'        strValidation = Trim(Refn) & " " & Trim(Seqf) & " " & Trim(CCRf) & " " & Format(.Fields("sysdttm"), "YY-MM-DD hh:nn")
'        vslName = .Fields("vslcde") & ""
'        Printer.Font = "Courier 12cpi"
'        Printer.FontSize = 10
'        Printer.Print " "
'        'Printer.Print Space(23) & "REF " & Refn & " SEQ " & Seqf & Space(2) & CCRf
'        'Printer.Print " "
'        'Printer.Print " "
'        Printer.Print Space(74) & DateTime
'        Printer.Print " "
'        Printer.Print " "
'        Printer.Print Space(4) & strExporter & Space(13) & vslName & Space(3) & _
'                        Trim(Mid(strEntry, 1, 8) & " " & _
'                        Mid(strEntry, 9, 8) & " " & _
'                        Mid(strEntry, 17, 8) & " " & _
'                        Mid(strEntry, 25, 8) & " " & _
'                        Mid(strEntry, 33, 8))
'        Printer.Print Space(45) & Trim(Mid(strEntry, 41, 8) & " " & _
'                        Mid(strEntry, 49, 8) & " " & _
'                        Mid(strEntry, 57, 8) & " " & _
'                        Mid(strEntry, 65, 8) & " " & _
'                        Mid(strEntry, 73, 8))
'        Printer.Print " "
'        Printer.Print " "
'        Printer.Print " "
'        Printer.Print Space(2) & .Fields("commod")
'        If CSng(.Fields("dolrte")) > 0 Then
'            WhfRteAmt = CSng(.Fields("dolrte"))
'        End If
'        sngVat = 0
'        sngWtx = 0
'        Do While Not .EOF
'            If .Fields("vatcde") <> "1" Then
'                sngVat = sngVat + CSng(.Fields("arrvat"))
'            End If
'            If .Fields("vatcde") = "2" Or .Fields("vatcde") = "4" Or .Fields("vatcde") = "6" Then
'                sngWtx = sngWtx + CSng(.Fields("arrtax"))
'            End If
'            sngArrastre = CSng(.Fields("arramt")) + CSng(.Fields("ovzamt")) + CSng(.Fields("dgramt")) + CSng(.Fields("arrvat")) - CSng(.Fields("arrtax"))
'            sngWArrastre = CSng(.Fields("arramt")) + CSng(.Fields("ovzamt")) + CSng(.Fields("dgramt"))
'            sngTArrastre = sngTArrastre + sngArrastre
'            sngTWArrastre = sngTWArrastre + sngWArrastre
'            If .Fields("whfcde") = "0" Then
'                sngWharfage = CSng(.Fields("whfamt"))
'            Else
'                sngWharfage = 0
'            End If
'            sngTWharfage = sngTWharfage + sngWharfage
'            If sngArrastre > 0 Then
'                strArrastre = Format(sngArrastre, "###,###,###.#0")
'            Else
'                strArrastre = " "
'            End If
'            If sngWArrastre > 0 Then
'                strWArrastre = Format(sngWArrastre, "###,###,###.#0")
'            Else
'                strWArrastre = " "
'            End If
'            If sngWharfage > 0 Then
'                strWharfage = Format(sngWharfage, "###,###,###.#0")
'            Else
'                strWharfage = " "
'            End If
'            strSize = .Fields("cntsze")
'            strCtnnum = .Fields("cntnum")
'            If CSng(.Fields("ovzamt")) > 0 Then
'                RevTonnage = Format(CSng(.Fields("revton")), "###,###,###.#0")
'            Else
'                RevTonnage = ""
'            End If
'            Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & RevTonnage & Space(2) & strArrastre & Space(29) & strArrastre
'            ctrCnt = ctrCnt - 1
'            .MoveNext
'        Loop
'        If ctrCnt > 0 Then
'            For X = 1 To ctrCnt
'                Printer.Print " "
'            Next
'        End If
'        If sngTArrastre > 0 Then
'            strTArrastre = Format(sngTArrastre, "###,###,###.#0")
'        Else
'            strTArrastre = " "
'        End If
'        If sngTWArrastre > 0 Then
'            strTWArrastre = Format(sngTWArrastre, "###,###,###.#0")
'        Else
'            strTWArrastre = " "
'        End If
'        If sngTWharfage > 0 Then
'            strTWharfage = Format(sngTWharfage, "###,###,###.#0")
'        Else
'            strTWharfage = " "
'        End If
'        If sngVat > 0 Then
'            If sngWtx > 0 Then
'                tmpString = "VAT INCLUSIVE LESS W/TAX"
'            Else
'                tmpString = "VAT INCLUSIVE"
'            End If
'        Else
'            tmpString = "ZERO RATED VAT"
'        End If
'        'Printer.Print " "  ' 3
'        Printer.Print " "  ' 3
'        Printer.Print " "  ' 3
'        Printer.Print Space(6) & Space(15) & Space(17) & Space(17) & Trim(tmpString) & Space(10) & strTArrastre
'        tmpString = NumToText(CCur(sngTArrastre))
'        Word1 = Mid(tmpString, 1, 35)
'        If Len(Trim(Mid(tmpString, 35, 1))) <> 0 And Len(Trim(Mid(tmpString, 36, 1))) <> 0 Then
'            Word1 = Trim(Word1) & "-"
'        End If
'        Word2 = Mid(tmpString, 36, 35)
'        If Len(Trim(Mid(tmpString, 71, 1))) <> 0 And Len(Trim(Mid(tmpString, 72, 1))) <> 0 Then
'            Word2 = Trim(Word2) & "-"
'        End If
'        Word3 = Mid(tmpString, 72, 35)
'        Printer.Print " "
'        Printer.Print " "
'        Printer.Print Space(46) & Word1
'        Printer.Print Space(2) & strRemarks & Space(6) & strRemarkOut & Space(7) & Word2
'
'        If DomesticMode Then
'            Printer.Print "  DOMESTIC" & Space(36) & Word3
'        Else
'            Printer.Print "  FOREIGN " & Space(36) & Word3
'        End If
'        Printer.Print " "
'        tmpString = strChqAmt & " CK    " & strCshAmt & " CS"
'        Printer.Print Space(44) & tmpString
'        tmpString = strAdrAmt & " AD"
'        Printer.Print Space(5) & UserName  ' & Space(26) & tmpString
'        If blnChkno1 Then
'            tmp1 = StrChk1
'        Else
'            tmp1 = " "
'        End If
'        If blnChkno2 Then
'            If Len(tmp1) > 0 Then
'                tmp2 = ", " & StrChk2
'            Else
'                tmp2 = " " & StrChk2
'            End If
'        Else
'            tmp2 = " "
'        End If
'        Printer.Print Space(44) & Trim(tmp1) & tmp2
'        If blnChkno3 Then
'            tmp1 = StrChk3
'        Else
'            tmp1 = " "
'        End If
'        If blnChkno4 Then
'            If Len(tmp1) > 0 Then
'                tmp2 = ", " & StrChk4
'            Else
'                tmp2 = " " & StrChk4
'            End If
'        Else
'            tmp2 = " "
'        End If
'        Printer.Print Space(44) & Trim(tmp1) & tmp2
'        If blnChkno5 Then
'            tmp1 = StrChk5
'        Else
'            tmp1 = " "
'        End If
'        Printer.Print Space(44) & tmp1
'        Printer.Print Space(44) & strValidation
'        Printer.Print ""
'        Printer.Print ""
'        Printer.Print Space(5) & "REF " & Refn & " SEQ " & Seqf & Space(2) & CCRf
'        Printer.FontSize = 10
'        Printer.EndDoc
'    End With
'End If
'
'CD.Close
'Set CD = Nothing
'
'End Sub

Private Function NumToText(dblValue As Currency) As String
    Static ones(0 To 9) As String
    Static teens(0 To 9) As String
    Static tens(0 To 9) As String
    Static thousands(0 To 4) As String
    Dim i As Integer, nPosition As Integer
    Dim nDigit As Integer, bAllZeros As Integer
    Dim strResult As String, strTemp As String
    Dim tmpBuff As String
    Dim strSign As String
    Dim negativeSign As Boolean

    ones(0) = "zero"
    ones(1) = "one"
    ones(2) = "two"
    ones(3) = "three"
    ones(4) = "four"
    ones(5) = "five"
    ones(6) = "six"
    ones(7) = "seven"
    ones(8) = "eight"
    ones(9) = "nine"

    teens(0) = "ten"
    teens(1) = "eleven"
    teens(2) = "twelve"
    teens(3) = "thirteen"
    teens(4) = "fourteen"
    teens(5) = "fifteen"
    teens(6) = "sixteen"
    teens(7) = "seventeen"
    teens(8) = "eighteen"
    teens(9) = "nineteen"

    tens(0) = ""
    tens(1) = "ten"
    tens(2) = "twenty"
    tens(3) = "thirty"
    tens(4) = "forty"
    tens(5) = "fifty"
    tens(6) = "sixty"
    tens(7) = "seventy"
    tens(8) = "eighty"
    tens(9) = "ninety"
    
    thousands(0) = ""
    thousands(1) = "thousand"
    thousands(2) = "million"
    thousands(3) = "billion"
    thousands(4) = "trillion"

    'Trap errors
    On Error GoTo NumToTextError
    'Get fractional part
    If dblValue < 0 Then
        negativeSign = True
        dblValue = Abs(dblValue)
    Else
        negativeSign = False
    End If
    strResult = "& " & Format((dblValue - Int(dblValue)) * 100, "00") & "/100"
    If negativeSign Then
        strSign = "NEGATIVE "
    Else
        strSign = ""
    End If
    strTemp = CStr(Int(dblValue))
    'Iterate through string
    For i = Len(strTemp) To 1 Step -1
        'Get value of this digit
        nDigit = Val(Mid$(strTemp, i, 1))
        'Get column position
        nPosition = (Len(strTemp) - i) + 1
        'Action depends on 1's, 10's or 100's column
        Select Case (nPosition Mod 3)
            Case 1  '1's position
                bAllZeros = False
                If i = 1 Then
                    tmpBuff = ones(nDigit) & " "
                ElseIf Mid$(strTemp, i - 1, 1) = "1" Then
                    tmpBuff = teens(nDigit) & " "
                    i = i - 1   'Skip tens position
                ElseIf nDigit > 0 Then
                    tmpBuff = ones(nDigit) & " "
                Else
                    'If next 10s & 100s columns are also
                    'zero, then don't show 'thousands'
                    bAllZeros = True
                    If i > 1 Then
                        If Mid$(strTemp, i - 1, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    If i > 2 Then
                        If Mid$(strTemp, i - 2, 1) <> "0" Then
                            bAllZeros = False
                        End If
                    End If
                    tmpBuff = ""
                End If
                If bAllZeros = False And nPosition > 1 Then
                    tmpBuff = tmpBuff & thousands(nPosition / 3) & " "
                End If
                strResult = tmpBuff & strResult
            Case 2  'Tens position
                If nDigit > 0 Then
                    strResult = tens(nDigit) & " " & strResult
                End If
            Case 0  'Hundreds position
                If nDigit > 0 Then
                    strResult = ones(nDigit) & " hundred " & strResult
                End If
        End Select
    Next i
    'Convert first letter to upper case
    If Len(strResult) > 0 Then
        strResult = UCase$(Left$(strResult, 1)) & Mid$(strResult, 2)
    End If

EndNumToText:
    'Return result
    NumToText = Trim(strSign) & strResult
    Exit Function

NumToTextError:
    strResult = "#Error#"
    Resume EndNumToText
End Function
Private Sub ConvertSize(ByRef TLength, ByRef TWidth, _
    ByRef THeight, TUms, TLng As Single, TWdth As Single, THght As Single)
 'COMPUTATION FOR OVERSIZE AMOUNTS
If (TLng > 0) And (TWdth > 0) And (THght > 0) Then
    If TUms <> "C" Then
        TLength = Round((TLng / 2.54), 2)
        TWidth = Round((TWdth / 2.54), 2)
        THeight = Round((THght / 2.54), 2)
    Else
        TLength = Round(TLng, 2)
        TWidth = Round(TWdth, 2)
        THeight = Round(THght, 2)
    End If
End If
End Sub

'Public Sub ReadConfig()
'Dim Xcnt As Integer
'Open App.Path & "\" & "Conn.cfg" For Binary Access Read As #1
'
'Do While Not EOF(1)
'    Xcnt = Xcnt + 1
'    Select Case Xcnt
'        Case 1
'            Line Input #1, sqlConBilling
'        Case 2
'            Line Input #1, sqlConNavis
'    End Select
'Loop
'End Sub

Public Function ConnectToNavis() As Boolean '(ByVal pCnnStr As String) As Boolean
Dim errBilling As ADODB.Error
Dim lsErrStr As String
   'sharon
    ' Open the database.
    On Error GoTo err_Connect
    Set gcnnNavis = New ADODB.Connection
    'Call ReadConfig
    gcnnNavis.Open sqlConNavis
    
    gcnnNavis.Open "Provider=sqloledb" & _
        ";Data Source=sbitc-db" & _
        ";Initial Catalog=apex" & _
        ";User ID=tosadmin;Password=tosadmin"

'    gcnnNavis.Open "Provider=sqloledb" & _
'        ";Data Source=sbitc-dev" & _
'        ";Initial Catalog=apex" & _
'        ";User ID=tosadmin;Password=password"
    
    '";Integrated Security=SSPI"
    gbNavis = True
    ConnectToNavis = True
   
    Exit Function
    
err_Connect:
    ConnectToNavis = False: gbConnected = False
    For Each errBilling In gcnnNavis.Errors
        With errBilling
            lsErrStr = "Connection Error. " & .Description & vbLf & _
            "Verify Log On then retry."
        End With
        MsgBox lsErrStr, vbCritical
    Next
End Function

Private Sub GrantOOGPermission(ByVal strContNo As String, ByVal intCCRNum As String)
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    On Error Resume Next
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "upnew_GrantOOGPermission"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("ccrnum", adInteger, adParamInput, 4, intCCRNum)
        .Parameters.Append .CreateParameter("cntnum", adVarChar, adParamInput, 12, strContNo)
        .Execute
    End With
    Set cmd = Nothing
End Sub
Private Sub GrantDGPermission(ByVal strContNo As String, ByVal intCCRNum As String)
    Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    On Error Resume Next
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "upnew_GrantDGPermission"
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("ccrnum", adInteger, adParamInput, 4, intCCRNum)
        .Parameters.Append .CreateParameter("cntnum", adVarChar, adParamInput, 12, strContNo)
        .Execute
    End With
    Set cmd = Nothing
End Sub

'PRNH - Query for checking and extracting container number and its details in NAVIS
Private Sub CheckContainerInNAVIS(ByVal cntNum As String)
    Dim rstExp As ADODB.Recordset
    Dim strQuery As String
        
    ConnectToNavis
    
    
    On Error GoTo err
    Set rstExp = New ADODB.Recordset
    
    With rstExp
        strQuery = "SET NOCOUNT ON; select a.id,  CAST(ROUND(d.length_mm * 0.00328083, 0) AS INTEGER) AS contsize," & _
            "a.freight_kind , e.flex_string02, a.category " & _
            "from inv_unit a " & _
            "inner join inv_unit_fcy_visit b on b.unit_gkey=a.gkey " & _
            "inner join argo_carrier_visit c on c.gkey=b.actual_ob_cv " & _
            "INNER JOIN ref_equipment d ON a.id = d.id_full " & _
            "INNER JOIN vsl_vessel_visit_details e on c.cvcvd_gkey = e.vvd_gkey " & _
            "where a.category in ('EXPRT','TRNSHP') and a.visit_state='1ACTIVE' and " & _
            "a.id = '" & cntNum & "'"

        .Open strQuery, gcnnNavis, adOpenForwardOnly, adLockReadOnly
            
        If Not .BOF = True Or Not .EOF = True Then
            'utxtSze.Value = .Fields("contsize")
            'utxtFEmp.Text = IIf(.Fields("freight_kind") = "FCL", "F", "E")
            lblCompCode.Caption = IIf(IsNull(.Fields("flex_string02")), "", .Fields("flex_string02"))
            
            If Trim(lblCompCode.Caption) = "" Then MsgBox "Company Code not Indicated. Verify with Operations"
        Else
        
            MsgBox "Container number does not exist in NAVIS. Verify with Operations"
            EnableForNewContainer (True)
            utxtPref.Text = ""
            utxtNo.Text = ""
            utxtPref.SetFocus
        End If
    End With
    
    
Exit Sub
err:

MsgBox "Error retrieving container details. Error message: " & err.Description
End Sub


Private Sub GetSparcsN4Host()
    
    On Error GoTo err
    Dim rstSparcsN4Host As ADODB.Recordset
    Dim strSparcsN4Host As String
    
    Set rstSparcsN4Host = New ADODB.Recordset
    
    strSparcsN4Host = "SELECT * " & _
                       "FROM SparcsN4Host " & _
                       "WHERE status='ACT'"

    rstSparcsN4Host.Open strSparcsN4Host, gcnnBilling, adOpenForwardOnly, adLockReadOnly
    
    If rstSparcsN4Host.BOF Then
        'MsgBox "WARNING:B.L. " & Trim(strBillNo) & " not found!", vbCritical, "Cargo Manifest"
        'MsgBox "Please inform your SUPERVISOR!", vbInformation, "Cargo Manifest"
        Exit Sub
    End If
    
    With rstSparcsN4Host
        .MoveFirst
        strN4Server = Trim(.Fields("hstnam"))
        strN4Authorization = Trim(.Fields("Authorization"))
        strN4UserName = Trim(.Fields("username"))
        strN4Password = Trim(.Fields("password"))
    End With
    Set rstSparcsN4Host = Nothing
    
    Exit Sub
    
err:
    MsgBox "error in retrieving N4 config"
End Sub


