VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSubicINVDE01 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CY Invoice System"
   ClientHeight    =   11145
   ClientLeft      =   15
   ClientTop       =   585
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSubicINVDE01.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid grdCustomers 
      Height          =   6135
      Left            =   1080
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   10821
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColor       =   -2147483624
      FocusRect       =   2
      GridLines       =   2
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
   Begin MSFlexGridLib.MSFlexGrid grdRates 
      Height          =   5895
      Left            =   840
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   10398
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColor       =   -2147483624
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
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
   Begin VB.Frame fraSubHead 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   600
      TabIndex        =   67
      Top             =   2040
      Width           =   8175
      Begin VB.ComboBox cmbChgTyp 
         BackColor       =   &H80000018&
         Height          =   465
         ItemData        =   "frmSubicINVDE01.frx":0442
         Left            =   2760
         List            =   "frmSubicINVDE01.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Width           =   5055
      End
      Begin VB.ComboBox cmbCargo 
         BackColor       =   &H80000018&
         Height          =   465
         ItemData        =   "frmSubicINVDE01.frx":0446
         Left            =   2760
         List            =   "frmSubicINVDE01.frx":0459
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Charge Type"
         Height          =   420
         Left            =   600
         TabIndex        =   69
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Cargo"
         Height          =   420
         Left            =   1560
         TabIndex        =   68
         Top             =   480
         Width           =   1035
      End
   End
   Begin VB.Frame fraDetails 
      Enabled         =   0   'False
      Height          =   1935
      Left            =   600
      TabIndex        =   21
      Top             =   4200
      Width           =   14295
      Begin VB.TextBox txtTelNum 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   420
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   25
         Top             =   1800
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.TextBox txtDysHrs 
         BackColor       =   &H80000018&
         Height          =   420
         Left            =   11640
         TabIndex        =   27
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtRateDesc 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2760
         MaxLength       =   55
         TabIndex        =   24
         Top             =   1320
         Width           =   7575
      End
      Begin VB.TextBox txtQuantity 
         BackColor       =   &H80000018&
         Height          =   420
         Left            =   11640
         TabIndex        =   26
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtSize 
         BackColor       =   &H80000018&
         Height          =   420
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   23
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtRateCode 
         BackColor       =   &H80000018&
         Height          =   420
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblRteTyp 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   270
         Left            =   3720
         TabIndex        =   64
         Top             =   840
         Width           =   150
      End
      Begin VB.Label lblTelNum 
         AutoSize        =   -1  'True
         Caption         =   "Telephone"
         ForeColor       =   &H80000011&
         Height          =   300
         Left            =   1080
         TabIndex        =   63
         Top             =   1800
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lblAmount 
         AutoSize        =   -1  'True
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   11760
         TabIndex        =   34
         Top             =   1440
         Width           =   165
      End
      Begin VB.Label Label17 
         Caption         =   "Amount       "
         Height          =   375
         Left            =   10560
         TabIndex        =   33
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Days / Hours "
         Height          =   420
         Left            =   9480
         TabIndex        =   32
         Top             =   840
         Width           =   2145
      End
      Begin VB.Label Label15 
         Caption         =   "Description "
         Height          =   375
         Left            =   720
         TabIndex        =   31
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
         Height          =   420
         Left            =   10080
         TabIndex        =   30
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label13 
         Caption         =   "Size        "
         Height          =   375
         Left            =   1800
         TabIndex        =   29
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Rate Code   "
         Height          =   375
         Left            =   1080
         TabIndex        =   28
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   14640
      Top             =   10080
   End
   Begin Crystal.CrystalReport crCYInv 
      Left            =   11520
      Top             =   10200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
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
      Left            =   600
      TabIndex        =   50
      Top             =   9360
      Width           =   14295
   End
   Begin VB.Frame fraHeading 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   600
      TabIndex        =   15
      Top             =   0
      Width           =   14295
      Begin VB.TextBox txtVoyNum 
         BackColor       =   &H80000018&
         Height          =   420
         Left            =   11040
         MaxLength       =   20
         TabIndex        =   12
         Top             =   2040
         Width           =   2535
      End
      Begin VB.ComboBox cmbVat 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frmSubicINVDE01.frx":04B3
         Left            =   11040
         List            =   "frmSubicINVDE01.frx":04C3
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1560
         Width           =   3255
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H80000018&
         Height          =   405
         Left            =   2760
         MaxLength       =   165
         TabIndex        =   3
         Top             =   1560
         Width           =   7215
      End
      Begin MSMask.MaskEdBox mskVslArv 
         Height          =   375
         Left            =   11040
         TabIndex        =   9
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   -2147483624
         MaxLength       =   10
         Mask            =   "####/##/##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbGtyCde 
         BackColor       =   &H80000018&
         Height          =   420
         ItemData        =   "frmSubicINVDE01.frx":050D
         Left            =   11040
         List            =   "frmSubicINVDE01.frx":0517
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3000
         Width           =   2655
      End
      Begin VB.ComboBox cmbNonCon 
         BackColor       =   &H80000018&
         Height          =   420
         ItemData        =   "frmSubicINVDE01.frx":0537
         Left            =   11040
         List            =   "frmSubicINVDE01.frx":0544
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtSADays 
         BackColor       =   &H80000018&
         Height          =   420
         Left            =   9480
         MaxLength       =   8
         TabIndex        =   8
         Top             =   3360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDiscnt 
         BackColor       =   &H80000018&
         Height          =   420
         Left            =   8040
         TabIndex        =   7
         Text            =   ".00"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtCusNum 
         BackColor       =   &H80000018&
         Height          =   420
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   0
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtAgent 
         BackColor       =   &H80000018&
         Height          =   420
         Left            =   2760
         MaxLength       =   29
         TabIndex        =   1
         Top             =   585
         Width           =   4935
      End
      Begin VB.TextBox txtInvNum 
         BackColor       =   &H80000018&
         Height          =   420
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtRegNum 
         BackColor       =   &H80000018&
         Height          =   420
         Left            =   2760
         MaxLength       =   12
         TabIndex        =   6
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtVslNam 
         BackColor       =   &H80000018&
         Height          =   420
         Left            =   11040
         MaxLength       =   20
         TabIndex        =   10
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Voyage Number"
         Height          =   300
         Left            =   8640
         TabIndex        =   73
         Top             =   2040
         Width           =   2145
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Vessel Arrival"
         Height          =   300
         Left            =   8400
         TabIndex        =   66
         Top             =   600
         Width           =   2310
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Invoice Details"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   615
         Left            =   0
         TabIndex        =   65
         Top             =   3600
         Width           =   14295
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Bill Category"
         Height          =   300
         Left            =   8760
         TabIndex        =   62
         Top             =   3000
         Width           =   2145
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Non-containerized"
         Height          =   300
         Left            =   8040
         TabIndex        =   61
         Top             =   2520
         Width           =   2805
      End
      Begin VB.Label Label10 
         Caption         =   "No. of days (SA)"
         Height          =   375
         Left            =   8640
         TabIndex        =   60
         Top             =   3360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label9 
         Caption         =   "Registry"
         Height          =   375
         Left            =   1200
         TabIndex        =   54
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Discount"
         Height          =   375
         Left            =   6600
         TabIndex        =   53
         Top             =   3480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblCusNam 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4440
         TabIndex        =   51
         Top             =   120
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agent/Addressee"
         Height          =   420
         Left            =   120
         TabIndex        =   49
         Top             =   585
         Width           =   2475
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "VAT"
         Height          =   300
         Left            =   10320
         TabIndex        =   20
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Customer Number"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Invoice Number "
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Bill Remarks        "
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lblVslNam 
         Caption         =   "Vessel Name "
         Height          =   375
         Left            =   8880
         TabIndex        =   16
         Top             =   1080
         Width           =   1935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdEntries 
      Height          =   3135
      Left            =   600
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   6240
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColor       =   -2147483624
      Enabled         =   0   'False
      FocusRect       =   0
      GridLines       =   2
      ScrollBars      =   2
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
   Begin VB.Frame fraTotals 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   10440
      TabIndex        =   41
      Top             =   6840
      Width           =   4335
      Begin VB.Label lblTax 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3300
         TabIndex        =   72
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(Tax)"
         Height          =   300
         Left            =   960
         TabIndex        =   71
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label lblVat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3300
         TabIndex        =   45
         Top             =   840
         Width           =   915
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3480
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "VAT"
         Height          =   300
         Left            =   1200
         TabIndex        =   43
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   300
         Left            =   960
         TabIndex        =   42
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Label lblF11 
      AutoSize        =   -1  'True
      Caption         =   "F11 = Charge/Cargo Type"
      ForeColor       =   &H80000011&
      Height          =   300
      Left            =   7320
      TabIndex        =   70
      Top             =   9960
      Width           =   3795
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13800
      TabIndex        =   58
      Top             =   10320
      Width           =   1125
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   55
      Top             =   10320
      Width           =   1575
   End
   Begin VB.Label lblF12 
      AutoSize        =   -1  'True
      Caption         =   "F12 = Edit Header"
      ForeColor       =   &H80000011&
      Height          =   300
      Left            =   11880
      TabIndex        =   48
      Top             =   9960
      Width           =   2805
   End
   Begin VB.Label lblF4 
      AutoSize        =   -1  'True
      Caption         =   "F4 = Picklist"
      ForeColor       =   &H80000011&
      Height          =   300
      Left            =   3600
      TabIndex        =   40
      Top             =   9600
      Width           =   2145
   End
   Begin VB.Label lblF10 
      AutoSize        =   -1  'True
      Caption         =   "F10 = Save and Print"
      ForeColor       =   &H80000011&
      Height          =   300
      Left            =   3600
      TabIndex        =   39
      Top             =   9960
      Width           =   3300
   End
   Begin VB.Label lblF8 
      AutoSize        =   -1  'True
      Caption         =   "F8 = Edit Detail"
      ForeColor       =   &H80000011&
      Height          =   300
      Left            =   600
      TabIndex        =   38
      Top             =   9960
      Width           =   2640
   End
   Begin VB.Label lblF7 
      AutoSize        =   -1  'True
      Caption         =   "F7 = Delete Detail"
      ForeColor       =   &H80000011&
      Height          =   300
      Left            =   11880
      TabIndex        =   37
      Top             =   9600
      Width           =   2970
   End
   Begin VB.Label lblF6 
      AutoSize        =   -1  'True
      Caption         =   "F6 = Add Detail"
      ForeColor       =   &H80000011&
      Height          =   300
      Left            =   7320
      TabIndex        =   36
      Top             =   9600
      Width           =   2475
   End
   Begin VB.Label lblF3 
      AutoSize        =   -1  'True
      Caption         =   "F3 = Exit"
      Height          =   300
      Left            =   600
      TabIndex        =   35
      Top             =   9600
      Width           =   1485
   End
   Begin VB.Label lblComputerName 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10320
      TabIndex        =   57
      Top             =   10320
      Width           =   1935
   End
   Begin VB.Label lblUserid 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   56
      Top             =   10320
      Width           =   1935
   End
   Begin VB.Label lblMessages 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   600
      TabIndex        =   59
      Top             =   10320
      Width           =   7815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmSubicINVDE01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'  Accumulators
Dim vDetlAmt As Currency ' Detail amount displayed on lblAmount (fraDetails)
Dim vRteAmnt As Currency
Dim vTmpRTag As String * 1
Dim tmpDiscount As Currency
'  Swithces
Dim vChkCust As Boolean '  Determines if cmdGetCust was used for validating customer
Dim vChkRate As Boolean '  Determines if cmdGetRate was used for validating rate
Dim vAddSwch As Boolean
Dim vEdtSwch As Boolean
Dim vSavSwch As Boolean
'Dim StrtRate As Boolean  ' Determines whether the rate's picklist is to be filled for the first time
Dim vEditHdr As Boolean  ' Determines if invoice header is being edited
Dim vEditSub As Boolean  ' Determines if invoice subhead(chargetype, cargo) is being edited
Dim WithDisc As Boolean  ' Determines if an amount has been discounted
Dim vEscSwch As Boolean  ' Determines whether addition or editing of a detail is discontinued
'  Row Counter
Dim X As Integer         ' Checks if grid contains details

Private Sub cmbCargo_GotFocus()
    SendKeys "%{DOWN}"
End Sub

Private Sub cmbCargo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn And vEditSub = True Then
    SetRatesPicklist (Trim(Left(cmbChgTyp, 2)))
    Call EnableFunctions
    fraSubHead.Enabled = False
    grdEntries.Enabled = True
    vChkRate = False
    vEditSub = False
    grdEntries.HighLight = flexHighlightAlways
'    grdEntries.Col = 0: grdEntries.ColSel = 5
    grdEntries.SetFocus
  Else
    Call FieldAdvance(KeyCode, cmbChgTyp, txtRegNum)
  End If
End Sub

Private Sub cmbChgTyp_GotFocus()
    SendKeys "%{DOWN}"
End Sub

Private Sub cmbChgTyp_KeyDown(KeyCode As Integer, Shift As Integer)
    If vEditSub = True And KeyCode = vbKeyUp Then
      cmbChgTyp.SetFocus
    Else
      Call FieldAdvance(KeyCode, txtRemark, cmbCargo)
    End If
End Sub

Private Sub cmbGtyCde_GotFocus()
  SendKeys "%{DOWN}"
End Sub

Private Sub cmbGtyCde_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyUp Then cmbNonCon.SetFocus
  If KeyCode = vbKeyReturn Then SaveHeaderEntries
End Sub

Private Sub cmbNonCon_GotFocus()
  SendKeys "%{DOWN}"
End Sub

Private Sub cmbNonCon_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtVoyNum, cmbGtyCde)
End Sub

Private Sub ValidateCustomer()
    vChkCust = True
    Call SearchCustomerPicklist
    If Trim(lblCusNam) = "" Then
        MsgBox "Please specify a valid customer code.", vbInformation, "Error Message"
        txtCusNum.SetFocus
        SendKeys "{HOME}": SendKeys "+{END}"
    Else
        txtAgent.SetFocus
    End If
End Sub

Private Sub ValidateRate()
'    If StrtRate = False Then Call SetRatesPicklist(Trim(Left(cmbChgTyp, 2)))
    vRteAmnt = 0
    Call SearchRatesPicklist
    If vRteAmnt = 0 Then
        vChkRate = False
        MsgBox "Either the rate code or container size you entered does not exist.  Press F4 " & _
            "key to view the rate code's picklist.", vbCritical, "Rate Code Error"
        txtRateCode.SetFocus
    Else
        vChkRate = True
        txtTelNum = ""
        lblTelNum.ForeColor = &H80000011
        txtTelNum.Enabled = False
        txtRateDesc.SetFocus
    End If
End Sub

Private Sub cmbVat_GotFocus()
    SendKeys "%{DOWN}"
End Sub

Private Sub cmbVat_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtVslNam, txtVoyNum)
End Sub

Private Sub Form_Load()
    lblUserid = Trim(gUserID)
    lblComputerName = zCurrentComputer
    lblDate = Format(Date, "YYYY/MM/DD")
'    mskVslArv.Text = Format(Date, "YYYY/MM/DD")
    vChkRate = False
    Call SetChargeTypes
    cmbVat.ListIndex = 0
    cmbChgTyp.ListIndex = 0
    cmbNonCon.ListIndex = 0
    cmbGtyCde.ListIndex = 0
    cmbCargo.ListIndex = 0
    Call StartGrid
    Call SetCustomerPicklist
End Sub

Private Sub SetChargeTypes()
    With cmbChgTyp
      .AddItem "VB|Vessel Billing", 0
      .AddItem "MC|Miscellaneous Charges", 1
      .AddItem "VC|Cranage", 2
      .AddItem "AN|Anchorage", 3
      .AddItem "CB|Cargo Billing (Arrastre)", 4
      .AddItem "SS|Stripping/Stuffing", 5
      .AddItem "ST|Storage", 6
'        .AddItem "AN|Anchorage", 0
'        .AddItem "CB|Cargo Billing (Arrastre)", 1
'        .AddItem "MC|Miscellaneous Charges", 2
'        .AddItem "SS|Stripping/Stuffing", 3
'        .AddItem "ST|Storage", 4
'        .AddItem "VB|Vessel Billing", 5
'        .AddItem "VC|Cranage", 6
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Msg As String
    Msg = "Do you want to exit the program?"
    If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, "Exit") = vbNo Then Cancel = True
End Sub

Private Sub grdCustomers_GotFocus()
    fraHeading.Enabled = False
    fraSubHead.Enabled = False
End Sub

Private Sub grdCustomers_KeyDown(KeyCode As Integer, Shift As Integer)
  With grdCustomers
      If KeyCode = vbKeyReturn Then
        .Col = 1
        If .Text = "     < REFRESH PICKLIST >" Then
            txtCusNum = ""
            .Clear
            .Rows = 2
            .Refresh
            Call SetCustomerPicklist
            Exit Sub
        End If
        fraHeading.Enabled = True
        fraSubHead.Enabled = True
        .Col = 0: txtCusNum = Trim(.Text)
        .Col = 1: lblCusNam = Trim(.Text)
        .Col = 2: txtAgent = Trim(.Text)
        .Visible = False
        txtAgent.SetFocus
        vChkCust = True     ' Customer has been validated
      ElseIf KeyCode = vbKeyEscape Then
          fraHeading.Enabled = True
          fraSubHead.Enabled = True
          .Visible = False
          txtCusNum.SetFocus
      End If
  End With
End Sub

Private Sub grdEntries_GotFocus()
  lblMessages = "Use function keys"
End Sub

Private Sub grdEntries_LostFocus()
  lblMessages = ""
End Sub

Private Sub grdRates_GotFocus()
  fraDetails.Enabled = False
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mskVslArv_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub mskVslArv_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtRegNum, txtVslNam)
End Sub

Private Sub tmrTime_Timer()
    lblTime.Caption = Format(Time, "HH:MM")
End Sub

Private Sub txtAgent_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtAgent_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtCusNum, txtInvNum)
End Sub

Private Sub txtAgent_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCusNum_Change()
    vChkCust = False    ' Any changes in customer # resets checker
End Sub
Private Sub txtCusNum_GotFocus()
    lblF4.ForeColor = &H80000012
    lblMessages = "Press F4 to view customer list."
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtCusNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then   ' View Customer Picklist
        With grdCustomers
            .Col = 0: .Row = 1: .ColSel = 2
            .SelectionMode = flexSelectionByRow
            .HighLight = flexHighlightAlways
            .Visible = True
            .SetFocus
        End With
    End If
    If KeyCode = vbKeyReturn Then
        ValidateCustomer
    End If
End Sub

Private Sub txtCusNum_LostFocus()
    lblF4.ForeColor = &H80000011
    lblMessages = ""
End Sub

Private Sub txtDiscnt_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtDiscnt_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtRemark, txtSADays)
End Sub

Private Sub txtDiscnt_LostFocus()
    If Not IsNumeric(txtDiscnt) Then txtDiscnt = 0
    txtDiscnt = Format(txtDiscnt, "###.#0")
End Sub

Private Sub txtDysHrs_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtDysHrs_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call EscapeDetail
        Case vbKeyUp
            txtQuantity.SetFocus
        Case vbKeyReturn
            SaveDetailEntries
    End Select
End Sub

Private Sub txtInvNum_GotFocus()
    txtInvNum = Format(txtInvNum)
    SendKeys "{HOME}": SendKeys "+{END}"
    lblMessages = "Enter 0 if an invoice should not be released."
End Sub

Private Sub txtInvNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
      If Trim(txtInvNum.Text) <> "" Then
           If isExist_OR(Trim(txtInvNum.Text)) = False Then
                Call FieldAdvance(KeyCode, txtAgent, txtRemark)
            Else
               MsgBox "Invoice No. Already Exist!", vbOKOnly + vbInformation, "Error"
               txtInvNum.SetFocus
               SendKeys "{HOME}": SendKeys "+{END}"
            End If
      End If
    End If
End Sub

Private Sub txtInvNum_LostFocus()
    lblMessages = ""
End Sub

Private Sub txtQuantity_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtRateCode_Change()
    vChkRate = False
End Sub

Private Sub txtRateCode_GotFocus()
    lblF4.ForeColor = &H80000012
    SendKeys "{HOME}": SendKeys "+{END}"
    lblMessages = "Press F4 to view valid rate codes."
End Sub

Private Sub txtRateCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vReply As Integer
    
    Select Case KeyCode
        Case vbKeyEscape
            Call EscapeDetail
        Case vbKeyF4
'            If StrtRate = False Then    ' If picklist is loaded for the first time populate grdRates else
'                StrtRate = True
'                Call SetRatesPicklist(Trim(Left(cmbChgTyp, 2)))
'            Else
                grdRates.Col = 0
                grdRates.Row = 1
                grdRates.ColSel = 3
                grdRates.SelectionMode = flexSelectionByRow
                grdRates.HighLight = flexHighlightAlways
'            End If
            grdRates.Visible = True
            grdRates.SetFocus
        Case vbKeyReturn
            txtSize.SetFocus
        Case Else
    End Select
    
End Sub

Private Sub txtRateCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRateCode_LostFocus()
    lblF4.ForeColor = &H80000011
    lblMessages = ""
End Sub

Private Sub txtQuantity_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtRateDesc, txtDysHrs)
    If KeyCode = vbKeyEscape Then EscapeDetail
End Sub

Private Sub txtQuantity_LostFocus()
    If Trim(txtQuantity) = "" Or Not IsNumeric(txtQuantity) Then
        txtQuantity = 0
    End If
    txtQuantity = Format(txtQuantity, "#####.#0")
End Sub

Private Sub SaveDetailEntries()
    
    If Trim(txtDysHrs) = "" Or Not IsNumeric(txtDysHrs) Then txtDysHrs = 0
    txtDysHrs = Format(txtDysHrs, "#####.#0")
    
    '   Validate Rate Code
    If txtRateCode = "" Then
        MsgBox "Rate code should not be blank.  Press F4 key for picklist.", vbExclamation, "Rate Code Error"
        txtRateCode.SetFocus
        Exit Sub
    End If
    If vChkRate = False Then
        MsgBox "Validate rate code before saving detail.", vbInformation, "Message"
        txtRateCode.SetFocus
        Exit Sub
    End If
    
    Call Computations
    vChkRate = False
    grdEntries.Enabled = True
    If MsgBox("Do you want to save this detail?", vbYesNo, "Save") = vbYes Then
        fraDetails.Enabled = False
        Call EnableFunctions
        Call SaveToGrid
        Call DisplayTotal
    Else
        If vEdtSwch = True Or vAddSwch = True Then '   If edited details are not saved, or adding
            vEdtSwch = False                        '   was not continued, then focus back to grid
            vAddSwch = False
            grdEntries.Col = 0
            grdEntries.ColSel = 5
            grdEntries.HighLight = flexHighlightAlways
            grdEntries.Refresh
            Call grdEntries_RowColChange
            Call EnableFunctions
            grdEntries.SetFocus
        Else  '  gridentries is empty
            grdEntries.Enabled = False
            txtRateCode.SetFocus
        End If
    End If

End Sub

Private Sub txtRateDesc_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtRateDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtSize, txtQuantity)
    If KeyCode = vbKeyEscape Then EscapeDetail
End Sub

Private Sub txtRateDesc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRegNum_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtRegNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, cmbCargo, mskVslArv)
End Sub

Private Sub txtRegNum_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRemark_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtSADays_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtSADays_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtDiscnt, cmbChgTyp)
End Sub

Private Sub txtSize_Change()
    vChkRate = False
End Sub

Private Sub txtSize_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtTelNum_GotFocus()
  SendKeys "{HOME}": SendKeys "+{END}"
End Sub

'Private Sub txtVAT_GotFocus()
'    SendKeys "{HOME}": SendKeys "+{END}"
'End Sub

Private Sub txtRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    If Len(Trim(txtRemark)) = 43 Then Beep
    Call FieldAdvance(KeyCode, txtInvNum, cmbChgTyp)
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub CancelHeading()
    fraHeading.Enabled = True
    fraSubHead.Enabled = True
    ' Increment invoice # only after saving & printing
    If vSavSwch = True And txtInvNum <> 0 Then
        txtInvNum = txtInvNum + 1: vSavSwch = False
    Else
        txtInvNum = ""
    End If
    txtRemark.Text = ""
    cmbChgTyp.ListIndex = 0
    cmbCargo.ListIndex = 0
    txtRegNum = ""
    txtVslNam = ""
    txtVoyNum = ""
    mskVslArv.Text = "____/__/__"
'    txtVAT = "N"
    cmbVat.ListIndex = 0
    cmbNonCon.ListIndex = 0
    cmbGtyCde.ListIndex = 0
    txtCusNum.SetFocus
End Sub

Private Sub CancelDetail()
    Dim vReply As Integer

    If X > 0 Then
        vReply = MsgBox("This will end your data entry.  All unsaved data will be lost.  Continue?", vbYesNo + vbDefaultButton2, "CY Invoice System")
        If vReply = vbNo Then
            Exit Sub
        End If
    End If
    grdEntries.Enabled = True
    grdEntries.Clear
    grdEntries.Rows = 2
    grdEntries.Refresh
    X = 0
    Call StartGrid
    grdEntries.Enabled = False
    lblTotal.Caption = "0.00"
    lblVat.Caption = "0.000"
    lblTax.Caption = "0.000"
    Call DisableFunctions
    fraDetails.Enabled = True
    txtRateCode = ""
    txtSize = ""
    txtRateDesc = ""
    txtQuantity = ""
    txtDysHrs = ""
    lblAmount.Caption = ""
    txtTelNum = ""
    fraDetails.Enabled = False
    Call CancelHeading
End Sub

Private Sub EscapeDetail() ' called when user discontinues adding or editing a detail
    vEscSwch = True
    If X > 0 Then
        vAddSwch = False
        vEdtSwch = False
        Call EnableFunctions
        grdEntries.Row = X
        grdEntries.Col = 0
        grdEntries.ColSel = 5
        grdEntries.SelectionMode = flexSelectionByRow
        grdEntries.HighLight = flexHighlightAlways
        Call grdEntries_RowColChange
        grdEntries.SetFocus
    Else
        Call AddDetail   ' clear all detail textboxes
        vAddSwch = False
        fraDetails.Enabled = False
        fraHeading.Enabled = True
        fraSubHead.Enabled = True
        txtCusNum.SetFocus
    End If
End Sub

Private Sub txtSize_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Call EscapeDetail
        Case vbKeyUp
            txtRateCode.SetFocus
        Case vbKeyReturn
            ValidateRate
    End Select
End Sub

Private Sub txtVoyNum_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtVoyNum_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, cmbVat, cmbNonCon)
End Sub

Private Sub txtVoyNum_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

'Private Sub txtVAT_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call FieldAdvance(KeyCode, txtVslNam, cmbNonCon)
'End Sub

'Private Sub txtVAT_KeyPress(KeyAscii As Integer)
'     KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub txtVslNam_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtVslNam_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, mskVslArv, cmbVat)
End Sub

Private Sub txtVslNam_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub SetCustomerPicklist()
    Dim rstCustomers As New ADODB.Recordset
    Dim RowCount As Integer

    Dim ColHdgs(2) As String
    Dim ColCounter As Integer
    
    ColHdgs(0) = " Code"
    ColHdgs(1) = "              Name"
    ColHdgs(2) = "           Agent"
    
    With grdCustomers
        For ColCounter = 0 To 2
            .Col = ColCounter: .Row = 0: .Text = ColHdgs(ColCounter)
        Next
        .RowHeight(0) = 350
        .ColWidth(0) = 1300
        .ColWidth(1) = 5550
        .ColWidth(2) = 6000
        .HighLight = flexHighlightAlways
        .Refresh
        rstCustomers.Open "Select * from CUSTOMER order by cusnam", gcnnBilling, , , adCmdText
            
        RowCount = 1
        Do While Not rstCustomers.EOF
            If RowCount > 1 Then
                .AddItem ("")
            End If
            .RowHeight(RowCount) = 350
            .Row = RowCount
            .Col = 0: .CellAlignment = 4
            .Text = "" & Trim(rstCustomers.Fields("cuscde"))
            .Col = 1: .CellAlignment = 1
            .Text = "" & Trim(rstCustomers.Fields("cusnam"))
            .Col = 2: .CellAlignment = 1
            .Text = "" & Trim(rstCustomers.Fields("careof"))
            RowCount = RowCount + 1
            rstCustomers.MoveNext
        Loop
        rstCustomers.Close
'-------------------------
        .AddItem ("")
        .Row = RowCount
        .RowHeight(RowCount) = 350
        .Col = 1
        .Text = "     < REFRESH PICKLIST >"
        .Refresh
'-------------------------
        .Col = 0: .Row = 1: .ColSel = 2
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
    End With
End Sub

Private Sub SetRatesPicklist(pChgTyp As String)
    Dim rstRate As New ADODB.Recordset
    Dim RowCount As Integer

    Dim ColHdgs(4) As String
    Dim ColCounter As Integer
        
    ColHdgs(0) = "  Code"
    ColHdgs(1) = "Sz"
    ColHdgs(2) = "          Description"
    ColHdgs(3) = "  Amount"
    ColHdgs(4) = "Type"
       
    grdRates.Clear
    grdRates.Rows = 2
    grdRates.Refresh
    
    For ColCounter = 0 To 4
        grdRates.Col = ColCounter
        grdRates.Row = 0
        grdRates.Text = ColHdgs(ColCounter)
    Next
    
    grdRates.RowHeight(0) = 350
    grdRates.ColWidth(0) = 1200
    grdRates.ColWidth(1) = 550
    grdRates.ColWidth(2) = 7800
    grdRates.ColWidth(3) = 1420
    grdRates.ColWidth(4) = 0
   
    grdRates.HighLight = flexHighlightAlways
    grdRates.Refresh
    
    rstRate.Open "Select * from CYRate where cyr_biltyp = '" & pChgTyp & "' order by cyr_rtecde", gcnnBilling, , , adCmdText
            
    RowCount = 1
'    rstRate.MoveFirst
    With rstRate
        Do While Not .EOF
            If RowCount > 1 Then
                grdRates.AddItem ("")
            End If

            grdRates.RowHeight(RowCount) = 350
            grdRates.Row = RowCount
            grdRates.Col = 0: grdRates.Text = "" & Trim(!cyr_rtecde)
            grdRates.Col = 1: grdRates.Text = "" & Trim(!cyr_cntsze)
            grdRates.Col = 2: grdRates.Text = "" & Trim(!cyr_rtedsc): grdRates.CellAlignment = flexAlignLeftCenter
            grdRates.Col = 3: grdRates.Text = Format(!cyr_rteamt, "###,###.#0")
            grdRates.Col = 4: grdRates.Text = "" & Trim(!cyr_biltyp)
            RowCount = RowCount + 1
            rstRate.MoveNext
        Loop
            .Close
    End With
   
    grdRates.AddItem ("")
    grdRates.Row = RowCount
    grdRates.RowHeight(RowCount) = 310
    grdRates.Col = 2
    grdRates.Text = "     < REFRESH PICKLIST >"
    grdRates.Refresh
    grdRates.Col = 0
    grdRates.Row = 1
    grdRates.ColSel = 4
    grdRates.SelectionMode = flexSelectionByRow
    grdRates.HighLight = flexHighlightAlways
'    StrtRate = True ' determines if grdrates has been set
    
End Sub

Private Sub grdRates_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        With grdRates
            fraDetails.Enabled = True
            .Col = 0
            txtRateCode = Trim(.Text)
            .Col = 1
            txtSize = Trim(.Text)
            .Col = 2
            
            If .Text = "     < REFRESH PICKLIST >" Then
                txtRateCode = ""
                .Clear
                .Rows = 2
                .Refresh
                Call SetRatesPicklist(Trim(Left(cmbChgTyp, 2)))
                Exit Sub
            End If

            txtRateDesc = Trim(.Text)
            .Col = 3
            vRteAmnt = .Text
            .Visible = False
            lblF4.ForeColor = &H80000011
            txtRateDesc.SetFocus
            vChkRate = True     ' Rate is valid
            txtTelNum = ""
            lblTelNum.ForeColor = &H80000011
            txtTelNum.Enabled = False
            .Col = 4
            lblRteTyp = Trim(Mid(cmbChgTyp, 4)) 'ReturnRateType(Trim(.Text))
        End With
    End If
    If KeyCode = vbKeyEscape Then
        fraDetails.Enabled = True
        grdRates.Visible = False
        txtRateCode.SetFocus
    End If
End Sub

Private Sub SaveToGrid()

    If vEdtSwch = False Then        '  Check if in edit mode
    
        vAddSwch = True
        X = X + 1   ' Accumulate the number of items on the grid
        If X > 1 Then
            grdEntries.AddItem ("")
        End If
        grdEntries.RowHeight(X) = 350
        grdEntries.Refresh
        grdEntries.Row = X
    
    End If
    
    grdEntries.Col = 0
    grdEntries.Text = Trim(txtRateCode)
    grdEntries.Col = 1
    grdEntries.Text = Trim(txtSize)
    grdEntries.Col = 2
    grdEntries.Text = Format(txtQuantity, "#####.#0")
    grdEntries.Col = 3
    grdEntries.Text = Format(txtDysHrs, "#####.#0")
    grdEntries.Col = 4
    grdEntries.Text = Format(vDetlAmt, "#,###,###.#0")
    grdEntries.Col = 5
    grdEntries.Text = Trim(txtRateDesc)
' --------------
    grdEntries.Col = 6
    grdEntries.Text = Trim(txtTelNum)
    grdEntries.Col = 7
    grdEntries.Text = Trim(lblRteTyp)
    grdEntries.Col = 8
    grdEntries.Text = Trim(cmbCargo.Text)
' --------------
    grdEntries.Row = X
    grdEntries.Col = 0
    grdEntries.ColSel = 7
    vAddSwch = False
    vEdtSwch = False
    grdEntries.SelectionMode = flexSelectionByRow
    grdEntries.HighLight = flexHighlightAlways
    grdEntries.SetFocus
    grdEntries.Refresh
    
End Sub

Private Sub AddDetail()
    grdEntries.HighLight = flexHighlightNever
    vAddSwch = True
    txtRateCode = ""
    txtSize = ""
    txtRateDesc = ""
    txtQuantity = ""
    txtDysHrs = ""
    lblAmount = ""
    txtTelNum = ""
    lblRteTyp = ""
    txtRateCode.SetFocus
End Sub

Private Sub SearchRatesPicklist()
    With grdRates
         .Col = 0: .Row = 1
         Do Until .Row = .Rows - 1
             If Trim(txtRateCode) = Trim(.Text) Then
                .Col = 1
                If txtSize = "" & Trim(.Text) Then
                    .Col = 2: txtRateDesc = Trim(.Text)
                    .Col = 3: vRteAmnt = .Text
                    .Col = 4: lblRteTyp = Trim(Mid(cmbChgTyp, 4)) 'ReturnRateType(Trim(.Text))
                    Exit Do
                End If
             End If
             .Col = 0: .Row = .Row + 1
         Loop
    End With
End Sub

Private Sub SearchCustomerPicklist()
    txtCusNum = Format(txtCusNum, "000000")
    With grdCustomers
         .Col = 0: .Row = 1
         lblCusNam = "": txtAgent = ""
         Do Until .Row = .Rows - 1
             If CStr(txtCusNum) = Trim(.Text) Then
                .Col = 1: lblCusNam = Trim(.Text)
                .Col = 2: txtAgent = Trim(.Text)
                Exit Do
            End If
            .Col = 0: .Row = .Row + 1
         Loop
    End With
End Sub

Private Sub StartGrid()
    Dim ColHdgs(5) As String
    Dim ColCounter As Integer
    
    ColHdgs(0) = "  Code"
    ColHdgs(1) = "Sz"
    ColHdgs(2) = "  Qty"
    ColHdgs(3) = "Dys/Hrs"
    ColHdgs(4) = " Amount"
    ColHdgs(5) = "    Description"
    With grdEntries
        For ColCounter = 0 To 5
            .Col = ColCounter: .Row = 0: .Text = ColHdgs(ColCounter)
        Next
        .RowHeight(0) = 350
        .ColWidth(0) = 1300
        .ColWidth(1) = 550
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        .ColWidth(4) = 1600
        .ColWidth(5) = 3630
        .HighLight = flexHighlightNever
        .Refresh
    End With
End Sub

Private Sub grdEntries_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF6                                '   F6=ADD
            If grdEntries.Rows = 16 Then
                MsgBox "You have reached the maximum number of details. " & _
                    "Print this invoice first then save next entries using another invoice.", _
                    vbInformation, "Add"
                Exit Sub
            End If
            Call DisableFunctions
            Call AddDetail
        Case vbKeyF7                                '   F7=DELETE
            Call DeleteDetail
        Case vbKeyF8                                '   F8=EDIT
            grdEntries.HighLight = flexHighlightNever
            Call DisableFunctions
            txtRateCode.SetFocus
            vEdtSwch = True
        Case vbKeyF10                               '   F10=SAVE/PRINT
            If MsgBox("Do you want to save and print invoice no. " & Trim(txtInvNum) & " ?", vbYesNo, "Save & Print Message") = vbYes Then
                Call SaveToDatabase
            Else
                grdEntries.SetFocus
            End If
        Case vbKeyF11                               '   F11=EDIT SUBHEAD
            vEditSub = True
            Call EditHeader
'            fraSubHead.Enabled = True
            cmbChgTyp.SetFocus
        Case vbKeyF12                               '   F12=EDIT ALL HEADER
            tmpDiscount = CCur(txtDiscnt)
            vEditHdr = True
            Call EditHeader
            fraHeading.Enabled = True
            txtCusNum.SetFocus
    End Select
End Sub

Private Sub EditHeader()
    fraSubHead.Enabled = True
    grdEntries.HighLight = flexHighlightNever
    Call DisableFunctions
    fraDetails.Enabled = False
End Sub

Private Sub ChangeHeader()
    fraHeading.Enabled = False
    fraSubHead.Enabled = False
    grdEntries.Enabled = True
    grdEntries.SetFocus
    grdEntries.Row = 0
    
    Do Until grdEntries.Row = X                      ' Gather from grid the necessary data for
        grdEntries.Row = grdEntries.Row + 1            ' possible changes in computation
        grdEntries.Col = 3
        txtDysHrs = grdEntries.Text ' dyshrs
        grdEntries.Col = 4
        vDetlAmt = grdEntries.Text  ' item amount
        Call EditDiscount
        grdEntries.Text = Format(vDetlAmt, "###,###.#0")
    Loop
    
    If CCur(txtDiscnt) = 0 Then
        WithDisc = False
    End If
    Call DisplayTotal
    vEditHdr = False
    Call EnableFunctions
    Call grdEntries_RowColChange
End Sub

Private Sub EditDiscount()
    SearchRatesPicklist
    If WithDisc = True Then
        vDetlAmt = vDetlAmt + (tmpDiscount * CCur(txtDysHrs) * vRteAmnt)
    End If
    If CCur(txtDiscnt) > 0 Then
        vDetlAmt = vDetlAmt - (CCur(txtDiscnt) * CCur(txtDysHrs) * vRteAmnt)
        WithDisc = True
    End If
End Sub

Private Sub grdEntries_RowColChange()
    If vAddSwch = True Or vEdtSwch = True Or vEditHdr = True Or vSavSwch = True Then
        Exit Sub
    End If
    If X > 0 Then
        txtRateCode = grdEntries.TextMatrix(grdEntries.Row, 0)
        txtSize = grdEntries.TextMatrix(grdEntries.Row, 1)
        txtQuantity = grdEntries.TextMatrix(grdEntries.Row, 2)
        txtDysHrs = grdEntries.TextMatrix(grdEntries.Row, 3)
        lblAmount.Caption = grdEntries.TextMatrix(grdEntries.Row, 4)
        txtRateDesc = grdEntries.TextMatrix(grdEntries.Row, 5)
        txtTelNum = grdEntries.TextMatrix(grdEntries.Row, 6)
        lblRteTyp = grdEntries.TextMatrix(grdEntries.Row, 7)
    End If
End Sub

Private Sub SaveHeaderEntries()
    Dim rstInvoice As New ADODB.Recordset
    
    If MsgBox("Are all entries correct?", vbYesNo, "CY Invoice System") = vbNo Then
        txtCusNum.SetFocus: Exit Sub
    End If
    
'   VALIDATE HEADER ENTRIES BEFORE SAVING
    '   Check if vessel arrival is specified
    If Not IsDate(mskVslArv.Text) Then
      MsgBox "Please specify the vessel's arrival date", vbInformation, Me.Caption
      mskVslArv.SetFocus
      Exit Sub
    End If
    
    '   Check if customer is valid
    If Not vChkCust Then
        MsgBox "Validate customer before saving.", vbInformation, "Invoice Error"
        txtCusNum.SetFocus
        Exit Sub
    End If
    
    '   Validate Invoice Number
    If Trim(txtInvNum) <> 0 Then
        If Trim(txtInvNum) = "" Or Not IsNumeric(txtInvNum) Then
            MsgBox "Specify a valid invoice number.", vbExclamation, "Invoice Error"
            txtInvNum.SetFocus
            Exit Sub
        End If
        rstInvoice.Open "Select * from INVICT where invnum = '" & Trim(txtInvNum) & _
                "'", gcnnBilling, , , adCmdText
        If Not rstInvoice.EOF Then
            rstInvoice.Close
            MsgBox "The invoice number you entered has been used.", vbExclamation, "Invoice Error"
            txtInvNum.SetFocus
            Exit Sub
        Else
            rstInvoice.Close
        End If
    End If
    
    '------------- HEADER ENTRIES ARE ALL VALID
    SetRatesPicklist (Trim(Left(cmbChgTyp, 2)))
    fraHeading.Enabled = False
    fraSubHead.Enabled = False
    fraDetails.Enabled = True
    vChkRate = False
    txtRateCode.SetFocus
    ' Check if header is being edited to manipulate any changes in detail
    If vEditHdr = True Then Call ChangeHeader
End Sub

Private Sub EnableFunctions()
    fraDetails.Enabled = False
    grdEntries.Enabled = True
    lblF6.ForeColor = &H80000012
    lblF7.ForeColor = &H80000012
    lblF8.ForeColor = &H80000012
    lblF10.ForeColor = &H80000012
    lblF11.ForeColor = &H80000012
    lblF12.ForeColor = &H80000012
End Sub

Private Sub DisableFunctions()
    fraDetails.Enabled = True
    grdEntries.Enabled = False
    lblF6.ForeColor = &H80000011
    lblF7.ForeColor = &H80000011
    lblF8.ForeColor = &H80000011
    lblF10.ForeColor = &H80000011
    lblF11.ForeColor = &H80000011
    lblF12.ForeColor = &H80000011
End Sub

Private Sub DeleteDetail()
    Dim vReply As Integer
    
    vReply = MsgBox("Delete this detail?", vbYesNo + vbDefaultButton2, "Delete")
    If vReply = vbNo Then
        Exit Sub
    End If
    If X = 1 Then
        MsgBox "Deleting the last item is not allowed.  Press F8 to edit this detail instead.", vbInformation, "Delete"
    Else
        grdEntries.RemoveItem (grdEntries.Row)
        X = X - 1
        grdEntries.Refresh
        grdEntries.Row = X
        Call grdEntries_RowColChange
        Call DisplayTotal
    End If
End Sub

Private Sub Computations()
   
    vDetlAmt = 0
    ' Detail Amount
    If txtQuantity > 0 And txtDysHrs > 0 Then
        vDetlAmt = vRteAmnt * txtQuantity * txtDysHrs
    End If
    If txtQuantity = 0 And txtDysHrs > 0 Then
        vDetlAmt = vRteAmnt * txtDysHrs
    End If
    If txtQuantity > 0 And txtDysHrs = 0 Then
        vDetlAmt = vRteAmnt * txtQuantity
    End If
    lblAmount.Caption = Format(vDetlAmt, "###,###.#0")
    
End Sub

Private Sub DisplayTotal()
    Dim vTtlAmnt As Currency
    Dim vTtlVat As Currency
    Dim vTtlTax As Currency

    vAddSwch = True     ' Use boolean vAddSwch to bypass grdentries_rowcolchange
    grdEntries.Row = 0
    Do Until grdEntries.Row = X
        grdEntries.Row = grdEntries.Row + 1
        grdEntries.Col = 4
        vTtlAmnt = vTtlAmnt + grdEntries.Text
    Loop
    lblTotal.Caption = Format(vTtlAmnt, "###,###,###.#0")
    
'    If txtVAT = "Y" Then
'        vTtlVat = vTtlAmnt * 0.1
'    End If
    vTtlVat = 0
    vTtlTax = 0
    Select Case cmbVat.ListIndex
'       Case 0
        Case 1
            vTtlVat = vTtlAmnt * 0.1
        Case 2
            vTtlTax = vTtlAmnt * 0.01
        Case 3
            vTtlVat = vTtlAmnt * 0.1
            vTtlTax = vTtlAmnt * 0.01
    End Select
    lblVat.Caption = Format(vTtlVat, "#,###.##0")
    lblTax.Caption = Format(vTtlTax, "#,###.##0")
    
    grdEntries.Row = X
    grdEntries.Col = 0
    grdEntries.ColSel = 5
    grdEntries.HighLight = flexHighlightAlways
    
    vAddSwch = False    ' Reset vAddSwch
    
End Sub

Private Sub SaveToDatabase()    'Save all details from the grid
    Dim rstInvict As New ADODB.Recordset
    Dim rstInvcyb As New ADODB.Recordset
    Dim tempAmt As Currency
    Dim blnPrint As Boolean
    
    Dim tempRate As String ' for checking rate type
    Dim tempSize As String '
    
    Dim InvNo As Long
    Dim VRefnum As Long

    
    Dim itemCtr As Integer

    vSavSwch = True
    VRefnum = gzGetRefNum("INV")
    'InvNo = Trim(txtInvNum.Text)
    grdEntries.Row = 0
    Screen.MousePointer = vbHourglass
    With rstInvcyb
        .Open "Invcyb", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
        Do Until grdEntries.Row = X
            itemCtr = itemCtr + 1
            .AddNew
            grdEntries.Row = grdEntries.Row + 1
            grdEntries.Col = 0
            .Fields("rtecde") = Trim(grdEntries.Text)
            If itemCtr = 1 Then tempRate = grdEntries.Text
            grdEntries.Col = 1
            .Fields("cntsze") = Trim(grdEntries.Text)
            If itemCtr = 1 Then tempSize = grdEntries.Text
            grdEntries.Col = 2
            .Fields("qty") = grdEntries.Text
            grdEntries.Col = 3
            .Fields("dyshrs") = grdEntries.Text
            grdEntries.Col = 4
            .Fields("invamt") = grdEntries.Text
            tempAmt = grdEntries.Text   ' Get the detail amount for further computations
            grdEntries.Col = 5
            .Fields("rtedsc") = Trim(grdEntries.Text)
            grdEntries.Col = 6
            .Fields("vatcde") = cmbVat.ListIndex
            If cmbVat.ListIndex = 1 Or cmbVat.ListIndex = 3 Then
'            If txtVAT = "Y" Then
                .Fields("vatcde") = "1"
                .Fields("invvat") = tempAmt * 0.1
            Else
'                .Fields("vatcde") = ""
                .Fields("invvat") = 0
            End If
            If cmbVat.ListIndex = 2 Or cmbVat.ListIndex = 3 Then
                .Fields("invtax") = tempAmt * 0.01
            Else
                .Fields("invtax") = 0
            End If
            grdEntries.Col = 8
            .Fields("cargo") = UCase(Trim(grdEntries.Text))
            .Fields("discnt") = txtDiscnt
            .Fields("invremark") = ""
            .Fields("refnum") = VRefnum
            '.Fields("invnum") = InvNo
            .Fields("itmnum") = itemCtr
            .Fields("sysdttm") = gzGetSysDate
            .Fields("status") = ""
            .Fields("rectag") = ""
            .Fields("userid") = gUserID
            .Fields("updcde") = "A"
            .Update
        Loop
            .Close
    End With
    With rstInvict
        .Open "Invict", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
        .AddNew
        .Fields("refnum") = VRefnum
        '.Fields("invnum") = InvNo

        
'        If Trim(CheckRateType(tempRate, tempSize)) = "SA" Then ' Invoice number
'            .Fields("invnum") = 0: blnPrint = False
'        Else
        .Fields("invnum") = Trim(txtInvNum.Text): blnPrint = True
'        End If
        .Fields("cuscde") = txtCusNum
        .Fields("cusnam") = Trim(lblCusNam)
        .Fields("invdttm") = gzGetSysDate
        .Fields("vslnam") = Trim(txtVslNam)
        .Fields("regnum") = Trim(txtRegNum)
        .Fields("voynum") = Trim(txtVoyNum)
        .Fields("invremark") = Trim(txtRemark) & "|" & Trim(txtAgent)
        .Fields("invamt") = lblTotal.Caption
        .Fields("invvat") = lblVat.Caption
        .Fields("invtax") = lblTax.Caption
        .Fields("status") = ""
'        Call GetRecordTag
        If Trim(cmbGtyCde.Text) = "On Account" Then
            .Fields("gtycde") = "N"
        Else
            .Fields("gtycde") = "Y"
        End If
        Select Case Trim(cmbNonCon.Text)
          Case "NA"
            .Fields("noncnt") = "0"
          Case "Basin"
            .Fields("noncnt") = "1"
          Case "Berthside"
            .Fields("noncnt") = "2"
        End Select
        If IsDate(mskVslArv.Text) Then
            .Fields("arrival") = CDate(mskVslArv)
        End If
        .Fields("rectag") = "" ' vTmpRTag
        .Fields("userid") = gUserID
        .Fields("updcde") = "A"
        .Fields("cfscy") = "1"
        .Update
        .Close
    End With
        
    X = 0

    On Error Resume Next
    
'    If blnPrint Then
         'PRINT INVOICE
    crCYInv.ReportFileName = App.Path & "\SubicInvoice1.rpt"
    crCYInv.CopiesToPrinter = 1
    crCYInv.ParameterFields(1) = "InvoiceNo; " & Trim(txtInvNum.Text) & ";TRUE"
'    Else 'PRINT SA
'        MsgBox "You will be printing a STATEMENT OF ACCOUNT. Please make necessary " & _
'                "adjustments now. Click OK to start printing.", vbInformation, "Save Message"
'        crCYInv.ReportFileName = App.Path & "\SubicINVSA.rpt"
'        crCYInv.CopiesToPrinter = 4
'        crCYInv.ParameterFields(1) = "Reference; " & Trim(InvNo) & ";TRUE"
'        crCYInv.ParameterFields(2) = "NumDays; " & Trim(txtSADays) & ";TRUE"
'    End If
    crCYInv.ProgressDialog = False
    crCYInv.Action = 1
    
    Call CancelDetail
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub FieldAdvance(pKeycode As Integer, pPrvCtl As Control, pNxtCtl As Control)
    Select Case pKeycode
        Case vbKeyUp
            pPrvCtl.SetFocus
        Case vbKeyReturn
             pNxtCtl.SetFocus
        End Select
End Sub

Public Function isExist_OR(ByVal Invnum As Long) As Boolean
  Dim rst As New ADODB.Recordset
  
  rst.Open "Invict", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
  
  With rst
    .Find "invnum=" & Invnum, , adSearchForward
       If Not .EOF Then
            isExist_OR = True
            Exit Function
       Else
            isExist_OR = False
       End If
  End With
  rst.Close
  Set rst = Nothing
End Function

