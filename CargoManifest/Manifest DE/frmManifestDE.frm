VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmManifestDE 
   Caption         =   "Cargo Manifest"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15240
   Icon            =   "frmManifestDE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbCarCde 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmManifestDE.frx":0442
      Left            =   2160
      List            =   "frmManifestDE.frx":0444
      TabIndex        =   1
      Text            =   "[Select Carrier]"
      Top             =   720
      Width           =   6615
   End
   Begin VB.TextBox txtPO 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   4935
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13800
      TabIndex        =   39
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   38
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      TabIndex        =   37
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8813
      TabIndex        =   36
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5213
      TabIndex        =   35
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6413
      TabIndex        =   34
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7606
      TabIndex        =   33
      Top             =   9840
      Width           =   1215
   End
   Begin VB.TextBox txtBroker 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   3600
      Width           =   3735
   End
   Begin VB.TextBox txtBillNo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtConDescr 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   8880
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   480
      Width           =   6135
   End
   Begin VB.TextBox txtConsignee 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   3120
      Width           =   3735
   End
   Begin VB.TextBox txtVoyageNo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txtRegistryNo 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox txtVesselName 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1200
      Width           =   4575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   18
      Top             =   9840
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "List of Container/s"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   4455
      Left            =   240
      TabIndex        =   19
      Top             =   5160
      Width           =   14775
      Begin VB.ComboBox cmbFullEmpty 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmManifestDE.frx":0446
         Left            =   2400
         List            =   "frmManifestDE.frx":0450
         TabIndex        =   15
         Text            =   "FCL"
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox txtWeight 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cmbSize 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmManifestDE.frx":045E
         Left            =   2400
         List            =   "frmManifestDE.frx":046B
         TabIndex        =   16
         Text            =   "[Select size]"
         Top             =   3360
         Width           =   2295
      End
      Begin VB.ComboBox cmbType 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmManifestDE.frx":047B
         Left            =   2400
         List            =   "frmManifestDE.frx":048B
         TabIndex        =   13
         Text            =   "[Select type]"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox txtSealNo 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtConNo 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   360
         Width           =   2295
      End
      Begin MSFlexGridLib.MSFlexGrid flxgrdConNo 
         Height          =   3615
         Left            =   5160
         TabIndex        =   17
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   6376
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label17 
         Caption         =   "Full / Empty :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Weight :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblTotalCon 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   12600
         TabIndex        =   43
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label lbl45 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   10680
         TabIndex        =   42
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label lbl40 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8760
         TabIndex        =   41
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label lbl20 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6840
         TabIndex        =   40
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Type :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   30
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Size :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   29
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Seal Number :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   28
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Container Number :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   2055
      End
   End
   Begin MSMask.MaskEdBox mskArrivalDte 
      Height          =   375
      Left            =   2160
      TabIndex        =   46
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskLastDischarge 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/MM/dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   47
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label15 
      Caption         =   "Port of Origin :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   45
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label14 
      Caption         =   "Last Discharge  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   44
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label13 
      Caption         =   "Broker :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   32
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Arrival Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   31
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Bill of Lading :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Consignee :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   25
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Container Name\Description :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   24
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Voyage Number :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Registry Number :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Vessel Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Carrier :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   720
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuClear 
         Caption         =   "C&lear"
      End
      Begin VB.Menu mnuClearDetails 
         Caption         =   "Clear &Details"
      End
      Begin VB.Menu mnuAddNewDet 
         Caption         =   "Add &New Details"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuReport 
         Caption         =   "&Report"
      End
   End
End
Attribute VB_Name = "frmManifestDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstManifestHead As ADODB.Recordset
Dim rstManifestDet As ADODB.Recordset
Dim rstCarriers As ADODB.Recordset
Dim intCounter As Integer
Dim intRowNo As Integer
Dim blnFlag As Boolean 'Used to edit details data
Dim blnNew As Boolean
Dim blnAlternateColor As Boolean
    

Private Sub cmbFullEmpty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        cmbSize.SetFocus
    End If
End Sub

Private Sub cmbFullEmpty_LostFocus()
    If blnFlag = False Then Exit Sub
    With flxgrdConNo
        .Row = Trim(intRowNo)
        .Col = 5
        .Text = cmbFullEmpty.Text
    End With
End Sub

Private Sub cmdDelete_Click()
    Dim rstDltRecord As New ADODB.Recordset
    Dim strResponse As String
    
    strResponse = ""
    If blnFlag = True Then
        strResponse = MsgBox("Are you sure you want to delete Container No. " & txtConNo.Text & "?", vbQuestion + vbYesNo, "Cargo Manifest")
        If strResponse = vbYes Then
            If Allow_Delete = True Then
                rstDltRecord.Open "DELETE FROM CargoMDet WHERE bilnum='" & Trim(txtBillNo.Text) & "' AND ctnnum='" & Trim(txtConNo.Text) & "'", gcnnBilling, adOpenDynamic, adLockOptimistic
                MsgBox "Container No. " & txtConNo.Text & " was deleted!", vbInformation, "Cargo Manifest"
                Set rstDltRecord = Nothing
                Call cmdFirst_Click
            End If
        End If
        blnFlag = False
    Else
        strResponse = MsgBox("Are you sure you want to delete Bill No. " & txtBillNo.Text & "?", vbQuestion + vbYesNo, "Cargo Manifest")
        If strResponse = vbYes Then
            If Allow_Delete = True Then
                If Delete_ManifestDetail(Trim(txtBillNo.Text)) = True Then
                    rstDltRecord.Open "DELETE FROM CargoMHead WHERE bilnum='" & Trim(txtBillNo.Text) & "'", gcnnBilling, adOpenDynamic, adLockOptimistic
                    MsgBox "Bill No. " & txtBillNo.Text & " was deleted!", vbInformation, "Cargo Manifest"
                    Set rstDltRecord = Nothing
                End If
            End If
        End If
        Call cmdFirst_Click
    End If
End Sub

Private Function Allow_Delete() As Boolean
    Dim rstAllowDelete As New ADODB.Recordset
    
    If blnFlag = True Then
        rstAllowDelete.Open ("SELECT ctnnum FROM CargoMdet WHERE bilnum='" & Trim(txtBillNo.Text) & "' AND ctnnum='" & Trim(txtConNo.Text) & "' AND gpsnum IS NULL"), gcnnBilling, adOpenDynamic, adLockOptimistic
        If Not rstAllowDelete.BOF Then
            Allow_Delete = True
        Else
            MsgBox "Cannot delete,this container is already billed!", vbInformation, "Cargo Manifest"
            Allow_Delete = False
        End If
    Else
        rstAllowDelete.Open "SELECT ctnnum FROM CargoMdet WHERE bilnum='" & Trim(txtBillNo.Text) & "' AND gpsnum IS NOT NULL", gcnnBilling, adOpenDynamic, adLockOptimistic
        If Not rstAllowDelete.BOF Then
            MsgBox "Cannot delete,this Bill No. has billed container/s!", vbInformation, "Cargo Manifest"
            Allow_Delete = False
        Else
            Allow_Delete = True
        End If
    End If
    Set rstAllowDelete = Nothing
End Function

Private Sub cmdFirst_Click()
    If rstManifestHead.BOF Then Exit Sub
    rstManifestHead.MoveFirst
    If rstManifestHead.BOF Then Exit Sub
    Call InitializeDetails
    flxgrdConNo.Clear
    flxgrdConNo.Refresh
    Call Format_flxgrdConNo
    Call RetrieveRecords
End Sub

Private Sub cmdLast_Click()
    If rstManifestHead.EOF Then Exit Sub
    rstManifestHead.MoveLast
    If rstManifestHead.EOF Then Exit Sub
    Call InitializeDetails
    flxgrdConNo.Clear
    flxgrdConNo.Refresh
    Call Format_flxgrdConNo
    Call RetrieveRecords
End Sub

Private Sub cmdNew_Click()
    blnNew = True
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
    cmdSearch.Enabled = False
    Call InitializeControls
End Sub

Private Sub cmdNext_Click()
    If rstManifestHead.EOF Then Exit Sub
    rstManifestHead.MoveNext
    If rstManifestHead.EOF Then Exit Sub
    Call InitializeDetails
    flxgrdConNo.Clear
    flxgrdConNo.Refresh
    Call Format_flxgrdConNo
    Call RetrieveRecords
End Sub

Private Sub cmdPrevious_Click()
    If rstManifestHead.BOF Then Exit Sub
    rstManifestHead.MovePrevious
    If rstManifestHead.BOF Then Exit Sub
    Call InitializeDetails
    flxgrdConNo.Clear
    flxgrdConNo.Refresh
    Call Format_flxgrdConNo
    Call RetrieveRecords
End Sub

Private Sub cmdSave_Click()
    Dim errBilling As ADODB.Error
    Dim lsErrStr As String
    Dim intSavingCntr As Integer
    Dim intSavingCntr2 As Integer
    Dim strConNo As String
    Dim strSealNo As String
    Dim strType As String
    Dim strSize As String
    Dim strWeight As String
    Dim strFullEmpty As String
    Dim strGatepassNo As String
    Dim strResponse As String
    
    strResponse = MsgBox("Are you sure you want to save this record?", vbYesNo + vbQuestion, "Cargo Manifest")
    If strResponse = vbNo Then
        Call cmdFirst_Click
        Exit Sub
    End If
      
    intSavingCntr = 0
    intSavingCntr2 = 0
    strConNo = ""
    strSealNo = ""
    strType = ""
    strSize = ""
    
    strConNo = flxgrdConNo.TextMatrix(1, 0)
    If Trim(txtBillNo.Text) = "" Or Trim(cmbCarCde.Text) = "" Or Trim(txtVesselName.Text) = "" Or Trim(txtRegistryNo.Text) = "" Or Trim(txtVoyageNo.Text) = "" Or Trim(txtConsignee.Text) = "" & _
       Trim(txtConDescr.Text) = "" Or Trim(strConNo) = "" Then
        MsgBox "Please provide valid entries!", vbExclamation, "Cargo Manifest"
        Call InitializeControls
        Exit Sub
    End If
    
    If Trim(mskArrivalDte.Text) <> "____-__-__" Then
        If Not IsDate(mskArrivalDte) Then
            MsgBox "Arrival date is not a valid data!", vbExclamation, "Cargo Manifest"
            mskArrivalDte.SetFocus
            Exit Sub
        End If
    End If
            
'On Error GoTo Trap:
    
    If blnNew = True Then
        'Insert header data
        With rstManifestHead
            .AddNew
            .Fields("bilnum") = txtBillNo.Text
            .Fields("carcde") = Mid(Trim(cmbCarCde.Text), 1, 6)
            .Fields("vslname") = txtVesselName.Text
            .Fields("po") = txtPO.Text
            .Fields("regnum") = txtRegistryNo.Text
            .Fields("voynum") = txtVoyageNo.Text
            .Fields("consignee") = txtConsignee.Text
            .Fields("ctnNameDesc") = txtConDescr.Text
            .Fields("broker") = txtBroker.Text
            If mskArrivalDte.Text <> "____-__-__" Then
                .Fields("arvdte") = mskArrivalDte.Text
            End If
            If mskLastDischarge.Text <> "____-__-__" Then
                .Fields("dischargedte") = mskLastDischarge.Text
            End If
            .Fields("sysdte") = gzGetSysDate
            .Fields("userid") = zCurrentUser
            .Update
        End With
        
        'Insert container details
        If flxgrdConNo.Rows >= 2 Then
            Do While flxgrdConNo.Rows > intSavingCntr + 1
                flxgrdConNo.Row = intSavingCntr + 1
                strConNo = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 0)
                strSealNo = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 1)
                strType = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 2)
                strSize = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 3)
                strWeight = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 4)
                strFullEmpty = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 5)
                strGatepassNo = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 6)
                If Trim(strConNo) = "" Then
                    MsgBox "Please provide valid entries!", vbExclamation, "Cargo Maintenance"
                    Call InitializeControls
                    Exit Sub
                End If
                With rstManifestDet
                    .AddNew
                    .Fields("bilnum") = txtBillNo.Text
                    .Fields("ctnnum") = Trim(strConNo)
                    .Fields("ctntype") = Trim(strType)
                    .Fields("ctnsze") = Trim(strSize)
                    .Fields("ctnweight") = Trim(strWeight)
                    .Fields("fullempty") = Trim(strFullEmpty)
                    .Fields("silnum") = Trim(strSealNo)
                    .Fields("regnum") = txtRegistryNo.Text
                    If Trim(strGatepassNo) <> "" Then
                        .Fields("gpsnum") = CLng(Trim(strGatepassNo))
                    End If
                    .Fields("sysdte") = gzGetSysDate
                    .Fields("userid") = zCurrentUser
                    .Update
                End With
                intSavingCntr = intSavingCntr + 1
            Loop
        End If
        
        MsgBox "Saving successful!", vbInformation, "Cargo Manifest"
        Call cmdFirst_Click
    Else
        'Insert header data
        With rstManifestHead
            .Fields("bilnum") = txtBillNo.Text
            .Fields("carcde") = Mid(Trim(cmbCarCde.Text), 1, 6)
            .Fields("vslname") = txtVesselName.Text
            .Fields("po") = Trim(txtPO.Text)
            .Fields("regnum") = txtRegistryNo.Text
            .Fields("voynum") = txtVoyageNo.Text
            .Fields("consignee") = txtConsignee.Text
            .Fields("ctnNameDesc") = txtConDescr.Text
            .Fields("broker") = txtBroker.Text
            If mskArrivalDte.Text <> "____-__-__" Then
                .Fields("arvdte") = mskArrivalDte.Text
            End If
            If mskLastDischarge.Text <> "____-__-__" Then
                .Fields("dischargedte") = mskLastDischarge.Text
            End If
            .Fields("sysdte") = gzGetSysDate
            .Fields("userid") = zCurrentUser
            .Update
        End With
        
        'Insert container details
        'Instead of updating delete records first and then insert new one.
         If Delete_ManifestDetail(Trim(txtBillNo.Text)) = False Then
            MsgBox "Manifest details was not updated!", vbExclamation, "Cargo Manifest"
            Exit Sub
        End If
        If flxgrdConNo.Rows >= 2 Then
            Do While flxgrdConNo.Rows > intSavingCntr + 1
                flxgrdConNo.Row = intSavingCntr + 1
                strConNo = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 0)
                strSealNo = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 1)
                strType = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 2)
                strSize = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 3)
                strWeight = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 4)
                strFullEmpty = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 5)
                strGatepassNo = flxgrdConNo.TextMatrix(flxgrdConNo.Row, 6)
                If Trim(strConNo) = "" Then
                    MsgBox "Please provide valid entries!", vbExclamation, "Cargo Maintenance"
                    Call InitializeControls
                    Exit Sub
                End If
                With rstManifestDet
                    .AddNew
                    .Fields("bilnum") = txtBillNo.Text
                    .Fields("ctnnum") = Trim(strConNo)
                    .Fields("ctntype") = Trim(strType)
                    .Fields("ctnsze") = Trim(strSize)
                    .Fields("ctnweight") = Trim(strWeight)
                    .Fields("fullempty") = Trim(strFullEmpty)
                    .Fields("silnum") = Trim(strSealNo)
                    .Fields("regnum") = txtRegistryNo.Text
                    If Trim(strGatepassNo) <> "" Then
                        .Fields("gpsnum") = CLng(Trim(strGatepassNo))
                    End If
                    .Fields("sysdte") = gzGetSysDate
                    .Fields("userid") = zCurrentUser
                    .Update
                End With
                intSavingCntr = intSavingCntr + 1
            Loop
        End If
        
        MsgBox "Saving successful!", vbInformation, "Cargo Manifest"
    End If
    blnNew = False
    Exit Sub

'Trap:
'    For Each errBilling In gcnnBilling.Errors
'        With errBilling
'            lsErrStr = "Saving Error:" & .Description
'        End With
'        MsgBox lsErrStr, vbCritical, "Error"
'    Next
End Sub

Private Function Delete_ManifestDetail(ByVal strBillno As String) As Boolean
    Dim rstDltMDet As New ADODB.Recordset
On Error GoTo Here:
    rstDltMDet.Open "DELETE FROM CargoMDet WHERE bilnum='" & Trim(strBillno) & "'", gcnnBilling, adOpenDynamic, adLockOptimistic
    Delete_ManifestDetail = True
    Set rstDltMDet = Nothing
    Exit Function
Here:
    Delete_ManifestDetail = False
End Function

Private Sub cmdSearch_Click()
    rstManifestHead.MoveFirst
    If rstManifestHead.BOF Then Exit Sub
    Do While Not rstManifestHead.EOF
        If Trim(txtBillNo.Text) = Trim(rstManifestHead.Fields("bilnum")) Then
            flxgrdConNo.Clear
            flxgrdConNo.Refresh
            Call Format_flxgrdConNo
            Call InitializeDetails
            Call RetrieveRecords
            Call Count_Con_Sze
            Exit Sub
        End If
        rstManifestHead.MoveNext
    Loop
    MsgBox "Bill No. " & Trim(txtBillNo.Text) & " does not exist!", vbInformation, "Cargo Manifest"
    Call cmdFirst_Click
End Sub

Private Sub flxgrdConNo_Click()
    intRowNo = 0
    txtConNo.Text = flxgrdConNo.TextMatrix(flxgrdConNo.RowSel, 0)
    txtSealNo.Text = flxgrdConNo.TextMatrix(flxgrdConNo.RowSel, 1)
    cmbType.Text = flxgrdConNo.TextMatrix(flxgrdConNo.RowSel, 2)
    cmbSize.Text = flxgrdConNo.TextMatrix(flxgrdConNo.RowSel, 3)
    txtWeight.Text = flxgrdConNo.TextMatrix(flxgrdConNo.RowSel, 4)
    cmbFullEmpty.Text = flxgrdConNo.TextMatrix(flxgrdConNo.RowSel, 5)
    If flxgrdConNo.TextMatrix(flxgrdConNo.RowSel, 6) <> "" Then
        txtConNo.Enabled = False
        txtSealNo.Enabled = False
        txtWeight.Enabled = False
        cmbType.Enabled = False
        cmbSize.Enabled = False
        cmbFullEmpty.Enabled = False
    Else
        txtConNo.Enabled = True
        txtSealNo.Enabled = True
        txtWeight.Enabled = True
        cmbType.Enabled = True
        cmbSize.Enabled = True
        cmbFullEmpty.Enabled = True
    End If
    intRowNo = flxgrdConNo.RowSel
    blnFlag = True
End Sub

Private Sub Form_Activate()
    blnAlternateColor = False
    blnNew = False
    txtBillNo.SetFocus
    intCounter = 0
    blnFlag = False
    Call InitializeControls
    
    Set rstManifestHead = New ADODB.Recordset
    Set rstManifestDet = New ADODB.Recordset
    
    rstManifestHead.Open "CargoMHead", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
    rstManifestDet.Open "CargoMDet", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
    
    If rstManifestHead.BOF Then
        Call cmdNew_Click
    Else
        Call cmdFirst_Click
    End If
End Sub

Private Sub Format_flxgrdConNo()
    With flxgrdConNo
        .Rows = 2
        .Cols = 7
        .RowHeight(0) = 300
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1
        .Row = 0:
        .Col = 0: .Text = "Container No.": .FontWidth = 7:  .ColWidth(0) = 2000
        .Col = 1: .Text = "Seal No.": .FontWidth = 7: .ColWidth(1) = 1500
        .Col = 2: .Text = "Type": .FontWidth = 7:  .ColWidth(2) = 700
        .Col = 3: .Text = "Size": .FontWidth = 7: .ColWidth(3) = 700
        .Col = 4: .Text = "Weight": .FontWidth = 7: .ColWidth(4) = 1500
        .Col = 5: .Text = "Full/Empty": .FontWidth = 7: .ColWidth(5) = 1500
        .Col = 6: .Text = "Gatepass No.": .FontWidth = 7: .ColWidth(6) = 2000
    End With
End Sub

Private Sub InitializeDetails()
    intCounter = 0
    blnFlag = False
    txtConNo.Text = ""
    txtSealNo.Text = ""
    txtWeight.Text = ""
    If blnNew = True Then
        cmbType.Text = "[Select type]"
        cmbSize.Text = "[Select size]"
        cmbFullEmpty.Text = "FCL"
    Else
        cmbType.Text = ""
        cmbSize.Text = ""
        cmbFullEmpty.Text = ""
    End If
End Sub

Private Sub InitializeControls()
    txtBillNo.SetFocus
    intCounter = 0
    blnFlag = False
    txtBillNo.Text = ""
    txtVesselName.Text = ""
    txtPO.Text = ""
    txtRegistryNo.Text = ""
    txtVoyageNo.Text = ""
    txtConsignee.Text = ""
    txtBroker.Text = ""
    mskArrivalDte.Text = "____-__-__"
    mskLastDischarge.Text = "____-__-__"
    txtConDescr.Text = ""
    txtConNo.Text = ""
    txtSealNo.Text = ""
    If blnNew = True Then
        cmbType.Text = "[Select type]"
        cmbSize.Text = "[Select size]"
        cmbFullEmpty.Text = "FCL"
    Else
        cmbType.Text = ""
        cmbSize.Text = ""
        cmbFullEmpty.Text = ""
    End If
    lbl20.Caption = ""
    lbl40.Caption = ""
    lbl45.Caption = ""
    lblTotalCon.Caption = ""
    flxgrdConNo.Clear
    flxgrdConNo.Refresh
    Call Format_flxgrdConNo
    
    'Populate Shipping Line Selections
    If Not gcnnBilling Is Nothing Then
     Call Populate_Carriers
    End If
End Sub

Private Sub Populate_Carriers()
    
    Set rstCarriers = New ADODB.Recordset
    
    rstCarriers.Open "SELECT cuscde,cusnam FROM Customer ORDER BY cusnam", gcnnBilling, adOpenDynamic
    
    If Not rstCarriers.BOF Then
        With cmbCarCde
            .Clear
            .Text = "[Select Carrier]"
            .AddItem "[Select Carrier]"
            rstCarriers.MoveFirst
            Do While Not rstCarriers.EOF
                .AddItem rstCarriers.Fields("cuscde") & " | " & rstCarriers.Fields("cusnam")
                rstCarriers.MoveNext
            Loop
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstManifestHead = Nothing
    Set rstManifestDet = Nothing
End Sub

Private Sub MaskEdBox2_Change()

End Sub

Private Sub mnuAddNewDet_Click()
    txtConNo.Enabled = True
    txtSealNo.Enabled = True
    cmbType.Enabled = True
    cmbSize.Enabled = True
    txtConNo.SetFocus
    intCounter = flxgrdConNo.Rows - 1
    blnFlag = False
    txtConNo.Text = ""
    txtSealNo.Text = ""
    cmbType.Text = "[Select type]"
    cmbSize.Text = "[Select size]"
End Sub

Private Sub mnuClear_Click()
    Call InitializeControls
End Sub

Private Sub mnuClearDetails_Click()
    txtConNo.SetFocus
    intCounter = 0
    blnFlag = False
    txtConNo.Text = ""
    txtSealNo.Text = ""
    cmbType.Text = "[Select type]"
    cmbSize.Text = "[Select size]"
    lbl20.Caption = ""
    lbl40.Caption = ""
    lbl45.Caption = ""
    lblTotalCon.Caption = ""
    flxgrdConNo.Clear
    flxgrdConNo.Refresh
    Call Format_flxgrdConNo
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuReport_Click()
    frmReport.Show 1
End Sub

Private Sub mnuSave_Click()
    Call cmdSave_Click
End Sub

Private Sub mnuUtilities_Click()

End Sub

Private Sub mskLastDischarge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        mskArrivalDte.SetFocus
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        txtConDescr.SetFocus
    End If
End Sub

Private Sub txtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        If Trim(txtBillNo.Text) <> "" Then
            If blnNew = True Then Call Chk_BL_If_Exist
        End If
    End If
End Sub

Private Function Chk_BL_If_Exist()
    Dim rstChkBill As ADODB.Recordset
    
    Set rstChkBill = New ADODB.Recordset
        
    rstChkBill.Open "SELECT bilnum FROM CargoMHead WHERE bilnum='" & Trim(txtBillNo.Text) & "'", gcnnBilling, adOpenForwardOnly, adLockReadOnly
    If rstChkBill.BOF Then
        cmbCarCde.SetFocus
    Else
        MsgBox "This B.L. already exist!", vbInformation, "Cargo Manifest"
        txtBillNo.SetFocus
    End If
    Set rstChkBill = Nothing
End Function

Private Sub txtBillNo_LostFocus()
    If Trim(txtBillNo.Text) <> "" Then
        If blnNew = True Then Chk_BL_If_Exist
    End If
End Sub

Private Sub txtCarrier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtBillNo.SetFocus
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        txtVesselName.SetFocus
    End If
End Sub

Private Sub txtPO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtVesselName.SetFocus
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        txtRegistryNo.SetFocus
    End If
End Sub

Private Sub txtVesselName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        cmbCarCde.SetFocus
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        txtPO.SetFocus
    End If
End Sub

Private Sub txtRegistryNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtPO.SetFocus
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        txtVoyageNo.SetFocus
    End If
End Sub

Private Sub txtVoyageNo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyUp Then
        txtRegistryNo.SetFocus
    End If
     If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        txtConsignee.SetFocus
    End If
End Sub

Private Sub txtConsignee_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtVoyageNo.SetFocus
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        txtBroker.SetFocus
    End If
End Sub

Private Sub txtBroker_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtConsignee.SetFocus
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        mskArrivalDte.SetFocus
    End If
End Sub

Private Sub mskArrivalDte_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtBroker.SetFocus
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        mskLastDischarge.SetFocus
    End If
End Sub

Private Sub txtCondescr_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        txtConNo.SetFocus
    End If
End Sub

Private Sub txtConNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtConDescr.SetFocus
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        If Chk_ConNo = True Then
            txtConNo.SetFocus
        Else
            txtSealNo.SetFocus
        End If
    End If
End Sub

Private Function Chk_ConNo() As Boolean
    Dim intChkCntr As Integer
    
    intChkCntr = 0
    If flxgrdConNo.Rows >= 2 Then
        Do While flxgrdConNo.Rows > intChkCntr + 1
            flxgrdConNo.Row = intChkCntr + 1
            If txtConNo.Text = Trim(flxgrdConNo.TextMatrix(flxgrdConNo.Row, 0)) Then
                MsgBox "Container No. " & Trim(txtConNo.Text) & " already exist!", vbInformation, "Cargo Maitenance"
                Chk_ConNo = True
                Exit Function
            End If
            intChkCntr = intChkCntr + 1
        Loop
        Chk_ConNo = False
    End If
End Function

Private Sub txtSealNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        txtConNo.SetFocus
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        cmbType.SetFocus
    End If
End Sub

Private Sub cmbType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtWeight.SetFocus
    End If
End Sub

Private Sub flxgrdConNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        cmbSize.SetFocus
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        cmdSave.SetFocus
    End If
End Sub

Private Sub cmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        flxgrdConNo.SetFocus
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        txtBillNo.SetFocus
    End If
End Sub

Private Sub txtConNo_LostFocus()
    If blnFlag = False Then Exit Sub
    With flxgrdConNo
        .Row = Trim(intRowNo)
        .Col = 0
        .Text = txtConNo.Text
    End With
End Sub

Private Sub txtSealNo_LostFocus()
    If blnFlag = False Then Exit Sub
    With flxgrdConNo
        .Row = Trim(intRowNo)
        .Col = 1
        .Text = txtSealNo.Text
    End With
End Sub

Private Sub cmbSize_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        If txtConNo.Text = "" Or txtSealNo.Text = "" Or cmbType.Text = "" Or cmbSize.Text = "" Then
            MsgBox "Please provide valid entries!", vbExclamation, "Cargo Manifest"
            txtConNo.SetFocus
            Exit Sub
        End If
        With flxgrdConNo
            If blnNew = True Then
                If intCounter > 0 Then
                    .AddItem ""
                End If
                .Row = intCounter + 1
            Else
                .Row = .RowSel
                intCounter = .RowSel - 1
            End If
            If blnAlternateColor = False Then
                blnAlternateColor = True
            Else
                blnAlternateColor = False
            End If
            .Col = 0: .Text = txtConNo.Text
            Call Alternate_BckColor(intCounter + 1, 0)
            .Col = 1: .Text = txtSealNo.Text
            Call Alternate_BckColor(intCounter + 1, 1)
            .Col = 2: .Text = cmbType.Text
            Call Alternate_BckColor(intCounter + 1, 2)
            .Col = 3: .Text = cmbSize.Text
            Call Alternate_BckColor(intCounter + 1, 3)
            .Col = 4: .Text = txtWeight.Text
            Call Alternate_BckColor(intCounter + 1, 4)
            .Col = 5: .Text = cmbFullEmpty.Text
            Call Alternate_BckColor(intCounter + 1, 5)
            .Col = 6
            Call Alternate_BckColor(intCounter + 1, 6)
            If blnNew = True Then
                txtConNo.Text = ""
                txtSealNo.Text = ""
                txtWeight.Text = ""
                cmbFullEmpty.Text = "FCL"
                cmbType.Text = "[Select type]"
                cmbSize.Text = "[Select size]"
                intCounter = intCounter + 1
            End If
        End With
        Call Count_Con_Sze
        txtConNo.SetFocus
    End If
End Sub

Private Sub cmbSize_LostFocus()
    If blnFlag = False Then Exit Sub
    With flxgrdConNo
        .Row = Trim(intRowNo)
        .Col = 3
        .Text = cmbSize.Text
    End With
End Sub

Private Sub cmbType_LostFocus()
    If blnFlag = False Then Exit Sub
    With flxgrdConNo
        .Row = Trim(intRowNo)
        .Col = 2
        .Text = cmbType.Text
    End With
End Sub

Private Sub RetrieveRecords()
    Dim rstManifest2 As New ADODB.Recordset
    Dim intCntr As Integer
    Dim strDate As String
    
    cmdNew.Enabled = True
    cmdDelete.Enabled = True
    cmdSearch.Enabled = True
    cmbType.Text = ""
    cmbSize.Text = ""
    blnNew = False
    
    strDate = ""
    'Get header data
    With rstManifestHead
        txtBillNo.Text = Trim(.Fields("bilnum"))
        'Get Carrier Name
        If Not rstCarriers Is Nothing Then
            With rstCarriers
                .MoveFirst
                .Find "cuscde=" & Trim(rstManifestHead.Fields("carcde"))
                cmbCarCde.Text = rstCarriers.Fields("cuscde") & " | " & rstCarriers.Fields("cusnam")
            End With
        End If
        txtPO.Text = IIf(IsNull(.Fields("po")), "", Trim(.Fields("po")))
        txtVesselName.Text = Trim(.Fields("vslname"))
        txtRegistryNo.Text = Trim(.Fields("regnum"))
        txtVoyageNo.Text = Trim(.Fields("voynum"))
        txtConsignee.Text = Trim(.Fields("consignee"))
        If Not IsNull(Trim(.Fields("broker"))) Then
            txtBroker.Text = Trim(.Fields("broker"))
        Else
            txtBroker.Text = ""
        End If
        strDate = ""
        If Not IsNull(Trim(.Fields("arvdte"))) Then
            strDate = Format(Trim(.Fields("arvdte")), "ddmmyyyy")
            mskArrivalDte.SelText = Right(Trim(strDate), 4) & Mid(Trim(strDate), 3, 2) & Left(Trim(strDate), 2)
        Else
            mskArrivalDte.Text = "____-__-__"
        End If
        strDate = ""
        If Not IsNull(Trim(.Fields("dischargedte"))) Then
            strDate = Format(Trim(.Fields("dischargedte")), "ddmmyyyy")
            mskLastDischarge.SelText = Right(Trim(strDate), 4) & Mid(Trim(strDate), 3, 2) & Left(Trim(strDate), 2)
        Else
            mskLastDischarge.Text = "____-__-__"
        End If
        txtConDescr.Text = Trim(.Fields("ctnNameDesc"))
    End With
    
    'Get detail data
    intCntr = 0
    rstManifest2.Open "SELECT * FROM CargoMDet WHERE bilnum='" & Trim(txtBillNo.Text) & "'", gcnnBilling, adOpenDynamic, adLockOptimistic
    If Not rstManifest2.BOF Then
        Do
            With flxgrdConNo
                If intCntr > 0 Then
                    .AddItem ""
                End If
                .Row = intCntr + 1
                If blnAlternateColor = False Then
                    blnAlternateColor = True
                Else
                    blnAlternateColor = False
                End If
                .Col = 0: .Text = rstManifest2.Fields("ctnnum")
                Call Alternate_BckColor(intCntr + 1, 0)
                .Col = 1: .Text = IIf(IsNull(rstManifest2.Fields("silnum")), "", rstManifest2.Fields("silnum"))
                Call Alternate_BckColor(intCntr + 1, 1)
                .Col = 2: .Text = rstManifest2.Fields("ctntype")
                Call Alternate_BckColor(intCntr + 1, 2)
                .Col = 3: .Text = rstManifest2.Fields("ctnsze")
                Call Alternate_BckColor(intCntr + 1, 3)
                .Col = 4: .Text = IIf(IsNull(rstManifest2.Fields("ctnweight")), "", rstManifest2.Fields("ctnweight"))
                Call Alternate_BckColor(intCntr + 1, 4)
                .Col = 5: .Text = IIf(IsNull(rstManifest2.Fields("fullempty")), "", rstManifest2.Fields("fullempty"))
                Call Alternate_BckColor(intCntr + 1, 5)
                .Col = 6
                If Not IsNull(rstManifest2.Fields("gpsnum")) Then
                    .Text = rstManifest2.Fields("gpsnum")
                End If
                Call Alternate_BckColor(intCntr + 1, 6)
            End With
            intCntr = intCntr + 1
            rstManifest2.MoveNext
        Loop Until rstManifest2.EOF
        Call Count_Con_Sze
    End If
End Sub

Private Sub Alternate_BckColor(ByVal intRow As Integer, ByVal intCol As Integer)
    flxgrdConNo.Row = intRow
    flxgrdConNo.Col = intCol
    If blnAlternateColor = False Then
        flxgrdConNo.CellBackColor = "&HFFFFFF"
    Else
        flxgrdConNo.CellBackColor = "&HC0FFFF"
    End If
End Sub

Private Sub Count_Con_Sze()
    Dim intCountCntr As Integer
    Dim intCnt20 As Integer
    Dim intCnt40 As Integer
    Dim intCnt45 As Integer
    Dim intTotalCon As Integer
        
    intCountCntr = 0
    intCnt20 = 0
    intCnt40 = 0
    intCnt45 = 0
    intTotalCon = 0
    
    If flxgrdConNo.Rows >= 2 Then
        Do While flxgrdConNo.Rows > intCountCntr + 1
            flxgrdConNo.Row = intCountCntr + 1
            Select Case Trim(flxgrdConNo.TextMatrix(flxgrdConNo.Row, 3))
                Case 20
                    intCnt20 = intCnt20 + 1
                Case 40
                    intCnt40 = intCnt40 + 1
                Case 45
                    intCnt45 = intCnt45 + 1
            End Select
            intCountCntr = intCountCntr + 1
        Loop
    End If
    
    intTotalCon = intCnt20 + intCnt40 + intCnt45
    
    'Dislay count result
    lbl20.Caption = "  " & CStr(intCnt20) & " X 20"
    lbl40.Caption = "  " & CStr(intCnt40) & " X 40"
    lbl45.Caption = "  " & CStr(intCnt45) & " X 45"
    lblTotalCon.Caption = "  Total : " & CStr(intTotalCon)
End Sub

Private Sub txtWeight_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
        cmbType.SetFocus
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        cmbFullEmpty.SetFocus
    End If
End Sub

Private Sub txtWeight_LostFocus()
    If blnFlag = False Then Exit Sub
    With flxgrdConNo
        .Row = Trim(intRowNo)
        .Col = 4
        .Text = txtWeight.Text
    End With
End Sub
