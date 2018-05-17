VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCYSCCR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CY Special Services CCR Issuance v2"
   ClientHeight    =   10860
   ClientLeft      =   75
   ClientTop       =   660
   ClientWidth     =   15270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CYSCCR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10860
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabrecord 
      Height          =   11055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15285
      _ExtentX        =   26961
      _ExtentY        =   19500
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Data Entry"
      TabPicture(0)   =   "CYSCCR.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label55"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label29"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label17"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label46"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label56"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label58"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label65"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label66"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label67"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label68"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label71"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label76"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label77"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label25"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "grdCCRTran"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "tabTran"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "grdCCRDtls"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "fraPayment"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "frmExporter"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "fraCustomer"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "fraControl"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "fraRemarks"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Frame1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Frame4"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "Payment Detail"
      TabPicture(1)   =   "CYSCCR.frx":0166
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblServer"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame4 
         Caption         =   "Company Code"
         ForeColor       =   &H00004080&
         Height          =   855
         Left            =   120
         TabIndex        =   208
         Top             =   8400
         Width           =   4080
         Begin VB.ComboBox cmbCompCode 
            Height          =   405
            ItemData        =   "CYSCCR.frx":0182
            Left            =   240
            List            =   "CYSCCR.frx":018C
            TabIndex        =   122
            Top             =   360
            Width           =   3735
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Initial CCR Number "
         ForeColor       =   &H00004080&
         Height          =   1680
         Left            =   4320
         TabIndex        =   41
         Top             =   8400
         Width           =   4365
         Begin VB.TextBox txtCCRNumberISI 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   405
            Left            =   240
            TabIndex        =   210
            Text            =   "CCR Number"
            Top             =   1200
            Width           =   3855
         End
         Begin VB.TextBox txtCCRNumber 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   405
            Left            =   240
            TabIndex        =   42
            Text            =   "CCR Number"
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label Label84 
            Caption         =   "SBITC:"
            Height          =   255
            Left            =   240
            TabIndex        =   212
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label83 
            Caption         =   "ISI:"
            Height          =   255
            Left            =   240
            TabIndex        =   211
            Top             =   960
            Width           =   495
         End
      End
      Begin VB.Frame fraRemarks 
         Caption         =   " Remarks "
         ForeColor       =   &H00004080&
         Height          =   915
         Left            =   120
         TabIndex        =   40
         Top             =   7440
         Width           =   8520
         Begin VB.TextBox txtRemark 
            Height          =   420
            Left            =   240
            MaxLength       =   50
            TabIndex        =   121
            Top             =   300
            Width           =   8100
         End
      End
      Begin VB.Frame fraControl 
         Height          =   705
         Left            =   120
         TabIndex        =   39
         Top             =   9240
         Width           =   4050
         Begin VB.CheckBox chkNewCCR 
            Appearance      =   0  'Flat
            Caption         =   "&Add To Detail"
            ForeColor       =   &H80000008&
            Height          =   390
            Left            =   960
            TabIndex        =   123
            Top             =   240
            Value           =   1  'Checked
            Width           =   1875
         End
      End
      Begin VB.Frame fraCustomer 
         Caption         =   " Customer-Broker"
         ForeColor       =   &H00004080&
         Height          =   1335
         Left            =   120
         TabIndex        =   31
         Top             =   6120
         Width           =   8520
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8520
            TabIndex        =   37
            Text            =   "Text2"
            Top             =   720
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8520
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   240
            Width           =   2460
         End
         Begin VB.CheckBox chkVAT 
            Appearance      =   0  'Flat
            Caption         =   "&VAT"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   240
            TabIndex        =   35
            Top             =   960
            Value           =   1  'Checked
            Width           =   840
         End
         Begin VB.CheckBox chkWTax 
            Appearance      =   0  'Flat
            Caption         =   "&W/Tax"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3480
            TabIndex        =   34
            Top             =   960
            Width           =   1290
         End
         Begin VB.CheckBox chkGuarantee 
            Appearance      =   0  'Flat
            Caption         =   "&Under Guarantee"
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   5880
            TabIndex        =   33
            Top             =   960
            Width           =   2415
         End
         Begin VB.ComboBox cboVATRate 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   390
            ItemData        =   "CYSCCR.frx":019C
            Left            =   1200
            List            =   "CYSCCR.frx":01A6
            TabIndex        =   32
            Top             =   840
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtCusName 
            Height          =   465
            Left            =   120
            TabIndex        =   120
            Top             =   360
            Width           =   8220
            _ExtentX        =   14499
            _ExtentY        =   820
            _Version        =   393216
            BackColor       =   16777215
            AutoTab         =   -1  'True
            MaxLength       =   50
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
      End
      Begin VB.Frame frmExporter 
         Caption         =   "Exporter-Consignee"
         ForeColor       =   &H00004080&
         Height          =   915
         Left            =   120
         TabIndex        =   30
         Top             =   5160
         Width           =   8520
         Begin VB.TextBox txtImporter 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   120
            MaxLength       =   50
            TabIndex        =   119
            Top             =   360
            Width           =   8220
         End
      End
      Begin VB.Frame fraPayment 
         Enabled         =   0   'False
         Height          =   6255
         Left            =   8760
         TabIndex        =   1
         Top             =   4560
         Width           =   6405
         Begin VB.TextBox txtLog 
            Height          =   735
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   2
            Text            =   "CYSCCR.frx":01B9
            Top             =   5280
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save|Print"
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
            Height          =   840
            Left            =   4320
            TabIndex        =   20
            Top             =   5280
            Width           =   1890
         End
         Begin MSMask.MaskEdBox txtCshAmt 
            Height          =   390
            Left            =   2287
            TabIndex        =   3
            Top             =   675
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   688
            _Version        =   393216
            ForeColor       =   16711680
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            TabIndex        =   5
            Top             =   2400
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            TabIndex        =   8
            Top             =   2850
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            TabIndex        =   11
            Top             =   3300
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            TabIndex        =   14
            Top             =   3750
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            TabIndex        =   17
            Top             =   4200
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            Left            =   4410
            TabIndex        =   7
            Top             =   2400
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            Left            =   4410
            TabIndex        =   10
            Top             =   2850
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            Left            =   4410
            TabIndex        =   13
            Top             =   3300
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            Left            =   4410
            TabIndex        =   16
            Top             =   3750
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            Left            =   4410
            TabIndex        =   19
            Top             =   4200
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            Left            =   2280
            TabIndex        =   6
            Top             =   2400
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            Left            =   2280
            TabIndex        =   9
            Top             =   2850
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            Left            =   2280
            TabIndex        =   12
            Top             =   3300
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            Left            =   2280
            TabIndex        =   15
            Top             =   3750
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            Left            =   2280
            TabIndex        =   18
            Top             =   4200
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   688
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtADRAmt 
            Height          =   390
            Left            =   2280
            TabIndex        =   4
            Top             =   1560
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   688
            _Version        =   393216
            ForeColor       =   16711680
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0.00"
            PromptChar      =   " "
         End
         Begin VB.Label Label82 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ADR AMT"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   2280
            TabIndex        =   209
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label45 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CHANGE"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   4440
            TabIndex        =   29
            Top             =   300
            Width           =   1845
         End
         Begin VB.Label Label50 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "AMOUNT DUE"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   150
            TabIndex        =   28
            Top             =   300
            Width           =   2070
         End
         Begin VB.Label Label48 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CASH AMT"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   2280
            TabIndex        =   27
            Top             =   300
            Width           =   2055
         End
         Begin VB.Label Label34 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CHECK AMT"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   150
            TabIndex        =   26
            Top             =   2025
            Width           =   2070
         End
         Begin VB.Label Label52 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CHECK BANK"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   4410
            TabIndex        =   25
            Top             =   2025
            Width           =   1845
         End
         Begin VB.Label Label53 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CHECK NO"
            ForeColor       =   &H00004080&
            Height          =   315
            Left            =   2280
            TabIndex        =   24
            Top             =   2025
            Width           =   2055
         End
         Begin VB.Label lblChkTot 
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
            Height          =   390
            Left            =   150
            TabIndex        =   23
            Top             =   4650
            Width           =   2190
         End
         Begin VB.Label lblChange 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   4440
            TabIndex        =   22
            Top             =   675
            Width           =   1845
         End
         Begin VB.Label lblAmtDue 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   120
            TabIndex        =   21
            Top             =   675
            Width           =   2070
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdCCRDtls 
         Height          =   1680
         Left            =   8760
         TabIndex        =   38
         Top             =   495
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2963
         _Version        =   393216
         Rows            =   25
         Cols            =   8
         FixedCols       =   0
         ForeColorFixed  =   16512
         BackColorSel    =   65535
         ForeColorSel    =   0
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "CONTAINER No.     | SIZE |  ENTRY No. | REGISTRY No.  | TOTAL  | LENGTH | WIDTH | HEIGHT"
      End
      Begin TabDlg.SSTab tabTran 
         Height          =   4980
         Left            =   120
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   120
         Width           =   8520
         _ExtentX        =   15028
         _ExtentY        =   8784
         _Version        =   393216
         Tabs            =   6
         Tab             =   4
         TabsPerRow      =   6
         TabHeight       =   706
         ForeColor       =   16512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "ARR"
         TabPicture(0)   =   "CYSCCR.frx":01BD
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cboDanger"
         Tab(0).Control(1)=   "Frame6"
         Tab(0).Control(2)=   "Frame7"
         Tab(0).Control(3)=   "Frame5"
         Tab(0).Control(4)=   "Label73"
         Tab(0).Control(5)=   "lblArrPrevAmt"
         Tab(0).Control(6)=   "Label60"
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "STO"
         TabPicture(1)   =   "CYSCCR.frx":01D9
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblStoPluginDate"
         Tab(1).Control(1)=   "Label81"
         Tab(1).Control(2)=   "lblStoPrevPay"
         Tab(1).Control(3)=   "Label75"
         Tab(1).Control(4)=   "Label57"
         Tab(1).Control(5)=   "Label13"
         Tab(1).Control(6)=   "Label49"
         Tab(1).Control(7)=   "Label47"
         Tab(1).Control(8)=   "lblStoEntryNo"
         Tab(1).Control(9)=   "lblStoRegNo"
         Tab(1).Control(10)=   "lblStoCCRNo"
         Tab(1).Control(11)=   "Label18"
         Tab(1).Control(12)=   "lblStoValidUntil"
         Tab(1).Control(13)=   "Label12"
         Tab(1).Control(14)=   "mskExpStorageIN"
         Tab(1).Control(15)=   "txtStoExtDate"
         Tab(1).Control(16)=   "txtSTOCCRNo"
         Tab(1).Control(17)=   "txtStoContSz"
         Tab(1).Control(18)=   "txtStoContNo"
         Tab(1).Control(19)=   "z"
         Tab(1).Control(20)=   "Frame2"
         Tab(1).ControlCount=   21
         TabCaption(2)   =   "RFR"
         TabPicture(2)   =   "CYSCCR.frx":01F5
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lblRfrHrs"
         Tab(2).Control(1)=   "Label78"
         Tab(2).Control(2)=   "lblRfrPrevPay"
         Tab(2).Control(3)=   "Label54"
         Tab(2).Control(4)=   "Label37"
         Tab(2).Control(5)=   "Label27"
         Tab(2).Control(6)=   "lblRfrValidUntil"
         Tab(2).Control(7)=   "Label10"
         Tab(2).Control(8)=   "Label14"
         Tab(2).Control(9)=   "Label15"
         Tab(2).Control(10)=   "Label35"
         Tab(2).Control(11)=   "txtRfrPlugInDate"
         Tab(2).Control(12)=   "txtRfrRegNo"
         Tab(2).Control(13)=   "txtRfrEntryNo"
         Tab(2).Control(14)=   "txtRfrContSz"
         Tab(2).Control(15)=   "txtRfrExtDate"
         Tab(2).Control(16)=   "txtRfrContNo"
         Tab(2).Control(17)=   "Frame8"
         Tab(2).ControlCount=   18
         TabCaption(3)   =   "SOC"
         TabPicture(3)   =   "CYSCCR.frx":0211
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lblSOContSz"
         Tab(3).Control(1)=   "lblSOFulEmp"
         Tab(3).Control(2)=   "Label64"
         Tab(3).Control(3)=   "lblManifest(0)"
         Tab(3).Control(4)=   "lblManifest(11)"
         Tab(3).Control(5)=   "Label39"
         Tab(3).Control(6)=   "Label38"
         Tab(3).Control(7)=   "Label21"
         Tab(3).Control(8)=   "Label19"
         Tab(3).Control(9)=   "mskStoEndDate"
         Tab(3).Control(10)=   "mskStoStrtDate"
         Tab(3).Control(11)=   "txtSOCCRNo"
         Tab(3).Control(12)=   "txtSOContNo"
         Tab(3).Control(13)=   "txtSOVessel"
         Tab(3).ControlCount=   14
         TabCaption(4)   =   "MSC"
         TabPicture(4)   =   "CYSCCR.frx":022D
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "Label36"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).Control(1)=   "lblMscAmount"
         Tab(4).Control(1).Enabled=   0   'False
         Tab(4).Control(2)=   "lblMscRateUOM"
         Tab(4).Control(2).Enabled=   0   'False
         Tab(4).Control(3)=   "Label28"
         Tab(4).Control(3).Enabled=   0   'False
         Tab(4).Control(4)=   "lblMscRateAmt"
         Tab(4).Control(4).Enabled=   0   'False
         Tab(4).Control(5)=   "Label26"
         Tab(4).Control(5).Enabled=   0   'False
         Tab(4).Control(6)=   "Label11"
         Tab(4).Control(6).Enabled=   0   'False
         Tab(4).Control(7)=   "lblMScRateDesc"
         Tab(4).Control(7).Enabled=   0   'False
         Tab(4).Control(8)=   "Label5"
         Tab(4).Control(8).Enabled=   0   'False
         Tab(4).Control(9)=   "Label4"
         Tab(4).Control(9).Enabled=   0   'False
         Tab(4).Control(10)=   "Label16"
         Tab(4).Control(10).Enabled=   0   'False
         Tab(4).Control(11)=   "txtMscRateCode"
         Tab(4).Control(11).Enabled=   0   'False
         Tab(4).Control(12)=   "txtMscCCRNo"
         Tab(4).Control(12).Enabled=   0   'False
         Tab(4).Control(13)=   "txtMscQty"
         Tab(4).Control(13).Enabled=   0   'False
         Tab(4).Control(14)=   "txtMscContSz"
         Tab(4).Control(14).Enabled=   0   'False
         Tab(4).Control(15)=   "txtMscContNo"
         Tab(4).Control(15).Enabled=   0   'False
         Tab(4).ControlCount=   16
         TabCaption(5)   =   "OTH"
         TabPicture(5)   =   "CYSCCR.frx":0249
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame3"
         Tab(5).Control(1)=   "txtOthVessel"
         Tab(5).Control(2)=   "txtOthContSz"
         Tab(5).Control(3)=   "txtOthContNo"
         Tab(5).Control(4)=   "txtOTHCCRNo"
         Tab(5).Control(5)=   "txtOthAmount"
         Tab(5).Control(6)=   "txtOthEntryNo"
         Tab(5).Control(7)=   "txtOthRegNo"
         Tab(5).Control(8)=   "txtOthFulEmp"
         Tab(5).Control(9)=   "Label42"
         Tab(5).Control(10)=   "Label63"
         Tab(5).Control(11)=   "Label7"
         Tab(5).Control(12)=   "Label20"
         Tab(5).Control(13)=   "Label40"
         Tab(5).Control(14)=   "Label41"
         Tab(5).Control(15)=   "Label72"
         Tab(5).ControlCount=   16
         Begin VB.ComboBox cboDanger 
            Height          =   405
            Left            =   -72720
            Style           =   2  'Dropdown List
            TabIndex        =   109
            Top             =   4485
            Width           =   6075
         End
         Begin VB.Frame Frame6 
            Height          =   2940
            Left            =   -69840
            TabIndex        =   96
            Top             =   1440
            Width           =   3255
            Begin VB.TextBox txtARRUOM 
               Enabled         =   0   'False
               Height          =   375
               Left            =   1350
               MaxLength       =   1
               TabIndex        =   98
               Text            =   "I"
               Top             =   1950
               Width           =   315
            End
            Begin VB.CheckBox chkARROvz 
               Appearance      =   0  'Flat
               Caption         =   "Oversize?"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   1440
               TabIndex        =   97
               Top             =   225
               Width           =   1605
            End
            Begin MSMask.MaskEdBox txtARROvzLen 
               Height          =   390
               Left            =   1350
               TabIndex        =   99
               Top             =   600
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   688
               _Version        =   393216
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtARROvzWid 
               Height          =   390
               Left            =   1350
               TabIndex        =   100
               Top             =   1050
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   688
               _Version        =   393216
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtARROvzHgt 
               Height          =   390
               Left            =   1350
               TabIndex        =   101
               Top             =   1500
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   688
               _Version        =   393216
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label62 
               BackStyle       =   0  'Transparent
               Caption         =   "C/I"
               Height          =   315
               Left            =   1920
               TabIndex        =   108
               Top             =   2025
               Width           =   765
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "Length"
               Height          =   315
               Left            =   150
               TabIndex        =   107
               Top             =   675
               Width           =   1215
            End
            Begin VB.Label Label23 
               BackStyle       =   0  'Transparent
               Caption         =   "Width"
               Height          =   315
               Left            =   150
               TabIndex        =   106
               Top             =   1125
               Width           =   1215
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Height"
               Height          =   315
               Left            =   150
               TabIndex        =   105
               Top             =   1575
               Width           =   1215
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "UOM"
               Height          =   315
               Left            =   150
               TabIndex        =   104
               Top             =   2025
               Width           =   1215
            End
            Begin VB.Label Label79 
               BackStyle       =   0  'Transparent
               Caption         =   "RevTon"
               Height          =   315
               Left            =   150
               TabIndex        =   103
               Top             =   2475
               Width           =   1140
            End
            Begin VB.Label lblArrRevTon 
               BorderStyle     =   1  'Fixed Single
               Height          =   390
               Left            =   1350
               TabIndex        =   102
               Top             =   2400
               Width           =   1755
            End
         End
         Begin VB.Frame Frame7 
            Height          =   885
            Left            =   -74880
            TabIndex        =   91
            Top             =   480
            Width           =   8295
            Begin VB.OptionButton optArrImpExp 
               Appearance      =   0  'Flat
               Caption         =   "E&xport"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   1
               Left            =   2160
               TabIndex        =   93
               Top             =   360
               Width           =   1770
            End
            Begin VB.OptionButton optArrImpExp 
               Appearance      =   0  'Flat
               Caption         =   "&Import"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   0
               Left            =   360
               TabIndex        =   92
               Top             =   360
               Value           =   -1  'True
               Width           =   1740
            End
            Begin MSMask.MaskEdBox txtARRCCRNo 
               Height          =   390
               Left            =   6360
               TabIndex        =   94
               Top             =   240
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   688
               _Version        =   393216
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "########"
               PromptChar      =   " "
            End
            Begin VB.Label lblARRCCRNo 
               BackStyle       =   0  'Transparent
               Caption         =   "Gatepass No"
               Height          =   315
               Left            =   4680
               TabIndex        =   95
               Top             =   360
               Width           =   1650
            End
         End
         Begin VB.Frame Frame2 
            Height          =   765
            Left            =   -74775
            TabIndex        =   88
            Top             =   480
            Width           =   4710
            Begin VB.OptionButton optStoImpExp 
               Caption         =   "&Import"
               Height          =   315
               Index           =   0
               Left            =   360
               TabIndex        =   90
               Top             =   360
               Value           =   -1  'True
               Width           =   1740
            End
            Begin VB.OptionButton optStoImpExp 
               Appearance      =   0  'Flat
               Caption         =   "E&xport"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   1
               Left            =   2280
               TabIndex        =   89
               Top             =   360
               Width           =   1890
            End
         End
         Begin VB.Frame Frame3 
            Height          =   2970
            Left            =   -70080
            TabIndex        =   75
            Top             =   600
            Width           =   3255
            Begin VB.CheckBox chkOthOvz 
               Appearance      =   0  'Flat
               Caption         =   "Oversize?"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   840
               TabIndex        =   77
               Top             =   240
               Width           =   1965
            End
            Begin VB.TextBox txtOthUOM 
               Enabled         =   0   'False
               Height          =   375
               Left            =   1350
               MaxLength       =   1
               TabIndex        =   76
               Text            =   "I"
               Top             =   1980
               Width           =   435
            End
            Begin MSMask.MaskEdBox txtOthOvzLen 
               Height          =   390
               Left            =   1350
               TabIndex        =   78
               Top             =   630
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   688
               _Version        =   393216
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtOthOvzWid 
               Height          =   390
               Left            =   1350
               TabIndex        =   79
               Top             =   1080
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   688
               _Version        =   393216
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtOthOvzHgt 
               Height          =   390
               Left            =   1350
               TabIndex        =   80
               Top             =   1530
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   688
               _Version        =   393216
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label43 
               BackStyle       =   0  'Transparent
               Caption         =   "UOM"
               Height          =   315
               Left            =   150
               TabIndex        =   87
               Top             =   2055
               Width           =   1215
            End
            Begin VB.Label Label44 
               BackStyle       =   0  'Transparent
               Caption         =   "Height"
               Height          =   315
               Left            =   150
               TabIndex        =   86
               Top             =   1605
               Width           =   1215
            End
            Begin VB.Label Label61 
               BackStyle       =   0  'Transparent
               Caption         =   "Width"
               Height          =   315
               Left            =   150
               TabIndex        =   85
               Top             =   1155
               Width           =   1215
            End
            Begin VB.Label Label69 
               BackStyle       =   0  'Transparent
               Caption         =   "Length"
               Height          =   315
               Left            =   150
               TabIndex        =   84
               Top             =   705
               Width           =   1215
            End
            Begin VB.Label Label70 
               BackStyle       =   0  'Transparent
               Caption         =   "C/I"
               Height          =   315
               Left            =   2160
               TabIndex        =   83
               Top             =   2175
               Width           =   765
            End
            Begin VB.Label lblOthRevTon 
               BorderStyle     =   1  'Fixed Single
               Height          =   390
               Left            =   1350
               TabIndex        =   82
               Top             =   2430
               Width           =   1755
            End
            Begin VB.Label Label80 
               BackStyle       =   0  'Transparent
               Caption         =   "RevTon"
               Height          =   315
               Left            =   150
               TabIndex        =   81
               Top             =   2505
               Width           =   1140
            End
         End
         Begin VB.TextBox txtOthVessel 
            Height          =   405
            Left            =   -73050
            MaxLength       =   6
            TabIndex        =   74
            Top             =   3300
            Width           =   2025
         End
         Begin VB.Frame z 
            Height          =   2970
            Left            =   -69960
            TabIndex        =   61
            Top             =   480
            Width           =   3255
            Begin VB.CheckBox chkStoOvz 
               Appearance      =   0  'Flat
               Caption         =   "Oversize?"
               ForeColor       =   &H80000008&
               Height          =   315
               Left            =   720
               TabIndex        =   63
               Top             =   240
               Width           =   1965
            End
            Begin VB.TextBox txtStoUOM 
               Enabled         =   0   'False
               Height          =   375
               Left            =   1320
               MaxLength       =   1
               TabIndex        =   62
               Text            =   "C"
               Top             =   1980
               Width           =   435
            End
            Begin MSMask.MaskEdBox txtStoOvzLen 
               Height          =   390
               Left            =   1320
               TabIndex        =   64
               Top             =   630
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   688
               _Version        =   393216
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtStoOvzWid 
               Height          =   390
               Left            =   1320
               TabIndex        =   65
               Top             =   1080
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   688
               _Version        =   393216
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtStoOvzHgt 
               Height          =   390
               Left            =   1320
               TabIndex        =   66
               Top             =   1530
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   688
               _Version        =   393216
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   " "
            End
            Begin VB.Label Label30 
               BackStyle       =   0  'Transparent
               Caption         =   "UOM"
               Height          =   315
               Left            =   150
               TabIndex        =   73
               Top             =   2055
               Width           =   1215
            End
            Begin VB.Label Label31 
               BackStyle       =   0  'Transparent
               Caption         =   "Height"
               Height          =   315
               Left            =   150
               TabIndex        =   72
               Top             =   1605
               Width           =   1215
            End
            Begin VB.Label Label32 
               BackStyle       =   0  'Transparent
               Caption         =   "Width"
               Height          =   315
               Left            =   150
               TabIndex        =   71
               Top             =   1155
               Width           =   1215
            End
            Begin VB.Label Label33 
               BackStyle       =   0  'Transparent
               Caption         =   "Length"
               Height          =   315
               Left            =   150
               TabIndex        =   70
               Top             =   705
               Width           =   1215
            End
            Begin VB.Label Label74 
               BackStyle       =   0  'Transparent
               Caption         =   "C/I"
               Height          =   315
               Left            =   2040
               TabIndex        =   69
               Top             =   2055
               Width           =   765
            End
            Begin VB.Label Label51 
               BackStyle       =   0  'Transparent
               Caption         =   "RevTon"
               Height          =   315
               Left            =   150
               TabIndex        =   68
               Top             =   2505
               Width           =   1140
            End
            Begin VB.Label lblStoRevTon 
               BorderStyle     =   1  'Fixed Single
               Height          =   390
               Left            =   1320
               TabIndex        =   67
               Top             =   2430
               Width           =   1755
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   " Container"
            ForeColor       =   &H00004080&
            Height          =   2355
            Left            =   -74880
            TabIndex        =   50
            Top             =   1440
            Width           =   4560
            Begin VB.TextBox Text3 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   8280
               TabIndex        =   52
               Text            =   "Text1"
               Top             =   720
               Visible         =   0   'False
               Width           =   300
            End
            Begin VB.TextBox Text4 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   8160
               TabIndex        =   51
               Text            =   "Text2"
               Top             =   720
               Visible         =   0   'False
               Width           =   255
            End
            Begin MSMask.MaskEdBox txtARRContNo 
               Height          =   390
               Left            =   1710
               TabIndex        =   53
               Top             =   360
               Width           =   2640
               _ExtentX        =   4657
               _ExtentY        =   688
               _Version        =   393216
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   12
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">AAAAAAAAAAAA"
               PromptChar      =   " "
            End
            Begin MSMask.MaskEdBox txtARRContSz 
               Height          =   390
               Left            =   1710
               TabIndex        =   54
               Top             =   840
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   688
               _Version        =   393216
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "##"
               PromptChar      =   " "
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Container #"
               Height          =   315
               Left            =   120
               TabIndex        =   60
               Top             =   435
               Width           =   1485
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Size"
               Height          =   315
               Left            =   120
               TabIndex        =   59
               Top             =   960
               Width           =   1065
            End
            Begin VB.Label lblARREntryNo 
               BorderStyle     =   1  'Fixed Single
               Height          =   390
               Left            =   1710
               TabIndex        =   58
               Top             =   1320
               Width           =   2640
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Entry No."
               Height          =   315
               Left            =   120
               TabIndex        =   57
               Top             =   1365
               Width           =   1425
            End
            Begin VB.Label lblARRRegNo 
               BorderStyle     =   1  'Fixed Single
               Height          =   390
               Left            =   1710
               TabIndex        =   56
               Top             =   1800
               Width           =   2640
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Registry No."
               Height          =   315
               Left            =   120
               TabIndex        =   55
               Top             =   1875
               Width           =   1665
            End
         End
         Begin VB.Frame Frame8 
            Height          =   765
            Left            =   -74760
            TabIndex        =   45
            Top             =   480
            Width           =   7935
            Begin VB.OptionButton optStoImpExp 
               Appearance      =   0  'Flat
               Caption         =   "E&xport"
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   2
               Left            =   1920
               TabIndex        =   47
               Top             =   360
               Width           =   1890
            End
            Begin VB.OptionButton optStoImpExp 
               Caption         =   "&Import"
               Height          =   315
               Index           =   3
               Left            =   360
               TabIndex        =   46
               Top             =   360
               Value           =   -1  'True
               Width           =   1260
            End
            Begin MSMask.MaskEdBox txtRFRCCRNo 
               Height          =   390
               Left            =   6255
               TabIndex        =   48
               Top             =   240
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   688
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "########"
               PromptChar      =   " "
            End
            Begin VB.Label Label59 
               BackStyle       =   0  'Transparent
               Caption         =   "Gatepass No"
               Height          =   315
               Left            =   4440
               TabIndex        =   49
               Top             =   315
               Width           =   1635
            End
         End
         Begin VB.TextBox txtSOVessel 
            Enabled         =   0   'False
            Height          =   465
            Left            =   -72360
            MaxLength       =   6
            TabIndex        =   44
            Top             =   3600
            Visible         =   0   'False
            Width           =   1665
         End
         Begin MSMask.MaskEdBox txtStoContNo 
            Height          =   390
            Left            =   -72825
            TabIndex        =   110
            Top             =   1755
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   688
            _Version        =   393216
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">AAAAAAAAAAAA"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtStoContSz 
            Height          =   390
            Left            =   -72825
            TabIndex        =   111
            Top             =   2205
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtSTOCCRNo 
            Height          =   390
            Left            =   -72840
            TabIndex        =   112
            Top             =   1320
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtRfrContNo 
            Height          =   390
            Left            =   -72645
            TabIndex        =   114
            Top             =   1395
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtMscContNo 
            Height          =   390
            Left            =   2250
            TabIndex        =   115
            Top             =   1200
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">AAAAAAAAAAAA"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtMscContSz 
            Height          =   390
            Left            =   3960
            TabIndex        =   117
            Top             =   1800
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtMscQty 
            Height          =   390
            Left            =   2250
            TabIndex        =   118
            Top             =   3375
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtOthContSz 
            Height          =   390
            Left            =   -73050
            TabIndex        =   124
            Top             =   1725
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
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
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtStoExtDate 
            Height          =   390
            Left            =   -72840
            TabIndex        =   125
            Top             =   4005
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   688
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtOthContNo 
            Height          =   390
            Left            =   -73050
            TabIndex        =   126
            Top             =   1200
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">AAAAAAAAAAAA"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtOTHCCRNo 
            Height          =   390
            Left            =   -73050
            TabIndex        =   127
            Top             =   675
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtOthAmount 
            Height          =   390
            Left            =   -73080
            TabIndex        =   128
            Top             =   4080
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0.00"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtOthEntryNo 
            Height          =   390
            Left            =   -73050
            TabIndex        =   129
            Top             =   2250
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">AAAAAAAAAAAA"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtOthRegNo 
            Height          =   390
            Left            =   -73050
            TabIndex        =   130
            Top             =   2775
            Width           =   2700
            _ExtentX        =   4763
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">AAAAAAAAAAAA"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtOthFulEmp 
            Height          =   390
            Left            =   -72240
            TabIndex        =   131
            Top             =   1725
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtMscCCRNo 
            Height          =   390
            Left            =   2250
            TabIndex        =   113
            Top             =   720
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtMscRateCode 
            Height          =   390
            Left            =   2250
            TabIndex        =   116
            Top             =   1800
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">AAAAAA"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtRfrExtDate 
            Height          =   390
            Left            =   -72645
            TabIndex        =   132
            Top             =   4020
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   688
            _Version        =   393216
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtRfrContSz 
            Height          =   390
            Left            =   -68700
            TabIndex        =   133
            Top             =   1395
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
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
            Mask            =   "##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtRfrEntryNo 
            Height          =   390
            Left            =   -72645
            TabIndex        =   134
            Top             =   1920
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtRfrRegNo 
            Height          =   390
            Left            =   -72645
            TabIndex        =   135
            Top             =   2445
            Width           =   3060
            _ExtentX        =   5398
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">&&&&&&&&&&&&"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtRfrPlugInDate 
            Height          =   390
            Left            =   -72600
            TabIndex        =   136
            Top             =   2970
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   688
            _Version        =   393216
            Enabled         =   0   'False
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskExpStorageIN 
            Height          =   390
            Left            =   -72840
            TabIndex        =   137
            Top             =   3555
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   688
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtSOContNo 
            Height          =   390
            Left            =   -72360
            TabIndex        =   138
            Top             =   825
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">AAAAAAAAAAAA"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtSOCCRNo 
            Height          =   390
            Left            =   -68280
            TabIndex        =   139
            Top             =   840
            Visible         =   0   'False
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   688
            _Version        =   393216
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "########"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskStoStrtDate 
            Height          =   375
            Left            =   -72360
            TabIndex        =   140
            Top             =   2505
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-##"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox mskStoEndDate 
            Height          =   375
            Left            =   -72360
            TabIndex        =   141
            Top             =   3045
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-##"
            PromptChar      =   " "
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Freight Kind(F/E)"
            Height          =   315
            Left            =   -74760
            TabIndex        =   218
            Top             =   1890
            Width           =   2415
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Vessel"
            Enabled         =   0   'False
            Height          =   315
            Left            =   -74760
            TabIndex        =   217
            Top             =   3600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            Height          =   315
            Left            =   -74760
            TabIndex        =   216
            Top             =   1350
            Width           =   1815
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "Container"
            Height          =   315
            Left            =   -74760
            TabIndex        =   215
            Top             =   825
            Width           =   1815
         End
         Begin VB.Label lblManifest 
            Caption         =   "Storage Start Date"
            Height          =   375
            Index           =   11
            Left            =   -74760
            TabIndex        =   214
            Top             =   2415
            Width           =   2175
         End
         Begin VB.Label lblManifest 
            Caption         =   "Storage End Date"
            Height          =   375
            Index           =   0
            Left            =   -74760
            TabIndex        =   213
            Top             =   3015
            Width           =   2175
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Valid Until"
            Height          =   315
            Left            =   -74775
            TabIndex        =   190
            Top             =   3630
            Width           =   1890
         End
         Begin VB.Label Label35 
            BackStyle       =   0  'Transparent
            Caption         =   "Sz"
            Height          =   315
            Left            =   -69225
            TabIndex        =   189
            Top             =   1470
            Width           =   390
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Plug-In"
            Height          =   315
            Left            =   -74820
            TabIndex        =   188
            Top             =   3045
            Width           =   2115
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Valid Until"
            Height          =   315
            Left            =   -74820
            TabIndex        =   187
            Top             =   3570
            Width           =   2115
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Extend Until"
            Height          =   315
            Left            =   -74820
            TabIndex        =   186
            Top             =   4095
            Width           =   2115
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity"
            Height          =   315
            Left            =   225
            TabIndex        =   185
            Top             =   3450
            Width           =   1965
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Rate Code"
            Height          =   315
            Left            =   225
            TabIndex        =   184
            Top             =   1875
            Width           =   1965
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Container"
            Height          =   315
            Left            =   225
            TabIndex        =   183
            Top             =   1350
            Width           =   1845
         End
         Begin VB.Label lblMScRateDesc 
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   2250
            TabIndex        =   182
            Top             =   2325
            Width           =   6000
         End
         Begin VB.Label lblStoValidUntil 
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   -72825
            TabIndex        =   181
            Top             =   3675
            Width           =   2640
         End
         Begin VB.Label lblRfrValidUntil 
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   -72600
            TabIndex        =   180
            Top             =   3480
            Width           =   3105
         End
         Begin VB.Label Label73 
            BackStyle       =   0  'Transparent
            Caption         =   "Danger Class"
            Height          =   315
            Left            =   -74775
            TabIndex        =   179
            Top             =   4560
            Width           =   1665
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Container #"
            Height          =   315
            Left            =   -74775
            TabIndex        =   178
            Top             =   1830
            Width           =   1890
         End
         Begin VB.Label lblStoCCRNo 
            BackStyle       =   0  'Transparent
            Caption         =   "Gatepass No"
            Height          =   315
            Left            =   -74775
            TabIndex        =   177
            Top             =   1380
            Width           =   1890
         End
         Begin VB.Label lblStoRegNo 
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   -72825
            TabIndex        =   176
            Top             =   3105
            Width           =   2640
         End
         Begin VB.Label lblStoEntryNo 
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   -72825
            TabIndex        =   175
            Top             =   2655
            Width           =   2640
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "Registry No."
            Height          =   315
            Left            =   -74775
            TabIndex        =   174
            Top             =   3180
            Width           =   1665
         End
         Begin VB.Label Label49 
            BackStyle       =   0  'Transparent
            Caption         =   "Entry No."
            Height          =   315
            Left            =   -74775
            TabIndex        =   173
            Top             =   2730
            Width           =   1665
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "Container No."
            Height          =   315
            Left            =   -74820
            TabIndex        =   172
            Top             =   1470
            Width           =   2115
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "Registry No."
            Height          =   315
            Left            =   -74820
            TabIndex        =   171
            Top             =   2520
            Width           =   2115
         End
         Begin VB.Label Label54 
            BackStyle       =   0  'Transparent
            Caption         =   "Entry No."
            Height          =   315
            Left            =   -74820
            TabIndex        =   170
            Top             =   1995
            Width           =   2115
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            Height          =   315
            Left            =   -74775
            TabIndex        =   169
            Top             =   2280
            Width           =   1665
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "Extend To"
            Height          =   315
            Left            =   -74775
            TabIndex        =   168
            Top             =   4080
            Width           =   1890
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "Size / FE"
            Height          =   315
            Left            =   -74775
            TabIndex        =   167
            Top             =   1800
            Width           =   1665
         End
         Begin VB.Label Label63 
            BackStyle       =   0  'Transparent
            Caption         =   "Container No."
            Height          =   315
            Left            =   -74775
            TabIndex        =   166
            Top             =   1275
            Width           =   1665
         End
         Begin VB.Label lblArrPrevAmt 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   -73815
            TabIndex        =   165
            Top             =   3960
            Width           =   2385
         End
         Begin VB.Label Label60 
            BackStyle       =   0  'Transparent
            Caption         =   "Paid"
            Height          =   315
            Left            =   -74640
            TabIndex        =   164
            Top             =   4035
            Width           =   765
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "GPS / CCR"
            Height          =   315
            Left            =   -74775
            TabIndex        =   163
            Top             =   750
            Width           =   1665
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   315
            Left            =   -74775
            TabIndex        =   162
            Top             =   4125
            Width           =   1140
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "Registry No."
            Height          =   315
            Left            =   -74775
            TabIndex        =   161
            Top             =   2850
            Width           =   1665
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "Entry No."
            Height          =   315
            Left            =   -74775
            TabIndex        =   160
            Top             =   2325
            Width           =   1665
         End
         Begin VB.Label Label72 
            BackStyle       =   0  'Transparent
            Caption         =   "Vessel"
            Height          =   315
            Left            =   -74775
            TabIndex        =   159
            Top             =   3375
            Width           =   1665
         End
         Begin VB.Label Label75 
            BackStyle       =   0  'Transparent
            Caption         =   "Paid"
            Height          =   315
            Left            =   -70020
            TabIndex        =   158
            Top             =   3600
            Width           =   765
         End
         Begin VB.Label lblStoPrevPay 
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
            Height          =   390
            Left            =   -69360
            TabIndex        =   157
            Top             =   3525
            Width           =   2685
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Gatepass #"
            Height          =   315
            Left            =   225
            TabIndex        =   156
            Top             =   825
            Width           =   1965
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   315
            Left            =   225
            TabIndex        =   155
            Top             =   2400
            Width           =   1965
         End
         Begin VB.Label lblMscRateAmt 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   2250
            TabIndex        =   154
            Top             =   2850
            Width           =   2505
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "Rate / UMS"
            Height          =   315
            Left            =   225
            TabIndex        =   153
            Top             =   2925
            Width           =   1965
         End
         Begin VB.Label lblMscRateUOM 
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   4800
            TabIndex        =   152
            Top             =   2850
            Width           =   3360
         End
         Begin VB.Label lblMscAmount 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   2250
            TabIndex        =   151
            Top             =   3900
            Width           =   2505
         End
         Begin VB.Label Label36 
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            Height          =   315
            Left            =   225
            TabIndex        =   150
            Top             =   3975
            Width           =   1965
         End
         Begin VB.Label lblRfrPrevPay 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   -69405
            TabIndex        =   149
            Top             =   3495
            Width           =   2580
         End
         Begin VB.Label Label78 
            BackStyle       =   0  'Transparent
            Caption         =   "Paid"
            Height          =   315
            Left            =   -69240
            TabIndex        =   148
            Top             =   3045
            Width           =   840
         End
         Begin VB.Label lblRfrHrs 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   -69405
            TabIndex        =   147
            Top             =   4020
            Width           =   2580
         End
         Begin VB.Label Label64 
            BackStyle       =   0  'Transparent
            Caption         =   "CCR No."
            Enabled         =   0   'False
            Height          =   315
            Left            =   -69720
            TabIndex        =   146
            Top             =   840
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.Label lblSOFulEmp 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "E"
            Enabled         =   0   'False
            Height          =   390
            Left            =   -72360
            TabIndex        =   145
            Top             =   1935
            Width           =   540
         End
         Begin VB.Label lblSOContSz 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   -72360
            TabIndex        =   144
            Top             =   1395
            Width           =   540
         End
         Begin VB.Label Label81 
            BackStyle       =   0  'Transparent
            Caption         =   "Plug-In"
            Height          =   315
            Left            =   -74760
            TabIndex        =   143
            Top             =   4515
            Width           =   1875
         End
         Begin VB.Label lblStoPluginDate 
            BorderStyle     =   1  'Fixed Single
            Height          =   390
            Left            =   -72840
            TabIndex        =   142
            Top             =   4440
            Width           =   2610
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdCCRTran 
         Height          =   1620
         Left            =   8760
         TabIndex        =   191
         Top             =   2520
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   2858
         _Version        =   393216
         Cols            =   34
         ForeColorFixed  =   16512
         BackColorSel    =   65535
         ForeColorSel    =   -2147483640
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         Enabled         =   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "  | RATE| AMOUNT|  VAT   |TAX     | OVZ       | TOTAL     |CUSTOMER NAME| # "
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
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CCR DETAILS"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   8760
         TabIndex        =   207
         Top             =   2160
         Width           =   6405
      End
      Begin VB.Label Label77 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " PREPARE PAYMENT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3240
         TabIndex        =   206
         Top             =   10560
         Width           =   2145
      End
      Begin VB.Label Label76 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   2760
         TabIndex        =   205
         Top             =   10560
         Width           =   450
      End
      Begin VB.Label Label71 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " VOID CCR DETAIL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6600
         TabIndex        =   204
         Top             =   10200
         Width           =   1785
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   6120
         TabIndex        =   203
         Top             =   10200
         Width           =   450
      End
      Begin VB.Label Label67 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " EXIT CCR ISSUANCE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6600
         TabIndex        =   202
         Top             =   10560
         Width           =   1905
      End
      Begin VB.Label Label66 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   6120
         TabIndex        =   201
         Top             =   10560
         Width           =   450
      End
      Begin VB.Label Label65 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SELECT TRANSACTION TYPE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3240
         TabIndex        =   200
         Top             =   10200
         Width           =   2625
      End
      Begin VB.Label Label58 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   2760
         TabIndex        =   199
         Top             =   10200
         Width           =   450
      End
      Begin VB.Label Label56 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SELECT CCR DETAIL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   600
         TabIndex        =   198
         Top             =   10560
         Width           =   1905
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   197
         Top             =   10560
         Width           =   450
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CONTAINER LIST"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   8760
         TabIndex        =   196
         Top             =   135
         Width           =   6405
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   195
         Top             =   10200
         Width           =   450
      End
      Begin VB.Label Label29 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " SELECT CONTAINER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   600
         TabIndex        =   194
         Top             =   10215
         Width           =   1905
      End
      Begin VB.Label lblServer 
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   -74640
         TabIndex        =   193
         Top             =   9975
         Width           =   6135
      End
      Begin VB.Label Label55 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PAYMENT DETAILS"
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   8760
         TabIndex        =   192
         Top             =   4200
         Width           =   6405
      End
   End
   Begin VB.Menu mnuMeu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuMenuGrid 
         Caption         =   "Select from &Grid"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuGridContainer 
         Caption         =   "Select from &Grid Container"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuMenuEdit 
         Caption         =   "&Edit selected item"
         Enabled         =   0   'False
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuMenuDelete 
         Caption         =   "&Delete selected item"
         Enabled         =   0   'False
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuF1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuTab 
         Caption         =   "Select &Next Tab"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuF3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuPayment 
         Caption         =   "Prepare &Payment"
         Enabled         =   0   'False
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuF4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuSave 
         Caption         =   "&Save / Print transaction"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuMenuExit 
         Caption         =   "E&xit"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmCYSCCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Const cDataEntry As Integer = 0
Const cPayment As Integer = 1

Private Enum enGridCol
    enCounter = 0
    enRateCode = 1
    enAmount = 2
    enVATAmt = 3
    enWTaxAmt = 4
    enOvzAmt = 5
    enTotalAmt = 6
    enCCRTag = 7
    encustomer = 8
    enratedescr = 9
    enCCRNo = 10
    enContNo = 11
    enContSz = 12
    enFulEmp = 13
    enEntryNo = 14
    enRegNo = 15
    enOvzLen = 16
    enOvzWid = 17
    enOvzHgt = 18
    enOvzUom = 19
    enRevTon = 20
    enDangerCode = 21
    enStoValidUntil = 22
    enRfrValidUntil = 23
    enStoDays = 24
    enQuantity = 25
    enVessel = 26
    enDangerAmt = 27
    enRemark = 28
    enShipLine = 29
    enGuaranty = 30
    enRfrHours = 31
    enimporter = 32
    
    'PRNH
    enCompCode = 33
End Enum

Dim nDwellDays As Integer
Const cVoid = "*VOID*"
Const cEmptyRfrDate = "    -  -     :  "

Const cNullDate = #12:00:00 AM#
Const cRTon20 As Currency = 27.95
Const cRTon40 As Currency = 63.75
Const cRTon45 As Currency = 76.38

Dim vRevTon As Single
Dim vRevTonRateArr, vRevTonRateSto As Currency
Dim vRevTonRateArrExp As Currency
Dim cRateCode, vCusCodeUnderG As String
Dim strSQL As String
Dim vStoDay, vTabOn As Integer
Dim vRfrHours, vCYMStoDay As Long

Dim bNewCCR, bArrImp, bStoImp, bVAT, bWTax, bUnderG, bARROversize, bStoOversize, bOTHOversize As Boolean
Dim nPtr, nCCRCounter As Integer
Dim nAmount, nVATAmount, nWTaxAmount, nTotalAmount As Currency
Dim bEscaped As Boolean

Dim clsCTCS, clsCCRPrinter As Object

Dim arrDetails() As Variant

Dim objrpt As New CrystalDataObject.CrystalComObject
Dim rstdata As ADODB.Recordset
Dim vGetData As Variant

Dim arrOversizeAmount As Currency
Dim arrDgAmount As Currency

Private Sub cboDanger_GotFocus()
    cboDanger.BackColor = &HFFFFC0
End Sub

Private Sub cboDanger_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
        Case vbKeySpace
            SendKeys ("{F4}")
        Case vbKeyReturn
            txtImporter.SetFocus
        Case Else
    End Select
End Sub

Private Sub cboDanger_LostFocus()
    cboDanger.BackColor = vbWindowBackground
End Sub

Private Sub chkARROvz_GotFocus()
    chkARROvz.BackColor = &HFFFFC0
End Sub

Private Sub chkARROvz_LostFocus()
    chkARROvz.BackColor = vbButtonFace
End Sub

Private Sub chkGuarantee_Click()
    bUnderG = (chkGuarantee.Value = 1)
    SendKeys "{TAB}"
    If bUnderG Then Call lzCustomerUnderG
End Sub

Private Sub chkGuarantee_GotFocus()
    chkGuarantee.BackColor = &HFFFFC0
End Sub

Private Sub chkGuarantee_KeyPress(KeyAscii As Integer)
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

Private Sub chkGuarantee_LostFocus()
    On Error Resume Next
    
    chkGuarantee.BackColor = vbButtonFace
    txtRemark.SetFocus
End Sub

Private Sub chkNewCCR_Click()
    bNewCCR = (chkNewCCR.Value = 1)
End Sub

Private Sub chkNewCCR_GotFocus()
    chkNewCCR.BackColor = &HFFFFC0
    If txtMscRateCode = "WEIGHT" Then
        chkNewCCR.Value = 1
        bNewCCR = True
    End If
End Sub

Private Sub chkNewCCR_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            txtRemark.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            'PRNH
            If Trim(cmbCompCode.Text) <> "" Then
                Call lzAddTran
            Else
                MsgBox "Please choose Company Code.", vbCritical, "Miscellaneous"
            End If
        Case Else
    End Select
End Sub

Private Sub chkARROvz_Click()
    bARROversize = (chkARROvz.Value = 1)
    txtARROvzLen.Enabled = bARROversize
    txtARROvzWid.Enabled = bARROversize
    txtARROvzHgt.Enabled = bARROversize
    txtARRUOM.Enabled = bARROversize
    If Not bARROversize Then lblArrRevTon = ""
    SendKeys ("{TAB}")
End Sub

Private Sub chkARROvz_KeyPress(KeyAscii As Integer)
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

Private Sub chkNewCCR_LostFocus()
    chkNewCCR.BackColor = vbButtonFace
    cmdSave.Enabled = True
'    txtARRCCRNo.SetFocus
End Sub

Private Sub chkOthOvz_Click()
    bOTHOversize = (chkOthOvz.Value = 1)
    txtOthOvzLen.Enabled = bOTHOversize
    txtOthOvzWid.Enabled = bOTHOversize
    txtOthOvzHgt.Enabled = bOTHOversize
    txtOthUOM.Enabled = bOTHOversize
    If Not bOTHOversize Then lblOthRevTon = ""
    SendKeys ("{TAB}")
End Sub

Private Sub chkOthOvz_GotFocus()
    chkOthOvz.BackColor = &HFFFFC0
End Sub

Private Sub chkOthOvz_KeyPress(KeyAscii As Integer)
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

Private Sub chkOthOvz_LostFocus()
    chkOthOvz.BackColor = vbButtonFace
End Sub

Private Sub chkStoOvz_Click()
    bStoOversize = (chkStoOvz.Value = 1)
    txtStoOvzLen.Enabled = bStoOversize
    txtStoOvzWid.Enabled = bStoOversize
    txtStoOvzHgt.Enabled = bStoOversize
    txtStoUOM.Enabled = bStoOversize
    If Not bStoOversize Then lblStoRevTon = ""
    SendKeys ("{TAB}")
End Sub

Private Sub chkStoOvz_GotFocus()
    chkStoOvz.BackColor = &HFFFFC0
End Sub

Private Sub chkStoOvz_KeyPress(KeyAscii As Integer)
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

Private Sub chkStoOvz_LostFocus()
    chkStoOvz.BackColor = vbButtonFace
End Sub

Private Sub chkVAT_Click()
    bVAT = (chkVAT.Value = 1)
    
    'For VAT Option
    If chkVAT.Value = 1 Then
        cboVATRate.Enabled = True
    Else
        cboVATRate.Enabled = False
    End If
    
    Call lzUpdateGridVAT
    SendKeys "{TAB}"
End Sub

Private Sub chkVAT_GotFocus()
    chkVAT.BackColor = &HFFFFC0
End Sub

Private Sub chkVAT_KeyPress(KeyAscii As Integer)
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

Private Sub chkVAT_LostFocus()
    chkVAT.BackColor = vbButtonFace
End Sub

Private Sub chkWTax_Click()
    bWTax = (chkWTax.Value = 1)
    Call lzUpdateGridWTax
    SendKeys "{TAB}"
End Sub

Private Sub chkWTax_GotFocus()
    chkWTax.BackColor = &HFFFFC0
End Sub

Private Sub chkWTax_KeyPress(KeyAscii As Integer)
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

Private Sub chkWTax_LostFocus()
    chkWTax.BackColor = vbButtonFace
End Sub

Private Sub cmbCompCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
    End Select
End Sub

Private Sub cmdSave_Click()
    lzSavePrint
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
    'tabrecord.Tab = cDataEntry + 1
    If chkGuarantee.Value = 1 Then
'        'NTBSPayment1.Amount_To_Pay = 0
        lblAmtDue = 0
    End If
    
    ''NTBSPayment1.SetFocus
End If

If KeyCode = vbKeyF2 Then
    If grdCCRDtls.Row > 0 Then
           grdCCRDtls.SetFocus
           SendKeys ("{RIGHT}")
        End If
    End If
End Sub

Private Sub Form_Load()
    ' connect to CTCS
    'Set clsCTCS = CreateObject("CTCS.cCTCS")
    'If Not clsCTCS.IsConnected Then clsCTCS.Connect

    'initialize
    vTabOn = 0
    Call lzInitialize
    ' populate combo boxes
    Call lzPopulateDangerClass
    ' Get user info
    Call lzGetUserInfo
    ConnectToNavis
    ' Get rates
    vRevTonRateArr = lzGetRateInfo("RTARIM")
    vRevTonRateSto = lzGetRateInfo("RTSTIM")
    vRevTonRateArrExp = lzGetRateInfo("RTAREX")
    '
    grdCCRTran.TextMatrix(grdCCRTran.Rows - 1, enCounter) = "**"
    
gUserID = UCase(gzCurrentUser)

'With 'NTBSPayment1
'    .UserID = gUserID
'    .ServerName = Trim(gINIServer)
'    .Database = Trim(gINIDatabase)
'    .TransactionFee = 20
'    .TransactionType = "CYS"
'
'If .Connect2Server(1) = False Then
'    MsgBox "There's a problem with your connection"
'    Exit Sub
'Else
'    '.Ini_CCRNum = lzGetNextCCR(gUserID)
'    .eCustomerCode = ""
'    .eCustomerName = ""
'    .CYImport_SpecialGatepass = False
'    .PaymentCustomerCode = ""
'    .PaymentCustomerName = ""
'    .CFS_CCRType = "D"
'    .Has_ePayment = False
'    .Has_ADRPayment = True
'    .Has_POSPayment = False
'    .Has_BankFundTransfer = False
'End If

'Get Windows System DIrectory
gSysDirectory = GetWindowsSystemDirectory
'----------------------------------------

lblServer = "Server : " & Trim(gINIServer)
tabrecord.Tab = 0

'Added Navis Project Team 10/28/2009
Call GetSparcsN4Host

'Added by Navis Project Team 11/04/2009
chkVAT.Visible = False
cboVATRate.Visible = False
chkWTax.Visible = False

DoEvents
'End With
End Sub

Private Sub Form_Terminate()
    On Error Resume Next
    clsCTCS.Disconnect
    Set clsCTCS = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = (MsgBox("Are you sure do you want to exit?", vbQuestion + vbYesNo, "SPECIAL SERVICE") = vbNo)
End Sub


Private Sub grdCCRDtls_KeyPress(KeyAscii As Integer)
        Select Case KeyAscii
        Case vbKeyEscape
            txtARRContSz.SetFocus
        Case vbKeyReturn
            txtARRContSz.SetFocus
        Case Else
    End Select

End Sub

Private Sub grdCCRDtls_RowColChange()
Dim nCurrentRow As Integer
Dim strGKey As String
Dim ReeferPaidThruDate As String
Dim sPluginDate As String
Dim sPaidThruDate As String
Dim rstExpDtl As ADODB.Recordset

nCurrentRow = grdCCRDtls.Row
Select Case tabTran.Tab
    Case 0
        If optArrImpExp(1).Value = True Then
            With grdCCRDtls
                txtARRContNo.Text = .TextMatrix(nCurrentRow, 0)
                txtARRContSz.Text = .TextMatrix(nCurrentRow, 1)
                lblArrPrevAmt.Caption = .TextMatrix(nCurrentRow, 2)
                txtARROvzLen.Text = .TextMatrix(nCurrentRow, 3)
                txtARROvzWid.Text = .TextMatrix(nCurrentRow, 4)
                txtARROvzHgt.Text = .TextMatrix(nCurrentRow, 5)
                Call Sparcs_DGCode(.TextMatrix(nCurrentRow, 0), "EXPRT")
            End With
            Call lzGetCYXArr(txtARRCCRNo)
        Else
            With grdCCRDtls
                txtARRContNo.Text = .TextMatrix(nCurrentRow, 0)
                txtARRContSz.Text = .TextMatrix(nCurrentRow, 1)
                lblARREntryNo.Caption = .TextMatrix(nCurrentRow, 2)
                lblARRRegNo.Caption = .TextMatrix(nCurrentRow, 3)
                lblArrPrevAmt.Caption = .TextMatrix(nCurrentRow, 4)
                txtARROvzLen.Text = .TextMatrix(nCurrentRow, 5)
                txtARROvzWid.Text = .TextMatrix(nCurrentRow, 6)
                txtARROvzHgt.Text = .TextMatrix(nCurrentRow, 7)
                Call Sparcs_DGCode(.TextMatrix(nCurrentRow, 0), "IMPRT")
            End With
            Call lzGetCYMArr(txtARRCCRNo)
        End If
        Call ComputeOOG
    Case 1
        With grdCCRDtls
        If optStoImpExp(1).Value = True Then
            txtStoContNo.Text = .TextMatrix(nCurrentRow, 0)
            txtStoContSz.Text = .TextMatrix(nCurrentRow, 1)
'            lblStoPrevPay.Caption = .TextMatrix(nCurrentRow, 2)
            txtStoOvzLen.Text = .TextMatrix(nCurrentRow, 3)
            txtStoOvzWid.Text = .TextMatrix(nCurrentRow, 4)
            txtStoOvzHgt.Text = .TextMatrix(nCurrentRow, 5)
            lblStoEntryNo.Caption = ""
            lblStoRegNo.Caption = ""
        Else
            txtStoContNo.Text = .TextMatrix(nCurrentRow, 0)
            txtStoContSz.Text = .TextMatrix(nCurrentRow, 1)
            lblStoEntryNo.Caption = .TextMatrix(nCurrentRow, 2)
            lblStoRegNo.Caption = .TextMatrix(nCurrentRow, 3)
            lblStoPrevPay.Caption = .TextMatrix(nCurrentRow, 4)
            txtStoOvzLen.Text = .TextMatrix(nCurrentRow, 5)
            txtStoOvzWid.Text = .TextMatrix(nCurrentRow, 6)
            txtStoOvzHgt.Text = .TextMatrix(nCurrentRow, 7)
        End If
        End With
        mskExpStorageIN.Text = "    -  -  "
        txtStoExtDate.Text = "    -  -  "
        lblStoValidUntil.Caption = ""
        lblStoPluginDate.Caption = ""
        'Added by Navis Project Team 11/06/2009
        strGKey = ""
        If optStoImpExp(0).Value = True Then ' import
            Call lzGetCYMSto(txtSTOCCRNo)
        Else 'export
            If txtStoContNo.Text <> "" Then
                strGKey = GetGKey(txtStoContNo.Text, "QUEUED", "STORAGE")
                Call Sparcs_LastDisch(txtStoContNo.Text, "STORAGE", "", strGKey)
            End If
            Me.mskExpStorageIN.Text = Format(dStorage, "yyyy-mm-dd")
            Set rstExpDtl = New ADODB.Recordset
            rstExpDtl.Open "SELECT cntnum, cntsze, exprtr, broker FROM ccrcyx WHERE ccrnum = '" & txtSTOCCRNo.Text & "'", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
            txtImporter.Text = rstExpDtl.Fields(2)
            txtCusName.Text = rstExpDtl.Fields(3)
            rstExpDtl.Close
            Set rstExpDtl = Nothing
        End If
        Call ComputeOOG
        
        'Check container if it is a reefer container
        ReeferPaidThruDate = GetLastDischargeDate(txtStoContNo.Text, "INVOICED", "REEFER")
        If ReeferPaidThruDate <> "1899-12-30 00:00:00" And ReeferPaidThruDate <> "" Then
            lblStoPluginDate.Caption = Format(ReeferPaidThruDate, "yyyy-mm-dd hh:mm")
        Else
            lblStoPluginDate.Caption = ""
        End If
    Case 2
        With grdCCRDtls
            txtRfrContNo.Text = .TextMatrix(nCurrentRow, 0)
            txtRfrContSz.Text = .TextMatrix(nCurrentRow, 1)
            If optStoImpExp(2).Value = True Then 'export
                lblRfrPrevPay.Caption = .TextMatrix(nCurrentRow, 2)
                txtRfrEntryNo.Text = ""
                txtRfrRegNo.Text = ""
                Call lzGetCYXRfr(txtRFRCCRNo.Text)
            Else 'import
                lblRfrPrevPay.Caption = .TextMatrix(nCurrentRow, 4)
                txtRfrEntryNo.Text = .TextMatrix(nCurrentRow, 2)
                txtRfrRegNo.Text = .TextMatrix(nCurrentRow, 3)
                Call lzGetCYMRfr(txtRFRCCRNo.Text)

            End If
        End With
        
        If txtRfrContNo.Text <> "" Then
            If optStoImpExp(3).Value Then ' import
                Call GetReeferDates(txtRfrContNo.Text, "INVOICED", sPluginDate, sPaidThruDate)
            ElseIf optStoImpExp(2).Value Then
                Call GetReeferDates(txtRfrContNo.Text, "INVOICED", sPluginDate, sPaidThruDate)
                If sPaidThruDate = "" Then
                    Call GetReeferDates(txtRfrContNo.Text, "QUEUED", sPluginDate, sPaidThruDate)
                End If
            End If
                    
            If sPluginDate <> "" And sPaidThruDate <> "" Then
                txtRfrPlugInDate.Text = Format(sPluginDate, "yyyy-mm-dd hh:mm")
                txtRfrPlugInDate.Enabled = False
                lblRfrValidUntil.Caption = Format(sPaidThruDate, "yyyy-mm-dd hh:mm")
            ElseIf sPluginDate <> "" And sPaidThruDate = "" Then
                txtRfrPlugInDate.Text = Format(sPluginDate, "yyyy-mm-dd hh:mm")
                txtRfrPlugInDate.Enabled = False
                lblRfrValidUntil.Caption = ""
            Else
                txtRfrPlugInDate.Text = cEmptyRfrDate
                txtRfrPlugInDate.Enabled = True
                lblRfrValidUntil.Caption = ""
            End If
            If lblRfrValidUntil.Caption = "" Then
                MsgBox ("This container has no valid until date. This container maybe not a reefer container for now.")
            End If
        End If
End Select


End Sub

Private Sub grdCCRTran_GotFocus()
    mnuMenuEdit.Enabled = True
    mnuMenuDelete.Enabled = True
End Sub

Private Sub grdCCRTran_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            tabTran.SetFocus
        Case vbKeyReturn
            Call lzInitializePay
        Case Else
    End Select
End Sub

Private Sub grdCCRTran_LostFocus()
    mnuMenuEdit.Enabled = False
    mnuMenuDelete.Enabled = False
End Sub

'Private Sub mnuaccess_Click()
'frmPaySetting.Show vbModal
'End Sub

Private Sub grdCCRTran_RowColChange()
  txtCCRNumber.Enabled = False
End Sub

Private Sub cboVATRate_Click()
    Call lzUpdateGridVAT
    SendKeys "{TAB}"
End Sub


Private Sub mnuGridContainer_Click()
    If grdCCRDtls.Row > 2 Then
       grdCCRDtls.SetFocus
       SendKeys ("{RIGHT}")
    End If
End Sub

Private Sub mnuMenuDelete_Click()
    Call lzDeleteItem
End Sub

Private Sub mnuMenuExit_Click()
    Unload Me
End Sub

Private Sub mnuMenuGrid_Click()
    If grdCCRTran.Rows > 2 Then
        grdCCRTran.SetFocus
        SendKeys ("{RIGHT}")
    End If
End Sub

Private Sub mnuMenuPayment_Click()
    Call lzInitializePay
End Sub

Private Sub mnuMenuSave_Click()
    Call lzSavePrint
End Sub

Private Sub mnuMenuTab_Click()
    With tabTran
        If .Tab = (.Tabs - 1) Then
            .Tab = 0
        Else
            .Tab = .Tab + 1
        End If
        vTabOn = .Tab
        grdCCRDtls.Clear
        Call lzEnableTab
        tabTran.SetFocus
    End With
End Sub

'Commented Navis Project Team 10/29/2009
'Private Sub NTBSPayment1_KeyDown(KeyCode As Integer, Shift As Integer)
'
'If KeyCode = vbKeyF4 Then
'        Call NTBSPayment1_SavePaymentClick
'End If
'
'End Sub
'
'Private Sub NTBSPayment1_SavePaymentClick()
'On Error Resume Next
'If MsgBox("Save Payment to Disk?", vbYesNo + vbQuestion + vbDefaultButton1, "Save?") = vbYes Then
'     Call lzSavePrint
'End If
'DoEvents
'End Sub

Private Sub mskExpStorageIN_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case vbKeyReturn
        
            chkStoOvz.SetFocus
'            chkStoOvz.Value = 1 'True
        Case Else
    End Select
End Sub


Private Sub mskStoEndDate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            'SendKeys ("{TAB}")
            txtImporter.SetFocus
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub optArrImpExp_Click(Index As Integer)
    optArrImpExp(Index).BackColor = vbButtonFace
    bArrImp = optArrImpExp(0).Value
    lblARRCCRNo = IIf(bArrImp, "Gatepass No", "CCR No")
    txtARRCCRNo.Enabled = True
End Sub

Private Sub optArrImpExp_GotFocus(Index As Integer)
    optArrImpExp(Index).BackColor = &H80000018
End Sub

Private Sub optArrImpExp_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            chkGuarantee.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub optArrImpExp_LostFocus(Index As Integer)
    optArrImpExp(Index).BackColor = vbButtonFace
    bArrImp = optArrImpExp(0).Value
    lblARRCCRNo = IIf(bArrImp, "Gatepass No", "CCR No")
    txtARRCCRNo.Enabled = True
End Sub

Private Sub optArrImpExp_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not (optArrImpExp(0).Value Or optArrImpExp(1).Value)
End Sub
'added by Navis Project Team 10/30/2009
Private Sub optStoImpExp_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To optStoImpExp.UBound
        If Index = i Then
            optStoImpExp(i).BackColor = &H80000018
        Else
            optStoImpExp(i).BackColor = vbButtonFace
        End If
    Next
    bStoImp = optStoImpExp(0).Value
    lblStoCCRNo = IIf(bStoImp, "Gatepass No", "CCR No")
    If optStoImpExp(0).Value = True Then
        Me.lblStoValidUntil.Visible = True
        Me.mskExpStorageIN.Visible = False
    ElseIf optStoImpExp(1).Value = True Then
        Me.lblStoValidUntil.Visible = False
        Me.mskExpStorageIN.Visible = True
    End If
'    Me.lblStoValidUntil.Visible = True
    txtSTOCCRNo.Enabled = True
    txtSTOCCRNo.SetFocus
End Sub

'commented by Navis Project Team 10/30/2009
'Private Sub optStoImpExp_GotFocus(Index As Integer)
'    optStoImpExp(i).BackColor = &H80000018
'End Sub

Private Sub optStoImpExp_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            chkGuarantee.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

'commented by Navis Project Team 10/30/2009
'Private Sub optStoImpExp_LostFocus(Index As Integer)
'    optStoImpExp(Index).BackColor = vbButtonFace
'    bStoImp = optStoImpExp(0).Value
'    lblStoCCRNo = IIf(bStoImp, "Gatepass No", "CCR No")
'    If optStoImpExp(0).Value = True Then
'        Me.lblStoValidUntil.Visible = True
'        Me.mskExpStorageIN.Visible = False
'    ElseIf optStoImpExp(1).Value = True Then
'        Me.lblStoValidUntil.Visible = False
'        Me.mskExpStorageIN.Visible = True
'    End If
''    Me.lblStoValidUntil.Visible = True
'    txtSTOCCRNo.Enabled = True
'    txtSTOCCRNo.SetFocus
'End Sub
Private Sub optStoImpExp_LostFocus(Index As Integer)
    optStoImpExp(Index).BackColor = vbButtonFace
    bStoImp = optStoImpExp(0).Value
    lblStoCCRNo = IIf(bStoImp, "Gatepass No", "CCR No")
    If optStoImpExp(0).Value = True Then
        Me.lblStoValidUntil.Visible = True
        Me.mskExpStorageIN.Visible = False
    ElseIf optStoImpExp(1).Value = True Then
        Me.lblStoValidUntil.Visible = False
        Me.mskExpStorageIN.Visible = True
    End If
''    Me.lblStoValidUntil.Visible = True
'    txtSTOCCRNo.Enabled = True
'    txtSTOCCRNo.SetFocus
End Sub

Private Sub optStoImpExp_Validate(Index As Integer, Cancel As Boolean)
'    Cancel = Not (optStoImpExp(0).Value Or optArrImpExp(1).Value)
End Sub

Private Sub tabrecord_Click(PreviousTab As Integer)
On Error Resume Next
Select Case tabrecord.Tab
    Case 0
        optArrImpExp(0).SetFocus
        txtCusName.SelStart = 0
        txtCusName.SelLength = txtCusName.MaxLength
    Case 1
        ''NTBSPayment1.SetFocus
End Select
End Sub

Private Sub tabTran_GotFocus()
    Select Case tabTran.Tab
        Case 0
            Call lzClearArr
        Case 1
            Call lzClearSto
        Case 2
            Call lzClearRfr
        Case 3
            Call lzClearSOC
            'Call lzClearSO
        Case 4
            Call lzClearMsc
        Case 5
            Call lzClearOth
'        Case 6
'            Call lzClearParking
        Case Else
    End Select
        grdCCRDtls.Clear
End Sub

Private Sub txtADRAmt_Change()
    lblChange.Caption = EvaluateChange
    lblChange.Caption = Format(lblChange.Caption, "###,###,##0.00")
End Sub

Private Sub txtARRCCRNo_GotFocus()
    With txtARRCCRNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARRCCRNo_KeyPress(KeyAscii As Integer)
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

Private Sub txtARRCCRNo_LostFocus()
    txtARRCCRNo.BackColor = vbWindowBackground
    
    If (Val(txtARRCCRNo)) > 0 Then
        If bArrImp Then
            Call ImportDetails(txtARRCCRNo)
            'Call LoadCYMArr(txtARRCCRNo)
'            Call lzGetCYMArr(txtARRCCRNo)
'            Call Sparcs_DGCode(txtARRContNo.Text, "IMPRT")
        Else
            Call TransDetails(txtARRCCRNo, "CYX")
            'Call lzGetCYMArr(txtARRCCRNo)
            'Edited by Navis Project Team 11/05/2009
'            Call lzGetCYXArr(txtARRCCRNo, txtARRContNo)
'            Call lzGetCYXArr(txtARRCCRNo)
'            Call Sparcs_DGCode(txtARRContNo.Text, "EXPRT")
            'Call LoadCYX(txtARRCCRNo, "CYX")
            'Call lzGetCYXArr(txtARRCCRNo)
        End If
'        Call ComputeOOG
        
    End If
End Sub

Private Sub txtARRContNo_GotFocus()
    With txtARRContNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARRContNo_KeyPress(KeyAscii As Integer)
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

Private Sub txtARRContNo_LostFocus()
    txtARRContNo.BackColor = vbWindowBackground
    'If Not bArrImp And (Trim(txtARRContNo) <> "") Then
    '    'Call lzGetCYXArr(txtARRCCRNo, txtARRContNo)
    '    Call lzGetCYXArr(txtARRCCRNo)
    'End If
    
     'added by Navis Team 11/04/2009
'    Call ComputeOOG
'    If bArrImp Then
'        Call Sparcs_DGCode(txtARRContNo, "IMPRT")
''        Call Sparcs_OOG(txtARRContNo, "IMPRT")
'        If optArrImpExp(1).Value = False And (txtARRContNo) <> "" Then
'            Call lzGetCYMArr(txtARRCCRNo)
'            'Call lzGetCYXArr(txtARRCCRNo) original code
'        End If
'    Else
'        Call Sparcs_DGCode(txtARRContNo, "EXPRT")
''        Call Sparcs_OOG(txtARRContNo, "EXPRT")
'        If optArrImpExp(1).Value = True And (txtARRContNo) <> "" Then
'            'Modified by Navis Project Team 11/06/2009
''            Call lzGetCYXArr(txtARRCCRNo, txtARRContNo)
'            Call lzGetCYXArr(txtARRCCRNo)
'            'Call lzGetCYXArr(txtARRCCRNo) original code
'        End If
'    End If
'
    txtARRContSz.SetFocus
    
End Sub

Private Sub txtARRContSz_GotFocus()
    With txtARRContSz
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARRContSz_KeyPress(KeyAscii As Integer)
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

Private Sub txtARRContSz_LostFocus()
    txtARRContSz.BackColor = vbWindowBackground
End Sub

Private Sub txtARRContSz_Validate(Cancel As Boolean)
    Cancel = InStr("20|40|45|", txtARRContSz & "|") = 0
    If Cancel Then MsgBox "Invalid container size. Please correct..."
End Sub

Private Sub txtARROvzHgt_GotFocus()
    With txtARROvzHgt
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARROvzHgt_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtARROvzHgt)
End Sub

Private Sub txtARROvzLen_GotFocus()
    With txtARROvzLen
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARROvzLen_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtARROvzLen)
End Sub

Private Sub txtARROvzWid_GotFocus()
    With txtARROvzWid
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARROvzWid_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtARROvzWid)
End Sub

Private Sub txtARRUOM_GotFocus()
    With txtARRUOM
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtARRUOM_Validate(Cancel As Boolean)
    Cancel = (txtARRUOM <> "I") And (txtARRUOM <> "C")
End Sub

Private Sub txtChkAmt_Change(Index As Integer)
    lblChange.Caption = EvaluateChange
    lblChange.Caption = Format(lblChange.Caption, "###,###,##0.00")
End Sub

Private Sub txtChkAmt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            If Index > 0 Then
                txtChkBank(Index - 1).SetFocus
            Else
                txtCshAmt.SetFocus
            End If
            KeyAscii = 0
        Case vbKeyReturn
            txtChkNo(Index).SetFocus
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtChkBank_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            txtChkNo(Index).SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            If Index = 4 Then
                If cmdSave.Enabled = True Then
                    cmdSave.SetFocus
                End If
            Else
                txtChkAmt(Index + 1).SetFocus
            End If
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtChkNo_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            txtChkAmt(Index).SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            txtChkBank(Index).SetFocus
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtCshAmt_Change()
    lblChange.Caption = EvaluateChange
    lblChange.Caption = Format(lblChange.Caption, "###,###,##0.00")
End Sub

Private Sub txtCshAmt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            txtChkAmt(0).SetFocus
            'SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
            
    End Select
End Sub
Private Sub txtCCRNumber_GotFocus()
  With txtCCRNumber
    .BackColor = &HFFFFFF
    .SelStart = 0
    .SelLength = Len(txtCCRNumber.Text)
    .BackColor = &HFFFFC0
  End With
End Sub

Private Sub txtCCRNumber_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(cmbCompCode.Text) <> "" Then
            Call ApplyCCR(txtCCRNumber.Text - 1, cmbCompCode.Text)
        Else
            MsgBox "Please select Company Code."
            cmbCompCode.SetFocus
        End If
    End If
End Sub

Private Sub txtCCRNumber_LostFocus()
  Dim a As Integer
  Dim rstCCR As ADODB.Recordset
  Set rstCCR = New ADODB.Recordset
  
  rstCCR.Open "SELECT dbo.fn_CYGetInitialCCBR ('CY', '" & Trim(UCase(gUserID)) & "')", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
  'rstCCR.Open "SELECT prvccr FROM SPLALLOC WHERE TELLER = '" & UCase(gUserID) & "'", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
  
  If Not rstCCR.EOF Then
    If rstCCR.Fields(0) <> CLng(txtCCRNumber.Text) - 1 Then
        a = MsgBox("Save changes to CCR Number?", vbYesNo, "CY Special Services")
        If a = 6 Then
            Call ApplyCCR(CLng(txtCCRNumber.Text) - 1, cmbCompCode.Text)
        Else
            Call lzGetUserInfo
        End If
    End If
  End If
  
'  If rstCCR.Fields("prvccr") <> txtCCRNumber.Text - 1 Then
'    a = MsgBox("Save changes to CCR Number?", vbYesNo, "CY Special Services")
'    If a = 6 Then
'      Call ApplyCCR(txtCCRNumber.Text - 1)
'    Else
'      Call lzGetUserInfo
'    End If
'  End If
  
  rstCCR.Close
  Set rstCCR = Nothing
  txtCCRNumber.Text = lzGetNextCCR(gUserID, cmbCompCode)
  txtCCRNumber.BackColor = vbButtonFace
End Sub

Private Sub txtCshAmt_LostFocus()
'Me.cmdSave.SetFocus

End Sub

Private Sub txtCusName_GotFocus()
    With txtCusName
        .Text = Trim(.Text)
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtCusName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            Call lzInitialize
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtCusName_LostFocus()
    txtCusName.BackColor = vbWindowBackground
    txtCusName = UCase(txtCusName)
End Sub

Private Sub lzPopulateDangerClass()
    With cboDanger
        .AddItem " " & Chr(124) & " None"
        .AddItem "1" & Chr(124) & " Explosives"
        .AddItem "2" & Chr(124) & " Gases"
        .AddItem "3" & Chr(124) & " Inflammable Liquid"
        .AddItem "4" & Chr(124) & " Inflammable Solids"
        .AddItem "5" & Chr(124) & " Oxidizing Agents/Organic Peroxides"
        .AddItem "6" & Chr(124) & " Poisonous(toxic) and Infectious Substances"
        .AddItem "7" & Chr(124) & " Radioactive Substances"
        .AddItem "8" & Chr(124) & " Corrosives"
        .AddItem "9" & Chr(124) & " Miscellaneous Dangerous Substances"
    End With
End Sub

Private Sub lzCustomerLookUp()
    frmCustPick.Show 1
    If gsCusCode <> "" Then
'        txtCusCode = gsCusCode
        txtCusName.Text = gsCusName
        If chkVAT.Visible And chkVAT.Enabled Then
            chkVAT.SetFocus
        End If
    Else
'        txtCusCode = Space(6)
        'txtCusCode.Enabled = False
        txtCusName.Enabled = True
    End If
End Sub

Private Sub lzCustomerUnderG()
    frmCustPick.Show 1
    If gsCusCode <> "" Then
        Text1.Text = gsCusCode 'vCusCodeUnderG = gsCusCode
        Text2.Text = gsCusName
        If chkVAT.Visible And chkVAT.Enabled Then
            chkVAT.SetFocus
        End If
    Else
        Text1.Text = Space(6) 'vCusCodeUnderG = Space(6)
        txtCusName.Enabled = True
    End If
End Sub

Private Sub lzInitialize()
Dim n As Integer
    bVAT = True: bWTax = False: chkVAT.Value = 1: chkWTax.Value = 0: cboVATRate.ListIndex = 0
    bNewCCR = True: chkNewCCR.Value = 1
    bUnderG = False
    nCCRCounter = 0: bArrImp = False: bStoImp = False
    With grdCCRTran
        .Enabled = False
        For n = 7 To .Cols - 1
            .ColWidth(n) = 0
        Next n
        .Rows = 1
        .AddItem ""
        .TextMatrix(.Row, enCounter) = "**"
    End With
    
    txtRemark = ""
        
    Call lzEnableTab
    txtImporter = ""
    txtCusName = ""
    
    'Added by Navis Project Team 11/05/2009
    lblAmtDue.Caption = ".00"
    txtCshAmt.Text = ".00"
    lblChange.Caption = ".00"
    For n = 0 To 4
        txtChkAmt(n) = ""
        txtChkNo(n) = Space(txtChkNo(n).MaxLength)
        txtChkBank(n) = Space(txtChkBank(n).MaxLength)
    Next n
    
    'PRNH
    txtADRAmt.Text = ".00"
End Sub

Private Sub txtARROvzHgt_KeyPress(KeyAscii As Integer)
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

Private Sub txtARROvzHgt_LostFocus()
    txtARROvzHgt.BackColor = vbWindowBackground
End Sub

Private Sub txtARROvzLen_KeyPress(KeyAscii As Integer)
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

Private Sub txtARROvzLen_LostFocus()
    txtARROvzLen.BackColor = vbWindowBackground
End Sub

Private Sub txtARROvzWid_KeyPress(KeyAscii As Integer)
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

Private Sub txtARROvzWid_LostFocus()
    txtARROvzWid.BackColor = vbWindowBackground
End Sub

'Private Sub txtDriver_GotFocus()
'    With txtDriver
'        .BackColor = &HFFFFC0
'        .SelStart = 0
'        .SelLength = .MaxLength
'    End With
'End Sub
'
'Private Sub txtDriver_KeyPress(KeyAscii As Integer)
'Select Case KeyAscii
'        Case vbKeyEscape
'            SendKeys ("+{TAB}")
'            KeyAscii = 0
'        Case vbKeyReturn
'            SendKeys ("{TAB}")
'            KeyAscii = 0
'        Case Else
'    End Select
'End Sub
'
'Private Sub txtDriver_LostFocus()
'txtDriver.BackColor = vbWindowBackground
'End Sub



Private Sub txtImporter_GotFocus()
    With txtImporter
        .Text = Trim(.Text)
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtImporter_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
        Case vbKeyEscape
            'Call lzInitialize
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtImporter_LostFocus()
    txtImporter.BackColor = vbWindowBackground
    txtImporter = UCase(txtImporter)
    txtCusName.SetFocus
    
End Sub

Private Sub txtMscCCRNo_GotFocus()
    With txtMscCCRNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMscCCRNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            chkGuarantee.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtMscCCRNo_LostFocus()
    txtMscCCRNo.BackColor = vbWindowBackground
    
    If txtMscCCRNo.Text <> "" Then
        Call lzGetCYMMsc(txtMscCCRNo)
        txtMscRateCode.SetFocus
    Else
        txtMscContNo.SetFocus
    End If
   
End Sub

Private Sub txtMscContNo_GotFocus()
    With txtMscContNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMscContNo_KeyPress(KeyAscii As Integer)
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

Private Sub txtMscContNo_LostFocus()
    txtMscContNo.BackColor = vbWindowBackground
End Sub

Private Sub txtMscContSz_GotFocus()
    With txtMscContSz
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMscContSz_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
            bEscaped = True
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtMscContSz_LostFocus()
    txtMscContSz.BackColor = vbWindowBackground
    If bEscaped Then
        bEscaped = False
    Else
        Call lzShowRate
        
    End If
End Sub

Private Sub txtMscQty_GotFocus()
    With txtMscQty
        .Text = Trim(.Text)
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMscQty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            KeyAscii = 0
            txtImporter.SetFocus
        Case Else
    End Select
End Sub

Private Sub txtMscQty_LostFocus()
    txtMscQty.BackColor = vbWindowBackground
    lblMscAmount = Format(CCur("0" & lblMscRateAmt) * CCur("0" & txtMscQty), "###,##0.00")
End Sub

Private Sub txtMscQty_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtMscQty)
End Sub

Private Sub txtMscRateCode_GotFocus()
    With txtMscRateCode
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtMSCRateCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 0 Then
        Select Case KeyCode
            Case vbKeyF4
                frmCYRate.Show 1
                txtMscRateCode = vRateCode
                txtMscContSz = vRateSz
                Call lzShowRate
            Case Else
        End Select
    End If
End Sub

Private Sub txtMscRateCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case 48 To 57, 65 To 90, 97 To 122
            'PRNH
            If Len(txtMscRateCode.Text) = 6 Then
                SendKeys ("{TAB}")
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub txtMscRateCode_LostFocus()
    txtMscRateCode.BackColor = vbWindowBackground
End Sub

Private Sub txtOthAmount_GotFocus()
    With txtOthAmount
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthAmount_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            KeyAscii = 0
            txtImporter.SetFocus
            
        Case Else
    End Select
End Sub

Private Sub txtOthAmount_LostFocus()
    txtOthAmount.BackColor = vbWindowBackground
End Sub

Private Sub txtOthAmount_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtOthAmount)
End Sub

Private Sub txtOTHCCRNo_GotFocus()
    With txtOTHCCRNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOTHCCRNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            chkGuarantee.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtOTHCCRNo_LostFocus()
    txtOTHCCRNo.BackColor = vbWindowBackground
End Sub

Private Sub txtOthContNo_GotFocus()
    With txtOthContNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthContNo_KeyPress(KeyAscii As Integer)
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

Private Sub txtOthContNo_LostFocus()
    txtOthContNo.BackColor = vbWindowBackground
End Sub

Private Sub txtOthContSz_GotFocus()
    With txtOthContSz
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthContSz_KeyPress(KeyAscii As Integer)
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

Private Sub txtOthContSz_LostFocus()
    txtOthContSz.BackColor = vbWindowBackground
'    Call GetOtherContainerInfo
End Sub

Private Sub txtOthContSz_Validate(Cancel As Boolean)
    Cancel = InStr("20|40|45|  |", txtOthContSz & "|") = 0
    If Cancel Then MsgBox "Invalid container size. Please correct..."
End Sub

Private Sub txtOthEntryNo_GotFocus()
    With txtOthEntryNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthEntryNo_KeyPress(KeyAscii As Integer)
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

Private Sub txtOthEntryNo_LostFocus()
    txtOthEntryNo.BackColor = vbWindowBackground
End Sub

Private Sub txtOthFulEmp_GotFocus()
    With txtOthFulEmp
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthFulEmp_KeyPress(KeyAscii As Integer)
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

Private Sub txtOthFulEmp_LostFocus()
    txtOthFulEmp.BackColor = vbWindowBackground
End Sub

Private Sub txtOthFulEmp_Validate(Cancel As Boolean)
    If Trim(txtOthFulEmp) <> "" Then
        Cancel = (txtOthFulEmp <> "F") And (txtOthFulEmp <> "E")
    End If
End Sub

Private Sub txtOthOvzHgt_GotFocus()
    With txtOthOvzHgt
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthOvzHgt_KeyPress(KeyAscii As Integer)
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

Private Sub txtOthOvzHgt_LostFocus()
    txtOthOvzHgt.BackColor = vbWindowBackground
End Sub

Private Sub txtOthOvzHgt_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtOthOvzHgt)
End Sub

Private Sub txtOthOvzLen_GotFocus()
    With txtOthOvzLen
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthOvzLen_KeyPress(KeyAscii As Integer)
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

Private Sub txtOthOvzLen_LostFocus()
    txtOthOvzLen.BackColor = vbWindowBackground
End Sub

Private Sub txtOthOvzLen_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtOthOvzLen)
End Sub

Private Sub txtOthOvzWid_GotFocus()
    With txtOthOvzWid
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthOvzWid_KeyPress(KeyAscii As Integer)
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

Private Sub txtOthOvzWid_LostFocus()
    txtOthOvzWid.BackColor = vbWindowBackground
End Sub

Private Sub txtOthOvzWid_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtOthOvzWid)
End Sub

Private Sub txtOthRegNo_GotFocus()
    With txtOthRegNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthRegNo_KeyPress(KeyAscii As Integer)
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

Private Sub txtOthRegNo_LostFocus()
    txtOthRegNo.BackColor = vbWindowBackground
End Sub

Private Sub txtOthUOM_GotFocus()
    With txtOthUOM
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthUOM_KeyPress(KeyAscii As Integer)
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

Private Sub txtOthUOM_LostFocus()
    txtOthUOM.BackColor = vbWindowBackground
    Call lzOthOversize
End Sub

Private Sub txtOthUOM_Validate(Cancel As Boolean)
    Cancel = (txtOthUOM <> "C") And (txtOthUOM <> "I")
End Sub

Private Sub txtOthVessel_GotFocus()
    With txtOthVessel
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtOthVessel_KeyPress(KeyAscii As Integer)
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

Private Sub txtOthVessel_LostFocus()
    txtOthVessel.BackColor = vbWindowBackground
    txtOthVessel = UCase(txtOthVessel)
End Sub

'Private Sub txtParkingAMT_GotFocus()
'With txtParkingAMT
'        .BackColor = &HFFFFC0
'        .SelStart = 0
'        .SelLength = .MaxLength
'    End With
'End Sub
'
'Private Sub txtParkingAMT_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case vbKeyEscape
'            SendKeys ("+{TAB}")
'        Case vbKeySpace
'            SendKeys ("{F4}")
'        Case vbKeyReturn
'            txtImporter.SetFocus
'            txtRemark.Text = "PARKING FEE"
'        Case Else
'    End Select
'End Sub
'
'Private Sub txtParkingAMT_LostFocus()
'txtParkingAMT.BackColor = vbWindowBackground
'End Sub

Private Sub txtRemark_GotFocus()
    
    'To check customer name on underguarantee transactions
    If bUnderG Then
        If Len(Trim(txtCusName.Text)) = 0 Then
            MsgBox "Customer name required...", vbInformation
            txtCusName.SetFocus
        End If
    End If
    
    With txtRemark
        'If Len(Trim(.Text)) = 0 Then
        If txtMscRateCode = "WEIGHT" Then
            .Text = "WEIGHING"
        ElseIf txtMscRateCode = "CERTIF" Then
            .Text = "CERTIFIED TRUE COPY"
        ElseIf txtMscRateCode = "CHANDL" Then
            .Text = "CHANDLING FEE"
        Else
            Select Case tabTran.Tab
                Case 0
                    .Text = "ADD'L ARRASTRE"
                Case 1
                    .Text = "STORAGE UP TO " & Format(txtStoExtDate, "YYYY/MM/DD")
                Case 2
                    .Text = "REEFER UP TO "
                Case 3
                    .Text = "Shipper's Owned Container"
                    '.Text = "SHUTOUT"
                Case 4
                    txtRemark.Text = vRateDesc
                    '.Text = "EQUIPMENT RENTAL / MISCELLANEOUS"
                Case 5
                    .Text = "OTHER SPECIAL SERVICE"
                Case Else
                    .Text = Space(.MaxLength)
            End Select
            
        End If
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            tabTran.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys "{TAB}"
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtRemark_LostFocus()
    txtRemark.BackColor = vbWindowBackground
    txtRemark = UCase(txtRemark)
    
End Sub

Private Sub txtRFRCCRNo_GotFocus()
    With txtRFRCCRNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRFRCCRNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            chkGuarantee.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtRFRCCRNo_LostFocus()
Dim sPluginDate As String
Dim sPaidThruDate As String

If Trim(txtRFRCCRNo.Text) <> "" Then
    txtRFRCCRNo.BackColor = vbWindowBackground
    If optStoImpExp(3).Value = True Then    'Import
        Call TransDetails(txtRFRCCRNo.Text, "CYM")
'        Call lzGetCYMRfr(txtRFRCCRNo)
    ElseIf optStoImpExp(2).Value = True And txtRFRCCRNo.Text <> "        " Then  'Export
'        Dim a As Integer
'        Dim rstExpDtl As ADODB.Recordset
            
        Call TransDetails(txtRFRCCRNo.Text, "CYX")
'        Call lzGetCYXRfr(txtRFRCCRNo)
'        Set rstExpDtl = New ADODB.Recordset
'
'        rstExpDtl.Open "SELECT cntnum, cntsze, exprtr, broker FROM ccrcyx WHERE ccrnum = '" & txtRFRCCRNo.Text & "'", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
'
'        If Not rstExpDtl.EOF Then
'            'Added by Navis Project Team 11/06/2009
'            txtRfrContNo.Text = rstExpDtl.Fields(0)
'            txtRfrContSz.Text = rstExpDtl.Fields(1)
'            txtImporter.Text = rstExpDtl.Fields(2)
'            txtCusName.Text = rstExpDtl.Fields(3)
'            'check if the container is really a reefer container in N4
'
'            Call GetReeferDates(rstExpDtl.Fields(0), "INVOICED", sPluginDate, sPaidThruDate)
'            If sPaidThruDate = "" Then
'                Call GetReeferDates(rstExpDtl.Fields(0), "QUEUED", sPluginDate, sPaidThruDate)
'            End If
'            If sPluginDate <> "" And sPaidThruDate <> "" Then
'                txtRfrPlugInDate.Text = Format(sPluginDate, "yyyy-mm-dd hh:mm")
'                lblRfrValidUntil.Caption = Format(sPaidThruDate, "yyyy-mm-dd hh:mm")
'                txtRfrPlugInDate.Enabled = False
'            ElseIf sPluginDate <> "" And sPaidThruDate = "" Then
'                txtRfrPlugInDate.Text = Format(sPluginDate, "yyyy-mm-dd hh:mm")
'                txtRfrPlugInDate.Enabled = False
'                lblRfrValidUntil.Caption = ""
'            Else
'                txtRfrPlugInDate.Text = cEmptyRfrDate
'                txtRfrPlugInDate.Enabled = True
'                lblRfrValidUntil.Caption = ""
'            End If
''            txtRfrPlugInDate.SetFocus
''            txtRfrPlugInDate.SelStart = 0
''            txtRfrPlugInDate.SelLength = Len(txtRfrPlugInDate.Text)
'        End If
    End If
    
    If txtRfrPlugInDate = cEmptyRfrDate Then
        If txtRfrPlugInDate.Enabled Then
            txtRfrPlugInDate.SetFocus
            txtRfrPlugInDate.SelStart = 0
            txtRfrPlugInDate.SelLength = Len(Me.txtRfrPlugInDate)
        End If
    Else
        txtRfrExtDate.SetFocus
    End If
End If

End Sub

Private Sub txtRfrContNo_GotFocus()
    With txtRfrContNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRfrContNo_KeyPress(KeyAscii As Integer)
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

Private Sub txtRfrContNo_LostFocus()
    txtRfrContNo.BackColor = vbWindowBackground
End Sub

Private Sub txtRfrContSz_GotFocus()
    With txtRfrContSz
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRfrContSz_KeyPress(KeyAscii As Integer)
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

Private Sub txtRfrContSz_LostFocus()
    txtRfrContSz.BackColor = vbWindowBackground
End Sub

Private Sub txtRfrEntryNo_GotFocus()
    With txtRfrEntryNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRfrEntryNo_KeyPress(KeyAscii As Integer)
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

Private Sub txtRfrEntryNo_LostFocus()
    txtRfrEntryNo.BackColor = vbWindowBackground
End Sub

Private Sub txtRfrExtDate_GotFocus()
    With txtRfrExtDate
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRfrExtDate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            bEscaped = True
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn
            If (txtRfrPlugInDate <> cEmptyRfrDate) And (txtRfrExtDate <> cEmptyRfrDate) Then
                txtImporter.SetFocus
            End If
        Case Else
    End Select
End Sub

Private Sub txtRfrExtDate_LostFocus()
    txtRfrExtDate.BackColor = vbWindowBackground
    If bEscaped Then
'        bEscaped = False
'        SendKeys ("+{TAB}")
    Else
        Call txtRfrExtDate_Validate(True)
        If IsDate(txtRfrExtDate.Text) Then
            nRfrExtDate = txtRfrExtDate.Text
        End If
    End If
End Sub

Private Sub txtRfrExtDate_Validate(Cancel As Boolean)
Dim vHrs As Long
    
    Cancel = True
    If bEscaped Then
        bEscaped = False
        Cancel = False
        SendKeys ("+{TAB}")
        Exit Sub
    End If
    
    If (txtRfrExtDate <> cEmptyRfrDate) And Not IsDate(txtRfrExtDate) Then
        MsgBox "Invalid date value. Please correct."
        txtRfrExtDate.SetFocus
    Else
        If (lblRfrValidUntil <> "" And lblRfrValidUntil <> cEmptyRfrDate) And IsDate(txtRfrExtDate) Then
            If (DateDiff("n", CDate(lblRfrValidUntil), CDate(txtRfrExtDate)) < 1) Or _
               (DateDiff("n", gzGetSysDate(), CDate(txtRfrExtDate)) < 0) Then
                MsgBox "Should be greater than system date/time.  Please correct..."
                txtRfrExtDate.SetFocus
            End If
        ElseIf (txtRfrPlugInDate <> cEmptyRfrDate) And IsDate(txtRfrExtDate) Then
            If (DateDiff("n", CDate(txtRfrPlugInDate), CDate(txtRfrExtDate)) < 1) Or _
               (DateDiff("n", gzGetSysDate(), CDate(txtRfrExtDate)) < 0) Then
                MsgBox "Should be greater than system date/time.  Please correct..."
                txtRfrExtDate.SetFocus
            End If
        End If
        If (lblRfrValidUntil <> "" And lblRfrValidUntil <> cEmptyRfrDate) And (txtRfrExtDate <> cEmptyRfrDate) _
            And IsDate(lblRfrValidUntil) And IsDate(txtRfrExtDate) Then
            Cancel = True
        ElseIf (txtRfrPlugInDate <> cEmptyRfrDate) And (txtRfrExtDate <> cEmptyRfrDate) _
            And IsDate(lblRfrValidUntil) And IsDate(txtRfrExtDate) Then
            Cancel = True
        End If
    End If
    If IsDate(txtRfrExtDate) Then
        If IsDate(lblRfrValidUntil) Then
             vHrs = DateDiff("h", CDate(lblRfrValidUntil), CDate(txtRfrExtDate))
             If vHrs > 0 Then
                 lblRfrHrs = vHrs & " hrs"
             Else
                 lblRfrHrs = ""
             End If
        ElseIf IsDate(txtRfrPlugInDate) Then
             vHrs = DateDiff("h", CDate(txtRfrPlugInDate), CDate(txtRfrExtDate))
             If vHrs > 0 Then
                 lblRfrHrs = vHrs & " hrs"
             Else
                 lblRfrHrs = ""
             End If
        End If
    End If


End Sub

Private Sub txtRfrPlugInDate_GotFocus()
    With txtRfrPlugInDate
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRfrPlugInDate_KeyPress(KeyAscii As Integer)
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

Private Sub txtRfrPlugInDate_LostFocus()
    txtRfrPlugInDate.BackColor = vbWindowBackground
    If lblRfrValidUntil = cEmptyRfrDate Then lblRfrValidUntil = txtRfrPlugInDate
End Sub

Private Sub txtRfrPlugInDate_Validate(Cancel As Boolean)
    Cancel = True
    If (txtRfrPlugInDate <> cEmptyRfrDate) And Not IsDate(txtRfrPlugInDate) Then
        MsgBox "Invalid date value. Please correct."
    Else
        If (txtRfrExtDate <> cEmptyRfrDate) And (txtRfrPlugInDate <> cEmptyRfrDate) Then
            If CDate(txtRfrPlugInDate) > CDate(txtRfrExtDate) Then
                MsgBox "Should be earlier than extension date.  Please correct..."
            Else
                Cancel = False
            End If
        Else
            Cancel = False
        End If
    End If

End Sub

Private Sub txtRfrRegNo_GotFocus()
    With txtRfrRegNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtRfrRegNo_KeyPress(KeyAscii As Integer)
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

Private Sub txtRfrRegNo_LostFocus()
    txtRfrRegNo.BackColor = vbWindowBackground
End Sub

Private Sub txtSOCCRNo_GotFocus()
    With txtSOCCRNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtSOCCRNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            chkGuarantee.SetFocus
            KeyAscii = 0
        Case vbKeyReturn
            SendKeys ("{TAB}")
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtSOCCRNo_LostFocus()
    txtSOCCRNo.BackColor = vbWindowBackground
    If txtSOCCRNo.Text <> "" Then ExportDetails (txtSOCCRNo.Text)
End Sub

Private Sub txtSOContNo_GotFocus()
    With txtSOContNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtSOContNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            tabTran.SetFocus
        Case vbKeyReturn
            'SendKeys ("{TAB}")
            'KeyAscii = 0
            mskStoEndDate.SetFocus
            mskStoEndDate.BackColor = &HFFFFC0
            mskStoEndDate.SelStart = 0
            mskStoEndDate.SelLength = mskStoEndDate.MaxLength
        Case Else
    End Select
End Sub

'Private Sub txtSOContNo_LostFocus()
'Dim w As New CWaitCursor
'Dim strGKey As String
'    txtSOContNo.BackColor = vbWindowBackground
'
'    ConnectToNavis
'    strGKey = GetGKey("APHU6623529", "QUEUED", "STORAGE")
'    Call Sparcs_LastDisch("APHU6623529", "STORAGE", "", strGKey)
'    'CRODate = Last Discharge + 7 days by default
'    'strDate = Format(DateAdd("d", 7, CDate(mskLastDischargeDate.Text)), "yyyy-mm-dd")
'    'mskCRODate.Text = strDate
''    With clsCTCS
''        w.SetCursor
''        Call .GetLatestMove(txtSOContNo)
''        If .ContSize > 0 Then
''            lblSOContSz = .ContSize
''            lblSOFulEmp = IIf(.ContFull, "F", "E")
''            'txtSOVessel = .GetLastCYXVessel(txtSOContNo)    ' change .GetLastCYXVessel to look from SQL file, not AS/400 file
''        Else
''            lblSOContSz = "": lblSOFulEmp = "": txtSOVessel = ""
''        End If
''        w.Restore
''    End With
'End Sub

Private Sub txtSOContNo_LostFocus()
Dim sSize As String
Dim sStoStart As String
Dim sStoEnd As String
Dim strResult As String
Dim w As New CWaitCursor

    txtSOContNo.BackColor = vbWindowBackground
    bSOCHasPaidThruDate = False
    If txtSOContNo.Text <> "" Then
        lblSOFulEmp = "E"
        strResult = Sparcs_GetSOC(txtSOContNo.Text, sSize, sStoStart, sStoEnd, bSOCHasPaidThruDate)
        If Len(sSize) = 0 Then
            MsgBox "Cannot proceed with payment. Container is either not in Navis or not a shipper's owned container.", vbExclamation, "Shipper's Owned Container"
            lblSOContSz = ""
            mskStoStrtDate.Text = "    -  -  "
            mskStoEndDate.Text = "    -  -  "
            lblSOContSz = ""
        Else
            lblSOContSz = sSize
            mskStoStrtDate.Text = Format(CDate(sStoStart), "yyyy-mm-dd")
            mskStoEndDate.Text = Format(gzGetSysDate(), "yyyy-mm-dd")
            mskStoEndDate.SetFocus
            mskStoEndDate.BackColor = &HFFFFC0
            mskStoEndDate.SelStart = 0
            mskStoEndDate.SelLength = mskStoEndDate.MaxLength
            If bSOCHasPaidThruDate = True Then  '06Nov2009
               MsgBox "Container was already billed!", vbInformation, "Shippers Owned Container"
            End If
        End If
    End If
     
End Sub
Private Sub txtARRUOM_KeyPress(KeyAscii As Integer)
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

Private Sub txtARRUOM_LostFocus()
    txtARRUOM.BackColor = vbWindowBackground
    txtARRUOM = UCase(txtARRUOM)
    Call lzArrOversize
End Sub

Private Sub lzGetCYMArr(ByVal pGatepass As String)
Dim cmd As ADODB.Command
Dim vOvrLen, vOvrWid, vOvrHgt As Long
Dim w As New CWaitCursor
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "upnew_getcymarrastre"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adInteger
        .Parameters(1).Value = CLng(pGatepass)
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adChar
        .Parameters(2).Direction = adParamOutput
        .Parameters(3).Type = adSmallInt
        .Parameters(3).Direction = adParamOutput
        .Parameters(4).Type = adInteger
        .Parameters(4).Direction = adParamOutput
        .Parameters(5).Type = adChar
        .Parameters(5).Direction = adParamOutput
        .Parameters(6).Type = adCurrency
        .Parameters(6).Direction = adParamOutput
        .Parameters(7).Type = adInteger
        .Parameters(7).Direction = adParamOutput
        .Parameters(8).Type = adInteger
        .Parameters(8).Direction = adParamOutput
        .Parameters(9).Type = adInteger
        .Parameters(9).Direction = adParamOutput
       
        .Execute
        
        If .Parameters(0) = 1 Then
            txtARRContNo = .Parameters(2)
            txtARRContSz = .Parameters(3)
            lblARREntryNo = "" & .Parameters(4)
            lblARRRegNo = "" & .Parameters(5)
            lblArrPrevAmt = Format(.Parameters(6), "##,###,##0.00")
            
            If .Parameters(7) > 0 Then
                vOvrLen = .Parameters(7)
                vOvrWid = .Parameters(8)
                vOvrHgt = .Parameters(9)
                    
                If vOvrLen + vOvrWid + vOvrHgt > 0 Then
                    Select Case txtARRContSz
                        Case "20"
                            If vOvrLen <= 240 Then vOvrLen = vOvrLen + 240
                        Case "40"
                            If vOvrLen <= 480 Then vOvrLen = vOvrLen + 480
                        Case "45"
                            If vOvrLen <= 540 Then vOvrLen = vOvrLen + 540
                        Case Else
                            vOvrLen = 0
                    End Select
                    If vOvrWid <= 96 Then vOvrWid = vOvrWid + 96
                    If vOvrHgt <= 102 Then vOvrHgt = vOvrHgt + 102
                    
                    txtARROvzLen = vOvrLen
                    txtARROvzWid = vOvrWid
                    txtARROvzHgt = vOvrHgt
                    chkARROvz.Value = 1
                End If
                
            Else
                'With clsCTCS
                '    w.SetCursor
                '    Call .GetLatestMove(txtARRContNo)
                '    If .ContSize <> "" Then
                '        txtARRContSz = .ContSize
                '        vOvrLen = .OverFore + .OverAft
                '        vOvrWid = .OverLeft + .OverRight
                '        vOvrHgt = .OverHeight
                            
                '        If vOvrLen + vOvrWid + vOvrHgt > 0 Then
                            
                            '/joe
                '            Select Case txtARRContSz
                '               Case "20"
                '                   txtARROvzLen = 240 + Round(vOvrLen / 2.54)
                '               Case "40"
                '                   txtARROvzLen = 480 + Round(vOvrLen / 2.54)
                '               Case "45"
                '                  txtARROvzLen = 540 + Round(vOvrLen / 2.54)
                '            End Select
                                                        
                '            txtARROvzWid = 96 + Round(vOvrWid / 2.54)
                '            txtARROvzHgt = 102 + Round(vOvrHgt / 2.54)
                '            txtARRUOM = "I"
                '            chkARROvz.Value = 1
                '        End If
                '    End If
                'End With
            End If
            
        Else
            txtARRContNo = Space(txtARRContNo.MaxLength)
            txtARRContSz = Space(txtARRContSz.MaxLength)
            lblARREntryNo = ""
            lblARRRegNo = ""
            lblArrPrevAmt = ""
        
        End If
    
    End With
    
End Sub

Public Function getCYMArr(ByVal pGatepassNo As String) As Recordset
Dim strSQL As String
Dim rsCYM As Recordset

Set rsCYM = New ADODB.Recordset

strSQL = "SELECT cntnum, cntsze, entnum, regnum, " & _
                 "pARRAMT=arramt+arrvat-arrtax, cntovl, cntovw, cntovh, " & _
                 "refnum, vatcde, gtycde,cnsgne,broker, ovzamt, dgramt " & _
         "FROM CYMGPS " & _
         "WHERE status <> 'CAN' AND gpsnum = " & pGatepassNo & ""

rsCYM.Open strSQL, gcnnBilling

Set getCYMArr = rsCYM



End Function

Public Function ExportDetails(ByVal pCCR As String)
Dim strSQL As String
Dim rsCyx As ADODB.Recordset

Set rsCyx = New ADODB.Recordset

strSQL = "SELECT cntnum, cntsze, fulemp, vslcde " & _
         "FROM CCRCyx " & _
         "WHERE status <> 'CAN' AND ccrnum = " & pCCR & ""

rsCyx.Open strSQL, gcnnBilling
txtSOContNo.Text = rsCyx.Fields(0)
lblSOContSz = rsCyx.Fields(1)
lblSOFulEmp = rsCyx.Fields(2)
txtSOVessel.Text = rsCyx.Fields(3)
'Set ExportDetails = rsCyx
End Function

Public Function ImportDetails(ByVal pGpsNum As String)
Dim rsImport As Recordset
Dim nRowCnt, nVatCode As Integer
Dim strUG, strCosigne, strBroker As String
Dim nRefNum As Long

On Error GoTo ErrImport

    Set rsImport = getCYMArr(pGpsNum)
    
    grdCCRDtls.Clear
    
    With rsImport
        While .EOF = False
            nRowCnt = nRowCnt + 1
            grdCCRDtls.Row = grdCCRDtls.Row + 1
            grdCCRDtls.TextMatrix(nRowCnt, 0) = .Fields(0) 'Container No.
            grdCCRDtls.TextMatrix(nRowCnt, 1) = .Fields(1) 'Container Size
            grdCCRDtls.TextMatrix(nRowCnt, 2) = .Fields(2) 'Entry No.
            grdCCRDtls.TextMatrix(nRowCnt, 3) = .Fields(3) 'Registry No.
            grdCCRDtls.TextMatrix(nRowCnt, 4) = Format(.Fields(4), "###,###,##0.00") 'Amount
            grdCCRDtls.TextMatrix(nRowCnt, 5) = .Fields(5) 'Length
            grdCCRDtls.TextMatrix(nRowCnt, 6) = .Fields(6) 'Width
            grdCCRDtls.TextMatrix(nRowCnt, 7) = .Fields(7) 'Height
            nRefNum = .Fields(8) 'refnum
            nVatCode = IIf(.Fields(9) = " ", 0, .Fields(9)) 'vat code
            strUG = IIf(IsNull(.Fields(10)) = True, "", .Fields(10)) ' under guarantee
            txtImporter.Text = IIf(IsNull(.Fields(11)) = True, "", .Fields(11))
            strBroker = IIf(IsNull(.Fields(12)) = True, "", .Fields(12))
            .MoveNext
        Wend
    End With
    
    grdCCRDtls.FormatString = "CONTAINER No.     |  SIZE   |  ENTRY No.        | REGISTRY No.  | TOTAL  | LENGTH | WIDTH | HEIGHT"

'Commented on 11/17/2004 JOE -------
'    'Display the Vat Code
'    Select Case nVatCode
'        Case 0
'            chkVAT.Value = 0
'            chkWTax.Value = 0
'        Case 1
'            chkVAT.Value = 1
'            chkWTax.Value = 0
'        Case 2
'            chkVAT.Value = 1
'            chkWTax.Value = 1
'        Case Else
'            chkVAT.Value = 1
'            chkWTax.Value = 1
'    End Select
'------------------------------------
    
    
    'Display the UnderGurantee status
    Select Case strUG
        Case "N"
            'chkGuarantee.Value = 0
            If strBroker <> "" Then
                txtCusName.Text = strBroker
            End If
        Case "Y"
            'chkGuarantee.Value = 1
            With getCustomer1("IMPORT", nRefNum)
                If .RecordCount > 0 Then
                     txtCusName.Text = IIf(IsNull(.Fields(1)) = True, "", .Fields(1))
                     Text1.Text = IIf(IsNull(.Fields(0)) = True, "", .Fields(0))
                     Text2.Text = IIf(IsNull(.Fields(1)) = True, "", .Fields(1))
                Else
                    If strBroker <> "" Then
                        txtCusName.Text = strBroker
                    End If
                End If
            End With
            
        Case Else
            'chkGuarantee.Value = 0
            If strBroker <> "" Then
                txtCusName.Text = strBroker
            End If
    End Select

Exit Function
ErrImport:
    MsgBox Err.Description, vbCritical, "SPECIAL SERVICES"

End Function



Public Sub LoadCYMArr(ByVal pCCRNum As String)
Dim rs As Recordset
Dim nRowCnt, nVatCde As Integer
Dim strGuarntyCde As String
Dim nRefNo As Long


Set rs = getCYMArr(Trim(txtARRCCRNo.Text))
grdCCRDtls.Clear

While rs.EOF = False
    nRowCnt = nRowCnt + 1
    grdCCRDtls.Row = grdCCRDtls.Row + 1
    grdCCRDtls.TextMatrix(nRowCnt, 0) = rs.Fields(0) 'Container No.
    grdCCRDtls.TextMatrix(nRowCnt, 1) = rs.Fields(1) 'Container Size
    grdCCRDtls.TextMatrix(nRowCnt, 2) = rs.Fields(2) 'Entry No.
    grdCCRDtls.TextMatrix(nRowCnt, 3) = rs.Fields(3) 'Registry No.
    grdCCRDtls.TextMatrix(nRowCnt, 4) = Format(rs.Fields(4), "###,###,##0.00") 'Amount
    grdCCRDtls.TextMatrix(nRowCnt, 5) = rs.Fields(5) 'Length
    grdCCRDtls.TextMatrix(nRowCnt, 6) = rs.Fields(6) 'Width
    grdCCRDtls.TextMatrix(nRowCnt, 7) = rs.Fields(7) 'Height
    nRefNo = rs.Fields(8) 'refnum
    nVatCde = IIf(rs.Fields(9) = " ", 0, rs.Fields(9)) 'vat code
    strGuarntyCde = IIf(IsNull(rs.Fields(10)) = True, "", rs.Fields(10)) ' under guarantee
    rs.MoveNext
Wend

grdCCRDtls.FormatString = "CONTAINER No.     |  SIZE   |  ENTRY No.        | REGISTRY No.  | TOTAL  | LENGTH | WIDTH | HEIGHT"

'Checked Vat Exempt
If nVatCde = 1 Then
    chkVAT.Value = 1
Else
    chkVAT.Value = 0
End If

'Checked Under Gurantee
If strGuarntyCde = "Y" Then
    chkGuarantee.Value = 1
Else
    chkGuarantee.Value = 0
End If

'Get Customer Name
Call getCustomer(nRefNo)




End Sub


Public Function RetrieveData(ByVal pCCRNum As String, ByVal pTrans As String) As Recordset

Dim rsTrans As Recordset

On Error GoTo ErrRetrieve
        
    Set rsTrans = New ADODB.Recordset
    
    If pTrans = "CYX" Then
    
         strSQL = "Select cntnum,cntsze, arramt+ovzamt+arrvat-arrtax As pArrAmt,cntovzl, " & _
                        "cntovzw, cntovzh,refnum,vatcde,guarntycde,exprtr,broker " & _
                  "From CCRcyx Where status <> 'CAN' AND ccrnum = " & pCCRNum & ""
    
    Else
        
        strSQL = "Select cntnum,cntsze, stoamt+ovzamt+stovat-stotax As pStoAmt, " & _
                         "cntovl,cntovw, cntovh,refnum,vatcde,gtycde,cnsgne,broker " & _
                  "From CYMgps Where status <> 'CAN' AND gpsnum = " & pCCRNum & ""
    
    End If
    
    rsTrans.Open strSQL, gcnnBilling, adOpenStatic, adLockReadOnly, adCmdText
    
    Set RetrieveData = rsTrans
    
Exit Function
ErrRetrieve:
    MsgBox Err.Description, vbCritical, "SPECIAL SERVICES"
    Set rsTrans = Nothing
    
End Function



Public Function getCYX(ByVal pCCRNum As String, ByVal pType As String) As Recordset

Dim strSQL As String
Dim rsCyx As Recordset

Set rsCyx = New ADODB.Recordset

If pType = "CYX" Then
   strSQL = "SELECT  cntnum,cntsze, arramt+ovzamt+arrvat-arrtax As pArrAmt, " & _
                    "cntovzl,cntovzw, cntovzh,refnum,vatcde,guarntycde,exprtr,broker " & _
            "FROM CCRCYX WHERE status <> 'CAN' AND ccrnum = " & pCCRNum & ""
 
Else
  strSQL = "SELECT  cntnum,cntsze, stoamt+ovzamt+stovat-stotax As pStoAmt, " & _
                  "cntovl,cntovw, cntovh,refnum,vatcde,gtycde " & _
           "FROM CYMgps WHERE status <> 'CAN' AND gpsnum = " & pCCRNum & ""
  
End If

rsCyx.Open strSQL, gcnnBilling

Set getCYX = rsCyx
Set rsCyx = Nothing

End Function

Public Sub TransDetails(ByVal pCCRNum As String, ByVal pTrans As String)

Dim rsExport As Recordset
Dim nRowCnt, nVatCode As Integer
Dim strUG  As String
Dim nRefNum As Long
Dim strExporter, strBroker As String

On Error GoTo ErrExport
    
    Set rsExport = RetrieveData(pCCRNum, pTrans)
    
    grdCCRDtls.Clear
    With rsExport
        While .EOF = False
            nRowCnt = nRowCnt + 1
            grdCCRDtls.Row = grdCCRDtls.Row + 1
            grdCCRDtls.TextMatrix(nRowCnt, 0) = .Fields(0)    'CONTAINER No.
            grdCCRDtls.TextMatrix(nRowCnt, 1) = .Fields(1)    'CONTAINER Size
            grdCCRDtls.TextMatrix(nRowCnt, 2) = Format(.Fields(2), "###,###,##0.00") 'AMOUNT
            grdCCRDtls.TextMatrix(nRowCnt, 3) = .Fields(3)    'LENGTH
            grdCCRDtls.TextMatrix(nRowCnt, 4) = .Fields(4)    'WIDTH
            grdCCRDtls.TextMatrix(nRowCnt, 5) = .Fields(5)    'HEIGHT
            nRefNum = .Fields(6)                              'REFERENCE No.
            nVatCode = IIf(.Fields(7) = " ", 0, .Fields(7))   'VAT CODE
            strUG = .Fields(8)                                'UNDERGURANTEE
            txtImporter.Text = IIf(IsNull(.Fields(9)) = True, "", .Fields(9))                 'EXPORTER
            strBroker = IIf(IsNull(.Fields(10)) = True, "", .Fields(10))                           'BROKER
            .MoveNext
        Wend
    End With

    'Display the Custom Format
     grdCCRDtls.FormatString = "CONTAINER #     |  SIZE |  TOTAL        | LENGTH   | WIDTH  | HEIGHT"
    txtARRContNo.Text = grdCCRDtls.TextMatrix(1, 0)

    'Display the Undergurantee status
    Select Case strUG
        Case "N"
            'chkGuarantee.Value = 0
            txtCusName.Text = strBroker
        Case "Y"
            'chkGuarantee.Value = 1
       
            With getCustomer1("EXPORT", nRefNum)
                 If .RecordCount > 0 Then
                     txtCusName.Text = IIf(IsNull(.Fields(1)) = True, "", .Fields(1))
                     Text1.Text = IIf(IsNull(.Fields(0)) = True, "", .Fields(0))
                     Text2.Text = IIf(IsNull(.Fields(1)) = True, "", .Fields(1))
                 Else
                    txtCusName.Text = strBroker
                 End If
            End With
                 
        Case Else
            'chkGuarantee.Value = 0
            txtCusName.Text = strBroker
    End Select

    'Display the Vat Code
    Select Case nVatCode
        Case 0
            chkVAT.Value = 0
            chkWTax.Value = 0
        Case 1
            chkVAT.Value = 1
            chkWTax.Value = 0
        Case 2
            chkVAT.Value = 1
            chkWTax.Value = 1
        Case Else
            chkVAT.Value = 1
            chkWTax.Value = 1
    End Select
    
Exit Sub
ErrExport:
    MsgBox Err.Description, vbCritical, "SPECIAL SERVICES"
    
End Sub

Public Sub LoadCYX(ByVal pCCRNum As String, ByVal pType As String)
Dim rsExport As Recordset
Dim nRowCnt, nVatCode As Integer
Dim strUG  As String
Dim nRefNum As Long
Dim strExporter, strBroker As String


Set rsExport = getCYX(pCCRNum, pType)

grdCCRDtls.Clear

With rsExport
    While .EOF = False
        nRowCnt = nRowCnt + 1
        grdCCRDtls.Row = grdCCRDtls.Row + 1
        grdCCRDtls.TextMatrix(nRowCnt, 0) = .Fields(0)    'CONTAINER No.
        grdCCRDtls.TextMatrix(nRowCnt, 1) = .Fields(1)    'CONTAINER Size
        grdCCRDtls.TextMatrix(nRowCnt, 2) = Format(.Fields(2), "###,###,##0.00") 'AMOUNT
        grdCCRDtls.TextMatrix(nRowCnt, 3) = .Fields(3)    'LENGTH
        grdCCRDtls.TextMatrix(nRowCnt, 4) = .Fields(4)    'WIDTH
        grdCCRDtls.TextMatrix(nRowCnt, 5) = .Fields(5)    'HEIGHT
        nRefNum = .Fields(6)                              'REFERENCE No.
        nVatCode = IIf(.Fields(7) = " ", 0, .Fields(7))   'VAT CODE
        strUG = .Fields(8)                                'UNDERGURANTEE
        txtImporter.Text = .Fields(9)                     'EXPORTER
        strBroker = .Fields(10)                           'BROKER
        .MoveNext
    Wend
End With

'CUSTOM FORMAT DATA GRID
grdCCRDtls.FormatString = "CONTAINER #     |  SIZE |  TOTAL        | LENGTH   | WIDTH  | HEIGHT"


If strUG = "N" Then
   chkGuarantee.Value = 0
   txtCusName.Text = strBroker
Else
   chkGuarantee.Value = 1
   
   With getCustomer1("EXPORT", nRefNum)
        txtCusName.Text = .Fields(1)
        Text1.Text = .Fields(0)
        Text2.Text = .Fields(1)
   End With
   
End If


If nVatCode = 1 Then
    chkVAT.Value = 1
Else
    chkVAT.Value = 0
End If


'If pType = "CYX" Then
'    If strUG = "Y" Then
'        chkGuarantee.Value = 1
'        Call getCustomer(nRefNum)
'    Else
'        Call getCustomer(nRefNum)
'        chkGuarantee.Value = 0
'    End If
'Else
'    If txtCusName.Text = "" Then
'       Call getCustomer(nRefNum)
'    End If
'End If
'
'Set rs = Nothing

End Sub

Public Function getCustomer1(ByVal strTrans As String, nRefNum As Long) As Recordset
    Dim rsCust As Recordset
    
    On Error GoTo ErrCustomer
            
        Set rsCust = New ADODB.Recordset
        
        Select Case strTrans
            Case "IMPORT"
                
                strSQL = "Select cuscde,cusnam From CYMpay Where refnum = " & nRefNum & ""
                
            Case "EXPORT"
                
                strSQL = "Select cuscde,cusnam From CCRpay Where refnum = " & nRefNum & ""
                
        End Select
        
        rsCust.Open strSQL, gcnnBilling, adOpenStatic, adLockReadOnly, adCmdText
        
        Set getCustomer1 = rsCust
        Set rsCust = Nothing
            
    Exit Function
ErrCustomer:
    
    MsgBox Err.Description, vbCritical, "SPECIAL SERVICES"
    Set rsCust = Nothing
    
End Function

Public Sub getCustomer(ByVal nRefNo As Long)
Dim rsCustomer As ADODB.Recordset

On Error GoTo ErrCustName
    Set rsCustomer = New ADODB.Recordset
    
    If optArrImpExp(0).Value = True Then
        strSQL = "SELECT cuscde,cusnam FROM cympay WHERE refnum = " & nRefNo & ""
    ElseIf optArrImpExp(1).Value = True Then
        strSQL = "SELECT cuscde,cusnam FROM ccrpay WHERE refnum = " & nRefNo & ""
    ElseIf optStoImpExp(0).Value = True Then
        strSQL = "SELECT cuscde,cusnam FROM cympay WHERE refnum = " & nRefNo & ""
    End If
             
    rsCustomer.Open strSQL, gcnnBilling, adOpenStatic, adLockReadOnly, adCmdText
    
    With rsCustomer
        If .RecordCount > 0 Then
            txtCusName.Text = Trim(.Fields(1))
            Text1.Text = Trim(.Fields(0))
            Text2.Text = Trim(.Fields(1))
        End If
    End With
    
    Set rsCustomer = Nothing
Exit Sub
ErrCustName:
    MsgBox Err.Description, vbCritical, "Error in retrieving customer"
    Err.Clear
End Sub


'Modified by Navis Project Team 11/06/2009
'Private Sub lzGetCYXArr(ByVal pGatepass As String, ByVal pCNTnum As String)
Private Sub lzGetCYXArr(ByVal pGatepass As String)
Dim cmd As ADODB.Command
Dim vOvrLen, vOvrWid, vOvrHgt As Long
Dim w As New CWaitCursor
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        'Modifed by Navis Project Team 11/06/2009
        '.CommandText = "upnew_getcyxarrastre"
        .CommandText = "upnew_getcyxarrastre1"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adInteger
        .Parameters(1).Value = CLng(pGatepass)
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adChar
        .Parameters(2).Direction = adParamOutput
        .Parameters(3).Type = adSmallInt
        .Parameters(3).Direction = adParamOutput
'        .Parameters(4).Type = adInteger
'        .Parameters(4).Direction = adParamOutput
        .Parameters(4).Type = adVarChar
        .Parameters(4).Direction = adParamOutput
        
        .Parameters(5).Type = adChar
        .Parameters(5).Direction = adParamOutput
        .Parameters(6).Type = adCurrency
        .Parameters(6).Direction = adParamOutput
        .Parameters(7).Type = adInteger
        .Parameters(7).Direction = adParamOutput
        .Parameters(8).Type = adInteger
        .Parameters(8).Direction = adParamOutput
        .Parameters(9).Type = adInteger
        .Parameters(9).Direction = adParamOutput
       
        .Execute
        
        If .Parameters(0) = 1 Then
            txtARRContNo = .Parameters(2)
            txtARRContSz = .Parameters(3)
            lblARREntryNo = "" & .Parameters(4)
            lblARRRegNo = "" & .Parameters(5)
            lblArrPrevAmt = Format(.Parameters(6), "##,###,##0.00")
            
            If .Parameters(7) > 0 Then
                vOvrLen = .Parameters(7)
                vOvrWid = .Parameters(8)
                vOvrHgt = .Parameters(9)
                    
                If vOvrLen + vOvrWid + vOvrHgt > 0 Then
                    Select Case txtARRContSz
                        Case "20"
                            If vOvrLen <= 240 Then vOvrLen = vOvrLen + 240
                        Case "40"
                            If vOvrLen <= 480 Then vOvrLen = vOvrLen + 480
                        Case "45"
                            If vOvrLen <= 540 Then vOvrLen = vOvrLen + 540
                        Case Else
                            vOvrLen = 0
                    End Select
                    If vOvrWid <= 96 Then vOvrWid = vOvrWid + 96
                    If vOvrHgt <= 102 Then vOvrHgt = vOvrHgt + 102
                    
                    txtARROvzLen = vOvrLen
                    txtARROvzWid = vOvrWid
                    txtARROvzHgt = vOvrHgt
                    chkARROvz.Value = 1
                End If
            End If
        End If
    End With
'Dim cmd As ADODB.Command
'Dim vOvrLen, vOvrWid, vOvrHgt As Long
'Dim w As New CWaitCursor
'
'    ' create command
'    Set cmd = New ADODB.Command
'    With cmd
'        Set .ActiveConnection = gcnnBilling
'        .CommandText = "upnew_getcyxarrastre1"
'        .CommandType = adCmdStoredProc
'
'        ' set parameters then execute
''        .Parameters(0).Direction = adParamReturnValue
' '       .Parameters(0).Type = adInteger                 ' CCR number
'
'        .Parameters(1).Value = CLng(pGatepass)
'        .Parameters(1).Direction = adParamInput
'        .Parameters(1).Type = adInteger                 ' CCR number
'        .Parameters(2).Type = adChar                    ' container number
'        .Parameters(2).Direction = adParamOutput
'        .Parameters(3).Type = adSmallInt                ' container size
'        .Parameters(3).Direction = adParamOutput
'        .Parameters(4).Type = adCurrency                ' arrastre amount
'        .Parameters(4).Direction = adParamOutput
'        .Parameters(5).Type = adInteger                 ' oversize length
'        .Parameters(5).Direction = adParamOutput
'        .Parameters(6).Type = adInteger                 ' oversize width
'        .Parameters(6).Direction = adParamOutput
'        .Parameters(7).Type = adInteger                 ' oversize height
'        .Parameters(7).Direction = adParamOutput
'        .Parameters(8).Value = CLng(pGatepass)
'        .Parameters(8).Direction = adParamInput
'        .Parameters(8).Value = pCNTnum
'
'        .Execute
'
'        If .Parameters(0) = 1 Then
'            txtARRContNo = .Parameters(2)
'            txtARRContSz = .Parameters(3)
'            lblArrPrevAmt = Format(.Parameters(4), "##,###,##0.00")
'
''            If .Parameters(5) > 0 Then
''                vOvrLen = .Parameters(5)
''                vOvrWid = .Parameters(6)
''                vOvrHgt = .Parameters(7)
''
''                If vOvrLen + vOvrWid + vOvrHgt > 0 Then
''                    Select Case txtARRContSz
''                        Case "20"
''                            If vOvrLen <= 240 Then vOvrLen = vOvrLen + 240
''                        Case "40"
''                            If vOvrLen <= 480 Then vOvrLen = vOvrLen + 480
''                        Case "45"
''                            If vOvrLen <= 540 Then vOvrLen = vOvrLen + 540
''                        Case Else
''                            vOvrLen = 0
''                    End Select
''                    If vOvrWid <= 96 Then vOvrWid = vOvrWid + 96
''                    If vOvrHgt <= 102 Then vOvrHgt = vOvrHgt + 102
''
''                    txtARROvzLen = vOvrLen
''                    txtARROvzWid = vOvrWid
''                    txtARROvzHgt = vOvrHgt
''                    chkARROvz.Value = 1
''                End If
''
''            Else
'                With clsCTCS
'                    w.SetCursor
'                    Call .GetLatestMove(txtARRContNo)
'                    If .ContSize <> "" Then
'                        txtARRContSz = .ContSize
'                        vOvrLen = .OverFore + .OverAft
'                        vOvrWid = .OverLeft + .OverRight
'                        vOvrHgt = .OverHeight
'
'                        If vOvrLen + vOvrWid + vOvrHgt > 0 Then
'
'                            '/joe
'                            Select Case txtARRContSz
'                               Case "20"
'                                   txtARROvzLen = 240 + Round(vOvrLen / 2.54)
'                               Case "40"
'                                   txtARROvzLen = 480 + Round(vOvrLen / 2.54)
'                               Case "45"
'                                   txtARROvzLen = 540 + Round(vOvrLen / 2.54)
'                            End Select
'
'                            txtARROvzWid = 96 + Round(vOvrWid / 2.54)
'                            txtARROvzHgt = 102 + Round(vOvrHgt / 2.54)
'                            txtARRUOM = "I"
'                            chkARROvz.Value = 1
'                        Else
'                            chkARROvz.Value = 0
'                        End If
'
'                        Dim i As Integer
'                        i = .GetDangerCode(txtARRContNo)
'                        cboDanger.ListIndex = i
'
'
''                        If vOvrLen + vOvrWid + vOvrHgt > 0 Then
''
''                            Select Case txtARRContSz
''                                Case "20"
''                                    If vOvrLen <= 240 Then vOvrLen = vOvrLen + 240
''                                Case "40"
''                                    If vOvrLen <= 480 Then vOvrLen = vOvrLen + 480
''                                Case "45"
''                                    If vOvrLen <= 540 Then vOvrLen = vOvrLen + 540
''                                Case Else
''                                    vOvrLen = 0
''                            End Select
''                            If vOvrWid <= 96 Then vOvrWid = vOvrWid + 96
''                            If vOvrHgt <= 96 Then vOvrHgt = vOvrHgt + 96
''
''                            txtARROvzLen = vOvrLen / 2.54
''                            txtARROvzWid = vOvrWid / 2.54
''                            txtARROvzHgt = vOvrHgt / 2.54
''                            txtARRUOM = "I"
''                            chkARROvz.Value = 1
''                        End If
'                    End If
'                End With
'            'End If
'
'        Else
'            txtARRContNo = Space(txtARRContNo.MaxLength)
'            txtARRContSz = Space(txtARRContSz.MaxLength)
'            lblArrPrevAmt = ""
'
'        End If
'
'    End With
    
End Sub

Private Function lzCheckCYXVAT(ByVal pCCBR As String, ByVal pContNum As String) As Double


    Dim rs As New ADODB.Recordset
    rs.Open "SELECT vatcde FROM CCRCYX WHERE ccrnum = " & pCCBR & " AND cntnum = '" & pContNum & "'" _
            , gcnnBilling, , , adCmdText
       
    If Not rs.EOF Then
        rs.MoveFirst
        If rs!vatcde = 1 Or rs!vatcde = 2 Then
            lzCheckCYXVAT = 0.12
        ElseIf rs!vatcde = 3 Or rs!vatcde = 4 Then
            lzCheckCYXVAT = 0.07
        Else
            lzCheckCYXVAT = 0
        End If
    Else
        lzCheckCYXVAT = 0
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

Private Function lzCheckCYMVAT(ByVal pGatepass As String) As Double

    Dim rs As New ADODB.Recordset
    rs.Open "SELECT vatcde FROM CYMGPS WHERE gpsnum = " & pGatepass _
        , gcnnBilling, , , adCmdText
       
    If Not rs.EOF Then
        rs.MoveFirst
        If rs!vatcde = 1 Or rs!vatcde = 2 Then
            lzCheckCYMVAT = 0.12
        ElseIf rs!vatcde = 3 Or rs!vatcde = 5 Then
            lzCheckCYMVAT = 0.07
        Else
            lzCheckCYMVAT = 0
        End If
    Else
        lzCheckCYMVAT = 0.07
    End If
    
    rs.Close
    Set rs = Nothing

    
End Function

Private Function lzCheckCYXTax(ByVal pCCBR As String, ByVal pContNum As String) As Double
    
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT vatcde FROM CCRCYX WHERE ccrnum = " & pCCBR & " AND cntnum = '" & pContNum & "'" _
            , gcnnBilling, , , adCmdText
       
    If Not rs.EOF Then
        rs.MoveFirst
        If rs!vatcde = 2 Or rs!vatcde = 4 Or rs!vatcde = 5 Then
            lzCheckCYXTax = 0.02
        Else
            lzCheckCYXTax = 0
        End If
    Else
        lzCheckCYXTax = 0
    End If
    
    rs.Close
    Set rs = Nothing
    
    
End Function

Private Function lzCheckCYMTax(ByVal pGatepass As String) As Double
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT vatcde FROM CYMGPS WHERE gpsnum = " & pGatepass _
        , gcnnBilling, , , adCmdText
       
    If Not rs.EOF Then
        rs.MoveFirst
        If rs!vatcde = 2 Or rs!vatcde = 4 Or rs!vatcde = 5 Then
            lzCheckCYMTax = 0.02
        Else
            lzCheckCYMTax = 0
        End If
    Else
        lzCheckCYMTax = 0
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

Private Sub lzGetCYMMsc(ByVal pGatepass As String)
Dim cmd As ADODB.Command
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getcymarrastre"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adInteger
        .Parameters(1).Value = CLng("0" & pGatepass)
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adChar
        .Parameters(2).Direction = adParamOutput
        .Parameters(3).Type = adSmallInt
        .Parameters(3).Direction = adParamOutput
        .Parameters(4).Type = adInteger
        .Parameters(4).Direction = adParamOutput
        .Parameters(5).Type = adChar
        .Parameters(5).Direction = adParamOutput
        .Parameters(6).Type = adCurrency
        .Parameters(6).Direction = adParamOutput
        
        'PRNH - Company Code
        .Parameters(7).Type = adChar
        .Parameters(7).Direction = adParamOutput
       
        .Execute
        
        If .Parameters(0) = 1 Then
            txtMscContNo = .Parameters(2)
            txtMscContSz = .Parameters(3)
            
            'PRNH - Company Code
            cmbCompCode.Text = Trim(.Parameters(7))
        Else
            txtMscContNo = Space(txtMscContNo.MaxLength)
            txtMscContSz = Space(txtMscContSz.MaxLength)
        End If
    
    End With
    
End Sub

Private Sub lzClearArr()
    txtARRCCRNo = Space(txtARRCCRNo.MaxLength)
    txtARRContNo = Space(txtARRContNo.MaxLength)
    txtARRContSz = Space(txtARRContSz.MaxLength)
    lblARREntryNo = ""
    lblARRRegNo = ""
    chkARROvz.Value = 0
    txtARROvzLen = ""
    txtARROvzWid = ""
    txtARROvzHgt = ""
    txtARRUOM = "I"
    cboDanger.ListIndex = 0
    lblArrPrevAmt = ""
    optArrImpExp(0).SetFocus
    grdCCRDtls.Clear
End Sub
'sharon begin
Private Sub lzClearSOC()
    txtSOContNo = Space(txtSOContNo.MaxLength)
    lblSOContSz = ""
    lblSOFulEmp = ""
    mskStoStrtDate.Text = "    -  -  "
    mskStoEndDate.Text = "    -  -  "
    txtSOContNo.SetFocus
    txtSOContNo.SelStart = 0
    txtSOContNo.SelLength = Len(txtSOContNo.Text)
    grdCCRDtls.Clear
End Sub
'sharon end
'Private Sub lzClearSO()
'    txtSOCCRNo = Space(txtSOCCRNo.MaxLength)
'    txtSOContNo = Space(txtSOContNo.MaxLength)
'    lblSOContSz = ""
'    lblSOFulEmp = ""
'    txtSOVessel = ""
'    txtSOCCRNo.SetFocus
'    txtSOCCRNo.SelStart = 0
'    txtSOCCRNo.SelLength = Len(txtSOCCRNo.Text)
'    grdCCRDtls.Clear
'End Sub

Private Sub lzClearRfr()
    txtRFRCCRNo = Space(txtRFRCCRNo.MaxLength)
    txtRfrContNo = Space(txtRfrContNo.MaxLength): txtRfrContNo.Enabled = False
    txtRfrContSz = Space(txtRfrContSz.MaxLength): txtRfrContSz.Enabled = False
    txtRfrEntryNo = Space(txtRfrEntryNo.MaxLength): txtRfrEntryNo.Enabled = False
    txtRfrRegNo = Space(txtRfrRegNo.MaxLength): txtRfrRegNo.Enabled = False
    txtRfrPlugInDate = cEmptyRfrDate: txtRfrPlugInDate.Enabled = False
    lblRfrValidUntil = ""
    lblRfrPrevPay = ""
    txtRfrExtDate = cEmptyRfrDate
    txtRFRCCRNo.SetFocus
    grdCCRDtls.Clear
End Sub

Private Sub lzClearSto()
    txtSTOCCRNo = Space(txtSTOCCRNo.MaxLength)
    txtStoContNo = Space(txtStoContNo.MaxLength)
    txtStoContSz = Space(txtStoContSz.MaxLength)
    lblStoEntryNo = ""
    lblStoRegNo = ""
    lblStoValidUntil = ""
    txtStoExtDate = Format(gzGetSysDate(), "YYYY-MM-DD")
    chkStoOvz.Value = 0
    txtStoOvzLen = ""
    txtStoOvzWid = ""
    txtStoOvzHgt = ""
    txtStoUOM = "I"
    lblStoRevTon = ""
    lblStoPrevPay = ""
    optStoImpExp(0).SetFocus
    'txtSTOCCRNo.SetFocus
    grdCCRDtls.Clear
'    'Added by Navis Project Team 10/30/2009
'    chkVAT.Value = 0
End Sub

Private Sub lzClearMsc()
    bEscaped = False
    txtMscCCRNo = Space(txtMscCCRNo.MaxLength)
    txtMscContNo = Space(txtMscContNo.MaxLength)
    txtMscContSz = Space(txtMscContSz.MaxLength)
    txtMscRateCode = Space(txtMscRateCode.MaxLength)
    lblMScRateDesc = ""
    lblMscRateAmt = ""
    lblMscRateUOM = ""
    txtMscQty = Space(txtMscQty.MaxLength)
    lblMscAmount = ""
    txtMscCCRNo.SetFocus
    
    'PRNH
    cmbCompCode.Text = ""
    
    grdCCRDtls.Clear
End Sub

Private Sub lzClearOth()
    txtOTHCCRNo = Space(txtOTHCCRNo.MaxLength)
    txtOthContNo = Space(txtOthContNo.MaxLength)
    txtOthContSz = Space(txtOthContSz.MaxLength)
    txtOthFulEmp = Space(txtOthFulEmp.MaxLength)
    txtOthEntryNo = Space(txtOthEntryNo.MaxLength)
    txtOthRegNo = Space(txtOthRegNo.MaxLength)
    txtOthVessel = Space(txtOthVessel.MaxLength)
    chkOthOvz.Value = 0
    txtOthOvzLen = ""
    txtOthOvzWid = ""
    txtOthOvzHgt = ""
    txtOthUOM = "I"
    txtOthAmount = ""
    txtOTHCCRNo.SetFocus
    grdCCRDtls.Clear
End Sub

'Private Sub lzClearParking()
'    txtTruckPLT = Space(txtTruckPLT.MaxLength)
'    txtDriver = Space(txtDriver.MaxLength)
'    txtTruckMake = Space(txtTruckMake.MaxLength)
'    txtTruckPLT.SetFocus
'    grdCCRDtls.Clear
'End Sub

Private Sub lzGetUserInfo()
    txtCCRNumber = lzGetNextCCR(gzCurrentUser, "SBITC")
    
    'PRNH
    txtCCRNumberISI = lzGetNextCCR(gzCurrentUser, "ISI")
End Sub

Private Function lzGetPreviousPayment(ByVal pGpsNum As Long, ByVal pCNTnum As String) As Currency
Dim cmd As ADODB.Command

'create a command object.
Set cmd = New ADODB.Command
With cmd
    
    Set .ActiveConnection = gcnnBilling
    .CommandText = "upnew_getcymstorage_payment"
    .CommandType = adCmdStoredProc
    
    'set parameters then execute the command object
    .Parameters(0).Direction = adParamReturnValue
    .Parameters(0).Type = adInteger
    .Parameters(1).Type = adInteger
    .Parameters(1).Value = pGpsNum                   'Gate Pass Number
    .Parameters(1).Direction = adParamInput
    
    .Parameters(2).Type = adChar
    .Parameters(2).Value = pCNTnum                   'Container Number
    .Parameters(2).Direction = adParamInput
    
    .Execute
            
    nStoAmt = IIf(IsNull(.Parameters(3)), 0, .Parameters(3))
    nStoVat = IIf(IsNull(.Parameters(4)), 0, .Parameters(4))
    nStoTax = IIf(IsNull(.Parameters(5)), 0, .Parameters(5))
    nOvzAmt = IIf(IsNull(.Parameters(6)), 0, .Parameters(6))
    lzGetPreviousPayment = IIf(IsNull(.Parameters(7)), 0, .Parameters(7))
    nStoDay = lzGetPreviousPayment

End With

End Function


Private Function lzGetRateInfo(ByVal pRTECDE As String, Optional ByVal pCNTSZE As String = "NIL") As Currency
Dim cmd As ADODB.Command
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getcyrateinfo"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        '.Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar                    ' rate code
        .Parameters(1).Value = pRTECDE
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adChar                    ' container size
        .Parameters(2).Value = IIf(pCNTSZE = "NIL", "  ", pCNTSZE)
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adChar                    ' rate type
        .Parameters(3).Direction = adParamOutput
        .Parameters(4).Type = adChar                    ' rate description
        .Parameters(4).Direction = adParamOutput
        .Parameters(5).Type = adCurrency                ' rate amount
        .Parameters(5).Direction = adParamOutput
        .Parameters(6).Type = adChar                    ' unit of measure
        .Parameters(6).Direction = adParamOutput
       
        .Execute
        
        lzGetRateInfo = IIf(IsNull(.Parameters(5)), 0, .Parameters(5))
        
        vRateCode = pRTECDE
        vRateSz = pCNTSZE
        vRateDesc = "" & .Parameters(4)
        vRateAmount = lzGetRateInfo
        vRateUOM = "" & .Parameters(6)
    End With
End Function

Private Sub lzUpdateGridVAT()
Dim n, i As Integer
    n = grdCCRTran.Rows
    If n > 2 Then
        With grdCCRTran
            lblAmtDue = ""
            For i = 1 To (n - 2)
                If .TextMatrix(i, enRateCode) <> cVoid Then
                nAmount = CCur(.TextMatrix(i, enAmount))
                nTotalAmount = CCur(.TextMatrix(i, enTotalAmt))
                If bVAT Then
                    nVATAmount = nAmount * 0.1
                    .TextMatrix(i, enVATAmt) = Format(nVATAmount, "##,##0.00")
                    nTotalAmount = nTotalAmount + nVATAmount
                Else
                    nVATAmount = CCur("0" & .TextMatrix(i, enVATAmt))
                    nTotalAmount = nTotalAmount - nVATAmount
                    nVATAmount = 0
                    .TextMatrix(i, enVATAmt) = ""
                End If
                .TextMatrix(i, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
                Call lzAddToTotal
                End If
            Next
        End With
    End If
End Sub

Private Sub lzUpdateGridWTax()
Dim n, i As Integer
    n = grdCCRTran.Rows
    If n > 2 Then
        With grdCCRTran
            lblAmtDue = ""
            For i = 1 To (n - 2)
                If .TextMatrix(i, enRateCode) <> cVoid Then
                    nAmount = CCur(.TextMatrix(i, enAmount))
                    nTotalAmount = CCur(.TextMatrix(i, enTotalAmt))
                    If bWTax Then
                        nWTaxAmount = nAmount * 0.02
                        .TextMatrix(i, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
                        nTotalAmount = nTotalAmount - nWTaxAmount
                    Else
                        nWTaxAmount = CCur("0" & .TextMatrix(i, enWTaxAmt))
                        nTotalAmount = nTotalAmount + nWTaxAmount
                        nWTaxAmount = 0
                        .TextMatrix(i, enWTaxAmt) = ""
                    End If
                    .TextMatrix(i, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
                    Call lzAddToTotal
                End If
            Next
        End With
    End If
End Sub

Private Sub lzAddToTotal()
Dim curTotal As Currency
    curTotal = CCur("0" & lblAmtDue)
    lblAmtDue = Format(curTotal + nTotalAmount, "#,###,##0.00")
    lblAmtDue.Refresh
End Sub

Private Sub lzComputeArr()
Dim curArrOvzAmt, curArrDanger As Currency
Dim nDiv As Single
Dim vatRate As Double

   'VAT Rate to be used
   'Added by Navis Project Team 11/04/2009
    If chkVAT.Visible = True And chkVAT.Value = 1 Then
        If Trim(cboVATRate.Text) = "5%" Then      '4% VAT Rate
            vatRate = 0.05
        Else                                       '10% VAT Rate
            vatRate = 0.12
        End If
    Else
        vatRate = 0
    End If
'    If Trim(cboVATRate.Text) = "5%" Then      '4% VAT Rate
'        vatRate = 0.05
'    Else                                       '10% VAT Rate
'        vatRate = 0.12
'    End If
    
    'Edited by Navis Project Team 11/05/2009
    ' compute new
    If txtARRContSz = "20" Then
        cRateCode = IIf(bArrImp, "CBIMP1", "CBEXP1")
    ElseIf txtARRContSz = "40" Then
        cRateCode = IIf(bArrImp, "CBIMP2", "CBEXP2")
    ElseIf txtARRContSz = "45" Then
        cRateCode = IIf(bArrImp, "CBIMP3", "CBEXP3")
    End If
'    cRateCode = IIf(bArrImp, "CBIMPA", "CBEXPA")
    
    nAmount = lzGetRateInfo(cRateCode, txtARRContSz.Text)
'    nAmount = lzGetRateInfo(cRateCode)
    
    If bARROversize Then
        curArrOvzAmt = lzArrOversize()
        nAmount = nAmount + curArrOvzAmt
    End If
    curArrDanger = lzArrDanger(nAmount)
    nAmount = Round(nAmount + curArrDanger, 2)
    
    Dim origVATRate, origTaxRate, origVATAmount, origTaxAmount As Double
    
    origVATRate = IIf(bArrImp, lzCheckCYMVAT(txtARRCCRNo.Text), lzCheckCYXVAT(txtARRCCRNo.Text, txtARRContNo.Text))
    origTaxRate = IIf(bArrImp, lzCheckCYMTax(txtARRCCRNo.Text), lzCheckCYXTax(txtARRCCRNo.Text, txtARRContNo.Text))
    
    Dim origBasicAmount As Double
    origBasicAmount = CCur("0" & lblArrPrevAmt)
    If origVATRate > 0 Then
      If origTaxRate > 0 Then
        nDiv = 1 + origVATRate - 0.02 '1.08
      Else
        nDiv = 1 + origVATRate '1.1
      End If
      origBasicAmount = Round(origBasicAmount / nDiv, 2)
    Else
      nDiv = 0
    End If
    
    
    nAmount = nAmount - origBasicAmount
    nVATAmount = IIf(bVAT, Round(nAmount * vatRate, 2), 0)
    nWTaxAmount = IIf(bWTax, Round(nAmount * 0.02, 2), 0) 'Modify Tax to 2%
    nTotalAmount = Round(nAmount + nVATAmount - nWTaxAmount, 2)
    
'    nTotalAmount = Round(nTotalAmount - CCur("0" & lblArrPrevAmt), 2)
'
'    nAmount = nTotalAmount
'    If bVAT Then
'      If bWTax Then
'        nDiv = 1 + vatRate - 0.02 '1.08
'      Else
'        nDiv = 1 + vatRate '1.1
'      End If
'      nAmount = Round(nAmount / nDiv, 2)
'    Else
'      nDiv = 0
'    End If
'
'    If bVAT Then
'      nVATAmount = nAmount * vatRate
'    Else
'      nVATAmount = 0
'    End If
'
'    If bWTax Then
'      nWTaxAmount = nAmount * 0.02
'    Else
'      nWTaxAmount = 0
'    End If
    
    If nTotalAmount > 0 Then
        ' new ccr
        If bNewCCR Or (nCCRCounter > 11) Then '7
            bNewCCR = False
            nCCRCounter = 1
        Else
            nCCRCounter = nCCRCounter + 1
        End If
        
        
        With grdCCRTran
            If HasDuplicateGatePass(txtARRCCRNo.Text) = False Or HasDuplicateContainerNum(txtARRContNo.Text) = False Then
                nPtr = .Rows
                .AddItem nPtr
                
                If nCCRCounter = 1 Then
                    .CellForeColor = &HFF0000
                End If
                
                .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
                .TextMatrix(nPtr - 1, enRateCode) = cRateCode
                .TextMatrix(nPtr - 1, enAmount) = Format(nAmount, "###,##0.00")
                If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
                If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
                'Edited by Navis Project Team 11/04/2009
                .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
'                .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount + curArrOvzAmt, "##,##,##0.00")
                
                .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
                .TextMatrix(nPtr - 1, encustomer) = txtCusName.Text
                .TextMatrix(nPtr - 1, enimporter) = txtImporter.Text
                .TextMatrix(nPtr - 1, enCCRNo) = txtARRCCRNo
                .TextMatrix(nPtr - 1, enContNo) = txtARRContNo
                .TextMatrix(nPtr - 1, enContSz) = txtARRContSz
                .TextMatrix(nPtr - 1, enEntryNo) = lblARREntryNo
                .TextMatrix(nPtr - 1, enRegNo) = lblARRRegNo
                If bARROversize Then
                    .TextMatrix(nPtr - 1, enOvzLen) = txtARROvzLen
                    .TextMatrix(nPtr - 1, enOvzWid) = txtARROvzWid
                    .TextMatrix(nPtr - 1, enOvzHgt) = txtARROvzHgt
                    .TextMatrix(nPtr - 1, enOvzUom) = txtARRUOM
                    .TextMatrix(nPtr - 1, enRevTon) = lblArrRevTon
                    .TextMatrix(nPtr - 1, enOvzAmt) = Format(curArrOvzAmt, "###,##0.00")
                End If
                If Left(cboDanger, 1) <> " " Then
                    .TextMatrix(nPtr - 1, enDangerCode) = cboDanger
                    .TextMatrix(nPtr - 1, enDangerAmt) = curArrDanger
                End If
                .TextMatrix(nPtr - 1, enRemark) = txtRemark
                
                .TextMatrix(nPtr, enCounter) = "**"
                .Row = nPtr
            Else
                MsgBox "Gatepass Number has been used!", vbInformation, "Import/Export Arrastre"
                lzClearArr
                Exit Sub
            End If
        End With
        'Added by Navis Project Team 11/04/2009
        'Added the computation of OOG in total payments
        'nTotalAmount = nTotalAmount + curArrOvzAmt
        
        Call lzAddToTotal
        Call lzClearArr
    Else
        MsgBox "No additional charges computed...", vbInformation
        tabTran.SetFocus
    End If
End Sub

Public Function HasDuplicateGatePass(ByVal pValue As String) As Boolean
Dim rowctr As Integer

pValue = UCase(Trim(pValue))
HasDuplicateGatePass = False

If grdCCRTran.Rows >= 0 Then
 For rowctr = 1 To grdCCRTran.Rows - 1
     If pValue = UCase(Trim(grdCCRTran.TextMatrix(rowctr, enCCRNo))) And grdCCRTran.TextMatrix(rowctr, enRateCode) = cRateCode Then
         HasDuplicateGatePass = True
         Exit Function
     End If
 Next rowctr
End If
End Function

Public Function HasDuplicateContainerNum(ByVal pValue As String) As Boolean
Dim rowctr As Integer

pValue = UCase(Trim(pValue))
HasDuplicateContainerNum = False

If grdCCRTran.Rows >= 0 Then
 For rowctr = 1 To grdCCRTran.Rows - 1
     If pValue = UCase(Trim(grdCCRTran.TextMatrix(rowctr, enContNo))) And grdCCRTran.TextMatrix(rowctr, enRateCode) = cRateCode Then
         HasDuplicateContainerNum = True
         Exit Function
     End If
 Next rowctr
End If
End Function
'sharon begin
Private Sub lzComputeSOC()
'    nPtr = grdCCRTran.Rows
'    grdCCRTran.AddItem nPtr
Dim vatRate As Double
Dim curStorage As Currency
Dim curLOLO As Currency
Dim nDwellDays As Integer

    If bSOCHasPaidThruDate = False Then
        Call lzComputeSOCLOLO(txtSOContNo.Text)
        Call lzComputeSOCStorage
    Else
        MsgBox "Container was already billed!", vbInformation, "Shippers Owned Container"
        Call lzClearSOC
    End If
    

End Sub
Private Sub lzComputeSOCLOLO(ByVal pContNo As String)
'    nPtr = grdCCRTran.Rows
'    grdCCRTran.AddItem nPtr
    Dim vatRate As Double
    Dim curStorage As Currency
    Dim curLOLORate As Currency
    Dim nDwellDays As Integer
    Dim strRateCode As String
    
    'get LOLO
    strRateCode = "MCLIF3"
    curLOLORate = lzGetRateInfo(strRateCode)
    nTotalAmount = curLOLORate
    If nTotalAmount > 0 Then
        ' new ccr
        If bNewCCR Or (nCCRCounter > 11) Then '7
            bNewCCR = False
            nCCRCounter = 1
        Else
            nCCRCounter = nCCRCounter + 1
        End If
    
        With grdCCRTran
            If HasDuplicateContNo(pContNo) = False Then
                nPtr = .Rows
                .AddItem nPtr
                
                If nCCRCounter = 1 Then
                    .CellForeColor = &HFF0000
                End If
                .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
                .TextMatrix(nPtr - 1, enRateCode) = strRateCode
                .TextMatrix(nPtr - 1, enAmount) = Format(curLOLORate, "###,##0.00")
                If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(0, "##,##0.00")
                If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(0, "##,##0.00")
                .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
                .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
                .TextMatrix(nPtr - 1, encustomer) = txtCusName.Text
                .TextMatrix(nPtr - 1, enimporter) = txtImporter.Text

                .TextMatrix(nPtr - 1, enCCRNo) = txtSOCCRNo
                .TextMatrix(nPtr - 1, enContNo) = txtSOContNo
                .TextMatrix(nPtr - 1, enContSz) = lblSOContSz
                .TextMatrix(nPtr - 1, enFulEmp) = lblSOFulEmp
                .TextMatrix(nPtr - 1, enEntryNo) = lblARREntryNo
                .TextMatrix(nPtr - 1, enVessel) = txtSOVessel
                .TextMatrix(nPtr - 1, enRemark) = txtRemark
                
                .TextMatrix(nPtr, enCounter) = "**"
                .Row = nPtr
            Else
                MsgBox "Container Number has been used!", vbInformation, "Shippers Owned Container"
                Call lzClearSOC
                Exit Sub
            End If
        End With
        Call lzAddToTotal
    Else
        MsgBox "No additional charges computed...", vbInformation
        txtSOCCRNo.SetFocus
    End If
End Sub
Private Sub lzComputeSOCStorage()
    Dim vatRate As Double
    Dim curStorageRate As Currency
    Dim curLOLO As Currency
   'VAT Rate to be used
'    If Trim(cboVATRate.Text) = "5%" Then      '4% VAT Rate
'        vatRate = 0.05
'    Else                                       '10% VAT Rate
'        vatRate = 0.12
'    End If

    nVATAmount = 0
    nWTaxAmount = 0
    
    
    
    'get storage rate per day
    Select Case lblSOContSz
            Case "20"
                cRateCode = "STODO1"
                curStorageRate = lzGetRateInfo(cRateCode, "20")
            Case "40"
                cRateCode = "STODO2"
                curStorageRate = lzGetRateInfo(cRateCode, "40")
            Case "45"
                cRateCode = "STODO3"
                curStorageRate = lzGetRateInfo(cRateCode, "45")
    End Select
    'compute storage
    
    nDwellDays = (DateDiff("d", CDate(mskStoStrtDate), CDate(mskStoEndDate))) + 1 'inclusive of gate in date
    
    nTotalAmount = (curStorageRate * nDwellDays)
    
    
'    nAmount = lzGetRateInfo(cRateCode)
'    nTotalAmount = Round(nAmount + nVATAmount - nWTaxAmount, 2)

    If nTotalAmount > 0 Then
        ' new ccr
        If bNewCCR Or (nCCRCounter > 11) Then '7
            bNewCCR = False
            nCCRCounter = 1
        Else
            nCCRCounter = nCCRCounter + 1
        End If
    
        With grdCCRTran
            If HasDuplicateContNo(txtSOContNo.Text) = False Then
                nPtr = .Rows
                .AddItem nPtr
                
                If nCCRCounter = 1 Then
                    .CellForeColor = &HFF0000
                End If
                .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
                .TextMatrix(nPtr - 1, enRateCode) = cRateCode
                .TextMatrix(nPtr - 1, enAmount) = Format(curStorageRate, "###,##0.00")
                If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
                If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
                .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
                .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
                .TextMatrix(nPtr - 1, encustomer) = txtCusName.Text
                .TextMatrix(nPtr - 1, enimporter) = txtImporter.Text

                .TextMatrix(nPtr - 1, enCCRNo) = txtSOCCRNo
                .TextMatrix(nPtr - 1, enContNo) = txtSOContNo
                .TextMatrix(nPtr - 1, enContSz) = lblSOContSz
                .TextMatrix(nPtr - 1, enFulEmp) = lblSOFulEmp
                .TextMatrix(nPtr - 1, enEntryNo) = lblARREntryNo
                .TextMatrix(nPtr - 1, enVessel) = txtSOVessel
                .TextMatrix(nPtr - 1, enRemark) = txtRemark
                
                .TextMatrix(nPtr, enCounter) = "**"
                .Row = nPtr
            Else
                MsgBox "Container Number has been used!", vbInformation, "Shippers Owned Container"
                Call lzClearSOC
                Exit Sub
            End If
        End With
        Call lzAddToTotal
        Call lzClearSOC
    Else
        MsgBox "No additional charges computed...", vbInformation
        txtSOCCRNo.SetFocus
    End If
End Sub

'sharon end
'Private Sub lzComputeSO()
''    nPtr = grdCCRTran.Rows
''    grdCCRTran.AddItem nPtr
'Dim vatRate As Double
'
'   'VAT Rate to be used
'    If Trim(cboVATRate.Text) = "5%" Then      '4% VAT Rate
'        vatRate = 0.05
'    Else                                       '10% VAT Rate
'        vatRate = 0.12
'    End If
'
'    cRateCode = IIf(lblSOFulEmp = "F", "SOF", "SOE")
'    nAmount = lzGetRateInfo(cRateCode)
'    nVATAmount = IIf(bVAT, Round(nAmount * vatRate, 2), 0)
''    nWTaxAmount = IIf(bWTax, nAmount * 0.1, 0)
'    nWTaxAmount = IIf(bWTax, Round(nAmount * 0.02, 2), 0) 'Modify Tax to 2%
'    nTotalAmount = Round(nAmount + nVATAmount - nWTaxAmount, 2)
'
'    If nTotalAmount > 0 Then
'        ' new ccr
'        If bNewCCR Or (nCCRCounter > 11) Then '7
'            bNewCCR = False
'            nCCRCounter = 1
'        Else
'            nCCRCounter = nCCRCounter + 1
'        End If
'
'        With grdCCRTran
'            If HasDuplicateContNo(txtSOContNo.Text) = False Then
'                nPtr = .Rows
'                .AddItem nPtr
'
'                If nCCRCounter = 1 Then
'                    .CellForeColor = &HFF0000
'                End If
'                .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
'                .TextMatrix(nPtr - 1, enRateCode) = cRateCode
'                .TextMatrix(nPtr - 1, enAmount) = Format(nAmount, "###,##0.00")
'                If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
'                If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
'                .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
'                .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
'                .TextMatrix(nPtr - 1, encustomer) = txtCusName.Text
'                .TextMatrix(nPtr - 1, enimporter) = txtImporter.Text
'
'                .TextMatrix(nPtr - 1, enCCRNo) = txtSOCCRNo
'                .TextMatrix(nPtr - 1, enContNo) = txtSOContNo
'                .TextMatrix(nPtr - 1, enContSz) = lblSOContSz
'                .TextMatrix(nPtr - 1, enFulEmp) = lblSOFulEmp
'                .TextMatrix(nPtr - 1, enEntryNo) = lblARREntryNo
'                .TextMatrix(nPtr - 1, enVessel) = txtSOVessel
'                .TextMatrix(nPtr - 1, enRemark) = txtRemark
'
'                .TextMatrix(nPtr, enCounter) = "**"
'                .Row = nPtr
'            Else
'                MsgBox "Container Number has been used!", vbInformation, "Shutout"
'                Call lzClearSO
'                Exit Sub
'            End If
'        End With
'        Call lzAddToTotal
'        Call lzClearSO
'    Else
'        MsgBox "No additional charges computed...", vbInformation
'        txtSOCCRNo.SetFocus
'    End If
'End Sub

Private Function lzArrOversize() As Currency
Dim pLength, pWidth, pHeight As Single
    
    pLength = CSng(txtARROvzLen)
    pWidth = CSng(txtARROvzWid)
    pHeight = CSng(txtARROvzHgt)
    
    If txtARRUOM = "C" Then
       pLength = pLength / 2.54
       pWidth = pWidth / 2.54
       pHeight = pHeight / 2.54
    End If
    
    vRevTon = ((pLength * pWidth * pHeight) / 1728) / 40
    Select Case txtARRContSz
        Case "20"
            If vRevTon >= cRTon20 Then vRevTon = vRevTon - cRTon20
        Case "40"
            If vRevTon >= cRTon40 Then vRevTon = vRevTon - cRTon40
        Case "45"
            If vRevTon >= cRTon45 Then vRevTon = vRevTon - cRTon45
        Case Else
            vRevTon = 0
    End Select
    vRevTon = Round(vRevTon, 2)
    If vRevTon > 0 Then
        lblArrRevTon = Format(vRevTon, "###,##0.00")
    Else
        lblArrRevTon = ""
    End If
    'Added by Navis Project Team 11/04/2009
    'Find the correct arrastre rate for oversize container
    If bArrImp Then
        vRevTonRateArr = lzGetRateInfo("CBIMPA")
    Else
        vRevTonRateArrExp = lzGetRateInfo("CBEXPA")
    End If
    
    If bArrImp Then
        lzArrOversize = vRevTon * vRevTonRateArr
    Else
        lzArrOversize = vRevTon * vRevTonRateArrExp
    End If

End Function

Private Function lzArrDanger(ByVal pAmt As Currency) As Currency
Dim sDangerCode As String * 1
    sDangerCode = Left(cboDanger, 1)
    Select Case sDangerCode
        Case "1", "6", "8"
            lzArrDanger = pAmt * 0.5
        Case "2", "3", "4", "7"
            lzArrDanger = pAmt * 0.25
        Case "5", "9"
            lzArrDanger = pAmt * 0.1
        Case Else
            lzArrDanger = 0
    End Select
End Function

Private Sub txtSOVessel_GotFocus()
    With txtSOVessel
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtSOVessel_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys "+{TAB}"
            KeyAscii = 0
        Case vbKeyReturn
            txtImporter.SetFocus
            KeyAscii = 0
        Case Else
    End Select
End Sub

Private Sub txtSOVessel_LostFocus()
    txtSOVessel.BackColor = vbWindowBackground
    txtSOVessel = UCase(txtSOVessel)
End Sub

Private Sub txtSOVessel_Validate(Cancel As Boolean)
    txtSOVessel = UCase(txtSOVessel)
End Sub

Private Sub txtSTOCCRNo_GotFocus()
    
    With txtSTOCCRNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtSTOCCRNo_KeyPress(KeyAscii As Integer)
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

Private Sub txtSTOCCRNo_LostFocus()
    Dim strGKey As String
    Dim ReeferPaidThruDate As String
    txtSTOCCRNo.BackColor = vbWindowBackground
    If txtSTOCCRNo.Text <> "" Then
        If bStoImp And (Val(txtSTOCCRNo)) > 0 Then
'            Call lzGetCYMSto(txtSTOCCRNo)
            Call TransDetails(txtSTOCCRNo, "CYM")
        Else
            Call TransDetails(txtSTOCCRNo, "CYX")
'            If optStoImpExp(1).Value = True And txtSTOCCRNo.Text <> "" Then
'              Dim a As Integer
'              Dim rstExpDtl As ADODB.Recordset
'              Set rstExpDtl = New ADODB.Recordset
'
'                rstExpDtl.Open "SELECT cntnum, cntsze, exprtr, broker FROM ccrcyx WHERE ccrnum = '" & txtSTOCCRNo.Text & "'", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
'
'                If Not rstExpDtl.EOF Then
'                    Me.txtStoContNo.Text = rstExpDtl.Fields(0)
'                    Me.txtStoContSz.Text = rstExpDtl.Fields(1)
'                    txtImporter.Text = rstExpDtl.Fields(2)
'                    txtCusName.Text = rstExpDtl.Fields(3)
'
'                    'ConnectToNavis
'                    Call ComputeOOG
'                    strGKey = ""
'                    strGKey = GetGKey(Me.txtStoContNo.Text, "QUEUED", "STORAGE")
'                    Call Sparcs_LastDisch(Me.txtStoContNo.Text, "STORAGE", "", strGKey)
'                    Me.mskExpStorageIN.Text = Format(dStorage, "yyyy-mm-dd")
'                    'Added by Navis Project Team 11/2009
'                    'Check container if it is a reefer container
'                    ReeferPaidThruDate = GetLastDischargeDate(txtStoContNo.Text, "INVOICED", "REEFER")
'                    If ReeferPaidThruDate <> "1899-12-30 00:00:00" And ReeferPaidThruDate <> "" Then
'                        lblStoPluginDate.Caption = Format(ReeferPaidThruDate, "yyyy-mm-dd hh:mm")
'                    Else
'                        lblStoPluginDate.Caption = ""
'                    End If
'                End If
'            End If
        End If
    End If
End Sub

Private Sub lzGetCYMSto(ByVal pGatepass As String)
Dim vOvrLen, vOvrWid, vOvrHgt As Long
Dim cmd As ADODB.Command
Dim w As New CWaitCursor
Dim strGKey As String
Dim ReeferPaidThruDate As String
Dim dbStorageValidUntil As String
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "upnew_getcymstorage"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue       ' 1 if succesfull, else 0
        .Parameters(1).Type = adInteger
        .Parameters(1).Value = CLng(pGatepass)
        .Parameters(1).Direction = adParamInput             ' gatepass number
        .Parameters(2).Type = adChar
        .Parameters(2).Direction = adParamOutput            ' container number
        .Parameters(3).Type = adSmallInt
        .Parameters(3).Direction = adParamOutput            ' container size
        .Parameters(4).Type = adInteger
        .Parameters(4).Direction = adParamOutput            ' entry number
        .Parameters(5).Type = adChar
        .Parameters(5).Direction = adParamOutput            ' registry number
        .Parameters(6).Type = adDate
        .Parameters(6).Direction = adParamOutput            ' storage validity
        .Parameters(7).Type = adCurrency
        .Parameters(7).Direction = adParamOutput            ' storage amount paid
        .Parameters(8).Type = adInteger
        .Parameters(8).Direction = adParamOutput            ' oversize length
        .Parameters(9).Type = adInteger
        .Parameters(9).Direction = adParamOutput            ' oversize width
        .Parameters(10).Type = adInteger
        .Parameters(10).Direction = adParamOutput           ' oversize height
        .Parameters(11).Type = adInteger
        .Parameters(11).Direction = adParamOutput           ' storage days
       
        .Execute
        
        If .Parameters(0) = 1 Then
            txtStoContNo = .Parameters(2)
            txtStoContSz = .Parameters(3)
            lblStoEntryNo = "" & .Parameters(4)
            lblStoRegNo = "" & .Parameters(5)
            lblStoValidUntil = Format(.Parameters(6), "YYYY-MM-DD")
            dbStorageValidUntil = Format(.Parameters(6), "YYYY-MM-DD")
            lblStoPrevPay = Format(.Parameters(7), "##,###,##0.00")
            
            vCYMStoDay = .Parameters(11)
            If .Parameters(8) > 0 Then
                vOvrLen = .Parameters(8)
                vOvrWid = .Parameters(9)
                vOvrHgt = .Parameters(10)
                    
                If vOvrLen + vOvrWid + vOvrHgt > 0 Then
                    Select Case txtStoContSz
                        Case "20"
                            If vOvrLen < 240 Then vOvrLen = vOvrLen + 240
                        Case "40"
                            If vOvrLen < 480 Then vOvrLen = vOvrLen + 480
                        Case "45"
                            If vOvrLen < 540 Then vOvrLen = vOvrLen + 540
                        Case Else
                            vOvrLen = 0
                    End Select
                    If vOvrWid < 96 Then vOvrWid = vOvrWid + 96
                    If vOvrHgt < 102 Then vOvrHgt = vOvrHgt + 102
                    
                    txtStoOvzLen = vOvrLen
                    txtStoOvzWid = vOvrWid
                    txtStoOvzHgt = vOvrHgt
                    chkStoOvz.Value = 1
                End If
                vCYMStoDay = 0
                
            Else
'                With clsCTCS
'                    w.SetCursor
'                    Call .GetLatestMove(txtStoContNo)
'                    If .ContSize <> "" Then
'                        txtARRContSz = .ContSize
'                        vOvrLen = .OverFore + .OverAft
'                        vOvrWid = .OverLeft + .OverRight
'                        vOvrHgt = .OverHeight
'
'
'                        If vOvrLen + vOvrWid + vOvrHgt > 0 Then
'
'                            '/joe
'                            Select Case txtARRContSz
'                               Case "20"
'                                   txtStoOvzLen = 240 + Round(vOvrLen / 2.54)
'                               Case "40"
'                                   txtStoOvzLen = 480 + Round(vOvrLen / 2.54)
'                               Case "45"
'                                   txtStoOvzLen = 540 + Round(vOvrLen / 2.54)
'                            End Select
'
'                            txtStoOvzWid = 96 + Round(vOvrWid / 2.54)
'                            txtStoOvzHgt = 102 + Round(vOvrHgt / 2.54)
'                            txtStoUOM = "I"
'                            chkStoOvz.Value = 1
'                        End If
'
'                    End If
'                End With
            End If
            'Added by Navis Project Team 10/29/2009
            'get the OOG in N4
            Call ComputeOOG
            
'            ConnectToNavis
            strGKey = ""
            strGKey = GetGKey(Me.txtStoContNo.Text, "INVOICED", "STORAGE")
            If strGKey = "" Then
                strGKey = GetGKey(Me.txtStoContNo.Text, "QUEUED", "STORAGE")
                If strGKey = "" Then
                    strGKey = GetGKey(Me.txtStoContNo.Text, "PARTIAL", "STORAGE")
                End If
            End If
            If strGKey <> "" Then
                Call Sparcs_LastDisch(Me.txtStoContNo.Text, "STORAGE", "", strGKey)
                If DateDiff("d", Format(dStorage, "yyyy-mm-dd"), "12/30/1899 12:00:00 AM") <> 0 Then
                    lblStoValidUntil.Caption = Format(dStorage, "yyyy-mm-dd")
                ElseIf dbStorageValidUntil <> "" And DateDiff("d", dbStorageValidUntil, "12/30/1899") <> 0 Then
                    lblStoValidUntil.Caption = dbStorageValidUntil
                Else
                    lblStoValidUntil.Caption = ""
                End If
            ElseIf dbStorageValidUntil <> "" And DateDiff("d", dbStorageValidUntil, "12/30/1899") <> 0 Then
                lblStoValidUntil.Caption = dbStorageValidUntil
            Else
                lblStoValidUntil.Caption = ""
            End If
            
            'Added by Navis Project Team 11/2009
            'Check container if it is a reefer container
            ReeferPaidThruDate = GetLastDischargeDate(txtStoContNo.Text, "INVOICED", "REEFER")
            If ReeferPaidThruDate <> "1899-12-30 00:00:00" And ReeferPaidThruDate <> "" Then
                lblStoPluginDate.Caption = Format(ReeferPaidThruDate, "yyyy-mm-dd hh:mm")
            Else
                lblStoPluginDate.Caption = ""
            End If
            
'            lblRfrValidUntil = Format(dReefer, "yyyy-mm-dd hh:mm")
'
'            lblRfrPrevPay = Format(.Parameters(8), "##,###,##0.00")
        Else
            txtStoContNo = Space(txtStoContNo.MaxLength)
            txtStoContSz = Space(txtStoContSz.MaxLength)
            lblStoEntryNo = ""
            lblStoRegNo = ""
            lblStoValidUntil = ""
            lblStoPrevPay = ""
        End If
    End With
    Set cmd = Nothing
End Sub

Private Sub ComputeOOG()
    
    If frmCYSCCR.tabTran.Tab = 0 Then
        If txtARRContNo.Text <> "" Then
            Call Sparcs_OOG(txtARRContNo.Text, IIf(optArrImpExp(0).Value = True, "IMPRT", "EXPRT"))
            If txtARRContSz.Text <> "" Then
                If IsNumeric(txtARROvzHgt.Text) Or IsNumeric(txtARROvzLen.Text) Or IsNumeric(txtARROvzWid.Text) Then
                    If CInt(txtARROvzHgt.Text) > 0 Or CInt(txtARROvzLen.Text) > 0 Or CInt(txtARROvzWid.Text) > 0 Then
                        If txtARRContSz.Text = "40" Then
                            If CInt(txtARROvzLen.Text) > 0 Then
                                txtARROvzLen.Text = CInt(txtARROvzLen.Text) / 2.54
                            End If
                            txtARROvzLen.Text = Round((CDbl(txtARROvzLen.Text) + 480), 2)
                        ElseIf txtARRContSz.Text = "20" Then
                            If CInt(txtARROvzLen.Text) > 0 Then
                                txtARROvzLen.Text = CInt(txtARROvzLen.Text) / 2.54
                            End If
                            txtARROvzLen.Text = Round((CDbl(txtARROvzLen.Text) + 240), 2)
                        End If
                        If CInt(txtARROvzHgt.Text) > 0 Then
                            txtARROvzHgt.Text = CInt(txtARROvzHgt.Text) / 2.54
                        End If
                        txtARROvzHgt.Text = Round((CDbl(txtARROvzHgt.Text) + 102), 2)
                        
                        If CInt(txtARROvzWid.Text) > 0 Then
                            txtARROvzWid.Text = CInt(txtARROvzWid.Text) / 2.54
                        End If
                        txtARROvzWid.Text = Round((CDbl(txtARROvzWid.Text) + 96), 2)
                        chkARROvz.Value = 1
    '                    Call lzArrOversize
                    End If
                End If
            End If
        End If
    ElseIf frmCYSCCR.tabTran.Tab = 1 Then
        If txtStoContNo.Text <> "" Then
            Call Sparcs_OOG(txtStoContNo.Text, IIf(optStoImpExp(0).Value = True, "IMPRT", "EXPRT"))
            If txtStoContSz.Text <> "" Then
                If IsNumeric(txtStoOvzHgt.Text) And IsNumeric(txtStoOvzLen.Text) And IsNumeric(txtStoOvzWid.Text) Then
                    If CInt(txtStoOvzHgt.Text) > 0 Or CInt(txtStoOvzLen.Text) > 0 Or CInt(txtStoOvzWid.Text) > 0 Then
                        If txtStoContSz.Text = "40" Then
                            If CInt(txtStoOvzLen.Text) > 0 Then
                                txtStoOvzLen.Text = CInt(txtStoOvzLen.Text) / 2.54
                            End If
                            txtStoOvzLen.Text = Round((CDbl(txtStoOvzLen.Text) + 480), 2)
                        ElseIf txtStoContSz.Text = "20" Then
                            If CInt(txtStoOvzLen.Text) > 0 Then
                                txtStoOvzLen.Text = CInt(txtStoOvzLen.Text) / 2.54
                            End If
                            txtStoOvzLen.Text = Round((CDbl(txtStoOvzLen.Text) + 240), 2)
                        End If
                        If CInt(txtStoOvzHgt.Text) > 0 Then
                            txtStoOvzHgt.Text = CInt(txtStoOvzHgt.Text) / 2.54
                        End If
                        txtStoOvzHgt.Text = Round((CDbl(txtStoOvzHgt.Text) + 102), 2)
                        
                        If CInt(txtStoOvzWid.Text) > 0 Then
                            txtStoOvzWid.Text = CInt(txtStoOvzWid.Text) / 2.54
                        End If
                        txtStoOvzWid.Text = Round((CDbl(txtStoOvzWid.Text) + 96), 2)
                        chkStoOvz.Value = 1
                        Call lzStoOversize
                    End If
                End If
            End If
        End If
    ElseIf frmCYSCCR.tabTran.Tab = 5 Then
        If txtOthContNo.Text <> "" Then
            Call Sparcs_GETCONTOOG(txtOthContNo.Text)
            If txtOthContSz.Text <> "" Then
                If CInt(txtOthOvzHgt.Text) > 0 Or CInt(txtOthOvzLen.Text) > 0 Or CInt(txtOthOvzWid.Text) > 0 Then
                    If txtOthContSz.Text = "40" Then
                        If CInt(txtOthOvzLen.Text) > 0 Then
                            txtOthOvzLen.Text = CInt(txtOthOvzLen.Text) / 2.54
                        End If
                        txtOthOvzLen.Text = Round((CDbl(txtOthOvzLen.Text) + 480), 2)
                    ElseIf txtOthContSz.Text = "20" Then
                        If CInt(txtOthOvzLen.Text) > 0 Then
                            txtOthOvzLen.Text = CInt(txtOthOvzLen.Text) / 2.54
                        End If
                        txtOthOvzLen.Text = Round((CDbl(txtOthOvzLen.Text) + 240), 2)
                    End If
                    If CInt(txtOthOvzHgt.Text) > 0 Then
                        txtOthOvzHgt.Text = CInt(txtOthOvzHgt.Text) / 2.54
                    End If
                    txtOthOvzHgt.Text = Round((CDbl(txtOthOvzHgt.Text) + 102), 2)
    
                    If CInt(txtOthOvzWid.Text) > 0 Then
                        txtOthOvzWid.Text = CInt(txtOthOvzWid.Text) / 2.54
                    End If
                    txtOthOvzWid.Text = Round((CDbl(txtOthOvzWid.Text) + 96), 2)
                    chkOthOvz.Value = 1
                    Call lzOthOversize
                End If
            End If
        End If
    End If
End Sub

Private Sub lzGetCYXRfr(ByVal pGatepass As String)
    Dim a As Integer
    Dim rstExpDtl As ADODB.Recordset
    Dim sPluginDate As String
    Dim sPaidThruDate As String
    
    
    Set rstExpDtl = New ADODB.Recordset

    rstExpDtl.Open "SELECT cntnum, cntsze, exprtr, broker FROM ccrcyx WHERE ccrnum = '" & txtRFRCCRNo.Text & "'", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText

    If Not rstExpDtl.EOF Then
        'Added by Navis Project Team 11/06/2009
        txtRfrContNo.Text = rstExpDtl.Fields(0)
        txtRfrContSz.Text = rstExpDtl.Fields(1)
        txtImporter.Text = rstExpDtl.Fields(2)
        txtCusName.Text = rstExpDtl.Fields(3)
        'check if the container is really a reefer container in N4
    
        Call GetReeferDates(rstExpDtl.Fields(0), "INVOICED", sPluginDate, sPaidThruDate)
        If sPaidThruDate = "" Then
            Call GetReeferDates(rstExpDtl.Fields(0), "QUEUED", sPluginDate, sPaidThruDate)
        End If
        If sPluginDate <> "" And sPaidThruDate <> "" Then
            txtRfrPlugInDate.Text = Format(sPluginDate, "yyyy-mm-dd hh:mm")
            lblRfrValidUntil.Caption = Format(sPaidThruDate, "yyyy-mm-dd hh:mm")
            txtRfrPlugInDate.Enabled = False
        ElseIf sPluginDate <> "" And sPaidThruDate = "" Then
            txtRfrPlugInDate.Text = Format(sPluginDate, "yyyy-mm-dd hh:mm")
            txtRfrPlugInDate.Enabled = False
            lblRfrValidUntil.Caption = ""
        Else
            txtRfrPlugInDate.Text = cEmptyRfrDate
            txtRfrPlugInDate.Enabled = True
            lblRfrValidUntil.Caption = ""
        End If
    End If
End Sub
Private Sub lzGetCYMRfr(ByVal pGatepass As String)
Dim cmd As ADODB.Command
Dim strGKey As String
Dim sPluginDate As String
Dim sPaidThruDate As String
    If Trim(txtRFRCCRNo) <> "" Then
        ' create command
        Set cmd = New ADODB.Command
        With cmd
            Set .ActiveConnection = gcnnBilling
            .CommandText = "up_getcymreefer"
            .CommandType = adCmdStoredProc
        
            ' set parameters then execute
            .Parameters(0).Direction = adParamReturnValue       ' 1 if succesfull, else 0
            .Parameters(1).Type = adInteger
            .Parameters(1).Value = CLng("0" & pGatepass)
            .Parameters(1).Direction = adParamInput             ' gatepass number
            .Parameters(2).Type = adChar
            .Parameters(2).Direction = adParamOutput            ' container number
            .Parameters(3).Type = adSmallInt
            .Parameters(3).Direction = adParamOutput            ' container size
            .Parameters(4).Type = adInteger
            .Parameters(4).Direction = adParamOutput            ' entry number
            .Parameters(5).Type = adChar
            .Parameters(5).Direction = adParamOutput            ' registry number
            .Parameters(6).Type = adDate
            .Parameters(6).Direction = adParamOutput            ' plugin date
            .Parameters(7).Type = adDate
            .Parameters(7).Direction = adParamOutput            ' valid until
            .Parameters(8).Type = adCurrency
            .Parameters(8).Direction = adParamOutput            ' reefer amount paid

            .Execute
            
            If .Parameters(0) = 1 Then
                txtRfrContNo = .Parameters(2): txtRfrContNo.Enabled = False
                txtRfrContSz = .Parameters(3): txtRfrContSz.Enabled = False
                txtRfrEntryNo = Trim(Str(.Parameters(4))): txtRfrEntryNo.Enabled = False
                txtRfrRegNo = Trim(.Parameters(5)): txtRfrRegNo.Enabled = False
                If .Parameters(6) <> cNullDate Then
                    txtRfrPlugInDate = Format(.Parameters(6), "YYYY-MM-DD hh:mm")
                    txtRfrPlugInDate.Enabled = False
                Else
                    txtRfrPlugInDate = cEmptyRfrDate
                    txtRfrPlugInDate.Enabled = True
                End If
                
                
'                ConnectToNavis
                'Modified by Project Navis Team 11/05/2009
'                strGKey = ""
'                strGKey = GetGKey(Me.txtRfrContNo.Text, "INVOICED", "REEFER")
'                Call Sparcs_LastDisch(Me.txtRfrContNo.Text, "REEFER", "", strGKey)
'                Me.txtRfrPlugInDate.Text = Format(dReefer, "yyyy-mm-dd hh:mm")
'                lblRfrValidUntil = Format(dReefer, "yyyy-mm-dd hh:mm")
                lblRfrPrevPay = Format(.Parameters(8), "##,###,##0.00")
                Call GetReeferDates(txtRfrContNo.Text, "INVOICED", sPluginDate, sPaidThruDate)
                If optStoImpExp(3).Value = 1 Then
                    If sPluginDate <> "" And sPaidThruDate <> "" Then
                        txtRfrPlugInDate.Text = Format(sPluginDate, "yyyy-mm-dd hh:mm")
                        lblRfrValidUntil.Caption = Format(sPaidThruDate, "yyyy-mm-dd hh:mm")
                    Else
                        MsgBox "Container is not a reefer type."
                    End If
                    
                End If

                
                Exit Sub
            End If
            
        End With
        Set cmd = Nothing
    
    End If
    
    txtRfrContNo = Space(txtRfrContNo.MaxLength): txtRfrContNo.Enabled = True
    txtRfrContSz = Space(txtRfrContSz.MaxLength): txtRfrContSz.Enabled = True
    txtRfrEntryNo = Space(txtRfrEntryNo.MaxLength): txtRfrEntryNo.Enabled = True
    txtRfrRegNo = Space(txtRfrRegNo.MaxLength): txtRfrRegNo.Enabled = True
    txtRfrPlugInDate = cEmptyRfrDate: txtRfrPlugInDate.Enabled = True
    lblRfrValidUntil = ""
    lblRfrPrevPay = ""
    
    
    
End Sub

Private Sub txtStoContNo_GotFocus()
    With txtStoContNo
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoContNo_KeyPress(KeyAscii As Integer)
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

Private Sub txtStoContNo_LostFocus()
    txtStoContNo.BackColor = vbWindowBackground
End Sub

Private Sub txtStoContSz_GotFocus()
    With txtStoContSz
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoContSz_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeyReturn

            If optStoImpExp(0).Value = True Then
                SendKeys ("{TAB}")
                KeyAscii = 0
            ElseIf optStoImpExp(1).Value = True Then
                mskExpStorageIN.SetFocus
                mskExpStorageIN.SelStart = 0
                mskExpStorageIN.SelLength = Len(mskExpStorageIN.Text)
            End If
        Case Else
    End Select
End Sub

Private Sub txtStoContSz_LostFocus()
    txtStoContSz.BackColor = vbWindowBackground
End Sub

Private Sub txtStoContSz_Validate(Cancel As Boolean)
    Cancel = InStr("20|40|45|", txtStoContSz & "|") = 0
    If Cancel Then MsgBox "Invalid container size. Please correct..."
End Sub

Private Sub txtStoExtDate_GotFocus()
    With txtStoExtDate
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoExtDate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            SendKeys ("+{TAB}")
            KeyAscii = 0
        Case vbKeySpace
            SendKeys ("{F4}")
        Case vbKeyReturn
            If lzStoExtDateValid() Then txtImporter.SetFocus
        Case Else
    End Select
End Sub

Private Sub txtStoExtDate_LostFocus()
    txtStoExtDate.BackColor = vbWindowBackground
    Call txtStoExtDate_Validate(True)
    nStoExtDate = txtStoExtDate.Text
End Sub

Private Sub txtStoExtDate_Validate(Cancel As Boolean)
    Cancel = Not lzStoExtDateValid()
End Sub

Private Sub txtStoOvzHgt_GotFocus()
    With txtStoOvzHgt
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoOvzHgt_KeyPress(KeyAscii As Integer)
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

Private Sub txtStoOvzHgt_LostFocus()
    txtStoOvzHgt.BackColor = vbWindowBackground
End Sub

Private Sub txtStoOvzHgt_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtStoOvzHgt)
End Sub

Private Sub txtStoOvzLen_GotFocus()
    With txtStoOvzLen
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoOvzLen_KeyPress(KeyAscii As Integer)
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

Private Sub txtStoOvzLen_LostFocus()
    txtStoOvzLen.BackColor = vbWindowBackground
End Sub

Private Sub txtStoOvzLen_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtStoOvzLen)
End Sub

Private Sub txtStoOvzWid_GotFocus()
    With txtStoOvzWid
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoOvzWid_KeyPress(KeyAscii As Integer)
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

Private Sub txtStoOvzWid_LostFocus()
    txtStoOvzWid.BackColor = vbWindowBackground
End Sub

Private Sub txtStoOvzWid_Validate(Cancel As Boolean)
    Cancel = Not IsNumeric("0" & txtStoOvzWid)
End Sub

Private Sub txtStoUOM_GotFocus()
    With txtStoUOM
        .BackColor = &HFFFFC0
        .SelStart = 0
        .SelLength = .MaxLength
    End With
End Sub

Private Sub txtStoUOM_KeyPress(KeyAscii As Integer)
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

Private Sub txtStoUOM_LostFocus()
    txtStoUOM.BackColor = vbWindowBackground
    Call lzStoOversize
End Sub

Private Sub lzComputeSto()
Dim curStoOvzAmt As Currency
Dim curCompOvzAmt As Currency
Dim curTotOvzAmt As Currency
Dim nPrevStoDays As Integer
Dim vatRate As Double
    
    'VAT Rate to be used
    If chkVAT.Visible = True And chkVAT.Value = 1 Then
        If Trim(cboVATRate.Text) = "5%" Then      '4% VAT Rate
            vatRate = 0.05
        Else                                       '10% VAT Rate
            vatRate = 0.12
        End If
    Else
        vatRate = 0
    End If
    
    ' compute
    'Edited by Navis Project Team 11/05/2009
    'check if the container is reefer
'    If txtStoContSz = "20" Then
'        cRateCode = IIf(bStoImp, "STOIM1", "STOEX1")
'    ElseIf txtStoContSz = "40" Then
'        cRateCode = IIf(bStoImp, "STOIM2", "STOEX1")
'    ElseIf txtStoContSz = "45" Then
'        cRateCode = IIf(bStoImp, "STOIM3", "STOEX1")
'    End If
    If lblStoPluginDate.Caption <> "1899-12-30 00:00:00" And lblStoPluginDate.Caption <> "" Then   ' rate code for reefer container
        If txtStoContSz = "20" Then
            cRateCode = IIf(bStoImp, "STOIM4", "STOEX4")
        ElseIf txtStoContSz = "40" Then
            cRateCode = IIf(bStoImp, "STOIM5", "STOEX5")
        ElseIf txtStoContSz = "45" Then
            cRateCode = IIf(bStoImp, "STOIM6", "STOEX6")
        End If
    ElseIf chkStoOvz.Value = 1 Then
        If txtStoContSz = "20" Then
            cRateCode = "STOOH1"
        ElseIf txtStoContSz = "40" Then
            cRateCode = "STOOH2"
        ElseIf txtStoContSz = "45" Then
            cRateCode = "STOOH3"
        End If
    Else 'Rate code for normal container
        If txtStoContSz = "20" Then
            cRateCode = IIf(bStoImp, "STOIM1", "STOEX1")
        ElseIf txtStoContSz = "40" Then
            cRateCode = IIf(bStoImp, "STOIM2", "STOEX2")
        ElseIf txtStoContSz = "45" Then
            cRateCode = IIf(bStoImp, "STOIM3", "STOEX3")
        End If
    End If

    
    'cRateCode = IIf(bStoImp, "IMST", "EXST")
    nAmount = lzGetRateInfo(cRateCode, txtStoContSz)
    If optStoImpExp(0).Value = True Then
        vStoDay = DateDiff("d", CDate(lblStoValidUntil), CDate(txtStoExtDate))
    ElseIf optStoImpExp(1).Value = True Then
        vStoDay = DateDiff("d", CDate(Me.mskExpStorageIN.Text), CDate(txtStoExtDate))
    End If
    nAmount = nAmount * vStoDay
    If bStoOversize Then
        'curStoOvzAmt = lzStoOversize() ' previous code
        curCompOvzAmt = lzStoOversize()
        
        'enhancedment
        'nPrevStoDays = lzGetPreviousPayment(Trim(txtSTOCCRNo), Trim(txtStoContNo))
        
        If nPrevStoDays <> 0 Then
           curTotOvzAmt = nPrevStoDays * (curCompOvzAmt / vStoDay)
           curStoOvzAmt = (curTotOvzAmt + curCompOvzAmt) - nOvzAmt
        Else
           curStoOvzAmt = curCompOvzAmt 'JOE-092004
        End If
        'nAmount = nAmount + curStoOvzAmt ' previous code
        
    End If
    nVATAmount = IIf(bVAT, Round((nAmount + curStoOvzAmt) * vatRate, 2), 0)
'    nWTaxAmount = IIf(bWTax, nAmount * 0.01, 0)
    nWTaxAmount = IIf(bWTax, Round((nAmount + curStoOvzAmt) * 0.02, 2), 0) 'Modify Tax to 2%
    nTotalAmount = Round(nAmount + nVATAmount - nWTaxAmount, 2)
    'nTotalAmount = nTotalAmount - CCur("0" & lblStoPrevPay)
        
    If nTotalAmount > 0 Then
        ' new ccr
        If bNewCCR Or (nCCRCounter > 11) Then '7
            bNewCCR = False
            nCCRCounter = 1
        Else
            nCCRCounter = nCCRCounter + 1
        End If
        
        With grdCCRTran
            If HasDuplicateGatePass(txtSTOCCRNo.Text) = False Then
                nPtr = .Rows
                .AddItem nPtr
                
                If nCCRCounter = 1 Then
                    .CellForeColor = &HFF0000
                End If
                
                .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
                .TextMatrix(nPtr - 1, enRateCode) = cRateCode
                .TextMatrix(nPtr - 1, enAmount) = Format(nAmount, "###,##0.00")
                If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
                If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
                If curStoOvzAmt > 0 Then .TextMatrix(nPtr - 1, enOvzAmt) = Format(curStoOvzAmt, "##,##0.00")
'                .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount + curStoOvzAmt, "##,##,##0.00")
                .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
                .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
                .TextMatrix(nPtr - 1, encustomer) = txtCusName.Text
                .TextMatrix(nPtr - 1, enimporter) = txtImporter.Text

                .TextMatrix(nPtr - 1, enCCRNo) = txtSTOCCRNo
                .TextMatrix(nPtr - 1, enContNo) = txtStoContNo
                .TextMatrix(nPtr - 1, enContSz) = txtStoContSz
                .TextMatrix(nPtr - 1, enEntryNo) = lblStoEntryNo
                .TextMatrix(nPtr - 1, enRegNo) = lblStoRegNo
                If bStoOversize Then
                    .TextMatrix(nPtr - 1, enOvzLen) = txtStoOvzLen
                    .TextMatrix(nPtr - 1, enOvzWid) = txtStoOvzWid
                    .TextMatrix(nPtr - 1, enOvzHgt) = txtStoOvzHgt
                    .TextMatrix(nPtr - 1, enOvzUom) = txtStoUOM
                    .TextMatrix(nPtr - 1, enRevTon) = lblStoRevTon
                    .TextMatrix(nPtr - 1, enOvzAmt) = Format(curStoOvzAmt, "###,##0.00")
                End If
                .TextMatrix(nPtr - 1, enStoValidUntil) = txtStoExtDate
                .TextMatrix(nPtr - 1, enStoDays) = vStoDay
                .TextMatrix(nPtr - 1, enRemark) = txtRemark
                
                .TextMatrix(nPtr, enCounter) = "**"
                .Row = nPtr
            Else
                MsgBox "Gatepass Number has been used!", vbInformation, "Storage Computation"
                Call lzClearSto
                Exit Sub
            End If
        End With
        'Added by Navis Project Team 11/04/2009
        'Added the computation of OOG in total payments
        'nTotalAmount = nTotalAmount + curStoOvzAmt
        
        Call lzAddToTotal
        Call lzClearSto
    Else
        MsgBox "No additional charges computed...", vbInformation
        txtSTOCCRNo.SetFocus
    End If
    
    
End Sub

Private Function lzStoOversize() As Currency
Dim pLength, pWidth, pHeight As Single
    pLength = CSng("0" & txtStoOvzLen)
    pWidth = CSng("0" & txtStoOvzWid)
    pHeight = CSng("0" & txtStoOvzHgt)
    
    If txtStoUOM = "C" Then
       pLength = pLength / 2.54
       pWidth = pWidth / 2.54
       pHeight = pHeight / 2.54
    End If
    
    vRevTon = pLength * pWidth * pHeight / 1728 / 40
    Select Case txtStoContSz
        Case "20"
            If vRevTon >= cRTon20 Then vRevTon = vRevTon - cRTon20
        Case "40"
            If vRevTon >= cRTon40 Then vRevTon = vRevTon - cRTon40
        Case "45"
            If vRevTon >= cRTon45 Then vRevTon = vRevTon - cRTon45
        Case Else
            vRevTon = 0
    End Select
    vRevTon = Round(vRevTon, 2)
    If vRevTon > 0 Then
        lblStoRevTon = Format(vRevTon, "###,##0.00")
    Else
        lblStoRevTon = ""
    End If
    
    'Added by Navis Project Team 11/04/2009
    'Find the correct storage rate for oversize container
'    vRevTonRateSto = lzGetRateInfo("STOOH1", txtStoContSz.Text)
    Select Case txtStoContSz
        Case "20"
            vRevTonRateSto = lzGetRateInfo("STOOH1", txtStoContSz.Text)
        Case "40"
            vRevTonRateSto = lzGetRateInfo("STOOH2", txtStoContSz.Text)
        Case "45"
            vRevTonRateSto = lzGetRateInfo("STOOH3", txtStoContSz.Text)
    End Select
    
    lzStoOversize = vRevTon * vRevTonRateSto * (vCYMStoDay + vStoDay)
    
End Function

Private Function lzStoExtDateValid() As Boolean
    lzStoExtDateValid = False
    If Not IsDate(txtStoExtDate) Then
        MsgBox "Invalid date.  Please correct..."
    Else
        If IsDate(lblStoValidUntil) Then
            If (DateDiff("d", CDate(lblStoValidUntil), CDate(txtStoExtDate)) < 1) Or _
               (DateDiff("d", gzGetSysDate(), CDate(txtStoExtDate)) < 0) Then
                MsgBox "Extension date cannot be less than date today.  Please correct..."
            Else
                lzStoExtDateValid = True
            End If
            
        ElseIf IsDate(mskExpStorageIN.Text) Then
            If (DateDiff("d", CDate(mskExpStorageIN.Text), CDate(txtStoExtDate)) < 1) Or _
               (DateDiff("d", gzGetSysDate(), CDate(txtStoExtDate)) < 0) Then
                MsgBox "Extension date cannot be less than date today.  Please correct..."
            Else
                lzStoExtDateValid = True
            End If
            
        End If
    End If
End Function

Private Sub txtStoUOM_Validate(Cancel As Boolean)
    Cancel = (txtStoUOM <> "I") And (txtStoUOM <> "C")
End Sub

Private Sub lzComputeMsc()
Dim vatRate As Double

   'VAT Rate to be used
    If chkVAT.Visible = True And chkVAT.Value = 1 Then
        If Trim(cboVATRate.Text) = "5%" Then      '4% VAT Rate
            vatRate = 0.05
        Else                                       '10% VAT Rate
            vatRate = 0.12
        End If
    Else
        vatRate = 0
    End If
    
  ' compute
  nAmount = CCur("0" & lblMscAmount)
  nVATAmount = IIf(bVAT, Round(nAmount * vatRate, 2), 0)
  ' nWTaxAmount = IIf(bWTax, nAmount * 0.01, 0)
  nWTaxAmount = IIf(bWTax, Round(nAmount * 0.02, 2), 0) 'Modify Tax to 2%
  nTotalAmount = Round(nAmount + nVATAmount - nWTaxAmount, 2)
        
  If nTotalAmount > 0 Then
    ' new ccr
    If bNewCCR Or (nCCRCounter > 11) Then '7
        bNewCCR = False
        nCCRCounter = 1
    Else
        nCCRCounter = nCCRCounter + 1
    End If
    
    With grdCCRTran
        If HasDuplicateContNo(txtMscContNo.Text) = False Then
            nPtr = .Rows
            .AddItem nPtr
            
            If nCCRCounter = 1 Then
                .CellForeColor = &HFF0000
            End If
                
            .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
            .TextMatrix(nPtr - 1, enRateCode) = txtMscRateCode
            .TextMatrix(nPtr - 1, enAmount) = Format(nAmount, "###,##0.00")
            If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
            If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
            .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
            .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
            .TextMatrix(nPtr - 1, encustomer) = txtCusName.Text
            .TextMatrix(nPtr - 1, enimporter) = txtImporter.Text
            .TextMatrix(nPtr - 1, enratedescr) = lblMScRateDesc
            
            .TextMatrix(nPtr - 1, enCCRNo) = txtMscCCRNo
            .TextMatrix(nPtr - 1, enContNo) = txtMscContNo
            .TextMatrix(nPtr - 1, enContSz) = txtMscContSz
            .TextMatrix(nPtr - 1, enQuantity) = txtMscQty
            .TextMatrix(nPtr - 1, enRemark) = txtRemark
            
            'PRNH
            .TextMatrix(nPtr - 1, enCompCode) = cmbCompCode.Text
            
            .TextMatrix(nPtr, enCounter) = "**"
            .Row = nPtr
        Else
            MsgBox "Container Number has been used!", vbInformation, "Miscellaneous"
            Call lzClearMsc
            Exit Sub
        End If
    End With
    Call lzAddToTotal
    Call lzClearMsc
Else
    MsgBox "No additional charges computed...", vbInformation
    txtMscCCRNo.SetFocus
End If
'End With
End Sub

Public Function HasDuplicateContNo(ByVal pValue As String) As Boolean
Dim rowctr As Integer

pValue = UCase(Trim(pValue))
HasDuplicateContNo = False

If grdCCRTran.Rows >= 0 Then
 For rowctr = 1 To grdCCRTran.Rows - 1
     If pValue = UCase(Trim(grdCCRTran.TextMatrix(rowctr, enContNo))) And grdCCRTran.TextMatrix(rowctr, enRateCode) = txtMscRateCode Then
         HasDuplicateContNo = True
         Exit Function
     End If
 Next rowctr
End If
End Function

'Private Sub lzComputeParking()
'Dim vatRate As Double
'
'   'VAT Rate to be used
'    If Trim(cboVATRate.Text) = "5%" Then      '4% VAT Rate
'        vatRate = 0.05
'    Else                                       '10% VAT Rate
'        vatRate = 0.12
'    End If
'
'    ' compute
'    On Error GoTo err_Amt
'    nAmount = CCur("0" & txtParkingAMT)
'    nVATAmount = IIf(bVAT, Round(nAmount * vatRate, 2), 0)
'    'nWTaxAmount = IIf(bWTax, nAmount * 0.01, 0)
'    nWTaxAmount = IIf(bWTax, Round(nAmount * 0.02, 2), 0) 'Modify tax to 2%
'    nTotalAmount = Round(nAmount + nVATAmount - nWTaxAmount, 2)
'
'    If nTotalAmount > 0 Then
'        ' new ccr
'        If bNewCCR Or (nCCRCounter > 11) Then '7
'            bNewCCR = False
'            nCCRCounter = 1
'        Else
'            nCCRCounter = nCCRCounter + 1
'        End If
'
'        With grdCCRTran
'            'If HasDuplicateContNo(txtOthContNo.Text) = False Then
'                nPtr = .Rows
'                .AddItem nPtr
'
'                If nCCRCounter = 1 Then
'                    .CellForeColor = &HFF0000
'                End If
'
'                .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
'                .TextMatrix(nPtr - 1, enRateCode) = "PARKING"
'                .TextMatrix(nPtr - 1, enAmount) = Format(nAmount, "###,##0.00")
'                If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
'                If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
'                .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
'                .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
'                .TextMatrix(nPtr - 1, encustomer) = Trim(txtCusName.Text)
'                .TextMatrix(nPtr - 1, enimporter) = Trim(txtImporter.Text)
'
'                .TextMatrix(nPtr - 1, enContNo) = txtTruckPLT.Text
'                .TextMatrix(nPtr - 1, enVessel) = txtTruckMake.Text
'                .TextMatrix(nPtr - 1, enRemark) = txtRemark.Text & "/" & txtDriver.Text
'
'                .TextMatrix(nPtr, enCounter) = "**"
'                .Row = nPtr
'            'Else
'            '    MsgBox "Gatepass No./Container No. has been used!", vbInformation, "Others"
'            '    Call lzClearOth
'            '    Exit Sub
'            'End If
'        End With
'        Call lzAddToTotal
'        Call lzClearOth
'    Else
'        MsgBox "No additional charges computed...", vbInformation
'        txtOTHCCRNo.SetFocus
'    End If
'    Exit Sub
'
'err_Amt:
'        MsgBox "Invalid amount.  Please re-enter.", vbInformation
'        On Error GoTo 0
'        txtOthAmount.SetFocus
'End Sub

Private Sub lzComputeOth()
Dim vatRate As Double

   'VAT Rate to be used
    If chkVAT.Visible = True And chkVAT.Value = 1 Then
        If Trim(cboVATRate.Text) = "5%" Then      '4% VAT Rate
            vatRate = 0.05
        Else                                       '10% VAT Rate
            vatRate = 0.12
        End If
    Else
        vatRate = 0
    End If
    
    ' compute
    On Error GoTo err_Amt
    nAmount = CCur("0" & txtOthAmount)
    If bOTHOversize Then
        Call lzOthOversize
    End If
    nVATAmount = IIf(bVAT, Round(nAmount * vatRate, 2), 0)
'    nWTaxAmount = IIf(bWTax, nAmount * 0.01, 0)
    nWTaxAmount = IIf(bWTax, Round(nAmount * 0.02, 2), 0) 'Modify tax to 2%
    nTotalAmount = Round(nAmount + nVATAmount - nWTaxAmount, 2)
        
    If nTotalAmount > 0 Then
        ' new ccr
        If bNewCCR Or (nCCRCounter > 11) Then '7
            bNewCCR = False
            nCCRCounter = 1
        Else
            nCCRCounter = nCCRCounter + 1
        End If
        
        With grdCCRTran
            If HasDuplicateContNo(txtOthContNo.Text) = False Then
                nPtr = .Rows
                .AddItem nPtr
                    
                If nCCRCounter = 1 Then
                    .CellForeColor = &HFF0000
                End If
                           
                .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
                .TextMatrix(nPtr - 1, enRateCode) = "OTHERS"
                .TextMatrix(nPtr - 1, enAmount) = Format(nAmount, "###,##0.00")
                If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
                If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
                .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
                .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
                .TextMatrix(nPtr - 1, encustomer) = Trim(txtCusName.Text)
                .TextMatrix(nPtr - 1, enimporter) = Trim(txtImporter.Text)

                .TextMatrix(nPtr - 1, enCCRNo) = txtOTHCCRNo
                .TextMatrix(nPtr - 1, enContNo) = txtOthContNo
                .TextMatrix(nPtr - 1, enContSz) = txtOthContSz
                .TextMatrix(nPtr - 1, enFulEmp) = txtOthFulEmp
                .TextMatrix(nPtr - 1, enEntryNo) = txtOthEntryNo
                .TextMatrix(nPtr - 1, enRegNo) = txtOthRegNo
                If bOTHOversize Then
                    .TextMatrix(nPtr - 1, enOvzLen) = txtOthOvzLen
                    .TextMatrix(nPtr - 1, enOvzWid) = txtOthOvzWid
                    .TextMatrix(nPtr - 1, enOvzHgt) = txtOthOvzHgt
                    .TextMatrix(nPtr - 1, enOvzUom) = txtOthUOM
                    .TextMatrix(nPtr - 1, enRevTon) = lblOthRevTon
                    .TextMatrix(nPtr - 1, enOvzAmt) = 0
                End If
                .TextMatrix(nPtr - 1, enVessel) = txtOthVessel
                .TextMatrix(nPtr - 1, enRemark) = txtRemark
                
                'PRNH - Company Code
                .TextMatrix(nPtr - 1, enCompCode) = cmbCompCode.Text
                
                .TextMatrix(nPtr, enCounter) = "**"
                .Row = nPtr
            Else
                MsgBox "Gatepass No./Container No. has been used!", vbInformation, "Others"
                Call lzClearOth
                Exit Sub
            End If
        End With
        Call lzAddToTotal
        Call lzClearOth
    Else
        MsgBox "No additional charges computed...", vbInformation
        txtOTHCCRNo.SetFocus
    End If
    Exit Sub

err_Amt:
        MsgBox "Invalid amount.  Please re-enter.", vbInformation
        On Error GoTo 0
        txtOthAmount.SetFocus
End Sub

Private Sub lzOthOversize()
Dim pLength, pWidth, pHeight As Single
    
    pLength = CSng(txtOthOvzLen)
    pWidth = CSng(txtOthOvzWid)
    pHeight = CSng(txtOthOvzHgt)
    
    If txtOthUOM = "C" Then
       pLength = pLength / 2.54
       pWidth = pWidth / 2.54
       pHeight = pHeight / 2.54
    End If
    
    vRevTon = pLength * pWidth * pHeight / 1728 / 40
    Select Case txtOthContSz
        Case "20"
            If vRevTon >= cRTon20 Then vRevTon = vRevTon - cRTon20
        Case "40"
            If vRevTon >= cRTon40 Then vRevTon = vRevTon - cRTon40
        Case "45"
            If vRevTon >= cRTon45 Then vRevTon = vRevTon - cRTon45
        Case Else
            vRevTon = 0
    End Select
    vRevTon = Round(vRevTon, 2)
    If vRevTon > 0 Then
        lblOthRevTon = Format(vRevTon, "###,##0.00")
    Else
        lblOthRevTon = ""
    End If

End Sub

Private Sub lzShowRate()
    Call lzGetRateInfo(txtMscRateCode, txtMscContSz)
    If vRateCode <> "" Then
        txtMscRateCode = vRateCode
        lblMScRateDesc = vRateDesc
        txtMscContSz = vRateSz
        lblMscRateAmt = Format(vRateAmount, "###,##0.00")
        lblMscRateUOM = vRateUOM
        lblMscAmount = Format(CCur("0" & lblMscRateAmt) * CCur("0" & txtMscQty), "###,##0.00")
        If (Trim(txtMscQty) = "") And (vRateAmount > 0) Then txtMscQty = 1
        txtMscQty.SetFocus
    End If
End Sub

Private Sub lzComputeRfr()
Dim h, m As Single
Dim bDateError As Boolean
Dim vatRate As Double

    'VAT Rate to be used
    If chkVAT.Visible = True And chkVAT.Value = 1 Then
        If Trim(cboVATRate.Text) = "5%" Then      '4% VAT Rate
            vatRate = 0.05
        Else                                       '10% VAT Rate
            vatRate = 0.12
        End If
    Else
        vatRate = 0
    End If
    
    bDateError = (txtRfrPlugInDate = cEmptyRfrDate) Or _
                 (lblRfrValidUntil = cEmptyRfrDate) Or _
                 (txtRfrExtDate = cEmptyRfrDate) Or _
                 (lblRfrValidUntil = "")
    If bDateError Then
       MsgBox "One or more date is invalid. Please correct.", vbExclamation
       Exit Sub
    End If
    
    ' compute
    If txtRfrContSz.Text = "20" Then
        cRateCode = "MCRFC2" '"IMRF"
    Else
        cRateCode = "MCRFC3" '"IMRF"
    End If
    
    'Edited by Navis Project Team
    'nAmount = lzGetRateInfo(cRateCode, txtRfrContSz)
    nAmount = lzGetRateInfo(cRateCode, IIf(txtRfrContSz.Text = "20", txtRfrContSz.Text, "40"))
    
    h = DateDiff("n", CDate(lblRfrValidUntil), CDate(txtRfrExtDate))
    
    h = h / 60
    m = h - Fix(h)
    vRfrHours = Fix(h / 6) * 6
    m = m + ((h / 6) - Fix(h / 6))
    If m > 0 Then vRfrHours = vRfrHours + 6
    If vRfrHours < 6 Then vRfrHours = 6
    txtRfrExtDate = Format(DateAdd("h", vRfrHours, CDate(lblRfrValidUntil)), "YYYY-MM-DD hh:mm")
    nAmount = nAmount * (vRfrHours / 6)
    nVATAmount = IIf(bVAT, Round(nAmount * vatRate, 2), 0)
'    nWTaxAmount = IIf(bWTax, nAmount * 0.01, 0)
     nWTaxAmount = IIf(bWTax, Round(nAmount * 0.02, 2), 0) 'Modify Tax to 2%
    nTotalAmount = Round(nAmount + nVATAmount - nWTaxAmount, 2)
        
    If nTotalAmount > 0 Then
        ' new ccr
        If bNewCCR Or (nCCRCounter > 11) Then '7
            bNewCCR = False
            nCCRCounter = 1
        Else
            nCCRCounter = nCCRCounter + 1
        End If
        
        With grdCCRTran
            If HasDuplicateGatePass(txtRFRCCRNo.Text) = False Then
                nPtr = .Rows
                .AddItem nPtr
                
                If nCCRCounter = 1 Then
                    .CellForeColor = &HFF0000
                End If
            
                .TextMatrix(nPtr - 1, enCounter) = nPtr - 1
                .TextMatrix(nPtr - 1, enRateCode) = cRateCode
                .TextMatrix(nPtr - 1, enAmount) = Format(nAmount, "###,##0.00")
                If nVATAmount > 0 Then .TextMatrix(nPtr - 1, enVATAmt) = Format(nVATAmount, "##,##0.00")
                If nWTaxAmount > 0 Then .TextMatrix(nPtr - 1, enWTaxAmt) = Format(nWTaxAmount, "##,##0.00")
                .TextMatrix(nPtr - 1, enTotalAmt) = Format(nTotalAmount, "##,##,##0.00")
                .TextMatrix(nPtr - 1, enCCRTag) = IIf(nCCRCounter = 1, "*", " ")
                .TextMatrix(nPtr - 1, encustomer) = txtCusName.Text
                .TextMatrix(nPtr - 1, enimporter) = txtImporter.Text

                .TextMatrix(nPtr - 1, enCCRNo) = txtRFRCCRNo
                .TextMatrix(nPtr - 1, enContNo) = txtRfrContNo
                .TextMatrix(nPtr - 1, enContSz) = txtRfrContSz
                .TextMatrix(nPtr - 1, enEntryNo) = txtRfrEntryNo
                .TextMatrix(nPtr - 1, enRegNo) = txtRfrRegNo
                .TextMatrix(nPtr - 1, enRfrHours) = vRfrHours
                .TextMatrix(nPtr - 1, enRfrValidUntil) = txtRfrExtDate
                
                txtRemark = "REEFER UP TO " & Format(txtRfrExtDate, "YYYY/MM/DD hh:mm")
                .TextMatrix(nPtr - 1, enRemark) = txtRemark
                
                .TextMatrix(nPtr, enCounter) = "**"
                .Row = nPtr
            Else
                MsgBox "Gatepass Number has been used!", vbInformation, "Reefer"
                Call lzClearRfr
                Exit Sub
            End If
        End With
        Call lzAddToTotal
        'MsgBox "Reefer extended by " & vRfrHours & " hours until " & txtRfrExtDate, vbInformation
        Call lzClearRfr
    Else
        MsgBox "No additional charges computed...", vbInformation
        txtRFRCCRNo.SetFocus
    End If
End Sub

Private Sub lzInitializePay()
Dim n As Integer
Dim intdtl As Integer

    If Not fraPayment.Enabled Then fraPayment.Enabled = True
    If chkGuarantee.Value = 0 Then
        txtCshAmt.Text = lblAmtDue
    Else
        txtCshAmt.Text = ".00"
    End If
    For n = 0 To 4
        txtChkAmt(n) = ".00"
        txtChkNo(n) = Space(txtChkNo(n).MaxLength)
        txtChkBank(n) = Space(txtChkBank(n).MaxLength)
    Next n
    lblChkTot = ".00"
    lblChange = ".00"
    txtCshAmt.SetFocus
    txtCshAmt.SelStart = 0
    txtCshAmt.SelLength = Len(txtCshAmt.Text)

End Sub

Private Sub lzDeleteItem()
Dim bNewCCR As Boolean
Dim n As Integer
    With grdCCRTran
        If (.Row < .Rows) And (.TextMatrix(.Row, enRateCode) <> cVoid) Then
            bNewCCR = (.TextMatrix(.Row, enCCRTag) = "*")
            Call lzLessFromTotal
            .RemoveItem .Row
            .AddItem "", .Row
            .TextMatrix(.Row, enRateCode) = cVoid
            If bNewCCR Then
                n = .Row + 1
                While n < (.Rows - 1)
                    If (.TextMatrix(n, enRateCode) <> cVoid) Then
                        .TextMatrix(n, enCCRTag) = "*"
                        n = .Rows
                    Else
                        n = n + 1
                    End If
                Wend
            End If
        End If
        .SetFocus
    End With
End Sub

Private Sub lzLessFromTotal()
Dim curTotal As Currency
    curTotal = CCur("0" & lblAmtDue)
    nTotalAmount = CCur("0" & grdCCRTran.TextMatrix(grdCCRTran.Row, enTotalAmt))
    lblAmtDue = Format(curTotal - nTotalAmount, "#,###,##0.00")
    lblAmtDue.Refresh
End Sub

Private Sub lzAddTran()
    If Not grdCCRTran.Enabled Then grdCCRTran.Enabled = True
    If Not mnuMenuPayment.Enabled Then mnuMenuPayment.Enabled = True
    
    Select Case tabTran.Tab
        Case 0
            Call lzComputeArr
        Case 1
            Call lzComputeSto
        Case 2
            Call lzComputeRfr
        Case 3
            Call lzComputeSOC
            'Call lzComputeSO
        Case 4
            Call lzComputeMsc
        Case 5
            Call lzComputeOth
'        Case 6
'            Call lzComputeParking
        Case Else
    End Select
        
    chkNewCCR.Value = IIf(bNewCCR, 1, 0)
End Sub

Private Sub lzEnableTab()
Dim n As Integer
    With tabTran
        For n = 0 To .Tabs - 1
            .TabEnabled(n) = IIf(n = vTabOn, True, False)
        Next n
        .Tab = vTabOn
    End With
End Sub

'Private Function lzGetNextCCR(ByVal pUserID As String) As Long
'Dim cmd As ADODB.Command
'    ' create command
'    Set cmd = New ADODB.Command
'    With cmd
'        Set .ActiveConnection = gcnnBilling
'        .CommandText = "up_getnextspl"
'        .CommandType = adCmdStoredProc
'
'        .Parameters(1).Type = adChar
'        .Parameters(1).Value = pUserID
'        .Parameters(1).Direction = adParamInput             ' user id
'        .Parameters(2).Type = adInteger
'        .Parameters(2).Direction = adParamOutput            ' next ccr
'
'        .Execute
'        lzGetNextCCR = .Parameters(2)
'
'    End With
'  '  Set cmd = Nothing
'End Function

Private Function lzGetNextCCR(ByVal pUserID As String, ByVal pCOMPCODE As String) As Long
    Dim cmd As ADODB.Command
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getnextspl"
        .CommandType = adCmdStoredProc
        
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUserID
        .Parameters(1).Direction = adParamInput             ' user id
        
         'PRNH
        .Parameters(2).Type = adChar
        .Parameters(2).Value = pCOMPCODE
        .Parameters(2).Direction = adParamInput            ' company code
        
        .Parameters(3).Type = adInteger
        .Parameters(3).Direction = adParamOutput            ' next ccr
    
        .Execute
        
        lzGetNextCCR = .Parameters(3)
    
    End With
    Set cmd = Nothing

End Function
    
Private Function lzCCRValid(ByVal pUserID As String, ByVal pCCRNo As Long, ByVal pCOMPCODE As String) As Boolean
Dim cmd As ADODB.Command
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_chkvalidspl"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = UCase(Trim(pUserID))
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adInteger
        .Parameters(2).Value = pCCRNo
        .Parameters(2).Direction = adParamInput
       
        'PRNH
        .Parameters(3).Type = adChar
        .Parameters(3).Value = pCOMPCODE
        .Parameters(3).Direction = adParamInput
       
        .Execute
        
        lzCCRValid = (.Parameters(0) > 0)
    
    End With
    
End Function

Private Function lzGetControlNo() As Long
Dim cmd As ADODB.Command
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getcontrolno"
        .CommandType = adCmdStoredProc
        
        .Parameters(1).Type = adChar
        .Parameters(1).Value = "CCR"
        .Parameters(1).Direction = adParamInput             '
        .Parameters(2).Type = adInteger
        .Parameters(2).Direction = adParamOutput            ' control number
    
        .Execute
        
        lzGetControlNo = .Parameters(2)
    
    End With
    Set cmd = Nothing
End Function

Private Sub lzApplyCCR(ByVal pUserID As String, ByVal pCCRNo As Long, pCOMPCODE As String)
Dim cmd As ADODB.Command
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_applyccrspl"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUserID
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adInteger
        .Parameters(2).Value = pCCRNo
        .Parameters(2).Direction = adParamInput

        'PRNH
        .Parameters(3).Type = adChar
        .Parameters(3).Value = pCOMPCODE
        .Parameters(3).Direction = adParamInput
        .Execute

    End With
    
End Sub

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
        Set prmGetADRBal = .CreateParameter(, adChar, adParamInput, 6, pCode)
        .Parameters.Append prmGetADRBal
        Set prmGetADRBal = .CreateParameter("pTYPE", adCurrency, adParamOutput)
        .Parameters.Append prmGetADRBal
        .Execute
        If Not IsNull(.Parameters("pTYPE")) Then
            lzGetADRBal = .Parameters("pTYPE")
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

Private Function lzApplyADR(ByVal pCUSCDE As String, _
                            ByVal pREFTYP As String, _
                            ByVal pRefnum As Long, _
                            ByVal pAdrAmt As Currency, _
                            ByVal pUserID As String, _
                            ByVal pREMARK As String) As Long
                            
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
        .Parameters(1).Value = " "
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adChar
        .Parameters(2).Value = pREFTYP
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adNumeric
        .Parameters(3).Value = pRefnum
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Value = pAdrAmt
        .Parameters(4).Direction = adParamInput
        .Parameters(5).Type = adChar
        .Parameters(5).Value = pREMARK
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

Private Sub lzSavePrint()
    Dim cmd As ADODB.Command
    Dim vRef, vSeq, vItem, vCCR, n As Long
    Dim bValidCCR As Boolean
    Dim v As Long
    Dim c As New CWaitCursor
    Dim strGKey As String
    Dim bIsOOG As Boolean
    Dim bIsDG As Boolean
    Dim strContainerNo As String
    Dim strOOGGranted As String
    Dim strDGGranted As String
    Dim intCCRNum As String
    Dim N4CurrentCategory As String
    Dim bHasUnitOut As Boolean
    'Added by AGH 01222010
    Dim bHasCashAmount As Boolean
    Dim bHasCheckAmount As Boolean
    Dim bPaymentOK As Boolean

    'PRNH
    Dim bHasADRAmount As Boolean
    
    
    'Dim clsCCRReprint As Object
    ' validate required info
    bIsOOG = False
    bIsDG = False
    strContainerNo = ""
'    Set clsCCRReprint = CreateObject("CCRPR03.clsCCRPR03")
    If Len(Trim(txtCusName)) = 0 Then
        MsgBox "Customer name required...", vbInformation
        txtCusName.SetFocus
        Exit Sub
    End If
    
    'Added by AGH 012222010
    bHasCashAmount = False
    bHasCheckAmount = False
    bPaymentOK = False
    
    'PRNH
    bHasADRAmount = False
    If Len(Trim(txtCshAmt.Text)) > 0 Then
        If IsNumeric(Trim(txtCshAmt.Text)) = True And CCur("0" & Trim(txtCshAmt.Text)) > 0 Then
            bHasCashAmount = True
        End If
    End If
    
    'Removed by PRNH
'    If Len(Trim(txtChkAmt(0).Text)) > 0 Then
'        If IsNumeric(Trim(txtChkAmt(0).Text)) = True And CCur("0" & Trim(txtChkAmt(0).Text)) > 0 Then
'            bHasCheckAmount = True
'        End If
'    End If
'
'    If Len(Trim(txtChkAmt(1).Text)) > 0 And bHasCheckAmount = False Then
'        If IsNumeric(Trim(txtChkAmt(1).Text)) = True And CCur("0" & Trim(txtChkAmt(1).Text)) > 0 Then
'            bHasCheckAmount = True
'        End If
'    End If
'
'    If Len(Trim(txtChkAmt(2).Text)) > 0 And bHasCheckAmount = False Then
'        If IsNumeric(Trim(txtChkAmt(2).Text)) = True And CCur("0" & Trim(txtChkAmt(2).Text)) > 0 Then
'            bHasCheckAmount = True
'        End If
'    End If
'
'    If Len(Trim(txtChkAmt(3).Text)) > 0 And bHasCheckAmount = False Then
'        If IsNumeric(Trim(txtChkAmt(3).Text)) = True And CCur("0" & Trim(txtChkAmt(3).Text)) > 0 Then
'            bHasCheckAmount = True
'        End If
'    End If
'
'    If Len(Trim(txtChkAmt(4).Text)) > 0 And bHasCheckAmount = False Then
'        If IsNumeric(Trim(txtChkAmt(4).Text)) = True And CCur("0" & Trim(txtChkAmt(4).Text)) > 0 Then
'            bHasCheckAmount = True
'        End If
'    End If

    'PRNH - Simplified version of the above
    Dim ctr As Integer
    For ctr = 0 To 4
        If Len(Trim(txtChkAmt(ctr).Text)) > 0 And bHasCheckAmount = False Then
            If IsNumeric(Trim(txtChkAmt(ctr).Text)) = True And CCur("0" & Trim(txtChkAmt(ctr).Text)) > 0 Then
                bHasCheckAmount = True
        End If
    End If
    Next
    
    
    'prnh - ADR Validation
     If Len(Trim(txtADRAmt.Text)) > 0 Then
        If IsNumeric(Trim(txtADRAmt.Text)) = True And CCur("0" & Trim(txtADRAmt.Text)) > 0 Then
            bHasADRAmount = True
        End If
    End If
    
    If chkGuarantee.Value = 1 Then
        bPaymentOK = True
    Else
        If bHasCashAmount = True Or bHasCheckAmount = True Or bHasADRAmount = True Then
            bPaymentOK = True
        Else
            bPaymentOK = False
        End If
    End If
    
    
    If bPaymentOK = False Then
        MsgBox "Payment is required...", vbInformation
        txtCshAmt.SetFocus
        Exit Sub
    End If
    bValidCCR = False
    vCCR = lzGetNextCCR(gUserID, grdCCRTran.TextMatrix(1, enCompCode))
    
    'PRNH - Check if valid CCR
    bValidCCR = lzCCRValid(gUserID, vCCR, grdCCRTran.TextMatrix(1, enCompCode))
    
    While Not bValidCCR
        'v = CLng("0" & Trim(InputBox("Enter CCR number: ", , Str(vCCR))))
        vNextCCR = vCCR
        frmNextCCR.Show 1
        v = vNextCCR
        
        If v > 0 Then
            'temporary hardcoded
            bValidCCR = lzCCRValid(gUserID, v, grdCCRTran.TextMatrix(1, enCompCode))
            If bValidCCR Then vCCR = v
        Else
            bValidCCR = True
            txtCshAmt.SetFocus
            Exit Sub
        End If
    Wend
    
    vRef = lzGetControlNo
    If vRef <= 0 Then
        MsgBox "CCR control number request error.  Please try again later.", vbInformation
        Exit Sub
    End If
    
    ' create command
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        
        ' write payment first
        .CommandText = "up_wrtccrpay"
        .CommandType = adCmdStoredProc
        
        .Parameters(1).Type = adInteger
        .Parameters(1).Value = vRef
        .Parameters(1).Direction = adParamInput             ' control number
        .Parameters(2).Type = adChar
        If bUnderG Then
            '.Parameters(2).Value = vCusCodeUnderG
            .Parameters(2).Value = Text1.Text
        Else
            .Parameters(2).Value = ""
        End If
        .Parameters(2).Direction = adParamInput             '
        .Parameters(3).Type = adChar
        .Parameters(3).Value = txtCusName
        .Parameters(3).Direction = adParamInput
        '
        'Edited AGH 01202010
        .Parameters(4).Type = adCurrency
'        .Parameters(4).Value = CCur("0" & txtCshAmt) + CCur("0" & txtChkAmt(0)) + CCur("0" & txtChkAmt(1)) + _
'                               CCur("0" & txtChkAmt(2)) + CCur("0" & txtChkAmt(3)) + CCur("0" & txtChkAmt(4))
        .Parameters(4).Value = CCur("0" & txtCshAmt)
        .Parameters(4).Direction = adParamInput             '
        
        .Parameters(5).Type = adCurrency
        'PRNH - ADR Amount
        .Parameters(5).Value = CCur("0" & txtADRAmt.Text)
        .Parameters(5).Direction = adParamInput             '
        
        .Parameters(6).Type = adInteger
        .Parameters(6).Value = 0
        .Parameters(6).Direction = adParamInput             '
        .Parameters(7).Type = adCurrency
        .Parameters(7).Value = CCur("0" & lblChange)
        .Parameters(7).Direction = adParamInput             '
        .Parameters(8).Type = adChar
        .Parameters(8).Value = txtChkNo(0)
        .Parameters(8).Direction = adParamInput             '
        .Parameters(9).Type = adChar
        .Parameters(9).Value = txtChkNo(1)
        .Parameters(9).Direction = adParamInput             '
        .Parameters(10).Type = adChar
        .Parameters(10).Value = txtChkNo(2)
        .Parameters(10).Direction = adParamInput             '
        .Parameters(11).Type = adChar
        .Parameters(11).Value = txtChkNo(3)
        .Parameters(11).Direction = adParamInput             '
        .Parameters(12).Type = adChar
        .Parameters(12).Value = "" 'txtChkNo(4).Text
        .Parameters(12).Direction = adParamInput             '
        .Parameters(13).Type = adNumeric
        .Parameters(13).Value = CCur("0" & txtChkAmt(0))
        .Parameters(13).Direction = adParamInput             '
        .Parameters(14).Type = adNumeric
        .Parameters(14).Value = CCur("0" & txtChkAmt(1))
        .Parameters(14).Direction = adParamInput             '
        .Parameters(15).Type = adNumeric
        .Parameters(15).Value = CCur("0" & txtChkAmt(2))
        .Parameters(15).Direction = adParamInput             '
        .Parameters(16).Type = adNumeric
        .Parameters(16).Value = CCur("0" & txtChkAmt(3))
        .Parameters(16).Direction = adParamInput             '
        .Parameters(17).Type = adNumeric
        .Parameters(17).Value = CCur("0" & txtChkAmt(4))
        .Parameters(17).Direction = adParamInput             '
        .Parameters(18).Type = adChar
        .Parameters(18).Value = txtChkBank(0)
        .Parameters(18).Direction = adParamInput             '
        .Parameters(19).Type = adChar
        .Parameters(19).Value = txtChkBank(1)
        .Parameters(19).Direction = adParamInput             '
        .Parameters(20).Type = adChar
        .Parameters(20).Value = txtChkBank(2)
        .Parameters(20).Direction = adParamInput             '
        .Parameters(21).Type = adChar
        .Parameters(21).Value = txtChkBank(3)
        .Parameters(21).Direction = adParamInput             '
        .Parameters(22).Type = adChar
        .Parameters(22).Value = txtChkBank(4)
        .Parameters(22).Direction = adParamInput             '
        .Parameters(23).Type = adChar
        .Parameters(23).Value = gUserID
        .Parameters(23).Direction = adParamInput             '

        .Execute
    
        ' write details next
        .CommandText = "up_wrtccrdtl"
        .CommandType = adCmdStoredProc
    
        vSeq = 0: vItem = 0: vCCR = vCCR - 1
        For n = 1 To (grdCCRTran.Rows - 2)
            If (grdCCRTran.TextMatrix(n, enRateCode) <> cVoid) Then
                If (n = 1) Or (grdCCRTran.TextMatrix(n, enCCRTag) = "*") Then
                    vSeq = vSeq + 1
                    vCCR = vCCR + 1
                    vItem = 0
                    
                    ' re check allocation
                    'temporary hardcoded
                    bValidCCR = lzCCRValid(gUserID, vCCR, grdCCRTran.TextMatrix(n, enCompCode))
                    If Not bValidCCR Then
                        While Not bValidCCR
                            vNextCCR = vCCR
                            frmNextCCR.Show 1
                            v = vNextCCR
                            If v > 0 Then
                                'temporary hardcoded
                                bValidCCR = lzCCRValid(gUserID, v, grdCCRTran.TextMatrix(n, enCompCode))
                                If bValidCCR Then vCCR = v
                            Else
                                bValidCCR = True
                                MsgBox "Please void this transaction!!" & vbCrLf _
                                        & "Reference No. " & Val(vRef), vbCritical
                                
                                
                                Call lzApplyCCR(gUserID, vCCR, grdCCRTran.TextMatrix(n, enCompCode))
                                Call lzInitialize
                                Call lzGetUserInfo
                                
                                GoTo tag_ReInit
                            End If
                        Wend
                    End If
                    
                End If
                vItem = vItem + 1
                If vSeq = 0 Then vSeq = 1
                ' set parameters then execute
                .Parameters(1).Type = adInteger
                .Parameters(1).Value = vRef
                .Parameters(1).Direction = adParamInput             '
                .Parameters(2).Type = adInteger
                .Parameters(2).Value = vSeq
                .Parameters(2).Direction = adParamInput             '
                .Parameters(3).Type = adInteger
                .Parameters(3).Value = vItem
                .Parameters(3).Direction = adParamInput             '
                .Parameters(4).Type = adInteger
                .Parameters(4).Value = vCCR
                .Parameters(4).Direction = adParamInput             '
                .Parameters(5).Type = adChar
                .Parameters(5).Value = Trim(grdCCRTran.TextMatrix(n, enRateCode)) ' Error part
                .Parameters(5).Direction = adParamInput
                
                .Parameters(6).Type = adNumeric
                'nDwellDays
                If nDwellDays > 0 And (Trim(grdCCRTran.TextMatrix(n, enRateCode)) = "STODO1" Or Trim(grdCCRTran.TextMatrix(n, enRateCode)) = "STODO2" Or Trim(grdCCRTran.TextMatrix(n, enRateCode)) = "STODO3") Then
                    .Parameters(6).Value = CCur("0" & grdCCRTran.TextMatrix(n, enAmount)) * nDwellDays
                    nDwellDays = 0
                Else
                    .Parameters(6).Value = CCur("0" & grdCCRTran.TextMatrix(n, enAmount))
                End If
                .Parameters(6).Direction = adParamInput             '
                
                .Parameters(7).Type = adNumeric
                .Parameters(7).Value = CCur("0" & grdCCRTran.TextMatrix(n, enVATAmt))
                .Parameters(7).Direction = adParamInput             '
                .Parameters(8).Type = adNumeric
                .Parameters(8).Value = CCur("0" & grdCCRTran.TextMatrix(n, enWTaxAmt))
                .Parameters(8).Direction = adParamInput             '
                .Parameters(9).Type = adInteger
                .Parameters(9).Value = CLng("0" & grdCCRTran.TextMatrix(n, enCCRNo))
                If CStr(grdCCRTran.TextMatrix(n, enCCRNo)) <> "" Then
                    intCCRNum = grdCCRTran.TextMatrix(n, enCCRNo)
                Else
                    intCCRNum = 0
                End If
                .Parameters(9).Direction = adParamInput             '
                .Parameters(10).Type = adChar
                .Parameters(10).Value = grdCCRTran.TextMatrix(n, enContNo)    'B
                .Parameters(10).Direction = adParamInput             '
                strContainerNo = grdCCRTran.TextMatrix(n, enContNo)
                .Parameters(11).Type = adNumeric
                .Parameters(11).Value = CLng("0" & grdCCRTran.TextMatrix(n, enContSz))
                .Parameters(11).Direction = adParamInput             '
                .Parameters(12).Type = adChar
                .Parameters(12).Value = grdCCRTran.TextMatrix(n, enFulEmp)  'C
                .Parameters(12).Direction = adParamInput             '
                .Parameters(13).Type = adInteger
                .Parameters(13).Value = CLng("0" & grdCCRTran.TextMatrix(n, enEntryNo))
                .Parameters(13).Direction = adParamInput             '
                .Parameters(14).Type = adChar
                .Parameters(14).Value = grdCCRTran.TextMatrix(n, enRegNo)  'D
                .Parameters(14).Direction = adParamInput
                .Parameters(15).Type = adNumeric
                .Parameters(15).Value = CCur("0" & grdCCRTran.TextMatrix(n, enOvzLen))
                .Parameters(15).Direction = adParamInput             '
                .Parameters(16).Type = adNumeric
                .Parameters(16).Value = CCur("0" & grdCCRTran.TextMatrix(n, enOvzWid))
                .Parameters(16).Direction = adParamInput             '
                .Parameters(17).Type = adNumeric
                .Parameters(17).Value = CCur("0" & grdCCRTran.TextMatrix(n, enOvzHgt))
                .Parameters(17).Direction = adParamInput             '
                'Added of Navis Project Team 10/29/2009
                If CCur("0" & grdCCRTran.TextMatrix(n, enOvzLen)) > 0 And CCur("0" & grdCCRTran.TextMatrix(n, enOvzWid)) > 0 And CCur("0" & grdCCRTran.TextMatrix(n, enOvzHgt)) > 0 Then
                    bIsOOG = True
                End If
                
                .Parameters(18).Type = adChar
                .Parameters(18).Value = grdCCRTran.TextMatrix(n, enOvzUom) 'F
                .Parameters(18).Direction = adParamInput             '
                .Parameters(19).Type = adNumeric
                .Parameters(19).Value = CCur("0" & grdCCRTran.TextMatrix(n, enRevTon))
                .Parameters(19).Direction = adParamInput             '
                .Parameters(20).Type = adChar
                .Parameters(20).Value = Left(grdCCRTran.TextMatrix(n, enDangerCode), 1)  'G
                .Parameters(20).Direction = adParamInput             '
                .Parameters(21).Type = adDate
                .Parameters(21).Value = CDate(IIf(IsDate(grdCCRTran.TextMatrix(n, enStoValidUntil)), grdCCRTran.TextMatrix(n, enStoValidUntil), cNullDate))
                .Parameters(21).Direction = adParamInput             '
                .Parameters(22).Type = adDate
                .Parameters(22).Value = CDate(IIf(IsDate(grdCCRTran.TextMatrix(n, enRfrValidUntil)), grdCCRTran.TextMatrix(n, enRfrValidUntil), cNullDate))
                .Parameters(22).Direction = adParamInput             '
                .Parameters(23).Type = adNumeric
                .Parameters(23).Value = CLng("0" & grdCCRTran.TextMatrix(n, enStoDays))
                .Parameters(23).Direction = adParamInput             '
                .Parameters(24).Type = adNumeric
                .Parameters(24).Value = CCur("0" & grdCCRTran.TextMatrix(n, enQuantity))
                .Parameters(24).Direction = adParamInput             '
                .Parameters(25).Type = adChar
                .Parameters(25).Value = grdCCRTran.TextMatrix(n, enVessel)  'H
                .Parameters(25).Direction = adParamInput             '
                .Parameters(26).Type = adNumeric
                If CCur("0" & Trim(grdCCRTran.TextMatrix(n, enDangerAmt))) > 0 Then
                    .Parameters(26).Value = CCur("0" & grdCCRTran.TextMatrix(n, enAmount))
                    .Parameters(6).Value = 0
                    'Added by Navis Project Team 11/04/2009
                    bIsDG = True
                Else
                    .Parameters(26).Value = 0
                    'Added by Navis Project Team 11/04/2009
                    bIsDG = False
                End If
                .Parameters(26).Direction = adParamInput             '
                .Parameters(27).Type = adNumeric
                If Trim(grdCCRTran.TextMatrix(n, enOvzAmt)) <> "" Then
                    .Parameters(27).Value = CCur("0" & grdCCRTran.TextMatrix(n, enAmount))
                    .Parameters(6).Value = 0
                Else
                    .Parameters(27).Value = 0
                End If
                .Parameters(27).Direction = adParamInput             '
                .Parameters(28).Type = adVarChar
                .Parameters(28).Size = 30
                .Parameters(28).Value = Left(grdCCRTran.TextMatrix(n, enRemark), 30) 'I
                .Parameters(28).Direction = adParamInput             '
                .Parameters(29).Type = adChar
                .Parameters(29).Value = grdCCRTran.TextMatrix(n, enShipLine)   'J
                .Parameters(29).Direction = adParamInput             '
                .Parameters(30).Type = adChar
                .Parameters(30).Value = IIf(bUnderG, "Y", " ")
                .Parameters(30).Direction = adParamInput             '
                .Parameters(31).Type = adNumeric
                .Parameters(31).Value = CLng("0" & grdCCRTran.TextMatrix(n, enRfrHours))
                .Parameters(31).Direction = adParamInput             '
                .Parameters(32).Type = adChar
                .Parameters(32).Value = Left(Trim(gUserID), 10)
                .Parameters(32).Direction = adParamInput             '
                
                 'PRNH - Company Code
                .Parameters(33).Type = adVarChar
                .Parameters(33).Size = 10
                .Parameters(33).Value = grdCCRTran.TextMatrix(n, enCompCode)
                .Execute
        
            End If
            'Save to Sparcs
'            ConnectToNavis
            If strContainerNo <> "" Then
                If Left(Trim(grdCCRTran.TextMatrix(n, enRateCode)), 5) = "MCRFC" Then
                    strGKey = GetGKey(strContainerNo, "INVOICED", "REEFER")
                    Call SavePaymentToSparcs(strContainerNo, "REEFER", nRfrExtDate, strGKey)
                    'Call ReleaseHold(grdCCRTran.TextMatrix(n, enContNo))
                ElseIf InStr(1, Trim(grdCCRTran.TextMatrix(n, enRateCode)), "STOIM") > 0 Or InStr(1, Trim(grdCCRTran.TextMatrix(n, enRateCode)), "STOEX") > 0 Then
                    If InStr(1, Trim(grdCCRTran.TextMatrix(n, enRateCode)), "STOEX") > 0 Then
                        Call GetContainerLastestCategory(strContainerNo, N4CurrentCategory, bHasUnitOut)
                        If N4CurrentCategory = "EXPRT" And bHasUnitOut = False Then
                            strGKey = GetGKey(grdCCRTran.TextMatrix(n, enContNo), "INVOICED", "STORAGE")
                            Call SavePaymentToSparcs(grdCCRTran.TextMatrix(n, enContNo), "STORAGE", nStoExtDate, strGKey)
                        End If
                    Else
                        strGKey = GetGKey(grdCCRTran.TextMatrix(n, enContNo), "INVOICED", "STORAGE")
                        Call SavePaymentToSparcs(grdCCRTran.TextMatrix(n, enContNo), "STORAGE", nStoExtDate, strGKey)
                    End If
                    
                    'Commented by Navis Project Team 11/13/2009
                    'If bIsOOG Then
                    '   Call ReleaseOOG(strContainerNo)
                    'End If
                ElseIf Trim(grdCCRTran.TextMatrix(n, enRateCode)) = "MCTRUS" Then
                    WeighHold (grdCCRTran.TextMatrix(n, enContNo))
                ElseIf Trim(grdCCRTran.TextMatrix(n, enRateCode)) = "STODO1" Or Trim(grdCCRTran.TextMatrix(n, enRateCode)) = "STODO2" Or Trim(grdCCRTran.TextMatrix(n, enRateCode)) = "STODO3" Then
                    'SOC lift grant BILLING permission at N4
                    If ((n - 1) >= 0) And Trim(grdCCRTran.TextMatrix(n - 1, enRateCode)) = "MCLIF3" Then
                        'lift billing hold for SOC container
                        ReleaseBilling ((grdCCRTran.TextMatrix(n, enContNo)))
                    End If
    
                ElseIf InStr(1, Trim(grdCCRTran.TextMatrix(n, enRateCode)), "CBIMP") > 0 Or InStr(1, Trim(grdCCRTran.TextMatrix(n, enRateCode)), "CBEXP") > 0 Then

                    If InStr(1, Trim(grdCCRTran.TextMatrix(n, enRateCode)), "CBEXP") > 0 Then
                        Call GetContainerLastestCategory(strContainerNo, N4CurrentCategory, bHasUnitOut)
                    End If
                    
                    If bIsOOG Then
                        If strContainerNo <> "" Then
                        
                            If InStr(1, Trim(grdCCRTran.TextMatrix(n, enRateCode)), "CBEXP") > 0 Then
                                If N4CurrentCategory = "EXPRT" And bHasUnitOut = False Then
                                    strOOGGranted = ReleaseOOG(strContainerNo)
                                    If strOOGGranted = "0" Then
                                        Call GrantOOGPermission(strContainerNo, intCCRNum)
                                    End If
                                End If
                            Else
                                strOOGGranted = ReleaseOOG(strContainerNo)
'                                If N4CurrentCategory = "IMPRT" And bHasUnitOut = False Then
'                                End If
                            End If
                        End If
                    End If
                    If bIsDG Then
                        If strContainerNo <> "" Then
                        
                            If InStr(1, Trim(grdCCRTran.TextMatrix(n, enRateCode)), "CBEXP") > 0 Then
                                If N4CurrentCategory = "EXPRT" And bHasUnitOut = False Then
                                    strDGGranted = ReleaseDG(strContainerNo)
                                    If strDGGranted = "0" Then
                                        Call GrantDGPermission(strContainerNo, intCCRNum)
                                    End If
                                End If
                            Else
                                strDGGranted = ReleaseDG(strContainerNo)
'                                If N4CurrentCategory = "IMPRT" And bHasUnitOut = False Then
'                                End If
                            End If
                        End If
                    End If
                    
                End If
            End If
            
            'PRNH
            Call lzApplyCCR(gUserID, vCCR, grdCCRTran.TextMatrix(n, enCompCode))
        Next n
        
        'PRNH - OLD
        'Call lzApplyCCR(gUserID, vCCR)
        Call lzInitialize
        Call lzGetUserInfo
        Call OutCCRPC(Trim(vRef))
        'Added by Navis Project Team 11/07/2009
        'Added printing module
'        With clsCCRReprint
'            'On Error GoTo err_Reprint
'            .CCRSupervisor vSupervisor
'            .CCRNumber = CLng("0" & vCCR)
''            .PrintCCR CLng("0" & Trim(vRef))
'            .PreviewCCR CLng("0" & Trim(vRef))
'        End With

        ' call printing module
'        Call lzPrintCCR(vRef, vCCR)
tag_ReInit:
    
    End With
    Set cmd = Nothing
    txtCshAmt.ForeColor = vbWindowBackground
    txtCusName.SetFocus
    
    Exit Sub

err_Save:
    MsgBox "Error in saving this transaction...", vbExclamation
    txtCshAmt.SetFocus
End Sub

'Private Sub OutCCRPC(pRefnum As Long, pSeqnum As Long, pCustomer As String, pAdrAmt As String, pCashAmt As String, _
'                                 pChqAmt As String, pChkno1 As Boolean, pChkno2 As Boolean, pChkno3 As Boolean, pChkno4 As Boolean, _
'                                 pChkno5 As Boolean)
Private Sub OutCCRPC(pRefnum As Long)
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

Dim vslName As String * 10
Dim X As Integer
Dim strRemarks As String * 15
Dim rsCCRDetail As ADODB.Recordset
Dim strRemarkOut As String * 16
Dim remark1 As String
Dim remark3 As String
Dim UserName As String
Dim strValidation As String * 35

Dim sRateCode As String
Dim sRateDescription As String * 29
Dim sRateAmount As Currency
Dim sDays As String
Dim docRefNo As String * 6
Dim sValidUntil As String
Dim sValidUntilText As String
Dim Amount As Currency
Dim TotalAmount As Currency
Dim TotalVatAmount As Currency
Dim TotalTaxAmount As Currency
Dim TotalCheckAmount As Currency
Dim strChqAmt As String
Dim strCshAmt As String
Dim rsCCRPay As ADODB.Recordset
Dim strAdrAmt As String

ctrCnt = 11
On Error Resume Next
    Set rsCCRPay = New ADODB.Recordset
    rsCCRPay.Open "SELECT cusnam, userid From CCRPay WHERE refnum = " & Trim(CStr(pRefnum)), _
            gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
    If rsCCRPay.BOF <> True And rsCCRPay.EOF <> True Then
        strExporter = rsCCRPay.Fields("cusnam")
        UserName = Trim(UCase(rsCCRPay.Fields("userid") & "")) & Space(8)
    End If
    rsCCRPay.Close
    
Set rsCCRDetail = New ADODB.Recordset
rsCCRDetail.Open "SELECT * From CCRdtl WHERE refnum = " & Trim(CStr(pRefnum)) & "" _
        & " order by itmnum", _
        gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
If rsCCRDetail.BOF <> True And rsCCRDetail.EOF <> True Then
    With rsCCRDetail
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

        If .Fields("guarntycde") = "Y" Then
            remark3 = "U/G"
        Else
            remark3 = " "
        End If
        'UserName = Trim(UCase(.Fields("userid") & "")) & Space(8) & Trim(UCase(.Fields("supvsr") & ""))
        Refn = Format(.Fields("refnum"), "000000")
        Seqf = Trim(Format(.Fields("seqnum"), "0000"))
        CCRf = Format(.Fields("ccrnum"), "000000")
        DateTime = Format(.Fields("sysdttm"), "     YYYY-MM-DD hh:nn")
        
'        strRemarks = Mid(.Fields("remark") & "", 1, 15)
        'strRemarkOut = Trim(Mid(remark1, 1, 6)) & Trim(Mid(remark2, 1, 7)) & Trim(Mid(remark3, 1, 3))
        strRemarkOut = Trim(Mid(remark1, 1, 6)) & Trim(Mid(remark3, 1, 3)) & Space(7)
        strEntry = .Fields("entnum")
        strValidation = Trim(Refn) & " " & Trim(Seqf) & " " & Trim(CCRf) & " " & Format(.Fields("sysdttm"), "YY-MM-DD hh:nn")
        vslName = .Fields("vslcde") & ""
        
        Printer.Font = "Courier 12cpi"
        Printer.FontSize = 10

        Printer.Print " "

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
        'Printer.Print Space(2) & .Fields("commod")
        Printer.Print Space(32)

        Do While Not .EOF
            strSize = .Fields("cntsze")
            strCtnnum = .Fields("cntnum")
            sRateDescription = Mid(Trim(.Fields("descr")), 1, 25)
            sDays = ""
            docRefNo = .Fields("docRefNo")
            If .Fields("chargetyp") = "IMST" Then
                If IsNull(.Fields("stordys")) = False And CInt(.Fields("stordys")) <> 0 Then
                    sDays = .Fields("stordys")
                End If
            ElseIf .Fields("chargetyp") = "IMRF" Then
                If IsNull(.Fields("rfrhrs")) = False And CInt(.Fields("rfrhrs")) <> 0 Then
                    sDays = .Fields("rfrhrs")
                End If
            End If
            
        
            If .Fields("chargetyp") <> "IMST" And .Fields("chargetyp") <> "EXST" Then
                sValidUntilText = Space(11)
                sValidUntil = ""
                
            Else
                sValidUntilText = "VALID UNTIL"
                sValidUntil = .Fields("enstodttm") & IIf(.Fields("chargetyp") = "IMRF", .Fields("rfrhrs"), "")
            End If
            
            Amount = CSng(.Fields("amt")) + CSng(.Fields("dgramt")) + CSng(.Fields("ovzamt")) + CSng(.Fields("vatamt")) - CSng(.Fields("wtax"))
            'Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & RevTonnage & Space(2) & strArrastre & Space(29) & strArrastre
            'sharon 05Nov2009 Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & RevTonnage & Space(2) & strArrastre & Space(2) & strWgh & Space(26) & Format(CDbl(strArrastre) + CDbl(strWgh), "###,###,###.#0")
            'printing of container numbers, container size, amount, rate code and rate description
            Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & sDays & Space(2) & sRateDescription & Space(2) & docRefNo & Space(2) & sValidUntilText & Space(2); sValidUntil & Space(6) & Format(CDbl(Amount), "###,###,###.#0")
            TotalAmount = TotalAmount + Amount
            TotalVatAmount = TotalVatAmount + CSng(.Fields("vatamt"))
            TotalTaxAmount = TotalTaxAmount + CSng(.Fields("wtax"))
            strRemarks = Trim(.Fields("remark"))
            ctrCnt = ctrCnt - 1
            .MoveNext
        Loop
        If ctrCnt > 0 Then
            For X = 1 To ctrCnt
                Printer.Print " "
            Next
        End If
    End With

    If TotalVatAmount > 0 Then
        If TotalTaxAmount > 0 Then
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
    Printer.Print Space(6) & Space(15) & Space(17) & Space(17) & Trim(tmpString) & Space(10) & Format(CDbl(TotalAmount), "###,###,###.#0")
    tmpString = NumToText(CCur(TotalAmount))
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

'        If DomesticMode Then
'            Printer.Print "  DOMESTIC" & Space(36) & Word3
'        Else
'            Printer.Print "  FOREIGN " & Space(36) & Word3
'        End If
    Printer.Print Space(46) & Word3

    Printer.Print " "


    Set rsCCRPay = New ADODB.Recordset
    rsCCRPay.Open "SELECT * From CCRPay WHERE refnum = " & Trim(CStr(pRefnum)), _
            gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
    If rsCCRPay.BOF <> True And rsCCRPay.EOF <> True Then
        With rsCCRPay
        'get the Header/footer data
            If IsNull(.Fields("chkamt1")) = False Then
                TotalCheckAmount = TotalCheckAmount + CSng(.Fields("chkamt1"))
            End If
            If IsNull(.Fields("chkamt2")) = False Then
                TotalCheckAmount = TotalCheckAmount + CSng(.Fields("chkamt2"))
            End If
            If IsNull(.Fields("chkamt3")) = False Then
                TotalCheckAmount = TotalCheckAmount + CSng(.Fields("chkamt3"))
            End If
            If IsNull(.Fields("chkamt4")) = False Then
                TotalCheckAmount = TotalCheckAmount + CSng(.Fields("chkamt4"))
            End If
            If IsNull(.Fields("chkamt5")) = False Then
                TotalCheckAmount = TotalCheckAmount + CSng(.Fields("chkamt5"))
            End If
            
            strChqAmt = Format(TotalCheckAmount, "###,###.00")
            strCshAmt = Format(.Fields("cshamt"), "###,###.00")
            
            tmpString = strChqAmt & " CK    " & strCshAmt & " CS"
            Printer.Print Space(44) & tmpString
            
            strAdrAmt = Format(.Fields("adramt"), "###,###.00")
            
            tmpString = strAdrAmt & " AD"
            
            UserName = .Fields("userid")
            
            Printer.Print Space(5) & UserName  ' & Space(26) & tmpString
            
            If IsNull(.Fields("chkamt1")) = False Then
                tmp1 = Trim(.Fields("chkno1"))
            Else
                tmp1 = " "
            End If
            
            If IsNull(.Fields("chkamt2")) = False Then
                If Len(tmp1) > 0 Then
                    tmp2 = ", " & Trim(.Fields("chkno2"))
                Else
                    tmp2 = " " & Trim(.Fields("chkno2"))
                End If
            Else
                tmp2 = " "
            End If
            Printer.Print Space(44) & Trim(tmp1) & tmp2

            If IsNull(.Fields("chkamt3")) = False Then
                tmp1 = Trim(.Fields("chkno3"))
            Else
                tmp1 = " "
            End If
            
            If IsNull(.Fields("chkamt4")) = False Then
                If Len(tmp1) > 0 Then
                    tmp2 = ", " & Trim(.Fields("chkno4"))
                Else
                    tmp2 = " " & Trim(.Fields("chkno4"))
                End If
            Else
                tmp2 = " "
            End If

            Printer.Print Space(44) & Trim(tmp1) & tmp2
            
            If IsNull(.Fields("chkamt5")) = False Then
                tmp1 = Trim(.Fields("chkno5"))
            Else
                tmp1 = " "
            End If

            Printer.Print Space(44) & tmp1
            Printer.Print Space(44) & strValidation
            Printer.Print ""
            Printer.Print ""
            Printer.Print Space(5) & "REF " & Refn & " SEQ " & Seqf & Space(2) & CCRf
        
        End With
        rsCCRPay.Close
        Set rsCCRPay = Nothing
    End If

    
    Printer.FontSize = 10
    Printer.EndDoc

End If

rsCCRDetail.Close
Set rsCCRDetail = Nothing

End Sub

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

'Private Sub txtTruckMake_GotFocus()
'    With txtTruckMake
'        .BackColor = &HFFFFC0
'        .SelStart = 0
'        .SelLength = .MaxLength
'    End With
'End Sub
'
'Private Sub txtTruckMake_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case vbKeyEscape
'            SendKeys ("+{TAB}")
'            KeyAscii = 0
'        Case vbKeyReturn
'            SendKeys ("{TAB}")
'            KeyAscii = 0
'        Case Else
'    End Select
'End Sub
'
'Private Sub txtTruckMake_LostFocus()
'txtTruckMake.BackColor = vbWindowBackground
'End Sub

'Private Sub txtTruckPLT_GotFocus()
'    With txtTruckPLT
'        .BackColor = &HFFFFC0
'        .SelStart = 0
'        .SelLength = .MaxLength
'    End With
'End Sub
'
'Private Sub txtTruckPLT_KeyPress(KeyAscii As Integer)
'Select Case KeyAscii
'        Case vbKeyEscape
'            SendKeys ("+{TAB}")
'            KeyAscii = 0
'        Case vbKeyReturn
'            SendKeys ("{TAB}")
'            KeyAscii = 0
'        Case Else
'    End Select
'End Sub
'
'Private Sub txtTruckPLT_LostFocus()
'txtTruckPLT.BackColor = vbWindowBackground
'End Sub

Public Sub ApplyCCR(CTA As Long, compCode As String)
    'PRNH _Retrieved from old but working version
    Dim cmd As ADODB.Command
    
    ' create command
    Set cmd = New ADODB.Command
    With cmd
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_applyccrspl"
        .CommandType = adCmdStoredProc
    
        ' set parameters then execute
        .Parameters(1).Type = adChar
        .Parameters(1).Value = zCurrentUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adInteger
        .Parameters(2).Value = CTA
        .Parameters(2).Direction = adParamInput
        
        'PRNH
        .Parameters(3).Type = adChar
        .Parameters(3).Value = compCode
        .Parameters(3).Direction = adParamInput

        .Execute

    End With

'PRNH - Commented out due to obsolete functions used
'  'Dim NDate As Date
'  'Dim NxtNo As Long
'  Dim tp As New Recordset
'  'Dim StartCCR As Long
'  'Dim EndCCR As Long
'  Dim PrvCCR As Long
'  Dim dte1 As New Command
'  Dim param1 As New ADODB.Parameter
'  Dim param2 As New ADODB.Parameter
'  Dim param3 As New ADODB.Parameter
'  'Dim tp1 As New Recordset
'
'  Set tp = New ADODB.Recordset
'  tp.Open "SELECT dbo.fn_CYGetInitialCCBR ('CY', '" & UCase(gUserID) & "')", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
'  'tp.Open "SELECT * FROM SPLALLOC WHERE TELLER = '" & UCase(gUserID) & "'", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
'
'  If Not tp.EOF Then
'      'StartCCR = tp.Fields("strccr")
'      'EndCCR = tp.Fields("endccr")
'      PrvCCR = tp.Fields(0)
'  Else
'      'StartCCR = 0
'      'EndCCR = 0
'      PrvCCR = 0
'  End If
'
'  'If (CTA <= EndCCR) And (CTA > PrvCCR) And (CTA >= StartCCR) Then
'  If CTA > PrvCCR Then
'    dte1.ActiveConnection = gcnnBilling
'    dte1.CommandText = "dbo.sp_CYUpdateCCBRAlloc"
'    dte1.CommandType = adCmdStoredProc
'    Set param1 = dte1.CreateParameter("ccbrnum", adBigInt, adParamInput, 18, CTA)
'    dte1.Parameters.Append param1
'    Set param2 = dte1.CreateParameter("grouptype", adVarChar, adParamInput, 10, "CY")
'    dte1.Parameters.Append param2
'    Set param3 = dte1.CreateParameter("teller", adVarChar, adParamInput, 20, Trim(UCase(gUserID)))
'    dte1.Parameters.Append param3
'    dte1.Execute
'
'    Set dte1 = Nothing
''      tp.Fields("prvccr") = CTA
''
''      tp1.Open "SELECT GETDATE()", gcnnBilling, , , adCmdText
''
''      tp.Fields("prvdte") = tp1.Fields(0)
''      tp.Update
'
'      'commented by Navis Project Team 10/29/2009
'      'NTBSPayment1.Ini_CCRNum = lzGetNextCCR(gUserID)
'      'txtCCRNumber.Text = NTBSPayment1.Ini_CCRNum
'
'  Else
'      MsgBox "CCR# " & CTA & " was already used &/or is allocated to other user!", vbInformation + vbOKOnly, "CYS Allocation"
'  End If
'
'  tp.Close
'  Set tp = Nothing
End Sub

Public Function ConnectToNavis() As Boolean '(ByVal pCnnStr As String) As Boolean
Dim errBilling As ADODB.Error
Dim lsErrStr As String
   
    ' Open the database.
    On Error GoTo err_Connect
'    Set gcnnNavis = New ADODB.Connection
 '   gcnnNavis.Open "Provider=sqloledb" & _
  '      ";Data Source=itss-appscript" & _
   '     ";Initial Catalog=apex" & _
    '    ";Integrated Security=SSPI"
        
    Set gcnnNavis = New ADODB.Connection
    gcnnNavis.Open "Provider=sqloledb" & _
        ";Data Source=sbitc-db" & _
        ";Initial Catalog=apex" & _
        ";User ID=tosadmin;Password=tosadmin"
        
'    Set gcnnNavis = New ADODB.Connection
'    gcnnNavis.Open "Provider=sqloledb" & _
'        ";Data Source=sbitc-dev" & _
'        ";Initial Catalog=apex" & _
'        ";User ID=sa_ictsi;password=Ictsi123"
'
    gbNavis = True
    ConnectToNavis = True
    
    'PRNH
    If Not ConnectToNavis Then
        MsgBox "Cannot connect to NAVIS. Please contact IT."
    End If
    Exit Function
    
err_Connect:
    ConnectToNavis = False: gbConnected = False
    For Each errBilling In gcnnBilling.Errors
        With errBilling
            lsErrStr = "Connection Error. " & .Description & vbLf & _
            "Verify Log On then retry.  Contact MIS for assistance."
        End With
        MsgBox lsErrStr, vbCritical
    Next
End Function

'Added Navis Project Team 10/28/2009
Private Sub GetSparcsN4Host()
    Dim rstSparcsN4Host As ADODB.Recordset
    Dim strSparcsN4Host As String
    
    Set rstSparcsN4Host = New ADODB.Recordset
    
    strSparcsN4Host = ""
    strSparcsN4Host = "SELECT * " & _
                       "FROM SparcsN4Host " & _
                       "WHERE status='ACT'"

    rstSparcsN4Host.Open strSparcsN4Host, gcnnBilling, adOpenForwardOnly, adLockReadOnly
    
    If rstSparcsN4Host.BOF Then
        Exit Sub
    End If
    
    With rstSparcsN4Host
        .MoveFirst
        strN4Url = Trim(.Fields("hstnam"))
        strN4Authorization = Trim(.Fields("Authorization"))
        strN4UserName = Trim(.Fields("username"))
        strN4Password = Trim(.Fields("password"))
    End With
    Set rstSparcsN4Host = Nothing
End Sub

Private Sub GetOtherContainerInfo()

    If txtOthContNo.Text <> "" Then
        Call ComputeOOG
    End If
End Sub

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

'PRNH - Evaluating Change
Private Function EvaluateChange() As Currency
    Dim TotalAmount As Currency
    Dim AmtPAy As Currency
    Dim lngCtrX As Long
    TotalAmount = 0
    AmtPAy = 0
    lngCtrX = 0
    
    If IsNumeric(txtCshAmt) Then
        TotalAmount = TotalAmount + CCur(txtCshAmt)
    End If
    For lngCtrX = 0 To 4
        If IsNumeric(txtChkAmt(lngCtrX).Text) Then
            TotalAmount = TotalAmount + CCur(txtChkAmt(lngCtrX).Text)
        End If
    Next
    If IsNumeric(txtADRAmt.Text) Then
        TotalAmount = TotalAmount + CCur(txtADRAmt.Text)
    End If
    AmtPAy = CCur(lblAmtDue.Caption)
    EvaluateChange = TotalAmount - AmtPAy
    If EvaluateChange >= 0 Then
        If bUnderG Then
            cmdSave.Enabled = False
        Else
            cmdSave.Enabled = True
        End If
    Else
        cmdSave.Enabled = False
    End If
End Function
