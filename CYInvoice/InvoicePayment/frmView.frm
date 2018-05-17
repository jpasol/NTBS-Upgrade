VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmView 
   Caption         =   "Payments Made"
   ClientHeight    =   12090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15135
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12090
   ScaleWidth      =   15135
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "View Options"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6000
         Picture         =   "frmView.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton optViewBy 
         Caption         =   "By Invoice"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker DTPDate 
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   360
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   22740993
         CurrentDate     =   36958
      End
      Begin VB.OptionButton optViewBy 
         Caption         =   "By Date of Payments"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton optViewBy 
         Caption         =   "By Customer "
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   3375
      End
      Begin VB.OptionButton optViewBy 
         Caption         =   "By Official Reciepts"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3375
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msgPayments 
      Height          =   8655
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   15266
      _Version        =   393216
      FixedRows       =   0
      ForeColorFixed  =   16711680
      BackColorSel    =   16711680
      ForeColorSel    =   65535
      GridColorFixed  =   -2147483636
      WordWrap        =   -1  'True
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   4
      _Band(0).ColHeader=   1
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rst As New ADODB.Recordset
Dim cnn As New ADODB.Connection
Dim sParent As String
Dim sChild As String
Dim sRelationship As String

Private Sub cmdPreview_Click()
     rst.Open "Shape (SHAPE  {" & sParent & " } as cmdParent " _
               & "APPEND ({ " & sChild & " }  as cmdChild " _
               & "RELATE " & sRelationship & " )) AS rsRelationship ", _
              cnn, adOpenDynamic, adLockReadOnly, adCmdText
    Set msgPayments.DataSource = rst
    rst.Close
End Sub

Private Sub Form_Load()
 
 Set cnn = New ADODB.Connection
 Set rst = New ADODB.Recordset
 
 cnn.Provider = "MSDataShape"
 cnn = "data Provider=sqloledb;Data Source=NTBS;Initial Catalog=BILLING;Integrated Security=SSPI"
 cnn.Open cnn
 
 DTPDate.Value = Date
 sParent = ""
 sChild = ""
 sRelationship = ""
 
 rst.StayInSync = False
 optViewBy(0).Value = 1
  With msgPayments
     .Parent
     .Width = Screen.Width - 500
     .Height = Screen.Height - 3000
     .Left = 250
     .Top = 2500
  End With
  Call optViewBy_Click(1)
  Call cmdPreview_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If MsgBox("Close View Payment Window?", vbOKCancel + vbQuestion, "Close") = vbOK Then
        Cancel = 0
        Unload Me
  Else
     Cancel = 1
  End If
End Sub

Private Sub optViewBy_Click(Index As Integer)
  Select Case Index
      Case 0  ' By date of payments
           sParent = " Select cuscde, ornum,ortype,totalamt from invpayhdr order by ornum "
           sChild = "   select distinct ornum, invpaydtl.invnum, invict.invamt,invict.invvat,invict.invtax,(invict.invamt + invict.invvat - invict.invtax) as TotalInvoiceAmount, payamt from invict, invpaydtl" _
                        & " Where invpaydtl.invnum = invict.invnum "
            sRelationship = " ornum to ornum"

      Case 1 ' By OR
            sParent = " Select cuscde, ornum,ortype,totalamt from invpayhdr order by ornum "
            sChild = "   select distinct ornum, invpaydtl.invnum, invict.invamt,invict.invvat,invict.invtax,(invict.invamt + invict.invvat - invict.invtax) as TotalInvoiceAmount, payamt from invict, invpaydtl" _
                        & " Where invpaydtl.invnum = invict.invnum "
            sRelationship = " ornum to ornum"

      Case 2 ' By Customer
            sParent = "Select CUSCDE from CUSTOMER "
            sChild = " Select * from vueorlist "
            sRelationship = " cuscde to cuscde "

      Case 3 ' By Invoice
            sParent = " Select invnum,invamt as Ampount, invtax as WTAX, (invamt+invvat-invtax) as Total from viewfinalorlist "
            sChild = " select invpaydtl.ornum, invpaydtl.invnum, invpaydtl.payamt as PayAmount, invpayhdr.ortype, invpaydtl.paydate from invpaydtl, invpayhdr " _
             & "where invpaydtl.ornum = invpayhdr.ornum"
            sRelationship = "invnum to invnum"
  End Select
End Sub
