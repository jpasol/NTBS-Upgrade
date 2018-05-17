VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListORpayment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Payments Made"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12540
   Icon            =   "frmListORpayment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Sort"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   12135
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
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   3375
      End
      Begin VB.OptionButton optViewBy 
         Caption         =   "O.R. #"
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
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   3375
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
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
         Left            =   9360
         Picture         =   "frmListORpayment.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
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
         Left            =   6960
         Picture         =   "frmListORpayment.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optViewBy 
         Caption         =   "By Date and Time  of Payment"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   4935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grd_Payments 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   13573
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   65535
      BackColorSel    =   8454016
      ForeColorSel    =   16711680
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
   Begin MSComCtl2.DTPicker DTPdate 
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   1560
      Width           =   4455
      _ExtentX        =   7858
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
      Format          =   22806528
      CurrentDate     =   36966
   End
   Begin VB.Label lblContest 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "View Date "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   7665
   End
End
Attribute VB_Name = "frmListORpayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim rs As ADODB.Recordset
Dim Dateformat As String


Private Sub cmdClose_Click()
 rs.Close
 Set rs = Nothing
 Unload Me
End Sub

Private Sub cmdPreview_Click()
  Call List_PaidBills
End Sub

Private Sub DTPdate_Change()
   Dateformat = " WHERE Year(ordate)=" & Year(DTPdate.Value) & " AND Month(ordate)=" & Month(DTPdate.Value) & " AND Day(ordate)=" & Day(DTPdate.Value)
   sSql = "select ORNUM, totalamt,availamt, ortype, cuscde,ordate from invpayhdr" & Dateformat
   If rs.State = adStateOpen Then
       rs.Close
   End If
   rs.CursorLocation = adUseClient
   rs.Open sSql, gcnnBilling, adOpenStatic, , adCmdText
   Call cmdPreview_Click
   optViewBy(0).Value = True
End Sub

Private Sub DTPdate_Click()
   Dateformat = " WHERE Year(ordate)=" & Year(DTPdate.Value) & " AND Month(ordate)=" & Month(DTPdate.Value) & " AND Day(ordate)=" & Day(DTPdate.Value)
   sSql = "select ORNUM, totalamt,availamt, ortype, cuscde,ordate from invpayhdr" & Dateformat
   If rs.State = adStateOpen Then
       rs.Close
   End If
  rs.CursorLocation = adUseClient
  rs.Open sSql, gcnnBilling, adOpenStatic, , adCmdText
  Call cmdPreview_Click
  optViewBy(0).Value = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Call Initialize
   Dateformat = " WHERE Year(ordate)=" & Year(DTPdate.Value) & " AND Month(ordate)=" & Month(DTPdate.Value) & " AND Day(ordate)=" & Day(DTPdate.Value)
   Set rs = New ADODB.Recordset
   sSql = "select ORNUM, totalamt,availamt, ortype, cuscde,ordate from invpayhdr" & Dateformat
   rs.CursorLocation = adUseClient
   rs.Open sSql, gcnnBilling, adOpenStatic, , adCmdText
   Call cmdPreview_Click
End Sub

Private Sub optViewBy_Click(Index As Integer)
 Select Case Index
   Case 0 ' By Date
        rs.Sort = "ordate DESC"
   Case 1 ' By OR#
     rs.Sort = "ornum "
   Case 2 ' By Customer
     rs.Sort = "cuscde"
 End Select
End Sub

Private Sub Initialize()
  With grd_Payments
       .Cols = 6
       .Rows = 2
       .Height = 7696
       .Width = 12255
 End With
 frmListORpayment.Width = 12630
 frmListORpayment.Height = 10290
 optViewBy(0).Value = True
 DTPdate.Value = Date
End Sub
Private Sub Initialized_Grid()
  Dim colhdgs(6) As String
  Dim col As Integer
  
  colhdgs(0) = " O.R. # "
  colhdgs(1) = " Total Amount "
  colhdgs(2) = "Avail.Credit Amt "
  colhdgs(3) = " OR Type "
  colhdgs(4) = " Customer Code "
  colhdgs(5) = " OR Date/Time "
  With grd_Payments
        .Clear
        .Rows = 2
        .row = 0
  
    For col = 0 To 5
        .col = col: .Text = colhdgs(col): .CellAlignment = 4: .CellFontBold = True: .CellForeColor = &HFFFF&: .CellBackColor = &HC00000
    Next col
  
        .ColWidth(0) = 1500
        .ColWidth(1) = 2000
        .ColWidth(2) = 2500
        .ColWidth(3) = 1300
        .ColWidth(4) = 2000
        .ColWidth(5) = 2700
  End With
End Sub

Public Sub List_PaidBills()
 Dim rowcount, col As Integer
 Dim blnToggle As Boolean
 Dim RowColor  As Long
 
  MousePointer = 11
  grd_Payments.Rows = 2
  grd_Payments.Clear
  Call Initialized_Grid
  
  If rs.RecordCount < 1 Then
          MsgBox "Empty list of Payments for " & DTPdate.Value, vbOKOnly + vbInformation, "No Paid Bills"
           
   ElseIf rs.RecordCount > 0 Then   'Populate the grid
      rowcount = 1
      With grd_Payments
        .Visible = False
        Do While Not rs.EOF
            If rowcount > 1 Then
               .AddItem ""  ' add another row
            End If
            
            RowColor = IIf(blnToggle = False, &H80000018, vbWhite)
            .RowHeight(rowcount) = 250
            .row = rowcount
            .col = 0: .Text = rs!ornum: .CellAlignment = 4: .CellBackColor = RowColor
            .col = 1: .Text = Format(rs!TotalAmt, "###,###,###.#0"): .CellBackColor = RowColor
            .col = 2: .Text = Format(rs!AvailAMT, "###,###,###.#0"): .CellBackColor = RowColor
            .col = 3: .Text = UCase(rs!ortype): .CellBackColor = RowColor: .CellAlignment = 4 'Center
            .col = 4: .Text = Trim(rs!cuscde): .CellBackColor = RowColor: .CellAlignment = 4
            '.col = 5: .Text = Format(rs!ORDate, "YYYY/MM/DD"): .CellBackColor = RowColor
            .col = 5: .Text = rs!ORDate: .CellBackColor = RowColor

             rowcount = rowcount + 1
            rs.MoveNext
            blnToggle = Not blnToggle
           Loop
        .Visible = True
        'set row selection
        .Enabled = True
        .col = 0
        .row = 1
        .ColSel = 5
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
      End With
    End If
Me.MousePointer = 0
End Sub


