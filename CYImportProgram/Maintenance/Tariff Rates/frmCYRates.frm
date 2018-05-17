VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCYRates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CY Billing Rates File Maintenance"
   ClientHeight    =   10905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "IBM3270 - 1254"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10905
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame fraSearch 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   2760
      Width           =   15375
      Begin VB.TextBox txtFndCde 
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   1385
      End
      Begin VB.TextBox txtFndSze 
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   1610
         TabIndex        =   30
         Top             =   360
         Width           =   505
      End
      Begin VB.TextBox txtFndDsc 
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   2100
         TabIndex        =   32
         Top             =   360
         Width           =   5550
      End
      Begin VB.TextBox txtFndAmt 
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   7630
         TabIndex        =   34
         Top             =   360
         Width           =   1640
      End
      Begin VB.TextBox txtFndUsr 
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   9260
         TabIndex        =   36
         Top             =   360
         Width           =   1820
      End
      Begin VB.TextBox txtFndDte 
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   11050
         TabIndex        =   38
         Top             =   360
         Width           =   3870
      End
      Begin VB.Label lblFndCde 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CODE"
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   0
         Width           =   1385
      End
      Begin VB.Label lblFndSze 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sz"
         Height          =   375
         Left            =   1610
         TabIndex        =   37
         Top             =   0
         Width           =   505
      End
      Begin VB.Label lblFndDsc 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DESCRIPTION"
         Height          =   375
         Left            =   2100
         TabIndex        =   35
         Top             =   0
         Width           =   5550
      End
      Begin VB.Label lblFndAmt 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AMOUNT"
         Height          =   375
         Left            =   7630
         TabIndex        =   33
         Top             =   0
         Width           =   1640
      End
      Begin VB.Label lblUsr 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "USERID"
         Height          =   375
         Left            =   9260
         TabIndex        =   31
         Top             =   0
         Width           =   1820
      End
      Begin VB.Label lblDte 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DATE"
         Height          =   375
         Left            =   11050
         TabIndex        =   29
         Top             =   0
         Width           =   3870
      End
   End
   Begin VB.Frame fraFields 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   14685
      Begin VB.ComboBox cmbBilTyp 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   10200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox txtRteCde 
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtCntSze 
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   2
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox txtRteDsc 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         MaxLength       =   55
         TabIndex        =   4
         Top             =   1680
         Width           =   7575
      End
      Begin VB.TextBox txtRteTyp 
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   3000
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtUms 
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   10200
         MaxLength       =   10
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtRteAmt 
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   3000
         MaxLength       =   12
         TabIndex        =   5
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtActCde 
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   10200
         MaxLength       =   5
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rate Code"
         Height          =   300
         Left            =   1320
         TabIndex        =   26
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Container Size"
         Height          =   300
         Left            =   480
         TabIndex        =   25
         Top             =   720
         Width           =   2310
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Rate Description"
         Height          =   300
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   2640
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Rate Type"
         Height          =   300
         Left            =   1200
         TabIndex        =   23
         Top             =   1200
         Width           =   1485
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Unit of Measure"
         Height          =   300
         Left            =   7560
         TabIndex        =   22
         Top             =   240
         Width           =   2475
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Rate Amount"
         Height          =   300
         Left            =   840
         TabIndex        =   21
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Bill Type"
         Height          =   300
         Left            =   8520
         TabIndex        =   20
         Top             =   1200
         Width           =   1485
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Account Code"
         Height          =   300
         Left            =   8040
         TabIndex        =   19
         Top             =   720
         Width           =   1980
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "User ID"
         Height          =   300
         Left            =   8880
         TabIndex        =   18
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label lblUserid 
         AutoSize        =   -1  'True
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   10680
         TabIndex        =   17
         Top             =   2160
         Width           =   165
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdRates 
      Height          =   6255
      Left            =   360
      TabIndex        =   9
      Top             =   3480
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   11033
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   250
      BackColor       =   -2147483624
      BackColorSel    =   -2147483646
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLines       =   2
      ScrollBars      =   2
      SelectionMode   =   1
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "F3 = E&xit"
      Height          =   615
      Left            =   12960
      TabIndex        =   16
      Top             =   10080
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   615
      Left            =   10920
      TabIndex        =   15
      Top             =   10080
      Width           =   2055
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   615
      Left            =   8760
      TabIndex        =   14
      Top             =   10080
      Width           =   2175
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   615
      Left            =   6720
      TabIndex        =   13
      Top             =   10080
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   4560
      TabIndex        =   12
      Top             =   10080
      Width           =   2170
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   615
      Left            =   2400
      TabIndex        =   11
      Top             =   10080
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   10080
      Width           =   2055
   End
   Begin VB.Label lblNumRow 
      AutoSize        =   -1  'True
      Caption         =   "0 rows"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   40
      Top             =   9720
      Width           =   465
   End
End
Attribute VB_Name = "frmCYRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SortBy As Integer '0-Code, 1-Size, 2-Description, 3-Amount, 4-Userid, 5-Date
Dim blnEdit As Boolean
Dim blnFill As Boolean

Private Sub cmbBilTyp_GotFocus()
    SendKeys "%{DOWN}"
End Sub

Private Sub Form_Load()
    blnFill = False
    blnEdit = False
    fraFields.Enabled = False
    SortBy = 0
    With cmbBilTyp
        .AddItem "AN|Anchorage", 0
        .AddItem "CB|Cargo Billing (Arrastre)", 1
        .AddItem "MC|Miscellaneous Charges", 2
        .AddItem "SS|Stripping/Stuffing", 3
        .AddItem "ST|Storage", 4
        .AddItem "VB|Vessel Billing", 5
        .AddItem "VC|Cranage", 6
    End With
    cmbBilTyp.ListIndex = 0
    Call SetHeading
    Call FillCells
    cmdAdd.Enabled = True
    cmdSave.Enabled = False
    cmdEdit.Enabled = True
    cmdCancel.Enabled = False
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
    cmdExit.Enabled = True
End Sub

Private Sub SetHeading()
  With grdRates
    .ColWidth(0) = 1350: .ColWidth(1) = 500: .ColWidth(2) = 0
    .ColWidth(3) = 5530: .ColWidth(4) = 1600: .ColWidth(5) = 0
    .ColWidth(6) = 0: .ColWidth(7) = 0: .ColWidth(8) = 1800
    .ColWidth(9) = 3650
  End With
End Sub

Private Sub FillCells()
    Dim rstAdo As New ADODB.Recordset
    Dim strSQL As String
 
    Screen.MousePointer = vbHourglass
    blnFill = False
    Select Case SortBy
      Case 0
        strSQL = "Select * from CYRATE order by cyr_rtecde"
      Case 1
        strSQL = "Select * from CYRATE order by cyr_cntsze"
      Case 2
        strSQL = "Select * from CYRATE order by cyr_rtedsc"
      Case 3
        strSQL = "Select * from CYRATE order by cyr_rteamt"
      Case 4
        strSQL = "Select * from CYRATE order by cyr_userid"
      Case 5
        strSQL = "Select * from CYRATE order by cyr_sysdte"
    End Select
    
    rstAdo.Open strSQL, gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
    
    With grdRates
      .Clear
      .Visible = False
      .Rows = 1
      .Row = 0
      Do While Not rstAdo.EOF
          .Col = 0: .Text = rstAdo!cyr_rtecde
          .Col = 1: .CellAlignment = 1: .Text = "" & rstAdo!cyr_cntsze
          .Col = 2: .Text = "" & rstAdo!cyr_rtetyp
          .Col = 3: .CellAlignment = 1: .Text = "" & rstAdo!cyr_rtedsc
          .Col = 4: .Text = "" & Format(rstAdo!cyr_rteamt, "##,###.#0")
          .Col = 5: .Text = "" & rstAdo!cyr_uomcde
          .Col = 6: .Text = "" & rstAdo!cyr_acccde
          .Col = 7: .Text = "" & GetBilTypIndex(rstAdo!cyr_biltyp)
          .Col = 8: .Text = rstAdo!cyr_userid
          .Col = 9: .Text = Format(rstAdo!cyr_sysdte, "yyyy/mm/dd hh:mm:ssAM/PM")
          rstAdo.MoveNext
          If Not rstAdo.EOF Then
            .Rows = .Rows + 1
            .Row = .Row + 1
          End If
      Loop
      blnFill = True
      .Visible = True
      .Refresh
      lblNumRow = .Rows & " row(s)"
      'set highlight to the first cell
      .Col = 0: .Row = 0: .ColSel = 9
    End With
    rstAdo.Close
    Screen.MousePointer = vbDefault
    
End Sub

Private Function GetBilTypIndex(pBilTyp As String) As Integer
    Dim intIdx As Integer
    Select Case pBilTyp
        Case "AN"
            intIdx = 0
        Case "CB"
            intIdx = 1
        Case "MC"
            intIdx = 2
        Case "SS"
            intIdx = 3
        Case "ST"
            intIdx = 4
        Case "VB"
            intIdx = 5
        Case "VC"
            intIdx = 6
    End Select
    GetBilTypIndex = intIdx
End Function

Private Sub grdRates_GotFocus()
    cmdSave.Enabled = False
    txtFndCde = ""
    txtFndSze = ""
    txtFndDsc = ""
    txtFndAmt = ""
    txtFndUsr = ""
    txtFndDte = ""
End Sub

Private Sub grdRates_RowColChange()
    grdRates.Enabled = True
    fraFields.Enabled = False
    fraSearch.Enabled = True
    cmdAdd.Enabled = True
    cmdEdit.Enabled = True
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
    cmdExit.Enabled = True
    If blnFill = True Then
        txtRteCde = grdRates.TextMatrix(grdRates.Row, 0)
        txtCntSze = Trim(grdRates.TextMatrix(grdRates.Row, 1))
        txtRteTyp = Trim(grdRates.TextMatrix(grdRates.Row, 2))
        txtRteDsc = Trim(grdRates.TextMatrix(grdRates.Row, 3))
        txtRteAmt = Trim(grdRates.TextMatrix(grdRates.Row, 4))
        txtUms = grdRates.TextMatrix(grdRates.Row, 5)
        txtActCde = grdRates.TextMatrix(grdRates.Row, 6)
'        txtBilTyp = grdRates.TextMatrix(grdRates.Row, 7)
        cmbBilTyp.ListIndex = grdRates.TextMatrix(grdRates.Row, 7)
        lblUserid = grdRates.TextMatrix(grdRates.Row, 8)
    End If
End Sub

Private Sub grdRates_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then Call cmdExit_Click
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdAdd_Click()
    ' Initialize text boxes for entry
    txtRteCde = ""
    txtCntSze = ""
    txtRteTyp = ""
    txtRteDsc = ""
    txtRteAmt = ""
    txtUms = ""
    txtActCde = ""
'    txtBilTyp = ""
    cmbBilTyp.ListIndex = 0
    lblUserid = UCase(zCurrentUser)
    fraSearch.Enabled = False
    grdRates.Enabled = False
    fraFields.Enabled = True
    cmdAdd.Enabled = False
    cmdCancel.Enabled = True
    cmdEdit.Enabled = False
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    cmdExit.Enabled = False
    cmdSave.Enabled = False
    txtRteCde.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim rstAdo As New ADODB.Recordset
    
'    If Len(Trim(txtBilTyp)) = 0 Then
'        MsgBox "Please specify a valid bill type.", vbInformation, "Save Message"
'        txtBilTyp.SetFocus
'        Exit Sub
'    End If
    If txtRteAmt = "" Then txtRteAmt = 0
    ' Saving only for edit
    If blnEdit = True Then
        Call SaveEdit
        blnEdit = False
        Exit Sub
    End If
    ' Saving only for add
    If lFindRec(txtRteCde, txtCntSze) Then
        MsgBox "A record with similar rate code and container size already exists.", vbExclamation, "Error"
        txtRteCde.SetFocus
        Exit Sub
    End If
    With rstAdo
        .Open "CYRATE", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
        .AddNew
        .Fields("cyr_rtecde") = txtRteCde
        .Fields("cyr_cntsze") = txtCntSze
        .Fields("cyr_rtetyp") = txtRteTyp
        .Fields("cyr_rtedsc") = txtRteDsc
        .Fields("cyr_rteamt") = txtRteAmt
        .Fields("cyr_uomcde") = txtUms
        .Fields("cyr_acccde") = txtActCde
        .Fields("cyr_biltyp") = Left(Trim(cmbBilTyp.Text), 2)  ' txtBilTyp
        .Fields("cyr_updcde") = ""
        .Fields("cyr_sysdte") = gzGetSysDate
        .Fields("cyr_userid") = UCase(zCurrentUser)
        .Update
        .Close
    End With
    Call FillCells
    grdRates.SetFocus
End Sub

Private Sub SaveEdit()
    Dim rstAdo As New ADODB.Recordset
    Dim intResponse As Integer
    
        intResponse = MsgBox("Save changes?", vbYesNo, "Save")
        If intResponse = vbNo Then
            txtRteCde.Enabled = True
            txtCntSze.Enabled = True
            Call grdRates_RowColChange
            grdRates.SetFocus
            Exit Sub
        End If
        
        With rstAdo
            .Open "Select * from CYRate where cyr_rtecde = '" & _
                Trim(txtRteCde) & "' and cyr_cntsze = '" & Trim(txtCntSze) & _
                "'", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
            .Fields("cyr_rtecde") = txtRteCde
            .Fields("cyr_cntsze") = txtCntSze
            .Fields("cyr_rtetyp") = txtRteTyp
            .Fields("cyr_rtedsc") = txtRteDsc
            .Fields("cyr_rteamt") = txtRteAmt
            .Fields("cyr_uomcde") = txtUms
            .Fields("cyr_acccde") = txtActCde
            .Fields("cyr_biltyp") = Left(Trim(cmbBilTyp.Text), 2)  ' txtBilTyp
            .Fields("cyr_updcde") = ""
            .Fields("cyr_sysdte") = gzGetSysDate
            .Fields("cyr_userid") = UCase(zCurrentUser)
            .Update
            .Close
        End With
        txtRteCde.Enabled = True
        txtCntSze.Enabled = True
        Call FillCells
        grdRates.SetFocus
End Sub

Private Sub cmdEdit_Click()
    grdRates.Enabled = False
    fraSearch.Enabled = False
    fraFields.Enabled = True
    txtRteCde.Enabled = False
    txtCntSze.Enabled = False
    lblUserid.Caption = UCase(zCurrentUser)
    txtRteTyp.SetFocus
    cmdEdit.Enabled = False
    cmdAdd.Enabled = False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    cmdPrint.Enabled = False
    cmdExit.Enabled = False
    blnEdit = True
End Sub

Private Sub cmdDelete_Click()
    Dim rstAdo As New ADODB.Recordset
    Dim intResponse As Integer
    
    cmdAdd.Enabled = False
    cmdEdit.Enabled = False
    cmdPrint.Enabled = False
    cmdExit.Enabled = False
    intResponse = MsgBox("Delete this record?", vbYesNo + vbDefaultButton2, "Delete")
    If intResponse = vbNo Then
        Call grdRates_RowColChange
        grdRates.SetFocus
        Exit Sub
    End If
    With rstAdo
      .Open "Select * from CYRate where cyr_rtecde = '" & _
               Trim(txtRteCde) & "' and cyr_cntsze = '" & Trim(txtCntSze) & _
               "'", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
      .Delete
      .Close
    End With
    Call FillCells
    grdRates.SetFocus
End Sub

Private Sub cmdCancel_Click()
    txtRteCde.Enabled = True
    txtCntSze.Enabled = True
    blnEdit = False
    Call grdRates_RowColChange
    grdRates.SetFocus
    cmdSave.Enabled = False
End Sub

Private Sub cmdPrint_Click()
    frmPrintPrev.Show vbModal
End Sub

Private Sub lblDte_Click()
    If SortBy = 5 Then Exit Sub
    If MsgBox("Sort records by date?", vbYesNo + vbQuestion, "Sort") = vbNo Then Exit Sub
    SortBy = 5
    FillCells
End Sub

Private Sub lblFndAmt_Click()
    If SortBy = 3 Then Exit Sub
    If MsgBox("Sort records by rate amount?", vbYesNo + vbQuestion, "Sort") = vbNo Then Exit Sub
    SortBy = 3
    FillCells
End Sub

Private Sub lblFndCde_Click()
    If SortBy = 0 Then Exit Sub
    If MsgBox("Sort records by rate code?", vbYesNo + vbQuestion, "Sort") = vbNo Then Exit Sub
    SortBy = 0
    FillCells
End Sub

Private Sub lblFndDsc_Click()
    If SortBy = 2 Then Exit Sub
    If MsgBox("Sort records by rate description?", vbYesNo + vbQuestion, "Sort") = vbNo Then Exit Sub
    SortBy = 2
    FillCells
End Sub

Private Sub lblFndSze_Click()
    If SortBy = 1 Then Exit Sub
    If MsgBox("Sort records by rate size?", vbYesNo + vbQuestion, "Sort") = vbNo Then Exit Sub
    SortBy = 1
    FillCells
End Sub

Private Sub lblUsr_Click()
    If SortBy = 4 Then Exit Sub
    If MsgBox("Sort records by user ID?", vbYesNo + vbQuestion, "Sort") = vbNo Then Exit Sub
    SortBy = 4
    FillCells
End Sub

Private Sub txtActCde_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtUms, cmbBilTyp)
End Sub

'Private Sub txtBilTyp_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call FieldAdvance(KeyCode, txtActCde, txtBilTyp)
'End Sub
'
'Private Sub txtBilTyp_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

Private Sub txtCntSze_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtRteCde, txtRteTyp)
End Sub

Private Sub txtFndAmt_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) >= " " And Chr(KeyAscii) <= "~" Then
        SmartType txtFndAmt, KeyAscii, 4
        KeyAscii = 0
    End If
End Sub

Private Sub txtFndCde_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) >= " " And Chr(KeyAscii) <= "~" Then
        SmartType txtFndCde, KeyAscii, 0
        KeyAscii = 0
    End If
End Sub

Private Sub txtFndDsc_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) >= " " And Chr(KeyAscii) <= "~" Then
        SmartType txtFndDsc, KeyAscii, 3
        KeyAscii = 0
    End If
End Sub

Private Sub txtFndDte_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) >= " " And Chr(KeyAscii) <= "~" Then
        SmartType txtFndDte, KeyAscii, 9
        KeyAscii = 0
    End If
End Sub

Private Sub txtFndSze_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) >= " " And Chr(KeyAscii) <= "~" Then
        SmartType txtFndSze, KeyAscii, 1
        KeyAscii = 0
    End If
End Sub

Private Sub txtFndUsr_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) >= " " And Chr(KeyAscii) <= "~" Then
        SmartType txtFndUsr, KeyAscii, 8
        KeyAscii = 0
    End If
End Sub

Private Sub txtRteAmt_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtRteDsc, txtUms)
End Sub

Private Sub txtRteAmt_LostFocus()
    If txtRteAmt = "" Or Not IsNumeric(txtRteAmt) Then txtRteAmt = "0"
    txtRteAmt = Format(txtRteAmt, "##,###.#0")
End Sub

Private Sub txtRteCde_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtRteCde, txtCntSze)
End Sub

Private Sub txtRteCde_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRteCde_LostFocus()
    If (txtRteCde <> "") And cmdSave.Enabled = False Then cmdSave.Enabled = True
    If (txtRteCde = "") And cmdSave.Enabled = True Then cmdSave.Enabled = False
End Sub

Private Sub txtRteDsc_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtRteTyp, txtRteAmt)
End Sub

Private Sub txtRteDsc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRteTyp_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtCntSze, txtRteDsc)
End Sub

Private Sub txtRteTyp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUms_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtRteAmt, txtActCde)
End Sub

Private Sub txtUms_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub FieldAdvance(pKeycode As Integer, pPreviousCtl As Control, pNextCtl As Control)
    Select Case pKeycode
        Case vbKeyUp
            If pPreviousCtl.Enabled = True Then
                pPreviousCtl.SetFocus
            End If
        Case vbKeyDown
            pNextCtl.SetFocus
        Case vbKeyReturn
            pNextCtl.SetFocus
        Case vbKeyEscape
            Call cmdCancel_Click
    End Select
End Sub

Private Function lFindRec(pstrRate As String, pstrSize As String) As Boolean
Dim rstRate As New ADODB.Recordset
    With rstRate
        .Open "Select * from CYRate where cyr_rtecde = '" & _
                     Trim(pstrRate) & "' and cyr_cntsze = '" & Trim(pstrSize) & _
                     "'", gcnnBilling, , , adCmdText
        lFindRec = Not .EOF
        .Close
    End With
End Function

Public Sub SmartType(c As Object, n As Integer, pColNum As Integer)
    Dim i As Integer
    Dim l As Integer
    Dim s As Integer
    Dim t As String

    s = c.SelStart
    l = c.SelLength
    t = c.Text

    t = Left(t, s) & Chr(n) & Right(t, Len(t) - s)
    s = s + 1
    i = 0

    Do While (i < grdRates.Rows - 1) _
        And StrComp(Left(t, s), Left(grdRates.TextMatrix(i, pColNum), s), vbTextCompare) <> 0
        i = i + 1
    Loop
    If UCase(Left(t, s)) = UCase(Left(grdRates.TextMatrix(i, pColNum), s)) Then
        t = Trim(grdRates.TextMatrix(i, pColNum))
        l = Len(t) - s
    Else
        t = Left(t, s)
        l = 0
    End If

    c.Text = t
    c.SelStart = s
    c.SelLength = l
    grdRates.TopRow = i
    grdRates.Row = i
    grdRates.Col = 0
    grdRates.ColSel = 9
End Sub
