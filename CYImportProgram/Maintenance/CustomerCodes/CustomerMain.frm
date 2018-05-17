VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmCustomerMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Master File Maintenance"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   945
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CustomerMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptCustList 
      Left            =   14760
      Top             =   9600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
   End
   Begin MSFlexGridLib.MSFlexGrid grdCustomers 
      Height          =   5655
      Left            =   360
      TabIndex        =   19
      Top             =   4320
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   9975
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
   End
   Begin VB.Frame fraDetails 
      Enabled         =   0   'False
      Height          =   3615
      Left            =   360
      TabIndex        =   8
      Top             =   0
      Width           =   14625
      Begin VB.TextBox txtCustCde 
         BackColor       =   &H80000018&
         Height          =   345
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtCustNam 
         BackColor       =   &H80000018&
         Height          =   345
         Left            =   2040
         MaxLength       =   40
         TabIndex        =   1
         Top             =   840
         Width           =   6735
      End
      Begin VB.TextBox txtAgent 
         BackColor       =   &H80000018&
         Height          =   345
         Left            =   2040
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1320
         Width           =   6735
      End
      Begin VB.TextBox txtAdd1 
         BackColor       =   &H80000018&
         Height          =   345
         Left            =   2040
         TabIndex        =   3
         Top             =   1800
         Width           =   6735
      End
      Begin VB.TextBox txtAdd2 
         BackColor       =   &H80000018&
         Height          =   345
         Left            =   2040
         TabIndex        =   4
         Top             =   2160
         Width           =   6735
      End
      Begin VB.TextBox txtAdd3 
         BackColor       =   &H80000018&
         Height          =   345
         Left            =   2040
         TabIndex        =   5
         Top             =   2520
         Width           =   6735
      End
      Begin VB.TextBox txtTelFax 
         BackColor       =   &H80000018&
         Height          =   345
         Left            =   2040
         TabIndex        =   6
         Top             =   3000
         Width           =   4575
      End
      Begin VB.TextBox txtCustTyp 
         BackColor       =   &H80000018&
         Height          =   345
         Left            =   10680
         MaxLength       =   3
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblDateTime 
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
         Height          =   375
         Left            =   10680
         TabIndex        =   18
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label lblUserid 
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   10680
         TabIndex        =   17
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Code     "
         Height          =   375
         Left            =   480
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Name     "
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Address  "
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Type      "
         Height          =   375
         Left            =   9120
         TabIndex        =   13
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Agent    "
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Tel/Fax  "
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "UserID    "
         Height          =   375
         Left            =   9120
         TabIndex        =   10
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Date/Time "
         Height          =   375
         Left            =   9120
         TabIndex        =   9
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Frame fraSearch 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   25
      Top             =   3480
      Width           =   15255
      Begin VB.TextBox txtFndTyp 
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   13750
         MaxLength       =   3
         TabIndex        =   33
         Top             =   480
         Width           =   1240
      End
      Begin VB.TextBox txtFndAgt 
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   7705
         MaxLength       =   40
         TabIndex        =   32
         Top             =   480
         Width           =   6060
      End
      Begin VB.TextBox txtFndNam 
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   1695
         MaxLength       =   40
         TabIndex        =   31
         Top             =   480
         Width           =   6025
      End
      Begin VB.TextBox txtFndCde 
         ForeColor       =   &H00C00000&
         Height          =   420
         Left            =   380
         MaxLength       =   6
         TabIndex        =   30
         Top             =   480
         Width           =   1330
      End
      Begin VB.Label lblCode 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CODE"
         Height          =   375
         Left            =   380
         TabIndex        =   29
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblCusNam 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NAME"
         Height          =   375
         Left            =   1695
         TabIndex        =   28
         Top             =   120
         Width           =   6025
      End
      Begin VB.Label lblAgent 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AGENT"
         Height          =   375
         Left            =   7705
         TabIndex        =   27
         Top             =   120
         Width           =   6060
      End
      Begin VB.Label lblType 
         Alignment       =   2  'Center
         BackColor       =   &H80000011&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TYPE"
         Height          =   375
         Left            =   13750
         TabIndex        =   26
         Top             =   120
         Width           =   1240
      End
   End
   Begin VB.Label lblNumRow 
      AutoSize        =   -1  'True
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
      Left            =   480
      TabIndex        =   34
      Top             =   9960
      Width           =   45
   End
   Begin VB.Label lblF9 
      AutoSize        =   -1  'True
      Caption         =   "F9 = Print"
      Height          =   300
      Left            =   13320
      TabIndex        =   24
      Top             =   10440
      Width           =   1650
   End
   Begin VB.Label lblF8 
      AutoSize        =   -1  'True
      Caption         =   "F8 = Edit"
      Height          =   300
      Left            =   10080
      TabIndex        =   23
      Top             =   10440
      Width           =   1485
   End
   Begin VB.Label lblF7 
      AutoSize        =   -1  'True
      Caption         =   "F7 = Delete"
      Height          =   300
      Left            =   6600
      TabIndex        =   22
      Top             =   10440
      Width           =   1815
   End
   Begin VB.Label lblF6 
      AutoSize        =   -1  'True
      Caption         =   "F6 = Add"
      Height          =   300
      Left            =   3480
      TabIndex        =   21
      Top             =   10440
      Width           =   1320
   End
   Begin VB.Label lblF3 
      AutoSize        =   -1  'True
      Caption         =   "F3 = Exit"
      Height          =   300
      Left            =   360
      TabIndex        =   20
      Top             =   10440
      Width           =   1485
   End
   Begin VB.Line Line2 
      X1              =   360
      X2              =   15000
      Y1              =   10320
      Y2              =   10320
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmCustomerMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AddSwch As Boolean
Dim EdtSwch As Boolean
Dim DelSwch As Boolean
Dim F03Swch As Boolean
Dim ValidCode As Boolean
Dim Arrange As Integer

Private Sub Form_Load()
    AddSwch = False
    EdtSwch = False
    DelSwch = False
    F03Swch = False
    ValidCode = False
    Arrange = 0   ' by default, arrange grid by code
    Call SetGrid
End Sub

Private Sub SetGrid()
    Dim ColCounter As Integer
    
    With grdCustomers
        .Rows = 1
        .ColWidth(0) = 1300
        .ColWidth(1) = 6010
        .ColWidth(2) = 6050
        .ColWidth(3) = 0 '2500
        .ColWidth(4) = 0 '2500
        .ColWidth(5) = 0 '2500
        .ColWidth(6) = 0 '2500
        .ColWidth(7) = 990
        .ColWidth(8) = 0 '2000
        .ColWidth(9) = 0 '3000
        .Refresh
    End With
    
    Call FillGrid
    
End Sub

Private Sub FillGrid()
    Dim rstCustomers As New ADODB.Recordset
    Dim strSQL As String
    
    Screen.MousePointer = vbHourglass
    
    Select Case Arrange
      Case 0  ' by code
        strSQL = "Select * from CUSTOMER order by cuscde"
      Case 1  ' by name
        strSQL = "Select * from CUSTOMER order by cusnam"
      Case 2  ' by agent
        strSQL = "Select * from CUSTOMER order by careof,cusnam"
      Case 3  ' by type
        strSQL = "Select * from CUSTOMER order by custyp,cusnam"
    End Select
        
    With grdCustomers
        rstCustomers.Open strSQL, gcnnBilling, , , adCmdText
        .Clear
        .Refresh
        .Visible = False
        .Rows = 1
        .Row = 0
        Do Until rstCustomers.EOF
            .Col = 0: .CellAlignment = flexAlignLeftCenter
            .Text = "" & rstCustomers!cuscde
            .Col = 1: .CellAlignment = flexAlignLeftCenter
            .Text = "" & Trim(rstCustomers!cusnam)
            .Col = 2: .CellAlignment = flexAlignLeftCenter
            .Text = "" & Trim(rstCustomers!careof)
            .Col = 3
            .Text = "" & Trim(rstCustomers!cusad1)
            .Col = 4
            .Text = "" & Trim(rstCustomers!cusad2)
            .Col = 5
            .Text = "" & Trim(rstCustomers!cusad3)
            .Col = 6
            .Text = "" & Trim(rstCustomers!telfax)
            .Col = 7
            .Text = "" & Trim(rstCustomers!custyp)
            .Col = 8
            .Text = "" & Trim(rstCustomers!Userid)
            .Col = 9
            .Text = "" & rstCustomers!sysdte
            rstCustomers.MoveNext
            If Not rstCustomers.EOF Then
              .Rows = .Rows + 1
              .Row = .Row + 1
            End If
         Loop
        rstCustomers.Close
        lblNumRow = .Rows & " row(s)"
        .Visible = True
        .Col = 0
        .Row = 0
        .ColSel = 9
        .HighLight = flexHighlightAlways
        .Refresh
    End With
    
    Screen.MousePointer = vbDefault
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    gcnnBilling.Close
End Sub

Private Sub grdCustomers_GotFocus()
  txtFndCde = "": txtFndNam = "": txtFndAgt = "": txtFndTyp = ""
End Sub

Private Sub lblAgent_Click()
  If Arrange = 2 Then Exit Sub
  If MsgBox("Sort records by agent?", vbYesNo + vbQuestion, "Sort") = vbNo Then Exit Sub
  Arrange = 2
  FillGrid
End Sub

Private Sub lblCode_Click()
  If Arrange = 0 Then Exit Sub
  If MsgBox("Sort records by customer code?", vbYesNo + vbQuestion, "Sort") = vbNo Then Exit Sub
  Arrange = 0
  FillGrid
End Sub

Private Sub lblCusNam_Click()
  If Arrange = 1 Then Exit Sub
  If MsgBox("Sort records by customer name?", vbYesNo + vbQuestion, "Sort") = vbNo Then Exit Sub
  Arrange = 1
  FillGrid
End Sub

Private Sub lblType_Click()
  If Arrange = 3 Then Exit Sub
  If MsgBox("Sort records by customer type?", vbYesNo + vbQuestion, "Sort") = vbNo Then Exit Sub
  Arrange = 3
  FillGrid
End Sub

Private Sub mnuAdd_Click()
  If AddSwch = False And EdtSwch = False Then ClearDetails
End Sub

Private Sub mnuDelete_Click()
  If AddSwch = False And EdtSwch = False Then DelCustomer
End Sub

Private Sub mnuEdit_Click()
    If AddSwch = True Or EdtSwch = True Then Exit Sub
    EdtSwch = True
    Call DisableFunctions
    txtCustCde.Enabled = False
    lblUserid.Caption = gUserID
    lblDateTime.Caption = gzGetSysDate
    txtCustNam.SetFocus
End Sub

Private Sub mnuExit_Click()
  ExitProgram
End Sub

Private Sub mnuPrint_Click()
  frmCustomerList.Show vbModal
End Sub

Private Sub txtAdd1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtAgent, txtAdd2)
End Sub

Private Sub txtAdd1_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtAdd2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtAdd1, txtAdd3)
End Sub

Private Sub txtAdd2_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtAdd3_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtAdd2, txtTelFax)
End Sub

Private Sub txtAdd3_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
                
Private Sub txtAgent_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtCustNam, txtAdd1)
End Sub

Private Sub txtAgent_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCustCde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
        Call txtCustCde_LostFocus
        If ValidCode = True Then
            txtCustNam.SetFocus
        End If
    End If
    If KeyCode = vbKeyEscape Then
        F03Swch = True
        Call EnableFunctions
    End If
    If KeyCode = vbKeyF3 Then
        F03Swch = True
        Call ExitProgram
    End If
End Sub

Private Sub txtCustNam_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtCustCde, txtAgent)
End Sub

Private Sub txtCustNam_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCustTyp_KeyDown(KeyCode As Integer, Shift As Integer)
' *** F03Swch is set to true to bypass txtCustTyp's lostfocus event *** '
    Select Case KeyCode
        Case vbKeyReturn
            Call txtCustTyp_LostFocus
            F03Swch = True
        Case vbKeyUp
            F03Swch = True
            txtTelFax.SetFocus
        Case vbKeyEscape
            F03Swch = True
            Call EnableFunctions
        Case vbKeyF3
            F03Swch = True
            Call ExitProgram
    End Select
End Sub

Private Sub txtCustTyp_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFndAgt_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) >= " " And Chr(KeyAscii) <= "~" Then
        SmartType txtFndAgt, KeyAscii, 2
        KeyAscii = 0
    End If
End Sub

Private Sub txtFndCde_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) >= " " And Chr(KeyAscii) <= "~" Then
        SmartType txtFndCde, KeyAscii, 0
        KeyAscii = 0
    End If
End Sub
Private Sub txtFndNam_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) >= " " And Chr(KeyAscii) <= "~" Then
        SmartType txtFndNam, KeyAscii, 1
        KeyAscii = 0
    End If
End Sub

Private Sub txtFndTyp_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) >= " " And Chr(KeyAscii) <= "~" Then
        SmartType txtFndTyp, KeyAscii, 7
        KeyAscii = 0
    End If
End Sub

Private Sub txtTelFax_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, txtAdd3, txtCustTyp)
End Sub

'Private Sub grdCustomers_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyF3                    ' EXIT PROGRAM
'            Call ExitProgram
'        Case vbKeyF6                    ' ADD CUSTOMER
'            Call ClearDetails
'        Case vbKeyF7                    ' DELETE CUSTOMER
'            Call DelCustomer
'        Case vbKeyF8                    ' EDIT CUSTOMER
'            EdtSwch = True
'            Call DisableFunctions
'            txtCustCde.Enabled = False
'            lblUserid.Caption = gUserID
'            lblDateTime.Caption = gzGetSysDate
'            txtCustNam.SetFocus
'        Case vbKeyF9                   ' PRINT CUSTOMER LIST
'            frmCustomerList.Show vbModal
'    End Select
'End Sub

Private Sub ExitProgram()
    Unload Me
End Sub

Private Sub grdCustomers_RowColChange()
    txtCustCde = Trim(grdCustomers.TextMatrix(grdCustomers.Row, 0))
    txtCustNam = Trim(grdCustomers.TextMatrix(grdCustomers.Row, 1))
    txtAgent = Trim(grdCustomers.TextMatrix(grdCustomers.Row, 2))
    txtAdd1 = Trim(grdCustomers.TextMatrix(grdCustomers.Row, 3))
    txtAdd2 = Trim(grdCustomers.TextMatrix(grdCustomers.Row, 4))
    txtAdd3 = Trim(grdCustomers.TextMatrix(grdCustomers.Row, 5))
    txtTelFax = Trim(grdCustomers.TextMatrix(grdCustomers.Row, 6))
    txtCustTyp = Trim(grdCustomers.TextMatrix(grdCustomers.Row, 7))
    lblUserid = Trim(grdCustomers.TextMatrix(grdCustomers.Row, 8))
    lblDateTime = Trim(grdCustomers.TextMatrix(grdCustomers.Row, 9))
End Sub

Private Sub txtCustCde_LostFocus()
    ValidCode = False
    If F03Swch = True Then
        F03Swch = False
        Exit Sub
    End If
    If txtCustCde = "" Or IsNumeric(txtCustCde) = False Then
        MsgBox "Enter a valid customer code.", vbExclamation, "Customer Code Error"
        txtCustCde.SetFocus
        Exit Sub
    End If
    If AddSwch = True Then  ' Check if customer code already exists
        If FindCode(txtCustCde) Then
            MsgBox "This customer code already exists. ", vbExclamation, "Customer Code Error"
            txtCustCde.SetFocus
            Exit Sub
        End If
    End If
    txtCustCde.Text = Format(txtCustCde, "000000")
    ValidCode = True
End Sub

Private Sub txtCustTyp_LostFocus()
    Dim Reply As Integer
    
    If F03Swch = True Then
        F03Swch = False
        Exit Sub
    End If
    Reply = MsgBox("Save this record ?", vbYesNo, "Save")
    If Reply = vbYes Then
        If AddSwch = True Then
            Call AddCustomer
        ElseIf EdtSwch = True Then
            Call SetRecord
        End If
    End If
    Call EnableFunctions
    AddSwch = False
    EdtSwch = False
End Sub

Private Sub DisableFunctions()
    grdCustomers.HighLight = flexHighlightNever
    fraSearch.Enabled = False
    grdCustomers.Enabled = False
    fraDetails.Enabled = True
    lblF6.ForeColor = &H80000011
    lblF7.ForeColor = &H80000011
    lblF8.ForeColor = &H80000011
    lblF9.ForeColor = &H80000011
End Sub

Private Sub EnableFunctions()
    lblF6.ForeColor = &H80000012
    lblF7.ForeColor = &H80000012
    lblF8.ForeColor = &H80000012
    lblF9.ForeColor = &H80000012
    fraSearch.Enabled = True
    grdCustomers.Enabled = True
    grdCustomers.Col = 0
    grdCustomers.ColSel = 9
    grdCustomers.HighLight = flexHighlightAlways
    txtCustCde.Enabled = True
    fraDetails.Enabled = False
    grdCustomers.SetFocus
    AddSwch = False
    EdtSwch = False
    Call grdCustomers_RowColChange
End Sub

Private Sub ClearDetails()
    Call DisableFunctions
    AddSwch = True
    txtCustCde = ""
    txtCustNam = ""
    txtAgent = ""
    txtAdd1 = ""
    txtAdd2 = ""
    txtAdd3 = ""
    txtTelFax = ""
    txtCustTyp = ""
    lblUserid.Caption = gUserID
    lblDateTime.Caption = gzGetSysDate
    txtCustCde.SetFocus
End Sub

Private Sub AddCustomer()
    Dim rstCustomers As New ADODB.Recordset
        
    With rstCustomers
        .Open "CUSTOMER", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
        .AddNew
        .Fields("cuscde") = Trim(txtCustCde)
        .Fields("cusnam") = Trim(txtCustNam)
        .Fields("careof") = Trim(txtAgent)
        .Fields("cusad1") = Trim(txtAdd1)
        .Fields("cusad2") = Trim(txtAdd2)
        .Fields("cusad3") = Trim(txtAdd3)
        .Fields("telfax") = Trim(txtTelFax)
        .Fields("custyp") = Trim(txtCustTyp)
        .Fields("userid") = Trim(lblUserid)
        .Fields("sysdte") = gzGetSysDate
        .Fields("status") = ""
        .Fields("updcde") = ""
        .Update
        .Close
    End With
    Call FillGrid
    Call EnableFunctions
End Sub

Private Sub DelCustomer()
    Dim Reply As Integer
    
    Reply = MsgBox("Delete this record ?", vbYesNo + vbDefaultButton2, "Delete")
    If Reply = vbNo Then
        Exit Sub
    End If
    If FindADR(txtCustCde) Then
        MsgBox "This customer has ADR transactions, it cannot be deleted.", vbInformation, "Delete Error"
    Else
        DelSwch = True
        Call SetRecord
        DelSwch = False
    End If
End Sub

Private Sub SetRecord() ' Sets a pointer to specific record, delete when requested
    Dim rstCustomers As New ADODB.Recordset
        
    With rstCustomers
        .Open "Select * from CUSTOMER where cuscde = '" & Trim(txtCustCde) & _
        "'", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
        If DelSwch = True Then
            .Delete
        ElseIf EdtSwch = True Then
            .Fields("cusnam") = Trim(txtCustNam)
            .Fields("careof") = Trim(txtAgent)
            .Fields("cusad1") = Trim(txtAdd1)
            .Fields("cusad2") = Trim(txtAdd2)
            .Fields("cusad3") = Trim(txtAdd3)
            .Fields("telfax") = Trim(txtTelFax)
            .Fields("custyp") = Trim(txtCustTyp)
            .Fields("userid") = Trim(lblUserid)
            .Fields("sysdte") = gzGetSysDate
            .Fields("status") = ""
            .Fields("updcde") = "U"
            .Update
        End If
        .Close
    End With
    Call FillGrid
    
End Sub

Private Sub FieldAdvance(pKeycode As Integer, pPreviousCtl As Control, pNextCtl As Control)
    Select Case pKeycode
        Case vbKeyReturn
            pNextCtl.SetFocus
        Case vbKeyDown
            pNextCtl.SetFocus
        Case vbKeyUp
            If pPreviousCtl.Enabled = True Then
                pPreviousCtl.SetFocus
            End If
        Case vbKeyEscape
            Call EnableFunctions
    End Select
End Sub

Private Function FindADR(strCustCode As String) As Boolean
    Dim rstADR As New ADODB.Recordset
    
    With rstADR
        .Open "Select * from ADRFLE where cuscde = '" & Trim(strCustCode) & _
              "'", gcnnBilling, , , adCmdText
        FindADR = Not .EOF
        .Close
    End With
End Function

Private Function FindCode(strCustCode As String) As Boolean
    Dim rstCode As New ADODB.Recordset
    
    With rstCode
        .Open "Select * from CUSTOMER where cuscde = '" & Trim(strCustCode) & _
              "'", gcnnBilling, , , adCmdText
        FindCode = Not .EOF
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

    Do While (i < grdCustomers.Rows - 1) _
        And StrComp(Left(t, s), Left(grdCustomers.TextMatrix(i, pColNum), s), vbTextCompare) <> 0
        i = i + 1
    Loop
    If UCase(Left(t, s)) = UCase(Left(grdCustomers.TextMatrix(i, pColNum), s)) Then
        t = Trim(grdCustomers.TextMatrix(i, pColNum))
        l = Len(t) - s
    Else
        t = Left(t, s)
        l = 0
    End If

    c.Text = t
    c.SelStart = s
    c.SelLength = l
    grdCustomers.TopRow = i
    grdCustomers.Row = i
    grdCustomers.Col = 0
    grdCustomers.ColSel = 9
End Sub
