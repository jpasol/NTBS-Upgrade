VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmSubicINVCorrection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CY Invoice Correction & Cancellation"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   15270
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
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame fraHeading 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   1080
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   12975
      Begin MSFlexGridLib.MSFlexGrid grdCustomers 
         Height          =   5775
         Left            =   600
         TabIndex        =   24
         Top             =   2880
         Visible         =   0   'False
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   10186
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         FocusRect       =   2
         ScrollBars      =   2
         SelectionMode   =   1
      End
      Begin VB.TextBox txtCusNum 
         Height          =   420
         Left            =   3360
         MaxLength       =   6
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtAgent 
         BackColor       =   &H80000014&
         Height          =   420
         Left            =   3360
         MaxLength       =   29
         TabIndex        =   3
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox txtVslNam 
         BackColor       =   &H80000014&
         Height          =   420
         Left            =   3360
         MaxLength       =   20
         TabIndex        =   7
         Top             =   3240
         Width           =   4935
      End
      Begin VB.TextBox txtRegistry 
         BackColor       =   &H80000014&
         Height          =   420
         Left            =   3360
         MaxLength       =   8
         TabIndex        =   6
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtRemark 
         BackColor       =   &H80000014&
         Height          =   405
         Left            =   3360
         MaxLength       =   165
         TabIndex        =   5
         Top             =   2040
         Width           =   9135
      End
      Begin VB.TextBox txtInvNum 
         BackColor       =   &H80000014&
         Height          =   390
         Left            =   3360
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1440
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mskVslArv 
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   3840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "####/##/##"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbNonCon 
         Height          =   420
         ItemData        =   "frmSubicINVCorrection.frx":0000
         Left            =   3360
         List            =   "frmSubicINVCorrection.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4440
         Width           =   2295
      End
      Begin VB.ComboBox cmbGtyCde 
         Height          =   420
         ItemData        =   "frmSubicINVCorrection.frx":0027
         Left            =   3360
         List            =   "frmSubicINVCorrection.frx":0031
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   5040
         Width           =   2775
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Bill Category"
         Height          =   300
         Left            =   720
         TabIndex        =   31
         Top             =   5040
         Width           =   2145
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Non-containerized"
         Height          =   300
         Left            =   120
         TabIndex        =   30
         Top             =   4440
         Width           =   2805
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Arrival Date"
         Height          =   300
         Left            =   960
         TabIndex        =   29
         Top             =   3840
         Width           =   1980
      End
      Begin VB.Label lblTax 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.000"
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   11670
         TabIndex        =   28
         Top             =   6240
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(Tax)"
         Height          =   300
         Left            =   9480
         TabIndex        =   27
         Top             =   6240
         Width           =   825
      End
      Begin VB.Label lblCusNam 
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   25
         Top             =   240
         Width           =   7695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agent"
         Height          =   300
         Left            =   2040
         TabIndex        =   23
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vessel Name"
         Height          =   300
         Left            =   1080
         TabIndex        =   22
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Registry"
         Height          =   300
         Left            =   1560
         TabIndex        =   21
         Top             =   2640
         Width           =   1320
      End
      Begin VB.Label lblVat 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.000"
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   11670
         TabIndex        =   17
         Top             =   5760
         Width           =   825
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   11835
         TabIndex        =   16
         Top             =   5280
         Width           =   660
      End
      Begin VB.Label lblInvTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   300
         Left            =   9480
         TabIndex        =   15
         Top             =   5280
         Width           =   825
      End
      Begin VB.Label lblInvVAT 
         AutoSize        =   -1  'True
         Caption         =   "VAT"
         Height          =   300
         Left            =   9840
         TabIndex        =   14
         Top             =   5760
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bill Remarks"
         Height          =   300
         Left            =   960
         TabIndex        =   13
         Top             =   2040
         Width           =   1980
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Invoice Number"
         Height          =   300
         Left            =   600
         TabIndex        =   12
         Top             =   1440
         Width           =   2310
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Customer Number"
         Height          =   300
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   2475
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Width           =   15135
      Begin VB.TextBox txtRefNum 
         BackColor       =   &H80000014&
         Height          =   375
         Left            =   4320
         MaxLength       =   8
         TabIndex        =   0
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Enter Reference Number"
         Height          =   375
         Left            =   360
         TabIndex        =   33
         Top             =   480
         Width           =   3735
      End
   End
   Begin VB.Label lblF4 
      Caption         =   "F4 = Customer Picklist"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   3360
      TabIndex        =   26
      Top             =   10200
      Width           =   3855
   End
   Begin VB.Label lblF10 
      Caption         =   "F10 = Save"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   11880
      TabIndex        =   20
      Top             =   10200
      Width           =   1935
   End
   Begin VB.Label lblF7 
      Caption         =   "F7 = Cancel Invoice"
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   7800
      TabIndex        =   19
      Top             =   10200
      Width           =   3255
   End
   Begin VB.Label lblF3 
      Caption         =   "F3 = Exit"
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   10200
      Width           =   1695
   End
End
Attribute VB_Name = "frmSubicINVCorrection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CmbFill As Boolean
Public tempInv As String


Private Sub cmbGtyCde_GotFocus()
    SendKeys "%{DOWN}"
End Sub

Private Sub cmbGtyCde_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeydownEvents(KeyCode, cmbNonCon, txtCusNum)
End Sub

Private Sub cmbNonCon_GotFocus()
    SendKeys "%{DOWN}"
End Sub

Private Sub cmbNonCon_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeydownEvents(KeyCode, mskVslArv, cmbGtyCde)
End Sub

Private Sub Form_Load()
    CmbFill = False
    tempInv = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim msg
    msg = "Do you want to exit the program?"
    If MsgBox(msg, vbQuestion + vbYesNo, "Exit") = vbNo Then Cancel = True
End Sub

Private Sub mskVslArv_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub mskVslArv_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeydownEvents(KeyCode, txtVslNam, cmbNonCon)
End Sub

Private Sub txtAgent_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtAgent_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeydownEvents(KeyCode, txtCusNum, txtInvNum)
End Sub

Private Sub txtAgent_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCusNum_GotFocus()
    lblF4.ForeColor = &H80000012
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtCusNum_LostFocus()
    lblF4.ForeColor = &H80000011
End Sub

Private Sub txtInvNum_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtInvNum_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim msg As String
 Dim isExist As Boolean
 Dim has_partial As Boolean
 Dim isCancelled  As Boolean
 msg = ""
 
 If txtInvNum.Text <> "" Or IsNumeric(txtInvNum.Text) Then
    isExist = HasDuplicate(Trim(txtInvNum.Text))
    has_partial = Invoice_Has_PartialPayment(Trim(Str(txtInvNum.Text)), msg)
    isCancelled = Is_InvoiceCancelled(Trim(Str(txtInvNum.Text)))

If KeyCode = 13 And isCancelled = True Then
      MsgBox "Cannot update/modify Invoice # " & txtInvNum.Text & Chr(13) & " This is a Cancelled Invoice! ", vbOKOnly + vbInformation, "Error"
      txtInvNum.Text = tempInv
      SendKeys "{HOME}": SendKeys "+{END}"

ElseIf KeyCode = 13 And has_partial = True Then
      MsgBox "Cannot update/modify Invoice # " & txtInvNum.Text & Chr(13) & msg, vbOKOnly + vbInformation, "Error"
      txtInvNum.Text = tempInv
      SendKeys "{HOME}": SendKeys "+{END}"
      
ElseIf KeyCode = 13 And isExist = True And tempInv <> txtInvNum Then
      MsgBox "Cannot update/modify Invoice # " & txtInvNum.Text & Chr(13) & " Invoice # already exist ", vbOKOnly + vbInformation, "Error"
      txtInvNum.Text = tempInv
      SendKeys "{HOME}": SendKeys "+{END}"

Else
   Call KeydownEvents(KeyCode, txtAgent, txtRemark)
End If
End If
End Sub

Private Sub txtRefNum_GotFocus()
    txtRefNum = txtRefNum
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtRegistry_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtRegistry_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeydownEvents(KeyCode, txtRemark, txtVslNam)
End Sub

Private Sub txtRegistry_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRemark_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeydownEvents(KeyCode, txtInvNum, txtRegistry)
End Sub

Private Sub txtVslNam_GotFocus()
    SendKeys "{HOME}": SendKeys "+{END}"
End Sub

Private Sub txtVslNam_KeyDown(KeyCode As Integer, Shift As Integer)
    Call KeydownEvents(KeyCode, txtRegistry, mskVslArv)
End Sub

Private Sub KeydownEvents(pKeycode As Integer, pPreviousCtl As Control, pNextCtl As Control)
    Select Case pKeycode
        Case vbKeyUp
            pPreviousCtl.SetFocus
        Case vbKeyReturn
            pNextCtl.SetFocus
        Case vbKeyF3
            Call NewReference
        Case vbKeyF7
            Call CancelInvoice
        Case vbKeyF10
            Call SaveInvoice
    End Select
End Sub

Private Sub txtRefNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then Unload Me
    If KeyCode = vbKeyReturn Then Call CheckReference
End Sub

Private Sub txtVslNam_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub CheckReference()
    Dim rstReference As New ADODB.Recordset
    Dim tempRemark As String
    Dim RemarkPos As Integer
    Dim msg As String
    Dim isExist As Boolean
    Dim has_partial As Boolean
    Dim isCancelled  As Boolean
    msg = ""

    
    If Len(Trim(txtRefNum)) = 0 Or Not IsNumeric(txtRefNum) Then
        MsgBox "Specify a valid reference number.", vbInformation, "Error Message"
        txtRefNum.SetFocus
        Exit Sub
    End If
    
    isExist = IsRefNumExist(Trim(txtRefNum.Text))
    has_partial = Has_PartialPayment(Trim(Str(txtRefNum.Text)), msg)
    isCancelled = Is_Cancelled(Trim(Str(txtRefNum.Text)))
    
  ' if Ref # already exist  then check if there is a partial payment made
If isExist = False Then
       MsgBox "Invoice does not exist!", vbExclamation, "Error"
       SendKeys "{HOME}": SendKeys "+{END}"
       Exit Sub
    
ElseIf isExist = True And isCancelled = True Then
       MsgBox "Invoice has already been cancelled.", vbExclamation, "Invoice Cancelled"
       SendKeys "{HOME}": SendKeys "+{END}"
       Exit Sub
       
ElseIf has_partial = True Then
       MsgBox msg, vbExclamation, "Invoice Cancelled"
       SendKeys "{HOME}": SendKeys "+{END}"
        Exit Sub
        
ElseIf (isExist = True And has_partial = False) And isCancelled = False Then
    With rstReference
        .Open "Select * from INVICT where refnum = '" & _
            Trim(txtRefNum) & "' and cfscy = '" & "1" & _
            "'", gcnnBilling, , , adCmdText
'        If Not .EOF Then
'            If !Status = "CAN" Then
'                MsgBox "Invoice has already been cancelled.", vbExclamation, "Invoice Cancelled"
'                txtRefNum.SetFocus
'                .Close
'                Exit Sub
'            End If
'        Else
'            MsgBox "Invoice does not exist.", vbExclamation, "Error"
'            txtRefNum.SetFocus
'            .Close
'            Exit Sub
'        End If
    
        ' The following commands are executed if reference entered is valid
        lblF7.ForeColor = &H80000012
        lblF10.ForeColor = &H80000012
        txtRefNum.Enabled = False
        tempInv = !invnum
        txtCusNum = !cuscde
        lblCusNam = Trim(!cusnam)
        txtInvNum = !invnum
        txtRegistry = "" & Trim(!regnum)
        txtVslNam = "" & Trim(!vslnam)
        lblTotal.Caption = Format(!invamt, "#,###.#0")
        lblVat.Caption = Format(!invvat, "#,###.##0")
        lblTax.Caption = Format(!invtax, "#,###.##0")
        tempRemark = "" & !invremark
        txtRemark = ""
        txtAgent = ""
        If Trim(tempRemark) <> "|" Then  ' when invremark is not empty
            RemarkPos = InStr(1, Trim(tempRemark), "|")
            If RemarkPos = 1 Then        ' when only the agent exists
                txtAgent = Trim(Mid(tempRemark, 2))
            Else
                txtRemark = Trim(Mid(tempRemark, 1, RemarkPos - 1))
                txtAgent = Trim(Mid(tempRemark, RemarkPos + 1))
            End If
        End If
        If Not IsNull(!arrival) Then
           mskVslArv.Text = Format(!arrival, "yyyy/mm/dd")
        End If
        Select Case !noncnt
            Case "0"
                cmbNonCon.ListIndex = 0
            Case "1"
                cmbNonCon.ListIndex = 1
            Case "2"
                cmbNonCon.ListIndex = 2
        End Select
        If !gtycde = "N" Then
            cmbGtyCde.ListIndex = 0
        Else
            cmbGtyCde.ListIndex = 1
        End If
        .Close
    End With
    fraHeading.Visible = True
End If
    
End Sub

Private Sub txtCusNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Then
        Call ViewPicklist
    Else
        Call KeydownEvents(KeyCode, txtCusNum, txtAgent)
    End If
End Sub

Private Sub ViewPicklist()
    If Not CmbFill Then SetPicklist
    With grdCustomers
        .Visible = True
        .Col = 0
        .Row = 1
        .ColSel = 1
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
        .SetFocus
    End With
    lblF7.ForeColor = &H80000011: lblF10.ForeColor = &H80000011
End Sub

Private Sub SearchCustomerPicklist()
Dim ctr As Integer
    ctr = 0
    lblCusNam = "" ': txtAgent = ""
    On Error Resume Next
    With grdCustomers
         .Col = 0: .Row = 1
         Do Until ctr = .Rows - 1
             If Trim(txtCusNum) = Trim(.Text) Then
                .Col = 1: lblCusNam.Caption = Trim(.Text)
                .Col = 2: If Trim(txtAgent) = "" Then txtAgent = Trim(.Text)
                Exit Do
            End If
            .Col = 0: .Row = .Row + 1
            ctr = ctr + 1
         Loop
    End With
End Sub

Private Sub grdCustomers_KeyDown(KeyCode As Integer, Shift As Integer)
    With grdCustomers
        If KeyCode = vbKeyReturn Then
            .Col = 0: txtCusNum = Trim(.Text)
            .Col = 1: lblCusNam = Trim(.Text)
            .Col = 2: txtAgent = Trim(.Text)
            .Visible = False
            lblF7.ForeColor = &H80000012: lblF10.ForeColor = &H80000012
            txtAgent.SetFocus
        ElseIf KeyCode = vbKeyEscape Then
            .Visible = False
            lblF7.ForeColor = &H80000012: lblF10.ForeColor = &H80000012
            txtCusNum.SetFocus
        End If
    End With
End Sub

Private Sub SetPicklist()
    Dim rstCustomers As New ADODB.Recordset
    Dim rowCount As Integer
        
    CmbFill = True
    rowCount = 1
    With grdCustomers
        .Col = 0: .Row = 0: .Text = "  Code"
        .Col = 1: .Row = 0: .Text = "                  Name"
        .RowHeight(0) = 350
        .ColWidth(0) = 1300
        .ColWidth(1) = 7490
        .ColWidth(2) = 0
        .HighLight = flexHighlightAlways
        .Refresh
        rstCustomers.Open "Select * from CUSTOMER order by cusnam", gcnnBilling, , , adCmdText
    
        Do While Not rstCustomers.EOF
            If rowCount > 1 Then
                .AddItem ("")
            End If
            .RowHeight(rowCount) = 350
            .Row = rowCount
            .Col = 0
            .CellAlignment = 4
            .Text = Trim(rstCustomers.Fields("cuscde"))
            .Col = 1: .CellAlignment = 1
            .Text = Trim(rstCustomers.Fields("cusnam"))
            .Col = 2: .CellAlignment = 1
            .Text = Trim(rstCustomers.Fields("careof"))
            rowCount = rowCount + 1
            rstCustomers.MoveNext
        Loop
        rstCustomers.Close
    End With
End Sub

Private Sub CancelInvoice()
    Dim rstHeading As New ADODB.Recordset
    Dim rstDetail As New ADODB.Recordset
    Dim vReply As Integer
    
vReply = MsgBox("Cancel this invoice?", vbYesNo + vbDefaultButton2, "Cancel")
If vReply = vbYes Then
        With rstHeading
           .Open "Select * from INVICT where refnum = '" & Trim(txtRefNum) & _
            "'", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
            !Status = "CAN"
            !Userid = gUserID
            .Update
            .Close
        End With
        With rstDetail
            .Open "Select * from INVCYB where refnum = '" & Trim(txtRefNum) & _
            "'", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
            Do While Not .EOF
                !Status = "CAN"
                !Userid = gUserID
                .Update
                .MoveNext
            Loop
            .Close
        End With
        Call NewReference
End If
End Sub

Private Sub SaveInvoice()
    Dim rstHeading As New ADODB.Recordset
    Dim rstDetail As New ADODB.Recordset
    Dim vReply As Integer
    
    
    If Trim(txtCusNum) = "" Then
        MsgBox "Please specify a valid customer code.", vbInformation, "Error Message"
        txtCusNum.SetFocus
        Exit Sub
    End If
    If Trim(txtInvNum) = "" Or Not IsNumeric(txtInvNum) Then
        MsgBox "Enter a valid invoice number.", vbExclamation, "Error Message"
        txtInvNum.SetFocus
        Exit Sub
    End If
    
    vReply = MsgBox("Update this invoice?", vbYesNo, "Save")
    If vReply = vbNo Then
        Exit Sub
    End If
    
    With rstHeading
       .Open "Select * from INVICT where refnum = '" & Trim(txtRefNum) & _
        "'", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
        .Fields("cuscde") = Trim(txtCusNum)
        .Fields("cusnam") = Trim(lblCusNam)
        .Fields("invnum") = txtInvNum
        .Fields("regnum") = Trim(txtRegistry)
        .Fields("vslnam") = Trim(txtVslNam)
        .Fields("invremark") = Trim(txtRemark) & "|" & Trim(txtAgent)
        .Fields("updcde") = "U"
        If IsDate(mskVslArv.Text) Then
            .Fields("arrival") = CDate(mskVslArv.Text)
        End If
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
        .Fields("userid") = gUserID
        .Update
        .Close
    End With
    With rstDetail
        .Open "Select * from INVCYB where refnum = '" & Trim(txtRefNum) & _
        "'", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdText
        Do While Not .EOF
            .Fields("updcde") = "U"
            .Fields("userid") = gUserID
            .Update
            .MoveNext
        Loop
            .Close
    End With
    Call NewReference

               
End Sub

Private Sub NewReference()
    fraHeading.Visible = False
    lblF7.ForeColor = &H80000011
    lblF10.ForeColor = &H80000011
    txtRefNum.Enabled = True
    txtRefNum.SetFocus
End Sub


Private Function IsRefNumExist(ByVal ref As Long) As Boolean
  Dim rst As New ADODB.Recordset
  
  rst.Open "select refnum,invnum from invict where refnum=" & ref, gcnnBilling, , , adCmdText
  
  If rst.EOF Then
        IsRefNumExist = False
  Else
        IsRefNumExist = True
  End If
  rst.Close
  Set rst = Nothing
End Function

Private Function HasDuplicate(ByVal inv As Long) As Boolean
  Dim rst As New ADODB.Recordset
  
  'rst.Open "select count(invnum)as NoInv from invict where invnum=" & inv, gcnnBilling, , , adCmdText
  rst.Open "select invnum from invict where invnum=" & inv, gcnnBilling, , , adCmdText

  If rst.EOF Then
        HasDuplicate = False
  Else
        HasDuplicate = True
  End If
  rst.Close
  Set rst = Nothing
End Function

Private Function Has_PartialPayment(ByVal ref As Long, ByRef messg As String) As Boolean
  Dim rst As New ADODB.Recordset
  Dim tInvamt As Single
  
  rst.Open "select invnum,isnull(invamt,0) as iAmt ,isnull(invtax,0) as Tax,isnull(invvat,0)as VAT,isnull(totalpay,0) as Tpay from invict where refnum=" & ref, gcnnBilling, , , adCmdText
  
  messg = ""
 If Not rst.EOF Then
       tInvamt = rst!iamt + rst!VAT - rst!Tax
        If rst!Tpay = tInvamt Then
                messg = "Invoice Correction and Cancellation will not be allowed"
                messg = messg & Chr(13) & " Invoice # " & rst!invnum & " is fully paid already"
                Has_PartialPayment = True
 
        ElseIf rst!Tpay > 0 Then
                messg = "Invoice Correction and Cancellation will not be allowed"
                messg = messg & Chr(13) & " Invoice # " & rst!invnum & " Has a partial payment"
                Has_PartialPayment = True
 
        End If
  Else
        Has_PartialPayment = False
 End If
 rst.Close
 Set rst = Nothing

End Function
Private Function Is_Cancelled(ByVal ref As String) As Boolean
 Dim rst As New ADODB.Recordset
    rst.Open "Select * from INVICT where refnum =" & ref & " and cfscy = " & "1", gcnnBilling, , , adCmdText
    If Not rst.EOF Then
         If rst!Status = "CAN" Then
                Is_Cancelled = True
         Else
               Is_Cancelled = False
         End If
    End If
    rst.Close
    Set rst = Nothing
End Function
Private Function Is_InvoiceCancelled(ByVal inv As String) As Boolean
 Dim rst As New ADODB.Recordset
    rst.Open "Select invnum,status from INVICT where invnum=" & inv & " and cfscy = " & "1", gcnnBilling, , , adCmdText
    If Not rst.EOF Then
         If rst!Status = "CAN" Then
                Is_InvoiceCancelled = True
         Else
               Is_InvoiceCancelled = False
         End If
    End If
    rst.Close
    Set rst = Nothing
End Function

Private Function Invoice_Has_PartialPayment(ByVal inv As Long, ByRef messg As String) As Boolean
  Dim rst As New ADODB.Recordset
  Dim tInvamt As Single
  
  rst.Open "select invnum,status,isnull(invamt,0) as iAmt ,isnull(invtax,0) as Tax,isnull(invvat,0)as VAT,isnull(totalpay,0) as Tpay from invict where invnum=" & inv, gcnnBilling, , , adCmdText
  
  messg = ""
 If Not rst.EOF Then
       tInvamt = rst!iamt + rst!VAT - rst!Tax
        If rst!Tpay = tInvamt Then
                messg = "Invoice Correction and Cancellation will not be allowed"
                messg = messg & Chr(13) & " Invoice # " & rst!invnum & " is fully paid already"
                Invoice_Has_PartialPayment = True
 
        ElseIf rst!Tpay > 0 Then
                messg = "Invoice Correction and Cancellation will not be allowed"
                messg = messg & Chr(13) & " Invoice # " & rst!invnum & " Has a partial payment"
                Invoice_Has_PartialPayment = True
 
        End If
  Else
        Invoice_Has_PartialPayment = False
 End If
 rst.Close
 Set rst = Nothing
End Function

