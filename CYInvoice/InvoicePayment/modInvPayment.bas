Attribute VB_Name = "modInvPayment"
       
'Public Const strCnn = "Provider=sqloledb;Data Source=NTBS;Initial Catalog=BILLING;Integrated Security=SSPI"

Public gcnnBilling As ADODB.Connection
Public gbConnected As Boolean
Public rsUnsettled As ADODB.Recordset

Public Type OR_Payment
  ornum As Long
  ortype As String * 3
  cuscde As String * 6
  CheckAMT1 As Currency
  CheckNo1 As String * 10
  CheckBnk1 As String * 10
  CheckAMT2 As Currency
  CheckNo2 As String * 10
  CheckBnk2 As String * 10
  CashAMT As Currency
  TotalAmt As Currency
  AvailAMT As Currency
  ORDate As Date
  Userid As String * 10
End Type


Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function WNetGetUser Lib "mpr" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long

Public Function zCurrentComputer() As String
Dim lSize As Long
Dim sBuffer As String
    sBuffer = Space$(15& + 1)
    lSize = Len(sBuffer)
    zCurrentComputer = ""
    If GetComputerName(sBuffer, lSize) Then
        zCurrentComputer = Left$(sBuffer, lSize)
    End If
End Function

Public Function zCurrentUser() As String
Dim lpUserName As String * 64
    If WNetGetUser("", lpUserName, Len(lpUserName)) Then
        zCurrentUser = ""
    Else
        zCurrentUser = Left(lpUserName, InStr(lpUserName, Chr(0)) - 1)
    End If
End Function

Public Sub FilterRecordset(ByVal pCustomer As String)
Dim sSql_Unpaid  As String
 sSql_Unpaid = "Select  invnum,cuscde,cusnam,invamt,isnull(invtax,0),isnull(invvat,0),isnull(totalpay,0),invdttm,status,(invamt + isnull(invvat,0) - isnull(invtax,0)) as invtotal, " _
   & " ((invamt+ isnull(invvat,0) - isnull(invtax,0))- isnull(totalPay,0)) as Balance From invict " _
   & " Where ((invamt + IsNull(invvat, 0) - IsNull(invtax, 0)) <> IsNull(totalpay, 0) Or totalpay Is Null) " _
   & " AND (UPPER(LTRIM(status)) not like 'CAN' or status is NULL) " _
   & " order by cuscde,invnum"
     
   If rsUnsettled.State = adStateOpen Then
      rsUnsettled.Filter = adFilterNone
      rsUnsettled.Requery
   Else
       rsUnsettled.Open sSql_Unpaid, gcnnBilling, adOpenKeyset, , adCmdText
   End If
   
   If UCase(Trim(pCustomer)) <> "NONE" Then
       rsUnsettled.Filter = "cusnam='" & pCustomer & "'"
   Else
      rsUnsettled.Filter = adFilterNone
   End If

End Sub

Public Sub Inialized_Grid()
  Dim colhdgs(6) As String
  Dim col As Integer
  
  colhdgs(0) = " Inv. No. "
  colhdgs(1) = " Inv. Amount "
  colhdgs(2) = "Customer Code"
  colhdgs(3) = " Customer Name "
  colhdgs(4) = " Date Issued "
  colhdgs(5) = " Balance "
  With frmMain.grd_InvList
        .Clear
        .Rows = 2
        .row = 0
  
    For col = 0 To 5
        .col = col: .Text = colhdgs(col): .CellAlignment = 4: .CellFontBold = True: .CellForeColor = &HFFFF&: .CellBackColor = &HC00000
    Next col
  
        .ColWidth(0) = 1500
        .ColWidth(1) = 2300
        .ColWidth(2) = 1900
        .ColWidth(3) = 4500
        .ColWidth(4) = 2800
        .ColWidth(5) = 2000
  End With
   
End Sub



Public Sub List_UnpaidBills()
 Dim rowcount, col As Integer
 Dim blnToggle As Boolean
 Dim RowColor  As Long
 
  frmMain.MousePointer = 11
   frmMain.grd_InvList.Rows = 2

   If rsUnsettled.RecordCount < 1 Then
         frmMain.grd_InvList.Clear
         Call Inialized_Grid
          frmMain.staStatus.Panels(1).Text = "Records: NONE... All Bills have been settled "
          MsgBox "Empty list of Invoice for Customer " & frmMain.cmbcust.Text, vbOKOnly + vbInformation, "Customer Bill "
           
   ElseIf rsUnsettled.RecordCount > 0 Then   'Populate the grid
      rowcount = 1
    With frmMain.grd_InvList
       frmMain.staStatus.Panels(1).Text = "Please wait...Query List to Database"
       blnToggle = False
       .Visible = False
       frmMain.AutoRedraw = True
        
        Do While Not rsUnsettled.EOF
            If rowcount > 1 Then
               .AddItem ""  ' add another row
            End If
            
            RowColor = IIf(blnToggle = False, &H80000018, vbWhite)
            .RowHeight(rowcount) = 250
            .row = rowcount
            .col = 0: .Text = Trim(rsUnsettled!invnum): .CellAlignment = 4: .CellBackColor = RowColor
            .col = 1: .Text = Format(rsUnsettled!invamt, "###,###,###.#0"): .CellBackColor = RowColor
            .col = 2: .Text = rsUnsettled!cuscde: .CellAlignment = 4: .CellBackColor = RowColor
            .col = 3: .Text = UCase(rsUnsettled!cusnam): .CellBackColor = RowColor
            .col = 4: .Text = Trim(rsUnsettled!invdttm): .CellBackColor = RowColor
            .col = 5: .Text = Format(rsUnsettled!balance, "(###,###,###.#0)"):  .CellBackColor = RowColor
             rowcount = rowcount + 1
            rsUnsettled.MoveNext
            blnToggle = Not blnToggle
           Loop
        frmMain.AutoRedraw = False
        .Visible = True
      
        frmMain.staStatus.Panels(1).Text = "Records: " & .Rows - 1
        'set row selection
        .Enabled = True
        .col = 0
        .row = 1
        .ColSel = 5
        .SelectionMode = flexSelectionByRow
        .HighLight = flexHighlightAlways
      End With
    End If
frmMain.MousePointer = 0
End Sub


'--------------------------------------------------------------------
' Function      :   gzGetSysDate()
' Parameters    :   none
' Returns       :   DateTime    -> Server Date and Time
'--------------------------------------------------------------------
Public Function gzGetSysDate() As Date
Dim cmdGetSysDate As ADODB.Command
Dim prmGetSysDate As ADODB.Parameter
Dim X As Date
    
    ' create command
    Set cmdGetSysDate = New ADODB.Command
    Set prmGetSysDate = New ADODB.Parameter
    With cmdGetSysDate
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getsysdate"
        .CommandType = adCmdStoredProc
        Set prmGetSysDate = .CreateParameter("pDATE", adDate, adParamOutput)
        .Parameters.Append prmGetSysDate
        .Execute
        gzGetSysDate = .Parameters("pDATE")
    End With
End Function

