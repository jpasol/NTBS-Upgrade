VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCCRPr03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'   ** Print Variable

Dim prtRefnum As Long
Dim prtSeqnum As Long
Dim prtCcrnum As Long
Dim prtDate As String
Dim prtTime As String
Dim prtCustomer As String
Dim prtVessel As String

'   ** Print Total Variable

Dim prtSubTotal As Long
Dim prtVat As Long
Dim prtWTax As Long
Dim prtGrandTotal As Long
Dim prtWords As String
Dim prtCash As String
Dim prtAdr As String
Dim prtChk As String


'local variable(s) to hold property value(s)
Private mvarCCRNumber As Variant 'local copy
Private mvarPrinterReference As Integer 'local copy
Public Function PrintCCR(pRefnum As Long) As Boolean

    Dim CRL As Recordset
    Dim ADR As Recordset
    Dim DETTl As Recordset

    VE.getCCRList (pRefnum)
    Set CRL = VE.rsgetCCRList
    
    VE.getAdramt (pRefnum)
    Set ADR = VE.rsgetAdramt

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
    prtDate = ""
    prtTime = ""

    If ADR.Fields("adramt") <> 0 Or Not IsNull(ADR.Fields("adramt")) Then
        AdrAmt = ADR.Fields("adramt")
    End If
    If ADR.Fields("chkamt1") <> 0 Or Not IsNull(ADR.Fields("chkamt1")) Then
        ChkAmt1 = ADR.Fields("chkamt1")
        ChkTotal = ChkTotal + ChkAmt1
    End If
    If ADR.Fields("chkamt2") <> 0 Or Not IsNull(ADR.Fields("chkamt2")) Then
        ChkAmt2 = ADR.Fields("chkamt2")
        ChkTotal = ChkTotal + ChkAmt2
    End If
    If ADR.Fields("chkamt3") <> 0 Or Not IsNull(ADR.Fields("chkamt3")) Then
        ChkAmt3 = ADR.Fields("chkamt3")
        ChkTotal = ChkTotal + ChkAmt3
    End If
    If ADR.Fields("chkamt4") <> 0 Or Not IsNull(ADR.Fields("chkamt4")) Then
        ChkAmt4 = ADR.Fields("chkamt4")
        ChkTotal = ChkTotal + ChkAmt4
    End If
    If ADR.Fields("chkamt5") <> 0 Or Not IsNull(ADR.Fields("chkamt5")) Then
        ChkAmt5 = ADR.Fields("chkamt5")
        ChkTotal = ChkTotal + ChkAmt5
    End If
    If ADR.Fields("cshamt") <> 0 Or Not IsNull(ADR.Fields("cshamt")) Then
        CashAmt = ADR.Fields("cshamt")
    End If
    
    ADR.Close
    Set ADR = Nothing

    With CRL
        Do While Not .EOF
              DetailTl = 0
        DetailAmt = 0
        TotalAmt = 0
        sngTempAmt = 0
        strCshAmt = ".00"
        strChqAmt = ".00"
        strAdrAmt = ".00"
        prtCash = ""
        prtAdr = ""
        VE.getTotal .Fields("refnum"), .Fields("seqnum")
        Set DETTl = VE.rsgetTotal
        DetailTl = DETTl.Fields("TotalAmt")
        DetailAmt = DETTl.Fields("TotalAmt")
        TotalAmt = DETTl.Fields("totalamt")
'   **  Checks Liquidation of ADR, Cheque and Cash
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
        If ChkAmt1 <> 0 Then
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
        End If
        If sngTempAmt = TotalAmt Then
            GoTo NextCCRTag
        End If
'   ** Cheque 2
        If ChkAmt2 <> 0 Then
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
        End If
        If sngTempAmt = TotalAmt Then
            GoTo NextCCRTag
        End If
'   ** Cheque 3
        If ChkAmt3 <> 0 Then
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
        End If
        If sngTempAmt = TotalAmt Then
            GoTo NextCCRTag
        End If
'   ** Cheque 4
        If ChkAmt4 <> 0 Then
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
        End If
        If sngTempAmt = TotalAmt Then
            GoTo NextCCRTag
        End If
'   ** Cheque 5
        If ChkAmt5 <> 0 Then
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
        End If
        If sngTempAmt = TotalAmt Then
            GoTo NextCCRTag
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
NextCCRTag:
            If mvarCCRNumber <> 0 Then
                If .Fields("ccrnum") = mvarCCRNumber Then
                    Call OutCCR(.Fields("refnum"), .Fields("seqnum"), .Fields("cusnam"), strAdrAmt, strCshAmt, strChqAmt, _
                                blnChkno1, blnChkno2, blnChkno3, blnChkno4, blnChkno5)
                End If
            Else
                Call OutCCR(.Fields("refnum"), .Fields("seqnum"), .Fields("cusnam"), strAdrAmt, strCshAmt, strChqAmt, _
                    blnChkno1, blnChkno2, blnChkno3, blnChkno4, blnChkno5)
            End If
            .MoveNext
            DETTl.Close
            Set DETTl = Nothing
        Loop
    .Close
    Set CRL = Nothing
    End With
End Function
Public Sub PreviewCCR(pRefnum As Long)
Dim CRL As Recordset
Dim ADR As Recordset
Dim DETTl As Recordset
VE.getCCRList (pRefnum)
Set CRL = VE.rsgetCCRList
VE.getAdramt (pRefnum)
Set ADR = VE.rsgetAdramt
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
    prtDate = ""
    prtTime = ""
    If ADR.Fields("adramt") <> 0 Or Not IsNull(ADR.Fields("adramt")) Then
        AdrAmt = ADR.Fields("adramt")
    End If
    If ADR.Fields("chkamt1") <> 0 Or Not IsNull(ADR.Fields("chkamt1")) Then
        ChkAmt1 = ADR.Fields("chkamt1")
        ChkTotal = ChkTotal + ChkAmt1
    End If
    If ADR.Fields("chkamt2") <> 0 Or Not IsNull(ADR.Fields("chkamt2")) Then
        ChkAmt2 = ADR.Fields("chkamt2")
        ChkTotal = ChkTotal + ChkAmt2
    End If
    If ADR.Fields("chkamt3") <> 0 Or Not IsNull(ADR.Fields("chkamt3")) Then
        ChkAmt3 = ADR.Fields("chkamt3")
        ChkTotal = ChkTotal + ChkAmt3
    End If
    If ADR.Fields("chkamt4") <> 0 Or Not IsNull(ADR.Fields("chkamt4")) Then
        ChkAmt4 = ADR.Fields("chkamt4")
        ChkTotal = ChkTotal + ChkAmt4
    End If
    If ADR.Fields("chkamt5") <> 0 Or Not IsNull(ADR.Fields("chkamt5")) Then
        ChkAmt5 = ADR.Fields("chkamt5")
        ChkTotal = ChkTotal + ChkAmt5
    End If
    If ADR.Fields("cshamt") <> 0 Or Not IsNull(ADR.Fields("cshamt")) Then
        CashAmt = ADR.Fields("cshamt")
    End If
ADR.Close
Set ADR = Nothing
With CRL
    Do While Not .EOF
        lngRcount = .RecordCount
        DetailTl = 0
        DetailAmt = 0
        TotalAmt = 0
        sngTempAmt = 0
        strCshAmt = ".00"
        strChqAmt = ".00"
        strAdrAmt = ".00"
        prtCash = ""
        prtAdr = ""
        VE.getTotal .Fields("refnum"), .Fields("seqnum")
        Set DETTl = VE.rsgetTotal
        DetailTl = DETTl.Fields("TotalAmt")
        DetailAmt = DETTl.Fields("TotalAmt")
        TotalAmt = DETTl.Fields("totalamt")
'   **  Checks Liquidation of ADR, Cheque and Cash
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
            GoTo NextTag
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
        If ChkAmt1 <> 0 Then
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
        End If
        If sngTempAmt = TotalAmt Then
            GoTo NextTag
        End If
'   ** Cheque 2
        If ChkAmt2 <> 0 Then
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
        End If
        If sngTempAmt = TotalAmt Then
            GoTo NextTag
        End If
'   ** Cheque 3
        If ChkAmt3 <> 0 Then
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
        End If
        If sngTempAmt = TotalAmt Then
            GoTo NextTag
        End If
'   ** Cheque 4
        If ChkAmt4 <> 0 Then
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
        End If
        If sngTempAmt = TotalAmt Then
            GoTo NextTag
        End If
'   ** Cheque 5
        If ChkAmt5 <> 0 Then
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
        End If
        If sngTempAmt = TotalAmt Then
            GoTo NextTag
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
NextTag:
        RefNum = .Fields("refnum")
        SeqNum = .Fields("seqnum")
        Customer = .Fields("cusnam")
        If mvarCCRNumber <> 0 Then
            If .Fields("ccrnum") = mvarCCRNumber Then
                frmCCRPr03.Show vbModal
                Exit Sub
            End If
        End If
        .MoveNext
    DETTl.Close
    
    Set DETTl = Nothing
    Loop
.Close
Set CRL = Nothing
End With
End Sub
Public Property Let PrinterReference(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrinterReference = 5
    mvarPrinterReference = vData
    Set Printer = Printers(vData)
End Property
Public Property Get PrinterReference() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrinterReference
    PrinterReference = mvarPrinterReference
End Property
Public Property Let CCRNumber(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CCRNumber = 5
    mvarCCRNumber = vData
End Property
Public Property Set CCRNumber(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.CCRNumber = Form1
    Set mvarCCRNumber = vData
End Property
Public Property Get CCRNumber() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CCRNumber
    If IsObject(mvarCCRNumber) Then
        Set CCRNumber = mvarCCRNumber
    Else
        CCRNumber = mvarCCRNumber
    End If
End Property
Private Sub OutCCR(pRefnum As Long, pSeqnum As Long, pCustomer As String, pAdrAmt As String, pCashAmt As String, _
                                 pChqAmt As String, pChkno1 As Boolean, pChkno2 As Boolean, pChkno3 As Boolean, pChkno4 As Boolean, _
                                 pChkno5 As Boolean)
    Dim x As Printer
    Dim PrinterName As String
    Dim f As Integer
    Dim CCR As rptCCRPr03
    Set CCR = New rptCCRPr03
    CCR.EnableParameterPrompting = False
    CCR.ParameterFields(1).AddCurrentValue (pRefnum)
    CCR.ParameterFields(2).AddCurrentValue (pSeqnum)
    CCR.ParameterFields(3).AddCurrentValue (pCustomer)
    CCR.ParameterFields(4).AddCurrentValue (pAdrAmt)
    CCR.ParameterFields(5).AddCurrentValue (pCashAmt)
    CCR.ParameterFields(6).AddCurrentValue (pChqAmt)
    CCR.ParameterFields(7).AddCurrentValue (pChkno1)
    CCR.ParameterFields(8).AddCurrentValue (pChkno2)
    CCR.ParameterFields(9).AddCurrentValue (pChkno3)
    CCR.ParameterFields(10).AddCurrentValue (pChkno4)
    CCR.ParameterFields(11).AddCurrentValue (pChkno5)
    CCR.ParameterFields(12).AddCurrentValue (strSupervisor)
'    CCR.TxtSupervisor.SetText strSupervisor
    CCR.DisplayProgressDialog = False
    CCR.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    CCR.PaperOrientation = crPortrait
    CCR.PrintOut False, 1
End Sub

Public Sub CCRSupervisor(ByVal vData As String)
     strSupervisor = vData
End Sub


