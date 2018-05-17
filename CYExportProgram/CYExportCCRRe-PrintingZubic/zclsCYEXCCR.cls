VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zclsCYEXCCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Dim StrChk1 As String * 30
Dim StrChk2 As String * 30
Dim StrChk3 As String * 30
Dim StrChk4 As String * 30
Dim StrChk5 As String * 30
'local variable(s) to hold property value(s)
Private mvarRefNumber As Long 'local copy
Private mvarCCRNumber As Long 'local copy
'local variable(s) to hold property value(s)
Private mvarPrinterReference As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarUserid As String 'local copy
Public Property Let Userid(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Userid = 5
    mvarUserid = vData
    gUserid = UCase(vData)
End Property
Public Property Get Userid() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Userid
    Userid = UCase(mvarUserid)
End Property
Public Sub Execute()
    frmPrinter.Show vbModal
    frmREPRINT.Show vbModal
End Sub
Public Property Let PrinterReference(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PrinterReference = 5
    mvarPrinterReference = vData
End Property
Public Property Set PrinterReference(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.PrinterReference = Form1
    Set mvarPrinterReference = vData
End Property
Public Property Get PrinterReference() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PrinterReference
    If IsObject(mvarPrinterReference) Then
        Set PrinterReference = mvarPrinterReference
    Else
        PrinterReference = mvarPrinterReference
    End If
End Property
Public Property Let CCRNumber(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CCRNumber = 5
    mvarCCRNumber = vData
End Property
Public Property Get CCRNumber() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CCRNumber
    CCRNumber = mvarCCRNumber
End Property
Public Property Let RefNumber(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RefNumber = 5
    mvarRefNumber = vData
End Property
Public Property Get RefNumber() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RefNumber
    RefNumber = mvarRefNumber
End Property
Public Sub PreviewCCR(pRefnum As Long)
Dim CRL As Recordset
Dim ADR As Recordset
Dim DETTl As Recordset
DE.getCCRList (pRefnum)
Set CRL = DE.rsgetCCRList
DE.getAdrAmt (pRefnum)
Set ADR = DE.rsgetAdrAmt
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
        DE.getTotal .Fields("refnum"), .Fields("seqnum")
        Set DETTl = DE.rsgetTotal
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
    '   **  Preview Gatepass
NextTag:
        RefNum = .Fields("refnum")
        SeqNum = .Fields("seqnum")
        Customer = .Fields("cusnam") & ""
'        frmCYEXCCR.Show vbModal
        .MoveNext
    DETTl.Close
    Set DETTl = Nothing
    Loop
.Close
Set CRL = Nothing
End With
End Sub
Public Function PrintCCR(pRefnum As Long) As Boolean
    Dim CRL As Recordset
    Dim ADR As Recordset
    Dim DETTl As Recordset
    DE.getCCRList (pRefnum)
    Set CRL = DE.rsgetCCRList
    DE.getAdrAmt (pRefnum)
    Set ADR = DE.rsgetAdrAmt
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
    
    
    If ADR.Fields("adramt") <> 0 Or Not IsNull(ADR.Fields("adramt")) Then
        AdrAmt = ADR.Fields("adramt")
    End If
    If ADR.Fields("chkamt1") <> 0 Or Not IsNull(ADR.Fields("chkamt1")) Then
        ChkAmt1 = ADR.Fields("chkamt1")
        ChkTotal = ChkTotal + ChkAmt1
        StrChk1 = ADR.Fields("chkno1") & " " & Format(ADR.Fields("chkamt1"), "#########.#0") & " " & ADR.Fields("chkbnk1")
    End If
    If ADR.Fields("chkamt2") <> 0 Or Not IsNull(ADR.Fields("chkamt2")) Then
        ChkAmt2 = ADR.Fields("chkamt2")
        ChkTotal = ChkTotal + ChkAmt2
        StrChk2 = ADR.Fields("chkno2") & " " & Format(ADR.Fields("chkamt2"), "#########.#0") & " " & ADR.Fields("chkbnk2")
    End If
    If ADR.Fields("chkamt3") <> 0 Or Not IsNull(ADR.Fields("chkamt3")) Then
        ChkAmt3 = ADR.Fields("chkamt3")
        ChkTotal = ChkTotal + ChkAmt3
        StrChk3 = ADR.Fields("chkno3") & " " & Format(ADR.Fields("chkamt3"), "#########.#0") & " " & ADR.Fields("chkbnk3")
    End If
    If ADR.Fields("chkamt4") <> 0 Or Not IsNull(ADR.Fields("chkamt4")) Then
        ChkAmt4 = ADR.Fields("chkamt4")
        ChkTotal = ChkTotal + ChkAmt4
        StrChk4 = ADR.Fields("chkno4") & " " & Format(ADR.Fields("chkamt4"), "#########.#0") & " " & ADR.Fields("chkbnk4")
    End If
    If ADR.Fields("chkamt5") <> 0 Or Not IsNull(ADR.Fields("chkamt5")) Then
        ChkAmt5 = ADR.Fields("chkamt5")
        ChkTotal = ChkTotal + ChkAmt5
        StrChk5 = ADR.Fields("chkno5") & " " & Format(ADR.Fields("chkamt5"), "#########.#0") & " " & ADR.Fields("chkbnk5")
    End If
    If ADR.Fields("cshamt") <> 0 Or Not IsNull(ADR.Fields("cshamt")) Then
        CashAmt = ADR.Fields("cshamt")
    End If
    ADR.Close
    Set ADR = Nothing
    
    With CRL
        Do While Not .EOF
            If .Fields("ccrmod") = "D" Then
                DomesticMode = True
            Else
                DomesticMode = False
            End If
            DetailTl = 0
            DetailAmt = 0
            TotalAmt = 0
            sngTempAmt = 0
            
            strCshAmt = ".00"
            strChqAmt = ".00"
            strAdrAmt = ".00"
            
            blnChkno1 = False
            blnChkno2 = False
            blnChkno3 = False
            blnChkno4 = False
            blnChkno5 = False
            DE.getTotal .Fields("refnum"), .Fields("seqnum")
            Set DETTl = DE.rsgetTotal
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
                If sngTempAmt = TotalAmt Then
                    GoTo NextCCRTag
                End If
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
                If DetailAmt > ChkAmt1 And DetailAmt <> 0 Then
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
                If sngTempAmt = TotalAmt Then
                    GoTo NextCCRTag
                End If
            End If
            '   ** Cheque 2
            If ChkAmt2 <> 0 And DetailAmt <> 0 Then
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
                If sngTempAmt = TotalAmt Then
                    GoTo NextCCRTag
                End If
            End If
            '   ** Cheque 3
            If ChkAmt3 <> 0 And DetailAmt <> 0 Then
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
                If sngTempAmt = TotalAmt Then
                    GoTo NextCCRTag
                End If
            End If
            '   ** Cheque 4
            If ChkAmt4 <> 0 And DetailAmt <> 0 Then
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
                If sngTempAmt = TotalAmt Then
                    GoTo NextCCRTag
                End If
            End If
            '   ** Cheque 5
            If ChkAmt5 <> 0 And ChkAmt5 <> 0 Then
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
                If sngTempAmt = TotalAmt Then
                    GoTo NextCCRTag
                End If
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
                    Call OutCCRPC(.Fields("refnum"), .Fields("seqnum"), .Fields("cusnam"), strAdrAmt, strCshAmt, strChqAmt, _
                                blnChkno1, blnChkno2, blnChkno3, blnChkno4, blnChkno5)
                End If
            Else
                Call OutCCRPC(.Fields("refnum"), .Fields("seqnum"), .Fields("cusnam") & "", strAdrAmt, strCshAmt, strChqAmt, _
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
Private Sub OutCCRPC(pRefnum As Long, pSeqnum As Long, pCustomer As String, pAdrAmt As String, pCashAmt As String, _
                                 pChqAmt As String, pChkno1 As Boolean, pChkno2 As Boolean, pChkno3 As Boolean, pChkno4 As Boolean, _
                                 pChkno5 As Boolean)
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
Dim strSize As String * 3
Dim strCtnnum As String * 12
Dim strArrastre As String * 12
Dim strWArrastre As String * 12
Dim strWharfage As String * 12
Dim strTArrastre As String * 12
Dim strTWArrastre As String * 12
Dim strTWharfage As String * 12
Dim sngArrastre As Currency
Dim sngWArrastre  As Currency
Dim sngWharfage As Currency
Dim sngTArrastre As Single
Dim sngTWArrastre As Single
Dim sngTWharfage As Single
Dim sngVat As Single
Dim sngWtx As Single
Dim vslName As String * 10
Dim x As Integer
Dim strRemarks As String * 15
Dim strRemarks30 As String * 36
Dim strRemarkOut As String * 16
Dim CD As Recordset
Dim WhfRteAmt As Currency
Dim remark1 As String
Dim remark2  As String
Dim remark3 As String
Dim UserName As String
Dim strValidation As String * 35
Dim RevTonnage As String * 16
WhfRteAmt = 0
ctrCnt = 11
DE.CCRDetails pRefnum, pSeqnum
Set CD = DE.rsCCRDetails
If CD.RecordCount > 0 Then
    With CD
        Select Case .Fields("vatcde")
        Case "0"
            remark1 = "0Vat"
        Case "2"
            remark1 = "1%Less"
        Case "3"
            remark1 = "6%Whd"
        Case Else
            remark1 = " "
        End Select
        Select Case .Fields("whfcde")
        Case "1"
            remark2 = "BOI"
        Case "2"
            remark2 = "PEZA"
        Case "3"
            remark2 = "Napocor"
        Case "4"
            remark2 = "WhfPd."
        Case "5"
            remark2 = "PPC"
        Case "6"
            remark2 = "ShutOut"
        Case Else
            remark2 = " "
        End Select
        If .Fields("guarntycde") = "Y" Then
            remark3 = "U/G"
        Else
            remark3 = " "
        End If
        'UserName = Trim(UCase(.Fields("userid") & ""))
        UserName = Trim(UCase(.Fields("userid") & "")) & Space(8) & Trim(UCase(.Fields("supvsr") & ""))
        Refn = Format(.Fields("refnum"), "000000")
        Seqf = Trim(Format(.Fields("seqnum"), "0000"))
        CCRf = Format(.Fields("ccrnum"), "000000")
        DateTime = Format(.Fields("sysdttm"), "     YYYY-MM-DD hh:nn")
        strExporter = Mid(.Fields("broker"), 1, 15) & "/" & Mid(.Fields("exprtr"), 1, 15)
        strRemarks = Mid(.Fields("remark") & "", 1, 15)
        strRemarkOut = Trim(Mid(remark1, 1, 6)) & Trim(Mid(remark2, 1, 7)) & Trim(Mid(remark3, 1, 3))
        strEntry = .Fields("entnum") & ""
        strValidation = Trim(Refn) & " " & Trim(Seqf) & " " & Trim(CCRf) & " " & Format(.Fields("sysdttm"), "YY-MM-DD hh:nn")
        vslName = .Fields("vslcde") & ""
        Printer.Font = "Courier 12cpi"
        Printer.FontSize = 10
        Printer.Print " "
        'Printer.Print Space(23) & "REF " & Refn & " SEQ " & Seqf & Space(2) & CCRf
        'Printer.Print " "
        'Printer.Print " "
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
        Printer.Print Space(2) & .Fields("commod")
        If CSng(.Fields("dolrte")) > 0 Then
            WhfRteAmt = CSng(.Fields("dolrte"))
        End If
        sngVat = 0
        sngWtx = 0
        Do While Not .EOF
            If .Fields("vatcde") <> "1" Then
                sngVat = sngVat + CSng(.Fields("arrvat"))
            End If
            If .Fields("vatcde") = "2" Or .Fields("vatcde") = "4" Or .Fields("vatcde") = "6" Then
                sngWtx = sngWtx + CSng(.Fields("arrtax"))
            End If
            sngArrastre = CSng(.Fields("arramt")) + CSng(.Fields("ovzamt")) + CSng(.Fields("dgramt")) + CSng(.Fields("arrvat")) - CSng(.Fields("arrtax"))
            sngWArrastre = CSng(.Fields("arramt")) + CSng(.Fields("ovzamt")) + CSng(.Fields("dgramt"))
            sngTArrastre = sngTArrastre + sngArrastre
            sngTWArrastre = sngTWArrastre + sngWArrastre
            If .Fields("whfcde") = "0" Then
                sngWharfage = CSng(.Fields("whfamt"))
            Else
                sngWharfage = 0
            End If
            sngTWharfage = sngTWharfage + sngWharfage
            If sngArrastre > 0 Then
                strArrastre = Format(sngArrastre, "###,###,###.#0")
            Else
                strArrastre = " "
            End If
            If sngWArrastre > 0 Then
                strWArrastre = Format(sngWArrastre, "###,###,###.#0")
            Else
                strWArrastre = " "
            End If
            If sngWharfage > 0 Then
                strWharfage = Format(sngWharfage, "###,###,###.#0")
            Else
                strWharfage = " "
            End If
            strSize = .Fields("cntsze")
            strCtnnum = .Fields("cntnum")
            If CSng(.Fields("ovzamt")) > 0 Then
                RevTonnage = Format(CSng(.Fields("revton")), "###,###,###.#0")
            Else
                RevTonnage = ""
            End If
            Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & RevTonnage & Space(2) & strArrastre & Space(29) & strArrastre
            ctrCnt = ctrCnt - 1
            .MoveNext
        Loop
        If ctrCnt > 0 Then
            For x = 1 To ctrCnt
                Printer.Print " "
            Next
        End If
        If sngTArrastre > 0 Then
            strTArrastre = Format(sngTArrastre, "###,###,###.#0")
        Else
            strTArrastre = " "
        End If
        If sngTWArrastre > 0 Then
            strTWArrastre = Format(sngTWArrastre, "###,###,###.#0")
        Else
            strTWArrastre = " "
        End If
        If sngTWharfage > 0 Then
            strTWharfage = Format(sngTWharfage, "###,###,###.#0")
        Else
            strTWharfage = " "
        End If
        If sngVat > 0 Then
            If sngWtx > 0 Then
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
        Printer.Print Space(6) & Space(15) & Space(17) & Space(17) & Trim(tmpString) & Space(10) & strTArrastre
        tmpString = NumToText(CCur(sngTArrastre))
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
        
        If DomesticMode Then
            Printer.Print "  DOMESTIC" & Space(36) & Word3
        Else
            Printer.Print "  FOREIGN " & Space(36) & Word3
        End If
        Printer.Print " "
        tmpString = strChqAmt & " CK    " & strCshAmt & " CS"
        Printer.Print Space(44) & tmpString
        tmpString = strAdrAmt & " AD"
        Printer.Print Space(5) & UserName  ' & Space(26) & tmpString
        If blnChkno1 Then
            tmp1 = StrChk1
        Else
            tmp1 = " "
        End If
        If blnChkno2 Then
            If Len(tmp1) > 0 Then
                tmp2 = ", " & StrChk2
            Else
                tmp2 = " " & StrChk2
            End If
        Else
            tmp2 = " "
        End If
        Printer.Print Space(44) & Trim(tmp1) & tmp2
        If blnChkno3 Then
            tmp1 = StrChk3
        Else
            tmp1 = " "
        End If
        If blnChkno4 Then
            If Len(tmp1) > 0 Then
                tmp2 = ", " & StrChk4
            Else
                tmp2 = " " & StrChk4
            End If
        Else
            tmp2 = " "
        End If
        Printer.Print Space(44) & Trim(tmp1) & tmp2
        If blnChkno5 Then
            tmp1 = StrChk5
        Else
            tmp1 = " "
        End If
        Printer.Print Space(44) & tmp1
        Printer.Print Space(44) & strValidation
        Printer.Print ""
        Printer.Print ""
        Printer.Print Space(5) & "REF " & Refn & " SEQ " & Seqf & Space(2) & CCRf
        Printer.FontSize = 10
        Printer.EndDoc
        
'        Printer.Print " "
'        Printer.Print Space(23) & "REF " & Refn & " SEQ " & Seqf & Space(2) & CCRf
'        Printer.Print " "  'N
'        Printer.Print Space(72) & DateTime
'        Printer.Print " "
'        Printer.Print " "
'        Printer.Print Space(2) & strExporter & Space(13) & vslName & Space(3) & _
'                        Trim(Mid(strEntry, 1, 8) & " " & _
'                        Mid(strEntry, 9, 8) & " " & _
'                        Mid(strEntry, 17, 8) & " " & _
'                        Mid(strEntry, 25, 8) & " " & _
'                        Mid(strEntry, 33, 8))
'        Printer.Print Space(45) & Trim(Mid(strEntry, 41, 8) & " " & _
'                        Mid(strEntry, 49, 8) & " " & _
'                        Mid(strEntry, 57, 8) & " " & _
'                        Mid(strEntry, 65, 8) & " " & _
'                        Mid(strEntry, 73, 8))
'        Printer.Print " "
'        Printer.Print " "
'        Printer.Print " "
'        Printer.Print Space(2) & .Fields("commod")
'        If CSng(.Fields("dolrte")) > 0 Then
'            WhfRteAmt = CSng(.Fields("dolrte"))
'        End If
'        sngVat = 0
'        sngWtx = 0
'        Do While Not .EOF
'            If .Fields("vatcde") <> "0" Then
'                sngVat = sngVat + CSng(.Fields("arrvat"))
'            End If
'            If .Fields("vatcde") = "2" Or .Fields("vatcde") = "3" Then
'                sngWtx = sngWtx + CSng(.Fields("arrtax"))
'            End If
'            sngArrastre = CSng(.Fields("arramt")) + CSng(.Fields("ovzamt")) + CSng(.Fields("dgramt")) + CSng(.Fields("arrvat")) - CSng(.Fields("arrtax"))
'            sngWArrastre = CSng(.Fields("arramt")) + CSng(.Fields("ovzamt")) + CSng(.Fields("dgramt"))
'            sngTArrastre = sngTArrastre + sngArrastre
'            sngTWArrastre = sngTWArrastre + sngWArrastre
'            If .Fields("whfcde") = "0" Then
'                sngWharfage = CSng(.Fields("whfamt"))
'            Else
'                sngWharfage = 0
'            End If
'            sngTWharfage = sngTWharfage + sngWharfage
'            If sngArrastre > 0 Then
'                strArrastre = Format(sngArrastre, "###,###,###.#0")
'            Else
'                strArrastre = " "
'            End If
'            If sngWArrastre > 0 Then
'                strWArrastre = Format(sngWArrastre, "###,###,###.#0")
'            Else
'                strWArrastre = " "
'            End If
'            If sngWharfage > 0 Then
'                strWharfage = Format(sngWharfage, "###,###,###.#0")
'            Else
'                strWharfage = " "
'            End If
'            strSize = .Fields("cntsze")
'            strCtnnum = .Fields("cntnum")
'            If CSng(.Fields("ovzamt")) > 0 Then
'                RevTonnage = Format(CSng(.Fields("revton")), "###,###,###.#0")
'            Else
'                RevTonnage = ""
'            End If
'            Printer.Print Space(2) & strSize & Space(1) & strCtnnum & Space(2) & RevTonnage & Space(2) & strArrastre & Space(29) & strArrastre
'            ctrCnt = ctrCnt - 1
'            .MoveNext
'        Loop
'        If ctrCnt > 0 Then
'            For x = 1 To ctrCnt
'                Printer.Print " "
'            Next
'        End If
'        If sngTArrastre > 0 Then
'            strTArrastre = Format(sngTArrastre, "###,###,###.#0")
'        Else
'            strTArrastre = " "
'        End If
'        If sngTWArrastre > 0 Then
'            strTWArrastre = Format(sngTWArrastre, "###,###,###.#0")
'        Else
'            strTWArrastre = " "
'        End If
'        If sngTWharfage > 0 Then
'            strTWharfage = Format(sngTWharfage, "###,###,###.#0")
'        Else
'            strTWharfage = " "
'        End If
'        If sngVat > 0 Then
'            If sngWtx > 0 Then
'                tmpString = "VAT INCLUSIVE LESS W/TAX"
'            Else
'                tmpString = "VAT INCLUSIVE"
'            End If
'        Else
'            tmpString = "ZERO RATED VAT"
'        End If
'        Printer.Print " " 'N
'        Printer.Print " " 'N
'        Printer.Print " " 'N
'        Printer.Print Space(6) & Space(15) & Space(17) & Space(17) & Trim(tmpString) & Space(10) & strTArrastre
'        'Printer.Print ""
'        tmpString = NumToText(CCur(sngTArrastre))
'        Word1 = Mid(tmpString, 1, 35)
'        If Len(Trim(Mid(tmpString, 35, 1))) <> 0 And Len(Trim(Mid(tmpString, 36, 1))) <> 0 Then
'            Word1 = Trim(Word1) & "-"
'        End If
'        Word2 = Mid(tmpString, 36, 35)
'        If Len(Trim(Mid(tmpString, 71, 1))) <> 0 And Len(Trim(Mid(tmpString, 72, 1))) <> 0 Then
'            Word2 = Trim(Word2) & "-"
'        End If
'        Word3 = Mid(tmpString, 72, 35)
'        Printer.Print " " 'N
'        Printer.Print Space(46) & Word1
'        Printer.Print Space(2) & strRemarks & Space(6) & strRemarkOut & Space(7) & Word2
'
'        If DomesticMode Then
'            Printer.Print "  DOMESTIC" & Space(36) & Word3
'        Else
'            Printer.Print "  FOREIGN " & Space(36) & Word3
'        End If
'        Printer.Print " " 'N
'        tmpString = strChqAmt & " CK    " & strCshAmt & " CS"
'        Printer.Print Space(44) & tmpString
'        tmpString = strAdrAmt & " AD"
'        Printer.Print Space(5) & UserName  ' & Space(26) & tmpString N
'        'Printer.Print "" N
'        If blnChkno1 Then
'            tmp1 = StrChk1
'        Else
'            tmp1 = " "
'        End If
'        If blnChkno2 Then
'            If Len(tmp1) > 0 Then
'                tmp2 = ", " & StrChk2
'            Else
'                tmp2 = " " & StrChk2
'            End If
'        Else
'            tmp2 = " "
'        End If
'        Printer.Print Space(44) & Trim(tmp1) & tmp2
'        If blnChkno3 Then
'            tmp1 = StrChk3
'        Else
'            tmp1 = " "
'        End If
'        If blnChkno4 Then
'            If Len(tmp1) > 0 Then
'                tmp2 = ", " & StrChk4
'            Else
'                tmp2 = " " & StrChk4
'            End If
'        Else
'            tmp2 = " "
'        End If
'        Printer.Print Space(44) & Trim(tmp1) & tmp2
'        If blnChkno5 Then
'            tmp1 = StrChk5
'        Else
'            tmp1 = " "
'        End If
'        Printer.Print Space(44) & tmp1
'        Printer.Print Space(44) & strValidation
'        Printer.FontSize = 10
'        Printer.EndDoc
    End With
End If
CD.Close
Set CD = Nothing
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
    strResult = "and " & Format((dblValue - Int(dblValue)) * 100, "00") & "/100"
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



