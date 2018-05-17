VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmReprint 
   Caption         =   "Reprint Gatepass"
   ClientHeight    =   9405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rptCYMPR01 
      Left            =   14640
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
      WindowShowProgressCtls=   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      Begin VB.CommandButton cmdReprint 
         Caption         =   "&Reprint"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   4680
         TabIndex        =   1
         Top             =   840
         Width           =   1575
      End
      Begin MSMask.MaskEdBox mskReference 
         Height          =   405
         Left            =   2640
         TabIndex        =   2
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   714
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "########"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskSequence 
         Height          =   405
         Left            =   2640
         TabIndex        =   3
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   393216
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
         Format          =   "##"
         PromptChar      =   " "
      End
      Begin VB.Label lblMain 
         Alignment       =   1  'Right Justify
         Caption         =   "Reference:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblMain 
         Alignment       =   1  'Right Justify
         Caption         =   "Sequence:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmReprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const cNullDate As Date = #12:00:00 AM#
Dim rstCYMPay As ADODB.Recordset
Dim rstCYMGps As ADODB.Recordset

Private Type Payments
    POS As Currency
    ADR As Currency
    CheckAmt1 As Currency
    CheckAmt2 As Currency
    CheckAmt3 As Currency
    CheckAmt4 As Currency
    CheckAmt5 As Currency
    CheckNo1 As String * 10
    CheckNo2 As String * 10
    CheckNo3 As String * 10
    CheckNo4 As String * 10
    CheckNo5 As String * 10
    CheckBnk1 As String * 10
    CheckBnk2 As String * 10
    CheckBnk3 As String * 10
    CheckBnk4 As String * 10
    CheckBnk5 As String * 10
    Cash As Currency
    Change As Currency
    TotalPayment As Currency
    RemainingPayment As Currency
    Customer As String * 30
End Type

Private Type Details
    Arrastre As Currency
    ArrastreVAT As Currency
    ArrastreTAX As Currency
    Storage As Currency
    StorageVAT As Currency
    StorageTAX As Currency
    Weighing As Currency
    WeighingVAT As Currency
    WeighingTAX As Currency
    Reefer As Currency
    ReeferVAT As Currency
    ReeferTAX As Currency
    Wharfage As Currency
    UnderGuarantee As String * 1
    ArrastreNet As Currency
    StorageNet As Currency
    WeighingNet As Currency
    ReeferNet As Currency
    TotalCharge As Currency
    TotalNet As Currency
    DueICTSI As Currency
    DueICTSIWords As Currency
    Gatepass As Long
    Reference As Long
    Sequence As Integer
    SysDate As Date
    Consignee As String * 30
    Broker As String * 30
    Registry As String * 12
    Location As String * 20
    VoyageNo As String * 20
    SMBAPermitNo As String * 20
    CustomPermitNo As String * 20
    EntryNo As Long
    BillNum As String * 22
    VesselCode As String * 7
    PortofOrig As String * 15
    DeclaredWeight As String * 15
    PDIGNo As String * 15
    ContainerNo As String * 22
    ContainerSize As Integer
    LastDischarge As Date
    ShippingLine As String * 7
    OrderSupplier As String * 8
    SealNumber As String * 8
    FullEmp As String * 1
    CRODate As Date
    FreeUntil As Date
    StorageEnd As Date
    StorageDay As Integer
    PlugIn As Date
    PlugOut As Date
    RevenueTon As Currency
    Discount As Currency
    StorageAMT As Currency
    DiscountAMT As Currency
    Oversize As Currency
    Commodity As String * 30
    UserID As String * 10
    ForExam As String * 1
    Remark As String * 30
    CustomsGuard As String * 1
    ConsCode As String * 1
    VATCode As String * 1
    ForWeighing As String * 1
    DangerClass As String * 1
    BoatNote As String * 8
    strRevenueTon As String * 7
    strArrastreLessOversize As String * 10
    strStorageAmt As String * 10
    strWeighingNet As String * 7
    strReeferNet As String * 10
    strTotalNet As String * 11
    strWharfage As String * 7
    strTotalCharge As String * 11
    strOversize As String * 10
    strDiscount As String * 10
    strToText As String * 50
End Type

Dim Detail As Details
Dim Payment As Payments
Dim intResponse As Integer

Dim curArrastre As Currency
Dim curArrastreVat As Currency
Dim curArrastreNoVat As Currency
Dim curArrastreTax As Currency
Dim curStorage As Currency
Dim curStorageVat As Currency
Dim curStorageNoVat As Currency
Dim curStorageTax As Currency
Dim curReefer As Currency
Dim curReeferVat As Currency
Dim curReeferNoVat As Currency
Dim curReeferTax As Currency
Dim curWeighing As Currency
Dim curWeighingVat As Currency
Dim curWeighingNoVat As Currency
Dim curWeighingTax As Currency
Dim curADRAmount As Currency
Dim curADRArrastre As Currency
Dim curADRStorage As Currency
Dim curADRWeighing As Currency
Dim curADRReefer As Currency

Dim dtmDateFrom As Date
Dim dtmDateTo As Date

Private Sub cmdReprint_Click()
    Call GetTotalPaymentAmounts
    Call GetTotalChargePerDetail
End Sub

Private Function LiquidatePaymentTypes(pType As Integer) As String
    Dim curPOSApplied As Currency
    Dim curADRApplied As Currency
    Dim curCheck1Applied As Currency
    Dim curCheck2Applied As Currency
    Dim curCheck3Applied As Currency
    Dim curCheck4Applied As Currency
    Dim curCheck5Applied As Currency
    Dim curCashApplied As Currency
    
    LiquidatePaymentTypes = ""
    With Payment
        Select Case pType
            Case 1
                'ADR
                If Detail.DueICTSI > 0 Then
                    If .ADR > 0 Then
                        If .ADR >= Detail.DueICTSI Then
                            curADRApplied = Detail.DueICTSI
                            .ADR = .ADR - Detail.DueICTSI
                        ElseIf .ADR < Detail.DueICTSI Then
                            curADRApplied = .ADR
                            .ADR = 0
                        End If
                        Detail.DueICTSI = Detail.DueICTSI - curADRApplied
                        LiquidatePaymentTypes = CStr(Format(curADRApplied, "####,##0.00"))
                        LiquidatePaymentTypes = Left(LiquidatePaymentTypes & Space(11), 11)
                        Exit Function
                    End If
                End If
            Case 2
                'CHECK1
                If Detail.DueICTSI > 0 Then
                    If .CheckAmt1 > 0 Then
                        If .CheckAmt1 >= Detail.DueICTSI Then
                            curCheck1Applied = Detail.DueICTSI
                            .CheckAmt1 = .CheckAmt1 - Detail.DueICTSI
                        ElseIf .CheckAmt1 < Detail.DueICTSI Then
                            curCheck1Applied = .CheckAmt1
                            .CheckAmt1 = 0
                        End If
                        Detail.DueICTSI = Detail.DueICTSI - curCheck1Applied
                        LiquidatePaymentTypes = CStr(Format(curCheck1Applied, "####,##0.00"))
                        LiquidatePaymentTypes = Left(LiquidatePaymentTypes & Space(11), 11)
                        Exit Function
                    End If
                End If
            Case 3
                'CHECK2
                If Detail.DueICTSI > 0 Then
                    If .CheckAmt2 > 0 Then
                        If .CheckAmt2 >= Detail.DueICTSI Then
                            curCheck2Applied = Detail.DueICTSI
                            .CheckAmt2 = .CheckAmt2 - Detail.DueICTSI
                        ElseIf .CheckAmt2 < Detail.DueICTSI Then
                            curCheck2Applied = .CheckAmt2
                            .CheckAmt2 = 0
                        End If
                        Detail.DueICTSI = Detail.DueICTSI - curCheck2Applied
                        LiquidatePaymentTypes = CStr(Format(curCheck2Applied, "####,##0.00"))
                        LiquidatePaymentTypes = Left(LiquidatePaymentTypes & Space(11), 11)
                        Exit Function
                    End If
                End If
            Case 4
                'CHECK3
                If Detail.DueICTSI > 0 Then
                    If .CheckAmt3 > 0 Then
                        If .CheckAmt3 >= Detail.DueICTSI Then
                            curCheck3Applied = Detail.DueICTSI
                            .CheckAmt3 = .CheckAmt3 - Detail.DueICTSI
                        ElseIf .CheckAmt3 < Detail.DueICTSI Then
                            curCheck3Applied = .CheckAmt3
                            .CheckAmt3 = 0
                        End If
                        Detail.DueICTSI = Detail.DueICTSI - curCheck3Applied
                        LiquidatePaymentTypes = CStr(Format(curCheck3Applied, "####,##0.00"))
                        LiquidatePaymentTypes = Left(LiquidatePaymentTypes & Space(11), 11)
                        Exit Function
                    End If
                End If
            Case 5
                'CHECK4
                If Detail.DueICTSI > 0 Then
                    If .CheckAmt4 > 0 Then
                        If .CheckAmt4 >= Detail.DueICTSI Then
                            curCheck4Applied = Detail.DueICTSI
                            .CheckAmt4 = .CheckAmt4 - Detail.DueICTSI
                        ElseIf .CheckAmt4 < Detail.DueICTSI Then
                            curCheck4Applied = .CheckAmt4
                            .CheckAmt4 = 0
                        End If
                        Detail.DueICTSI = Detail.DueICTSI - curCheck4Applied
                        LiquidatePaymentTypes = CStr(Format(curCheck4Applied, "####,##0.00"))
                        LiquidatePaymentTypes = Left(LiquidatePaymentTypes & Space(11), 11)
                        Exit Function
                    End If
                End If
            Case 6
                'CHECK5
                If Detail.DueICTSI > 0 Then
                    If .CheckAmt5 > 0 Then
                        If .CheckAmt5 >= Detail.DueICTSI Then
                            curCheck5Applied = Detail.DueICTSI
                            .CheckAmt5 = .CheckAmt5 - Detail.DueICTSI
                        ElseIf .CheckAmt5 < Detail.DueICTSI Then
                            curCheck5Applied = .CheckAmt5
                            .CheckAmt5 = 0
                        End If
                        Detail.DueICTSI = Detail.DueICTSI - curCheck5Applied
                        LiquidatePaymentTypes = CStr(Format(curCheck5Applied, "####,##0.00"))
                        LiquidatePaymentTypes = Left(LiquidatePaymentTypes & Space(11), 11)
                        Exit Function
                    End If
                End If
            Case 7
                'CASH
                If .RemainingPayment > 0 Then
                    If .Cash > 0 Then
                        If .Cash >= Detail.DueICTSI Then
                            curCashApplied = Detail.DueICTSI
                            .Cash = .Cash - Detail.DueICTSI
                        ElseIf .Cash < Detail.DueICTSI Then
                            curCashApplied = .Cash
                            .Cash = 0
                        End If
                         Detail.DueICTSI = Detail.DueICTSI - curCashApplied
                         LiquidatePaymentTypes = CStr(Format(curCashApplied, "####,##0.00"))
                         LiquidatePaymentTypes = Left(LiquidatePaymentTypes & Space(11), 11)
                         Exit Function
                    End If
                End If
          Case 8
                'POS
                If .RemainingPayment > 0 Then
                         LiquidatePaymentTypes = CStr(Format(.POS, "####,##0.00"))
                         LiquidatePaymentTypes = Left(LiquidatePaymentTypes & Space(11), 11)
                         Exit Function
                End If
        End Select
    End With
End Function

Private Sub GetTotalChargePerDetail()
    Dim curTotalCharge As Currency
    Set rstCYMGps = New ADODB.Recordset
    rstCYMGps.LockType = adLockOptimistic
    rstCYMGps.CursorType = adOpenStatic
    rstCYMGps.Open "Select * from CYMGps where refnum= " & CLng(mskReference) & " order by seqnum", gcnnBilling, , , adCmdText
    
    rstCYMGps.MoveFirst
    Do While Not rstCYMGps.EOF
        With Detail
            .Arrastre = rstCYMGps.Fields("arramt")
            .ArrastreVAT = rstCYMGps.Fields("arrvat")
            .ArrastreTAX = rstCYMGps.Fields("arrtax")
            .Storage = rstCYMGps.Fields("stoamt")
            .StorageVAT = rstCYMGps.Fields("stovat")
            .StorageTAX = rstCYMGps.Fields("stotax")
            .Weighing = rstCYMGps.Fields("wghamt")
            .WeighingVAT = rstCYMGps.Fields("wghvat")
            .WeighingTAX = rstCYMGps.Fields("wghtax")
            .Reefer = rstCYMGps.Fields("rframt")
            .ReeferVAT = rstCYMGps.Fields("rfrvat")
            .ReeferTAX = rstCYMGps.Fields("rfrtax")
            .Wharfage = rstCYMGps.Fields("whfamt")
            .UnderGuarantee = Trim(rstCYMGps.Fields("gtycde"))
            
            .ArrastreNet = .Arrastre + .ArrastreVAT - .ArrastreTAX
            .StorageNet = .Storage + .StorageVAT - .StorageTAX
            .WeighingNet = .Weighing + .WeighingVAT - .WeighingTAX
            .ReeferNet = .Reefer + .ReeferVAT - .ReeferTAX
            .TotalCharge = .ArrastreNet + .StorageNet + .WeighingNet + .ReeferNet + .Wharfage
            .TotalNet = .ArrastreNet + .StorageNet + .WeighingNet + .ReeferNet
            
            .Gatepass = rstCYMGps.Fields("gpsnum")
            .Reference = rstCYMGps.Fields("refnum")
            .Sequence = rstCYMGps.Fields("seqnum")
            .SysDate = rstCYMGps.Fields("sysdte")
            .Consignee = Left(rstCYMGps.Fields("cnsgne") & Space(30), 30)
            .Broker = Left(rstCYMGps.Fields("broker") & Space(30), 30)
            .Registry = Left(rstCYMGps.Fields("regnum") & Space(12), 12)
            .EntryNo = rstCYMGps.Fields("entnum")
            .Location = Left(rstCYMGps.Fields("location") & Space(10), 10)
            .VoyageNo = Left(rstCYMGps.Fields("voyageno") & Space(10), 10)
            .SMBAPermitNo = Left(rstCYMGps.Fields("sbmapn") & Space(10), 10)
            .CustomPermitNo = Left(rstCYMGps.Fields("custompn") & Space(10), 10)
            .BillNum = Left(rstCYMGps.Fields("bilnum") & Space(22), 22)
            .VesselCode = rstCYMGps.Fields("vslcde")
            .PortofOrig = Left(rstCYMGps.Fields("prtorg") & Space(15), 15)
            .DeclaredWeight = Left(rstCYMGps.Fields("dclwgt") & Space(15), 15)
            .PDIGNo = Left(rstCYMGps.Fields("pdigno") & Space(15), 15)
            .ContainerNo = Left(rstCYMGps.Fields("cntnum") & Space(12), 12)
            .ContainerSize = rstCYMGps.Fields("cntsze")
            .LastDischarge = rstCYMGps.Fields("lstdch")
            .ShippingLine = rstCYMGps.Fields("shplin")
            .OrderSupplier = rstCYMGps.Fields("ordsup")
            .SealNumber = rstCYMGps.Fields("silnum")
            .FullEmp = rstCYMGps.Fields("fulemp")
            .CRODate = rstCYMGps.Fields("crodte")
            .FreeUntil = rstCYMGps.Fields("freeuntil")
            .StorageEnd = rstCYMGps.Fields("stoend")
            .StorageDay = rstCYMGps.Fields("stoday")
            .PlugIn = IIf(IsNull(rstCYMGps.Fields("plugin")), cNullDate, rstCYMGps.Fields("plugin"))
            .PlugOut = IIf(IsNull(rstCYMGps.Fields("plugou")), cNullDate, rstCYMGps.Fields("plugou"))
            .RevenueTon = rstCYMGps.Fields("revton")
            .Discount = rstCYMGps.Fields("pctdsc")
            
            If .Discount > 0 And .Discount <> 1 Then
                .StorageAMT = ((.StorageNet / (100 - (.Discount * 100))) * 100)
            ElseIf .Discount = 1 Then
                .StorageAMT = 0
            Else
                .StorageAMT = .Storage + .StorageVAT - .StorageTAX
            End If
            
            .Oversize = rstCYMGps.Fields("ovzamt")
            .Commodity = Left(rstCYMGps.Fields("commod") & Space(30), 30)
            .UserID = rstCYMGps.Fields("userid")
            .ForExam = rstCYMGps.Fields("forexm")
            .Remark = rstCYMGps.Fields("remark")
            .CustomsGuard = rstCYMGps.Fields("cusgrd")
            .ConsCode = rstCYMGps.Fields("conscde")
            .VATCode = rstCYMGps.Fields("vatcde")
            .DangerClass = rstCYMGps.Fields("dgrcls")
            .BoatNote = rstCYMGps.Fields("boatnt")
            
            Select Case .UnderGuarantee
                Case "A"
                    .DueICTSI = .TotalCharge - .ArrastreNet
                Case "B"
                    .DueICTSI = .TotalCharge - .StorageNet
                Case "C"
                    .DueICTSI = .TotalCharge - .WeighingNet
                Case "D"
                    .DueICTSI = .TotalCharge - .ReeferNet
                Case "E"
                    .DueICTSI = .TotalCharge - .ArrastreNet - .StorageNet
                Case "F"
                    .DueICTSI = .TotalCharge - .ArrastreNet - .WeighingNet
                Case "G"
                    .DueICTSI = .TotalCharge - .ArrastreNet - .ReeferNet
                Case "H"
                    .DueICTSI = .TotalCharge - .StorageNet - .WeighingNet
                Case "I"
                    .DueICTSI = .TotalCharge - .StorageNet - .ReeferNet
                Case "J"
                    .DueICTSI = .TotalCharge - .WeighingNet - .ReeferNet
                Case "K"
                    .DueICTSI = .TotalCharge - .WeighingNet - .ArrastreNet - .StorageNet
                Case "L"
                    .DueICTSI = .TotalCharge - .ReeferNet - .ArrastreNet - .StorageNet
                Case "M"
                    .DueICTSI = .TotalCharge - .ReeferNet - .WeighingNet - .StorageNet
                Case "N"
                    .DueICTSI = .TotalCharge - .ReeferNet - .WeighingNet - .StorageNet - .ArrastreNet
                Case Else
                    .DueICTSI = .TotalCharge
            End Select
            .DueICTSIWords = .DueICTSI - .Wharfage
            
            If .Sequence = CInt(mskSequence) Then
                Call PrintGatePass
            End If
           
            rstCYMGps.MoveNext
        End With
    Loop
    On Error GoTo ErrorEndDoc
        Printer.EndDoc
    On Error GoTo 0
    rstCYMGps.Close
    Exit Sub
ErrorEndDoc:
    intResponse = MsgBox("Error printing...", vbExclamation + vbDefaultButton2 + vbAbortRetryIgnore, "Error!")
    If intResponse = vbAbort Then
        Unload Me
    ElseIf (intResponse = vbRetry) Or (intResponse = vbIgnore) Then
        Resume
    End If
End Sub

Private Sub PrintGatePass()
    Dim strToText As String
    Dim strPayment As String
    Dim blnChk1Printed As Boolean
    Dim blnChk2Printed As Boolean
    Dim blnChk3Printed As Boolean
    Dim blnChk4Printed As Boolean
    Dim blnChk5Printed As Boolean
    
    With Detail
        On Error GoTo ErrPrinting
            Printer.FontName = "Arial"
            Printer.FontSize = 11
            Printer.PrintQuality = vbPRPQDraft
            Printer.Print
            Printer.Print
            'Printer.Print
            Printer.Print Space(80); .Reference; Space(1); .Sequence; Space(1); .Gatepass
            Printer.Print
            Printer.Print Space(125); Format(.SysDate, "yyyy/mm/dd"); Space(5); Format(.SysDate, "hh:mm:ss")
            'Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print Space(8); Left(.Consignee, 30); Space(21); Left(.Registry, 10); Space(12); Left(.VoyageNo, 10); Space(20); Left(.CustomPermitNo, 10)
            Printer.Print
            'Printer.Print
            Printer.Print Space(8); Left(.Broker, 32); Space(30); Left(.BillNum, 20); Space(30); Left(.SMBAPermitNo, 10)
            Printer.Print
            'Printer.Print
            Printer.Print Space(8); Left(.VesselCode, 10); Space(23); Left(.PortofOrig, 3); Space(46); Format(.LastDischarge, "yyyy/mm/dd"); Space(20); .DeclaredWeight
            Printer.Print
            Printer.Print Space(8); Left(.ContainerNo, 12); Space(17); Left(.ContainerSize, 2); Space(7);
            If .FullEmp = "F" Then
                Printer.Print "FULL";
            Else
                Printer.Print "EMPTY";
            End If
            Printer.Print Space(10); Left(.Location, 10); Space(30); Left(.ShippingLine, 7)
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print Space(38); Space(42); ' Format(.FreeUntil, "yyyy/mm/dd"); Space(32);
            
            If .PlugIn = cNullDate Then
                Printer.Print Space(10);
                Printer.Print Space(8);
                Printer.Print Space(5)
            Else
                Printer.Print Format(.PlugIn, "yyyy/mm/dd");
                Printer.Print Space(3);
                Printer.Print Format(.PlugIn, "hh:mm")
            End If
           
            'Printer.Print
            If .PlugOut <> cNullDate Then
                Printer.Print Space(69); .StorageDay; Space(44);
                Printer.Print DateDiff("h", .PlugIn, .PlugOut);
            Else
                Printer.Print Space(127);
            End If
            Printer.Print Space(7);
            Printer.Print Format(.CRODate, "yyyy/mm/dd")
            Printer.Print Space(38); Space(42); 'Format(.StorageEnd, "yyyy/mm/dd"); Space(32);
            If .PlugOut = cNullDate Then
                Printer.Print Space(10);
                Printer.Print Space(8);
                Printer.Print Space(5);
            Else
                Printer.Print Format(.PlugOut, "yyyy/mm/dd");
                Printer.Print Space(3);
                Printer.Print Format(.PlugOut, "hh:mm")
            End If
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print
            If .RevenueTon > 0 Then
                RSet .strRevenueTon = CStr(Format(.RevenueTon, "###0.00"))
                Printer.Print Left(.strRevenueTon & Space(7), 7); Space(16);
            Else
                Printer.Print Space(7); Space(21);
            End If
                
            If .Oversize > 0 Then
                RSet .strArrastreLessOversize = CStr(Format(.ArrastreNet - .Oversize, "###,##0.00"))
                Printer.Print Left(.strArrastreLessOversize & Space(10), 10);
            Else
                RSet .strArrastreLessOversize = CStr(Format(.ArrastreNet, "###,##0.00"))
                Printer.Print Left(.strArrastreLessOversize & Space(10), 10);
            End If
            Printer.Print Space(21);
            
            RSet .strStorageAmt = CStr(Format(.StorageAMT, "###,##0.00"))
            Printer.Print Left(.strStorageAmt & Space(10), 10);
            Printer.Print Space(21);
            
            RSet .strReeferNet = CStr(Format(.ReeferNet, "###,##0.00"))
            Printer.Print Left(.strReeferNet & Space(10), 10);
            Printer.Print Space(21);
            
            RSet .strTotalCharge = CStr(Format(.TotalCharge, "####,##0.00"))
            Printer.Print Left(.strTotalCharge & Space(11), 11);
            Printer.Print
            Printer.Print
            Printer.Print Space(9); Space(8);
            If .Oversize > 0 Then
                RSet .strOversize = CStr(Format(.Oversize, "###,##0.00"))
                Printer.Print Left(.strOversize & Space(10), 10);
            Else
                Printer.Print Space(10);
            End If
            
            Printer.Print Space(8);
            If .Discount > 0 Then
                RSet .strDiscount = CStr(Format(.Discount, "##,##0.000"))
                Printer.Print Left(.strDiscount & Space(10), 10);
            Else
                Printer.Print Space(10)
            End If
            
            'for total
            'Printer.Print
            Printer.Print
            Printer.Print Space(29);
            Printer.Print Space(21);
            Select Case .VATCode
                Case Space(1), "4"
                    Printer.Print Space(20); "ZERO RATED VAT     ";
                Case "1"
                    Printer.Print Space(20); "VAT INCLUSIVE      ";
                Case "2", "3", "5"
                    Printer.Print Space(20); "VAT INCL. LESS WTAX";
            End Select
            
           'Printer.Print Space(42);
            
            RSet .strTotalCharge = CStr(Format(.TotalCharge, "####,##0.00"))
            Printer.Print Space(31); Left(.strTotalCharge & Space(11), 11)
            Printer.Print
            Printer.Print Space(5); Left(.Commodity, 20);
            Printer.Print Space(20);
           'Printer.Print Left(.UserID, 10);
            Printer.Print Space(10);
            Printer.Print Space(10);
            strToText = NumToText(.DueICTSIWords)
            Printer.Print Left(strToText, 45)
            Printer.Print Space(75); Mid(strToText, 40)
            
            'Printer.Print
            'Printer.Print
            Printer.Print
            Printer.Print Space(92); Format(.SysDate, "yyyymmdd"); Space(1); Format(.SysDate, "hhmm"); Space(1); _
                                   .Reference; Space(1); .Sequence; Space(1); .Gatepass
            'Printer.Print
            Printer.Print Space(5); Left(.UserID, 10); Space(14); Left(gbSupervisor, 15);
            Printer.Print Space(50);
            Select Case .UnderGuarantee
                Case Space(1)
                    Printer.Print "        "
                Case "A"
                    Printer.Print "U/G ARR "
                Case "B"
                    Printer.Print "U/G STRG"
                Case "C"
                    Printer.Print "U/G WGH "
                Case "D"
                    Printer.Print "U/G RFR "
                Case "N"
                    Printer.Print "ALL     "
                Case Else
                    Printer.Print "U/G     "
            End Select
            'Printer.Print Space(75)
            'Printer.Print Space(5)
            'Printer.Print
            If .DangerClass = Space(1) Then
                Printer.Print Space(4);
            Else
                Printer.Print "DC " & .DangerClass;
            End If
            
            strPayment = LiquidatePaymentTypes(2)
            Printer.Print Space(70);
            If strPayment <> "" Then
                Printer.Print Left(strPayment, 11);
                Printer.Print Space(1);
                Printer.Print "CK";
                blnChk1Printed = True
            Else
                Printer.Print Space(14);
                blnChk1Printed = False
            End If
            
            strPayment = LiquidatePaymentTypes(3)
            Printer.Print Space(20);
            If strPayment <> "" Then
                Printer.Print Left(strPayment, 11);
                Printer.Print Space(1);
                Printer.Print "CK";
                blnChk2Printed = True
            Else
                Printer.Print Space(14);
                blnChk2Printed = False
            End If
                
            Printer.Print Space(10);
            Printer.Print
            strPayment = LiquidatePaymentTypes(4)
            Printer.Print Space(70);
            If strPayment <> "" Then
                Printer.Print Left(strPayment, 11);
                Printer.Print Space(1);
                Printer.Print "CK";
                blnChk3Printed = True
            Else
                Printer.Print Space(14);
                blnChk3Printed = False
            End If
            'Printer.Print
            strPayment = LiquidatePaymentTypes(5)
            Printer.Print Space(10);
            If strPayment <> "" Then
                Printer.Print Left(strPayment, 11);
                Printer.Print Space(1);
                Printer.Print "CK";
                blnChk4Printed = True
            Else
                Printer.Print Space(14);
                blnChk4Printed = False
            End If
            Printer.Print
            strPayment = LiquidatePaymentTypes(6)
            Printer.Print Space(70);
            If strPayment <> "" Then
                Printer.Print Left(strPayment, 11);
                Printer.Print Space(1);
                Printer.Print "CK"
                blnChk5Printed = True
            Else
                Printer.Print Space(14)
                blnChk5Printed = False
            End If
            
            Printer.Print Space(92);
            If blnChk1Printed Then
                Printer.Print Left((Payment.CheckNo1 & Space(10)), 10);
                Printer.Print Space(1);
            End If
            If blnChk2Printed Then
                Printer.Print Left((Payment.CheckNo2 & Space(10)), 10);
                Printer.Print Space(1);
            End If
            If blnChk3Printed Then
                Printer.Print Left((Payment.CheckNo3 & Space(10)), 10);
                Printer.Print Space(1);
            End If
            If blnChk4Printed Then
                Printer.Print Left((Payment.CheckNo4 & Space(10)), 10);
                Printer.Print Space(1);
            End If
            If blnChk5Printed Then
                Printer.Print Left((Payment.CheckNo5 & Space(10)), 10);
                Printer.Print Space(1);
            End If
            Printer.Print Space(5)
            
            strPayment = LiquidatePaymentTypes(7)
            Printer.Print Space(70);
            If strPayment <> "" Then
                Printer.Print Left(strPayment, 11);
                Printer.Print Space(1);
                Printer.Print "CS";
            Else
                Printer.Print Space(14);
            End If
            Printer.Print Space(5);
            If blnChk1Printed Then
                Printer.Print Left((Payment.CheckBnk1 & Space(10)), 10);
                Printer.Print Space(1);
            End If
            If blnChk2Printed Then
                Printer.Print Left((Payment.CheckBnk2 & Space(10)), 10);
                Printer.Print Space(1);
            End If
            If blnChk3Printed Then
                Printer.Print Left((Payment.CheckBnk3 & Space(10)), 10);
                Printer.Print Space(1);
            End If
            If blnChk4Printed Then
                Printer.Print Left((Payment.CheckBnk4 & Space(10)), 10);
                Printer.Print Space(1);
            End If
            If blnChk5Printed Then
                Printer.Print Left((Payment.CheckBnk5 & Space(10)), 10);
                Printer.Print Space(1)
            Else
                Printer.Print Space(1)
            End If
            Printer.Print
            Printer.Print
            Printer.Print Space(110);
            Printer.Print .Remark
            If .CustomsGuard = "Y" Then
                Printer.Print Space(109); "/Underguard  "
            Else
                Printer.Print Space(109); Space(13)
            End If
            
            Printer.Print Space(4);
            Printer.Print
            Printer.Print
            Printer.Print Space(75); "CCR VALID UNTIL "; Space(10)
            'If .CRODate > .StorageEnd Then
            '    Printer.Print Format(.StorageEnd, "yyyy/mm/dd")
            'Else
            '    Printer.Print Format(.CRODate, "yyyy/mm/dd")
            'End If
            Printer.NewPage
        End With
    On Error GoTo 0
    Exit Sub

ErrPrinting:
    intResponse = MsgBox("Error printing...", vbExclamation + vbDefaultButton2 + vbAbortRetryIgnore, "Error!")
    If intResponse = vbAbort Then
        Unload Me
    ElseIf (intResponse = vbRetry) Or (intResponse = vbIgnore) Then
        Resume
    End If
End Sub
        
        
'
'
'
'
'
'
'            Printer.FontName = "Arial"
'            Printer.FontSize = 11
'            Printer.PrintQuality = vbPRPQDraft
'            Printer.Print
'            Printer.Print
'            'Printer.Print
'            Printer.Print Space(80); .Reference; Space(1); .Sequence; Space(1); .Gatepass
'            Printer.Print
'            Printer.Print Space(122); Format(.SysDate, "yyyy/mm/dd"); Space(5); Format(.SysDate, "hh:mm:ss")
'            Printer.Print
'            Printer.Print
'            Printer.Print
'            Printer.Print
'            'Printer.Print
'            Printer.Print Space(8); Left(.Consignee, 30); Space(18); Left(.Registry, 10); Space(12); Left(.VoyageNo, 10); Space(16); Left(.CustomPermitNo, 10)
'            Printer.Print
'            Printer.Print Space(8); Left(.Broker, 30); Space(26); Left(.BillNum, 20); Space(30); Left(.SMBAPermitNo, 10)
'            Printer.Print
'            'Printer.Print
'            Printer.Print Space(8); Left(.VesselCode, 7); Space(26); Left(.PortofOrig, 3); Space(46); Format(.LastDischarge, "yyyy/mm/dd"); Space(15); .DeclaredWeight
'            Printer.Print
'            Printer.Print Space(8); Left(.ContainerNo, 12); Space(17); Left(.ContainerSize, 2); Space(7);
'            If .FullEmp = "F" Then
'                Printer.Print "FULL";
'            Else
'                Printer.Print "EMPTY";
'            End If
'            Printer.Print Space(10); Left(.Location, 10); Space(20); Left(.ShippingLine, 7)
'            Printer.Print
'            Printer.Print
'            Printer.Print Space(35); Format(.FreeUntil, "yyyy/mm/dd"); Space(32);
'
'            If .PlugIn = cNullDate Then
'                Printer.Print Space(10);
'                Printer.Print Space(8);
'                Printer.Print Space(5)
'            Else
'                Printer.Print Format(.PlugIn, "yyyy/mm/dd");
'                Printer.Print Space(3);
'                Printer.Print Format(.PlugIn, "hh:mm")
'            End If
'
'            Printer.Print
'            Printer.Print Space(35); Format(.StorageEnd, "yyyy/mm/dd");
'            Printer.Print Space(21); .StorageDay; Space(6);
'            If .PlugOut = cNullDate Then
'                Printer.Print Space(10);
'                Printer.Print Space(8);
'                Printer.Print Space(5);
'            Else
'                Printer.Print Format(.PlugOut, "yyyy/mm/dd");
'                Printer.Print Space(3);
'                Printer.Print Format(.PlugOut, "hh:mm");
'            End If
'
'            If .PlugOut <> cNullDate Then
'                Printer.Print Space(7);
'                Printer.Print DateDiff("h", .PlugIn, .PlugOut);
'            Else
'                Printer.Print Space(3);
'                Printer.Print Space(5);
'            End If
'            Printer.Print Space(7);
'            Printer.Print Format(.CRODate, "yyyy/mm/dd")
'            Printer.Print
'            Printer.Print
'            Printer.Print
'            Printer.Print
'            If .RevenueTon > 0 Then
'                RSet .strRevenueTon = CStr(Format(.RevenueTon, "###0.00"))
'                Printer.Print Left(.strRevenueTon & Space(7), 7); Space(16);
'            Else
'                Printer.Print Space(7); Space(21);
'            End If
'
'            If .Oversize > 0 Then
'                RSet .strArrastreLessOversize = CStr(Format(.ArrastreNet - .Oversize, "###,##0.00"))
'                Printer.Print Left(.strArrastreLessOversize & Space(10), 10);
'            Else
'                RSet .strArrastreLessOversize = CStr(Format(.ArrastreNet, "###,##0.00"))
'                Printer.Print Left(.strArrastreLessOversize & Space(10), 10);
'            End If
'            Printer.Print Space(21);
'
'            RSet .strStorageAmt = CStr(Format(.StorageAMT, "###,##0.00"))
'            Printer.Print Left(.strStorageAmt & Space(10), 10);
'            Printer.Print Space(21);
'
'            RSet .strReeferNet = CStr(Format(.ReeferNet, "###,##0.00"))
'            Printer.Print Left(.strReeferNet & Space(10), 10);
'            Printer.Print Space(21);
'
'            RSet .strTotalCharge = CStr(Format(.TotalCharge, "####,##0.00"))
'            Printer.Print Left(.strTotalCharge & Space(11), 11);
'            Printer.Print
'            Printer.Print
'            Printer.Print Space(9); Space(8);
'            If .Oversize > 0 Then
'                RSet .strOversize = CStr(Format(.Oversize, "###,##0.00"))
'                Printer.Print Left(.strOversize & Space(10), 10);
'            Else
'                Printer.Print Space(10);
'            End If
'
'            Printer.Print Space(8);
'            If .Discount > 0 Then
'                RSet .strDiscount = CStr(Format(.Discount, "##,##0.000"))
'                Printer.Print Left(.strDiscount & Space(10), 10);
'            Else
'                Printer.Print Space(10)
'            End If
'
'            'for total
'            'Printer.Print
'            Printer.Print
'            Printer.Print Space(29);
'           'RSet .strArrastreLessOversize = CStr(Format(.ArrastreNet, "###,##0.00"))
'           'Printer.Print Left(.strArrastreLessOversize & Space(10), 10);
'           'Printer.Print Space(21);
'           '
'           'RSet .strStorageAmt = CStr(Format(.StorageNet, "###,##0.00"))
'           'Printer.Print Left(.strStorageAmt & Space(10), 10);
'           'Printer.Print Space(21);
'           '
'           'RSet .strReeferNet = CStr(Format(.ReeferNet, "###,##0.00"))
'           'Printer.Print Left(.strReeferNet & Space(10), 10);
'           'printer.Print Space(21);
'
'            Printer.Print Space(21);
'            Select Case .VATCode
'                Case Space(1), "4"
'                    Printer.Print Space(20); "ZERO RATED VAT     ";
'                Case "1"
'                    Printer.Print Space(20); "VAT INCLUSIVE      ";
'                Case "2", "3", "5"
'                    Printer.Print Space(20); "VAT INCL. LESS WTAX";
'            End Select
'
'           'Printer.Print Space(42);
'
'            RSet .strTotalCharge = CStr(Format(.TotalCharge, "####,##0.00"))
'            Printer.Print Space(31); Left(.strTotalCharge & Space(11), 11)
'            'Printer.Print
'            Printer.Print
'            Printer.Print Space(5); Left(.Commodity, 20);
'            Printer.Print Space(20);
'           'Printer.Print Left(.UserID, 10);
'            Printer.Print Space(10);
'            Printer.Print Space(10);
'            strToText = NumToText(.DueICTSIWords)
'            Printer.Print Left(strToText, 45)
'            Printer.Print Space(75); Mid(strToText, 40)
'
'            'Printer.Print
'            'Printer.Print
'            Printer.Print
'            Printer.Print Space(90); Format(.SysDate, "yyyymmdd"); Space(1); Format(.SysDate, "hhmm"); Space(1); _
'                                   .Reference; Space(1); .Sequence; Space(1); .Gatepass
'
'            Printer.Print Space(5); Left(.UserID, 10); Space(5); Left(gbSupervisor, 15);
'            Printer.Print Space(50);
'            Select Case .UnderGuarantee
'                Case Space(1)
'                    Printer.Print "        "
'                Case "A"
'                    Printer.Print "U/G ARR "
'                Case "B"
'                    Printer.Print "U/G STRG"
'                Case "C"
'                    Printer.Print "U/G WGH "
'                Case "D"
'                    Printer.Print "U/G RFR "
'                Case "N"
'                    Printer.Print "ALL     "
'                Case Else
'                    Printer.Print "U/G     "
'            End Select
'            'Printer.Print Space(75)
'            'Printer.Print Space(5)
'            'Printer.Print
'            If .DangerClass = Space(1) Then
'                Printer.Print Space(4);
'            Else
'                Printer.Print "DC " & .DangerClass;
'            End If
'
'            strPayment = LiquidatePaymentTypes(2)
'            Printer.Print Space(64);
'            If strPayment <> "" Then
'                Printer.Print Left(strPayment, 11);
'                Printer.Print Space(1);
'                Printer.Print "CK";
'                blnChk1Printed = True
'            Else
'                Printer.Print Space(14);
'                blnChk1Printed = False
'            End If
'
'            strPayment = LiquidatePaymentTypes(3)
'            Printer.Print Space(20);
'            If strPayment <> "" Then
'                Printer.Print Left(strPayment, 11);
'                Printer.Print Space(1);
'                Printer.Print "CK";
'                blnChk2Printed = True
'            Else
'                Printer.Print Space(14);
'                blnChk2Printed = False
'            End If
'
'            Printer.Print Space(10);
'            Printer.Print
'            strPayment = LiquidatePaymentTypes(4)
'            Printer.Print Space(68);
'            If strPayment <> "" Then
'                Printer.Print Left(strPayment, 11);
'                Printer.Print Space(1);
'                Printer.Print "CK";
'                blnChk3Printed = True
'            Else
'                Printer.Print Space(14);
'                blnChk3Printed = False
'            End If
'            'Printer.Print
'            strPayment = LiquidatePaymentTypes(5)
'            Printer.Print Space(10);
'            If strPayment <> "" Then
'                Printer.Print Left(strPayment, 11);
'                Printer.Print Space(1);
'                Printer.Print "CK";
'                blnChk4Printed = True
'            Else
'                Printer.Print Space(14);
'                blnChk4Printed = False
'            End If
'            Printer.Print
'            strPayment = LiquidatePaymentTypes(6)
'            Printer.Print Space(68);
'            If strPayment <> "" Then
'                Printer.Print Left(strPayment, 11);
'                Printer.Print Space(1);
'                Printer.Print "CK"
'                blnChk5Printed = True
'            Else
'                Printer.Print Space(14)
'                blnChk5Printed = False
'            End If
'
'            Printer.Print Space(89);
'            If blnChk1Printed Then
'                Printer.Print Left((Payment.CheckNo1 & Space(10)), 10);
'                Printer.Print Space(1);
'            End If
'            If blnChk2Printed Then
'                Printer.Print Left((Payment.CheckNo2 & Space(10)), 10);
'                Printer.Print Space(1);
'            End If
'            If blnChk3Printed Then
'                Printer.Print Left((Payment.CheckNo3 & Space(10)), 10);
'                Printer.Print Space(1);
'            End If
'            If blnChk4Printed Then
'                Printer.Print Left((Payment.CheckNo4 & Space(10)), 10);
'                Printer.Print Space(1);
'            End If
'            If blnChk5Printed Then
'                Printer.Print Left((Payment.CheckNo5 & Space(10)), 10);
'                Printer.Print Space(1);
'            End If
'            Printer.Print Space(5)
'
'            strPayment = LiquidatePaymentTypes(7)
'            Printer.Print Space(68);
'            If strPayment <> "" Then
'                Printer.Print Left(strPayment, 11);
'                Printer.Print Space(1);
'                Printer.Print "CS";
'            Else
'                Printer.Print Space(14);
'            End If
'            Printer.Print Space(5);
'            If blnChk1Printed Then
'                Printer.Print Left((Payment.CheckBnk1 & Space(10)), 10);
'                Printer.Print Space(1);
'            End If
'            If blnChk2Printed Then
'                Printer.Print Left((Payment.CheckBnk2 & Space(10)), 10);
'                Printer.Print Space(1);
'            End If
'            If blnChk3Printed Then
'                Printer.Print Left((Payment.CheckBnk3 & Space(10)), 10);
'                Printer.Print Space(1);
'            End If
'            If blnChk4Printed Then
'                Printer.Print Left((Payment.CheckBnk4 & Space(10)), 10);
'                Printer.Print Space(1);
'            End If
'            If blnChk5Printed Then
'                Printer.Print Left((Payment.CheckBnk5 & Space(10)), 10);
'                Printer.Print Space(1)
'            Else
'                Printer.Print Space(1)
'            End If
'            Printer.Print
'            'Printer.Print
'            Printer.Print Space(110);
'            Printer.Print .Remark
'            If .CustomsGuard = "Y" Then
'                Printer.Print Space(75); "/Underguard  "
'            Else
'                Printer.Print Space(75); Space(13)
'            End If
'
'            Printer.Print Space(4);
'            Printer.Print
'            Printer.Print
'            Printer.Print Space(75); "CCR VALID UNTIL ";
'            If .CRODate > .StorageEnd Then
'                Printer.Print Format(.StorageEnd, "yyyy/mm/dd")
'            Else
'                Printer.Print Format(.CRODate, "yyyy/mm/dd")
'            End If
'
'            Printer.NewPage
'        End With
'    On Error GoTo 0
'    Exit Sub
'
'ErrPrinting:
'    intResponse = MsgBox("Error printing...", vbExclamation + vbDefaultButton2 + vbAbortRetryIgnore, "Error!")
'    If intResponse = vbAbort Then
'        Unload Me
'    ElseIf (intResponse = vbRetry) Or (intResponse = vbIgnore) Then
'        Resume
'    End If
'End Sub

Private Sub GetTotalPaymentAmounts()
    Set rstCYMPay = New ADODB.Recordset
    rstCYMPay.LockType = adLockOptimistic
    rstCYMPay.CursorType = adOpenStatic
    rstCYMPay.Open "Select * from CYMPAY where Refnum= " & CLng(mskReference), gcnnBilling, , , adCmdText
    
    With Payment
        If IsNull(rstCYMPay.Fields("ftramt")) Then
            .POS = 0
        Else
            .POS = rstCYMPay.Fields("ftramt")
        End If
        .ADR = rstCYMPay.Fields("adramt")
        .CheckAmt1 = rstCYMPay.Fields("chkamt1")
        .CheckAmt2 = rstCYMPay.Fields("chkamt2")
        .CheckAmt3 = rstCYMPay.Fields("chkamt3")
        .CheckAmt4 = rstCYMPay.Fields("chkamt4")
        .CheckAmt5 = rstCYMPay.Fields("chkamt5")
        .CheckNo1 = rstCYMPay.Fields("chkno1")
        .CheckNo2 = rstCYMPay.Fields("chkno2")
        .CheckNo3 = rstCYMPay.Fields("chkno3")
        .CheckNo4 = rstCYMPay.Fields("chkno4")
        .CheckNo5 = rstCYMPay.Fields("chkno5")
        .CheckBnk1 = rstCYMPay.Fields("chkbnk1")
        .CheckBnk2 = rstCYMPay.Fields("chkbnk2")
        .CheckBnk3 = rstCYMPay.Fields("chkbnk3")
        .CheckBnk4 = rstCYMPay.Fields("chkbnk4")
        .CheckBnk5 = rstCYMPay.Fields("chkbnk5")
        .Cash = rstCYMPay.Fields("cshamt")
        .Change = rstCYMPay.Fields("chgamt")
        .Customer = rstCYMPay.Fields("cusnam")
        .TotalPayment = .POS + .ADR + .CheckAmt1 + .CheckAmt2 + .CheckAmt3 + .CheckAmt4 + .CheckAmt5 _
                                + .Cash - .Change
        .RemainingPayment = .TotalPayment
    End With
    rstCYMPay.Close
End Sub

Private Sub Form_Activate()
    mskReference.SetFocus
End Sub

Private Sub mskReference_KeyDown(KeyCode As Integer, Shift As Integer)
    Call FieldAdvance(KeyCode, mskReference, mskSequence)
End Sub

Private Sub mskSequence_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        'Call cmdReprint_Click
        cmdReprint.SetFocus
    Else
        Call FieldAdvance(KeyCode, mskReference, cmdReprint)
    End If
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

Public Function gzGetADRPaid(ByVal pUser As String, ByVal pFrom As Date, ByVal pTo As Date) As Currency
    Dim cmdGetADRPaid As ADODB.Command
    Dim prmGetADRPaid As ADODB.Parameter

    ' create command
    Set cmdGetADRPaid = New ADODB.Command
    Set prmGetADRPaid = New ADODB.Parameter
    With cmdGetADRPaid
        Set .ActiveConnection = gcnnBilling
        .CommandText = "up_getadrpaid"
        .CommandType = adCmdStoredProc
        
        .Parameters(0).Direction = adParamReturnValue
        .Parameters(1).Type = adChar
        .Parameters(1).Value = pUser
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adDate
        .Parameters(2).Value = pFrom
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adDate
        .Parameters(3).Value = pTo
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Direction = adParamOutput
        .Execute

        If IsNull(.Parameters(4)) Then
            gzGetADRPaid = 0
        Else
            gzGetADRPaid = .Parameters(4)
        End If
    End With
End Function

Public Function lzApplyADR(ByVal pCUSCDE As String, _
                            ByVal pREFTYP As String, _
                            ByVal pREFNUM As Long, _
                            ByVal pADRAMT As Currency, _
                            ByVal pUSERID As String, _
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
        .Parameters(1).Value = pCUSCDE
        .Parameters(1).Direction = adParamInput
        .Parameters(2).Type = adChar
        .Parameters(2).Value = pREFTYP
        .Parameters(2).Direction = adParamInput
        .Parameters(3).Type = adNumeric
        .Parameters(3).Value = pREFNUM
        .Parameters(3).Direction = adParamInput
        .Parameters(4).Type = adCurrency
        .Parameters(4).Value = pADRAMT
        .Parameters(4).Direction = adParamInput
        .Parameters(5).Type = adChar
        .Parameters(5).Value = pREMARK
        .Parameters(5).Direction = adParamInput
        .Parameters(6).Type = adChar
        .Parameters(6).Value = pUSERID
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

Private Sub FieldAdvance(pKeyCode As Integer, pPreviousControl As Control, pNextControl As Control)
    Select Case pKeyCode
        Case vbKeyDown
            If (TypeOf pNextControl Is TextBox) Or (TypeOf pNextControl Is MaskEdBox) Then
                pNextControl.SelStart = 0
                pNextControl.SelLength = pNextControl.MaxLength
            End If
            pNextControl.SetFocus
        Case vbKeyReturn
            If (TypeOf pNextControl Is TextBox) Or (TypeOf pNextControl Is MaskEdBox) Then
                pNextControl.SelStart = 0
                pNextControl.SelLength = pNextControl.MaxLength
            End If
            pNextControl.SetFocus
        Case vbKeyUp
            If (TypeOf pPreviousControl Is TextBox) Or (TypeOf pPreviousControl Is MaskEdBox) Then
                pPreviousControl.SelStart = 0
                pPreviousControl.SelLength = pPreviousControl.MaxLength
            End If
            pPreviousControl.SetFocus
         Case vbKeyF3
            intResponse = MsgBox("Do you really want to Exit?", vbYesNo + vbCritical, "Quit Program")
            If intResponse = vbYes Then
                Unload Me
            End If
    End Select
End Sub

