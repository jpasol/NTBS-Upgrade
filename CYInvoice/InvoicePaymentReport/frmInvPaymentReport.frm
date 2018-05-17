VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmInvPayReport 
   Caption         =   "CYInvoice Payment Reports"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   2175
   ClientWidth     =   10515
   ForeColor       =   &H00C00000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   10515
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      MouseIcon       =   "frmInvPaymentReport.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   9360
      Width           =   5535
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      MouseIcon       =   "frmInvPaymentReport.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   8520
      Width           =   5535
   End
   Begin VB.ComboBox cmbcustomer 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      ItemData        =   "frmInvPaymentReport.frx":0614
      Left            =   1200
      List            =   "frmInvPaymentReport.frx":0616
      MouseIcon       =   "frmInvPaymentReport.frx":0618
      MousePointer    =   99  'Custom
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Frame fradate 
      Caption         =   "Parameter"
      ForeColor       =   &H00C00000&
      Height          =   6135
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   5655
      Begin MSACAL.Calendar startDate 
         Height          =   2295
         Left            =   600
         TabIndex        =   3
         Top             =   840
         Width           =   4335
         _Version        =   524288
         _ExtentX        =   7646
         _ExtentY        =   4048
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2000
         Month           =   11
         Day             =   8
         DayLength       =   1
         MonthLength     =   0
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSACAL.Calendar endDate 
         Height          =   2295
         Left            =   720
         TabIndex        =   4
         Top             =   3720
         Width           =   4095
         _Version        =   524288
         _ExtentX        =   7223
         _ExtentY        =   4048
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2000
         Month           =   11
         Day             =   8
         DayLength       =   1
         MonthLength     =   0
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblend 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "End  Date"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   5175
      End
      Begin VB.Label lblstart 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Start Date"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.Frame fraReports 
      Caption         =   "Reports"
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5775
      Begin VB.ComboBox cmbReports 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   390
         ItemData        =   "frmInvPaymentReport.frx":0922
         Left            =   240
         List            =   "frmInvPaymentReport.frx":0924
         MouseIcon       =   "frmInvPaymentReport.frx":0926
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   5175
      End
   End
   Begin MSComctlLib.StatusBar statbar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   10635
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10319
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "11/11/2009"
         EndProperty
      EndProperty
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   10455
      Left            =   6000
      TabIndex        =   7
      Top             =   120
      Width           =   9000
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   960
   End
End
Attribute VB_Name = "frmInvPayReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbcustomer_GotFocus()
   statbar1.Panels(1).Text = " Choose A Customer "
   SendKeys "%{down}"
End Sub

Private Sub cmbReports_Click()
    Select Case cmbReports.ListIndex
            Case 0 ' Unpaid Bills
                    fradate.Visible = False
            
            Case Else
                    fradate.Visible = True
    End Select
End Sub

Private Sub cmbReports_GotFocus()
   statbar1.Panels(1).Text = " Choose A Report to View "
    SendKeys "%{down}"
End Sub

Private Sub cmbReports_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys "{TAB}"
  End If
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdClose_GotFocus()
  statbar1.Panels(1).Text = " Exit to  Invoice Payment Report "
End Sub

Private Sub cmdPreview_Click()
  Dim sDate1  As String
  Dim edate1 As String
    
    sDate1 = Format(startDate.Month & "/" & startDate.Day & "/" & startDate.Year, "mm,dd,yyyy")
    edate1 = Format(endDate.Month & "/" & endDate.Day & "/" & endDate.Year, "mm,dd,yyyy")
    
    If startDate.Value > endDate.Value And cmbReports.ListIndex <> 0 Then

        MsgBox " End Date must be Greater then Start Date ", vbOKOnly + vbInformation, " Invalid Date Range"
        endDate.Refresh
        startDate.Refresh
        endDate.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
  Screen.MousePointer = vbHourglass
  statbar1.Panels(1).Text = "Please Wait..Generating Reporting"
  Select Case cmbReports.ListIndex
    
        Case 0    'Report - UnPaid Bills
                Call Unpaid_RprtPreview
    
        Case 1    ' Summary of Inv. Payment
                 Call InvSummary_RprtPreview(sDate1, edate1)
                
        Case 2    'Detailed Report - Invoice Payment
                Call detailed_rprtPreview(sDate1, edate1)
                
        Case 3    ' Summary of O.R. (Isssuance)
                Call ORList_RprtPreview(sDate1, edate1)
                
        Case 4     ' Summary of Invoice with Adjustment
                Call InvAdjustment_RprtPreview(sDate1, edate1)
            
    End Select
    statbar1.Panels(1).Text = "Finished Generating Report"
    Screen.MousePointer = vbDefault
    DoEvents
If CRViewer1.IsBusy = False Then
        CRViewer1.Zoom (2)
End If
End Sub

Private Sub Unpaid_RprtPreview()
    Dim Unpaid_Rprt As New CrystalReport1
    Dim str1 As String
     
   Select Case cmbcustomer.ListIndex
              Case 0  ' All customer
                    Unpaid_Rprt.itxtCusnam.SetText ("Summary of UnPaid Bills of all Customers")
                    Unpaid_Rprt.Section8.Suppress = False
                
              Case Else  ' particular customer
                     str1 = "'" & Mid(cmbcustomer.Text, 10, Len(cmbcustomer.Text)) & "'"
                     str1 = "'" & Mid(cmbcustomer.Text, 1, 6) & "'"
                     Unpaid_Rprt.itxtCusnam.SetText (" UnPaid Bills ")
                     Unpaid_Rprt.Section8.Suppress = True
                     Unpaid_Rprt.RecordSelectionFormula = Unpaid_Rprt.RecordSelectionFormula & _
                                   " AND {INVICT.cuscde}= " & str1
  End Select
    CRViewer1.ReportSource = Unpaid_Rprt
    CRViewer1.ViewReport
 
End Sub
Private Sub InvSummary_RprtPreview(ByVal sdate As String, ByVal edate As String)
  Dim InvSummary_Rprt As New CrystalReport4
    'Summary Report - Invoice Payment
    
    InvSummary_Rprt.RecordSelectionFormula = ""
    
    Select Case cmbcustomer.ListIndex
    
         Case 0  ' All customer
                InvSummary_Rprt.RecordSelectionFormula = "DATE({viewfinalorlist.ordate}) >= DATE(" & Format(sdate, "yyyy,mm,dd") & ")" _
                             & " AND DATE({viewfinalorlist.ordate}) <= DATE(" & Format(edate, "yyyy,mm,dd") & ")"

          Case Else   'Particular Customer
                InvSummary_Rprt.RecordSelectionFormula = "DATE({viewfinalorlist.ordate}) >= DATE(" & Format(sdate, "yyyy,mm,dd") & ")" _
                               & " AND DATE({viewfinalorlist.ordate}) <= DATE(" & Format(edate, "yyyy,mm,dd") & ")" _
                               & " AND {viewfinalorlist.cuscde} = " & "'" & Mid(cmbcustomer.Text, 1, 6) & "'"
    End Select
    With InvSummary_Rprt
        .itxtTotalAdj.SetText Format(GetTotalAdjustement(), "###,###,###.#0")
        .itxtPayAmount.SetText Format(GetTotalPayAmount(), "###,###,###.#0")
        .itxtheader.SetText ("Summary Report- Invoice Payment")
        .itxtrange.SetText ("FOR THE PERIOD: " & Format(sdate, "mmm dd, yyyy") & "   TO   " & Format(edate, "mmm dd, yyyy"))
        .itxtcustomer.SetText ("( " & cmbcustomer.Text & " )")
    End With
    CRViewer1.ReportSource = InvSummary_Rprt
    CRViewer1.ViewReport
End Sub


Private Sub detailed_rprtPreview(ByVal sdate As String, ByVal edate As String)
    Dim Dtlpaid_Rprt As New CrystalReport2
    'Detailed Report - Invoice Payment
    
    Dtlpaid_Rprt.RecordSelectionFormula = ""
    
    Select Case cmbcustomer.ListIndex
         Case 0  ' All customer
              Dtlpaid_Rprt.RecordSelectionFormula = "{INVPAYHDR.ortype} <> 'ADJ' AND  DATE({INVPAYHDR.ordate}) >= DATE(" & Format(sdate, "yyyy,mm,dd") & ")" _
                               & " and DATE({INVPAYHDR.ordate}) <= DATE(" & Format(edate, "yyyy,mm,dd") & ")"
              Dtlpaid_Rprt.Section3.Suppress = False
              Dtlpaid_Rprt.Section8.Suppress = False
                
          Case Else   'Particular Customer
              Dtlpaid_Rprt.Section3.Suppress = True
              Dtlpaid_Rprt.Section8.Suppress = True
              Dtlpaid_Rprt.RecordSelectionFormula = "{INVPAYHDR.ortype} <> 'ADJ' AND  DATE({INVPAYHDR.ordate}) >= DATE(" & Format(sdate, "yyyy,mm,dd") & ")" _
                               & " AND DATE({INVPAYHDR.ordate}) <= DATE(" & Format(edate, "yyyy,mm,dd") & ")" _
                               & " AND {INVICT.cuscde} = " & "'" & Mid(cmbcustomer.Text, 1, 6) & "'"
    End Select
    Dtlpaid_Rprt.itxtheader.SetText ("Detailed Summary Report- Invoice Payment")
    Dtlpaid_Rprt.itxtrange.SetText ("FOR THE PERIOD: " & Format(sdate, "mmm dd, yyyy") & "   TO   " & Format(edate, "mmm dd, yyyy"))
    Dtlpaid_Rprt.itxtcustomer.SetText ("( " & cmbcustomer.Text & " )")
    CRViewer1.ReportSource = Dtlpaid_Rprt
    CRViewer1.ViewReport
End Sub
Private Sub InvAdjustment_RprtPreview(ByVal sdate As String, ByVal edate As String)
    Dim InvAdjustment_Rprt As New CrystalReport5
    'Summary Report of Invoice with Adjustment
    
    InvAdjustment_Rprt.RecordSelectionFormula = ""
    
    Select Case cmbcustomer.ListIndex
         Case 0  ' All customer
              InvAdjustment_Rprt.RecordSelectionFormula = "DATE({vueTAdjust.paydate}) >= DATE(" & Format(sdate, "yyyy,mm,dd") & ")" _
                               & " and DATE({vueTAdjust.paydate}) <= DATE(" & Format(edate, "yyyy,mm,dd") & ")"
                
          Case Else   'Particular Customer
              InvAdjustment_Rprt.RecordSelectionFormula = "DATE({vueTAdjust.paydate}) >= DATE(" & Format(sdate, "yyyy,mm,dd") & ")" _
                               & " AND DATE({vueTAdjust.paydate}) <= DATE(" & Format(edate, "yyyy,mm,dd") & ")" _
                               & " AND {INVICT.cuscde} = " & "'" & Mid(cmbcustomer.Text, 1, 6) & "'"
    End Select
    InvAdjustment_Rprt.itxtheader.SetText ("Summary Report of Invoice with Adjustment")
    InvAdjustment_Rprt.itxtrange.SetText ("FOR THE PERIOD: " & Format(sdate, "mmm dd, yyyy") & "   TO   " & Format(edate, "mmm dd, yyyy"))
    InvAdjustment_Rprt.itxtcustomer.SetText ("( " & cmbcustomer.Text & " )")
    CRViewer1.ReportSource = InvAdjustment_Rprt
    CRViewer1.ViewReport
End Sub

Private Sub ORList_RprtPreview(ByVal sdate As String, ByVal edate As String)
    Dim ORList As New CrystalReport3

    ORList.RecordSelectionFormula = ""
  With ORList
    Select Case cmbcustomer.ListIndex
         Case 0  ' All customer
              .RecordSelectionFormula = "DATE({viewfinalorlist.ordate}) >= DATE(" & Format(sdate, "yyyy,mm,dd") & ")" _
                               & " AND DATE({viewfinalorlist.ordate}) <= DATE(" & Format(edate, "yyyy,mm,dd") & ")"
                    
         Case Else   'Particular Customer
              .RecordSelectionFormula = "DATE({viewfinalorlist.ordate}) >= DATE(" & Format(sdate, "yyyy,mm,dd") & ")" _
                               & " AND DATE({viewfinalorlist.ordate}) <= DATE(" & Format(edate, "yyyy,mm,dd") & ")" _
                               & " AND {viewfinalorlist.cuscde} = " & "'" & Mid(cmbcustomer.Text, 1, 6) & "'"
    End Select
    .itxtTotalAdj.SetText Format(GetTotalAdjustement(), "###,###,###.#0")
    .itxtPayAmount.SetText Format(GetTotalPayAmount(), "###,###,###.#0")
    .itxtheader.SetText ("Summary Report of O.R. ")
    .itxtcustomer.SetText ("( " & cmbcustomer.Text & " )")
    .itxtrange.SetText ("FOR THE PERIOD: " & Format(sdate, "mmm dd, yyyy") & "   TO   " & Format(edate, "mmm dd, yyyy"))
    CRViewer1.ReportSource = ORList
    CRViewer1.ViewReport
 End With
End Sub
Private Function GetTotalAdjustement() As Currency
  Dim rst As ADODB.Recordset
  Dim nDayDepart As Long
  Dim sStart$, sEnd$, Ssql$
  Dim sCustomer As String
  Dim TotalADJ As Currency
  nDayDepart = endDate.Value - startDate.Value
  sStart = Format(startDate.Value, "YYYY-MM-DD")
  sEnd = Format(endDate.Value, "YYYY-MM-DD")
  sCustomer = IIf(cmbcustomer.ListIndex = 0, "", " AND cuscde=" & "'" & Mid(cmbcustomer.Text, 1, 6) & "'")
  If nDayDepart = 0 Then
    Ssql = " select distinct invnum,TAdjustment from viewfinalorlist " _
        & " WHERE  (year(ordate)= " & Year(startDate.Value) _
        & "  AND Month(ordate)=" & Month(startDate.Value) _
        & "  AND day(ordate)= " & Day(startDate.Value) & ")" _
        & sCustomer
  Else
    sEnd = Format(endDate.Value + 1, "YYYY-MM-DD")
    Ssql = "select distinct invnum,TAdjustment from viewfinalorlist" _
        & " Where  ordate>= '" & sStart & "'" _
        & " AND ordate< '" & sEnd & "'" _
        & sCustomer
  End If
    Set rst = New ADODB.Recordset
    rst.Open Ssql, gcnnBilling, , , adCmdText
    TotalADJ = 0
  With rst
    If Not .EOF Then
        While Not .EOF
           TotalADJ = TotalADJ + IIf(IsNull(!TAdjustment), 0, !TAdjustment)
           .MoveNext
        Wend
    End If
  End With
    GetTotalAdjustement = TotalADJ
    rst.Close
    Set rst = Nothing
End Function

Private Function GetTotalPayAmount() As Currency
  Dim rst As ADODB.Recordset
  Dim nDayDepart As Long
  Dim sStart$, sEnd$, Ssql$
  Dim sCustomer As String
  Dim TotalPayAmt As Currency
  nDayDepart = endDate.Value - startDate.Value
  sStart = Format(startDate.Value, "YYYY-MM-DD")
  sEnd = Format(endDate.Value, "YYYY-MM-DD")
  sCustomer = IIf(cmbcustomer.ListIndex = 0, "", " AND cuscde=" & "'" & Mid(cmbcustomer.Text, 1, 6) & "'")
  If nDayDepart = 0 Then
    Ssql = " select distinct ornum,invnum,payamt from viewfinalorlist " _
        & " WHERE  (year(ordate)= " & Year(startDate.Value) _
        & "  AND Month(ordate)=" & Month(startDate.Value) _
        & "  AND day(ordate)= " & Day(startDate.Value) & ")" _
        & sCustomer
  Else
    sEnd = Format(endDate.Value + 1, "YYYY-MM-DD")
    Ssql = " select distinct ornum,invnum,payamt from viewfinalorlist " _
        & " Where  ordate>= '" & sStart & "'" _
        & " AND ordate< '" & sEnd & "'" _
        & sCustomer
  End If
    Set rst = New ADODB.Recordset
    rst.Open Ssql, gcnnBilling, , , adCmdText
    TotalPayAmt = 0
  With rst
    If Not .EOF Then
        While Not .EOF
           TotalPayAmt = TotalPayAmt + IIf(IsNull(!Payamt), 0, !Payamt)
           .MoveNext
        Wend
    End If
  End With
    GetTotalPayAmount = TotalPayAmt
    rst.Close
    Set rst = Nothing
End Function


Private Sub initialize()
    Dim rst As New ADODB.Recordset
    rst.Open "customer", gcnnBilling, adOpenForwardOnly, , adCmdTable
    
    statbar1.Panels(1).Text = "Invoice Payment Reports "
    statbar1.Panels(2).Text = zCurrentComputer
    statbar1.Panels(3).Text = "User : " & zCurrentUser
    
    cmbReports.AddItem " UnPaid Bills"
    cmbReports.AddItem " Summary of Invoice Payment "
    cmbReports.AddItem " Detailed Report - Invoice Payment"
    cmbReports.AddItem " Summary of O.R. "
    cmbReports.AddItem " Summary of Invoice with Adjustment"
    cmbcustomer.AddItem "All Customers"
    cmbcustomer.ListIndex = 0
    cmbReports.ListIndex = 0
    startDate.Value = Date
    startDate.Year = Year(Date)
    startDate.Month = Month(Date)
    endDate.Value = Date
    endDate.Year = Year(Date)
    endDate.Month = Month(Date)
    endDate.Day = Day(Date)
    
    startDate.Day = 1
    startDate.Day = 1


    lblstart.Caption = "Start Date - " & Format(startDate.Month & " " & startDate.Day & " " & startDate.Year, "mmmm dd,yyyy")
    lblend.Caption = "End Date - " & Format(endDate.Month & " " & endDate.Day & " " & endDate.Year, "mmmm dd,yyyy")

    fradate.Visible = False
    
    Do While Not rst.EOF
        cmbcustomer.AddItem rst.Fields("cuscde").Value & " - " & rst.Fields("cusNam").Value
        rst.MoveNext
    Loop
    rst.Close
    Set rst = Nothing
End Sub

Private Sub cmdPreview_GotFocus()
  statbar1.Panels(1).Text = " Click to Preview Report "
End Sub

Private Sub endDate_GotFocus()
    statbar1.Panels(1).Text = " Specify the End Date of Report "
End Sub

Private Sub Form_GotFocus()
  statbar1.Panels(1).Text = " Invoice Payment Report "
End Sub

Private Sub startDate_Click()
    lblstart.Caption = "Start Date - " & Format(startDate.Month & " " & startDate.Day & " " & startDate.Year, "mmmm dd,yyyy")
End Sub

Private Sub startDate_GotFocus()
 statbar1.Panels(1).Text = " Specify the Start Date of Report  "
End Sub

Private Sub startDate_NewMonth()
      startDate.Day = 1
      startDate.Value = startDate.Day
      lblstart.Caption = "Start Date - " & Format(startDate.Month & " " & startDate.Day & " " & startDate.Year, "mmmm dd,yyyy")
End Sub

Private Sub startDate_NewYear()
      startDate.Value = startDate.Month
      lblstart.Caption = "Start Date - " & Format(startDate.Month & " " & startDate.Day & " " & startDate.Year, "mmmm dd,yyyy")
End Sub

Private Sub endDate_Click()
    lblend.Caption = "End Date - " & Format(endDate.Month & " " & endDate.Day & " " & endDate.Year, "mmmm dd,yyyy")
End Sub

Private Sub endDate_NewMonth()
     endDate.Day = 1
     endDate.Value = endDate.Day
    lblend.Caption = "End Date - " & Format(endDate.Month & " " & endDate.Day & " " & endDate.Year, "mmmm dd,yyyy")
End Sub

Private Sub endDate_NewYear()
    endDate.Value = endDate.Day
    lblend.Caption = "End Date - " & Format(endDate.Month & " " & endDate.Day & " " & endDate.Year, "mmmm dd,yyyy")
End Sub

Private Sub Form_Load()
  Call initialize
  cmdPreview.Default = True
End Sub
