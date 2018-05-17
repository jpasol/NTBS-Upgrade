VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "CRVIEWER.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCYUgty 
   Caption         =   "SBITC Extraction of Underguarantee Bill"
   ClientHeight    =   10455
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCYUgty.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10455
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Height          =   420
      ItemData        =   "frmCYUgty.frx":014A
      Left            =   450
      List            =   "frmCYUgty.frx":0154
      TabIndex        =   11
      Text            =   "SBITC"
      Top             =   4200
      Width           =   2895
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   10080
      Visible         =   0   'False
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox txtInvNum 
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   9735
      Left            =   3720
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Preview of Invoices Generated"
      Top             =   240
      Width           =   11295
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview Invoice(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Picture         =   "frmCYUgty.frx":0164
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4900
      Width           =   2895
   End
   Begin MSMask.MaskEdBox mskStrDte 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   12582912
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####/##/##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskEndDte 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   12582912
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      Caption         =   "End date"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Start date"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblChgTyp 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1200
      TabIndex        =   6
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Underguarantee Bills"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   7
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Start Invoice No."
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuMenuExit 
         Caption         =   "E&xit"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&ChargeType"
      Begin VB.Menu mnuChgTyp 
         Caption         =   "&Export"
         Index           =   0
      End
      Begin VB.Menu mnuChgTyp 
         Caption         =   "&Import"
         Index           =   1
      End
      Begin VB.Menu mnuChgTyp 
         Caption         =   "&Special Services"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmCYUgty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim ChrgTyp As Integer
Dim intStoDys As Integer
Dim intRfrHrs As Integer

Const EX = 0    ' Export
Const IM = 1    ' Import
Const SP = 2    ' Special Services

Private Sub cmdPreview_Click()
    If Not IsDate(mskEndDte.Text) Or Not IsDate(mskStrDte.Text) Or _
       Not IsNumeric(Trim(txtInvNum)) Then
        MsgBox "Specify valid entries.", vbExclamation, "Error"
        mskStrDte.SetFocus: Exit Sub
    End If
    
    If ValidInvNum Then
        MsgBox "The invoice number specified has been used.", vbExclamation, "Error"
        txtInvNum.SetFocus: Exit Sub
    End If
    
    If MsgBox("Are all entries correct?  Continue generating " & UCase(lblChgTyp) & " invoice(s)?", _
       vbYesNo + vbDefaultButton2 + vbQuestion, "SBITC Underguarantee " & lblChgTyp) = vbYes Then TransferData
    Screen.MousePointer = vbDefault
    ProgressBar1.Visible = False
    mskStrDte.SetFocus
    
End Sub

Private Function ValidInvNum() As Boolean
    Dim rsINVICT As New ADODB.Recordset
    Dim sql As String
    
    sql = "Select invnum from INVICT where (invnum='" & Trim(txtInvNum) & "') AND (CompanyCode = '" & Combo1.Text & "') AND (status <> 'CAN')"
    With rsINVICT
        .Open sql, gcnnBilling, , , adCmdText
        If Not (.BOF And .EOF) Then
           ValidInvNum = True
        Else
           ValidInvNum = False
        End If
        .Close
    End With
 End Function


Private Sub Form_Load()
    mskStrDte.Text = Format(Now, "YYYY/MM/DD")
    mskEndDte.Text = Format(Now, "YYYY/MM/DD")
    ChrgTyp = EX    ' set default to export
    ConnectToBilling
End Sub


Private Sub mnuChgTyp_Click(Index As Integer)
    Select Case Index
        Case EX
            ChrgTyp = EX: lblChgTyp = "Export"
        Case IM
            ChrgTyp = IM: lblChgTyp = "Import"
        Case SP
            ChrgTyp = SP: lblChgTyp = "Special Services"
    End Select
End Sub

Private Sub mnuMenuExit_Click()
    If gbConnected Then gcnnBilling.Close
    Unload Me
End Sub

Private Sub TransferData()
    Dim rsUgty As New ADODB.Recordset
    Dim strSQL As String
    Dim tmpEndDte As Date
'    Dim prvRefNum As Long
    Dim tmpInvNum As Long
    Dim prvChgTyp As String
    Dim prvChgSze As String
    Dim prvCusCde As String
    Dim newCusCde As String
    
    Dim tmpRefNum As Long   ' used to retain ref# value for detail table
    Dim strCustmr As String
    Dim tmpItmNum As Integer
    Dim tmpTtlAmt As Currency
    Dim tmpTtlVat As Currency
    Dim tmpTtlTax As Currency
    Dim tmpAddMrk As String
    
'    index 0=Arrastre; 1=Storage; 2=Reefer
    Dim tmpRecTag As String ' based from the biltyp of rates
    Dim tmpRegNum As String
    Dim tmpVslCde As String
    Dim tmpUgType As String
    Dim tmpReference As Long
    Dim prvWithVAT As Boolean   ' previous with vat
    Dim curWithVAT As Boolean   ' current with vat
    Dim prvVATCde As String * 1
    Dim curVATCde As String * 1
    Dim tmpRteCde As String
    Dim tmpCntSze As String
    Dim tmpCntNum As String

    tmpInvNum = txtInvNum
    tmpEndDte = DateAdd("d", 1, CDate(mskEndDte.Text))
    tmpRteCde = ""
    prvCusCde = ""
    newCusCde = ""
    
'    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    
    Select Case ChrgTyp
'   ---------------------------- E X P O R T ------------------------------
        Case EX
       
            strSQL = "Select * from CCRCyx where (status <> 'CAN') and CompanyCode = '" & Combo1.Text & "' and " & _
                "(guarntycde = 'Y') and " & _
                "cast('" & mskStrDte.Text & "' as datetime) <= sysdttm and " & _
                "('" & tmpEndDte & "' > sysdttm) order by refnum"
                
            With rsUgty
            
                .Open strSQL, gcnnBilling, , , adCmdText
                
                If .EOF Then
                    MsgBox "No available data for the specified dates.", vbInformation, "Export Underguarantee"
                    .Close
                    Exit Sub
                End If
                
                tmpRefNum = gzGetRefNum("INV")
                
                Do Until .EOF
                
'                   Save first ref# for comparison
                    If prvCusCde = "" Then prvCusCde = GetCustomerCode(.Fields("refnum"), "E")
                    'If prvRefNum = 0 Then prvRefNum = .Fields("refnum")
                    tmpItmNum = tmpItmNum + 1
                    If tmpItmNum = 1 Then prvVATCde = .Fields("vatcde")
                    
                    tmpRteCde = ""
                    Select Case Trim(.Fields("cntsze"))
                      Case "20"
                        tmpRteCde = "CBEXP1"
                      Case "40"
                        tmpRteCde = "CBEXP2"
                      Case "45"
                        tmpRteCde = "CBEXP3"
                    End Select
                                        
                    Call SaveINVCYB(tmpRefNum, tmpItmNum, tmpRteCde, .Fields("cntsze"), _
                        .Fields("cntnum"), .Fields("ccrnum"), _
                        .Fields("arramt") + .Fields("ovzamt") + .Fields("dgramt") + .Fields("wghamt"), _
                        .Fields("arrvat"), .Fields("ovzamt"), .Fields("dgramt"), _
                        .Fields("revton"), "WGH", .Fields("arrtax"), .Fields("vatcde"), 1)

                    tmpTtlAmt = tmpTtlAmt + .Fields("arramt") + .Fields("ovzamt") + _
                                                  .Fields("dgramt") + .Fields("wghamt")
                    tmpTtlVat = tmpTtlVat + .Fields("arrvat")
                    tmpTtlTax = tmpTtlTax + .Fields("arrtax")
                        
                    .MoveNext
                    
                    If .EOF Then GoTo SaveINVICTex
                    
                    newCusCde = GetCustomerCode(.Fields("refnum"), "E")
                    
                    If (prvCusCde <> newCusCde) Or (tmpItmNum > 16) Or _
                       (prvVATCde <> .Fields("vatcde")) Then
SaveINVICTex:
                        Call SaveINVICT(prvCusCde, tmpRefNum, tmpInvNum, "", "", "ARRASTRE", tmpTtlAmt, _
                                        tmpTtlVat, "2", "EXPORT", tmpTtlTax)
                        If Not .EOF Then
                            prvCusCde = newCusCde
                            'prvRefNum = .Fields("refnum")
                            tmpRefNum = gzGetRefNum("INV")
                            tmpInvNum = tmpInvNum + 1
                            tmpItmNum = 0: tmpTtlAmt = 0: tmpTtlVat = 0: tmpTtlTax = 0
                        End If
    
                    End If
                    
                Loop
                .Close
                
            End With
            
'   ---------------------------- I M P O R T ------------------------------
        Case IM
        
        strSQL = "SELECT * FROM CYMGPS WHERE (status <> 'CAN') and CompanyCode = '" & Combo1.Text & "' and " & _
            "(gtycde <> '') and " & _
            "(left(broker, 3) <> 'BOC') and " & _
            " cast('" & mskStrDte.Text & "' as datetime) <= sysdte and " & _
            " ('" & tmpEndDte & "' > sysdte) order by gpsnum"
            
        With rsUgty

            .Open strSQL, gcnnBilling, , , adCmdText
            
            If .EOF Then
                MsgBox "No available data for the specified dates.", vbInformation, "Import Underguarantee"
                .Close
                Exit Sub
            End If
            
            tmpRefNum = gzGetRefNum("INV")
            
            Do Until .EOF
               
               If prvCusCde = "" Then prvCusCde = GetCustomerCode(.Fields("refnum"), "I")
               'If prvRefNum = 0 Then prvRefNum = .Fields("refnum")

               If (InStr("|AEFGKLN", .Fields("gtycde")) > 1 And .Fields("arramt") > 0) Or _
                  (InStr("|BEHIKLN", .Fields("gtycde")) > 1 And .Fields("stoamt") > 0) Or _
                  (InStr("|CFHJKMN", .Fields("gtycde")) > 1 And .Fields("wghamt") > 0) Or _
                  (InStr("|DGIJLMN", .Fields("gtycde")) > 1 And .Fields("rframt") > 0) Then
                                        
                    ' ARRASTRE UG
                    If InStr("|AEFGKLN", .Fields("gtycde")) > 1 And .Fields("arramt") > 0 Then
                                                    
                        tmpItmNum = tmpItmNum + 1
                        If tmpItmNum = 1 Then prvVATCde = .Fields("vatcde")
                        
                        tmpRteCde = ""
                        Select Case Trim(.Fields("cntsze"))
                          Case "20"
                            tmpRteCde = "CBIMP1"
                          Case "40"
                            tmpRteCde = "CBIMP2"
                          Case "45"
                            tmpRteCde = "CBIMP3"
                        End Select
                        
                        Call SaveINVCYB(tmpRefNum, tmpItmNum, tmpRteCde, _
                          .Fields("cntsze"), .Fields("cntnum"), .Fields("gpsnum"), _
                          .Fields("arramt"), .Fields("arrvat"), .Fields("ovzamt"), _
                          .Fields("dgramt"), .Fields("revton"), "", .Fields("arrtax"), .Fields("vatcde"), 1)
                            
                        tmpRegNum = .Fields("regnum")
                        tmpVslCde = .Fields("vslcde")
                        tmpRecTag = ""
                        tmpUgType = ""
                        tmpTtlAmt = tmpTtlAmt + .Fields("arramt")
                        tmpTtlVat = tmpTtlVat + .Fields("arrvat")
                        tmpTtlTax = tmpTtlTax + .Fields("arrtax")
                        ''
                        'compcode = .Fields("CompanyCode")
                        
                    End If
                    
                    ' STORAGE UG
                    If InStr("|BEHIKLN", .Fields("gtycde")) > 1 And .Fields("stoamt") > 0 Then
                            
                        tmpItmNum = tmpItmNum + 1
                        If tmpItmNum = 1 Then prvVATCde = .Fields("vatcde")
                        
                        tmpRteCde = ""
                        If Not IsNull(.Fields("plugin")) And Not IsNull(.Fields("plugou")) Then
                          ' if storage has plugin and plugout use rate code for
                          '     import storage reefer containers
                          Select Case Trim(.Fields("cntsze"))
                            Case "20"
                              tmpRteCde = "STOIM4"
                            Case "40"
                              tmpRteCde = "STOIM5"
                            Case "45"
                              tmpRteCde = "STOIM6"
                          End Select
                        Else
                          Select Case Trim(.Fields("cntsze"))
                            Case "20"
                              tmpRteCde = "STOIM1"
                            Case "40"
                              tmpRteCde = "STOIM2"
                            Case "45"
                              tmpRteCde = "STOIM3"
                          End Select
                        End If
                        
                        intStoDys = .Fields("stoday")
                            
                        Call SaveINVCYB(tmpRefNum, tmpItmNum, tmpRteCde, .Fields("cntsze"), _
                          .Fields("cntnum"), .Fields("gpsnum"), .Fields("stoamt"), _
                          .Fields("stovat"), 0, 0, 0, "STO", .Fields("stotax"), .Fields("vatcde"), 1)
                          
                        tmpRegNum = .Fields("regnum")
                        tmpVslCde = .Fields("vslcde")
                        tmpRecTag = ""
                        tmpUgType = ""
                        tmpTtlAmt = tmpTtlAmt + .Fields("stoamt")
                        tmpTtlVat = tmpTtlVat + .Fields("stovat")
                        tmpTtlTax = tmpTtlTax + .Fields("stotax")
                        
                    End If
                    
                    ' REEFER UG
                    If InStr("|DGIJLMN", .Fields("gtycde")) > 1 And .Fields("rframt") > 0 Then
                       
                        tmpItmNum = tmpItmNum + 1
                        If tmpItmNum = 1 Then prvVATCde = .Fields("vatcde")
                        
                        intRfrHrs = 0
                        If Not IsNull(.Fields("plugin")) And Not IsNull(.Fields("plugou")) Then
                            intRfrHrs = Fix(DateDiff("h", .Fields("plugin"), .Fields("plugou")))
                        End If

'                       Reefer hours not more than 1 day takes MCRFC1 as rate code,
'                             and disregards container size
                        If intRfrHrs <= 24 Then
                          tmpRteCde = "MCRFC1"
                          tmpCntSze = ""
                        Else
                          tmpCntSze = Trim(.Fields("cntsze"))
                          Select Case tmpCntSze
                            Case "20"
                              tmpRteCde = "MCRFC2"
                            Case "40"
                              tmpRteCde = "MCRFC3"
                            Case "45"
                              tmpRteCde = "MCRFC6"
                          End Select
                        End If
                        
                        Call SaveINVCYB(tmpRefNum, tmpItmNum, tmpRteCde, tmpCntSze, _
                          .Fields("cntnum"), .Fields("gpsnum"), .Fields("rframt"), _
                          .Fields("rfrvat"), 0, 0, 0, "RFR", .Fields("rfrtax"), .Fields("vatcde"), 1)
                        
                        tmpRegNum = .Fields("regnum")
                        tmpVslCde = .Fields("vslcde")
                        tmpRecTag = ""
                        tmpUgType = ""
                        tmpTtlAmt = tmpTtlAmt + .Fields("rframt")
                        tmpTtlVat = tmpTtlVat + .Fields("rfrvat")
                        tmpTtlTax = tmpTtlTax + .Fields("rfrtax")
                        
                    End If

               End If
               .MoveNext
               If .EOF Then GoTo SaveINVICTim
               
               newCusCde = GetCustomerCode(.Fields("refnum"), "I")
               
               If (prvCusCde <> newCusCde) Or (tmpItmNum > 16) Or _
                  (prvVATCde <> .Fields("vatcde")) Then
SaveINVICTim:
                  Call SaveINVICT(prvCusCde, tmpRefNum, tmpInvNum, tmpVslCde, tmpRegNum, tmpUgType, _
                                    tmpTtlAmt, tmpTtlVat, tmpRecTag, "IMPORT", tmpTtlTax)
                                    
                  If Not .EOF Then
                    prvCusCde = newCusCde
                    'prvRefNum = .Fields("refnum")
                    tmpRefNum = gzGetRefNum("INV")
                    tmpInvNum = tmpInvNum + 1
                    tmpItmNum = 0: tmpTtlAmt = 0: tmpTtlVat = 0: tmpTtlTax = 0
                  End If
                  
               End If
               
            Loop
            
            .Close
            
        End With
'   ---------------------------- SPECIAL SERVICES ------------------------------
        Case SP
        
            prvChgTyp = ""
            prvChgSze = ""
            strSQL = "Select * from CCRdtl where (status <> 'CAN') and CompanyCode = '" & Combo1.Text & "' and " & _
                "(guarntycde = 'Y') and " & _
                "cast('" & mskStrDte.Text & "' as datetime) <= sysdttm and " & _
                "('" & tmpEndDte & "' > sysdttm) order by refnum"
                
            With rsUgty
            
                .Open strSQL, gcnnBilling, , , adCmdText
                
                If .EOF Then
                    MsgBox "No available data for the specified dates.", vbInformation, "Export Underguarantee"
                    .Close
                    Exit Sub
                End If
                
                tmpRefNum = gzGetRefNum("INV")
                
                Do Until .EOF
                
                    'save first ref# for comparison
                    If prvCusCde = "" Then prvCusCde = GetCustomerCode(.Fields("refnum"), "E")
                    'If prvRefNum = 0 Then prvRefNum = .Fields("refnum")
                    tmpItmNum = tmpItmNum + 1
                    
                    prvVATCde = ""
                    If tmpItmNum = 1 And Not IsNull(.Fields("vatcde")) Then prvVATCde = .Fields("vatcde")
                    
                    If IsNull(Trim(.Fields("cntsze"))) Or Trim(.Fields("cntsze")) = "0" Then
                        tmpCntSze = ""
                    Else
                        tmpCntSze = Trim(.Fields("cntsze"))
                    End If
                    
                    tmpCntNum = ""
                    If Not IsNull(Trim(.Fields("cntnum"))) Then tmpCntNum = Trim(.Fields("cntnum"))
                    
                    tmpAddMrk = ""
                    
                    Call SaveINVCYB(tmpRefNum, tmpItmNum, .Fields("chargetyp"), _
                        Trim(tmpCntSze), Trim(tmpCntNum), .Fields("ccrnum"), _
                        .Fields("amt") + .Fields("ovzamt") + .Fields("dgramt"), _
                        .Fields("vatamt"), .Fields("ovzamt"), .Fields("dgramt"), _
                        .Fields("revton"), tmpAddMrk, .Fields("wtax"), prvVATCde, .Fields("quantity"))
                                            
                    tmpRegNum = "": tmpVslCde = "": tmpUgType = ""
                    If Not IsNull(.Fields("regnum")) Then tmpRegNum = Trim(.Fields("regnum"))
                    If Not IsNull(.Fields("vslcde")) Then tmpVslCde = Trim(.Fields("vslcde"))
                    If Not IsNull(.Fields("descr")) Then tmpUgType = Trim(.Fields("descr"))
                    tmpRecTag = "" 'GetRecordTag(.Fields("chargetyp"))
                    tmpTtlAmt = tmpTtlAmt + .Fields("amt") + .Fields("ovzamt") + _
                                .Fields("dgramt")
                    tmpTtlVat = tmpTtlVat + .Fields("vatamt")
                    tmpTtlTax = tmpTtlTax + .Fields("wtax")
                              
                    .MoveNext
                    
                    If .EOF Then GoTo SaveINVICTsp
                    
                    newCusCde = GetCustomerCode(.Fields("refnum"), "E")
                    
                    If prvCusCde <> newCusCde Or tmpItmNum > 16 Then
SaveINVICTsp:
                        Call SaveINVICT(prvCusCde, tmpRefNum, tmpInvNum, tmpVslCde, tmpRegNum, tmpUgType, tmpTtlAmt, _
                                        tmpTtlVat, tmpRecTag, "SPECIAL SERVICE", tmpTtlTax)
                        If Not .EOF Then
                            prvCusCde = newCusCde
                            'prvRefNum = .Fields("refnum")
                            tmpRefNum = gzGetRefNum("INV")
                            tmpInvNum = tmpInvNum + 1
                            tmpItmNum = 0: tmpTtlAmt = 0: tmpTtlVat = 0: tmpTtlTax = 0
                        End If
    
                    End If
                    
                Loop
                .Close
                
            End With
'   ---------------------------------- E L S E -------------------------------
        Case Else
            Exit Sub
            
    End Select
    
    Call PreviewOutput(tmpInvNum)
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
errHandler:
    Screen.MousePointer = vbDefault
    MsgBox "Error #" & Err.Number & " : " & Err.Description, vbCritical, "Error"
    MsgBox "Contact MIS for assistance.", vbCritical, "Error"
    Call mnuMenuExit_Click
    
End Sub

Private Sub SaveINVICT(pCusCde As String, pRefNum As Long, pInvNum As Long, pVslNam As String, _
                       pRegNum As String, pChgTyp As String, pTtlAmt As Currency, pTtlVat As Currency, _
                       pRecTag As String, pImpExp As String, pTtlTax As Currency)
                       
Dim rsINVICT As New ADODB.Recordset
Dim strCustmr As String
        
        With rsINVICT
            .Open "INVICT", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
            .AddNew
            .Fields("refnum") = pRefNum
            .Fields("invnum") = pInvNum
'            If pImpExp = "EXPORT" Or pImpExp = "SPECIAL SERVICE" Then
'                strCustmr = Trim(GetCustomerCode(pPrvRef, "E"))
'            Else
'                strCustmr = Trim(GetCustomerCode(pPrvRef, "I"))
'            End If
            .Fields("cuscde") = pCusCde
            .Fields("cusnam") = Trim(gzGetCustomerInfo(.Fields("cuscde")).cusnam)
            .Fields("invdttm") = gzGetSysDate
            .Fields("vslnam") = Trim(pVslNam)
            .Fields("regnum") = Trim(pRegNum)
            .Fields("invremark") = "U/G " & Trim(pChgTyp) & " - " & _
                                    Trim(pImpExp) & "|" & Trim(gzGetCustomerInfo(.Fields("cuscde")).careof)
            .Fields("invamt") = pTtlAmt
            .Fields("invvat") = pTtlVat
            .Fields("invtax") = pTtlTax
            .Fields("gtycde") = "Y"
            .Fields("status") = Space(1)
            .Fields("rectag") = pRecTag
            .Fields("userid") = zCurrentUser
            .Fields("updcde") = "A"
            .Fields("cfscy") = "1"
            .Fields("effdte") = ""
            .Fields("CompanyCode") = Combo1.Text
            .Update
            .Close
        End With
        If ProgressBar1.Value > 90 Then ProgressBar1.Value = 0
        ProgressBar1.Value = ProgressBar1.Value + 2
        
End Sub

Private Sub SaveINVCYB(pRefNum As Long, pItmNum As Integer, pRteCde As String, _
                       pCntSze As String, pCntNum As String, pGpsNum As Long, _
                       pInvAmt As Currency, pInvVat As Currency, pOvzAmt As Currency, _
                       pDgrAmt As Currency, pRevTon As Currency, pAddRmk As String, _
                       pInvTax As Currency, pVatCde As String, pQty As Currency)
                       
  Dim rsINVCYB As New ADODB.Recordset

    With rsINVCYB
            .Open "INVCYB", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
            .AddNew
            .Fields("refnum") = pRefNum
            .Fields("itmnum") = pItmNum
            .Fields("rtecde") = pRteCde
            .Fields("cntsze") = pCntSze
            .Fields("qty") = pQty
            .Fields("rtedsc") = "CCR#" & pGpsNum & " " & Trim(pCntNum) & "-" & Trim(pCntSze) & "'"
            If pOvzAmt > 0 Then
                .Fields("rtedsc") = .Fields("rtedsc") & " W/ O.H., REVTON = " & pRevTon
            End If
            If pDgrAmt > 0 Then
                .Fields("rtedsc") = .Fields("rtedsc") & " W/ ADDL. DANGER CHARGE"
            End If
            Select Case Trim(pAddRmk)
                Case ""
                    .Fields("dyshrs") = 1
                Case "STO"
                    .Fields("rtedsc") = .Fields("rtedsc") & "-" & intStoDys & " DAY(S)"
                    .Fields("dyshrs") = intStoDys
                Case "RFR"
                    .Fields("rtedsc") = .Fields("rtedsc") & "-" & intRfrHrs & " HR(S)"
                    .Fields("dyshrs") = intRfrHrs / 6
                Case "WGH"
                    .Fields("rtedsc") = .Fields("rtedsc") & " W/ADDL. WEIGHING CHARGE"
            End Select
            If pInvVat = 0 And pInvTax = 0 Then
                pVatCde = "0"
            ElseIf pInvVat > 0 And pInvTax = 0 Then
                pVatCde = "1"
            ElseIf pInvVat > 0 And pInvTax > 0 Then
                pVatCde = "2"
            ElseIf pInvVat = 0 And pInvTax > 0 Then
                pVatCde = "3"
            End If
            .Fields("vatcde") = pVatCde
            .Fields("invamt") = pInvAmt
            .Fields("invvat") = pInvVat
            .Fields("invtax") = pInvTax
            .Fields("sysdttm") = gzGetSysDate
            .Fields("status") = Space(1)
            .Fields("rectag") = Space(1)
            .Fields("userid") = zCurrentUser
            .Fields("updcde") = "A"
            .Fields("discnt") = 0
            .Fields("invremark") = ""
            .Fields("cargo") = "NA"
            .Fields("CompanyCode") = Combo1.Text
            .Update
            .Close
    End With
    If ProgressBar1.Value > 90 Then ProgressBar1.Value = 0
    ProgressBar1.Value = ProgressBar1.Value + 2

End Sub
' Returns customer code
' pImpExp string stores "I" for import or "E" for export
Private Function GetCustomerCode(pRefNum As Long, pImpExp As String) As String
    Dim rsPay As New ADODB.Recordset
    Dim strSQL As String
    
    If pImpExp = "I" Then
        strSQL = "Select * from CYMPay where (refnum= " & pRefNum & ")"
    Else
        strSQL = "Select * from CCRPay where (refnum= " & pRefNum & ")"
    End If
    With rsPay
        .Open strSQL, gcnnBilling, , , adCmdText
        If Not .EOF Then
            GetCustomerCode = Format(.Fields("cuscde"), "000000")
        End If
        .Close
    End With
End Function

Private Sub PreviewOutput(pLstInv As Long)

    Dim crInvoice As New crSubicRpt
    'Dim crInvoice As New CrystalReport1
    Dim tmpStartInv As Long

    tmpStartInv = CLng(txtInvNum)

    Do Until tmpStartInv > pLstInv
        crInvoice.ParameterFields(1).AddCurrentValue (tmpStartInv)
        crInvoice.ParameterFields(2).AddCurrentValue (Combo1.Text)
        tmpStartInv = tmpStartInv + 1
        If ProgressBar1.Value > 90 Then ProgressBar1.Value = 0
        ProgressBar1.Value = ProgressBar1.Value + 2
    Loop
    'CRViewer.ReportSource = crInvoice
    CRViewer.ReportSource = crInvoice
    ProgressBar1.Value = 100
    CRViewer.ViewReport
End Sub
 
Private Function GetReference(pInvNum As Long) As Long
    Dim rsINVICT As New ADODB.Recordset
    With rsINVICT
        .Open "Select refnum from INVICT where (invnum = " & pInvNum & ")", _
                gcnnBilling, , , adCmdText
        If Not .EOF Then GetReference = !refnum
        .Close
    End With
End Function

Private Function GetRecordTag(pRate As String) As String
  Dim rsCode As New ADODB.Recordset
  Dim tmpTag As String * 1
  
  With rsCode
      .Open "Select * from CYRate where cyr_rtecde = '" & Trim(pRate) & _
          "'", gcnnBilling, , , adCmdText
      Select Case Trim(!cyr_biltyp)
          Case "SO"
              tmpTag = "1"
          Case "AR"
              tmpTag = "2"
          Case "SF"
              tmpTag = "3"
          Case "VB"
              tmpTag = "4"
          Case "AN"
              tmpTag = "5"
          Case ""
              tmpTag = "0"
      End Select
      .Close
      GetRecordTag = tmpTag
  End With
End Function

