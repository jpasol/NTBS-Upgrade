VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCYUgty 
   Caption         =   "SBITC Extraction of Underguarantee Bill"
   ClientHeight    =   10455
   ClientLeft      =   165
   ClientTop       =   735
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
      Picture         =   "frmCYUgty.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
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
    
    If Not ValidInvNum Then
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
    
    sql = "Select * from INVICT where (invnum='" & Trim(txtInvNum) & "')"
    With rsINVICT
        .Open sql, gcnnBilling, , , adCmdText
        If .EOF Then ValidInvNum = True
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
    Dim prvRefNum As Long
'    Dim tmpInvNum(3) As Long
'    Dim tmpRefInvNum As Long
    Dim tmpInvNum As Long
    Dim prvChgTyp As String
    Dim prvChgSze As String
    
    Dim tmpRefNum As Long   ' used to retain ref# value for detail table
    Dim strCustmr As String
    Dim tmpItmNum(3) As Integer
    Dim tmpTtlAmt(3) As Currency
    Dim tmpTtlVat(3) As Currency
    Dim tmpTtlTax(3) As Currency
    Dim tmpAddMrk As String
    
    Dim ARRpass As Integer
    Dim STOpass As Integer
    Dim WGHpass As Integer
    Dim RFRpass As Integer
    
'    index 0=Arrastre; 1=Storage; 2=Weighing; 3=Reefer
    Dim tmpRecTag(3) As String ' based from the biltyp of rates
    Dim tmpRegNum(3) As String
    Dim tmpVslCde(3) As String
    Dim tmpUgType(3) As String
    Dim tmpReference(3) As Long
    Dim prvWithVAT(3) As Boolean   ' previous with vat
    Dim curWithVAT(3) As Boolean   ' current with vat
    Dim prvVATCde(3) As String * 1
    Dim curVATCde(3) As String * 1
    Dim tmpRteCde As String
    Dim tmpCntSze As String
    
'    tmpInvNum(0) = txtInvNum
    tmpInvNum = txtInvNum
    tmpEndDte = DateAdd("d", 1, CDate(mskEndDte.Text))
    tmpRteCde = ""
    
'    On Error GoTo errHandler
    Screen.MousePointer = vbHourglass
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    
    Select Case ChrgTyp
'   ---------------------------- E X P O R T ------------------------------
        Case EX
        
            strSQL = "Select * from CCRCyx where (status <> 'CAN') and " & _
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
                    If prvRefNum = 0 Then prvRefNum = .Fields("refnum")
                    tmpItmNum(0) = tmpItmNum(0) + 1
                    If tmpItmNum(0) = 1 Then prvVATCde(0) = .Fields("vatcde")
                    
                    tmpRteCde = ""
                    Select Case Trim(.Fields("cntsze"))
                      Case "20"
                        tmpRteCde = "CBEXP1"
                      Case "40"
                        tmpRteCde = "CBEXP2"
                      Case "45"
                        tmpRteCde = "CBEXP3"
                    End Select
                                        
                    Call SaveINVCYB(tmpRefNum, tmpItmNum(0), tmpRteCde, .Fields("cntsze"), _
                        .Fields("cntnum"), .Fields("ccrnum"), _
                        .Fields("arramt") + .Fields("ovzamt") + .Fields("dgramt"), _
                        .Fields("arrvat"), .Fields("ovzamt"), .Fields("dgramt"), _
                        .Fields("revton"), "", .Fields("arrtax"), .Fields("vatcde"), 1)

                    tmpTtlAmt(0) = tmpTtlAmt(0) + .Fields("arramt") + .Fields("ovzamt") + _
                                                  .Fields("dgramt")
                    tmpTtlVat(0) = tmpTtlVat(0) + .Fields("arrvat")
                    tmpTtlTax(0) = tmpTtlTax(0) + .Fields("arrtax")
                        
                    .MoveNext
                    
                    If .EOF Then GoTo SaveINVICTex
                    
                    If prvRefNum <> .Fields("refnum") Or tmpItmNum(0) > 14 Or _
                       (prvVATCde(0) <> .Fields("vatcde")) Then
SaveINVICTex:
                        Call SaveINVICT(prvRefNum, tmpRefNum, tmpInvNum, "", "", "ARRASTRE", tmpTtlAmt(0), _
                                        tmpTtlVat(0), "2", "EXPORT", tmpTtlTax(0))
                        If Not .EOF Then
                            prvRefNum = .Fields("refnum")
                            tmpRefNum = gzGetRefNum("INV")
                            tmpInvNum = tmpInvNum + 1
                            tmpItmNum(0) = 0: tmpTtlAmt(0) = 0: tmpTtlVat(0) = 0: tmpTtlTax(0) = 0
                        End If
    
                    End If
                    
                Loop
                .Close
                
            End With
            
'   ---------------------------- I M P O R T ------------------------------
        Case IM
        
        strSQL = "SELECT * FROM CYMGPS WHERE (status <> 'CAN') and " & _
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
            
            Do Until .EOF

               If prvRefNum = 0 Then prvRefNum = rsUgty.Fields("refnum")

               If (InStr("|AEFGKLN", .Fields("gtycde")) > 1 And .Fields("arramt") > 0 And ARRpass = 0) Or _
                  (InStr("|BEHIKLN", .Fields("gtycde")) > 1 And .Fields("stoamt") > 0 And STOpass = 0) Or _
                  (InStr("|CFHJKMN", .Fields("gtycde")) > 1 And .Fields("wghamt") > 0 And WGHpass = 0) Or _
                  (InStr("|DGIJLMN", .Fields("gtycde")) > 1 And .Fields("rframt") > 0 And RFRpass = 0) Then
                                        
                    ' ARRASTRE UG
                    If InStr("|AEFGKLN", .Fields("gtycde")) > 1 And ARRpass = 0 And _
                       .Fields("arramt") > 0 Then
                                                    
                        tmpItmNum(0) = tmpItmNum(0) + 1
                   
                        If tmpItmNum(0) = 1 Then
                          tmpReference(0) = gzGetRefNum("INV")
                          prvVATCde(0) = .Fields("vatcde")
                        End If
                        
                        If prvVATCde(0) <> .Fields("vatcde") Then GoTo SaveINVICTim
                        
                        tmpRteCde = ""
                        Select Case Trim(.Fields("cntsze"))
                          Case "20"
                            tmpRteCde = "CBIMP1"
                          Case "40"
                            tmpRteCde = "CBIMP2"
                          Case "45"
                            tmpRteCde = "CBIMP3"
                        End Select
                        
                        Call SaveINVCYB(tmpReference(0), tmpItmNum(0), tmpRteCde, _
                          .Fields("cntsze"), .Fields("cntnum"), .Fields("gpsnum"), _
                          .Fields("arramt"), .Fields("arrvat"), .Fields("ovzamt"), _
                          .Fields("dgramt"), .Fields("revton"), "", .Fields("arrtax"), .Fields("vatcde"), 1)
                            
                        tmpRegNum(0) = .Fields("regnum")
                        tmpVslCde(0) = .Fields("vslcde")
                        tmpRecTag(0) = "2"
                        tmpUgType(0) = "ARRASTRE"
                        tmpTtlAmt(0) = tmpTtlAmt(0) + .Fields("arramt")
                        tmpTtlVat(0) = tmpTtlVat(0) + .Fields("arrvat")
                        tmpTtlTax(0) = tmpTtlTax(0) + .Fields("arrtax")
                        ARRpass = 1
                        
                    End If
                    
                    ' STORAGE UG
                    If InStr("|BEHIKLN", .Fields("gtycde")) > 1 And STOpass = 0 And _
                       .Fields("stoamt") > 0 Then
                            
                        tmpItmNum(1) = tmpItmNum(1) + 1
                        If tmpItmNum(1) = 1 Then
                           tmpReference(1) = gzGetRefNum("INV")
                           prvVATCde(1) = .Fields("vatcde")
                        End If

                        If prvVATCde(1) <> .Fields("vatcde") Then GoTo SaveINVICTim
                        
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
                            
                        Call SaveINVCYB(tmpReference(1), tmpItmNum(1), tmpRteCde, .Fields("cntsze"), _
                          .Fields("cntnum"), .Fields("gpsnum"), .Fields("stoamt"), _
                          .Fields("stovat"), 0, 0, 0, "STO", .Fields("stotax"), .Fields("vatcde"), 1)
                          
                        tmpRegNum(1) = .Fields("regnum")
                        tmpVslCde(1) = .Fields("vslcde")
                        tmpRecTag(1) = "3"
                        tmpUgType(1) = "STORAGE"
                        tmpTtlAmt(1) = tmpTtlAmt(1) + .Fields("stoamt")
                        tmpTtlVat(1) = tmpTtlVat(1) + .Fields("stovat")
                        tmpTtlTax(1) = tmpTtlTax(1) + .Fields("stotax")
                        STOpass = 1
                        
                    End If
                    
'                    ' WEIGHING UG
'                    If InStr("|CFHJKMN", .Fields("gtycde")) > 1 And WGHpass = 0 And _
'                       .Fields("wghamt") > 0 Then
'
'                        tmpItmNum(2) = tmpItmNum(2) + 1
'                        If tmpItmNum(2) = 1 Then
'                           prvVATCde(2) = .Fields("vatcde")
'                        End If
'
'                        If prvVATCde(2) <> .Fields("vatcde") Then GoTo SaveINVICTim
'
'                        Call SaveINVCYB(tmpItmNum(2), "WEIGHT", .Fields("cntsze"), _
'                          .Fields("cntnum"), .Fields("gpsnum"), .Fields("wghamt"), _
'                          .Fields("wghvat"), 0, 0, 0, "", .Fields("wghtax"), .Fields("vatcde"))
'
'                        tmpRegNum(2) = .Fields("regnum")
'                        tmpVslCde(2) = .Fields("vslcde")
'                        tmpRecTag(2) = "4"
'                        tmpUgType(2) = "WEIGHING"
'                        tmpTtlAmt(2) = tmpTtlAmt(2) + .Fields("wghamt")
'                        tmpTtlVat(2) = tmpTtlVat(2) + .Fields("wghvat")
'                        tmpTtlTax(2) = tmpTtlTax(2) + .Fields("wghtax")
'                        WGHpass = 1
'                    End If
                    
                    ' REEFER UG
                    If InStr("|DGIJLMN", .Fields("gtycde")) > 1 And RFRpass = 0 And _
                       .Fields("rframt") > 0 Then
                       
                        tmpItmNum(3) = tmpItmNum(3) + 1
                        If tmpItmNum(3) = 1 Then
                            tmpReference(3) = gzGetRefNum("INV")
                            prvVATCde(3) = .Fields("vatcde")
                        End If
                        
                        If prvVATCde(3) <> .Fields("vatcde") Then GoTo SaveINVICTim
                        
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
                        
                        Call SaveINVCYB(tmpReference(3), tmpItmNum(3), tmpRteCde, tmpCntSze, _
                          .Fields("cntnum"), .Fields("gpsnum"), .Fields("rframt"), _
                          .Fields("rfrvat"), 0, 0, 0, "RFR", .Fields("rfrtax"), .Fields("vatcde"), 1)
                        
                        tmpRegNum(3) = .Fields("regnum")
                        tmpVslCde(3) = .Fields("vslcde")
                        tmpRecTag(3) = "1"
                        tmpUgType(3) = "REEFER"
                        tmpTtlAmt(3) = tmpTtlAmt(3) + .Fields("rframt")
                        tmpTtlVat(3) = tmpTtlVat(3) + .Fields("rfrvat")
                        tmpTtlTax(3) = tmpTtlTax(3) + .Fields("rfrtax")
                        RFRpass = 1
                        
                    End If

               End If
               .MoveNext
               ARRpass = 0: STOpass = 0: WGHpass = 0: RFRpass = 0
               If .EOF Then GoTo SaveINVICTim
               
               If (prvRefNum <> .Fields("refnum")) Or (tmpItmNum(0) > 14) Or _
                  (tmpItmNum(1) > 14) Or (tmpItmNum(2) > 14) Or (tmpItmNum(3) > 14) Then
SaveINVICTim:
                   If tmpTtlAmt(0) > 0 Then
                        Call SaveINVICT(prvRefNum, tmpReference(0), tmpInvNum, tmpVslCde(0), tmpRegNum(0), tmpUgType(0), _
                                    tmpTtlAmt(0), tmpTtlVat(0), tmpRecTag(0), "IMPORT", tmpTtlTax(0))
                        tmpInvNum = tmpInvNum + 1
                        tmpItmNum(0) = 0: tmpTtlAmt(0) = 0: tmpTtlVat(0) = 0: tmpTtlTax(0) = 0
                   End If
                   If tmpTtlAmt(1) > 0 Then
                        Call SaveINVICT(prvRefNum, tmpReference(1), tmpInvNum, tmpVslCde(1), tmpRegNum(1), tmpUgType(1), _
                                    tmpTtlAmt(1), tmpTtlVat(1), tmpRecTag(1), "IMPORT", tmpTtlTax(1))
                        tmpInvNum = tmpInvNum + 1
                        tmpItmNum(1) = 0: tmpTtlAmt(1) = 0: tmpTtlVat(1) = 0: tmpTtlTax(1) = 0
                   End If
'                   If tmpTtlAmt(2) > 0 Then
'                        Call SaveINVICT(prvRefNum, tmpInvNum(2), tmpVslCde(2), tmpRegNum(2), tmpUgType(2), _
'                                    tmpTtlAmt(2), tmpTtlVat(2), tmpRecTag(2), "IMPORT", tmpTtlTax(2))
'                        tmpInvNum = tmpInvNum + 1
'                        tmpItmNum(2) = 0: tmpTtlAmt(2) = 0: tmpTtlVat(2) = 0: tmpTtlTax(2) = 0
'                   End If
                   If tmpTtlAmt(3) > 0 Then
                        Call SaveINVICT(prvRefNum, tmpReference(3), tmpInvNum, tmpVslCde(3), tmpRegNum(3), tmpUgType(3), _
                                    tmpTtlAmt(3), tmpTtlVat(3), tmpRecTag(3), "IMPORT", tmpTtlTax(3))
                        tmpInvNum = tmpInvNum + 1
                        tmpItmNum(3) = 0: tmpTtlAmt(3) = 0: tmpTtlVat(3) = 0: tmpTtlTax(3) = 0
                   End If
                   
                   If Not rsUgty.EOF Then prvRefNum = .Fields("refnum")
                   
               End If
               
            Loop
            
            .Close
            
        End With
'   ---------------------------- SPECIAL SERVICES ------------------------------
        Case SP
        
            prvChgTyp = ""
            prvChgSze = ""
            strSQL = "Select * from CCRdtl where (status <> 'CAN') and " & _
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
                    If prvRefNum = 0 Then prvRefNum = .Fields("refnum")
                    If prvChgTyp = "" Then prvChgTyp = .Fields("chargetyp")
                    If prvChgSze = "" Then prvChgSze = .Fields("cntsze")
                    tmpItmNum(0) = tmpItmNum(0) + 1
                    
                    prvVATCde(0) = ""
                    If tmpItmNum(0) = 1 And Not IsNull(.Fields("vatcde")) Then
                      prvVATCde(0) = .Fields("vatcde")
                    End If
                    
                    If IsNull(Trim(.Fields("cntsze"))) Or Trim(.Fields("cntsze")) = "0" Then
                        tmpCntSze = ""
                    Else
                        tmpCntSze = Trim(.Fields("cntsze"))
                    End If
                    
                    
                    tmpAddMrk = ""
                    
                    Call SaveINVCYB(tmpRefNum, tmpItmNum(0), .Fields("chargetyp"), _
                        Trim(tmpCntSze), .Fields("cntnum"), .Fields("ccrnum"), _
                        .Fields("amt") + .Fields("ovzamt") + .Fields("dgramt"), _
                        .Fields("vatamt"), .Fields("ovzamt"), .Fields("dgramt"), _
                        .Fields("revton"), tmpAddMrk, .Fields("wtax"), prvVATCde(0), .Fields("quantity"))
                                            
                    tmpRegNum(0) = "": tmpVslCde(0) = "": tmpUgType(0) = ""
                    If Not IsNull(.Fields("regnum")) Then tmpRegNum(0) = Trim(.Fields("regnum"))
                    If Not IsNull(.Fields("vslcde")) Then tmpVslCde(0) = Trim(.Fields("vslcde"))
                    If Not IsNull(.Fields("descr")) Then tmpUgType(0) = Trim(.Fields("descr"))
                    tmpRecTag(0) = GetRecordTag(.Fields("chargetyp"))
                    tmpTtlAmt(0) = tmpTtlAmt(0) + .Fields("amt") + .Fields("ovzamt") + _
                                                  .Fields("dgramt")
                    tmpTtlVat(0) = tmpTtlVat(0) + .Fields("vatamt")
                    tmpTtlTax(0) = tmpTtlTax(0) + .Fields("wtax")
                              
                    .MoveNext
                    
                    If .EOF Then GoTo SaveINVICTsp
                    
                    If prvRefNum <> .Fields("refnum") Or tmpItmNum(0) > 14 Then
SaveINVICTsp:
                        Call SaveINVICT(prvRefNum, tmpRefNum, tmpInvNum, tmpVslCde(0), tmpRegNum(0), tmpUgType(0), tmpTtlAmt(0), _
                                        tmpTtlVat(0), tmpRecTag(0), "SPECIAL SERVICE", 0)
                        If Not .EOF Then
                            prvChgTyp = .Fields("chargetyp")
                            prvChgSze = .Fields("cntsze")
                            prvRefNum = .Fields("refnum")
                            tmpRefNum = gzGetRefNum("INV")
                            tmpInvNum = tmpInvNum + 1
                            tmpItmNum(0) = 0: tmpTtlAmt(0) = 0: tmpTtlVat(0) = 0: tmpTtlTax(0) = 0
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

Private Sub SaveINVICT(pPrvRef As Long, pRefNum As Long, pInvNum As Long, pVslNam As String, _
                       pRegNum As String, pChgTyp As String, pTtlAmt As Currency, pTtlVat As Currency, _
                       pRecTag As String, pImpExp As String, pTtlTax As Currency)
                       
Dim rsINVICT As New ADODB.Recordset
Dim strCustmr As String
        
        With rsINVICT
            .Open "INVICT", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
            .AddNew
            .Fields("refnum") = pRefNum
            .Fields("invnum") = pInvNum
            If pImpExp = "EXPORT" Or pImpExp = "SPECIAL SERVICE" Then
                strCustmr = Trim(GetCustomerCode(pPrvRef, "E"))
            Else
                strCustmr = Trim(GetCustomerCode(pPrvRef, "I"))
            End If
            .Fields("cuscde") = strCustmr
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
    Dim tmpStartInv As Long

    tmpStartInv = CLng(txtInvNum)

    Do Until tmpStartInv > pLstInv
        crInvoice.ParameterFields(1).AddCurrentValue (tmpStartInv)
        tmpStartInv = tmpStartInv + 1
        If ProgressBar1.Value > 90 Then ProgressBar1.Value = 0
        ProgressBar1.Value = ProgressBar1.Value + 2
    Loop
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
