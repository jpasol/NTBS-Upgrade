VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmCYUgty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CY Undeguarantee Bills"
   ClientHeight    =   10830
   ClientLeft      =   150
   ClientTop       =   720
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10830
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
      Caption         =   "&Preview"
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
      Top             =   4200
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
      Left            =   1140
      TabIndex        =   6
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Underguarantee           Bills"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
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
    
    If MsgBox("Are all entries correct?  Continue generating invoice?", vbYesNo + vbDefaultButton2 + vbQuestion, "CY Underguarantee") = vbYes Then TransferData
    
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
        'Case SP
         '   ChrgTyp = SP: lblChgTyp = "Special Service"
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
    Dim tmpInvNum As Long
    
    Dim tmpRefNum As Long   ' used to retain ref# value for detail table
    Dim strCustmr As String
    Dim tmpItmNum(3) As Integer
    Dim tmpTtlAmt(3) As Currency
    Dim tmpTtlVat(3) As Currency
    
    Dim ARRpass As Integer
    Dim STOpass As Integer
    Dim WGHpass As Integer
    Dim RFRpass As Integer
'    Dim DGRpass As Integer  ' danger amt
'    Dim OVZpass As Integer  ' oversize amt
    
    ' index 0=AR; 1=ST; 2=WG; 3=RF
    Dim tmpRecTag(3) As String ' based from the biltyp of rates
    Dim tmpRegNum(3) As String
    Dim tmpVslCde(3) As String
    Dim tmpUgType(3) As String
    Dim tmpReference(3) As Long
    Dim prvWithVAT(3) As Boolean   ' previous with vat
    Dim curWithVAT(3) As Boolean   ' current with vat

    tmpInvNum = txtInvNum
    tmpEndDte = DateAdd("d", 1, CDate(mskEndDte.Text))
    
    On Error GoTo errHandler
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
'               tmpRefNum = 1
                
                Do Until .EOF
                
                    'save first ref# for comparison
                    If prvRefNum = 0 Then prvRefNum = .Fields("refnum")
                    tmpItmNum(0) = tmpItmNum(0) + 1
                    prvWithVAT(0) = False: curWithVAT(0) = False
                    
                    Call SaveINVCYB(tmpRefNum, tmpItmNum(0), "EXAR", .Fields("cntsze"), _
                        .Fields("cntnum"), .Fields("ccrnum"), _
                        .Fields("arramt") + .Fields("ovzamt") + .Fields("dgramt"), _
                        .Fields("arrvat"), .Fields("ovzamt"), .Fields("dgramt"), _
                        .Fields("revton"), "")
                        
                    If .Fields("arrvat") > 0 Then prvWithVAT(0) = True
                    tmpTtlAmt(0) = tmpTtlAmt(0) + .Fields("arramt") + .Fields("ovzamt") + _
                                                  .Fields("dgramt")
                    tmpTtlVat(0) = tmpTtlVat(0) + .Fields("arrvat")
                        
                    .MoveNext
                    
                    If .EOF Then GoTo SaveINVICTex
                    
                    If .Fields("arrvat") > 0 Then curWithVAT(0) = True
                    
                    If prvRefNum <> .Fields("refnum") Or tmpItmNum(0) > 14 Or _
                       (prvWithVAT(0) <> curWithVAT(0)) Then
SaveINVICTex:
                        Call SaveINVICT(prvRefNum, tmpRefNum, tmpInvNum, "", "", "ARRASTRE", tmpTtlAmt(0), _
                                        tmpTtlVat(0), "2", "EXPORT")
                        If Not .EOF Then
                            prvRefNum = .Fields("refnum")
                            tmpRefNum = gzGetRefNum("INV")
'                           tmpRefNum = tmpRefNum + 1   ' TEMPORARY ONLY'''''''''''
                            tmpInvNum = tmpInvNum + 1
                            tmpItmNum(0) = 0: tmpTtlAmt(0) = 0: tmpTtlVat(0) = 0
                        End If
    
                    End If
                    
                Loop
                .Close
                
            End With
            
'   ---------------------------- I M P O R T ------------------------------
        Case IM
        
      'TEMPORARY ONLY
'       tmpReference(0) = 100: tmpReference(1) = 200: tmpReference(2) = 300: tmpReference(3) = 400

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
'                           tmpReference(0) = tmpReference(0) + 1
                            tmpReference(0) = gzGetRefNum("INV")
                            prvWithVAT(0) = False
                            If .Fields("arrvat") > 0 Then prvWithVAT(0) = True
                        End If
                        
                        curWithVAT(0) = False
                        If .Fields("arrvat") > 0 Then curWithVAT(0) = True
                        
                        If prvWithVAT(0) <> curWithVAT(0) Then GoTo SaveINVICTim
                        
                        Call SaveINVCYB(tmpReference(0), tmpItmNum(0), "IMAR", _
                          .Fields("cntsze"), .Fields("cntnum"), .Fields("gpsnum"), _
                          .Fields("arramt"), .Fields("arrvat"), .Fields("ovzamt"), _
                          .Fields("dgramt"), .Fields("revton"), "")
                            
                        tmpRegNum(0) = .Fields("regnum")
                        tmpVslCde(0) = .Fields("vslcde")
                        tmpRecTag(0) = "2"
                        tmpUgType(0) = "ARRASTRE"
                        tmpTtlAmt(0) = tmpTtlAmt(0) + .Fields("arramt")
                        tmpTtlVat(0) = tmpTtlVat(0) + .Fields("arrvat")
                        ARRpass = 1
                        
'                       arrastre ug with oversize or dgramt is separate invoice
                        If .Fields("ovzamt") > 0 Or .Fields("dgramt") > 0 Then
                            GoTo SaveINVICTim
                        End If
                        
                    End If
                    
                    ' STORAGE UG
                    If InStr("|BEHIKLN", .Fields("gtycde")) > 1 And STOpass = 0 And _
                       .Fields("stoamt") > 0 Then
                            
                        tmpItmNum(1) = tmpItmNum(1) + 1
                        If tmpItmNum(1) = 1 Then
'                           tmpReference(1) = tmpReference(1) + 1
                            tmpReference(1) = gzGetRefNum("INV")
                            prvWithVAT(1) = False
                            If .Fields("arrvat") > 0 Then prvWithVAT(1) = True
                        End If

                        curWithVAT(1) = False
                        If .Fields("arrvat") > 0 Then curWithVAT(1) = True
                        
                        If prvWithVAT(1) <> curWithVAT(1) Then GoTo SaveINVICTim
                        
                        intStoDys = .Fields("stoday")
                            
                        Call SaveINVCYB(tmpReference(1), tmpItmNum(1), "IMST", .Fields("cntsze"), _
                          .Fields("cntnum"), .Fields("gpsnum"), .Fields("stoamt"), _
                          .Fields("stovat"), 0, 0, 0, "STO")
                          
                        tmpRegNum(1) = .Fields("regnum")
                        tmpVslCde(1) = .Fields("vslcde")
                        tmpRecTag(1) = "3"
                        tmpUgType(1) = "STORAGE"
                        tmpTtlAmt(1) = tmpTtlAmt(1) + .Fields("stoamt")
                        tmpTtlVat(1) = tmpTtlVat(1) + .Fields("stovat")
                        STOpass = 1
                        
                    End If
                    
                    ' WEIGHING UG
                    If InStr("|CFHJKMN", .Fields("gtycde")) > 1 And WGHpass = 0 And _
                       .Fields("wghamt") > 0 Then
                       
                        tmpItmNum(2) = tmpItmNum(2) + 1
                        If tmpItmNum(2) = 1 Then
'                           tmpReference(2) = tmpReference(2) + 1
                            tmpReference(2) = gzGetRefNum("INV")
                            prvWithVAT(2) = False
                            If .Fields("arrvat") > 0 Then prvWithVAT(2) = True
                        End If
    
                        curWithVAT(2) = False
                        If .Fields("arrvat") > 0 Then curWithVAT(2) = True
                        
                        If prvWithVAT(2) <> curWithVAT(2) Then GoTo SaveINVICTim
                        
                        Call SaveINVCYB(tmpReference(2), tmpItmNum(2), "WEIGHT", .Fields("cntsze"), _
                          .Fields("cntnum"), .Fields("gpsnum"), .Fields("wghamt"), _
                          .Fields("wghvat"), 0, 0, 0, "")
                       
                        tmpRegNum(2) = .Fields("regnum")
                        tmpVslCde(2) = .Fields("vslcde")
                        tmpRecTag(2) = "4"
                        tmpUgType(2) = "WEIGHING"
                        tmpTtlAmt(2) = tmpTtlAmt(2) + .Fields("wghamt")
                        tmpTtlVat(2) = tmpTtlVat(2) + .Fields("wghvat")
                        WGHpass = 1
                            
                    End If
                    
                    ' REEFER UG
                    If InStr("|DGIJLMN", .Fields("gtycde")) > 1 And RFRpass = 0 And _
                       .Fields("rframt") > 0 Then
                       
                        tmpItmNum(3) = tmpItmNum(3) + 1
                        If tmpItmNum(3) = 1 Then
'                           tmpReference(3) = tmpReference(3) + 1
                            tmpReference(3) = gzGetRefNum("INV")
                            prvWithVAT(3) = False
                            If .Fields("arrvat") > 0 Then prvWithVAT(3) = True
                        End If
                        
                        curWithVAT(3) = False
                        If .Fields("arrvat") > 0 Then curWithVAT(3) = True
                        
                        If prvWithVAT(3) <> curWithVAT(3) Then GoTo SaveINVICTim
                        
                        intRfrHrs = 0
                        If Not IsNull(.Fields("plugin")) And Not IsNull(.Fields("plugou")) Then
                            intRfrHrs = Fix(DateDiff("h", .Fields("plugin"), .Fields("plugou")))
                        End If
                        
                        Call SaveINVCYB(tmpReference(3), tmpItmNum(3), "IMRF", .Fields("cntsze"), _
                          .Fields("cntnum"), .Fields("gpsnum"), .Fields("rframt"), _
                          .Fields("rfrvat"), 0, 0, 0, "RFR")
                        
                        tmpRegNum(3) = .Fields("regnum")
                        tmpVslCde(3) = .Fields("vslcde")
                        tmpRecTag(3) = "1"
                        tmpUgType(3) = "REEFER"
                        tmpTtlAmt(3) = tmpTtlAmt(3) + .Fields("rframt")
                        tmpTtlVat(3) = tmpTtlVat(3) + .Fields("rfrvat")
                        RFRpass = 1
                        
                    End If

               End If
               .MoveNext
               ARRpass = 0: STOpass = 0: WGHpass = 0: RFRpass = 0 ': OVZpass = 0: DGRpass = 0
               If .EOF Then GoTo SaveINVICTim
               
               If (prvRefNum <> .Fields("refnum")) Or (tmpItmNum(0) > 14) Or _
                  (tmpItmNum(1) > 14) Or (tmpItmNum(2) > 14) Or (tmpItmNum(3) > 14) Then
SaveINVICTim:
                   If tmpTtlAmt(0) > 0 Then
                        Call SaveINVICT(prvRefNum, tmpReference(0), tmpInvNum, tmpVslCde(0), tmpRegNum(0), tmpUgType(0), _
                                    tmpTtlAmt(0), tmpTtlVat(0), tmpRecTag(0), "IMPORT")
                        tmpInvNum = tmpInvNum + 1
                        tmpItmNum(0) = 0: tmpTtlAmt(0) = 0: tmpTtlVat(0) = 0
                   End If
                   If tmpTtlAmt(1) > 0 Then
                        Call SaveINVICT(prvRefNum, tmpReference(1), tmpInvNum, tmpVslCde(1), tmpRegNum(1), tmpUgType(1), _
                                    tmpTtlAmt(1), tmpTtlVat(1), tmpRecTag(1), "IMPORT")
                        tmpInvNum = tmpInvNum + 1
                        tmpItmNum(1) = 0: tmpTtlAmt(1) = 0: tmpTtlVat(1) = 0
                   End If
                   If tmpTtlAmt(2) > 0 Then
                        Call SaveINVICT(prvRefNum, tmpReference(2), tmpInvNum, tmpVslCde(2), tmpRegNum(2), tmpUgType(2), _
                                    tmpTtlAmt(2), tmpTtlVat(2), tmpRecTag(2), "IMPORT")
                        tmpInvNum = tmpInvNum + 1
                        tmpItmNum(2) = 0: tmpTtlAmt(2) = 0: tmpTtlVat(2) = 0
                   End If
                   If tmpTtlAmt(3) > 0 Then
                        Call SaveINVICT(prvRefNum, tmpReference(3), tmpInvNum, tmpVslCde(3), tmpRegNum(3), tmpUgType(3), _
                                    tmpTtlAmt(3), tmpTtlVat(3), tmpRecTag(3), "IMPORT")
                        tmpInvNum = tmpInvNum + 1
                        tmpItmNum(3) = 0: tmpTtlAmt(3) = 0: tmpTtlVat(3) = 0
                   End If
                   
                   If Not rsUgty.EOF Then prvRefNum = .Fields("refnum")
                   
               End If
               
'               ARRpass = 0: STOpass = 0: WGHpass = 0: RFRpass = 0: OVZpass = 0: DGRpass = 0
                
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
                       pRecTag As String, pImpExp As String)
                       
Dim rsINVICT As New ADODB.Recordset
Dim strCustmr As String
    
        With rsINVICT
            .Open "INVICT", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
            .AddNew
            .Fields("refnum") = pRefNum
            .Fields("invnum") = pInvNum
            If pImpExp = "IMPORT" Then
                strCustmr = Trim(GetCustomerInfo(pPrvRef, "I"))
            Else
                strCustmr = Trim(GetCustomerInfo(pPrvRef, "E"))
            End If
            .Fields("cuscde") = Left(strCustmr, 6)
            .Fields("cusnam") = Mid(strCustmr, 8)
            .Fields("invdttm") = gzGetSysDate
            .Fields("vslnam") = Trim(pVslNam)
            .Fields("regnum") = Trim(pRegNum)
            .Fields("invremark") = "U/G " & Trim(pChgTyp) & " - " & _
                                    Trim(pImpExp) & "|" & GetAgent(.Fields("cuscde"))
            .Fields("invamt") = pTtlAmt
            .Fields("invvat") = pTtlVat
            .Fields("invtax") = 0
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
                       pDgrAmt As Currency, pRevTon As Currency, pAddRmk As String)
                       
Dim rsINVCYB As New ADODB.Recordset

    With rsINVCYB
            .Open "INVCYB", gcnnBilling, adOpenDynamic, adLockOptimistic, adCmdTable
            .AddNew
            .Fields("refnum") = pRefNum
            .Fields("itmnum") = pItmNum
            .Fields("rtecde") = pRteCde
            If pRteCde = "WEIGHT" Then
                .Fields("cntsze") = ""
            Else
                .Fields("cntsze") = pCntSze
            End If
            .Fields("qty") = 1
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
            If pInvVat > 0 Then
                .Fields("vatcde") = "1"
            Else
                .Fields("vatcde") = Space(1)
            End If
            .Fields("invamt") = pInvAmt
            .Fields("invvat") = pInvVat
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

' Returns customer code and name (e.g. 900080|'K' LINE (KAWASAKI KISEN KAISHA, LTD.)
' pImpExp string stores "I" for import or "E" for export
Private Function GetCustomerInfo(pRefNum As Long, pImpExp As String) As String
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
            If Trim(.Fields("cuscde")) <> "" Then GetCustomerInfo = Format(.Fields("cuscde"), "000000") & "|" & Trim(.Fields("cusnam"))
        End If
        .Close
    End With
End Function

Private Sub PreviewOutput(pLstInv As Long)
    Dim crCYInvoice As New crCYInvoice
    Dim tmpStartInv As Long
    
    tmpStartInv = CLng(txtInvNum)
    
    Do Until tmpStartInv > pLstInv
        crCYInvoice.ParameterFields(1).AddCurrentValue GetReference(tmpStartInv)
        tmpStartInv = tmpStartInv + 1
        If ProgressBar1.Value > 90 Then ProgressBar1.Value = 0
        ProgressBar1.Value = ProgressBar1.Value + 2
    Loop
    CRViewer.ReportSource = crCYInvoice
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

Private Function GetAgent(pCusCde As String) As String
    Dim rsCust As New ADODB.Recordset
    With rsCust
        .Open "Select careof from CUSTOMER where (cuscde='" & pCusCde & "')", _
                gcnnBilling, , , adCmdText
        If Not .EOF Then GetAgent = !careof
        .Close
    End With
End Function
