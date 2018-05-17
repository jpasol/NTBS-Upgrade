VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmReport 
   Caption         =   "Cargo Manifest Report"
   ClientHeight    =   10680
   ClientLeft      =   225
   ClientTop       =   615
   ClientWidth     =   15240
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbVessel 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   3
      Text            =   "[Select Vessel Name]"
      Top             =   1080
      Width           =   4455
   End
   Begin VB.ComboBox cmbShippingLine 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3600
      TabIndex        =   2
      Text            =   "[Select Carrier]"
      Top             =   600
      Width           =   8895
   End
   Begin MSMask.MaskEdBox mskArrivalTo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskArrivalFrom 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BackColor       =   12648447
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13920
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdView 
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
      Height          =   375
      Left            =   12600
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cmbPayStatus 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmReport.frx":0442
      Left            =   8160
      List            =   "frmReport.frx":044F
      TabIndex        =   4
      Text            =   "[Select Payment Status]"
      Top             =   1080
      Width           =   4335
   End
   Begin CRVIEWERLibCtl.CRViewer crvManifest 
      Height          =   9135
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   15255
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
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
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "Arrival Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   12600
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdView_Click()
    Dim rptManifest As New rptManifest
    Dim strSelection As String
    
    Screen.MousePointer = vbHourglass
    
    If Trim(cmbPayStatus.Text) = "[Select Payment Status]" Then
        Screen.MousePointer = vbDefault
        MsgBox "Please select a criteria!", vbInformation, "Cargo Manifest Report"
        Exit Sub
    End If
       
    With rptManifest
'        .txtPayStatus.SetText Trim(cmbPayStatus.Text)
'        .txtDteFrom.SetText Trim(mskArrivalFrom.Text)
'        .txtDteTo.SetText Trim(mskArrivalTo.Text)
'        .txtGpsNum.SetText "Gatepass Number"
    End With
    
    strSelection = ""
       
    strSelection = "(date({CargoMHead.arvdte}) >= date(" & Left(mskArrivalFrom.Text, 4) & "," & _
                    Mid(mskArrivalFrom.Text, 6, 2) & "," & Mid(mskArrivalFrom.Text, 9, 2) & ")) AND " & _
                   "(time({CargoMDet.sysdte}) >= time(" & "00,00,00" & ")) AND " & _
                   "(date({CargoMHead.arvdte}) <= date(" & Left(mskArrivalTo.Text, 4) & "," & _
                    Mid(mskArrivalTo.Text, 6, 2) & "," & Mid(mskArrivalTo.Text, 9, 2) & ")) AND " & _
                   "(time({CargoMDet.sysdte}) <= time(" & "23,58,59" & ")) "
                       
    If Trim(cmbPayStatus.Text) = "Billed Container/s" Then
        'Display all billed container/s
        strSelection = strSelection & " AND (NOT ISNULL({CargoMDet.gpsnum}))"
    ElseIf Trim(cmbPayStatus.Text) = "Un-billed Container/s" Then
        'rptManifest.txtGpsNum.SetText ""
        'Display all un-billed container/s
        strSelection = strSelection & " AND (ISNULL({CargoMDet.gpsnum}))"
    End If
    
    If Trim(cmbShippingLine.Text) <> "[Select Carrier]" And Trim(cmbShippingLine.Text) <> "" And Trim(cmbVessel.Text) <> "[Select Vessel Name]" And Trim(cmbVessel.Text) <> "" Then
        strSelection = strSelection & " AND ({CargoMHead.carcde}='" & Mid(Trim(cmbShippingLine.Text), 1, 6) & "')" & " AND ({CargoMHead.vslname}='" & Trim(cmbVessel.Text) & "')"
    ElseIf Trim(cmbShippingLine.Text) <> "[Select Carrier]" And Trim(cmbShippingLine.Text) <> "" Then
        strSelection = strSelection & " AND ({CargoMHead.carcde}='" & Mid(Trim(cmbShippingLine.Text), 1, 6) & "')"
    ElseIf Trim(cmbVessel.Text) <> "[Select Vessel Name]" And Trim(cmbVessel.Text) <> "" Then
        strSelection = strSelection & " AND ({CargoMHead.vslname}='" & Trim(cmbVessel.Text) & "')"
    End If
    
    rptManifest.RecordSelectionFormula = strSelection
    crvManifest.ReportSource = rptManifest
    crvManifest.ViewReport
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim strDate As String
    Dim rstSelections As New ADODB.Recordset
    
    strDate = Format(Now, "mmddyyyy")
    mskArrivalFrom.SelText = Right(Trim(strDate), 4) & Mid(Trim(strDate), 1, 2) & Mid(Trim(strDate), 3, 2)
    mskArrivalTo.SelText = Right(Trim(strDate), 4) & Mid(Trim(strDate), 1, 2) & Mid(Trim(strDate), 3, 2)
        
    If Not gcnnBilling Is Nothing Then
        'Populate Shipping Line Selections
        rstSelections.Open "SELECT cuscde,cusnam FROM Customer ORDER BY cusnam", gcnnBilling, adOpenDynamic
        cmbShippingLine.Clear
        cmbShippingLine.Text = "[Select Carrier]"
        cmbShippingLine.AddItem "[Select Carrier]"
        If Not rstSelections.BOF Then
            rstSelections.MoveFirst
            Do While Not rstSelections.EOF
                cmbShippingLine.AddItem rstSelections.Fields("cuscde") & " | " & rstSelections.Fields("cusnam")
                rstSelections.MoveNext
            Loop
            rstSelections.Close
        End If
        'Populate Vessel Name Selections
        rstSelections.Open "SELECT DISTINCT vslname FROM CargoMHead ORDER BY vslname", gcnnBilling
        cmbVessel.Clear
        cmbVessel.Text = "[Select Vessel Name]"
        cmbVessel.AddItem "[Select Vessel Name]"
        If Not rstSelections.BOF Then
            rstSelections.MoveFirst
            Do While Not rstSelections.EOF
                cmbVessel.AddItem rstSelections.Fields("vslname")
                rstSelections.MoveNext
            Loop
            rstSelections.Close
        End If
    End If
End Sub

