VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#2.2#0"; "crviewer.dll"
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
   Begin VB.TextBox txtVesselName 
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
      Text            =   "[Type Vessel Name]"
      ToolTipText     =   "Type Vessel Name"
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtVoyageNo 
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
      Left            =   6480
      TabIndex        =   3
      Text            =   "[Type Voyage Number]"
      ToolTipText     =   "Type Voyage Number"
      Top             =   600
      Width           =   2895
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
      Left            =   13800
      TabIndex        =   10
      Top             =   600
      Width           =   1335
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
      Left            =   12480
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin MSMask.MaskEdBox mskDteTo 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
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
      Format          =   "yyyy-mm-dd"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDteFrom 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "yyyy-MM-dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
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
      Format          =   "yyyy-mm-dd"
      Mask            =   "####-##-##"
      PromptChar      =   "_"
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
      ItemData        =   "frmReport.frx":000C
      Left            =   9480
      List            =   "frmReport.frx":0019
      TabIndex        =   4
      Text            =   "[Select payment status]"
      Top             =   600
      Width           =   2895
   End
   Begin CRVIEWERLibCtl.CRViewer crvManifest 
      Height          =   9495
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   15015
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
   Begin VB.Label Label1 
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "Option"
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
      Left            =   12480
      TabIndex        =   9
      Top             =   120
      Width           =   2655
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
      Width           =   8775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "Date Interval"
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
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdView_Click()
    Dim rptManifest As New rptManifest
    Dim strSelection As String
    
    If Trim(cmbPayStatus.Text) = "[Select payment status]" Then
        MsgBox "Please select a criteria!", vbInformation, "Cargo Manifest Report"
        Exit Sub
    End If
       
    With rptManifest
        .txtPayStatus.SetText Trim(cmbPayStatus.Text)
        .txtDteFrom.SetText Trim(mskDteFrom.Text)
        .txtDteTo.SetText Trim(mskDteTo.Text)
        .txtGpsNum.SetText "Gatepass Number"
    End With
    
    strSelection = ""
       
    strSelection = "(date({CargoMDet.sysdte}) >= date(" & Left(mskDteFrom.Text, 4) & "," & _
                    Mid(mskDteFrom.Text, 9, 2) & "," & Mid(mskDteFrom.Text, 6, 2) & ")) AND " & _
                   "(time({CargoMDet.sysdte}) >= time(" & "00,00,00" & ")) AND " & _
                   "(date({CargoMDet.sysdte}) <= date(" & Left(mskDteTo.Text, 4) & "," & _
                    Mid(mskDteTo.Text, 9, 2) & "," & Mid(mskDteTo.Text, 6, 2) & ")) AND " & _
                   "(time({CargoMDet.sysdte}) <= time(" & "23,58,59" & ")) "
                   
    If Trim(cmbPayStatus.Text) = "Billed Container/s" Then
        'Display all billed container/s
        strSelection = strSelection & " AND (NOT ISNULL({CargoMDet.gpsnum}))"
    ElseIf Trim(cmbPayStatus.Text) = "Un-billed Container/s" Then
        rptManifest.txtGpsNum.SetText ""
        'Display all un-billed container/s
        strSelection = strSelection & " AND (ISNULL({CargoMDet.gpsnum}))"
    End If
    
    If Trim(txtVesselName.Text) <> "[Type Vessel Name]" And Trim(txtVesselName.Text) <> "" And Trim(txtVoyageNo.Text) <> "[Type Voyage Number]" And Trim(txtVoyageNo.Text) <> "" Then
        strSelection = strSelection & " AND ({CargoMHead.vslname}='" & Trim(txtVesselName.Text) & "')" & " AND ({CargoMHead.voynum}='" & Trim(txtVoyageNo.Text) & "')"
    ElseIf Trim(txtVesselName.Text) <> "[Type Vessel Name]" And Trim(txtVesselName.Text) <> "" Then
        strSelection = strSelection & " AND ({CargoMHead.vslname}='" & Trim(txtVesselName.Text) & "')"
    ElseIf Trim(txtVoyageNo.Text) <> "[Type Voyage Number]" And Trim(txtVoyageNo.Text) <> "" Then
        strSelection = strSelection & " AND ({CargoMHead.voynum}='" & Trim(txtVoyageNo.Text) & "')"
    End If
      
    rptManifest.RecordSelectionFormula = strSelection
    crvManifest.ReportSource = rptManifest
    crvManifest.ViewReport
End Sub

Private Sub Form_Load()
    Dim strDate As String
    
    strDate = Format(Now, "mmddyyyy")
    mskDteFrom.SelText = Right(Trim(strDate), 4) & Mid(Trim(strDate), 3, 2) & Left(Trim(strDate), 2)
    mskDteTo.SelText = Right(Trim(strDate), 4) & Mid(Trim(strDate), 3, 2) & Left(Trim(strDate), 2)
End Sub
