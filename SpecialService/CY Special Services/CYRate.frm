VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCYRate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CY Special Services Tariff Rates"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14535
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CYRate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   14535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid grdCYRate 
      Height          =   8790
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   14430
      _ExtentX        =   25453
      _ExtentY        =   15505
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   65535
      ForeColorSel    =   -2147483640
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   "RATE    |^ SZ |^ TYPE |DESCRIPTION                             |AMOUNT       | UOM        "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ESC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   12480
      TabIndex        =   6
      Top             =   9000
      Width           =   930
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EXIT "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   480
      Left            =   13440
      TabIndex        =   5
      Top             =   9000
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   6720
      TabIndex        =   4
      Top             =   9000
      Width           =   1290
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SELECT FROM LIST"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   480
      Left            =   8040
      TabIndex        =   3
      Top             =   9000
      Width           =   3705
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UP / DN ARROW KEYS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   75
      TabIndex        =   2
      Top             =   9000
      Width           =   1440
   End
   Begin VB.Label Label56 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NAVIGATE UP / DOWN LIST"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   480
      Left            =   1560
      TabIndex        =   1
      Top             =   9000
      Width           =   4425
   End
End
Attribute VB_Name = "frmCYRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rst As ADODB.Recordset

Private Sub Form_Load()
    Call lzLoadRate
End Sub

Private Sub lzLoadRate()
Dim sSQL As String
Dim nPtr As Integer
Dim w As New CWaitCursor
    w.SetCursor
    Set rst = New ADODB.Recordset
    sSQL = "SELECT cyr_rtecde, cyr_cntsze, cyr_rtetyp, cyr_rtedsc, cyr_rteamt, cyr_uomcde "
    sSQL = sSQL & "FROM CYRate "
    sSQL = sSQL & "WHERE (cyr_rtetyp = 'EQP') OR (cyr_rtetyp = 'SPL')"
    sSQL = sSQL & "ORDER BY cyr_rtecde, cyr_cntsze "
    rst.Open sSQL, gcnnBilling, adOpenDynamic, adLockReadOnly, adCmdText
    If Not rst.EOF Then grdCYRate.Rows = 1
    grdCYRate.Visible = False
    While Not rst.EOF
        With grdCYRate
            nPtr = .Rows
            .AddItem "", nPtr
            .Row = nPtr
            .Col = 0: .Text = "" & rst!cyr_rtecde
            .Col = 1: .Text = "" & rst!cyr_cntsze
            .Col = 2: .Text = "" & rst!cyr_rtetyp
            .Col = 3: .Text = "" & rst!cyr_rtedsc
            .Col = 4: .Text = Format("" & rst!cyr_rteamt, "###,##0.00")
            .Col = 5: .Text = "" & rst!cyr_uomcde
        End With
        rst.MoveNext
    Wend
    With grdCYRate
        nPtr = .Rows
        .AddItem "", nPtr
        .Row = nPtr
        .TextMatrix(.Row, 0) = "<RELOAD>"
        .Row = 1: .Col = 0
        .Visible = True
    End With
    SendKeys ("{RIGHT}")
    On Error Resume Next
    rst.Close
    w.Restore
End Sub

Private Sub grdCYRate_DblClick()
    Call lzReturn
End Sub

Private Sub grdCYRate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyEscape
            vRateCode = ""
            vRateSz = ""
            vRateDesc = ""
            vRateAmount = 0
            vRateUOM = ""
            Unload Me
        Case vbKeyReturn
            Call lzReturn
        Case Else
    End Select
End Sub

Private Sub lzReturn()
    With grdCYRate
        vRateCode = .TextMatrix(.Row, 0)
        If vRateCode = "<RELOAD>" Then
            Call lzLoadRate
        Else
            vRateSz = .TextMatrix(.Row, 1)
            vRateDesc = .TextMatrix(.Row, 3)
            vRateAmount = CCur(.TextMatrix(.Row, 4))
            vRateUOM = .TextMatrix(.Row, 5)
            Unload Me
        End If
    End With
End Sub

