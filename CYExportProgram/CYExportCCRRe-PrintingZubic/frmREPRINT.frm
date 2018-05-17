VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmREPRINT 
   BorderStyle     =   0  'None
   Caption         =   "( SUBIC - ZCCRCYREPRT ) CY Export CCR Re-Printing"
   ClientHeight    =   10770
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "IBM3270 - 1254"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   10275
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12832
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3704
            Picture         =   "frmREPRINT.frx":0000
            Text            =   "CCRCYREPRT"
            TextSave        =   "CCRCYREPRT"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "7/11/00"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "8:41 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SST 
      Height          =   10815
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   19076
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "By Reference Number"
      TabPicture(0)   =   "frmREPRINT.frx":27B2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdByccr"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "flxRef"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdExit"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdPrinter"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdPrint"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdSearch"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "By CCR Number"
      TabPicture(1)   =   "frmREPRINT.frx":27CE
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtReference"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label8"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdSearch2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdPrint2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdPrinter2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdExit2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame4"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "flxCCR"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmdByref"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtTab(1)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.TextBox txtTab 
         Height          =   420
         Index           =   1
         Left            =   240
         MaxLength       =   8
         TabIndex        =   24
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtTab 
         Height          =   420
         Index           =   0
         Left            =   -74760
         MaxLength       =   8
         TabIndex        =   0
         Top             =   600
         Width           =   3855
      End
      Begin VB.CommandButton cmdByccr 
         Caption         =   "F11 - By CCR Number"
         Height          =   615
         Left            =   -63600
         TabIndex        =   1
         Top             =   480
         Width           =   3735
      End
      Begin VB.CommandButton cmdByref 
         Caption         =   "F11 - By Reference"
         Height          =   615
         Left            =   11400
         TabIndex        =   23
         Top             =   480
         Width           =   3735
      End
      Begin MSFlexGridLib.MSFlexGrid flxRef 
         Height          =   7695
         Left            =   -74880
         TabIndex        =   2
         Top             =   1680
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   13573
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   "Opt |Reference|Sequence| CCR No. | Exporter Name     | Broker Name      |  Date           "
      End
      Begin MSFlexGridLib.MSFlexGrid flxCCR 
         Height          =   7575
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   15015
         _ExtentX        =   26485
         _ExtentY        =   13361
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         FocusRect       =   2
         SelectionMode   =   1
         FormatString    =   "No. | Container No. | Size |     Exporter Name     |      Broker Name   |  Date & Time     "
      End
      Begin VB.Frame Frame4 
         Height          =   135
         Left            =   120
         TabIndex        =   18
         Top             =   9360
         Width           =   15015
      End
      Begin VB.CommandButton cmdExit2 
         Caption         =   "F3 - E&xit"
         Height          =   615
         Left            =   13320
         TabIndex        =   17
         Top             =   9600
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrinter2 
         Caption         =   "F5 - Change Printer"
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   9600
         Width           =   3615
      End
      Begin VB.CommandButton cmdPrint2 
         Caption         =   "F7 - &Print CCR"
         Height          =   615
         Left            =   3840
         TabIndex        =   15
         Top             =   9600
         Width           =   3135
      End
      Begin VB.CommandButton cmdSearch2 
         Caption         =   "F6 - Another CCR"
         Height          =   615
         Left            =   7080
         TabIndex        =   14
         Top             =   9600
         Width           =   6135
      End
      Begin VB.Frame Frame3 
         Height          =   135
         Left            =   -74880
         TabIndex        =   13
         Top             =   9360
         Width           =   15015
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "F3 - E&xit"
         Height          =   615
         Left            =   -61680
         TabIndex        =   6
         Top             =   9600
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrinter 
         Caption         =   "F5 - Change Printer"
         Height          =   615
         Left            =   -74880
         TabIndex        =   3
         Top             =   9600
         Width           =   3615
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "F7 - &Print"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -71160
         TabIndex        =   4
         Top             =   9600
         Width           =   3135
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "F6 - Another Reference"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -67920
         TabIndex        =   5
         Top             =   9600
         Width           =   6135
      End
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   15015
      End
      Begin VB.Frame Frame1 
         Height          =   135
         Left            =   -74880
         TabIndex        =   11
         Top             =   1080
         Width           =   15015
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reference No."
         Height          =   375
         Left            =   4200
         TabIndex        =   26
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label txtReference 
         BorderStyle     =   1  'Fixed Single
         Height          =   420
         Left            =   4200
         TabIndex        =   25
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "     C C R   D E T A I L S"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   15015
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "(1)"
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   -74640
         TabIndex        =   21
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "     - Print"
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   1320
         Width           =   15015
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CCR Number"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reference"
         Height          =   375
         Left            =   -74760
         TabIndex        =   9
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu PrinterChange 
         Caption         =   "Change Printer"
         Checked         =   -1  'True
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu FileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmREPRINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flxR As Integer
Dim flxC As Integer
Dim CCR As zclsCYEXCCR
Private Sub cmdByccr_Click()
    Call ByCCR
    SST.Tab = 1
    Call AnotherCCR
    txtTab(1).SetFocus
End Sub
Private Sub cmdByref_Click()
    Call ByReference
    SST.Tab = 0
    Call AnotherReference
    txtTab(0).SetFocus
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdExit2_Click()
    Unload Me
End Sub
Private Sub cmdPrint_Click()
' ** Printing
    Call ReadAndPrintFlxRef
    Call AnotherReference
End Sub
Private Sub cmdPrint2_Click()
    Call PrintCCR
    Call AnotherCCR
End Sub
Private Sub cmdPrinter_Click()
    frmPrinter.Show vbModal
    SB.Panels(3).Text = Printer.DeviceName
End Sub
Private Sub Command3_Click()
    frmPrinter.Show vbModal
End Sub
Private Sub Command4_Click()
    Unload Me
End Sub
Private Sub cmdPrinter2_Click()
    frmPrinter.Show vbModal
    SB.Panels(3).Text = Printer.DeviceName
End Sub
Private Sub cmdSearch_Click()
    Call AnotherReference
End Sub
Private Sub cmdSearch2_Click()
    Call AnotherCCR
End Sub
Private Sub flxRef_KeyPress(KeyAscii As Integer)
    With flxRef
        If KeyAscii <> 13 Then
            If KeyAscii <> 8 Then
                If UCase(Chr(KeyAscii)) = "1" Then
                    .TextMatrix(.Row, 0) = "1"
                    If .Row < (flxR - 1) Then
                        .Row = .Row + 1
                        .Col = 0
                        .ColSel = 6
                    End If
                Else
                    Beep
                    KeyAscii = 0
                End If
            Else
                .TextMatrix(.Row, 0) = " "
            End If
        End If
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If SST.Tab = 0 Then
        Select Case KeyCode
            Case vbKeyF3
                If cmdExit.Enabled Then
                    Call cmdExit_Click
                End If
            Case vbKeyF5
                If cmdPrinter.Enabled Then
                    Call cmdPrinter_Click
                End If
            Case vbKeyF6
                If cmdSearch.Enabled Then
                    Call cmdSearch_Click
                End If
            Case vbKeyF7
                If cmdPrint.Enabled Then
                    Call cmdPrint_Click
                End If
            Case vbKeyF11
                If cmdByccr.Enabled Then
                    Call cmdByccr_Click
                End If
        End Select
    Else
        Select Case KeyCode
            Case vbKeyF3
                If cmdExit2.Enabled Then
                    Call cmdExit2_Click
                End If
            Case vbKeyF5
                If cmdPrinter2.Enabled Then
                    Call cmdPrinter2_Click
                End If
            Case vbKeyF6
                If cmdSearch2.Enabled Then
                    Call cmdSearch2_Click
                End If
            Case vbKeyF7
                If cmdPrint2.Enabled Then
                    Call cmdPrint2_Click
                End If
            Case vbKeyF11
                If cmdByref.Enabled Then
                    Call cmdByref_Click
                End If
        End Select
    End If
End Sub

Private Sub Form_Load()
    Dim info As Recordset
    DE.GetUserInfo
    Set info = DE.rsGetUserInfo
    With info
        SB.Panels(1).Text = .Fields("Workstation")
        SB.Panels(2).Text = gUserid
        SB.Panels(3).Text = Printer.DeviceName
    End With
    info.Close
    Set info = Nothing
    Set CCR = New zclsCYEXCCR
    SST.Tab = 0
    Call ByReference
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set CCR = Nothing
End Sub
Private Sub txtTab_Change(Index As Integer)
    txtReference.Caption = " "
End Sub
Private Sub txtTab_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    If Index = 0 Then
        If FillGridByReference Then
            Tab0 (False)
            Tab1 (False)
            cmdPrint.Enabled = True
            cmdPrint.TabStop = True
            cmdSearch.Enabled = True
            cmdSearch.TabStop = True
            flxRef.Enabled = True
            flxRef.TabStop = True
            flxRef.SetFocus
        Else
            MsgBox " No Records Found ! ", vbExclamation + vbOKOnly, "Search Result"
            Tab0 (False)
            Tab1 (False)
            txtTab(0).TabStop = True
            txtTab(0).Enabled = True
            txtTab(0).SetFocus
        End If
    Else
        If FillGridByCCR Then
            Tab0 (False)
            Tab1 (False)
            cmdPrint2.Enabled = True
            cmdPrint2.TabStop = True
            cmdSearch2.Enabled = True
            cmdSearch2.TabStop = True
            flxCCR.Enabled = True
            flxCCR.TabStop = True
            flxCCR.SetFocus
        Else
            MsgBox " No Records Found ! ", vbExclamation + vbOKOnly, "Search Result"
            Tab0 (False)
            Tab1 (False)
            txtTab(1).TabStop = True
            txtTab(1).Enabled = True
            txtTab(1).SetFocus
        End If
    End If
End If
End Sub
Public Function FillGridByReference() As Boolean
    FillGridByReference = False
    Dim rs As Recordset
    DE.GetRefDetails CLng(Trim(txtTab(0).Text))
    Set rs = DE.rsGetRefDetails
    flxRef.Clear
    flxRef.Rows = 2
    flxRef.Cols = 7
    flxRef.FormatString = "Opt |Reference|Sequence| CCR No. | Exporter Name     | Broker Name      |  Date           "
    flxR = 1
    If rs.RecordCount > 0 Then
        With rs
            Do While Not .EOF
                If flxR > 1 Then
                    flxRef.AddItem " "
                End If
                flxRef.TextMatrix(flxR, 0) = " "
                flxRef.TextMatrix(flxR, 1) = .Fields("refnum") & ""
                flxRef.TextMatrix(flxR, 2) = .Fields("seqnum") & ""
                flxRef.TextMatrix(flxR, 3) = .Fields("ccrnum") & ""
                flxRef.TextMatrix(flxR, 4) = .Fields("exprtr") & ""
                flxRef.TextMatrix(flxR, 5) = .Fields("broker") & ""
                flxRef.TextMatrix(flxR, 6) = Format(.Fields("sysdttm") & "", "YYYY-MM-DD hh:nn")
                flxR = flxR + 1
                .MoveNext
            Loop
            FillGridByReference = True
        End With
    Else
        FillGridByReference = False
    End If
    rs.Close
    Set rs = Nothing
    flxRef.Row = 1
    flxRef.Col = 0
    flxRef.ColSel = 6
End Function
Public Sub Tab0(Mode As Boolean)
    txtTab(0).TabStop = Mode
    txtTab(0).Enabled = Mode
    flxRef.TabStop = Mode
    flxRef.Enabled = Mode
    cmdPrinter.TabStop = Mode
    cmdPrint.TabStop = Mode
    cmdPrint.Enabled = Mode
    cmdSearch.TabStop = Mode
    cmdSearch.Enabled = Mode
    cmdExit.TabStop = Mode
End Sub
Public Sub Tab1(Mode As Boolean)
    txtTab(1).TabStop = Mode
    txtTab(1).Enabled = Mode
    flxCCR.Enabled = Mode
    flxCCR.TabStop = Mode
    cmdPrinter2.TabStop = Mode
    cmdPrint2.TabStop = Mode
    cmdPrint2.Enabled = Mode
    cmdSearch2.TabStop = Mode
    cmdSearch2.Enabled = Mode
    cmdExit2.TabStop = Mode
End Sub
Public Sub ByReference()
    Tab0 (False)
    Tab1 (False)
    txtTab(0).TabStop = True
    txtTab(0).Enabled = True
End Sub
Public Sub ByCCR()
    Tab0 (False)
    Tab1 (False)
    txtTab(1).TabStop = True
    txtTab(1).Enabled = True
End Sub
Public Function ReadAndPrintFlxRef() As Boolean
    Dim x As Integer
    If flxR > 1 Then
        With flxRef
            For x = 1 To (flxR - 1)
                If Trim(.TextMatrix(x, 0)) = "1" Then
                    CCR.CCRNumber = CLng(Trim(.TextMatrix(x, 3)))
                    CCR.PrintCCR CLng(Trim(.TextMatrix(x, 1)))
                    .TextMatrix(x, 0) = " "
                End If
            Next
        End With
    End If
    ReadAndPrintFlxRef = True
End Function
Public Sub AnotherReference()
    flxRef.Clear
    flxRef.Rows = 2
    flxRef.Cols = 7
    flxRef.FormatString = "Opt |Reference|Sequence| CCR No. | Exporter Name     | Broker Name      |  Date           "
    Tab0 (False)
    Tab1 (False)
    txtTab(0).Enabled = True
    txtTab(0).TabStop = True
    txtTab(0).SetFocus
End Sub
Private Function FillGridByCCR() As Boolean
    FillGridByCCR = False
    Dim rs As Recordset
    DE.GetDetails CLng(Trim(txtTab(1).Text))
    Set rs = DE.rsGetDetails
    flxCCR.Clear
    flxCCR.Rows = 2
    flxCCR.Cols = 6
    flxCCR.FormatString = "No. | Container No. | Size |     Exporter Name     |      Broker Name   |  Date & Time     "
    flxC = 1
    If rs.RecordCount > 0 Then
        With rs
            txtReference.Caption = Trim(UCase(.Fields("refnum") & ""))
            Do While Not .EOF
                If flxC > 1 Then
                    flxCCR.AddItem " "
                End If
                flxCCR.TextMatrix(flxC, 0) = .Fields("itmnum") & ""
                flxCCR.TextMatrix(flxC, 1) = .Fields("cntnum") & ""
                flxCCR.TextMatrix(flxC, 2) = .Fields("cntsze") & ""
                flxCCR.TextMatrix(flxC, 3) = .Fields("exprtr") & ""
                flxCCR.TextMatrix(flxC, 4) = .Fields("broker") & ""
                flxCCR.TextMatrix(flxC, 5) = Format(.Fields("sysdttm") & "", "YYYY-MM-DD hh:nn")
                flxC = flxC + 1
                .MoveNext
            Loop
            FillGridByCCR = True
        End With
    Else
    FillGridByCCR = False
    End If
    rs.Close
    Set rs = Nothing
    flxCCR.Row = 1
    flxCCR.Col = 0
    flxCCR.ColSel = 5
End Function
Public Sub AnotherCCR()
    flxCCR.Clear
    flxCCR.Rows = 2
    flxCCR.Cols = 6
    flxCCR.FormatString = "No. | Container No. | Size |     Exporter Name     |      Broker Name   |  Date & Time     "
    Tab0 (False)
    Tab1 (False)
    txtTab(1).Enabled = True
    txtTab(1).TabStop = True
    txtTab(1).SetFocus
End Sub
Private Sub PrintCCR()
    CCR.CCRNumber = CLng(Trim(txtTab(1).Text))
    CCR.PrintCCR (CLng(Trim(txtReference.Caption)))
End Sub
Private Sub txtTab_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 8 Then
        If Not IsNumeric(Chr(KeyAscii)) Then
            Beep
            KeyAscii = 0
        End If
    End If
End Sub
