VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "CY Invoice Payment"
   ClientHeight    =   10815
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   15045
   Icon            =   "frmORMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10815
   ScaleWidth      =   15045
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   8760
      Width           =   14655
      Begin VB.CommandButton cmd_inventry 
         Caption         =   "F4 - INVOICE ENTRY"
         Height          =   975
         Left            =   600
         Picture         =   "frmORMain.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmd_adj 
         Caption         =   "F5 - INVOICE ADJ"
         Height          =   975
         Left            =   3360
         Picture         =   "frmORMain.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "F3 - EXIT"
         Height          =   975
         Left            =   6120
         Picture         =   "frmORMain.frx":09CE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cmbcust 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   420
         Left            =   10080
         MouseIcon       =   "frmORMain.frx":0E10
         MousePointer    =   99  'Custom
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Customer List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   10080
         TabIndex        =   8
         Top             =   360
         Width           =   4335
      End
   End
   Begin ComctlLib.StatusBar staStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   10440
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   18415
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Computer Name"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "User"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "2:33 PM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grd_InvList 
      Height          =   855
      Left            =   50
      TabIndex        =   4
      Top             =   840
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   1508
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ForeColor       =   0
      ForeColorFixed  =   65535
      BackColorSel    =   8454016
      ForeColorSel    =   16711680
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "List of Unpaid Bills"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   14775
   End
   Begin VB.Menu mneInvoice 
      Caption         =   "Menu"
      Index           =   0
      Begin VB.Menu mnePayment 
         Caption         =   "Invoice &Payment"
         Index           =   0
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnePayment 
         Caption         =   "Invoice Payment &Adjustment"
         Index           =   1
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnePayment 
         Caption         =   "&View Payments Made"
         Index           =   2
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnePayment 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnePayment 
         Caption         =   "E&xit"
         Index           =   4
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PROGRAM    : Invoice Payment for Subic
'Version    : 1.1
'Processing : Payment Processing, Invoice Adjustment, Applying of Available Amt as payment
'Programmer : Raquel Ocampo
'Date       : Feb. 2001



Option Explicit

Private Sub frm_shortcutkey(ByVal pKeyCode As Long)
    Select Case pKeyCode
        Case vbKeyF3
            Call cmdexit_Click
        Case vbKeyF4
            Call cmd_inventry_Click
        Case vbKeyF5
            Call cmd_adj_Click
        Case vbKeyF12 ' show List of Payments
          frmListORpayment.Show vbModal
    End Select
End Sub


Private Sub cmbcust_Click()
  Call FilterRecordset(cmbcust.List(cmbcust.ListIndex))
  Call Inialized_Grid
  Call List_UnpaidBills
End Sub

Private Sub cmbcust_GotFocus()
    SendKeys "%{DOWN}", False
    DoEvents
End Sub

Private Sub cmd_adj_Click()
    FrmInvAdjust.Show vbModal
    Call FilterRecordset(cmbcust.List(cmbcust.ListIndex))
    Call Inialized_Grid
    Call List_UnpaidBills
End Sub

Private Sub cmd_adj_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
            Case vbKeyF3, vbKeyF4, vbKeyF5
                Call frm_shortcutkey(KeyCode)
             End Select
End Sub

Private Sub cmd_inventry_Click()
    frmInvEntry.Show vbModal
    Call FilterRecordset(cmbcust.List(cmbcust.ListIndex))
    Call List_UnpaidBills
End Sub

Private Sub cmd_inventry_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
            Case vbKeyF3, vbKeyF4, vbKeyF5
                Call frm_shortcutkey(KeyCode)
             End Select
End Sub

Private Sub cmdexit_Click()
      Unload Me
End Sub

Private Sub cmdexit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyF3, vbKeyF4, vbKeyF5
                Call frm_shortcutkey(KeyCode)
             End Select
             
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF12
                Call frm_shortcutkey(KeyCode)
    End Select
End Sub

Private Sub Form_Load()
    Set rsUnsettled = New ADODB.Recordset
    Call InitializeWindow
    Call list_customer
    Call FilterRecordset("NONE")
    Call List_UnpaidBills
End Sub

Private Sub InitializeWindow()
  With staStatus
     .Panels(1).Text = ""
     .Panels(2).Text = UCase(zCurrentComputer)
     .Panels(3).Text = UCase(zCurrentUser)
     .Panels(4).ToolTipText = Format(Date, "dddd,mm/dd/yyyy")
  End With
   Label7.Top = 240
   Label7.Left = 100
   Label7.Width = Screen.Width - 200
   Me.WindowState = 2 'Maximize

  With grd_InvList
       .Cols = 6
       .Rows = 2
       .Left = Label7.Left
       .Top = 800
       .Width = Label7.Width
       .Height = 7815
  End With
End Sub

Private Sub list_customer()
    Dim cust As New ADODB.Recordset
    
     cust.Open "select distinct Upper(cusnam) as name from customer ", gcnnBilling, , , adCmdText
        cmbcust.AddItem "None"
        Do While Not cust.EOF
           cmbcust.AddItem UCase(cust!Name)
           cust.MoveNext
        Loop
        cmbcust.Text = cmbcust.List(0)
        cust.Close
        Set cust = Nothing
End Sub

 

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If MsgBox("Exit Invoice Payment Window?", vbOKCancel + vbQuestion, "Exit?") = vbOK Then
        Cancel = 0
        Set rsUnsettled = Nothing
        Unload Me
  Else
      Cancel = 1
     grd_InvList.SetFocus
  End If
End Sub

Private Sub mnePayment_Click(Index As Integer)
 Select Case Index
   Case 0 ' Invoice Payment Entry
        frmInvEntry.Show vbModal
   Case 1 ' Invoice Adjustment
        FrmInvAdjust.Show vbModal
   Case 2 ' View Payments Made
         frmListORpayment.Show vbModal
   Case 4 ' Exit
      Unload Me
 End Select
End Sub
