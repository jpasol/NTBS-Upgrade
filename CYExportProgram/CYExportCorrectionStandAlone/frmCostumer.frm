VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCostumer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Costumer Listing"
   ClientHeight    =   6390
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9090
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
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
   ScaleHeight     =   6390
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCode 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "F12 - Cancel"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5640
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid flxCustomer 
      Height          =   4575
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8070
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   400
      BackColor       =   16777215
      ForeColor       =   0
      GridColor       =   8388608
      AllowBigSelection=   0   'False
      GridLines       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   120
      MaxLength       =   50
      TabIndex        =   0
      Top             =   480
      Width           =   5895
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Enter - OK"
      BeginProperty Font 
         Name            =   "IBM3270 - 1254"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   -120
      TabIndex        =   3
      Top             =   -120
      Width           =   9135
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6240
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "IBM3270 - 1254"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   5895
      End
   End
End
Attribute VB_Name = "frmCostumer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TransCancelled As Boolean
Public flxGotFocus As Boolean
Public fillingCustomer As Boolean
Public ListCount As Long
Public CustomerCode As Long
Public CustomerName As String
Private Sub cmdCancel_Click()
    frmCostumer.CustomerCode = 0
    frmCostumer.CustomerName = " "
    TransCancelled = True
    Me.Hide
End Sub
Private Sub flxCustomer_GotFocus()
    flxGotFocus = True
    Call flxfocusedSetting
    flxCustomer.TopRow = flxCustomer.Row
    txtName.Text = flxCustomer.TextMatrix(flxCustomer.Row, 1)
    txtCode.Text = flxCustomer.TextMatrix(flxCustomer.Row, 2)
End Sub
Private Sub flxCustomer_LostFocus()
    flxGotFocus = False
    Call flxUnfocusedSetting
End Sub
Private Sub flxCustomer_RowColChange()
'    If Not fillingCustomer Then
    If flxGotFocus Then
        If TypeOf Screen.ActiveControl Is MSFlexGrid Then
            txtName.Text = flxCustomer.TextMatrix(flxCustomer.Row, 1)
            txtCode.Text = flxCustomer.TextMatrix(flxCustomer.Row, 2)
        End If
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF12
            Call cmdCancel_Click
        Case 13
            Call OKButton_Click
    End Select
End Sub
Private Sub Form_Load()
    flxCustomer.ScrollTrack = True
    'Call flxUnfocusedSetting
    flxGotFocus = False
    flxCustomer.BackColorSel = &HC0C0C0
    flxCustomer.ForeColorSel = &H8000000E
End Sub
Private Sub OKButton_Click()
    CustomerCode = Val(txtCode.Text)
    CustomerName = txtName.Text
    TransCancelled = False
    Me.Hide
End Sub
Public Sub SmartType(c As Object, n As Integer)
    Dim i As Integer
    Dim l As Integer
    Dim s As Integer
    Dim t As String
    Dim q As String
    
    s = txtName.SelStart
    l = txtName.SelLength
    t = txtName.Text
    q = txtCode.Text
    
    t = Left(t, s) & Chr(n) & Right(t, Len(t) - s)
    s = s + 1
    i = 0
    
    Do While (i < ListCount - 1) _
        And StrComp(Left(t, s), Left(flxCustomer.TextMatrix(i, 1), s), vbTextCompare) <> 0
        i = i + 1
    Loop
    If UCase(Left(t, s)) = UCase(Left(flxCustomer.TextMatrix(i, 1), s)) Then
        t = Trim(flxCustomer.TextMatrix(i, 1)) ' Combo1.List(i)
        q = Trim(flxCustomer.TextMatrix(i, 2))
        l = Len(t) - s
    Else
        t = Left(t, s)
        q = " "
        l = 0
    End If

    Call flxUnfocusedSetting
    
    txtName.Text = t
    txtCode.Text = q
    txtName.SelStart = s
    txtName.SelLength = l
    flxCustomer.TopRow = i
    flxCustomer.Row = i
    flxCustomer.Col = 0
    flxCustomer.ColSel = 2
End Sub
Public Sub SmartTypeCode(c As Object, n As Integer)
    Dim i As Integer
    Dim l As Integer
    Dim s As Integer
    Dim t As String
    Dim q As String
    
    s = txtCode.SelStart
    l = txtCode.SelLength
    t = txtCode.Text
    q = txtName.Text
    t = Left(t, s) & Chr(n) & Right(t, Len(t) - s)
    s = s + 1
    i = 0
    Do While (i < ListCount - 1) _
        And StrComp(Left(t, s), Left(flxCustomer.TextMatrix(i, 2), s), vbTextCompare) <> 0
        i = i + 1
    Loop
    If UCase(Left(t, s)) = UCase(Left(flxCustomer.TextMatrix(i, 2), s)) Then
        t = Trim(flxCustomer.TextMatrix(i, 2)) ' Combo1.List(i)
        q = Trim(flxCustomer.TextMatrix(i, 1))
        l = Len(t) - s
    Else
        t = Left(t, s)
        q = " "
        l = 0
    End If
    
    Call flxUnfocusedSetting
    
    txtCode.Text = t
    txtName.Text = q
    txtCode.SelStart = s
    txtCode.SelLength = l
    flxCustomer.TopRow = i
    flxCustomer.Row = i
    flxCustomer.Col = 0
    flxCustomer.ColSel = 2
End Sub
Private Sub txtCode_GotFocus()
    txtCode.BackColor = &H8000000E
    
    txtCode.SelStart = 0
    txtCode.SelLength = Len(Trim(txtName.Text))
End Sub
Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40
            SendKeys "{Tab}", True
        Case 38
            SendKeys "+{Tab}", True
    End Select
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) >= " " And Chr(KeyAscii) <= "~" Then
        SmartTypeCode txtCode, KeyAscii
        KeyAscii = 0
    End If
End Sub

Private Sub txtCode_LostFocus()
    txtCode.BackColor = &H8000000F
End Sub
Private Sub txtName_GotFocus()
    txtName.BackColor = &H8000000E
    txtName.Refresh
    txtName.SelStart = 0
    txtName.SelLength = Len(Trim(txtName.Text))
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 40
            SendKeys "{Tab}", True
        Case 38
            SendKeys "+{Tab}", True
    End Select
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) >= " " And Chr(KeyAscii) <= "~" Then
        SmartType txtName, KeyAscii
        KeyAscii = 0
    End If
End Sub
Public Sub flxUnfocusedSetting()
    flxCustomer.BackColorSel = &HC0C0C0
    flxCustomer.ForeColorSel = &H8000000E
End Sub
Public Sub flxfocusedSetting()
    flxCustomer.BackColorSel = &H8000000D
    flxCustomer.ForeColorSel = &H8000000E
End Sub
Private Sub txtName_LostFocus()
    txtName.BackColor = &H8000000F
    txtName.Refresh
End Sub
