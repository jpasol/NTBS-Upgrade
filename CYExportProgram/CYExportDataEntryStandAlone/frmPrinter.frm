VERSION 5.00
Begin VB.Form frmPrinter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Printer Selection"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8415
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8295
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   8415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "F3 - &Cancel"
      Height          =   615
      Left            =   4680
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ENTER - &OK"
      Height          =   615
      Left            =   6600
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Use the <Up> and <Down> Arrow Keys to select a printer"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Default Printer"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "frmPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PreviousPrinter As Integer
Private Sub cmbPrinter_Click()
    Set Printer = Printers(cmbPrinter.ListIndex)
    Printer.Orientation = 2 'landscape
    PrinterRef = cmbPrinter.ListIndex
End Sub
Private Sub cmbPrinter_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Call cmdCancel_Click
        Case vbKeyReturn
            Call cmdOk_Click
    End Select
End Sub
Private Sub cmdCancel_Click()
    Set Printer = Printers(PreviousPrinter)
    Unload Me
End Sub
Private Sub cmdOk_Click()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Call cmdCancel_Click
        Case vbKeyReturn
            Call cmdOk_Click
    End Select
End Sub
Private Sub Form_Load()
    Dim Pr As Printer
    Dim ref As Integer
    Dim refTouse As Integer
    Dim strRef As String * 2
    ref = 0
    For Each Pr In Printers
        strRef = Str(ref + 1)
        cmbPrinter.AddItem strRef & "| " & Pr.DeviceName

        If Pr.DeviceName = Printer.DeviceName Then
            refTouse = ref
        End If
        ref = ref + 1
    Next Pr
    cmbPrinter.ListIndex = refTouse
    PreviousPrinter = refTouse
End Sub
