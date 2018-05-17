VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1170
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   11235
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBackground 
      Height          =   555
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   1635
      TabIndex        =   3
      Top             =   0
      Width           =   1695
   End
   Begin VB.PictureBox picMainStatus 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11235
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   675
      Width           =   11235
      Begin VB.PictureBox picSideBar 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3555
         Left            =   0
         ScaleHeight     =   3555
         ScaleWidth      =   255
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   -75
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Timer tmrClock 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   9525
         Top             =   0
      End
      Begin prjBilling.uAnimButton btnStart 
         Height          =   420
         Left            =   15
         TabIndex        =   2
         Top             =   35
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   741
         Interval        =   50
      End
      Begin VB.Label lblStatusComputer 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4575
         TabIndex        =   8
         Top             =   50
         Width           =   1455
      End
      Begin VB.Label lblStatusTime 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8175
         TabIndex        =   7
         Top             =   50
         Width           =   1005
      End
      Begin VB.Label lblStatusDate 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6585
         TabIndex        =   6
         Top             =   50
         Width           =   1455
      End
      Begin VB.Label lblStatusUser 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   380
         Left            =   3000
         TabIndex        =   5
         Top             =   50
         Width           =   1455
      End
      Begin VB.Label lblStatusMsg 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   380
         Left            =   1560
         TabIndex        =   4
         Top             =   50
         Width           =   1335
      End
      Begin VB.Image imgStart 
         Height          =   7500
         Left            =   0
         Picture         =   "Main.frx":0000
         Top             =   0
         Visible         =   0   'False
         Width           =   1260
      End
   End
   Begin MSComctlLib.ImageList ilsMenuIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":A852
            Key             =   """DataEntry"""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":A9AC
            Key             =   """ShutDown"""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":D15E
            Key             =   """Timer"""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":D2B8
            Key             =   """NetPrinter"""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":FA6A
            Key             =   """Printer"""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1221C
            Key             =   """FindInDoc"""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":149CE
            Key             =   """LogOff"""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":14E20
            Key             =   """TechSupport"""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":15272
            Key             =   """Help"""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":153CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":15526
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":15840
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":15B5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":15CB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1677E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1683F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":16999
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Used to transfer side logo onto the owner-draw menu:
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' To ensure we shut-down when choose close whilst the button is down:
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_CLOSE = &HF060

' The popup menu object:
Private WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1

Dim maxhgt As Long, maxwid As Long
Dim pwid, phgt As Integer

Private Sub Form_Load()

Set m_cMenu = New cPopupMenu
    ' setup status bar
    Me.Icon = ilsMenuIcons.ListImages(17).Picture
    With picBackground
        .Picture = LoadPicture(App.Path & "\" & gBackground)
        .ScaleMode = 3
        .Visible = False
        .AutoSize = True
        .AutoRedraw = True
        pwid = .ScaleWidth
        phgt = .ScaleHeight
    End With
    lzShowClock
    lzShowUser
    tmrClock.Enabled = True
    ' setup menu
    zSetupMenu
    Me.Refresh
End Sub

Private Sub Form_Resize()
Dim lmsglen As Long
    Me.WindowState = vbMaximized
    maxhgt = Height \ Screen.TwipsPerPixelY
    maxwid = Width \ Screen.TwipsPerPixelX
   
    lblStatusTime.left = Me.Width - lblStatusTime.Width - 170
    lblStatusDate.left = lblStatusTime.left - lblStatusDate.Width - 30
    lblStatusComputer.left = lblStatusDate.left - lblStatusComputer.Width - 30
    lblStatusUser.left = lblStatusComputer.left - lblStatusUser.Width - 30
    lblStatusMsg.left = btnStart.Width + 30
    lmsglen = lblStatusUser.left - lblStatusMsg.left - 30
    If lmsglen > 0 Then lblStatusMsg.Width = lmsglen
End Sub

Private Sub Form_Paint()
Dim i%, j%, x%
Dim phDC&, frmhdc&
    phDC& = picBackground.hDC
    frmhdc& = Me.hDC
    For j% = 0 To maxhgt Step phgt
        For i% = 0 To maxwid Step pwid
            x% = BitBlt(frmhdc&, i%, j%, pwid, phgt, phDC&, 0, 0, &HCC0020)
        Next
    Next
End Sub

Private Sub Form_Terminate()
    If gbConnected Then
        On Error Resume Next
        gcnnBilling.Close
    End If
    End
End Sub

Private Sub m_cMenu_DrawItem(ByVal hDC As Long, ByVal lMenuIndex As Long, lLeft As Long, lTop As Long, lRight As Long, lBottom As Long, ByVal bSelected As Boolean, ByVal bChecked As Boolean, ByVal bDisabled As Boolean, bDoDefault As Boolean)
Dim lW As Long
   ' The DrawItem event for Owner Draw menu items either allows you
   ' to draw the entire item, or just to do some new drawing then
   ' let the standard method do its stuff.  This is useful if you
   ' want to add a graphic to the left or right of the menu item.

   ' Here we draw the relevant part of the side bar
   ' logo to the left of the menu then offset the
   ' left position so the rest of the menu draws
   ' after it:
   lW = picSideBar.Width \ Screen.TwipsPerPixelX
   BitBlt hDC, lLeft, lTop, lW, lBottom - lTop, picSideBar.hDC, 0, lTop, vbSrcCopy
   lLeft = lLeft + lW + 1
   bDoDefault = True
End Sub

Private Sub m_cMenu_ItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
   ' Show the user what's been highlighted.  In a real application you
   ' would want to make these captions more descriptive:
      
    'lblInfo.Caption = m_cMenu.Caption(ItemNumber)
   
End Sub

Private Sub m_cMenu_MeasureItem(ByVal lMenuIndex As Long, lWidth As Long, lHeight As Long)
   ' When a menu item is owner-draw, it will raise this event to request
   ' its size.  lWidth as lHeight will be already filled in with the
   ' size of the menu item as it would be if the standard drawing method
   ' was used.
   
   ' Here we check if the item being measured is in the main pop-up menu;
   ' if it is we add the width to accommodate the side bar logo:
   If m_cMenu.hMenu(1) = m_cMenu.hMenu(lMenuIndex) Then
      ' Add the side bar width:
      lWidth = lWidth + picSideBar.Width \ Screen.TwipsPerPixelX + 1
   End If
End Sub

Private Sub mnuMainFileClose_Click()
    If Forms.Count > 1 Then
        Unload Forms(0).ActiveForm
    Else
        Unload Me
    End If
End Sub

Private Sub tmrClock_Timer()
    lzShowClock
    DoEvents
End Sub

Private Sub lzShowClock()
    lblStatusDate.Caption = Format(Now, "mm/dd/yyyy"): lblStatusDate.Refresh
    lblStatusTime.Caption = Format(Now, "hh:mm"): lblStatusTime.Refresh
End Sub

Private Sub lzShowUser()
    lblStatusUser.Caption = gUserID: lblStatusUser.Refresh
    lblStatusComputer.Caption = gComputer: lblStatusComputer.Refresh
End Sub

Private Sub zSetupMenu()
Dim i As Long
Dim iI As Long
   
    ' Set up animated button:
    With btnStart
        .BackColor = &H80&
        .XCells = 1: .YCells = 20
        .DefaultCell = 1
        Set .Picture = imgStart.Picture
        '.CellSteps(10) = 10
        .Interval = 50
    End With

    ' Set up pop-up menu:
    Set m_cMenu = New cPopupMenu
'   Set m_cMenu = CreateObject("cPopupMenu")
    With m_cMenu
        ' Set up for cPopupMenu:
        .hWndOwner = Me.hwnd
        .ImageList = ilsMenuIcons
        .HeaderStyle = ecnmHeaderCaptionBar
        
        ' Now add the menu items.  The items in the main menu are all
        ' set to to OwnerDraw so we can add the side bar logo.  See
        ' further description in DrawItem and MeasureItem events.
        
        '=============
        ' Import
        '=============
        i = .AddItem("-Import"): .OwnerDraw(i) = True
        i = .AddItem("&Gatepass Issuance", , , , 10, , , "CYM"): .OwnerDraw(i) = True
        .AddItem "-Data Entry", , , i
        .AddItem "&Import Gatepass Issuance", , , i, 0, , , "CYMDE01"
        .AddItem "&Empty and Shutout Container", , , i, 0, , , "CYEDE01"
        .AddItem "&Correction and Cancellation", , , i, 0, , , "CYCDE01"
        '----------------------
        .AddItem "-Inquiry", , , i
        .AddItem "Te&ller Collection Inquiry", , , i, 5, , , "CYMIN02"
        .AddItem "Inquir&y by Container", , , i, 5, , , "CYMIN04"
        '----------------------
        .AddItem "-Reports", , , i
        .AddItem "Import Gatepass Reprinting", , , i, 4, , , "CYMPR01"
        .AddItem "&Assessor/Teller Turnover Report", , , i, 4, , , "CYMPR18"
        .AddItem "Daily Summary Report (Auditors Copy) per Reference", , , i, 4, , , "CYMPR13"
        .AddItem "Daily Summary Report (Auditors Copy) per CCR", , , i, 4, , , "CYMPR14"
        .AddItem "Import Gatepass Summary Report", , , i, 4, , , "CYMPR05"
        .AddItem "Cancelle&d Gatepass Report", , , i, 4, , , "CYMPR12"
        .AddItem "Stora&ge Collection Report", , , i, 4, , , "CYMSTOR"
        '.AddItem "Mont&hly Report", , , i, 4, , , "CYMPR24"
        .AddItem "Inquire by Gatepass", , , i, 4, , , "CYMIN10"
        .AddItem "Underguarantee Report", , , i, 4, , , "CYMPR25"
        .AddItem "Teller Cash/Checks Turn-Over Report", , , i, 4, , , "CYMTURN"
        '----------------------
        .AddItem "-File Maintenance", , , i
       .AddItem "Teller Gatepass Allocation", , , i, 12, , , "CYMALOC"
       .AddItem "Tariff Rates", , , i, 12, , , "CYMRATE"
       .AddItem "Customer Codes", , , i, 12, , , "CYMCUST"
        '=============
        ' EXPORT
        '=============
        i = .AddItem("-Export"): .OwnerDraw(i) = True
        i = .AddItem("&CY Export Transaction", , , , 10, , , "CCR"): .OwnerDraw(i) = True
        .AddItem "-Data Entry", , , i
        .AddItem "&Export CCR Issuance", , , i, 0, , , "CCRDE01"
        .AddItem "Export CCR Correction/Voiding", , , i, 0, , , "CCRDE08"
        '----------------------
        .AddItem "-Inquiry", , , i
        .AddItem "Teller Collection Inquiry", , , i, 5, , , "CCRIN01"
        .AddItem "Payment Transaction Inquiry", , , i, 5, , , "CCRIN02"
        '----------------------
        .AddItem "-Re-Printing", , , i
        .AddItem "CY Export CCR Reprinting", , , i, 4, , , "CCRPR01"
        .AddItem "Daily / Monthly Report", , , i, 4, , , "CCRRPT"
        '=============
        ' SPECIAL SERVICES
        '=============
        i = .AddItem("-Special Services"): .OwnerDraw(i) = True
        i = .AddItem("CY &SplSvcs Transaction", , , , 10, , , "CYS"): .OwnerDraw(i) = True
        .AddItem "-Data Entry", , , i
        .AddItem "&SplSvcs CCR Issuance", , , i, 0, , , "CCRCYS"
        .AddItem "&CCR Correction/Voiding", , , i, 0, , , "CCRCYC"
        .AddItem "-Printing", , , i
        .AddItem "Sp&lSvcs CCR Reprinting", , , i, 4, , , "CCRSPR"
        .AddItem "&Reports", , , i, 4, , , "CYSREP"
        
        '=============
        ' INVOICE
        '=============
        i = .AddItem("-Invoice"): .OwnerDraw(i) = True
        i = .AddItem("C&Y Invoice", , , , 10, , , "INV"): .OwnerDraw(i) = True
        .AddItem "-Data Entry", , , i
        .AddItem "&Invoice Issuance", , , i, 0, , , "INVDE01"
        .AddItem "Invoice &Payment", , , i, 0, , , "PAYINV"
        .AddItem "&Correction and Cancellation", , , i, 0, , , "INVDE02"
        .AddItem "-Printing", , , i
        .AddItem "&Reports", , , i, 4, , , "INVPR01"
        .AddItem "Re&printing", , , i, 4, , , "INVPR02"
        .AddItem " Invoice Paymemt Report", , , i, 4, , , "PAYRPT"
        
        '-------------------------------------------------------------------------
        i = .AddItem("-Help"): .OwnerDraw(i) = True
        i = .AddItem("&Help", , , , 8, , , "Help"): .OwnerDraw(i) = True
        i = .AddItem("&Tech Support", , , , 7, , , "TechSupport"): .OwnerDraw(i) = True
        
        i = .AddItem("-"): .OwnerDraw(i) = True
        i = .AddItem("&Change Printer", , , , 3, , , "PRINTER"): .OwnerDraw(i) = True
        i = .AddItem("&Log Out...", , , , 6, , , "LogOut"): .OwnerDraw(i) = True
        i = .AddItem("Sh&ut Down...", , , , 1, , , "ShutDown"): .OwnerDraw(i) = True
   
    End With
   
    ' Now prepare the side bar.
    ' Firstly, evaluate the menu item's height in the main menu:
    Dim lHeight As Long, LT As Long
    For i = 1 To m_cMenu.Count
        ' Check if item is in the main menu:
        If (m_cMenu.hMenu(i) = m_cMenu.hMenu(1)) Then
           ' Add the item:
           lHeight = lHeight + m_cMenu.MenuItemHeight(i)
           LT = LT + 1
        End If
    Next i
   
    ' We use a PictureBox to hold the side logo here for convenience,
    ' however, you could use CreateCompatibleDC and CreateCompatibleBitmap
    ' to create a memory DC to hold this to avoid having the extra control.
    picSideBar.Height = lHeight * Screen.TwipsPerPixelY
    ' Draw a gradient into it.  Here I stole the code directly from the
    ' SideLogo/Fonts at any angle project for simplicity:
    Dim c As New cLogo
    With c
        .DrawingObject = picSideBar
        .StartColor = &H3399&
        .EndColor = &H0&
        .Caption = " NT BILLING"
        ilsMenuIcons.ListImages(1).Draw 0, 0, 0
        .hImageList = ilsMenuIcons.hImageList
        .IconIndex = 14
        .Draw
    End With

End Sub

Private Sub btnStart_Click()
Dim lIndex As Long

Dim cCYMDE01 As Object
Dim cCYEDE01 As Object
Dim cCYCDE01 As Object
Dim cCYMIN02 As Object
Dim cCYMIN04 As Object
Dim cCYMPR01 As Object
Dim cCYMALOC As Object
Dim cCYMRATE As Object
Dim cCYMCUST As Object
Dim cCYMPR18 As Object
Dim cCYMPR07 As Object
Dim cCYMPR24 As Object
Dim cCYMPR12 As Object
Dim cCYMPR13 As Object
Dim cCYMPR14 As Object
Dim cCYMPR05 As Object
Dim cCYMPR11 As Object
Dim cCYMPR25 As Object
Dim cCYMSTOR As Object
Dim cCYMIN10 As Object
Dim cContEntry  As Object
Dim cCYMTURN As Object

Dim cCCRDE01 As Object
Dim cCCRDE08 As Object
Dim cCCRIN01 As Object
Dim cCCRIN02 As Object
Dim cCCRPR01 As Object
Dim cCCRRPT As Object

Dim cCCRCYS As Object  ' Special Services
Dim cCCRCYC As Object


Dim cINVDE01 As Object
Dim cINVDE02 As Object
Dim cINVPR01 As Object
Dim cINVPR02 As Object
Dim cINVPAY As Object
Dim cPAYRPT As Object
    ' Show the popup menu and get the item the user clicks:
    
  '  lIndex = m_cMenu.ShowPopupMenu(picMainStatus.left, picMainStatus.tOp, picMainStatus.left, picMainStatus.tOp, Me.ScaleWidth - picMainStatus.left - picMainStatus.Width, picMainStatus.tOp + picMainStatus.Height, False)
   lIndex = m_cMenu.ShowPopupMenu( _
        picMainStatus.left, picMainStatus.tOp, picMainStatus.left, picMainStatus.tOp, _
        Me.ScaleWidth - picMainStatus.left - picMainStatus.Width, picMainStatus.tOp, False)
'    lIndex = m_cMenu.ShowPopupMenu(picMainStatus.left, picMainStatus.tOp, , , , , False)

    If (lIndex > 0) Then
        Me.Refresh
        Select Case m_cMenu.ItemKey(lIndex)
'-------------------------------------------------------------------------------------
            Case "CYMDE01"
                Set cCYMDE01 = CreateObject("prjManifestCont.clsCYMDE01")
                With cCYMDE01
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(gbSupervisor)
                    Call .Disconnect
                End With
                Set cCYMDE01 = Nothing
            Case "CYEDE01"
                Set cCYEDE01 = CreateObject("prjEmptyCont.clsCYEDE01")
                With cCYEDE01
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(gbSupervisor)
                    Call .Disconnect
                End With
                Set cCYEDE01 = Nothing
            Case "CYCDE01"
                Set cCYCDE01 = CreateObject("prjCYCancelCorrect.clsCYMDE03")
                With cCYCDE01
                    Call .ConnectByStr(gConnStr)
                    Call .Execute
                    Call .Disconnect
                End With
                Set cCYCDE01 = Nothing
            Case "CYMIN02"
                Set cCYMIN02 = CreateObject("CYMTellerInquiry.clsCYMTellerInquiry")
                With cCYMIN02
                    Call .ConnectByStr(gConnStr)
                         .USERID = zCurrentUser()
                    Call .Execute
                    Call .Disconnect
                End With
                Set cCYMIN02 = Nothing
            Case "CYMIN04"
                Set cCYMIN04 = CreateObject("CYMGpassTrans.clsCYMGpassTrans")
                With cCYMIN04
                    Call .Execute
                End With
                Set cCYMIN04 = Nothing
            Case "CYMPR01"
                Set cCYMPR01 = CreateObject("prjReprint.clsCYMPR01")
                With cCYMPR01
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(gbSupervisor)
                    Call .Disconnect
                End With
                Set cCYMPR01 = Nothing
            Case "CYMPR18"
                Set cCYMPR18 = CreateObject("prjCYMReport.cCYSRPT")
                With cCYMPR18
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(0)
                    Call .Disconnect
                End With
                Set cCYMPR18 = Nothing
            Case "CYMPR07"
                Set cCYMPR07 = CreateObject("prjCYMReport.cCYSRPT")
                With cCYMPR07
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(8)
                    Call .Disconnect
                End With
                Set cCYMPR07 = Nothing
            Case "CYMPR24"
                Set cCYMPR24 = CreateObject("prjCYMReport.cCYSRPT")
                With cCYMPR24
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(6)
                    Call .Disconnect
                End With
                Set cCYMPR24 = Nothing
            Case "CYMPR12"
                Set cCYMPR12 = CreateObject("prjCYMReport.cCYSRPT")
                With cCYMPR12
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(4)
                    Call .Disconnect
                End With
                Set cCYMPR12 = Nothing
            Case "CYMSTOR"
                Set cCYMSTOR = CreateObject("prjCYMReport.cCYSRPT")
                With cCYMSTOR
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(5)
                    Call .Disconnect
                End With
                Set cCYMSTOR = Nothing
            Case "CYMPR13"
                 Set cCYMPR13 = CreateObject("prjCYMReport.cCYSRPT")
                With cCYMPR13
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(1)
                    Call .Disconnect
                End With
                Set cCYMPR13 = Nothing
            Case "CYMPR14"
                Set cCYMPR14 = CreateObject("prjCYMReport.cCYSRPT")
                With cCYMPR14
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(2)
                    Call .Disconnect
                End With
                Set cCYMPR14 = Nothing
            Case "CYMPR05"
                Set cCYMPR05 = CreateObject("prjCYMReport.cCYSRPT")
                With cCYMPR05
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(3)
                    Call .Disconnect
                End With
                Set cCYMPR05 = Nothing
             Case "CYMPR11"
                Set cCYMPR11 = CreateObject("prjCYMReport.cCYSRPT")
                With cCYMPR11
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(4)
                    Call .Disconnect
                End With
                Set cCYMPR11 = Nothing
             Case "CYMTURN"
                Set cCYMTURN = CreateObject("prjCYMReport.cCYSRPT")
                With cCYMTURN
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(8)
                    Call .Disconnect
                End With
                Set cCYMTURN = Nothing
'            Case "CCRFM03"
'                Set cCCRFM03 = CreateObject("CYRatesMaintenance.clsCYRates")
'                With cCCRFM03
'                    Call .ConnectByStr(gConnStr)
'                    Call .Execute
'                    Call .Disconnect
'                End With
'                Set cCCRFM03 = Nothing
'            Case "CCRFM01"
'                Set cCCRFM01 = CreateObject("CCRAllocation.clsCCRAllocation")
'                With cCCRFM01
'                    Call .ConnectByStr(gConnStr)
'                    Call .Execute
'                    Call .Disconnect
'                End With
'                Set cCCRFM01 = Nothing
            Case "CYMCONTENTRY"
                Set cContEntry = CreateObject("prjCYMReport.cCYSRPT")
                With cContEntry
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(9)
                    Call .Disconnect
                End With
                Set cContEntry = Nothing
            Case "CYMIN10"
                Set cCYMIN10 = CreateObject("prjCYMReport.cCYSRPT")
                With cCYMIN10
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(6)
                    Call .Disconnect
                End With
                Set cCYMIN10 = Nothing
            Case "CYMPR25"
                Set cCYMPR25 = CreateObject("prjCYMReport.cCYSRPT")
                With cCYMPR25
                    Call .ConnectByStr(gConnStr)
                    Call .Execute(7)
                    Call .Disconnect
                End With
                Set cCYMPR25 = Nothing
            Case "CYMALOC"
                Set cCYMALOC = CreateObject("CCRAllocation.clsCCRAllocation")
                With cCYMALOC
                    Call .ConnectByStr(gConnStr)
                         .USERID = zCurrentUser()
                    Call .Execute
                    Call .Disconnect
                End With
                Set cCYMALOC = Nothing
            Case "CYMRATE"
                Set cCYMRATE = CreateObject("CYRatesMaintenance.clsCYRates")
                With cCYMRATE
                    Call .ConnectByStr(gConnStr)
                         .USERID = zCurrentUser()
                    Call .Execute
                    Call .Disconnect
                End With
                Set cCYMRATE = Nothing
            Case "CYMCUST"
                Set cCYMCUST = CreateObject("CustomerMaintenance.clsCustMaintenance")
                With cCYMCUST
                    Call .ConnectByStr(gConnStr)
                         .USERID = zCurrentUser()
                    Call .Execute
                    Call .Disconnect
                End With
            Set cCYMCUST = Nothing
'------------------------------------------------------------------------------
'  SPECIAL SERVICES
'------------------------------------------------------------------------------
            Case "CCRCYS"
                Set cCCRCYS = CreateObject("SubicCYSCCR.cCYSCCR")
                With cCCRCYS
                    Call .CCRSuper(gbSupervisor)
                    Call .ConnectByStr(gConnStr, gUserID)
                    Call .Execute
                    Call .Disconnect
                End With
                Set cCCRCYS = Nothing
            Case "CCRCYC"
                Set cCCRCYC = CreateObject("CYSCorrection.cCYSCorrection")
                With cCCRCYC
                    Call .ConnectByStr(gConnStr, gUserID)
                    Call .Execute
                    Call .Disconnect
                End With
                Set cCCRCYC = Nothing
            Case "CCRSPR"
                Set cCCRCYS = CreateObject("SubicCYSCCR.cCYSCCR")
                With cCCRCYS
                    Call .CCRSuper(gbSupervisor)
                    Call .ConnectByStr(gConnStr, gUserID)
                    Call .ReprintCCR
                    Call .Disconnect
                End With
                Set cCCRCYS = Nothing
            Case "CYSREP"
                Set cCCRCYS = CreateObject("CYSReports.cCYSRPT")
                With cCCRCYS
                         .USERID = zCurrentUser()
                    Call .Execute
                End With
                Set cCCRCYS = Nothing
'------------------------------------------------------------------------------
            Case "CCRDE01"
                Set cCCRDE01 = CreateObject("CCRDE06.clsCCRDE06")
                cCCRDE01.Supervisor = gbSupervisor
                cCCRDE01.USERID = zCurrentUser()
                cCCRDE01.ConnectByStr gConnStr
                cCCRDE01.Execute
                cCCRDE01.Disconnect
                Set cCCRDE01 = Nothing
            Case "CCRDE08"
                Set cCCRDE08 = CreateObject("CCRDE08.clsCCRDE08")
                cCCRDE08.USERID = zCurrentUser()
                cCCRDE08.ConnectByStr gConnStr
                cCCRDE08.Execute
                cCCRDE08.Disconnect
                Set cCCRDE08 = Nothing
            Case "CCRIN01"
                Set cCCRIN01 = CreateObject("CYXINQ02.clsCYXINQ01")
                cCCRIN01.USERID = zCurrentUser()
                cCCRIN01.Execute
                Set cCCRIN01 = Nothing
            Case "CCRIN02"
                Set cCCRIN02 = CreateObject("CYXINQ02.clsCYXINQ02")
                cCCRIN02.USERID = zCurrentUser()
                cCCRIN02.Execute
                Set cCCRIN02 = Nothing
            Case "CCRPR01"
                Set cCCRPR01 = CreateObject("ZCCRCYREPRT.ZclsCYEXCCR")
                cCCRPR01.USERID = zCurrentUser()
                cCCRPR01.Execute
                Set cCCRPR01 = Nothing
            Case "CCRRPT"
                Set cCCRRPT = CreateObject("zcCCRRPT.zclsCCRRPT")
                cCCRRPT.USERID = zCurrentUser()
                cCCRRPT.Execute
                Set cCCRRPT = Nothing

'------------------------------------------------------------------------------
            Case "INVDE01"
                On Error GoTo errINVDE01
                Set cINVDE01 = CreateObject("SubicINVDE01.clsSubicINVDE01")
                With cINVDE01
                    .ConnectByStr (gConnStr)
                    .USERID = zCurrentUser()
                    .Execute
                    .Disconnect
                End With
                Set cINVDE01 = Nothing
                Exit Sub
errINVDE01: MsgBox "Error creating object 'INVDE01'. Contact MIS for assistance."
            Exit Sub
            
            Case "PAYINV"
                On Error GoTo errPAYINV
                Set cINVPAY = CreateObject("INVPAYMENT.clsINVPAYMENT")
                With cINVPAY
                    .ConnectByStr (gConnStr)
                    .USERID = zCurrentUser()
                    .Execute
                    .Disconnect
                End With
                Set cINVPAY = Nothing
                Exit Sub
errPAYINV: MsgBox "Error creating object 'INVPAYMENT'. Contact MIS for assistance."
            Exit Sub
            
            Case "PAYRPT"
                On Error GoTo errINVRPT
                Set cPAYRPT = CreateObject("INVPAYREPORT.clsCOMMON")
                With cPAYRPT
                    .ConnectByStr (gConnStr)
                    .USERID = zCurrentUser()
                    .Execute
                    .Disconnect
                End With
                Set cPAYRPT = Nothing
                Exit Sub
errINVRPT: MsgBox "Error creating object 'PAYRPT'. Contact MIS for assistance."
            Exit Sub


            Case "INVDE02"
                On Error GoTo errINVDE02
                Set cINVDE02 = CreateObject("SubicINVCorrection.clsCYInvCorr")
                With cINVDE02
                    .ConnectByStr (gConnStr)
                    .USERID = zCurrentUser()
                    .Execute
                    .Disconnect
                End With
                Set cINVDE02 = Nothing
                Exit Sub
errINVDE02: MsgBox "Error creating object 'INVDE02'. Contact MIS for assistance."
            Exit Sub
                
            Case "INVPR01"
                On Error GoTo errINVPR01
                Set cINVPR01 = CreateObject("SubicINVReports.clsSubicINVReports")
                With cINVPR01
                    .USERID = zCurrentUser()
                    .Execute
                End With
                Set cINVPR01 = Nothing
                Exit Sub
errINVPR01: MsgBox "Error creating object 'INVPR01'. Contact MIS for assistance."
            Exit Sub
                
            Case "INVPR02"
                On Error GoTo errINVPR02
                Set cINVPR02 = CreateObject("SubicINVReprint.clsSubicINVReprint")
                cINVPR02.Execute
                Set cINVPR02 = Nothing
                Exit Sub
errINVPR02: MsgBox "Error creating object 'INVPR02'. Contact MIS for assistance."
            Exit Sub
          


            
'------------------------------------------------------------------------------
            Case "PRINTER"
                Call Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @2")
            Case "LogOut"
                PostMessage Me.hwnd, WM_SYSCOMMAND, SC_CLOSE, 0
                gShutDown = False
                Unload Me
            Case "ShutDown"
                If MsgBox("Are you sure you want to close the computer?", vbYesNo, "Shut Down System") = vbYes Then
                    ' If we unload here directly, we will have a problem
                    ' because the button code will not terminate.  sigh...
                    PostMessage Me.hwnd, WM_SYSCOMMAND, SC_CLOSE, 0
                    gShutDown = True
                    Unload Me
                End If
            Case Else
                MsgBox "Not yet installed...", vbOKOnly, "Missing Module"
        End Select
    Else
       ' lIndex=0 :: cancelled the menu.
    End If
End Sub

