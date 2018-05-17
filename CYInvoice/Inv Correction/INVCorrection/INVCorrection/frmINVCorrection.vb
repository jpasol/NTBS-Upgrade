Option Explicit On 

Public Class frmINVCorrection
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents grpTitle As System.Windows.Forms.GroupBox
    Friend WithEvents grpHeader As System.Windows.Forms.GroupBox
    Friend WithEvents grpDetails As System.Windows.Forms.GroupBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtORNum As System.Windows.Forms.TextBox
    Friend WithEvents btnGet As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents statPanelUser As System.Windows.Forms.StatusBarPanel
    Friend WithEvents statPanelDate As System.Windows.Forms.StatusBarPanel
    Friend WithEvents statPanelTime As System.Windows.Forms.StatusBarPanel
    Friend WithEvents statbarINVCorretion As System.Windows.Forms.StatusBar
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents picINVCorrection As System.Windows.Forms.PictureBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtCusCde As System.Windows.Forms.TextBox
    Friend WithEvents txtChequeBank2 As System.Windows.Forms.TextBox
    Friend WithEvents txtChequeBank1 As System.Windows.Forms.TextBox
    Friend WithEvents txtChequeNo2 As System.Windows.Forms.TextBox
    Friend WithEvents txtChequeNo1 As System.Windows.Forms.TextBox
    Friend WithEvents txtChequeAmt2 As System.Windows.Forms.TextBox
    Friend WithEvents txtChequeAmt1 As System.Windows.Forms.TextBox
    Friend WithEvents txtCreditAmt As System.Windows.Forms.TextBox
    Friend WithEvents txtCashAmt As System.Windows.Forms.TextBox
    Friend WithEvents txtAmtTotal As System.Windows.Forms.TextBox
    Friend WithEvents txtChequeAmtTotal As System.Windows.Forms.TextBox
    Friend WithEvents dgINVPayDtl As System.Windows.Forms.DataGrid
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents lblPayAmt As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtTeller As System.Windows.Forms.TextBox
    Friend WithEvents txtORDate As System.Windows.Forms.TextBox
    Friend WithEvents lblCustomerName As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmINVCorrection))
        Me.grpTitle = New System.Windows.Forms.GroupBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.picINVCorrection = New System.Windows.Forms.PictureBox
        Me.grpHeader = New System.Windows.Forms.GroupBox
        Me.lblCustomerName = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.txtORDate = New System.Windows.Forms.TextBox
        Me.txtTeller = New System.Windows.Forms.TextBox
        Me.lblPayAmt = New System.Windows.Forms.Label
        Me.btnClear = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnExit = New System.Windows.Forms.Button
        Me.txtCusCde = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtCreditAmt = New System.Windows.Forms.TextBox
        Me.txtCashAmt = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtAmtTotal = New System.Windows.Forms.TextBox
        Me.txtChequeAmtTotal = New System.Windows.Forms.TextBox
        Me.txtChequeBank2 = New System.Windows.Forms.TextBox
        Me.txtChequeBank1 = New System.Windows.Forms.TextBox
        Me.txtChequeNo2 = New System.Windows.Forms.TextBox
        Me.txtChequeNo1 = New System.Windows.Forms.TextBox
        Me.txtChequeAmt2 = New System.Windows.Forms.TextBox
        Me.txtChequeAmt1 = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnGet = New System.Windows.Forms.Button
        Me.txtORNum = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.grpDetails = New System.Windows.Forms.GroupBox
        Me.dgINVPayDtl = New System.Windows.Forms.DataGrid
        Me.statbarINVCorretion = New System.Windows.Forms.StatusBar
        Me.statPanelUser = New System.Windows.Forms.StatusBarPanel
        Me.statPanelDate = New System.Windows.Forms.StatusBarPanel
        Me.statPanelTime = New System.Windows.Forms.StatusBarPanel
        Me.grpTitle.SuspendLayout()
        Me.grpHeader.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.grpDetails.SuspendLayout()
        CType(Me.dgINVPayDtl, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.statPanelUser, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.statPanelDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.statPanelTime, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'grpTitle
        '
        Me.grpTitle.Controls.Add(Me.Label13)
        Me.grpTitle.Controls.Add(Me.picINVCorrection)
        Me.grpTitle.Location = New System.Drawing.Point(7, 2)
        Me.grpTitle.Name = "grpTitle"
        Me.grpTitle.Size = New System.Drawing.Size(823, 89)
        Me.grpTitle.TabIndex = 0
        Me.grpTitle.TabStop = False
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Font = New System.Drawing.Font("Papyrus", 33.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.SaddleBrown
        Me.Label13.Location = New System.Drawing.Point(400, 14)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(408, 64)
        Me.Label13.TabIndex = 1
        Me.Label13.Text = "Invoice Correction"
        '
        'picINVCorrection
        '
        Me.picINVCorrection.Image = CType(resources.GetObject("picINVCorrection.Image"), System.Drawing.Image)
        Me.picINVCorrection.Location = New System.Drawing.Point(8, 13)
        Me.picINVCorrection.Name = "picINVCorrection"
        Me.picINVCorrection.Size = New System.Drawing.Size(360, 72)
        Me.picINVCorrection.TabIndex = 0
        Me.picINVCorrection.TabStop = False
        '
        'grpHeader
        '
        Me.grpHeader.Controls.Add(Me.lblCustomerName)
        Me.grpHeader.Controls.Add(Me.Label15)
        Me.grpHeader.Controls.Add(Me.Label14)
        Me.grpHeader.Controls.Add(Me.txtORDate)
        Me.grpHeader.Controls.Add(Me.txtTeller)
        Me.grpHeader.Controls.Add(Me.lblPayAmt)
        Me.grpHeader.Controls.Add(Me.btnClear)
        Me.grpHeader.Controls.Add(Me.btnSave)
        Me.grpHeader.Controls.Add(Me.btnExit)
        Me.grpHeader.Controls.Add(Me.txtCusCde)
        Me.grpHeader.Controls.Add(Me.Label12)
        Me.grpHeader.Controls.Add(Me.Label11)
        Me.grpHeader.Controls.Add(Me.Label10)
        Me.grpHeader.Controls.Add(Me.txtCreditAmt)
        Me.grpHeader.Controls.Add(Me.txtCashAmt)
        Me.grpHeader.Controls.Add(Me.Label7)
        Me.grpHeader.Controls.Add(Me.Label6)
        Me.grpHeader.Controls.Add(Me.txtAmtTotal)
        Me.grpHeader.Controls.Add(Me.txtChequeAmtTotal)
        Me.grpHeader.Controls.Add(Me.txtChequeBank2)
        Me.grpHeader.Controls.Add(Me.txtChequeBank1)
        Me.grpHeader.Controls.Add(Me.txtChequeNo2)
        Me.grpHeader.Controls.Add(Me.txtChequeNo1)
        Me.grpHeader.Controls.Add(Me.txtChequeAmt2)
        Me.grpHeader.Controls.Add(Me.txtChequeAmt1)
        Me.grpHeader.Controls.Add(Me.Label9)
        Me.grpHeader.Controls.Add(Me.Label8)
        Me.grpHeader.Controls.Add(Me.Label5)
        Me.grpHeader.Controls.Add(Me.Label4)
        Me.grpHeader.Controls.Add(Me.Label3)
        Me.grpHeader.Controls.Add(Me.Panel1)
        Me.grpHeader.Controls.Add(Me.Label2)
        Me.grpHeader.Location = New System.Drawing.Point(7, 92)
        Me.grpHeader.Name = "grpHeader"
        Me.grpHeader.Size = New System.Drawing.Size(823, 256)
        Me.grpHeader.TabIndex = 1
        Me.grpHeader.TabStop = False
        '
        'lblCustomerName
        '
        Me.lblCustomerName.Font = New System.Drawing.Font("Lucida Sans", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCustomerName.Location = New System.Drawing.Point(328, 85)
        Me.lblCustomerName.Name = "lblCustomerName"
        Me.lblCustomerName.Size = New System.Drawing.Size(480, 16)
        Me.lblCustomerName.TabIndex = 31
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(592, 205)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(72, 16)
        Me.Label15.TabIndex = 30
        Me.Label15.Text = "Teller ID :"
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(592, 229)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(64, 16)
        Me.Label14.TabIndex = 29
        Me.Label14.Text = "OR Date :"
        '
        'txtORDate
        '
        Me.txtORDate.BackColor = System.Drawing.Color.Bisque
        Me.txtORDate.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtORDate.Location = New System.Drawing.Point(664, 224)
        Me.txtORDate.Name = "txtORDate"
        Me.txtORDate.ReadOnly = True
        Me.txtORDate.Size = New System.Drawing.Size(144, 21)
        Me.txtORDate.TabIndex = 28
        Me.txtORDate.TabStop = False
        Me.txtORDate.Text = ""
        '
        'txtTeller
        '
        Me.txtTeller.BackColor = System.Drawing.Color.Bisque
        Me.txtTeller.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTeller.Location = New System.Drawing.Point(664, 200)
        Me.txtTeller.Name = "txtTeller"
        Me.txtTeller.ReadOnly = True
        Me.txtTeller.Size = New System.Drawing.Size(144, 21)
        Me.txtTeller.TabIndex = 27
        Me.txtTeller.TabStop = False
        Me.txtTeller.Text = ""
        '
        'lblPayAmt
        '
        Me.lblPayAmt.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPayAmt.Location = New System.Drawing.Point(336, 224)
        Me.lblPayAmt.Name = "lblPayAmt"
        Me.lblPayAmt.Size = New System.Drawing.Size(56, 16)
        Me.lblPayAmt.TabIndex = 26
        Me.lblPayAmt.Text = "0"
        Me.lblPayAmt.Visible = False
        '
        'btnClear
        '
        Me.btnClear.Font = New System.Drawing.Font("Lucida Sans", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClear.Image = CType(resources.GetObject("btnClear.Image"), System.Drawing.Image)
        Me.btnClear.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnClear.Location = New System.Drawing.Point(691, 14)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(60, 58)
        Me.btnClear.TabIndex = 11
        Me.btnClear.Text = "CLEAR"
        Me.btnClear.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'btnSave
        '
        Me.btnSave.Enabled = False
        Me.btnSave.Font = New System.Drawing.Font("Lucida Sans", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Image = CType(resources.GetObject("btnSave.Image"), System.Drawing.Image)
        Me.btnSave.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnSave.Location = New System.Drawing.Point(630, 14)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(60, 58)
        Me.btnSave.TabIndex = 10
        Me.btnSave.Text = "SAVE"
        Me.btnSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'btnExit
        '
        Me.btnExit.Font = New System.Drawing.Font("Lucida Sans", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Image = CType(resources.GetObject("btnExit.Image"), System.Drawing.Image)
        Me.btnExit.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnExit.Location = New System.Drawing.Point(752, 14)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(60, 58)
        Me.btnExit.TabIndex = 12
        Me.btnExit.Text = "CLOSE"
        Me.btnExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'txtCusCde
        '
        Me.txtCusCde.BackColor = System.Drawing.Color.Bisque
        Me.txtCusCde.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCusCde.Location = New System.Drawing.Point(168, 80)
        Me.txtCusCde.Name = "txtCusCde"
        Me.txtCusCde.ReadOnly = True
        Me.txtCusCde.Size = New System.Drawing.Size(144, 21)
        Me.txtCusCde.TabIndex = 0
        Me.txtCusCde.TabStop = False
        Me.txtCusCde.Text = ""
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(592, 109)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(56, 16)
        Me.Label12.TabIndex = 25
        Me.Label12.Text = "Bank :"
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(328, 133)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(88, 16)
        Me.Label11.TabIndex = 24
        Me.Label11.Text = "Cheque No. :"
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(8, 133)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(120, 16)
        Me.Label10.TabIndex = 23
        Me.Label10.Text = "Cheque Amount :"
        '
        'txtCreditAmt
        '
        Me.txtCreditAmt.BackColor = System.Drawing.Color.Bisque
        Me.txtCreditAmt.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCreditAmt.Location = New System.Drawing.Point(168, 224)
        Me.txtCreditAmt.Name = "txtCreditAmt"
        Me.txtCreditAmt.ReadOnly = True
        Me.txtCreditAmt.Size = New System.Drawing.Size(144, 21)
        Me.txtCreditAmt.TabIndex = 0
        Me.txtCreditAmt.TabStop = False
        Me.txtCreditAmt.Text = "0.00"
        Me.txtCreditAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtCashAmt
        '
        Me.txtCashAmt.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCashAmt.Location = New System.Drawing.Point(168, 176)
        Me.txtCashAmt.Name = "txtCashAmt"
        Me.txtCashAmt.Size = New System.Drawing.Size(144, 21)
        Me.txtCashAmt.TabIndex = 9
        Me.txtCashAmt.Text = "0.00"
        Me.txtCashAmt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(8, 157)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(152, 16)
        Me.Label7.TabIndex = 20
        Me.Label7.Text = "Total Cheque Amount :"
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(8, 229)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(120, 16)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Credit Amount :"
        '
        'txtAmtTotal
        '
        Me.txtAmtTotal.BackColor = System.Drawing.Color.Bisque
        Me.txtAmtTotal.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAmtTotal.Location = New System.Drawing.Point(168, 200)
        Me.txtAmtTotal.Name = "txtAmtTotal"
        Me.txtAmtTotal.ReadOnly = True
        Me.txtAmtTotal.Size = New System.Drawing.Size(144, 21)
        Me.txtAmtTotal.TabIndex = 0
        Me.txtAmtTotal.TabStop = False
        Me.txtAmtTotal.Text = "0.00"
        Me.txtAmtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtChequeAmtTotal
        '
        Me.txtChequeAmtTotal.BackColor = System.Drawing.Color.Bisque
        Me.txtChequeAmtTotal.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChequeAmtTotal.Location = New System.Drawing.Point(168, 152)
        Me.txtChequeAmtTotal.Name = "txtChequeAmtTotal"
        Me.txtChequeAmtTotal.ReadOnly = True
        Me.txtChequeAmtTotal.Size = New System.Drawing.Size(144, 21)
        Me.txtChequeAmtTotal.TabIndex = 0
        Me.txtChequeAmtTotal.TabStop = False
        Me.txtChequeAmtTotal.Text = "0.00"
        Me.txtChequeAmtTotal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtChequeBank2
        '
        Me.txtChequeBank2.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChequeBank2.Location = New System.Drawing.Point(664, 128)
        Me.txtChequeBank2.Name = "txtChequeBank2"
        Me.txtChequeBank2.Size = New System.Drawing.Size(144, 21)
        Me.txtChequeBank2.TabIndex = 8
        Me.txtChequeBank2.Text = ""
        '
        'txtChequeBank1
        '
        Me.txtChequeBank1.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChequeBank1.Location = New System.Drawing.Point(664, 104)
        Me.txtChequeBank1.Name = "txtChequeBank1"
        Me.txtChequeBank1.Size = New System.Drawing.Size(144, 21)
        Me.txtChequeBank1.TabIndex = 5
        Me.txtChequeBank1.Text = ""
        '
        'txtChequeNo2
        '
        Me.txtChequeNo2.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChequeNo2.Location = New System.Drawing.Point(432, 128)
        Me.txtChequeNo2.Name = "txtChequeNo2"
        Me.txtChequeNo2.Size = New System.Drawing.Size(144, 21)
        Me.txtChequeNo2.TabIndex = 7
        Me.txtChequeNo2.Text = ""
        Me.txtChequeNo2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtChequeNo1
        '
        Me.txtChequeNo1.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChequeNo1.Location = New System.Drawing.Point(432, 104)
        Me.txtChequeNo1.Name = "txtChequeNo1"
        Me.txtChequeNo1.Size = New System.Drawing.Size(144, 21)
        Me.txtChequeNo1.TabIndex = 4
        Me.txtChequeNo1.Text = ""
        Me.txtChequeNo1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtChequeAmt2
        '
        Me.txtChequeAmt2.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChequeAmt2.Location = New System.Drawing.Point(168, 128)
        Me.txtChequeAmt2.Name = "txtChequeAmt2"
        Me.txtChequeAmt2.Size = New System.Drawing.Size(144, 21)
        Me.txtChequeAmt2.TabIndex = 6
        Me.txtChequeAmt2.Text = "0.00"
        Me.txtChequeAmt2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtChequeAmt1
        '
        Me.txtChequeAmt1.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChequeAmt1.Location = New System.Drawing.Point(168, 104)
        Me.txtChequeAmt1.Name = "txtChequeAmt1"
        Me.txtChequeAmt1.Size = New System.Drawing.Size(144, 21)
        Me.txtChequeAmt1.TabIndex = 3
        Me.txtChequeAmt1.Text = "0.00"
        Me.txtChequeAmt1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(592, 133)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(56, 16)
        Me.Label9.TabIndex = 10
        Me.Label9.Text = "Bank :"
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(328, 109)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(88, 16)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "Cheque No. :"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(8, 109)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(120, 16)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "Cheque Amount :"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(8, 181)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(120, 16)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Cash Amount :"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(8, 205)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Total Amount :"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Bisque
        Me.Panel1.Controls.Add(Me.btnGet)
        Me.Panel1.Controls.Add(Me.txtORNum)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(9, 14)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(303, 48)
        Me.Panel1.TabIndex = 0
        '
        'btnGet
        '
        Me.btnGet.BackColor = System.Drawing.Color.Tan
        Me.btnGet.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGet.Location = New System.Drawing.Point(222, 13)
        Me.btnGet.Name = "btnGet"
        Me.btnGet.Size = New System.Drawing.Size(72, 22)
        Me.btnGet.TabIndex = 2
        Me.btnGet.Text = "GET"
        '
        'txtORNum
        '
        Me.txtORNum.BackColor = System.Drawing.Color.White
        Me.txtORNum.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtORNum.Location = New System.Drawing.Point(96, 14)
        Me.txtORNum.MaxLength = 5
        Me.txtORNum.Name = "txtORNum"
        Me.txtORNum.Size = New System.Drawing.Size(112, 22)
        Me.txtORNum.TabIndex = 1
        Me.txtORNum.Text = ""
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "OR Number :"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 85)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(120, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Customer Code :"
        '
        'grpDetails
        '
        Me.grpDetails.Controls.Add(Me.dgINVPayDtl)
        Me.grpDetails.Location = New System.Drawing.Point(7, 350)
        Me.grpDetails.Name = "grpDetails"
        Me.grpDetails.Size = New System.Drawing.Size(823, 306)
        Me.grpDetails.TabIndex = 2
        Me.grpDetails.TabStop = False
        '
        'dgINVPayDtl
        '
        Me.dgINVPayDtl.CaptionBackColor = System.Drawing.Color.Bisque
        Me.dgINVPayDtl.CaptionFont = New System.Drawing.Font("Lucida Sans", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgINVPayDtl.CaptionForeColor = System.Drawing.SystemColors.ControlText
        Me.dgINVPayDtl.CaptionText = "INVOICE PAYMENT DETAILS"
        Me.dgINVPayDtl.DataMember = ""
        Me.dgINVPayDtl.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dgINVPayDtl.LinkColor = System.Drawing.Color.Bisque
        Me.dgINVPayDtl.Location = New System.Drawing.Point(9, 15)
        Me.dgINVPayDtl.Name = "dgINVPayDtl"
        Me.dgINVPayDtl.ReadOnly = True
        Me.dgINVPayDtl.Size = New System.Drawing.Size(804, 281)
        Me.dgINVPayDtl.TabIndex = 0
        Me.dgINVPayDtl.TabStop = False
        '
        'statbarINVCorretion
        '
        Me.statbarINVCorretion.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.statbarINVCorretion.Location = New System.Drawing.Point(0, 663)
        Me.statbarINVCorretion.Name = "statbarINVCorretion"
        Me.statbarINVCorretion.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.statPanelUser, Me.statPanelDate, Me.statPanelTime})
        Me.statbarINVCorretion.ShowPanels = True
        Me.statbarINVCorretion.Size = New System.Drawing.Size(836, 22)
        Me.statbarINVCorretion.SizingGrip = False
        Me.statbarINVCorretion.TabIndex = 5
        '
        'statPanelUser
        '
        Me.statPanelUser.Width = 420
        '
        'statPanelDate
        '
        Me.statPanelDate.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.statPanelDate.Width = 250
        '
        'statPanelTime
        '
        Me.statPanelTime.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.statPanelTime.Width = 165
        '
        'frmINVCorrection
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Tan
        Me.ClientSize = New System.Drawing.Size(836, 685)
        Me.Controls.Add(Me.statbarINVCorretion)
        Me.Controls.Add(Me.grpDetails)
        Me.Controls.Add(Me.grpHeader)
        Me.Controls.Add(Me.grpTitle)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmINVCorrection"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "INVOICE CORRECTION"
        Me.grpTitle.ResumeLayout(False)
        Me.grpHeader.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.grpDetails.ResumeLayout(False)
        CType(Me.dgINVPayDtl, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.statPanelUser, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.statPanelDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.statPanelTime, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private dtrowHdr As DataRow
    Private dtrowDtl As DataRow
    Private clsINVCorrection As clsINVCorrection
    Private strSQL As String
    Private ts As DataGridTableStyle

    Private Sub frmINVCorrection_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetStatusBar()
        txtORNum.Focus()
    End Sub

    Private Sub SetStatusBar()
        statPanelDate.Text = CType(FormatDateTime(Today(), DateFormat.LongDate), String) & " "
        statPanelTime.Text = CType(TimeValue(Now()), String) & " "
        statPanelUser.Text = " User Name : " & UCase(zCurrentUser())
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        End
    End Sub

    '*************************** R E T R I V E  D A T A ***************************
    Private Sub btnGet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGet.Click
        If Not IsNumeric(txtORNum.Text) Then
            MsgBox("Invalid OR Number", MsgBoxStyle.Exclamation, "INVALID")
            txtORNum.Focus()
            Exit Sub
        End If
        clsINVCorrection = New clsINVCorrection
        strSQL = "SELECT * FROM INVPAYHDR WHERE ORNum = " & txtORNum.Text
        clsINVCorrection.PopulateDataSet(strSQL, "Header")
        If dsINVPayHdr.Tables.Count > 0 Then
            If dsINVPayHdr.Tables(0).Rows.Count > 0 Then
                PopulateGrpHdr()
                txtORNum.ReadOnly = True
                txtORNum.BackColor = Color.Tan
                strSQL = "SELECT INVNUM,INVAmt,PAYAmt,PAYDate,RBalance FROM INVPAYDTL WHERE ORNum = " & txtORNum.Text
                clsINVCorrection.PopulateDataSet(strSQL, "Detail")
                If dsINVPayDtl.Tables.Count > 0 Then
                    PopulatedgINVPayDtl()

                    Dim dtView As DataView
                    Dim dtrowView As DataRowView
                    Dim decTotalPay As Decimal = 0
                    If dsINVPayDtl.Tables(0).Rows.Count > 0 Then
                        dtView = dsINVPayDtl.Tables(0).DefaultView
                        For Each dtrowView In dtView
                            decTotalPay = decTotalPay + CType(dtrowView("PAYAmt"), Decimal)
                        Next
                        lblPayAmt.Text = decTotalPay
                    End If
                End If
            Else
                MsgBox("Invalid OR Number", MsgBoxStyle.Exclamation, "INVALID")
                txtORNum.Focus()
            End If
        End If
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        ClearForm()
    End Sub

    '*************************** U P D A T E  I N V P A Y H D R ***************************
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim ValFields As Boolean
        ValFields = ValidateFields()
        If ValFields = False Then
            Exit Sub
        End If

        Dim strResponse As String = MsgBox("Do you want to Save the Changes you made to OR # " & txtORNum.Text, MsgBoxStyle.OKCancel, "INVOICE CORRECTION")
        If strResponse = vbCancel Then
            Exit Sub
        End If

        Dim strSQL As String = "UPDATE INVPAYHDR SET " & _
                               "CheckAmt1 = " & txtChequeAmt1.Text & _
                               ",CheckAMT2 = " & txtChequeAmt2.Text & _
                               ",CheckNo1 = " & txtChequeNo1.Text & _
                               ",CheckNo2 = " & txtChequeNo2.Text & _
                               ",CheckBnk1 = '" & txtChequeBank1.Text & "'" & _
                               ",CheckBnk2 = '" & txtChequeBank2.Text & "'" & _
                               ",CashAMT = " & CStr(CDec(txtCashAmt.Text)) & _
                               ",TotalAMT = " & CStr(CDec(txtAmtTotal.Text)) & _
                               ",AvailAMT = " & CStr(CDec(txtCreditAmt.Text)) & _
                               " WHERE ORNum = " & txtORNum.Text

        Dim isUpdate As Boolean
        clsINVCorrection = New clsINVCorrection
        isUpdate = clsINVCorrection.UpdateINVPAYHDR(strSQL)

        If isUpdate = True Then
            MsgBox("Invoice is Updated.", MsgBoxStyle.Exclamation, "")
        Else
            MsgBox("Invoice is not Updated.", MsgBoxStyle.Critical, "ERROR")
        End If
    End Sub

    '*************************** F I E L D  V A L I D A T I O N ***************************
    Private Function ValChe1() As Boolean
        If IsNumeric(txtChequeAmt1.Text) = True Then
            If txtChequeAmt1.Text > 0 Then
                If IsNumeric(txtChequeNo1.Text) = False Then
                    MsgBox("Invalid Cheque Number.", MsgBoxStyle.Critical, "INVALID")
                    txtChequeNo1.Focus()
                    Return False
                ElseIf CType(txtChequeNo1.Text, Integer) = 0 Then
                    MsgBox("Invalid Cheque Number.", MsgBoxStyle.Critical, "INVALID")
                    txtChequeNo1.Focus()
                    Return False
                End If
                If txtChequeBank1.Text = "" Then
                    MsgBox("Invalid Cheque Bank.", MsgBoxStyle.Critical, "INVALID")
                    txtChequeBank1.Focus()
                    Return False
                End If
                Return True
            Else
                txtChequeNo1.Text = 0
                txtChequeBank1.Text = ""
                Return True
            End If
        Else
            MsgBox("Invalid Cheque Amount.", MsgBoxStyle.Critical, "INVALID")
            Return False
        End If
    End Function

    Private Function ValChe2() As Boolean
        If IsNumeric(txtChequeAmt2.Text) = True Then
            If txtChequeAmt2.Text > 0 Then
                If IsNumeric(txtChequeNo2.Text) = False Then
                    MsgBox("Invalid Cheque Number.", MsgBoxStyle.Critical, "INVALID")
                    txtChequeNo2.Focus()
                    Return False
                ElseIf CType(txtChequeNo2.Text, Integer) = 0 Then
                    MsgBox("Invalid Cheque Number.", MsgBoxStyle.Critical, "INVALID")
                    txtChequeNo2.Focus()
                    Return False
                End If
                If txtChequeBank2.Text = "" Then
                    MsgBox("Invalid Cheque Bank.", MsgBoxStyle.Critical, "INVALID")
                    txtChequeBank2.Focus()
                    Return False
                End If
                Return True
            Else
                txtChequeNo2.Text = 0
                txtChequeBank2.Text = ""
                Return True
            End If
        Else
            MsgBox("Invalid Cheque Amount.", MsgBoxStyle.Critical, "INVALID")
            Return False
        End If
    End Function

    Private Function ValidateFields() As Boolean
        Dim ValCheque1 As Boolean = ValChe1()
        Dim ValCheque2 As Boolean = ValChe2()

        If ValCheque1 = False Or ValCheque2 = False Or IsNumeric(txtCashAmt.Text) = False Then
            If IsNumeric(txtCashAmt.Text) = False Then
                MsgBox("Invalid Cash Amount.", MsgBoxStyle.Critical, "INVALID")
                txtCashAmt.Focus()
            End If
            Return False
        End If

        If CType(txtAmtTotal.Text, Decimal) >= CType(lblPayAmt.Text, Decimal) Then
            Return True
        Else
            Return False
        End If
    End Function

    '*************************** P O P U L A T E  D A T A G R I D ***************************
    Private Sub PopulatedgINVPayDtl()
        Dim dtabDetails As New DataTable
        dtabDetails = dsINVPayDtl.Tables(0)
        dgINVPayDtl.TableStyles.Clear()
        ts = New DataGridTableStyle
        ts.MappingName = dtabDetails.ToString
        dgINVPayDtl.TableStyles.Add(ts)

        With dgINVPayDtl
            .DataSource = dtabDetails
            .AlternatingBackColor = Color.AliceBlue
            .BackColor = Color.White
            .TableStyles(0).RowHeaderWidth = 15
            .TableStyles(0).GridColumnStyles.Item(0).Width = 133
            .TableStyles(0).GridColumnStyles.Item(1).Width = 175
            .TableStyles(0).GridColumnStyles.Item(2).Width = 175
            .TableStyles(0).GridColumnStyles.Item(3).Width = 110
            .TableStyles(0).GridColumnStyles.Item(4).Width = 175
            .TableStyles(0).GridColumnStyles.Item(0).HeaderText = "Invoice Number"
            .TableStyles(0).GridColumnStyles.Item(1).HeaderText = "Invoice Amount"
            .TableStyles(0).GridColumnStyles.Item(2).HeaderText = "Amount Paid"
            .TableStyles(0).GridColumnStyles.Item(3).HeaderText = "Date"
            .TableStyles(0).GridColumnStyles.Item(4).HeaderText = "Balance"
            .TableStyles(0).SelectionBackColor = .TableStyles(0).BackColor.LightSalmon
            .TableStyles(0).SelectionForeColor = .TableStyles(0).ForeColor.Black
            .TableStyles(0).AlternatingBackColor = .TableStyles(0).BackColor.AntiqueWhite
            .TableStyles(0).BackColor = .TableStyles(0).BackColor.White
            .TableStyles(0).HeaderBackColor = .TableStyles(0).BackColor.Sienna
            .TableStyles(0).HeaderForeColor = .TableStyles(0).ForeColor.White
            .TableStyles(0).HeaderFont = New System.Drawing.Font("Tahoma", 9.0F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, System.Byte))
        End With
    End Sub

    Private Sub PopulateGrpHdr()
        txtChequeAmt1.Focus()
        btnSave.Enabled = True
        clsINVCorrection = New clsINVCorrection
        dtrowHdr = clsINVCorrection.RetriveINVPAYHDR(txtORNum.Text)

        txtCusCde.Text = dtrowHdr.Item("cuscde")
        txtORDate.Text = IsStrNull(Format(dtrowHdr.Item("ORDate"), "MM/dd/yyyy"))
        txtTeller.Text = UCase(IsStrNull(dtrowHdr.Item("userid")))
        txtChequeAmt1.Text = IsZeroAmt(dtrowHdr.Item("CheckAmt1"))
        txtChequeAmt2.Text = IsZeroAmt(dtrowHdr.Item("CheckAMT2"))
        txtChequeNo1.Text = dtrowHdr.Item("CheckNo1")
        txtChequeNo2.Text = dtrowHdr.Item("CheckNo2")
        txtChequeBank1.Text = IsStrNull(dtrowHdr.Item("CheckBnk1"))
        txtChequeBank2.Text = IsStrNull(dtrowHdr.Item("CheckBnk2"))
        txtChequeAmtTotal.Text = IsZeroAmt(CType(dtrowHdr.Item("CheckAmt1"), Decimal) + CType(dtrowHdr.Item("CheckAMT2"), Decimal))
        txtCashAmt.Text = IsZeroAmt(dtrowHdr.Item("CashAMT"))
        txtAmtTotal.Text = IsZeroAmt(dtrowHdr.Item("TotalAMT"))
        txtCreditAmt.Text = IsZeroAmt(dtrowHdr.Item("AvailAMT"))
        lblCustomerName.Text = clsINVCorrection.GetCustomerName("SELECT cusnam FROM Customer WHERE cuscde=" & txtCusCde.Text)

    End Sub

    Private Function IsZeroAmt(ByVal objVal As Object) As String
        If IsDBNull(objVal) = True Or CType(objVal, Decimal) = 0 Then
            Return "0.00"
        Else
            Return CType(objVal, String)
        End If
    End Function

    Private Function IsStrNull(ByVal objVal As Object) As String
        If IsDBNull(objVal) = True Then
            Return ""
        Else
            Return Trim(objVal)
        End If
    End Function

    '*************************** C L E A R  F O R M ***************************
    Private Sub ClearForm()
        btnSave.Enabled = False
        txtORNum.Focus()
        dgINVPayDtl.DataSource = Nothing
        txtTeller.Text = ""
        txtORDate.Text = ""
        txtCusCde.Text = ""
        txtChequeAmt1.Text = "0.00"
        txtChequeAmt2.Text = "0.00"
        txtChequeNo1.Text = ""
        txtChequeNo2.Text = ""
        txtChequeBank1.Text = ""
        txtChequeBank2.Text = ""
        txtChequeAmtTotal.Text = "0.00"
        txtCashAmt.Text = "0.00"
        txtAmtTotal.Text = "0.00"
        txtCreditAmt.Text = "0.00"
        txtORNum.Text = ""
        txtORNum.ReadOnly = False
        txtORNum.BackColor = Color.White
        lblCustomerName.Text = ""
    End Sub

    Private Function getToString(ByRef strValue) As String
        If strValue.Trim = "" Then
            Return "NULL"
        Else
            Return "'" & strValue.Trim & "'"
        End If
    End Function

#Region "LOST FOCUS"
    Private Sub txtChequeAmt1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtChequeAmt1.LostFocus
        If txtChequeAmt1.Text = "" Then
            txtChequeAmt1.Text = "0.00"
        End If
        If IsNumeric(txtChequeAmt1.Text) = False Then
            MsgBox("Invalid Cheque Amount.", MsgBoxStyle.Critical, "INVALID")
            txtChequeAmt1.Focus()
            Exit Sub
        End If
        txtAmtTotal.Text = FormatNumber(CType(txtChequeAmt1.Text, Decimal) + CType(txtChequeAmt2.Text, Decimal) + CType(txtCashAmt.Text, Decimal), 2)
        txtCreditAmt.Text = FormatNumber(CType(txtAmtTotal.Text, Decimal) - CType(lblPayAmt.Text, Decimal), 2)
        txtChequeAmtTotal.Text = FormatNumber(CType(txtChequeAmt1.Text, Decimal) + CType(txtChequeAmt2.Text, Decimal), 2)
        If txtCreditAmt.Text < 0 Then
            MsgBox("Total Invoice Payment is Higher than Total Invoice Amount.", MsgBoxStyle.Critical, "INVALID")
            txtChequeAmt1.Focus()
        ElseIf txtCreditAmt.Text = 0 Then
            txtCreditAmt.Text = "0.00"
        End If
    End Sub

    Private Sub txtChequeAmt2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtChequeAmt2.LostFocus
        If txtChequeAmt2.Text = "" Then
            txtChequeAmt2.Text = "0.00"
        End If
        If IsNumeric(txtChequeAmt2.Text) = False Then
            MsgBox("Invalid Cheque Amount.", MsgBoxStyle.Critical, "INVALID")
            txtChequeAmt2.Focus()
            Exit Sub
        End If
        txtAmtTotal.Text = FormatNumber(CType(txtChequeAmt1.Text, Decimal) + CType(txtChequeAmt2.Text, Decimal) + CType(txtCashAmt.Text, Decimal), 2)
        txtCreditAmt.Text = FormatNumber(CType(txtAmtTotal.Text, Decimal) - CType(lblPayAmt.Text, Decimal), 2)
        txtChequeAmtTotal.Text = FormatNumber(CType(txtChequeAmt1.Text, Decimal) + CType(txtChequeAmt2.Text, Decimal), 2)
        If txtCreditAmt.Text < 0 Then
            MsgBox("Total Invoice Payment is Higher than Total Invoice Amount.", MsgBoxStyle.Critical, "INVALID")
            txtChequeAmt2.Focus()
        ElseIf txtCreditAmt.Text = 0 Then
            txtCreditAmt.Text = "0.00"
        End If
    End Sub

    Private Sub txtCashAmt_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCashAmt.LostFocus
        If txtCashAmt.Text = "" Then
            txtCashAmt.Text = "0.00"
        End If
        If IsNumeric(txtCashAmt.Text) = False Then
            MsgBox("Invalid Cash Amount.", MsgBoxStyle.Critical, "INVALID")
            txtCashAmt.Focus()
            Exit Sub
        End If
        txtAmtTotal.Text = FormatNumber(CType(txtChequeAmt1.Text, Decimal) + CType(txtChequeAmt2.Text, Decimal) + CType(txtCashAmt.Text, Decimal), 2)
        txtCreditAmt.Text = FormatNumber(CType(txtAmtTotal.Text, Decimal) - CType(lblPayAmt.Text, Decimal), 2)
        If txtCreditAmt.Text < 0 Then
            MsgBox("Total Invoice Payment is Higher than Total Invoice Amount.", MsgBoxStyle.Critical, "INVALID")
            txtCashAmt.Focus()
        ElseIf txtCreditAmt.Text = 0 Then
            txtCreditAmt.Text = "0.00"
        End If
    End Sub
#End Region

#Region "KEY DOWN"
    Private Sub txtCashAmt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCashAmt.KeyDown
        If e.KeyCode = 13 Then
            btnSave.Focus()
        End If
    End Sub

    Private Sub txtChequeAmt1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChequeAmt1.KeyDown
        If e.KeyCode = 13 Then
            txtChequeNo1.Focus()
        End If
    End Sub

    Private Sub txtChequeNo1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChequeNo1.KeyDown
        If e.KeyCode = 13 Then
            txtChequeBank1.Focus()
        End If
    End Sub

    Private Sub txtChequeBank1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChequeBank1.KeyDown
        If e.KeyCode = 13 Then
            txtChequeAmt2.Focus()
        End If
    End Sub

    Private Sub txtChequeAmt2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChequeAmt2.KeyDown
        If e.KeyCode = 13 Then
            txtChequeNo2.Focus()
        End If
    End Sub

    Private Sub txtChequeNo2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChequeNo2.KeyDown
        If e.KeyCode = 13 Then
            txtChequeBank2.Focus()
        End If
    End Sub

    Private Sub txtChequeBank2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtChequeBank2.KeyDown
        If e.KeyCode = 13 Then
            txtCashAmt.Focus()
        End If
    End Sub

    Private Sub txtORNum_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtORNum.KeyDown
        If e.KeyCode = 13 Then
            btnGet.Focus()
        End If
    End Sub
#End Region

End Class

