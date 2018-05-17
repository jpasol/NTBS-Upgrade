Option Explicit On 

Public Class frmCAC
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtExCash As System.Windows.Forms.TextBox
    Friend WithEvents txtTotCash As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents btnCLOSE As System.Windows.Forms.Button
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents dtePeriod As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents cmbTransType As System.Windows.Forms.ComboBox
    Friend WithEvents txtTellerID As System.Windows.Forms.TextBox
    Friend WithEvents txtTimeTo As System.Windows.Forms.TextBox
    Friend WithEvents txtTimeFrom As System.Windows.Forms.TextBox
    Friend WithEvents txtAmtLeft As System.Windows.Forms.TextBox
    Friend WithEvents txtExcessCheque As System.Windows.Forms.TextBox
    Friend WithEvents txtTotCheque As System.Windows.Forms.TextBox
    Friend WithEvents txtGrandTot As System.Windows.Forms.TextBox
    Friend WithEvents txtRemarks As System.Windows.Forms.TextBox
    Friend WithEvents statCACBar As System.Windows.Forms.StatusBar
    Friend WithEvents statPanelUser As System.Windows.Forms.StatusBarPanel
    Friend WithEvents statPanelDate As System.Windows.Forms.StatusBarPanel
    Friend WithEvents statPanelTime As System.Windows.Forms.StatusBarPanel
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents txt1000 As System.Windows.Forms.TextBox
    Friend WithEvents txt500 As System.Windows.Forms.TextBox
    Friend WithEvents txt200 As System.Windows.Forms.TextBox
    Friend WithEvents txt100 As System.Windows.Forms.TextBox
    Friend WithEvents txt50 As System.Windows.Forms.TextBox
    Friend WithEvents txt20 As System.Windows.Forms.TextBox
    Friend WithEvents txt10 As System.Windows.Forms.TextBox
    Friend WithEvents txt5 As System.Windows.Forms.TextBox
    Friend WithEvents txt1 As System.Windows.Forms.TextBox
    Friend WithEvents txt010 As System.Windows.Forms.TextBox
    Friend WithEvents txt005 As System.Windows.Forms.TextBox
    Friend WithEvents txt001 As System.Windows.Forms.TextBox
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents txtTot1000 As System.Windows.Forms.TextBox
    Friend WithEvents txtTot500 As System.Windows.Forms.TextBox
    Friend WithEvents txtTot200 As System.Windows.Forms.TextBox
    Friend WithEvents txtTot100 As System.Windows.Forms.TextBox
    Friend WithEvents txtTot50 As System.Windows.Forms.TextBox
    Friend WithEvents txtTot20 As System.Windows.Forms.TextBox
    Friend WithEvents txtTot10 As System.Windows.Forms.TextBox
    Friend WithEvents txtTot5 As System.Windows.Forms.TextBox
    Friend WithEvents txtTot1 As System.Windows.Forms.TextBox
    Friend WithEvents txtTot025 As System.Windows.Forms.TextBox
    Friend WithEvents txtTot010 As System.Windows.Forms.TextBox
    Friend WithEvents txtTot005 As System.Windows.Forms.TextBox
    Friend WithEvents txtTot001 As System.Windows.Forms.TextBox
    Friend WithEvents txt025 As System.Windows.Forms.TextBox
    Friend WithEvents gbCriteria As System.Windows.Forms.GroupBox
    Friend WithEvents gbCash As System.Windows.Forms.GroupBox
    Friend WithEvents gbCheque As System.Windows.Forms.GroupBox
    Friend WithEvents gbControl As System.Windows.Forms.GroupBox
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents dgChequeStat As System.Windows.Forms.DataGrid
    Friend WithEvents dgChequeDetails As System.Windows.Forms.DataGrid
    Friend WithEvents lblBatch As System.Windows.Forms.Label
    Friend WithEvents lstTimeRange As System.Windows.Forms.ListBox
    Friend WithEvents lblTimeTo As System.Windows.Forms.Label
    Friend WithEvents lblTimeFrom As System.Windows.Forms.Label
    Friend WithEvents lblTimeRange As System.Windows.Forms.Label
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents gbHeader As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCAC))
        Me.gbCriteria = New System.Windows.Forms.GroupBox
        Me.lstTimeRange = New System.Windows.Forms.ListBox
        Me.lblBatch = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.cmbTransType = New System.Windows.Forms.ComboBox
        Me.txtTellerID = New System.Windows.Forms.TextBox
        Me.txtTimeTo = New System.Windows.Forms.TextBox
        Me.txtTimeFrom = New System.Windows.Forms.TextBox
        Me.lblTimeTo = New System.Windows.Forms.Label
        Me.lblTimeFrom = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.dtePeriod = New System.Windows.Forms.DateTimePicker
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.lblTimeRange = New System.Windows.Forms.Label
        Me.gbCash = New System.Windows.Forms.GroupBox
        Me.Label35 = New System.Windows.Forms.Label
        Me.Label34 = New System.Windows.Forms.Label
        Me.txt001 = New System.Windows.Forms.TextBox
        Me.txt005 = New System.Windows.Forms.TextBox
        Me.txt010 = New System.Windows.Forms.TextBox
        Me.txt025 = New System.Windows.Forms.TextBox
        Me.txt1 = New System.Windows.Forms.TextBox
        Me.txt5 = New System.Windows.Forms.TextBox
        Me.txt10 = New System.Windows.Forms.TextBox
        Me.txt20 = New System.Windows.Forms.TextBox
        Me.txt50 = New System.Windows.Forms.TextBox
        Me.txt100 = New System.Windows.Forms.TextBox
        Me.txt200 = New System.Windows.Forms.TextBox
        Me.txt500 = New System.Windows.Forms.TextBox
        Me.txt1000 = New System.Windows.Forms.TextBox
        Me.txtTot001 = New System.Windows.Forms.TextBox
        Me.txtTot005 = New System.Windows.Forms.TextBox
        Me.txtTot010 = New System.Windows.Forms.TextBox
        Me.txtTot025 = New System.Windows.Forms.TextBox
        Me.Label33 = New System.Windows.Forms.Label
        Me.Label32 = New System.Windows.Forms.Label
        Me.Label31 = New System.Windows.Forms.Label
        Me.Label30 = New System.Windows.Forms.Label
        Me.Label29 = New System.Windows.Forms.Label
        Me.Label28 = New System.Windows.Forms.Label
        Me.Label27 = New System.Windows.Forms.Label
        Me.Label26 = New System.Windows.Forms.Label
        Me.Label25 = New System.Windows.Forms.Label
        Me.Label24 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.txtTot1 = New System.Windows.Forms.TextBox
        Me.txtTot5 = New System.Windows.Forms.TextBox
        Me.txtTot10 = New System.Windows.Forms.TextBox
        Me.txtTot20 = New System.Windows.Forms.TextBox
        Me.txtTot50 = New System.Windows.Forms.TextBox
        Me.txtTot100 = New System.Windows.Forms.TextBox
        Me.txtTot200 = New System.Windows.Forms.TextBox
        Me.txtTot500 = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.txtTot1000 = New System.Windows.Forms.TextBox
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtAmtLeft = New System.Windows.Forms.TextBox
        Me.txtTotCash = New System.Windows.Forms.TextBox
        Me.txtExCash = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.gbCheque = New System.Windows.Forms.GroupBox
        Me.txtExcessCheque = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtTotCheque = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.dgChequeStat = New System.Windows.Forms.DataGrid
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.dgChequeDetails = New System.Windows.Forms.DataGrid
        Me.gbControl = New System.Windows.Forms.GroupBox
        Me.Label36 = New System.Windows.Forms.Label
        Me.txtGrandTot = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtRemarks = New System.Windows.Forms.TextBox
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnCLOSE = New System.Windows.Forms.Button
        Me.statCACBar = New System.Windows.Forms.StatusBar
        Me.statPanelUser = New System.Windows.Forms.StatusBarPanel
        Me.statPanelDate = New System.Windows.Forms.StatusBarPanel
        Me.statPanelTime = New System.Windows.Forms.StatusBarPanel
        Me.lblID = New System.Windows.Forms.Label
        Me.gbHeader = New System.Windows.Forms.GroupBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.gbCriteria.SuspendLayout()
        Me.gbCash.SuspendLayout()
        Me.gbCheque.SuspendLayout()
        CType(Me.dgChequeStat, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgChequeDetails, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbControl.SuspendLayout()
        CType(Me.statPanelUser, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.statPanelDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.statPanelTime, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbHeader.SuspendLayout()
        Me.SuspendLayout()
        '
        'gbCriteria
        '
        Me.gbCriteria.AccessibleDescription = resources.GetString("gbCriteria.AccessibleDescription")
        Me.gbCriteria.AccessibleName = resources.GetString("gbCriteria.AccessibleName")
        Me.gbCriteria.Anchor = CType(resources.GetObject("gbCriteria.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.gbCriteria.BackgroundImage = CType(resources.GetObject("gbCriteria.BackgroundImage"), System.Drawing.Image)
        Me.gbCriteria.Controls.Add(Me.lstTimeRange)
        Me.gbCriteria.Controls.Add(Me.lblBatch)
        Me.gbCriteria.Controls.Add(Me.Label19)
        Me.gbCriteria.Controls.Add(Me.Label18)
        Me.gbCriteria.Controls.Add(Me.cmbTransType)
        Me.gbCriteria.Controls.Add(Me.txtTellerID)
        Me.gbCriteria.Controls.Add(Me.txtTimeTo)
        Me.gbCriteria.Controls.Add(Me.txtTimeFrom)
        Me.gbCriteria.Controls.Add(Me.lblTimeTo)
        Me.gbCriteria.Controls.Add(Me.lblTimeFrom)
        Me.gbCriteria.Controls.Add(Me.Label15)
        Me.gbCriteria.Controls.Add(Me.dtePeriod)
        Me.gbCriteria.Controls.Add(Me.Label14)
        Me.gbCriteria.Controls.Add(Me.Label13)
        Me.gbCriteria.Controls.Add(Me.lblTimeRange)
        Me.gbCriteria.Dock = CType(resources.GetObject("gbCriteria.Dock"), System.Windows.Forms.DockStyle)
        Me.gbCriteria.Enabled = CType(resources.GetObject("gbCriteria.Enabled"), Boolean)
        Me.gbCriteria.Font = CType(resources.GetObject("gbCriteria.Font"), System.Drawing.Font)
        Me.gbCriteria.ImeMode = CType(resources.GetObject("gbCriteria.ImeMode"), System.Windows.Forms.ImeMode)
        Me.gbCriteria.Location = CType(resources.GetObject("gbCriteria.Location"), System.Drawing.Point)
        Me.gbCriteria.Name = "gbCriteria"
        Me.gbCriteria.RightToLeft = CType(resources.GetObject("gbCriteria.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.gbCriteria.Size = CType(resources.GetObject("gbCriteria.Size"), System.Drawing.Size)
        Me.gbCriteria.TabIndex = CType(resources.GetObject("gbCriteria.TabIndex"), Integer)
        Me.gbCriteria.TabStop = False
        Me.gbCriteria.Text = resources.GetString("gbCriteria.Text")
        Me.gbCriteria.Visible = CType(resources.GetObject("gbCriteria.Visible"), Boolean)
        '
        'lstTimeRange
        '
        Me.lstTimeRange.AccessibleDescription = resources.GetString("lstTimeRange.AccessibleDescription")
        Me.lstTimeRange.AccessibleName = resources.GetString("lstTimeRange.AccessibleName")
        Me.lstTimeRange.Anchor = CType(resources.GetObject("lstTimeRange.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.lstTimeRange.BackgroundImage = CType(resources.GetObject("lstTimeRange.BackgroundImage"), System.Drawing.Image)
        Me.lstTimeRange.ColumnWidth = CType(resources.GetObject("lstTimeRange.ColumnWidth"), Integer)
        Me.lstTimeRange.Dock = CType(resources.GetObject("lstTimeRange.Dock"), System.Windows.Forms.DockStyle)
        Me.lstTimeRange.Enabled = CType(resources.GetObject("lstTimeRange.Enabled"), Boolean)
        Me.lstTimeRange.Font = CType(resources.GetObject("lstTimeRange.Font"), System.Drawing.Font)
        Me.lstTimeRange.HorizontalExtent = CType(resources.GetObject("lstTimeRange.HorizontalExtent"), Integer)
        Me.lstTimeRange.HorizontalScrollbar = CType(resources.GetObject("lstTimeRange.HorizontalScrollbar"), Boolean)
        Me.lstTimeRange.ImeMode = CType(resources.GetObject("lstTimeRange.ImeMode"), System.Windows.Forms.ImeMode)
        Me.lstTimeRange.IntegralHeight = CType(resources.GetObject("lstTimeRange.IntegralHeight"), Boolean)
        Me.lstTimeRange.ItemHeight = CType(resources.GetObject("lstTimeRange.ItemHeight"), Integer)
        Me.lstTimeRange.Location = CType(resources.GetObject("lstTimeRange.Location"), System.Drawing.Point)
        Me.lstTimeRange.Name = "lstTimeRange"
        Me.lstTimeRange.RightToLeft = CType(resources.GetObject("lstTimeRange.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.lstTimeRange.ScrollAlwaysVisible = CType(resources.GetObject("lstTimeRange.ScrollAlwaysVisible"), Boolean)
        Me.lstTimeRange.Size = CType(resources.GetObject("lstTimeRange.Size"), System.Drawing.Size)
        Me.lstTimeRange.TabIndex = CType(resources.GetObject("lstTimeRange.TabIndex"), Integer)
        Me.lstTimeRange.TabStop = False
        Me.lstTimeRange.Visible = CType(resources.GetObject("lstTimeRange.Visible"), Boolean)
        '
        'lblBatch
        '
        Me.lblBatch.AccessibleDescription = resources.GetString("lblBatch.AccessibleDescription")
        Me.lblBatch.AccessibleName = resources.GetString("lblBatch.AccessibleName")
        Me.lblBatch.Anchor = CType(resources.GetObject("lblBatch.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.lblBatch.AutoSize = CType(resources.GetObject("lblBatch.AutoSize"), Boolean)
        Me.lblBatch.Dock = CType(resources.GetObject("lblBatch.Dock"), System.Windows.Forms.DockStyle)
        Me.lblBatch.Enabled = CType(resources.GetObject("lblBatch.Enabled"), Boolean)
        Me.lblBatch.Font = CType(resources.GetObject("lblBatch.Font"), System.Drawing.Font)
        Me.lblBatch.ForeColor = System.Drawing.SystemColors.Info
        Me.lblBatch.Image = CType(resources.GetObject("lblBatch.Image"), System.Drawing.Image)
        Me.lblBatch.ImageAlign = CType(resources.GetObject("lblBatch.ImageAlign"), System.Drawing.ContentAlignment)
        Me.lblBatch.ImageIndex = CType(resources.GetObject("lblBatch.ImageIndex"), Integer)
        Me.lblBatch.ImeMode = CType(resources.GetObject("lblBatch.ImeMode"), System.Windows.Forms.ImeMode)
        Me.lblBatch.Location = CType(resources.GetObject("lblBatch.Location"), System.Drawing.Point)
        Me.lblBatch.Name = "lblBatch"
        Me.lblBatch.RightToLeft = CType(resources.GetObject("lblBatch.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.lblBatch.Size = CType(resources.GetObject("lblBatch.Size"), System.Drawing.Size)
        Me.lblBatch.TabIndex = CType(resources.GetObject("lblBatch.TabIndex"), Integer)
        Me.lblBatch.Text = resources.GetString("lblBatch.Text")
        Me.lblBatch.TextAlign = CType(resources.GetObject("lblBatch.TextAlign"), System.Drawing.ContentAlignment)
        Me.lblBatch.Visible = CType(resources.GetObject("lblBatch.Visible"), Boolean)
        '
        'Label19
        '
        Me.Label19.AccessibleDescription = resources.GetString("Label19.AccessibleDescription")
        Me.Label19.AccessibleName = resources.GetString("Label19.AccessibleName")
        Me.Label19.Anchor = CType(resources.GetObject("Label19.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label19.AutoSize = CType(resources.GetObject("Label19.AutoSize"), Boolean)
        Me.Label19.Dock = CType(resources.GetObject("Label19.Dock"), System.Windows.Forms.DockStyle)
        Me.Label19.Enabled = CType(resources.GetObject("Label19.Enabled"), Boolean)
        Me.Label19.Font = CType(resources.GetObject("Label19.Font"), System.Drawing.Font)
        Me.Label19.ForeColor = System.Drawing.SystemColors.Info
        Me.Label19.Image = CType(resources.GetObject("Label19.Image"), System.Drawing.Image)
        Me.Label19.ImageAlign = CType(resources.GetObject("Label19.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label19.ImageIndex = CType(resources.GetObject("Label19.ImageIndex"), Integer)
        Me.Label19.ImeMode = CType(resources.GetObject("Label19.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label19.Location = CType(resources.GetObject("Label19.Location"), System.Drawing.Point)
        Me.Label19.Name = "Label19"
        Me.Label19.RightToLeft = CType(resources.GetObject("Label19.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label19.Size = CType(resources.GetObject("Label19.Size"), System.Drawing.Size)
        Me.Label19.TabIndex = CType(resources.GetObject("Label19.TabIndex"), Integer)
        Me.Label19.Text = resources.GetString("Label19.Text")
        Me.Label19.TextAlign = CType(resources.GetObject("Label19.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label19.Visible = CType(resources.GetObject("Label19.Visible"), Boolean)
        '
        'Label18
        '
        Me.Label18.AccessibleDescription = resources.GetString("Label18.AccessibleDescription")
        Me.Label18.AccessibleName = resources.GetString("Label18.AccessibleName")
        Me.Label18.Anchor = CType(resources.GetObject("Label18.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label18.AutoSize = CType(resources.GetObject("Label18.AutoSize"), Boolean)
        Me.Label18.Dock = CType(resources.GetObject("Label18.Dock"), System.Windows.Forms.DockStyle)
        Me.Label18.Enabled = CType(resources.GetObject("Label18.Enabled"), Boolean)
        Me.Label18.Font = CType(resources.GetObject("Label18.Font"), System.Drawing.Font)
        Me.Label18.ForeColor = System.Drawing.SystemColors.Info
        Me.Label18.Image = CType(resources.GetObject("Label18.Image"), System.Drawing.Image)
        Me.Label18.ImageAlign = CType(resources.GetObject("Label18.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label18.ImageIndex = CType(resources.GetObject("Label18.ImageIndex"), Integer)
        Me.Label18.ImeMode = CType(resources.GetObject("Label18.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label18.Location = CType(resources.GetObject("Label18.Location"), System.Drawing.Point)
        Me.Label18.Name = "Label18"
        Me.Label18.RightToLeft = CType(resources.GetObject("Label18.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label18.Size = CType(resources.GetObject("Label18.Size"), System.Drawing.Size)
        Me.Label18.TabIndex = CType(resources.GetObject("Label18.TabIndex"), Integer)
        Me.Label18.Text = resources.GetString("Label18.Text")
        Me.Label18.TextAlign = CType(resources.GetObject("Label18.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label18.Visible = CType(resources.GetObject("Label18.Visible"), Boolean)
        '
        'cmbTransType
        '
        Me.cmbTransType.AccessibleDescription = resources.GetString("cmbTransType.AccessibleDescription")
        Me.cmbTransType.AccessibleName = resources.GetString("cmbTransType.AccessibleName")
        Me.cmbTransType.Anchor = CType(resources.GetObject("cmbTransType.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.cmbTransType.BackgroundImage = CType(resources.GetObject("cmbTransType.BackgroundImage"), System.Drawing.Image)
        Me.cmbTransType.Dock = CType(resources.GetObject("cmbTransType.Dock"), System.Windows.Forms.DockStyle)
        Me.cmbTransType.Enabled = CType(resources.GetObject("cmbTransType.Enabled"), Boolean)
        Me.cmbTransType.Font = CType(resources.GetObject("cmbTransType.Font"), System.Drawing.Font)
        Me.cmbTransType.ImeMode = CType(resources.GetObject("cmbTransType.ImeMode"), System.Windows.Forms.ImeMode)
        Me.cmbTransType.IntegralHeight = CType(resources.GetObject("cmbTransType.IntegralHeight"), Boolean)
        Me.cmbTransType.ItemHeight = CType(resources.GetObject("cmbTransType.ItemHeight"), Integer)
        Me.cmbTransType.Items.AddRange(New Object() {resources.GetString("cmbTransType.Items"), resources.GetString("cmbTransType.Items1"), resources.GetString("cmbTransType.Items2"), resources.GetString("cmbTransType.Items3"), resources.GetString("cmbTransType.Items4")})
        Me.cmbTransType.Location = CType(resources.GetObject("cmbTransType.Location"), System.Drawing.Point)
        Me.cmbTransType.MaxDropDownItems = CType(resources.GetObject("cmbTransType.MaxDropDownItems"), Integer)
        Me.cmbTransType.MaxLength = CType(resources.GetObject("cmbTransType.MaxLength"), Integer)
        Me.cmbTransType.Name = "cmbTransType"
        Me.cmbTransType.RightToLeft = CType(resources.GetObject("cmbTransType.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.cmbTransType.Size = CType(resources.GetObject("cmbTransType.Size"), System.Drawing.Size)
        Me.cmbTransType.TabIndex = CType(resources.GetObject("cmbTransType.TabIndex"), Integer)
        Me.cmbTransType.TabStop = False
        Me.cmbTransType.Text = resources.GetString("cmbTransType.Text")
        Me.cmbTransType.Visible = CType(resources.GetObject("cmbTransType.Visible"), Boolean)
        '
        'txtTellerID
        '
        Me.txtTellerID.AccessibleDescription = resources.GetString("txtTellerID.AccessibleDescription")
        Me.txtTellerID.AccessibleName = resources.GetString("txtTellerID.AccessibleName")
        Me.txtTellerID.Anchor = CType(resources.GetObject("txtTellerID.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTellerID.AutoSize = CType(resources.GetObject("txtTellerID.AutoSize"), Boolean)
        Me.txtTellerID.BackgroundImage = CType(resources.GetObject("txtTellerID.BackgroundImage"), System.Drawing.Image)
        Me.txtTellerID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtTellerID.Dock = CType(resources.GetObject("txtTellerID.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTellerID.Enabled = CType(resources.GetObject("txtTellerID.Enabled"), Boolean)
        Me.txtTellerID.Font = CType(resources.GetObject("txtTellerID.Font"), System.Drawing.Font)
        Me.txtTellerID.ImeMode = CType(resources.GetObject("txtTellerID.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTellerID.Location = CType(resources.GetObject("txtTellerID.Location"), System.Drawing.Point)
        Me.txtTellerID.MaxLength = CType(resources.GetObject("txtTellerID.MaxLength"), Integer)
        Me.txtTellerID.Multiline = CType(resources.GetObject("txtTellerID.Multiline"), Boolean)
        Me.txtTellerID.Name = "txtTellerID"
        Me.txtTellerID.PasswordChar = CType(resources.GetObject("txtTellerID.PasswordChar"), Char)
        Me.txtTellerID.RightToLeft = CType(resources.GetObject("txtTellerID.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTellerID.ScrollBars = CType(resources.GetObject("txtTellerID.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTellerID.Size = CType(resources.GetObject("txtTellerID.Size"), System.Drawing.Size)
        Me.txtTellerID.TabIndex = CType(resources.GetObject("txtTellerID.TabIndex"), Integer)
        Me.txtTellerID.TabStop = False
        Me.txtTellerID.Text = resources.GetString("txtTellerID.Text")
        Me.txtTellerID.TextAlign = CType(resources.GetObject("txtTellerID.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTellerID.Visible = CType(resources.GetObject("txtTellerID.Visible"), Boolean)
        Me.txtTellerID.WordWrap = CType(resources.GetObject("txtTellerID.WordWrap"), Boolean)
        '
        'txtTimeTo
        '
        Me.txtTimeTo.AccessibleDescription = resources.GetString("txtTimeTo.AccessibleDescription")
        Me.txtTimeTo.AccessibleName = resources.GetString("txtTimeTo.AccessibleName")
        Me.txtTimeTo.Anchor = CType(resources.GetObject("txtTimeTo.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTimeTo.AutoSize = CType(resources.GetObject("txtTimeTo.AutoSize"), Boolean)
        Me.txtTimeTo.BackgroundImage = CType(resources.GetObject("txtTimeTo.BackgroundImage"), System.Drawing.Image)
        Me.txtTimeTo.Dock = CType(resources.GetObject("txtTimeTo.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTimeTo.Enabled = CType(resources.GetObject("txtTimeTo.Enabled"), Boolean)
        Me.txtTimeTo.Font = CType(resources.GetObject("txtTimeTo.Font"), System.Drawing.Font)
        Me.txtTimeTo.ImeMode = CType(resources.GetObject("txtTimeTo.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTimeTo.Location = CType(resources.GetObject("txtTimeTo.Location"), System.Drawing.Point)
        Me.txtTimeTo.MaxLength = CType(resources.GetObject("txtTimeTo.MaxLength"), Integer)
        Me.txtTimeTo.Multiline = CType(resources.GetObject("txtTimeTo.Multiline"), Boolean)
        Me.txtTimeTo.Name = "txtTimeTo"
        Me.txtTimeTo.PasswordChar = CType(resources.GetObject("txtTimeTo.PasswordChar"), Char)
        Me.txtTimeTo.RightToLeft = CType(resources.GetObject("txtTimeTo.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTimeTo.ScrollBars = CType(resources.GetObject("txtTimeTo.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTimeTo.Size = CType(resources.GetObject("txtTimeTo.Size"), System.Drawing.Size)
        Me.txtTimeTo.TabIndex = CType(resources.GetObject("txtTimeTo.TabIndex"), Integer)
        Me.txtTimeTo.TabStop = False
        Me.txtTimeTo.Text = resources.GetString("txtTimeTo.Text")
        Me.txtTimeTo.TextAlign = CType(resources.GetObject("txtTimeTo.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTimeTo.Visible = CType(resources.GetObject("txtTimeTo.Visible"), Boolean)
        Me.txtTimeTo.WordWrap = CType(resources.GetObject("txtTimeTo.WordWrap"), Boolean)
        '
        'txtTimeFrom
        '
        Me.txtTimeFrom.AccessibleDescription = resources.GetString("txtTimeFrom.AccessibleDescription")
        Me.txtTimeFrom.AccessibleName = resources.GetString("txtTimeFrom.AccessibleName")
        Me.txtTimeFrom.Anchor = CType(resources.GetObject("txtTimeFrom.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTimeFrom.AutoSize = CType(resources.GetObject("txtTimeFrom.AutoSize"), Boolean)
        Me.txtTimeFrom.BackgroundImage = CType(resources.GetObject("txtTimeFrom.BackgroundImage"), System.Drawing.Image)
        Me.txtTimeFrom.Dock = CType(resources.GetObject("txtTimeFrom.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTimeFrom.Enabled = CType(resources.GetObject("txtTimeFrom.Enabled"), Boolean)
        Me.txtTimeFrom.Font = CType(resources.GetObject("txtTimeFrom.Font"), System.Drawing.Font)
        Me.txtTimeFrom.ImeMode = CType(resources.GetObject("txtTimeFrom.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTimeFrom.Location = CType(resources.GetObject("txtTimeFrom.Location"), System.Drawing.Point)
        Me.txtTimeFrom.MaxLength = CType(resources.GetObject("txtTimeFrom.MaxLength"), Integer)
        Me.txtTimeFrom.Multiline = CType(resources.GetObject("txtTimeFrom.Multiline"), Boolean)
        Me.txtTimeFrom.Name = "txtTimeFrom"
        Me.txtTimeFrom.PasswordChar = CType(resources.GetObject("txtTimeFrom.PasswordChar"), Char)
        Me.txtTimeFrom.RightToLeft = CType(resources.GetObject("txtTimeFrom.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTimeFrom.ScrollBars = CType(resources.GetObject("txtTimeFrom.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTimeFrom.Size = CType(resources.GetObject("txtTimeFrom.Size"), System.Drawing.Size)
        Me.txtTimeFrom.TabIndex = CType(resources.GetObject("txtTimeFrom.TabIndex"), Integer)
        Me.txtTimeFrom.TabStop = False
        Me.txtTimeFrom.Text = resources.GetString("txtTimeFrom.Text")
        Me.txtTimeFrom.TextAlign = CType(resources.GetObject("txtTimeFrom.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTimeFrom.Visible = CType(resources.GetObject("txtTimeFrom.Visible"), Boolean)
        Me.txtTimeFrom.WordWrap = CType(resources.GetObject("txtTimeFrom.WordWrap"), Boolean)
        '
        'lblTimeTo
        '
        Me.lblTimeTo.AccessibleDescription = resources.GetString("lblTimeTo.AccessibleDescription")
        Me.lblTimeTo.AccessibleName = resources.GetString("lblTimeTo.AccessibleName")
        Me.lblTimeTo.Anchor = CType(resources.GetObject("lblTimeTo.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.lblTimeTo.AutoSize = CType(resources.GetObject("lblTimeTo.AutoSize"), Boolean)
        Me.lblTimeTo.Dock = CType(resources.GetObject("lblTimeTo.Dock"), System.Windows.Forms.DockStyle)
        Me.lblTimeTo.Enabled = CType(resources.GetObject("lblTimeTo.Enabled"), Boolean)
        Me.lblTimeTo.Font = CType(resources.GetObject("lblTimeTo.Font"), System.Drawing.Font)
        Me.lblTimeTo.ForeColor = System.Drawing.SystemColors.Info
        Me.lblTimeTo.Image = CType(resources.GetObject("lblTimeTo.Image"), System.Drawing.Image)
        Me.lblTimeTo.ImageAlign = CType(resources.GetObject("lblTimeTo.ImageAlign"), System.Drawing.ContentAlignment)
        Me.lblTimeTo.ImageIndex = CType(resources.GetObject("lblTimeTo.ImageIndex"), Integer)
        Me.lblTimeTo.ImeMode = CType(resources.GetObject("lblTimeTo.ImeMode"), System.Windows.Forms.ImeMode)
        Me.lblTimeTo.Location = CType(resources.GetObject("lblTimeTo.Location"), System.Drawing.Point)
        Me.lblTimeTo.Name = "lblTimeTo"
        Me.lblTimeTo.RightToLeft = CType(resources.GetObject("lblTimeTo.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.lblTimeTo.Size = CType(resources.GetObject("lblTimeTo.Size"), System.Drawing.Size)
        Me.lblTimeTo.TabIndex = CType(resources.GetObject("lblTimeTo.TabIndex"), Integer)
        Me.lblTimeTo.Text = resources.GetString("lblTimeTo.Text")
        Me.lblTimeTo.TextAlign = CType(resources.GetObject("lblTimeTo.TextAlign"), System.Drawing.ContentAlignment)
        Me.lblTimeTo.Visible = CType(resources.GetObject("lblTimeTo.Visible"), Boolean)
        '
        'lblTimeFrom
        '
        Me.lblTimeFrom.AccessibleDescription = resources.GetString("lblTimeFrom.AccessibleDescription")
        Me.lblTimeFrom.AccessibleName = resources.GetString("lblTimeFrom.AccessibleName")
        Me.lblTimeFrom.Anchor = CType(resources.GetObject("lblTimeFrom.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.lblTimeFrom.AutoSize = CType(resources.GetObject("lblTimeFrom.AutoSize"), Boolean)
        Me.lblTimeFrom.Dock = CType(resources.GetObject("lblTimeFrom.Dock"), System.Windows.Forms.DockStyle)
        Me.lblTimeFrom.Enabled = CType(resources.GetObject("lblTimeFrom.Enabled"), Boolean)
        Me.lblTimeFrom.Font = CType(resources.GetObject("lblTimeFrom.Font"), System.Drawing.Font)
        Me.lblTimeFrom.ForeColor = System.Drawing.SystemColors.Info
        Me.lblTimeFrom.Image = CType(resources.GetObject("lblTimeFrom.Image"), System.Drawing.Image)
        Me.lblTimeFrom.ImageAlign = CType(resources.GetObject("lblTimeFrom.ImageAlign"), System.Drawing.ContentAlignment)
        Me.lblTimeFrom.ImageIndex = CType(resources.GetObject("lblTimeFrom.ImageIndex"), Integer)
        Me.lblTimeFrom.ImeMode = CType(resources.GetObject("lblTimeFrom.ImeMode"), System.Windows.Forms.ImeMode)
        Me.lblTimeFrom.Location = CType(resources.GetObject("lblTimeFrom.Location"), System.Drawing.Point)
        Me.lblTimeFrom.Name = "lblTimeFrom"
        Me.lblTimeFrom.RightToLeft = CType(resources.GetObject("lblTimeFrom.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.lblTimeFrom.Size = CType(resources.GetObject("lblTimeFrom.Size"), System.Drawing.Size)
        Me.lblTimeFrom.TabIndex = CType(resources.GetObject("lblTimeFrom.TabIndex"), Integer)
        Me.lblTimeFrom.Text = resources.GetString("lblTimeFrom.Text")
        Me.lblTimeFrom.TextAlign = CType(resources.GetObject("lblTimeFrom.TextAlign"), System.Drawing.ContentAlignment)
        Me.lblTimeFrom.Visible = CType(resources.GetObject("lblTimeFrom.Visible"), Boolean)
        '
        'Label15
        '
        Me.Label15.AccessibleDescription = resources.GetString("Label15.AccessibleDescription")
        Me.Label15.AccessibleName = resources.GetString("Label15.AccessibleName")
        Me.Label15.Anchor = CType(resources.GetObject("Label15.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label15.AutoSize = CType(resources.GetObject("Label15.AutoSize"), Boolean)
        Me.Label15.Dock = CType(resources.GetObject("Label15.Dock"), System.Windows.Forms.DockStyle)
        Me.Label15.Enabled = CType(resources.GetObject("Label15.Enabled"), Boolean)
        Me.Label15.Font = CType(resources.GetObject("Label15.Font"), System.Drawing.Font)
        Me.Label15.ForeColor = System.Drawing.SystemColors.Info
        Me.Label15.Image = CType(resources.GetObject("Label15.Image"), System.Drawing.Image)
        Me.Label15.ImageAlign = CType(resources.GetObject("Label15.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label15.ImageIndex = CType(resources.GetObject("Label15.ImageIndex"), Integer)
        Me.Label15.ImeMode = CType(resources.GetObject("Label15.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label15.Location = CType(resources.GetObject("Label15.Location"), System.Drawing.Point)
        Me.Label15.Name = "Label15"
        Me.Label15.RightToLeft = CType(resources.GetObject("Label15.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label15.Size = CType(resources.GetObject("Label15.Size"), System.Drawing.Size)
        Me.Label15.TabIndex = CType(resources.GetObject("Label15.TabIndex"), Integer)
        Me.Label15.Text = resources.GetString("Label15.Text")
        Me.Label15.TextAlign = CType(resources.GetObject("Label15.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label15.Visible = CType(resources.GetObject("Label15.Visible"), Boolean)
        '
        'dtePeriod
        '
        Me.dtePeriod.AccessibleDescription = resources.GetString("dtePeriod.AccessibleDescription")
        Me.dtePeriod.AccessibleName = resources.GetString("dtePeriod.AccessibleName")
        Me.dtePeriod.AccessibleRole = System.Windows.Forms.AccessibleRole.MenuPopup
        Me.dtePeriod.Anchor = CType(resources.GetObject("dtePeriod.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.dtePeriod.BackgroundImage = CType(resources.GetObject("dtePeriod.BackgroundImage"), System.Drawing.Image)
        Me.dtePeriod.CalendarFont = CType(resources.GetObject("dtePeriod.CalendarFont"), System.Drawing.Font)
        Me.dtePeriod.CalendarMonthBackground = System.Drawing.Color.LightSteelBlue
        Me.dtePeriod.CalendarTitleBackColor = System.Drawing.Color.SteelBlue
        Me.dtePeriod.Cursor = System.Windows.Forms.Cursors.Hand
        Me.dtePeriod.CustomFormat = "mm/dd/yy"
        Me.dtePeriod.Dock = CType(resources.GetObject("dtePeriod.Dock"), System.Windows.Forms.DockStyle)
        Me.dtePeriod.DropDownAlign = CType(resources.GetObject("dtePeriod.DropDownAlign"), System.Windows.Forms.LeftRightAlignment)
        Me.dtePeriod.Enabled = CType(resources.GetObject("dtePeriod.Enabled"), Boolean)
        Me.dtePeriod.Font = CType(resources.GetObject("dtePeriod.Font"), System.Drawing.Font)
        Me.dtePeriod.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtePeriod.ImeMode = CType(resources.GetObject("dtePeriod.ImeMode"), System.Windows.Forms.ImeMode)
        Me.dtePeriod.Location = CType(resources.GetObject("dtePeriod.Location"), System.Drawing.Point)
        Me.dtePeriod.Name = "dtePeriod"
        Me.dtePeriod.RightToLeft = CType(resources.GetObject("dtePeriod.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.dtePeriod.Size = CType(resources.GetObject("dtePeriod.Size"), System.Drawing.Size)
        Me.dtePeriod.TabIndex = CType(resources.GetObject("dtePeriod.TabIndex"), Integer)
        Me.dtePeriod.TabStop = False
        Me.dtePeriod.Visible = CType(resources.GetObject("dtePeriod.Visible"), Boolean)
        '
        'Label14
        '
        Me.Label14.AccessibleDescription = resources.GetString("Label14.AccessibleDescription")
        Me.Label14.AccessibleName = resources.GetString("Label14.AccessibleName")
        Me.Label14.Anchor = CType(resources.GetObject("Label14.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label14.AutoSize = CType(resources.GetObject("Label14.AutoSize"), Boolean)
        Me.Label14.Dock = CType(resources.GetObject("Label14.Dock"), System.Windows.Forms.DockStyle)
        Me.Label14.Enabled = CType(resources.GetObject("Label14.Enabled"), Boolean)
        Me.Label14.Font = CType(resources.GetObject("Label14.Font"), System.Drawing.Font)
        Me.Label14.ForeColor = System.Drawing.SystemColors.Info
        Me.Label14.Image = CType(resources.GetObject("Label14.Image"), System.Drawing.Image)
        Me.Label14.ImageAlign = CType(resources.GetObject("Label14.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label14.ImageIndex = CType(resources.GetObject("Label14.ImageIndex"), Integer)
        Me.Label14.ImeMode = CType(resources.GetObject("Label14.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label14.Location = CType(resources.GetObject("Label14.Location"), System.Drawing.Point)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = CType(resources.GetObject("Label14.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label14.Size = CType(resources.GetObject("Label14.Size"), System.Drawing.Size)
        Me.Label14.TabIndex = CType(resources.GetObject("Label14.TabIndex"), Integer)
        Me.Label14.Text = resources.GetString("Label14.Text")
        Me.Label14.TextAlign = CType(resources.GetObject("Label14.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label14.Visible = CType(resources.GetObject("Label14.Visible"), Boolean)
        '
        'Label13
        '
        Me.Label13.AccessibleDescription = resources.GetString("Label13.AccessibleDescription")
        Me.Label13.AccessibleName = resources.GetString("Label13.AccessibleName")
        Me.Label13.Anchor = CType(resources.GetObject("Label13.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label13.AutoSize = CType(resources.GetObject("Label13.AutoSize"), Boolean)
        Me.Label13.Dock = CType(resources.GetObject("Label13.Dock"), System.Windows.Forms.DockStyle)
        Me.Label13.Enabled = CType(resources.GetObject("Label13.Enabled"), Boolean)
        Me.Label13.Font = CType(resources.GetObject("Label13.Font"), System.Drawing.Font)
        Me.Label13.ForeColor = System.Drawing.SystemColors.Info
        Me.Label13.Image = CType(resources.GetObject("Label13.Image"), System.Drawing.Image)
        Me.Label13.ImageAlign = CType(resources.GetObject("Label13.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label13.ImageIndex = CType(resources.GetObject("Label13.ImageIndex"), Integer)
        Me.Label13.ImeMode = CType(resources.GetObject("Label13.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label13.Location = CType(resources.GetObject("Label13.Location"), System.Drawing.Point)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = CType(resources.GetObject("Label13.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label13.Size = CType(resources.GetObject("Label13.Size"), System.Drawing.Size)
        Me.Label13.TabIndex = CType(resources.GetObject("Label13.TabIndex"), Integer)
        Me.Label13.Text = resources.GetString("Label13.Text")
        Me.Label13.TextAlign = CType(resources.GetObject("Label13.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label13.Visible = CType(resources.GetObject("Label13.Visible"), Boolean)
        '
        'lblTimeRange
        '
        Me.lblTimeRange.AccessibleDescription = resources.GetString("lblTimeRange.AccessibleDescription")
        Me.lblTimeRange.AccessibleName = resources.GetString("lblTimeRange.AccessibleName")
        Me.lblTimeRange.Anchor = CType(resources.GetObject("lblTimeRange.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.lblTimeRange.AutoSize = CType(resources.GetObject("lblTimeRange.AutoSize"), Boolean)
        Me.lblTimeRange.Dock = CType(resources.GetObject("lblTimeRange.Dock"), System.Windows.Forms.DockStyle)
        Me.lblTimeRange.Enabled = CType(resources.GetObject("lblTimeRange.Enabled"), Boolean)
        Me.lblTimeRange.Font = CType(resources.GetObject("lblTimeRange.Font"), System.Drawing.Font)
        Me.lblTimeRange.ForeColor = System.Drawing.SystemColors.Info
        Me.lblTimeRange.Image = CType(resources.GetObject("lblTimeRange.Image"), System.Drawing.Image)
        Me.lblTimeRange.ImageAlign = CType(resources.GetObject("lblTimeRange.ImageAlign"), System.Drawing.ContentAlignment)
        Me.lblTimeRange.ImageIndex = CType(resources.GetObject("lblTimeRange.ImageIndex"), Integer)
        Me.lblTimeRange.ImeMode = CType(resources.GetObject("lblTimeRange.ImeMode"), System.Windows.Forms.ImeMode)
        Me.lblTimeRange.Location = CType(resources.GetObject("lblTimeRange.Location"), System.Drawing.Point)
        Me.lblTimeRange.Name = "lblTimeRange"
        Me.lblTimeRange.RightToLeft = CType(resources.GetObject("lblTimeRange.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.lblTimeRange.Size = CType(resources.GetObject("lblTimeRange.Size"), System.Drawing.Size)
        Me.lblTimeRange.TabIndex = CType(resources.GetObject("lblTimeRange.TabIndex"), Integer)
        Me.lblTimeRange.Text = resources.GetString("lblTimeRange.Text")
        Me.lblTimeRange.TextAlign = CType(resources.GetObject("lblTimeRange.TextAlign"), System.Drawing.ContentAlignment)
        Me.lblTimeRange.Visible = CType(resources.GetObject("lblTimeRange.Visible"), Boolean)
        '
        'gbCash
        '
        Me.gbCash.AccessibleDescription = resources.GetString("gbCash.AccessibleDescription")
        Me.gbCash.AccessibleName = resources.GetString("gbCash.AccessibleName")
        Me.gbCash.Anchor = CType(resources.GetObject("gbCash.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.gbCash.BackgroundImage = CType(resources.GetObject("gbCash.BackgroundImage"), System.Drawing.Image)
        Me.gbCash.Controls.Add(Me.Label35)
        Me.gbCash.Controls.Add(Me.Label34)
        Me.gbCash.Controls.Add(Me.txt001)
        Me.gbCash.Controls.Add(Me.txt005)
        Me.gbCash.Controls.Add(Me.txt010)
        Me.gbCash.Controls.Add(Me.txt025)
        Me.gbCash.Controls.Add(Me.txt1)
        Me.gbCash.Controls.Add(Me.txt5)
        Me.gbCash.Controls.Add(Me.txt10)
        Me.gbCash.Controls.Add(Me.txt20)
        Me.gbCash.Controls.Add(Me.txt50)
        Me.gbCash.Controls.Add(Me.txt100)
        Me.gbCash.Controls.Add(Me.txt200)
        Me.gbCash.Controls.Add(Me.txt500)
        Me.gbCash.Controls.Add(Me.txt1000)
        Me.gbCash.Controls.Add(Me.txtTot001)
        Me.gbCash.Controls.Add(Me.txtTot005)
        Me.gbCash.Controls.Add(Me.txtTot010)
        Me.gbCash.Controls.Add(Me.txtTot025)
        Me.gbCash.Controls.Add(Me.Label33)
        Me.gbCash.Controls.Add(Me.Label32)
        Me.gbCash.Controls.Add(Me.Label31)
        Me.gbCash.Controls.Add(Me.Label30)
        Me.gbCash.Controls.Add(Me.Label29)
        Me.gbCash.Controls.Add(Me.Label28)
        Me.gbCash.Controls.Add(Me.Label27)
        Me.gbCash.Controls.Add(Me.Label26)
        Me.gbCash.Controls.Add(Me.Label25)
        Me.gbCash.Controls.Add(Me.Label24)
        Me.gbCash.Controls.Add(Me.Label23)
        Me.gbCash.Controls.Add(Me.txtTot1)
        Me.gbCash.Controls.Add(Me.txtTot5)
        Me.gbCash.Controls.Add(Me.txtTot10)
        Me.gbCash.Controls.Add(Me.txtTot20)
        Me.gbCash.Controls.Add(Me.txtTot50)
        Me.gbCash.Controls.Add(Me.txtTot100)
        Me.gbCash.Controls.Add(Me.txtTot200)
        Me.gbCash.Controls.Add(Me.txtTot500)
        Me.gbCash.Controls.Add(Me.Label22)
        Me.gbCash.Controls.Add(Me.txtTot1000)
        Me.gbCash.Controls.Add(Me.Label21)
        Me.gbCash.Controls.Add(Me.Label20)
        Me.gbCash.Controls.Add(Me.Label10)
        Me.gbCash.Controls.Add(Me.txtAmtLeft)
        Me.gbCash.Controls.Add(Me.txtTotCash)
        Me.gbCash.Controls.Add(Me.txtExCash)
        Me.gbCash.Controls.Add(Me.Label3)
        Me.gbCash.Controls.Add(Me.Label2)
        Me.gbCash.Controls.Add(Me.Label1)
        Me.gbCash.Dock = CType(resources.GetObject("gbCash.Dock"), System.Windows.Forms.DockStyle)
        Me.gbCash.Enabled = CType(resources.GetObject("gbCash.Enabled"), Boolean)
        Me.gbCash.Font = CType(resources.GetObject("gbCash.Font"), System.Drawing.Font)
        Me.gbCash.ImeMode = CType(resources.GetObject("gbCash.ImeMode"), System.Windows.Forms.ImeMode)
        Me.gbCash.Location = CType(resources.GetObject("gbCash.Location"), System.Drawing.Point)
        Me.gbCash.Name = "gbCash"
        Me.gbCash.RightToLeft = CType(resources.GetObject("gbCash.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.gbCash.Size = CType(resources.GetObject("gbCash.Size"), System.Drawing.Size)
        Me.gbCash.TabIndex = CType(resources.GetObject("gbCash.TabIndex"), Integer)
        Me.gbCash.TabStop = False
        Me.gbCash.Text = resources.GetString("gbCash.Text")
        Me.gbCash.Visible = CType(resources.GetObject("gbCash.Visible"), Boolean)
        '
        'Label35
        '
        Me.Label35.AccessibleDescription = resources.GetString("Label35.AccessibleDescription")
        Me.Label35.AccessibleName = resources.GetString("Label35.AccessibleName")
        Me.Label35.Anchor = CType(resources.GetObject("Label35.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label35.AutoSize = CType(resources.GetObject("Label35.AutoSize"), Boolean)
        Me.Label35.BackColor = System.Drawing.Color.MidnightBlue
        Me.Label35.Dock = CType(resources.GetObject("Label35.Dock"), System.Windows.Forms.DockStyle)
        Me.Label35.Enabled = CType(resources.GetObject("Label35.Enabled"), Boolean)
        Me.Label35.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label35.Font = CType(resources.GetObject("Label35.Font"), System.Drawing.Font)
        Me.Label35.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label35.Image = CType(resources.GetObject("Label35.Image"), System.Drawing.Image)
        Me.Label35.ImageAlign = CType(resources.GetObject("Label35.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label35.ImageIndex = CType(resources.GetObject("Label35.ImageIndex"), Integer)
        Me.Label35.ImeMode = CType(resources.GetObject("Label35.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label35.Location = CType(resources.GetObject("Label35.Location"), System.Drawing.Point)
        Me.Label35.Name = "Label35"
        Me.Label35.RightToLeft = CType(resources.GetObject("Label35.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label35.Size = CType(resources.GetObject("Label35.Size"), System.Drawing.Size)
        Me.Label35.TabIndex = CType(resources.GetObject("Label35.TabIndex"), Integer)
        Me.Label35.Text = resources.GetString("Label35.Text")
        Me.Label35.TextAlign = CType(resources.GetObject("Label35.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label35.Visible = CType(resources.GetObject("Label35.Visible"), Boolean)
        '
        'Label34
        '
        Me.Label34.AccessibleDescription = resources.GetString("Label34.AccessibleDescription")
        Me.Label34.AccessibleName = resources.GetString("Label34.AccessibleName")
        Me.Label34.Anchor = CType(resources.GetObject("Label34.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label34.AutoSize = CType(resources.GetObject("Label34.AutoSize"), Boolean)
        Me.Label34.BackColor = System.Drawing.Color.MidnightBlue
        Me.Label34.Dock = CType(resources.GetObject("Label34.Dock"), System.Windows.Forms.DockStyle)
        Me.Label34.Enabled = CType(resources.GetObject("Label34.Enabled"), Boolean)
        Me.Label34.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label34.Font = CType(resources.GetObject("Label34.Font"), System.Drawing.Font)
        Me.Label34.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label34.Image = CType(resources.GetObject("Label34.Image"), System.Drawing.Image)
        Me.Label34.ImageAlign = CType(resources.GetObject("Label34.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label34.ImageIndex = CType(resources.GetObject("Label34.ImageIndex"), Integer)
        Me.Label34.ImeMode = CType(resources.GetObject("Label34.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label34.Location = CType(resources.GetObject("Label34.Location"), System.Drawing.Point)
        Me.Label34.Name = "Label34"
        Me.Label34.RightToLeft = CType(resources.GetObject("Label34.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label34.Size = CType(resources.GetObject("Label34.Size"), System.Drawing.Size)
        Me.Label34.TabIndex = CType(resources.GetObject("Label34.TabIndex"), Integer)
        Me.Label34.Text = resources.GetString("Label34.Text")
        Me.Label34.TextAlign = CType(resources.GetObject("Label34.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label34.Visible = CType(resources.GetObject("Label34.Visible"), Boolean)
        '
        'txt001
        '
        Me.txt001.AccessibleDescription = resources.GetString("txt001.AccessibleDescription")
        Me.txt001.AccessibleName = resources.GetString("txt001.AccessibleName")
        Me.txt001.Anchor = CType(resources.GetObject("txt001.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txt001.AutoSize = CType(resources.GetObject("txt001.AutoSize"), Boolean)
        Me.txt001.BackColor = System.Drawing.Color.White
        Me.txt001.BackgroundImage = CType(resources.GetObject("txt001.BackgroundImage"), System.Drawing.Image)
        Me.txt001.Dock = CType(resources.GetObject("txt001.Dock"), System.Windows.Forms.DockStyle)
        Me.txt001.Enabled = CType(resources.GetObject("txt001.Enabled"), Boolean)
        Me.txt001.Font = CType(resources.GetObject("txt001.Font"), System.Drawing.Font)
        Me.txt001.ImeMode = CType(resources.GetObject("txt001.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txt001.Location = CType(resources.GetObject("txt001.Location"), System.Drawing.Point)
        Me.txt001.MaxLength = CType(resources.GetObject("txt001.MaxLength"), Integer)
        Me.txt001.Multiline = CType(resources.GetObject("txt001.Multiline"), Boolean)
        Me.txt001.Name = "txt001"
        Me.txt001.PasswordChar = CType(resources.GetObject("txt001.PasswordChar"), Char)
        Me.txt001.RightToLeft = CType(resources.GetObject("txt001.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txt001.ScrollBars = CType(resources.GetObject("txt001.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txt001.Size = CType(resources.GetObject("txt001.Size"), System.Drawing.Size)
        Me.txt001.TabIndex = CType(resources.GetObject("txt001.TabIndex"), Integer)
        Me.txt001.TabStop = False
        Me.txt001.Text = resources.GetString("txt001.Text")
        Me.txt001.TextAlign = CType(resources.GetObject("txt001.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txt001.Visible = CType(resources.GetObject("txt001.Visible"), Boolean)
        Me.txt001.WordWrap = CType(resources.GetObject("txt001.WordWrap"), Boolean)
        '
        'txt005
        '
        Me.txt005.AccessibleDescription = resources.GetString("txt005.AccessibleDescription")
        Me.txt005.AccessibleName = resources.GetString("txt005.AccessibleName")
        Me.txt005.Anchor = CType(resources.GetObject("txt005.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txt005.AutoSize = CType(resources.GetObject("txt005.AutoSize"), Boolean)
        Me.txt005.BackColor = System.Drawing.Color.White
        Me.txt005.BackgroundImage = CType(resources.GetObject("txt005.BackgroundImage"), System.Drawing.Image)
        Me.txt005.Dock = CType(resources.GetObject("txt005.Dock"), System.Windows.Forms.DockStyle)
        Me.txt005.Enabled = CType(resources.GetObject("txt005.Enabled"), Boolean)
        Me.txt005.Font = CType(resources.GetObject("txt005.Font"), System.Drawing.Font)
        Me.txt005.ImeMode = CType(resources.GetObject("txt005.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txt005.Location = CType(resources.GetObject("txt005.Location"), System.Drawing.Point)
        Me.txt005.MaxLength = CType(resources.GetObject("txt005.MaxLength"), Integer)
        Me.txt005.Multiline = CType(resources.GetObject("txt005.Multiline"), Boolean)
        Me.txt005.Name = "txt005"
        Me.txt005.PasswordChar = CType(resources.GetObject("txt005.PasswordChar"), Char)
        Me.txt005.RightToLeft = CType(resources.GetObject("txt005.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txt005.ScrollBars = CType(resources.GetObject("txt005.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txt005.Size = CType(resources.GetObject("txt005.Size"), System.Drawing.Size)
        Me.txt005.TabIndex = CType(resources.GetObject("txt005.TabIndex"), Integer)
        Me.txt005.TabStop = False
        Me.txt005.Text = resources.GetString("txt005.Text")
        Me.txt005.TextAlign = CType(resources.GetObject("txt005.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txt005.Visible = CType(resources.GetObject("txt005.Visible"), Boolean)
        Me.txt005.WordWrap = CType(resources.GetObject("txt005.WordWrap"), Boolean)
        '
        'txt010
        '
        Me.txt010.AccessibleDescription = resources.GetString("txt010.AccessibleDescription")
        Me.txt010.AccessibleName = resources.GetString("txt010.AccessibleName")
        Me.txt010.Anchor = CType(resources.GetObject("txt010.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txt010.AutoSize = CType(resources.GetObject("txt010.AutoSize"), Boolean)
        Me.txt010.BackColor = System.Drawing.Color.White
        Me.txt010.BackgroundImage = CType(resources.GetObject("txt010.BackgroundImage"), System.Drawing.Image)
        Me.txt010.Dock = CType(resources.GetObject("txt010.Dock"), System.Windows.Forms.DockStyle)
        Me.txt010.Enabled = CType(resources.GetObject("txt010.Enabled"), Boolean)
        Me.txt010.Font = CType(resources.GetObject("txt010.Font"), System.Drawing.Font)
        Me.txt010.ImeMode = CType(resources.GetObject("txt010.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txt010.Location = CType(resources.GetObject("txt010.Location"), System.Drawing.Point)
        Me.txt010.MaxLength = CType(resources.GetObject("txt010.MaxLength"), Integer)
        Me.txt010.Multiline = CType(resources.GetObject("txt010.Multiline"), Boolean)
        Me.txt010.Name = "txt010"
        Me.txt010.PasswordChar = CType(resources.GetObject("txt010.PasswordChar"), Char)
        Me.txt010.RightToLeft = CType(resources.GetObject("txt010.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txt010.ScrollBars = CType(resources.GetObject("txt010.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txt010.Size = CType(resources.GetObject("txt010.Size"), System.Drawing.Size)
        Me.txt010.TabIndex = CType(resources.GetObject("txt010.TabIndex"), Integer)
        Me.txt010.TabStop = False
        Me.txt010.Text = resources.GetString("txt010.Text")
        Me.txt010.TextAlign = CType(resources.GetObject("txt010.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txt010.Visible = CType(resources.GetObject("txt010.Visible"), Boolean)
        Me.txt010.WordWrap = CType(resources.GetObject("txt010.WordWrap"), Boolean)
        '
        'txt025
        '
        Me.txt025.AccessibleDescription = resources.GetString("txt025.AccessibleDescription")
        Me.txt025.AccessibleName = resources.GetString("txt025.AccessibleName")
        Me.txt025.Anchor = CType(resources.GetObject("txt025.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txt025.AutoSize = CType(resources.GetObject("txt025.AutoSize"), Boolean)
        Me.txt025.BackColor = System.Drawing.Color.White
        Me.txt025.BackgroundImage = CType(resources.GetObject("txt025.BackgroundImage"), System.Drawing.Image)
        Me.txt025.Dock = CType(resources.GetObject("txt025.Dock"), System.Windows.Forms.DockStyle)
        Me.txt025.Enabled = CType(resources.GetObject("txt025.Enabled"), Boolean)
        Me.txt025.Font = CType(resources.GetObject("txt025.Font"), System.Drawing.Font)
        Me.txt025.ImeMode = CType(resources.GetObject("txt025.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txt025.Location = CType(resources.GetObject("txt025.Location"), System.Drawing.Point)
        Me.txt025.MaxLength = CType(resources.GetObject("txt025.MaxLength"), Integer)
        Me.txt025.Multiline = CType(resources.GetObject("txt025.Multiline"), Boolean)
        Me.txt025.Name = "txt025"
        Me.txt025.PasswordChar = CType(resources.GetObject("txt025.PasswordChar"), Char)
        Me.txt025.RightToLeft = CType(resources.GetObject("txt025.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txt025.ScrollBars = CType(resources.GetObject("txt025.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txt025.Size = CType(resources.GetObject("txt025.Size"), System.Drawing.Size)
        Me.txt025.TabIndex = CType(resources.GetObject("txt025.TabIndex"), Integer)
        Me.txt025.TabStop = False
        Me.txt025.Text = resources.GetString("txt025.Text")
        Me.txt025.TextAlign = CType(resources.GetObject("txt025.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txt025.Visible = CType(resources.GetObject("txt025.Visible"), Boolean)
        Me.txt025.WordWrap = CType(resources.GetObject("txt025.WordWrap"), Boolean)
        '
        'txt1
        '
        Me.txt1.AccessibleDescription = resources.GetString("txt1.AccessibleDescription")
        Me.txt1.AccessibleName = resources.GetString("txt1.AccessibleName")
        Me.txt1.Anchor = CType(resources.GetObject("txt1.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txt1.AutoSize = CType(resources.GetObject("txt1.AutoSize"), Boolean)
        Me.txt1.BackColor = System.Drawing.Color.White
        Me.txt1.BackgroundImage = CType(resources.GetObject("txt1.BackgroundImage"), System.Drawing.Image)
        Me.txt1.Dock = CType(resources.GetObject("txt1.Dock"), System.Windows.Forms.DockStyle)
        Me.txt1.Enabled = CType(resources.GetObject("txt1.Enabled"), Boolean)
        Me.txt1.Font = CType(resources.GetObject("txt1.Font"), System.Drawing.Font)
        Me.txt1.ImeMode = CType(resources.GetObject("txt1.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txt1.Location = CType(resources.GetObject("txt1.Location"), System.Drawing.Point)
        Me.txt1.MaxLength = CType(resources.GetObject("txt1.MaxLength"), Integer)
        Me.txt1.Multiline = CType(resources.GetObject("txt1.Multiline"), Boolean)
        Me.txt1.Name = "txt1"
        Me.txt1.PasswordChar = CType(resources.GetObject("txt1.PasswordChar"), Char)
        Me.txt1.RightToLeft = CType(resources.GetObject("txt1.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txt1.ScrollBars = CType(resources.GetObject("txt1.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txt1.Size = CType(resources.GetObject("txt1.Size"), System.Drawing.Size)
        Me.txt1.TabIndex = CType(resources.GetObject("txt1.TabIndex"), Integer)
        Me.txt1.TabStop = False
        Me.txt1.Text = resources.GetString("txt1.Text")
        Me.txt1.TextAlign = CType(resources.GetObject("txt1.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txt1.Visible = CType(resources.GetObject("txt1.Visible"), Boolean)
        Me.txt1.WordWrap = CType(resources.GetObject("txt1.WordWrap"), Boolean)
        '
        'txt5
        '
        Me.txt5.AccessibleDescription = resources.GetString("txt5.AccessibleDescription")
        Me.txt5.AccessibleName = resources.GetString("txt5.AccessibleName")
        Me.txt5.Anchor = CType(resources.GetObject("txt5.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txt5.AutoSize = CType(resources.GetObject("txt5.AutoSize"), Boolean)
        Me.txt5.BackColor = System.Drawing.Color.White
        Me.txt5.BackgroundImage = CType(resources.GetObject("txt5.BackgroundImage"), System.Drawing.Image)
        Me.txt5.Dock = CType(resources.GetObject("txt5.Dock"), System.Windows.Forms.DockStyle)
        Me.txt5.Enabled = CType(resources.GetObject("txt5.Enabled"), Boolean)
        Me.txt5.Font = CType(resources.GetObject("txt5.Font"), System.Drawing.Font)
        Me.txt5.ImeMode = CType(resources.GetObject("txt5.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txt5.Location = CType(resources.GetObject("txt5.Location"), System.Drawing.Point)
        Me.txt5.MaxLength = CType(resources.GetObject("txt5.MaxLength"), Integer)
        Me.txt5.Multiline = CType(resources.GetObject("txt5.Multiline"), Boolean)
        Me.txt5.Name = "txt5"
        Me.txt5.PasswordChar = CType(resources.GetObject("txt5.PasswordChar"), Char)
        Me.txt5.RightToLeft = CType(resources.GetObject("txt5.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txt5.ScrollBars = CType(resources.GetObject("txt5.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txt5.Size = CType(resources.GetObject("txt5.Size"), System.Drawing.Size)
        Me.txt5.TabIndex = CType(resources.GetObject("txt5.TabIndex"), Integer)
        Me.txt5.TabStop = False
        Me.txt5.Text = resources.GetString("txt5.Text")
        Me.txt5.TextAlign = CType(resources.GetObject("txt5.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txt5.Visible = CType(resources.GetObject("txt5.Visible"), Boolean)
        Me.txt5.WordWrap = CType(resources.GetObject("txt5.WordWrap"), Boolean)
        '
        'txt10
        '
        Me.txt10.AccessibleDescription = resources.GetString("txt10.AccessibleDescription")
        Me.txt10.AccessibleName = resources.GetString("txt10.AccessibleName")
        Me.txt10.Anchor = CType(resources.GetObject("txt10.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txt10.AutoSize = CType(resources.GetObject("txt10.AutoSize"), Boolean)
        Me.txt10.BackColor = System.Drawing.Color.White
        Me.txt10.BackgroundImage = CType(resources.GetObject("txt10.BackgroundImage"), System.Drawing.Image)
        Me.txt10.Dock = CType(resources.GetObject("txt10.Dock"), System.Windows.Forms.DockStyle)
        Me.txt10.Enabled = CType(resources.GetObject("txt10.Enabled"), Boolean)
        Me.txt10.Font = CType(resources.GetObject("txt10.Font"), System.Drawing.Font)
        Me.txt10.ImeMode = CType(resources.GetObject("txt10.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txt10.Location = CType(resources.GetObject("txt10.Location"), System.Drawing.Point)
        Me.txt10.MaxLength = CType(resources.GetObject("txt10.MaxLength"), Integer)
        Me.txt10.Multiline = CType(resources.GetObject("txt10.Multiline"), Boolean)
        Me.txt10.Name = "txt10"
        Me.txt10.PasswordChar = CType(resources.GetObject("txt10.PasswordChar"), Char)
        Me.txt10.RightToLeft = CType(resources.GetObject("txt10.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txt10.ScrollBars = CType(resources.GetObject("txt10.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txt10.Size = CType(resources.GetObject("txt10.Size"), System.Drawing.Size)
        Me.txt10.TabIndex = CType(resources.GetObject("txt10.TabIndex"), Integer)
        Me.txt10.TabStop = False
        Me.txt10.Text = resources.GetString("txt10.Text")
        Me.txt10.TextAlign = CType(resources.GetObject("txt10.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txt10.Visible = CType(resources.GetObject("txt10.Visible"), Boolean)
        Me.txt10.WordWrap = CType(resources.GetObject("txt10.WordWrap"), Boolean)
        '
        'txt20
        '
        Me.txt20.AccessibleDescription = resources.GetString("txt20.AccessibleDescription")
        Me.txt20.AccessibleName = resources.GetString("txt20.AccessibleName")
        Me.txt20.Anchor = CType(resources.GetObject("txt20.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txt20.AutoSize = CType(resources.GetObject("txt20.AutoSize"), Boolean)
        Me.txt20.BackColor = System.Drawing.Color.White
        Me.txt20.BackgroundImage = CType(resources.GetObject("txt20.BackgroundImage"), System.Drawing.Image)
        Me.txt20.Dock = CType(resources.GetObject("txt20.Dock"), System.Windows.Forms.DockStyle)
        Me.txt20.Enabled = CType(resources.GetObject("txt20.Enabled"), Boolean)
        Me.txt20.Font = CType(resources.GetObject("txt20.Font"), System.Drawing.Font)
        Me.txt20.ImeMode = CType(resources.GetObject("txt20.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txt20.Location = CType(resources.GetObject("txt20.Location"), System.Drawing.Point)
        Me.txt20.MaxLength = CType(resources.GetObject("txt20.MaxLength"), Integer)
        Me.txt20.Multiline = CType(resources.GetObject("txt20.Multiline"), Boolean)
        Me.txt20.Name = "txt20"
        Me.txt20.PasswordChar = CType(resources.GetObject("txt20.PasswordChar"), Char)
        Me.txt20.RightToLeft = CType(resources.GetObject("txt20.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txt20.ScrollBars = CType(resources.GetObject("txt20.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txt20.Size = CType(resources.GetObject("txt20.Size"), System.Drawing.Size)
        Me.txt20.TabIndex = CType(resources.GetObject("txt20.TabIndex"), Integer)
        Me.txt20.TabStop = False
        Me.txt20.Text = resources.GetString("txt20.Text")
        Me.txt20.TextAlign = CType(resources.GetObject("txt20.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txt20.Visible = CType(resources.GetObject("txt20.Visible"), Boolean)
        Me.txt20.WordWrap = CType(resources.GetObject("txt20.WordWrap"), Boolean)
        '
        'txt50
        '
        Me.txt50.AccessibleDescription = resources.GetString("txt50.AccessibleDescription")
        Me.txt50.AccessibleName = resources.GetString("txt50.AccessibleName")
        Me.txt50.Anchor = CType(resources.GetObject("txt50.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txt50.AutoSize = CType(resources.GetObject("txt50.AutoSize"), Boolean)
        Me.txt50.BackColor = System.Drawing.Color.White
        Me.txt50.BackgroundImage = CType(resources.GetObject("txt50.BackgroundImage"), System.Drawing.Image)
        Me.txt50.Dock = CType(resources.GetObject("txt50.Dock"), System.Windows.Forms.DockStyle)
        Me.txt50.Enabled = CType(resources.GetObject("txt50.Enabled"), Boolean)
        Me.txt50.Font = CType(resources.GetObject("txt50.Font"), System.Drawing.Font)
        Me.txt50.ImeMode = CType(resources.GetObject("txt50.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txt50.Location = CType(resources.GetObject("txt50.Location"), System.Drawing.Point)
        Me.txt50.MaxLength = CType(resources.GetObject("txt50.MaxLength"), Integer)
        Me.txt50.Multiline = CType(resources.GetObject("txt50.Multiline"), Boolean)
        Me.txt50.Name = "txt50"
        Me.txt50.PasswordChar = CType(resources.GetObject("txt50.PasswordChar"), Char)
        Me.txt50.RightToLeft = CType(resources.GetObject("txt50.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txt50.ScrollBars = CType(resources.GetObject("txt50.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txt50.Size = CType(resources.GetObject("txt50.Size"), System.Drawing.Size)
        Me.txt50.TabIndex = CType(resources.GetObject("txt50.TabIndex"), Integer)
        Me.txt50.TabStop = False
        Me.txt50.Text = resources.GetString("txt50.Text")
        Me.txt50.TextAlign = CType(resources.GetObject("txt50.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txt50.Visible = CType(resources.GetObject("txt50.Visible"), Boolean)
        Me.txt50.WordWrap = CType(resources.GetObject("txt50.WordWrap"), Boolean)
        '
        'txt100
        '
        Me.txt100.AccessibleDescription = resources.GetString("txt100.AccessibleDescription")
        Me.txt100.AccessibleName = resources.GetString("txt100.AccessibleName")
        Me.txt100.Anchor = CType(resources.GetObject("txt100.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txt100.AutoSize = CType(resources.GetObject("txt100.AutoSize"), Boolean)
        Me.txt100.BackColor = System.Drawing.Color.White
        Me.txt100.BackgroundImage = CType(resources.GetObject("txt100.BackgroundImage"), System.Drawing.Image)
        Me.txt100.Dock = CType(resources.GetObject("txt100.Dock"), System.Windows.Forms.DockStyle)
        Me.txt100.Enabled = CType(resources.GetObject("txt100.Enabled"), Boolean)
        Me.txt100.Font = CType(resources.GetObject("txt100.Font"), System.Drawing.Font)
        Me.txt100.ImeMode = CType(resources.GetObject("txt100.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txt100.Location = CType(resources.GetObject("txt100.Location"), System.Drawing.Point)
        Me.txt100.MaxLength = CType(resources.GetObject("txt100.MaxLength"), Integer)
        Me.txt100.Multiline = CType(resources.GetObject("txt100.Multiline"), Boolean)
        Me.txt100.Name = "txt100"
        Me.txt100.PasswordChar = CType(resources.GetObject("txt100.PasswordChar"), Char)
        Me.txt100.RightToLeft = CType(resources.GetObject("txt100.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txt100.ScrollBars = CType(resources.GetObject("txt100.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txt100.Size = CType(resources.GetObject("txt100.Size"), System.Drawing.Size)
        Me.txt100.TabIndex = CType(resources.GetObject("txt100.TabIndex"), Integer)
        Me.txt100.TabStop = False
        Me.txt100.Text = resources.GetString("txt100.Text")
        Me.txt100.TextAlign = CType(resources.GetObject("txt100.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txt100.Visible = CType(resources.GetObject("txt100.Visible"), Boolean)
        Me.txt100.WordWrap = CType(resources.GetObject("txt100.WordWrap"), Boolean)
        '
        'txt200
        '
        Me.txt200.AccessibleDescription = resources.GetString("txt200.AccessibleDescription")
        Me.txt200.AccessibleName = resources.GetString("txt200.AccessibleName")
        Me.txt200.Anchor = CType(resources.GetObject("txt200.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txt200.AutoSize = CType(resources.GetObject("txt200.AutoSize"), Boolean)
        Me.txt200.BackColor = System.Drawing.Color.White
        Me.txt200.BackgroundImage = CType(resources.GetObject("txt200.BackgroundImage"), System.Drawing.Image)
        Me.txt200.Dock = CType(resources.GetObject("txt200.Dock"), System.Windows.Forms.DockStyle)
        Me.txt200.Enabled = CType(resources.GetObject("txt200.Enabled"), Boolean)
        Me.txt200.Font = CType(resources.GetObject("txt200.Font"), System.Drawing.Font)
        Me.txt200.ImeMode = CType(resources.GetObject("txt200.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txt200.Location = CType(resources.GetObject("txt200.Location"), System.Drawing.Point)
        Me.txt200.MaxLength = CType(resources.GetObject("txt200.MaxLength"), Integer)
        Me.txt200.Multiline = CType(resources.GetObject("txt200.Multiline"), Boolean)
        Me.txt200.Name = "txt200"
        Me.txt200.PasswordChar = CType(resources.GetObject("txt200.PasswordChar"), Char)
        Me.txt200.RightToLeft = CType(resources.GetObject("txt200.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txt200.ScrollBars = CType(resources.GetObject("txt200.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txt200.Size = CType(resources.GetObject("txt200.Size"), System.Drawing.Size)
        Me.txt200.TabIndex = CType(resources.GetObject("txt200.TabIndex"), Integer)
        Me.txt200.TabStop = False
        Me.txt200.Text = resources.GetString("txt200.Text")
        Me.txt200.TextAlign = CType(resources.GetObject("txt200.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txt200.Visible = CType(resources.GetObject("txt200.Visible"), Boolean)
        Me.txt200.WordWrap = CType(resources.GetObject("txt200.WordWrap"), Boolean)
        '
        'txt500
        '
        Me.txt500.AccessibleDescription = resources.GetString("txt500.AccessibleDescription")
        Me.txt500.AccessibleName = resources.GetString("txt500.AccessibleName")
        Me.txt500.Anchor = CType(resources.GetObject("txt500.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txt500.AutoSize = CType(resources.GetObject("txt500.AutoSize"), Boolean)
        Me.txt500.BackColor = System.Drawing.Color.White
        Me.txt500.BackgroundImage = CType(resources.GetObject("txt500.BackgroundImage"), System.Drawing.Image)
        Me.txt500.Dock = CType(resources.GetObject("txt500.Dock"), System.Windows.Forms.DockStyle)
        Me.txt500.Enabled = CType(resources.GetObject("txt500.Enabled"), Boolean)
        Me.txt500.Font = CType(resources.GetObject("txt500.Font"), System.Drawing.Font)
        Me.txt500.ImeMode = CType(resources.GetObject("txt500.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txt500.Location = CType(resources.GetObject("txt500.Location"), System.Drawing.Point)
        Me.txt500.MaxLength = CType(resources.GetObject("txt500.MaxLength"), Integer)
        Me.txt500.Multiline = CType(resources.GetObject("txt500.Multiline"), Boolean)
        Me.txt500.Name = "txt500"
        Me.txt500.PasswordChar = CType(resources.GetObject("txt500.PasswordChar"), Char)
        Me.txt500.RightToLeft = CType(resources.GetObject("txt500.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txt500.ScrollBars = CType(resources.GetObject("txt500.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txt500.Size = CType(resources.GetObject("txt500.Size"), System.Drawing.Size)
        Me.txt500.TabIndex = CType(resources.GetObject("txt500.TabIndex"), Integer)
        Me.txt500.TabStop = False
        Me.txt500.Text = resources.GetString("txt500.Text")
        Me.txt500.TextAlign = CType(resources.GetObject("txt500.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txt500.Visible = CType(resources.GetObject("txt500.Visible"), Boolean)
        Me.txt500.WordWrap = CType(resources.GetObject("txt500.WordWrap"), Boolean)
        '
        'txt1000
        '
        Me.txt1000.AccessibleDescription = resources.GetString("txt1000.AccessibleDescription")
        Me.txt1000.AccessibleName = resources.GetString("txt1000.AccessibleName")
        Me.txt1000.Anchor = CType(resources.GetObject("txt1000.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txt1000.AutoSize = CType(resources.GetObject("txt1000.AutoSize"), Boolean)
        Me.txt1000.BackColor = System.Drawing.Color.White
        Me.txt1000.BackgroundImage = CType(resources.GetObject("txt1000.BackgroundImage"), System.Drawing.Image)
        Me.txt1000.Dock = CType(resources.GetObject("txt1000.Dock"), System.Windows.Forms.DockStyle)
        Me.txt1000.Enabled = CType(resources.GetObject("txt1000.Enabled"), Boolean)
        Me.txt1000.Font = CType(resources.GetObject("txt1000.Font"), System.Drawing.Font)
        Me.txt1000.ImeMode = CType(resources.GetObject("txt1000.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txt1000.Location = CType(resources.GetObject("txt1000.Location"), System.Drawing.Point)
        Me.txt1000.MaxLength = CType(resources.GetObject("txt1000.MaxLength"), Integer)
        Me.txt1000.Multiline = CType(resources.GetObject("txt1000.Multiline"), Boolean)
        Me.txt1000.Name = "txt1000"
        Me.txt1000.PasswordChar = CType(resources.GetObject("txt1000.PasswordChar"), Char)
        Me.txt1000.RightToLeft = CType(resources.GetObject("txt1000.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txt1000.ScrollBars = CType(resources.GetObject("txt1000.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txt1000.Size = CType(resources.GetObject("txt1000.Size"), System.Drawing.Size)
        Me.txt1000.TabIndex = CType(resources.GetObject("txt1000.TabIndex"), Integer)
        Me.txt1000.TabStop = False
        Me.txt1000.Text = resources.GetString("txt1000.Text")
        Me.txt1000.TextAlign = CType(resources.GetObject("txt1000.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txt1000.Visible = CType(resources.GetObject("txt1000.Visible"), Boolean)
        Me.txt1000.WordWrap = CType(resources.GetObject("txt1000.WordWrap"), Boolean)
        '
        'txtTot001
        '
        Me.txtTot001.AccessibleDescription = resources.GetString("txtTot001.AccessibleDescription")
        Me.txtTot001.AccessibleName = resources.GetString("txtTot001.AccessibleName")
        Me.txtTot001.Anchor = CType(resources.GetObject("txtTot001.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTot001.AutoSize = CType(resources.GetObject("txtTot001.AutoSize"), Boolean)
        Me.txtTot001.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTot001.BackgroundImage = CType(resources.GetObject("txtTot001.BackgroundImage"), System.Drawing.Image)
        Me.txtTot001.Dock = CType(resources.GetObject("txtTot001.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTot001.Enabled = CType(resources.GetObject("txtTot001.Enabled"), Boolean)
        Me.txtTot001.Font = CType(resources.GetObject("txtTot001.Font"), System.Drawing.Font)
        Me.txtTot001.ImeMode = CType(resources.GetObject("txtTot001.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTot001.Location = CType(resources.GetObject("txtTot001.Location"), System.Drawing.Point)
        Me.txtTot001.MaxLength = CType(resources.GetObject("txtTot001.MaxLength"), Integer)
        Me.txtTot001.Multiline = CType(resources.GetObject("txtTot001.Multiline"), Boolean)
        Me.txtTot001.Name = "txtTot001"
        Me.txtTot001.PasswordChar = CType(resources.GetObject("txtTot001.PasswordChar"), Char)
        Me.txtTot001.ReadOnly = True
        Me.txtTot001.RightToLeft = CType(resources.GetObject("txtTot001.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTot001.ScrollBars = CType(resources.GetObject("txtTot001.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTot001.Size = CType(resources.GetObject("txtTot001.Size"), System.Drawing.Size)
        Me.txtTot001.TabIndex = CType(resources.GetObject("txtTot001.TabIndex"), Integer)
        Me.txtTot001.TabStop = False
        Me.txtTot001.Text = resources.GetString("txtTot001.Text")
        Me.txtTot001.TextAlign = CType(resources.GetObject("txtTot001.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTot001.Visible = CType(resources.GetObject("txtTot001.Visible"), Boolean)
        Me.txtTot001.WordWrap = CType(resources.GetObject("txtTot001.WordWrap"), Boolean)
        '
        'txtTot005
        '
        Me.txtTot005.AccessibleDescription = resources.GetString("txtTot005.AccessibleDescription")
        Me.txtTot005.AccessibleName = resources.GetString("txtTot005.AccessibleName")
        Me.txtTot005.Anchor = CType(resources.GetObject("txtTot005.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTot005.AutoSize = CType(resources.GetObject("txtTot005.AutoSize"), Boolean)
        Me.txtTot005.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTot005.BackgroundImage = CType(resources.GetObject("txtTot005.BackgroundImage"), System.Drawing.Image)
        Me.txtTot005.Dock = CType(resources.GetObject("txtTot005.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTot005.Enabled = CType(resources.GetObject("txtTot005.Enabled"), Boolean)
        Me.txtTot005.Font = CType(resources.GetObject("txtTot005.Font"), System.Drawing.Font)
        Me.txtTot005.ImeMode = CType(resources.GetObject("txtTot005.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTot005.Location = CType(resources.GetObject("txtTot005.Location"), System.Drawing.Point)
        Me.txtTot005.MaxLength = CType(resources.GetObject("txtTot005.MaxLength"), Integer)
        Me.txtTot005.Multiline = CType(resources.GetObject("txtTot005.Multiline"), Boolean)
        Me.txtTot005.Name = "txtTot005"
        Me.txtTot005.PasswordChar = CType(resources.GetObject("txtTot005.PasswordChar"), Char)
        Me.txtTot005.ReadOnly = True
        Me.txtTot005.RightToLeft = CType(resources.GetObject("txtTot005.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTot005.ScrollBars = CType(resources.GetObject("txtTot005.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTot005.Size = CType(resources.GetObject("txtTot005.Size"), System.Drawing.Size)
        Me.txtTot005.TabIndex = CType(resources.GetObject("txtTot005.TabIndex"), Integer)
        Me.txtTot005.TabStop = False
        Me.txtTot005.Text = resources.GetString("txtTot005.Text")
        Me.txtTot005.TextAlign = CType(resources.GetObject("txtTot005.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTot005.Visible = CType(resources.GetObject("txtTot005.Visible"), Boolean)
        Me.txtTot005.WordWrap = CType(resources.GetObject("txtTot005.WordWrap"), Boolean)
        '
        'txtTot010
        '
        Me.txtTot010.AccessibleDescription = resources.GetString("txtTot010.AccessibleDescription")
        Me.txtTot010.AccessibleName = resources.GetString("txtTot010.AccessibleName")
        Me.txtTot010.Anchor = CType(resources.GetObject("txtTot010.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTot010.AutoSize = CType(resources.GetObject("txtTot010.AutoSize"), Boolean)
        Me.txtTot010.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTot010.BackgroundImage = CType(resources.GetObject("txtTot010.BackgroundImage"), System.Drawing.Image)
        Me.txtTot010.Dock = CType(resources.GetObject("txtTot010.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTot010.Enabled = CType(resources.GetObject("txtTot010.Enabled"), Boolean)
        Me.txtTot010.Font = CType(resources.GetObject("txtTot010.Font"), System.Drawing.Font)
        Me.txtTot010.ImeMode = CType(resources.GetObject("txtTot010.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTot010.Location = CType(resources.GetObject("txtTot010.Location"), System.Drawing.Point)
        Me.txtTot010.MaxLength = CType(resources.GetObject("txtTot010.MaxLength"), Integer)
        Me.txtTot010.Multiline = CType(resources.GetObject("txtTot010.Multiline"), Boolean)
        Me.txtTot010.Name = "txtTot010"
        Me.txtTot010.PasswordChar = CType(resources.GetObject("txtTot010.PasswordChar"), Char)
        Me.txtTot010.ReadOnly = True
        Me.txtTot010.RightToLeft = CType(resources.GetObject("txtTot010.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTot010.ScrollBars = CType(resources.GetObject("txtTot010.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTot010.Size = CType(resources.GetObject("txtTot010.Size"), System.Drawing.Size)
        Me.txtTot010.TabIndex = CType(resources.GetObject("txtTot010.TabIndex"), Integer)
        Me.txtTot010.TabStop = False
        Me.txtTot010.Text = resources.GetString("txtTot010.Text")
        Me.txtTot010.TextAlign = CType(resources.GetObject("txtTot010.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTot010.Visible = CType(resources.GetObject("txtTot010.Visible"), Boolean)
        Me.txtTot010.WordWrap = CType(resources.GetObject("txtTot010.WordWrap"), Boolean)
        '
        'txtTot025
        '
        Me.txtTot025.AccessibleDescription = resources.GetString("txtTot025.AccessibleDescription")
        Me.txtTot025.AccessibleName = resources.GetString("txtTot025.AccessibleName")
        Me.txtTot025.Anchor = CType(resources.GetObject("txtTot025.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTot025.AutoSize = CType(resources.GetObject("txtTot025.AutoSize"), Boolean)
        Me.txtTot025.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTot025.BackgroundImage = CType(resources.GetObject("txtTot025.BackgroundImage"), System.Drawing.Image)
        Me.txtTot025.Dock = CType(resources.GetObject("txtTot025.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTot025.Enabled = CType(resources.GetObject("txtTot025.Enabled"), Boolean)
        Me.txtTot025.Font = CType(resources.GetObject("txtTot025.Font"), System.Drawing.Font)
        Me.txtTot025.ImeMode = CType(resources.GetObject("txtTot025.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTot025.Location = CType(resources.GetObject("txtTot025.Location"), System.Drawing.Point)
        Me.txtTot025.MaxLength = CType(resources.GetObject("txtTot025.MaxLength"), Integer)
        Me.txtTot025.Multiline = CType(resources.GetObject("txtTot025.Multiline"), Boolean)
        Me.txtTot025.Name = "txtTot025"
        Me.txtTot025.PasswordChar = CType(resources.GetObject("txtTot025.PasswordChar"), Char)
        Me.txtTot025.ReadOnly = True
        Me.txtTot025.RightToLeft = CType(resources.GetObject("txtTot025.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTot025.ScrollBars = CType(resources.GetObject("txtTot025.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTot025.Size = CType(resources.GetObject("txtTot025.Size"), System.Drawing.Size)
        Me.txtTot025.TabIndex = CType(resources.GetObject("txtTot025.TabIndex"), Integer)
        Me.txtTot025.TabStop = False
        Me.txtTot025.Text = resources.GetString("txtTot025.Text")
        Me.txtTot025.TextAlign = CType(resources.GetObject("txtTot025.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTot025.Visible = CType(resources.GetObject("txtTot025.Visible"), Boolean)
        Me.txtTot025.WordWrap = CType(resources.GetObject("txtTot025.WordWrap"), Boolean)
        '
        'Label33
        '
        Me.Label33.AccessibleDescription = resources.GetString("Label33.AccessibleDescription")
        Me.Label33.AccessibleName = resources.GetString("Label33.AccessibleName")
        Me.Label33.Anchor = CType(resources.GetObject("Label33.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label33.AutoSize = CType(resources.GetObject("Label33.AutoSize"), Boolean)
        Me.Label33.Dock = CType(resources.GetObject("Label33.Dock"), System.Windows.Forms.DockStyle)
        Me.Label33.Enabled = CType(resources.GetObject("Label33.Enabled"), Boolean)
        Me.Label33.Font = CType(resources.GetObject("Label33.Font"), System.Drawing.Font)
        Me.Label33.ForeColor = System.Drawing.SystemColors.Info
        Me.Label33.Image = CType(resources.GetObject("Label33.Image"), System.Drawing.Image)
        Me.Label33.ImageAlign = CType(resources.GetObject("Label33.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label33.ImageIndex = CType(resources.GetObject("Label33.ImageIndex"), Integer)
        Me.Label33.ImeMode = CType(resources.GetObject("Label33.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label33.Location = CType(resources.GetObject("Label33.Location"), System.Drawing.Point)
        Me.Label33.Name = "Label33"
        Me.Label33.RightToLeft = CType(resources.GetObject("Label33.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label33.Size = CType(resources.GetObject("Label33.Size"), System.Drawing.Size)
        Me.Label33.TabIndex = CType(resources.GetObject("Label33.TabIndex"), Integer)
        Me.Label33.Text = resources.GetString("Label33.Text")
        Me.Label33.TextAlign = CType(resources.GetObject("Label33.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label33.Visible = CType(resources.GetObject("Label33.Visible"), Boolean)
        '
        'Label32
        '
        Me.Label32.AccessibleDescription = resources.GetString("Label32.AccessibleDescription")
        Me.Label32.AccessibleName = resources.GetString("Label32.AccessibleName")
        Me.Label32.Anchor = CType(resources.GetObject("Label32.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label32.AutoSize = CType(resources.GetObject("Label32.AutoSize"), Boolean)
        Me.Label32.Dock = CType(resources.GetObject("Label32.Dock"), System.Windows.Forms.DockStyle)
        Me.Label32.Enabled = CType(resources.GetObject("Label32.Enabled"), Boolean)
        Me.Label32.Font = CType(resources.GetObject("Label32.Font"), System.Drawing.Font)
        Me.Label32.ForeColor = System.Drawing.SystemColors.Info
        Me.Label32.Image = CType(resources.GetObject("Label32.Image"), System.Drawing.Image)
        Me.Label32.ImageAlign = CType(resources.GetObject("Label32.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label32.ImageIndex = CType(resources.GetObject("Label32.ImageIndex"), Integer)
        Me.Label32.ImeMode = CType(resources.GetObject("Label32.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label32.Location = CType(resources.GetObject("Label32.Location"), System.Drawing.Point)
        Me.Label32.Name = "Label32"
        Me.Label32.RightToLeft = CType(resources.GetObject("Label32.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label32.Size = CType(resources.GetObject("Label32.Size"), System.Drawing.Size)
        Me.Label32.TabIndex = CType(resources.GetObject("Label32.TabIndex"), Integer)
        Me.Label32.Text = resources.GetString("Label32.Text")
        Me.Label32.TextAlign = CType(resources.GetObject("Label32.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label32.Visible = CType(resources.GetObject("Label32.Visible"), Boolean)
        '
        'Label31
        '
        Me.Label31.AccessibleDescription = resources.GetString("Label31.AccessibleDescription")
        Me.Label31.AccessibleName = resources.GetString("Label31.AccessibleName")
        Me.Label31.Anchor = CType(resources.GetObject("Label31.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label31.AutoSize = CType(resources.GetObject("Label31.AutoSize"), Boolean)
        Me.Label31.Dock = CType(resources.GetObject("Label31.Dock"), System.Windows.Forms.DockStyle)
        Me.Label31.Enabled = CType(resources.GetObject("Label31.Enabled"), Boolean)
        Me.Label31.Font = CType(resources.GetObject("Label31.Font"), System.Drawing.Font)
        Me.Label31.ForeColor = System.Drawing.SystemColors.Info
        Me.Label31.Image = CType(resources.GetObject("Label31.Image"), System.Drawing.Image)
        Me.Label31.ImageAlign = CType(resources.GetObject("Label31.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label31.ImageIndex = CType(resources.GetObject("Label31.ImageIndex"), Integer)
        Me.Label31.ImeMode = CType(resources.GetObject("Label31.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label31.Location = CType(resources.GetObject("Label31.Location"), System.Drawing.Point)
        Me.Label31.Name = "Label31"
        Me.Label31.RightToLeft = CType(resources.GetObject("Label31.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label31.Size = CType(resources.GetObject("Label31.Size"), System.Drawing.Size)
        Me.Label31.TabIndex = CType(resources.GetObject("Label31.TabIndex"), Integer)
        Me.Label31.Text = resources.GetString("Label31.Text")
        Me.Label31.TextAlign = CType(resources.GetObject("Label31.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label31.Visible = CType(resources.GetObject("Label31.Visible"), Boolean)
        '
        'Label30
        '
        Me.Label30.AccessibleDescription = resources.GetString("Label30.AccessibleDescription")
        Me.Label30.AccessibleName = resources.GetString("Label30.AccessibleName")
        Me.Label30.Anchor = CType(resources.GetObject("Label30.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label30.AutoSize = CType(resources.GetObject("Label30.AutoSize"), Boolean)
        Me.Label30.Dock = CType(resources.GetObject("Label30.Dock"), System.Windows.Forms.DockStyle)
        Me.Label30.Enabled = CType(resources.GetObject("Label30.Enabled"), Boolean)
        Me.Label30.Font = CType(resources.GetObject("Label30.Font"), System.Drawing.Font)
        Me.Label30.ForeColor = System.Drawing.SystemColors.Info
        Me.Label30.Image = CType(resources.GetObject("Label30.Image"), System.Drawing.Image)
        Me.Label30.ImageAlign = CType(resources.GetObject("Label30.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label30.ImageIndex = CType(resources.GetObject("Label30.ImageIndex"), Integer)
        Me.Label30.ImeMode = CType(resources.GetObject("Label30.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label30.Location = CType(resources.GetObject("Label30.Location"), System.Drawing.Point)
        Me.Label30.Name = "Label30"
        Me.Label30.RightToLeft = CType(resources.GetObject("Label30.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label30.Size = CType(resources.GetObject("Label30.Size"), System.Drawing.Size)
        Me.Label30.TabIndex = CType(resources.GetObject("Label30.TabIndex"), Integer)
        Me.Label30.Text = resources.GetString("Label30.Text")
        Me.Label30.TextAlign = CType(resources.GetObject("Label30.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label30.Visible = CType(resources.GetObject("Label30.Visible"), Boolean)
        '
        'Label29
        '
        Me.Label29.AccessibleDescription = resources.GetString("Label29.AccessibleDescription")
        Me.Label29.AccessibleName = resources.GetString("Label29.AccessibleName")
        Me.Label29.Anchor = CType(resources.GetObject("Label29.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label29.AutoSize = CType(resources.GetObject("Label29.AutoSize"), Boolean)
        Me.Label29.Dock = CType(resources.GetObject("Label29.Dock"), System.Windows.Forms.DockStyle)
        Me.Label29.Enabled = CType(resources.GetObject("Label29.Enabled"), Boolean)
        Me.Label29.Font = CType(resources.GetObject("Label29.Font"), System.Drawing.Font)
        Me.Label29.ForeColor = System.Drawing.SystemColors.Info
        Me.Label29.Image = CType(resources.GetObject("Label29.Image"), System.Drawing.Image)
        Me.Label29.ImageAlign = CType(resources.GetObject("Label29.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label29.ImageIndex = CType(resources.GetObject("Label29.ImageIndex"), Integer)
        Me.Label29.ImeMode = CType(resources.GetObject("Label29.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label29.Location = CType(resources.GetObject("Label29.Location"), System.Drawing.Point)
        Me.Label29.Name = "Label29"
        Me.Label29.RightToLeft = CType(resources.GetObject("Label29.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label29.Size = CType(resources.GetObject("Label29.Size"), System.Drawing.Size)
        Me.Label29.TabIndex = CType(resources.GetObject("Label29.TabIndex"), Integer)
        Me.Label29.Text = resources.GetString("Label29.Text")
        Me.Label29.TextAlign = CType(resources.GetObject("Label29.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label29.Visible = CType(resources.GetObject("Label29.Visible"), Boolean)
        '
        'Label28
        '
        Me.Label28.AccessibleDescription = resources.GetString("Label28.AccessibleDescription")
        Me.Label28.AccessibleName = resources.GetString("Label28.AccessibleName")
        Me.Label28.Anchor = CType(resources.GetObject("Label28.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label28.AutoSize = CType(resources.GetObject("Label28.AutoSize"), Boolean)
        Me.Label28.Dock = CType(resources.GetObject("Label28.Dock"), System.Windows.Forms.DockStyle)
        Me.Label28.Enabled = CType(resources.GetObject("Label28.Enabled"), Boolean)
        Me.Label28.Font = CType(resources.GetObject("Label28.Font"), System.Drawing.Font)
        Me.Label28.ForeColor = System.Drawing.SystemColors.Info
        Me.Label28.Image = CType(resources.GetObject("Label28.Image"), System.Drawing.Image)
        Me.Label28.ImageAlign = CType(resources.GetObject("Label28.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label28.ImageIndex = CType(resources.GetObject("Label28.ImageIndex"), Integer)
        Me.Label28.ImeMode = CType(resources.GetObject("Label28.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label28.Location = CType(resources.GetObject("Label28.Location"), System.Drawing.Point)
        Me.Label28.Name = "Label28"
        Me.Label28.RightToLeft = CType(resources.GetObject("Label28.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label28.Size = CType(resources.GetObject("Label28.Size"), System.Drawing.Size)
        Me.Label28.TabIndex = CType(resources.GetObject("Label28.TabIndex"), Integer)
        Me.Label28.Text = resources.GetString("Label28.Text")
        Me.Label28.TextAlign = CType(resources.GetObject("Label28.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label28.Visible = CType(resources.GetObject("Label28.Visible"), Boolean)
        '
        'Label27
        '
        Me.Label27.AccessibleDescription = resources.GetString("Label27.AccessibleDescription")
        Me.Label27.AccessibleName = resources.GetString("Label27.AccessibleName")
        Me.Label27.Anchor = CType(resources.GetObject("Label27.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label27.AutoSize = CType(resources.GetObject("Label27.AutoSize"), Boolean)
        Me.Label27.Dock = CType(resources.GetObject("Label27.Dock"), System.Windows.Forms.DockStyle)
        Me.Label27.Enabled = CType(resources.GetObject("Label27.Enabled"), Boolean)
        Me.Label27.Font = CType(resources.GetObject("Label27.Font"), System.Drawing.Font)
        Me.Label27.ForeColor = System.Drawing.SystemColors.Info
        Me.Label27.Image = CType(resources.GetObject("Label27.Image"), System.Drawing.Image)
        Me.Label27.ImageAlign = CType(resources.GetObject("Label27.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label27.ImageIndex = CType(resources.GetObject("Label27.ImageIndex"), Integer)
        Me.Label27.ImeMode = CType(resources.GetObject("Label27.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label27.Location = CType(resources.GetObject("Label27.Location"), System.Drawing.Point)
        Me.Label27.Name = "Label27"
        Me.Label27.RightToLeft = CType(resources.GetObject("Label27.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label27.Size = CType(resources.GetObject("Label27.Size"), System.Drawing.Size)
        Me.Label27.TabIndex = CType(resources.GetObject("Label27.TabIndex"), Integer)
        Me.Label27.Text = resources.GetString("Label27.Text")
        Me.Label27.TextAlign = CType(resources.GetObject("Label27.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label27.Visible = CType(resources.GetObject("Label27.Visible"), Boolean)
        '
        'Label26
        '
        Me.Label26.AccessibleDescription = resources.GetString("Label26.AccessibleDescription")
        Me.Label26.AccessibleName = resources.GetString("Label26.AccessibleName")
        Me.Label26.Anchor = CType(resources.GetObject("Label26.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label26.AutoSize = CType(resources.GetObject("Label26.AutoSize"), Boolean)
        Me.Label26.Dock = CType(resources.GetObject("Label26.Dock"), System.Windows.Forms.DockStyle)
        Me.Label26.Enabled = CType(resources.GetObject("Label26.Enabled"), Boolean)
        Me.Label26.Font = CType(resources.GetObject("Label26.Font"), System.Drawing.Font)
        Me.Label26.ForeColor = System.Drawing.SystemColors.Info
        Me.Label26.Image = CType(resources.GetObject("Label26.Image"), System.Drawing.Image)
        Me.Label26.ImageAlign = CType(resources.GetObject("Label26.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label26.ImageIndex = CType(resources.GetObject("Label26.ImageIndex"), Integer)
        Me.Label26.ImeMode = CType(resources.GetObject("Label26.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label26.Location = CType(resources.GetObject("Label26.Location"), System.Drawing.Point)
        Me.Label26.Name = "Label26"
        Me.Label26.RightToLeft = CType(resources.GetObject("Label26.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label26.Size = CType(resources.GetObject("Label26.Size"), System.Drawing.Size)
        Me.Label26.TabIndex = CType(resources.GetObject("Label26.TabIndex"), Integer)
        Me.Label26.Text = resources.GetString("Label26.Text")
        Me.Label26.TextAlign = CType(resources.GetObject("Label26.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label26.Visible = CType(resources.GetObject("Label26.Visible"), Boolean)
        '
        'Label25
        '
        Me.Label25.AccessibleDescription = resources.GetString("Label25.AccessibleDescription")
        Me.Label25.AccessibleName = resources.GetString("Label25.AccessibleName")
        Me.Label25.Anchor = CType(resources.GetObject("Label25.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label25.AutoSize = CType(resources.GetObject("Label25.AutoSize"), Boolean)
        Me.Label25.Dock = CType(resources.GetObject("Label25.Dock"), System.Windows.Forms.DockStyle)
        Me.Label25.Enabled = CType(resources.GetObject("Label25.Enabled"), Boolean)
        Me.Label25.Font = CType(resources.GetObject("Label25.Font"), System.Drawing.Font)
        Me.Label25.ForeColor = System.Drawing.SystemColors.Info
        Me.Label25.Image = CType(resources.GetObject("Label25.Image"), System.Drawing.Image)
        Me.Label25.ImageAlign = CType(resources.GetObject("Label25.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label25.ImageIndex = CType(resources.GetObject("Label25.ImageIndex"), Integer)
        Me.Label25.ImeMode = CType(resources.GetObject("Label25.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label25.Location = CType(resources.GetObject("Label25.Location"), System.Drawing.Point)
        Me.Label25.Name = "Label25"
        Me.Label25.RightToLeft = CType(resources.GetObject("Label25.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label25.Size = CType(resources.GetObject("Label25.Size"), System.Drawing.Size)
        Me.Label25.TabIndex = CType(resources.GetObject("Label25.TabIndex"), Integer)
        Me.Label25.Text = resources.GetString("Label25.Text")
        Me.Label25.TextAlign = CType(resources.GetObject("Label25.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label25.Visible = CType(resources.GetObject("Label25.Visible"), Boolean)
        '
        'Label24
        '
        Me.Label24.AccessibleDescription = resources.GetString("Label24.AccessibleDescription")
        Me.Label24.AccessibleName = resources.GetString("Label24.AccessibleName")
        Me.Label24.Anchor = CType(resources.GetObject("Label24.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label24.AutoSize = CType(resources.GetObject("Label24.AutoSize"), Boolean)
        Me.Label24.Dock = CType(resources.GetObject("Label24.Dock"), System.Windows.Forms.DockStyle)
        Me.Label24.Enabled = CType(resources.GetObject("Label24.Enabled"), Boolean)
        Me.Label24.Font = CType(resources.GetObject("Label24.Font"), System.Drawing.Font)
        Me.Label24.ForeColor = System.Drawing.SystemColors.Info
        Me.Label24.Image = CType(resources.GetObject("Label24.Image"), System.Drawing.Image)
        Me.Label24.ImageAlign = CType(resources.GetObject("Label24.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label24.ImageIndex = CType(resources.GetObject("Label24.ImageIndex"), Integer)
        Me.Label24.ImeMode = CType(resources.GetObject("Label24.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label24.Location = CType(resources.GetObject("Label24.Location"), System.Drawing.Point)
        Me.Label24.Name = "Label24"
        Me.Label24.RightToLeft = CType(resources.GetObject("Label24.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label24.Size = CType(resources.GetObject("Label24.Size"), System.Drawing.Size)
        Me.Label24.TabIndex = CType(resources.GetObject("Label24.TabIndex"), Integer)
        Me.Label24.Text = resources.GetString("Label24.Text")
        Me.Label24.TextAlign = CType(resources.GetObject("Label24.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label24.Visible = CType(resources.GetObject("Label24.Visible"), Boolean)
        '
        'Label23
        '
        Me.Label23.AccessibleDescription = resources.GetString("Label23.AccessibleDescription")
        Me.Label23.AccessibleName = resources.GetString("Label23.AccessibleName")
        Me.Label23.Anchor = CType(resources.GetObject("Label23.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label23.AutoSize = CType(resources.GetObject("Label23.AutoSize"), Boolean)
        Me.Label23.BackColor = System.Drawing.Color.LightSlateGray
        Me.Label23.Dock = CType(resources.GetObject("Label23.Dock"), System.Windows.Forms.DockStyle)
        Me.Label23.Enabled = CType(resources.GetObject("Label23.Enabled"), Boolean)
        Me.Label23.Font = CType(resources.GetObject("Label23.Font"), System.Drawing.Font)
        Me.Label23.ForeColor = System.Drawing.SystemColors.Info
        Me.Label23.Image = CType(resources.GetObject("Label23.Image"), System.Drawing.Image)
        Me.Label23.ImageAlign = CType(resources.GetObject("Label23.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label23.ImageIndex = CType(resources.GetObject("Label23.ImageIndex"), Integer)
        Me.Label23.ImeMode = CType(resources.GetObject("Label23.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label23.Location = CType(resources.GetObject("Label23.Location"), System.Drawing.Point)
        Me.Label23.Name = "Label23"
        Me.Label23.RightToLeft = CType(resources.GetObject("Label23.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label23.Size = CType(resources.GetObject("Label23.Size"), System.Drawing.Size)
        Me.Label23.TabIndex = CType(resources.GetObject("Label23.TabIndex"), Integer)
        Me.Label23.Text = resources.GetString("Label23.Text")
        Me.Label23.TextAlign = CType(resources.GetObject("Label23.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label23.Visible = CType(resources.GetObject("Label23.Visible"), Boolean)
        '
        'txtTot1
        '
        Me.txtTot1.AccessibleDescription = resources.GetString("txtTot1.AccessibleDescription")
        Me.txtTot1.AccessibleName = resources.GetString("txtTot1.AccessibleName")
        Me.txtTot1.Anchor = CType(resources.GetObject("txtTot1.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTot1.AutoSize = CType(resources.GetObject("txtTot1.AutoSize"), Boolean)
        Me.txtTot1.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTot1.BackgroundImage = CType(resources.GetObject("txtTot1.BackgroundImage"), System.Drawing.Image)
        Me.txtTot1.Dock = CType(resources.GetObject("txtTot1.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTot1.Enabled = CType(resources.GetObject("txtTot1.Enabled"), Boolean)
        Me.txtTot1.Font = CType(resources.GetObject("txtTot1.Font"), System.Drawing.Font)
        Me.txtTot1.ImeMode = CType(resources.GetObject("txtTot1.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTot1.Location = CType(resources.GetObject("txtTot1.Location"), System.Drawing.Point)
        Me.txtTot1.MaxLength = CType(resources.GetObject("txtTot1.MaxLength"), Integer)
        Me.txtTot1.Multiline = CType(resources.GetObject("txtTot1.Multiline"), Boolean)
        Me.txtTot1.Name = "txtTot1"
        Me.txtTot1.PasswordChar = CType(resources.GetObject("txtTot1.PasswordChar"), Char)
        Me.txtTot1.ReadOnly = True
        Me.txtTot1.RightToLeft = CType(resources.GetObject("txtTot1.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTot1.ScrollBars = CType(resources.GetObject("txtTot1.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTot1.Size = CType(resources.GetObject("txtTot1.Size"), System.Drawing.Size)
        Me.txtTot1.TabIndex = CType(resources.GetObject("txtTot1.TabIndex"), Integer)
        Me.txtTot1.TabStop = False
        Me.txtTot1.Text = resources.GetString("txtTot1.Text")
        Me.txtTot1.TextAlign = CType(resources.GetObject("txtTot1.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTot1.Visible = CType(resources.GetObject("txtTot1.Visible"), Boolean)
        Me.txtTot1.WordWrap = CType(resources.GetObject("txtTot1.WordWrap"), Boolean)
        '
        'txtTot5
        '
        Me.txtTot5.AccessibleDescription = resources.GetString("txtTot5.AccessibleDescription")
        Me.txtTot5.AccessibleName = resources.GetString("txtTot5.AccessibleName")
        Me.txtTot5.Anchor = CType(resources.GetObject("txtTot5.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTot5.AutoSize = CType(resources.GetObject("txtTot5.AutoSize"), Boolean)
        Me.txtTot5.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTot5.BackgroundImage = CType(resources.GetObject("txtTot5.BackgroundImage"), System.Drawing.Image)
        Me.txtTot5.Dock = CType(resources.GetObject("txtTot5.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTot5.Enabled = CType(resources.GetObject("txtTot5.Enabled"), Boolean)
        Me.txtTot5.Font = CType(resources.GetObject("txtTot5.Font"), System.Drawing.Font)
        Me.txtTot5.ImeMode = CType(resources.GetObject("txtTot5.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTot5.Location = CType(resources.GetObject("txtTot5.Location"), System.Drawing.Point)
        Me.txtTot5.MaxLength = CType(resources.GetObject("txtTot5.MaxLength"), Integer)
        Me.txtTot5.Multiline = CType(resources.GetObject("txtTot5.Multiline"), Boolean)
        Me.txtTot5.Name = "txtTot5"
        Me.txtTot5.PasswordChar = CType(resources.GetObject("txtTot5.PasswordChar"), Char)
        Me.txtTot5.ReadOnly = True
        Me.txtTot5.RightToLeft = CType(resources.GetObject("txtTot5.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTot5.ScrollBars = CType(resources.GetObject("txtTot5.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTot5.Size = CType(resources.GetObject("txtTot5.Size"), System.Drawing.Size)
        Me.txtTot5.TabIndex = CType(resources.GetObject("txtTot5.TabIndex"), Integer)
        Me.txtTot5.TabStop = False
        Me.txtTot5.Text = resources.GetString("txtTot5.Text")
        Me.txtTot5.TextAlign = CType(resources.GetObject("txtTot5.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTot5.Visible = CType(resources.GetObject("txtTot5.Visible"), Boolean)
        Me.txtTot5.WordWrap = CType(resources.GetObject("txtTot5.WordWrap"), Boolean)
        '
        'txtTot10
        '
        Me.txtTot10.AccessibleDescription = resources.GetString("txtTot10.AccessibleDescription")
        Me.txtTot10.AccessibleName = resources.GetString("txtTot10.AccessibleName")
        Me.txtTot10.Anchor = CType(resources.GetObject("txtTot10.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTot10.AutoSize = CType(resources.GetObject("txtTot10.AutoSize"), Boolean)
        Me.txtTot10.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTot10.BackgroundImage = CType(resources.GetObject("txtTot10.BackgroundImage"), System.Drawing.Image)
        Me.txtTot10.Dock = CType(resources.GetObject("txtTot10.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTot10.Enabled = CType(resources.GetObject("txtTot10.Enabled"), Boolean)
        Me.txtTot10.Font = CType(resources.GetObject("txtTot10.Font"), System.Drawing.Font)
        Me.txtTot10.ImeMode = CType(resources.GetObject("txtTot10.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTot10.Location = CType(resources.GetObject("txtTot10.Location"), System.Drawing.Point)
        Me.txtTot10.MaxLength = CType(resources.GetObject("txtTot10.MaxLength"), Integer)
        Me.txtTot10.Multiline = CType(resources.GetObject("txtTot10.Multiline"), Boolean)
        Me.txtTot10.Name = "txtTot10"
        Me.txtTot10.PasswordChar = CType(resources.GetObject("txtTot10.PasswordChar"), Char)
        Me.txtTot10.ReadOnly = True
        Me.txtTot10.RightToLeft = CType(resources.GetObject("txtTot10.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTot10.ScrollBars = CType(resources.GetObject("txtTot10.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTot10.Size = CType(resources.GetObject("txtTot10.Size"), System.Drawing.Size)
        Me.txtTot10.TabIndex = CType(resources.GetObject("txtTot10.TabIndex"), Integer)
        Me.txtTot10.TabStop = False
        Me.txtTot10.Text = resources.GetString("txtTot10.Text")
        Me.txtTot10.TextAlign = CType(resources.GetObject("txtTot10.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTot10.Visible = CType(resources.GetObject("txtTot10.Visible"), Boolean)
        Me.txtTot10.WordWrap = CType(resources.GetObject("txtTot10.WordWrap"), Boolean)
        '
        'txtTot20
        '
        Me.txtTot20.AccessibleDescription = resources.GetString("txtTot20.AccessibleDescription")
        Me.txtTot20.AccessibleName = resources.GetString("txtTot20.AccessibleName")
        Me.txtTot20.Anchor = CType(resources.GetObject("txtTot20.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTot20.AutoSize = CType(resources.GetObject("txtTot20.AutoSize"), Boolean)
        Me.txtTot20.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTot20.BackgroundImage = CType(resources.GetObject("txtTot20.BackgroundImage"), System.Drawing.Image)
        Me.txtTot20.Dock = CType(resources.GetObject("txtTot20.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTot20.Enabled = CType(resources.GetObject("txtTot20.Enabled"), Boolean)
        Me.txtTot20.Font = CType(resources.GetObject("txtTot20.Font"), System.Drawing.Font)
        Me.txtTot20.ImeMode = CType(resources.GetObject("txtTot20.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTot20.Location = CType(resources.GetObject("txtTot20.Location"), System.Drawing.Point)
        Me.txtTot20.MaxLength = CType(resources.GetObject("txtTot20.MaxLength"), Integer)
        Me.txtTot20.Multiline = CType(resources.GetObject("txtTot20.Multiline"), Boolean)
        Me.txtTot20.Name = "txtTot20"
        Me.txtTot20.PasswordChar = CType(resources.GetObject("txtTot20.PasswordChar"), Char)
        Me.txtTot20.ReadOnly = True
        Me.txtTot20.RightToLeft = CType(resources.GetObject("txtTot20.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTot20.ScrollBars = CType(resources.GetObject("txtTot20.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTot20.Size = CType(resources.GetObject("txtTot20.Size"), System.Drawing.Size)
        Me.txtTot20.TabIndex = CType(resources.GetObject("txtTot20.TabIndex"), Integer)
        Me.txtTot20.TabStop = False
        Me.txtTot20.Text = resources.GetString("txtTot20.Text")
        Me.txtTot20.TextAlign = CType(resources.GetObject("txtTot20.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTot20.Visible = CType(resources.GetObject("txtTot20.Visible"), Boolean)
        Me.txtTot20.WordWrap = CType(resources.GetObject("txtTot20.WordWrap"), Boolean)
        '
        'txtTot50
        '
        Me.txtTot50.AccessibleDescription = resources.GetString("txtTot50.AccessibleDescription")
        Me.txtTot50.AccessibleName = resources.GetString("txtTot50.AccessibleName")
        Me.txtTot50.Anchor = CType(resources.GetObject("txtTot50.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTot50.AutoSize = CType(resources.GetObject("txtTot50.AutoSize"), Boolean)
        Me.txtTot50.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTot50.BackgroundImage = CType(resources.GetObject("txtTot50.BackgroundImage"), System.Drawing.Image)
        Me.txtTot50.Dock = CType(resources.GetObject("txtTot50.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTot50.Enabled = CType(resources.GetObject("txtTot50.Enabled"), Boolean)
        Me.txtTot50.Font = CType(resources.GetObject("txtTot50.Font"), System.Drawing.Font)
        Me.txtTot50.ImeMode = CType(resources.GetObject("txtTot50.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTot50.Location = CType(resources.GetObject("txtTot50.Location"), System.Drawing.Point)
        Me.txtTot50.MaxLength = CType(resources.GetObject("txtTot50.MaxLength"), Integer)
        Me.txtTot50.Multiline = CType(resources.GetObject("txtTot50.Multiline"), Boolean)
        Me.txtTot50.Name = "txtTot50"
        Me.txtTot50.PasswordChar = CType(resources.GetObject("txtTot50.PasswordChar"), Char)
        Me.txtTot50.ReadOnly = True
        Me.txtTot50.RightToLeft = CType(resources.GetObject("txtTot50.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTot50.ScrollBars = CType(resources.GetObject("txtTot50.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTot50.Size = CType(resources.GetObject("txtTot50.Size"), System.Drawing.Size)
        Me.txtTot50.TabIndex = CType(resources.GetObject("txtTot50.TabIndex"), Integer)
        Me.txtTot50.TabStop = False
        Me.txtTot50.Text = resources.GetString("txtTot50.Text")
        Me.txtTot50.TextAlign = CType(resources.GetObject("txtTot50.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTot50.Visible = CType(resources.GetObject("txtTot50.Visible"), Boolean)
        Me.txtTot50.WordWrap = CType(resources.GetObject("txtTot50.WordWrap"), Boolean)
        '
        'txtTot100
        '
        Me.txtTot100.AccessibleDescription = resources.GetString("txtTot100.AccessibleDescription")
        Me.txtTot100.AccessibleName = resources.GetString("txtTot100.AccessibleName")
        Me.txtTot100.Anchor = CType(resources.GetObject("txtTot100.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTot100.AutoSize = CType(resources.GetObject("txtTot100.AutoSize"), Boolean)
        Me.txtTot100.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTot100.BackgroundImage = CType(resources.GetObject("txtTot100.BackgroundImage"), System.Drawing.Image)
        Me.txtTot100.Dock = CType(resources.GetObject("txtTot100.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTot100.Enabled = CType(resources.GetObject("txtTot100.Enabled"), Boolean)
        Me.txtTot100.Font = CType(resources.GetObject("txtTot100.Font"), System.Drawing.Font)
        Me.txtTot100.ImeMode = CType(resources.GetObject("txtTot100.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTot100.Location = CType(resources.GetObject("txtTot100.Location"), System.Drawing.Point)
        Me.txtTot100.MaxLength = CType(resources.GetObject("txtTot100.MaxLength"), Integer)
        Me.txtTot100.Multiline = CType(resources.GetObject("txtTot100.Multiline"), Boolean)
        Me.txtTot100.Name = "txtTot100"
        Me.txtTot100.PasswordChar = CType(resources.GetObject("txtTot100.PasswordChar"), Char)
        Me.txtTot100.ReadOnly = True
        Me.txtTot100.RightToLeft = CType(resources.GetObject("txtTot100.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTot100.ScrollBars = CType(resources.GetObject("txtTot100.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTot100.Size = CType(resources.GetObject("txtTot100.Size"), System.Drawing.Size)
        Me.txtTot100.TabIndex = CType(resources.GetObject("txtTot100.TabIndex"), Integer)
        Me.txtTot100.TabStop = False
        Me.txtTot100.Text = resources.GetString("txtTot100.Text")
        Me.txtTot100.TextAlign = CType(resources.GetObject("txtTot100.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTot100.Visible = CType(resources.GetObject("txtTot100.Visible"), Boolean)
        Me.txtTot100.WordWrap = CType(resources.GetObject("txtTot100.WordWrap"), Boolean)
        '
        'txtTot200
        '
        Me.txtTot200.AccessibleDescription = resources.GetString("txtTot200.AccessibleDescription")
        Me.txtTot200.AccessibleName = resources.GetString("txtTot200.AccessibleName")
        Me.txtTot200.Anchor = CType(resources.GetObject("txtTot200.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTot200.AutoSize = CType(resources.GetObject("txtTot200.AutoSize"), Boolean)
        Me.txtTot200.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTot200.BackgroundImage = CType(resources.GetObject("txtTot200.BackgroundImage"), System.Drawing.Image)
        Me.txtTot200.Dock = CType(resources.GetObject("txtTot200.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTot200.Enabled = CType(resources.GetObject("txtTot200.Enabled"), Boolean)
        Me.txtTot200.Font = CType(resources.GetObject("txtTot200.Font"), System.Drawing.Font)
        Me.txtTot200.ImeMode = CType(resources.GetObject("txtTot200.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTot200.Location = CType(resources.GetObject("txtTot200.Location"), System.Drawing.Point)
        Me.txtTot200.MaxLength = CType(resources.GetObject("txtTot200.MaxLength"), Integer)
        Me.txtTot200.Multiline = CType(resources.GetObject("txtTot200.Multiline"), Boolean)
        Me.txtTot200.Name = "txtTot200"
        Me.txtTot200.PasswordChar = CType(resources.GetObject("txtTot200.PasswordChar"), Char)
        Me.txtTot200.ReadOnly = True
        Me.txtTot200.RightToLeft = CType(resources.GetObject("txtTot200.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTot200.ScrollBars = CType(resources.GetObject("txtTot200.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTot200.Size = CType(resources.GetObject("txtTot200.Size"), System.Drawing.Size)
        Me.txtTot200.TabIndex = CType(resources.GetObject("txtTot200.TabIndex"), Integer)
        Me.txtTot200.TabStop = False
        Me.txtTot200.Text = resources.GetString("txtTot200.Text")
        Me.txtTot200.TextAlign = CType(resources.GetObject("txtTot200.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTot200.Visible = CType(resources.GetObject("txtTot200.Visible"), Boolean)
        Me.txtTot200.WordWrap = CType(resources.GetObject("txtTot200.WordWrap"), Boolean)
        '
        'txtTot500
        '
        Me.txtTot500.AccessibleDescription = resources.GetString("txtTot500.AccessibleDescription")
        Me.txtTot500.AccessibleName = resources.GetString("txtTot500.AccessibleName")
        Me.txtTot500.Anchor = CType(resources.GetObject("txtTot500.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTot500.AutoSize = CType(resources.GetObject("txtTot500.AutoSize"), Boolean)
        Me.txtTot500.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTot500.BackgroundImage = CType(resources.GetObject("txtTot500.BackgroundImage"), System.Drawing.Image)
        Me.txtTot500.Dock = CType(resources.GetObject("txtTot500.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTot500.Enabled = CType(resources.GetObject("txtTot500.Enabled"), Boolean)
        Me.txtTot500.Font = CType(resources.GetObject("txtTot500.Font"), System.Drawing.Font)
        Me.txtTot500.ImeMode = CType(resources.GetObject("txtTot500.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTot500.Location = CType(resources.GetObject("txtTot500.Location"), System.Drawing.Point)
        Me.txtTot500.MaxLength = CType(resources.GetObject("txtTot500.MaxLength"), Integer)
        Me.txtTot500.Multiline = CType(resources.GetObject("txtTot500.Multiline"), Boolean)
        Me.txtTot500.Name = "txtTot500"
        Me.txtTot500.PasswordChar = CType(resources.GetObject("txtTot500.PasswordChar"), Char)
        Me.txtTot500.ReadOnly = True
        Me.txtTot500.RightToLeft = CType(resources.GetObject("txtTot500.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTot500.ScrollBars = CType(resources.GetObject("txtTot500.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTot500.Size = CType(resources.GetObject("txtTot500.Size"), System.Drawing.Size)
        Me.txtTot500.TabIndex = CType(resources.GetObject("txtTot500.TabIndex"), Integer)
        Me.txtTot500.TabStop = False
        Me.txtTot500.Text = resources.GetString("txtTot500.Text")
        Me.txtTot500.TextAlign = CType(resources.GetObject("txtTot500.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTot500.Visible = CType(resources.GetObject("txtTot500.Visible"), Boolean)
        Me.txtTot500.WordWrap = CType(resources.GetObject("txtTot500.WordWrap"), Boolean)
        '
        'Label22
        '
        Me.Label22.AccessibleDescription = resources.GetString("Label22.AccessibleDescription")
        Me.Label22.AccessibleName = resources.GetString("Label22.AccessibleName")
        Me.Label22.Anchor = CType(resources.GetObject("Label22.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label22.AutoSize = CType(resources.GetObject("Label22.AutoSize"), Boolean)
        Me.Label22.Dock = CType(resources.GetObject("Label22.Dock"), System.Windows.Forms.DockStyle)
        Me.Label22.Enabled = CType(resources.GetObject("Label22.Enabled"), Boolean)
        Me.Label22.Font = CType(resources.GetObject("Label22.Font"), System.Drawing.Font)
        Me.Label22.ForeColor = System.Drawing.SystemColors.Info
        Me.Label22.Image = CType(resources.GetObject("Label22.Image"), System.Drawing.Image)
        Me.Label22.ImageAlign = CType(resources.GetObject("Label22.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label22.ImageIndex = CType(resources.GetObject("Label22.ImageIndex"), Integer)
        Me.Label22.ImeMode = CType(resources.GetObject("Label22.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label22.Location = CType(resources.GetObject("Label22.Location"), System.Drawing.Point)
        Me.Label22.Name = "Label22"
        Me.Label22.RightToLeft = CType(resources.GetObject("Label22.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label22.Size = CType(resources.GetObject("Label22.Size"), System.Drawing.Size)
        Me.Label22.TabIndex = CType(resources.GetObject("Label22.TabIndex"), Integer)
        Me.Label22.Text = resources.GetString("Label22.Text")
        Me.Label22.TextAlign = CType(resources.GetObject("Label22.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label22.Visible = CType(resources.GetObject("Label22.Visible"), Boolean)
        '
        'txtTot1000
        '
        Me.txtTot1000.AccessibleDescription = resources.GetString("txtTot1000.AccessibleDescription")
        Me.txtTot1000.AccessibleName = resources.GetString("txtTot1000.AccessibleName")
        Me.txtTot1000.Anchor = CType(resources.GetObject("txtTot1000.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTot1000.AutoSize = CType(resources.GetObject("txtTot1000.AutoSize"), Boolean)
        Me.txtTot1000.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTot1000.BackgroundImage = CType(resources.GetObject("txtTot1000.BackgroundImage"), System.Drawing.Image)
        Me.txtTot1000.Dock = CType(resources.GetObject("txtTot1000.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTot1000.Enabled = CType(resources.GetObject("txtTot1000.Enabled"), Boolean)
        Me.txtTot1000.Font = CType(resources.GetObject("txtTot1000.Font"), System.Drawing.Font)
        Me.txtTot1000.ImeMode = CType(resources.GetObject("txtTot1000.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTot1000.Location = CType(resources.GetObject("txtTot1000.Location"), System.Drawing.Point)
        Me.txtTot1000.MaxLength = CType(resources.GetObject("txtTot1000.MaxLength"), Integer)
        Me.txtTot1000.Multiline = CType(resources.GetObject("txtTot1000.Multiline"), Boolean)
        Me.txtTot1000.Name = "txtTot1000"
        Me.txtTot1000.PasswordChar = CType(resources.GetObject("txtTot1000.PasswordChar"), Char)
        Me.txtTot1000.ReadOnly = True
        Me.txtTot1000.RightToLeft = CType(resources.GetObject("txtTot1000.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTot1000.ScrollBars = CType(resources.GetObject("txtTot1000.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTot1000.Size = CType(resources.GetObject("txtTot1000.Size"), System.Drawing.Size)
        Me.txtTot1000.TabIndex = CType(resources.GetObject("txtTot1000.TabIndex"), Integer)
        Me.txtTot1000.TabStop = False
        Me.txtTot1000.Text = resources.GetString("txtTot1000.Text")
        Me.txtTot1000.TextAlign = CType(resources.GetObject("txtTot1000.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTot1000.Visible = CType(resources.GetObject("txtTot1000.Visible"), Boolean)
        Me.txtTot1000.WordWrap = CType(resources.GetObject("txtTot1000.WordWrap"), Boolean)
        '
        'Label21
        '
        Me.Label21.AccessibleDescription = resources.GetString("Label21.AccessibleDescription")
        Me.Label21.AccessibleName = resources.GetString("Label21.AccessibleName")
        Me.Label21.Anchor = CType(resources.GetObject("Label21.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label21.AutoSize = CType(resources.GetObject("Label21.AutoSize"), Boolean)
        Me.Label21.BackColor = System.Drawing.Color.MidnightBlue
        Me.Label21.Dock = CType(resources.GetObject("Label21.Dock"), System.Windows.Forms.DockStyle)
        Me.Label21.Enabled = CType(resources.GetObject("Label21.Enabled"), Boolean)
        Me.Label21.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label21.Font = CType(resources.GetObject("Label21.Font"), System.Drawing.Font)
        Me.Label21.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label21.Image = CType(resources.GetObject("Label21.Image"), System.Drawing.Image)
        Me.Label21.ImageAlign = CType(resources.GetObject("Label21.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label21.ImageIndex = CType(resources.GetObject("Label21.ImageIndex"), Integer)
        Me.Label21.ImeMode = CType(resources.GetObject("Label21.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label21.Location = CType(resources.GetObject("Label21.Location"), System.Drawing.Point)
        Me.Label21.Name = "Label21"
        Me.Label21.RightToLeft = CType(resources.GetObject("Label21.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label21.Size = CType(resources.GetObject("Label21.Size"), System.Drawing.Size)
        Me.Label21.TabIndex = CType(resources.GetObject("Label21.TabIndex"), Integer)
        Me.Label21.Text = resources.GetString("Label21.Text")
        Me.Label21.TextAlign = CType(resources.GetObject("Label21.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label21.Visible = CType(resources.GetObject("Label21.Visible"), Boolean)
        '
        'Label20
        '
        Me.Label20.AccessibleDescription = resources.GetString("Label20.AccessibleDescription")
        Me.Label20.AccessibleName = resources.GetString("Label20.AccessibleName")
        Me.Label20.Anchor = CType(resources.GetObject("Label20.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label20.AutoSize = CType(resources.GetObject("Label20.AutoSize"), Boolean)
        Me.Label20.Dock = CType(resources.GetObject("Label20.Dock"), System.Windows.Forms.DockStyle)
        Me.Label20.Enabled = CType(resources.GetObject("Label20.Enabled"), Boolean)
        Me.Label20.Font = CType(resources.GetObject("Label20.Font"), System.Drawing.Font)
        Me.Label20.ForeColor = System.Drawing.SystemColors.Info
        Me.Label20.Image = CType(resources.GetObject("Label20.Image"), System.Drawing.Image)
        Me.Label20.ImageAlign = CType(resources.GetObject("Label20.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label20.ImageIndex = CType(resources.GetObject("Label20.ImageIndex"), Integer)
        Me.Label20.ImeMode = CType(resources.GetObject("Label20.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label20.Location = CType(resources.GetObject("Label20.Location"), System.Drawing.Point)
        Me.Label20.Name = "Label20"
        Me.Label20.RightToLeft = CType(resources.GetObject("Label20.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label20.Size = CType(resources.GetObject("Label20.Size"), System.Drawing.Size)
        Me.Label20.TabIndex = CType(resources.GetObject("Label20.TabIndex"), Integer)
        Me.Label20.Text = resources.GetString("Label20.Text")
        Me.Label20.TextAlign = CType(resources.GetObject("Label20.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label20.Visible = CType(resources.GetObject("Label20.Visible"), Boolean)
        '
        'Label10
        '
        Me.Label10.AccessibleDescription = resources.GetString("Label10.AccessibleDescription")
        Me.Label10.AccessibleName = resources.GetString("Label10.AccessibleName")
        Me.Label10.Anchor = CType(resources.GetObject("Label10.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label10.AutoSize = CType(resources.GetObject("Label10.AutoSize"), Boolean)
        Me.Label10.Dock = CType(resources.GetObject("Label10.Dock"), System.Windows.Forms.DockStyle)
        Me.Label10.Enabled = CType(resources.GetObject("Label10.Enabled"), Boolean)
        Me.Label10.Font = CType(resources.GetObject("Label10.Font"), System.Drawing.Font)
        Me.Label10.ForeColor = System.Drawing.SystemColors.Info
        Me.Label10.Image = CType(resources.GetObject("Label10.Image"), System.Drawing.Image)
        Me.Label10.ImageAlign = CType(resources.GetObject("Label10.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label10.ImageIndex = CType(resources.GetObject("Label10.ImageIndex"), Integer)
        Me.Label10.ImeMode = CType(resources.GetObject("Label10.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label10.Location = CType(resources.GetObject("Label10.Location"), System.Drawing.Point)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = CType(resources.GetObject("Label10.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label10.Size = CType(resources.GetObject("Label10.Size"), System.Drawing.Size)
        Me.Label10.TabIndex = CType(resources.GetObject("Label10.TabIndex"), Integer)
        Me.Label10.Text = resources.GetString("Label10.Text")
        Me.Label10.TextAlign = CType(resources.GetObject("Label10.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label10.Visible = CType(resources.GetObject("Label10.Visible"), Boolean)
        '
        'txtAmtLeft
        '
        Me.txtAmtLeft.AccessibleDescription = resources.GetString("txtAmtLeft.AccessibleDescription")
        Me.txtAmtLeft.AccessibleName = resources.GetString("txtAmtLeft.AccessibleName")
        Me.txtAmtLeft.Anchor = CType(resources.GetObject("txtAmtLeft.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtAmtLeft.AutoSize = CType(resources.GetObject("txtAmtLeft.AutoSize"), Boolean)
        Me.txtAmtLeft.BackColor = System.Drawing.Color.LightSteelBlue
        Me.txtAmtLeft.BackgroundImage = CType(resources.GetObject("txtAmtLeft.BackgroundImage"), System.Drawing.Image)
        Me.txtAmtLeft.Dock = CType(resources.GetObject("txtAmtLeft.Dock"), System.Windows.Forms.DockStyle)
        Me.txtAmtLeft.Enabled = CType(resources.GetObject("txtAmtLeft.Enabled"), Boolean)
        Me.txtAmtLeft.Font = CType(resources.GetObject("txtAmtLeft.Font"), System.Drawing.Font)
        Me.txtAmtLeft.ImeMode = CType(resources.GetObject("txtAmtLeft.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtAmtLeft.Location = CType(resources.GetObject("txtAmtLeft.Location"), System.Drawing.Point)
        Me.txtAmtLeft.MaxLength = CType(resources.GetObject("txtAmtLeft.MaxLength"), Integer)
        Me.txtAmtLeft.Multiline = CType(resources.GetObject("txtAmtLeft.Multiline"), Boolean)
        Me.txtAmtLeft.Name = "txtAmtLeft"
        Me.txtAmtLeft.PasswordChar = CType(resources.GetObject("txtAmtLeft.PasswordChar"), Char)
        Me.txtAmtLeft.ReadOnly = True
        Me.txtAmtLeft.RightToLeft = CType(resources.GetObject("txtAmtLeft.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtAmtLeft.ScrollBars = CType(resources.GetObject("txtAmtLeft.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtAmtLeft.Size = CType(resources.GetObject("txtAmtLeft.Size"), System.Drawing.Size)
        Me.txtAmtLeft.TabIndex = CType(resources.GetObject("txtAmtLeft.TabIndex"), Integer)
        Me.txtAmtLeft.TabStop = False
        Me.txtAmtLeft.Text = resources.GetString("txtAmtLeft.Text")
        Me.txtAmtLeft.TextAlign = CType(resources.GetObject("txtAmtLeft.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtAmtLeft.Visible = CType(resources.GetObject("txtAmtLeft.Visible"), Boolean)
        Me.txtAmtLeft.WordWrap = CType(resources.GetObject("txtAmtLeft.WordWrap"), Boolean)
        '
        'txtTotCash
        '
        Me.txtTotCash.AccessibleDescription = resources.GetString("txtTotCash.AccessibleDescription")
        Me.txtTotCash.AccessibleName = resources.GetString("txtTotCash.AccessibleName")
        Me.txtTotCash.Anchor = CType(resources.GetObject("txtTotCash.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTotCash.AutoSize = CType(resources.GetObject("txtTotCash.AutoSize"), Boolean)
        Me.txtTotCash.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTotCash.BackgroundImage = CType(resources.GetObject("txtTotCash.BackgroundImage"), System.Drawing.Image)
        Me.txtTotCash.Dock = CType(resources.GetObject("txtTotCash.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTotCash.Enabled = CType(resources.GetObject("txtTotCash.Enabled"), Boolean)
        Me.txtTotCash.Font = CType(resources.GetObject("txtTotCash.Font"), System.Drawing.Font)
        Me.txtTotCash.ImeMode = CType(resources.GetObject("txtTotCash.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTotCash.Location = CType(resources.GetObject("txtTotCash.Location"), System.Drawing.Point)
        Me.txtTotCash.MaxLength = CType(resources.GetObject("txtTotCash.MaxLength"), Integer)
        Me.txtTotCash.Multiline = CType(resources.GetObject("txtTotCash.Multiline"), Boolean)
        Me.txtTotCash.Name = "txtTotCash"
        Me.txtTotCash.PasswordChar = CType(resources.GetObject("txtTotCash.PasswordChar"), Char)
        Me.txtTotCash.ReadOnly = True
        Me.txtTotCash.RightToLeft = CType(resources.GetObject("txtTotCash.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTotCash.ScrollBars = CType(resources.GetObject("txtTotCash.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTotCash.Size = CType(resources.GetObject("txtTotCash.Size"), System.Drawing.Size)
        Me.txtTotCash.TabIndex = CType(resources.GetObject("txtTotCash.TabIndex"), Integer)
        Me.txtTotCash.TabStop = False
        Me.txtTotCash.Text = resources.GetString("txtTotCash.Text")
        Me.txtTotCash.TextAlign = CType(resources.GetObject("txtTotCash.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTotCash.Visible = CType(resources.GetObject("txtTotCash.Visible"), Boolean)
        Me.txtTotCash.WordWrap = CType(resources.GetObject("txtTotCash.WordWrap"), Boolean)
        '
        'txtExCash
        '
        Me.txtExCash.AccessibleDescription = resources.GetString("txtExCash.AccessibleDescription")
        Me.txtExCash.AccessibleName = resources.GetString("txtExCash.AccessibleName")
        Me.txtExCash.Anchor = CType(resources.GetObject("txtExCash.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtExCash.AutoSize = CType(resources.GetObject("txtExCash.AutoSize"), Boolean)
        Me.txtExCash.BackColor = System.Drawing.Color.LightSteelBlue
        Me.txtExCash.BackgroundImage = CType(resources.GetObject("txtExCash.BackgroundImage"), System.Drawing.Image)
        Me.txtExCash.Dock = CType(resources.GetObject("txtExCash.Dock"), System.Windows.Forms.DockStyle)
        Me.txtExCash.Enabled = CType(resources.GetObject("txtExCash.Enabled"), Boolean)
        Me.txtExCash.Font = CType(resources.GetObject("txtExCash.Font"), System.Drawing.Font)
        Me.txtExCash.ImeMode = CType(resources.GetObject("txtExCash.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtExCash.Location = CType(resources.GetObject("txtExCash.Location"), System.Drawing.Point)
        Me.txtExCash.MaxLength = CType(resources.GetObject("txtExCash.MaxLength"), Integer)
        Me.txtExCash.Multiline = CType(resources.GetObject("txtExCash.Multiline"), Boolean)
        Me.txtExCash.Name = "txtExCash"
        Me.txtExCash.PasswordChar = CType(resources.GetObject("txtExCash.PasswordChar"), Char)
        Me.txtExCash.ReadOnly = True
        Me.txtExCash.RightToLeft = CType(resources.GetObject("txtExCash.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtExCash.ScrollBars = CType(resources.GetObject("txtExCash.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtExCash.Size = CType(resources.GetObject("txtExCash.Size"), System.Drawing.Size)
        Me.txtExCash.TabIndex = CType(resources.GetObject("txtExCash.TabIndex"), Integer)
        Me.txtExCash.TabStop = False
        Me.txtExCash.Text = resources.GetString("txtExCash.Text")
        Me.txtExCash.TextAlign = CType(resources.GetObject("txtExCash.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtExCash.Visible = CType(resources.GetObject("txtExCash.Visible"), Boolean)
        Me.txtExCash.WordWrap = CType(resources.GetObject("txtExCash.WordWrap"), Boolean)
        '
        'Label3
        '
        Me.Label3.AccessibleDescription = resources.GetString("Label3.AccessibleDescription")
        Me.Label3.AccessibleName = resources.GetString("Label3.AccessibleName")
        Me.Label3.Anchor = CType(resources.GetObject("Label3.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = CType(resources.GetObject("Label3.AutoSize"), Boolean)
        Me.Label3.Dock = CType(resources.GetObject("Label3.Dock"), System.Windows.Forms.DockStyle)
        Me.Label3.Enabled = CType(resources.GetObject("Label3.Enabled"), Boolean)
        Me.Label3.Font = CType(resources.GetObject("Label3.Font"), System.Drawing.Font)
        Me.Label3.ForeColor = System.Drawing.SystemColors.Info
        Me.Label3.Image = CType(resources.GetObject("Label3.Image"), System.Drawing.Image)
        Me.Label3.ImageAlign = CType(resources.GetObject("Label3.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label3.ImageIndex = CType(resources.GetObject("Label3.ImageIndex"), Integer)
        Me.Label3.ImeMode = CType(resources.GetObject("Label3.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label3.Location = CType(resources.GetObject("Label3.Location"), System.Drawing.Point)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = CType(resources.GetObject("Label3.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label3.Size = CType(resources.GetObject("Label3.Size"), System.Drawing.Size)
        Me.Label3.TabIndex = CType(resources.GetObject("Label3.TabIndex"), Integer)
        Me.Label3.Text = resources.GetString("Label3.Text")
        Me.Label3.TextAlign = CType(resources.GetObject("Label3.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label3.Visible = CType(resources.GetObject("Label3.Visible"), Boolean)
        '
        'Label2
        '
        Me.Label2.AccessibleDescription = resources.GetString("Label2.AccessibleDescription")
        Me.Label2.AccessibleName = resources.GetString("Label2.AccessibleName")
        Me.Label2.Anchor = CType(resources.GetObject("Label2.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = CType(resources.GetObject("Label2.AutoSize"), Boolean)
        Me.Label2.Dock = CType(resources.GetObject("Label2.Dock"), System.Windows.Forms.DockStyle)
        Me.Label2.Enabled = CType(resources.GetObject("Label2.Enabled"), Boolean)
        Me.Label2.Font = CType(resources.GetObject("Label2.Font"), System.Drawing.Font)
        Me.Label2.ForeColor = System.Drawing.SystemColors.Info
        Me.Label2.Image = CType(resources.GetObject("Label2.Image"), System.Drawing.Image)
        Me.Label2.ImageAlign = CType(resources.GetObject("Label2.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label2.ImageIndex = CType(resources.GetObject("Label2.ImageIndex"), Integer)
        Me.Label2.ImeMode = CType(resources.GetObject("Label2.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label2.Location = CType(resources.GetObject("Label2.Location"), System.Drawing.Point)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = CType(resources.GetObject("Label2.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label2.Size = CType(resources.GetObject("Label2.Size"), System.Drawing.Size)
        Me.Label2.TabIndex = CType(resources.GetObject("Label2.TabIndex"), Integer)
        Me.Label2.Text = resources.GetString("Label2.Text")
        Me.Label2.TextAlign = CType(resources.GetObject("Label2.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label2.Visible = CType(resources.GetObject("Label2.Visible"), Boolean)
        '
        'Label1
        '
        Me.Label1.AccessibleDescription = resources.GetString("Label1.AccessibleDescription")
        Me.Label1.AccessibleName = resources.GetString("Label1.AccessibleName")
        Me.Label1.Anchor = CType(resources.GetObject("Label1.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = CType(resources.GetObject("Label1.AutoSize"), Boolean)
        Me.Label1.BackColor = System.Drawing.Color.SteelBlue
        Me.Label1.Dock = CType(resources.GetObject("Label1.Dock"), System.Windows.Forms.DockStyle)
        Me.Label1.Enabled = CType(resources.GetObject("Label1.Enabled"), Boolean)
        Me.Label1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label1.Font = CType(resources.GetObject("Label1.Font"), System.Drawing.Font)
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label1.Image = CType(resources.GetObject("Label1.Image"), System.Drawing.Image)
        Me.Label1.ImageAlign = CType(resources.GetObject("Label1.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label1.ImageIndex = CType(resources.GetObject("Label1.ImageIndex"), Integer)
        Me.Label1.ImeMode = CType(resources.GetObject("Label1.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label1.Location = CType(resources.GetObject("Label1.Location"), System.Drawing.Point)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = CType(resources.GetObject("Label1.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label1.Size = CType(resources.GetObject("Label1.Size"), System.Drawing.Size)
        Me.Label1.TabIndex = CType(resources.GetObject("Label1.TabIndex"), Integer)
        Me.Label1.Text = resources.GetString("Label1.Text")
        Me.Label1.TextAlign = CType(resources.GetObject("Label1.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label1.Visible = CType(resources.GetObject("Label1.Visible"), Boolean)
        '
        'gbCheque
        '
        Me.gbCheque.AccessibleDescription = resources.GetString("gbCheque.AccessibleDescription")
        Me.gbCheque.AccessibleName = resources.GetString("gbCheque.AccessibleName")
        Me.gbCheque.Anchor = CType(resources.GetObject("gbCheque.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.gbCheque.BackgroundImage = CType(resources.GetObject("gbCheque.BackgroundImage"), System.Drawing.Image)
        Me.gbCheque.Controls.Add(Me.txtExcessCheque)
        Me.gbCheque.Controls.Add(Me.Label9)
        Me.gbCheque.Controls.Add(Me.txtTotCheque)
        Me.gbCheque.Controls.Add(Me.Label8)
        Me.gbCheque.Controls.Add(Me.Label6)
        Me.gbCheque.Controls.Add(Me.dgChequeStat)
        Me.gbCheque.Controls.Add(Me.Label5)
        Me.gbCheque.Controls.Add(Me.Label4)
        Me.gbCheque.Controls.Add(Me.dgChequeDetails)
        Me.gbCheque.Dock = CType(resources.GetObject("gbCheque.Dock"), System.Windows.Forms.DockStyle)
        Me.gbCheque.Enabled = CType(resources.GetObject("gbCheque.Enabled"), Boolean)
        Me.gbCheque.Font = CType(resources.GetObject("gbCheque.Font"), System.Drawing.Font)
        Me.gbCheque.ImeMode = CType(resources.GetObject("gbCheque.ImeMode"), System.Windows.Forms.ImeMode)
        Me.gbCheque.Location = CType(resources.GetObject("gbCheque.Location"), System.Drawing.Point)
        Me.gbCheque.Name = "gbCheque"
        Me.gbCheque.RightToLeft = CType(resources.GetObject("gbCheque.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.gbCheque.Size = CType(resources.GetObject("gbCheque.Size"), System.Drawing.Size)
        Me.gbCheque.TabIndex = CType(resources.GetObject("gbCheque.TabIndex"), Integer)
        Me.gbCheque.TabStop = False
        Me.gbCheque.Text = resources.GetString("gbCheque.Text")
        Me.gbCheque.Visible = CType(resources.GetObject("gbCheque.Visible"), Boolean)
        '
        'txtExcessCheque
        '
        Me.txtExcessCheque.AccessibleDescription = resources.GetString("txtExcessCheque.AccessibleDescription")
        Me.txtExcessCheque.AccessibleName = resources.GetString("txtExcessCheque.AccessibleName")
        Me.txtExcessCheque.Anchor = CType(resources.GetObject("txtExcessCheque.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtExcessCheque.AutoSize = CType(resources.GetObject("txtExcessCheque.AutoSize"), Boolean)
        Me.txtExcessCheque.BackColor = System.Drawing.Color.LightSteelBlue
        Me.txtExcessCheque.BackgroundImage = CType(resources.GetObject("txtExcessCheque.BackgroundImage"), System.Drawing.Image)
        Me.txtExcessCheque.Dock = CType(resources.GetObject("txtExcessCheque.Dock"), System.Windows.Forms.DockStyle)
        Me.txtExcessCheque.Enabled = CType(resources.GetObject("txtExcessCheque.Enabled"), Boolean)
        Me.txtExcessCheque.Font = CType(resources.GetObject("txtExcessCheque.Font"), System.Drawing.Font)
        Me.txtExcessCheque.ImeMode = CType(resources.GetObject("txtExcessCheque.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtExcessCheque.Location = CType(resources.GetObject("txtExcessCheque.Location"), System.Drawing.Point)
        Me.txtExcessCheque.MaxLength = CType(resources.GetObject("txtExcessCheque.MaxLength"), Integer)
        Me.txtExcessCheque.Multiline = CType(resources.GetObject("txtExcessCheque.Multiline"), Boolean)
        Me.txtExcessCheque.Name = "txtExcessCheque"
        Me.txtExcessCheque.PasswordChar = CType(resources.GetObject("txtExcessCheque.PasswordChar"), Char)
        Me.txtExcessCheque.ReadOnly = True
        Me.txtExcessCheque.RightToLeft = CType(resources.GetObject("txtExcessCheque.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtExcessCheque.ScrollBars = CType(resources.GetObject("txtExcessCheque.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtExcessCheque.Size = CType(resources.GetObject("txtExcessCheque.Size"), System.Drawing.Size)
        Me.txtExcessCheque.TabIndex = CType(resources.GetObject("txtExcessCheque.TabIndex"), Integer)
        Me.txtExcessCheque.TabStop = False
        Me.txtExcessCheque.Text = resources.GetString("txtExcessCheque.Text")
        Me.txtExcessCheque.TextAlign = CType(resources.GetObject("txtExcessCheque.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtExcessCheque.Visible = CType(resources.GetObject("txtExcessCheque.Visible"), Boolean)
        Me.txtExcessCheque.WordWrap = CType(resources.GetObject("txtExcessCheque.WordWrap"), Boolean)
        '
        'Label9
        '
        Me.Label9.AccessibleDescription = resources.GetString("Label9.AccessibleDescription")
        Me.Label9.AccessibleName = resources.GetString("Label9.AccessibleName")
        Me.Label9.Anchor = CType(resources.GetObject("Label9.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label9.AutoSize = CType(resources.GetObject("Label9.AutoSize"), Boolean)
        Me.Label9.Dock = CType(resources.GetObject("Label9.Dock"), System.Windows.Forms.DockStyle)
        Me.Label9.Enabled = CType(resources.GetObject("Label9.Enabled"), Boolean)
        Me.Label9.Font = CType(resources.GetObject("Label9.Font"), System.Drawing.Font)
        Me.Label9.ForeColor = System.Drawing.SystemColors.Info
        Me.Label9.Image = CType(resources.GetObject("Label9.Image"), System.Drawing.Image)
        Me.Label9.ImageAlign = CType(resources.GetObject("Label9.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label9.ImageIndex = CType(resources.GetObject("Label9.ImageIndex"), Integer)
        Me.Label9.ImeMode = CType(resources.GetObject("Label9.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label9.Location = CType(resources.GetObject("Label9.Location"), System.Drawing.Point)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = CType(resources.GetObject("Label9.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label9.Size = CType(resources.GetObject("Label9.Size"), System.Drawing.Size)
        Me.Label9.TabIndex = CType(resources.GetObject("Label9.TabIndex"), Integer)
        Me.Label9.Text = resources.GetString("Label9.Text")
        Me.Label9.TextAlign = CType(resources.GetObject("Label9.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label9.Visible = CType(resources.GetObject("Label9.Visible"), Boolean)
        '
        'txtTotCheque
        '
        Me.txtTotCheque.AccessibleDescription = resources.GetString("txtTotCheque.AccessibleDescription")
        Me.txtTotCheque.AccessibleName = resources.GetString("txtTotCheque.AccessibleName")
        Me.txtTotCheque.Anchor = CType(resources.GetObject("txtTotCheque.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtTotCheque.AutoSize = CType(resources.GetObject("txtTotCheque.AutoSize"), Boolean)
        Me.txtTotCheque.BackColor = System.Drawing.Color.AliceBlue
        Me.txtTotCheque.BackgroundImage = CType(resources.GetObject("txtTotCheque.BackgroundImage"), System.Drawing.Image)
        Me.txtTotCheque.Dock = CType(resources.GetObject("txtTotCheque.Dock"), System.Windows.Forms.DockStyle)
        Me.txtTotCheque.Enabled = CType(resources.GetObject("txtTotCheque.Enabled"), Boolean)
        Me.txtTotCheque.Font = CType(resources.GetObject("txtTotCheque.Font"), System.Drawing.Font)
        Me.txtTotCheque.ImeMode = CType(resources.GetObject("txtTotCheque.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtTotCheque.Location = CType(resources.GetObject("txtTotCheque.Location"), System.Drawing.Point)
        Me.txtTotCheque.MaxLength = CType(resources.GetObject("txtTotCheque.MaxLength"), Integer)
        Me.txtTotCheque.Multiline = CType(resources.GetObject("txtTotCheque.Multiline"), Boolean)
        Me.txtTotCheque.Name = "txtTotCheque"
        Me.txtTotCheque.PasswordChar = CType(resources.GetObject("txtTotCheque.PasswordChar"), Char)
        Me.txtTotCheque.ReadOnly = True
        Me.txtTotCheque.RightToLeft = CType(resources.GetObject("txtTotCheque.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtTotCheque.ScrollBars = CType(resources.GetObject("txtTotCheque.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtTotCheque.Size = CType(resources.GetObject("txtTotCheque.Size"), System.Drawing.Size)
        Me.txtTotCheque.TabIndex = CType(resources.GetObject("txtTotCheque.TabIndex"), Integer)
        Me.txtTotCheque.TabStop = False
        Me.txtTotCheque.Text = resources.GetString("txtTotCheque.Text")
        Me.txtTotCheque.TextAlign = CType(resources.GetObject("txtTotCheque.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtTotCheque.Visible = CType(resources.GetObject("txtTotCheque.Visible"), Boolean)
        Me.txtTotCheque.WordWrap = CType(resources.GetObject("txtTotCheque.WordWrap"), Boolean)
        '
        'Label8
        '
        Me.Label8.AccessibleDescription = resources.GetString("Label8.AccessibleDescription")
        Me.Label8.AccessibleName = resources.GetString("Label8.AccessibleName")
        Me.Label8.Anchor = CType(resources.GetObject("Label8.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = CType(resources.GetObject("Label8.AutoSize"), Boolean)
        Me.Label8.Dock = CType(resources.GetObject("Label8.Dock"), System.Windows.Forms.DockStyle)
        Me.Label8.Enabled = CType(resources.GetObject("Label8.Enabled"), Boolean)
        Me.Label8.Font = CType(resources.GetObject("Label8.Font"), System.Drawing.Font)
        Me.Label8.ForeColor = System.Drawing.SystemColors.Info
        Me.Label8.Image = CType(resources.GetObject("Label8.Image"), System.Drawing.Image)
        Me.Label8.ImageAlign = CType(resources.GetObject("Label8.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label8.ImageIndex = CType(resources.GetObject("Label8.ImageIndex"), Integer)
        Me.Label8.ImeMode = CType(resources.GetObject("Label8.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label8.Location = CType(resources.GetObject("Label8.Location"), System.Drawing.Point)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = CType(resources.GetObject("Label8.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label8.Size = CType(resources.GetObject("Label8.Size"), System.Drawing.Size)
        Me.Label8.TabIndex = CType(resources.GetObject("Label8.TabIndex"), Integer)
        Me.Label8.Text = resources.GetString("Label8.Text")
        Me.Label8.TextAlign = CType(resources.GetObject("Label8.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label8.Visible = CType(resources.GetObject("Label8.Visible"), Boolean)
        '
        'Label6
        '
        Me.Label6.AccessibleDescription = resources.GetString("Label6.AccessibleDescription")
        Me.Label6.AccessibleName = resources.GetString("Label6.AccessibleName")
        Me.Label6.Anchor = CType(resources.GetObject("Label6.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = CType(resources.GetObject("Label6.AutoSize"), Boolean)
        Me.Label6.Dock = CType(resources.GetObject("Label6.Dock"), System.Windows.Forms.DockStyle)
        Me.Label6.Enabled = CType(resources.GetObject("Label6.Enabled"), Boolean)
        Me.Label6.Font = CType(resources.GetObject("Label6.Font"), System.Drawing.Font)
        Me.Label6.ForeColor = System.Drawing.SystemColors.Info
        Me.Label6.Image = CType(resources.GetObject("Label6.Image"), System.Drawing.Image)
        Me.Label6.ImageAlign = CType(resources.GetObject("Label6.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label6.ImageIndex = CType(resources.GetObject("Label6.ImageIndex"), Integer)
        Me.Label6.ImeMode = CType(resources.GetObject("Label6.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label6.Location = CType(resources.GetObject("Label6.Location"), System.Drawing.Point)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = CType(resources.GetObject("Label6.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label6.Size = CType(resources.GetObject("Label6.Size"), System.Drawing.Size)
        Me.Label6.TabIndex = CType(resources.GetObject("Label6.TabIndex"), Integer)
        Me.Label6.Text = resources.GetString("Label6.Text")
        Me.Label6.TextAlign = CType(resources.GetObject("Label6.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label6.Visible = CType(resources.GetObject("Label6.Visible"), Boolean)
        '
        'dgChequeStat
        '
        Me.dgChequeStat.AccessibleDescription = resources.GetString("dgChequeStat.AccessibleDescription")
        Me.dgChequeStat.AccessibleName = resources.GetString("dgChequeStat.AccessibleName")
        Me.dgChequeStat.AlternatingBackColor = System.Drawing.Color.LightGray
        Me.dgChequeStat.Anchor = CType(resources.GetObject("dgChequeStat.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.dgChequeStat.BackColor = System.Drawing.Color.Gainsboro
        Me.dgChequeStat.BackgroundColor = System.Drawing.Color.Silver
        Me.dgChequeStat.BackgroundImage = CType(resources.GetObject("dgChequeStat.BackgroundImage"), System.Drawing.Image)
        Me.dgChequeStat.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgChequeStat.CaptionBackColor = System.Drawing.Color.LightSteelBlue
        Me.dgChequeStat.CaptionFont = CType(resources.GetObject("dgChequeStat.CaptionFont"), System.Drawing.Font)
        Me.dgChequeStat.CaptionForeColor = System.Drawing.Color.MidnightBlue
        Me.dgChequeStat.CaptionText = resources.GetString("dgChequeStat.CaptionText")
        Me.dgChequeStat.DataMember = ""
        Me.dgChequeStat.Dock = CType(resources.GetObject("dgChequeStat.Dock"), System.Windows.Forms.DockStyle)
        Me.dgChequeStat.Enabled = CType(resources.GetObject("dgChequeStat.Enabled"), Boolean)
        Me.dgChequeStat.FlatMode = True
        Me.dgChequeStat.Font = CType(resources.GetObject("dgChequeStat.Font"), System.Drawing.Font)
        Me.dgChequeStat.ForeColor = System.Drawing.Color.Black
        Me.dgChequeStat.GridLineColor = System.Drawing.Color.DimGray
        Me.dgChequeStat.GridLineStyle = System.Windows.Forms.DataGridLineStyle.None
        Me.dgChequeStat.HeaderBackColor = System.Drawing.Color.MidnightBlue
        Me.dgChequeStat.HeaderFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.dgChequeStat.HeaderForeColor = System.Drawing.Color.White
        Me.dgChequeStat.ImeMode = CType(resources.GetObject("dgChequeStat.ImeMode"), System.Windows.Forms.ImeMode)
        Me.dgChequeStat.LinkColor = System.Drawing.Color.MidnightBlue
        Me.dgChequeStat.Location = CType(resources.GetObject("dgChequeStat.Location"), System.Drawing.Point)
        Me.dgChequeStat.Name = "dgChequeStat"
        Me.dgChequeStat.ParentRowsBackColor = System.Drawing.Color.DarkGray
        Me.dgChequeStat.ParentRowsForeColor = System.Drawing.Color.Black
        Me.dgChequeStat.ReadOnly = True
        Me.dgChequeStat.RightToLeft = CType(resources.GetObject("dgChequeStat.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.dgChequeStat.SelectionBackColor = System.Drawing.Color.CadetBlue
        Me.dgChequeStat.SelectionForeColor = System.Drawing.Color.White
        Me.dgChequeStat.Size = CType(resources.GetObject("dgChequeStat.Size"), System.Drawing.Size)
        Me.dgChequeStat.TabIndex = CType(resources.GetObject("dgChequeStat.TabIndex"), Integer)
        Me.dgChequeStat.Visible = CType(resources.GetObject("dgChequeStat.Visible"), Boolean)
        '
        'Label5
        '
        Me.Label5.AccessibleDescription = resources.GetString("Label5.AccessibleDescription")
        Me.Label5.AccessibleName = resources.GetString("Label5.AccessibleName")
        Me.Label5.Anchor = CType(resources.GetObject("Label5.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = CType(resources.GetObject("Label5.AutoSize"), Boolean)
        Me.Label5.Dock = CType(resources.GetObject("Label5.Dock"), System.Windows.Forms.DockStyle)
        Me.Label5.Enabled = CType(resources.GetObject("Label5.Enabled"), Boolean)
        Me.Label5.Font = CType(resources.GetObject("Label5.Font"), System.Drawing.Font)
        Me.Label5.ForeColor = System.Drawing.SystemColors.Info
        Me.Label5.Image = CType(resources.GetObject("Label5.Image"), System.Drawing.Image)
        Me.Label5.ImageAlign = CType(resources.GetObject("Label5.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label5.ImageIndex = CType(resources.GetObject("Label5.ImageIndex"), Integer)
        Me.Label5.ImeMode = CType(resources.GetObject("Label5.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label5.Location = CType(resources.GetObject("Label5.Location"), System.Drawing.Point)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = CType(resources.GetObject("Label5.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label5.Size = CType(resources.GetObject("Label5.Size"), System.Drawing.Size)
        Me.Label5.TabIndex = CType(resources.GetObject("Label5.TabIndex"), Integer)
        Me.Label5.Text = resources.GetString("Label5.Text")
        Me.Label5.TextAlign = CType(resources.GetObject("Label5.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label5.Visible = CType(resources.GetObject("Label5.Visible"), Boolean)
        '
        'Label4
        '
        Me.Label4.AccessibleDescription = resources.GetString("Label4.AccessibleDescription")
        Me.Label4.AccessibleName = resources.GetString("Label4.AccessibleName")
        Me.Label4.Anchor = CType(resources.GetObject("Label4.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = CType(resources.GetObject("Label4.AutoSize"), Boolean)
        Me.Label4.BackColor = System.Drawing.Color.SteelBlue
        Me.Label4.Dock = CType(resources.GetObject("Label4.Dock"), System.Windows.Forms.DockStyle)
        Me.Label4.Enabled = CType(resources.GetObject("Label4.Enabled"), Boolean)
        Me.Label4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label4.Font = CType(resources.GetObject("Label4.Font"), System.Drawing.Font)
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label4.Image = CType(resources.GetObject("Label4.Image"), System.Drawing.Image)
        Me.Label4.ImageAlign = CType(resources.GetObject("Label4.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label4.ImageIndex = CType(resources.GetObject("Label4.ImageIndex"), Integer)
        Me.Label4.ImeMode = CType(resources.GetObject("Label4.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label4.Location = CType(resources.GetObject("Label4.Location"), System.Drawing.Point)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = CType(resources.GetObject("Label4.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label4.Size = CType(resources.GetObject("Label4.Size"), System.Drawing.Size)
        Me.Label4.TabIndex = CType(resources.GetObject("Label4.TabIndex"), Integer)
        Me.Label4.Text = resources.GetString("Label4.Text")
        Me.Label4.TextAlign = CType(resources.GetObject("Label4.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label4.Visible = CType(resources.GetObject("Label4.Visible"), Boolean)
        '
        'dgChequeDetails
        '
        Me.dgChequeDetails.AccessibleDescription = resources.GetString("dgChequeDetails.AccessibleDescription")
        Me.dgChequeDetails.AccessibleName = resources.GetString("dgChequeDetails.AccessibleName")
        Me.dgChequeDetails.AlternatingBackColor = System.Drawing.Color.LightGray
        Me.dgChequeDetails.Anchor = CType(resources.GetObject("dgChequeDetails.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.dgChequeDetails.BackColor = System.Drawing.Color.White
        Me.dgChequeDetails.BackgroundColor = System.Drawing.Color.Silver
        Me.dgChequeDetails.BackgroundImage = CType(resources.GetObject("dgChequeDetails.BackgroundImage"), System.Drawing.Image)
        Me.dgChequeDetails.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgChequeDetails.CaptionBackColor = System.Drawing.Color.LightSteelBlue
        Me.dgChequeDetails.CaptionFont = CType(resources.GetObject("dgChequeDetails.CaptionFont"), System.Drawing.Font)
        Me.dgChequeDetails.CaptionForeColor = System.Drawing.Color.MidnightBlue
        Me.dgChequeDetails.CaptionText = resources.GetString("dgChequeDetails.CaptionText")
        Me.dgChequeDetails.DataMember = ""
        Me.dgChequeDetails.Dock = CType(resources.GetObject("dgChequeDetails.Dock"), System.Windows.Forms.DockStyle)
        Me.dgChequeDetails.Enabled = CType(resources.GetObject("dgChequeDetails.Enabled"), Boolean)
        Me.dgChequeDetails.FlatMode = True
        Me.dgChequeDetails.Font = CType(resources.GetObject("dgChequeDetails.Font"), System.Drawing.Font)
        Me.dgChequeDetails.ForeColor = System.Drawing.Color.Black
        Me.dgChequeDetails.GridLineColor = System.Drawing.Color.DimGray
        Me.dgChequeDetails.GridLineStyle = System.Windows.Forms.DataGridLineStyle.None
        Me.dgChequeDetails.HeaderBackColor = System.Drawing.Color.MidnightBlue
        Me.dgChequeDetails.HeaderFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.dgChequeDetails.HeaderForeColor = System.Drawing.Color.White
        Me.dgChequeDetails.ImeMode = CType(resources.GetObject("dgChequeDetails.ImeMode"), System.Windows.Forms.ImeMode)
        Me.dgChequeDetails.LinkColor = System.Drawing.Color.MidnightBlue
        Me.dgChequeDetails.Location = CType(resources.GetObject("dgChequeDetails.Location"), System.Drawing.Point)
        Me.dgChequeDetails.Name = "dgChequeDetails"
        Me.dgChequeDetails.ParentRowsBackColor = System.Drawing.Color.DarkGray
        Me.dgChequeDetails.ParentRowsForeColor = System.Drawing.Color.Black
        Me.dgChequeDetails.PreferredColumnWidth = 100
        Me.dgChequeDetails.ReadOnly = True
        Me.dgChequeDetails.RightToLeft = CType(resources.GetObject("dgChequeDetails.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.dgChequeDetails.SelectionBackColor = System.Drawing.Color.CadetBlue
        Me.dgChequeDetails.SelectionForeColor = System.Drawing.Color.White
        Me.dgChequeDetails.Size = CType(resources.GetObject("dgChequeDetails.Size"), System.Drawing.Size)
        Me.dgChequeDetails.TabIndex = CType(resources.GetObject("dgChequeDetails.TabIndex"), Integer)
        Me.dgChequeDetails.Visible = CType(resources.GetObject("dgChequeDetails.Visible"), Boolean)
        '
        'gbControl
        '
        Me.gbControl.AccessibleDescription = resources.GetString("gbControl.AccessibleDescription")
        Me.gbControl.AccessibleName = resources.GetString("gbControl.AccessibleName")
        Me.gbControl.Anchor = CType(resources.GetObject("gbControl.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.gbControl.BackgroundImage = CType(resources.GetObject("gbControl.BackgroundImage"), System.Drawing.Image)
        Me.gbControl.Controls.Add(Me.Label36)
        Me.gbControl.Controls.Add(Me.txtGrandTot)
        Me.gbControl.Controls.Add(Me.Label11)
        Me.gbControl.Controls.Add(Me.txtRemarks)
        Me.gbControl.Controls.Add(Me.btnNew)
        Me.gbControl.Controls.Add(Me.btnSave)
        Me.gbControl.Controls.Add(Me.btnPrint)
        Me.gbControl.Controls.Add(Me.btnCLOSE)
        Me.gbControl.Dock = CType(resources.GetObject("gbControl.Dock"), System.Windows.Forms.DockStyle)
        Me.gbControl.Enabled = CType(resources.GetObject("gbControl.Enabled"), Boolean)
        Me.gbControl.Font = CType(resources.GetObject("gbControl.Font"), System.Drawing.Font)
        Me.gbControl.ImeMode = CType(resources.GetObject("gbControl.ImeMode"), System.Windows.Forms.ImeMode)
        Me.gbControl.Location = CType(resources.GetObject("gbControl.Location"), System.Drawing.Point)
        Me.gbControl.Name = "gbControl"
        Me.gbControl.RightToLeft = CType(resources.GetObject("gbControl.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.gbControl.Size = CType(resources.GetObject("gbControl.Size"), System.Drawing.Size)
        Me.gbControl.TabIndex = CType(resources.GetObject("gbControl.TabIndex"), Integer)
        Me.gbControl.TabStop = False
        Me.gbControl.Text = resources.GetString("gbControl.Text")
        Me.gbControl.Visible = CType(resources.GetObject("gbControl.Visible"), Boolean)
        '
        'Label36
        '
        Me.Label36.AccessibleDescription = resources.GetString("Label36.AccessibleDescription")
        Me.Label36.AccessibleName = resources.GetString("Label36.AccessibleName")
        Me.Label36.Anchor = CType(resources.GetObject("Label36.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label36.AutoSize = CType(resources.GetObject("Label36.AutoSize"), Boolean)
        Me.Label36.Dock = CType(resources.GetObject("Label36.Dock"), System.Windows.Forms.DockStyle)
        Me.Label36.Enabled = CType(resources.GetObject("Label36.Enabled"), Boolean)
        Me.Label36.Font = CType(resources.GetObject("Label36.Font"), System.Drawing.Font)
        Me.Label36.ForeColor = System.Drawing.SystemColors.Info
        Me.Label36.Image = CType(resources.GetObject("Label36.Image"), System.Drawing.Image)
        Me.Label36.ImageAlign = CType(resources.GetObject("Label36.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label36.ImageIndex = CType(resources.GetObject("Label36.ImageIndex"), Integer)
        Me.Label36.ImeMode = CType(resources.GetObject("Label36.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label36.Location = CType(resources.GetObject("Label36.Location"), System.Drawing.Point)
        Me.Label36.Name = "Label36"
        Me.Label36.RightToLeft = CType(resources.GetObject("Label36.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label36.Size = CType(resources.GetObject("Label36.Size"), System.Drawing.Size)
        Me.Label36.TabIndex = CType(resources.GetObject("Label36.TabIndex"), Integer)
        Me.Label36.Text = resources.GetString("Label36.Text")
        Me.Label36.TextAlign = CType(resources.GetObject("Label36.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label36.Visible = CType(resources.GetObject("Label36.Visible"), Boolean)
        '
        'txtGrandTot
        '
        Me.txtGrandTot.AccessibleDescription = resources.GetString("txtGrandTot.AccessibleDescription")
        Me.txtGrandTot.AccessibleName = resources.GetString("txtGrandTot.AccessibleName")
        Me.txtGrandTot.Anchor = CType(resources.GetObject("txtGrandTot.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtGrandTot.AutoSize = CType(resources.GetObject("txtGrandTot.AutoSize"), Boolean)
        Me.txtGrandTot.BackColor = System.Drawing.Color.LightSteelBlue
        Me.txtGrandTot.BackgroundImage = CType(resources.GetObject("txtGrandTot.BackgroundImage"), System.Drawing.Image)
        Me.txtGrandTot.Dock = CType(resources.GetObject("txtGrandTot.Dock"), System.Windows.Forms.DockStyle)
        Me.txtGrandTot.Enabled = CType(resources.GetObject("txtGrandTot.Enabled"), Boolean)
        Me.txtGrandTot.Font = CType(resources.GetObject("txtGrandTot.Font"), System.Drawing.Font)
        Me.txtGrandTot.ImeMode = CType(resources.GetObject("txtGrandTot.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtGrandTot.Location = CType(resources.GetObject("txtGrandTot.Location"), System.Drawing.Point)
        Me.txtGrandTot.MaxLength = CType(resources.GetObject("txtGrandTot.MaxLength"), Integer)
        Me.txtGrandTot.Multiline = CType(resources.GetObject("txtGrandTot.Multiline"), Boolean)
        Me.txtGrandTot.Name = "txtGrandTot"
        Me.txtGrandTot.PasswordChar = CType(resources.GetObject("txtGrandTot.PasswordChar"), Char)
        Me.txtGrandTot.ReadOnly = True
        Me.txtGrandTot.RightToLeft = CType(resources.GetObject("txtGrandTot.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtGrandTot.ScrollBars = CType(resources.GetObject("txtGrandTot.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtGrandTot.Size = CType(resources.GetObject("txtGrandTot.Size"), System.Drawing.Size)
        Me.txtGrandTot.TabIndex = CType(resources.GetObject("txtGrandTot.TabIndex"), Integer)
        Me.txtGrandTot.TabStop = False
        Me.txtGrandTot.Text = resources.GetString("txtGrandTot.Text")
        Me.txtGrandTot.TextAlign = CType(resources.GetObject("txtGrandTot.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtGrandTot.Visible = CType(resources.GetObject("txtGrandTot.Visible"), Boolean)
        Me.txtGrandTot.WordWrap = CType(resources.GetObject("txtGrandTot.WordWrap"), Boolean)
        '
        'Label11
        '
        Me.Label11.AccessibleDescription = resources.GetString("Label11.AccessibleDescription")
        Me.Label11.AccessibleName = resources.GetString("Label11.AccessibleName")
        Me.Label11.Anchor = CType(resources.GetObject("Label11.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label11.AutoSize = CType(resources.GetObject("Label11.AutoSize"), Boolean)
        Me.Label11.Dock = CType(resources.GetObject("Label11.Dock"), System.Windows.Forms.DockStyle)
        Me.Label11.Enabled = CType(resources.GetObject("Label11.Enabled"), Boolean)
        Me.Label11.Font = CType(resources.GetObject("Label11.Font"), System.Drawing.Font)
        Me.Label11.ForeColor = System.Drawing.SystemColors.Info
        Me.Label11.Image = CType(resources.GetObject("Label11.Image"), System.Drawing.Image)
        Me.Label11.ImageAlign = CType(resources.GetObject("Label11.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label11.ImageIndex = CType(resources.GetObject("Label11.ImageIndex"), Integer)
        Me.Label11.ImeMode = CType(resources.GetObject("Label11.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label11.Location = CType(resources.GetObject("Label11.Location"), System.Drawing.Point)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = CType(resources.GetObject("Label11.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label11.Size = CType(resources.GetObject("Label11.Size"), System.Drawing.Size)
        Me.Label11.TabIndex = CType(resources.GetObject("Label11.TabIndex"), Integer)
        Me.Label11.Text = resources.GetString("Label11.Text")
        Me.Label11.TextAlign = CType(resources.GetObject("Label11.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label11.Visible = CType(resources.GetObject("Label11.Visible"), Boolean)
        '
        'txtRemarks
        '
        Me.txtRemarks.AccessibleDescription = resources.GetString("txtRemarks.AccessibleDescription")
        Me.txtRemarks.AccessibleName = resources.GetString("txtRemarks.AccessibleName")
        Me.txtRemarks.Anchor = CType(resources.GetObject("txtRemarks.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.txtRemarks.AutoSize = CType(resources.GetObject("txtRemarks.AutoSize"), Boolean)
        Me.txtRemarks.BackgroundImage = CType(resources.GetObject("txtRemarks.BackgroundImage"), System.Drawing.Image)
        Me.txtRemarks.Dock = CType(resources.GetObject("txtRemarks.Dock"), System.Windows.Forms.DockStyle)
        Me.txtRemarks.Enabled = CType(resources.GetObject("txtRemarks.Enabled"), Boolean)
        Me.txtRemarks.Font = CType(resources.GetObject("txtRemarks.Font"), System.Drawing.Font)
        Me.txtRemarks.ImeMode = CType(resources.GetObject("txtRemarks.ImeMode"), System.Windows.Forms.ImeMode)
        Me.txtRemarks.Location = CType(resources.GetObject("txtRemarks.Location"), System.Drawing.Point)
        Me.txtRemarks.MaxLength = CType(resources.GetObject("txtRemarks.MaxLength"), Integer)
        Me.txtRemarks.Multiline = CType(resources.GetObject("txtRemarks.Multiline"), Boolean)
        Me.txtRemarks.Name = "txtRemarks"
        Me.txtRemarks.PasswordChar = CType(resources.GetObject("txtRemarks.PasswordChar"), Char)
        Me.txtRemarks.RightToLeft = CType(resources.GetObject("txtRemarks.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.txtRemarks.ScrollBars = CType(resources.GetObject("txtRemarks.ScrollBars"), System.Windows.Forms.ScrollBars)
        Me.txtRemarks.Size = CType(resources.GetObject("txtRemarks.Size"), System.Drawing.Size)
        Me.txtRemarks.TabIndex = CType(resources.GetObject("txtRemarks.TabIndex"), Integer)
        Me.txtRemarks.TabStop = False
        Me.txtRemarks.Text = resources.GetString("txtRemarks.Text")
        Me.txtRemarks.TextAlign = CType(resources.GetObject("txtRemarks.TextAlign"), System.Windows.Forms.HorizontalAlignment)
        Me.txtRemarks.Visible = CType(resources.GetObject("txtRemarks.Visible"), Boolean)
        Me.txtRemarks.WordWrap = CType(resources.GetObject("txtRemarks.WordWrap"), Boolean)
        '
        'btnNew
        '
        Me.btnNew.AccessibleDescription = resources.GetString("btnNew.AccessibleDescription")
        Me.btnNew.AccessibleName = resources.GetString("btnNew.AccessibleName")
        Me.btnNew.Anchor = CType(resources.GetObject("btnNew.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btnNew.BackgroundImage = CType(resources.GetObject("btnNew.BackgroundImage"), System.Drawing.Image)
        Me.btnNew.Dock = CType(resources.GetObject("btnNew.Dock"), System.Windows.Forms.DockStyle)
        Me.btnNew.Enabled = CType(resources.GetObject("btnNew.Enabled"), Boolean)
        Me.btnNew.FlatStyle = CType(resources.GetObject("btnNew.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btnNew.Font = CType(resources.GetObject("btnNew.Font"), System.Drawing.Font)
        Me.btnNew.Image = CType(resources.GetObject("btnNew.Image"), System.Drawing.Image)
        Me.btnNew.ImageAlign = CType(resources.GetObject("btnNew.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btnNew.ImageIndex = CType(resources.GetObject("btnNew.ImageIndex"), Integer)
        Me.btnNew.ImeMode = CType(resources.GetObject("btnNew.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btnNew.Location = CType(resources.GetObject("btnNew.Location"), System.Drawing.Point)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.RightToLeft = CType(resources.GetObject("btnNew.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btnNew.Size = CType(resources.GetObject("btnNew.Size"), System.Drawing.Size)
        Me.btnNew.TabIndex = CType(resources.GetObject("btnNew.TabIndex"), Integer)
        Me.btnNew.Text = resources.GetString("btnNew.Text")
        Me.btnNew.TextAlign = CType(resources.GetObject("btnNew.TextAlign"), System.Drawing.ContentAlignment)
        Me.btnNew.Visible = CType(resources.GetObject("btnNew.Visible"), Boolean)
        '
        'btnSave
        '
        Me.btnSave.AccessibleDescription = resources.GetString("btnSave.AccessibleDescription")
        Me.btnSave.AccessibleName = resources.GetString("btnSave.AccessibleName")
        Me.btnSave.Anchor = CType(resources.GetObject("btnSave.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btnSave.BackgroundImage = CType(resources.GetObject("btnSave.BackgroundImage"), System.Drawing.Image)
        Me.btnSave.Dock = CType(resources.GetObject("btnSave.Dock"), System.Windows.Forms.DockStyle)
        Me.btnSave.Enabled = CType(resources.GetObject("btnSave.Enabled"), Boolean)
        Me.btnSave.FlatStyle = CType(resources.GetObject("btnSave.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btnSave.Font = CType(resources.GetObject("btnSave.Font"), System.Drawing.Font)
        Me.btnSave.Image = CType(resources.GetObject("btnSave.Image"), System.Drawing.Image)
        Me.btnSave.ImageAlign = CType(resources.GetObject("btnSave.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btnSave.ImageIndex = CType(resources.GetObject("btnSave.ImageIndex"), Integer)
        Me.btnSave.ImeMode = CType(resources.GetObject("btnSave.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btnSave.Location = CType(resources.GetObject("btnSave.Location"), System.Drawing.Point)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.RightToLeft = CType(resources.GetObject("btnSave.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btnSave.Size = CType(resources.GetObject("btnSave.Size"), System.Drawing.Size)
        Me.btnSave.TabIndex = CType(resources.GetObject("btnSave.TabIndex"), Integer)
        Me.btnSave.Text = resources.GetString("btnSave.Text")
        Me.btnSave.TextAlign = CType(resources.GetObject("btnSave.TextAlign"), System.Drawing.ContentAlignment)
        Me.btnSave.Visible = CType(resources.GetObject("btnSave.Visible"), Boolean)
        '
        'btnPrint
        '
        Me.btnPrint.AccessibleDescription = resources.GetString("btnPrint.AccessibleDescription")
        Me.btnPrint.AccessibleName = resources.GetString("btnPrint.AccessibleName")
        Me.btnPrint.Anchor = CType(resources.GetObject("btnPrint.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btnPrint.BackgroundImage = CType(resources.GetObject("btnPrint.BackgroundImage"), System.Drawing.Image)
        Me.btnPrint.Dock = CType(resources.GetObject("btnPrint.Dock"), System.Windows.Forms.DockStyle)
        Me.btnPrint.Enabled = CType(resources.GetObject("btnPrint.Enabled"), Boolean)
        Me.btnPrint.FlatStyle = CType(resources.GetObject("btnPrint.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btnPrint.Font = CType(resources.GetObject("btnPrint.Font"), System.Drawing.Font)
        Me.btnPrint.Image = CType(resources.GetObject("btnPrint.Image"), System.Drawing.Image)
        Me.btnPrint.ImageAlign = CType(resources.GetObject("btnPrint.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btnPrint.ImageIndex = CType(resources.GetObject("btnPrint.ImageIndex"), Integer)
        Me.btnPrint.ImeMode = CType(resources.GetObject("btnPrint.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btnPrint.Location = CType(resources.GetObject("btnPrint.Location"), System.Drawing.Point)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.RightToLeft = CType(resources.GetObject("btnPrint.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btnPrint.Size = CType(resources.GetObject("btnPrint.Size"), System.Drawing.Size)
        Me.btnPrint.TabIndex = CType(resources.GetObject("btnPrint.TabIndex"), Integer)
        Me.btnPrint.Text = resources.GetString("btnPrint.Text")
        Me.btnPrint.TextAlign = CType(resources.GetObject("btnPrint.TextAlign"), System.Drawing.ContentAlignment)
        Me.btnPrint.Visible = CType(resources.GetObject("btnPrint.Visible"), Boolean)
        '
        'btnCLOSE
        '
        Me.btnCLOSE.AccessibleDescription = resources.GetString("btnCLOSE.AccessibleDescription")
        Me.btnCLOSE.AccessibleName = resources.GetString("btnCLOSE.AccessibleName")
        Me.btnCLOSE.Anchor = CType(resources.GetObject("btnCLOSE.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.btnCLOSE.BackgroundImage = CType(resources.GetObject("btnCLOSE.BackgroundImage"), System.Drawing.Image)
        Me.btnCLOSE.Dock = CType(resources.GetObject("btnCLOSE.Dock"), System.Windows.Forms.DockStyle)
        Me.btnCLOSE.Enabled = CType(resources.GetObject("btnCLOSE.Enabled"), Boolean)
        Me.btnCLOSE.FlatStyle = CType(resources.GetObject("btnCLOSE.FlatStyle"), System.Windows.Forms.FlatStyle)
        Me.btnCLOSE.Font = CType(resources.GetObject("btnCLOSE.Font"), System.Drawing.Font)
        Me.btnCLOSE.Image = CType(resources.GetObject("btnCLOSE.Image"), System.Drawing.Image)
        Me.btnCLOSE.ImageAlign = CType(resources.GetObject("btnCLOSE.ImageAlign"), System.Drawing.ContentAlignment)
        Me.btnCLOSE.ImageIndex = CType(resources.GetObject("btnCLOSE.ImageIndex"), Integer)
        Me.btnCLOSE.ImeMode = CType(resources.GetObject("btnCLOSE.ImeMode"), System.Windows.Forms.ImeMode)
        Me.btnCLOSE.Location = CType(resources.GetObject("btnCLOSE.Location"), System.Drawing.Point)
        Me.btnCLOSE.Name = "btnCLOSE"
        Me.btnCLOSE.RightToLeft = CType(resources.GetObject("btnCLOSE.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.btnCLOSE.Size = CType(resources.GetObject("btnCLOSE.Size"), System.Drawing.Size)
        Me.btnCLOSE.TabIndex = CType(resources.GetObject("btnCLOSE.TabIndex"), Integer)
        Me.btnCLOSE.Text = resources.GetString("btnCLOSE.Text")
        Me.btnCLOSE.TextAlign = CType(resources.GetObject("btnCLOSE.TextAlign"), System.Drawing.ContentAlignment)
        Me.btnCLOSE.Visible = CType(resources.GetObject("btnCLOSE.Visible"), Boolean)
        '
        'statCACBar
        '
        Me.statCACBar.AccessibleDescription = resources.GetString("statCACBar.AccessibleDescription")
        Me.statCACBar.AccessibleName = resources.GetString("statCACBar.AccessibleName")
        Me.statCACBar.Anchor = CType(resources.GetObject("statCACBar.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.statCACBar.BackgroundImage = CType(resources.GetObject("statCACBar.BackgroundImage"), System.Drawing.Image)
        Me.statCACBar.Dock = CType(resources.GetObject("statCACBar.Dock"), System.Windows.Forms.DockStyle)
        Me.statCACBar.Enabled = CType(resources.GetObject("statCACBar.Enabled"), Boolean)
        Me.statCACBar.Font = CType(resources.GetObject("statCACBar.Font"), System.Drawing.Font)
        Me.statCACBar.ImeMode = CType(resources.GetObject("statCACBar.ImeMode"), System.Windows.Forms.ImeMode)
        Me.statCACBar.Location = CType(resources.GetObject("statCACBar.Location"), System.Drawing.Point)
        Me.statCACBar.Name = "statCACBar"
        Me.statCACBar.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.statPanelUser, Me.statPanelDate, Me.statPanelTime})
        Me.statCACBar.RightToLeft = CType(resources.GetObject("statCACBar.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.statCACBar.ShowPanels = True
        Me.statCACBar.Size = CType(resources.GetObject("statCACBar.Size"), System.Drawing.Size)
        Me.statCACBar.SizingGrip = False
        Me.statCACBar.TabIndex = CType(resources.GetObject("statCACBar.TabIndex"), Integer)
        Me.statCACBar.Text = resources.GetString("statCACBar.Text")
        Me.statCACBar.Visible = CType(resources.GetObject("statCACBar.Visible"), Boolean)
        '
        'statPanelUser
        '
        Me.statPanelUser.Alignment = CType(resources.GetObject("statPanelUser.Alignment"), System.Windows.Forms.HorizontalAlignment)
        Me.statPanelUser.Icon = CType(resources.GetObject("statPanelUser.Icon"), System.Drawing.Icon)
        Me.statPanelUser.MinWidth = CType(resources.GetObject("statPanelUser.MinWidth"), Integer)
        Me.statPanelUser.Text = resources.GetString("statPanelUser.Text")
        Me.statPanelUser.ToolTipText = resources.GetString("statPanelUser.ToolTipText")
        Me.statPanelUser.Width = CType(resources.GetObject("statPanelUser.Width"), Integer)
        '
        'statPanelDate
        '
        Me.statPanelDate.Alignment = CType(resources.GetObject("statPanelDate.Alignment"), System.Windows.Forms.HorizontalAlignment)
        Me.statPanelDate.Icon = CType(resources.GetObject("statPanelDate.Icon"), System.Drawing.Icon)
        Me.statPanelDate.MinWidth = CType(resources.GetObject("statPanelDate.MinWidth"), Integer)
        Me.statPanelDate.Text = resources.GetString("statPanelDate.Text")
        Me.statPanelDate.ToolTipText = resources.GetString("statPanelDate.ToolTipText")
        Me.statPanelDate.Width = CType(resources.GetObject("statPanelDate.Width"), Integer)
        '
        'statPanelTime
        '
        Me.statPanelTime.Alignment = CType(resources.GetObject("statPanelTime.Alignment"), System.Windows.Forms.HorizontalAlignment)
        Me.statPanelTime.Icon = CType(resources.GetObject("statPanelTime.Icon"), System.Drawing.Icon)
        Me.statPanelTime.MinWidth = CType(resources.GetObject("statPanelTime.MinWidth"), Integer)
        Me.statPanelTime.Text = resources.GetString("statPanelTime.Text")
        Me.statPanelTime.ToolTipText = resources.GetString("statPanelTime.ToolTipText")
        Me.statPanelTime.Width = CType(resources.GetObject("statPanelTime.Width"), Integer)
        '
        'lblID
        '
        Me.lblID.AccessibleDescription = resources.GetString("lblID.AccessibleDescription")
        Me.lblID.AccessibleName = resources.GetString("lblID.AccessibleName")
        Me.lblID.Anchor = CType(resources.GetObject("lblID.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.lblID.AutoSize = CType(resources.GetObject("lblID.AutoSize"), Boolean)
        Me.lblID.Dock = CType(resources.GetObject("lblID.Dock"), System.Windows.Forms.DockStyle)
        Me.lblID.Enabled = CType(resources.GetObject("lblID.Enabled"), Boolean)
        Me.lblID.Font = CType(resources.GetObject("lblID.Font"), System.Drawing.Font)
        Me.lblID.ForeColor = System.Drawing.SystemColors.Info
        Me.lblID.Image = CType(resources.GetObject("lblID.Image"), System.Drawing.Image)
        Me.lblID.ImageAlign = CType(resources.GetObject("lblID.ImageAlign"), System.Drawing.ContentAlignment)
        Me.lblID.ImageIndex = CType(resources.GetObject("lblID.ImageIndex"), Integer)
        Me.lblID.ImeMode = CType(resources.GetObject("lblID.ImeMode"), System.Windows.Forms.ImeMode)
        Me.lblID.Location = CType(resources.GetObject("lblID.Location"), System.Drawing.Point)
        Me.lblID.Name = "lblID"
        Me.lblID.RightToLeft = CType(resources.GetObject("lblID.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.lblID.Size = CType(resources.GetObject("lblID.Size"), System.Drawing.Size)
        Me.lblID.TabIndex = CType(resources.GetObject("lblID.TabIndex"), Integer)
        Me.lblID.Text = resources.GetString("lblID.Text")
        Me.lblID.TextAlign = CType(resources.GetObject("lblID.TextAlign"), System.Drawing.ContentAlignment)
        Me.lblID.Visible = CType(resources.GetObject("lblID.Visible"), Boolean)
        '
        'gbHeader
        '
        Me.gbHeader.AccessibleDescription = resources.GetString("gbHeader.AccessibleDescription")
        Me.gbHeader.AccessibleName = resources.GetString("gbHeader.AccessibleName")
        Me.gbHeader.Anchor = CType(resources.GetObject("gbHeader.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.gbHeader.BackgroundImage = CType(resources.GetObject("gbHeader.BackgroundImage"), System.Drawing.Image)
        Me.gbHeader.Controls.Add(Me.Label7)
        Me.gbHeader.Controls.Add(Me.PictureBox1)
        Me.gbHeader.Dock = CType(resources.GetObject("gbHeader.Dock"), System.Windows.Forms.DockStyle)
        Me.gbHeader.Enabled = CType(resources.GetObject("gbHeader.Enabled"), Boolean)
        Me.gbHeader.Font = CType(resources.GetObject("gbHeader.Font"), System.Drawing.Font)
        Me.gbHeader.ImeMode = CType(resources.GetObject("gbHeader.ImeMode"), System.Windows.Forms.ImeMode)
        Me.gbHeader.Location = CType(resources.GetObject("gbHeader.Location"), System.Drawing.Point)
        Me.gbHeader.Name = "gbHeader"
        Me.gbHeader.RightToLeft = CType(resources.GetObject("gbHeader.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.gbHeader.Size = CType(resources.GetObject("gbHeader.Size"), System.Drawing.Size)
        Me.gbHeader.TabIndex = CType(resources.GetObject("gbHeader.TabIndex"), Integer)
        Me.gbHeader.TabStop = False
        Me.gbHeader.Text = resources.GetString("gbHeader.Text")
        Me.gbHeader.Visible = CType(resources.GetObject("gbHeader.Visible"), Boolean)
        '
        'Label7
        '
        Me.Label7.AccessibleDescription = resources.GetString("Label7.AccessibleDescription")
        Me.Label7.AccessibleName = resources.GetString("Label7.AccessibleName")
        Me.Label7.Anchor = CType(resources.GetObject("Label7.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = CType(resources.GetObject("Label7.AutoSize"), Boolean)
        Me.Label7.Dock = CType(resources.GetObject("Label7.Dock"), System.Windows.Forms.DockStyle)
        Me.Label7.Enabled = CType(resources.GetObject("Label7.Enabled"), Boolean)
        Me.Label7.Font = CType(resources.GetObject("Label7.Font"), System.Drawing.Font)
        Me.Label7.ForeColor = System.Drawing.Color.AliceBlue
        Me.Label7.Image = CType(resources.GetObject("Label7.Image"), System.Drawing.Image)
        Me.Label7.ImageAlign = CType(resources.GetObject("Label7.ImageAlign"), System.Drawing.ContentAlignment)
        Me.Label7.ImageIndex = CType(resources.GetObject("Label7.ImageIndex"), Integer)
        Me.Label7.ImeMode = CType(resources.GetObject("Label7.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Label7.Location = CType(resources.GetObject("Label7.Location"), System.Drawing.Point)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = CType(resources.GetObject("Label7.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.Label7.Size = CType(resources.GetObject("Label7.Size"), System.Drawing.Size)
        Me.Label7.TabIndex = CType(resources.GetObject("Label7.TabIndex"), Integer)
        Me.Label7.Text = resources.GetString("Label7.Text")
        Me.Label7.TextAlign = CType(resources.GetObject("Label7.TextAlign"), System.Drawing.ContentAlignment)
        Me.Label7.Visible = CType(resources.GetObject("Label7.Visible"), Boolean)
        '
        'PictureBox1
        '
        Me.PictureBox1.AccessibleDescription = resources.GetString("PictureBox1.AccessibleDescription")
        Me.PictureBox1.AccessibleName = resources.GetString("PictureBox1.AccessibleName")
        Me.PictureBox1.Anchor = CType(resources.GetObject("PictureBox1.Anchor"), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.Dock = CType(resources.GetObject("PictureBox1.Dock"), System.Windows.Forms.DockStyle)
        Me.PictureBox1.Enabled = CType(resources.GetObject("PictureBox1.Enabled"), Boolean)
        Me.PictureBox1.Font = CType(resources.GetObject("PictureBox1.Font"), System.Drawing.Font)
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.ImeMode = CType(resources.GetObject("PictureBox1.ImeMode"), System.Windows.Forms.ImeMode)
        Me.PictureBox1.Location = CType(resources.GetObject("PictureBox1.Location"), System.Drawing.Point)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.RightToLeft = CType(resources.GetObject("PictureBox1.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.PictureBox1.Size = CType(resources.GetObject("PictureBox1.Size"), System.Drawing.Size)
        Me.PictureBox1.SizeMode = CType(resources.GetObject("PictureBox1.SizeMode"), System.Windows.Forms.PictureBoxSizeMode)
        Me.PictureBox1.TabIndex = CType(resources.GetObject("PictureBox1.TabIndex"), Integer)
        Me.PictureBox1.TabStop = False
        Me.PictureBox1.Text = resources.GetString("PictureBox1.Text")
        Me.PictureBox1.Visible = CType(resources.GetObject("PictureBox1.Visible"), Boolean)
        '
        'frmCAC
        '
        Me.AccessibleDescription = resources.GetString("$this.AccessibleDescription")
        Me.AccessibleName = resources.GetString("$this.AccessibleName")
        Me.AutoScale = False
        Me.AutoScaleBaseSize = CType(resources.GetObject("$this.AutoScaleBaseSize"), System.Drawing.Size)
        Me.AutoScroll = CType(resources.GetObject("$this.AutoScroll"), Boolean)
        Me.AutoScrollMargin = CType(resources.GetObject("$this.AutoScrollMargin"), System.Drawing.Size)
        Me.AutoScrollMinSize = CType(resources.GetObject("$this.AutoScrollMinSize"), System.Drawing.Size)
        Me.BackColor = System.Drawing.Color.LightSlateGray
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = CType(resources.GetObject("$this.ClientSize"), System.Drawing.Size)
        Me.Controls.Add(Me.gbHeader)
        Me.Controls.Add(Me.lblID)
        Me.Controls.Add(Me.statCACBar)
        Me.Controls.Add(Me.gbControl)
        Me.Controls.Add(Me.gbCheque)
        Me.Controls.Add(Me.gbCash)
        Me.Controls.Add(Me.gbCriteria)
        Me.Enabled = CType(resources.GetObject("$this.Enabled"), Boolean)
        Me.Font = CType(resources.GetObject("$this.Font"), System.Drawing.Font)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.ImeMode = CType(resources.GetObject("$this.ImeMode"), System.Windows.Forms.ImeMode)
        Me.Location = CType(resources.GetObject("$this.Location"), System.Drawing.Point)
        Me.MaximumSize = CType(resources.GetObject("$this.MaximumSize"), System.Drawing.Size)
        Me.MinimumSize = CType(resources.GetObject("$this.MinimumSize"), System.Drawing.Size)
        Me.Name = "frmCAC"
        Me.RightToLeft = CType(resources.GetObject("$this.RightToLeft"), System.Windows.Forms.RightToLeft)
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = CType(resources.GetObject("$this.StartPosition"), System.Windows.Forms.FormStartPosition)
        Me.Text = resources.GetString("$this.Text")
        Me.gbCriteria.ResumeLayout(False)
        Me.gbCash.ResumeLayout(False)
        Me.gbCheque.ResumeLayout(False)
        CType(Me.dgChequeStat, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgChequeDetails, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbControl.ResumeLayout(False)
        CType(Me.statPanelUser, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.statPanelDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.statPanelTime, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbHeader.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private dtabCash As New DataTable
    Private CallClsCAC As New clsCAC
    Private ValDis As Boolean
    Private ts As DataGridTableStyle
    Private decGrCshAmt As Decimal = 0
    Private decGrChgAmt As Decimal = 0
    Private decExCheque As Decimal = 0
    Private decAmCheque As Decimal = 0
    Private decGrandTotal As Decimal = 0


    Private Sub frmCAC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetStatusBar()
        txtTellerID.Text = zCurrentUser()
        cmbTransType.SelectedItem = "All Transaction"
        cmbTransType.Focus()
    End Sub

    Private Sub SetStatusBar()
        statPanelDate.Text = CType(FormatDateTime(Today(), DateFormat.LongDate), String) & " "
        statPanelTime.Text = CType(TimeValue(Now()), String) & " "
        statPanelUser.Text = " User Name : " & UCase(zCurrentUser())
    End Sub

#Region "CASH COLLECTION"

    Private Sub TotalCash()
        Dim decTotCash As Decimal = (CType(txtTot1000.Text, Decimal)) + (CType(txtTot500.Text, Decimal)) + (CType(txtTot200.Text, Decimal)) + (CType(txtTot100.Text, Decimal)) + (CType(txtTot50.Text, Decimal)) + (CType(txtTot20.Text, Decimal)) + (CType(txtTot10.Text, Decimal)) + (CType(txtTot5.Text, Decimal)) + (CType(txtTot1.Text, Decimal)) + (CType(txtTot025.Text, Decimal)) + (CType(txtTot010.Text, Decimal)) + (CType(txtTot005.Text, Decimal)) + (CType(txtTot001.Text, Decimal))
        txtTotCash.Text = FormatNumber(CType(decTotCash, String), 2)
        txtAmtLeft.Text = FormatNumber(CType(((decGrCshAmt - decGrChgAmt) - decTotCash), String), 2)
    End Sub

    Private Sub txt1000_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1000.LostFocus
        ValDis = CallClsCAC.NumVal(txt1000.Text)
        If ValDis = True Then
            txtTot1000.Text = CType(FormatNumber((CType(txt1000.Text, Decimal) * 1000), 2), String)
            TotalCash()
        Else
            MsgBox("Pls. Input a Numeric Value", MsgBoxStyle.Exclamation, "Invalid")
            txt1000.Text = 0
            txt1000.Focus()
            txtTot1000.Text = "0.00"
        End If
    End Sub

    Private Sub txt500_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt500.LostFocus
        ValDis = CallClsCAC.NumVal(txt500.Text)
        If ValDis = True Then
            txtTot500.Text = CType(FormatNumber((CType(txt500.Text, Decimal) * 500), 2), String)
            TotalCash()
        Else
            MsgBox("Pls. Input a Numeric Value", MsgBoxStyle.Exclamation, "Invalid")
            txt500.Text = 0
            txt500.Focus()
            txtTot500.Text = "0.00"
        End If
    End Sub

    Private Sub txt200_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt200.LostFocus
        ValDis = CallClsCAC.NumVal(txt200.Text)
        If ValDis = True Then
            txtTot200.Text = CType(FormatNumber((CType(txt200.Text, Decimal) * 200), 2), String)
            TotalCash()
        Else
            MsgBox("Pls. Input a Numeric Value", MsgBoxStyle.Exclamation, "Invalid")
            txt200.Text = 0
            txt200.Focus()
            txtTot200.Text = "0.00"
        End If
    End Sub

    Private Sub txt100_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt100.LostFocus
        ValDis = CallClsCAC.NumVal(txt100.Text)
        If ValDis = True Then
            txtTot100.Text = CType(FormatNumber((CType(txt100.Text, Decimal) * 100), 2), String)
            TotalCash()
        Else
            MsgBox("Pls. Input a Numeric Value", MsgBoxStyle.Exclamation, "Invalid")
            txt100.Text = 0
            txt100.Focus()
            txtTot100.Text = "0.00"
        End If
    End Sub

    Private Sub txt50_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt50.LostFocus
        ValDis = CallClsCAC.NumVal(txt50.Text)
        If ValDis = True Then
            txtTot50.Text = CType(FormatNumber((CType(txt50.Text, Decimal) * 50), 2), String)
            TotalCash()
        Else
            MsgBox("Pls. Input a Numeric Value", MsgBoxStyle.Exclamation, "Invalid")
            txt50.Text = 0
            txt50.Focus()
            txtTot50.Text = "0.00"
        End If
    End Sub

    Private Sub txt10_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt10.LostFocus
        ValDis = CallClsCAC.NumVal(txt10.Text)
        If ValDis = True Then
            txtTot10.Text = CType(FormatNumber((CType(txt10.Text, Decimal) * 10), 2), String)
            TotalCash()
        Else
            MsgBox("Pls. Input a Numeric Value", MsgBoxStyle.Exclamation, "Invalid")
            txt10.Text = 0
            txt10.Focus()
            txtTot10.Text = "0.00"
        End If
    End Sub

    Private Sub txt20_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt20.LostFocus
        ValDis = CallClsCAC.NumVal(txt20.Text)
        If ValDis = True Then
            txtTot20.Text = CType(FormatNumber((CType(txt20.Text, Decimal) * 20), 2), String)
            TotalCash()
        Else
            MsgBox("Pls. Input a Numeric Value", MsgBoxStyle.Exclamation, "Invalid")
            txt20.Text = 0
            txt20.Focus()
            txtTot20.Text = "0.00"
        End If
    End Sub

    Private Sub txt5_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt5.LostFocus
        ValDis = CallClsCAC.NumVal(txt5.Text)
        If ValDis = True Then
            txtTot5.Text = CType(FormatNumber((CType(txt5.Text, Decimal) * 5), 2), String)
            TotalCash()
        Else
            MsgBox("Pls. Input a Numeric Value", MsgBoxStyle.Exclamation, "Invalid")
            txt5.Text = 0
            txt5.Focus()
            txtTot5.Text = "0.00"
        End If
    End Sub

    Private Sub txt1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt1.LostFocus
        ValDis = CallClsCAC.NumVal(txt1.Text)
        If ValDis = True Then
            txtTot1.Text = CType(FormatNumber((CType(txt1.Text, Decimal) * 1), 2), String)
            TotalCash()
        Else
            MsgBox("Pls. Input a Numeric Value", MsgBoxStyle.Exclamation, "Invalid")
            txt1.Text = 0
            txt1.Focus()
            txtTot1.Text = "0.00"
        End If
    End Sub

    Private Sub txt025_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt025.LostFocus
        ValDis = CallClsCAC.NumVal(txt025.Text)
        If ValDis = True Then
            txtTot025.Text = CType(FormatNumber((CType(txt025.Text, Decimal) * 0.25), 2), String)
            TotalCash()
        Else
            MsgBox("Pls. Input a Numeric Value", MsgBoxStyle.Exclamation, "Invalid")
            txt025.Text = 0
            txt025.Focus()
            txtTot025.Text = "0.00"
        End If
    End Sub

    Private Sub txt010_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt010.LostFocus
        ValDis = CallClsCAC.NumVal(txt010.Text)
        If ValDis = True Then
            txtTot010.Text = CType(FormatNumber((CType(txt010.Text, Decimal) * 0.1), 2), String)
            TotalCash()
        Else
            MsgBox("Pls. Input a Numeric Value", MsgBoxStyle.Exclamation, "Invalid")
            txt010.Text = 0
            txt010.Focus()
            txtTot010.Text = "0.00"
        End If
    End Sub

    Private Sub txt005_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt005.LostFocus
        ValDis = CallClsCAC.NumVal(txt005.Text)
        If ValDis = True Then
            txtTot005.Text = CType(FormatNumber((CType(txt005.Text, Decimal) * 0.05), 2), String)
            TotalCash()
        Else
            MsgBox("Pls. Input a Numeric Value", MsgBoxStyle.Exclamation, "Invalid")
            txt005.Text = 0
            txt005.Focus()
            txtTot005.Text = "0.00"
        End If
    End Sub

    Private Sub txt001_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txt001.LostFocus
        ValDis = CallClsCAC.NumVal(txt001.Text)
        If ValDis = True Then
            txtTot001.Text = CType(FormatNumber((CType(txt001.Text, Decimal) * 0.01), 2), String)
            TotalCash()
        Else
            MsgBox("Pls. Input a Numeric Value", MsgBoxStyle.Exclamation, "Invalid")
            txt001.Text = 0
            txt001.Focus()
            txtTot001.Text = "0.00"
        End If
    End Sub

    Private Sub txt1000_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt1000.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txt500.Focus()
        End If
    End Sub

    Private Sub txt20_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt20.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txt10.Focus()
        End If
    End Sub

    Private Sub txt001_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt001.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txtRemarks.Focus()
        End If
    End Sub

    Private Sub txt005_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt005.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txt001.Focus()
        End If
    End Sub

    Private Sub txt010_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt010.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txt005.Focus()
        End If
    End Sub

    Private Sub txt025_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt025.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txt010.Focus()
        End If
    End Sub

    Private Sub txt1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt1.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txt025.Focus()
        End If
    End Sub

    Private Sub txt10_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt10.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txt5.Focus()
        End If
    End Sub

    Private Sub txt100_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt100.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txt50.Focus()
        End If
    End Sub

    Private Sub txt200_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt200.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txt100.Focus()
        End If
    End Sub

    Private Sub txt5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt5.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txt1.Focus()
        End If
    End Sub

    Private Sub txt50_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt50.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txt20.Focus()
        End If
    End Sub

    Private Sub txt500_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt500.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txt200.Focus()
        End If
    End Sub

#End Region

#Region "CRITERIA"

    Private Sub cmbTransType_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbTransType.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            ValDis = CallClsCAC.StrVal(cmbTransType.Text)
            If ValDis = False Then
                MsgBox("Invalid Transaction Type", MsgBoxStyle.Critical, "Invalid")
                cmbTransType.Focus()
            Else
                txtTellerID.Focus()
            End If
        End If
    End Sub

    Private Sub txtTellerID_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTellerID.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            ValDis = CallClsCAC.StrVal(txtTellerID.Text)
            If ValDis = False Then
                MsgBox("Invalid Teller ID", MsgBoxStyle.Critical)
                txtTellerID.Focus()
            Else
                dtePeriod.Focus()
            End If
        End If
    End Sub

    Private Sub dtePeriod_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtePeriod.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            If IsDate(dtePeriod.Text) = False Then
                MsgBox("Invalid Date", MsgBoxStyle.Critical, "Invalid")
                dtePeriod.Focus()
            Else
                CACExisting()
                If dsTurnOverSlip.Tables(0).Rows.Count > 0 Then
                    PopulatelstTimeRange()
                    lblTimeTo.Visible = False
                    lblTimeFrom.Visible = False
                    lblTimeRange.Visible = True
                    lstTimeRange.Visible = True
                    lstTimeRange.Focus()
                Else
                    MsgBox("There are no Time Ranges Saved for this Submission Date.", MsgBoxStyle.Information, "")
                    txtTimeFrom.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub txtTimeFrom_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTimeFrom.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            txtTimeTo.Focus()
        End If
    End Sub

    Private Sub txtTimeTo_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTimeTo.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            ValDis = CallClsCAC.TimeVal(txtTimeFrom.Text, txtTimeTo.Text)
            If ValDis = True Then
                ValDis = ValgbCriteria()
                If ValDis = True Then
                    If lblBatch.Text <> "" Then
                        Dim strResponse As String
                        strResponse = MsgBox("This particular Time range has already been used by Teller " & UCase(txtTellerID.Text) & _
                                             ". If you save this slip, the Data will be overwritten. Proceed Anyway?", MsgBoxStyle.OKCancel, "WARNING")
                        If strResponse = vbOK Then
                            txt1000.Focus()
                            RetrieveData()
                        ElseIf strResponse = vbCancel Then
                            ClearForm()
                        End If
                    Else
                        txt1000.Focus()
                        RetrieveData()
                    End If
                End If
            Else
                MsgBox("Invalid Time Range", MsgBoxStyle.Critical, "Invalid")
                txtTimeFrom.Focus()
                txtTimeFrom.Text = ""
                txtTimeTo.Text = ""
            End If
        End If
    End Sub

    Private Sub lstTimeRange_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lstTimeRange.KeyDown
        If e.KeyCode = 13 Or e.KeyCode = 9 Then
            If IsNothing(lstTimeRange.SelectedItem) Then
                MsgBox("Invalid Time Range", MsgBoxStyle.Critical, "Invalid")
            Else
                txtTimeFrom.Focus()
                cmbTransType.Enabled = False
                cmbTransType.BackColor = Color.AliceBlue
                txtTellerID.ReadOnly = True
                txtTellerID.BackColor = Color.AliceBlue
                dtePeriod.Enabled = False
                dtePeriod.BackColor = Color.AliceBlue
                PopulateFields()
                RetrieveData()
                TotalCash()
                txtTimeFrom.ReadOnly = False
                txtTimeFrom.BackColor = Color.White
                txtTimeTo.ReadOnly = False
                txtTimeTo.BackColor = Color.White
            End If
        ElseIf e.KeyCode = 27 Then
            txtTimeFrom.Focus()
            lblTimeTo.Visible = True
            lblTimeFrom.Visible = True
            lblTimeRange.Visible = False
            lstTimeRange.Visible = False
        End If
    End Sub

    Private Function ValgbCriteria() As Boolean
        If cmbTransType.Text <> "" And txtTellerID.Text <> "" And IsDate(dtePeriod.Text) = True Then
            Return True
        Else
            MsgBox("Pls. Complete the Required Fields.", MsgBoxStyle.Critical)
            If cmbTransType.Text = "" Then
                cmbTransType.Focus()
            ElseIf txtTellerID.Text = "" Then
                txtTellerID.Focus()
            ElseIf IsDate(dtePeriod.Text) = False Then
                dtePeriod.Focus()
            End If
            Return False
        End If
    End Function

    Private Sub txtTellerID_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTellerID.Click
        cmbTransType.Focus()
    End Sub

    Private Sub txtTimeFrom_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTimeFrom.Click
        cmbTransType.Focus()
    End Sub

    Private Sub txtTimeTo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTimeTo.Click
        cmbTransType.Focus()
    End Sub

#End Region

    Private Sub PopulateFields()
        Dim objListItem As ListItem
        objListItem = CType(lstTimeRange.SelectedItem, ListItem)

        Dim drTurnOver As DataRow = CallClsCAC.GetTurnOver(objListItem.ID)
        If Not IsNothing(drTurnOver) Then
            lblID.Text = drTurnOver.Item("ID") 'HIDDEN VALUE TURNOVERSLIP RECORD FIELD "ID"
            lblBatch.Text = "BATCH NUMBER " & drTurnOver.Item("Batch")
            txt1000.Text = drTurnOver.Item("P1000")
            txt500.Text = drTurnOver.Item("P500")
            txt200.Text = drTurnOver.Item("P200")
            txt100.Text = drTurnOver.Item("P100")
            txt50.Text = drTurnOver.Item("P50")
            txt20.Text = drTurnOver.Item("P20")
            txt10.Text = drTurnOver.Item("P10")
            txt5.Text = drTurnOver.Item("P5")
            txt1.Text = drTurnOver.Item("P1")
            txt025.Text = drTurnOver.Item("P025")
            txt010.Text = drTurnOver.Item("P010")
            txt005.Text = drTurnOver.Item("P005")
            txt001.Text = drTurnOver.Item("P001")
            txtTot1000.Text = FormatNumber(CType((CType(txt1000.Text, Decimal) * 1000), String), 2)
            txtTot500.Text = FormatNumber(CType((CType(txt500.Text, Decimal) * 500), String), 2)
            txtTot200.Text = FormatNumber(CType((CType(txt200.Text, Decimal) * 200), String), 2)
            txtTot100.Text = FormatNumber(CType((CType(txt100.Text, Decimal) * 100), String), 2)
            txtTot50.Text = FormatNumber(CType((CType(txt50.Text, Decimal) * 50), String), 2)
            txtTot20.Text = FormatNumber(CType((CType(txt20.Text, Decimal) * 20), String), 2)
            txtTot10.Text = FormatNumber(CType((CType(txt10.Text, Decimal) * 10), String), 2)
            txtTot5.Text = FormatNumber(CType((CType(txt5.Text, Decimal) * 5), String), 2)
            txtTot1.Text = FormatNumber(CType((CType(txt1.Text, Decimal) * 1), String), 2)
            txtTot025.Text = FormatNumber(CType((CType(txt025.Text, Decimal) * 0.25), String), 2)
            txtTot010.Text = FormatNumber(CType((CType(txt010.Text, Decimal) * 0.1), String), 2)
            txtTot005.Text = FormatNumber(CType((CType(txt005.Text, Decimal) * 0.05), String), 2)
            txtTot001.Text = FormatNumber(CType((CType(txt001.Text, Decimal) * 0.01), String), 2)
            txtTimeFrom.Text = drTurnOver.Item("TimeFrom")
            txtTimeTo.Text = drTurnOver.Item("TimeTo")
            lblTimeTo.Visible = True
            lblTimeFrom.Visible = True
            lblTimeRange.Visible = False
            lstTimeRange.Visible = False
        End If
    End Sub

    Private Sub CACExisting()
        Dim strSQL As String
        Dim TransTyp As String
        Dim intBatch As String

        Select Case cmbTransType.SelectedItem
            Case "All Transaction"
                TransTyp = "ALL"
            Case "Import"
                TransTyp = "CYM"
            Case "Export"
                TransTyp = "CYX"
            Case "Special Services"
                TransTyp = "CYS"
            Case "Invoice"
                TransTyp = "INV"
        End Select

        strSQL = "SELECT * FROM TurnOverSlip WHERE TransTyp = '" & TransTyp & "' AND TellerID = '" & txtTellerID.Text & "'" & " AND BatchDate = '" & dtePeriod.Text & "'"

        CallClsCAC.RetrieveTurnOverSlip(strSQL)
    End Sub

    Private Sub PopulatelstTimeRange()
        Dim dv As DataView
        Dim drv As DataRowView
        Dim objListItem As ListItem
        Dim strTimeRange As String

        dv = dsTurnOverSlip.Tables(0).DefaultView
        lstTimeRange.Items.Clear()
        For Each drv In dv
            strTimeRange = drv("TimeFrom") & " - " & drv("TimeTo")
            objListItem = New ListItem(strTimeRange, drv("ID"))
            lstTimeRange.Items.Add(objListItem)
        Next
        lstTimeRange.SetSelected(0, True)
    End Sub

    Private Sub RetrieveData()
        Dim dteTo As Date = CType(dtePeriod.Text & " " & txtTimeTo.Text, Date)
        Dim dteFrom As Date = CType(dtePeriod.Text & " " & txtTimeFrom.Text, Date)
        Dim strSQL As String
        Select Case cmbTransType.SelectedItem
            Case "All Transaction"
                strSQL = "SELECT cshamt,chgamt,chkbnk1,chkbnk2,chkbnk3,chkbnk4,chkbnk5,chkamt1,chkamt2," & _
                         "chkamt3,chkamt4,chkamt5,chkno1,chkno2,chkno3,chkno4,chkno5 FROM CYMPay WHERE status <> 'CAN' AND UPPER(userid) = " & _
                         UCase(CallClsCAC.getToString(txtTellerID.Text)) & " AND sysdttm  >= CAST('" & CType(FormatDateTime(dteFrom, DateFormat.GeneralDate), String) & _
                         "' AS SMALLDATETIME) AND sysdttm  <= CAST('" & CType(FormatDateTime(dteTo, DateFormat.GeneralDate), String) & "' AS SMALLDATETIME)"
                CallClsCAC.RetrieveCAC(strSQL)
                PopulatedgChequeALL("Import")

                strSQL = "SELECT cshamt,chgamt,chkbnk1,chkbnk2,chkbnk3,chkbnk4,chkbnk5,chkamt1,chkamt2," & _
                         "chkamt3,chkamt4,chkamt5,chkno1,chkno2,chkno3,chkno4,chkno5 FROM CCRpay WHERE ccrtyp = '1' AND status <> 'CAN' AND UPPER(userid) = " & _
                         UCase(CallClsCAC.getToString(txtTellerID.Text)) & " AND sysdttm  >= CAST('" & CType(FormatDateTime(dteFrom, DateFormat.GeneralDate), String) & _
                         "' AS SMALLDATETIME) AND sysdttm  <= CAST('" & CType(FormatDateTime(dteTo, DateFormat.GeneralDate), String) & "' AS SMALLDATETIME)"
                CallClsCAC.RetrieveCAC(strSQL)
                PopulatedgChequeALL("Export")

                strSQL = "SELECT cshamt,chgamt,chkbnk1,chkbnk2,chkbnk3,chkbnk4,chkbnk5,chkamt1,chkamt2," & _
                         "chkamt3,chkamt4,chkamt5,chkno1,chkno2,chkno3,chkno4,chkno5 FROM CCRpay AS PAY INNER JOIN CCRdtl AS DTL ON PAY.refnum = DTL.refnum " & _
                         "WHERE DTL.guarntycde <> 'Y' AND PAY.ccrtyp = '2' AND PAY.status <> 'CAN' AND UPPER(PAY.userid) = " & _
                         UCase(CallClsCAC.getToString(txtTellerID.Text)) & " AND PAY.sysdttm  >= CAST('" & CType(FormatDateTime(dteFrom, DateFormat.GeneralDate), String) & _
                         "' AS SMALLDATETIME) AND PAY.sysdttm  <= CAST('" & CType(FormatDateTime(dteTo, DateFormat.GeneralDate), String) & "' AS SMALLDATETIME) " & _
                         "GROUP BY cshamt,chgamt,chkbnk1,chkbnk2,chkbnk3,chkbnk4,chkbnk5,chkamt1,chkamt2,chkamt3,chkamt4,chkamt5,chkno1,chkno2,chkno3,chkno4,chkno5"
                CallClsCAC.RetrieveCAC(strSQL)
                PopulatedgChequeALL("Special Services")

                strSQL = "SELECT CashAMT,AvailAMT,CheckAmt1,CheckAmt2,CheckBnk1,CheckBnk2,CheckNo1,CheckNo2" & _
                         " FROM INVPAYHDR as INV inner join INVPAYDTL as PAY on INV.ORNUM = PAY.ORNUM inner join INVICT as ICT on PAY.INVNUM = ICT.INVNUM WHERE ICT.status <>'CAN'" & _
                         " AND UPPER(INV.userid) = " & UCase(CallClsCAC.getToString(txtTellerID.Text)) & " AND ORDate  >= CAST('" & CType(FormatDateTime(dteFrom, DateFormat.GeneralDate), String) & _
                         "' AS SMALLDATETIME) AND ORDate  <= CAST('" & CType(FormatDateTime(dteTo, DateFormat.GeneralDate), String) & "' AS SMALLDATETIME)" & _
                         " GROUP BY CashAMT,AvailAMT,CheckAmt1,CheckAmt2,CheckBnk1,CheckBnk2,CheckNo1,CheckNo2"
                CallClsCAC.RetrieveCAC(strSQL)
                PopulatedgChequeALL("Invoice")

                If dtabDetails.Rows.Count = 0 Then
                    ClearForm()
                    MsgBox("No Records Found", MsgBoxStyle.Information)
                Else
                    PopulatedgChequeALL2()
                End If

            Case "Import"
                strSQL = "SELECT cshamt,chgamt,chkbnk1,chkbnk2,chkbnk3,chkbnk4,chkbnk5,chkamt1,chkamt2," & _
                         "chkamt3,chkamt4,chkamt5,chkno1,chkno2,chkno3,chkno4,chkno5 FROM CYMPay WHERE status <> 'CAN' AND UPPER(userid) = " & _
                         UCase(CallClsCAC.getToString(txtTellerID.Text)) & " AND sysdttm  >= CAST('" & CType(FormatDateTime(dteFrom, DateFormat.GeneralDate), String) & _
                         "' AS SMALLDATETIME) AND sysdttm  <= CAST('" & CType(FormatDateTime(dteTo, DateFormat.GeneralDate), String) & "' AS SMALLDATETIME)"
                CallClsCAC.RetrieveCAC(strSQL)
                If dsCAC.Tables(0).Rows.Count > 0 Then
                    PopulatedgChequeMXS()
                    DisabledgbCriteria()
                Else
                    ClearForm()
                    MsgBox("No Records Found", MsgBoxStyle.Information)
                End If
            Case "Export"
                strSQL = "SELECT cshamt,chgamt,chkbnk1,chkbnk2,chkbnk3,chkbnk4,chkbnk5,chkamt1,chkamt2," & _
                         "chkamt3,chkamt4,chkamt5,chkno1,chkno2,chkno3,chkno4,chkno5 FROM CCRpay WHERE ccrtyp = '1' AND status <> 'CAN' AND UPPER(userid) = " & _
                         UCase(CallClsCAC.getToString(txtTellerID.Text)) & " AND sysdttm  >= CAST('" & CType(FormatDateTime(dteFrom, DateFormat.GeneralDate), String) & _
                         "' AS SMALLDATETIME) AND sysdttm  <= CAST('" & CType(FormatDateTime(dteTo, DateFormat.GeneralDate), String) & "' AS SMALLDATETIME)"
                CallClsCAC.RetrieveCAC(strSQL)
                If dsCAC.Tables(0).Rows.Count > 0 Then
                    PopulatedgChequeMXS()
                    DisabledgbCriteria()
                Else
                    ClearForm()
                    MsgBox("No Records Found", MsgBoxStyle.Information)
                End If
            Case "Special Services"
                strSQL = "SELECT cshamt,chgamt,chkbnk1,chkbnk2,chkbnk3,chkbnk4,chkbnk5,chkamt1,chkamt2," & _
                         "chkamt3,chkamt4,chkamt5,chkno1,chkno2,chkno3,chkno4,chkno5 FROM CCRpay AS PAY INNER JOIN CCRdtl AS DTL ON PAY.refnum = DTL.refnum " & _
                         "WHERE DTL.guarntycde <> 'Y' AND PAY.ccrtyp = '2' AND PAY.status <> 'CAN' AND UPPER(PAY.userid) = " & _
                         UCase(CallClsCAC.getToString(txtTellerID.Text)) & " AND PAY.sysdttm  >= CAST('" & CType(FormatDateTime(dteFrom, DateFormat.GeneralDate), String) & _
                         "' AS SMALLDATETIME) AND PAY.sysdttm  <= CAST('" & CType(FormatDateTime(dteTo, DateFormat.GeneralDate), String) & "' AS SMALLDATETIME) " & _
                         "GROUP BY cshamt,chgamt,chkbnk1,chkbnk2,chkbnk3,chkbnk4,chkbnk5,chkamt1,chkamt2,chkamt3,chkamt4,chkamt5,chkno1,chkno2,chkno3,chkno4,chkno5"
                CallClsCAC.RetrieveCAC(strSQL)
                If dsCAC.Tables(0).Rows.Count > 0 Then
                    PopulatedgChequeMXS()
                    DisabledgbCriteria()
                Else
                    ClearForm()
                    MsgBox("No Records Found", MsgBoxStyle.Information)
                End If
            Case "Invoice"
                strSQL = "SELECT CashAMT,AvailAMT,CheckAmt1,CheckAmt2,CheckBnk1,CheckBnk2,CheckNo1,CheckNo2" & _
                         " FROM INVPAYHDR as INV inner join INVPAYDTL as PAY on INV.ORNUM = PAY.ORNUM inner join INVICT as ICT on PAY.INVNUM = ICT.INVNUM WHERE ICT.status <>'CAN'" & _
                         " AND UPPER(INV.userid) = " & UCase(CallClsCAC.getToString(txtTellerID.Text)) & " AND ORDate  >= CAST('" & CType(FormatDateTime(dteFrom, DateFormat.GeneralDate), String) & _
                         "' AS SMALLDATETIME) AND ORDate  <= CAST('" & CType(FormatDateTime(dteTo, DateFormat.GeneralDate), String) & "' AS SMALLDATETIME)" & _
                         " GROUP BY CashAMT,AvailAMT,CheckAmt1,CheckAmt2,CheckBnk1,CheckBnk2,CheckNo1,CheckNo2"
                CallClsCAC.RetrieveCAC(strSQL)
                If dsCAC.Tables(0).Rows.Count > 0 Then
                    PopulatedgChequeINV()
                    DisabledgbCriteria()
                Else
                    ClearForm()
                    MsgBox("No Records Found", MsgBoxStyle.Information)
                End If
        End Select
    End Sub

    Private Sub PopulatedgChequeMXS()
        'Populate datatable
        dtabDetails = New DataTable

        If dtabDetails.Columns.Contains("Bank") = True Then
            dtabDetails.Columns.Remove("Bank")
        End If
        dtabDetails.Columns.Add("Bank", Type.GetType("System.String"))

        If dtabDetails.Columns.Contains("Cheque No.") = True Then
            dtabDetails.Columns.Remove("Cheque No.")
        End If
        dtabDetails.Columns.Add("Cheque No.", Type.GetType("System.String"))

        If dtabDetails.Columns.Contains("Amount") = True Then
            dtabDetails.Columns.Remove("Amount")
        End If
        dtabDetails.Columns.Add("Amount", Type.GetType("System.Decimal"))

        If dtabDetails.Columns.Contains("Excess") = True Then
            dtabDetails.Columns.Remove("Excess")
        End If
        dtabDetails.Columns.Add("Excess", Type.GetType("System.Decimal"))


        Dim dtabdsCAC As New DataTable
        dtabdsCAC = dsCAC.Tables(0)
        decExCheque = 0
        decAmCheque = 0
        decGrandTotal = 0
        decGrCshAmt = 0
        decGrChgAmt = 0

        If dtabdsCAC.Rows.Count > 0 Then
            Dim lngCtr As Long = 0

            Do While lngCtr < dtabdsCAC.Rows.Count
                '------- Expected Grand Total (add the cash amount)
                decGrCshAmt = decGrCshAmt + dtabdsCAC.Rows(lngCtr)("cshamt")
                decGrChgAmt = decGrChgAmt + dtabdsCAC.Rows(lngCtr)("chgamt")

                '------- POPULATING dgChequeDetails
                If dtabdsCAC.Rows(lngCtr)("chkamt1") <> 0 Then
                    Dim dtarow As DataRow
                    dtarow = dtabDetails.NewRow
                    dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk1")
                    dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno1")
                    dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt1")
                    decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt1")
                    If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 Then
                        dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                        decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                    Else
                        dtarow("Excess") = 0
                    End If
                    dtabDetails.Rows.Add(dtarow)
                End If

                If dtabdsCAC.Rows(lngCtr)("chkamt2") <> 0 Then
                    Dim dtarow As DataRow
                    dtarow = dtabDetails.NewRow
                    dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk2")
                    dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno2")
                    dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt2")
                    decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt2")
                    If dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                        dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                        decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                    Else
                        dtarow("Excess") = 0
                    End If
                    dtabDetails.Rows.Add(dtarow)
                End If

                If dtabdsCAC.Rows(lngCtr)("chkamt3") <> 0 Then
                    Dim dtarow As DataRow
                    dtarow = dtabDetails.NewRow
                    dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk3")
                    dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno3")
                    dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt3")
                    decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt3")
                    If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 And dtabdsCAC.Rows(lngCtr)("chkbnk2") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                        dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                        decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                    Else
                        dtarow("Excess") = 0
                    End If
                    dtabDetails.Rows.Add(dtarow)
                End If

                If dtabdsCAC.Rows(lngCtr)("chkamt4") <> 0 Then
                    Dim dtarow As DataRow
                    dtarow = dtabDetails.NewRow
                    dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk4")
                    dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno4")
                    dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt4")
                    decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt4")
                    If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 And dtabdsCAC.Rows(lngCtr)("chkbnk3") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk2") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                        dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                        decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                    Else
                        dtarow("Excess") = 0
                    End If
                    dtabDetails.Rows.Add(dtarow)
                End If

                If dtabdsCAC.Rows(lngCtr)("chkamt5") <> 0 Then
                    Dim dtarow As DataRow
                    dtarow = dtabDetails.NewRow
                    dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk5")
                    dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno5")
                    dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt5")
                    decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt5")
                    If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 And dtabdsCAC.Rows(lngCtr)("chkbnk4") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk3") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk2") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                        dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                        decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                    Else
                        dtarow("Excess") = 0
                    End If
                    dtabDetails.Rows.Add(dtarow)
                End If

                '------- Expected Grand Total (add cheque amount)
                lngCtr += 1
            Loop
        End If

        ts = New DataGridTableStyle
        ts.MappingName = dtabDetails.ToString
        dgChequeDetails.TableStyles.Add(ts)

        With dgChequeDetails
            .DataSource = dtabDetails
            .AlternatingBackColor = Color.AliceBlue
            .BackColor = Color.White
            .TableStyles(0).RowHeaderWidth = 15
            .TableStyles(0).GridColumnStyles.Item(0).Width = 75
            .TableStyles(0).GridColumnStyles.Item(1).Width = 106
            .TableStyles(0).GridColumnStyles.Item(2).Width = 110
            .TableStyles(0).GridColumnStyles.Item(3).Width = 110
            .TableStyles(0).SelectionBackColor = .TableStyles(0).BackColor.PowderBlue
            .TableStyles(0).SelectionForeColor = .TableStyles(0).ForeColor.Black
            .TableStyles(0).AlternatingBackColor = .TableStyles(0).BackColor.AliceBlue
            .TableStyles(0).BackColor = .TableStyles(0).BackColor.White
            .TableStyles(0).HeaderBackColor = .TableStyles(0).BackColor.MidnightBlue
            .TableStyles(0).HeaderForeColor = .TableStyles(0).ForeColor.White
            .TableStyles(0).HeaderFont = New System.Drawing.Font("Tahoma", 8.0F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, System.Byte))
        End With


        '------- POPULATING dgChequeStat
        Dim dtabStat As New DataTable

        If dtabStat.Columns.Contains("Bank") = True Then
            dtabStat.Columns.Remove("Bank")
        End If
        dtabStat.Columns.Add("Bank", Type.GetType("System.String"))

        If dtabStat.Columns.Contains("Pieces") = True Then
            dtabStat.Columns.Remove("Pieces")
        End If
        dtabStat.Columns.Add("Pieces", Type.GetType("System.Int32"))

        If dtabStat.Columns.Contains("Amount") = True Then
            dtabStat.Columns.Remove("Amount")
        End If
        dtabStat.Columns.Add("Amount", Type.GetType("System.Decimal"))



        If dtabDetails.Rows.Count > 0 Then
            Dim dv As DataView
            Dim drv As DataRowView

            dv = dtabDetails.DefaultView
            dv.Sort = "Bank"

            Dim strPreBank As String = ""
            Dim decAmt As Decimal = 0
            Dim intPcs As Integer = 0

            For Each drv In dv
                If Trim(drv("Bank")) = Trim(strPreBank) Then
                    decAmt = decAmt + drv("Amount")
                    intPcs += 1
                Else
                    If strPreBank <> "" Then
                        Dim dtarow As DataRow
                        dtarow = dtabStat.NewRow
                        dtarow("Bank") = strPreBank
                        dtarow("Pieces") = intPcs
                        dtarow("Amount") = decAmt
                        dtabStat.Rows.Add(dtarow)
                    End If
                    strPreBank = drv("Bank")
                    decAmt = drv("Amount")
                    intPcs = 1
                End If
            Next
            '---- add final row of dgChequeStat
            Dim dtarow1 As DataRow
            dtarow1 = dtabStat.NewRow
            dtarow1("Bank") = strPreBank
            dtarow1("Pieces") = intPcs
            dtarow1("Amount") = decAmt
            dtabStat.Rows.Add(dtarow1)
        End If

        ts = New DataGridTableStyle
        ts.MappingName = dtabStat.ToString
        dgChequeStat.TableStyles.Add(ts)

        With dgChequeStat
            .DataSource = dtabStat
            .AlternatingBackColor = Color.AliceBlue
            .BackColor = Color.White
            .TableStyles(0).RowHeaderWidth = 15
            .TableStyles(0).GridColumnStyles.Item(0).Width = 75
            .TableStyles(0).GridColumnStyles.Item(1).Width = 106
            .TableStyles(0).GridColumnStyles.Item(2).Width = 220
            .TableStyles(0).SelectionBackColor = .TableStyles(0).BackColor.PowderBlue
            .TableStyles(0).SelectionForeColor = .TableStyles(0).ForeColor.Black
            .TableStyles(0).AlternatingBackColor = .TableStyles(0).BackColor.AliceBlue
            .TableStyles(0).BackColor = .TableStyles(0).BackColor.White
            .TableStyles(0).HeaderBackColor = .TableStyles(0).BackColor.MidnightBlue
            .TableStyles(0).HeaderForeColor = .TableStyles(0).ForeColor.White
            .TableStyles(0).HeaderFont = New System.Drawing.Font("Tahoma", 8.0F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, System.Byte))
        End With


        txtExcessCheque.Text = FormatNumber(CType(decExCheque, String), 2)
        txtTotCheque.Text = FormatNumber(CType(decAmCheque, String), 2)
        txtGrandTot.Text = FormatNumber(CType(((decGrCshAmt + decAmCheque) - decGrChgAmt), String), 2)
        txtExCash.Text = FormatNumber(CType((decGrCshAmt - decGrChgAmt), String), 2)
        txtAmtLeft.Text = FormatNumber(CType((decGrCshAmt - decGrChgAmt), String), 2)

    End Sub

    Private Sub PopulatedgChequeINV()
        'Populate datatable
        dtabDetails = New DataTable

        If dtabDetails.Columns.Contains("Bank") = True Then
            dtabDetails.Columns.Remove("Bank")
        End If
        dtabDetails.Columns.Add("Bank", Type.GetType("System.String"))

        If dtabDetails.Columns.Contains("Cheque No.") = True Then
            dtabDetails.Columns.Remove("Cheque No.")
        End If
        dtabDetails.Columns.Add("Cheque No.", Type.GetType("System.String"))

        If dtabDetails.Columns.Contains("Amount") = True Then
            dtabDetails.Columns.Remove("Amount")
        End If
        dtabDetails.Columns.Add("Amount", Type.GetType("System.Decimal"))

        If dtabDetails.Columns.Contains("Excess") = True Then
            dtabDetails.Columns.Remove("Excess")
        End If
        dtabDetails.Columns.Add("Excess", Type.GetType("System.Decimal"))


        Dim dtabdsCAC As New DataTable
        dtabdsCAC = dsCAC.Tables(0)
        decExCheque = 0
        decAmCheque = 0
        decGrandTotal = 0
        decGrCshAmt = 0
        decGrChgAmt = 0

        If dtabdsCAC.Rows.Count > 0 Then
            Dim lngCtr As Long = 0

            Do While lngCtr < dtabdsCAC.Rows.Count
                '------- Expected Grand Total (add the cash amount)
                decGrCshAmt = decGrCshAmt + dtabdsCAC.Rows(lngCtr)("CashAMT")
                'decGrChgAmt = decGrChgAmt + dtabdsCAC.Rows(lngCtr)("AvailAMT")

                '------- POPULATING dgChequeDetails
                If dtabdsCAC.Rows(lngCtr)("CheckAmt1") <> 0 Then
                    Dim dtarow As DataRow
                    dtarow = dtabDetails.NewRow
                    dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("CheckBnk1")
                    dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("CheckNo1")
                    dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("CheckAmt1")
                    decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("CheckAmt1")
                    'If dtabdsCAC.Rows(lngCtr)("CashAMT") = 0 Then
                    '    dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("AvailAMT")
                    '    decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("AvailAMT")
                    'Else
                    '    dtarow("Excess") = 0
                    'End If
                    dtarow("Excess") = 0
                    dtabDetails.Rows.Add(dtarow)
                End If

                If dtabdsCAC.Rows(lngCtr)("CheckAmt2") <> 0 Then
                    Dim dtarow As DataRow
                    dtarow = dtabDetails.NewRow
                    dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("CheckBnk2")
                    dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("CheckNo2")
                    dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("CheckAmt2")
                    decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("CheckAmt2")
                    'If dtabdsCAC.Rows(lngCtr)("CheckBnk1") = "" And dtabdsCAC.Rows(lngCtr)("CashAMT") = 0 Then
                    '    dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("AvailAMT")
                    '    decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("AvailAMT")
                    'Else
                    '    dtarow("Excess") = 0
                    'End If
                    dtarow("Excess") = 0
                    dtabDetails.Rows.Add(dtarow)
                End If

                '------- Expected Grand Total (add cheque amount)
                lngCtr += 1
            Loop
        End If

        ts = New DataGridTableStyle
        ts.MappingName = dtabDetails.ToString
        dgChequeDetails.TableStyles.Add(ts)

        With dgChequeDetails
            .DataSource = dtabDetails
            .AlternatingBackColor = Color.AliceBlue
            .BackColor = Color.White
            .TableStyles(0).RowHeaderWidth = 15
            .TableStyles(0).GridColumnStyles.Item(0).Width = 75
            .TableStyles(0).GridColumnStyles.Item(1).Width = 106
            .TableStyles(0).GridColumnStyles.Item(2).Width = 110
            .TableStyles(0).GridColumnStyles.Item(3).Width = 110
            .TableStyles(0).SelectionBackColor = .TableStyles(0).BackColor.PowderBlue
            .TableStyles(0).SelectionForeColor = .TableStyles(0).ForeColor.Black
            .TableStyles(0).AlternatingBackColor = .TableStyles(0).BackColor.AliceBlue
            .TableStyles(0).BackColor = .TableStyles(0).BackColor.White
            .TableStyles(0).HeaderBackColor = .TableStyles(0).BackColor.MidnightBlue
            .TableStyles(0).HeaderForeColor = .TableStyles(0).ForeColor.White
            .TableStyles(0).HeaderFont = New System.Drawing.Font("Tahoma", 8.0F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, System.Byte))
        End With


        '------- POPULATING dgChequeStat
        Dim dtabStat As New DataTable

        If dtabStat.Columns.Contains("Bank") = True Then
            dtabStat.Columns.Remove("Bank")
        End If
        dtabStat.Columns.Add("Bank", Type.GetType("System.String"))

        If dtabStat.Columns.Contains("Pieces") = True Then
            dtabStat.Columns.Remove("Pieces")
        End If
        dtabStat.Columns.Add("Pieces", Type.GetType("System.Int32"))

        If dtabStat.Columns.Contains("Amount") = True Then
            dtabStat.Columns.Remove("Amount")
        End If
        dtabStat.Columns.Add("Amount", Type.GetType("System.Decimal"))

        If dtabDetails.Rows.Count > 0 Then
            Dim dv As DataView
            Dim drv As DataRowView

            dv = dtabDetails.DefaultView
            dv.Sort = "Bank"

            Dim strPreBank As String = ""
            Dim decAmt As Decimal = 0
            Dim intPcs As Integer = 0

            For Each drv In dv
                If Trim(drv("Bank")) = Trim(strPreBank) Then
                    decAmt = decAmt + drv("Amount")
                    intPcs += 1
                Else
                    If strPreBank <> "" Then
                        Dim dtarow As DataRow
                        dtarow = dtabStat.NewRow
                        dtarow("Bank") = strPreBank
                        dtarow("Pieces") = intPcs
                        dtarow("Amount") = decAmt
                        dtabStat.Rows.Add(dtarow)
                    End If
                    strPreBank = drv("Bank")
                    decAmt = drv("Amount")
                    intPcs = 1
                End If
            Next
            '---- add final row of dgChequeStat
            Dim dtarow1 As DataRow
            dtarow1 = dtabStat.NewRow
            dtarow1("Bank") = strPreBank
            dtarow1("Pieces") = intPcs
            dtarow1("Amount") = decAmt
            dtabStat.Rows.Add(dtarow1)
        End If

        ts = New DataGridTableStyle
        ts.MappingName = dtabStat.ToString
        dgChequeStat.TableStyles.Add(ts)

        With dgChequeStat
            .DataSource = dtabStat
            .AlternatingBackColor = Color.AliceBlue
            .BackColor = Color.White
            .TableStyles(0).RowHeaderWidth = 15
            .TableStyles(0).GridColumnStyles.Item(0).Width = 75
            .TableStyles(0).GridColumnStyles.Item(1).Width = 106
            .TableStyles(0).GridColumnStyles.Item(2).Width = 220
            .TableStyles(0).SelectionBackColor = .TableStyles(0).BackColor.PowderBlue
            .TableStyles(0).SelectionForeColor = .TableStyles(0).ForeColor.Black
            .TableStyles(0).AlternatingBackColor = .TableStyles(0).BackColor.AliceBlue
            .TableStyles(0).BackColor = .TableStyles(0).BackColor.White
            .TableStyles(0).HeaderBackColor = .TableStyles(0).BackColor.MidnightBlue
            .TableStyles(0).HeaderForeColor = .TableStyles(0).ForeColor.White
            .TableStyles(0).HeaderFont = New System.Drawing.Font("Tahoma", 8.0F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, System.Byte))
        End With


        txtExcessCheque.Text = FormatNumber(CType(decExCheque, String), 2)
        txtTotCheque.Text = FormatNumber(CType(decAmCheque, String), 2)
        txtGrandTot.Text = FormatNumber(CType(((decGrCshAmt + decAmCheque) - decGrChgAmt), String), 2)
        txtExCash.Text = FormatNumber(CType((decGrCshAmt - decGrChgAmt), String), 2)
        txtAmtLeft.Text = FormatNumber(CType((decGrCshAmt - decGrChgAmt), String), 2)
    End Sub

    Private Sub PopulatedgChequeALL(ByVal strTransType)

        Select Case strTransType
            '*****************************************************
            'IMPORT 
            '*****************************************************
            Case "Import"
                dtabDetails = New DataTable
                decExCheque = 0
                decAmCheque = 0
                decGrandTotal = 0
                decGrCshAmt = 0
                decGrChgAmt = 0

                If dtabDetails.Columns.Contains("Bank") = True Then
                    dtabDetails.Columns.Remove("Bank")
                End If
                dtabDetails.Columns.Add("Bank", Type.GetType("System.String"))

                If dtabDetails.Columns.Contains("Cheque No.") = True Then
                    dtabDetails.Columns.Remove("Cheque No.")
                End If
                dtabDetails.Columns.Add("Cheque No.", Type.GetType("System.String"))

                If dtabDetails.Columns.Contains("Amount") = True Then
                    dtabDetails.Columns.Remove("Amount")
                End If
                dtabDetails.Columns.Add("Amount", Type.GetType("System.Decimal"))

                If dtabDetails.Columns.Contains("Excess") = True Then
                    dtabDetails.Columns.Remove("Excess")
                End If
                dtabDetails.Columns.Add("Excess", Type.GetType("System.Decimal"))

                Dim dtabdsCAC As New DataTable
                dtabdsCAC = dsCAC.Tables(0)

                If dtabdsCAC.Rows.Count > 0 Then
                    Dim lngCtr As Long = 0

                    Do While lngCtr < dtabdsCAC.Rows.Count
                        '------- Expected Grand Total (add the cash amount)
                        decGrCshAmt = decGrCshAmt + dtabdsCAC.Rows(lngCtr)("cshamt")
                        decGrChgAmt = decGrChgAmt + dtabdsCAC.Rows(lngCtr)("chgamt")

                        '------- POPULATING dgChequeDetails
                        If dtabdsCAC.Rows(lngCtr)("chkamt1") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk1")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno1")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt1")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt1")
                            If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        If dtabdsCAC.Rows(lngCtr)("chkamt2") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk2")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno2")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt2")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt2")
                            If dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        If dtabdsCAC.Rows(lngCtr)("chkamt3") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk3")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno3")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt3")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt3")
                            If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 And dtabdsCAC.Rows(lngCtr)("chkbnk2") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        If dtabdsCAC.Rows(lngCtr)("chkamt4") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk4")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno4")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt4")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt4")
                            If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 And dtabdsCAC.Rows(lngCtr)("chkbnk3") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk2") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        If dtabdsCAC.Rows(lngCtr)("chkamt5") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk5")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno5")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt5")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt5")
                            If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 And dtabdsCAC.Rows(lngCtr)("chkbnk4") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk3") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk2") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        '------- Expected Grand Total (add cheque amount)
                        lngCtr += 1
                    Loop
                End If

                '*****************************************************
                'EXPORT
                '*****************************************************
            Case "Export"
                Dim dtabdsCAC As New DataTable
                dtabdsCAC = dsCAC.Tables(0)

                If dtabdsCAC.Rows.Count > 0 Then
                    Dim lngCtr As Long = 0

                    Do While lngCtr < dtabdsCAC.Rows.Count
                        '------- Expected Grand Total (add the cash amount)
                        decGrCshAmt = decGrCshAmt + dtabdsCAC.Rows(lngCtr)("cshamt")
                        decGrChgAmt = decGrChgAmt + dtabdsCAC.Rows(lngCtr)("chgamt")

                        '------- POPULATING dgChequeDetails
                        If dtabdsCAC.Rows(lngCtr)("chkamt1") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk1")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno1")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt1")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt1")
                            If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        If dtabdsCAC.Rows(lngCtr)("chkamt2") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk2")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno2")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt2")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt2")
                            If dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        If dtabdsCAC.Rows(lngCtr)("chkamt3") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk3")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno3")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt3")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt3")
                            If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 And dtabdsCAC.Rows(lngCtr)("chkbnk2") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        If dtabdsCAC.Rows(lngCtr)("chkamt4") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk4")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno4")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt4")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt4")
                            If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 And dtabdsCAC.Rows(lngCtr)("chkbnk3") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk2") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        If dtabdsCAC.Rows(lngCtr)("chkamt5") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk5")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno5")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt5")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt5")
                            If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 And dtabdsCAC.Rows(lngCtr)("chkbnk4") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk3") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk2") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        '------- Expected Grand Total (add cheque amount)
                        lngCtr += 1
                    Loop
                End If

                '*****************************************************
                'SPECIAL SERVICES
                '*****************************************************

            Case "Special Services"
                Dim dtabdsCAC As New DataTable
                dtabdsCAC = dsCAC.Tables(0)

                If dtabdsCAC.Rows.Count > 0 Then
                    Dim lngCtr As Long = 0

                    Do While lngCtr < dtabdsCAC.Rows.Count
                        '------- Expected Grand Total (add the cash amount)
                        decGrCshAmt = decGrCshAmt + dtabdsCAC.Rows(lngCtr)("cshamt")
                        decGrChgAmt = decGrChgAmt + dtabdsCAC.Rows(lngCtr)("chgamt")

                        '------- POPULATING dgChequeDetails
                        If dtabdsCAC.Rows(lngCtr)("chkamt1") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk1")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno1")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt1")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt1")
                            If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        If dtabdsCAC.Rows(lngCtr)("chkamt2") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk2")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno2")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt2")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt2")
                            If dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        If dtabdsCAC.Rows(lngCtr)("chkamt3") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk3")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno3")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt3")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt3")
                            If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 And dtabdsCAC.Rows(lngCtr)("chkbnk2") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        If dtabdsCAC.Rows(lngCtr)("chkamt4") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk4")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno4")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt4")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt4")
                            If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 And dtabdsCAC.Rows(lngCtr)("chkbnk3") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk2") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        If dtabdsCAC.Rows(lngCtr)("chkamt5") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("chkbnk5")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("chkno5")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("chkamt5")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("chkamt5")
                            If dtabdsCAC.Rows(lngCtr)("cshamt") = 0 And dtabdsCAC.Rows(lngCtr)("chkbnk4") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk3") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk2") = "" And dtabdsCAC.Rows(lngCtr)("chkbnk1") = "" Then
                                dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("chgamt")
                                decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("chgamt")
                            Else
                                dtarow("Excess") = 0
                            End If
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        '------- Expected Grand Total (add cheque amount)
                        lngCtr += 1
                    Loop
                End If

                '*****************************************************
                'INVOIVE
                '*****************************************************
            Case "Invoice"
                Dim dtabdsCAC As New DataTable
                dtabdsCAC = dsCAC.Tables(0)

                If dtabdsCAC.Rows.Count > 0 Then
                    Dim lngCtr As Long = 0

                    Do While lngCtr < dtabdsCAC.Rows.Count
                        '------- Expected Grand Total (add the cash amount)
                        decGrCshAmt = decGrCshAmt + dtabdsCAC.Rows(lngCtr)("CashAMT")
                        'decGrChgAmt = decGrChgAmt + dtabdsCAC.Rows(lngCtr)("AvailAMT")

                        '------- POPULATING dgChequeDetails
                        If dtabdsCAC.Rows(lngCtr)("CheckAmt1") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("CheckBnk1")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("CheckNo1")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("CheckAmt1")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("CheckAmt1")
                            'If dtabdsCAC.Rows(lngCtr)("CashAMT") = 0 Then
                            '    dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("AvailAMT")
                            '    decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("AvailAMT")
                            'Else
                            '    dtarow("Excess") = 0
                            'End If
                            dtarow("Excess") = 0
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        If dtabdsCAC.Rows(lngCtr)("CheckAmt2") <> 0 Then
                            Dim dtarow As DataRow
                            dtarow = dtabDetails.NewRow
                            dtarow("Bank") = dtabdsCAC.Rows(lngCtr)("CheckBnk2")
                            dtarow("Cheque No.") = dtabdsCAC.Rows(lngCtr)("CheckNo2")
                            dtarow("Amount") = dtabdsCAC.Rows(lngCtr)("CheckAmt2")
                            decAmCheque = decAmCheque + dtabdsCAC.Rows(lngCtr)("CheckAmt2")
                            'If dtabdsCAC.Rows(lngCtr)("CheckBnk1") = "" And dtabdsCAC.Rows(lngCtr)("CashAMT") = 0 Then
                            '    dtarow("Excess") = dtabdsCAC.Rows(lngCtr)("AvailAMT")
                            '    decExCheque = decExCheque + dtabdsCAC.Rows(lngCtr)("AvailAMT")
                            'Else
                            '    dtarow("Excess") = 0
                            'End If
                            dtarow("Excess") = 0
                            dtabDetails.Rows.Add(dtarow)
                        End If

                        '------- Expected Grand Total (add cheque amount)
                        lngCtr += 1
                    Loop
                End If

                ts = New DataGridTableStyle
                ts.MappingName = dtabDetails.ToString
                dgChequeDetails.TableStyles.Add(ts)

                With dgChequeDetails
                    .DataSource = dtabDetails
                    .AlternatingBackColor = Color.AliceBlue
                    .BackColor = Color.White
                    .TableStyles(0).RowHeaderWidth = 15
                    .TableStyles(0).GridColumnStyles.Item(0).Width = 75
                    .TableStyles(0).GridColumnStyles.Item(1).Width = 106
                    .TableStyles(0).GridColumnStyles.Item(2).Width = 110
                    .TableStyles(0).GridColumnStyles.Item(3).Width = 110
                    .TableStyles(0).SelectionBackColor = .TableStyles(0).BackColor.PowderBlue
                    .TableStyles(0).SelectionForeColor = .TableStyles(0).ForeColor.Black
                    .TableStyles(0).AlternatingBackColor = .TableStyles(0).BackColor.AliceBlue
                    .TableStyles(0).BackColor = .TableStyles(0).BackColor.White
                    .TableStyles(0).HeaderBackColor = .TableStyles(0).BackColor.MidnightBlue
                    .TableStyles(0).HeaderForeColor = .TableStyles(0).ForeColor.White
                    .TableStyles(0).HeaderFont = New System.Drawing.Font("Tahoma", 8.0F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, System.Byte))
                End With
        End Select

        txtExcessCheque.Text = FormatNumber(CType(decExCheque, String), 2)
        txtTotCheque.Text = FormatNumber(CType(decAmCheque, String), 2)
        txtGrandTot.Text = FormatNumber(CType(((decGrCshAmt + decAmCheque) - decGrChgAmt), String), 2)
        txtExCash.Text = FormatNumber(CType((decGrCshAmt - decGrChgAmt), String), 2)
        txtAmtLeft.Text = FormatNumber(CType((decGrCshAmt - decGrChgAmt), String), 2)

    End Sub

    Private Sub PopulatedgChequeALL2()
        '------- POPULATING dgChequeStat
        Dim dtabStat As New DataTable

        If dtabStat.Columns.Contains("Bank") = True Then
            dtabStat.Columns.Remove("Bank")
        End If
        dtabStat.Columns.Add("Bank", Type.GetType("System.String"))

        If dtabStat.Columns.Contains("Pieces") = True Then
            dtabStat.Columns.Remove("Pieces")
        End If
        dtabStat.Columns.Add("Pieces", Type.GetType("System.Int32"))

        If dtabStat.Columns.Contains("Amount") = True Then
            dtabStat.Columns.Remove("Amount")
        End If
        dtabStat.Columns.Add("Amount", Type.GetType("System.Decimal"))

        If dtabDetails.Rows.Count > 0 Then
            Dim dv As DataView
            Dim drv As DataRowView

            dv = dtabDetails.DefaultView
            dv.Sort = "Bank"

            Dim strPreBank As String = ""
            Dim decAmt As Decimal = 0
            Dim intPcs As Integer = 0

            For Each drv In dv
                If Trim(drv("Bank")) = Trim(strPreBank) Then
                    decAmt = decAmt + drv("Amount")
                    intPcs += 1
                Else
                    If strPreBank <> "" Then
                        Dim dtarow As DataRow
                        dtarow = dtabStat.NewRow
                        dtarow("Bank") = strPreBank
                        dtarow("Pieces") = intPcs
                        dtarow("Amount") = decAmt
                        dtabStat.Rows.Add(dtarow)
                    End If
                    strPreBank = drv("Bank")
                    decAmt = drv("Amount")
                    intPcs = 1
                End If
            Next
            '---- add final row of dgChequeStat
            Dim dtarow1 As DataRow
            dtarow1 = dtabStat.NewRow
            dtarow1("Bank") = strPreBank
            dtarow1("Pieces") = intPcs
            dtarow1("Amount") = decAmt
            dtabStat.Rows.Add(dtarow1)
        End If

        ts = New DataGridTableStyle
        ts.MappingName = dtabStat.ToString
        dgChequeStat.TableStyles.Add(ts)

        With dgChequeStat
            .DataSource = dtabStat
            .AlternatingBackColor = Color.AliceBlue
            .BackColor = Color.White
            .TableStyles(0).RowHeaderWidth = 15
            .TableStyles(0).GridColumnStyles.Item(0).Width = 75
            .TableStyles(0).GridColumnStyles.Item(1).Width = 106
            .TableStyles(0).GridColumnStyles.Item(2).Width = 220
            .TableStyles(0).SelectionBackColor = .TableStyles(0).BackColor.PowderBlue
            .TableStyles(0).SelectionForeColor = .TableStyles(0).ForeColor.Black
            .TableStyles(0).AlternatingBackColor = .TableStyles(0).BackColor.AliceBlue
            .TableStyles(0).BackColor = .TableStyles(0).BackColor.White
            .TableStyles(0).HeaderBackColor = .TableStyles(0).BackColor.MidnightBlue
            .TableStyles(0).HeaderForeColor = .TableStyles(0).ForeColor.White
            .TableStyles(0).HeaderFont = New System.Drawing.Font("Tahoma", 8.0F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, System.Byte))
        End With

    End Sub

    Private Sub ClearForm()
        '--- gbCriteria ---
        cmbTransType.SelectedItem = "All Transaction"
        cmbTransType.Enabled = True
        cmbTransType.BackColor = Color.White
        cmbTransType.Focus()
        txtTellerID.Text = zCurrentUser()
        txtTellerID.ReadOnly = False
        txtTellerID.BackColor = Color.White
        dtePeriod.Text = Today()
        dtePeriod.Enabled = True
        dtePeriod.BackColor = Color.White
        txtTimeFrom.Text = ""
        txtTimeFrom.ReadOnly = False
        txtTimeFrom.BackColor = Color.White
        txtTimeTo.Text = ""
        txtTimeTo.ReadOnly = False
        txtTimeTo.BackColor = Color.White
        lblTimeTo.Visible = True
        lblTimeFrom.Visible = True
        lblTimeRange.Visible = False
        lstTimeRange.Visible = False
        lblBatch.Text = ""
        lblID.Text = ""
        '--- gbCash ---
        decGrCshAmt = 0
        decGrChgAmt = 0
        txt1000.Text = "0"
        txt500.Text = "0"
        txt200.Text = "0"
        txt100.Text = "0"
        txt50.Text = "0"
        txt20.Text = "0"
        txt10.Text = "0"
        txt5.Text = "0"
        txt1.Text = "0"
        txt025.Text = "0"
        txt010.Text = "0"
        txt005.Text = "0"
        txt001.Text = "0"
        txtTot1000.Text = "0.00"
        txtTot500.Text = "0.00"
        txtTot200.Text = "0.00"
        txtTot100.Text = "0.00"
        txtTot50.Text = "0.00"
        txtTot20.Text = "0.00"
        txtTot10.Text = "0.00"
        txtTot5.Text = "0.00"
        txtTot1.Text = "0.00"
        txtTot025.Text = "0.00"
        txtTot010.Text = "0.00"
        txtTot005.Text = "0.00"
        txtTot001.Text = "0.00"
        txtTotCash.Text = "0.00"
        txtAmtLeft.Text = "0.00"
        txtExCash.Text = "0.00"
        '--- gbCheque ---
        dgChequeDetails.DataSource = Nothing
        dgChequeStat.DataSource = Nothing
        txtTotCheque.Text = "0.00"
        txtExcessCheque.Text = "0.00"
        '--- gbControl ---
        txtGrandTot.Text = "0.00"
        txtRemarks.Text = ""
    End Sub

    Private Sub EnabledgbCash()
        txt1000.ReadOnly = False
        txt500.ReadOnly = False
        txt200.ReadOnly = False
        txt100.ReadOnly = False
        txt50.ReadOnly = False
        txt20.ReadOnly = False
        txt10.ReadOnly = False
        txt5.ReadOnly = False
        txt1.ReadOnly = False
        txt025.ReadOnly = False
        txt010.ReadOnly = False
        txt005.ReadOnly = False
        txt001.ReadOnly = False
        txt1000.BackColor = Color.White
        txt500.BackColor = Color.White
        txt200.BackColor = Color.White
        txt100.BackColor = Color.White
        txt50.BackColor = Color.White
        txt20.BackColor = Color.White
        txt10.BackColor = Color.White
        txt5.BackColor = Color.White
        txt1.BackColor = Color.White
        txt025.BackColor = Color.White
        txt010.BackColor = Color.White
        txt005.BackColor = Color.White
        txt001.BackColor = Color.White
    End Sub

    Private Sub DisEnabledgbCash()
        txt1000.ReadOnly = True
        txt500.ReadOnly = True
        txt200.ReadOnly = True
        txt100.ReadOnly = True
        txt50.ReadOnly = True
        txt20.ReadOnly = True
        txt10.ReadOnly = True
        txt5.ReadOnly = True
        txt1.ReadOnly = True
        txt025.ReadOnly = True
        txt010.ReadOnly = True
        txt005.ReadOnly = True
        txt001.ReadOnly = True
        txt1000.BackColor = Color.AliceBlue
        txt500.BackColor = Color.AliceBlue
        txt200.BackColor = Color.AliceBlue
        txt100.BackColor = Color.AliceBlue
        txt50.BackColor = Color.AliceBlue
        txt20.BackColor = Color.AliceBlue
        txt10.BackColor = Color.AliceBlue
        txt5.BackColor = Color.AliceBlue
        txt1.BackColor = Color.AliceBlue
        txt025.BackColor = Color.AliceBlue
        txt010.BackColor = Color.AliceBlue
        txt005.BackColor = Color.AliceBlue
        txt001.BackColor = Color.AliceBlue
    End Sub

    Private Sub DisabledgbCriteria()
        cmbTransType.Enabled = False
        cmbTransType.BackColor = Color.AliceBlue
        txtTellerID.ReadOnly = True
        txtTellerID.BackColor = Color.AliceBlue
        dtePeriod.Enabled = False
        dtePeriod.BackColor = Color.AliceBlue
        txtTimeFrom.ReadOnly = True
        txtTimeFrom.BackColor = Color.AliceBlue
        txtTimeTo.ReadOnly = True
        txtTimeTo.BackColor = Color.AliceBlue
    End Sub

    Private Function ValidateFields() As Boolean
        If CType(txtGrandTot.Text, Decimal) = 0 Or CType(txtExCash.Text, Decimal) <> CType(txtTotCash.Text, Decimal) Then
            If CType(txtGrandTot.Text, Decimal) = 0 Then
                MsgBox("Expected Cash and Cheque Collection Total should not be P0.00.", MsgBoxStyle.Information, "")
            ElseIf CType(txtExCash.Text, Decimal) > CType(txtTotCash.Text, Decimal) Then
                MsgBox("Please verify, your Remitted Cash cannot be less than the Expected Cash!", MsgBoxStyle.Information, "")
            ElseIf CType(txtExCash.Text, Decimal) < CType(txtTotCash.Text, Decimal) Then
                MsgBox("Please verify, your Remitted Cash exceeds the Expected Cash!", MsgBoxStyle.Information, "")
            End If
            Return False
        ElseIf CType(txtGrandTot.Text, Decimal) > 0 And CType(txtExCash.Text, Decimal) = CType(txtTotCash.Text, Decimal) Then
            Return True
        End If
    End Function

    Private Sub PopulatedtabCashDetails()
        dtabCashDetails = New DataTable

        If dtabCashDetails.Columns.Contains("Denomination") = True Then
            dtabCashDetails.Columns.Remove("Denomination")
        End If
        dtabCashDetails.Columns.Add("Denomination", Type.GetType("System.String"))

        If dtabCashDetails.Columns.Contains("Quantity") = True Then
            dtabCashDetails.Columns.Remove("Quantity")
        End If
        dtabCashDetails.Columns.Add("Quantity", Type.GetType("System.Int32"))

        If dtabCashDetails.Columns.Contains("Amount") = True Then
            dtabCashDetails.Columns.Remove("Amount")
        End If
        dtabCashDetails.Columns.Add("Amount", Type.GetType("System.Decimal"))

        Dim dtarow As DataRow
        dtarow = dtabCashDetails.NewRow
        dtarow("Denomination") = "P1000"
        dtarow("Quantity") = CType(txt1000.Text, Integer)
        dtarow("Amount") = CType(txtTot1000.Text, Decimal)
        dtabCashDetails.Rows.Add(dtarow)

        dtarow = dtabCashDetails.NewRow
        dtarow("Denomination") = "P500"
        dtarow("Quantity") = CType(txt500.Text, Integer)
        dtarow("Amount") = CType(txtTot500.Text, Decimal)
        dtabCashDetails.Rows.Add(dtarow)

        dtarow = dtabCashDetails.NewRow
        dtarow("Denomination") = "P200"
        dtarow("Quantity") = CType(txt200.Text, Integer)
        dtarow("Amount") = CType(txtTot200.Text, Decimal)
        dtabCashDetails.Rows.Add(dtarow)

        dtarow = dtabCashDetails.NewRow
        dtarow("Denomination") = "P100"
        dtarow("Quantity") = CType(txt100.Text, Integer)
        dtarow("Amount") = CType(txtTot100.Text, Decimal)
        dtabCashDetails.Rows.Add(dtarow)

        dtarow = dtabCashDetails.NewRow
        dtarow("Denomination") = "P50"
        dtarow("Quantity") = CType(txt50.Text, Integer)
        dtarow("Amount") = CType(txtTot50.Text, Decimal)
        dtabCashDetails.Rows.Add(dtarow)

        dtarow = dtabCashDetails.NewRow
        dtarow("Denomination") = "P20"
        dtarow("Quantity") = CType(txt20.Text, Integer)
        dtarow("Amount") = CType(txtTot20.Text, Decimal)
        dtabCashDetails.Rows.Add(dtarow)

        dtarow = dtabCashDetails.NewRow
        dtarow("Denomination") = "P10"
        dtarow("Quantity") = CType(txt10.Text, Integer)
        dtarow("Amount") = CType(txtTot10.Text, Decimal)
        dtabCashDetails.Rows.Add(dtarow)

        dtarow = dtabCashDetails.NewRow
        dtarow("Denomination") = "P5"
        dtarow("Quantity") = CType(txt5.Text, Integer)
        dtarow("Amount") = CType(txtTot5.Text, Decimal)
        dtabCashDetails.Rows.Add(dtarow)

        dtarow = dtabCashDetails.NewRow
        dtarow("Denomination") = "P1"
        dtarow("Quantity") = CType(txt1.Text, Integer)
        dtarow("Amount") = CType(txtTot1.Text, Decimal)
        dtabCashDetails.Rows.Add(dtarow)

        dtarow = dtabCashDetails.NewRow
        dtarow("Denomination") = "P025"
        dtarow("Quantity") = CType(txt025.Text, Integer)
        dtarow("Amount") = CType(txtTot025.Text, Decimal)
        dtabCashDetails.Rows.Add(dtarow)

        dtarow = dtabCashDetails.NewRow
        dtarow("Denomination") = "P010"
        dtarow("Quantity") = CType(txt010.Text, Integer)
        dtarow("Amount") = CType(txtTot010.Text, Decimal)
        dtabCashDetails.Rows.Add(dtarow)

        dtarow = dtabCashDetails.NewRow
        dtarow("Denomination") = "P005"
        dtarow("Quantity") = CType(txt005.Text, Integer)
        dtarow("Amount") = CType(txtTot005.Text, Decimal)
        dtabCashDetails.Rows.Add(dtarow)

        dtarow = dtabCashDetails.NewRow
        dtarow("Denomination") = "P001"
        dtarow("Quantity") = CType(txt001.Text, Integer)
        dtarow("Amount") = CType(txtTot001.Text, Decimal)
        dtabCashDetails.Rows.Add(dtarow)

    End Sub

#Region "Buttons"

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim strSQL As String
        Dim TransTyp As String
        Dim intBatch As String

        Select Case cmbTransType.SelectedItem
            Case "All Transaction"
                TransTyp = "ALL"
            Case "Import"
                TransTyp = "CYM"
            Case "Export"
                TransTyp = "CYX"
            Case "Special Services"
                TransTyp = "CYS"
            Case "Invoice"
                TransTyp = "INV"
        End Select

        If lblBatch.Text = "" Or lblID.Text = "" Then
            Dim ValFields As Boolean = ValidateFields()
            If ValFields = True Then
                strSQL = "SELECT Batch FROM TurnOverSlip WHERE TransTyp = " & CallClsCAC.getToString(TransTyp) & " AND " & _
                         "TellerID = " & CallClsCAC.getToString(UCase(txtTellerID.Text)) & " AND " & _
                         "CAST(BatchDate as SMALLDATETIME) = CAST(" & CallClsCAC.getToString(dtePeriod.Text) & " AS SMALLDATETIME)"

                intBatch = CallClsCAC.GetBatch(strSQL)

                If intBatch = "Application Cannot Retrive Data From Database" Then
                    MsgBox(intBatch, MsgBoxStyle.Critical, "ERROR")
                    Exit Sub
                End If

                intBatch = CType((CType(intBatch, Integer) + 1), String)

                strSQL = "INSERT INTO TurnOverSlip(BatchDate,TransTyp,TellerID,TimeFrom,TimeTo,P1000,P500,P200,P100,P50,P20,P10,P5,P1,P025,P010,P005,P001,Remarks,Batch) VALUES(" & _
                         CallClsCAC.getToString(dtePeriod.Text) & "," & _
                         CallClsCAC.getToString(TransTyp) & "," & _
                         CallClsCAC.getToString(UCase(txtTellerID.Text)) & "," & _
                         CallClsCAC.getToString(CType(TimeValue(txtTimeFrom.Text), String)) & "," & _
                         CallClsCAC.getToString(CType(TimeValue(txtTimeTo.Text), String)) & "," & _
                         txt1000.Text & "," & _
                         txt500.Text & "," & _
                         txt200.Text & "," & _
                         txt100.Text & "," & _
                         txt50.Text & "," & _
                         txt20.Text & "," & _
                         txt10.Text & "," & _
                         txt5.Text & "," & _
                         txt1.Text & "," & _
                         txt025.Text & "," & _
                         txt010.Text & "," & _
                         txt005.Text & "," & _
                         txt001.Text & "," & _
                         CallClsCAC.getToString(txtRemarks.Text) & "," & _
                         intBatch & ")"

                Dim isSave As Boolean = CallClsCAC.SaveCAC(strSQL)

                If isSave = True Then
                    MsgBox("Record is Saved", MsgBoxStyle.Information, "")
                Else
                    MsgBox("Record is Not Saved", MsgBoxStyle.Critical, "ERROR")
                End If
            End If
        ElseIf lblBatch.Text <> "" And lblID.Text <> "" Then
            Dim ValFields As Boolean = ValidateFields()
            If ValFields = True Then
                strSQL = "UPDATE TurnOverSlip SET TimeFrom = " & CallClsCAC.getToString(CType(TimeValue(txtTimeFrom.Text), String)) & _
                         ",TimeTo = " & CallClsCAC.getToString(CType(TimeValue(txtTimeTo.Text), String)) & _
                         ",P1000 = " & txt1000.Text & _
                         ",P500 = " & txt500.Text & _
                         ",P200 = " & txt200.Text & _
                         ",P100 = " & txt100.Text & _
                         ",P50 = " & txt50.Text & _
                         ",P20 = " & txt20.Text & _
                         ",P10 = " & txt10.Text & _
                         ",P5 = " & txt5.Text & _
                         ",P1 = " & txt1.Text & _
                         ",P025 = " & txt025.Text & _
                         ",P010 = " & txt010.Text & _
                         ",P005 = " & txt005.Text & _
                         ",P001 = " & txt001.Text & _
                         ",Remarks = " & CallClsCAC.getToString(txtRemarks.Text) & _
                         " WHERE ID = " & lblID.Text

                Dim isUpdate As Boolean = CallClsCAC.SaveCAC(strSQL)

                If isUpdate = True Then
                    MsgBox("Record is Updated", MsgBoxStyle.Information, "")
                Else
                    MsgBox("Record is Not Updated", MsgBoxStyle.Critical, "ERROR")
                End If
            End If
        End If
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        ClearForm()
    End Sub

    Private Sub btnCLOSE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCLOSE.Click
        End
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Cursor = Cursors.WaitCursor
        If CType(txtGrandTot.Text, Decimal) = 0 Then
            MsgBox("There is no Cash and Collection Report to Display.", MsgBoxStyle.Exclamation, "Invalid")
            Cursor = Cursors.Default
            Exit Sub
        End If
        PopulatedtabCashDetails()

        Dim frmRptCAC As New frmRptCAC
        Dim RptCAC As New rptCAC

        RptCAC.OpenSubreport("rptCash.rpt").SetDataSource(dtabCashDetails)
        RptCAC.OpenSubreport("rptCheque.rpt").SetDataSource(dtabDetails)
        RptCAC.SetParameterValue("strDate", Trim(dtePeriod.Text))
        RptCAC.SetParameterValue("strTimeRange", Trim(txtTimeFrom.Text & " - " & txtTimeTo.Text))
        RptCAC.SetParameterValue("strTranType", Trim(cmbTransType.SelectedItem))
        RptCAC.SetParameterValue("strCurDate", Trim(CType(Today(), String)))
        RptCAC.SetParameterValue("strRemarks", Trim(txtRemarks.Text))
        RptCAC.SetParameterValue("strUserID", UCase(Trim(txtTellerID.Text)))
        RptCAC.SetParameterValue("numGrandTot", Trim(txtGrandTot.Text))

        frmRptCAC.crvReports.ReportSource = RptCAC
        Cursor = Cursors.Default
        frmRptCAC.ShowDialog()
        Cursor = Cursors.Default
    End Sub

#End Region


End Class
