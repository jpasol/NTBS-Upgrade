Public Class frmRptCAC
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
    Friend WithEvents statCACBar As System.Windows.Forms.StatusBar
    Friend WithEvents statPanelUser As System.Windows.Forms.StatusBarPanel
    Friend WithEvents statPanelDate As System.Windows.Forms.StatusBarPanel
    Friend WithEvents statPanelTime As System.Windows.Forms.StatusBarPanel
    Friend WithEvents gbHeader As System.Windows.Forms.GroupBox
    Friend WithEvents btnCLOSE As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents crvReports As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmRptCAC))
        Me.statCACBar = New System.Windows.Forms.StatusBar
        Me.statPanelUser = New System.Windows.Forms.StatusBarPanel
        Me.statPanelDate = New System.Windows.Forms.StatusBarPanel
        Me.statPanelTime = New System.Windows.Forms.StatusBarPanel
        Me.gbHeader = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.btnCLOSE = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.crvReports = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        CType(Me.statPanelUser, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.statPanelDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.statPanelTime, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbHeader.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'statCACBar
        '
        Me.statCACBar.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.statCACBar.Location = New System.Drawing.Point(0, 724)
        Me.statCACBar.Name = "statCACBar"
        Me.statCACBar.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.statPanelUser, Me.statPanelDate, Me.statPanelTime})
        Me.statCACBar.ShowPanels = True
        Me.statCACBar.Size = New System.Drawing.Size(921, 22)
        Me.statCACBar.SizingGrip = False
        Me.statCACBar.TabIndex = 5
        '
        'statPanelUser
        '
        Me.statPanelUser.Width = 550
        '
        'statPanelDate
        '
        Me.statPanelDate.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.statPanelDate.Width = 250
        '
        'statPanelTime
        '
        Me.statPanelTime.Alignment = System.Windows.Forms.HorizontalAlignment.Right
        Me.statPanelTime.Width = 120
        '
        'gbHeader
        '
        Me.gbHeader.Controls.Add(Me.Label1)
        Me.gbHeader.Controls.Add(Me.PictureBox1)
        Me.gbHeader.Controls.Add(Me.btnCLOSE)
        Me.gbHeader.Location = New System.Drawing.Point(8, 3)
        Me.gbHeader.Name = "gbHeader"
        Me.gbHeader.Size = New System.Drawing.Size(905, 95)
        Me.gbHeader.TabIndex = 23
        Me.gbHeader.TabStop = False
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.AliceBlue
        Me.Label1.Location = New System.Drawing.Point(392, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(384, 72)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "SBITC Cash and Cheque Collection Report"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(11, 16)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(368, 72)
        Me.PictureBox1.TabIndex = 7
        Me.PictureBox1.TabStop = False
        '
        'btnCLOSE
        '
        Me.btnCLOSE.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold)
        Me.btnCLOSE.Image = CType(resources.GetObject("btnCLOSE.Image"), System.Drawing.Image)
        Me.btnCLOSE.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnCLOSE.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnCLOSE.Location = New System.Drawing.Point(820, 14)
        Me.btnCLOSE.Name = "btnCLOSE"
        Me.btnCLOSE.Size = New System.Drawing.Size(75, 73)
        Me.btnCLOSE.TabIndex = 6
        Me.btnCLOSE.Text = "CLOSE                                                                            " & _
        ""
        Me.btnCLOSE.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.crvReports)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 104)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(905, 608)
        Me.GroupBox1.TabIndex = 24
        Me.GroupBox1.TabStop = False
        '
        'crvReports
        '
        Me.crvReports.ActiveViewIndex = -1
        Me.crvReports.DisplayGroupTree = False
        Me.crvReports.Location = New System.Drawing.Point(8, 16)
        Me.crvReports.Name = "crvReports"
        Me.crvReports.ReportSource = Nothing
        Me.crvReports.Size = New System.Drawing.Size(888, 584)
        Me.crvReports.TabIndex = 1
        '
        'frmRptCAC
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.LightSlateGray
        Me.ClientSize = New System.Drawing.Size(921, 746)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.gbHeader)
        Me.Controls.Add(Me.statCACBar)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmRptCAC"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cash and Cheque Collection Report"
        CType(Me.statPanelUser, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.statPanelDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.statPanelTime, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbHeader.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmRptCAC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SetStatusBar()
    End Sub

    Private Sub SetStatusBar()
        statPanelDate.Text = CType(FormatDateTime(Today(), DateFormat.LongDate), String) & " "
        statPanelTime.Text = CType(TimeValue(Now()), String) & " "
        statPanelUser.Text = " User Name : " & UCase(zCurrentUser())
    End Sub

    Private Sub btnCLOSE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCLOSE.Click
        Me.Close()
    End Sub

    'Private Sub PopulateRpt()
    '    Dim frmCAC As frmCAC
    '    Dim RptCash As New rptCash
    '    Dim RptCheque As New rptCheque
    '    Dim RptCAC As New rptCAC

    '    RptCash.SetDataSource(dtabCashDetails)

    '    RptCheque.SetDataSource(dtabDetails)

    '    RptCAC.SetParameterValue("strDate", Trim(frmCAC.dtePeriod.Value))
    '    RptCAC.SetParameterValue("strTimeRange", Trim(frmCAC.txtTimeFrom.Text & " - " & frmCAC.txtTimeTo.Text))
    '    RptCAC.SetParameterValue("strTranType", Trim(frmCAC.cmbTransType.SelectedItem))
    '    RptCAC.SetParameterValue("strCurDate", Trim(CType(Today(), String)))
    '    RptCAC.SetParameterValue("strRemarks", Trim(frmCAC.txtRemarks.Text))
    '    RptCAC.SetParameterValue("strUserID", Trim(frmCAC.txtTellerID.Text))
    '    RptCAC.SetParameterValue("numGrandTot", Trim(frmCAC.txtGrandTot.Text))

    '    crvReports.ReportSource = RptCAC
    '    Cursor = Cursors.Default
    'End Sub
End Class
