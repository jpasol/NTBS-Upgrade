Public Class frmReports
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
    Friend WithEvents grpCriteria As System.Windows.Forms.GroupBox
    Friend WithEvents grpTrnDates As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents crvViewer As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents cmbInOut As System.Windows.Forms.ComboBox
    Friend WithEvents cmbFullEmpty As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSize As System.Windows.Forms.ComboBox
    Friend WithEvents dtpDateFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpDateTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.crvViewer = New CrystalDecisions.Windows.Forms.CrystalReportViewer
        Me.grpCriteria = New System.Windows.Forms.GroupBox
        Me.cmbInOut = New System.Windows.Forms.ComboBox
        Me.cmbFullEmpty = New System.Windows.Forms.ComboBox
        Me.cmbSize = New System.Windows.Forms.ComboBox
        Me.grpTrnDates = New System.Windows.Forms.GroupBox
        Me.dtpDateFrom = New System.Windows.Forms.DateTimePicker
        Me.dtpDateTo = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnView = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.grpCriteria.SuspendLayout()
        Me.grpTrnDates.SuspendLayout()
        Me.SuspendLayout()
        '
        'crvViewer
        '
        Me.crvViewer.ActiveViewIndex = -1
        Me.crvViewer.DisplayGroupTree = False
        Me.crvViewer.Location = New System.Drawing.Point(264, 8)
        Me.crvViewer.Name = "crvViewer"
        Me.crvViewer.ReportSource = Nothing
        Me.crvViewer.Size = New System.Drawing.Size(592, 672)
        Me.crvViewer.TabIndex = 0
        '
        'grpCriteria
        '
        Me.grpCriteria.BackColor = System.Drawing.Color.WhiteSmoke
        Me.grpCriteria.Controls.Add(Me.Label6)
        Me.grpCriteria.Controls.Add(Me.Label5)
        Me.grpCriteria.Controls.Add(Me.Label4)
        Me.grpCriteria.Controls.Add(Me.cmbSize)
        Me.grpCriteria.Controls.Add(Me.cmbFullEmpty)
        Me.grpCriteria.Controls.Add(Me.cmbInOut)
        Me.grpCriteria.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpCriteria.Location = New System.Drawing.Point(8, 128)
        Me.grpCriteria.Name = "grpCriteria"
        Me.grpCriteria.Size = New System.Drawing.Size(248, 192)
        Me.grpCriteria.TabIndex = 1
        Me.grpCriteria.TabStop = False
        Me.grpCriteria.Text = "Viewing Criteria"
        '
        'cmbInOut
        '
        Me.cmbInOut.Items.AddRange(New Object() {"ALL", "INBOUND", "OUTBOUND"})
        Me.cmbInOut.Location = New System.Drawing.Point(16, 40)
        Me.cmbInOut.Name = "cmbInOut"
        Me.cmbInOut.Size = New System.Drawing.Size(192, 21)
        Me.cmbInOut.TabIndex = 0
        Me.cmbInOut.Text = "ALL"
        '
        'cmbFullEmpty
        '
        Me.cmbFullEmpty.Items.AddRange(New Object() {"ALL", "FULL", "EMPTY"})
        Me.cmbFullEmpty.Location = New System.Drawing.Point(16, 96)
        Me.cmbFullEmpty.Name = "cmbFullEmpty"
        Me.cmbFullEmpty.Size = New System.Drawing.Size(192, 21)
        Me.cmbFullEmpty.TabIndex = 1
        Me.cmbFullEmpty.Text = "ALL"
        '
        'cmbSize
        '
        Me.cmbSize.Items.AddRange(New Object() {"ALL", "20", "40", "45"})
        Me.cmbSize.Location = New System.Drawing.Point(16, 152)
        Me.cmbSize.Name = "cmbSize"
        Me.cmbSize.Size = New System.Drawing.Size(192, 21)
        Me.cmbSize.TabIndex = 2
        Me.cmbSize.Text = "ALL"
        '
        'grpTrnDates
        '
        Me.grpTrnDates.Controls.Add(Me.Label2)
        Me.grpTrnDates.Controls.Add(Me.dtpDateTo)
        Me.grpTrnDates.Controls.Add(Me.dtpDateFrom)
        Me.grpTrnDates.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpTrnDates.Location = New System.Drawing.Point(8, 64)
        Me.grpTrnDates.Name = "grpTrnDates"
        Me.grpTrnDates.Size = New System.Drawing.Size(248, 56)
        Me.grpTrnDates.TabIndex = 2
        Me.grpTrnDates.TabStop = False
        Me.grpTrnDates.Text = "Date Billed"
        '
        'dtpDateFrom
        '
        Me.dtpDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpDateFrom.Location = New System.Drawing.Point(8, 24)
        Me.dtpDateFrom.Name = "dtpDateFrom"
        Me.dtpDateFrom.Size = New System.Drawing.Size(96, 20)
        Me.dtpDateFrom.TabIndex = 0
        '
        'dtpDateTo
        '
        Me.dtpDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpDateTo.Location = New System.Drawing.Point(144, 24)
        Me.dtpDateTo.Name = "dtpDateTo"
        Me.dtpDateTo.Size = New System.Drawing.Size(96, 20)
        Me.dtpDateTo.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(248, 48)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Container Information Summary"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.SkyBlue
        Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label4.Location = New System.Drawing.Point(16, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(192, 16)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "INBOUND\OUTBOUND"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.SkyBlue
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.Location = New System.Drawing.Point(16, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(192, 16)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "FULL/EMPTY"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.SkyBlue
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label6.Location = New System.Drawing.Point(16, 136)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(192, 16)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "CONTAINER SIZE"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.SteelBlue
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label7.Location = New System.Drawing.Point(0, 48)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(208, 8)
        Me.Label7.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(104, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "TO"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btnView
        '
        Me.btnView.BackColor = System.Drawing.Color.Silver
        Me.btnView.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnView.Location = New System.Drawing.Point(176, 328)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(80, 24)
        Me.btnView.TabIndex = 7
        Me.btnView.Text = "&View"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.SteelBlue
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Location = New System.Drawing.Point(56, 360)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(208, 8)
        Me.Label3.TabIndex = 8
        '
        'frmReports
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(864, 686)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnView)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.grpTrnDates)
        Me.Controls.Add(Me.grpCriteria)
        Me.Controls.Add(Me.crvViewer)
        Me.Name = "frmReports"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "NTBS and OP Reports"
        Me.grpCriteria.ResumeLayout(False)
        Me.grpTrnDates.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private dtabCtnInfo As dsReports.CtnInfoDataTable

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        If cmbInOut.Text = "ALL" Then
            Dim dtabObj As New DataTable

            dtabObj = Get_CCRCyx(dtpDateFrom.Text, dtpDateTo.Text)
        ElseIf cmbInOut.Text = "IMPORT" Then
        Else
        End If

        If cmbFullEmpty.Text = "ALL" Then

        ElseIf cmbFullEmpty.Text = "FULL" Then
        Else
        End If

        Select Case Trim(cmbSize.Text)
            Case "20"
            Case "40"
            Case "45"
            Case Else

        End Select

        DisConnect()
        ViewReport()
    End Sub

    Private Sub Add_To_CtnInfo(ByVal strCtnNo As String, ByVal intCtnSze As Integer, ByVal strFE As String, ByVal strIO As String, ByVal strCCRGps As String)
        Dim rowCtnInfo As dsReports.CtnInfoRow

        rowCtnInfo = dtabCtnInfo.NewCtnInfoRow

        With rowCtnInfo
            .CtnNo = Trim(strCtnNo)
            .CtnSze = intCtnSze
            .FullEmpty = Trim(strFE)
            .InOut = Trim(strIO)
            .CCRGps = Trim(strCCRGps)
        End With

        dtabCtnInfo.Rows.Add(rowCtnInfo)
    End Sub

    Private Sub ViewReport()
        Dim rptCtnInfo As New rptContainerInfo

        rptCtnInfo.SetDataSource(dtabCtnInfo)
        crvViewer.ReportSource = rptCtnInfo
        crvViewer.Show()
    End Sub
End Class
