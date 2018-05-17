Imports System.Configuration.ConfigurationSettings

Public Class frmMain
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
    Friend WithEvents btnUploadManifest As System.Windows.Forms.Button
    Friend WithEvents btnEncodeManifest As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents StatusBar As System.Windows.Forms.StatusBar
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel3 As System.Windows.Forms.StatusBarPanel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.btnUploadManifest = New System.Windows.Forms.Button
        Me.btnEncodeManifest = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.StatusBar = New System.Windows.Forms.StatusBar
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel3 = New System.Windows.Forms.StatusBarPanel
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnUploadManifest
        '
        Me.btnUploadManifest.BackColor = System.Drawing.Color.SlateGray
        Me.btnUploadManifest.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUploadManifest.Location = New System.Drawing.Point(32, 72)
        Me.btnUploadManifest.Name = "btnUploadManifest"
        Me.btnUploadManifest.Size = New System.Drawing.Size(112, 32)
        Me.btnUploadManifest.TabIndex = 0
        Me.btnUploadManifest.Text = "Upload Manifest"
        '
        'btnEncodeManifest
        '
        Me.btnEncodeManifest.BackColor = System.Drawing.Color.SlateGray
        Me.btnEncodeManifest.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnEncodeManifest.Location = New System.Drawing.Point(32, 112)
        Me.btnEncodeManifest.Name = "btnEncodeManifest"
        Me.btnEncodeManifest.Size = New System.Drawing.Size(112, 32)
        Me.btnEncodeManifest.TabIndex = 1
        Me.btnEncodeManifest.Text = "Encode Manifest"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.SteelBlue
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(296, 24)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "SBITC Cargo Manifest System "
        '
        'btnClose
        '
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnClose.Location = New System.Drawing.Point(304, 0)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(16, 16)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "X"
        '
        'StatusBar
        '
        Me.StatusBar.Location = New System.Drawing.Point(0, 154)
        Me.StatusBar.Name = "StatusBar"
        Me.StatusBar.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarPanel1, Me.StatusBarPanel2, Me.StatusBarPanel3})
        Me.StatusBar.ShowPanels = True
        Me.StatusBar.Size = New System.Drawing.Size(320, 22)
        Me.StatusBar.TabIndex = 4
        '
        'StatusBarPanel1
        '
        Me.StatusBarPanel1.Width = 80
        '
        'StatusBarPanel2
        '
        Me.StatusBarPanel2.Width = 80
        '
        'StatusBarPanel3
        '
        Me.StatusBarPanel3.Width = 150
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(320, 176)
        Me.ControlBox = False
        Me.Controls.Add(Me.StatusBar)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnEncodeManifest)
        Me.Controls.Add(Me.btnUploadManifest)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private strServer As String = CType(AppSettings("Server"), String)
    Private strDatabase As String = CType(AppSettings("Database"), String)

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnEncodeManifest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEncodeManifest.Click
        'Dim objManifest As prjManifestDE.clsManifestDE
        'Dim gconnstr As String

        ''Set objManifest = New prjManifestDE.clsManifestDE

        'objManifest = CreateObject("prjManifestDE.clsManifestDE")

        'gconnstr = "Provider=sqloledb" & _
        '    ";Data Source=" & Trim(strServer) & _
        '    ";Initial Catalog=" & Trim(strDatabase) & _
        '    ";Integrated Security=SSPI"

        'With objManifest
        '    .ConnectByStr(gconnstr)
        '    .Execute()
        '    .Disconnect()
        'End With

        'objManifest = Nothing
    End Sub

    Private Sub btnUploadManifest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUploadManifest.Click
        Dim objUpload As New Upload_Manifest.clsUpload

        With objUpload
            .strServer = Trim(strServer)
            .strDatabase = Trim(strDatabase)
            .main()
        End With
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        StatusBarPanel1.Text = strServer
        StatusBarPanel2.Text = strDatabase
        StatusBarPanel3.Text = Now
    End Sub
End Class
