Imports System.IO
Imports Microsoft.VisualBasic

Public Class frmUploading
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
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents btnUpload As System.Windows.Forms.Button
    Friend WithEvents lstErrors As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblUploadingStatus As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnOpenBOLFile As System.Windows.Forms.Button
    Friend WithEvents btnOpenGENFile As System.Windows.Forms.Button
    Friend WithEvents btnOpenCTNFile As System.Windows.Forms.Button
    Friend WithEvents txtBOLFilePath As System.Windows.Forms.TextBox
    Friend WithEvents txtGENFilePath As System.Windows.Forms.TextBox
    Friend WithEvents txtCTNFilePath As System.Windows.Forms.TextBox
    Friend WithEvents StatusBar As System.Windows.Forms.StatusBar
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel3 As System.Windows.Forms.StatusBarPanel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmUploading))
        Me.btnUpload = New System.Windows.Forms.Button
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.btnOpenBOLFile = New System.Windows.Forms.Button
        Me.lstErrors = New System.Windows.Forms.ListBox
        Me.txtBOLFilePath = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.lblUploadingStatus = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtGENFilePath = New System.Windows.Forms.TextBox
        Me.btnOpenGENFile = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtCTNFilePath = New System.Windows.Forms.TextBox
        Me.btnOpenCTNFile = New System.Windows.Forms.Button
        Me.StatusBar = New System.Windows.Forms.StatusBar
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel3 = New System.Windows.Forms.StatusBarPanel
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnUpload
        '
        Me.btnUpload.BackColor = System.Drawing.Color.SlateGray
        Me.btnUpload.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnUpload.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnUpload.Location = New System.Drawing.Point(8, 176)
        Me.btnUpload.Name = "btnUpload"
        Me.btnUpload.Size = New System.Drawing.Size(112, 24)
        Me.btnUpload.TabIndex = 1
        Me.btnUpload.Text = "Upload to NTBS"
        '
        'btnOpenBOLFile
        '
        Me.btnOpenBOLFile.BackColor = System.Drawing.Color.SlateGray
        Me.btnOpenBOLFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnOpenBOLFile.Font = New System.Drawing.Font("Arial Black", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpenBOLFile.Location = New System.Drawing.Point(360, 40)
        Me.btnOpenBOLFile.Name = "btnOpenBOLFile"
        Me.btnOpenBOLFile.Size = New System.Drawing.Size(32, 22)
        Me.btnOpenBOLFile.TabIndex = 2
        Me.btnOpenBOLFile.Text = "..."
        '
        'lstErrors
        '
        Me.lstErrors.BackColor = System.Drawing.Color.AliceBlue
        Me.lstErrors.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstErrors.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstErrors.ItemHeight = 14
        Me.lstErrors.Location = New System.Drawing.Point(8, 232)
        Me.lstErrors.Name = "lstErrors"
        Me.lstErrors.Size = New System.Drawing.Size(384, 156)
        Me.lstErrors.TabIndex = 3
        '
        'txtBOLFilePath
        '
        Me.txtBOLFilePath.BackColor = System.Drawing.Color.AliceBlue
        Me.txtBOLFilePath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBOLFilePath.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBOLFilePath.Location = New System.Drawing.Point(8, 40)
        Me.txtBOLFilePath.Name = "txtBOLFilePath"
        Me.txtBOLFilePath.ReadOnly = True
        Me.txtBOLFilePath.Size = New System.Drawing.Size(352, 22)
        Me.txtBOLFilePath.TabIndex = 4
        Me.txtBOLFilePath.Text = ""
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(128, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "BOL File Path"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 216)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(144, 16)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Uploading Results Log"
        '
        'lblUploadingStatus
        '
        Me.lblUploadingStatus.Font = New System.Drawing.Font("Lucida Sans Unicode", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUploadingStatus.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(192, Byte))
        Me.lblUploadingStatus.Location = New System.Drawing.Point(128, 176)
        Me.lblUploadingStatus.Name = "lblUploadingStatus"
        Me.lblUploadingStatus.Size = New System.Drawing.Size(264, 24)
        Me.lblUploadingStatus.TabIndex = 6
        Me.lblUploadingStatus.Text = "<-- Click this button to start Uploading"
        Me.lblUploadingStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.Label3.Location = New System.Drawing.Point(8, 120)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(128, 16)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "GEN File Path"
        '
        'txtGENFilePath
        '
        Me.txtGENFilePath.BackColor = System.Drawing.Color.AliceBlue
        Me.txtGENFilePath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtGENFilePath.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGENFilePath.Location = New System.Drawing.Point(8, 136)
        Me.txtGENFilePath.Name = "txtGENFilePath"
        Me.txtGENFilePath.ReadOnly = True
        Me.txtGENFilePath.Size = New System.Drawing.Size(352, 22)
        Me.txtGENFilePath.TabIndex = 9
        Me.txtGENFilePath.Text = ""
        '
        'btnOpenGENFile
        '
        Me.btnOpenGENFile.BackColor = System.Drawing.Color.SlateGray
        Me.btnOpenGENFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnOpenGENFile.Font = New System.Drawing.Font("Arial Black", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpenGENFile.Location = New System.Drawing.Point(360, 136)
        Me.btnOpenGENFile.Name = "btnOpenGENFile"
        Me.btnOpenGENFile.Size = New System.Drawing.Size(32, 22)
        Me.btnOpenGENFile.TabIndex = 8
        Me.btnOpenGENFile.Text = "..."
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(64, Byte))
        Me.Label4.Location = New System.Drawing.Point(8, 72)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 16)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "CTN File Path"
        '
        'txtCTNFilePath
        '
        Me.txtCTNFilePath.BackColor = System.Drawing.Color.AliceBlue
        Me.txtCTNFilePath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCTNFilePath.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCTNFilePath.Location = New System.Drawing.Point(8, 88)
        Me.txtCTNFilePath.Name = "txtCTNFilePath"
        Me.txtCTNFilePath.ReadOnly = True
        Me.txtCTNFilePath.Size = New System.Drawing.Size(352, 22)
        Me.txtCTNFilePath.TabIndex = 12
        Me.txtCTNFilePath.Text = ""
        '
        'btnOpenCTNFile
        '
        Me.btnOpenCTNFile.BackColor = System.Drawing.Color.SlateGray
        Me.btnOpenCTNFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnOpenCTNFile.Font = New System.Drawing.Font("Arial Black", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpenCTNFile.Location = New System.Drawing.Point(360, 88)
        Me.btnOpenCTNFile.Name = "btnOpenCTNFile"
        Me.btnOpenCTNFile.Size = New System.Drawing.Size(32, 22)
        Me.btnOpenCTNFile.TabIndex = 11
        Me.btnOpenCTNFile.Text = "..."
        '
        'StatusBar
        '
        Me.StatusBar.Location = New System.Drawing.Point(0, 408)
        Me.StatusBar.Name = "StatusBar"
        Me.StatusBar.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarPanel1, Me.StatusBarPanel2, Me.StatusBarPanel3})
        Me.StatusBar.ShowPanels = True
        Me.StatusBar.Size = New System.Drawing.Size(400, 22)
        Me.StatusBar.TabIndex = 14
        '
        'StatusBarPanel3
        '
        Me.StatusBarPanel3.Width = 200
        '
        'frmUploading
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.SteelBlue
        Me.ClientSize = New System.Drawing.Size(400, 430)
        Me.Controls.Add(Me.StatusBar)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtCTNFilePath)
        Me.Controls.Add(Me.btnOpenCTNFile)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtGENFilePath)
        Me.Controls.Add(Me.btnOpenGENFile)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblUploadingStatus)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtBOLFilePath)
        Me.Controls.Add(Me.lstErrors)
        Me.Controls.Add(Me.btnOpenBOLFile)
        Me.Controls.Add(Me.btnUpload)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmUploading"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cargo Manifest Uploading Facility"
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Object
    Private objUpload As New clsUpload

    'Variables
    Private strFileName As String = ""

    Private Sub btnOpenBOLFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenBOLFile.Click
        Dim openFile As New System.Windows.Forms.OpenFileDialog

        lblUploadingStatus.Text = "<-- Click this button to start Uploading"

        openFile.DefaultExt = "bol"
        openFile.Filter = "Text documents (*.bol)|*.bol"

        openFile.ShowDialog()

        If openFile.FileNames.Length > 0 Then
            If openFile.CheckPathExists = True Then
                If openFile.CheckFileExists = True Then
                    txtBOLFilePath.Text = openFile.FileName
                Else
                    MsgBox("File not valid!", MsgBoxStyle.OKOnly, "BOL File")
                End If
            Else
                MsgBox("Path not valid!", MsgBoxStyle.OKOnly, "BOL Path")
            End If
        End If
    End Sub

    Private Sub btnOpenCTNFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenCTNFile.Click
        Dim openFile As New System.Windows.Forms.OpenFileDialog

        lblUploadingStatus.Text = "<-- Click this button to start Uploading"

        openFile.DefaultExt = "ctn"
        openFile.Filter = "Text documents (*.ctn)|*.ctn"

        openFile.ShowDialog()

        If openFile.FileNames.Length > 0 Then
            If openFile.CheckPathExists = True Then
                If openFile.CheckFileExists = True Then
                    txtCTNFilePath.Text = openFile.FileName
                Else
                    MsgBox("File not valid!", MsgBoxStyle.OKOnly, "CTN File")
                End If
            Else
                MsgBox("Path not valid!", MsgBoxStyle.OKOnly, "CTN Path")
            End If
        End If
    End Sub

    Private Sub btnOpenGENFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenGENFile.Click
        Dim openFile As New System.Windows.Forms.OpenFileDialog

        lblUploadingStatus.Text = "<-- Click this button to start Uploading"

        openFile.DefaultExt = "gen"
        openFile.Filter = "Text documents (*.gen)|*.gen"

        openFile.ShowDialog()

        If openFile.FileNames.Length > 0 Then
            If openFile.CheckPathExists = True Then
                If openFile.CheckFileExists = True Then
                    txtGENFilePath.Text = openFile.FileName
                Else
                    MsgBox("File not valid!", MsgBoxStyle.OKOnly, "GEN File")
                End If
            Else
                MsgBox("Path not valid!", MsgBoxStyle.OKOnly, "GEN Path")
            End If
        End If
    End Sub

    Private Sub btnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpload.Click
        If Trim(txtBOLFilePath.Text) <> "" And Trim(txtCTNFilePath.Text) <> "" And Trim(txtGENFilePath.Text) <> "" Then
            'Check if files belong to the same shipping line
            If Chk_ValidFiles() = False Then
                MsgBox("Uploading can't proceed due to inconsistent file entries!", MsgBoxStyle.Exclamation, "Uploading Restrictions")
                Exit Sub
            End If
            Dim strResponse As String = ""
            strResponse = MsgBox("Are you sure you want to upload these files?", MsgBoxStyle.YesNo, "Uploading Process")
            If strResponse = vbYes Then
                lblUploadingStatus.Text = "Application Uploading . . ."
                System.Windows.Forms.Application.DoEvents()
                'Upload BOL
                Call UploadBOL()
                'Upload CTN
                Call UploadCTN()
                'Upload GEN
                Call UploadGEN()
                lblUploadingStatus.Text = "F I N I S H E D"
            End If
        Else
            MsgBox("Please provide valid entries!", MsgBoxStyle.Exclamation, "Uploading Requirements")
        End If
    End Sub

    Private Function Chk_ValidFiles() As Boolean
        'Get BOL File Name
        strFileName = Mid(Trim(txtBOLFilePath.Text), Trim(txtBOLFilePath.Text).Length - 10, 7)

        'Check if CTN filename is the same with BOL File Name
        If InStr(Trim(txtCTNFilePath.Text), Trim(strFileName)) = 0 Then
            Return False
        End If

        'Check if GEN filename is the same with CTN and BOL
        If InStr(Trim(txtGENFilePath.Text), Trim(strFileName)) = 0 Then
            Return False
        End If

        Return True
    End Function

    Private Sub UploadBOL()
        Dim oFile As File
        Dim oRead As StreamReader
        Dim strLine As String = ""
        Dim arrLine As Array
        Dim strCarCde As String = ""

        oRead = oFile.OpenText(Trim(txtBOLFilePath.Text))

        lstErrors.Items.Clear()
        While oRead.Peek <> -1
            strLine = oRead.ReadLine()
            'Remove unnecessary characters
            'strLine = Replace(strLine, Mid(Trim(strLine), 1, 1), " ")
            strLine = CommaDelimitedLine(strLine, ",", """")
            strLine = Replace(strLine, "'", "''")
            arrLine = strLine.Split(CChar(","))

            'Legend:
            'arrLine(2) = Bill No.
            'arrLine(1) = Registry No.
            'arrLine(16) = Consignee
            'arrLine(22) = Broker
            'arrLine(44) = Container Description
            'arrLine(27) = Port of Origin

            'Check if Bill No. already exist in CargoMHead table 
            If objUpload.Check_Header(Trim(arrLine(2).ToString)) = True Then
                'Log
                lstErrors.Items.Add("BL No. " & Trim(arrLine(2).ToString) & " already exist!")
                System.Windows.Forms.Application.DoEvents()
            Else
                'Get Carrier Code
                If Mid(Trim(strFileName), 1, 3) = "APL" Then
                    strCarCde = "100001"
                ElseIf Mid(Trim(strFileName), 1, 3) = "MSK" Then
                    strCarCde = "100062"
                ElseIf Mid(Trim(strFileName), 1, 3) = "WHL" Then
                    strCarCde = "100090"
                Else
                    MsgBox("Uploading of data from Shipping lines other than APL,MSK and WHL is not allowed!", MsgBoxStyle.Information, "Header Uploading Restriction")
                    MsgBox("For this transaction, kindly use the encoding facility!", MsgBoxStyle.Information, "Header Uploading Restriction")
                    Exit Sub
                End If
                'Insert Header data
                objUpload.Insert_Header(arrLine(2).ToString, strCarCde, _
                                        arrLine(27).ToString, arrLine(1).ToString, _
                                        arrLine(16).ToString, arrLine(44).ToString, _
                                        arrLine(22).ToString, zCurrentUser)
                End If
        End While

        oFile = Nothing
        oRead = Nothing
    End Sub

    Private Sub UploadCTN()
        Dim oFile As File
        Dim oRead As StreamReader
        Dim strLine As String = ""
        Dim arrLine As Array

        oRead = oFile.OpenText(Trim(txtCTNFilePath.Text))

        While oRead.Peek <> -1
            strLine = oRead.ReadLine()
            'Remove unnecessary characters
            strLine = Replace(strLine, Mid(Trim(strLine), 1, 1), " ")
            arrLine = strLine.Split(CChar(","))

            'Legend:
            'arrLine(1) = Registry Number
            'arrLine(2) = Bill No.
            'arrLine(3) = Container No.
            'arrline(4) = Container Size
            'arrline(5) = Full or Empty
            'arrLine(7) = SilNo.

            'Check if Bill No. exist in CargoMHead table
            If objUpload.Check_Header(Trim(arrLine(2).ToString)) = True Then
                'Check if Bill No. and Container No. already exist in CargoMDet
                If objUpload.Check_Detail(Trim(arrLine(2).ToString), Trim(arrLine(3).ToString)) = True Then
                    'Log
                    lstErrors.Items.Add("Container No. " & Trim(arrLine(3).ToString) & " from BL No. " & Trim(arrLine(2).ToString) & " already exist!")
                    System.Windows.Forms.Application.DoEvents()
                Else
                    'Insert Detail data
                    objUpload.Insert_Detail(arrLine(2).ToString, arrLine(3).ToString, _
                                            "", arrLine(4).ToString, "", arrLine(5).ToString, _
                                            arrLine(7).ToString, arrLine(1).ToString, zCurrentUser)
                End If
            Else
                'Log 
                lstErrors.Items.Add("Can't add Container No. " & Trim(arrLine(3).ToString) & " from BL No. " & Trim(arrLine(2).ToString) & " because Header data does not exist!")
                System.Windows.Forms.Application.DoEvents()
            End If
        End While

        oFile = Nothing
        oRead = Nothing
    End Sub

    Private Sub UploadGEN()
        Dim oFile As File
        Dim oRead As StreamReader
        Dim strLine As String = ""
        Dim arrLine As Array
        Dim dteArrival As Date
        Dim strArrival As String = ""

        oRead = oFile.OpenText(Trim(txtGENFilePath.Text))

        While oRead.Peek <> -1
            strLine = oRead.ReadLine()
            'Remove unnecessary characters
            strLine = Replace(strLine, Mid(Trim(strLine), 1, 1), " ")
            arrLine = strLine.Split(CChar(","))

            'Legend:
            'arrLine(1) = Registry No.
            'arrLine(2) = Arrival Date
            'arrLine(13) = Vessel Name
            'arrLine(18) = Voyage No.

            'Format Arrival date
            strArrival = Mid(Trim(arrLine(2).ToString), 5, 2) & "/" & Mid(Trim(arrLine(2).ToString), 7, 2) & "/" & Mid(Trim(arrLine(2).ToString), 1, 4)
            dteArrival = CType(strArrival, Date)

            'Check if Registry No. exist in CargoMHead table 
            If objUpload.Check_HeaderByRegNum(Trim(arrLine(1).ToString)) = True Then
                'Update Header data
                objUpload.Update_Header(arrLine(1).ToString, arrLine(13).ToString, _
                                        dteArrival, arrLine(18).ToString)
            Else
                'Log
                lstErrors.Items.Add("Header data with Reg. No. " & Trim(arrLine(1).ToString) & " does not exist!")
                System.Windows.Forms.Application.DoEvents()
            End If
        End While

        oFile = Nothing
        oRead = Nothing

        lstErrors.Items.Add("-- F I N I S H E D --")
    End Sub

    Private Sub frmUploading_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        StatusBarPanel1.Text = strSrvr
        StatusBarPanel2.Text = strDB
        StatusBarPanel3.Text = Now
    End Sub

    Private Function CommaDelimitedLine(ByVal CurrentLine As String, ByVal Delimiter As String, ByVal Qualifier As String) As String
        Dim i As Integer
        Dim blnStart As Boolean = False
        Dim Ch As Char
        Dim strLine As String = ""

        For i = 1 To Len(CurrentLine)
            Ch = Mid(CurrentLine, i, 1)
            If Ch = Qualifier Then
                If blnStart = True Then
                    blnStart = False
                Else
                    blnStart = True
                End If
            ElseIf Ch = Delimiter Then
                If blnStart = False Then
                    strLine = strLine & Ch
                Else
                    strLine = strLine & " "
                End If
            Else
                strLine = strLine & Ch
            End If
        Next

        CommaDelimitedLine = strLine
    End Function
End Class
