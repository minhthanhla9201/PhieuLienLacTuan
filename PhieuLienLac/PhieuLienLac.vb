Imports System.IO

Public Class PhieuLienLac
    Private configManager As ConfigManager

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboFileTongHop.SelectedIndexChanged
        If Not String.IsNullOrEmpty(CboFileTongHop.SelectedItem) Then
            Dim tuan As String = Path.GetFileNameWithoutExtension(CboFileTongHop.SelectedItem).Split(" ")(1).Trim
            TxtExportFolder.Text = Path.Combine(configManager.GetConfig.OutputFolder, "Tuan " & tuan)
        Else
            TxtExportFolder.Text = configManager.GetConfig.OutputFolder
        End If
    End Sub

    Private Sub PhieuLienLac_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Application.EnableVisualStyles()
        configManager = New ConfigManager()

        Try
            'Load config
            configManager.LoadConfig()
            TxtFolderPath.Text = configManager.GetConfig.InputFolder
            TxtExportFolder.Text = configManager.GetConfig.OutputFolder

            LoadComboboxPhieuTongHop()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Lỗi")
        End Try
    End Sub

    Private Sub BtnSelectFolder_Click(sender As Object, e As EventArgs) Handles BtnSelectFolder.Click
        Using dialog As New FolderBrowserDialog()
            dialog.Description = "Chọn thư mục chứa file"
            dialog.ShowNewFolderButton = True

            If dialog.ShowDialog() = DialogResult.OK Then
                Dim selectedPath As String = dialog.SelectedPath
                TxtFolderPath.Text = selectedPath

                'Load lại danh sách khi chọn lại folder
                LoadComboboxPhieuTongHop()
            End If
        End Using
    End Sub

    Private Sub BtnSelectFolderExport_Click(sender As Object, e As EventArgs) Handles BtnSelectFolderExport.Click
        Using dialog As New FolderBrowserDialog()
            dialog.Description = "Chọn thư mục để lưu file"
            dialog.ShowNewFolderButton = True

            If dialog.ShowDialog() = DialogResult.OK Then
                Dim selectedPath As String = dialog.SelectedPath
                TxtExportFolder.Text = selectedPath
            End If
        End Using
    End Sub

    Private Sub BtnTaoPhieu_Click(sender As Object, e As EventArgs) Handles BtnTaoPhieu.Click
        XuatPhieuLienLacTuan()
    End Sub

#Region "Load dữ liệu"
    Private Sub LoadComboboxPhieuTongHop()
        CboFileTongHop.Items.Clear()

        Dim folderPath = TxtFolderPath.Text
        If Directory.Exists(folderPath) Then
            Dim excelFiles = Directory.GetFiles(TxtFolderPath.Text, "*.xls*") ' .xls và .xlsx

            ' Đưa danh sách file vào combobox
            For Each file In excelFiles
                CboFileTongHop.Items.Add(Path.GetFileName(file))
            Next
        End If
    End Sub
#End Region

#Region "Xử lý xuất dữ liệu"
    Private Async Sub XuatPhieuLienLacTuan()
        Dim excelReader As New ExcelManager
        Try
            If Not CheckInput() Then
                Exit Sub
            End If

            If Directory.Exists(TxtExportFolder.Text) Then
                Dim excelFiles = Directory.GetFiles(TxtExportFolder.Text, "*.jpeg*")
                If excelFiles.Count > 0 Then
                    Dim result As DialogResult = MessageBox.Show("Thư mục xuất đã có file xuất trước đó, nếu tiếp tục thì file cũ sẽ mất. Bạn muốn tiếp tục không?",
                                                                 "Xác nhận",
                                                                 MessageBoxButtons.YesNo,
                                                                 MessageBoxIcon.Question)

                    If result <> DialogResult.Yes Then
                        Exit Sub
                    End If
                End If
            End If

            Me.Cursor = Cursors.WaitCursor
            Me.Enabled = False
            LblStatus.Text = $"Đang xử lý đọc file liên lạc tổng hợp..."
            ProgressBar1.Visible = True
            LblStatus.Visible = True

            Dim stuList As List(Of Student) = excelReader.ReadExcelFile(Path.Combine(TxtFolderPath.Text, CboFileTongHop.SelectedItem))

            LblStatus.Text = $"Bắt đầu xử lý tạo phiếu liên lạc"
            ProgressBar1.Maximum = stuList.Count
            ProgressBar1.Value = 0

            Dim status As String = "Đang tạo phiếu liên lạc"
            Dim progress = New Progress(Of Integer)(
            Sub(value)
                ProgressBar1.Value = value
                LblStatus.Text = $"{status} {value}/{stuList.Count}"
                If value = stuList.Count Then
                    status = "Đang xuất hình chụp phiếu"
                End If
            End Sub)

            'Tạo folder nếu chưa có
            If Not Directory.Exists(TxtExportFolder.Text) Then
                Directory.CreateDirectory(TxtExportFolder.Text)
            End If

            Dim outputPath As String = Path.Combine(TxtExportFolder.Text, $"PhieuLienLac_Tuan {stuList(0).Tuan}.xlsx")
            If ChkXuatPhieuLoi.Checked AndAlso File.Exists(outputPath) Then
                Await excelReader.ExportPhieuLoi(configManager.GetConfig.TemplatePath, outputPath, stuList, progress)
            Else
                Await excelReader.ExportToTemplate(configManager.GetConfig.TemplatePath, outputPath,
                                                   stuList, progress, ChkXuatPhieuLoi.Checked)
            End If

            'Lưu lại config để xuất lần sau
            configManager.GetConfig.InputFolder = TxtFolderPath.Text
            configManager.GetConfig.OutputFolder = Path.GetDirectoryName(TxtExportFolder.Text)
            configManager.SaveConfig()
        Catch ex As Exception
            MsgBox(ex.Message & vbNewLine & ex.StackTrace, MsgBoxStyle.Critical, "Lỗi")
        Finally
            Me.Enabled = True
            Me.Cursor = Cursors.Default
            ProgressBar1.Visible = False
            LblStatus.Visible = False
        End Try
    End Sub

    Private Function CheckInput() As Boolean
        If String.IsNullOrEmpty(TxtFolderPath.Text) Then
            MsgBox("Chưa chọn thư mục chứa file liên lạc tổng hợp", MsgBoxStyle.Exclamation)
            TxtFolderPath.Focus()
            Return False
        End If

        If String.IsNullOrEmpty(CboFileTongHop.SelectedItem) Then
            MsgBox("Chưa chọn file liên lạc tổng hợp", MsgBoxStyle.Exclamation)
            CboFileTongHop.Focus()
            Return False
        End If

        If Not File.Exists(Path.Combine(TxtFolderPath.Text, CboFileTongHop.SelectedItem)) Then
            MsgBox("File liên lạc không tồn tại hoặc đã bị xóa", MsgBoxStyle.Exclamation)
            CboFileTongHop.Focus()
            Return False
        End If

        If String.IsNullOrEmpty(TxtExportFolder.Text) Then
            MsgBox("Chưa chọn thư mục để lưu phiếu xuất", MsgBoxStyle.Exclamation)
            TxtExportFolder.Focus()
            Return False
        End If

        Return True
    End Function
#End Region
End Class
