<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PhieuLienLac
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CboFileTongHop = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TxtFolderPath = New System.Windows.Forms.TextBox()
        Me.BtnSelectFolder = New System.Windows.Forms.Button()
        Me.BtnTaoPhieu = New System.Windows.Forms.Button()
        Me.BtnSelectFolderExport = New System.Windows.Forms.Button()
        Me.TxtExportFolder = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.LblStatus = New System.Windows.Forms.Label()
        Me.ChkXuatPhieuLoi = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(47, 50)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "File Tổng Hợp"
        '
        'CboFileTongHop
        '
        Me.CboFileTongHop.FormattingEnabled = True
        Me.CboFileTongHop.Location = New System.Drawing.Point(140, 47)
        Me.CboFileTongHop.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CboFileTongHop.Name = "CboFileTongHop"
        Me.CboFileTongHop.Size = New System.Drawing.Size(522, 24)
        Me.CboFileTongHop.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(20, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(116, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Thư Mục Chứa File"
        '
        'TxtFolderPath
        '
        Me.TxtFolderPath.Location = New System.Drawing.Point(140, 15)
        Me.TxtFolderPath.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TxtFolderPath.Name = "TxtFolderPath"
        Me.TxtFolderPath.Size = New System.Drawing.Size(522, 23)
        Me.TxtFolderPath.TabIndex = 3
        '
        'BtnSelectFolder
        '
        Me.BtnSelectFolder.Location = New System.Drawing.Point(668, 13)
        Me.BtnSelectFolder.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.BtnSelectFolder.Name = "BtnSelectFolder"
        Me.BtnSelectFolder.Size = New System.Drawing.Size(35, 28)
        Me.BtnSelectFolder.TabIndex = 4
        Me.BtnSelectFolder.Text = "..."
        Me.BtnSelectFolder.UseVisualStyleBackColor = True
        '
        'BtnTaoPhieu
        '
        Me.BtnTaoPhieu.Location = New System.Drawing.Point(610, 125)
        Me.BtnTaoPhieu.Name = "BtnTaoPhieu"
        Me.BtnTaoPhieu.Size = New System.Drawing.Size(93, 56)
        Me.BtnTaoPhieu.TabIndex = 5
        Me.BtnTaoPhieu.Text = "Xuất Phiếu Liên Lạc"
        Me.BtnTaoPhieu.UseVisualStyleBackColor = True
        '
        'BtnSelectFolderExport
        '
        Me.BtnSelectFolderExport.Location = New System.Drawing.Point(668, 77)
        Me.BtnSelectFolderExport.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.BtnSelectFolderExport.Name = "BtnSelectFolderExport"
        Me.BtnSelectFolderExport.Size = New System.Drawing.Size(35, 28)
        Me.BtnSelectFolderExport.TabIndex = 8
        Me.BtnSelectFolderExport.Text = "..."
        Me.BtnSelectFolderExport.UseVisualStyleBackColor = True
        '
        'TxtExportFolder
        '
        Me.TxtExportFolder.Location = New System.Drawing.Point(140, 79)
        Me.TxtExportFolder.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TxtExportFolder.Name = "TxtExportFolder"
        Me.TxtExportFolder.Size = New System.Drawing.Size(522, 23)
        Me.TxtExportFolder.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(46, 82)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Thư Mục Xuất"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(140, 206)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(474, 23)
        Me.ProgressBar1.TabIndex = 9
        Me.ProgressBar1.Visible = False
        '
        'LblStatus
        '
        Me.LblStatus.AutoSize = True
        Me.LblStatus.ForeColor = System.Drawing.Color.RoyalBlue
        Me.LblStatus.Location = New System.Drawing.Point(137, 234)
        Me.LblStatus.Name = "LblStatus"
        Me.LblStatus.Size = New System.Drawing.Size(20, 16)
        Me.LblStatus.TabIndex = 10
        Me.LblStatus.Text = "..."
        Me.LblStatus.Visible = False
        '
        'ChkXuatPhieuLoi
        '
        Me.ChkXuatPhieuLoi.AutoSize = True
        Me.ChkXuatPhieuLoi.Location = New System.Drawing.Point(140, 125)
        Me.ChkXuatPhieuLoi.Name = "ChkXuatPhieuLoi"
        Me.ChkXuatPhieuLoi.Size = New System.Drawing.Size(122, 20)
        Me.ChkXuatPhieuLoi.TabIndex = 11
        Me.ChkXuatPhieuLoi.Text = "Xuất lại phiếu lỗi"
        Me.ChkXuatPhieuLoi.UseVisualStyleBackColor = True
        '
        'PhieuLienLac
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(755, 259)
        Me.Controls.Add(Me.ChkXuatPhieuLoi)
        Me.Controls.Add(Me.LblStatus)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.BtnSelectFolderExport)
        Me.Controls.Add(Me.TxtExportFolder)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.BtnTaoPhieu)
        Me.Controls.Add(Me.BtnSelectFolder)
        Me.Controls.Add(Me.TxtFolderPath)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CboFileTongHop)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "PhieuLienLac"
        Me.Text = "Phiếu Liên Lạc"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents CboFileTongHop As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents TxtFolderPath As TextBox
    Friend WithEvents BtnSelectFolder As Button
    Friend WithEvents BtnTaoPhieu As Button
    Friend WithEvents BtnSelectFolderExport As Button
    Friend WithEvents TxtExportFolder As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents LblStatus As Label
    Friend WithEvents ChkXuatPhieuLoi As CheckBox
End Class
