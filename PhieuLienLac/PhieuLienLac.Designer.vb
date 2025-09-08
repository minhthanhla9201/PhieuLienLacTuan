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
        Me.components = New System.ComponentModel.Container()
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
        Me.ChkInLaiPhieu = New System.Windows.Forms.CheckBox()
        Me.TxtPhieuInLai = New System.Windows.Forms.TextBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TxtPhieuLienLacPath = New System.Windows.Forms.TextBox()
        Me.BtnSelectPhieuLienLac = New System.Windows.Forms.Button()
        Me.BtnInLai = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(43, 52)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "File Tổng Hợp"
        '
        'CboFileTongHop
        '
        Me.CboFileTongHop.FormattingEnabled = True
        Me.CboFileTongHop.Location = New System.Drawing.Point(136, 49)
        Me.CboFileTongHop.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.CboFileTongHop.Name = "CboFileTongHop"
        Me.CboFileTongHop.Size = New System.Drawing.Size(522, 24)
        Me.CboFileTongHop.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(16, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(116, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Thư Mục Chứa File"
        '
        'TxtFolderPath
        '
        Me.TxtFolderPath.Location = New System.Drawing.Point(136, 17)
        Me.TxtFolderPath.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TxtFolderPath.Name = "TxtFolderPath"
        Me.TxtFolderPath.Size = New System.Drawing.Size(522, 23)
        Me.TxtFolderPath.TabIndex = 3
        '
        'BtnSelectFolder
        '
        Me.BtnSelectFolder.Location = New System.Drawing.Point(664, 15)
        Me.BtnSelectFolder.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.BtnSelectFolder.Name = "BtnSelectFolder"
        Me.BtnSelectFolder.Size = New System.Drawing.Size(35, 28)
        Me.BtnSelectFolder.TabIndex = 4
        Me.BtnSelectFolder.Text = "..."
        Me.BtnSelectFolder.UseVisualStyleBackColor = True
        '
        'BtnTaoPhieu
        '
        Me.BtnTaoPhieu.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BtnTaoPhieu.Location = New System.Drawing.Point(717, 32)
        Me.BtnTaoPhieu.Name = "BtnTaoPhieu"
        Me.BtnTaoPhieu.Size = New System.Drawing.Size(93, 56)
        Me.BtnTaoPhieu.TabIndex = 5
        Me.BtnTaoPhieu.Text = "Xuất Phiếu Liên Lạc"
        Me.BtnTaoPhieu.UseVisualStyleBackColor = True
        '
        'BtnSelectFolderExport
        '
        Me.BtnSelectFolderExport.Location = New System.Drawing.Point(664, 79)
        Me.BtnSelectFolderExport.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.BtnSelectFolderExport.Name = "BtnSelectFolderExport"
        Me.BtnSelectFolderExport.Size = New System.Drawing.Size(35, 28)
        Me.BtnSelectFolderExport.TabIndex = 8
        Me.BtnSelectFolderExport.Text = "..."
        Me.BtnSelectFolderExport.UseVisualStyleBackColor = True
        '
        'TxtExportFolder
        '
        Me.TxtExportFolder.Location = New System.Drawing.Point(136, 81)
        Me.TxtExportFolder.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TxtExportFolder.Name = "TxtExportFolder"
        Me.TxtExportFolder.Size = New System.Drawing.Size(522, 23)
        Me.TxtExportFolder.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(42, 84)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Thư Mục Xuất"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(136, 244)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(474, 23)
        Me.ProgressBar1.TabIndex = 9
        Me.ProgressBar1.Visible = False
        '
        'LblStatus
        '
        Me.LblStatus.AutoSize = True
        Me.LblStatus.ForeColor = System.Drawing.Color.RoyalBlue
        Me.LblStatus.Location = New System.Drawing.Point(133, 272)
        Me.LblStatus.Name = "LblStatus"
        Me.LblStatus.Size = New System.Drawing.Size(20, 16)
        Me.LblStatus.TabIndex = 10
        Me.LblStatus.Text = "..."
        Me.LblStatus.Visible = False
        '
        'ChkXuatPhieuLoi
        '
        Me.ChkXuatPhieuLoi.AutoSize = True
        Me.ChkXuatPhieuLoi.Location = New System.Drawing.Point(136, 111)
        Me.ChkXuatPhieuLoi.Name = "ChkXuatPhieuLoi"
        Me.ChkXuatPhieuLoi.Size = New System.Drawing.Size(122, 20)
        Me.ChkXuatPhieuLoi.TabIndex = 11
        Me.ChkXuatPhieuLoi.Text = "Xuất lại phiếu lỗi"
        Me.ChkXuatPhieuLoi.UseVisualStyleBackColor = True
        '
        'ChkInLaiPhieu
        '
        Me.ChkInLaiPhieu.AutoSize = True
        Me.ChkInLaiPhieu.Location = New System.Drawing.Point(39, 155)
        Me.ChkInLaiPhieu.Name = "ChkInLaiPhieu"
        Me.ChkInLaiPhieu.Size = New System.Drawing.Size(90, 20)
        Me.ChkInLaiPhieu.TabIndex = 12
        Me.ChkInLaiPhieu.Text = "In lại phiếu"
        Me.ChkInLaiPhieu.UseVisualStyleBackColor = True
        '
        'TxtPhieuInLai
        '
        Me.TxtPhieuInLai.Enabled = False
        Me.TxtPhieuInLai.Location = New System.Drawing.Point(136, 154)
        Me.TxtPhieuInLai.Name = "TxtPhieuInLai"
        Me.TxtPhieuInLai.Size = New System.Drawing.Size(522, 23)
        Me.TxtPhieuInLai.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.TxtPhieuInLai, "Nhập số thứ tự học sinh cách nhau bằng dấu phải để in lại")
        '
        'ToolTip1
        '
        Me.ToolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.TxtPhieuLienLacPath)
        Me.Panel1.Controls.Add(Me.TxtPhieuInLai)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.ChkInLaiPhieu)
        Me.Panel1.Controls.Add(Me.CboFileTongHop)
        Me.Panel1.Controls.Add(Me.ChkXuatPhieuLoi)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.LblStatus)
        Me.Panel1.Controls.Add(Me.TxtFolderPath)
        Me.Panel1.Controls.Add(Me.ProgressBar1)
        Me.Panel1.Controls.Add(Me.BtnSelectFolder)
        Me.Panel1.Controls.Add(Me.BtnSelectPhieuLienLac)
        Me.Panel1.Controls.Add(Me.BtnSelectFolderExport)
        Me.Panel1.Controls.Add(Me.BtnInLai)
        Me.Panel1.Controls.Add(Me.BtnTaoPhieu)
        Me.Panel1.Controls.Add(Me.TxtExportFolder)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(828, 308)
        Me.Panel1.TabIndex = 14
        '
        'TxtPhieuLienLacPath
        '
        Me.TxtPhieuLienLacPath.Enabled = False
        Me.TxtPhieuLienLacPath.Location = New System.Drawing.Point(136, 183)
        Me.TxtPhieuLienLacPath.Name = "TxtPhieuLienLacPath"
        Me.TxtPhieuLienLacPath.Size = New System.Drawing.Size(522, 23)
        Me.TxtPhieuLienLacPath.TabIndex = 14
        '
        'BtnSelectPhieuLienLac
        '
        Me.BtnSelectPhieuLienLac.Enabled = False
        Me.BtnSelectPhieuLienLac.Location = New System.Drawing.Point(664, 180)
        Me.BtnSelectPhieuLienLac.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.BtnSelectPhieuLienLac.Name = "BtnSelectPhieuLienLac"
        Me.BtnSelectPhieuLienLac.Size = New System.Drawing.Size(35, 28)
        Me.BtnSelectPhieuLienLac.TabIndex = 8
        Me.BtnSelectPhieuLienLac.Text = "..."
        Me.BtnSelectPhieuLienLac.UseVisualStyleBackColor = True
        '
        'BtnInLai
        '
        Me.BtnInLai.Cursor = System.Windows.Forms.Cursors.Hand
        Me.BtnInLai.Enabled = False
        Me.BtnInLai.Location = New System.Drawing.Point(717, 146)
        Me.BtnInLai.Name = "BtnInLai"
        Me.BtnInLai.Size = New System.Drawing.Size(93, 56)
        Me.BtnInLai.TabIndex = 5
        Me.BtnInLai.Text = "In Lại Phiếu Liên Lạc"
        Me.BtnInLai.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(11, 186)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(118, 16)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "File liên lạc đã xuất"
        '
        'PhieuLienLac
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(828, 308)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "PhieuLienLac"
        Me.Text = "Phiếu Liên Lạc Tuần"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

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
    Friend WithEvents ChkInLaiPhieu As CheckBox
    Friend WithEvents TxtPhieuInLai As TextBox
    Friend WithEvents ToolTip1 As ToolTip
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label4 As Label
    Friend WithEvents TxtPhieuLienLacPath As TextBox
    Friend WithEvents BtnSelectPhieuLienLac As Button
    Friend WithEvents BtnInLai As Button
End Class
