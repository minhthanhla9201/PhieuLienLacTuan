Public Class Student
    Public Property Tuan As String
    Public Property STT As Integer
    Public Property HoTen As String
    Public Property GioiTinh As String
    Public Property NgaySinh As String
    Public Property NoiSinh As String
    Public Property Lop As String
    Public Property DanToc As String
    Public Property NhanXetKetQuaHocTap As String
    Public Property NhanXetKetQuaRenLuyen As String
    Public Property TongSoNgayVang As Integer
    Public Property KetQuaHocTap As String
    Public Property KetQuaRenLuyen As String
    Public Property NhanXetGVCN As String
    Public Property PhieuLoi As String
    Public Function XuatLai() As Boolean
        If Not String.IsNullOrEmpty(PhieuLoi) Then
            Return True
        End If
        Return False
    End Function
    Public Overrides Function ToString() As String
        Return $"[{STT}] {HoTen}, {GioiTinh}, NS: {NgaySinh}, " &
               $"Nơi sinh: {NoiSinh}, Lớp: {Lop}, Dân tộc: {DanToc}, " &
               $"NX HT: {NhanXetKetQuaHocTap}, NX RL: {NhanXetKetQuaRenLuyen}, " &
               $"Vắng: {TongSoNgayVang}, HT: {KetQuaHocTap}, RL: {KetQuaRenLuyen}, " &
               $"Nhận xét: {NhanXetGVCN}, Xuất lại: {PhieuLoi}"
    End Function
End Class
