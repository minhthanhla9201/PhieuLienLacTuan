Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.IO

Public Class ExcelManager
    Public Const INPUT_ROWINDEX_START_HYPERLINK As Integer = 2  'Dòng bắt đầu xuất hyperlink
    Public Const INPUT_ROWINDEX_START As Integer = 4    'Dòng bắt đầu đọc nội dung
    Public Const INPUT_ROWINDEX_NXKQHT As Integer = 13  'Nhận xét KQHT
    Public Const INPUT_ROWINDEX_NXKQRL As Integer = 15  'Nhận xét KQRL
    Public Const INPUT_ROWINDEX_NXGVCN As Integer = 21  'Nhận xét GVCN

    ''' <summary>
    ''' Đọc thông tin phiếu liên lạc
    ''' </summary>
    ''' <param name="filePath"></param>
    Public Function ReadExcelFile(filePath As String) As List(Of Student)
        Dim excelApp As Excel.Application = Nothing
        Dim workbook As Excel.Workbook = Nothing
        Dim worksheet As Excel.Worksheet = Nothing
        Dim sheets As Excel.Sheets = Nothing
        Dim range As Excel.Range = Nothing

        Dim studentList As New List(Of Student)

        Try
            ' Mở Excel Application (ẩn đi)
            excelApp = New Excel.Application()
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            ' Mở file Excel
            workbook = excelApp.Workbooks.Open(filePath)
            sheets = workbook.Sheets
            worksheet = sheets(1) ' Lấy sheet đầu tiên
            range = worksheet.UsedRange    ' Lấy vùng có dữ liệu

            Dim rowCount As Integer = range.Rows.Count
            Dim colCount As Integer = range.Columns.Count

            'Xử lý lấy tuần
            Dim tuan As String = Path.GetFileNameWithoutExtension(filePath).Split(" ")(1).Trim

            ' Read data
            For rowIndex As Integer = INPUT_ROWINDEX_START To rowCount
                'Stop if not exist STT
                If worksheet.Range($"A{rowIndex}").Value2 Is Nothing Then
                    Exit For
                End If

                Dim s As New Student With {
                    .Tuan = tuan,
                    .STT = If(worksheet.Range($"A{rowIndex}").Value2 IsNot Nothing, CInt(worksheet.Range($"A{rowIndex}").Value2), 0),
                    .HoTen = Convert.ToString(worksheet.Range($"B{rowIndex}").Value2),
                    .GioiTinh = Convert.ToString(worksheet.Range($"C{rowIndex}").Value2),
                    .NgaySinh = Convert.ToString(worksheet.Range($"D{rowIndex}").Value2),
                    .NoiSinh = Convert.ToString(worksheet.Range($"E{rowIndex}").Value2),
                    .Lop = Convert.ToString(worksheet.Range($"F{rowIndex}").Value2),
                    .DanToc = Convert.ToString(worksheet.Range($"G{rowIndex}").Value2),
                    .NhanXetKetQuaHocTap = Convert.ToString(worksheet.Range($"H{rowIndex}").Value2),
                    .NhanXetKetQuaRenLuyen = Convert.ToString(worksheet.Range($"I{rowIndex}").Value2),
                    .TongSoNgayVang = If(worksheet.Range($"J{rowIndex}").Value2 IsNot Nothing, CInt(worksheet.Range($"J{rowIndex}").Value2), 0),
                    .KetQuaHocTap = Convert.ToString(worksheet.Range($"K{rowIndex}").Value2),
                    .KetQuaRenLuyen = Convert.ToString(worksheet.Range($"L{rowIndex}").Value2),
                    .NhanXetGVCN = Convert.ToString(worksheet.Range($"M{rowIndex}").Value2),
                    .PhieuLoi = Convert.ToString(worksheet.Range($"N{rowIndex}").Value2)
                }

                studentList.Add(s)
                Console.WriteLine(s.ToString)
            Next

            Return studentList
        Catch ex As Exception
            Throw ex
        Finally
            If range IsNot Nothing Then
                Marshal.ReleaseComObject(range)
            End If

            If worksheet IsNot Nothing Then
                Marshal.ReleaseComObject(worksheet)
            End If

            If sheets IsNot Nothing Then
                Marshal.ReleaseComObject(sheets)
            End If

            If workbook IsNot Nothing Then
                workbook.Close(False)
                Marshal.ReleaseComObject(workbook)
            End If

            If excelApp IsNot Nothing Then
                excelApp.Quit()
                Marshal.ReleaseComObject(excelApp)
            End If

            ' Buộc Garbage Collection dọn dẹp
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    Public Async Function ExportToTemplate(templatePath As String, outputPath As String, studenList As List(Of Student),
                                           progress As IProgress(Of Integer), ByVal xuatPhieuLoi As Boolean) As Task
        Await Task.Run(Sub()
                           Dim excelApp As Excel.Application = Nothing
                           Dim workbook As Excel.Workbook = Nothing
                           Dim templateSheet As Excel.Worksheet = Nothing
                           Dim hyperlinkSheet As Excel.Worksheet = Nothing
                           Dim sheets As Excel.Sheets = Nothing

                           Dim printRange As Excel.Range = Nothing
                           Dim chartObj As Excel.ChartObject = Nothing
                           Dim chart As Excel.Chart = Nothing

                           Try
                               ' Mở Excel
                               excelApp = New Excel.Application()
                               excelApp.Visible = False
                               excelApp.DisplayAlerts = False
                               excelApp.ErrorCheckingOptions.NumberAsText = False

                               ' Mở template
                               Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
                               workbook = excelApp.Workbooks.Open(Path.Combine(baseDir, templatePath))
                               sheets = workbook.Sheets
                               templateSheet = CType(sheets("Template"), Excel.Worksheet)
                               hyperlinkSheet = CType(sheets("DanhSach"), Excel.Worksheet)

                               For i As Integer = 0 To studenList.Count - 1
                                   ' Copy sheet mẫu
                                   Dim newSheet As Excel.Worksheet = CType(workbook.Sheets("Template"), Excel.Worksheet)
                                   newSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                                   Dim stuSheet As Excel.Worksheet = CType(workbook.Sheets(workbook.Sheets.Count), Excel.Worksheet)

                                   Dim stu = studenList(i)
                                   ' Đặt tên sheet theo học sinh
                                   Try
                                       stuSheet.Name = stu.HoTen.Substring(0, Math.Min(25, stu.HoTen.Length))
                                   Catch
                                       stuSheet.Name = "HS" & stu.STT
                                   End Try

                                   ' Ghi dữ liệu vào các ô (ví dụ theo vị trí trong template)
                                   stuSheet.Range("L4").Value = "'" & stu.Tuan.PadLeft(2, "0")
                                   stuSheet.Range("D6").Value = stu.HoTen
                                   stuSheet.Range("D7").Value = stu.GioiTinh
                                   stuSheet.Range("D8").Value = stu.NgaySinh
                                   stuSheet.Range("M6").Value = stu.Lop
                                   stuSheet.Range("M7").Value = stu.DanToc
                                   stuSheet.Range("M8").Value = stu.NoiSinh
                                   stuSheet.Range("A13").WrapText = False
                                   stuSheet.Range("A15").WrapText = False
                                   stuSheet.Range("F16").Value = stu.TongSoNgayVang
                                   stuSheet.Range("F18").Value = stu.KetQuaHocTap
                                   stuSheet.Range("F19").Value = stu.KetQuaRenLuyen
                                   stuSheet.Range("A21").WrapText = False

                                   Dim nxGVCN() As String = stu.NhanXetGVCN.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                                   writeNhanXet(stuSheet, INPUT_ROWINDEX_NXGVCN, nxGVCN)

                                   Dim nxKQRL() As String = stu.NhanXetKetQuaRenLuyen.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                                   writeNhanXet(stuSheet, INPUT_ROWINDEX_NXKQRL, nxKQRL)

                                   Dim nxKQHT() As String = stu.NhanXetKetQuaHocTap.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                                   writeNhanXet(stuSheet, INPUT_ROWINDEX_NXKQHT, nxKQHT)

                                   hyperlinkSheet.Range($"A{INPUT_ROWINDEX_START_HYPERLINK + i}").Value = stu.STT
                                   hyperlinkSheet.Range($"B{INPUT_ROWINDEX_START_HYPERLINK + i}").Value = stu.HoTen
                                   hyperlinkSheet.Hyperlinks.Add(
                                       Anchor:=hyperlinkSheet.Range($"B{INPUT_ROWINDEX_START_HYPERLINK + i}"),
                                       Address:="",
                                       SubAddress:=$"'{stu.HoTen}'!A1",
                                       ScreenTip:=$"Nhảy đến sheet {stu.HoTen}",
                                       TextToDisplay:=stu.HoTen)

                                   progress.Report(i + 1)
                               Next

                               'Xóa sheet template
                               templateSheet.Delete()

                               hyperlinkSheet.Activate()

                               ' Lưu ra file mới
                               workbook.SaveAs(outputPath)

                               Console.WriteLine("Xuất dữ liệu ra Excel thành công: " & outputPath)

                               Console.WriteLine("Bắt đầu xuất hình ảnh")
                               workbook = excelApp.Workbooks.Open(outputPath)
                               sheets = workbook.Sheets

                               For i As Integer = 0 To studenList.Count - 1
                                   Dim stuSheet As Excel.Worksheet = CType(sheets(studenList(i).HoTen), Excel.Worksheet)
                                   stuSheet.Activate()

                                   Dim printArea As String = stuSheet.PageSetup.PrintArea
                                   If String.IsNullOrEmpty(printArea) Then
                                       ' Nếu không có vùng in, sử dụng UsedRange hoặc vùng mặc định
                                       printRange = stuSheet.UsedRange
                                       Console.WriteLine("Không có vùng in được thiết lập. Sử dụng UsedRange: " & printRange.Address)
                                   Else
                                       printRange = stuSheet.Range(printArea)
                                       Console.WriteLine("Sử dụng vùng in: " & printArea)
                                   End If

                                   Dim maxRetries As Integer = 3
                                   Dim retryCount As Integer = 0
                                   Dim success As Boolean = False

                                   While retryCount < maxRetries AndAlso Not success
                                       Try
                                           printRange.CopyPicture(Excel.XlPictureAppearance.xlPrinter, Excel.XlCopyPictureFormat.xlPicture)
                                           success = True
                                       Catch ex As Exception
                                           retryCount += 1
                                           Console.WriteLine($"Lỗi CopyPicture (thử {retryCount}/{maxRetries}): {ex.Message}")
                                           If retryCount = maxRetries Then
                                               Throw New Exception($"Không thể sao chép vùng in cho học sinh {studenList(i).HoTen} sau {maxRetries} lần thử", ex)
                                           End If
                                           Threading.Thread.Sleep(500) ' Chờ 500ms trước khi thử lại
                                       End Try
                                   End While

                                   ' Tạo chart tạm để chứa ảnh
                                   Threading.Thread.Sleep(1000)
                                   chartObj = stuSheet.ChartObjects().Add(0, 0, printRange.Width, printRange.Height)
                                   chart = chartObj.Chart
                                   chart.Paste()

                                   ' Xuất chart thành file ảnh
                                   Dim imageFile As String = Path.Combine(Path.GetDirectoryName(outputPath), $"{studenList(i).STT.ToString.PadLeft(2, "0") & "_" & studenList(i).HoTen}.jpeg")
                                   chart.Export(Filename:=imageFile, FilterName:="JPEG")
                                   Console.WriteLine("Xuất hình ảnh thành công: " & imageFile)

                                   progress.Report(i + 1)
                               Next

                           Catch ex As Exception
                               Console.WriteLine("Lỗi: " & ex.Message)
                               Throw ex
                           Finally
                               If chart IsNot Nothing Then
                                   Marshal.ReleaseComObject(chart)
                               End If
                               If chartObj IsNot Nothing Then
                                   Marshal.ReleaseComObject(chartObj)
                               End If
                               If printRange IsNot Nothing Then
                                   Marshal.ReleaseComObject(printRange)
                               End If

                               If templateSheet IsNot Nothing Then
                                   Marshal.ReleaseComObject(templateSheet)
                               End If

                               If sheets IsNot Nothing Then
                                   Marshal.ReleaseComObject(sheets)
                               End If

                               If workbook IsNot Nothing Then
                                   workbook.Close(False)
                                   Marshal.ReleaseComObject(workbook)
                               End If

                               If excelApp IsNot Nothing Then
                                   excelApp.Quit()
                                   Marshal.ReleaseComObject(excelApp)
                               End If

                               ' Buộc Garbage Collection dọn dẹp
                               GC.Collect()
                               GC.WaitForPendingFinalizers()
                           End Try
                       End Sub)
    End Function

    Public Async Function ExportPhieuLoi(ByVal templatePath As String, ByVal outputPath As String, ByVal studenList As List(Of Student),
                                         progress As IProgress(Of Integer)) As Task
        Await Task.Run(Sub()
                           Dim excelApp As Excel.Application = Nothing
                           Dim workbook As Excel.Workbook = Nothing
                           Dim hyperlinkSheet As Excel.Worksheet = Nothing
                           Dim sheets As Excel.Sheets = Nothing

                           Dim templateWorkbook As Excel.Workbook = Nothing
                           Dim templateSheet As Excel.Worksheet = Nothing

                           Dim printRange As Excel.Range = Nothing
                           Dim chartObj As Excel.ChartObject = Nothing
                           Dim chart As Excel.Chart = Nothing

                           Try
                               ' Mở Excel
                               excelApp = New Excel.Application()
                               excelApp.Visible = False
                               excelApp.DisplayAlerts = False
                               excelApp.ErrorCheckingOptions.NumberAsText = False

                               workbook = excelApp.Workbooks.Open(outputPath)
                               sheets = workbook.Sheets

                               Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
                               templateWorkbook = excelApp.Workbooks.Open(Path.Combine(baseDir, templatePath))
                               templateSheet = CType(templateWorkbook.Sheets("Template"), Excel.Worksheet)

                               For i As Integer = 0 To studenList.Count - 1
                                   If Not studenList(i).XuatLai Then
                                       Continue For
                                   End If

                                   Dim stu = studenList(i)
                                   Dim stuSheet As Excel.Worksheet = CType(sheets(stu.HoTen), Excel.Worksheet)
                                   stuSheet.Delete()

                                   templateSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                                   stuSheet = CType(workbook.Sheets(workbook.Sheets.Count), Excel.Worksheet)
                                   stuSheet.Name = stu.HoTen

                                   ' Ghi dữ liệu vào các ô (ví dụ theo vị trí trong template)
                                   stuSheet.Range("L4").Value = "'" & stu.Tuan.PadLeft(2, "0")
                                   stuSheet.Range("D6").Value = stu.HoTen
                                   stuSheet.Range("D7").Value = stu.GioiTinh
                                   stuSheet.Range("D8").Value = stu.NgaySinh
                                   stuSheet.Range("M6").Value = stu.Lop
                                   stuSheet.Range("M7").Value = stu.DanToc
                                   stuSheet.Range("M8").Value = stu.NoiSinh
                                   stuSheet.Range("A13").WrapText = False
                                   stuSheet.Range("A15").WrapText = False
                                   stuSheet.Range("F16").Value = stu.TongSoNgayVang
                                   stuSheet.Range("F18").Value = stu.KetQuaHocTap
                                   stuSheet.Range("F19").Value = stu.KetQuaRenLuyen
                                   stuSheet.Range("A21").WrapText = False

                                   Dim nxGVCN() As String = stu.NhanXetGVCN.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                                   writeNhanXet(stuSheet, INPUT_ROWINDEX_NXGVCN, nxGVCN)

                                   Dim nxKQRL() As String = stu.NhanXetKetQuaRenLuyen.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                                   writeNhanXet(stuSheet, INPUT_ROWINDEX_NXKQRL, nxKQRL)

                                   Dim nxKQHT() As String = stu.NhanXetKetQuaHocTap.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                                   writeNhanXet(stuSheet, INPUT_ROWINDEX_NXKQHT, nxKQHT)

                                   stuSheet.Activate()
                                   Threading.Thread.Sleep(500)

                                   Dim printArea As String = stuSheet.PageSetup.PrintArea
                                   If String.IsNullOrEmpty(printArea) Then
                                       ' Nếu không có vùng in, sử dụng UsedRange hoặc vùng mặc định
                                       printRange = stuSheet.UsedRange
                                       Console.WriteLine("Không có vùng in được thiết lập. Sử dụng UsedRange: " & printRange.Address)
                                   Else
                                       printRange = stuSheet.Range(printArea)
                                       Console.WriteLine("Sử dụng vùng in: " & printArea)
                                   End If

                                   Dim maxRetries As Integer = 3
                                   Dim retryCount As Integer = 0
                                   Dim success As Boolean = False

                                   While retryCount < maxRetries AndAlso Not success
                                       Try
                                           printRange.CopyPicture(Excel.XlPictureAppearance.xlPrinter, Excel.XlCopyPictureFormat.xlPicture)
                                           success = True
                                       Catch ex As Exception
                                           retryCount += 1
                                           Console.WriteLine($"Lỗi CopyPicture (thử {retryCount}/{maxRetries}): {ex.Message}")
                                           If retryCount = maxRetries Then
                                               Throw New Exception($"Không thể sao chép vùng in cho học sinh {stu.HoTen} sau {maxRetries} lần thử", ex)
                                           End If
                                           Threading.Thread.Sleep(500) ' Chờ 500ms trước khi thử lại
                                       End Try
                                   End While

                                   ' Tạo chart tạm để chứa ảnh
                                   Threading.Thread.Sleep(500)
                                   chartObj = stuSheet.ChartObjects().Add(0, 0, printRange.Width, printRange.Height)
                                   chart = chartObj.Chart
                                   chart.Paste()

                                   ' Xuất chart thành file ảnh
                                   Dim imageFile As String = Path.Combine(Path.GetDirectoryName(outputPath), $"{stu.STT.ToString.PadLeft(2, "0") & "_" & stu.HoTen}.jpeg")
                                   chart.Export(Filename:=imageFile, FilterName:="JPEG")
                                   Console.WriteLine("Xuất hình ảnh thành công: " & imageFile)

                                   progress.Report(i + 1)
                               Next

                               ' Lưu ra file mới
                               workbook.Save()

                               Console.WriteLine("Xuất dữ liệu lại ra Excel thành công: " & outputPath)

                           Catch ex As Exception
                               Console.WriteLine("Lỗi: " & ex.Message)
                               Throw ex
                           Finally
                               If chart IsNot Nothing Then
                                   Marshal.ReleaseComObject(chart)
                               End If
                               If chartObj IsNot Nothing Then
                                   Marshal.ReleaseComObject(chartObj)
                               End If
                               If printRange IsNot Nothing Then
                                   Marshal.ReleaseComObject(printRange)
                               End If

                               If templateSheet IsNot Nothing Then
                                   Marshal.ReleaseComObject(templateSheet)
                               End If

                               If sheets IsNot Nothing Then
                                   Marshal.ReleaseComObject(sheets)
                               End If

                               If templateWorkbook IsNot Nothing Then
                                   templateWorkbook.Close(False)
                                   Marshal.ReleaseComObject(templateWorkbook)
                               End If

                               If workbook IsNot Nothing Then
                                   workbook.Close(False)
                                   Marshal.ReleaseComObject(workbook)
                               End If

                               If excelApp IsNot Nothing Then
                                   excelApp.Quit()
                                   Marshal.ReleaseComObject(excelApp)
                               End If

                               ' Buộc Garbage Collection dọn dẹp
                               GC.Collect()
                               GC.WaitForPendingFinalizers()
                           End Try

                       End Sub)
    End Function

    Public Async Function InPhieuLai(ByVal phieuLienLacPath As String, ByVal studenList As String(),
                                     progress As IProgress(Of Integer)) As Task
        Await Task.Run(Sub()
                           Dim excelApp As Excel.Application = Nothing
                           Dim workbook As Excel.Workbook = Nothing
                           Dim sheets As Excel.Sheets = Nothing

                           Dim printRange As Excel.Range = Nothing
                           Dim chartObj As Excel.ChartObject = Nothing
                           Dim chart As Excel.Chart = Nothing

                           Try
                               ' Mở Excel
                               excelApp = New Excel.Application()
                               excelApp.Visible = False
                               excelApp.DisplayAlerts = False
                               excelApp.ErrorCheckingOptions.NumberAsText = False

                               workbook = excelApp.Workbooks.Open(phieuLienLacPath)
                               sheets = workbook.Sheets

                               Dim tempSheet As Excel.Worksheet = CType(sheets("DanhSach"), Excel.Worksheet)
                               Dim range As Excel.Range = tempSheet.UsedRange

                               Dim rowCount As Integer = range.Rows.Count

                               'Them 0 nếu nhập chưa đủ
                               For i As Integer = 0 To studenList.Count - 1
                                   studenList(i) = studenList(i).PadLeft(2, "0")
                               Next

                               Dim students As New List(Of Student)
                               For rowIndex As Integer = 1 To rowCount
                                   Dim sttCheck As String = Convert.ToString(tempSheet.Range("A" & rowIndex).Value2)
                                   If sttCheck IsNot Nothing AndAlso studenList.Contains(sttCheck.PadLeft(2, "0")) Then
                                       Dim stu As New Student
                                       stu.STT = sttCheck
                                       stu.HoTen = Convert.ToString(tempSheet.Range("B" & rowIndex).Value2)
                                       students.Add(stu)
                                   End If
                               Next

                               For Each s As Student In students
                                   tempSheet = CType(sheets(s.HoTen), Excel.Worksheet)
                                   tempSheet.Activate()

                                   Threading.Thread.Sleep(500)

                                   Dim printArea As String = tempSheet.PageSetup.PrintArea
                                   If String.IsNullOrEmpty(printArea) Then
                                       printRange = tempSheet.UsedRange
                                   Else
                                       printRange = tempSheet.Range(printArea)
                                   End If

                                   Dim maxRetries As Integer = 3
                                   Dim retryCount As Integer = 0
                                   Dim success As Boolean = False

                                   While retryCount < maxRetries AndAlso Not success
                                       Try
                                           printRange.CopyPicture(Excel.XlPictureAppearance.xlPrinter, Excel.XlCopyPictureFormat.xlPicture)
                                           success = True
                                       Catch ex As Exception
                                           retryCount += 1
                                           Console.WriteLine($"Lỗi CopyPicture (thử {retryCount}/{maxRetries}): {ex.Message}")
                                           If retryCount = maxRetries Then
                                               Throw New Exception($"Không thể sao chép vùng in cho học sinh {s.HoTen} sau {maxRetries} lần thử", ex)
                                           End If
                                           Threading.Thread.Sleep(500) ' Chờ 500ms trước khi thử lại
                                       End Try
                                   End While

                                   ' Tạo chart tạm để chứa ảnh
                                   Threading.Thread.Sleep(500)
                                   chartObj = tempSheet.ChartObjects().Add(0, 0, printRange.Width, printRange.Height)
                                   chart = chartObj.Chart
                                   chart.Paste()

                                   ' Xuất chart thành file ảnh
                                   Dim imageFile As String = Path.Combine(Path.GetDirectoryName(phieuLienLacPath), $"{s.STT.ToString.PadLeft(2, "0") & "_" & s.HoTen}.jpeg")
                                   chart.Export(Filename:=imageFile, FilterName:="JPEG")
                                   Console.WriteLine("Xuất hình ảnh thành công: " & imageFile)
                               Next

                               Console.WriteLine("In lại file thành công: " & phieuLienLacPath)

                           Catch ex As Exception
                               Console.WriteLine("Lỗi: " & ex.Message)
                               Throw ex
                           Finally
                               If chart IsNot Nothing Then
                                   Marshal.ReleaseComObject(chart)
                               End If
                               If chartObj IsNot Nothing Then
                                   Marshal.ReleaseComObject(chartObj)
                               End If
                               If printRange IsNot Nothing Then
                                   Marshal.ReleaseComObject(printRange)
                               End If

                               If sheets IsNot Nothing Then
                                   Marshal.ReleaseComObject(sheets)
                               End If

                               If workbook IsNot Nothing Then
                                   workbook.Close(False)
                                   Marshal.ReleaseComObject(workbook)
                               End If

                               If excelApp IsNot Nothing Then
                                   excelApp.Quit()
                                   Marshal.ReleaseComObject(excelApp)
                               End If

                               ' Buộc Garbage Collection dọn dẹp
                               GC.Collect()
                               GC.WaitForPendingFinalizers()
                           End Try
                       End Sub)
    End Function

    ''' <summary>
    ''' Xuất nhận xét
    ''' </summary>
    ''' <param name="stuSheet"></param>
    ''' <param name="rowIndex"></param>
    ''' <param name="nx"></param>
    Private Sub writeNhanXet(stuSheet As Excel.Worksheet, ByVal rowIndex As Integer, nx As String())
        If nx.Count <= 1 Then
            If nx.Count = 1 Then
                stuSheet.Range($"A{rowIndex}").Value = nx(0)
            Else
                stuSheet.Range($"A{rowIndex}").Value = Nothing
            End If
        Else
            stuSheet.Range($"A{rowIndex}").Value = nx(0)
            Dim rowToCopy As Excel.Range = CType(stuSheet.Rows(rowIndex), Excel.Range)
            rowToCopy.Copy()
            For j As Integer = 1 To nx.Count - 1
                Dim insertRow As Excel.Range = CType(stuSheet.Rows(j + rowIndex), Excel.Range)
                insertRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown)
                stuSheet.Range($"A{rowIndex + j}").Value = nx(j)
            Next
        End If
    End Sub
End Class
