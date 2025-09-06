Imports System.IO

Public Class ConfigManager
    Private configFile As String = "config.ini"
    Private _config As ConfigInfo

    Public ReadOnly Property GetConfig As ConfigInfo
        Get
            Return _config
        End Get
    End Property

    Public Sub New()
        If Not File.Exists(configFile) Then
            _config = New ConfigInfo
            SaveConfig()
        Else
            _config = New ConfigInfo
        End If
    End Sub

    ''' <summary>
    ''' Xuất config ra file
    ''' </summary>
    Public Sub SaveConfig()
        Using writer As New StreamWriter(configFile, False)
            writer.WriteLine("TemplatePath=" & _config.TemplatePath)
            writer.WriteLine("InputFolder=" & _config.InputFolder)
            writer.WriteLine("OutputFolder=" & _config.OutputFolder)
            writer.WriteLine("LastWeekExport=" & _config.LastWeekExport)
        End Using
    End Sub

    ''' <summary>
    ''' Đọc file config
    ''' </summary>
    Public Sub LoadConfig()
        If Not File.Exists(configFile) Then
            Exit Sub
        End If

        Dim lines() As String = File.ReadAllLines(configFile)
        For Each line In lines
            If String.IsNullOrWhiteSpace(line) OrElse Not line.Contains("=") Then Continue For

            Dim parts() As String = line.Split(New Char() {"="c}, 2) ' chia thành 2 phần
            Dim key As String = parts(0).Trim()
            Dim value As String = parts(1).Trim()

            Select Case key
                Case "TemplatePath"
                    _config.TemplatePath = value
                Case "InputFolder"
                    _config.InputFolder = value
                Case "OutputFolder"
                    _config.OutputFolder = value
                Case "LastWeekExport"
                    _config.LastWeekExport = value
            End Select
        Next
    End Sub
End Class

Public Class ConfigInfo
    Public Property TemplatePath As String = "MauPhieuLienLacLop6A2.xlsx"
    Public Property InputFolder As String
    Public Property OutputFolder As String
    Public Property LastWeekExport As String
End Class