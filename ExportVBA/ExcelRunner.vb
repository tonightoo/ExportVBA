Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports ExportVBA.Domain
Imports System.Reflection

Public Class ExcelRunner
    Implements IDisposable

    Private disposedValue As Boolean

    Private XlApplication As Microsoft.Office.Interop.Excel.Application = Nothing

    Private XlWorkbooks As Workbooks = Nothing

    Private XlBook As Workbook = Nothing

    Private ExcelInstance As Excel

    Public Sub New(ByVal excelFilePath As String)
    End Sub

    Public Sub New(ByVal excel As Excel)
        Dim appPath As String = Environment.CurrentDirectory
        XlApplication = New Microsoft.Office.Interop.Excel.Application()
        XlWorkbooks = XlApplication.Workbooks
        XlBook = XlWorkbooks.Open($"{appPath}\{excel.FileName}")
        Me.ExcelInstance = excel
    End Sub

    Public Sub Run()
        With Me.ExcelInstance
            XlApplication.Run($"{ .FileName}!{ .TargetModule.MethodName}")
        End With
    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: マネージド状態を破棄します (マネージド オブジェクト)
            End If

            ' TODO: アンマネージド リソース (アンマネージド オブジェクト) を解放し、ファイナライザーをオーバーライドします
            ' TODO: 大きなフィールドを null に設定します

            If XlBook IsNot Nothing Then
                XlBook.Close()
                Marshal.ReleaseComObject(XlBook)
            End If

            If XlWorkbooks IsNot Nothing Then
                Marshal.ReleaseComObject(XlWorkbooks)
            End If

            If XlApplication IsNot Nothing Then
                XlApplication.Quit()
                Marshal.ReleaseComObject(XlApplication)
            End If

            disposedValue = True
        End If
    End Sub

    ' ' TODO: 'Dispose(disposing As Boolean)' にアンマネージド リソースを解放するコードが含まれる場合にのみ、ファイナライザーをオーバーライドします
    ' Protected Overrides Sub Finalize()
    '     ' このコードを変更しないでください。クリーンアップ コードを 'Dispose(disposing As Boolean)' メソッドに記述します
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' このコードを変更しないでください。クリーンアップ コードを 'Dispose(disposing As Boolean)' メソッドに記述します
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class
