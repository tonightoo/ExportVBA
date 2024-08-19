Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.Xml.Serialization

Namespace Config

    ''' <summary>
    ''' XMLの設定ファイルを取扱うクラス
    ''' </summary>
    Public Class XmlConfigService : Implements IConfigService

        ''' <summary>
        ''' 設定ファイルのパス
        ''' </summary>
        ''' <returns></returns>
        Private Property FilePath As String

        Public Sub New(ByVal filePath As String)
            Me.FilePath = filePath
        End Sub

        ''' <summary>
        ''' オブジェクトから設定ファイルを作成する
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="obj"></param>
        Public Sub Save(Of T)(ByVal obj As T) Implements IConfigService.Save
            Using fs = New FileStream(FilePath, FileMode.Append)
                Dim xs As New XmlSerializer(GetType(T))
                xs.Serialize(fs, obj)
            End Using
        End Sub

        ''' <summary>
        ''' 設定ファイルを読み込み設定用オブジェクトを返す
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <returns></returns>
        Public Function Load(Of T)() As T Implements IConfigService.Load
            Dim obj As Object

            Using fs = New FileStream(FilePath, FileMode.Open)
                Dim xs = New XmlSerializer(GetType(T))
                obj = DirectCast(xs.Deserialize(fs), T)
            End Using

            Return obj
        End Function
    End Class
End Namespace
