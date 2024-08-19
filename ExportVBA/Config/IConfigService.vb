Namespace Config
    ''' <summary>
    ''' 設定を取り扱うクラス
    ''' </summary>
    Public Interface IConfigService

        ''' <summary>
        ''' 設定ファイルを読み込む
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <returns></returns>
        Function Load(Of T)() As T

        ''' <summary>
        ''' 設定ファイルを書き出す
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        Sub Save(Of T)(ByVal obj As T)

    End Interface
End Namespace
