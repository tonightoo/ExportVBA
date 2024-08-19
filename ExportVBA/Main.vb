Imports Microsoft.Office.Interop.Excel
Imports ExportVBA.Config
Imports Org.BouncyCastle.Asn1.Tsp

Module Main

    ''' <summary>
    ''' Config
    ''' </summary>
    ''' <returns></returns>
    Private Property Config As ApplicationConfig

    Private Const CONFIG_PATH As String = ".\Config\ApplicationConfig.xml"

    Sub Main()

        Try

            '設定を読み込む
            Dim configService As IConfigService = New XmlConfigService(CONFIG_PATH)
            Config = configService.Load(Of ApplicationConfig)()

            Using runner = New ExcelRunner(Config.TargetExcel)
                runner.Run()
            End Using

        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try


    End Sub

End Module
