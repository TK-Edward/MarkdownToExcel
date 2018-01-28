Module MarkdownToExcel

Private Property mdEncoding As String = "utf-8"

Public Enum ModeState
    None = 0
    DirectMode
    BulkMode
End Enum

    Sub Main()

        Try
            ' getArguments
            Dim args As String() = System.Environment.GetCommandLineArgs()

            ' アプリケーションパスを取得
            Dim appPath As String = args(0)
            Dim curDirPath As String = System.IO.Path.GetDirectoryName(appPath)
            System.Environment.CurrentDirectory = curDirPath

            ConvertInBulkMode()

            Console.WriteLine("処理が正常に完了しました。")

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            Console.WriteLine("Enterを押下で終了します...")
            Console.ReadLine()
        End Try

    End Sub


Private Function GetMarkdownText(ByVal path As String) As String

    Dim sr As System.IO.StreamReader = Nothing
    Dim strMd As String = Nothing

    Try
        sr = New System.IO.StreamReader(path, System.Text.Encoding.GetEncoding(mdEncoding))
        strMd = sr.ReadToEnd()
    Catch ex As Exception
        Throw
    Finally
        If Not sr Is Nothing Then
            sr.Close()
        End If
    End Try

    Return strMd

End Function

Private Sub ConvertInBulkMode()
    ' Markdownパーサー
    Dim mdp As MarkdownParser = Nothing
    ' EXCEL
    Dim cts As New CreateTestSpecification()

    Try
        ' 設定を取得
        Dim xConf As XElement = XElement.Load("Config.xml")
        ' 変換対象設定の取得
        Dim xConvDataList As IEnumerable(Of XElement) = From x As XElement In xConf.Element("import").Elements("book").Elements("sheet")
        'ファイル名に使用できない文字を取得
        Dim invalidChars As Char() = System.IO.Path.GetInvalidPathChars
        ' Markdownパーサー
        mdp = New MarkdownParser

        For Each xConvData As XElement In xConvDataList
            ' レイアウト取得
            Dim xLayout As XElement = (From x As XElement In xConf.Elements("layout") Where x.@name = xConvData.@layout).First()
            ' Markdown取得
            Dim path As String = xConvData.@src
            If Not IsSafePath(path, False) Then
                Console.WriteLine(path)
                Throw New Exception("不正なパスです。")
            End If
            Dim mdPath As String = GetFilePath(xConvData.@src)
            Dim strMd As String = GetMarkdownText(mdPath)
            Dim strXMd As String = "<body>" & vbCrLf & mdp.Transform(strMd) & "</body>"
            strXMd = strXMd.Replace("<hr>", "<hr/>")
            Console.WriteLine(strXMd)
            Dim xMarkdown As XElement = XElement.Parse(strXMd)
            cts.Convert(xConvData.Parent.@name, GetFileName(mdPath, False), xMarkdown, xLayout)
        Next

        Dim outputPath As String = (From x As XElement In xConf.Elements("output")).First().@src
        cts.Save(GetFilePath(outputPath))
    Catch ex As Exception
    Finally
        If Not mdp Is Nothing Then
            mdp.Dispose()
        End If
    End Try

End Sub

Private Sub ConvertInDirectMode(ByVal src As String)

End Sub


End Module