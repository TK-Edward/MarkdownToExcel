Module Common

Public Function IsSafePath(ByVal path As String, ByVal isFileName As Boolean) As Boolean
    If (String.IsNullOrEmpty(path)) Then
        ' null か、空文字列は不正とする
        Return False
    End If

    ' 使えない文字があるかチェック
    Dim invalidChars As Char()
    If (isFileName) Then
        invalidChars = System.IO.Path.GetInvalidFileNameChars()
    Else
        invalidChars = System.IO.Path.GetInvalidPathChars()
    End If

    If (path.IndexOfAny(invalidChars) >= 0) Then
        ' 使えない文字がある
        Return False
    End If

    ' 使えないファイル名を含んでいないかチェック
    If System.Text.RegularExpressions.Regex.IsMatch(path _
                                        , "(^|\\|/)(CON|PRN|AUX|NUL|CLOCK\$|COM[0-9]|LPT[0-9])(\.|\\|/|$)" _
                                        , System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
        ' 使えないファイル名がある
        Return False
    End If
    Return True
End Function

Public Function GetFilePath(ByVal path As String) As String
    If Not System.IO.Path.IsPathRooted(path) Then
        ' ファイル名が絶対パスである場合は返す
        Return path
    Else
        ' 相対パスから絶対パスを取得し返す
        Return System.IO.Path.GetFullPath(path)
    End If
End Function

Public Enum FileSystemType
    None = 0
    File
    Directory
End Enum

Public Function GetFileSystemType(ByVal path As String)
    If System.IO.File.Exists(path) Then
        Return FileSystemType.File
    ElseIf System.IO.Directory.Exists(path) Then
        Return FileSystemType.Directory
    Else
        Return FileSystemType.None
    End If
End Function

Public Function GetFileName(ByVal path As String, ByVal ExtensionFlg As Boolean) As String
    If ExtensionFlg Then
        Return System.IO.Path.GetFileName(path)
    Else
        Return System.IO.Path.GetFileNameWithoutExtension(path)
    End If
End Function

Public Function IsMarkdown(ByVal path As String) As Boolean
    Dim extension As String = System.IO.Path.GetExtension(path)
    Select Case extension.ToUpper
        Case "MARKDOWM", "MD"
            Return True
    End Select
    Return False
End Function

End Module
