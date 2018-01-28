Imports System.IO
Imports System.Reflection
Imports System.Text
Imports Microsoft.ClearScript.V8

Public Class MarkdownParser

    Public Class poyo

        Public Property _num As Integer
        Public Property _hoge As String
        Public Property _hogehoge As String

        Sub New(ByVal num As Integer, ByVal hoge As String, ByVal hogehoge As String)
            _num = num
            _hoge = hoge
            _hogehoge = hogehoge
        End Sub

    End Class

    Private _jsEngine As V8ScriptEngine = Nothing
    Private _disposed As Boolean = Nothing

    Public Sub New()
        _jsEngine = New V8ScriptEngine()
        _disposed = False

        ' Import marked.js
        Dim sr As New StreamReader( _
            Assembly.GetExecutingAssembly().GetManifestResourceStream("MD2E.marked.min.js") _
            , Encoding.GetEncoding("utf-8")
        )
        _jsEngine.Execute(sr.ReadToEnd)
        sr.Close()
    End Sub

    Public Function Transform(ByVal markdown As String) As String
        If _jsEngine Is Nothing Then
            Throw New ArgumentNullException("jsEngine")
        End If

        ' Import Converting Target
        _jsEngine.Script.markdown = markdown

        ' Return converted item
        Return _jsEngine.Evaluate("marked(markdown);")
    End Function

    Public Sub Dispose()
        If _disposed Then
            Return
        End If
        _disposed = True

        If _jsEngine Is Nothing Then
            Return
        End If
        _jsEngine.Dispose()
        _jsEngine = Nothing
    End Sub

End Class
