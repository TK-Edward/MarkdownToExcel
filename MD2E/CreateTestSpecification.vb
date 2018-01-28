Imports ClosedXML.Excel

Public Class CreateTestSpecification

Private Property _books As Dictionary(Of String, XLWorkbook)
Private Property _sheets As List(Of String)

Private Property _extention As String = ".xlsx"

Public Structure CPoint
    Private _x As Integer
    Private _y As Integer
    Public ReadOnly Property x() As Integer
        Get
            Return Me._x
        End Get
    End Property
    Public ReadOnly Property y() As Integer
        Get
            Return Me._y
        End Get
    End Property
    Public Sub New(ByVal x As Integer, ByVal y As Integer)
        Me._x = x
        Me._y = y
    End Sub
End Structure

Public Structure CPointRange
    Private _beginPoint As CPoint
    Private _endPoint As CPoint
    Public ReadOnly Property BeginPoint() As CPoint
        Get
            Return Me._beginPoint
        End Get
    End Property
    Public ReadOnly Property EndPoint() As CPoint
        Get
            Return Me._endPoint
        End Get
    End Property
    Public ReadOnly Property Range() As Integer()
        Get
            Return {_beginPoint.x, _beginPoint.y, _endPoint.x, _endPoint.y}
        End Get
    End Property
    Public ReadOnly Property firstCellRow() As Integer
        Get
            Return _beginPoint.x
        End Get
    End Property
    Public ReadOnly Property firstCellCol() As Integer
        Get
            Return _beginPoint.y
        End Get
    End Property
    Public ReadOnly Property lastCellRow() As Integer
        Get
            Return _endPoint.x
        End Get
    End Property
    Public ReadOnly Property lastCellCol() As Integer
        Get
            Return _endPoint.y
        End Get
    End Property
    Public Sub New(ByVal beginPoint As CPoint, ByVal endPoint As CPoint)
        Me._beginPoint = beginPoint
        Me._endPoint = endPoint
    End Sub
    Public Sub New(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)
        Me._beginPoint = New CPoint(x1, y1)
        Me._endPoint = New CPoint(x2, y2)
    End Sub
End Structure

Public Sub New()
    _books = New Dictionary(Of String, XLWorkbook)
End Sub

Public Sub WriteFixdSection(ByVal sheet As IXLWorksheet, ByVal layout As XElement)
    Dim items As IEnumerable(Of XElement) = From x As XElement In layout.Element("fixdSection").Elements("item")
    Dim startCell As IXLCell = sheet.Cell(items.First().Parent.@startCell)
    If startCell Is Nothing OrElse startCell.Equals("") Then
        startCell = sheet.LastCellUsed
    End If
    For Each item As XElement In items

        Dim lblFlg As Boolean = False
        If Not item.@type Is Nothing AndAlso item.@type.Equals("lbl") Then
            lblFlg = True
        End If

        Dim range As IXLRange = sheet.Range(item.@cell)

        Dim firstCellCol As Integer = startCell.Address.ColumnNumber + range.FirstColumn.ColumnNumber - 1
        Dim firstCellRow As Integer = startCell.Address.RowNumber + range.FirstRow.RowNumber - 1
        Dim lastCellCol As Integer = startCell.Address.ColumnNumber + range.LastColumn.ColumnNumber - 1
        Dim lastCellRow As Integer = startCell.Address.RowNumber + range.LastRow.RowNumber - 1

        range = sheet.Range(firstCellRow, firstCellCol, lastCellRow, lastCellCol)

        With sheet.Range(range.RangeAddress)
            .Merge(False)
            .SetValue(item.Value)
            .Style.Border.OutsideBorder = XLBorderStyleValues.Thin

            If lblFlg Then
                .Style.Fill.BackgroundColor = XLColor.LightGray
            End If
        End With
    Next
End Sub

Public Sub WriteTableSection(ByVal sheet As IXLWorksheet, ByVal layout As XElement, ByVal data As XElement)

    Dim items As IEnumerable(Of XElement) = From x As XElement In layout.Element("tableSection").Element("header").Elements("item")
    Dim startCell As IXLCell = sheet.Cell(items.First().Parent.@startCell)
    If startCell Is Nothing OrElse startCell.Equals("") Then
        startCell = sheet.Cell("A" & (sheet.RowsUsed().Last.RowNumber + 1))
    End If
    Dim range As IXLRange = sheet.Range(startCell, startCell)

    Dim startRow As Integer = startCell.WorksheetRow.RowNumber
    Dim startCol As Integer = startCell.WorksheetColumn.ColumnNumber

    Dim currentRow As Integer = startRow
    Dim currentCol As Integer = startCol

    Dim tblData As New DataTable
    tblData.Columns.Add("H1")
    tblData.Columns.Add("H2")
    tblData.Columns.Add("H3")
    tblData.Columns.Add("H4")
    tblData.Columns.Add("H5")
    tblData.Columns.Add("H6")
    tblData.Columns.Add("H7")

    For Each dataItem As XElement In data.Elements
        For Each col As DataColumn In tblData.Columns
            If dataItem.Name.ToString.ToUpper.Equals(col.ColumnName) AndAlso tblData.Columns.IndexOf(col.ColumnName) = 1 Then
                Dim nRow As DataRow = tblData.NewRow()
                    nRow(col.ColumnName) = dataItem.Value
                    tblData.Rows.Add(nRow)
            ElseIf dataItem.Name.ToString.ToUpper.Equals(col.ColumnName) Then
                If tblData.Rows(tblData.Rows.Count - 1)(col.ColumnName).ToString.Equals("") Then
                        tblData.Rows(tblData.Rows.Count - 1)(col.ColumnName) = dataItem.Value
                Else
                    Dim nRow As DataRow = tblData.NewRow()
                    nRow(col.ColumnName) = dataItem.Value
                    tblData.Rows.Add(nRow)
                End If
            End If
        Next
        'Select Case dataItem.Name.ToString.ToUpper
        '    Case "H1"
        '        Dim nRow As DataRow = tblData.NewRow()
        '        nRow("H1") = dataItem.Value
        '        tblData.Rows.Add(nRow)
        '    Case "H2"
        '        If tblData.Rows(tblData.Rows.Count - 1)("H2").ToString.Equals("") Then
        '            tblData.Rows(tblData.Rows.Count - 1)("H2") = dataItem.Value
        '        Else
        '            Dim nRow As DataRow = tblData.NewRow()
        '            nRow("H2") = dataItem.Value
        '            tblData.Rows.Add(nRow)
        '        End If
        '    Case "H3"
        '        If tblData.Rows(tblData.Rows.Count - 1)("H3").ToString.Equals("") Then
        '            tblData.Rows(tblData.Rows.Count - 1)("H3") = dataItem.Value
        '        Else
        '            Dim nRow As DataRow = tblData.NewRow()
        '            nRow("H3") = dataItem.Value
        '            tblData.Rows.Add(nRow)
        '        End If
        'End Select
    Next

    Dim layoutTbl As New DataTable
    layoutTbl.Columns.Add("firstCellCol")
    layoutTbl.Columns.Add("firstCellRow")
    layoutTbl.Columns.Add("lastCellCol")
    layoutTbl.Columns.Add("lastCellRow")

    Dim firstCellCol As Integer = range.FirstColumn.ColumnNumber
    Dim firstCellRow As Integer = range.FirstRow.RowNumber
    Dim lastCellCol As Integer = range.LastColumn.ColumnNumber
    Dim lastCellRow As Integer = range.LastRow.RowNumber
    For Each item As XElement In items
        lastCellCol = firstCellCol + Integer.Parse(item.@rangeWidth) - 1
        lastCellRow = firstCellRow + Integer.Parse(item.@rangeHeight) - 1

        Dim row As DataRow = layoutTbl.NewRow()
        row("firstCellCol") = firstCellCol
        row("firstCellRow") = firstCellRow
        row("lastCellCol") = lastCellCol
        row("lastCellRow") = lastCellRow
        layoutTbl.Rows.Add(row)

        firstCellCol = lastCellCol + 1
    Next

    Dim writeRange As IXLRange = Nothing
    For Each dtRow As DataRow In tblData.Rows

        For Each layoutRow As DataRow In layoutTbl.Rows
            Dim rowIndex As Integer = tblData.Rows.IndexOf(dtRow)
            Dim colIndex As Integer = layoutTbl.Rows.IndexOf(layoutRow)

            writeRange = sheet.Range(layoutRow("firstCellRow") + rowIndex, _
                                        layoutRow("firstCellCol"), _
                                        layoutRow("lastCellRow") + rowIndex, _
                                        layoutRow("lastCellCol"))

            If tblData.Columns.Count > colIndex Then
                With sheet.Range(writeRange.RangeAddress)
                    .Merge(False)
                    .SetValue(dtRow(colIndex).ToString)
                    .Style.Border.OutsideBorder = XLBorderStyleValues.Thin
                End With
            Else
                With sheet.Range(writeRange.RangeAddress)
                    .Merge(False)
                    .Style.Border.OutsideBorder = XLBorderStyleValues.Thin
                End With
            End If
        Next
    Next

End Sub

Public Sub WriteTableSection(ByVal sheet As IXLWorksheet, ByVal layout As XElement)

    Dim items As IEnumerable(Of XElement) = From x As XElement In layout.Element("tableSection").Element("header").Elements("item")
    Dim startCell As IXLCell = sheet.Cell(items.First().Parent.@startCell)
    If startCell Is Nothing OrElse startCell.Equals("") Then
        startCell = sheet.Cell("A" & (sheet.RowsUsed().Last.RowNumber + 1))
    End If
    Dim range As IXLRange = sheet.Range(startCell, startCell)

    Dim startRow As Integer = startCell.WorksheetRow.RowNumber
    Dim startCol As Integer = startCell.WorksheetColumn.ColumnNumber

    Dim firstCellCol As Integer = range.FirstColumn.ColumnNumber
    Dim firstCellRow As Integer = range.FirstRow.RowNumber
    Dim lastCellCol As Integer = range.LastColumn.ColumnNumber
    Dim lastCellRow As Integer = range.LastRow.RowNumber
    For Each item As XElement In items
        lastCellCol = firstCellCol + Integer.Parse(item.@rangeWidth) - 1
        lastCellRow = firstCellRow + Integer.Parse(item.@rangeHeight) - 1

        range = sheet.Range(firstCellRow, firstCellCol, lastCellRow, lastCellCol)

        With sheet.Range(range.RangeAddress)
            .Merge(False)
            .SetValue(item.Value)
            .Style.Fill.BackgroundColor = XLColor.LightBlue
            .Style.Border.OutsideBorder = XLBorderStyleValues.Thin
        End With

        firstCellCol = lastCellCol + 1
    Next
End Sub

Public Sub Convert(ByVal bookName As String, ByVal sheetName As String, ByVal xMdData As XElement, ByVal xLayout As XElement)

    Dim worksheet As IXLWorksheet = Nothing
    If Not _books.ContainsKey(bookName) Then
        _books.Add(bookName, New XLWorkbook)
        worksheet = _books(bookName).AddWorksheet(sheetName)
    ElseIf Not _books(bookName).TryGetWorksheet(sheetName, worksheet) Then
        worksheet = _books(bookName).AddWorksheet(sheetName)
    End If

    WriteFixdSection(worksheet, xLayout)
    WriteTableSection(worksheet, xLayout)
    WriteTableSection(worksheet, xLayout, xMdData)

End Sub

Public Sub Save(ByVal dirPath As String)
    For Each book As KeyValuePair(Of String, XLWorkbook) In _books
        book.Value.SaveAs(System.IO.Path.Combine(dirPath, book.Key & _extention))
    Next
End Sub

End Class
