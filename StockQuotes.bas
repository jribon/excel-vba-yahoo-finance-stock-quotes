Attribute VB_Name = "StockQuotes"
Option Explicit

Private Const MAIN_SHEET As String = "Dashboard"
Private Const YAHOO_URL As String = "http://finance.yahoo.com/d/quotes.csv?s=[s]&f=[f]"

Private colCount As Integer
Private rowCount As Integer


' Primary Macro.
' This is the application's entry point (when clicking on the "Refresh" button).
Public Sub Refresh()
    Dim url As String
    
    Call ComputeTableSize   ' count number of rows and columns
    Call ClearTable         ' clear imported data
    Call ComputeTableSize   ' recompute size after clearing the data
    
    ' get symbol and tag list and return valid URL
    url = BuildURL
    
    ' remove empty row and column headers
    Call DeleteEmptyHeaders
    
    ' download online data from Yahoo Finance
    Call LoadExternalData(url)
    
    ' resize columns to fit their content
    Call ResizeColumns
End Sub


' Count number of rows and columns
Private Sub ComputeTableSize()
    With ActiveWorkbook.Sheets(MAIN_SHEET)
        rowCount = .UsedRange.Rows.Count
        colCount = .UsedRange.Columns.Count
    End With
End Sub


' Clear imported data while preserving column and row headers.
Private Sub ClearTable()
    If rowCount > 2 And colCount > 2 Then
        ' Only clear table if there are some values
        With ActiveWorkbook.Sheets(MAIN_SHEET)
            .Range(.Cells(3, 3), .Cells(rowCount, colCount)).ClearContents
        End With
    End If
End Sub


' Resize columns to fit content of cells.
Private Sub ResizeColumns()
    With ActiveWorkbook.Sheets(MAIN_SHEET)
        .Range("C3").CurrentRegion.EntireColumn.AutoFit
    End With
End Sub


' Build a valid Yahoo Finance's URL based on the given symbols and tags.
Private Function BuildURL()
    Dim url As String
    Dim symbols As String
    Dim tags As String
    
    ' Concatenates symbols and tags
    symbols = GetSymbols
    tags = GetTags
    
    ' Inserts symbol and tag lists as parameters
    url = YAHOO_URL
    url = Replace(url, "[s]", symbols)    ' e.g. GOOGL+AAPL+FB+AMZN
    url = Replace(url, "[f]", tags)       ' e.g. s0n0l1d1t1
    
    BuildURL = url
End Function


' Concatenate symbols (aka. tickers).
' @return: string containing symbols separated by "+" characters.
Private Function GetSymbols() As String
    Dim i As Integer
    Dim s As String
    Dim symbols As String
    
    ' Loops on symbols and add them to the list.
    symbols = ""
    For i = 3 To rowCount
        s = ActiveWorkbook.Sheets(MAIN_SHEET).Cells(i, 1).Value
        
        ' concatenates only if there is a symbol.
        If s <> "" Then
            symbols = IIf(symbols = "", s, symbols & "+" & s)
        End If
    Next i
    
    GetSymbols = symbols
End Function


' Concatenate tags.
' @return: string containing tags (with no sepration character).
Private Function GetTags() As String
    Dim j As Integer
    Dim t As String
    Dim tags As String
    
    ' Loops on tags and add them to the list.
    tags = ""
    For j = 3 To colCount
        t = ActiveWorkbook.Sheets(MAIN_SHEET).Cells(1, j).Value
        tags = tags & t
    Next j
        
    GetTags = tags
End Function


' Delete the extra empty columns and rows.
Private Sub DeleteEmptyHeaders()
    Dim i As Integer
    
    With ActiveWorkbook.Sheets(MAIN_SHEET)
        ' Deletes empty rows (loop goes backward)
        For i = rowCount To 3 Step -1
            If .Cells(i, 1).Value = "" Then
                .Cells(i, 1).EntireRow.Delete
            End If
        Next i
        
        'Delete empty columns (loop goes backward)
        For i = colCount To 3 Step -1
            If .Cells(1, i).Value = "" Then
                .Cells(1, i).EntireColumn.Delete
            End If
        Next i
    End With
End Sub

' Download stock quotes from Yahoo Finance.
' @param url: valid URL with parameters, pointing to the Yahoo Finance API.
Private Sub LoadExternalData(url As String)
    Dim q As QueryTable
    Dim s As Worksheet
    Dim r As Range
        
    ' Avoids alert messages when replacing data
    Application.DisplayAlerts = False
    
    ' Set destination sheet and destination range for the returned data
    Set s = ActiveWorkbook.Sheets(MAIN_SHEET)
    Set r = s.Range("C3")
    
    ' Indicates that the result returned by the URL is a text file.
    url = "TEXT;" & url
    
    ' Fetch online data using QueryTable Object:
    '  - Create a new QueryTable using QueryTables.Add(URL, DestinationRange)
    '  - Use QueryTable.Refresh to send the request to Yahoo Finance API
    '  - Finally, delete the QueryTable using QueryTable.Delete
    Set q = s.QueryTables.Add(url, r)
    With q
        .RefreshStyle = xlOverwriteCells                    ' Replace current cells
        .BackgroundQuery = False                            ' Synchronous Query
        .TextFileParseType = xlDelimited                    ' Parsing Type (column  separated by  a character)
        .TextFileTextQualifier = xlTextQualifierDoubleQuote ' Column Name Delimiter ""
        .TextFileCommaDelimiter = True                      ' Column Separator
        .Refresh
    End With
    
    ' Destroys the QueryTable object (since it is used only once)
    q.Delete
    
    ' Re-enables alert messages
    Application.DisplayAlerts = True
End Sub

