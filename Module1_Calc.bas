
Sub Main

End Sub


Sub SortSheetsAndMakeIndex
    Dim oDoc As Object, oSheets As Object, i As Integer
    Dim aNames() As String, n As Integer
    oDoc = ThisComponent
    oSheets = oDoc.Sheets
    n = oSheets.getCount()

    ' Get sheet names
    ReDim aNames(n - 1)
    For i = 0 To n - 1
        aNames(i) = oSheets.getByIndex(i).Name
    Next i

    ' Sort names alphabetically
    Dim j As Integer, sTmp As String
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If UCase(aNames(i)) > UCase(aNames(j)) Then
                sTmp = aNames(i)
                aNames(i) = aNames(j)
                aNames(j) = sTmp
            End If
        Next j
    Next i

    ' Reorder sheets
    For i = 0 To n - 1
        oSheets.moveByName(aNames(i), i)
    Next i

    ' Remove existing Index sheet
    If oSheets.hasByName("Index") Then oSheets.removeByName("Index")

    ' Insert new Index sheet
    oSheets.insertNewByName("Index", 0)
    Dim oIndexSheet As Object
    oIndexSheet = oSheets.getByName("Index")

    ' Add hyperlinks to sheets
    Dim oCell As Object, oText As Object, oCursor As Object, oField As Object
    For i = 1 To n
        oCell = oIndexSheet.getCellByPosition(0, i - 1)
        oText = oCell.Text
        oCursor = oText.createTextCursor()
        oText.setString("")
        oField = oDoc.createInstance("com.sun.star.text.TextField.URL")
        oField.URL = "#" & "'" & aNames(i - 1) & "'.A1"
        oField.Representation = aNames(i - 1)
        oField.TargetFrame = "_self"
        oText.insertTextContent(oCursor, oField, False)
    Next i
End Sub


'------------------------------------------------------------------------------
' Function: MonthsWeeksDays
'
' Description:
'   Calculates the difference between two dates and expresses it as:
'     - years, months, weeks, and days (calendar-based)
'
' Parameters:
'   startDate       - Numeric or date value (e.g. cell A1 with a date)
'   endDate         - Numeric or date value (e.g. cell B1 with a later date)
'   [formatOpt]     - Optional string to control format:
'                     "long"  => e.g. "1 year, 2 months, 3 weeks, 4 days" (default)
'                     "short" => e.g. "1y 2m 3w 4d"
'   [showAllParts]  - Optional boolean:
'                     FALSE (default) => skip parts where value is 0
'                     TRUE            => show all parts, even 0s (e.g. "0 years, 0 months, 2 weeks")
'
' Usage in Calc:
'   =MonthsWeeksDays(A1, B1)
'   =MonthsWeeksDays(A1, B1, "short")
'   =MonthsWeeksDays(A1, B1, "long", TRUE)
'
' Notes:
'   - Returns "" if dates are missing
'   - Returns "Error: ..." on invalid inputs
'------------------------------------------------------------------------------

Function MonthsWeeksDays(startDate As Variant, endDate As Variant, Optional formatOpt As Variant, Optional showAllParts As Variant) As String
    Dim sd As Date, ed As Date, tmpDate As Date
    Dim totalDays As Long
    Dim years As Integer, months As Integer, weeks As Integer, days As Integer
    Dim result As String
    Dim fmt As String, showAll As Boolean

    On Error GoTo ErrorHandler

    ' Handle empty input
    If IsEmpty(startDate) Or IsEmpty(endDate) Then
        MonthsWeeksDays = ""
        Exit Function
    End If

    ' Accept numeric (serial) dates only
    If IsNumeric(startDate) Then
        sd = CDate(startDate)
    Else
        MonthsWeeksDays = "Error: Bad start"
        Exit Function
    End If

    If IsNumeric(endDate) Then
        ed = CDate(endDate)
    Else
        MonthsWeeksDays = "Error: Bad end"
        Exit Function
    End If

    If ed < sd Then
        MonthsWeeksDays = "Error: End < Start"
        Exit Function
    End If

    ' Handle options
    If IsMissing(formatOpt) Or IsEmpty(formatOpt) Then
        fmt = "long"
    Else
        fmt = LCase(Trim(CStr(formatOpt)))
    End If

    If IsMissing(showAllParts) Or IsEmpty(showAllParts) Then
        showAll = False
    Else
        showAll = CBool(showAllParts)
    End If

    ' Extract calendar parts
    Dim y1 As Integer, m1 As Integer, d1 As Integer
    Dim y2 As Integer, m2 As Integer, d2 As Integer

    y1 = Year(sd): m1 = Month(sd): d1 = Day(sd)
    y2 = Year(ed): m2 = Month(ed): d2 = Day(ed)

    ' Calculate year/month/day difference
    years = y2 - y1
    months = m2 - m1
    days = d2 - d1

    If days < 0 Then
        months = months - 1
        tmpDate = DateSerial(y2, m2, 1)
        days = ed - DateSerial(Year(tmpDate), Month(tmpDate), d1)
    End If

    If months < 0 Then
        years = years - 1
        months = months + 12
    End If

    tmpDate = DateAdd("yyyy", years, sd)
    tmpDate = DateAdd("m", months, tmpDate)
    totalDays = ed - tmpDate
    weeks = Int(totalDays / 7)
    days = totalDays Mod 7


    ' Remaining days
    totalDays = ed - tmpDate
    weeks = Int(totalDays / 7)
    days = totalDays Mod 7

    result = ""

    If fmt = "short" Then
        If showAll Or years > 0 Then result = result & years & "y "
        If showAll Or months > 0 Then result = result & months & "m "
        If showAll Or weeks > 0 Then result = result & weeks & "w "
        If showAll Or days > 0 Then result = result & days & "d"
    Else ' long
        If showAll Or years > 0 Then result = result & years & " year" & IIf(years <> 1, "s", "") & ", "
        If showAll Or months > 0 Then result = result & months & " month" & IIf(months <> 1, "s", "") & ", "
        If showAll Or weeks > 0 Then result = result & weeks & " week" & IIf(weeks <> 1, "s", "") & ", "
        If showAll Or days > 0 Then result = result & days & " day" & IIf(days <> 1, "s", "") & ", "
        If Right(result, 2) = ", " Then result = Left(result, Len(result) - 2)
    End If

    MonthsWeeksDays = Trim(result)
    Exit Function

ErrorHandler:
    MonthsWeeksDays = "Error"
End Function



'------------------------------------------------------------------------------
' Function: WordCount
'
' Description:
'   Counts the number of words in a given string.
'
' Parameters:
'   text (Variant) - Any text input from a cell (numbers will be converted)
'
' Returns:
'   Integer - word count (0 if blank or invalid)
'
' Usage in Calc:
'   =WordCount(A1)
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Function: WordCount
'
' Description:
'   Counts words in a string, handling spaces, tabs, and newlines
'
' Parameters:
'   text (Variant) - Input string or cell
'
' Returns:
'   Integer - Number of words (tokens separated by any whitespace)
'------------------------------------------------------------------------------

Function WordCount(text As Variant) As Integer
    Dim clean As String
    Dim i As Integer, count As Integer
    Dim words() As String

    On Error GoTo ErrorHandler

    If IsMissing(text) Or IsEmpty(text) Then
        WordCount = 0
        Exit Function
    End If

    ' Normalize input and replace all whitespace with a single space
    clean = CStr(text)
    clean = Replace(clean, Chr(13), " ")
    clean = Replace(clean, Chr(10), " ")
    clean = Replace(clean, Chr(9), " ")
    clean = Trim(clean)

    If clean = "" Then
        WordCount = 0
        Exit Function
    End If

    ' Split on space and count non-empty tokens
    words = Split(clean, " ")
    count = 0
    For i = LBound(words) To UBound(words)
        If Trim(words(i)) <> "" Then count = count + 1
    Next i

    WordCount = count
    Exit Function

ErrorHandler:
    WordCount = 0
End Function

Sub Macro1

End Sub


