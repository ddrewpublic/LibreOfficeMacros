
Sub Main

End Sub


Sub ShrinkAllTextSizesBy2
    Dim oDoc As Object : oDoc = ThisComponent
    Dim oText As Object : oText = oDoc.Text
    Dim oCursor As Object : oCursor = oText.createTextCursor()
    
    ' Select entire document
    oCursor.gotoStart(False)
    oCursor.gotoEnd(True)

    ' Reduce font size safely
    Dim originalSize As Single
    originalSize = oCursor.CharHeight

    If originalSize > 2 Then
        oCursor.CharHeight = originalSize - 2
    End If
End Sub


Sub ShrinkAllStylesBy2pt
    Dim oDoc As Object : oDoc = ThisComponent
    Dim oStyleFamilies As Object : oStyleFamilies = oDoc.StyleFamilies

    ' Handle paragraph styles
    If oStyleFamilies.hasByName("ParagraphStyles") Then
        Dim oParaStyles As Object : oParaStyles = oStyleFamilies.getByName("ParagraphStyles")
        Dim sName As String
        For Each sName In oParaStyles.getElementNames()
            Dim oParaStyle As Object : oParaStyle = oParaStyles.getByName(sName)
            If oParaStyle.CharHeight > 2 Then
                oParaStyle.CharHeight = oParaStyle.CharHeight - 1
            End If
        Next
    End If

    ' Handle character styles
    If oStyleFamilies.hasByName("CharacterStyles") Then
        Dim oCharStyles As Object : oCharStyles = oStyleFamilies.getByName("CharacterStyles")
        For Each sName In oCharStyles.getElementNames()
            Dim oCharStyle As Object : oCharStyle = oCharStyles.getByName(sName)
            If oCharStyle.CharHeight > 2 Then
                oCharStyle.CharHeight = oCharStyle.CharHeight - 1
            End If
        Next
    End If
End Sub


