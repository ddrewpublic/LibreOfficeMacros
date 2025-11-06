
Sub Main

End Sub

Sub ForceFontViaCharacterProperties
    Dim oDoc : oDoc = ThisComponent
    Dim oPages : oPages = oDoc.getDrawPages()
    Dim oPage, oShape, oTextCursor

    For iPage = 0 To oPages.getCount() - 1
        oPage = oPages.getByIndex(iPage)
        For iShape = 0 To oPage.getCount() - 1
            oShape = oPage.getByIndex(iShape)

            On Error Resume Next
            oTextCursor = oShape.Text.createTextCursor()
            oTextCursor.gotoStart(False)
            oTextCursor.gotoEnd(True)
            oTextCursor.CharFontName = "Liberation Sans"
            On Error GoTo 0
        Next iShape
    Next iPage
End Sub


