Attribute VB_Name = "Wide2Narrow"
Option Explicit

' *****************************
' * Wide2Narrow
' * Converts double-width numbers, letters, and puntuaction into half-width characters
' * Does not convert Greek symbols at this time
' *
' *
' * Copyright (c) 2001 Ryan Ginstrom
' *
' * You are free to do anything you want to this macro except sell it.
' * No warranties, etc. etc.
' *
' * 9/12/2001 -- Added characters to ReplaceArray
' * 3/07/2003 -- Refactored into sub-functions
' * 3/14/2004 -- Added choice of current story, current document, or all open documents
' *         - Wide2NarrowAll
' *         - Wide2NarrowDocument
' *         - Wide2NarrowStory
' *****************************

Public Sub Wide2NarrowAll() ' gets text in text boxes too

    Dim page As Object
    For Each page In Word.Documents
    
    page.Activate
    
        Wide2NarrowDocument
        
    Next page
    
End Sub

Public Sub Wide2NarrowDocument()
    
    ' For the current story
    Wide2NarrowStory
            
    ' "Ungroup" all the shapes in the document so that we can look at the text boxes
    While UngroupBoxes = True
    Wend

    Dim shpTemp As Object
    For Each shpTemp In ActiveDocument.Shapes
            shpTemp.Select
            Wide2NarrowStory
    Next shpTemp

End Sub

Public Sub Wide2NarrowStory()

    Dim code As Long ' character codes
    Dim OriginalStart As Long
    Dim OriginalEnd As Long
        
            
        OriginalStart = Selection.Start
        OriginalEnd = Selection.End
        
        Selection.Start = 0
        Selection.End = 0
        
        Wide2NarrowNumbers
        Wide2NarrowLetters
        Wide2NarrowSymbols
    
        ' return selection to original position
        Selection.Start = OriginalStart
        Selection.End = OriginalEnd

End Sub


Private Function UngroupBoxes() As Boolean
    
    UngroupBoxes = False

    Dim shpTemp As Object
    For Each shpTemp In ActiveDocument.Shapes
        If shpTemp.Type = msoGroup Then
            shpTemp.Ungroup
            UngroupBoxes = True
        End If
    Next shpTemp

End Function


Sub Wide2NarrowNumbers()

    With Selection.Find
        ' Initialize object
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchByte = False
        .MatchFuzzy = False
        
        ' convert numbers
        Dim code As Integer
        For code = Asc("0") To Asc("9")
            .Text = Chr(code)
            .Replacement.Text = Chr(code)
            .Execute Replace:=wdReplaceAll
        Next code
    End With
    
End Sub

Sub Wide2NarrowLetters()

    With Selection.Find
        ' Initialize object
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchByte = False
        .MatchFuzzy = False
        
        ' convert letters
        Dim code As Integer
        For code = Asc("a") To Asc("z")
            .Text = Chr(code)
            .Replacement.Text = Chr(code)
            .Execute Replace:=wdReplaceAll
        Next code
    End With
    
End Sub

Sub Wide2NarrowSymbols()

    Dim ReplaceArray As Variant
    Dim ReplaceString As Variant
    
    ' This is an array of special characters to replace
    ' You can add to or subtract from this list as needed
    ReplaceArray = _
    Array( _
        "ÅC,", "ÅD.", "ÅF:", "ÅG;", "ÅH?", "ÅI!", _
        "Å]-", "Å^/", "Å_\", "Åè\", "Åb|", "''", _
        "Åi(", "Åj)", "Åm[", "Ån]", "Åm[", "Ån]", "Åy[", "Åz]", "Åo{", "Åp}", "ÅÉ<", "ÅÑ>", _
        "Å{+", "Å|-", "Å~x", "ÅÅ=", _
        "Åê$", "Åñ*", "Åó@", "Åï&", "Åì%", "Åî#", "Å@ ")
    
    With Selection.Find
        ' Initialize object
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .MatchByte = False
        .MatchFuzzy = False
    
        ' convert misc symbols in replacearray
        For Each ReplaceString In ReplaceArray
            .Text = Left(ReplaceString, 1)
            .Replacement.Text = Right(ReplaceString, 1)
            .Execute Replace:=wdReplaceAll
        Next ReplaceString
    End With
    
End Sub


