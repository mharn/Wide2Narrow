Attribute VB_Name = "FindNextJ"
'======================================
' FindNextJ
'
' Copyright (C) Ryan Ginstrom
' Version 0.1
'
' MIT License (do whatever you want, just don't sue me)
'======================================

Option Explicit

Private Const SMARTQUOTE_1 As Integer = -32408
Private Const SMARTQUOTE_2 As Integer = -32409
Private Const SMARTSINGLEQUOTE_1 As Integer = -32410
Private Const SMARTSINGLEQUOTE_2 As Integer = -32411
Private Const HALF_KATA_START As Integer = 166
Private Const HALF_KATA_END As Integer = 221

Sub FindNextJ()
   
' Turn off screen updating
Application.ScreenUpdating = False
    
    With Selection
    
    Do While .End < .StoryLength
      
        .Collapse direction:=wdCollapseEnd
        .MoveRight Unit:=wdCharacter, Count:=1, Extend:=wdExtend
   
        If .LanguageDetected = True Then
            If .LanguageID = wdJapanese Then
                ' Turn on screen updating
                Application.ScreenUpdating = True
                Exit Sub
            End If
        ElseIf IsJ(.Text) Then
            ' Turn on screen updating
            Application.ScreenUpdating = True
            Exit Sub
        End If
            
    Loop
    
    ' We reached the end of the document without
    ' finding any Japanese characters
    .Collapse direction:=wdCollapseEnd
    
    End With
    
' Turn on screen updating
Application.ScreenUpdating = True
    
End Sub

Function IsJ(c As String) As Boolean

    If Len(c) = 0 Then
        IsJ = False
        Exit Function
    End If
       
    Dim charcode As Integer
    charcode = Asc(c)
    
    If charcode = SMARTQUOTE_1 Or _
            charcode = SMARTQUOTE_2 Or _
            charcode = SMARTSINGLEQUOTE_1 Or _
            charcode = SMARTSINGLEQUOTE_2 Then
            
        IsJ = False
        Exit Function
    End If
       
    If charcode < 0 Or _
            (charcode >= HALF_KATA_START And charcode <= HALF_KATA_END) Then
        IsJ = True
        Exit Function
    End If
    
    IsJ = False
    
End Function
