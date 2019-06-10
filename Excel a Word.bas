Attribute VB_Name = "Módulo1"
Sub CopiarWord()

' Copiar el grafico de excel.......
ActiveSheet.Range("A1:C10").Copy

Dim appWord As New Word.Application
Dim docWord As New Word.Document
Dim rng As Range

' Agregamos un nuevo documento....
With appWord
    .Visible = True
    Set docWord = .Documents.Add
    .Activate
End With
 With appWord.Selection
    ' Le damos formato al documento......
    .HomeKey unit:=wdLine, Extend:=wdExtend
    .ParagraphFormat.Alignment = wdAlignParagraphCenter
    .Font.Size = 18
    With .Font
    .Name = "Verdana"
    .Size = 18
    .Bold = True
    .Italic = False
    .Smallcaps = True
    End With



' Pegamos el grafico......
.TypeParagraph
.TypeParagraph
.Paste
End With

With docWord

End With
Set appWord = Nothing

End Sub
