Attribute VB_Name = "NewMacros"
Sub Pastelinktextonly()
'
' Pastelinktextonly Macro
'
'
'Selection.PasteSpecial link:=True, dataType:=wdPasteText, Placement:=wdInLine, DisplayAsIcon:=False
Dim objRange As Range

Set objRange = Selection.Range

objRange.PasteSpecial link:=True, dataType:=wdPasteText, Placement:=wdInLine, DisplayAsIcon:=False
'objRange.Paragraphs(1).Range.InlineShapes(1).LinkFormat.AutoUpdate = False
'objRange.Paragraphs(0).Range.InlineShapes(0).LinkFormat.AutoUpdate = False



Set objRange = Nothing



End Sub
