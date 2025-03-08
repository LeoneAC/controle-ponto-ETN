Attribute VB_Name = "AtalhoColarValores"
Sub ColarValores()
Attribute ColarValores.VB_Description = "Macro para colar conte�do em uma C�lula como Valores diretamente, evitando que apare�a a mensagem de refer�ncia cruzada."
Attribute ColarValores.VB_ProcData.VB_Invoke_Func = "V\n14"
'
' ColarValores Macro
' Macro para colar conte�do em uma C�lula como Valores diretamente, evitando que apare�a a mensagem de refer�ncia cruzada.
' Problemas: Usar esse atalho mata o hist�rico de mudan�as, ent�o o Ctrl+Z reseta depois de colar valores com o atalho
'
' Atalho do teclado: Ctrl+Shift+V
'

    ' Check if the selection is a range
    If TypeName(Selection) = "Range" Then
        On Error Resume Next 'Desativa a manipula��o de erros do VBA para que ele apenas ignore e siga em frente em caso de erros
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        On Error GoTo 0 'Reseta a manipula��o padr�o de erros do VBA
    End If

'''' C�digo abandonado (mas funcional no excel 2016)
'    Dim selectedRange As Range
'    Dim clipboard As MSForms.DataObject
'
'    ' Criar o objeto para acessar o clipboard
'    Set clipboard = New MSForms.DataObject
'    clipboard.GetFromClipboard
'    ' Check if the selection is a range
'    If TypeName(Selection) = "Range" Then
'        Set selectedRange = Selection
'        ' Obter o texto do clipboard
'        On Error Resume Next
'        selectedRange.Value = TimeValue(clipboard.GetText)
'        On Error GoTo 0
'    End If
    
'''' C�digo abandonado (mas funcional no excel 2016)
'    Dim selectedRange As Range
'    Dim objCP As Object
'
'    ' Criar o objeto para acessar o clipboard
'    Set objCP = CreateObject("HtmlFile")
'
'    ' Check if the selection is a range
'    If TypeName(Selection) = "Range" Then
'        Set selectedRange = Selection
'        ' Obter o texto do clipboard
'        On Error Resume Next
'        selectedRange.Value = TimeValue(objCP.ParentWindow.ClipboardData.GetData("text"))
'        On Error GoTo 0
'    End If
    
End Sub
