Attribute VB_Name = "NovoCalend�rio"
Option Explicit 'Adicione Option Explicit no in�cio do m�dulo para exigir declara��o expl�cita de vari�veis.

'Rotina para resetar a planilha para um novo preenchimento
Sub LimpaPlanilha()
    
    'Sugere salvar a planilha em um novo arquivo
    Select Case MsgBox("Iniciar um novo Controle de Horas apaga todos os dados dos meses e do calend�rio." & vbCrLf & vbCrLf & _
                        "� importante que voc� trabalhe em uma nova c�pia da planilha para n�o perder as informa��es do ano anterior." & vbCrLf & vbCrLf & _
                        "Deseja continuar em um novo arquivo?", _
                        vbYesNoCancel + vbQuestion, _
                        "Voc� est� iniciando o Controle de Horas do ano seguinte")
        Case vbYes
            'https://stackoverflow.com/q/64352055/9736020
            Dim suggestedFileName As String
            suggestedFileName = ThisWorkbook.Path & "\" & "Controle de Horas " & (ThisWorkbook.Sheets("DADOS").Range("ANO").Value + 1) & " - Com Macro"
            With Application.FileDialog(msoFileDialogSaveAs)
                Application.EnableEvents = False
                .InitialFileName = suggestedFileName
                .FilterIndex = 2 'Define o filtro de tipo de arquivo - 2: Pasta de trabalho habilitada para macros (.xlsm)
                .Title = "Salvar arquivo como"
                If .Show Then
                    ActiveWorkbook.SaveAs Filename:=.SelectedItems(1), FileFormat:=xlOpenXMLWorkbookMacroEnabled
                    'Continua com a macro ap�s salvamento
                Else
                    Exit Sub
                End If
                Application.EnableEvents = True
            End With
        Case vbNo
            Select Case MsgBox("Esta opera��o apagar� todos os dados do Calend�rio e das planilhas de meses.", _
                                vbOKCancel + vbExclamation, _
                                "Aten��o!")
                Case vbOK
                    'Nada a fazer, usu�rio confirmou que quer continuar sem salvar
                Case vbCancel
                    Exit Sub
            End Select
        Case vbCancel
            Exit Sub
    End Select
    
    ThisWorkbook.Activate
    ThisWorkbook.Unprotect
    
    Dim meses() As String: meses = Split("Jan,Fev,Mar,Abr,Mai,Jun,Jul,Ago,Set,Out,Nov,Dez", ",")
    Dim intervalos() As String: intervalos = Split("Pontos,Tratamento,Observa��es,Marca��o manual,Calend�rio", ",")
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim mes As Variant
    Dim ano As Integer: ano = Sheets("DADOS").Range("ANO").Value

    
    'Refazer os meses com base em uma c�pia de seguran�a
    Dim wsIndex As Integer
    Dim mes_anterior As Variant: mes_anterior = "BASE"
    Worksheets("BASE").Visible = xlSheetVisible
    For Each mes In meses
        If Not SheetExists(mes) Then
            wsIndex = Worksheets(mes_anterior).Index
            Worksheets("BASE").Copy After:=Worksheets(mes_anterior)
            Set ws = Worksheets(wsIndex + 1)
            ws.Name = mes
            ws.Tab.Color = False
        End If
        mes_anterior = mes
    Next mes
    Worksheets("BASE").Visible = xlSheetVeryHidden
    
    'Limpa todos os dados que devem ser limpos para um novo preenchimento da planilha
    Dim editableRange As AllowEditRange
    For Each mes In meses
        For Each editableRange In Worksheets(mes).Protection.AllowEditRanges
            If IsInArray(editableRange.Title, intervalos) Then
                editableRange.Range.ClearContents
            End If
        Next editableRange
    Next mes
    For Each editableRange In Worksheets("DADOS").Protection.AllowEditRanges
        If IsInArray(editableRange.Title, intervalos) Then
            Application.EnableEvents = False
            editableRange.Range.ClearContents
            Application.EnableEvents = True
        End If
    Next editableRange
    
    'Coloca o primeiro dia do ano no Calend�rio para configurar o ano da planilha e dar um exemplo de preenchimento ao usu�rio
    '01/01/25    FERIADO Confraterniza��o Universal - Feriado Nacional
    Set tbl = Sheets("DADOS").ListObjects("Calendario")
    Application.EnableEvents = False
    tbl.ListColumns("DIA").DataBodyRange.Cells(1, 1).Value = "01/01/" & (ano + 1)
    tbl.ListColumns("TIPO").DataBodyRange.Cells(1).Value = "FERIADO"
    tbl.ListColumns("DESCRI��O").DataBodyRange(1).Value = "Confraterniza��o Universal - Feriado Nacional"
    Application.EnableEvents = True
    
    'Oculta as linhas que s�o vazias
    Dim data As Range
    For Each mes In meses
        For Each data In Sheets(mes).Range("$C$2:$C$32")
            If data.Value = "" Then
                data.EntireRow.Hidden = True
            End If
        Next data
    Next mes
    
    'Oculta as colunas auxiliares (em constru��o)
    For Each mes In meses
        With Sheets(mes)
            .Range("B:B,H:K").EntireColumn.Hidden = True
            .Columns("Y").ColumnWidth = 0.1
            .Activate
            .Range("A1").Select
        End With
    Next mes
    
    'Proteger as planilhas e pasta de trabalho
    For Each mes In meses
        Sheets(mes).Protect AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    Next mes
    Sheets("DADOS").Protect AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    Sheets("EXEMPLO").Protect
    Sheets("BASE").Protect AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    ThisWorkbook.Protect
    
    'Volta para planilha DADOS
    Sheets("DADOS").Select
    
End Sub

'Verifica que se uma string est� contida num vetor
Public Function IsInArray(str As String, arr As Variant) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If str = arr(i) Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

'Verifica se uma planilha existe
Public Function SheetExists(sheetName As Variant) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next ws
    SheetExists = False

End Function
