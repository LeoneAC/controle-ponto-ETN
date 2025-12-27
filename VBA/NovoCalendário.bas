Attribute VB_Name = "NovoCalendário"
Option Explicit 'Adicione Option Explicit no início do módulo para exigir declaração explícita de variáveis

'Rotina para resetar a planilha para um novo preenchimento
Sub LimpaPlanilha()
    
    'Sugere salvar a planilha em um novo arquivo
    Select Case MsgBox("Iniciar um novo Controle de Horas apaga todos os dados dos meses e do calendário." & vbCrLf & vbCrLf & _
                        "É importante que você trabalhe em uma nova cópia da planilha para não perder as informações do ano anterior." & vbCrLf & vbCrLf & _
                        "Deseja continuar em um novo arquivo?", _
                        vbYesNoCancel + vbQuestion, _
                        "Você está iniciando o Controle de Horas do ano seguinte")
        Case vbYes
            'https://stackoverflow.com/q/64352055/9736020
            Dim suggestedFileName As String
            suggestedFileName = ThisWorkbook.Path & "\" & "Controle de Horas " & (DADOS.Range("ANO").Value + 1)
            With Application.FileDialog(msoFileDialogSaveAs)
                Application.EnableEvents = False
                .InitialFileName = suggestedFileName
                .FilterIndex = 2 'Define o filtro de tipo de arquivo - 2: Pasta de trabalho habilitada para macros (.xlsm)
                .Title = "Salvar arquivo como"
                If .Show Then
                    ActiveWorkbook.SaveAs Filename:=.SelectedItems(1), FileFormat:=xlOpenXMLWorkbookMacroEnabled
                    'Continua com a macro após salvamento
                Else
                    Exit Sub
                End If
                Application.EnableEvents = True
            End With
        Case vbNo
            Select Case MsgBox("Esta operação apagará todos os dados do Calendário e das planilhas de meses.", _
                                vbOKCancel + vbExclamation, _
                                "Atenção!")
                Case vbOK
                    'Nada a fazer, usuário confirmou que quer continuar sem salvar
                Case vbCancel
                    Exit Sub
            End Select
        Case vbCancel
            Exit Sub
    End Select
    
    ThisWorkbook.Activate
    ThisWorkbook.Unprotect
    
    Dim meses() As String: meses = Split("Jan,Fev,Mar,Abr,Mai,Jun,Jul,Ago,Set,Out,Nov,Dez", ",")
    Dim intervalos() As String: intervalos = Split("Pontos,Tratamento,Observações,Marcação manual,Calendário", ",")
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim mes As Variant
    'Se o usuário apagou todo o Calendário antes de clicar no botão, usa o Ano atual (de hoje)
    Dim ano As Integer: If IsNumeric(DADOS.Range("ANO").Value) Then ano = DADOS.Range("ANO").Value Else ano = Year(Date)

    'Desativa o cálculo automático das células, os eventos e a atualização gráfica para melhorar o desempenho durante a criação das planilhas
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    'Refazer os meses com base em uma cópia de segurança
    Dim wsIndex As Integer
    Dim vbComp As Object
    Dim mes_anterior As Variant: mes_anterior = BASE.Name
    BASE.Visible = xlSheetVisible
    For Each mes In meses
        If SheetExists(mes) Then
            Application.DisplayAlerts = False 'evita mensagem de confirmação
            Worksheets(mes).Delete
            Application.DisplayAlerts = True
        End If
        wsIndex = Worksheets(mes_anterior).Index
        BASE.Copy After:=Worksheets(mes_anterior)
        Set ws = Worksheets(wsIndex + 1)
        ws.Name = mes
        ws.Tab.Color = False
        mes_anterior = mes
    Next mes
    BASE.Visible = xlSheetVeryHidden
    
    'Limpa todos os dados que devem ser limpos para um novo preenchimento da planilha
    Dim editableRange As AllowEditRange
    For Each editableRange In DADOS.Protection.AllowEditRanges
        If IsInArray(editableRange.Title, intervalos) Then
            Application.EnableEvents = False
            editableRange.Range.ClearContents
            Application.EnableEvents = True
        End If
    Next editableRange
    
    'Reativa o cálculo automático da planilha para estabelecer valores iniciais e evitar problemas de compilação
    Application.Calculation = xlCalculationAutomatic
    
    'Coloca o primeiro dia do ano no Calendário para configurar o ano da planilha e dar um exemplo de preenchimento ao usuário
    '01/01/25    FERIADO Confraternização Universal - Feriado Nacional
    Set tbl = DADOS.ListObjects("Calendario")
    Application.EnableEvents = False
    tbl.ListColumns("DIA").DataBodyRange.Cells(1, 1).Value = "01/01/" & (ano + 1)
    tbl.ListColumns("TIPO").DataBodyRange.Cells(1).Value = "FERIADO"
    tbl.ListColumns("DESCRIÇÃO").DataBodyRange(1).Value = "Confraternização Universal - Feriado Nacional"
    Application.EnableEvents = True
    
    'Oculta as linhas que são vazias
    Dim data As Range
    For Each mes In meses
        For Each data In Sheets(mes).Range("$C$2:$C$32")
            If data.Value = "" Then
                data.EntireRow.Hidden = True
            End If
        Next data
    Next mes
    
    'Oculta as colunas auxiliares
    For Each mes In meses
        With Sheets(mes)
            .Range("B:B,H:K").EntireColumn.Hidden = True
            .Columns("Y").ColumnWidth = 0.1
            .Activate
            .Range("A1").Select
        End With
    Next mes
    
    'Proteger as planilhas e pasta de trabalho
    DesprotegerTodasPlanilhas 'Percisei fazer isso uma vez para forçar o comportamento correto dos Protects
    For Each mes In meses
        Sheets(mes).Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    Next mes
    DADOS.Protect AllowFormattingColumns:=False, AllowFormattingRows:=False, AllowFormattingCells:=False, DrawingObjects:=False
    EXEMPLO.Protect
    BASE.Protect AllowFormattingColumns:=True, AllowFormattingRows:=True
    ThisWorkbook.Protect
    
    'Volta para planilha DADOS
    DADOS.Select
    
    'Reativa os eventos e a atualização gráfica
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
End Sub

'Verifica que se uma string está contida num vetor
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

Sub DesprotegerTodasPlanilhas()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect
    Next ws
End Sub
