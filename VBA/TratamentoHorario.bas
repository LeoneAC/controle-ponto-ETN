Attribute VB_Name = "TratamentoHorario"
Option Explicit

'============================================================='
'''''''' DECLARAÇÃO DAS VARIÁVEIS E CONSTANTES GLOBAIS ''''''''
'============================================================='

''''Declaração dos horários notáveis como constantes globais deste módulo
''Comercial
Private Const ENTRADA_COM As Date = #7:30:00 AM#                'TimeValue("7:30")
Private Const SAIDA_ALMOCO_COM As Date = #12:00:00 PM#          'TimeValue("12:00")
Private Const LIM_SUP_SAIDA_ALMOCO_COM As Date = #12:30:00 PM#  'TimeValue("12:30") '=RETORNO_ALMOCO - "1:00"
Private Const LIM_INF_RETORNO_ALMOCO_COM As Date = #1:00:00 PM# 'TimeValue("13:00") '=SAIDA_ALMOCO + "1:00"
Private Const RETORNO_ALMOCO_COM As Date = #1:30:00 PM#         'TimeValue("13:30")
Private Const SAIDA_COM As Date = #5:00:00 PM#                  'TimeValue("17:00")
''Construção
Private Const ENTRADA_C As Date = #7:30:00 AM#                  'TimeValue("7:30")
Private Const SAIDA_ALMOCO_C As Date = #11:45:00 AM#            'TimeValue("11:45")
Private Const LIM_SUP_SAIDA_ALMOCO_C As Date = #12:45:00 PM#    'TimeValue("12:45")
Private Const LIM_INF_RETORNO_ALMOCO_C As Date = #12:45:00 PM#  'TimeValue("12:45")
Private Const RETORNO_ALMOCO_C As Date = #1:45:00 PM#           'TimeValue("13:45")
Private Const SAIDA_C As Date = #5:30:00 PM#                    'TimeValue("17:30")
''Construção (sexta)
Private Const ENTRADA_CS As Date = #7:00:00 AM#                 'TimeValue("7:00")
Private Const SAIDA_ALMOCO_CS As Date = #11:45:00 AM#           'TimeValue("11:45")
Private Const LIM_SUP_SAIDA_ALMOCO_CS As Date = #12:15:00 PM#   'TimeValue("12:15")
Private Const LIM_INF_RETORNO_ALMOCO_CS As Date = #12:45:00 PM# 'TimeValue("12:45")
Private Const RETORNO_ALMOCO_CS As Date = #1:15:00 PM#          'TimeValue("13:15")
Private Const SAIDA_CS As Date = #4:30:00 PM#                   'TimeValue("16:30")
''Almoço padrão de 1 hora (que serve para todos os tipos de pontos e todos os dias)
Private Const INI_ALMOCO_PADRAO_1H As Date = #12:00:00 PM#      'TimeValue("12:00")
Private Const FIM_ALMOCO_PADRAO_1H As Date = #1:00:00 PM#       'TimeValue("13:00")

''''Declaração das variáveis estáticas de horários
Private ENTRADA_                As Date
Private SAIDA_ALMOCO_           As Date
Private LIM_SUP_SAIDA_ALMOCO_   As Date
Private LIM_INF_RETORNO_ALMOCO_ As Date
Private RETORNO_ALMOCO_         As Date
Private SAIDA_                  As Date

''''Declaração das constantes que serão usadas como marcadores para definir com quais horários padrões as variáveis serão inicializadas
Private Const COM As Integer = 1 'Comercial
Private Const C As Integer = 2   'Construção
Private Const CS As Integer = 3  'Construção Sexta

''''Declaração das pseudo-constantes que servirão para evitar que os Ranges e Names precisem ser acessados e avaliados o tempo todo
Private FDS_ As String
Private FERIAS_ As String
Private FERIADO_ As String
Private DISPENSADO_ As String
Private MEIO_COMPENSADO_ As String
Private EXPEDIENTE_CORRIDO_ As String
'

'============================================================='
''''''''' DECLARAÇÃO DAS FUNÇÕES E ROTINAS AUXILIARES '''''''''
'============================================================='

'Procedimento para configurar os horários padrões de acordo com o tipo de ponto e dia
'A ideia dessa rotina é reduzir o número de execuções de inicialização e otimizar o tempo da planilha
Sub DefinirConstanteTempo(dia_semana As String, tipo_ponto As String)

    ''''Declaração da variável estática de flag que indicará com o atual tipo de ponto e dia (Comercial, Construção ou Construção-Sexta)
    Static TPD As Integer 'Por padrão, Inteiros estáticos são inicializados com 0, então na primeira chamada essa função sempre entrará em algum 'if', colocando algum horário padrão

    ''''Testa o caso atual
    Select Case tipo_ponto '---> Equivale a: ThisWorkbook.Names("TIPO_DE_PONTO").RefersToRange.Value
        Case "Comercial"
            If TPD <> COM Then 'Só define as variáveis se o caso mudou da última chamada
                TPD = COM
                ENTRADA_ = ENTRADA_COM
                SAIDA_ALMOCO_ = SAIDA_ALMOCO_COM
                LIM_SUP_SAIDA_ALMOCO_ = LIM_SUP_SAIDA_ALMOCO_COM
                LIM_INF_RETORNO_ALMOCO_ = LIM_INF_RETORNO_ALMOCO_COM
                RETORNO_ALMOCO_ = RETORNO_ALMOCO_COM
                SAIDA_ = SAIDA_COM
            End If
        Case Else '"Construção"
            Select Case dia_semana
                Case "Sexta"
                    If TPD <> CS Then
                        TPD = CS
                        ENTRADA_ = ENTRADA_CS
                        SAIDA_ALMOCO_ = SAIDA_ALMOCO_CS
                        LIM_SUP_SAIDA_ALMOCO_ = LIM_SUP_SAIDA_ALMOCO_CS
                        LIM_INF_RETORNO_ALMOCO_ = LIM_INF_RETORNO_ALMOCO_CS
                        RETORNO_ALMOCO_ = RETORNO_ALMOCO_CS
                        SAIDA_ = SAIDA_CS
                    End If
                Case Else
                    If TPD <> C Then
                        TPD = C
                        ENTRADA_ = ENTRADA_C
                        SAIDA_ALMOCO_ = SAIDA_ALMOCO_C
                        LIM_SUP_SAIDA_ALMOCO_ = LIM_SUP_SAIDA_ALMOCO_C
                        LIM_INF_RETORNO_ALMOCO_ = LIM_INF_RETORNO_ALMOCO_C
                        RETORNO_ALMOCO_ = RETORNO_ALMOCO_C
                        SAIDA_ = SAIDA_C
                        End If
            End Select
    End Select
End Sub

'Procedimento para ser chamado na abertura da planilha e inicializar os Tipos de Dias usados nos tratamentos
'A ideia dessa rotina é reduzir o número de execuções de inicialização e otimizar o tempo da planilha
Sub InicializaTiposDia()
    FDS_ = CStr(Evaluate(ThisWorkbook.Names("FIM_DE_SEMANA").RefersTo)) 'Não é possível obter o FIM_DE_SEMANA pelo Range, já que ele não referencia nenhum range
    FERIAS_ = DADOS.Range("FERIAS").Value
    FERIADO_ = DADOS.Range("FERIADO").Value
    DISPENSADO_ = DADOS.Range("DISPENSADO").Value
    MEIO_COMPENSADO_ = DADOS.Range("MEIO_COMPENSADO").Value
    EXPEDIENTE_CORRIDO_ = DADOS.Range("EXPEDIENTE_CORRIDO").Value
End Sub

'Função auxiliar para saber se o dia é do tipo férias (e afins) ou não
Function EhFerias(tipo_dia As String) As Boolean
    If FERIAS_ = "" Then InicializaTiposDia 'Testa a inicialização das FÉRIAS (que é chamada primeiro) porque alguns eventos (como erros não tratados) apagam os valores das variáveis inicializadas
    EhFerias = (tipo_dia = FERIAS_)
    'Será aqui que serão acrescentados os tipos de dia FÉRIAS* e o COMPENSADO_CRE1
End Function

'Função auxiliar para saber se o dia é de um tipo especial ou não
Function EhDiaEspecial(tipo_dia As String) As Boolean
    EhDiaEspecial = (tipo_dia = FDS_) _
                 Or (tipo_dia = FERIADO_) _
                 Or (tipo_dia = DISPENSADO_) _
                 Or (tipo_dia = MEIO_COMPENSADO_)
End Function

'Função auxiliar para arredondar os valores de tempo e evitar problemas de lixo de memória em operações
'Essa função arredonda os segundos para o minuto mais próximo e é necessária porque o Excel considera tempos como doubles (cheio de casas decimais)
'A lógica dela é: divide o tempo por 1 min para achar (em inteiros) a quantidade de minutos, arredonda essa quantidade e reconverte para o double
'que representa aquele tempo. Assim o valor fica pouco fora do valor incorreto, resetando os erros acumulados nas operações
Function Round2minute(tempo As Date) As Date
    Round2minute = Round(tempo / #12:01:00 AM#) * (#12:01:00 AM#) '#12:01:00 AM# == "0:01"
End Function

'============================================================='
''''''' DECLARAÇÃO DAS FUNÇÕES DE TRATAMENTO PRINCIPAIS '''''''
'============================================================='

''''Função para tratar o horário de entrada
Function EntradaTratada(dia_semana As String, tipo_dia As String, entrada As Date, saida_almoco_tratada As Date, tipo_ponto As String, calendar As Range) As Date
Application.Volatile (False)
'Debug.Print "EntradaTratada"
'
'As variáveis de entrada 'calendar' e 'tipo_ponto' têm função apenas de fazer essa UDF (User Defined Function) ser recalculada quando esses valores atualizam

''''Excel base desta função (ao utilizar, lembre-se de que elas podem conter erro de digitação ou estarem desatualizadas):
'=SE($B2=FERIAS;
'    0;
'    SE(OU($B2=FIM_DE_SEMANA;$B2=FERIADO;$B2=DISPENSADO;$B2=MEIO_COMPENSADO);
'        D2;
'        SE(TIPO_DE_PONTO="Comercial";
'            SE(D2>"12:00"+0;
'                I2;
'                SE(E(D2>="7:00"+0;D2<="7:35"+0);VALOR.TEMPO("7:30");D2));
'            SE(E(TIPO_DE_PONTO="Construção";$A2<>"Sexta");
'                SE(D2>"11:45"+0;
'                    I2;
'                    SE(E(D2>="7:00"+0;D2<="7:35"+0);VALOR.TEMPO("7:30");D2));
'                SE(D2>"11:45"+0;
'                    I2;
'                    SE(E(D2>="6:30"+0;D2<="7:05"+0);VALOR.TEMPO("7:00");D2))))))
    
    ''''Filtros iniciais
    'Se for férias (ou afins), impede que qualquer valor passe para frente
    If EhFerias(tipo_dia) Then
        EntradaTratada = 0
        Exit Function
    End If
    'Se for algum dia especial, ele não trata o horário, deixando passar a informação para frente
    If EhDiaEspecial(tipo_dia) Then
        EntradaTratada = entrada
        Exit Function
    End If
    'Na entrada, o dia de Expediente Corrido é tratado como um dia normal
    
    ''''Configura os horários padrões de acordo com o tipo de ponto e dia
    Call DefinirConstanteTempo(dia_semana, tipo_ponto)
    
    ''''Tratamento principal
    'Se o usuário fez apenas a parte da tarde + hora extra no almoço (entra depois da saída para almoço), então a entrada recebe a saída para almoço para zerar o turno da manhã
    If (entrada > SAIDA_ALMOCO_) Then
        EntradaTratada = saida_almoco_tratada
    'Se o usuário bateu ponto até meia hora antes ou até 5min depois do horário de entrada, arredonda para o início do expediente
    ElseIf (Round2minute(entrada) >= Round2minute(ENTRADA_ - TimeValue("0:30"))) And (Round2minute(entrada) <= Round2minute(ENTRADA_ + TimeValue("0:05"))) Then ' Foi necessário truncar os valores por erro de arredondamento
        EntradaTratada = ENTRADA_
    'O comportamento padrão é não tratar o ponto
    Else
        EntradaTratada = entrada
    End If
        
End Function

''''Função para tratar o horário de saída
Function SaidaTratada(dia_semana As String, tipo_dia As String, horarios As Range, retorno_almoco_tratado As Date, tipo_ponto As String, calendar As Range) As Date
Application.Volatile (False)
'Debug.Print "SaidaTratada"
'
'As variáveis de entrada 'calendar' e 'tipo_ponto' têm função apenas de fazer essa UDF (User Defined Function) ser recalculada quando esses valores atualizam

''''Excel base desta função (ao utilizar, lembre-se de que elas podem conter erro de digitação ou estarem desatualizadas):
'=SE($B2=FERIAS;
'    0;
'    SE(OU($B2=FIM_DE_SEMANA;$B2=FERIADO;$B2=DISPENSADO;$B2=MEIO_COMPENSADO);
'        G2;
'        SE(E($B2=EXPEDIENTE_CORRIDO;$E2=0;$F2=0);
'            G2;
'            SE(TIPO_DE_PONTO="Comercial";
'                SE(G2<"13:30"+0;
'                    J2;
'                    SE(E(G2>="16:55"+0;G2<="17:00"+0);VALOR.TEMPO("17:00");G2));
'                SE(E(TIPO_DE_PONTO="Construção";$A2<>"Sexta");
'                    SE(G2<"13:45"+0;
'                        J2;
'                        SE(E(G2>="17:25"+0;G2<="17:30"+0);VALOR.TEMPO("17:30");G2));
'                    SE(G2<"13:15"+0;
'                        J2;
'                        SE(E(G2>="16:25"+0;G2<="16:30"+0);VALOR.TEMPO("16:30");G2)))))))

    ''''Declaração das variáveis
    Dim saida_almoco As Date:     saida_almoco = horarios.Cells(1, 2).Value
    Dim retorno_almoco As Date:   retorno_almoco = horarios.Cells(1, 3).Value
    Dim saida As Date:            saida = horarios.Cells(1, 4).Value

    ''''Filtros iniciais
    'Se for férias (ou afins), impede que qualquer valor passe para frente
    If EhFerias(tipo_dia) Then
        SaidaTratada = 0
        Exit Function
    End If
    'Se for algum dia especial, ele não trata o horário, deixando passar a informação para frente
    If EhDiaEspecial(tipo_dia) Then
        SaidaTratada = saida
        Exit Function
    End If
    'Se for um dia de Expediente Corrido sem marcação de almoço, passa a informação sem tratar (se o ponto do almoço foi batido, o dia corrido é considerado um dia normal)
    If (tipo_dia = EXPEDIENTE_CORRIDO_) And (saida_almoco = 0) And (retorno_almoco = 0) Then
        SaidaTratada = saida
        Exit Function
    End If
    
    ''''Configura os horários padrões de acordo com o tipo de ponto e dia
    Call DefinirConstanteTempo(dia_semana, tipo_ponto)
    
    ''''Tratamento principal
    'Se o usuário fez apenas a parte da manhã + hora extra no almoço (sai antes do retorno do almoço), então a saída recebe o retorno do almoço para zerar o turno da tarde
    If (saida < RETORNO_ALMOCO_) Then
        SaidaTratada = retorno_almoco_tratado
    'Se o usuário bateu ponto na nos últimos 5min do horário, arredonda para o fim do expediente
    ElseIf (Round2minute(saida) >= Round2minute(SAIDA_ - TimeValue("00:05"))) And (saida <= SAIDA_) Then ' Foi necessário truncar os valores por erro de arredondamento
        SaidaTratada = SAIDA_
    'O comportamento padrão é não tratar o ponto
    Else
        SaidaTratada = saida
    End If

End Function

''''Função para tratar o horário de saída para almoço
Function SaidaAlmocoTratada(dia_semana As String, tipo_dia As String, horarios As Range, tipo_ponto As String, calendar As Range) As Date
Application.Volatile (False)
'Debug.Print "SaidaAlmocoTratada"
'
'As variáveis de entrada 'calendar' e 'tipo_ponto' têm função apenas de fazer essa UDF (User Defined Function) ser recalculada quando esses valores atualizam
'
''''Excel base desta função (ao utilizar, lembre-se de que elas podem conter erro de digitação ou estarem desatualizadas):
'=SE($B2=FERIAS;
'    0;
'    SE(OU($B2=FIM_DE_SEMANA;$B2=FERIADO;$B2=DISPENSADO;$B2=MEIO_COMPENSADO);
'        $E2;
'        SE(E(ÉCÉL.VAZIA($E2);ÉCÉL.VAZIA($F2));
'            SE(E(ÉCÉL.VAZIA($D2);ÉCÉL.VAZIA($G2));
'                $E2;
'                SE($B2=EXPEDIENTE_CORRIDO;
'                    0;
'                    SE(TIPO_DE_PONTO="Comercial";
'                        SE($G2<="12:30"+0;
'                            $G2;
'                            SE($G2<="13:30"+0;
'                                VALOR.TEMPO("12:30");
'                                VALOR.TEMPO("12:00")));
'                        SE($A2<>"Sexta";
'                            SE($G2<="12:45"+0;
'                                $G2;
'                                SE($G2<="13:45"+0;
'                                    VALOR.TEMPO("12:45");
'                                    VALOR.TEMPO("11:45")));
'                            SE($G2<="12:15"+0;
'                                $G2;
'                                SE($G2<="13:15"+0;
'                                    VALOR.TEMPO("12:15");
'                                    VALOR.TEMPO("11:45")))))));
'            SE(TIPO_DE_PONTO="Comercial";
'                SE($E2<"12:00"+0;
'                    $E2;
'                    SE($F2>"13:30"+0;
'                        SE($E2>"12:30"+0;
'                            VALOR.TEMPO("12:30");
'                            $E2);
'                        SE(($F2-$E2)<=VALOR.TEMPO("1:00");
'                            VALOR.TEMPO("12:00");
'                            $E2)));
'                SE($A2<>"Sexta";
'                    SE($E2<"11:45"+0;
'                        $E2;
'                        SE($F2>"13:45"+0;
'                            SE($E2>"12:45"+0;
'                                VALOR.TEMPO("12:45");
'                                $E2);
'                            SE(($F2-$E2)<=VALOR.TEMPO("1:00");
'                                VALOR.TEMPO("12:00");
'                                $E2)));
'                    SE($E2<"11:45"+0;
'                        $E2;
'                        SE($F2>"13:15"+0;
'                            SE($E2>"12:15"+0;
'                                VALOR.TEMPO("12:15");
'                                $E2);
'                            SE(($F2-$E2)<=VALOR.TEMPO("1:00");
'                                VALOR.TEMPO("12:00");
'                                $E2))))))))

    ''''Declaração das variáveis
    Dim entrada As Date:          entrada = horarios.Cells(1, 1).Value
    Dim saida_almoco As Date:     saida_almoco = horarios.Cells(1, 2).Value
    Dim retorno_almoco As Date:   retorno_almoco = horarios.Cells(1, 3).Value
    Dim saida As Date:            saida = horarios.Cells(1, 4).Value
    
    ''''Filtros iniciais
    'Se for férias (ou afins), impede que qualquer valor passe para frente
    If EhFerias(tipo_dia) Then
        SaidaAlmocoTratada = 0
        Exit Function
    End If
    'Se for algum dia especial, ele não trata o horário, deixando passar a informação para frente
    If EhDiaEspecial(tipo_dia) Then
        SaidaAlmocoTratada = saida_almoco
        Exit Function
    End If

    ''''Configura os horários padrões de acordo com o tipo de ponto e dia
    Call DefinirConstanteTempo(dia_semana, tipo_ponto)
    
    ''''Tratamento principal (perceba que o tratamento padrão é passar o valor sem tratamento: SaidaAlmocoTratada = saida_almoco)
    'Se o almoço estiver vazio
    If ((saida_almoco = 0) And (retorno_almoco = 0)) Then
        'Se o resto todo estiver vazio, passa (o vazio) sem tratar
        If ((entrada = 0) And (saida = 0)) Then
            SaidaAlmocoTratada = saida_almoco
        'Se for um dia de Expediente Corrido, anula o almoço (se o ponto do almoço foi batido, o dia corrido é considerado um dia normal)
        ElseIf (tipo_dia = EXPEDIENTE_CORRIDO_) Then
            SaidaAlmocoTratada = 0
        'Se apenas entrada e saída estiverem preenchidas
        Else
            'Se a saída for antes do fim do turno da manhã (antes do limite superior de saída para almoço), passa a saída para a saída do almoço sem tratar
            If (saida <= LIM_SUP_SAIDA_ALMOCO_) Then
                SaidaAlmocoTratada = saida
            'Se a saída for maior que o limite superior de saída, porém menor que o limite (superiror) de retorno do almoço, trava a saída do almoço no limite superior de saída para almoço
            ElseIf (saida <= RETORNO_ALMOCO_) Then
                SaidaAlmocoTratada = LIM_SUP_SAIDA_ALMOCO_
            'Se não, isto é, se a saída for depois do almoço, trava a saída do almoço no limite inferior de saída para almoço para que ela não influencie em hora extra
            Else
                SaidaAlmocoTratada = SAIDA_ALMOCO_
            End If
        End If
    'Se o almoço não estiver vazio (supõe que entrada e saída também não estão)
    Else
        'Se a saída para almoço for antes do início do almoço, passa sem tratar
        If (saida_almoco < SAIDA_ALMOCO_) Then
            SaidaAlmocoTratada = saida_almoco
        'Se não, se o retorno do almoço for depois do fim do almoço
        ElseIf (retorno_almoco > RETORNO_ALMOCO_) Then
            'Se a saída para almoço for depois do limite para saída do almoço, trava a saída do almoço no limite de saída para almoço
            If (saida_almoco > LIM_SUP_SAIDA_ALMOCO_) Then
                SaidaAlmocoTratada = LIM_SUP_SAIDA_ALMOCO_
            'Se não, passa sem tratar
            Else
                SaidaAlmocoTratada = saida_almoco
            End If
        'Se não, se a diferença entre saída e retorno do almoço for menor que 1h, usa o valor padrão para a saída para almoço
        ElseIf (Round2minute(retorno_almoco - saida_almoco) <= Round2minute(TimeValue("1:00"))) Then ' Foi necessário truncar os valores por erro de arredondamento
            SaidaAlmocoTratada = INI_ALMOCO_PADRAO_1H
        'Se não for nenhum caso especial, passa sem tratar
        Else
            SaidaAlmocoTratada = saida_almoco
        End If
    End If

End Function

''''Função para tratar o horário de retorno do almoço
Function RetornoAlmocoTratado(dia_semana As String, tipo_dia As String, horarios As Range, tipo_ponto As String, calendar As Range) As Date
Application.Volatile (False)
'Debug.Print "RetornoAlmocoTratado"
'
'As variáveis de entrada 'calendar' e 'tipo_ponto' têm função apenas de fazer essa UDF (User Defined Function) ser recalculada quando esses valores atualizam

''''Excel base desta função (ao utilizar, lembre-se de que elas podem conter erro de digitação ou estarem desatualizadas):
'=SE($B2=FERIAS;
'    0;
'    SE(OU($B2=FIM_DE_SEMANA;$B2=FERIADO;$B2=DISPENSADO;$B2=MEIO_COMPENSADO);
'        $F2;
'        SE(E(ÉCÉL.VAZIA($E2);ÉCÉL.VAZIA($F2));
'            SE(E(ÉCÉL.VAZIA($D2);ÉCÉL.VAZIA($G2));
'                $F2;
'                SE($B2=EXPEDIENTE_CORRIDO;
'                    0;
'                    SE(TIPO_DE_PONTO="Comercial";
'                        SE($D2>="13:00"+0;
'                            $D2;
'                            SE($D2>="12:00"+0;
'                                VALOR.TEMPO("13:00");
'                                VALOR.TEMPO("13:30")));
'                        SE($A2<>"Sexta";
'                            SE($D2>="12:45"+0;
'                                $D2;
'                                SE($D2>="11:45"+0;
'                                    VALOR.TEMPO("12:45");
'                                    VALOR.TEMPO("13:45")));
'                            SE($D2>="12:45"+0;
'                                $D2;
'                                SE($D2>="11:45"+0;
'                                    VALOR.TEMPO("12:45");
'                                    VALOR.TEMPO("13:15")))))));
'            SE(TIPO_DE_PONTO="Comercial";
'                SE($E2<"12:00"+0;
'                    SE($F2<"13:00"+0;
'                        VALOR.TEMPO("13:00");
'                        $F2);
'                    SE($F2>"13:30"+0;
'                        $F2;
'                        SE(($F2-$E2)<=VALOR.TEMPO("1:00");
'                            VALOR.TEMPO("13:00");
'                            $F2)));
'                SE($A2<>"Sexta";
'                    SE($E2<"11:45"+0;
'                        SE($F2<"12:45"+0;
'                            VALOR.TEMPO("12:45");
'                            $F2);
'                        SE($F2>"13:45"+0;
'                            $F2;
'                            SE(($F2-$E2)<=VALOR.TEMPO("1:00");
'                                VALOR.TEMPO("13:00");
'                                $F2)));
'                    SE($E2<"11:45"+0;
'                        SE($F2<"12:45"+0;
'                            VALOR.TEMPO("12:45");
'                            $F2);
'                        SE($F2>"13:15"+0;
'                            $F2;
'                            SE(($F2-$E2)<=VALOR.TEMPO("1:00");
'                                VALOR.TEMPO("13:00");
'                                $F2))))))))

    ''''Declaração das variáveis
    Dim entrada As Date:          entrada = horarios.Cells(1, 1).Value
    Dim saida_almoco As Date:     saida_almoco = horarios.Cells(1, 2).Value
    Dim retorno_almoco As Date:   retorno_almoco = horarios.Cells(1, 3).Value
    Dim saida As Date:            saida = horarios.Cells(1, 4).Value
    
    ''''Filtros iniciais
    'Se for férias (ou afins), impede que qualquer valor passe para frente
    If EhFerias(tipo_dia) Then
        RetornoAlmocoTratado = 0
        Exit Function
    End If
    'Se for algum dia especial, ele não trata o horário, deixando passar a informação para frente
    If EhDiaEspecial(tipo_dia) Then
        RetornoAlmocoTratado = retorno_almoco
        Exit Function
    End If
    
    ''''Configura os horários padrões de acordo com o tipo de ponto e dia
    Call DefinirConstanteTempo(dia_semana, tipo_ponto)
    
    ''''Tratamento principal (perceba que o tratamento padrão é passar o valor sem tratamento: RetornoAlmocoTratado = retorno_almoco)
    'Se o almoço estiver vazio
    If ((saida_almoco = 0) And (retorno_almoco = 0)) Then
        'Se o resto todo estiver vazio, passa (o vazio) sem tratar
        If ((entrada = 0) And (saida = 0)) Then
            RetornoAlmocoTratado = retorno_almoco
        'Se for um dia de Expediente Corrido, anula o almoço (se o ponto do almoço foi batido, o dia corrido é considerado um dia normal)
        ElseIf (tipo_dia = EXPEDIENTE_CORRIDO_) Then
            RetornoAlmocoTratado = 0
        'Se apenas entrada e saída estiverem preenchidas
        Else
            'Se a entrada for depois do começo do turno da tarde (depois do limite inferior de retorno do almoço), passa a entrada para o retorno do almoço sem tratar
            If (entrada >= LIM_INF_RETORNO_ALMOCO_) Then
                RetornoAlmocoTratado = entrada
            'Se a entrada for menor que o limite inferior de retorno, porém maior que o limite (inferior) de saída para almoço, trava o retorno do almoço no limite inferior de retorno do almoço
            ElseIf (entrada >= SAIDA_ALMOCO_) Then
                RetornoAlmocoTratado = LIM_INF_RETORNO_ALMOCO_
            'Se não, isto é, se a entrada for antes do almoço, trava o retorno do almoço no limite superior de retorno do almoço para que ele não influencie em hora extra
            Else
                RetornoAlmocoTratado = RETORNO_ALMOCO_
            End If
        End If
    'Se o almoço não estiver vazio (supõe que entrada e saída também não estão)
    Else
        'Se a saída para almoço for antes do início do almoço
        If (saida_almoco < SAIDA_ALMOCO_) Then
            'Se o retorno do almoço for antes do limite para retorno do almoço, trava o retorno do almoço no limite de retorno do almoço
            If (retorno_almoco < LIM_INF_RETORNO_ALMOCO_) Then
                RetornoAlmocoTratado = LIM_INF_RETORNO_ALMOCO_
            'Se não, passa sem tratar
            Else
                RetornoAlmocoTratado = retorno_almoco
            End If
        'Se não, se o retorno do almoço for depois do fim do almoço, passa sem tratar
        ElseIf (retorno_almoco > RETORNO_ALMOCO_) Then
            RetornoAlmocoTratado = retorno_almoco
        'Se não, se a diferença entre saída e retorno do almoço for menor que 1h, usa o valor padrão para o retorno do almoço
        ElseIf (Round2minute(retorno_almoco - saida_almoco) <= Round2minute(TimeValue("1:00"))) Then ' Foi necessário truncar os valores por erro de arredondamento
            RetornoAlmocoTratado = FIM_ALMOCO_PADRAO_1H
        'Se não for nenhum caso especial, passa sem tratar
        Else
            RetornoAlmocoTratado = retorno_almoco
        End If
    End If

End Function


