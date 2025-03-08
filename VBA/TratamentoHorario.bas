Attribute VB_Name = "TratamentoHorario"
Option Explicit

'============================================================='
'''''''' DECLARA��O DAS VARI�VEIS E CONSTANTES GLOBAIS ''''''''
'============================================================='

''''Declara��o dos hor�rios not�veis como constantes globais deste m�dulo
''Comercial
Private Const ENTRADA_COM As Date = #7:30:00 AM#                'TimeValue("7:30")
Private Const SAIDA_ALMOCO_COM As Date = #12:00:00 PM#          'TimeValue("12:00")
Private Const LIM_SUP_SAIDA_ALMOCO_COM As Date = #12:30:00 PM#  'TimeValue("12:30") '=RETORNO_ALMOCO - "1:00"
Private Const LIM_INF_RETORNO_ALMOCO_COM As Date = #1:00:00 PM# 'TimeValue("13:00") '=SAIDA_ALMOCO + "1:00"
Private Const RETORNO_ALMOCO_COM As Date = #1:30:00 PM#         'TimeValue("13:30")
Private Const SAIDA_COM As Date = #5:00:00 PM#                  'TimeValue("17:00")
''Constru��o
Private Const ENTRADA_C As Date = #7:30:00 AM#                  'TimeValue("7:30")
Private Const SAIDA_ALMOCO_C As Date = #11:45:00 AM#            'TimeValue("11:45")
Private Const LIM_SUP_SAIDA_ALMOCO_C As Date = #12:45:00 PM#    'TimeValue("12:45")
Private Const LIM_INF_RETORNO_ALMOCO_C As Date = #12:45:00 PM#  'TimeValue("12:45")
Private Const RETORNO_ALMOCO_C As Date = #1:45:00 PM#           'TimeValue("13:45")
Private Const SAIDA_C As Date = #5:30:00 PM#                    'TimeValue("17:30")
''Constru��o (sexta)
Private Const ENTRADA_CS As Date = #7:00:00 AM#                 'TimeValue("7:00")
Private Const SAIDA_ALMOCO_CS As Date = #11:45:00 AM#           'TimeValue("11:45")
Private Const LIM_SUP_SAIDA_ALMOCO_CS As Date = #12:15:00 PM#   'TimeValue("12:15")
Private Const LIM_INF_RETORNO_ALMOCO_CS As Date = #12:45:00 PM# 'TimeValue("12:45")
Private Const RETORNO_ALMOCO_CS As Date = #1:15:00 PM#          'TimeValue("13:15")
Private Const SAIDA_CS As Date = #4:30:00 PM#                   'TimeValue("16:30")
''Almo�o padr�o de 1 hora (que serve para todos os tipos de pontos e todos os dias)
Private Const INI_ALMOCO_PADRAO_1H As Date = #12:00:00 PM#      'TimeValue("12:00")
Private Const FIM_ALMOCO_PADRAO_1H As Date = #1:00:00 PM#       'TimeValue("13:00")

''''Declara��o das vari�veis est�ticas de hor�rios
Private ENTRADA_                As Date
Private SAIDA_ALMOCO_           As Date
Private LIM_SUP_SAIDA_ALMOCO_   As Date
Private LIM_INF_RETORNO_ALMOCO_ As Date
Private RETORNO_ALMOCO_         As Date
Private SAIDA_                  As Date

''''Declara��o das constantes que ser�o usadas como marcadores para definir com quais hor�rios padr�es as vari�veis ser�o inicializadas
Private Const COM As Integer = 1 'Comercial
Private Const C As Integer = 2   'Constru��o
Private Const CS As Integer = 3  'Constru��o Sexta

''''Declara��o das pseudo-constantes que servir�o para evitar que os Ranges e Names precisem ser acessados e avaliados o tempo todo
Private FDS_ As String
Private FERIAS_ As String
Private FERIADO_ As String
Private DISPENSADO_ As String
Private MEIO_COMPENSADO_ As String
Private EXPEDIENTE_CORRIDO_ As String
'

'============================================================='
''''''''' DECLARA��O DAS FUN��ES E ROTINAS AUXILIARES '''''''''
'============================================================='

'Procedimento para configurar os hor�rios padr�es de acordo com o tipo de ponto e dia
'A ideia dessa rotina � reduzir o n�mero de execu��es de inicializa��o e otimizar o tempo da planilha
Sub DefinirConstanteTempo(dia_semana As String, tipo_ponto As String)

    ''''Declara��o da vari�vel est�tica de flag que indicar� com o atual tipo de ponto e dia (Comercial, Constru��o ou Constru��o-Sexta)
    Static TPD As Integer 'Por padr�o, Inteiros est�ticos s�o inicializados com 0, ent�o na primeira chamada essa fun��o sempre entrar� em algum 'if', colocando algum hor�rio padr�o

    ''''Testa o caso atual
    Select Case tipo_ponto '---> Equivale a: ThisWorkbook.Names("TIPO_DE_PONTO").RefersToRange.Value
        Case "Comercial"
            If TPD <> COM Then 'S� define as vari�veis se o caso mudou da �ltima chamada
                TPD = COM
                ENTRADA_ = ENTRADA_COM
                SAIDA_ALMOCO_ = SAIDA_ALMOCO_COM
                LIM_SUP_SAIDA_ALMOCO_ = LIM_SUP_SAIDA_ALMOCO_COM
                LIM_INF_RETORNO_ALMOCO_ = LIM_INF_RETORNO_ALMOCO_COM
                RETORNO_ALMOCO_ = RETORNO_ALMOCO_COM
                SAIDA_ = SAIDA_COM
            End If
        Case Else '"Constru��o"
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
'A ideia dessa rotina � reduzir o n�mero de execu��es de inicializa��o e otimizar o tempo da planilha
Sub InicializaTiposDia()
    FDS_ = CStr(Evaluate(ThisWorkbook.Names("FIM_DE_SEMANA").RefersTo)) 'N�o � poss�vel obter o FIM_DE_SEMANA pelo Range, j� que ele n�o referencia nenhum range
    FERIAS_ = DADOS.Range("FERIAS").Value
    FERIADO_ = DADOS.Range("FERIADO").Value
    DISPENSADO_ = DADOS.Range("DISPENSADO").Value
    MEIO_COMPENSADO_ = DADOS.Range("MEIO_COMPENSADO").Value
    EXPEDIENTE_CORRIDO_ = DADOS.Range("EXPEDIENTE_CORRIDO").Value
End Sub

'Fun��o auxiliar para saber se o dia � do tipo f�rias (e afins) ou n�o
Function EhFerias(tipo_dia As String) As Boolean
    If FERIAS_ = "" Then InicializaTiposDia 'Testa a inicializa��o das F�RIAS (que � chamada primeiro) porque alguns eventos (como erros n�o tratados) apagam os valores das vari�veis inicializadas
    EhFerias = (tipo_dia = FERIAS_)
    'Ser� aqui que ser�o acrescentados os tipos de dia F�RIAS* e o COMPENSADO_CRE1
End Function

'Fun��o auxiliar para saber se o dia � de um tipo especial ou n�o
Function EhDiaEspecial(tipo_dia As String) As Boolean
    EhDiaEspecial = (tipo_dia = FDS_) _
                 Or (tipo_dia = FERIADO_) _
                 Or (tipo_dia = DISPENSADO_) _
                 Or (tipo_dia = MEIO_COMPENSADO_)
End Function

'Fun��o auxiliar para arredondar os valores de tempo e evitar problemas de lixo de mem�ria em opera��es
'Essa fun��o arredonda os segundos para o minuto mais pr�ximo e � necess�ria porque o Excel considera tempos como doubles (cheio de casas decimais)
'A l�gica dela �: divide o tempo por 1 min para achar (em inteiros) a quantidade de minutos, arredonda essa quantidade e reconverte para o double
'que representa aquele tempo. Assim o valor fica pouco fora do valor incorreto, resetando os erros acumulados nas opera��es
Function Round2minute(tempo As Date) As Date
    Round2minute = Round(tempo / #12:01:00 AM#) * (#12:01:00 AM#) '#12:01:00 AM# == "0:01"
End Function

'============================================================='
''''''' DECLARA��O DAS FUN��ES DE TRATAMENTO PRINCIPAIS '''''''
'============================================================='

''''Fun��o para tratar o hor�rio de entrada
Function EntradaTratada(dia_semana As String, tipo_dia As String, entrada As Date, saida_almoco_tratada As Date, tipo_ponto As String, calendar As Range) As Date
Application.Volatile (False)
'Debug.Print "EntradaTratada"
'
'As vari�veis de entrada 'calendar' e 'tipo_ponto' t�m fun��o apenas de fazer essa UDF (User Defined Function) ser recalculada quando esses valores atualizam

''''Excel base desta fun��o (ao utilizar, lembre-se de que elas podem conter erro de digita��o ou estarem desatualizadas):
'=SE($B2=FERIAS;
'    0;
'    SE(OU($B2=FIM_DE_SEMANA;$B2=FERIADO;$B2=DISPENSADO;$B2=MEIO_COMPENSADO);
'        D2;
'        SE(TIPO_DE_PONTO="Comercial";
'            SE(D2>"12:00"+0;
'                I2;
'                SE(E(D2>="7:00"+0;D2<="7:35"+0);VALOR.TEMPO("7:30");D2));
'            SE(E(TIPO_DE_PONTO="Constru��o";$A2<>"Sexta");
'                SE(D2>"11:45"+0;
'                    I2;
'                    SE(E(D2>="7:00"+0;D2<="7:35"+0);VALOR.TEMPO("7:30");D2));
'                SE(D2>"11:45"+0;
'                    I2;
'                    SE(E(D2>="6:30"+0;D2<="7:05"+0);VALOR.TEMPO("7:00");D2))))))
    
    ''''Filtros iniciais
    'Se for f�rias (ou afins), impede que qualquer valor passe para frente
    If EhFerias(tipo_dia) Then
        EntradaTratada = 0
        Exit Function
    End If
    'Se for algum dia especial, ele n�o trata o hor�rio, deixando passar a informa��o para frente
    If EhDiaEspecial(tipo_dia) Then
        EntradaTratada = entrada
        Exit Function
    End If
    'Na entrada, o dia de Expediente Corrido � tratado como um dia normal
    
    ''''Configura os hor�rios padr�es de acordo com o tipo de ponto e dia
    Call DefinirConstanteTempo(dia_semana, tipo_ponto)
    
    ''''Tratamento principal
    'Se o usu�rio fez apenas a parte da tarde + hora extra no almo�o (entra depois da sa�da para almo�o), ent�o a entrada recebe a sa�da para almo�o para zerar o turno da manh�
    If (entrada > SAIDA_ALMOCO_) Then
        EntradaTratada = saida_almoco_tratada
    'Se o usu�rio bateu ponto at� meia hora antes ou at� 5min depois do hor�rio de entrada, arredonda para o in�cio do expediente
    ElseIf (Round2minute(entrada) >= Round2minute(ENTRADA_ - TimeValue("0:30"))) And (Round2minute(entrada) <= Round2minute(ENTRADA_ + TimeValue("0:05"))) Then ' Foi necess�rio truncar os valores por erro de arredondamento
        EntradaTratada = ENTRADA_
    'O comportamento padr�o � n�o tratar o ponto
    Else
        EntradaTratada = entrada
    End If
        
End Function

''''Fun��o para tratar o hor�rio de sa�da
Function SaidaTratada(dia_semana As String, tipo_dia As String, horarios As Range, retorno_almoco_tratado As Date, tipo_ponto As String, calendar As Range) As Date
Application.Volatile (False)
'Debug.Print "SaidaTratada"
'
'As vari�veis de entrada 'calendar' e 'tipo_ponto' t�m fun��o apenas de fazer essa UDF (User Defined Function) ser recalculada quando esses valores atualizam

''''Excel base desta fun��o (ao utilizar, lembre-se de que elas podem conter erro de digita��o ou estarem desatualizadas):
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
'                SE(E(TIPO_DE_PONTO="Constru��o";$A2<>"Sexta");
'                    SE(G2<"13:45"+0;
'                        J2;
'                        SE(E(G2>="17:25"+0;G2<="17:30"+0);VALOR.TEMPO("17:30");G2));
'                    SE(G2<"13:15"+0;
'                        J2;
'                        SE(E(G2>="16:25"+0;G2<="16:30"+0);VALOR.TEMPO("16:30");G2)))))))

    ''''Declara��o das vari�veis
    Dim saida_almoco As Date:     saida_almoco = horarios.Cells(1, 2).Value
    Dim retorno_almoco As Date:   retorno_almoco = horarios.Cells(1, 3).Value
    Dim saida As Date:            saida = horarios.Cells(1, 4).Value

    ''''Filtros iniciais
    'Se for f�rias (ou afins), impede que qualquer valor passe para frente
    If EhFerias(tipo_dia) Then
        SaidaTratada = 0
        Exit Function
    End If
    'Se for algum dia especial, ele n�o trata o hor�rio, deixando passar a informa��o para frente
    If EhDiaEspecial(tipo_dia) Then
        SaidaTratada = saida
        Exit Function
    End If
    'Se for um dia de Expediente Corrido sem marca��o de almo�o, passa a informa��o sem tratar (se o ponto do almo�o foi batido, o dia corrido � considerado um dia normal)
    If (tipo_dia = EXPEDIENTE_CORRIDO_) And (saida_almoco = 0) And (retorno_almoco = 0) Then
        SaidaTratada = saida
        Exit Function
    End If
    
    ''''Configura os hor�rios padr�es de acordo com o tipo de ponto e dia
    Call DefinirConstanteTempo(dia_semana, tipo_ponto)
    
    ''''Tratamento principal
    'Se o usu�rio fez apenas a parte da manh� + hora extra no almo�o (sai antes do retorno do almo�o), ent�o a sa�da recebe o retorno do almo�o para zerar o turno da tarde
    If (saida < RETORNO_ALMOCO_) Then
        SaidaTratada = retorno_almoco_tratado
    'Se o usu�rio bateu ponto na nos �ltimos 5min do hor�rio, arredonda para o fim do expediente
    ElseIf (Round2minute(saida) >= Round2minute(SAIDA_ - TimeValue("00:05"))) And (saida <= SAIDA_) Then ' Foi necess�rio truncar os valores por erro de arredondamento
        SaidaTratada = SAIDA_
    'O comportamento padr�o � n�o tratar o ponto
    Else
        SaidaTratada = saida
    End If

End Function

''''Fun��o para tratar o hor�rio de sa�da para almo�o
Function SaidaAlmocoTratada(dia_semana As String, tipo_dia As String, horarios As Range, tipo_ponto As String, calendar As Range) As Date
Application.Volatile (False)
'Debug.Print "SaidaAlmocoTratada"
'
'As vari�veis de entrada 'calendar' e 'tipo_ponto' t�m fun��o apenas de fazer essa UDF (User Defined Function) ser recalculada quando esses valores atualizam
'
''''Excel base desta fun��o (ao utilizar, lembre-se de que elas podem conter erro de digita��o ou estarem desatualizadas):
'=SE($B2=FERIAS;
'    0;
'    SE(OU($B2=FIM_DE_SEMANA;$B2=FERIADO;$B2=DISPENSADO;$B2=MEIO_COMPENSADO);
'        $E2;
'        SE(E(�C�L.VAZIA($E2);�C�L.VAZIA($F2));
'            SE(E(�C�L.VAZIA($D2);�C�L.VAZIA($G2));
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

    ''''Declara��o das vari�veis
    Dim entrada As Date:          entrada = horarios.Cells(1, 1).Value
    Dim saida_almoco As Date:     saida_almoco = horarios.Cells(1, 2).Value
    Dim retorno_almoco As Date:   retorno_almoco = horarios.Cells(1, 3).Value
    Dim saida As Date:            saida = horarios.Cells(1, 4).Value
    
    ''''Filtros iniciais
    'Se for f�rias (ou afins), impede que qualquer valor passe para frente
    If EhFerias(tipo_dia) Then
        SaidaAlmocoTratada = 0
        Exit Function
    End If
    'Se for algum dia especial, ele n�o trata o hor�rio, deixando passar a informa��o para frente
    If EhDiaEspecial(tipo_dia) Then
        SaidaAlmocoTratada = saida_almoco
        Exit Function
    End If

    ''''Configura os hor�rios padr�es de acordo com o tipo de ponto e dia
    Call DefinirConstanteTempo(dia_semana, tipo_ponto)
    
    ''''Tratamento principal (perceba que o tratamento padr�o � passar o valor sem tratamento: SaidaAlmocoTratada = saida_almoco)
    'Se o almo�o estiver vazio
    If ((saida_almoco = 0) And (retorno_almoco = 0)) Then
        'Se o resto todo estiver vazio, passa (o vazio) sem tratar
        If ((entrada = 0) And (saida = 0)) Then
            SaidaAlmocoTratada = saida_almoco
        'Se for um dia de Expediente Corrido, anula o almo�o (se o ponto do almo�o foi batido, o dia corrido � considerado um dia normal)
        ElseIf (tipo_dia = EXPEDIENTE_CORRIDO_) Then
            SaidaAlmocoTratada = 0
        'Se apenas entrada e sa�da estiverem preenchidas
        Else
            'Se a sa�da for antes do fim do turno da manh� (antes do limite superior de sa�da para almo�o), passa a sa�da para a sa�da do almo�o sem tratar
            If (saida <= LIM_SUP_SAIDA_ALMOCO_) Then
                SaidaAlmocoTratada = saida
            'Se a sa�da for maior que o limite superior de sa�da, por�m menor que o limite (superiror) de retorno do almo�o, trava a sa�da do almo�o no limite superior de sa�da para almo�o
            ElseIf (saida <= RETORNO_ALMOCO_) Then
                SaidaAlmocoTratada = LIM_SUP_SAIDA_ALMOCO_
            'Se n�o, isto �, se a sa�da for depois do almo�o, trava a sa�da do almo�o no limite inferior de sa�da para almo�o para que ela n�o influencie em hora extra
            Else
                SaidaAlmocoTratada = SAIDA_ALMOCO_
            End If
        End If
    'Se o almo�o n�o estiver vazio (sup�e que entrada e sa�da tamb�m n�o est�o)
    Else
        'Se a sa�da para almo�o for antes do in�cio do almo�o, passa sem tratar
        If (saida_almoco < SAIDA_ALMOCO_) Then
            SaidaAlmocoTratada = saida_almoco
        'Se n�o, se o retorno do almo�o for depois do fim do almo�o
        ElseIf (retorno_almoco > RETORNO_ALMOCO_) Then
            'Se a sa�da para almo�o for depois do limite para sa�da do almo�o, trava a sa�da do almo�o no limite de sa�da para almo�o
            If (saida_almoco > LIM_SUP_SAIDA_ALMOCO_) Then
                SaidaAlmocoTratada = LIM_SUP_SAIDA_ALMOCO_
            'Se n�o, passa sem tratar
            Else
                SaidaAlmocoTratada = saida_almoco
            End If
        'Se n�o, se a diferen�a entre sa�da e retorno do almo�o for menor que 1h, usa o valor padr�o para a sa�da para almo�o
        ElseIf (Round2minute(retorno_almoco - saida_almoco) <= Round2minute(TimeValue("1:00"))) Then ' Foi necess�rio truncar os valores por erro de arredondamento
            SaidaAlmocoTratada = INI_ALMOCO_PADRAO_1H
        'Se n�o for nenhum caso especial, passa sem tratar
        Else
            SaidaAlmocoTratada = saida_almoco
        End If
    End If

End Function

''''Fun��o para tratar o hor�rio de retorno do almo�o
Function RetornoAlmocoTratado(dia_semana As String, tipo_dia As String, horarios As Range, tipo_ponto As String, calendar As Range) As Date
Application.Volatile (False)
'Debug.Print "RetornoAlmocoTratado"
'
'As vari�veis de entrada 'calendar' e 'tipo_ponto' t�m fun��o apenas de fazer essa UDF (User Defined Function) ser recalculada quando esses valores atualizam

''''Excel base desta fun��o (ao utilizar, lembre-se de que elas podem conter erro de digita��o ou estarem desatualizadas):
'=SE($B2=FERIAS;
'    0;
'    SE(OU($B2=FIM_DE_SEMANA;$B2=FERIADO;$B2=DISPENSADO;$B2=MEIO_COMPENSADO);
'        $F2;
'        SE(E(�C�L.VAZIA($E2);�C�L.VAZIA($F2));
'            SE(E(�C�L.VAZIA($D2);�C�L.VAZIA($G2));
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

    ''''Declara��o das vari�veis
    Dim entrada As Date:          entrada = horarios.Cells(1, 1).Value
    Dim saida_almoco As Date:     saida_almoco = horarios.Cells(1, 2).Value
    Dim retorno_almoco As Date:   retorno_almoco = horarios.Cells(1, 3).Value
    Dim saida As Date:            saida = horarios.Cells(1, 4).Value
    
    ''''Filtros iniciais
    'Se for f�rias (ou afins), impede que qualquer valor passe para frente
    If EhFerias(tipo_dia) Then
        RetornoAlmocoTratado = 0
        Exit Function
    End If
    'Se for algum dia especial, ele n�o trata o hor�rio, deixando passar a informa��o para frente
    If EhDiaEspecial(tipo_dia) Then
        RetornoAlmocoTratado = retorno_almoco
        Exit Function
    End If
    
    ''''Configura os hor�rios padr�es de acordo com o tipo de ponto e dia
    Call DefinirConstanteTempo(dia_semana, tipo_ponto)
    
    ''''Tratamento principal (perceba que o tratamento padr�o � passar o valor sem tratamento: RetornoAlmocoTratado = retorno_almoco)
    'Se o almo�o estiver vazio
    If ((saida_almoco = 0) And (retorno_almoco = 0)) Then
        'Se o resto todo estiver vazio, passa (o vazio) sem tratar
        If ((entrada = 0) And (saida = 0)) Then
            RetornoAlmocoTratado = retorno_almoco
        'Se for um dia de Expediente Corrido, anula o almo�o (se o ponto do almo�o foi batido, o dia corrido � considerado um dia normal)
        ElseIf (tipo_dia = EXPEDIENTE_CORRIDO_) Then
            RetornoAlmocoTratado = 0
        'Se apenas entrada e sa�da estiverem preenchidas
        Else
            'Se a entrada for depois do come�o do turno da tarde (depois do limite inferior de retorno do almo�o), passa a entrada para o retorno do almo�o sem tratar
            If (entrada >= LIM_INF_RETORNO_ALMOCO_) Then
                RetornoAlmocoTratado = entrada
            'Se a entrada for menor que o limite inferior de retorno, por�m maior que o limite (inferior) de sa�da para almo�o, trava o retorno do almo�o no limite inferior de retorno do almo�o
            ElseIf (entrada >= SAIDA_ALMOCO_) Then
                RetornoAlmocoTratado = LIM_INF_RETORNO_ALMOCO_
            'Se n�o, isto �, se a entrada for antes do almo�o, trava o retorno do almo�o no limite superior de retorno do almo�o para que ele n�o influencie em hora extra
            Else
                RetornoAlmocoTratado = RETORNO_ALMOCO_
            End If
        End If
    'Se o almo�o n�o estiver vazio (sup�e que entrada e sa�da tamb�m n�o est�o)
    Else
        'Se a sa�da para almo�o for antes do in�cio do almo�o
        If (saida_almoco < SAIDA_ALMOCO_) Then
            'Se o retorno do almo�o for antes do limite para retorno do almo�o, trava o retorno do almo�o no limite de retorno do almo�o
            If (retorno_almoco < LIM_INF_RETORNO_ALMOCO_) Then
                RetornoAlmocoTratado = LIM_INF_RETORNO_ALMOCO_
            'Se n�o, passa sem tratar
            Else
                RetornoAlmocoTratado = retorno_almoco
            End If
        'Se n�o, se o retorno do almo�o for depois do fim do almo�o, passa sem tratar
        ElseIf (retorno_almoco > RETORNO_ALMOCO_) Then
            RetornoAlmocoTratado = retorno_almoco
        'Se n�o, se a diferen�a entre sa�da e retorno do almo�o for menor que 1h, usa o valor padr�o para o retorno do almo�o
        ElseIf (Round2minute(retorno_almoco - saida_almoco) <= Round2minute(TimeValue("1:00"))) Then ' Foi necess�rio truncar os valores por erro de arredondamento
            RetornoAlmocoTratado = FIM_ALMOCO_PADRAO_1H
        'Se n�o for nenhum caso especial, passa sem tratar
        Else
            RetornoAlmocoTratado = retorno_almoco
        End If
    End If

End Function


