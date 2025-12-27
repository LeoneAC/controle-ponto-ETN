### :hash: <span style="color:#007AFF">Tags para pesquisa: `#BUG` `#FEATURE`</span>

---
---
---

# :package: `[v1.0.0]` <span style="color:#9900FF">(03/01/2025)</span>

##  **Versão inicial com Macro** **`[v1.0.0m]`**
  _Pendente de ser explicada_: quando tiver tempo, eu gostaria de fazer uma documentação mais detalhada da versão para que haja um compilado de recursos disponíveis na planilha, mas, de qualquer forma, ela foi criada para ser autoexplicativa, contando com uma área de _Informações_ na aba **DADOS** e exemplos de utilização na aba **EXEMPLO**.

## **Versão inicial sem Macro** **`[v1.0.0]`**

Fiz essa versão sem macros para aumentar ao máximo a compatibilidade com qualquer computador e para haver uma versão estável, que dará menos bugs por não ter macros envolvidas na configuração da planilha. No entanto, devido ao trabalho para manter as duas versões, a intenção é descontinuá-la em breve.

> [!TIP]
> Exatamente como a versão [`[v1.0.0m]`](#versão-inicial-com-macro-v100m), porém sem a macro para criar um novo ano (sem o botão).

> [!NOTE]
> Essa versão será postada no GitHub apenas no _release_.

---
---
---

# :package: `[v1.1.0]` <span style="color:#9900FF">(13/01/2025)</span>

Essa versão foi lançada rapidamente para corrigir o bug que causava erros graves na versão do Excel 2016, mas traz outras melhorias além da correção.

## :pencil2: <span style="color:#FFFFFF">Mudanças</span>

1. Versões `[v1.1.0]` (sem macro) e `[v1.1.0m]` (com macro) unificadas:
   - Estabeleci a versão com Macro como base e, já que a diferença entre elas passou a ser basicamente a parte 1 da aba **DADOS**, resolvi descontinuar a versão sem macros.
1. **`#FEATURE`** Adição do tipo de dia **FÉRIAS** que bloqueia qualquer marcação no seu dia (zera as horas no tratamento).
   - Um erro também é mostrado nos campos de preenchimento de ponto (em amarelo).
1. Tratamento dos tipos de dias **DISPENSADO** e **1/2 COMPENSADO** como feriados e fins de semana.
1. **`#BUG`** Retirada do `@` nas fórmulas e nas formatações condicionais para melhorar compatibilidade com versões antigas do Excel (2016, por exemplo).
1. **`#FEATURE`** Alteração do tratamento da _Saída para Almoço_ e da _Entrada do Almoço_ para permitir que esses horários sejam deixados em branco em dias normais.
1. Melhoria das explicações na aba **EXEMPLO**.
1. Acréscimo e retirada de exemplos na aba **EXEMPLO**.
    <ol type="I">
      <li>Acréscimo do uso de férias (tipo de dia **FÉRIAS**).</li>
      <li>Retirada do exemplo de erro na falta de tratamento.</li>
    </ol>
1. **`#FEATURE`** Exibição de erro na coluna de _Horas Trabalhadas_ quando ela está negativa.
1. Melhoria na explicação do texto de _Instruções_ da aba **DADOS**.

## :x: $\textcolor{red}{\textsf{Bugs conhecidos}}$

### `[b1]` `#BUG` _{Encontrado pelo Igor Jaloto}_

Na versão mais recente do Excel (Office 365), fazer o passo a passo de copiar e colar com:
   1. `Ctrl+V`
   1. depois: `Ctrl`
   1. depois: `V`

gera uma mensagem de aviso de fórmula recursiva no meio do processo.

Para resolver isso, foi possível fazer um atalho `Ctrl+Shift+V` com macro para contornar o problema (versão [`[v1.1.1m]`](#package-v111m)), mas ele exclui o histórico de `Ctrl+Z`. Por isso essa ideia ainda não foi aplicada.

Outra maneira de contornar o problema é usando o mouse, com **Colar especial** diretamente, mas esse método é pouco prático.

> [!IMPORTANT]
> Se esse caso acontecer com você, basta clicar em `OK` na mensagem que aparece na tela e continuar o processo de colar somente valores.

### `[b2]` `#BUG` _{Encontrado pelo Leone}_

No tratamento do horário de almoço, há um bug se a pessoa bater ponto `11:55` e `12:55` (no horário comercial, por exemplo), porque a tabela manterá os valores como estão (almoço == 1h), mas o SAP arredonda o `12:55` para `13:00` e gera hora negativa de `11:55` a `12:00`.

> [!WARNING]
> Se esse caso acontecer com você, será necessário manipular os horários manualmente sobrescrevendo as células nas colunas ocultas de _Tratamento de horários_ com valores manuais.

---

## :package: `[v1.1.1m]`

Tentativa de correção do bug [`[b1]`](#b1-bug-encontrado-pelo-igor-jaloto) da versão [`[v1.1.0]`](#package-v110-13012025) com macro no atalho `Ctrl+Shift+V` para colar apenas valores.

> [!NOTE]
>  Essa versão não será postada no GitHub.

---
---
---

# :package: `[v1.2.0-excel]` <span style="color:#9900FF">(10/02/2025)</span>

Essa versão, bora tenha demorado a ser oficialmente lançada, foi feita principalmente para corrigir os bugs encontrados pelos colegas.

Apesar disso, ela traz muitas outras melhorias e já esquematiza maneiras de implementar features sugeridas.

## :pencil2: <span style="color:#FFFFFF">Mudanças</span>

### &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Alterações de interface e usabilidade

1. Atualização de notas explicativas para melhorar entendimento:
   - Coluna **Tipo de dia**
   - Coluna **ABNO ou ENF1 Abonos**
   - Coluna **NAUT Presença Não Perm.**
   - Células de totais
1. Alteração nas *Instruções* da planilha **DADOS**:
   - Atualização/adição diversas nos textos das instruções.
   - Adição de uma sugestão de prompt para usar o ChatGPT-4.0 para gerar o _Calendário_.
     - *Curiosidade:* o prompt foi gerado com DeepSeek-V3, mas funcionou muito melhor no ChatGPT-4.0.
1. Adição do código **ENF1** na *Legenda* da planilha **DADOS**.
1. Troca das colunas **AUS ou JIN** de lugar com **ABNO**.
1. Troca do título da coluna **ABNO Abono gerencial** para **ABNO ou ENF1 Abonos**.
1. Atualização do texto de legenda da célula de saldo total para melhorar o entendimento.
1. Correção e melhoria do texto da seção *Novo Ano*.
1. Atualização de diversos itens na planilha **EXEMPLO** para melhorar o entendimento.
1. Redesign da planilha **EXEMPLO** para diminuir a poluição visual.
1. Adição de exemplos na planilha **EXEMPLO**:
    - Exemplo de divergência esperada de valores com o programa de pontos da Empresa.
        > :bulb: $\textcolor{#66CC00}{\textsf{Dica}}$\
        > É possível que as **Variações de Horas do Almoço** contenham a soma dos tratamentos do SAP. Nesse exemplo de marcação abaixo, em um ponto Comercial, você vê `+0:10` no Excel, mas esse valor é o resultado de `+0:30` hora extra do almoço somado a `-0:20` de atraso por ter saído antes de `12:00`. No SAP, você verá os dois horários e não apenas o resultado da soma.
        > 
        > | Entrada | Saída almoço| Retorno almoço | Saída |
        > | :-----: | :---------: | :------------: | :---: |
        > |  07:30  |    11:40    |      12:40     | 17:00 |
    - Exemplo de _bypass_ do tratamento (alteração manual do valor da célula nas colunas ocultas auxiliares de tratamento de pontos).
        > :bulb: $\textcolor{#66CC00}{\textsf{Dica}}$\
        > Se você precisar bypassar (passar por cima, anular) o tratamento de marcações de pontos, você pode desocultar as colunas ocultas auxiliares de tratamento de pontos, desproteger a planilha e inserir um valor manualmente numa célula. Nesse exemplo abaixo, se você quisesse tratar manualmente no SAP o tempo de `+0:29`, o tratamento automático do Excel continuaria arredondando o valor de `07:01` para `07:30`, portanto seria preciso substituir manualmente a fórmula da célula pelo horário `07:01`.
        > 
        > | Entrada | Saída almoço| Retorno almoço | Saída |
        > | :-----: | :---------: | :------------: | :---: |
        > |  07:01  |    12:00    |      13:30     | 17:00 |
        >
        > **Mas <ins>atenção</ins>:** Lembre-se de restaurar a fórmula da célula quando for criar uma planilha para o ano seguinte. [Veja a anotação abaixo](#atalho-shift).
1. **`#FEATURE`** Formatar células ocultas auxiliares de tratamento de pontos para ficarem coloridas quando não contiverem mais fórmulas.
    - Como o botão **Novo Controle de Horas** não reseta as fórmulas para que o usuário não perca a personalização que fez para si, uma alteração manual em alguma dessas células pode causar erro no ano seguinte. Destacar a célula que foi alterada para um valor fixo pode ajudar a identificar essas alterações.
      > :rotating_light: $\textcolor{red}{\textsf{Cuidado!}}$\
      > <a name="atalho-shift"></a>
      > Se o usuário alterar alguma célula nas colunas cinzas, que não são feitas para serem alteradas, ou, principalmente, nas colunas ocultas auxiliares de tratamento de pontos (onde é mais comum que o usuário faça alguma mudança manual), lembre-se de que, no ano seguinte, quando clicar no botão **Novo Controle de Horas**, essas alterações <ins>não serão desfeitas</ins> e, se necessário, as células precisam ter suas fórmulas reestabelecidas pelo próprio usuário. Se o usuário quiser uma maneira prática de reestabelecer a planilha ao seu "padrão de fábrica" (_default_), basta deletar todos as abas de meses antes de clicar no botão. Caso contrário, se o usuário quiser manter suas personalizações, mas corrigindo as células alteradas, ele precisa corrigir manualmente, copiando a fórmula de outra célula análoga não alterada para reestabelecer o comportamento padrão da célula alterada.

    ### Alterações diversas 

1. **`#BUG`** Correção da planilha **EXEMPLO** para que ela fique fixa nos exemplos do Tipo de Ponto Comercial.
1. Definição de Nomes (variáveis) para os Tipos de dia **FIM DE SEMANA** e **DIA ÚTIL** para manter um paralelismo na programação.
1. **`#BUG`** Correção definitiva de arredondamentos com a criação de uma função (`Round2minute`) para arredondar para o minuto mais próximo e eliminar acúmulo de erros de arredondamento que o Excel causa em operações com tempo. Essa mudança foi feita devido a alguns fatores:
    - Antes eu estava usando a função `=TRUNCAR(<tempo>,6)`, mas ela não resolveu alguns dos problemas, além de não ter um correspondente no VBA.
    - As **Variações de Horas do Almoço** na Construção apareciam com `00:00`.
    - A função `EntradaTratada` dava problema na sexta-feira ao receber um horário de `06:30` no ponto da Construção (não arredondava para `07:00`).
    - O valor do **Saldo total de horas livres no mês** acumulava erro de segundos e mostrava `16:20:59` ao invés de `16:21:00`.
1. **`#FEATURE`** Implementação do atalho `Shift+Ctrl+V` para _"Colar como Valores"_ diretamente.
    - O problema deste atalho é que ele "mata" o histórico de `Ctrl+Z` antes dele, então não é possível mais desfazer as ações da planilha de antes do atalho ser usado após usá-lo.
    - Esse atalho não é mais tão necessário, uma vez que o bug [`[b1]`](#b1-bug-encontrado-pelo-igor-jaloto) já foi corrigido.
    - Essa implementação engloba a ideia da versão [`[v1.1.1m]`](#package-v111m).
    > :warning: $\textcolor{#FFCC00}{\textsf{Aviso!}}$\
    > Se você usar o atalho `Shift+Ctrl+V` para _"Colar como Valores"_ diretamente, tenha ciência de que todo o histórico de modificações que podem ser desfeitas com `Ctrl+Z`, até o momento em que o atalho foi usado, será perdido.
1. Melhoria da robustez da macro `LimpaPlanilha` ao trocar `Sheets("DADOS")` por `DADOS`, etc.
1. `#BUG` Correção do comportamento do dia de **EXPEDIENTE CORRIDO** para seu tratamento original da versão `[v1.0.0m]`.
    - A versão [`[v1.1.0]`](#package-v110-13012025) havia gerado um problema no tratamento desse tipo de dia.
1. Alteração retroativa do nome das versões para o padrão de versionamento: `MAJOR.MINOR.PATCH`

## :heavy_check_mark: $\textcolor{#66CC00}{\textsf{Bugs corrigidos}}$

### `[b1]` `#BUG` _{Encontrado pelo Igor Jaloto}_

_Clique [aqui](#b1-bug-encontrado-pelo-igor-jaloto) para explicação do bug_ `[b1]`.

Foi possível corrigir o problema travando as células (formato `$A1`) das colunas à esquerda das células da coluna **Ausência para Comp. no mês** e liberando (formato `A1`) as células das colunas à direita desta.\
O mesmo procedimento foi feito para a coluna **CRE1 Compens. no mês** para evitar futuros problemas, embora esta não tenha apresentado problema porque as colunas à esquerda estão muito longe.

O atalho `Shift+Ctrl+V` foi mantido da versão [`[v1.1.1m]`](#package-v111m) por enquanto porque ele evita muitos passos na hora de colar como valores.

### `[b2]` `#BUG` _{Encontrado pelo Leone}_

_Clique [aqui](#b2-bug-encontrado-pelo-leone) para explicação do bug_ `[b2]`.

Esse bug foi corrigido junto com o [`[b4]`](#b4-bug-encontrado-pelo-diogo-costa), através da implementação inicial da feature [`[f5]`](#f5-feature-solicitada-pelo-leone).

### `[b3]` `#BUG` _{Encontrado pela Maria Tormin}_

Fórmula da célula de total estava levando em conta as colunas **ABNO** e **NAUT**.

Foi feita a correção da fórmula da célula de saldo total, retirando as colunas de **ABNO** e **NAUT** da conta.

> [!NOTE]
> O comportamento desse saldo é interpretativo, a ideia dele é indicar visualmente se você precisa fazer mais horas ou não para pagar o que está devendo. Como você consegue manipular os diversos tratamentos, todos eles entram nessa conta, com exceção de **ABNO** e **NAUT**:
> - **Horas não tratadas:** são saldos feitos justamente para serem manipulados entre tratamentos;
> - **CRE1, CRE2, CPHM:** são saldos manipuláveis durante o mês, então eles podem ser convertidos entre si;
> - **Pagamento de horas extras e descontos:** são horas que, teoricamente, você consegue trabalhar para ter ou para compensar, respectivamente;
> - **Abonos e Horas extras não permitidas (ABNO e NAUT):** são horas ganhas/perdidas, você não é capaz de convertê-las para outros tratamentos. Ou perde com **NAUT**, ou ganha com **ABNO**.
> 
> Para que isso fique mais claro, pense que se você está para receber 2h de horas extras no mês, mas se atrasa no último dia; nesse caso você passaria **PGHE** para **CRE1**. Por outro lado, suponha que estava devendo 2h de **CPHM**, mas precisou ir no médico e recebeu 4h de **ENF1**; nesse caso, o **ENF1** não afeta sua situação de ainda ter que compensar 2h de **CPHM**.

### `[b4]` `#BUG` _{Encontrado pelo Diogo Costa}_

Ao lançar só o período do almoço na entrada e saída, a planilha falha, errando os valores de horas positivas e negativas, por exemplo:
| Entrada | Saída almoço| Retorno almoço | Saída |
| :-----: | :---------: | :------------: | :---: |
|  06:51  |             |                | 11:45 | 

Esse bug foi corrigido junto com o [`[b2]`](#b2-bug-encontrado-pelo-leone), através da implementação inicial da feature [`[f5]`](#f5-feature-solicitada-pelo-leone).

## :x: $\textcolor{red}{\textsf{Bugs conhecidos}}$

Nenhum bug conhecido até o momento.

## :sparkles: $\textcolor{#00CCCC}{\textsf{Features implementadas}}$

### `[f5]` `#FEATURE` _{Solicitada pelo Leone}_

> Usar UDF (_User Defined Functions_) para transformar os códigos das células em funções do VBA, pois as coisas estão crescendo e o tratamento de possibilidades está ficando incompreensível.

Foram feitos códigos em VBA para as 4 colunas de pontos tratados para resolver os bugs [`[b2]`](#b2-bug-encontrado-pelo-leone) e [`[b4]`](#b4-bug-encontrado-pelo-diogo-costa). A partir daqui, quando necessário, basta fazer o mesmo para outras colunas.\
Foram criados ao todo 4 algoritmos que influenciaram nos valores das 4 colunas dos horários tratados (a descrição está no código VBA da Pasta de Trabalho e no arquivo [Algoritmo_tratamento_horarios.md](https://github.com/LeoneAC/controle-ponto-ETN/blob/main/Algoritmo_tratamento_horarios.md)).

Gerou-se um problema de otimização no código porque, devido ao fato de as funções dependerem das colunas **Dia** e **Tipo de Dia**, as quais são fórmulas dependentes da coluna **Data**; as funções estavam sendo chamadas para <ins>todas</ins> as linhas sempre que algum horário era lançado. Esse problema só foi resolvido ao se colocar as UDFs explicitamente como não voláteis (`Application.Volatile (False)`) e fazer as funções dependerem do _Tipo de Ponto_ e do _Calendário_ diretamente para que as mudanças destes (sempre que ocorressem) afetassem os tratamentos de pontos.

Usar evento de `Worksheet_Change` se mostrou menos eficiente nesse caso porque as UDFs só atualizavam com `Application.CalculateFull`, comando que faz <ins>todas</ins> as planilhas de <ins>todos</ins> os Excels abertas serem totalmente recalculadas.\
No fim, trabalhar sem UDF, só com funções Excel, provou ser uma das, senão a, maneiras mais otimizadas de tratar o ponto, todavia esse código equivalente ficou tão grande que vários outros problemas surgiram, como: falta de legibilidade, problemas para manutenções futuras e aumento de chance de erros de digitação como troca de horários padrões. Por isso, optei por sacrificar levemente o desempenho da planilha em prol de um código mais compreensível. Além disso, fiz um *rework* do código usando variáveis e constantes globais para reduzir inicializações excessivas e reduzir o custo computacional da planilha.

## :bulb: $\textcolor{#FFCC00}{\textsf{Features pendentes}}$

### `[f1]` `#FEATURE` _{Solicitada pelo Igor Jaloto}_

> Criação de um tipo de dia **COMPENSADO** para **CRE1** (o atual é **CRE2**) para ser inserido no _Calendário_ da planilha **DADOS**.

É possível de ser feito, mas exige cuidado para que as pessoas não se confundam com os tipos de dias. Será necessário:
- Criar nome na tabela _Legenda_ da planilha **DADOS**: **Compensado (CRE1)**.
- Alterar na tabela _Legenda_:
  - **Compensado** (só o texto) para **Compensado (CRE2)**.
  - **1/2 Compensado** (só o texto) para **1/2 Compensado (CRE2)**.
- Estudar a possibilidade de alterar em código, **COMPENSADO** para **COMPENSADO2**, para que possa ser usado **COMPENSADO1** ao invés de **COMPENSADO_CRE1**.
- Acrescentar uma linha de exemplo na planilha **EXEMPLO**.
- Alterar as *Instruções* da planilha **DADOS**.
- Adicionar regras de formatação condicional para o novo dia.
- Fazer os ajustes necessários de comportamento das funções com o novo tipo de dia:
  - Alterar a lógica das células que, em geral, será basicamente adicionar uma exceção do mesmo comportamento do tipo da exceção de **FERIAS**.
  - Alterar as funções `EhFerias` e `Inicializar`.

### `[f2]` `#FEATURE` _{Solicitada pelo Erik}_

> Exigir que apenas os dias iniciais e finais das férias sejam colocados no _Calendário_ da planilha **DADOS**, ao invés de cada um dos dias das férias.

É um recurso possível de ser feito, mas daria algum trabalho. Pensei num jeito de preencher o meio com o tipo de dia **FERIAS** se a célula de cima for **FERIAS** até que se encontre outro **FERIAS**. Já fiz um esboço inicial de como será o algoritmo para isso e será necessária a criação de um novo tipo de dia auxiliar chamado, por exemplo, de **FERIAS\***.\
Além disso, será necessário:
- Criar o tipo de dia **FERIAS\*** auxiliar não acessível ao usuário e que, a princípio, terá comportamento parecido com o **FIM_DE_SEMANA**.
- Alterar as *Instruções* da planilha **DADOS**.
- Acrescentar o novo tipo de dia na formatação condicional de **FERIAS**.
- Alterar a lógica da coluna `B` **Tipo de dia**.
- Acrescentar uma lógica para corrigir a exibição de texto na coluna auxiliar oculta `Y`.
- Fazer os ajustes necessários de comportamento das funções com o novo tipo de dia.
- Acrescentar **FERIAS\*** na função `EhFerias`.

### `[f3]` `#FEATURE` _{Solicitada pelo Diogo Costa}_

> Adição de uma coluna **Duração** que pudesse ser mostrada em formato decimal do Excel (`0,00`) para comparar com o SAP, isto é: `h,(mm x 100/60)`.

É possível de ser feito, o problema aqui é que seriam necessárias pelo menos 3 colunas (uma para cada variação de horas). Além disso, essa *feature* não resolve o problema de quando o SAP tem mais de 3 linhas e tratamento, por exemplo, uma pessoa que, no mesmo dia, chega tarde, sai cedo para o almoço, volta cedo do almoço e sai tarde; teriam 4 linhas para tratamento no SAP e uma das colunas seria a soma de duas.

No lugar dessa ideia, seria melhor, talvez, criar um botão de *toggle* que alterna entre esses dois formatos: decimal e horário.\
O que fazer:
- Criar o botão.
- Quando novas planilhas forem feitas, será necessário:
    - conferir os intervalos da formatação condicional e do VBA.
    - conferir como o botão será duplicado, se funcionará perfeitamente.
- Mostrar os intervalos de variação como "h,(mm x 100/60)" ao invés de hh:mm quando o botão for acionado.
- Decidir se os totais terão essa mudança de formatação também (a princípio, acredito que seria melhor que sim).
- Cuidar para que o clique force o recálculo das células afetadas.
- Adicionar exemplos na planilha **EXEMPLO**.
- Conversões necessárias:
  - Hora -> Decimal: `([h]:mm)*24` e formatar como `0,00`.
  - Decimal -> Hora: `0,00/24` e formatar como `[h]:mm`.

### `[f4]` `#FEATURE` _{Solicitada pelo Diogo Costa}_

> Automação das marcações manuais.

É um recurso difícil de implementar, seria possível fazer uma formatação condicional para colorir as marcações manuais, mas mais que isso seria necessário um poder de programação que exigiria algum tempo no VBA. Mesmo que através de algum botão, seria um recurso com algumas falhas para alocar a hora manual no meio das horas marcadas com ponto e acredito que muito pouco usado. O VBA também já está sobrecarregando um pouco a planilha e acho que colocar esse recurso automático e sem botão não seria muito saudável para o desempenho da planilha, pois exigiria monitorar o evento de mudança da planilha: `Worksheet_Change`.

No recurso de formatação condicional, seria necessário:
- Acrescentar um exemplo na planilha **EXEMPLO** para mostrar como a formatação funciona.
- Usar a seguinte regra de formatação:
    ```
    =OU(B1=SE($A1=$A$10:$A$14;$B$10:$B$14))
    ```
    onde:
    - `B1` é a 1ª célula do intervalo dos horários.
    - `$A1` é a 1ª data.
    - `$A$10:$A$14` Coluna de datas da tabela de marcação manual.
    - `$B$10:$B$14` Coluna de horários da tabela de marcação manual.

### `[f6]` `#FEATURE` _{Solicitada pelo Igor Jaloto}_

> Criar um botão que preencha o almoço com os pontos de saída e retorno do almoço (como se tivesse sido feita 1h de almoço apenas)

Esse botão exige macro, mas o próprio Jaloto conseguiu implementá-lo muito bem. A questão aqui é que estamos fazendo um Python para importar os dados direto do SAP, sem precisar ficar preenchendo, o que tornaria o botão quase desnecessário.

Por outro lado, esse botão ajudaria a planejar horas ou algo do tipo, preenchendo dias que ainda serão trabalhados.

Se ele for implementado, os pontos de mudança e de atenção seriam basicamente os da *feature* [`[f3]`](#f3-feature-solicitada-pelo-diogo-costa).


---
---
---

# :package: `[v1.2.1-excel]` <span style="color:#9900FF">(16/03/2025)</span>

## :pencil2: <span style="color:#FFFFFF">Mudanças</span>

Apenas correção do bug [`[b5]`](#b5-bug-encontrado-pelo-diogo-costa).

## :heavy_check_mark: $\textcolor{#66CC00}{\textsf{Bugs corrigidos}}$

### `[b5]` `#BUG` _{Encontrado pelo Diogo Costa}_

Caracterizado a _issue_ [(#1)](https://github.com/LeoneAC/controle-ponto-ETN/issues/1).

## :x: $\textcolor{red}{\textsf{Bugs conhecidos}}$

Nenhum bug conhecido até o momento.

## :sparkles: $\textcolor{#00CCCC}{\textsf{Features implementadas}}$

Nenhuma feature implementada nessa versão.

## :bulb: $\textcolor{#FFCC00}{\textsf{Features pendentes}}$

Apenas as da versão [`[v1.2.0-excel]`](#package-v120-excel-10022025)

---
---
---

# :package: `[v1.2.2-excel]` <span style="color:#9900FF">(27/12/2025)</span>

## :pencil2: <span style="color:#FFFFFF">Mudanças</span>

- Retorno da automação da célula `ANO` da aba `DADOS`:
    ```
    =SE($C$13<>"";ANO($C$13);"(Inserir data na primeira linha do Calendário)")
    ```
- Atualização do Calendário para 2026.
- 
- Adição do código INTER na coluna de Abono, com atualização da dica explicativa.
  - A aba `EXEMPLOS` foi atualizada para refletir essa mudança.
- Atualização da função `EntradaTratada` do VBA para corresponder ao novo comportamento do SAP que não mais arredonda a meia hora inicial para **7:30**
  - Troca de `ENTRADA_ - TimeValue("0:30")` para `ENTRADA_ - TimeValue("0:00")`
- Desativação da macro `AtalhoColarValores` do atalho `Ctrl+Shift+V`, pois as novas versões do Excel 365 implementaram o mesmo comportamento de colar apenas valores no mesmo atalho utilizado aqui. Por isso não há motivo de manter o atalho com a desvantagem de matar o `Ctrl+Z`, isto é, perdendo o histórico de alterações.
  - A aba `EXEMPLOS` foi atualizada para refletir essa mudança.
- Aplicação da correção do bug `[b5]` na aba `EXEMPLOS` também.
- Alteração do comportamento do botão de `Novo Controle de Horas`:
  - Correção do nome sugerido do arquivo no botão `Novo Controle de Horas` tirando `"- Com Macro"` no final.
  - Retirada da possibilidade de formatação das células (quando protegidas) de todas as planilhas para evitar desconfiguração no processo de cópia e cola no uso da planilha, deletando `AllowFormattingCells:=True` de todas as planilhas e tirando todas as possibilidades de edição da aba `BASE`.
  - Agora as planilhas não deletadas e recriadas ao invés de ter apenas os dados apagados, para que as "gambiarras" do usuário sejam desfeitas de um ano para outro e evite bugs difíceis de encontrar. (A ideia é que um lançamento manual de um dia em um ano não necessariamente será necessário no ano seguinte).
- Acrescentada formatação condicional para tirar `00:00` de dias que estão com tipo de dia em branco no calendário: 
    ```
    Fórmula: =E(L2=0;$B2=0)
    Formato: texto em cinza muito claro
    Aplica-se a: =$L$2:$P$32;$T$2:$T$32
    ```

## :heavy_check_mark: $\textcolor{#66CC00}{\textsf{Bugs corrigidos}}$

Nenhum bug corrigido nessa versão.

## :x: $\textcolor{red}{\textsf{Bugs conhecidos}}$

Nenhum bug conhecido até o momento.

## :sparkles: $\textcolor{#00CCCC}{\textsf{Features implementadas}}$

Nenhuma feature implementada nessa versão.

## :bulb: $\textcolor{#FFCC00}{\textsf{Features pendentes}}$

Apenas as da versão [`[v1.2.0-excel]`](#package-v120-excel-10022025)
