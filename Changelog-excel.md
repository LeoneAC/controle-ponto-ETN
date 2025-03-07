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