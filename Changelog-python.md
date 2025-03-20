### :hash: <span style="color:#007AFF">Tags para pesquisa: `#BUG` `#FEATURE`</span>

---
---
---

# :package: `[v1.0.0-python]` <span style="color:#9900FF">(22/01/2025)</span>

## Versão inicial
O *script* em Python basicamente pega as informações de marcações de ponto do arquivo `.pdf` gerado pelo SAP (Sistema de ponto da ETN), arquivo este normalmente chamado de `smart.pdf`, e as insere no arquivo Excel de controle de ponto, arquivo este normalmente chamado de `Controle de Horas YYYY.xlsm`.

O código é capaz de:
- encontrar o mês referente às marcações no próprio `.pdf`;
- inserir as informações no Excel quando este está fechado ou aberto (em uso);
- aceitar nomes de arquivos `.pdf` e `.xlsm` diferentes dos padrões sugeridos.

> [!TIP]
> As instruções de uso do executável `import_from_SAP.exe` serão colocadas no [README.md](https://github.com/LeoneAC/controle-ponto-ETN/blob/main/README.md).

> [!NOTE]
> Para mais informações de como *script* funciona, leia os comentários no próprio *script*.

## :x: $\textcolor{red}{\textsf{Bugs conhecidos}}$

A versão `[v1.0.0-python]` foi testada várias vezes usando Excel nas versões **2016** e **365** e todos as fragilidades encontradas foram corrigidas com exceção de:

### `[b1]` `#BUG` _{Encontrado pelo Leone}_

O programa não funciona quando o arquivo Excel está bloqueado para edição e com as macros desabilitadas.

Isso pode ocorrer principalmente quando é a primeira vez que você está abrindo a planilha, no entanto o Windows interpreta um arquivo já usado como novo quando você faz algumas ações como copiar para outro lugar, por exemplo.

### `[b2]` `#BUG` _{Encontrado pelo Leone}_

Quando há marcações bloqueadas ou a serem bloqueadas (5 ou mais marcações num mesmo dia), o programa deve falhar.

Não foi feita nenhuma correção em código para detectar esse tipo de caso.

---
---
---

# :package: `[v1.0.1-python]` <span style="color:#9900FF">(12/02/2025)</span>

Essa versão corrige um bug na leitura do arquivo `.pdf` que deixava a marcação passar quando a última linha da tabela era de informação de marcação.

## :pencil2: <span style="color:#FFFFFF">Mudanças</span>

1. Alteração retroativa do nome das versões para o padrão de versionamento: `MAJOR.MINOR.PATCH`
1. Adição do `-python` na versão para evitar ambiguidade com a versão geral do projeto e a versão do Excel.

## :heavy_check_mark: $\textcolor{#66CC00}{\textsf{Bugs corrigidos}}$

### `[b3]` `#BUG` _{Encontrado pelo Igor Jaloto}_

Quando a última linha da tabela inteira do `.pdf` continha uma informação de **Marcação** ao invés de qualquer outra, o código deixava essa linha passar e não puxava corretamente a informação. 

## :x: $\textcolor{red}{\textsf{Bugs conhecidos}}$

*Os mesmos da versão [`[v1.0.0-python]`](#package-v100-python-22012025)*.

## :sparkles: $\textcolor{#00CCCC}{\textsf{Features implementadas}}$

*Nenhuma*

## :bulb: $\textcolor{#FFCC00}{\textsf{Features pendentes}}$

*Nenhuma*

---
---
---

# :package: `[v1.0.2-python]` <span style="color:#9900FF">(19/03/2025)</span>

## :pencil2: <span style="color:#FFFFFF">Mudanças</span>

Apenas correção do bug [`[b2]`](#b2-bug-encontrado-pelo-leone).

## :heavy_check_mark: $\textcolor{#66CC00}{\textsf{Bugs corrigidos}}$

### `[b2]` `#BUG` _{Encontrado pelo Leone}_

Veja [`[b2]`](#b2-bug-encontrado-pelo-leone), caracterizado na _issue_ [(#5)](https://github.com/LeoneAC/controle-ponto-ETN/issues/5).

## :x: $\textcolor{red}{\textsf{Bugs conhecidos}}$

*Somente o bug [`[b1]`](#b1-bug-encontrado-pelo-leone)*.

## :sparkles: $\textcolor{#00CCCC}{\textsf{Features implementadas}}$

*Nenhuma*

## :bulb: $\textcolor{#FFCC00}{\textsf{Features pendentes}}$

*Nenhuma*