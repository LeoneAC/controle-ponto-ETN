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

---
---
---

# :package: `[v1.0.1-python]` <span style="color:#9900FF">(12/02/2025)</span>

Essa versão corrige um bug na leitura do arquivo `.pdf` que deixava a marcação passar quando a última linha da tabela era de informação de marcação.

## :pencil2: <span style="color:#FFFFFF">Mudanças</span>

1. Alteração retroativa do nome das versões para o padrão de versionamento: `MAJOR.MINOR.PATCH`
1. Adição do `-python` na versão para evitar ambiguidade com a versão geral do projeto e a versão do Excel.

## :heavy_check_mark: $\textcolor{#66CC00}{\textsf{Bugs corrigidos}}$

### `[b1]` `#BUG` _{Encontrado pelo Igor Jaloto}_

Quando a última linha da tabela inteira do `.pdf` continha uma informação de **Marcação** ao invés de qualquer outra, o código deixava essa linha passar e não puxava corretamente a informação. 

## :x: $\textcolor{red}{\textsf{Bugs conhecidos}}$

*Nenhum*

## :sparkles: $\textcolor{#00CCCC}{\textsf{Features implementadas}}$

*Nenhuma*

## :bulb: $\textcolor{#FFCC00}{\textsf{Features pendentes}}$

*Nenhuma*