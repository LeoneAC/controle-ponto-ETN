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
