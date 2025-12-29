## 📦 Baixe a última versão da planilha [clicando aqui](https://github.com/LeoneAC/controle-ponto-ETN/releases/latest/download/Controle-de-Horas.zip)

# Sobre o projeto (controle-ponto-ETN)

Esse projeto traz ferramentas para facilitar o Controle de Ponto na Eletronuclear S.A. (ETN).

De modo geral, é composto por:
- um <ins>**Excel**</ins> (`Controle de Horas YYYY.xlsm`) que ajuda a visualizar as marcações de ponto do funcionário e simula os tratamentos de ponto que o SAP (sistema que a empresa usa) faz, facilitando o tratamento e controle das horas;
- um <ins>**Executável**</ins> (`import_from_SAP.exe`) que importa para o Excel as informações de um arquivo `.pdf` gerado pelo SAP, facilitando o preenchimento dos horários.

_Observação:_ O arquivo <ins>Excel</ins> pode ser usado com preenchimento manual, sem o Executável.

# Sobre o Excel

O arquivo Excel foi pensado para ser autoexplicativo, contando com uma área de _Informações_ na aba **DADOS** e exemplos de utilização na aba **EXEMPLO**.\
Tenho intenção de fazer um Manual do Usuário em LaTeX, mas isso toma muito tempo e não sei se vou conseguir.

Dúvidas, sugestões e bugs podem ser postados na área de [Issues](https://github.com/LeoneAC/controle-ponto-ETN/issues), farei o possível para analisá-los e atendê-los, mas infelizmente não consigo dar prazos.

## Problemas para habilitar macros no Excel

Se a mensagem abaixo aparecer para você, por favor, clique no botão [`Saiba mais`](https://support.microsoft.com/pt-br/topic/uma-macro-potencialmente-perigosa-foi-bloqueada-0952faa0-37e7-4316-b61d-5b5ed6024216) e siga as instruções para desbloqueio do arquivo Excel.

> RISCO DE SEGURANÇA: A Microsoft bloqueou a execução de macros porque a origem deste arquivo não é confiável.

![segurança](https://github.com/user-attachments/assets/413803d2-cbd7-4705-a275-1941eb10a536)

# Sobre o Executável

O executável foi desenvolvido em Python e serve basicamente para facilitar a vida de quem usa a planilha em Excel, isto é, não é um recurso obrigatório para o funcionamento da planilha.

## Como utilizar o Executável

Para utilizar o `import_from_SAP.exe`, baixe do SAP o Comprovante de tempos em `.pdf` seguindo os seguintes passos:
1. Autoatendimento do Empregado;
1. Consultas;
1. \[Frequência\] Comprovante de Tempos;
1. Escolher o mês e ano;
1. Executar;
1. Baixe o .pdf em ![botão de download](https://github.com/user-attachments/assets/4fd38274-cce0-4436-bfab-1bf5c68f2b38).



Coloque numa mesma pasta o `.exe`, o `.pdf` e o `.xlsm` (como mostra a imagem abaixo) e execute o `import_from_SAP.exe` (duplo clique).  
A pasta pode conter outros arquivos, como os comprovantes de tempo dos outros meses, por exemplo.
> Organização da pasta
>```
>Controle de Horas/
>├── Controle de Horas 2026.xlsm
>├── import_from_SAP.exe
>├── smart.pdf
>└── (outros arquivos)
>```
![Organização_de pasta](https://github.com/user-attachments/assets/88f5127c-df76-42e2-bf15-940049ee1888)


## ❗ Observações importantes

- O arquivo é grande porque eu compacto tudo num mesmo `.exe` para facilitar o envio.
- É normal ele demorar a executar após você abrir o `.exe` (o computador precisa descompactar algumas informações do arquivo), então ele ficará piscando o cursor com a tela preta vazia, apenas aguarde ([#6](../../issues/6)).
- É normal ele demorar a fechar ao clicar `Enter` no fim do programa, você pode simplesmente esperar ou fechar a janela no ❌.
- O programa sobrescreve os dados importados do SAP para a sua planilha, ignorando se havia algo previamente escrito na célula ou não.
- Você pode usar o programa com a planilha aberta ou fechada, tanto faz, mas é imprescindível que ela esteja **salva com edição e macro habilitadas**.
  - Se é a primeira vez que você está abrindo a planilha ([#4](../../issues/4)), ela certamente ficará bloqueada para edição e com as macros desabilitadas. Nesse caso você precisa:
    1. abrir;
    1. permitir edição e macros;
    1. salvar;
    1. fechar o arquivo.
    
    Só então o programa que importa do SAP funcionará sem problemas.
> [!NOTE]
> Se você estiver recebendo erros ao usar o programa, veja [#7](../../issues/7).
- A versão `v1.0.3-python` do `.exe` foi testada extensivamente usando Excel na versão **365** e todas as fragilidades encontradas foram corrigidas com exceção de:
  - o problema supracitado de permissão de edição e macros;
  - quando há marcações bloqueadas ou a serem bloqueadas (5 ou mais marcações num mesmo dia). Nesse caso o programa deve falhar.

---
# Licença (License)

<p xmlns:cc="http://creativecommons.org/ns#" xmlns:dct="http://purl.org/dc/terms/"><a property="dct:title" rel="cc:attributionURL" href="https://github.com/LeoneAC/controle-ponto-ETN.git">controle-ponto-ETN</a> by <a rel="cc:attributionURL dct:creator" property="cc:attributionName" href="https://github.com/LeoneAC">Leone Andrade Campos</a> and Carlos Leonardo da Silva Xavier is licensed under <a href="https://creativecommons.org/licenses/by-nc-sa/4.0/?ref=chooser-v1" target="_blank" rel="license noopener noreferrer" style="display:inline-block;">Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International

<img style="height:22px!important;margin-left:3px;vertical-align:text-bottom;" src="https://mirrors.creativecommons.org/presskit/icons/cc.svg?ref=chooser-v1" alt=""><img style="height:22px!important;margin-left:3px;vertical-align:text-bottom;" src="https://mirrors.creativecommons.org/presskit/icons/by.svg?ref=chooser-v1" alt=""><img style="height:22px!important;margin-left:3px;vertical-align:text-bottom;" src="https://mirrors.creativecommons.org/presskit/icons/nc.svg?ref=chooser-v1" alt=""><img style="height:22px!important;margin-left:3px;vertical-align:text-bottom;" src="https://mirrors.creativecommons.org/presskit/icons/sa.svg?ref=chooser-v1" alt=""></a>

CC BY-NC-SA 4.0 - https://creativecommons.org/licenses/by-nc-sa/4.0/</p>
