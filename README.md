# Emissão de DARF — guia de uso

O programa lê **código da receita, valor principal, período de apuração e CNPJ** de quatro células de uma planilha .xlsx. Você conclui a identificação no [Sicalc da Receita Federal](https://sicalc.receita.fazenda.gov.br/sicalc/principal) pelo Google Chrome. Depois, a interface preenche o cálculo e salva **um PDF por execução** na pasta escolhida.

**Antes de pagar:** confira no PDF o CNPJ, o código, o período, o valor, o vencimento e os acréscimos. Gerar o PDF não efetua o pagamento.

## 1. Preparar o computador

Você precisa de:

- Windows, Google Chrome e Python 3 para Windows.
- Internet para acessar o Sicalc e instalar as bibliotecas na primeira abertura.
- A pasta do programa em um local onde você possa criar arquivos, como Documentos.

Mantenha juntos os arquivos Iniciar DARF.cmd, iniciar.pyw, gui.pyw, darf.py, interface_windows.py e requirements.txt. Baixe o [Python para Windows](https://www.python.org/downloads/windows/) se ele ainda não estiver instalado.

Para abrir o programa, dê **dois cliques em Iniciar DARF.cmd** no Explorador de Arquivos. Na primeira vez, uma janela preparará as bibliotecas automaticamente; aguarde até aparecer **Emissão de DARF**. Você não precisa digitar comandos no terminal. Nas próximas aberturas, a interface aparece diretamente.

## 2. Fazer um teste de leitura com a planilha incluída

Use a [planilha de teste](outputs/01a122b6-3b56-7401-abf8-273ad1432478/planilha_teste_darf.xlsx). Ela contém quatro abas: **Dados DARF**, **Lançamentos**, **Despesas** e **Histórico**. As três últimas têm registros fictícios para você testar a escolha de abas. A primeira contém os valores informados anteriormente, inclusive o CNPJ; ela é a única aba que deve fornecer os dados do DARF neste teste.

Na interface:

1. Em **Planilha .xlsx**, clique em **Procurar...** e escolha a planilha de teste. O programa carregará os nomes das quatro abas.
2. Indique as referências abaixo. Em cada linha, escolha **Dados DARF** no seletor de aba e digite somente a célula na caixa à direita.

| Linha da interface | Aba | Célula | Valor esperado |
| --- | --- | --- | --- |
| Código do imposto | Dados DARF | B4 | 3467-02 |
| Valor principal | Dados DARF | B5 | R$ 20,00 |
| Período de apuração | Dados DARF | B6 | 03/09/2026 |
| CNPJ | Dados DARF | B7 | CNPJ informado anteriormente, terminado em 0109 |

3. Clique em **Conferir dados**. A parte inferior da janela deverá mostrar código 3467-02, valor R$ 20,00, período 03/09/2026 e o CNPJ normalizado.
4. Se os botões da parte inferior não estiverem visíveis, use a barra de rolagem à direita ou a roda do mouse.

**Conferir dados apenas lê e valida a planilha.** Esse teste não abre o Sicalc e não emite um DARF. O botão de emissão só fica habilitado depois da conferência.

## 3. Usar sua própria planilha

O programa aceita arquivos locais **.xlsx** criados no Excel ou baixados do Google Sheets. As quatro informações podem estar na mesma aba ou em abas diferentes. Você escolhe uma aba e uma célula para cada campo; por exemplo, a terceira aba pode ter o código em A8 e a segunda aba pode ter o valor em A9.

| Informação | Como preencher a célula |
| --- | --- |
| Código da receita | Seis dígitos, com ou sem hífen: 3467-02 ou 346702. |
| Valor principal | Número positivo com até duas casas decimais. O número 20 significa R$ 20,00. |
| Período de apuração | 03092026 ou 03/09/2026. Para esse código, 03 indica o terceiro decêndio de setembro de 2026. |
| CNPJ | CNPJ válido de 14 dígitos, com ou sem pontuação. |

Para preservar zeros iniciais, formate **código, período e CNPJ como Texto** antes de digitar. Prefira valores digitados diretamente nas quatro células. Fórmulas sem resultado salvo no arquivo .xlsx aparecem vazias para o programa. Salve a planilha antes de clicar em **Conferir dados**; se alterar o arquivo, confira os dados novamente.

### Se você não tem Excel: use o Google Sheets

O Google Sheets [permite abrir e editar arquivos .xlsx](https://support.google.com/docs/answer/9331167?hl=pt-BR). Para usar a planilha de teste ou a sua:

1. Envie o arquivo .xlsx ao Google Drive e abra-o com o Google Sheets.
2. Edite as células necessárias. Preserve os campos de código, período e CNPJ como texto.
3. No Sheets, use **Arquivo > Fazer download > Microsoft Excel (.xlsx)**.
4. No programa, clique em **Procurar...** e selecione essa cópia baixada, normalmente na pasta Downloads.
5. Depois de cada alteração no Sheets, baixe uma **nova cópia .xlsx** e selecione-a novamente no programa. Clique em **Conferir dados** para verificar a versão escolhida.

A interface lê o arquivo **salvo no computador**. Um link do Google Sheets não substitui o caminho do arquivo, e alterações feitas na nuvem não atualizam automaticamente a cópia local.

## 4. Escolher onde o PDF será salvo

Em **Onde salvar o PDF?**, a primeira pasta é o destino final. Por padrão, ela é a pasta do programa. Clique em **Escolher pasta...** para alterá-la. Escolha uma pasta que já exista.

A segunda pasta é a de **downloads do Chrome**. Ela começa como a pasta Downloads do Windows. Altere-a somente se o Chrome baixar arquivos em outro lugar. O programa procura o PDF recebido nessa pasta e salva uma cópia com nome próprio no destino final.

O arquivo final tem nome semelhante a darf_3467-02_03092026_20261013_145333_762788.pdf. O Chrome pode manter outra cópia na pasta Downloads. Use o caminho mostrado pela interface ao concluir.

## 5. Emitir um DARF

Faça primeiro as etapas das seções 2 ou 3 e confira os dados exibidos. Depois:

1. Clique em **Abrir Sicalc no Chrome**. O programa abre o site no Chrome comum.
2. No site, clique em **Preenchimento rápido** e escolha **Pessoa Jurídica**.
3. Informe **o mesmo CNPJ mostrado na interface**. O botão **Copiar CNPJ** fica disponível após **Conferir dados**.
4. Conclua **“Sou humano”** e clique em **Continuar** no site.
5. Aguarde aparecer a página com o campo **código ou nome da receita**. Se o site ainda estiver na página **Contribuinte** ou mostrar um erro, resolva isso no Chrome antes de voltar ao programa.
6. Volte à interface e clique em **Já estou na tela do código — emitir DARF**. Não é preciso pressionar Enter nem clicar novamente no Chrome.
7. A interface se esconde, localiza a janela do Sicalc, ativa a aba e preenche código, período e valor. Ela calcula, seleciona o resultado e solicita o PDF. **Aguarde sem usar o mouse ou o teclado** nessa etapa.
8. Quando a interface reaparecer, confira a mensagem **PDF salvo em** e clique em **Abrir PDF salvo**. Revise o documento antes de qualquer pagamento.

O CAPTCHA e o CNPJ são preenchidos por você no Chrome comum. A automação começa depois que **Continuar** abre a tela do código. O Sicalc recusou o CAPTCHA em uma sessão de Chrome automatizada nos testes anteriores.

## 6. Resolver problemas comuns

| Situação | Verificação |
| --- | --- |
| A janela do programa não abre | Instale Python 3 para Windows. Se a preparação automática falhar, veja instalacao.log na pasta do programa. |
| O caminho da planilha é um link do Sheets | Baixe o arquivo em **Microsoft Excel (.xlsx)** e selecione a cópia local. |
| Uma aba ou célula não aparece | Confira se escolheu o arquivo certo. Se digitou o caminho manualmente, clique em **Carregar abas**. |
| Uma das quatro células está vazia ou inválida | Confira a aba, o endereço da célula e o valor salvo no .xlsx. Para fórmulas, prefira substituir pelo valor final digitado. |
| O botão de emitir está desabilitado | Preencha as quatro referências, escolha pastas existentes e clique em **Conferir dados**. |
| O Sicalc mostra “Contribuinte inválido” | Confira o CNPJ e conclua a identificação no Chrome. O programa só prossegue na tela do código. |
| O programa não encontra o Chrome ou o campo do código | Deixe uma única janela do Chrome com o Sicalc na tela do código e tente novamente. |
| O PDF não aparece | Confira onde o Chrome fez o download e ajuste a segunda pasta na interface. Veja também se o site abriu o PDF em outra aba. |

Se a emissão falhar, a interface reaparece com uma mensagem. Confira o estado atual do Sicalc antes de tentar novamente, pois uma segunda tentativa pode gerar outro documento.
