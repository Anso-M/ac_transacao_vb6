# Projeto VB6

## Descrição
Este projeto em Visual Basic 6 (VB6) é uma aplicação de gerenciamento de transações bancárias. A aplicação permite visualizar, adicionar, editar, excluir e exportar transações para um arquivo Excel. O projeto inclui funcionalidades de filtro para facilitar a consulta de transações específicas.

## Requisitos
- Microsoft Visual Basic 6.0
- Microsoft Excel (para exportação de dados)
- SQL Server Express

## Instalação
1. Clone este repositório:
   ```sh
   git clone https://github.com/seu-usuario/nome-do-projeto.git
   ```
   É possível também fazer o download dos arquivos.
<br><br/>
2. Abra o projeto no Microsoft Visual Basic 6.0.
<br><br/>
3. Configure as referências necessárias:
   - Adicione a biblioteca `Microsoft Excel 16.0 Object Library`.
   - Adicione a referência à biblioteca ADODB para manipulação do banco de dados.
<br><br/>
4. Rode os scripts SQL presentes nos arquivos do projeto.
<br><br/>
5. Certifique-se de que o banco de dados SQL Server esteja configurado corretamente conforme a string de conexão fornecida no código (`localhost\SQLEXPRESS`).
<br><br/>
6. Certifique-se de que o banco de dados possua no mínimo 1 registro na tabela `clientes`, pois as transações só podem ser cadastradas se houver algum cliente com um número de cartão válido.

## Funcionalidades do Código VB6

### Botões e Suas Funcionalidades

1. **Inserir**
   - **Descrição**: Insere uma nova transação no banco de dados. Preencha os campos na parte superior da interface e clique no botão 'Inserir'.
   - **Validações**:
     - Verifica se o número do cartão tem 16 dígitos e é numérico.
     - Verifica se a data está no formato brasileiro (dd/mm/yyyy).
     - Verifica se o valor da transação é numérico.
     - Substitui a vírgula por ponto decimal no valor da transação.
   - **Ações**:
     - Obtém o `id_cliente` com base no `numero_cartao`.
     - Insere a transação no banco de dados e atualiza o DataGrid.

2. **Editar**
   - **Descrição**: Edita uma transação existente no banco de dados. No DataGrid, clique na linha da transação que deseja editar. Os campos na parte superior da interface serão preenchidos após o clique. Altere os campos que deseja e clique no botão 'Editar'.
   - **Validações**:
     - Verifica se o número do cartão tem 16 dígitos e é numérico.
     - Verifica se a data está no formato brasileiro (dd/mm/yyyy).
     - Verifica se o valor da transação é numérico.
     - Substitui a vírgula por ponto decimal no valor da transação.
   - **Ações**:
     - Obtém o `id_cliente` com base no `numero_cartao`.
     - Atualiza a transação existente no banco de dados e atualiza o DataGrid.

3. **Excluir**
   - **Descrição**: Exclui uma transação selecionada do banco de dados. No DataGrid, clique na linha da transação que deseja excluir. Depois, clique em 'Excluir'. A aplicação vai pedir uma confirmação de exclusão. Depois da confirmação, a transação será deletada.
   - **Validações**:
     - Verifica se um `id_transacao` foi selecionado.
     - Pede confirmação ao usuário antes de excluir.
   - **Ações**:
     - Exclui a transação do banco de dados e atualiza o DataGrid.

4. **Limpar**
   - **Descrição**: Limpa todos os campos de texto da transação.

5. **Consultar**
   - **Descrição**: Consulta transações no banco de dados com base nos filtros fornecidos.
   - **Validações**:
     - Verifica se o número do cartão é numérico.
     - Valida a data no formato brasileiro.
     - Verifica se o valor da transação é numérico.
   - **Filtros**:
     - `numero_cartao`
     - `data_transacao`
     - `valor_transacao`

6. **LimparFiltros**
   - **Descrição**: Limpa todos os campos de filtro e atualiza o DataGrid com todos os dados.

7. **MostrarCategoria**
   - **Descrição**: Atualiza o DataGrid, adicionando uma coluna de categoria da transação.
   - **Validações**:
     - Verifica se a função de categoria está criada no banco de dados.

8. **Exportar para Excel**
   - **Descrição**: Exporta os dados de transações do último mês para um arquivo Excel.
   - **Validações**:
     - Verifica se a função de categoria está criada no banco de dados.
   - **Ações**:
     - Cria um objeto Excel e preenche as células com os dados do banco.
     - Exibe uma caixa de diálogo para o usuário escolher o local de salvamento do arquivo.

### Campos de Texto

- **txtIdTransacao (oculto)**: Armazena o ID da transação.
- **txtNumero**: Armazena o número do cartão.
- **txtValor**: Armazena o valor da transação.
- **txtData**: Armazena a data da transação no formato dd/mm/yyyy.
- **txtDescricao**: Armazena a descrição da transação.

### Campos de Filtro

- **txtFiltroNumeroCartao**: Filtra transações pelo número do cartão.
- **txtFiltroDataTransacao**: Filtra transações pela data.
- **txtFiltroValorTransacao**: Filtra transações pelo valor.

### Validações e Verificações

- **IsValidBrazilianDate**: Função que valida se a data está no formato brasileiro e se é uma data válida.
- **Verificação no DataGrid**: Verifica se o DataGrid contém registros antes de preencher os campos de texto.
- **Formatação de Data**: Formata a data no formato dd/mm/yyyy para exibição e yyyy-mm-dd para operações no banco de dados.

### Funções Adicionais

- **AtualizarGrid**: Atualiza o DataGrid com os dados mais recentes do banco de dados.
- **Error Handlers**: Exibe mensagens de erro para casos específicos de falha ao acessar ou modificar dados no DataGrid.

## Arquivos e Estrutura do Projeto
- `src/Banco.vbp`: Arquivo de projeto VB6.
- `src/Banco.frm`: Arquivo de formulário contendo o design e código do formulário.
- `src/Banco.vbw`: Arquivo de espaço de trabalho VB6.

## Licença
Este projeto está licenciado sob a [MIT License](LICENSE).
