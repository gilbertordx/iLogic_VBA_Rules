# Documentação das Regras iLogic

Esta documentação compila as informações e o código para as principais regras iLogic utilizadas na automação do Autodesk Inventor.

---

# Padrão de Boas Práticas para Scripts iLogic

- **Modularidade:** Separe funções auxiliares (validação, montagem de descrição, logging, etc.)
- **Validação:** Sempre valide parâmetros antes de gerar descrições ou atualizar iProperties
- **Atualização do Part Number:** Sempre que gerar uma nova descrição, atualize também o Part Number
- **Tratamento de Erros:** Use tratamento de exceções e logging para rastreabilidade
- **Clareza e Consistência:** Use nomes de variáveis claros e padronizados

## Exemplo de Estrutura Recomendada
```vbnet
Sub Main()
    ' Coleta e valida parâmetros
    Dim param1 As Double = PARAM1
    Dim param2 As String = PARAM2
    If Not ValidarParametros(param1, param2) Then
        MessageBox.Show("Parâmetros inválidos!", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Exit Sub
    End If
    ' Monta descrição
    Dim desc As String = MontarDescricao(param1, param2)
    iProperties.Value("Project", "Description") = desc
    iProperties.Value("Project", "Part Number") = desc
End Sub

Private Function ValidarParametros(p1 As Double, p2 As String) As Boolean
    If p1 <= 0 Then Return False
    If String.IsNullOrWhiteSpace(p2) Then Return False
    Return True
End Function

Private Function MontarDescricao(p1 As Double, p2 As String) As String
    Return $"DESCRIÇÃO PADRÃO {p1} | {p2}"
End Function
```

---

# Como Usar
1. Abra o arquivo de peça (.ipt) ou montagem (.iam) no Autodesk Inventor.
2. Vá para a aba **"Manage"** e clique em **"iLogic"** -> **"Add Rule"**.
3. Crie uma nova regra com o nome desejado.
4. Copie o código padronizado e cole na janela de edição da regra.
5. Salve a regra.
6. A regra pode ser executada manualmente, por evento (ex: "Before Save", "After Parameter Change") ou chamada por outra regra iLogic.

---

# Observações
- Sempre alinhe o nome dos parâmetros do modelo com o nome usado no script.
- Use logging para rastrear erros e facilitar manutenção.
- Atualize a documentação sempre que um novo padrão for adotado.

---

# Regra iLogic: _Gerar_Descricao_Tubo_

## Propósito
Esta regra iLogic foi desenvolvida para automatizar a geração da propriedade **"Description"** (Descrição) de peças do tipo _"TUBO"_ no Autodesk Inventor. Ela utiliza propriedades personalizadas e informações do material para criar uma descrição padronizada e informativa.

## Descrição
A regra coleta a _altura_, _largura_, _espessura_ e _comprimento_ de um tubo, além do _material_ atribuído à peça. Com essas informações, ela constrói uma *string de descrição formatada* que inclui as dimensões, o material (em maiúsculas) e o comprimento total do tubo, e então atribui essa string à propriedade **"Description"** (Categoria 'Project') da peça.

## Parâmetros/iProperties Necessários
Para o funcionamento correto desta regra, os seguintes parâmetros/iProperties devem estar definidos na peça:

*   `ALTURA`: Dimensão da altura do tubo.
*   `LARGURA`: Dimensão da largura do tubo.
*   `ESPESSURA`: Dimensão da espessura da parede do tubo.
*   `COMPRIMENTO`: Comprimento total do tubo.
*   `iProperties.Material()`: O material atribuído à peça.

## Formato de Saída (Exemplo)
"TUBO 30 × 30 × 1,2 | AÇO, CARBONO | COMP. 1740 mm"

## Código VBA
```vba
Dim materialNome As String = UCase(iProperties.Material())
Dim dimensoesTexto As String = ALTURA & " × " & LARGURA & " × " & ESPESSURA
Dim comprimentoTexto As String = COMPRIMENTO & " mm"

Dim descricaoFinal As String
descricaoFinal = "TUBO " & dimensoesTexto & " | " & materialNome & " | COMP. " & comprimentoTexto

iProperties.Value("Project", "Description") = descricaoFinal
```

## Como Usar
1.  Abra o *arquivo de peça (.ipt)* ou *montagem (.iam)* no Autodesk Inventor.
2.  Vá para a aba **"Manage"** e clique em **"iLogic"** -> **"Add Rule"**.
3.  Crie uma nova regra com o nome `Gerar_Descricao_Tubo`.
4.  Copie o **código VBA** fornecido acima e cole-o na janela de edição da regra.
5.  Certifique-se de que os parâmetros `ALTURA`, `LARGURA`, `ESPESSURA` e `COMPRIMENTO` estejam definidos na peça e que um material esteja atribuído.
6.  Salve a regra.
7.  A regra pode ser executada *manualmente*, por um *evento* (ex: "Before Save", "After Parameter Change") ou *chamada por outra regra iLogic*.

---

# Regra iLogic: _Gerar_Descricao_Trefilado_

## Propósito
Esta regra iLogic foi desenvolvida para automatizar a geração da propriedade **"Description"** (Descrição) de peças do tipo _"TREFILADO"_ no Autodesk Inventor. Ela converte um parâmetro de bitola de milímetros para polegadas e utiliza o comprimento para criar uma descrição padronizada.

## Descrição
A regra coleta os parâmetros `BITOLA` e `COMPRIMENTO` da peça. A `BITOLA` é acessada utilizando `ThisApplication.ActiveDocument.ComponentDefinition.Parameters.Item("BITOLA").ValueAsString`, que retorna a *representação de string formatada* do parâmetro (ex: "3/8""), e o símbolo de polegadas (") é adicionado. Em seguida, uma *string de descrição* é construída com o prefixo "TREFILADO", a bitola e o comprimento total, e então atribuída à propriedade **"Description"** (Categoria 'Project') da peça.

## Parâmetros/iProperties Necessários
Para o funcionamento correto desta regra, os seguintes parâmetros/iProperties devem estar definidos na peça:

*   `BITOLA`: Dimensão da bitola do trefilado (esperado que seu `ValueAsString` retorne a string nominal, ex: "3/8").
*   `COMPRIMENTO`: Comprimento total do trefilado em milímetros.

## Formato de Saída (Exemplo)
"TREFILADO 3/8" | COMP. 1740 mm"

## Código VBA
```vba
Dim bitolaTexto As String = ThisApplication.ActiveDocument.ComponentDefinition.Parameters.Item("BITOLA").ValueAsString & Chr(34) ' Obtém a representação de string nominal do parâmetro BITOLA

Dim comprimentoTexto As String = COMPRIMENTO & " mm"

Dim descricaoFinal As String
descricaoFinal = "TREFILADO " & bitolaTexto & " | COMP. " & comprimentoTexto

iProperties.Value("Project", "Description") = descricaoFinal
```

## Como Usar
1.  Abra o *arquivo de peça (.ipt)* no Autodesk Inventor.
2.  Vá para a aba **"Manage"** e clique em **"iLogic"** -> **"Add Rule"**.
3.  Crie uma nova regra com o nome `Gerar_Descricao_Trefilado`.
4.  Copie o **código VBA** fornecido acima e cole-o na janela de edição da regra.
5.  Certifique-se de que os parâmetros `BITOLA` (com a representação nominal correta) e `COMPRIMENTO` estejam definidos na peça.
6.  Salve a regra.
7.  A regra pode ser executada *manualmente*, por um *evento* (ex: "Before Save", "After Parameter Change") ou *chamada por outra regra iLogic*.

---

# Regra iLogic: _Preencher_iProperties_Projeto_

## Propósito
Esta regra iLogic tem como objetivo automatizar o preenchimento de várias propriedades de resumo (**iProperties** na aba _'Summary'_) de uma *peça* ou *montagem* no Autodesk Inventor. Ela garante que campos essenciais como **Título**, **Cliente**, **Categoria** e **Palavras-chave** sejam preenchidos de forma consistente.

## Descrição
A regra verifica se o campo **"Title"** (Título) na aba _'Summary'_ das iProperties está vazio. Se estiver, ela interage com o usuário através de *caixas de diálogo* (`InputBox`) para coletar informações como o título do projeto, o nome do cliente, uma categoria e palavras-chave adicionais. Além disso, ela automaticamente preenche o _Autor_ com o nome de usuário do sistema (`System.Environment.UserName`) e define um _Gerente fixo_ (`GILBERTO`).

A regra inclui tratamento para o *cancelamento das entradas mais críticas* (Título e Cliente), impedindo que a regra continue se essas informações não forem fornecidas.

## iProperties Preenchidas
Esta regra preenche as seguintes iProperties (na categoria 'Summary'):

*   **Title (Título):** Definido pelo usuário.
*   **Author (Autor):** Preenchido automaticamente com o nome de usuário do sistema (`System.Environment.UserName`).
*   **Manager (Gerente):** Definido como `GILBERTO`.
*   **Company (Empresa):** Definido pelo usuário (Cliente).
*   **Category (Categoria):** Definido pelo usuário.
*   **Keywords (Palavras-chave):** Definido pelo usuário.

## Interação com o Usuário
A regra solicitará as seguintes informações via `InputBox`:

*   **Título do projeto:** "Qual o título do projeto?:"
*   **Cliente:** "Qual o cliente?:"
*   **Categoria:** "Defina uma categoria para o projeto:"
*   **Palavras-chave:** "Adicione palavras-chave (separadas por vírgula):"

## Tratamento de Erros / Saída Antecipada
*   Se o usuário cancelar ou deixar o campo "**Título do projeto**" vazio, a regra será encerrada (`Exit Sub`).
*   Se o usuário cancelar ou deixar o campo "**Cliente**" vazio, a regra será encerrada (`Exit Sub`).

## Código VBA
```vba
If iProperties.Value("Summary", "Title") = "" Then

    ' -- Variáveis para armazenar as entradas do usuário --
    Dim projetoTitulo  As String
    Dim empresaCliente As String
    Dim projetoCategoria As String
    Dim projetoKeywords As String

    ' -- Perguntas ao usuário e validação de entrada --
    projetoTitulo = InputBox("Qual o título do projeto?:", "Definir Título do Projeto")
    If projetoTitulo = "" Then Exit Sub ' Sai da regra se o título for cancelado/vazio

    empresaCliente = InputBox("Qual o cliente?:", "Definir Cliente")
    If empresaCliente = "" Then Exit Sub ' Sai da regra se o cliente for cancelado/vazio

    projetoCategoria = InputBox("Defina uma categoria para o projeto:", "Definir Categoria")
    projetoKeywords = InputBox("Adicione palavras-chave (separadas por vírgula):", "Definir Palavras-Chave")

    ' -- Constantes para valores fixos --
    Const NOME_GERENTE As String = "GILBERTO"

    ' -- Grava iProperties na aba Resumo --
    iProperties.Value("Summary", "Title")     = projetoTitulo
    iProperties.Value("Summary", "Author")    = System.Environment.UserName ' Nome de usuário do sistema
    iProperties.Value("Summary", "Manager")   = NOME_GERENTE
    iProperties.Value("Summary", "Company")   = empresaCliente
    iProperties.Value("Summary", "Category")  = projetoCategoria
    iProperties.Value("Summary", "Keywords")  = projetoKeywords

End If
```

## Como Usar
1.  Abra o *arquivo de peça (.ipt)* ou *montagem (.iam)* no Autodesk Inventor.
2.  Vá para a aba **"Manage"** e clique em **"iLogic"** -> **"Add Rule"**.
3.  Crie uma nova regra com o nome `Preencher_iProperties_Projeto`.
4.  Copie o **código VBA** fornecido acima e cole-o na janela de edição da regra.
5.  Salve a regra.
6.  A regra pode ser executada *manualmente* ou configurada para ser executada *automaticamente* através de eventos (ex: "After Open Document", "Before Save Document"). Recomenda-se executá-la após a criação de um novo documento. 