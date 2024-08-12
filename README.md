# Tracker Cotação Dólar 

Este projeto é uma ferramenta para obter a cotação do dólar comercial a partir de um site específico e gerar um documento do Word com a informação coletada, que é então convertido para PDF. A aplicação utiliza o Selenium para automatizar a navegação na web, o Tkinter para interação com o usuário e o `python-docx` para manipulação de documentos do Word.

## Funcionalidades

- **Obtenção da cotação do dólar**: Navega para uma página web específica e coleta a cotação do dólar.
- **Captura de tela**: Realiza uma captura de tela da cotação do dólar.
- **Geração de documento**: Cria um documento do Word com a cotação do dólar e a captura de tela.
- **Conversão para PDF**: Converte o documento do Word para PDF.

## Como Usar

Você tem duas opções para usar este projeto:

### 1. Baixar e Executar o Executável

Se você preferir, pode baixar o arquivo executável [`exchangescraper.exe`](https://drive.google.com/drive/folders/19rOJTGfA2Tct5nLoAdHr-p9eVAwcBibs?usp=drive_link) e simplesmente executá-lo no seu computador. Com essa opção, **não há necessidade de instalar dependências** ou configurar o ambiente Python. Basta seguir os passos abaixo:

1. **Baixe o arquivo executável**:

    - [Clique aqui para baixar o `exchangescraper.exe`](https://drive.google.com/drive/folders/19rOJTGfA2Tct5nLoAdHr-p9eVAwcBibs?usp=drive_link)

2. **Execute o arquivo**:

    - No Windows, basta dar um duplo clique em `exchangescraper.exe` e seguir as instruções na interface gráfica.

### 2. Executar o Código Fonte

Se você preferir, pode clonar o repositório e executar o código Python diretamente. Esta opção é ideal para desenvolvedores que desejam entender ou modificar o código.

1. **Clone o repositório:**

    ```bash
    git clone https://github.com/SEU_USUARIO/NOME_DO_REPOSITORIO.git
    cd NOME_DO_REPOSITORIO
    ```

2. **Instale as dependências:**

    É recomendado usar um ambiente virtual. Crie e ative um ambiente virtual, depois instale as dependências:

    ```bash
    python -m venv venv
    source venv/bin/activate  # No Windows use `venv\Scripts\activate`
    pip install -r requirements.txt
    ```

3. **Baixe o ChromeDriver:**

    Faça o download do ChromeDriver compatível com a versão do seu Chrome [aqui](https://sites.google.com/chromium.org/driver/). Coloque o executável do ChromeDriver em um diretório que está no PATH ou especifique o caminho diretamente no código.

4. **Execute o script:**

    ```bash
    python main.py
    ```

5. **Siga as instruções na interface gráfica** para obter a cotação do dólar e gerar o documento.

## Estrutura do Código

- `CustomDialog`: Classe para criar uma janela de entrada para o nome do usuário.
- `ask_name()`: Função para exibir o `CustomDialog` e obter o nome do usuário.
- `start_driver()`: Configura e inicia o navegador Chrome com o Selenium.
- `get_exchange_rate()`: Navega para o site e coleta a cotação do dólar.
- `convert_docx_to_pdf()`: Converte o documento do Word para PDF.
- `create_exchange_rate_document()`: Cria um documento do Word com a cotação do dólar e o converte para PDF.
- `execute_tasks()`: Função principal que coordena a execução das tarefas.
- `main()`: Função principal que inicia o processo.


## Licença

Este projeto está licenciado sob a Licença MIT - veja o arquivo [LICENSE](LICENSE) para mais detalhes.

## Contato

Se você tiver alguma dúvida ou sugestão, sinta-se à vontade para me contatar.
