
# Automação de Relatórios e Envio de Emails com Selenium, Excel e Outlook

Este projeto realiza a automação de processos de geração de relatórios em uma plataforma web, faz o download desses relatórios e, em seguida, envia emails personalizados com anexos utilizando o Outlook. A automação é feita com Selenium para a navegação e manipulação da página web, e OpenPyXL para a manipulação dos arquivos Excel.

## Funcionalidades

1. **Login Automático**: Utilizando o Selenium, o script realiza o login na plataforma web, preenchendo os campos de login e senha.
2. **Extração de Relatórios**: Através da navegação automática, o relatório desejado é selecionado e baixado no formato Excel.
3. **Processamento de Planilhas**: Após o download, o arquivo Excel é movido para um diretório específico e processado, criando novas planilhas para cada instrutor, com informações organizadas.
4. **Envio de Emails**: Utilizando a integração com o Outlook, o script envia emails automáticos com os relatórios anexados para destinatários especificados em uma planilha.

## Requisitos

- Python 3.x
- Google Chrome
- [ChromeDriver](https://sites.google.com/a/chromium.org/chromedriver/downloads) compatível com a versão do Chrome instalada
- Pacotes Python:
  - `selenium`
  - `openpyxl`
  - `pandas`
  - `win32com.client`
  - `re`
  - `shutil`

Instale os pacotes necessários com o seguinte comando:

```bash
pip install selenium openpyxl pandas pywin32
```

## Como Usar

### Configuração

1. **ChromeDriver**: Baixe o ChromeDriver e coloque o caminho do executável no sistema ou no código.
2. **Planilha de Emails**: Crie uma planilha Excel com uma aba chamada "Dados", contendo colunas como Nome, Nome Completo e Email.
3. **Configuração do Outlook**: O Outlook deve estar configurado no computador que executa o script.

### Execução

1. **Navegação e Download de Relatórios**: O script realiza login em uma plataforma web, seleciona um relatório e faz o download para a pasta de downloads padrão. O campo de data é preenchido automaticamente com a data atual e uma data de fim especificada.

```python
navegador = webdriver.Chrome()
navegador.get("URL_DA_PLATAFORMA")
# Preencher campos de login, senha e instituição
# Clicar no botão de login
# Navegar até o relatório desejado
```

2. **Processamento de Arquivos Excel**: O arquivo baixado é movido para uma pasta específica, e as informações contidas nele são processadas. Para cada instrutor presente no relatório, uma nova planilha é criada e os dados são organizados.

```python
def processar_planilha(nome_arquivo):
    # Carregar a planilha, organizar dados e salvar novas planilhas para cada instrutor
```

3. **Envio de Emails com Outlook**: Com os relatórios processados, emails personalizados são enviados com os anexos correspondentes.

```python
for linha in range(2, len(sheet_selecionada['A']) + 1):
    nome = sheet_selecionada['A%s' % linha].value
    email = sheet_selecionada['C%s' % linha].value
    # Criar e enviar email com o anexo do relatório
```

### Customização

- **Modifique os campos de login e senha** no código conforme suas credenciais.
- **Caminho do diretório**: Ajuste o diretório de downloads e o diretório de destino conforme seu ambiente.
- **Datas e Campos Personalizados**: O código já possui suporte para inserir automaticamente a data atual, mas pode ser ajustado para outras necessidades de preenchimento de campos.

