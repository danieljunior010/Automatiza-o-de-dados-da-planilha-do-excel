# Automação de Cadastro de Produtos

Este projeto é um script Python que automatiza o processo de cadastro de produtos em uma aplicação utilizando as bibliotecas `openpyxl`, `pyperclip` e `pyautogui`. Ele lê os dados de uma planilha Excel e preenche automaticamente os campos correspondentes na interface da aplicação.

## Funcionalidades

- Leitura de uma planilha Excel (`produtos_ficticios.xlsx`), onde os dados dos produtos estão armazenados.
- Preenchimento automático de campos como Nome do Produto, Descrição, Categoria, Código do Produto, Peso, Dimensões, Preço, Estoque, Data de Validade, Cor, Tamanho, Material, Fabricante, País de Origem, Observações, Código de Barras, e Localização do Armazém.
- Navegação automática na interface gráfica da aplicação, simulando cliques e colando informações.
- Uso de delays (`sleep`) para garantir que o processo de automação siga o ritmo da aplicação.
  
## Tecnologias Utilizadas

- **Python**: Linguagem principal do script.
- **[openpyxl](https://openpyxl.readthedocs.io/en/stable/)**: Biblioteca para leitura e manipulação de arquivos Excel.
- **[pyperclip](https://pypi.org/project/pyperclip/)**: Biblioteca para copiar e colar textos no clipboard.
- **[pyautogui](https://pyautogui.readthedocs.io/en/latest/)**: Biblioteca para automação de mouse e teclado.

## Como Executar

1. **Instale as dependências**:
   Você precisará instalar as bibliotecas usadas no projeto. Execute o seguinte comando:
   ```bash
   pip install openpyxl pyperclip pyautogui
