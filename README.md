Este é um aplicativo desktop em Python com interface gráfica (GUI) que automatiza a coleta e preenchimento de dados de PLD (Preço da Luz de Desembaraço) do submercado SUL, obtidos do site da CCEE.

🎯 Funcionalidades Principais:
1. BUSCAR VALORES PLD-SUL
Usa Selenium WebDriver para acessar o site da CCEE (https://www.ccee.org.br)
Seleciona automaticamente dados "HORÁRIO" para um período específico
Filtra apenas dados do submercado SUL
Baixa arquivo (Excel ou CSV) e extrai os valores da 3ª coluna
Permite escolher entre data atual ou dia anterior
2. PREENCHER PLD
Carrega um arquivo Excel com estrutura PLD
Localiza linhas com a data selecionada (coluna A)
Preenche a coluna D com os valores do SUL obtidos
Abre o arquivo automaticamente após salvar
3. COPIAR VALORES
Copia os valores exibidos na tela para a área de transferência
📋 Componentes Técnicos:
Componente	Propósito
tkinter	Interface gráfica (janelas, botões, campos de texto)
Selenium	Automação do navegador Chrome
openpyxl	Manipulação de arquivos Excel
pandas	Processamento de dados (CSV/Excel)
glob	Busca de arquivos recentes no Download
⚙️ Recursos Internos:
Persistência: Salva o caminho do arquivo PLD em config.txt
Log de Erros: Registra erros com data/hora em log_erros.txt
Compatibilidade: Detecta se está sendo executado como executável ou script Python
Tratamento de Exceções: Try/except robusto para cada etapa crítica
🔧 Interface Gráfica:
Code
┌─ Monitor PLD - SUL ──────────────────┐
│ [Procurar] Arquivo PLD               │
│ ◉ Dia atual  ○ Dia anterior         │
│ [BUSCAR VALORES PLD-SUL]             │
│ ┌──────────────────────────────────┐ │
│ │ (área de texto com valores)      │ │
│ └──────────────────────────────────┘ │
│ [COPIAR VALORES]                     │
│ [PREENCHER PLD]                      │
└──────────────────────────────────────┘
💡 Caso de Uso:
Operador de energia que precisa automatizar a atualização de sua planilha PLD com dados do mercado CCEE, economizando tempo e reduzindo erros manuais.
