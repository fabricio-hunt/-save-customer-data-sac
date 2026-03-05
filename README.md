# RPA - Automação Sistema de SAC Bemol

Automação robusta para extração de dados de clientes (CPF, Nome, Telefone) do **Sistema de Vendas Bemol SAC**. O robô lê IDs de uma planilha Excel (ou arquivo TXT), os insere no sistema (Delphi/VCL), lê os resultados na tela via API Win32 (`WM_GETTEXT`) e os salva cumulativamente em um arquivo Excel, garantindo a retomada (resume) automática em caso de interrupção.

---

## 📋 Status atual

> **✅ Em Produção — Leitura de campos resolvida e scripts processando em lote com sucesso!**

| Etapa | Status |
|---|---|
| Ler arquivos de entrada (`.xlsx` ou `.txt`) | ✅ Funcionando |
| Conectar na janela "Sistema de Vendas" e gerir foco | ✅ Funcionando (DPI Aware) |
| Digitar ID no campo Cliente e pressionar Enter | ✅ Funcionando |
| Ler dados transparentemente (sem depender de Ctrl+C) | ✅ Funcionando (Win32 API `WM_GETTEXT`) |
| Salvar planilha de saída (`Clientes_Extraidos_SAC.xlsx`) | ✅ Funcionando |
| Retomar processamento (Pular clientes "OK") | ✅ Funcionando (Salva a cada 50 clientes) |

### Como o problema de leitura foi resolvido:
O sistema SAC é construído em Delphi e utiliza componentes customizados da classe `TCompEdit`. Anteriormente, tentar copiar os dados via `Ctrl+A/Ctrl+C` falhava na hora que o foco mudava. A solução implantada foi a utilização da biblioteca **`win32gui` com DpiAwareness** que pesquisa o Handle (`hwnd`) exato de cada campo com base em suas coordenadas fixas, executando em seguida a mensagem nativa **`WM_GETTEXT`**. Isso tornou a extração virtualmente instantânea (milissegundos) e muito menos suscetível a erros de tela.

---

## 🖥️ Ambiente

| Item | Detalhe |
|---|---|
| SO | Windows |
| Python | 3.11+ |
| Monitores | Suporta dual-monitors. O script detecta automaticamente deslocamentos com o uso de `detectar_coords.py`. |
| Sistema alvo | "Sistema de Vendas" / "Sistema de SAC" (Delphi/VCL) |
| Janela Win32 | Título da janela contendo _"Benchimol, Irmão & Cia Ltda"_ ou _"Sistema de Vendas"_ |

---

## 📁 Arquivos do projeto

```text
etl-lisnildo-sac/
├── main.py                      # Script principal (Lê do .xlsx e processa)
├── processar-erros.py           # Script para reprocessar erros lendo de um .txt
├── detectar_coords.py           # Utilitário que encontra a tela e detecta as coordenadas dinamicamente
├── diagnostico.py               # Diagnostica a estrutura Win32 da interface e classes de hwnd
├── Clientes_Fora_Padrao 5.xlsx  # Planilha de entrada original (Lida por main.py)
├── reprocessar_erros.txt        # TXT com ids falhos que devem ser lidos (processar-erros.py)
└── Clientes_Extraidos_SAC.xlsx  # Planilha final de output (Lida e gravada pelos 2 scripts)
```

---

## ⚙️ Instalação

```bash
pip install pyautogui openpyxl pywin32 pyperclip
```

---

## 🗺️ Coordenadas e Detecção de Tela

As posições (x, y) de cada campo precisam estar parametrizadas no array `COORDS` dentro dos scripts `main.py` e `processar-erros.py`. Para calcular automaticamente isso baseado na posição exata da tela e no monitor atual (Esquerdo ou Direito), use:

```bash
python detectar_coords.py
```
O script exibirá exatamente o bloco `COORDS = { ... }` que precisa ser preenchido no código. O código é _DPI-Aware_ para contornar escalas de tela variáveis (como zooms do Windows a 125/150%).

---

## 🚀 Uso e Processamento

### Processamento Principal (`main.py`)
Para executar em lote lendo a planilha inteira.
```bash
python main.py
```
* O fluxo identificará caso o `Clientes_Extraidos_SAC.xlsx` já possua registros com `Status='OK'` e os ignorará, permitindo que a execução interropida (por queda de energia, manual, etc) retorne exatamente de onde parou.
* A gravação do progresso acontece no disco **a cada 50 clientes**.

### Processamento de Erros ou Lista Separada (`processar-erros.py`)
Para rodar especificamente os registros salvos em `reprocessar_erros.txt` (um ID por linha). O script une os resultados, atualizando os status de falhas recentes na própria planilha `Clientes_Extraidos_SAC.xlsx`.
```bash
python processar-erros.py
```

### Modo de Teste / Debug
Verifica rapidamente se ele consegue ler os campos expostos em tela.
```bash
python main.py --debug
```

* **Nota de Emergência:** Para frear a automação abruptamente durante o movimento, deslize depressa o mouse para o **Canto Superior Esquerdo** da tela (ativando o FailSafe do PyAutoGUI).

---

## 📊 Estrutura da Planilha de Saída (`Clientes_Extraidos_SAC.xlsx`)

O layout final foi condensado para:

| Coluna | Descrição |
|---|---|
| `ID` | ID do cliente |
| `Nome` | Primeiro nome extraído do cliente |
| `Sobrenome` | O restante do nome |
| `CPF` | CPF formatado (000.000.000-00) ou CNPJ |
| `Telefone` | O campo correspondente a contato telefônico |
| `Status` | `OK` ou `ERRO: ...`/`Cliente não encontrado` |

---

## ⚡ Ajustes de Tempo de Resposta

No início dos scripts `main.py` e `processar-erros.py`, defina as tolerâncias caso a conexão do sistema seja lenta:
```python
PAUSA_APOS_BUSCA = 10.0   # segundos aguardando o sistema carregar visualmente o cliente
PAUSA_ENTRE_IDS  = 1.0    # pausa entre cada verificação e o início do próximo
```
