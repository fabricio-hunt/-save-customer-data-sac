# RPA - Automação Sistema de SAC Bemol

Automação para extração de dados de clientes (CPF, Nome, Telefone) do **Sistema de Vendas Bemol SAC**, lendo IDs de uma planilha Excel e salvando os resultados em outra planilha.

---

## 📋 Status atual

> **⚠️ Em desenvolvimento — parado no ponto de leitura dos campos**

| Etapa | Status |
|---|---|
| Ler planilha de entrada (`Clientes_Fora_Padrao 5.xlsx`) | ✅ Funcionando |
| Conectar na janela "Sistema de Vendas" | ✅ Funcionando |
| Digitar ID no campo Cliente e pressionar Enter | ✅ Funcionando |
| Sistema carrega os dados do cliente na tela | ✅ Funcionando |
| Ler CPF, Nome e Telefone da tela | ❌ **Pendente** |
| Salvar planilha de saída (`Clientes_Extraidos_SAC.xlsx`) | ⏳ Depende do item acima |

### Problema atual
O sistema usa controles **`TCompEdit`** (Delphi/VCL customizado) que bloqueiam:
- `WM_GETTEXT` — retorna vazio
- `WM_COPY` — não copia para clipboard
- `SetFocus` programático — retorna "Acesso negado"

O `Ctrl+A` + `Ctrl+C` **funciona manualmente** (confirmado colando no Bloco de Notas), mas o script não consegue copiar porque o **terminal/VS Code rouba o foco** antes do `Ctrl+C` ser executado.

**Próxima tentativa:** o script agora minimiza o terminal automaticamente antes de operar. Testar amanhã com `python main.py --debug`.

---

## 🖥️ Ambiente

| Item | Detalhe |
|---|---|
| SO | Windows 10/11 |
| Python | 3.11 **32-bit** (obrigatório — sistema SAC é 32-bit) |
| Monitors | 2 monitores — SAC ocupa a tela **esquerda** inteira |
| Sistema alvo | "Sistema de Vendas" / "Sistema de SAC" (Delphi/VCL) |
| Janela Win32 | Título: `Benchimol, Irmão & Cia Ltda...` / hwnd varia |
| Controles | Classe `TCompEdit` — não expõe texto via Win32 padrão |

---

## 📁 Arquivos do projeto

```
etl-lisnildo-sac/
├── main.py                      # Script principal do RPA (= rpa_bemol_sac.py)
├── window.py                    # Script utilitário — mostra posição do mouse
├── capturar_coords.py           # Captura coordenadas clicando nos campos
├── diagnostico.py               # Diagnóstico de janelas e sessão Windows
├── Clientes_Fora_Padrao 5.xlsx  # Planilha de entrada com coluna ID_CLIENTE
├── Clientes_Extraidos_SAC.xlsx  # Planilha de saída (gerada pelo RPA)
└── rpa_log.txt                  # Log de execução (gerado automaticamente)
```

---

## ⚙️ Instalação

```bash
# Python 32-bit obrigatório
# Baixe em: https://www.python.org/downloads/windows/
# Marque "Add to PATH" durante a instalação

pip install pyautogui openpyxl pywin32 pyperclip pynput
```

---

## 🗺️ Coordenadas dos campos (tela atual)

Capturadas com `python window.py` — **devem ser reconfirmadas se a janela mover**.

```python
COORDS = {
    "Cliente"    : (324, 74),
    "CPF"        : (449, 77),
    "Nome"       : (324, 98),
    "Telefone"   : (300, 120),
    "Limpar tudo": (648, 120),
}
```

Para recapturar:
```bash
python window.py
# Passe o mouse sobre cada campo e anote as coordenadas
```

---

## 🚀 Como usar

### 1. Testar leitura dos campos (DEBUG)
```bash
# Com um cliente carregado na tela do SAC:
python main.py --debug

# O terminal será minimizado automaticamente após 3 segundos.
# Verifique o output no rpa_log.txt
```

**Saída esperada quando funcionar:**
```
[INFO]   Cliente         lido='2090418'
[INFO]   CPF             lido='95508082204'
[INFO]   Nome            lido='PATRICIA ARAUJO OLIVEIRA'
[INFO]   Telefone        lido='92985942997'
```

### 2. Executar em massa
```bash
python main.py
```
- Aguarda 3 segundos (minimiza terminal)
- Processa os 3698 IDs da planilha
- Salva resultados em `Clientes_Extraidos_SAC.xlsx`
- **Para emergência:** mova o mouse para o **CANTO SUPERIOR ESQUERDO**

---

## 🔍 Histórico de tentativas para leitura dos campos

| Abordagem | Resultado |
|---|---|
| `pywinauto` + `children(class_name="Edit")` | ❌ Lista vazia — `TCompEdit` não é classe `Edit` padrão |
| `win32gui.EnumChildWindows` buscando `"Edit"` | ❌ Lista vazia — mesma razão |
| `WM_GETTEXT` direto no hwnd do `TCompEdit` | ❌ Retorna vazio |
| `WM_COPY` direto no hwnd | ❌ Não copia |
| `SetFocus` + `Ctrl+A` + `Ctrl+C` programático | ❌ "Acesso negado" (processo diferente) |
| `pyautogui.click` + `Ctrl+A` + `Ctrl+C` | ⚠️ Terminal rouba foco antes do `Ctrl+C` |
| OCR com Tesseract | ❌ Tesseract não instalado (sem compilador C++) |
| Click + minimizar terminal + `Ctrl+C` | 🔄 **Próxima tentativa** |

**Confirmado manualmente:** `Ctrl+A` + `Ctrl+C` no campo funciona (colagem no Bloco de Notas retorna o texto correto). O problema é exclusivamente de **foco da janela** durante a automação.

---

## 💡 Próximas abordagens caso o foco continue falhando

1. **Rodar o script de uma janela CMD separada** (não VS Code) — o VS Code pode estar interceptando o foco
2. **`pyautogui` com janela SAC maximizada** na tela inteira, terminal na outra tela
3. **Criar um `.bat`** que abre o script minimizado para não competir pelo foco
4. **Hook de teclado global** (`pynput`) para capturar o clipboard no momento exato
5. **Ler o processo de memória** do SAC via `ReadProcessMemory` (requer privilégios)

---

## 📊 Planilha de entrada

**Arquivo:** `Clientes_Fora_Padrao 5.xlsx`  
**Coluna obrigatória:** `ID_CLIENTE`  
**Total de IDs:** 3698 únicos

## 📊 Planilha de saída

**Arquivo:** `Clientes_Extraidos_SAC.xlsx`

| Coluna | Descrição |
|---|---|
| `ID_CLIENTE` | ID original da planilha de entrada |
| `CPF_CNPJ` | CPF formatado (000.000.000-00) ou CNPJ |
| `PRIMEIRO_NOME` | Primeira palavra do nome |
| `SOBRENOME` | Restante do nome |
| `CONTATO` | Primeiro telefone cadastrado |
| `STATUS` | `OK` ou `ERRO: <mensagem>` |

---

## ⚡ Configurações ajustáveis no `main.py`

```python
PAUSA_APOS_BUSCA = 10.0   # segundos aguardando o sistema carregar o cliente
PAUSA_ENTRE_IDS  = 1.0    # pausa entre cada cliente processado
```
