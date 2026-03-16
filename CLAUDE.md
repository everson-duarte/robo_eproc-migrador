# CLAUDE.md — Robo e-Proc Migrador

## O que é este projeto

Automação desktop em Python para migração de processos judiciais no sistema **e-Proc do TJSP** (Tribunal de Justiça de São Paulo). Usa Selenium para controlar o Chrome e CustomTkinter para a interface gráfica.

## Estrutura

```
ROBO_EPROC_MIGRADOR_V1.1/
├── main.py                  # Ponto de entrada — interface gráfica (CustomTkinter)
├── funcoes/
│   ├── eproc.py             # Lógica principal: migrador, migrador_sem_cpf, acessar_localizadores
│   ├── navegador.py         # Controle do Chrome via Selenium (driver global)
│   ├── logger.py            # Logger único "eproc_migrador" — arquivo + console
│   └── ui_utils.py          # CustomDialog: janela modal reutilizável (info/warning/error/yesno)
├── requirements.txt         # selenium, customtkinter, openpyxl
├── logs/migracao.log        # Gerado em runtime, ignorado pelo git
├── README.md
├── INSTRUCOES_USO.md
└── CODIGOS_ERRO_EPROC.md    # Dicionário de códigos de erro do e-Proc
```

## Dependências

```
selenium==4.27.1
customtkinter==5.2.2
openpyxl==3.1.5
```

Instalar com: `pip install -r requirements.txt`

Requer **Google Chrome** instalado. O ChromeDriver é gerenciado automaticamente pelo Selenium.

## Como executar

```bash
python main.py
```

O Chrome abre automaticamente e aguarda o login manual no e-Proc (timeout de 90s).

## Fluxo principal

1. `main.py` → `acessar_eproc()` abre o Chrome e aguarda login
2. Usuário seleciona função na sidebar e clica **Executar**
3. A função roda em thread separada (`threading.Thread`) para não travar a UI
4. O `cancel_event` (`threading.Event`) permite interrupção segura via botão **Parar**
5. Logs aparecem no painel escuro da UI e são gravados em `logs/migracao.log`

## Funções em `eproc.py`

| Função | Descrição |
|--------|-----------|
| `migrador()` | Migra processos de uma planilha Excel; pula partes sem CPF/CNPJ |
| `migrador_sem_cpf()` | Igual ao migrador, mas tenta tratar partes sem CPF selecionando "Parte SEM CPF" nos dropdowns |
| `acessar_localizadores()` | Navega até "Meus Localizadores" no e-Proc |
| `registrar_cancelamento(event)` | Registra o `cancel_event` para interrupção entre iterações |

## Formato da planilha Excel

Aba: **Planilha1**

| Coluna | Descrição |
|--------|-----------|
| `Processo` | Número do processo. Aceita incidente separado por `/` (ex: `1234567-89.2023.8.26.0100/1`) |
| `Status` | Preenchido automaticamente após cada tentativa |

## Timeout configurável

Definido via `eproc_module.TIMEOUT_MIGRACAO` (segundos). Opções na UI: 3 min / 5 min / 9 min.

## Padrões do projeto

- **Driver Selenium:** variável global em `navegador.py`; sempre acesse via `obter_driver()`
- **Logger:** use sempre `from .logger import logger` — nunca crie loggers novos
- **Diálogos:** use `CustomDialog` de `ui_utils.py` em vez de `tkinter.messagebox`
- **Cancelamento:** verifique `cancel_event.is_set()` entre operações longas em `eproc.py`
- **Threading:** toda ação demorada roda em thread daemon; atualizações de UI usam `app.after(0, ...)`

## O que NÃO versionar

- `logs/` — logs locais de execução
- `dist/` e `build/` — artefatos do PyInstaller
- `funcoes/__pycache__/` — cache do Python
- Planilhas `.xlsx` com dados reais de processos
