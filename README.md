# ROBO_EPROC_MIGRADOR

Automação para migração de processos no sistema e-Proc do TJSP.

## Requisitos

- Python 3.8 ou superior
- Google Chrome instalado
- ChromeDriver (o Selenium gerencia automaticamente)

## Instalação

1. Clone ou baixe este repositório

2. Instale as dependências:
```bash
pip install -r requirements.txt
```

## Como Usar

1. Execute o programa:
```bash
python main.py
```

2. O navegador Chrome será aberto automaticamente
3. Realize o login no e-Proc manualmente (se necessário)
4. Escolha o **Modo** e o **Timeout** na barra lateral e clique em **Migrador**
5. Clique em **Executar** e selecione o arquivo Excel com os processos

## Modos do Migrador

| Modo | Comportamento |
|------|--------------|
| **Padrão (pula sem CPF)** | Pula processos que apresentam "Pessoas sem CPF/CNPJ" |
| **Trata processos sem CPF** | Tenta selecionar "Parte SEM CPF" nos dropdowns e prosseguir com a migração |

## Formato do Arquivo Excel

Prepare um arquivo `.xlsx` com a aba chamada **Planilha1** contendo:

- **Processo**: Número do processo (pode incluir incidente separado por `/`, ex: `1234567-89.2023.8.26.0100/1`)
- **Status**: Será preenchido automaticamente pelo programa (criada automaticamente se não existir)

Exemplo:
| Processo | Status |
|----------|--------|
| 1234567-89.2023.8.26.0100 | |
| 2345678-90.2023.8.26.0100/1 | |

## Observações

- O programa mantém os dados de sessão do navegador em `eproc_user_data/`
- O status de cada processo é atualizado automaticamente na planilha Excel após cada tentativa
- Não feche o navegador manualmente durante o processamento
- Use o botão **Parar** na interface para interromper com segurança

## Estrutura do Projeto

```
ROBO_EPROC_MIGRADOR/
├── main.py                    # Interface gráfica + ponto de entrada
├── funcoes/
│   ├── navegador.py           # Controle do Chrome
│   ├── eproc.py               # Lógica de migração (migrador e migrador_sem_cpf)
│   └── ui_utils.py
├── requirements.txt
├── INSTRUCOES_USO.md
└── CODIGOS_ERRO_EPROC.md
```
