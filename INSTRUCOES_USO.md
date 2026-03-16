# Instruções de Uso - Robô Migrador e-Proc

## Visão Geral

Tudo é controlado por um único aplicativo: `main.py`.

A interface lateral permite escolher o **Modo** e o **Timeout** antes de executar.

---

## Modos do Migrador

### Modo 1 — Padrão (pula sem CPF)
**Quando usar:** Primeira rodada geral de migração.

**Comportamento:**
- Migra processos normalmente
- **Pula** processos que apresentam "Pessoas sem CPF/CNPJ"
- Registra `Não Migrado SEM CPF-CNPJ` na planilha

---

### Modo 2 — Trata processos sem CPF
**Quando usar:** Para a lista de processos que ficaram com "Não Migrado SEM CPF-CNPJ".

**Comportamento:**
- Migra processos normalmente
- Ao encontrar "Pessoas sem CPF/CNPJ":
  - Seleciona automaticamente a opção **"Parte SEM CPF"** nos dropdowns
  - Marca o checkbox de declaração automaticamente
  - Se conseguir tratar, prossegue com a migração
  - Se ainda houver erros impeditivos, registra o erro específico

---

## Como Executar

```powershell
python main.py
```

1. O Chrome abre automaticamente — faça login no e-Proc se necessário
2. Na sidebar, selecione o **Modo** e o **Timeout**
3. Clique em **Migrador** e depois em **Executar**
4. Selecione o arquivo Excel com os processos
5. O robô processa e atualiza a planilha automaticamente
6. Use **Parar** para interromper com segurança a qualquer momento

---

## Formato do Arquivo Excel

- Aba obrigatória: **Planilha1**
- Coluna **Processo**: número do processo (com ou sem incidente separado por `/`)
- Coluna **Status**: preenchida automaticamente (criada se não existir)

Exemplos de formatos aceitos na coluna Processo:
```
1234567-89.2023.8.26.0100
1234567-89.2023.8.26.0100/1
```

---

## Status Registrados na Planilha

| Status | Significado |
|--------|-------------|
| `Migrado com Sucesso` | Processo migrado |
| `Não Migrado SEM CPF-CNPJ` | Pulado por ter partes sem CPF (Modo Padrão) |
| `Sem CPF Tratado - Erro X: ...` | Tratou CPF mas ainda há erro impeditivo |
| `Não Migrado - Falha ao Tratar SEM CPF` | Não conseguiu selecionar "Parte SEM CPF" |
| `Erro X: Descrição` | Erro de negócio retornado pelo e-Proc |
| `ERRO SISTEMA: Falha na validação da requisição...` | Exceção do sistema após clicar em Migrar — processo pulado automaticamente |
| `ERRO SISTEMA: Timeout após migração` | Sistema não respondeu no tempo configurado |
| `ERRO SISTEMA: Timeout ao carregar dados` | Página demorou mais que 5 min para carregar |
| `ERRO SISTEMA: Falha ao processar` | Erro inesperado — robô tenta recuperar e continua |

---

## Timeout da Migração

Após clicar em Migrar, o robô aguarda o resultado por até o tempo configurado. Altere na sidebar antes de executar:

| Opção | Tempo | Quando usar |
|-------|-------|-------------|
| **3 min** | 180s | Primeira rodada — processa os simples rápido; complexos viram "Timeout" para retry |
| **5 min** | 300s | Padrão — bom equilíbrio |
| **9 min** | 540s | Retry dos complexos — processos com muitas páginas |

**Estratégia recomendada:**
1. Primeira rodada com **3 min** — rápido nos simples
2. Segunda rodada com **5 ou 9 min** — apenas os que ficaram com `ERRO SISTEMA: Timeout após migração`

---

## Processamento em Paralelo

Com hardware suficiente (16GB+ RAM), é possível rodar múltiplas instâncias simultaneamente, cada uma com sua própria planilha.

**Passo a passo:**

1. Divida a planilha em lotes (ex: 4 arquivos com ~1.650 processos cada)
2. Abra múltiplos terminais (`Ctrl + Shift + ``)
3. Execute `python main.py` em cada terminal e selecione um lote diferente

**Dica — Áreas de Trabalho Virtuais (Windows):**
- `Win + Ctrl + D` → Nova área de trabalho
- `Win + Ctrl + ←/→` → Alternar entre áreas
- `Win + Tab` → Visualizar todas

---

## Estimativa de Tempo

Referência: ~200 processos/hora por instância (pode variar conforme estabilidade do e-Proc).

| Instâncias | 6.600 processos |
|------------|-----------------|
| 1 | ~33 horas |
| 2 | ~17 horas |
| 3 | ~11 horas |
| 4 | ~8 horas |

---

## Regras Importantes

**Pode:**
- Usar mouse e teclado normalmente em outros programas
- Navegar na internet em outro navegador
- Trabalhar em outros monitores

**Não pode:**
- Clicar no Chrome que o robô está usando
- Abrir ou editar a mesma planilha Excel que o robô processa
- Rodar duas instâncias na mesma planilha

---

## Em Caso de Problema

1. Verifique os logs no terminal — o robô imprime o status de cada processo
2. Abra a planilha Excel para ver o status registrado
3. Use o botão **Parar** para interromper com segurança
4. Reexecute apenas os processos com erro (filtre pela coluna Status)
