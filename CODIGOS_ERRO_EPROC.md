# Códigos de Erro do EPROC

Códigos de erro de negócio retornados pelo e-Proc durante a migração. Aparecem na coluna **Status** no formato `Erro X: Descrição`.

## Códigos Conhecidos

| Código | Descrição |
|--------|-----------|
| 0 | Processo já Migrado |
| 1 | Erro INESPERADO NA MIGRAÇÃO |
| 5 | Processo NÃO CONSTA na base do SAJ |
| 6 | Migração NÃO HABILITADA para este processo |
| 13 | JUÍZO NÃO HABILITADO PARA MIGRAÇÃO |
| 14 | COMPETÊNCIA não cadastrada no EPROC |
| 20 | Existem MANDADOS PENDENTES |
| 23 | Existem ARs sem Devolução |
| 25 | Erros nas peças processuais |
| 34 | MIGRAÇÃO NÃO PERMITIDA |
| 35 | Erro no Cadastro de Parte |
| 36 | Inconsistência no Cadastro de Parte |
| 47 | Advogado ou Escritório não cadastrado no EPROC |
| 51 | Processo Será Migrado Após Solução do Incidente/Apenso |
| 60 | Migração NÃO HABILITADA, Processo Entranhado |
| 61 | Processo em Segundo Grau - Agravo/Recurso |
| 66 | Existem ARs Ag. Envio aos Correios |
| 89 | Existem mandados pendentes que ainda não constam na fila e situação: Ag. Cumprimento pelo Oficial |
| 95 | PROCESSO BAIXADO NÃO SERÁ MIGRADO |
| 96 | CNPJ NÃO INFORMADO |

## Como Adicionar Novos Códigos

Edite o dicionário `CODIGOS_ERRO_EPROC` em `funcoes/eproc.py`:

```python
CODIGOS_ERRO_EPROC = {
    # ...códigos existentes...
    99: "Descrição do novo erro",
}
```

## Exemplos de Saída na Planilha

**Erro único:**
```
Erro 61: Processo em Segundo Grau - Agravo/Recurso
```

**Múltiplos erros:**
```
Erro 25: Erros nas peças processuais; Erro 61: Processo em Segundo Grau - Agravo/Recurso
```

## Observações

- O robô captura **automaticamente** todos os códigos únicos exibidos na página
- Se o mesmo código aparecer várias vezes (ex: código 25 para várias páginas), é registrado **uma única vez**
- Erros são separados por `"; "` quando há mais de um
- Erros do sistema (não de negócio) aparecem como `ERRO SISTEMA: ...` — ver tabela completa em `INSTRUCOES_USO.md`
