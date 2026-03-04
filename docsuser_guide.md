# 📖 Guia do Usuário — Excel VBA Automation

## Pré-requisitos

- Microsoft Excel 365 (Windows)
- Macros habilitadas no Excel
- Python 3.8+ (opcional, para scripts auxiliares)

---

## 1. Configuração Inicial

### 1.1 Habilitar Macros
1. Abra o arquivo `dashboard.xlsm`
2. Clique em **Habilitar Conteúdo** na barra amarela
3. Ou vá em **Arquivo → Opções → Central de Confiabilidade → Configurações → Habilitar todas as macros**

### 1.2 Acessar o Editor VBA
- Atalho: `Alt + F11`
- Ou: **Desenvolvedor → Visual Basic**

> ⚠️ Se a aba **Desenvolvedor** não aparecer:
> Arquivo → Opções → Personalizar Faixa de Opções → marcar **Desenvolvedor**

### 1.3 Importar os módulos `.bas`
1. No Editor VBA, clique com botão direito em **VBAProject**
2. Selecione **Importar Arquivo**
3. Importe nesta ordem:
   - `Module_Utils.bas`
   - `Module_API.bas`
   - `Module_Automation.bas`

---

## 2. Módulos Disponíveis

### Module_Utils — Utilitários Gerais

| Função/Sub | Descrição |
|---|---|
| `SpeedOn` / `SpeedOff` | Liga/desliga otimizações de performance |
| `SheetExists(nome)` | Verifica se uma aba existe |
| `GetOrCreateSheet(nome)` | Cria ou retorna uma aba |
| `LastRow(ws, col)` | Última linha preenchida |
| `LastCol(ws, row)` | Última coluna preenchida |
| `FormatAsTable(ws, rng, nome)` | Formata intervalo como tabela |
| `LogError(source, num, desc)` | Registra erros na aba "Log" |
| `SaveBackup()` | Salva cópia com timestamp |
| `ToBRL(valor)` | Formata número como R$ |

### Module_API — Integração com APIs

| Função/Sub | Descrição |
|---|---|
| `HttpGet(endpoint)` | Requisição GET e retorna JSON |
| `HttpPost(endpoint, payload)` | Requisição POST com JSON |
| `ParseJsonValue(json, chave)` | Extrai valor de campo JSON |
| `FetchExchangeRate(from, to)` | Busca cotação de moeda |
| `PostSheetData(aba, endpoint)` | Envia dados de aba para API |
| `ImportApiData(endpoint, aba)` | Importa dados da API para aba |

### Module_Automation — Automação Principal

| Sub | Descrição |
|---|---|
| Macros de automação de tarefas repetitivas | Ver arquivo `Module_Automation.bas` |

---

## 3. Como Executar as Macros

### Via Atalho de Teclado
1. `Alt + F8` para abrir a lista de macros
2. Selecione a macro desejada
3. Clique em **Executar**

### Via Botão no Dashboard
- Clique nos botões disponíveis na aba **Dashboard**

### Via Editor VBA
1. `Alt + F11` para abrir o editor
2. Posicione o cursor dentro da Sub desejada
3. Pressione `F5` para executar

---

## 4. Configurar a API

Abra `Module_API.bas` e altere as constantes no topo:
```vba
Private Const BASE_URL As String = "https://sua-api.com/v1"
Private Const API_KEY  As String = "sua-chave-aqui"
```

---

## 5. Estrutura das Abas Esperadas

| Aba | Descrição |
|---|---|
| `Dashboard` | Visão geral com KPIs e gráficos |
| `Dados` | Dados brutos para processamento |
| `API_Data` | Dados importados via API |
| `Log` | Registro automático de erros |
| `Config` | Parâmetros configuráveis |

---

## 6. Solução de Problemas

| Problema | Solução |
|---|---|
| "Macros desabilitadas" | Habilite macros nas configurações (ver 1.1) |
| Erro ao chamar API | Verifique `BASE_URL` e `API_KEY` em `Module_API.bas` |
| Aba não encontrada | Verifique o nome exato da aba (case-sensitive) |
| Código lento | Certifique-se de chamar `SpeedOn` antes de loops grandes |
| Erro 429 na API | API com rate limit — adicione `Application.Wait` entre chamadas |
| `.bas` não importa | Certifique-se de que o arquivo não está aberto em outro programa |

---

## 7. Boas Práticas

- Sempre chame `SpeedOn` antes de loops e `SpeedOff` ao final
- Use `LogError` em todos os blocos `On Error`
- Faça backup antes de rodar macros destrutivas (`SaveBackup`)
- Nunca deixe `API_KEY` hardcoded em produção — use a aba `Config`

---

## 8. Links Úteis

- [Documentação VBA — Microsoft](https://learn.microsoft.com/office/vba/api/overview/excel)
- [MSXML2.XMLHTTP — Referência](https://learn.microsoft.com/previous-versions/windows/desktop/ms759148(v=vs.85))
- [Excel 365 — Novidades](https://support.microsoft.com/excel)
```

