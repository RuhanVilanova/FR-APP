# FR-APP — Conciliação de Despesas de Frota

Aplicação web (Databricks App) que automatiza a conciliação mensal de despesas de frota de veículos.
Lê as faturas em Excel/PDF de diferentes fornecedores (locadoras, rastreamento, pedágio, combustível,
manutenção), classifica cada lançamento em uma conta contábil, enriquece por placa com dados de
designação (centro de custo etc.) e gera a planilha consolidada **CONSO** pronta para lançamento.

O projeto nasceu da refatoração de ~14 scripts quase idênticos para uma arquitetura orientada a objetos.

## Destaques técnicos

- **Strategy** — cada família de fornecedor tem seu extrator (`extratores/`), subclasse de `ExtratorBase`.
- **Template Method** — `BaseFornecedorProcessor` (`processor.py`) define o fluxo fixo: localizar → extrair → designar → escrever.
- **Factory** — `FornecedorProcessorFactory` (`factory.py`) registra os fornecedores suportados.
- **Service** — `DesignacaoService` resolve placa → centro de custo, com fallback em base histórica e cache.
- Regras de classificação em cascata para manutenção (peça / mão de obra) com balanceamento guloso entre contas.
- Geração de Excel com `openpyxl`: colunas dinâmicas, fórmulas, estilos e bloco de rateio (OC) com arredondamento `Decimal`.
- Extração de dados de PDF com `pdfplumber`.
- I/O em Unity Catalog Volumes via Databricks SDK (`volume_io.py`).
- Frontend em HTML + Tailwind CSS (upload, execução por fornecedor, download dos resultados).

## Estrutura

| Arquivo | Papel |
| --- | --- |
| `server.py` | Entrypoint Flask do Databricks App (`app.yaml`) |
| `main.py` | Execução local (`python main.py`) |
| `config.py` | Caminhos, mês vigente e tabela de contas |
| `processor.py` / `factory.py` | Fluxo de processamento e registro de fornecedores |
| `extratores/` | Um extrator por família de fornecedor |
| `excel_writer.py` | Escrita do CONSO e bloco OC |
| `designacao.py` | Enriquecimento por placa |
| `volume_io.py` | Acesso ao UC Volume via SDK |
| `index.html` | Interface web |

## Como executar

1. `pip install -r requirements.txt`
2. Defina `VOLUME_URI` apontando para o seu UC Volume (ou ajuste o default em `config.py`).
3. Copie `Contas_ticket.example.json` para `Contas_ticket.json`.
4. Local: `python main.py` — Databricks App: deploy da pasta com `app.yaml`.

> **Nota:** esta é uma versão pública sanitizada. Números de contas contábeis, caminhos internos e
> mapeamentos reais foram substituídos por valores fictícios.
