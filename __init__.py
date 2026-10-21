"""
FR_POO
======

Pacote que unifica o processamento de consolidação de gastos de frotas
por fornecedor (antes eram ~14 scripts quase idênticos).

Arquitetura:
    - config.py         -> caminhos, mês atual, tabela de contas contábeis
    - utils.py           -> normalização de texto/cabeçalho, achar linha de cabeçalho
    - modelos.py          -> dataclass Lancamento (uma linha do CONSO final)
    - designacao.py      -> DesignacaoService (lookup PLACA -> AN/CCO/CC/TIPO, com fallback histórico)
    - excel_writer.py     -> escreve o Excel "CONSO ..." final
    - extratores/         -> um Strategy por "família" de fornecedor (extração + cálculo)
    - processor.py        -> BaseFornecedorProcessor (Template Method = fluxo fixo)
    - factory.py           -> FornecedorProcessorFactory (cria o processor certo)
    - main.py              -> orquestrador (roda todos os fornecedores configurados)
"""