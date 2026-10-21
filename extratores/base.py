"""Strategy: cada fornecedor (ou família de fornecedores) tem seu próprio Extrator,
responsável pela parte que REALMENTE varia de um fornecedor para outro:
    - onde/como ler os lançamentos (qual aba, qual coluna de placa/valor)
    - qual conta contábil usar em cada linha
    - qualquer cálculo específico (ex.: abatimento de estornos do EDENRED)

Tudo o que é comum (localizar arquivo, localizar designação, enriquecer com
AN/CCO/CC/TIPO, gerar o Excel final) fica no `BaseFornecedorProcessor`
(processor.py) e NÃO deve ser reimplementado aqui.
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path
from typing import List, Sequence

from ..modelos import Lancamento


class ExtratorBase(ABC):
    """Interface Strategy para extração + cálculo específico de um fornecedor."""

    #: Se True, o processor usa o fallback de designação histórica quando a
    #: placa não é achada na designação normal do mês. RELEMAR, por exemplo,
    #: usa False (comportamento original: nunca caía no histórico).
    usar_historico: bool = True

    #: Se True, o Excel final ganha a coluna SERVIÇO (UNIDAS e EDENRED).
    incluir_servico: bool = False

    #: Nomes de colunas extras (vindas de ``lanc.extras``) inseridas entre
    #: VALOR (ou SERVIÇO) e AN no Excel final.
    colunas_pos_valor: Sequence[str] = ()

    #: Nomes de colunas extras inseridas DEPOIS de VALIDAÇÃO no Excel final.
    colunas_pos_validacao: Sequence[str] = ()

    @abstractmethod
    def extrair(self, caminho_excel_fornecedor: Path) -> List[Lancamento]:
        """Lê o arquivo do fornecedor e retorna os Lancamento já com `conta`
        (e `servico`, quando aplicável) resolvidos. `an`/`cco`/`cc`/`tipo`
        ainda não devem ser preenchidos aqui -- isso é feito depois pelo
        DesignacaoService, no processor.
        """
        raise NotImplementedError

    def pos_processar(self, lancamentos: List[Lancamento]) -> List[Lancamento]:
        """Hook opcional, executado DEPOIS do enriquecimento com designação
        (ou seja, já com AN/CCO/CC preenchidos). Só o EDENRED usa isso, para
        abater linhas de 'Estorno' contra linhas de 'Pedágio' do mesmo grupo
        AN/CCO/CC. Por padrão, não faz nada.
        """
        return lancamentos