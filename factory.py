"""FornecedorProcessorFactory: ponto único que sabe, para cada fornecedor,
qual Extrator (Strategy) usar e com quais parâmetros -- ou seja, é aqui (e só
aqui) que fica a configuração que antes era o bloco `FORNECEDOR_CONFIG` no
topo de cada script.

Uso:
    processor = FornecedorProcessorFactory.criar("VAMOS")
    resultado = processor.processar()

    # ou todos de uma vez:
    for nome in FornecedorProcessorFactory.fornecedores_disponiveis():
        FornecedorProcessorFactory.criar(nome).processar()
"""
from __future__ import annotations

from typing import Callable, Dict, List

from .extratores import (
    ColunaSimplesExtractor,
    CobrancaRelacaoExtractor,
    DescricaoMapeadaExtractor,
    ListaAtivosExtractor,
    MultasTaxaExtractor,
    MultiSheetContaExtractor,
    TicketCombustivelExtractor,
    TicketManutencaoExtractor,
    TipoTransacaoExtractor,
)
from .processor import BaseFornecedorProcessor, GeomaticsProcessor


def _vamos() -> BaseFornecedorProcessor:
    return BaseFornecedorProcessor(
        nome_fornecedor="VAMOS",
        pasta_fornecedor="VAMOS",
        nome_excel_fornecedor="VAMOS LOCAÇÃO",
        extrator=ColunaSimplesExtractor(tipo_gasto="Locação de veículos", coluna_valor="total"),
    )


def _tks() -> BaseFornecedorProcessor:
    return BaseFornecedorProcessor(
        nome_fornecedor="TKS",
        pasta_fornecedor="TKS",
        nome_excel_fornecedor="TKS",
        extrator=ColunaSimplesExtractor(tipo_gasto="Locação de veículos", coluna_valor="valor"),
    )


def _lm() -> BaseFornecedorProcessor:
    return BaseFornecedorProcessor(
        nome_fornecedor="LM",
        pasta_fornecedor="LM",
        nome_excel_fornecedor="LM LOCAÇÃO SUPERVISOR E VENDEDOR",
        extrator=ColunaSimplesExtractor(tipo_gasto="Locação de veículos", coluna_valor="total fatura"),
    )


def _ald() -> BaseFornecedorProcessor:
    return BaseFornecedorProcessor(
        nome_fornecedor="ALD",
        pasta_fornecedor="ALD",
        nome_excel_fornecedor="ALD LOCAÇÃO",
        extrator=ColunaSimplesExtractor(tipo_gasto="Locação de veículos", coluna_valor="total", busca_flexivel=True),
    )


def _relemar_md() -> BaseFornecedorProcessor:
    return BaseFornecedorProcessor(
        nome_fornecedor="RELEMAR MD",
        pasta_fornecedor="RELEMAR",
        nome_excel_fornecedor="RELEMAR GESTÃO MD",
        extrator=ColunaSimplesExtractor(
            tipo_gasto="Serviços administrativos (externos)", coluna_valor="valor",
            limpar_placas=True, usar_historico=False,
        ),
    )


def _geomatics_cameras() -> BaseFornecedorProcessor:
    return GeomaticsProcessor(
        nome_fornecedor="GEOMATICS CAMERAS",
        pasta_fornecedor="CAMERAS",
        nome_excel_fornecedor="GEOMATICS CAMERAS",
        extrator=CobrancaRelacaoExtractor(tipo_gasto="Serviços administrativos (externos)"),
    )


def _geomatics_cameras_sensores() -> BaseFornecedorProcessor:
    return GeomaticsProcessor(
        nome_fornecedor="GEOMATICS CAMERAS E SENSORES",
        pasta_fornecedor="CAMERAS E SENSORES",
        nome_excel_fornecedor="GEOMATICS CAMERAS E SENSORES",
        extrator=CobrancaRelacaoExtractor(
            tipo_gasto="Serviços administrativos (externos)", aceitar_alias_placas=True
        ),
    )


def _geomatics_rastreadores() -> BaseFornecedorProcessor:
    return GeomaticsProcessor(
        nome_fornecedor="GEOMATICS RASTREADORES DE CARGA",
        pasta_fornecedor="RASTREADORES DE CARGA",
        nome_excel_fornecedor="GEOMATICS RASTREADOR DE CARGA",
        extrator=CobrancaRelacaoExtractor(
            tipo_gasto="Serviços administrativos (externos)", aceitar_alias_placas=True
        ),
    )


def _relemar_ativos() -> BaseFornecedorProcessor:
    return BaseFornecedorProcessor(
        nome_fornecedor="RELEMAR ATIVOS",
        pasta_fornecedor="RELEMAR",
        nome_excel_fornecedor="RELEMAR GESTÃO ATIVOS",
        extrator=ListaAtivosExtractor(tipo_gasto="Serviços administrativos (externos)"),
    )


def _relemar_despachante_multas() -> BaseFornecedorProcessor:
    return BaseFornecedorProcessor(
        nome_fornecedor="RELEMAR DESPACHANTE E MULTAS",
        pasta_fornecedor="RELEMAR",
        nome_excel_fornecedor="RELEMAR DESPACHANTE E MULTAS",
        extrator=MultasTaxaExtractor(
            tipo_gasto_multa="Multas estaduais por penalidades",
            tipo_gasto_taxa="Serviços administrativos (externos)",
        ),
    )


def _unidas() -> BaseFornecedorProcessor:
    return BaseFornecedorProcessor(
        nome_fornecedor="UNIDAS",
        pasta_fornecedor="UNIDAS",
        nome_excel_fornecedor="UNIDAS LEVES E PESADOS",
        extrator=DescricaoMapeadaExtractor(),
    )


def _macari() -> BaseFornecedorProcessor:
    return BaseFornecedorProcessor(
        nome_fornecedor="MACARI",
        pasta_fornecedor="MACARI",
        nome_excel_fornecedor="MACARI - 08.2026",
        extrator=MultiSheetContaExtractor(
            tipo_gasto_por_planilha={
                "nota fiscal": "Serviços administrativos (externos)",
                "reembolso": "Licenças e placas",
            }
        ),
    )


def _edenred() -> BaseFornecedorProcessor:
    return BaseFornecedorProcessor(
        nome_fornecedor="EDENRED PEDAGIO",
        pasta_fornecedor="EDENRED",
        nome_excel_fornecedor="EDENRED PEDAGIO",
        extrator=TipoTransacaoExtractor(),
    )


def _ticket_combustivel() -> BaseFornecedorProcessor:
    return BaseFornecedorProcessor(
        nome_fornecedor="TICKET COMBUSTIVEL",
        pasta_fornecedor="TICKET",
        nome_excel_fornecedor="TICKET COMBUSTIVEL",
        extrator=TicketCombustivelExtractor(),
    )


def _ticket_manutencao() -> BaseFornecedorProcessor:
    return BaseFornecedorProcessor(
        nome_fornecedor="TICKET MANUTENÇÃO",
        pasta_fornecedor="TICKET",
        nome_excel_fornecedor="TICKET MANUTENÇÃO",
        extrator=TicketManutencaoExtractor(),
    )


class FornecedorProcessorFactory:
    """Registry: nome do fornecedor -> função que constrói o processor pronto."""

    _REGISTRO: Dict[str, Callable[[], BaseFornecedorProcessor]] = {
        "VAMOS": _vamos,
        "TKS": _tks,
        "LM": _lm,
        "ALD": _ald,
        "RELEMAR_MD": _relemar_md,
        "GEOMATICS_CAMERAS": _geomatics_cameras,
        "GEOMATICS_CAMERAS_SENSORES": _geomatics_cameras_sensores,
        "GEOMATICS_RASTREADORES": _geomatics_rastreadores,
        "RELEMAR_ATIVOS": _relemar_ativos,
        "RELEMAR_DESPACHANTE_MULTAS": _relemar_despachante_multas,
        "UNIDAS": _unidas,
        "MACARI": _macari,
        "EDENRED": _edenred,
        "TICKET_COMBUSTIVEL": _ticket_combustivel,
        "TICKET_MANUTENCAO": _ticket_manutencao,
    }

    @classmethod
    def criar(cls, nome_fornecedor: str) -> BaseFornecedorProcessor:
        construtor = cls._REGISTRO.get(nome_fornecedor)
        if construtor is None:
            disponiveis = ", ".join(sorted(cls._REGISTRO))
            raise ValueError(f"Fornecedor '{nome_fornecedor}' não registrado. Disponíveis: {disponiveis}")
        return construtor()

    @classmethod
    def fornecedores_disponiveis(cls) -> List[str]:
        return sorted(cls._REGISTRO)

    @classmethod
    def registrar(cls, nome_fornecedor: str, construtor: Callable[[], BaseFornecedorProcessor]) -> None:
        """Permite registrar um NOVO fornecedor sem tocar nesse arquivo em outros
        lugares -- só chame isso a partir do seu próprio script/módulo."""
        cls._REGISTRO[nome_fornecedor] = construtor
