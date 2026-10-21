"""Modelo de dados comum a todos os fornecedores: uma linha da planilha CONSO."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, Optional


@dataclass
class Lancamento:
    """Uma linha do Excel 'CONSO ...' final.

    `conta` já vem resolvida pelo extrator (ele conhece o tipo de gasto certo
    para aquela linha). `an`, `cco`, `cc` e `tipo` começam vazios e são
    preenchidos depois pelo DesignacaoService (etapa de enriquecimento).
    """
    placa: str
    valor: Optional[float]
    conta: Optional[str]
    servico: Optional[str] = None  # só alguns fornecedores têm coluna SERVIÇO (UNIDAS, EDENRED)
    an: Optional[str] = None
    cco: Optional[str] = None
    cc: Optional[str] = None
    tipo: Optional[str] = None
    extras: Dict[str, Any] = field(default_factory=dict)

    def enriquecer_com_designacao(self, an, cco, cc, tipo) -> None:
        self.an, self.cco, self.cc, self.tipo = an, cco, cc, tipo

    def possui_designacao(self) -> bool:
        return self.an is not None or self.cco is not None or self.cc is not None or self.tipo is not None