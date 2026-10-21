"""ColunaSimplesExtractor: cobre os fornecedores cujo Excel é "uma linha por
lançamento", com uma coluna de PLACA e uma coluna de VALOR, variando só:

    - o nome da coluna de valor ("TOTAL", "VALOR", "TOTAL FATURA", ...)
    - se a busca do cabeçalho é simples (só a aba ativa, 1 linha) ou
      "flexível" (todas as abas, cabeçalho pode ocupar várias linhas -> ALD)
    - se a placa precisa ter o hífen removido (RELEMAR MD)
    - se usa o fallback de designação histórica (RELEMAR MD é o único que NÃO usa)

    Fornecedor  -> coluna_valor      | busca_flexivel | limpar_placas | usar_historico
    VAMOS       -> "total"           | não            | não           | sim
    TKS         -> "valor"           | não            | não           | sim
    LM          -> "total fatura"    | não            | não           | sim
    ALD         -> "total"           | sim            | não           | sim
    RELEMAR MD  -> "valor"           | não            | sim           | não
"""
from __future__ import annotations

from pathlib import Path
from typing import List

from openpyxl import load_workbook

from .. import utils
from ..config import conta_por_tipo_gasto
from ..modelos import Lancamento
from .base import ExtratorBase


class ColunaSimplesExtractor(ExtratorBase):
    def __init__(
        self,
        tipo_gasto: str,
        coluna_valor: str,
        coluna_placa: str = "placa",
        busca_flexivel: bool = False,
        limpar_placas: bool = False,
        usar_historico: bool = True,
    ):
        self.tipo_gasto = tipo_gasto
        self.coluna_valor = utils.normalizar_cabecalho(coluna_valor)
        self.coluna_placa = utils.normalizar_cabecalho(coluna_placa)
        self.busca_flexivel = busca_flexivel
        self.limpar_placas = limpar_placas
        self.usar_historico = usar_historico

    def extrair(self, caminho_excel_fornecedor: Path) -> List[Lancamento]:
        conta = conta_por_tipo_gasto(self.tipo_gasto)
        colunas_obrigatorias = {self.coluna_placa, self.coluna_valor}

        wb = load_workbook(caminho_excel_fornecedor, read_only=True, data_only=True)
        try:
            if self.busca_flexivel:
                planilha, linha_cabecalho, idx = utils.localizar_planilha_com_cabecalho(
                    wb, colunas_obrigatorias
                )
            else:
                planilha = wb.active
                cabecalhos, linha_cabecalho = utils.encontrar_linha_cabecalho(planilha, colunas_obrigatorias)
                if cabecalhos is None:
                    raise ValueError(
                        f"As colunas '{self.coluna_placa}' e '{self.coluna_valor}' não foram "
                        f"encontradas em: {caminho_excel_fornecedor.name}"
                    )
                idx = utils.indices_colunas(cabecalhos)

            lancamentos = []
            for linha in planilha.iter_rows(min_row=linha_cabecalho + 1, values_only=True):
                placa = linha[idx[self.coluna_placa]] if idx[self.coluna_placa] < len(linha) else None
                if placa is None:
                    continue
                if self.limpar_placas:
                    placa = utils.limpar_placa(placa)
                valor = linha[idx[self.coluna_valor]] if idx[self.coluna_valor] < len(linha) else None
                lancamentos.append(Lancamento(placa=placa, valor=valor, conta=conta))
            return lancamentos
        finally:
            wb.close()