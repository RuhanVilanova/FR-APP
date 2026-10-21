"""BaseFornecedorProcessor: o "Template Method" que era duplicado (com pequenas
variações) em todos os 14 scripts:

    1. localizar o(s) arquivo(s) do fornecedor
    2. carregar a designação (normal + histórico) daquele fornecedor/mês
    3. extrair os lançamentos (delegado ao Extrator -- Strategy)
    4. enriquecer cada lançamento com AN/CCO/CC/TIPO (registrando quem ficou
       sem designação)
    5. pós-processar (delegado ao Extrator; só o EDENRED faz algo aqui)
    6. gravar o Excel "CONSO ..." final

Isso NUNCA muda de fornecedor para fornecedor -- o que muda é só o Extrator
que a classe recebe no construtor (Strategy) e, opcionalmente, o nome do
arquivo de saída.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

from . import config, excel_writer
from .designacao import DesignacaoService
from .extratores.base import ExtratorBase
from .modelos import Lancamento


@dataclass
class ResultadoProcessamento:
    fornecedor: str
    arquivos_gerados: List[Path] = field(default_factory=list)
    placas_sem_designacao: List[dict] = field(default_factory=list)
    tipos_nao_mapeados: List[dict] = field(default_factory=list)


class BaseFornecedorProcessor:
    def __init__(
        self,
        nome_fornecedor: str,
        pasta_fornecedor: str,
        nome_excel_fornecedor: str,
        extrator: ExtratorBase,
        permitir_varios_arquivos: bool = False,
        pasta_base: Optional[Path] = None,
    ):
        self.nome_fornecedor = nome_fornecedor
        self.pasta_fornecedor = pasta_fornecedor
        self.nome_excel_fornecedor = nome_excel_fornecedor
        self.extrator = extrator
        self.permitir_varios_arquivos = permitir_varios_arquivos
        self.pasta_base = pasta_base or config.CAMINHO_CONTROLE_DESPESAS_FR

        self.designacao_service = DesignacaoService(config.CAMINHO_DESIGNACAO_HISTORICA)

    # -- hooks que uma subclasse raramente precisa sobrescrever ---------------
    def nome_arquivo_saida(self, caminho_excel_fornecedor: Path) -> Path:
        """Nome do CONSO gerado para CADA arquivo de entrada processado.
        Por padrão usa `nome_excel_fornecedor`; GEOMATICS sobrescreve para usar
        o nome de cada arquivo de entrada (já que processa vários arquivos)."""
        return caminho_excel_fornecedor.parent / f"CONSO {self.nome_excel_fornecedor}.xlsx"

    # -- fluxo fixo (Template Method) -----------------------------------------
    def processar(self) -> ResultadoProcessamento:
        from . import utils  # import local para evitar ciclo em tempo de import

        resultado = ResultadoProcessamento(fornecedor=self.nome_fornecedor)

        caminhos = utils.pegar_excel_fornecedor(
            self.pasta_base, self.pasta_fornecedor, self.nome_excel_fornecedor,
            permitir_varios=self.permitir_varios_arquivos,
        )
        if not self.permitir_varios_arquivos:
            caminhos = [caminhos]

        for caminho_excel in caminhos:
            self.designacao_service.carregar_designacao_do_fornecedor(caminho_excel)

            lancamentos = self.extrator.extrair(caminho_excel)
            self._enriquecer_com_designacao(lancamentos, resultado)
            lancamentos = self.extrator.pos_processar(lancamentos)

            if hasattr(self.extrator, "tipos_nao_mapeados"):
                resultado.tipos_nao_mapeados.extend(self.extrator.tipos_nao_mapeados)

            caminho_saida = self.nome_arquivo_saida(caminho_excel)
            excel_writer.escrever_conso(
                caminho_saida,
                lancamentos,
                incluir_servico=self.extrator.incluir_servico,
                colunas_pos_valor=self.extrator.colunas_pos_valor,
                colunas_pos_validacao=self.extrator.colunas_pos_validacao,
            )
            excel_writer.inserir_conso_no_fornecedor(caminho_saida, caminho_excel)
            caminho_saida.unlink(missing_ok=True)
            resultado.arquivos_gerados.append(caminho_excel)

            if hasattr(self.extrator, "arquivos_extras"):
                resultado.arquivos_gerados.extend(self.extrator.arquivos_extras)

        return resultado

    def _enriquecer_com_designacao(self, lancamentos: List[Lancamento], resultado: ResultadoProcessamento) -> None:
        for lanc in lancamentos:
            an, cco, cc, tipo = self.designacao_service.buscar(lanc.placa, usar_historico=self.extrator.usar_historico)
            lanc.enriquecer_com_designacao(an, cco, cc, tipo)
            if not lanc.possui_designacao():
                resultado.placas_sem_designacao.append({"placa": lanc.placa, "fornecedor": self.nome_fornecedor})


class GeomaticsProcessor(BaseFornecedorProcessor):
    """Única diferença real: processa VÁRIOS arquivos de entrada (um CONSO por
    arquivo, nomeado a partir do próprio arquivo de entrada, não do fornecedor)."""

    def __init__(self, *args, **kwargs):
        kwargs.setdefault("permitir_varios_arquivos", True)
        super().__init__(*args, **kwargs)

    def nome_arquivo_saida(self, caminho_excel_fornecedor: Path) -> Path:
        return caminho_excel_fornecedor.parent / f"CONSO_{caminho_excel_fornecedor.stem}.xlsx"
