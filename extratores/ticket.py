"""Extrator do fornecedor TICKET COMBUSTIVEL.

Lê uma planilha com colunas PLACA, VALOR e SERVICO. A conta contábil não é
fixa: ela é resolvida dinamicamente a partir do arquivo `Contas_ticket.json`,
que mapeia cada serviço para o ID de uma conta cadastrada em `config.GASTO_CONTAS`.

Quando um serviço ainda não tem conta definida (ou vem vazio), o backend expõe
essa pendência para o HTML, que pergunta ao usuário qual conta deve ser usada e
persiste a resposta no JSON para as próximas execuções.
"""
from __future__ import annotations

import json
import math
import re
from pathlib import Path
from typing import Dict, List, Optional

from openpyxl import load_workbook

from .. import config, utils
from ..modelos import Lancamento
from .base import ExtratorBase

_CHAVE_SEM_SERVICO = "__sem_servico__"
_ROTULO_SEM_SERVICO = "Sem serviço informado"


class TicketCombustivelExtractor(ExtratorBase):
    incluir_servico = True

    def __init__(self, caminho_contas_ticket: Optional[Path] = None):
        self.caminho_contas_ticket = caminho_contas_ticket or config.CAMINHO_CONTAS_TICKET
        self._ultimo_caminho_excel: Optional[Path] = None
        self.arquivos_extras: List[Path] = []

    def _garantir_arquivo_json(self) -> None:
        self.caminho_contas_ticket.parent.mkdir(parents=True, exist_ok=True)
        if self.caminho_contas_ticket.exists():
            return
        self.caminho_contas_ticket.write_text(
            json.dumps({"servicos": {}}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _rotulo_servico(self, servico) -> str:
        texto = "" if servico is None else str(servico).strip()
        return texto or _ROTULO_SEM_SERVICO

    def _valor_servico_para_json(self, servico) -> str:
        return "" if servico is None else str(servico).strip()

    def _chave_servico(self, servico) -> str:
        valor = self._valor_servico_para_json(servico)
        if not valor:
            return _CHAVE_SEM_SERVICO
        return utils.normalizar_texto(valor)

    def _carregar_mapeamentos(self) -> Dict[str, dict]:
        self._garantir_arquivo_json()
        try:
            dados = json.loads(self.caminho_contas_ticket.read_text(encoding="utf-8") or "{}")
        except json.JSONDecodeError as exc:
            raise ValueError(
                f"JSON inválido em {self.caminho_contas_ticket.name}: {exc.msg}"
            ) from exc

        servicos = dados.get("servicos", {})
        if not isinstance(servicos, dict):
            raise ValueError(
                f"Estrutura inválida em {self.caminho_contas_ticket.name}: esperado objeto 'servicos'"
            )
        return servicos

    def _salvar_mapeamentos(self, servicos: Dict[str, dict]) -> None:
        payload = {"servicos": servicos}
        self.caminho_contas_ticket.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def salvar_mapeamento(self, servico, conta_id: int | str) -> dict:
        conta_id_int = int(conta_id)
        config.tipo_gasto_por_id(conta_id_int)  # valida o ID

        servicos = self._carregar_mapeamentos()
        chave = self._chave_servico(servico)
        servicos[chave] = {
            "servico": self._valor_servico_para_json(servico),
            "conta_id": conta_id_int,
        }
        self._salvar_mapeamentos(servicos)
        return servicos[chave]

    def _localizar_planilha_ticket(self, workbook):
        return utils.localizar_planilha_com_cabecalho(workbook, {"placa", "valor", "servico"})

    def listar_servicos_pendentes(self, caminho_excel_fornecedor: Path) -> List[dict]:
        servicos_mapeados = self._carregar_mapeamentos()

        wb = load_workbook(caminho_excel_fornecedor, read_only=True, data_only=True)
        pendentes: Dict[str, dict] = {}
        try:
            planilha, linha_cabecalho, idx = self._localizar_planilha_ticket(wb)
            for linha in planilha.iter_rows(min_row=linha_cabecalho + 1, values_only=True):
                placa = linha[idx["placa"]] if idx["placa"] < len(linha) else None
                if placa is None:
                    continue

                servico = linha[idx["servico"]] if idx["servico"] < len(linha) else None
                chave = self._chave_servico(servico)
                if chave in servicos_mapeados or chave in pendentes:
                    continue

                pendentes[chave] = {
                    "chave": chave,
                    "servico": self._valor_servico_para_json(servico),
                    "servico_exibicao": self._rotulo_servico(servico),
                }
        finally:
            wb.close()

        return list(pendentes.values())

    def extrair(self, caminho_excel_fornecedor: Path) -> List[Lancamento]:
        self._ultimo_caminho_excel = caminho_excel_fornecedor
        servicos_mapeados = self._carregar_mapeamentos()
        pendentes: Dict[str, str] = {}

        wb = load_workbook(caminho_excel_fornecedor, read_only=True, data_only=True)
        lancamentos: List[Lancamento] = []
        try:
            planilha, linha_cabecalho, idx = self._localizar_planilha_ticket(wb)
            for linha in planilha.iter_rows(min_row=linha_cabecalho + 1, values_only=True):
                placa = linha[idx["placa"]] if idx["placa"] < len(linha) else None
                if placa is None:
                    continue

                valor = linha[idx["valor"]] if idx["valor"] < len(linha) else None
                servico = linha[idx["servico"]] if idx["servico"] < len(linha) else None
                chave = self._chave_servico(servico)
                mapeamento = servicos_mapeados.get(chave)

                if mapeamento is None:
                    pendentes[chave] = self._rotulo_servico(servico)
                    continue

                conta = config.conta_por_id(mapeamento["conta_id"])
                lancamentos.append(
                    Lancamento(
                        placa=placa,
                        valor=valor,
                        conta=conta,
                        servico=self._rotulo_servico(servico),
                    )
                )
        finally:
            wb.close()

        if pendentes:
            servicos = ", ".join(sorted(pendentes.values()))
            raise ValueError(
                "Serviços do TICKET COMBUSTIVEL sem conta cadastrada: "
                f"{servicos}. Cadastre-os pelo HTML antes de processar."
            )

        return lancamentos

    # ------------------------------------------------------------------
    # Honorarios: leitura do PDF e distribuicao da diferenca
    # ------------------------------------------------------------------

    def _localizar_pdf_por_tipo(self, pasta: Path, tipo: str) -> Optional[Path]:
        """Acha o PDF cujo nome contenha 'TICKET' e `tipo` (ex: COMBUSTIVEL, RECOLHA)."""
        tipo_upper = tipo.upper()
        for arq in sorted(pasta.glob("*.pdf")):
            nome_upper = arq.stem.upper()
            if "TICKET" in nome_upper and tipo_upper in nome_upper:
                return arq
        return None

    def _extrair_valor_total_pdf(self, caminho_pdf: Path) -> Optional[float]:
        """Extrai o VALOR TOTAL DA NOTA do PDF via pdfplumber."""
        try:
            import pdfplumber
        except ImportError as exc:
            raise ImportError(
                "Instale o pdfplumber para leitura do PDF do TICKET: "
                "pip install pdfplumber"
            ) from exc

        texto = ""
        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                texto += (pagina.extract_text() or "") + "\n"

        padrao = r'VALOR\s+TOTAL\s+(?:DA\s+)?NOTA[^\d]*(\d[\d.]*,\d{2})'
        match = re.search(padrao, texto, re.IGNORECASE)
        if not match:
            return None

        valor_str = match.group(1).replace(".", "").replace(",", ".")
        return float(valor_str)

    # ------------------------------------------------------------------
    # Distribuicao generica de valor entre placas
    # ------------------------------------------------------------------

    def _placas_unicas(self, lancamentos: List[Lancamento]) -> List[Lancamento]:
        """Retorna uma referencia por placa unica (primeira ocorrencia)."""
        refs: List[Lancamento] = []
        vistas: set[str] = set()
        for l in lancamentos:
            placa_str = str(l.placa).strip().upper()
            if placa_str not in vistas:
                vistas.add(placa_str)
                refs.append(l)
        return refs

    def _distribuir_valor(
        self,
        total: float,
        placas_ref: List[Lancamento],
        servico: str,
        conta: str,
    ) -> List[Lancamento]:
        """Distribui `total` igualmente entre as placas, arredondando pra cima
        (2 casas) e tirando o excesso da ultima placa."""
        n = len(placas_ref)
        valor_por_placa = total / n
        valor_arredondado = math.ceil(valor_por_placa * 100) / 100

        total_arredondado = round(valor_arredondado * n, 2)
        excesso = round(total_arredondado - total, 2)

        novos: List[Lancamento] = []
        for i, ref in enumerate(placas_ref):
            valor = valor_arredondado
            if i == n - 1:
                valor = round(valor - excesso, 2)
            novos.append(Lancamento(
                placa=ref.placa,
                valor=valor,
                conta=conta,
                servico=servico,
                an=ref.an,
                cco=ref.cco,
                cc=ref.cc,
                tipo=ref.tipo,
            ))
        return novos

    # ------------------------------------------------------------------
    # Honorarios (linhas extras no CONSO principal)
    # ------------------------------------------------------------------

    def _calcular_honorarios(
        self, lancamentos: List[Lancamento], placas_ref: List[Lancamento], pasta: Path,
    ) -> List[Lancamento]:
        """Retorna lancamentos de HONORARIOS se houver diferenca entre o PDF
        COMBUSTIVEL e a soma do Excel."""
        caminho_pdf = self._localizar_pdf_por_tipo(pasta, "COMBUSTIVEL")
        if caminho_pdf is None:
            print("  [TICKET] PDF COMBUSTIVEL nao encontrado, pulando honorarios.")
            return []

        valor_pdf = self._extrair_valor_total_pdf(caminho_pdf)
        if valor_pdf is None:
            print(f"  [TICKET] VALOR TOTAL DA NOTA nao encontrado em {caminho_pdf.name}")
            return []

        valor_conso = sum(l.valor for l in lancamentos if l.valor is not None)
        diferenca = round(valor_pdf - valor_conso, 2)

        if diferenca <= 0:
            print(f"  [TICKET] Sem diferenca para honorarios (PDF={valor_pdf:.2f}, CONSO={valor_conso:.2f})")
            return []

        conta = config.conta_por_tipo_gasto("Servi\u00e7os administrativos (externos)")
        novos = self._distribuir_valor(diferenca, placas_ref, "HONORARIOS", conta)

        total_hon = sum(l.valor for l in novos)
        print(f"  [TICKET] Honorarios: PDF={valor_pdf:.2f} | CONSO={valor_conso:.2f} | "
              f"Diferenca={diferenca:.2f} | {len(novos)} placas | Total={total_hon:.2f}")
        return novos

    # ------------------------------------------------------------------
    # CONSO RECOLHA (arquivo Excel separado)
    # ------------------------------------------------------------------

    def _gerar_conso_recolha(
        self, placas_ref: List[Lancamento], pasta: Path,
    ) -> Optional[Path]:
        """Gera CONSO RECOLHA distribuindo o valor total do PDF RECOLHA."""
        caminho_pdf = self._localizar_pdf_por_tipo(pasta, "RECOLHA")
        if caminho_pdf is None:
            print("  [TICKET] PDF RECOLHA nao encontrado, pulando CONSO RECOLHA.")
            return None

        valor_pdf = self._extrair_valor_total_pdf(caminho_pdf)
        if valor_pdf is None:
            print(f"  [TICKET] VALOR TOTAL DA NOTA nao encontrado em {caminho_pdf.name}")
            return None

        conta = config.conta_por_tipo_gasto("Servi\u00e7os administrativos (externos)")
        lancamentos_recolha = self._distribuir_valor(valor_pdf, placas_ref, "RECOLHA", conta)

        from datetime import date
        from .. import excel_writer

        hoje = date.today()
        ano = hoje.year if hoje.month > 1 else hoje.year - 1
        nome = f"RECOLHA - {config.MES_ATUAL}.{ano}.xlsx"
        caminho_saida = pasta / nome

        excel_writer.escrever_conso(caminho_saida, lancamentos_recolha, incluir_servico=True)

        total = sum(l.valor for l in lancamentos_recolha)
        print(f"  [TICKET] Recolha: PDF={valor_pdf:.2f} | {len(lancamentos_recolha)} placas | "
              f"Total={total:.2f} | -> {nome}")
        return caminho_saida

    # ------------------------------------------------------------------
    # pos_processar: honorarios + recolha
    # ------------------------------------------------------------------

    def pos_processar(self, lancamentos: List[Lancamento]) -> List[Lancamento]:
        """Adiciona linhas de HONORARIOS ao CONSO principal e gera o CONSO RECOLHA."""
        self.arquivos_extras.clear()

        if not self._ultimo_caminho_excel:
            return lancamentos

        pasta = self._ultimo_caminho_excel.parent
        placas_ref = self._placas_unicas(lancamentos)

        if not placas_ref:
            print("  [TICKET] Nenhuma placa para distribuir.")
            return lancamentos

        # 1. Honorarios (linhas extras no CONSO TICKET COMBUSTIVEL)
        honorarios = self._calcular_honorarios(lancamentos, placas_ref, pasta)

        # 2. CONSO RECOLHA (arquivo separado)
        caminho_recolha = self._gerar_conso_recolha(placas_ref, pasta)
        if caminho_recolha is not None:
            self.arquivos_extras.append(caminho_recolha)

        return lancamentos + honorarios


# ======================================================================
# TICKET MANUTENÇÃO
# ======================================================================

class TicketManutencaoExtractor(ExtratorBase):
    """Extrator do fornecedor TICKET MANUTENÇÃO.

    Lê uma planilha com colunas Placa, R$ Aprovado (valor principal),
    Tipo Manutenção, Grupo Peça, Peça, Tipo MO (extras pós-valor) e
    R$ Pç. Aprovado, R$ MO Aprovado (extras pós-validação).

    A conta contábil é resolvida por uma cascata de 9 passos sequenciais
    (``_resolver_contas``): cada linha recebe conta no primeiro passo que
    casa e NÃO é reavaliada nos passos seguintes.
    """

    incluir_servico = False

    def __init__(self):
        self._ultimo_caminho_excel: Optional[Path] = None
        self.arquivos_extras: List[Path] = []

    colunas_pos_valor = (
        "TIPO MANUTENÇÃO",
        "GRUPO PEÇA",
        "PEÇA",
        "TIPO MO",
    )
    colunas_pos_validacao = (
        "R$ PÇ. APROVADO",
        "R$ MO APROVADO",
    )

    # Colunas obrigatórias do Excel do fornecedor (serão normalizadas)
    _COLUNAS_OBRIGATORIAS = {
        "placa", "R$ Aprovado",
        "Tipo Manutenção", "Grupo Peça", "Peça", "Tipo MO",
        "R$ Pç. Aprovado", "R$ MO Aprovado",
    }

    # Mapeamento: chave normalizada do cabeçalho → nome da coluna em extras
    _MAPA_EXTRAS = {
        "tipomanutencao": "TIPO MANUTENÇÃO",
        "grupopeca":      "GRUPO PEÇA",
        "peca":           "PEÇA",
        "tipomo":         "TIPO MO",
        "rpcaprovado":    "R$ PÇ. APROVADO",
        "rmoaprovado":    "R$ MO APROVADO",
    }

    def extrair(self, caminho_excel_fornecedor: Path) -> List[Lancamento]:
        self._ultimo_caminho_excel = caminho_excel_fornecedor
        wb = load_workbook(caminho_excel_fornecedor, read_only=True, data_only=True)
        lancamentos: List[Lancamento] = []
        try:
            planilha, linha_cabecalho, idx = utils.localizar_planilha_com_cabecalho(
                wb, self._COLUNAS_OBRIGATORIAS,
            )
            for linha in planilha.iter_rows(min_row=linha_cabecalho + 1, values_only=True):
                placa = linha[idx["placa"]] if idx["placa"] < len(linha) else None
                if placa is None:
                    continue

                valor = linha[idx["raprovado"]] if idx["raprovado"] < len(linha) else None

                extras: Dict[str, object] = {}
                for chave_norm, nome_saida in self._MAPA_EXTRAS.items():
                    col_idx = idx.get(chave_norm)
                    if col_idx is not None and col_idx < len(linha):
                        extras[nome_saida] = linha[col_idx]
                    else:
                        extras[nome_saida] = None

                lancamentos.append(Lancamento(
                    placa=placa,
                    valor=valor,
                    conta=None,
                    extras=extras,
                ))
        finally:
            wb.close()

        self._resolver_contas(lancamentos)

        # Remover linhas com valor zero (não geram lançamento contábil)
        lancamentos = [l for l in lancamentos if l.valor and l.valor != 0]

        return lancamentos

    # ------------------------------------------------------------------
    #  Helpers de leitura de extras
    # ------------------------------------------------------------------

    @staticmethod
    def _extra_texto(lanc: Lancamento, chave: str) -> str:
        """Retorna o valor normalizado (sem acento, minúsculo) de um extra."""
        val = lanc.extras.get(chave)
        if val is None:
            return ""
        return utils.normalizar_texto(str(val))

    @staticmethod
    def _extra_numero(lanc: Lancamento, chave: str) -> float:
        """Retorna o valor numérico de um extra (0.0 se vazio/inválido)."""
        val = lanc.extras.get(chave)
        if val is None:
            return 0.0
        try:
            return float(val)
        except (TypeError, ValueError):
            return 0.0

    # ------------------------------------------------------------------
    #  Cascata de 9 passos — resolução de conta contábil
    # ------------------------------------------------------------------

    _PNEU_EXATOS = frozenset({"pneu", "pneu reformado", "rodizio de pneus"})

    def _resolver_contas(self, lancamentos: List[Lancamento]) -> None:
        """Atribui ``conta`` a cada lançamento seguindo 9 passos em cascata.

        Uma vez preenchida, a conta NÃO é sobrescrita por passos posteriores.
        Comparações de texto são case-insensitive e sem acentos (normalizar_texto).
        """
        # Pré-resolve as contas uma única vez
        conta_3  = config.conta_por_id(3)   # Lavagem e lubrificação
        conta_6  = config.conta_por_id(6)   # Transferência de Veículos
        conta_7  = config.conta_por_id(7)   # Manutenção de frotas
        conta_9  = config.conta_por_id(9)   # Peças de veículos
        conta_10 = config.conta_por_id(10)  # Pneus
        conta_15 = config.conta_por_id(15)  # Óleos e lubrificantes

        pendentes_9: List[Lancamento] = []

        # === Passos 1–8 (passagem única) =================================
        for lanc in lancamentos:
            peca = self._extra_texto(lanc, "PEÇA")

            # Passo 1 — Peça contém "lavagem"
            if "lavagem" in peca:
                lanc.conta = conta_3
                continue

            # Passo 2 — Peça exatamente PNEU / PNEU REFORMADO / RODIZIO DE PNEUS
            if peca in self._PNEU_EXATOS:
                lanc.conta = conta_10
                continue

            # Passo 3 — Peça contém lubri / fluido / oleo
            if any(t in peca for t in ("lubri", "fluido", "oleo")):
                lanc.conta = conta_15
                continue

            # Passo 4 — Peça contém socorro / guincho
            if any(t in peca for t in ("socorro", "guincho")):
                lanc.conta = conta_6
                continue

            # Passo 5 — Peça contém preventiva
            if "preventiva" in peca:
                lanc.conta = conta_7
                continue

            # Passo 6 — Tipo MO contém fornecimento
            tipo_mo = self._extra_texto(lanc, "TIPO MO")
            if "fornecimento" in tipo_mo:
                lanc.conta = conta_9
                continue

            # Passos 7–9 — Tipo Manutenção PREVENTIVA / CORRETIVA
            tipo_manut = self._extra_texto(lanc, "TIPO MANUTENÇÃO")
            if tipo_manut not in ("preventiva", "corretiva"):
                continue  # nenhuma regra se aplica

            pc = self._extra_numero(lanc, "R$ PÇ. APROVADO")
            mo = self._extra_numero(lanc, "R$ MO APROVADO")

            # 7.1 — custo exclusivo MO
            if pc == 0 and mo > 0:
                lanc.conta = conta_7
            # 7.2 — custo exclusivo peça
            elif pc > 0 and mo == 0:
                lanc.conta = conta_9
            # Passo 8 — PREVENTIVA mista
            elif pc > 0 and mo > 0 and tipo_manut == "preventiva":
                lanc.conta = conta_7
            # Passo 9 candidato — CORRETIVA mista
            elif pc > 0 and mo > 0 and tipo_manut == "corretiva":
                pendentes_9.append(lanc)

        # === Passo 9 — Balanceamento CORRETIVA mista =====================
        if not pendentes_9:
            return

        # Somas acumuladas dos passos 1–8
        soma_7 = sum(l.valor or 0 for l in lancamentos if l.conta == conta_7)
        soma_9 = sum(l.valor or 0 for l in lancamentos if l.conta == conta_9)

        # Ordenar por valor decrescente melhora o equilíbrio (Largest
        # Differencing Method — itens grandes são alocados primeiro,
        # reduzindo a chance de desbalanceamento no final).
        pendentes_9.sort(key=lambda l: l.valor or 0, reverse=True)

        for lanc in pendentes_9:
            valor = lanc.valor or 0
            if soma_7 <= soma_9:
                lanc.conta = conta_7
                soma_7 += valor
            else:
                lanc.conta = conta_9
                soma_9 += valor

    # ------------------------------------------------------------------
    #  Honorários: PDF NFS-e → distribuição da diferença
    # ------------------------------------------------------------------

    def _localizar_pdf_manutencao(self, pasta: Path) -> Optional[Path]:
        """Acha o PDF cujo nome contenha 'TICKET' e 'MANUTENCAO' (sem acento)."""
        for arq in sorted(pasta.glob("*.pdf")):
            nome_norm = utils.normalizar_texto(arq.stem)
            if "ticket" in nome_norm and "manutencao" in nome_norm:
                return arq
        return None

    def _extrair_valor_liquido_nfse(self, caminho_pdf: Path) -> Optional[float]:
        """Extrai VALOR LÍQUIDO DA NFS-e + IBS/CBS do PDF via pdfplumber."""
        try:
            import pdfplumber
        except ImportError as exc:
            raise ImportError(
                "Instale o pdfplumber para leitura do PDF do TICKET: "
                "pip install pdfplumber"
            ) from exc

        texto = ""
        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                texto += (pagina.extract_text() or "") + "\n"

        # Âncora: "VALOR LÍQUIDO DA NFS-e + IBS/CBS"
        padrao = r'VALOR\s+L[\u00cdI]QUIDO\s+DA\s+NFS.?e\s*\+\s*IBS\s*/?\s*CBS'
        match = re.search(padrao, texto, re.IGNORECASE)
        if not match:
            return None

        # Próxima linha com valores monetários — o último corresponde
        # à última coluna do cabeçalho (VALOR LÍQUIDO + IBS/CBS).
        trecho = texto[match.end():]
        for linha in trecho.split("\n"):
            linha = linha.strip()
            if not linha:
                continue
            valores = re.findall(r'(\d[\d.]*,\d{2})', linha)
            if valores:
                valor_str = valores[-1]
                return float(valor_str.replace(".", "").replace(",", "."))

        return None

    @staticmethod
    def _placas_unicas(lancamentos: List[Lancamento]) -> List[Lancamento]:
        """Retorna uma referência por placa única (primeira ocorrência)."""
        refs: List[Lancamento] = []
        vistas: set[str] = set()
        for l in lancamentos:
            placa_str = str(l.placa).strip().upper()
            if placa_str not in vistas:
                vistas.add(placa_str)
                refs.append(l)
        return refs

    def _distribuir_honorarios(
        self,
        total: float,
        placas_ref: List[Lancamento],
        conta: str,
    ) -> List[Lancamento]:
        """Distribui ``total`` igualmente entre as placas (arredonda p/ cima,
        tira o excesso da última).  Extras ficam vazios exceto TIPO MANUTENÇÃO."""
        n = len(placas_ref)
        valor_por_placa = total / n
        valor_arredondado = math.ceil(valor_por_placa * 100) / 100

        total_arredondado = round(valor_arredondado * n, 2)
        excesso = round(total_arredondado - total, 2)

        extras_hon = {
            "TIPO MANUTEN\u00c7\u00c3O": "HONOR\u00c1RIOS",
            "GRUPO PE\u00c7A": None,
            "PE\u00c7A": None,
            "TIPO MO": None,
            "R$ P\u00c7. APROVADO": None,
            "R$ MO APROVADO": None,
        }

        novos: List[Lancamento] = []
        for i, ref in enumerate(placas_ref):
            valor = valor_arredondado
            if i == n - 1:
                valor = round(valor - excesso, 2)
            novos.append(Lancamento(
                placa=ref.placa,
                valor=valor,
                conta=conta,
                an=ref.an,
                cco=ref.cco,
                cc=ref.cc,
                tipo=ref.tipo,
                extras=dict(extras_hon),
            ))
        return novos

    def _calcular_honorarios(
        self, lancamentos: List[Lancamento], placas_ref: List[Lancamento], pasta: Path,
    ) -> List[Lancamento]:
        """Retorna lançamentos de HONORÁRIOS se houver diferença entre o PDF
        MANUTENÇÃO e a soma do Excel."""
        caminho_pdf = self._localizar_pdf_manutencao(pasta)
        if caminho_pdf is None:
            print("  [TICKET MANUTEN\u00c7\u00c3O] PDF MANUTEN\u00c7\u00c3O n\u00e3o encontrado, pulando honor\u00e1rios.")
            return []

        valor_pdf = self._extrair_valor_liquido_nfse(caminho_pdf)
        if valor_pdf is None:
            print(f"  [TICKET MANUTEN\u00c7\u00c3O] VALOR L\u00cdQUIDO DA NFS-e + IBS/CBS n\u00e3o encontrado em {caminho_pdf.name}")
            return []

        valor_conso = sum(l.valor or 0 for l in lancamentos)
        diferenca = round(valor_pdf - valor_conso, 2)

        if diferenca <= 0:
            print(f"  [TICKET MANUTEN\u00c7\u00c3O] Sem diferen\u00e7a para honor\u00e1rios "
                  f"(PDF={valor_pdf:.2f}, CONSO={valor_conso:.2f})")
            return []

        conta = config.conta_por_tipo_gasto("Servi\u00e7os administrativos (externos)")
        novos = self._distribuir_honorarios(diferenca, placas_ref, conta)

        total_hon = sum(l.valor for l in novos)
        print(f"  [TICKET MANUTEN\u00c7\u00c3O] Honor\u00e1rios: PDF={valor_pdf:.2f} | CONSO={valor_conso:.2f} | "
              f"Diferen\u00e7a={diferenca:.2f} | {len(novos)} placas | Total={total_hon:.2f}")
        return novos

    def pos_processar(self, lancamentos: List[Lancamento]) -> List[Lancamento]:
        """Adiciona linhas de HONORÁRIOS ao CONSO principal."""
        self.arquivos_extras.clear()

        if not self._ultimo_caminho_excel:
            return lancamentos

        pasta = self._ultimo_caminho_excel.parent
        placas_ref = self._placas_unicas(lancamentos)

        if not placas_ref:
            print("  [TICKET MANUTEN\u00c7\u00c3O] Nenhuma placa para distribuir.")
            return lancamentos

        honorarios = self._calcular_honorarios(lancamentos, placas_ref, pasta)
        return lancamentos + honorarios
