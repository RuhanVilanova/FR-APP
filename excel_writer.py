"""Escreve o Excel 'CONSO ...' final a partir de uma lista de Lancamento.

Duas variações existiam no original:
  - sem coluna SERVIÇO: PLACA, VALOR, AN, CCO, CC, CONTA, TIPO, FLEX, VALIDAÇÃO
    (FLEX na coluna H, fórmula "089."&C&"."&D&"."&F&...)
  - com coluna SERVIÇO (UNIDAS, EDENRED): PLACA, VALOR, SERVIÇO, AN, CCO, CC, CONTA,
    TIPO, FLEX, VALIDAÇÃO (FLEX na coluna I, fórmula "089."&D&"."&E&"."&G&...)

Em vez de duplicar a lógica, escrevemos as colunas dinamicamente e calculamos a
fórmula do FLEX a partir das LETRAS reais das colunas AN/CCO/CONTA -- assim o
mesmo writer serve para os dois layouts.
"""
from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable, Sequence

from copy import copy

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, numbers
from openpyxl.utils import get_column_letter

from .modelos import Lancamento

_SUFIXO_FLEX = ".000.00000000.00000.00000"


def _calcular_flex(an, cco, conta) -> str:
    """Monta a conta FLEX concatenando AN, CCO e CONTA com prefixo 089."""
    return f"089.{an or ''}.{cco or ''}.{conta or ''}{_SUFIXO_FLEX}"


def escrever_conso(
    caminho_saida: Path,
    lancamentos: Iterable[Lancamento],
    incluir_servico: bool,
    colunas_pos_valor: Sequence[str] = (),
    colunas_pos_validacao: Sequence[str] = (),
) -> Path:
    """Gera o Excel CONSO com layout dinâmico.

    Ordem das colunas:
        PLACA | VALOR | [SERVIÇO] | [*colunas_pos_valor] | AN | CCO | CC |
        CONTA | TIPO | FLEX | VALIDAÇÃO | [*colunas_pos_validacao]

    Valores das colunas extras são lidos de ``lanc.extras[nome_coluna]``.
    """
    # -- montar lista de colunas ------------------------------------------------
    colunas = ["PLACA", "VALOR"]
    if incluir_servico:
        colunas.append("SERVIÇO")
    colunas.extend(colunas_pos_valor)
    colunas += ["AN", "CCO", "CC", "CONTA", "TIPO", "FLEX", "VALIDAÇÃO"]
    colunas.extend(colunas_pos_validacao)

    wb = Workbook()
    ws = wb.active
    ws.title = "CONSOLIDADO"
    ws.append(colunas)

    for lanc in lancamentos:
        flex_valor = _calcular_flex(lanc.an, lanc.cco, lanc.conta)
        linha = [lanc.placa, lanc.valor]
        if incluir_servico:
            linha.append(lanc.servico)
        for col in colunas_pos_valor:
            linha.append(lanc.extras.get(col))
        linha += [lanc.an, lanc.cco, lanc.cc, lanc.conta, lanc.tipo, flex_valor, None]
        for col in colunas_pos_validacao:
            linha.append(lanc.extras.get(col))
        ws.append(linha)

    # --- Formatação do cabeçalho: fundo azul claro + fonte branca em negrito ---
    azul_claro = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
    fonte_cabecalho = Font(color="FFFFFF", bold=True)
    for celula in ws[1]:
        celula.fill = azul_claro
        celula.font = fonte_cabecalho

    # --- Formato R$ na coluna VALOR e em colunas monetárias extras ----------
    fmt_brl = 'R$ #,##0.00'
    colunas_monetarias = {"VALOR"}
    for col in (*colunas_pos_valor, *colunas_pos_validacao):
        if col.upper().startswith("R$"):
            colunas_monetarias.add(col)

    for nome_col in colunas_monetarias:
        col_idx = colunas.index(nome_col) + 1
        for linha in range(2, ws.max_row + 1):
            ws.cell(linha, col_idx).number_format = fmt_brl

    # --- Auto-ajuste de largura das colunas ---
    for col_idx, nome_col in enumerate(colunas, start=1):
        largura_max = len(str(nome_col))
        for linha in range(2, ws.max_row + 1):
            valor = ws.cell(linha, col_idx).value
            if valor is not None:
                largura_max = max(largura_max, len(str(valor)))
        ws.column_dimensions[get_column_letter(col_idx)].width = largura_max + 4

    # --- Bloco OC (rateio por conta contábil / FLEX) ---
    gerar_bloco_oc(ws)

    wb.save(caminho_saida)
    wb.close()
    return caminho_saida


def gerar_bloco_oc(ws) -> None:
    """Gera o bloco de rateio OC ao lado direito dos dados do CONSO.

    Lê as colunas CONTA, FLEX e VALOR do CONSO já gravado no *worksheet*,
    agrupa por conta contábil (ordem de aparição) e, dentro de cada conta,
    lista os FLEX únicos com valor e percentual de rateio (QTD).

    O bloco é escrito **2 colunas após** a última coluna do CONSO.
    Se já existir um bloco OC anterior, ele é substituído.

    Decisões de implementação:
      - "Lançamento" = cada linha de dados do CONSO (1 linha = 1 ocorrência).
      - CONTA vazia/None → agrupada como ``(Sem Conta)``.
      - FLEX vazio/None → representado como ``(Sem FLEX)``.
      - Percentuais calculados com ``Decimal`` (ROUND_HALF_UP, 2 casas).
      - Ajuste de arredondamento aplicado na OC com maior contagem.
    """
    # ── 1. Descobrir colunas obrigatórias ──────────────────────────────
    headers: dict[str, int] = {}
    for cell in ws[1]:
        if cell.value is not None:
            headers[str(cell.value).strip().upper()] = cell.column

    col_conta = headers.get("CONTA")
    col_flex  = headers.get("FLEX")
    col_valor = headers.get("VALOR")

    if not all((col_conta, col_flex, col_valor)):
        return  # CONSO vazio ou sem colunas obrigatórias

    # ── 2. Limpar bloco OC existente (re-execução) ─────────────────────
    col_linha_existente = headers.get("LINHA")
    max_col_antes = ws.max_column
    if col_linha_existente:
        for row in range(1, ws.max_row + 1):
            for col in range(col_linha_existente, max_col_antes + 1):
                ws.cell(row, col).value = None

    # ── 3. Ler dados do CONSO ──────────────────────────────────────────
    contas_ordem: list[str] = []
    contas_data: dict[str, dict[str, dict]] = {}

    for row in range(2, ws.max_row + 1):
        conta_raw = ws.cell(row, col_conta).value
        flex_raw  = ws.cell(row, col_flex).value
        valor_raw = ws.cell(row, col_valor).value

        # Pular linhas completamente vazias (ex.: sobras de um bloco OC
        # anterior que estendeu max_row além dos dados do CONSO).
        if conta_raw is None and flex_raw is None and valor_raw is None:
            continue

        conta = str(conta_raw).strip() if conta_raw is not None else "(Sem Conta)"
        flex  = str(flex_raw).strip()  if flex_raw  is not None else "(Sem FLEX)"
        try:
            valor = Decimal(str(valor_raw))
        except Exception:
            valor = Decimal("0")

        if conta not in contas_data:
            contas_ordem.append(conta)
            contas_data[conta] = {}

        bucket = contas_data[conta]
        if flex not in bucket:
            bucket[flex] = {"soma": Decimal("0"), "contagem": 0}
        bucket[flex]["soma"] += valor
        bucket[flex]["contagem"] += 1

    if not contas_ordem:
        return

    # ── 4. Determinar coluna-início do bloco OC ───────────────────────
    _OC_HEADERS_SET = {"LINHA", "OC", "VALOR OC", "QTD"}
    if col_linha_existente:
        col_inicio = col_linha_existente
    else:
        last_conso_col = max(
            (c for h, c in headers.items() if h not in _OC_HEADERS_SET),
            default=1,
        )
        col_inicio = last_conso_col + 3  # 2 colunas vazias de gap

    # ── 5. Cabeçalhos do bloco OC ──────────────────────────────────────
    _oc_nomes = ["linha", "OC", "valor OC", "QTD"]
    azul = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
    fonte_cab = Font(color="FFFFFF", bold=True)
    for i, nome in enumerate(_oc_nomes):
        c = ws.cell(1, col_inicio + i, value=nome)
        c.fill = azul
        c.font = fonte_cab

    # ── 6. Gerar linhas (cabeçalho-conta + detalhe-FLEX) ──────────────
    fmt_brl = "R$ #,##0.00"
    fmt_qtd = "0.00"
    row_out = 2

    for num_linha, conta in enumerate(contas_ordem, start=1):
        flexes = contas_data[conta]
        soma_total = sum(d["soma"] for d in flexes.values())
        total_contagem = sum(d["contagem"] for d in flexes.values())

        # -- linha-cabeçalho da conta --
        ws.cell(row_out, col_inicio, value=num_linha)
        ws.cell(row_out, col_inicio + 1, value=conta)
        cv = ws.cell(row_out, col_inicio + 2, value=float(soma_total))
        cv.number_format = fmt_brl
        cq = ws.cell(row_out, col_inicio + 3, value=100.0)
        cq.number_format = fmt_qtd
        row_out += 1

        # -- calcular QTD (percentual por contagem de lançamentos) --
        flex_items = list(flexes.items())
        percentuais: dict[str, Decimal] = {}

        if total_contagem > 0:
            for flex, dados in flex_items:
                pct = (
                    Decimal(str(dados["contagem"]))
                    / Decimal(str(total_contagem))
                    * Decimal("100")
                ).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                percentuais[flex] = pct
        else:
            # fallback: distribuir igualmente
            n = len(flex_items)
            pct_base = (Decimal("100") / Decimal(str(n))).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            percentuais = {f: pct_base for f, _ in flex_items}

        # ajuste de arredondamento na OC com maior contagem
        soma_pct = sum(percentuais.values())
        diff = Decimal("100.00") - soma_pct
        if diff != Decimal("0"):
            maior_flex = (
                max(flex_items, key=lambda x: x[1]["contagem"])[0]
                if total_contagem > 0
                else flex_items[0][0]
            )
            percentuais[maior_flex] += diff

        # -- linhas-detalhe --
        for flex, dados in flex_items:
            ws.cell(row_out, col_inicio, value=num_linha)
            ws.cell(row_out, col_inicio + 1, value=flex)
            cv = ws.cell(row_out, col_inicio + 2, value=float(dados["soma"]))
            cv.number_format = fmt_brl
            cq = ws.cell(row_out, col_inicio + 3, value=float(percentuais[flex]))
            cq.number_format = fmt_qtd
            row_out += 1

    # ── 7. Auto-ajuste de largura das colunas OC ──────────────────────
    for i, nome in enumerate(_oc_nomes):
        col_idx = col_inicio + i
        largura = len(nome)
        for r in range(2, row_out):
            val = ws.cell(r, col_idx).value
            if val is not None:
                largura = max(largura, len(str(val)))
        ws.column_dimensions[get_column_letter(col_idx)].width = largura + 4


def gerar_bloco_oc_em_arquivo(caminho_fornecedor: Path) -> None:
    """Abre o Excel do fornecedor, localiza a aba ``conso`` e gera/atualiza o
    bloco de rateio OC.

    Usado pelo botão *Criar OC* para gerar o bloco sem reprocessar o CONSO.

    Raises:
        ValueError: se a aba ``conso`` não existir no arquivo.
    """
    wb = load_workbook(caminho_fornecedor)
    if "conso" not in wb.sheetnames:
        wb.close()
        raise ValueError(f"Aba 'conso' não encontrada em {caminho_fornecedor.name}")

    ws = wb["conso"]
    gerar_bloco_oc(ws)
    wb.save(caminho_fornecedor)
    wb.close()


def inserir_conso_no_fornecedor(
    caminho_conso: Path,
    caminho_fornecedor: Path,
) -> None:
    """Copia a planilha do CONSO como aba 'conso' dentro do Excel do fornecedor.

    Abre o CONSO gerado e o Excel original do fornecedor, cria (ou substitui)
    uma aba chamada ``conso`` com todo o conteúdo -- valores, fórmulas,
    estilos e larguras de coluna.
    """
    wb_conso = load_workbook(caminho_conso)
    ws_conso = wb_conso.active

    wb_forn = load_workbook(caminho_fornecedor)

    # Remove aba 'conso' anterior se existir (re-execução)
    if "conso" in wb_forn.sheetnames:
        del wb_forn["conso"]

    ws_nova = wb_forn.create_sheet("conso")

    # Copiar célula por célula (valor/fórmula + estilo)
    for row in ws_conso.iter_rows():
        for cell in row:
            nova = ws_nova.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                nova.font = copy(cell.font)
                nova.fill = copy(cell.fill)
                nova.number_format = cell.number_format
                nova.alignment = copy(cell.alignment)
                nova.border = copy(cell.border)

    # Copiar larguras das colunas
    for col_letter, dim in ws_conso.column_dimensions.items():
        ws_nova.column_dimensions[col_letter].width = dim.width

    wb_forn.save(caminho_fornecedor)
    wb_forn.close()
    wb_conso.close()