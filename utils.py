"""Funções utilitárias de normalização de texto e localização de dados em planilhas.

Consolida helpers que apareciam repetidos (às vezes com pequenas variações)
em praticamente todos os scripts originais: normalizarTexto, normalizarCabecalho,
encontrarLinhaCabecalho, localizarPlanilha, localizarValorRotulado, localizarCabecalho
(versão "flexível", multi-linha, usada pelo ALD), pegarExcelFornecedor e
pegarExcelDesignacao.
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Iterable, Optional


# ---------------------------------------------------------------------------
# Normalização
# ---------------------------------------------------------------------------
def normalizar_texto(texto) -> str:
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.strip().casefold()


def normalizar_cabecalho(texto) -> str:
    return re.sub(r"[^a-z0-9]", "", normalizar_texto(texto))


def normalizar_espacos(texto) -> str:
    """Usado pelo EDENRED para normalizar 'Tipo de Transação' (remove espaços duplicados)."""
    return re.sub(r"\s+", " ", normalizar_texto(texto))


def limpar_placa(placa):
    """Remove hífen de placas tipo 'ABC-1234' -> 'ABC1234' (RELEMAR e UNIDAS fazem isso)."""
    return placa.replace("-", "") if isinstance(placa, str) else placa


# ---------------------------------------------------------------------------
# Localização de cabeçalho dentro de uma planilha já aberta (openpyxl worksheet)
# ---------------------------------------------------------------------------
def encontrar_linha_cabecalho(planilha, colunas_esperadas: set[str]):
    """Varre linha a linha procurando uma que contenha todas as `colunas_esperadas`
    (já normalizadas). Retorna (linha_de_valores, numero_da_linha) ou (None, None).
    """
    for numero_linha, linha in enumerate(planilha.iter_rows(values_only=True), start=1):
        cabecalhos_normalizados = {
            normalizar_cabecalho(c) for c in linha if c is not None
        }
        if colunas_esperadas.issubset(cabecalhos_normalizados):
            return linha, numero_linha
    return None, None


def indices_colunas(cabecalhos) -> dict:
    """A partir de uma linha de cabeçalhos (tupla de valores), monta
    {cabecalho_normalizado: indice}.
    """
    return {
        normalizar_cabecalho(c): i
        for i, c in enumerate(cabecalhos)
        if c is not None
    }


def localizar_cabecalho_flexivel(planilha, colunas_obrigatorias: Iterable[str], linhas_maximas: int = 5):
    """Versão 'flexível' usada pelo ALD: procura um bloco de até `linhas_maximas`
    linhas consecutivas cuja UNIÃO dos valores contenha todas as colunas
    obrigatórias (cobre cabeçalhos quebrados em mais de uma linha).
    Retorna (numero_ultima_linha_do_cabecalho, indices) ou levanta ValueError.
    """
    colunas_normalizadas = {normalizar_cabecalho(c) for c in colunas_obrigatorias}
    linhas = list(planilha.iter_rows(values_only=True))

    for inicio in range(len(linhas)):
        indices = {}
        ultima_linha = inicio
        for deslocamento in range(linhas_maximas):
            idx = inicio + deslocamento
            if idx >= len(linhas):
                break
            for col_idx, valor in enumerate(linhas[idx]):
                if valor is None:
                    continue
                chave = normalizar_cabecalho(valor)
                if chave and chave not in indices:
                    indices[chave] = col_idx
            ultima_linha = idx
            if colunas_normalizadas.issubset(indices):
                return ultima_linha + 1, indices

    raise ValueError(f"As colunas {sorted(colunas_normalizadas)} não foram encontradas em: {planilha.title}")


def localizar_planilha_com_cabecalho(workbook, colunas_obrigatorias: Iterable[str]):
    """Tenta `localizar_cabecalho_flexivel` em cada aba do workbook até achar uma
    que contenha as colunas obrigatórias. Retorna (planilha, linha_cabecalho, indices).
    """
    for planilha in workbook.worksheets:
        try:
            linha_cabecalho, idx = localizar_cabecalho_flexivel(planilha, colunas_obrigatorias)
        except ValueError:
            continue
        return planilha, linha_cabecalho, idx

    abas = ", ".join(p.title for p in workbook.worksheets)
    raise ValueError(f"As colunas {sorted(colunas_obrigatorias)} não foram encontradas em nenhuma aba ({abas})")


def localizar_planilha_por_nome(workbook, nome_planilha: str):
    """Acha a primeira aba cujo título (normalizado) CONTÉM `nome_planilha` (normalizado)."""
    termo = normalizar_cabecalho(nome_planilha)
    for planilha in workbook.worksheets:
        if termo in normalizar_cabecalho(planilha.title):
            return planilha
    raise ValueError(f"A planilha '{nome_planilha}' não foi encontrada em: {workbook.properties.title or 'arquivo'}")


def localizar_valor_rotulado(planilha, rotulo: str):
    """Procura uma célula cujo valor == `rotulo` e retorna o valor da célula à direita.
    Usado pelas planilhas GEOMATICS (ex.: 'Valor por instalação: 150,00').
    """
    rotulo_normalizado = normalizar_cabecalho(rotulo)
    for linha in planilha.iter_rows():
        for celula in linha:
            if normalizar_cabecalho(celula.value) == rotulo_normalizado:
                valor = planilha.cell(celula.row, celula.column + 1).value
                if valor is not None:
                    return valor
    raise ValueError(f"O rótulo '{rotulo}' não foi encontrado na planilha '{planilha.title}'")


# ---------------------------------------------------------------------------
# Localização de arquivos no disco
# ---------------------------------------------------------------------------
def pegar_excel_fornecedor(pasta_base: Path, pasta_fornecedor: str, nome_excel_fornecedor: str,
                            permitir_varios: bool = False):
    """Localiza o(s) arquivo(s) .xlsx do fornecedor dentro de `pasta_base/pasta_fornecedor`,
    cujo nome contém `nome_excel_fornecedor` e que não seja um CONSO nem uma designação.

    Se `permitir_varios` for False (padrão), retorna o primeiro Path encontrado
    (comportamento de VAMOS/TKS/LM/ALD/RELEMAR/MACARI/EDENRED).
    Se True, retorna uma lista com todos os arquivos que baterem (comportamento
    de GEOMATICS, que processa vários arquivos de uma vez).
    """
    pasta_fornecedor_path = pasta_base / pasta_fornecedor
    termo_busca = normalizar_texto(nome_excel_fornecedor)

    if not pasta_fornecedor_path.exists():
        raise FileNotFoundError(f"Pasta do fornecedor não encontrada: {pasta_fornecedor_path}")

    arquivos_xlsx = sorted(pasta_fornecedor_path.glob("*.xlsx"))
    candidatos = [
        arquivo for arquivo in arquivos_xlsx
        if termo_busca in normalizar_texto(arquivo.stem)
        and not normalizar_texto(arquivo.stem).startswith("conso")
        and "designacao" not in normalizar_texto(arquivo.stem)
    ]

    if not candidatos:
        raise FileNotFoundError(
            f"Nenhum arquivo com '{nome_excel_fornecedor}' encontrado em: {pasta_fornecedor_path}"
        )

    return candidatos if permitir_varios else candidatos[0]


def pegar_excel_designacao(caminho_excel_fornecedor: Path) -> Path:
    """Procura o arquivo 'Designação...xlsx' na pasta-mãe (ou avó) do arquivo do fornecedor."""
    pastas_designacao = [
        caminho_excel_fornecedor.parent.parent,
        caminho_excel_fornecedor.parent.parent.parent,
    ]

    for pasta in pastas_designacao:
        candidatos = sorted(
            arquivo for arquivo in pasta.glob("*.xlsx")
            if "designacao" in normalizar_texto(arquivo.stem)
        )
        if candidatos:
            return candidatos[0]

    tentadas = ", ".join(str(p) for p in pastas_designacao)
    raise FileNotFoundError(f"Nenhum arquivo com 'Designação' encontrado nas pastas: {tentadas}")