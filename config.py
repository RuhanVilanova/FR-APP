"""Configuração global: caminhos (UC Volume) e tabela de contas contábeis."""
from __future__ import annotations

import os
from datetime import date
from pathlib import Path

# ---------------------------------------------------------------------------
# Caminhos base — Unity Catalog Volume
# No Databricks App, o volume é exposto via env var VOLUME_URI (app.yaml).
# O acesso ao volume é feito via Databricks SDK (volume_io.py).
# Os Path objects são usados para construção de caminhos; o I/O real usa SDK.
# ---------------------------------------------------------------------------
CAMINHO_VOLUME = Path(os.environ.get("VOLUME_URI", "/Volumes/<catalog>/<schema>/<volume>"))

_MESES = (
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
)

# Ano e mês dinâmicos a partir da data atual.
# Podem ser alterados em tempo de execução via `alterar_mes()`.
ANO_ATUAL = str(date.today().year)
MES_ATUAL = _MESES[date.today().month - 1]

_NUM_MES = date.today().month
CAMINHO_CONTROLE_DESPESAS_FR = (
    CAMINHO_VOLUME / "FR - FROTAS" / ANO_ATUAL / f"{_NUM_MES:02d}. {MES_ATUAL}"
)

CAMINHO_AUXILIARES = CAMINHO_VOLUME / "FR - FROTAS" / "AUXILIARES"
CAMINHO_DESIGNACAO_HISTORICA = CAMINHO_AUXILIARES / "DESIGNAÇÃO HISTORICA CLOUD.xlsx"
CAMINHO_CONTAS_TICKET = Path(__file__).resolve().parent / "Contas_ticket.json"


def _recalcular_caminhos(indice_mes: int) -> None:
    """Recalcula os caminhos globais para o mês indicado (0=JAN, 11=DEZ)."""
    global MES_ATUAL, CAMINHO_CONTROLE_DESPESAS_FR

    MES_ATUAL = _MESES[indice_mes]
    numero_pasta = indice_mes + 1

    CAMINHO_CONTROLE_DESPESAS_FR = (
        CAMINHO_VOLUME / "FR - FROTAS" / ANO_ATUAL / f"{numero_pasta:02d}. {MES_ATUAL}"
    )


def alterar_mes(indice: int) -> str:
    """Altera o mês vigente (0=JAN … 11=DEZ) e retorna o novo MES_ATUAL."""
    if not (0 <= indice < 12):
        raise ValueError(f"Índice de mês inválido: {indice}")
    _recalcular_caminhos(indice)
    return MES_ATUAL


def indice_mes_atual() -> int:
    """Retorna o índice (0–11) do mês vigente."""
    return _MESES.index(MES_ATUAL)


def meses_disponiveis() -> list[dict]:
    """Lista de meses para o frontend."""
    return [{"indice": i, "nome": nome} for i, nome in enumerate(_MESES)]

# ---------------------------------------------------------------------------
# Tabela de contas contábeis por tipo de gasto
# ---------------------------------------------------------------------------
# NOTA (versão pública): números de conta FICTÍCIOS — substitua pelos do seu plano de contas.
GASTO_CONTAS = {
    "Locação de veículos": {"id": 1, "conta": "90000001"},
    "Pedágios": {"id": 2, "conta": "90000002"},
    "Lavagem e lubrificação de veículos utilitários": {"id": 3, "conta": "90000003"},
    "Gasolina outros": {"id": 4, "conta": "90000004"},
    "Diesel": {"id": 5, "conta": "90000005"},
    "Transferência de Veículos": {"id": 6, "conta": "90000006"},
    "Manutenção de frotas e equipamentos de transporte (utilitários)": {"id": 7, "conta": "90000007"},
    "Mensagens": {"id": 8, "conta": "90000008"},
    "Peças de veículos": {"id": 9, "conta": "90000009"},
    "Pneus": {"id": 10, "conta": "90000010"},
    "Multas estaduais por penalidades": {"id": 11, "conta": "90000011"},
    "Serviços administrativos (externos)": {"id": 12, "conta": "90000012"},
    "Caminhões de distribuição a longo prazo": {"id": 13, "conta": "90000013"},
    "Licenças e placas": {"id": 14, "conta": "90000014"},
    "Óleos e lubrificantes para máquinas": {"id": 15, "conta": "90000015"},
    "Gastos com viagem Estacionamento": {"id": 16, "conta": "90000016"},
    "Controle de pragas": {"id": 17, "conta": "90000017"},
}


_ID_PARA_TIPO_GASTO = {
    dados["id"]: tipo_gasto
    for tipo_gasto, dados in GASTO_CONTAS.items()
}


def conta_por_tipo_gasto(tipo_gasto: str) -> str:
    """Busca a conta contábil de um tipo de gasto; levanta erro claro se não existir."""
    dados = GASTO_CONTAS.get(tipo_gasto)
    if dados is None:
        raise ValueError(f"Nenhuma conta cadastrada para o tipo de gasto: {tipo_gasto}")
    return dados["conta"]


def id_conta_por_tipo_gasto(tipo_gasto: str) -> int:
    """Retorna o ID interno da conta para um tipo de gasto cadastrado."""
    dados = GASTO_CONTAS.get(tipo_gasto)
    if dados is None:
        raise ValueError(f"Nenhuma conta cadastrada para o tipo de gasto: {tipo_gasto}")
    return dados["id"]


def tipo_gasto_por_id(conta_id: int | str) -> str:
    """Retorna o tipo de gasto a partir do ID interno da conta."""
    try:
        conta_id_int = int(conta_id)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"ID de conta inválido: {conta_id}") from exc

    tipo_gasto = _ID_PARA_TIPO_GASTO.get(conta_id_int)
    if tipo_gasto is None:
        raise ValueError(f"Nenhuma conta cadastrada para o ID: {conta_id_int}")
    return tipo_gasto


def conta_por_id(conta_id: int | str) -> str:
    """Retorna o número da conta contábil a partir do ID interno."""
    return conta_por_tipo_gasto(tipo_gasto_por_id(conta_id))


def contas_cadastradas() -> list[dict]:
    """Lista todas as contas no formato usado pelo backend/frontend."""
    return [
        {
            "id": dados["id"],
            "tipo_gasto": tipo_gasto,
            "conta": dados["conta"],
            "label": f'{dados["id"]}. {tipo_gasto} - {dados["conta"]}',
        }
        for tipo_gasto, dados in sorted(GASTO_CONTAS.items(), key=lambda item: item[1]["id"])
    ]