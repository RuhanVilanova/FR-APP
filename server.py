"""Servidor Flask — Databricks App para conciliação de frotas.

Formas de execução:
  1. Databricks App (script direto):  python server.py
  2. Como módulo do pacote:           python -m FR_POO.server

A porta é lida de DATABRICKS_APP_PORT (definida automaticamente pelo
Databricks Apps); em execução local usa 5000.

Este arquivo é o entrypoint do app.yaml e precisa conter TODAS as rotas
que o index.html consome.
"""

# ---------------------------------------------------------------------------
# PEP 366 — Bootstrap para execução como script direto.
# ---------------------------------------------------------------------------
from __future__ import annotations

if __package__ is None or __package__ == "":
    import sys
    from pathlib import Path as _Path

    _pkg_dir = _Path(__file__).resolve().parent
    _parent  = str(_pkg_dir.parent)
    if _parent not in sys.path:
        sys.path.insert(0, _parent)

    __package__ = _pkg_dir.name
    __import__(__package__)

# ---------------------------------------------------------------------------
import base64
import os
import shutil
import tempfile
import uuid
from pathlib import Path
from urllib.parse import unquote

from io import BytesIO

from flask import Flask, abort, jsonify, request, send_file, send_from_directory
from flask_cors import CORS

from . import config, excel_writer, utils, volume_io
from .factory import FornecedorProcessorFactory

# ---------------------------------------------------------------------------
# App Flask
# ---------------------------------------------------------------------------
_PASTA_PACOTE = Path(__file__).resolve().parent

app = Flask(__name__)
CORS(app)

PORT = int(os.environ.get("DATABRICKS_APP_PORT", 5000))


# ---------------------------------------------------------------------------
# Upload: roteamento de arquivos para pastas de fornecedores
# ---------------------------------------------------------------------------
_PASTAS_FORNECEDORES = sorted({
    "VAMOS", "TKS", "LM", "ALD", "RELEMAR", "CAMERAS",
    "CAMERAS E SENSORES", "RASTREADORES DE CARGA",
    "UNIDAS", "MACARI", "EDENRED", "TICKET",
})

_EXCEL_PARA_PASTA = [
    ("geomatics cameras e sensores", "CAMERAS E SENSORES"),
    ("geomatics rastreador",         "RASTREADORES DE CARGA"),
    ("geomatics camera",             "CAMERAS"),
    ("relemar",                      "RELEMAR"),
    ("vamos",                        "VAMOS"),
    ("unidas",                       "UNIDAS"),
    ("macari",                       "MACARI"),
    ("edenred",                      "EDENRED"),
    ("ticket",                       "TICKET"),
    ("tks",                          "TKS"),
    ("ald",                          "ALD"),
]

_upload_sessions: dict = {}


def _rotear_arquivo(nome: str, caminho_relativo: str = "") -> tuple:
    """Retorna (pasta_destino | None, tipo)."""
    nome_norm = utils.normalizar_texto(nome)

    if nome.startswith("~$") or nome_norm.startswith("conso"):
        return None, "ignorado"

    if "designacao" in nome_norm:
        return ".", "designacao_mensal"

    if caminho_relativo:
        for parte in Path(caminho_relativo).parts:
            parte_norm = utils.normalizar_texto(parte)
            for pf in _PASTAS_FORNECEDORES:
                if utils.normalizar_texto(pf) == parte_norm:
                    return pf, "fornecedor"

    for padrao, pasta in _EXCEL_PARA_PASTA:
        if padrao in nome_norm:
            ext = Path(nome).suffix.lower()
            return pasta, "fornecedor" if ext == ".xlsx" else "pdf"

    if Path(nome).suffix.lower() == ".pdf":
        return None, "pdf_desconhecido"

    return None, "desconhecido"


def _validar_xlsx(caminho_arquivo: Path) -> str | None:
    try:
        from openpyxl import load_workbook
        wb = load_workbook(caminho_arquivo, read_only=True)
        wb.close()
        return None
    except Exception as exc:
        return str(exc)


# ---------------------------------------------------------------------------
# Helpers TICKET
# ---------------------------------------------------------------------------
def _processor_ticket():
    return FornecedorProcessorFactory.criar("TICKET_COMBUSTIVEL")


def _caminhos_excel_ticket() -> list[Path]:
    processor = _processor_ticket()
    caminhos = utils.pegar_excel_fornecedor(
        processor.pasta_base,
        processor.pasta_fornecedor,
        processor.nome_excel_fornecedor,
        permitir_varios=processor.permitir_varios_arquivos,
    )
    return caminhos if isinstance(caminhos, list) else [caminhos]


def _pendencias_ticket() -> list[dict]:
    try:
        processor = _processor_ticket()
        caminhos = _caminhos_excel_ticket()
    except (ValueError, FileNotFoundError):
        return []

    extrator = processor.extrator
    if not hasattr(extrator, "listar_servicos_pendentes"):
        return []

    pendentes: list[dict] = []
    vistos: set[str] = set()
    for caminho in caminhos:
        for item in extrator.listar_servicos_pendentes(caminho):
            chave = item["chave"]
            if chave in vistos:
                continue
            vistos.add(chave)
            pendentes.append({
                "servico": item["servico"],
                "servico_exibicao": item["servico_exibicao"],
            })
    return pendentes


# ===========================================================================
# Rotas
# ===========================================================================

@app.route("/")
def index():
    return send_from_directory(str(_PASTA_PACOTE), "index.html")


# ---------------------------------------------------------------------------
# GET /api/fornecedores
# ---------------------------------------------------------------------------
@app.route("/api/fornecedores", methods=["GET"])
def listar_fornecedores():
    return jsonify({
        "mes_atual": config.MES_ATUAL,
        "indice_mes": config.indice_mes_atual(),
        "meses": config.meses_disponiveis(),
        "fornecedores": FornecedorProcessorFactory.fornecedores_disponiveis(),
        "pastas_fornecedores": _PASTAS_FORNECEDORES,
    })


# ---------------------------------------------------------------------------
# POST /api/executar
# ---------------------------------------------------------------------------
@app.route("/api/executar", methods=["POST"])
def executar():
    dados = request.get_json(silent=True) or {}
    lista_solicitada = dados.get("fornecedores") or []

    if lista_solicitada:
        disponiveis = set(FornecedorProcessorFactory.fornecedores_disponiveis())
        invalidos = [n for n in lista_solicitada if n not in disponiveis]
        if invalidos:
            return jsonify({
                "mes_atual": config.MES_ATUAL,
                "resultados": [],
                "erros": [{"fornecedor": n, "mensagem": f"'{n}' nao registrado"} for n in invalidos],
            }), 400
        fornecedores = lista_solicitada
    else:
        fornecedores = FornecedorProcessorFactory.fornecedores_disponiveis()

    resultados_json: list[dict] = []
    erros: list[dict] = []

    with volume_io.volume_local():
        # Checar pendências TICKET dentro do context (precisa dos arquivos locais)
        if "TICKET_COMBUSTIVEL" in fornecedores:
            pend = _pendencias_ticket()
            if pend:
                return jsonify({
                    "mes_atual": config.MES_ATUAL,
                    "pendencias_ticket": pend,
                    "contas_disponiveis": config.contas_cadastradas(),
                }), 409

        for nome in fornecedores:
            try:
                processor = FornecedorProcessorFactory.criar(nome)
                resultado = processor.processar()
                resultados_json.append({
                    "fornecedor":            resultado.fornecedor,
                    "chave":                 nome,
                    "arquivos_gerados":      [str(p) for p in resultado.arquivos_gerados],
                    "placas_sem_designacao": resultado.placas_sem_designacao,
                    "tipos_nao_mapeados":    resultado.tipos_nao_mapeados,
                })
            except Exception as exc:
                erros.append({"fornecedor": nome, "mensagem": str(exc)})

    return jsonify({
        "mes_atual":  config.MES_ATUAL,
        "resultados": resultados_json,
        "erros":      erros,
    })


# ---------------------------------------------------------------------------
# POST /api/criar-oc  —  gera bloco OC em CONSO existente
# ---------------------------------------------------------------------------
@app.route("/api/criar-oc", methods=["POST"])
def criar_oc():
    dados = request.get_json(silent=True) or {}
    lista_solicitada = dados.get("fornecedores") or []

    if lista_solicitada:
        disponiveis = set(FornecedorProcessorFactory.fornecedores_disponiveis())
        invalidos = [n for n in lista_solicitada if n not in disponiveis]
        if invalidos:
            return jsonify({
                "resultados": [],
                "erros": [{"fornecedor": n, "mensagem": f"'{n}' nao registrado"} for n in invalidos],
            }), 400
        fornecedores = lista_solicitada
    else:
        fornecedores = FornecedorProcessorFactory.fornecedores_disponiveis()

    resultados_json: list[dict] = []
    erros: list[dict] = []

    with volume_io.volume_local():
        for nome in fornecedores:
            try:
                processor = FornecedorProcessorFactory.criar(nome)
                caminhos = utils.pegar_excel_fornecedor(
                    processor.pasta_base,
                    processor.pasta_fornecedor,
                    processor.nome_excel_fornecedor,
                    permitir_varios=processor.permitir_varios_arquivos,
                )
                if not isinstance(caminhos, list):
                    caminhos = [caminhos]

                for caminho in caminhos:
                    excel_writer.gerar_bloco_oc_em_arquivo(caminho)
                    resultados_json.append({
                        "fornecedor": nome,
                        "arquivo": caminho.name,
                    })
            except Exception as exc:
                erros.append({"fornecedor": nome, "mensagem": str(exc)})

    return jsonify({
        "resultados": resultados_json,
        "erros": erros,
    })


# ---------------------------------------------------------------------------
# POST /api/ticket/mapear-servico
# ---------------------------------------------------------------------------
@app.route("/api/ticket/mapear-servico", methods=["POST"])
def mapear_servico_ticket():
    dados = request.get_json(silent=True) or {}
    if "conta_id" not in dados:
        return jsonify({"erro": "Campo 'conta_id' e obrigatorio"}), 400

    extrator = _processor_ticket().extrator
    if not hasattr(extrator, "salvar_mapeamento"):
        return jsonify({"erro": "Extrator do TICKET nao suporta salvamento"}), 500

    extrator.salvar_mapeamento(dados.get("servico"), dados["conta_id"])
    return jsonify({
        "pendencias_ticket": _pendencias_ticket(),
        "contas_disponiveis": config.contas_cadastradas(),
    })


# ---------------------------------------------------------------------------
# POST /api/alterar-mes
# ---------------------------------------------------------------------------
@app.route("/api/alterar-mes", methods=["POST"])
def alterar_mes():
    dados = request.get_json(silent=True) or {}
    indice = dados.get("indice")
    if indice is None:
        return jsonify({"erro": "Campo 'indice' e obrigatorio"}), 400
    try:
        config.alterar_mes(int(indice))
    except ValueError as exc:
        return jsonify({"erro": str(exc)}), 400
    return jsonify({
        "mes_atual": config.MES_ATUAL,
        "indice_mes": config.indice_mes_atual(),
        "pasta": str(config.CAMINHO_CONTROLE_DESPESAS_FR),
    })


# ---------------------------------------------------------------------------
# GET /api/arquivos  —  lista arquivos do mês selecionado
# ---------------------------------------------------------------------------
@app.route("/api/arquivos", methods=["GET"])
def listar_arquivos():
    mes_idx = request.args.get("mes")
    if mes_idx is None:
        return jsonify({"erro": "Parametro 'mes' e obrigatorio"}), 400
    try:
        mes_idx = int(mes_idx)
        nome_mes = config._MESES[mes_idx]
        pasta_mes = config.CAMINHO_VOLUME / "FR - FROTAS" / "2026" / f"{mes_idx + 1:02d}. {nome_mes}"
    except (ValueError, IndexError):
        return jsonify({"erro": "Indice de mes invalido"}), 400

    pasta_mes_str = str(pasta_mes)
    arquivos = []
    for item in volume_io.listar_recursivo(pasta_mes_str):
        rel = item["caminho"][len(pasta_mes_str):].lstrip("/")
        partes = rel.rsplit("/", 1)
        pasta_rel = partes[0] if len(partes) > 1 else ""
        arquivos.append({
            "nome": item["nome"],
            "pasta": pasta_rel,
            "caminho": item["caminho"],
            "tamanho": item["tamanho"],
        })

    aux_str = str(config.CAMINHO_VOLUME / "FR - FROTAS" / "AUXILIARES")
    for item in volume_io.listar_recursivo(aux_str):
        arquivos.append({
            "nome": item["nome"],
            "pasta": "AUXILIARES",
            "caminho": item["caminho"],
            "tamanho": item["tamanho"],
        })

    return jsonify({
        "mes": nome_mes,
        "pasta": str(pasta_mes),
        "arquivos": arquivos,
    })


# ---------------------------------------------------------------------------
# GET /api/download
# ---------------------------------------------------------------------------
@app.route("/api/download", methods=["GET"])
def download_arquivo():
    caminho = request.args.get("path", "")
    if not caminho:
        return jsonify({"erro": "Parametro 'path' e obrigatorio"}), 400

    caminho = unquote(caminho)
    nome = Path(caminho).name
    ext = Path(caminho).suffix.lower()
    if ext not in (".xlsx", ".pdf"):
        abort(404, "Tipo de arquivo nao permitido")

    try:
        conteudo = volume_io.download_bytes(caminho)
    except Exception:
        abort(404, "Arquivo nao encontrado")

    mime = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        if ext == ".xlsx"
        else "application/pdf"
    )
    return send_file(BytesIO(conteudo), as_attachment=True, download_name=nome, mimetype=mime)


# ---------------------------------------------------------------------------
# POST /api/upload  —  recebe arquivos base64, valida, roteia, salva em temp
# ---------------------------------------------------------------------------
@app.route("/api/upload", methods=["POST"])
def upload_arquivos():
    try:
        dados = request.get_json(silent=True) or {}
        modo = dados.get("modo", "pasta")
        mes = dados.get("mes")
        arquivos_in = dados.get("arquivos", [])

        if not arquivos_in:
            return jsonify({"erro": "Nenhum arquivo enviado"}), 400

        session_id = str(uuid.uuid4())
        temp_dir = Path(tempfile.mkdtemp(prefix="fr_upload_"))

        if modo == "designacao_historica":
            pasta_base = config.CAMINHO_AUXILIARES
        else:
            if mes is not None:
                config.alterar_mes(int(mes))
            pasta_base = config.CAMINHO_CONTROLE_DESPESAS_FR

        resultados = []
        for arq in arquivos_in:
            nome = arq["nome"]
            caminho_rel = arq.get("caminho_relativo", nome)
            dados_b64 = arq.get("dados", "")

            temp_path = temp_dir / nome
            try:
                temp_path.write_bytes(base64.b64decode(dados_b64))
            except Exception as exc:
                resultados.append({"nome": nome, "valido": False, "erro": f"Erro ao decodificar: {exc}"})
                continue

            erro_validacao = None
            if nome.lower().endswith(".xlsx"):
                erro_validacao = _validar_xlsx(temp_path)

            if modo == "designacao_mensal":
                destino_rel, tipo = ".", "designacao_mensal"
            elif modo == "designacao_historica":
                destino_rel, tipo = ".", "designacao_historica"
            else:
                destino_rel, tipo = _rotear_arquivo(nome, caminho_rel)

            if destino_rel is not None:
                destino_final = pasta_base / destino_rel / nome if destino_rel != "." else pasta_base / nome
                if modo == "designacao_historica":
                    destino_final = config.CAMINHO_DESIGNACAO_HISTORICA
                conflito = volume_io.arquivo_existe(str(destino_final))
            else:
                destino_final = None
                conflito = False

            resultados.append({
                "nome": nome,
                "destino": str(destino_final) if destino_final else None,
                "destino_exibicao": f"{destino_rel}/{nome}" if destino_rel and destino_rel != "." else nome,
                "tipo": tipo,
                "valido": erro_validacao is None,
                "erro": f"Arquivo não adicionado: remova o status do excel {nome} de interno para publico" if erro_validacao else None,
                "conflito": conflito,
                "temp": str(temp_path),
            })

        _upload_sessions[session_id] = {
            "temp_dir": str(temp_dir),
            "resultados": resultados,
            "modo": modo,
        }

        return jsonify({
            "sessao": session_id,
            "resultados": [
                {k: v for k, v in r.items() if k != "temp"}
                for r in resultados
            ],
        })
    except Exception as exc:
        return jsonify({"erro": f"Erro interno no upload: {exc}"}), 500


# ---------------------------------------------------------------------------
# POST /api/upload/confirmar  —  move arquivos validados para destino final
# ---------------------------------------------------------------------------
@app.route("/api/upload/confirmar", methods=["POST"])
def upload_confirmar():
    try:
        dados = request.get_json(silent=True) or {}
        session_id = dados.get("sessao")
        substituir = set(dados.get("substituir", []))

        session = _upload_sessions.pop(session_id, None)
        if not session:
            return jsonify({"erro": "Sessão expirada ou inválida"}), 400

        sucesso = []
        erros = []

        for r in session["resultados"]:
            nome = r["nome"]

            if not r["valido"]:
                erros.append({"nome": nome, "erro": r["erro"]})
                continue

            if r["destino"] is None:
                erros.append({"nome": nome, "erro": "Destino não identificado"})
                continue

            destino = Path(r["destino"])
            temp_path = Path(r["temp"])

            if volume_io.arquivo_existe(str(destino)) and nome not in substituir:
                erros.append({"nome": nome, "erro": "Conflito não resolvido (arquivo já existe)"})
                continue

            try:
                volume_io.upload_arquivo(temp_path, str(destino))
                sucesso.append({"nome": nome, "destino": str(destino)})
            except Exception as exc:
                erros.append({"nome": nome, "erro": str(exc)})

        try:
            shutil.rmtree(session["temp_dir"])
        except Exception:
            pass

        return jsonify({"sucesso": sucesso, "erros": erros})
    except Exception as exc:
        return jsonify({"erro": f"Erro interno ao confirmar upload: {exc}"}), 500


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    n_forn = len(FornecedorProcessorFactory.fornecedores_disponiveis())
    print(f"\n  FR Frotas — Backend rodando em http://localhost:{PORT}")
    print(f"  Mês vigente: {config.MES_ATUAL}")
    print(f"  Fornecedores: {n_forn}")
    print(f"  Pasta: {config.CAMINHO_CONTROLE_DESPESAS_FR}")
    print(f"  Ctrl+C para encerrar\n")
    app.run(debug=False, host="0.0.0.0", port=PORT)
