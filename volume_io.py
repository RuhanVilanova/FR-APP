"""Acesso ao UC Volume via Databricks SDK.

Centraliza TODAS as operações de I/O no Unity Catalog Volume, substituindo
acessos FUSE (pathlib direto em /Volumes/…) por chamadas REST via SDK.

Uso típico:
    from . import volume_io

    # operações pontuais
    volume_io.arquivo_existe("/Volumes/cat/sch/vol/arq.xlsx")
    conteudo = volume_io.download_bytes("/Volumes/cat/sch/vol/arq.xlsx")
    volume_io.upload_bytes("/Volumes/cat/sch/vol/arq.xlsx", dados)
    items = volume_io.listar_recursivo("/Volumes/cat/sch/vol/pasta")

    # processamento batch: baixa → temp local → processa → upload
    with volume_io.volume_local() as temp_root:
        # config.CAMINHO_* redirecionados para temp_root
        processar_normalmente()
    # upload automático dos arquivos novos/modificados
"""
from __future__ import annotations

import contextlib
import shutil
import tempfile
from pathlib import Path
from typing import Generator

from databricks.sdk import WorkspaceClient

from . import config

# ---------------------------------------------------------------------------
# Singleton do WorkspaceClient
# ---------------------------------------------------------------------------
_ws_client: WorkspaceClient | None = None


def _sdk() -> WorkspaceClient:
    global _ws_client
    if _ws_client is None:
        _ws_client = WorkspaceClient()
    return _ws_client


# ---------------------------------------------------------------------------
# Operações atômicas
# ---------------------------------------------------------------------------

def arquivo_existe(caminho_volume: str) -> bool:
    """Verifica se um arquivo existe no UC Volume via SDK."""
    try:
        resp = _sdk().files.get_status(caminho_volume)
        return not resp.is_dir
    except Exception:
        return False


def diretorio_existe(caminho_volume: str) -> bool:
    """Verifica se um diretório existe no UC Volume via SDK."""
    try:
        resp = _sdk().files.get_status(caminho_volume)
        return resp.is_dir
    except Exception:
        return False


def download_bytes(caminho_volume: str) -> bytes:
    """Baixa conteúdo de um arquivo do UC Volume."""
    resp = _sdk().files.download(caminho_volume)
    return resp.contents.read()


def upload_bytes(caminho_volume: str, dados: bytes) -> None:
    """Faz upload de bytes para o UC Volume."""
    from io import BytesIO
    _sdk().files.upload(caminho_volume, BytesIO(dados), overwrite=True)


def upload_arquivo(origem_local: Path, destino_volume: str) -> None:
    """Copia arquivo local para o UC Volume via SDK."""
    with open(origem_local, "rb") as f:
        _sdk().files.upload(destino_volume, f, overwrite=True)


def listar_recursivo(pasta: str) -> list[dict]:
    """Lista recursivamente arquivos em uma pasta do volume via SDK.

    Retorna lista de dicts com chaves: nome, caminho, tamanho.
    """
    resultado: list[dict] = []
    try:
        items = list(_sdk().files.list_directory_contents(pasta))
    except Exception:
        return resultado
    for item in items:
        if item.is_directory:
            resultado.extend(listar_recursivo(item.path))
        elif not item.name.startswith("~$"):
            resultado.append({
                "nome": item.name,
                "caminho": item.path,
                "tamanho": item.file_size or 0,
            })
    return resultado


def baixar_pasta(pasta_volume: str, pasta_local: Path) -> None:
    """Baixa recursivamente uma pasta do UC Volume para disco local via SDK."""
    pasta_local.mkdir(parents=True, exist_ok=True)
    try:
        items = list(_sdk().files.list_directory_contents(pasta_volume))
    except Exception:
        return
    for item in items:
        if item.is_directory:
            baixar_pasta(item.path, pasta_local / item.name)
        elif not item.name.startswith("~$"):
            try:
                resp = _sdk().files.download(item.path)
                (pasta_local / item.name).write_bytes(resp.contents.read())
            except Exception:
                pass


def upload_pasta_modificados(
    pasta_local: Path,
    pasta_volume_destino: str,
    snapshot_antes: dict[Path, float],
) -> None:
    """Faz upload dos arquivos novos ou modificados de volta para o volume."""
    for p in pasta_local.rglob("*"):
        if p.is_file() and not p.name.startswith("~$"):
            if p not in snapshot_antes or p.stat().st_mtime > snapshot_antes[p]:
                rel = p.relative_to(pasta_local)
                destino = f"{pasta_volume_destino}/{rel}"
                try:
                    with open(p, "rb") as f:
                        _sdk().files.upload(destino, f, overwrite=True)
                except Exception:
                    pass


def criar_diretorio(caminho_volume: str) -> None:
    """Cria um diretório no UC Volume (no-op se já existir)."""
    try:
        _sdk().files.create_directory(caminho_volume)
    except Exception:
        pass  # já existe ou erro silencioso


def garantir_estrutura_ano() -> None:
    """Cria a cadeia de pastas do ano atual se ainda não existir.

    Estrutura criada:
        FR - FROTAS / <ANO> / 01. JANEIRO / <fornecedor>
        FR - FROTAS / <ANO> / 02. FEVEREIRO / <fornecedor>
        …
        FR - FROTAS / <ANO> / 12. DEZEMBRO / <fornecedor>
    """
    pasta_ano = str(config.CAMINHO_VOLUME / "FR - FROTAS" / config.ANO_ATUAL)

    if diretorio_existe(pasta_ano):
        return  # ano já existe, nada a fazer

    _MESES = (
        "JANEIRO", "FEVEREIRO", "MAR\u00c7O", "ABRIL", "MAIO", "JUNHO",
        "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
    )
    _PASTAS_FORNECEDORES = sorted({
        "VAMOS", "TKS", "LM", "ALD", "RELEMAR", "CAMERAS",
        "CAMERAS E SENSORES", "RASTREADORES DE CARGA",
        "UNIDAS", "MACARI", "EDENRED", "TICKET",
    })

    criar_diretorio(pasta_ano)
    for idx, mes in enumerate(_MESES):
        pasta_mes = f"{pasta_ano}/{idx + 1:02d}. {mes}"
        criar_diretorio(pasta_mes)
        for fornecedor in _PASTAS_FORNECEDORES:
            criar_diretorio(f"{pasta_mes}/{fornecedor}")

    # Garantir pasta AUXILIARES também existe
    criar_diretorio(str(config.CAMINHO_AUXILIARES))


# ---------------------------------------------------------------------------
# Context manager: download → temp local → processa → upload
# ---------------------------------------------------------------------------

@contextlib.contextmanager
def volume_local() -> Generator[Path, None, None]:
    """Context manager que espelha o volume em temp local para processamento.

    1. Baixa a pasta do mês + AUXILIARES para um diretório temporário
    2. Redireciona config.CAMINHO_* para o temp (processamento lê/escreve localmente)
    3. Após o yield, faz upload de arquivos novos/modificados de volta para o volume
    4. Restaura config e limpa o temp
    """
    temp_root = Path(tempfile.mkdtemp(prefix="fr_proc_"))

    vol_original = config.CAMINHO_VOLUME
    ctrl_original = config.CAMINHO_CONTROLE_DESPESAS_FR
    aux_original = config.CAMINHO_AUXILIARES
    des_original = config.CAMINHO_DESIGNACAO_HISTORICA

    try:
        # Espelhar estrutura do volume localmente
        rel_ctrl = ctrl_original.relative_to(vol_original)
        pasta_mes_local = temp_root / rel_ctrl
        baixar_pasta(str(ctrl_original), pasta_mes_local)

        rel_aux = aux_original.relative_to(vol_original)
        aux_local = temp_root / rel_aux
        baixar_pasta(str(aux_original), aux_local)

        # Snapshot antes do processamento (detectar novos/modificados)
        antes: dict[Path, float] = {}
        for p in pasta_mes_local.rglob("*"):
            if p.is_file():
                antes[p] = p.stat().st_mtime

        # Redirecionar config para temp
        config.CAMINHO_VOLUME = temp_root
        config.CAMINHO_CONTROLE_DESPESAS_FR = pasta_mes_local
        config.CAMINHO_AUXILIARES = aux_local
        config.CAMINHO_DESIGNACAO_HISTORICA = aux_local / des_original.name

        yield temp_root

        # Upload arquivos novos ou modificados de volta para o volume
        upload_pasta_modificados(pasta_mes_local, str(ctrl_original), antes)

    finally:
        config.CAMINHO_VOLUME = vol_original
        config.CAMINHO_CONTROLE_DESPESAS_FR = ctrl_original
        config.CAMINHO_AUXILIARES = aux_original
        config.CAMINHO_DESIGNACAO_HISTORICA = des_original
        shutil.rmtree(temp_root, ignore_errors=True)
