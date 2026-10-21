    from __future__ import annotations
    """Orquestrador — conciliação de frotas (CLI + interface web local).

    Uso (de dentro da pasta FR_POO):
        python main.py                         # abre interface web no navegador
        python main.py --port 8080            # web na porta 8080
        python main.py --cli                  # roda TODOS via terminal
        python main.py --cli VAMOS TKS        # roda só os informados

    Ou como módulo (do diretório pai da FR_POO):
        python -m FR_POO.main
        python -m FR_POO.main --cli
    """

    # ---------------------------------------------------------------------------
    # Bootstrap: permite rodar como `python main.py` de dentro da pasta do pacote.
    # Quando executado assim, __package__ é None e imports relativos falham.
    # Resolvemos adicionando o diretório-pai ao sys.path.
    # ---------------------------------------------------------------------------
    if __package__ is None or __package__ == "":
        import sys as _sys
        from pathlib import Path as _Path

        _pkg_dir = _Path(__file__).resolve().parent
        _parent  = str(_pkg_dir.parent)
        if _parent not in _sys.path:
            _sys.path.insert(0, _parent)

        __package__ = _pkg_dir.name
        __import__(__package__)

    import base64
    import shutil
    import sys
    import tempfile
    import uuid
    import webbrowser
    from pathlib import Path
    from typing import List

    from . import config, excel_writer, utils, volume_io
    from .factory import FornecedorProcessorFactory
    from .processor import ResultadoProcessamento

    # --------------------------------------------------------------------------
    # Upload: roteamento de arquivos para pastas de fornecedores
    # --------------------------------------------------------------------------
    _PASTAS_FORNECEDORES = sorted({
        "VAMOS", "TKS", "LM", "ALD", "RELEMAR", "CAMERAS",
        "CAMERAS E SENSORES", "RASTREADORES DE CARGA",
        "UNIDAS", "MACARI", "EDENRED", "TICKET",
    })

    # Padrões de nome de Excel → pasta (mais longos primeiro para match correto)
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

    _upload_sessions: dict = {}  # session_id -> {temp_dir, arquivos}

    def _rotear_arquivo(nome: str, caminho_relativo: str = "") -> tuple:
        """Retorna (pasta_destino | None, tipo)."""
        nome_norm = utils.normalizar_texto(nome)

        if nome.startswith("~$") or nome_norm.startswith("conso"):
            return None, "ignorado"

        if "designacao" in nome_norm:
            return ".", "designacao_mensal"

        # Pasta reconhecível no caminho relativo?
        if caminho_relativo:
            for parte in Path(caminho_relativo).parts:
                parte_norm = utils.normalizar_texto(parte)
                for pf in _PASTAS_FORNECEDORES:
                    if utils.normalizar_texto(pf) == parte_norm:
                        return pf, "fornecedor"

        # Match pelo nome do arquivo
        for padrao, pasta in _EXCEL_PARA_PASTA:
            if padrao in nome_norm:
                ext = Path(nome).suffix.lower()
                return pasta, "fornecedor" if ext == ".xlsx" else "pdf"

        if Path(nome).suffix.lower() == ".pdf":
            return None, "pdf_desconhecido"

        return None, "desconhecido"

    def _validar_xlsx(caminho_arquivo: Path) -> str | None:
        """Tenta abrir o xlsx com openpyxl. Retorna None se OK, ou mensagem de erro."""
        try:
            from openpyxl import load_workbook
            wb = load_workbook(caminho_arquivo, read_only=True)
            wb.close()
            return None
        except Exception as exc:
            return str(exc)


    # ---------------------------------------------------------------------------
    # Helpers TICKET (compartilhados entre CLI e Web)
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


    # =========================================================================
    #  MODO CLI
    # =========================================================================
    def _resolver_pendencias_ticket_cli() -> None:
        """Se houver serviços do TICKET sem conta, pede ao usuário no terminal."""
        try:
            processor = _processor_ticket()
            caminhos = _caminhos_excel_ticket()
        except (ValueError, FileNotFoundError):
            return

        extrator = processor.extrator
        if not hasattr(extrator, "listar_servicos_pendentes"):
            return

        pendentes, vistos = [], set()
        for caminho in caminhos:
            for item in extrator.listar_servicos_pendentes(caminho):
                if item["chave"] not in vistos:
                    vistos.add(item["chave"])
                    pendentes.append(item)

        if not pendentes:
            return

        contas = config.contas_cadastradas()

        print("\n" + "=" * 60)
        print("  TICKET COMBUSTIVEL - Servicos sem conta cadastrada")
        print("=" * 60)
        print(f"\n  {len(pendentes)} servico(s) precisam de mapeamento.\n")
        print("  Contas disponiveis:")
        for c in contas:
            print(f"    {c['id']:>2}. {c['tipo_gasto']} - {c['conta']}")
        print()

        for i, p in enumerate(pendentes, 1):
            while True:
                try:
                    resp = input(
                        f"  [{i}/{len(pendentes)}] Servico: \"{p['servico_exibicao']}\"\n"
                        f"           ID da conta (1-{len(contas)}): "
                    ).strip()
                    conta_id = int(resp)
                    FornecedorProcessorFactory.criar("TICKET_COMBUSTIVEL").extrator.salvar_mapeamento(
                        p["servico"], conta_id,
                    )
                    tipo = config.tipo_gasto_por_id(conta_id)
                    print(f"           -> Salvo: {tipo}\n")
                    break
                except (ValueError, KeyError) as exc:
                    print(f"           [!] {exc}  Tente novamente.\n")

        print("  Todos os servicos mapeados. Prosseguindo...\n")


    def rodar_cli(fornecedores: List[str]) -> List[ResultadoProcessamento]:
        """Executa a conciliação no terminal."""
        print(f"\n  Mes de referencia: {config.MES_ATUAL}")
        print(f"  Pasta base: {config.CAMINHO_CONTROLE_DESPESAS_FR}\n")

        resultados: List[ResultadoProcessamento] = []
        total = len(fornecedores)

        with volume_io.volume_local():
            if "TICKET_COMBUSTIVEL" in fornecedores:
                _resolver_pendencias_ticket_cli()

            for idx, nome in enumerate(fornecedores, 1):
                print(f"[{idx}/{total}] {nome}...", end=" ")
                try:
                    processor = FornecedorProcessorFactory.criar(nome)
                    resultado = processor.processar()
                except Exception as exc:
                    print(f"ERRO: {exc}")
                    continue

                n_arq = len(resultado.arquivos_gerados)
                print(f"OK ({n_arq} arquivo{'s' if n_arq != 1 else ''})")
                for arquivo in resultado.arquivos_gerados:
                    print(f"      -> {arquivo}")
                resultados.append(resultado)

        _imprimir_relatorio(resultados)
        return resultados


    def _imprimir_relatorio(resultados: List[ResultadoProcessamento]) -> None:
        print("\n" + "-" * 60)
        print("  RELATORIO")
        print("-" * 60)

        placas = [(r.fornecedor, item["placa"]) for r in resultados for item in r.placas_sem_designacao]
        if placas:
            print(f"\n  Placas sem designacao ({len(placas)}):")
            for forn, placa in placas:
                print(f"    {forn:30s} | {placa}")
        else:
            print("\n  Nenhuma placa sem designacao.")

        tipos = [
            (r.fornecedor, item["tipo_transacao"], item["placa"])
            for r in resultados for item in r.tipos_nao_mapeados
        ]
        if tipos:
            print(f"\n  Tipos de transacao nao mapeados ({len(tipos)}):")
            for forn, tipo, placa in tipos:
                print(f"    {forn:30s} | {tipo:30s} | {placa}")

        print("\n" + "=" * 60)
        print(f"  Concluido: {len(resultados)} fornecedor(es) processado(s).")
        print("=" * 60 + "\n")


    # =========================================================================
    #  MODO WEB (servidor HTTP local — sem dependências externas)
    # =========================================================================
    def rodar_web(port: int = 5000) -> None:
        """Sobe um servidor HTTP local (stdlib) e abre o navegador."""
        import json as _json
        import mimetypes
        from functools import partial
        from http.server import HTTPServer, BaseHTTPRequestHandler
        from urllib.parse import unquote, urlparse, parse_qs

        _PASTA_PACOTE = Path(__file__).resolve().parent

        class _Handler(BaseHTTPRequestHandler):
            """Mini-router para as rotas da aplicação."""

            # ---------- helpers -------------------------------------------------
            def _json_response(self, data: dict, status: int = 200) -> None:
                body = _json.dumps(data, ensure_ascii=False).encode("utf-8")
                self.send_response(status)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def _read_body_json(self) -> dict:
                length = int(self.headers.get("Content-Length", 0))
                if length == 0:
                    return {}
                raw = self.rfile.read(length)
                try:
                    return _json.loads(raw)
                except _json.JSONDecodeError:
                    return {}

            def _serve_file(self, filepath: Path, content_type: str | None = None,
                            download_name: str | None = None) -> None:
                if not filepath.exists():
                    self.send_error(404)
                    return
                data = filepath.read_bytes()
                ct = content_type or mimetypes.guess_type(str(filepath))[0] or "application/octet-stream"
                self.send_response(200)
                self.send_header("Content-Type", ct)
                self.send_header("Content-Length", str(len(data)))
                if download_name:
                    self.send_header("Content-Disposition", f'attachment; filename="{download_name}"')
                self.end_headers()
                self.wfile.write(data)

            # ---------- GET -----------------------------------------------------
            def do_GET(self) -> None:
                parsed = urlparse(self.path)
                path = parsed.path.rstrip("/") or "/"

                if path == "/" or path == "/index.html":
                    self._serve_file(_PASTA_PACOTE / "index.html", "text/html; charset=utf-8")

                elif path == "/api/fornecedores":
                    self._json_response({
                        "mes_atual": config.MES_ATUAL,
                        "indice_mes": config.indice_mes_atual(),
                        "meses": config.meses_disponiveis(),
                        "fornecedores": FornecedorProcessorFactory.fornecedores_disponiveis(),
                        "pastas_fornecedores": _PASTAS_FORNECEDORES,
                    })

                elif path == "/api/download":
                    qs = parse_qs(parsed.query)
                    caminho = unquote(qs.get("path", [""])[0])
                    if not caminho:
                        self.send_error(400, "Parametro 'path' e obrigatorio")
                        return
                    ext = Path(caminho).suffix.lower()
                    if ext not in (".xlsx", ".pdf"):
                        self.send_error(404, "Tipo de arquivo nao permitido")
                        return
                    try:
                        conteudo = volume_io.download_bytes(caminho)
                    except Exception:
                        self.send_error(404, "Arquivo nao encontrado")
                        return
                    nome = Path(caminho).name
                    mime = (
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        if ext == ".xlsx"
                        else "application/pdf"
                    )
                    self.send_response(200)
                    self.send_header("Content-Type", mime)
                    self.send_header("Content-Length", str(len(conteudo)))
                    self.send_header("Content-Disposition", f'attachment; filename="{nome}"')
                    self.end_headers()
                    self.wfile.write(conteudo)

                elif path == "/api/arquivos":
                    qs = parse_qs(parsed.query)
                    mes_idx = qs.get("mes", [None])[0]
                    if mes_idx is None:
                        self.send_error(400, "Parametro 'mes' e obrigatorio")
                        return
                    try:
                        mes_idx = int(mes_idx)
                        nome_mes = config._MESES[mes_idx]
                        pasta_mes = config.CAMINHO_VOLUME / "FR - FROTAS" / config.ANO_ATUAL / f"{mes_idx + 1:02d}. {nome_mes}"
                    except (ValueError, IndexError):
                        self.send_error(400, "Indice de mes invalido")
                        return

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

                    # Incluir AUXILIARES
                    aux_str = str(config.CAMINHO_AUXILIARES)
                    for item in volume_io.listar_recursivo(aux_str):
                        arquivos.append({
                            "nome": item["nome"],
                            "pasta": "AUXILIARES",
                            "caminho": item["caminho"],
                            "tamanho": item["tamanho"],
                                })

                    self._json_response({
                        "mes": nome_mes,
                        "pasta": str(pasta_mes),
                        "arquivos": arquivos,
                    })

                else:
                    self.send_error(404)

            # ---------- POST ----------------------------------------------------
            def do_POST(self) -> None:
                parsed = urlparse(self.path)
                path = parsed.path.rstrip("/")

                if path == "/api/executar":
                    self._handle_executar()
                elif path == "/api/criar-oc":
                    self._handle_criar_oc()
                elif path == "/api/ticket/mapear-servico":
                    self._handle_mapear_ticket()
                elif path == "/api/alterar-mes":
                    self._handle_alterar_mes()
                elif path == "/api/upload":
                    self._handle_upload()
                elif path == "/api/upload/confirmar":
                    self._handle_upload_confirmar()
                else:
                    self.send_error(404)

            def _handle_executar(self) -> None:
                dados = self._read_body_json()
                lista_solicitada = dados.get("fornecedores") or []

                if lista_solicitada:
                    disponiveis = set(FornecedorProcessorFactory.fornecedores_disponiveis())
                    invalidos = [n for n in lista_solicitada if n not in disponiveis]
                    if invalidos:
                        return self._json_response({
                            "mes_atual": config.MES_ATUAL,
                            "resultados": [],
                            "erros": [{"fornecedor": n, "mensagem": f"'{n}' nao registrado"} for n in invalidos],
                        }, 400)
                    fornecedores = lista_solicitada
                else:
                    fornecedores = FornecedorProcessorFactory.fornecedores_disponiveis()

                resultados_json: list[dict] = []
                erros: list[dict] = []

                with volume_io.volume_local() as temp_root:
                    # Checar pendências TICKET dentro do context (precisa dos arquivos locais)
                    if "TICKET_COMBUSTIVEL" in fornecedores:
                        pend = _pendencias_ticket()
                        if pend:
                            return self._json_response({
                                "mes_atual": config.MES_ATUAL,
                                "pendencias_ticket": pend,
                                "contas_disponiveis": config.contas_cadastradas(),
                            }, 409)

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

                # Converter caminhos locais (temp) para caminhos do UC Volume
                temp_prefix = str(temp_root)
                vol_prefix = str(config.CAMINHO_VOLUME)
                for r in resultados_json:
                    r["arquivos_gerados"] = [
                        a.replace(temp_prefix, vol_prefix)
                        for a in r["arquivos_gerados"]
                    ]

                self._json_response({
                    "mes_atual":  config.MES_ATUAL,
                    "resultados": resultados_json,
                    "erros":      erros,
                })

            def _handle_criar_oc(self) -> None:
                dados = self._read_body_json()
                lista_solicitada = dados.get("fornecedores") or []

                if lista_solicitada:
                    disponiveis = set(FornecedorProcessorFactory.fornecedores_disponiveis())
                    invalidos = [n for n in lista_solicitada if n not in disponiveis]
                    if invalidos:
                        return self._json_response({
                            "resultados": [],
                            "erros": [{"fornecedor": n, "mensagem": f"'{n}' nao registrado"} for n in invalidos],
                        }, 400)
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

                self._json_response({
                    "resultados": resultados_json,
                    "erros": erros,
                })

            def _handle_mapear_ticket(self) -> None:
                dados = self._read_body_json()
                if "conta_id" not in dados:
                    self.send_error(400, "Campo 'conta_id' e obrigatorio")
                    return

                extrator = _processor_ticket().extrator
                if not hasattr(extrator, "salvar_mapeamento"):
                    self.send_error(500, "Extrator do TICKET nao suporta salvamento")
                    return

                extrator.salvar_mapeamento(dados.get("servico"), dados["conta_id"])
                self._json_response({
                    "pendencias_ticket": _pendencias_ticket(),
                    "contas_disponiveis": config.contas_cadastradas(),
                })

            # ---------- Upload de arquivos ------------------------------------
            def _handle_upload(self) -> None:
                """Recebe arquivos (base64), valida xlsx, roteia, salva em temp.
                Retorna análise com sessão para confirmação."""
                try:
                    self._handle_upload_impl()
                except Exception as exc:
                    self._json_response({"erro": f"Erro interno no upload: {exc}"}, 500)

            def _handle_upload_impl(self) -> None:
                dados = self._read_body_json()
                modo = dados.get("modo", "pasta")
                mes = dados.get("mes")
                arquivos_in = dados.get("arquivos", [])

                if not arquivos_in:
                    return self._json_response({"erro": "Nenhum arquivo enviado"}, 400)

                # Criar sessão temp
                session_id = str(uuid.uuid4())
                temp_dir = Path(tempfile.mkdtemp(prefix="fr_upload_"))

                # Calcular pasta-destino base
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

                    # Salvar em temp
                    temp_path = temp_dir / nome
                    try:
                        temp_path.write_bytes(base64.b64decode(dados_b64))
                    except Exception as exc:
                        resultados.append({"nome": nome, "valido": False, "erro": f"Erro ao decodificar: {exc}"})
                        continue

                    # Validar xlsx
                    erro_validacao = None
                    if nome.lower().endswith(".xlsx"):
                        erro_validacao = _validar_xlsx(temp_path)

                    # Rotear
                    if modo == "designacao_mensal":
                        destino_rel, tipo = ".", "designacao_mensal"
                    elif modo == "designacao_historica":
                        destino_rel, tipo = ".", "designacao_historica"
                    else:
                        destino_rel, tipo = _rotear_arquivo(nome, caminho_rel)

                    # Caminho final
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
                        "erro": f"Arquivo n\u00e3o adicionado: remova o status do excel {nome} de interno para publico" if erro_validacao else None,
                        "conflito": conflito,
                        "temp": str(temp_path),
                    })

                _upload_sessions[session_id] = {
                    "temp_dir": str(temp_dir),
                    "resultados": resultados,
                    "modo": modo,
                }

                self._json_response({
                    "sessao": session_id,
                    "resultados": [
                        {k: v for k, v in r.items() if k != "temp"}
                        for r in resultados
                    ],
                })

            def _handle_upload_confirmar(self) -> None:
                """Move arquivos validados de temp para destinos finais."""
                try:
                    self._handle_upload_confirmar_impl()
                except Exception as exc:
                    self._json_response({"erro": f"Erro interno ao confirmar upload: {exc}"}, 500)

            def _handle_upload_confirmar_impl(self) -> None:
                dados = self._read_body_json()
                session_id = dados.get("sessao")
                substituir = set(dados.get("substituir", []))

                session = _upload_sessions.pop(session_id, None)
                if not session:
                    return self._json_response({"erro": "Sess\u00e3o expirada ou inv\u00e1lida"}, 400)

                sucesso = []
                erros = []

                for r in session["resultados"]:
                    nome = r["nome"]

                    if not r["valido"]:
                        erros.append({"nome": nome, "erro": r["erro"]})
                        continue

                    if r["destino"] is None:
                        erros.append({"nome": nome, "erro": "Destino n\u00e3o identificado"})
                        continue

                    destino = Path(r["destino"])
                    temp_path = Path(r["temp"])

                    if volume_io.arquivo_existe(str(destino)) and nome not in substituir:
                        erros.append({"nome": nome, "erro": "Conflito n\u00e3o resolvido (arquivo j\u00e1 existe)"})
                        continue

                    try:
                        volume_io.upload_arquivo(temp_path, str(destino))
                        sucesso.append({"nome": nome, "destino": str(destino)})
                    except Exception as exc:
                        erros.append({"nome": nome, "erro": str(exc)})

                # Limpar temp
                try:
                    shutil.rmtree(session["temp_dir"])
                except Exception:
                    pass

                self._json_response({"sucesso": sucesso, "erros": erros})

            def _handle_alterar_mes(self) -> None:
                dados = self._read_body_json()
                indice = dados.get("indice")
                if indice is None:
                    self.send_error(400, "Campo 'indice' e obrigatorio")
                    return
                try:
                    config.alterar_mes(int(indice))
                except ValueError as exc:
                    self._json_response({"erro": str(exc)}, 400)
                    return
                self._json_response({
                    "mes_atual": config.MES_ATUAL,
                    "indice_mes": config.indice_mes_atual(),
                    "pasta": str(config.CAMINHO_CONTROLE_DESPESAS_FR),
                })

            # ---------- silenciar logs do BaseHTTPRequestHandler ----------------
            def log_message(self, fmt, *args) -> None:
                pass  # silencia GET/POST logs repetitivos

        # -- Iniciar -----------------------------------------------------------
        # Garantir estrutura de pastas do ano atual
        volume_io.garantir_estrutura_ano()

        n_forn = len(FornecedorProcessorFactory.fornecedores_disponiveis())
        print(f"\n  FR Frotas \u2014 Interface web em http://localhost:{port}")
        print(f"  Mes vigente: {config.MES_ATUAL}")
        print(f"  Fornecedores: {n_forn}")
        print(f"  Pasta: {config.CAMINHO_CONTROLE_DESPESAS_FR}")
        print(f"  Ctrl+C para encerrar\n")

        webbrowser.open(f"http://localhost:{port}")
        servidor = HTTPServer(("127.0.0.1", port), _Handler)
        try:
            servidor.serve_forever()
        except KeyboardInterrupt:
            print("\n  Servidor encerrado.")
            servidor.server_close()


    # =========================================================================
    #  Entrypoint
    # =========================================================================
    if __name__ == "__main__":
        args = sys.argv[1:]

        if "--cli" in args:
            args.remove("--cli")
            fornecedores = args if args else FornecedorProcessorFactory.fornecedores_disponiveis()
            rodar_cli(fornecedores)
        else:
            port = 5000
            if "--port" in args:
                idx = args.index("--port")
                port = int(args[idx + 1])
            rodar_web(port)
