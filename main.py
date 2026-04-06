import os
import signal
import sys
import threading
import time
import logging

from dotenv import dotenv_values
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

from distribuidor import DistribuidorArquivos


def obter_base_dir() -> str:
    # Quando empacotado em EXE, os arquivos externos passam a ser lidos ao lado do executavel.
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def carregar_config(base_dir: str) -> dict:
    caminho_config = os.path.join(base_dir, "config.env")

    if not os.path.exists(caminho_config):
        raise FileNotFoundError(f"Arquivo config.env nao encontrado em {caminho_config}.")

    config = dotenv_values(caminho_config)
    chaves_obrigatorias = ["PASTA_ENTRADA", "PASTA_EXCEL", "PASTA_RELATORIOS"]
    faltantes = [chave for chave in chaves_obrigatorias if not str(config.get(chave, "")).strip()]

    if faltantes:
        lista = ", ".join(faltantes)
        raise ValueError(f"As seguintes variaveis obrigatorias nao foram definidas no config.env: {lista}.")

    return config


def configurar_logging(base_dir: str) -> logging.Logger:
    # O log diario e escrito em arquivo e tambem exibido no console.
    pasta_logs = os.path.join(base_dir, "logs")
    os.makedirs(pasta_logs, exist_ok=True)

    logger = logging.getLogger("distribuidor_arquivos")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    logger.propagate = False

    formato = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    arquivo_log = os.path.join(pasta_logs, f"distribuidor_{time.strftime('%Y-%m-%d')}.log")

    handler_console = logging.StreamHandler(sys.stdout)
    handler_console.setFormatter(formato)

    handler_arquivo = logging.FileHandler(arquivo_log, encoding="utf-8")
    handler_arquivo.setFormatter(formato)

    logger.addHandler(handler_console)
    logger.addHandler(handler_arquivo)
    return logger


class ManipuladorEventos(FileSystemEventHandler):
    def __init__(self, distribuidor: DistribuidorArquivos, logger: logging.Logger):
        super().__init__()
        self.distribuidor = distribuidor
        self.logger = logger
        self._lock = threading.Lock()
        self._em_processamento = set()

    def _esta_na_pasta_monitorada(self, caminho: str) -> bool:
        pasta_monitorada = os.path.normcase(os.path.normpath(self.distribuidor.pasta_entrada))
        pasta_arquivo = os.path.normcase(os.path.normpath(os.path.dirname(caminho)))
        return pasta_monitorada == pasta_arquivo

    def _ignorar_arquivo(self, caminho: str) -> bool:
        nome = os.path.basename(caminho)

        if not nome:
            return True

        if nome.startswith("~$"):
            return True

        pasta_nao_identificados = os.path.normcase(os.path.normpath(self.distribuidor.pasta_nao_identificados))
        pasta_arquivo = os.path.normcase(os.path.normpath(os.path.dirname(caminho)))
        if pasta_arquivo == pasta_nao_identificados:
            return True

        return False

    def _processar_em_thread(self, caminho: str) -> None:
        caminho_normalizado = os.path.normcase(os.path.normpath(caminho))

        with self._lock:
            if caminho_normalizado in self._em_processamento:
                return
            self._em_processamento.add(caminho_normalizado)

        def alvo():
            try:
                self.logger.info("Novo arquivo detectado: %s", caminho)
                self.distribuidor.processar_arquivo(caminho)
            finally:
                with self._lock:
                    self._em_processamento.discard(caminho_normalizado)

        thread = threading.Thread(target=alvo, daemon=True)
        thread.start()

    def on_created(self, event):
        if event.is_directory:
            return

        if self._ignorar_arquivo(event.src_path):
            return

        if self._esta_na_pasta_monitorada(event.src_path):
            self._processar_em_thread(event.src_path)

    def on_moved(self, event):
        if event.is_directory:
            return

        if self._ignorar_arquivo(event.dest_path):
            return

        if self._esta_na_pasta_monitorada(event.dest_path):
            self._processar_em_thread(event.dest_path)


def executar() -> int:
    base_dir = obter_base_dir()
    logger = configurar_logging(base_dir)

    try:
        # Carrega configuracao e regras iniciais para validar o ambiente na partida.
        config = carregar_config(base_dir)
        distribuidor = DistribuidorArquivos(
            base_dir=base_dir,
            pasta_entrada=config["PASTA_ENTRADA"],
            pasta_excel=config["PASTA_EXCEL"],
            pasta_relatorios=config["PASTA_RELATORIOS"],
            logger=logger,
        )

        resumo = distribuidor.obter_resumo_regras()
        logger.info(
            "Empresas carregadas: %s | Rotas carregadas: %s | Palavras-chave carregadas: %s",
            resumo.total_empresas,
            resumo.total_rotas,
            resumo.total_palavras_chave,
        )
        logger.info("Monitorando pasta de entrada em %s", distribuidor.pasta_entrada)
    except Exception as erro:
        logger.exception("Falha ao iniciar o sistema: %s", erro)
        return 1

    observer = Observer()
    handler = ManipuladorEventos(distribuidor, logger)

    # O watchdog monitora apenas a pasta de entrada principal em tempo real.
    observer.schedule(handler, distribuidor.pasta_entrada, recursive=False)
    observer.start()

    evento_parada = threading.Event()

    def encerrar(*_args):
        if evento_parada.is_set():
            return

        # O encerramento seguro para o observer e dispara a exportacao do relatorio pendente.
        evento_parada.set()
        logger.info("Sinal de encerramento recebido. Finalizando monitoramento.")
        observer.stop()

    signal.signal(signal.SIGINT, encerrar)
    signal.signal(signal.SIGTERM, encerrar)

    try:
        while not evento_parada.is_set():
            time.sleep(1)
    except KeyboardInterrupt:
        encerrar()
    finally:
        observer.stop()
        observer.join()
        distribuidor.encerrar()

    return 0


if __name__ == "__main__":
    raise SystemExit(executar())