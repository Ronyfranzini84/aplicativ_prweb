from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import time
from pathlib import Path
from typing import Callable

import pandas as pd
from selenium.common.exceptions import (
    InvalidSessionIdException,
    TimeoutException,
    UnexpectedAlertPresentException,
    WebDriverException,
)
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver as ChromeDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


@dataclass(slots=True)
class PRWebConfig:
    empresa_gateway: int
    empresa_pr: int
    login: int
    senha: str
    filial: int
    atividade: str = "D"
    motivo: int = 35
    carga: int = 0
    url: str = "https://prweb01/bahia/gateway"
    webdriver_wait: int = 10
    alert_wait: int = 3


def _garantir_openpyxl() -> None:
    try:
        import openpyxl  # noqa: F401
    except ImportError as exc:
        raise RuntimeError(
            "O suporte a arquivos .xlsx requer a biblioteca 'openpyxl'. "
            "Instale com: pip install openpyxl"
        ) from exc


def carregar_dataframe(caminho_arquivo: str) -> pd.DataFrame:
    _garantir_openpyxl()
    return pd.read_excel(caminho_arquivo, engine="openpyxl")


def salvar_dataframe(df: pd.DataFrame, caminho_arquivo: str) -> str:
    _garantir_openpyxl()
    caminho = Path(caminho_arquivo)
    caminho.parent.mkdir(parents=True, exist_ok=True)

    try:
        df.to_excel(caminho, index=False, engine="openpyxl")
        return str(caminho)
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = caminho.with_name(f"{caminho.stem}_{timestamp}{caminho.suffix or '.xlsx'}")
        df.to_excel(fallback, index=False, engine="openpyxl")
        return str(fallback)


class PRWebAutomation:
    def __init__(self, config: PRWebConfig, status_callback: Callable[[str], None] | None = None):
        self.config = config
        self.status_callback = status_callback
        self.driver: WebDriver | None = None
        self.wait: WebDriverWait | None = None

    def _status(self, message: str) -> None:
        if self.status_callback:
            self.status_callback(message)

    def _require_driver(self) -> WebDriver:
        if self.driver is None:
            raise RuntimeError("Driver do navegador nao foi inicializado.")
        return self.driver

    def _require_wait(self) -> WebDriverWait:
        if self.wait is None:
            raise RuntimeError("WebDriverWait nao foi inicializado.")
        return self.wait

    def _start_driver(self) -> None:
        service = Service(ChromeDriverManager().install())
        driver = ChromeDriver(service=service)
        driver.implicitly_wait(2)
        driver.maximize_window()
        driver.get(self.config.url)
        self.driver = driver
        self.wait = WebDriverWait(driver, self.config.webdriver_wait)

    def _capturar_alertas(self, timeout: int | None = None) -> list[str]:
        driver = self._require_driver()
        mensagens: list[str] = []
        espera_alerta = WebDriverWait(driver, timeout or self.config.alert_wait)

        while True:
            try:
                espera_alerta.until(EC.alert_is_present())
                alerta = driver.switch_to.alert
                mensagens.append(alerta.text)
                alerta.accept()
                time.sleep(0.5)
            except TimeoutException:
                break
            except Exception as exc:
                mensagens.append(f"EXCEPTION: {exc}")
                break

        return mensagens

    @staticmethod
    def _normalizar_dataframe(pedidos_df: pd.DataFrame) -> pd.DataFrame:
        if "Pedidos" not in pedidos_df.columns or "desm" not in pedidos_df.columns:
            raise ValueError("A planilha deve conter as colunas 'Pedidos' e 'desm'.")

        pedidos = pedidos_df.copy()
        if "msg" not in pedidos.columns:
            pedidos["msg"] = ""
        if "status" not in pedidos.columns:
            pedidos["status"] = ""

        pedidos["Pedidos"] = (
            pedidos["Pedidos"]
            .astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )
        pedidos["desm"] = (
            pedidos["desm"]
            .astype(str)
            .str.strip()
            .replace({"nan": "", "None": ""})
            .str.replace(r"\.0$", "", regex=True)
        )
        return pedidos

    def _login_gateway(self) -> None:
        driver = self._require_driver()
        wait = self._require_wait()

        self._status("Abrindo login do gateway...")
        campo_emp_login = wait.until(EC.presence_of_element_located((By.NAME, "CD_EMPGCB_FUN")))
        campo_emp_login.clear()
        campo_emp_login.send_keys(str(self.config.empresa_gateway))

        campo_login = driver.find_element(By.NAME, "CD_FUN")
        campo_login.clear()
        campo_login.send_keys(str(self.config.login))

        campo_senha = driver.find_element(By.NAME, "CD_USRSGR_SNH_CPL")
        campo_senha.clear()
        campo_senha.send_keys(self.config.senha)

        driver.find_element(By.NAME, "NM_BOT_PRC").click()

    def _abrir_fluxo_transferencia(self) -> None:
        driver = self._require_driver()
        wait = self._require_wait()

        self._status("Navegando ate Cargas Fracionadas...")
        select_element = wait.until(EC.presence_of_element_located((By.NAME, "U01_DS_APL_VIS_SLC")))
        Select(select_element).select_by_visible_text("Roteirizacao")
        driver.find_element(By.NAME, "NM_BOT_PRC").click()

        campo_empresa = wait.until(EC.presence_of_element_located((By.NAME, "CD_EMPGCB")))
        campo_empresa.clear()
        campo_empresa.send_keys(str(self.config.empresa_pr))

        campo_filial = driver.find_element(By.NAME, "CD_FIL")
        campo_filial.clear()
        campo_filial.send_keys(str(self.config.filial))

        campo_atividade = driver.find_element(By.NAME, "CD_TAFIL")
        campo_atividade.clear()
        campo_atividade.send_keys(self.config.atividade)

        carga_entrega = driver.find_element(By.XPATH, "//input[@name='VR_RDO_SLC' and @value='11']")
        driver.execute_script("arguments[0].click();", carga_entrega)
        driver.find_element(By.ID, "NM_BOT_CON").click()

        cargas_fracionadas = wait.until(
            EC.presence_of_element_located((By.XPATH, "//input[@name='U01_VR_RDO_SLC' and @value='4']"))
        )
        driver.execute_script("arguments[0].click();", cargas_fracionadas)
        driver.find_element(By.ID, "NM_BOT_PRC").click()

        wait.until(EC.element_to_be_clickable((By.ID, "NM_BOT_TRA"))).click()
        wait.until(EC.element_to_be_clickable((By.ID, "NM_BOT_TRA"))).click()

        matricula = wait.until(EC.presence_of_element_located((By.NAME, "CD_EMPGCB_FUN")))
        matricula.clear()
        matricula.send_keys(str(self.config.empresa_gateway))

        campo_login = wait.until(EC.presence_of_element_located((By.NAME, "CD_FUN")))
        campo_login.clear()
        campo_login.send_keys(str(self.config.login))

        campo_senha = wait.until(EC.presence_of_element_located((By.NAME, "CD_USRSGR_SNH_CPL")))
        campo_senha.clear()
        campo_senha.send_keys(self.config.senha)

    def _limpar_tela(self) -> None:
        driver = self._require_driver()
        try:
            driver.find_element(By.ID, "NM_BOT_LIM").click()
            time.sleep(1)
        except Exception:
            pass

    def _processar_linha(self, pedidos: pd.DataFrame, index: int, pedido: str, desm: str) -> None:
        driver = self._require_driver()
        wait = self._require_wait()

        desm_limpo = str(desm).strip()
        if not desm_limpo or desm_limpo.lower() == "nan":
            desm_limpo = "0"
            pedidos.at[index, "desm"] = "0"

        campo_senha = wait.until(EC.presence_of_element_located((By.NAME, "CD_USRSGR_SNH_CPL")))
        campo_senha.clear()
        campo_senha.send_keys(self.config.senha)
   
        campo_pedido = driver.find_element(By.NAME, "CD_DOCTO")
        campo_pedido.clear()
        campo_pedido.send_keys(pedido)

        campo_desm = driver.find_element(By.NAME, "CD_PVEMCR_DSM")
        campo_desm.clear()
        campo_desm.send_keys(desm_limpo)
  
        driver.find_element(By.ID, "NM_BOT_PRC").click()
       

        mensagens = self._capturar_alertas(timeout=2)
        if mensagens:
            pedidos.at[index, "msg"] = " | ".join(mensagens)
            self._limpar_tela()
            return

        campo_motivo = driver.find_element(By.NAME, "CD_MMCETG")
        campo_motivo.clear()
        campo_motivo.send_keys(str(self.config.motivo))
        

        campo_carga = driver.find_element(By.NAME, "CD_CGAETG_DST")
        campo_carga.clear()
        campo_carga.send_keys(str(self.config.carga))
    

        driver.find_element(By.ID, "NM_BOT_PRC").click()
        time.sleep(1)

        mensagens = self._capturar_alertas(timeout=3)
        if mensagens:
            pedidos.at[index, "msg"] = " | ".join(mensagens)
        else:
            pedidos.at[index, "msg"] = "Processado sem erro"

        self._limpar_tela()

    def run(self, pedidos_df: pd.DataFrame) -> pd.DataFrame:
        pedidos = self._normalizar_dataframe(pedidos_df)

        self._status("Inicializando navegador...")
        self._start_driver()

        try:
            self._login_gateway()
            self._abrir_fluxo_transferencia()

            total = len(pedidos)
            for index, (pedido, desm) in enumerate(zip(pedidos["Pedidos"], pedidos["desm"])):
                self._status(f"Processando pedido {index + 1}/{total}...")
                try:
                    self._processar_linha(pedidos, index, str(pedido), str(desm))
                except UnexpectedAlertPresentException:
                    mensagens = self._capturar_alertas(timeout=3)
                    pedidos.at[index, "msg"] = (
                        " | ".join(mensagens)
                        if mensagens
                        else "Alerta inesperado durante processamento"
                    )
                    self._limpar_tela()
                except InvalidSessionIdException as exc:
                    pedidos.at[index, "msg"] = f"Sessao encerrada: {exc}"
                    break
                except WebDriverException as exc:
                    mensagens = self._capturar_alertas(timeout=2)
                    pedidos.at[index, "msg"] = " | ".join(mensagens) if mensagens else f"Erro Selenium: {exc}"
                    self._limpar_tela()
                except Exception as exc:
                    pedidos.at[index, "msg"] = f"Erro no processamento: {exc}"
                    self._limpar_tela()

            self._status("Processamento concluido.")
            return pedidos
        finally:
            if self.driver is not None:
                self.driver.quit()
                self.driver = None
                self.wait = None


def executar_transferencia(
    pedidos_df: pd.DataFrame,
    config: PRWebConfig,
    status_callback: Callable[[str], None] | None = None,
) -> pd.DataFrame:
    automacao = PRWebAutomation(config=config, status_callback=status_callback)
    return automacao.run(pedidos_df)


def _parse_args():
    import argparse

    parser = argparse.ArgumentParser(
        description="Processa transferencias de pedidos no PRWeb via Selenium.",
        epilog="Exemplo: python cargas.py pedidos.xlsx --carga 17404047 --filial 1200",
    )
    parser.add_argument("entrada", nargs="?", default="pedidos.xlsx",
                        help="Planilha .xlsx de entrada (default: pedidos.xlsx)")
    parser.add_argument("--saida", default="pedidos_resultado.xlsx",
                        help="Arquivo .xlsx de saida (default: pedidos_resultado.xlsx)")
    parser.add_argument("--empresa-gateway", type=int, default=29, metavar="N")
    parser.add_argument("--empresa-pr", type=int, default=21, metavar="N")
    parser.add_argument("--login", type=int, default=3471500, metavar="N")
    parser.add_argument("--senha", default="VAREJO1289")
    parser.add_argument("--filial", type=int, default=1200, metavar="N")
    parser.add_argument("--atividade", default="D")
    parser.add_argument("--motivo", type=int, default=35, metavar="N")
    parser.add_argument("--carga", type=int, default=17404047, metavar="N")
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()

    entrada = Path(args.entrada).expanduser()
    saida = Path(args.saida).expanduser()

    if not entrada.exists():
        raise SystemExit(
            f"Arquivo de entrada nao encontrado: {entrada}\n"
            "Use 'python cargas.py --help' para ver as opcoes ou execute o app_pyside6.py."
        )

    configuracao = PRWebConfig(
        empresa_gateway=args.empresa_gateway,
        empresa_pr=args.empresa_pr,
        login=args.login,
        senha=args.senha,
        filial=args.filial,
        atividade=args.atividade,
        motivo=args.motivo,
        carga=args.carga,
    )

    df_pedidos = carregar_dataframe(str(entrada))
    df_resultado = executar_transferencia(df_pedidos, configuracao, status_callback=print)
    caminho_salvo = salvar_dataframe(df_resultado, str(saida))
    print(f"Resultados salvos em: {caminho_salvo}")