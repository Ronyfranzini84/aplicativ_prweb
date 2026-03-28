from dataclasses import dataclass
import time
from typing import Callable, Optional

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


@dataclass
class PRWebConfig:
    empresa_gateway: int
    empresa_pr: int
    login: int
    senha: str
    filial: int
    atividade: str = "D"
    motivo: int = 35
    carga: int = 16335512
    url: str = "https://prweb01/bahia/gateway"


class PRWebAutomation:
    def __init__(self, config: PRWebConfig, status_callback: Optional[Callable[[str], None]] = None):
        self.config = config
        self.status_callback = status_callback
        self.driver = None

    def _status(self, message: str) -> None:
        if self.status_callback:
            self.status_callback(message)

    def _start_driver(self) -> None:
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service)
        self.driver.implicitly_wait(2)
        self.driver.maximize_window()
        self.driver.get(self.config.url)

    def _login_prweb(self) -> None:
        self.driver.find_element(By.XPATH, "/html/body/form[1]/table[1]/tbody/tr[1]/td[1]/input[1]").send_keys(
            self.config.empresa_gateway
        )
        self.driver.find_element(By.XPATH, "/html/body/form[1]/table[1]/tbody/tr[1]/td[2]/input[1]").send_keys(
            self.config.login
        )
        self.driver.find_element(By.XPATH, "/html/body/form[1]/table[1]/tbody/tr[1]/td[3]/b[1]/input").send_keys(
            self.config.senha
        )
        self.driver.find_element(By.XPATH, "/html/body/form[1]/table[2]/tbody/tr/td[2]/input").click()

    def _abrir_tela_pr(self) -> None:
        time.sleep(5)
        self.driver.find_element(By.XPATH, "/html/body/form[1]").click()
        select_element = self.driver.find_element(By.NAME, "U01_DS_APL_VIS_SLC")
        select = Select(select_element)
        select.select_by_value("9")
        self.driver.find_element(By.XPATH, "/html/body/form[1]/table[2]/tbody/tr/td[2]/input").click()
        time.sleep(5)

    def _login_pr(self) -> None:
        campo_empresa = self.driver.find_element(By.XPATH, "/html/body/form[1]/table[2]/tbody/tr/td[2]/input[1]")
        campo_empresa.clear()
        campo_empresa.send_keys(self.config.empresa_pr)
        self.driver.find_element(By.XPATH, "/html/body/form[1]/table[2]/tbody/tr/td[3]/input[1]").send_keys(self.config.filial)
        self.driver.find_element(By.XPATH, "/html/body/form[1]/table[2]/tbody/tr/td[4]/input[1]").send_keys(self.config.atividade)

    def _abrir_fluxo_transferencia(self) -> None:
        self.driver.find_element(By.XPATH, "/html/body/form[1]/table[3]/tbody/tr/td[1]/table/tbody/tr[11]/td[1]/input").click()
        self.driver.find_element(By.XPATH, '//*[@id="NM_BOT_CON"]').click()
        time.sleep(5)

        self.driver.find_element(By.XPATH, "/html/body/form[1]/table[3]/tbody/tr[5]/td[1]/input").click()
        self.driver.find_element(By.XPATH, '//*[@id="NM_BOT_PRC"]').click()
        time.sleep(3)

        self.driver.find_element(By.XPATH, '//*[@id="NM_BOT_TRA"]').click()
        time.sleep(1)
        self.driver.find_element(By.XPATH, '//*[@id="NM_BOT_TRA"]').click()
        time.sleep(1)

    def _preencher_cabecalho(self) -> None:
        campo_empresa = self.driver.find_element(By.XPATH, "/html/body/form[1]/table[3]/tbody/tr/td[5]/input")
        campo_empresa.click()
        campo_empresa.clear()
        campo_empresa.send_keys(self.config.empresa_gateway)
        self.driver.find_element(By.XPATH, "/html/body/form[1]/table[3]/tbody/tr/td[6]/input").send_keys(self.config.login)

    def run(self, pedidos_df: pd.DataFrame) -> pd.DataFrame:
        if "Pedidos" not in pedidos_df.columns or "desm" not in pedidos_df.columns:
            raise ValueError("A planilha deve conter as colunas 'Pedidos' e 'desm'.")

        pedidos = pedidos_df.copy()
        if "msg" not in pedidos.columns:
            pedidos["msg"] = ""

        self._status("Inicializando navegador...")
        self._start_driver()
        try:
            self._status("Realizando login no gateway...")
            self._login_prweb()
            self._abrir_tela_pr()
            self._login_pr()
            self._abrir_fluxo_transferencia()
            self._preencher_cabecalho()

            total = len(pedidos)
            for idx, (pedido, desm) in enumerate(zip(pedidos["Pedidos"], pedidos["desm"])):
                if pd.isna(desm) or str(desm).strip() == "":
                    pedidos.at[idx, "msg"] = "DESM vazio - linha ignorada"
                    continue

                self._status(f"Processando pedido {idx + 1}/{total}...")
                try:
                    self.driver.find_element(By.XPATH, "/html/body/form[1]/table[3]/tbody/tr/td[7]/b/input").send_keys(self.config.senha)
                    self.driver.find_element(By.XPATH, "/html/body/form[1]/table[4]/tbody/tr[1]/td[2]/input[1]").send_keys(pedido)
                    self.driver.find_element(By.XPATH, "/html/body/form[1]/table[4]/tbody/tr[1]/td[2]/input[2]").send_keys(desm)

                    self.driver.find_element(By.XPATH, '//*[@id="NM_BOT_PRC"]').click()
                    time.sleep(1)

                    self.driver.find_element(By.XPATH, "/html/body/form[1]/table[5]/tbody/tr[7]/td[2]/input[1]").send_keys(self.config.motivo)
                    self.driver.find_element(By.XPATH, "/html/body/form[1]/table[6]/tbody/tr/td/table/tbody/tr[1]/td[4]/input").send_keys(self.config.carga)

                    time.sleep(1)
                    self.driver.find_element(By.XPATH, '//*[@id="NM_BOT_PRC"]').click()
                    time.sleep(1)

                    try:
                        alerta = WebDriverWait(self.driver, 10).until(EC.alert_is_present())
                        texto_alerta = alerta.text
                        alerta.accept()
                        pedidos.at[idx, "msg"] = texto_alerta
                    except TimeoutException:
                        pedidos.at[idx, "msg"] = "Nenhum alerta apareceu"

                    self.driver.find_element(By.XPATH, '//*[@id="NM_BOT_LIM"]').click()
                    time.sleep(1)
                except Exception as exc:
                    pedidos.at[idx, "msg"] = f"Erro no processamento: {exc}"

            self._status("Processamento concluido.")
            return pedidos
        finally:
            if self.driver:
                self.driver.quit()


def carregar_dataframe(caminho_arquivo: str) -> pd.DataFrame:
    return pd.read_excel(caminho_arquivo)


def executar_transferencia(
    pedidos_df: pd.DataFrame,
    config: PRWebConfig,
    status_callback: Optional[Callable[[str], None]] = None,
) -> pd.DataFrame:
    automacao = PRWebAutomation(config=config, status_callback=status_callback)
    return automacao.run(pedidos_df)


if __name__ == "__main__":
    # Mantem um modo script para uso rapido em ambiente legado.
    configuracao = PRWebConfig(
        empresa_gateway=29,
        empresa_pr=21,
        login=2634333,
        senha="qwer7410",
        filial=1400,
        atividade="D",
        motivo=35,
        carga=16335512,
    )

    entrada = "pedidos.xlsx"
    saida = "pedidos_resultado.xlsx"

    df_pedidos = carregar_dataframe(entrada)
    df_resultado = executar_transferencia(df_pedidos, configuracao, status_callback=print)
    df_resultado.to_excel(saida, index=False)
    print(f"Resultados salvos em: {saida}")