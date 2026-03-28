from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
from PySide6.QtCore import QObject, Qt, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from cargas import PRWebConfig, executar_transferencia


class ProcessWorker(QObject):
    finished = Signal(object)
    error = Signal(str)
    progress = Signal(str)

    def __init__(self, dataframe: pd.DataFrame, config: PRWebConfig):
        super().__init__()
        self.dataframe = dataframe
        self.config = config

    def run(self) -> None:
        try:
            resultado = executar_transferencia(
                pedidos_df=self.dataframe,
                config=self.config,
                status_callback=self.progress.emit,
            )
            self.finished.emit(resultado)
        except Exception as exc:
            self.error.emit(str(exc))


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("App PRWeb - Cargas")
        self.resize(980, 700)

        self.signature_text = "© 2026 Rony Franzini"

        self.df_entrada: pd.DataFrame | None = None
        self.df_resultado: pd.DataFrame | None = None

        self.thread: QThread | None = None
        self.worker: ProcessWorker | None = None

        self.input_path_edit = QLineEdit()
        self.input_path_edit.setPlaceholderText("Selecione um arquivo .xlsx")

        self.output_path_edit = QLineEdit(str(Path.cwd() / "pedidos_resultado.xlsx"))

        self.empresa_gateway_spin = self._spinbox(0, 999999999, 29)
        self.empresa_pr_spin = self._spinbox(0, 999999999, 21)
        self.login_spin = self._spinbox(0, 999999999, 2634333)
        self.senha_edit = QLineEdit("qwer7410")
        self.senha_edit.setEchoMode(QLineEdit.Password)
        self.filial_spin = self._spinbox(0, 999999999, 1400)
        self.atividade_edit = QLineEdit("D")
        self.motivo_spin = self._spinbox(0, 999999999, 35)
        self.carga_spin = self._spinbox(0, 999999999, 16335512)

        self.only_msg_check = QCheckBox("Exportar apenas linhas com mensagem")

        self.import_btn = QPushButton("Importar Planilha")
        self.process_btn = QPushButton("Processar no PRWeb")
        self.export_btn = QPushButton("Exportar Resultado")
        self.export_btn.setEnabled(False)

        self.table = QTableWidget(0, 0)
        self.signature_label = QLabel(self.signature_text)
        self.signature_label.setObjectName("signatureLabel")

        self._build_ui()
        self._apply_styles()
        self._connect_signals()
        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage("Pronto")

    @staticmethod
    def _spinbox(minimum: int, maximum: int, value: int) -> QSpinBox:
        box = QSpinBox()
        box.setRange(minimum, maximum)
        box.setValue(value)
        return box

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)

        title = QLabel("ORGANIZADOR DE CARGAS PRWEB")
        title.setObjectName("mainTitle")
        title.setAlignment(Qt.AlignHCenter)

        arquivo_group = QGroupBox("Arquivos")
        arquivo_layout = QGridLayout(arquivo_group)

        browse_input_btn = QPushButton("...")
        browse_output_btn = QPushButton("...")
        browse_input_btn.clicked.connect(self._select_input_file)
        browse_output_btn.clicked.connect(self._select_output_file)

        arquivo_layout.addWidget(QLabel("Entrada (.xlsx):"), 0, 0)
        arquivo_layout.addWidget(self.input_path_edit, 0, 1)
        arquivo_layout.addWidget(browse_input_btn, 0, 2)

        arquivo_layout.addWidget(QLabel("Saida (.xlsx):"), 1, 0)
        arquivo_layout.addWidget(self.output_path_edit, 1, 1)
        arquivo_layout.addWidget(browse_output_btn, 1, 2)

        config_group = QGroupBox("Configuracao PRWeb")
        config_layout = QFormLayout(config_group)
        config_layout.addRow("Empresa Gateway:", self.empresa_gateway_spin)
        config_layout.addRow("Empresa PR:", self.empresa_pr_spin)
        config_layout.addRow("Login:", self.login_spin)
        config_layout.addRow("Senha:", self.senha_edit)
        config_layout.addRow("Filial:", self.filial_spin)
        config_layout.addRow("Atividade:", self.atividade_edit)
        config_layout.addRow("Motivo:", self.motivo_spin)
        config_layout.addRow("Carga:", self.carga_spin)

        buttons_layout = QHBoxLayout()
        buttons_layout.addWidget(self.import_btn)
        buttons_layout.addWidget(self.process_btn)
        buttons_layout.addWidget(self.export_btn)

        footer_layout = QHBoxLayout()
        footer_layout.addWidget(self.signature_label)
        footer_layout.addStretch(1)

        root_layout.addWidget(title)
        root_layout.addWidget(arquivo_group)
        root_layout.addWidget(config_group)
        root_layout.addWidget(self.only_msg_check)
        root_layout.addLayout(buttons_layout)
        root_layout.addWidget(self.table)
        root_layout.addLayout(footer_layout)

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 #9adcf4,
                    stop:0.5 #45b2df,
                    stop:1 #1b82be
                );
            }
            QWidget {
                color: #f8fbff;
                font-family: 'Segoe UI';
                font-size: 12px;
            }
            QLabel#mainTitle {
                font-size: 24px;
                font-weight: 700;
                color: #ffffff;
                padding: 8px 0 6px 0;
            }
            QGroupBox {
                border: 1px solid rgba(255, 255, 255, 0.45);
                border-radius: 12px;
                margin-top: 14px;
                background-color: rgba(0, 60, 100, 0.24);
                font-weight: 600;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
                color: #eef9ff;
            }
            QLineEdit, QSpinBox, QTableWidget {
                background-color: rgba(255, 255, 255, 0.96);
                color: #12384a;
                border: 1px solid #78c4e8;
                border-radius: 10px;
                padding: 6px;
            }
            QPushButton {
                background-color: #ffffff;
                color: #1372a7;
                border: 1px solid #3da8d8;
                border-radius: 11px;
                font-weight: 700;
                min-height: 34px;
                padding: 0 14px;
            }
            QPushButton:hover {
                background-color: #dff3fe;
            }
            QPushButton:pressed {
                background-color: #bce6fb;
            }
            QPushButton:disabled {
                color: #8ca8b8;
                background-color: #ebf4fa;
                border-color: #a9c8d8;
            }
            QTableWidget {
                gridline-color: #b5def2;
                selection-background-color: #77c8ec;
                selection-color: #083148;
            }
            QHeaderView::section {
                background-color: #e8f6ff;
                color: #12506f;
                border: 1px solid #b7ddf0;
                padding: 5px;
                font-weight: 700;
            }
            QCheckBox {
                spacing: 8px;
            }
            QLabel#signatureLabel {
                font-family: 'Segoe Script';
                font-size: 12px;
                color: #f0fbff;
                padding: 2px 0 0 2px;
            }
            QStatusBar {
                background-color: rgba(5, 62, 99, 0.48);
                color: #ecf9ff;
            }
            """
        )

    def _connect_signals(self) -> None:
        self.import_btn.clicked.connect(self._import_dataframe)
        self.process_btn.clicked.connect(self._processar)
        self.export_btn.clicked.connect(self._exportar)

    def _select_input_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Selecionar arquivo de entrada", "", "Excel (*.xlsx)")
        if path:
            self.input_path_edit.setText(path)

    def _select_output_file(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Salvar resultado", self.output_path_edit.text(), "Excel (*.xlsx)")
        if path:
            self.output_path_edit.setText(path)

    def _import_dataframe(self) -> None:
        path = self.input_path_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "Arquivo obrigatorio", "Selecione o arquivo de entrada.")
            return

        try:
            df = pd.read_excel(path)
        except Exception as exc:
            QMessageBox.critical(self, "Erro ao importar", str(exc))
            return

        if "Pedidos" not in df.columns or "desm" not in df.columns:
            QMessageBox.warning(
                self,
                "Colunas obrigatorias",
                "A planilha precisa conter as colunas 'Pedidos' e 'desm'.",
            )
            return

        self.df_entrada = df
        self.df_resultado = None
        self.export_btn.setEnabled(False)
        self._carregar_tabela(df.head(200))
        self.statusBar().showMessage(f"Planilha importada com {len(df)} linhas")

    def _build_config(self) -> PRWebConfig:
        return PRWebConfig(
            empresa_gateway=self.empresa_gateway_spin.value(),
            empresa_pr=self.empresa_pr_spin.value(),
            login=self.login_spin.value(),
            senha=self.senha_edit.text(),
            filial=self.filial_spin.value(),
            atividade=self.atividade_edit.text().strip() or "D",
            motivo=self.motivo_spin.value(),
            carga=self.carga_spin.value(),
        )

    def _processar(self) -> None:
        if self.df_entrada is None:
            QMessageBox.warning(self, "Sem dados", "Importe uma planilha antes de processar.")
            return

        if not self.senha_edit.text().strip():
            QMessageBox.warning(self, "Senha obrigatoria", "Informe a senha.")
            return

        config = self._build_config()
        self.statusBar().showMessage("Iniciando processamento...")
        self.process_btn.setEnabled(False)
        self.import_btn.setEnabled(False)

        self.thread = QThread(self)
        self.worker = ProcessWorker(self.df_entrada, config)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.statusBar().showMessage)
        self.worker.finished.connect(self._on_process_finished)
        self.worker.error.connect(self._on_process_error)

        self.worker.finished.connect(self.thread.quit)
        self.worker.error.connect(self.thread.quit)
        self.thread.finished.connect(self._on_thread_finished)

        self.thread.start()

    def _on_process_finished(self, resultado: pd.DataFrame) -> None:
        self.df_resultado = resultado
        self.export_btn.setEnabled(True)
        self._carregar_tabela(resultado.head(200))
        self.statusBar().showMessage("Processamento concluido")

    def _on_process_error(self, error_text: str) -> None:
        self.statusBar().showMessage("Falha no processamento")
        QMessageBox.critical(self, "Erro no processamento", error_text)

    def _on_thread_finished(self) -> None:
        self.process_btn.setEnabled(True)
        self.import_btn.setEnabled(True)
        if self.worker:
            self.worker.deleteLater()
        if self.thread:
            self.thread.deleteLater()
        self.worker = None
        self.thread = None

    def _carregar_tabela(self, df: pd.DataFrame) -> None:
        self.table.clear()
        self.table.setColumnCount(len(df.columns))
        self.table.setRowCount(len(df))
        self.table.setHorizontalHeaderLabels([str(col) for col in df.columns])

        for row_idx in range(len(df)):
            for col_idx, col_name in enumerate(df.columns):
                value = df.iloc[row_idx][col_name]
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(value)))

        self.table.resizeColumnsToContents()

    def _exportar(self) -> None:
        if self.df_resultado is None:
            QMessageBox.warning(self, "Sem resultado", "Nenhum resultado para exportar.")
            return

        output_path = self.output_path_edit.text().strip()
        if not output_path:
            QMessageBox.warning(self, "Arquivo de saida", "Defina o caminho de exportacao.")
            return

        df_export = self.df_resultado.copy()
        if self.only_msg_check.isChecked() and "msg" in df_export.columns:
            df_export = df_export[df_export["msg"].astype(str).str.strip() != ""]

        try:
            Path(output_path).parent.mkdir(parents=True, exist_ok=True)
            df_export.to_excel(output_path, index=False)
            self.statusBar().showMessage(f"Resultado exportado em: {output_path}")
        except Exception as exc:
            QMessageBox.critical(self, "Erro ao exportar", str(exc))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
