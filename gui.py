
"""Interface gráfica do Sistema de Ensalamento ABC Aprendiz.

A interface utiliza PySide6.
Toda a lógica de processamento pertence ao controller.py.

Funcionalidades:
- Selecionar planilha Excel
- Visualizar e pesquisar alunos
- Visualizar estrutura de turmas e capacidades
- Consultar inconsistências
- Executar ensalamento
- Visualizar o resultado antes de salvar
- Exportar resultado em um novo arquivo Excel
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QMenu,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import controller


# ==========================================================
# CORES DA IDENTIDADE VISUAL
# ==========================================================

AZUL = "#31529C"
AZUL_ESCURO = "#233D75"
VERDE = "#2EA5A0"
FUNDO = "#F5F7FA"
TEXTO = "#263347"
CINZA = "#697586"
VERMELHO = "#B42318"

FILTRO_EXCEL = "Planilhas Excel (*.xlsx)"


ESTILO = f"""
QMainWindow, QWidget {{
    background-color: {FUNDO};
    color: {TEXTO};
    font-family: "Segoe UI", Arial;
    font-size: 13px;
}}

QFrame#cabecalho {{
    background: white;
    border-bottom: 1px solid #E2E8F0;
}}

QFrame#cartao {{
    background: white;
    border: 1px solid #E2E8F0;
    border-radius: 12px;
}}

QLabel#titulo {{
    font-size: 22px;
    font-weight: bold;
    color: {AZUL_ESCURO};
}}

QLabel#subtitulo {{
    font-size: 12px;
    color: {CINZA};
}}

QLabel#numero {{
    font-size: 25px;
    font-weight: bold;
    color: {AZUL_ESCURO};
}}

QLabel#rotulo {{
    font-size: 12px;
    color: {CINZA};
}}

QPushButton {{
    background: white;
    border: 1px solid #CED8E4;
    border-radius: 8px;
    padding: 10px 16px;
    font-weight: bold;
}}

QPushButton:hover {{
    background: #EDF3FA;
}}

QPushButton:disabled {{
    background: #E9EDF2;
    color: #9AA5B3;
}}

QPushButton#principal {{
    color: white;
    background: {AZUL};
    border: 1px solid {AZUL};
}}

QPushButton#principal:hover {{
    background: {AZUL_ESCURO};
}}

QPushButton#principal:disabled {{
    color: #9AA5B3;
    background: #E9EDF2;
    border: 1px solid #E2E7ED;
}}

QPushButton#salvar {{
    color: white;
    background: {VERDE};
    border: 1px solid {VERDE};
}}

QPushButton#salvar:disabled {{
    color: #9AA5B3;
    background: #E9EDF2;
    border: 1px solid #E2E7ED;
}}

QLineEdit {{
    background: white;
    border: 1px solid #D8E1EB;
    border-radius: 8px;
    padding: 10px;
}}

QTabWidget::pane {{
    border: 1px solid #E2E8F0;
    background: white;
}}

QTabBar::tab {{
    background: #E9EEF4;
    padding: 12px 20px;
    margin-right: 3px;
}}

QTabBar::tab:selected {{
    background: white;
    color: {AZUL};
    font-weight: bold;
    border-bottom: 3px solid {VERDE};
}}

QTableWidget {{
    background: white;
    alternate-background-color: #F7F9FC;
    gridline-color: #E8EDF3;
    border: none;
}}

QHeaderView::section {{
    background: #EDF2F7;
    padding: 9px;
    border: none;
    font-weight: bold;
}}
"""


# ==========================================================
# COMPONENTES AUXILIARES
# ==========================================================

def criar_tabela():
    tabela = QTableWidget()

    tabela.setEditTriggers(
        QTableWidget.EditTrigger.NoEditTriggers
    )
    tabela.setSelectionBehavior(
        QTableWidget.SelectionBehavior.SelectRows
    )
    tabela.setAlternatingRowColors(True)
    tabela.verticalHeader().setVisible(False)

    tabela.setProperty("detalhes_rolaveis", True)
    tabela.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
    tabela.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
    tabela.setWordWrap(True)
    tabela.setTextElideMode(Qt.TextElideMode.ElideNone)
    tabela.horizontalHeader().sectionResized.connect(tabela.resizeRowsToContents)

    return tabela


def preencher_tabela(tabela, dataframe):
    """Mostra um DataFrame em uma tabela somente leitura."""

    tabela.setSortingEnabled(False)
    tabela.clear()

    if dataframe is None:
        tabela.setRowCount(0)
        tabela.setColumnCount(0)
        return

    df = dataframe.astype(object).fillna("")

    tabela.setRowCount(len(df))
    tabela.setColumnCount(len(df.columns))

    tabela.setHorizontalHeaderLabels(
        [str(coluna) for coluna in df.columns]
    )

    for i, linha in enumerate(
        df.itertuples(index=False, name=None)
    ):
        for j, valor in enumerate(linha):
            tabela.setItem(
                i,
                j,
                QTableWidgetItem(str(valor))
            )

    if tabela.property("detalhes_rolaveis"):
        larguras = {"Origem": 145, "Aba": 90, "Linha": 70, "Campos": 180,
                    "Nome": 450, "Sexo": 160, "Disponibilidade": 220, "Motivo": 850,
                    "Turma": 200, "Turno": 180, "Vagas": 150}
        for j, coluna in enumerate(df.columns):
            tabela.setColumnWidth(j, larguras.get(str(coluna), 150))
            for i in range(len(df)):
                tabela.item(i, j).setToolTip(tabela.item(i, j).text())
        tabela.resizeRowsToContents()
    tabela.setSortingEnabled(True)


def criar_cartao(titulo):
    cartao = QFrame()
    cartao.setObjectName("cartao")

    layout = QVBoxLayout(cartao)
    layout.setContentsMargins(18, 14, 18, 14)

    rotulo = QLabel(titulo)
    rotulo.setObjectName("rotulo")

    numero = QLabel("—")
    numero.setObjectName("numero")

    layout.addWidget(rotulo)
    layout.addWidget(numero)

    return cartao, numero


# ==========================================================
# JANELA PRINCIPAL
# ==========================================================

class SolverThread(QThread):
    concluido = Signal(str)
    falhou = Signal(str)

    def run(self):
        try:
            self.concluido.emit(controller.resolver())
        except Exception as exc:
            self.falhou.emit(str(exc))


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle(
            "ABC Aprendiz | Sistema de Ensalamento"
        )

        self.resize(1050, 700)
        self.setMinimumSize(760, 520)
        self.setStyleSheet(ESTILO)

        self._worker = None
        self._caminho = None
        self._dados = {}
        self._df_alunos = None
        self._resultado_disponivel = False

        self._construir_interface()

    def _construir_interface(self):

        central = QWidget()
        self.setCentralWidget(central)

        principal = QVBoxLayout(central)
        principal.setContentsMargins(0, 0, 0, 0)
        principal.setSpacing(0)

        # --------------------------------------------------
        # CABEÇALHO
        # --------------------------------------------------

        cabecalho = QFrame()
        cabecalho.setObjectName("cabecalho")

        cab_layout = QHBoxLayout(cabecalho)
        cab_layout.setContentsMargins(24, 16, 24, 16)

        logo = QLabel()
        logo.setFixedSize(90, 65)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)

        caminho_logo = (
            Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
            / "assets"
            / "logo_abc.png"
        )

        if caminho_logo.exists():
            imagem = QPixmap(str(caminho_logo))

            if not imagem.isNull():
                logo.setPixmap(
                    imagem.scaled(
                        85,
                        60,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation
                    )
                )

        if logo.pixmap().isNull():
            logo.setText("ABC\nAprendiz")
            logo.setStyleSheet(f"color: {AZUL}; font-weight: bold; font-size: 16px;")
        titulos = QVBoxLayout()

        titulo = QLabel("Sistema de Ensalamento")
        titulo.setObjectName("titulo")

        subtitulo = QLabel(
            "ABC Aprendiz | Organização automática de turmas"
        )
        subtitulo.setObjectName("subtitulo")

        titulos.addWidget(titulo)
        titulos.addWidget(subtitulo)

        cab_layout.addWidget(logo)
        cab_layout.addLayout(titulos)
        cab_layout.addStretch()

        principal.addWidget(cabecalho)

        # --------------------------------------------------
        # CONTEÚDO
        # --------------------------------------------------

        conteudo = QWidget()
        layout = QVBoxLayout(conteudo)

        layout.setContentsMargins(25, 22, 25, 20)
        layout.setSpacing(15)

        # BOTÕES

        botoes = QHBoxLayout()

        self._btn_abrir = QPushButton(
            "Selecionar planilha"
        )

        self._btn_resolver = QPushButton(
            "Realizar ensalamento"
        )

        self._btn_salvar = QPushButton(
            "Salvar / exportar"
        )

        self._btn_resolver.setObjectName("principal")
        self._btn_salvar.setObjectName("salvar")

        self._btn_resolver.setEnabled(False)
        self._btn_salvar.setEnabled(False)

        self._btn_abrir.clicked.connect(self._abrir)
        self._btn_resolver.clicked.connect(self._resolver)
        menu_salvar=QMenu(self)
        menu_salvar.addAction("Exportar novo Excel",self._salvar)
        menu_salvar.addAction("Salvar no Excel de entrada",self._salvar_na_entrada)
        self._btn_salvar.setMenu(menu_salvar)

        botoes.addWidget(self._btn_abrir)
        botoes.addWidget(self._btn_resolver)
        botoes.addWidget(self._btn_salvar)
        botoes.addStretch()

        layout.addLayout(botoes)

        # ARQUIVO

        self._arquivo = QLabel(
            "Nenhuma planilha selecionada."
        )
        self._arquivo.setObjectName("subtitulo")
        layout.addWidget(self._arquivo)

        # INDICADORES

        indicadores = QHBoxLayout()

        cartao, self._total_alunos = criar_cartao(
            "Estudantes"
        )
        indicadores.addWidget(cartao)

        cartao, self._total_turmas = criar_cartao(
            "Turmas"
        )
        indicadores.addWidget(cartao)

        cartao, self._total_vagas = criar_cartao(
            "Vagas totais"
        )
        indicadores.addWidget(cartao)

        cartao, self._total_pendencias = criar_cartao(
            "Ocorrências atuais"
        )
        indicadores.addWidget(cartao)

        layout.addLayout(indicadores)

        # --------------------------------------------------
        # ABAS
        # --------------------------------------------------

        self._abas = QTabWidget()
        self._abas.setUsesScrollButtons(False)

        # ABA 1 - ESTUDANTES

        pagina_alunos = QWidget()
        alunos_layout = QVBoxLayout(pagina_alunos)

        topo_alunos = QHBoxLayout()

        titulo_alunos = QLabel("Estudantes da planilha")
        topo_alunos.addWidget(titulo_alunos)
        topo_alunos.addStretch()

        self._busca = QLineEdit()
        self._busca.setPlaceholderText(
            "Buscar estudante pelo nome..."
        )
        self._busca.setFixedWidth(250)

        self._busca.textChanged.connect(
            self._filtrar_alunos
        )

        topo_alunos.addWidget(self._busca)
        alunos_layout.addLayout(topo_alunos)

        self._tabela_alunos = criar_tabela()
        alunos_layout.addWidget(self._tabela_alunos)

        self._abas.addTab(
            pagina_alunos,
            "Estudantes"
        )

        # ABA 2 - INCONSISTÊNCIAS

        pagina_pendencias = QWidget()
        pendencias_layout = QVBoxLayout(
            pagina_pendencias
        )

        self._aviso_pendencias = QLabel(
            "As inconsistências serão exibidas aqui "
            "quando estiverem disponíveis."
        )
        self._aviso_pendencias.setWordWrap(True)

        acoes_pendencias=QHBoxLayout()
        self._tipo_ocorrencia=QComboBox()
        self._tipo_ocorrencia.addItems(["Erros e avisos atuais","Tratamentos aplicados"])
        self._tipo_ocorrencia.currentIndexChanged.connect(self._atualizar_ocorrencias)
        self._btn_inconsistencias=QPushButton("Exportar inconsistências")
        self._btn_inconsistencias.setEnabled(False)
        self._btn_inconsistencias.clicked.connect(self._exportar_inconsistencias)
        acoes_pendencias.addWidget(self._tipo_ocorrencia)
        acoes_pendencias.addStretch()
        acoes_pendencias.addWidget(self._btn_inconsistencias)
        pendencias_layout.addLayout(acoes_pendencias)
        self._tabela_pendencias = criar_tabela()

        pendencias_layout.addWidget(
            self._aviso_pendencias
        )
        pendencias_layout.addWidget(
            self._tabela_pendencias
        )

        self._abas.addTab(
            pagina_pendencias,
            "Inconsistências"
        )

        # ABA 3 - ENSALAMENTO

        pagina_resultado = QWidget()
        resultado_layout = QVBoxLayout(
            pagina_resultado
        )

        self._aviso_resultado = QLabel(
            "O ensalamento aparecerá aqui "
            "após a otimização."
        )
        self._aviso_resultado.setWordWrap(True)

        self._tabela_resultado = criar_tabela()

        resultado_layout.addWidget(
            self._aviso_resultado
        )
        resultado_layout.addWidget(
            self._tabela_resultado
        )

        self._abas.addTab(
            pagina_resultado,
            "Ensalamento"
        )

        pagina_turmas = QWidget()
        turmas_layout = QVBoxLayout(pagina_turmas)
        self._tabela_turmas = criar_tabela()
        turmas_layout.addWidget(self._tabela_turmas)
        self._abas.addTab(pagina_turmas, "Turmas")
        pagina_resumo=QWidget()
        resumo_layout=QVBoxLayout(pagina_resumo)
        self._aviso_resumo=QLabel("O resumo aparecerá após um ensalamento válido.")
        self._aviso_resumo.setWordWrap(True)
        self._tabela_resumo=criar_tabela()
        self._btn_resumo=QPushButton("Exportar resumo (.txt)")
        self._btn_resumo.setEnabled(False)
        self._btn_resumo.clicked.connect(self._exportar_resumo)
        resumo_layout.addWidget(self._aviso_resumo)
        resumo_layout.addWidget(self._tabela_resumo)
        resumo_layout.addWidget(self._btn_resumo)
        self._abas.addTab(pagina_resumo,"Resumo")
        layout.addWidget(self._abas, 1)

        # STATUS

        self._status = QPlainTextEdit("Selecione um arquivo Excel para começar.")
        self._status.setReadOnly(True)
        self._status.setUndoRedoEnabled(False)
        self._status.setTabChangesFocus(True)
        self._status.setFrameShape(QFrame.Shape.NoFrame)
        self._status.document().setDocumentMargin(0)
        self._status.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        self._status.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._status.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._status.setFixedHeight(self._status.fontMetrics().lineSpacing() + 14)

        layout.addWidget(self._status)

        principal.addWidget(conteudo, 1)

    # ======================================================
    # AÇÕES
    # ======================================================

    def _abrir(self, caminho=None):
        if not isinstance(caminho, (str, Path)):
            caminho, _ = QFileDialog.getOpenFileName(
                self, "Selecionar planilha",
                str((Path(sys.executable).parent if getattr(sys, "frozen", False)
                     else Path(__file__).resolve().parent) / "dados"),
                FILTRO_EXCEL
            )

        if not caminho:
            return

        try:
            mensagem = controller.abrir(caminho)

            self._caminho = caminho
            self._tipo_ocorrencia.setCurrentIndex(0)
            self._resultado_disponivel = False

            self._arquivo.setText(
                f"Arquivo: {Path(caminho).name}"
            )

            self._btn_resolver.setEnabled(not controller._estado.problema.erros)
            self._btn_salvar.setEnabled(False)

            self._atualizar_dados()

            self._abas.setCurrentIndex(0)
            self._mostrar_status(mensagem)
            if controller._estado.problema.erros:
                erros=controller._estado.problema.erros
                titulo=("ATENÇÃO — VAGAS INSUFICIENTES" if any('Excedente:' in e for e in erros)
                        else "ATENÇÃO — CONFIGURAÇÃO INCOMPLETA" if any('Turmas' in e for e in erros)
                        else "ATENÇÃO — DADOS INCOMPLETOS")
                self._atencao(titulo,controller.orientacao_validacao())
            elif controller.resumo_ocorrencias()['avisos']:
                self._atencao("ATENÇÃO — POSSÍVEIS HOMÔNIMOS",
                              "Há nomes iguais na aba Alunos. Todos foram preservados. "
                              "Confira as linhas indicadas em Inconsistências. O ensalamento está disponível.")

        except Exception as exc:
            self._caminho = None
            self._df_alunos = None
            self._resultado_disponivel = False
            for indicador in (self._total_alunos,self._total_turmas,self._total_vagas,self._total_pendencias):
                indicador.setText("—")
            self._aviso_resultado.setText("Selecione uma planilha válida e execute o ensalamento.")
            self._btn_resolver.setEnabled(False)
            self._btn_salvar.setEnabled(False)
            self._btn_inconsistencias.setEnabled(False)
            self._btn_resumo.setEnabled(False)
            for tabela in (self._tabela_alunos, self._tabela_turmas, self._tabela_pendencias, self._tabela_resultado, self._tabela_resumo):
                preencher_tabela(tabela, None)
            self._arquivo.setText("Nenhuma planilha selecionada.")
            self._aviso_resumo.setText("O resumo aparecerá após um ensalamento válido.")
            self._erro(str(exc))

    def _atualizar_dados(self):
        """Atualiza a interface com os dados disponíveis."""

        estado = controller._estado

        if estado.problema is None:
            return

        self._dados = estado.problema.dados

        # ESTUDANTES

        self._df_alunos = controller.planilha_inicial()

        preencher_tabela(
            self._tabela_alunos,
            self._df_alunos
        )

        if self._df_alunos is not None:
            self._total_alunos.setText(
                str(len(self._df_alunos))
            )
        else:
            self._total_alunos.setText("—")

        # TURMAS

        self._total_turmas.setText("—")
        self._total_vagas.setText("—")

        estrutura = controller.estrutura_turmas()
        preencher_tabela(self._tabela_turmas, estrutura)
        self._total_turmas.setText(str(len(estrutura)))
        self._total_vagas.setText(str(int(estrutura['Vagas'].sum()) if len(estrutura) else 0))
        self._btn_resolver.setEnabled(not estado.problema.erros and self._worker is None)
        pendentes = len(estado.problema.indices_pendentes)
        self._btn_resolver.setText("Realizar ensalamento")

        # INCONSISTÊNCIAS

        df_pendencias = controller.painel_ocorrencias()
        self._atualizar_ocorrencias()
        self._total_pendencias.setText(str(len(df_pendencias)))
        resumo = controller.resumo_ocorrencias()
        self._aviso_pendencias.setText(
            f"Registros com dados inválidos: {resumo['registros_afetados']}. "
            f"Ocorrências atuais: {resumo['ocorrencias']}. Avisos: {resumo['avisos']}. "
            f"Impedimentos: {resumo['impedimentos']}. Tratamentos aplicados: {resumo['tratamentos']} (sem bloqueio). Linhas vazias ignoradas: {resumo['vazias']}. "
            f"Histórico preservado na exportação: {resumo['historico']} registros (fora da contagem atual). "
            "Possíveis homônimos são avisos. Dados incompletos bloqueiam todo o ensalamento."
        )

        # ENSALAMENTO

        df_resultado = self._dados.get(
            "Ensalamento"
        )

        resultado_valido = (
            self._resultado_disponivel
            and df_resultado is not None
            and estado.resultado == "resolvido"
        )

        if resultado_valido:
            preencher_tabela(
                self._tabela_resultado,
                df_resultado
            )

            self._aviso_resultado.setText(
                f"Resultado disponível: "
                f"{len(df_resultado)} estudantes alocados e verificados."
            )

        else:
            preencher_tabela(
                self._tabela_resultado,
                None
            )

            self._aviso_resultado.setText(
                "O resultado aparecerá após "
                "a execução do ensalamento."
            )

        self._btn_salvar.setEnabled(resultado_valido)
        self._btn_inconsistencias.setEnabled(self._worker is None)
        self._btn_resumo.setEnabled(resultado_valido and self._worker is None)
        preencher_tabela(self._tabela_resumo,controller.resumo_distribuicao())
        self._aviso_resumo.setText(
            controller.texto_resumo().split('\n')[2] if resultado_valido
            else "O resumo aparecerá após um ensalamento válido."
        )

    def _atualizar_ocorrencias(self, *_):
        quadro=(controller.painel_tratamentos() if self._tipo_ocorrencia.currentIndex()==1
                else controller.painel_ocorrencias())
        preencher_tabela(self._tabela_pendencias,quadro)

    def _filtrar_alunos(self, texto):
        if self._df_alunos is None:
            return

        texto = texto.strip()

        if not texto:
            preencher_tabela(
                self._tabela_alunos,
                self._df_alunos
            )
            return

        coluna = next(
            (
                col
                for col in self._df_alunos.columns
                if str(col).strip().lower()
                in ("nome", "nome do aluno")
            ),
            None
        )

        if coluna is None:
            return

        filtrados = self._df_alunos[
            self._df_alunos[coluna]
            .fillna("")
            .astype(str)
            .str.contains(
                texto,
                case=False,
                regex=False
            )
        ]

        preencher_tabela(
            self._tabela_alunos,
            filtrados
        )

    def _resolver(self):
        if self._caminho is None or self._worker is not None:
            return
        self._resultado_disponivel = False
        if controller._estado.problema.erros or controller._estado.problema.indices_pendentes:
            self._atencao("ATENÇÃO — PLANILHA INADEQUADA", controller.orientacao_validacao())
            return
        for button in (self._btn_resolver, self._btn_salvar, self._btn_abrir, self._btn_inconsistencias, self._btn_resumo):
            button.setEnabled(False)
        self._mostrar_status("Otimizando turmas. Aguarde...")
        self._worker = SolverThread(self)
        self._worker.concluido.connect(self._solver_concluido)
        self._worker.falhou.connect(self._solver_falhou)
        self._worker.finished.connect(self._solver_finalizado)
        self._worker.start()

    def _solver_concluido(self, mensagem):
        self._resultado_disponivel = controller._estado.resultado == "resolvido"
        self._atualizar_dados()
        self._abas.setCurrentIndex(2)
        self._mostrar_status(mensagem)
        if not getattr(self,"_verificacao_automatica",False):
            QMessageBox.information(self,"Ensalamento concluído",mensagem)

    def _solver_falhou(self, mensagem):
        self._resultado_disponivel = False
        self._atualizar_dados()
        self._erro(mensagem)

    def _solver_finalizado(self):
        self._worker.deleteLater()
        self._worker = None
        self._btn_abrir.setEnabled(True)
        self._atualizar_dados()

    def closeEvent(self, event):
        if self._worker is not None:
            event.ignore()
            self._mostrar_status("Aguarde a otimização terminar antes de fechar.")
        else:
            event.accept()

    def _salvar(self):
        if self._caminho is None:
            self._erro(
                "Nenhuma planilha selecionada."
            )
            return

        if not self._btn_salvar.isEnabled():
            self._erro(
                "Execute o ensalamento antes de exportar."
            )
            return

        origem = Path(self._caminho)
        destino, _ = QFileDialog.getSaveFileName(
            self, "Exportar ensalamento",
            str(origem.with_name(origem.stem + "_ensalamento.xlsx")), FILTRO_EXCEL
        )
        if not destino:
            return
        if not Path(destino).suffix:
            destino += ".xlsx"
        try:
            mensagem = controller.salvar(destino)

            self._mostrar_status(mensagem)

            QMessageBox.information(
                self,
                "Resultado exportado",
                mensagem
            )

        except Exception as exc:
            self._erro(str(exc))

    def _salvar_na_entrada(self):
        if not self._caminho or not self._btn_salvar.isEnabled():
            self._erro("Execute um ensalamento válido antes de salvar na entrada.")
            return
        resposta=QMessageBox.question(
            self,"Salvar no Excel de entrada",
            "As abas Ensalamento e Inconsistências do arquivo selecionado serão atualizadas. "
            "As abas de entrada serão preservadas.\n\nDeseja salvar nesse mesmo arquivo?",
            QMessageBox.StandardButton.Yes|QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if resposta!=QMessageBox.StandardButton.Yes:return
        try:
            mensagem=controller.salvar(self._caminho,sobrescrever_entrada=True)
            self._mostrar_status(mensagem)
            QMessageBox.information(self,"Resultado salvo",mensagem)
        except Exception as exc:self._erro(str(exc))

    def _exportar_inconsistencias(self):
        if not self._caminho or self._worker is not None:return
        origem=Path(self._caminho)
        destino,_=QFileDialog.getSaveFileName(self,"Exportar inconsistências",
                    str(origem.with_name(origem.stem+"_inconsistencias.xlsx")),FILTRO_EXCEL)
        if not destino:return
        if not Path(destino).suffix:destino+=".xlsx"
        try:
            mensagem=controller.exportar_inconsistencias(destino)
            self._mostrar_status(mensagem)
            QMessageBox.information(self,"Inconsistências exportadas",mensagem)
        except Exception as exc:self._erro(str(exc))

    def _exportar_resumo(self):
        if not self._caminho or not self._btn_resumo.isEnabled():return
        origem=Path(self._caminho)
        destino,_=QFileDialog.getSaveFileName(self,"Exportar resumo",
                    str(origem.with_name(origem.stem+"_resumo.txt")),"Relatório de texto (*.txt)")
        if not destino:return
        if not Path(destino).suffix:destino+=".txt"
        try:
            mensagem=controller.exportar_resumo(destino)
            self._mostrar_status(mensagem)
            QMessageBox.information(self,"Resumo exportado",mensagem)
        except Exception as exc:self._erro(str(exc))

    # ======================================================
    # MENSAGENS
    # ======================================================

    def _mostrar_status(self, mensagem, erro=False):
        cor = VERMELHO if erro else AZUL_ESCURO

        self._status.setStyleSheet(
            f"color: {cor}; padding: 6px;"
        )

        self._status.setPlainText(mensagem)
        linhas_visiveis = 3 if erro else 1
        self._status.setFixedHeight(self._status.fontMetrics().lineSpacing() * linhas_visiveis + 14)
        self._status.verticalScrollBar().setValue(0)

    def _atencao(self, titulo, mensagem):
        self._mostrar_status(mensagem, erro=True)
        if not getattr(self, "_verificacao_automatica", False):
            QMessageBox.warning(self, titulo, mensagem)

    def _erro(self, mensagem):
        self._mostrar_status(
            f"Erro: {mensagem}",
            erro=True
        )

        if getattr(self, "_verificacao_automatica", False):
            self._erro_verificacao = mensagem
            return
        QMessageBox.critical(self, "Erro", mensagem)


# ==========================================================
# INICIALIZAÇÃO
# ==========================================================

def run() -> int:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()

    if '--verificar-planilha' in sys.argv:
        _verificar_fluxo(app, window)
    else:
        _trazer_para_frente(window)

    return app.exec()


def _verificar_fluxo(app, window):
    """Diagnóstico interno: executa o mesmo fluxo Qt em uma cópia de teste explícita."""
    import argparse
    import json
    import time
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--verificar-planilha', required=True)
    parser.add_argument('--relatorio', required=True)
    parser.add_argument('--exportar-para')
    parser.add_argument('--salvar-na-entrada',action='store_true')
    args, _ = parser.parse_known_args(sys.argv[1:])
    caminho = str(Path(args.verificar_planilha).resolve())
    relatorio = Path(args.relatorio)
    window._verificacao_automatica = True
    window._erro_verificacao = None
    started = time.monotonic()

    def registrar(fase, **extra):
        payload = {'fase': fase, 'empacotado': bool(getattr(sys, 'frozen', False)), **extra}
        relatorio.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')

    def conferir():
        if window._worker is not None:
            QTimer.singleShot(100, conferir)
            return
        try:
            if window._erro_verificacao:
                raise RuntimeError(window._erro_verificacao)
            if not window._btn_salvar.isEnabled():
                raise RuntimeError('A interface não disponibilizou resultado salvável.')
            p = controller._estado.problema
            controller.verificar_resultado(p, p.dados['Ensalamento'])
            origem_hash=controller.digest(caminho)
            destino=Path(args.exportar_para) if args.exportar_para else Path(caminho).with_name(Path(caminho).stem+'_ensalamento.xlsx')
            controller.salvar(destino)
            assert controller.digest(caminho)==origem_hash
            resumo=controller.resumo_distribuicao()
            assert window._btn_resumo.isEnabled() and window._tabela_resumo.rowCount()==len(p.turmas)
            assert int(resumo['Alocados'].sum())==len(p.alunos)
            controller.exportar_resumo(destino.with_suffix('.txt'))
            if args.salvar_na_entrada:
                controller.salvar(caminho,sobrescrever_entrada=True)
            registrar('concluido', resumo_exportado=True, salvar_na_entrada=args.salvar_na_entrada, ok=True, entrada_preservada=True, arquivo_exportado=str(destino), status=controller._estado.status_solver,
                      lidos=len(p.alunos_importados), alocados=window._tabela_resultado.rowCount(),
                      pendentes=len(p.indices_pendentes), previa_visivel=window._abas.currentIndex()==2,
                      botao_salvar_habilitado=True, segundos=time.monotonic()-started)
            app.exit(0)
        except Exception as exc:
            registrar('erro', ok=False, erro=str(exc))
            app.exit(1)

    def iniciar():
        try:
            registrar('carregando')
            window._abrir(caminho)
            if window._erro_verificacao:
                raise RuntimeError(window._erro_verificacao)
            if not window._btn_resolver.isEnabled():
                p=controller._estado.problema
                assert not window._btn_salvar.isEnabled()
                assert controller._estado.resultado is None
                assert window._btn_inconsistencias.isEnabled()
                origem_hash=controller.digest(caminho)
                destino=Path(args.exportar_para) if args.exportar_para else Path(caminho).with_name(Path(caminho).stem+'_inconsistencias.xlsx')
                controller.exportar_inconsistencias(destino)
                assert controller.digest(caminho)==origem_hash
                registrar('bloqueado',inconsistencias_exportadas=True,ok=True,impedimentos=p.erros,pendentes=len(p.indices_pendentes),
                          vazias=len(p.indices_vazios),alocados=0,solver_executado=False)
                app.exit(0)
                return
            registrar('otimizando')
            window._resolver()
            QTimer.singleShot(100, conferir)
        except Exception as exc:
            registrar('erro', ok=False, erro=str(exc))
            app.exit(1)
    QTimer.singleShot(0, iniciar)


def _trazer_para_frente(window: QMainWindow) -> None:
    """Preserva o comportamento original do executável."""

    window.raise_()
    window.activateWindow()


if __name__ == "__main__":
    raise SystemExit(run())


