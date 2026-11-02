"""Teste da interface real PySide6 com diálogos controlados e loop Qt ativo."""
import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
import time
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox
from PySide6.QtCore import QTimer
from PySide6.QtGui import QFontDatabase
from openpyxl import load_workbook
import controller
import gui
from test_sistema import fixture, alunos, ROOT, WORK

class InterfaceTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(dir=WORK)
        self.dialogos=[patch.object(QMessageBox,n).start() for n in ('warning','information','critical')]
        self.addCleanup(patch.stopall)

    def tearDown(self):
        self.tmp.cleanup()

    @classmethod
    def setUpClass(cls):
        cls.app=QApplication.instance() or QApplication([])
        QFontDatabase.addApplicationFont('C:/Windows/Fonts/segoeui.ttf')
        cls.app.setStyle('Fusion')

    def test_fluxo_thread_visualizacao_salvamento(self):
        path=Path(self.tmp.name)/'interface_sintetica.xlsx'
        fixture(path,alunos())
        w=gui.MainWindow()
        w.show()
        self.assertFalse(w._btn_salvar.isEnabled())
        with patch.object(QFileDialog,'getOpenFileName',return_value=(str(path),'')):
            w._abrir()
        self.assertEqual(w._tabela_alunos.rowCount(),11)
        self.assertEqual(w._tabela_turmas.rowCount(),3)
        self.assertTrue(w._btn_resolver.isEnabled())
        ticks=[]
        timer=QTimer(); timer.timeout.connect(lambda:ticks.append(time.monotonic())); timer.start(1)
        w._resolver()
        self.assertFalse(w._btn_abrir.isEnabled())
        self.assertFalse(w._btn_salvar.isEnabled())
        deadline=time.monotonic()+15
        while w._worker is not None and time.monotonic()<deadline:
            self.app.processEvents()
            time.sleep(.001)
        timer.stop()
        self.assertIsNone(w._worker)
        self.assertGreater(len(ticks),0)
        self.assertTrue(w._btn_salvar.isEnabled())
        self.assertEqual(w._tabela_resultado.rowCount(),11)
        self.assertEqual(w._abas.currentIndex(),2)
        with patch.object(QFileDialog,'getSaveFileName',return_value=(str(path.with_name('exportado.xlsx')),'')),patch.object(QMessageBox,'information'):
            w._salvar()
        wb=load_workbook(path.with_name('exportado.xlsx'))
        self.assertEqual(wb['Ensalamento'].max_row,12)
        wb.close()
        w.grab().save(str(Path(self.tmp.name)/'interface.png'))
        w.close()

    def test_interface_sem_turmas_bloqueia_solver(self):
        path=Path(self.tmp.name)/'sem_turmas_sintetica.xlsx'
        fixture(path,alunos(),include=False)
        w=gui.MainWindow()
        with patch.object(QFileDialog,'getOpenFileName',return_value=(str(path),'')):
            w._abrir()
        self.assertFalse(w._btn_resolver.isEnabled())
        self.assertFalse(w._btn_salvar.isEnabled())
        self.assertGreater(w._tabela_pendencias.rowCount(),0)
        w.close()

    def test_historico_duplicado_nao_e_exibido_como_erro_atual(self):
        path=Path(self.tmp.name)/'historico.xlsx'
        fixture(path,[('Ana','Feminino','Manhã')]*2)
        wb=load_workbook(path)
        ws=wb.create_sheet('Inconsistências')
        ws.append(controller.OCORRENCIAS)
        ws.append(['Ana',None,None,'Nome duplicado'])
        wb.save(path);wb.close()
        w=gui.MainWindow()
        with patch.object(QFileDialog,'getOpenFileName',return_value=(str(path),'')):
            w._abrir()
        self.assertEqual(w._tabela_pendencias.rowCount(),1)
        self.assertEqual(w._tabela_pendencias.horizontalHeaderItem(0).text(),'Origem')
        self.assertEqual(w._total_pendencias.text(),'1')
        self.assertIn('Impedimentos: 0',w._aviso_pendencias.text())
        self.assertTrue(w._btn_resolver.isEnabled())
        w.close()

    def test_interface_bloqueia_todos_quando_ha_dados_incompletos(self):
        path=Path(self.tmp.name)/'pendentes.xlsx'
        fixture(path,[('Ana','Feminino','Manhã'),('Bia',None,'Manhã')])
        w=gui.MainWindow()
        with patch.object(QFileDialog,'getOpenFileName',return_value=(str(path),'')):
            w._abrir()
        self.assertFalse(w._btn_resolver.isEnabled())
        self.assertEqual(w._btn_resolver.text(),'Realizar ensalamento')
        self.assertIn('Registros com dados inválidos: 1',w._aviso_pendencias.text())
        self.assertEqual(w._tabela_alunos.rowCount(),2)
        w.close()

    def test_fluxo_completo_oito_planilhas_fornecidas(self):
        import shutil
        downloads=ROOT.parents[1]/'Downloads'
        sources=[downloads/f'instancia_{i}{suffix}.xlsx' for suffix in ('','_alunos_tratados') for i in range(1,5)]
        if not all(path.is_file() for path in sources):
            self.skipTest('Planilhas fornecidas indisponíveis neste computador.')
        for source in sources:
            with self.subTest(arquivo=source.name):
                original=controller.digest(source)
                target=Path(self.tmp.name)/source.name
                shutil.copy2(source,target)
                w=gui.MainWindow()
                errors=[]
                with patch.object(QFileDialog,'getOpenFileName',return_value=(str(target),'')),patch.object(QMessageBox,'critical',side_effect=lambda *args:errors.append(str(args[-1]))):
                    w._abrir()
                    self.assertFalse(errors)
                    invalida=source.name=='instancia_4.xlsx' or '_alunos_tratados' in source.name
                    if invalida:
                        self.assertFalse(w._btn_resolver.isEnabled())
                        self.assertFalse(w._btn_salvar.isEnabled())
                        self.assertIsNone(controller._estado.resultado)
                        w.close()
                        self.assertEqual(controller.digest(source),original)
                        continue
                    self.assertTrue(w._btn_resolver.isEnabled())
                    w._resolver()
                    deadline=time.monotonic()+15
                    while w._worker is not None and time.monotonic()<deadline:
                        self.app.processEvents();time.sleep(.001)
                    self.assertIsNone(w._worker)
                    self.assertFalse(errors)
                    self.assertTrue(w._btn_salvar.isEnabled())
                    i=int(source.name.split('_')[1].split('.')[0])
                    self.assertEqual(w._tabela_resultado.rowCount(),{1:120,2:160,3:165,4:147}[i])
                    with patch.object(QFileDialog,'getSaveFileName',return_value=(str(target.with_name(target.stem+'_resultado.xlsx')),'')),patch.object(QMessageBox,'information'):
                        w._salvar()
                    self.assertFalse(errors)
                wb=load_workbook(target.with_name(target.stem+'_resultado.xlsx'),read_only=True)
                self.assertIn('Ensalamento',wb.sheetnames)
                self.assertIn('Inconsistências',wb.sheetnames)
                wb.close();w.close()
                self.assertEqual(controller.digest(source),original)

    def test_reimportacao_invalida_valida_e_erro_limpa_interface(self):
        path=Path(self.tmp.name)/'transicoes.xlsx';fixture(path,[('Ana',None,'Manhã')])
        w=gui.MainWindow();w._abrir(str(path))
        self.assertFalse(w._btn_resolver.isEnabled());self.assertFalse(w._btn_salvar.isEnabled())
        fixture(path,alunos());w._abrir(str(path))
        self.assertTrue(w._btn_resolver.isEnabled());self.assertFalse(w._btn_salvar.isEnabled())
        w._abrir(str(path.with_name('ausente.xlsx')))
        self.assertFalse(w._btn_resolver.isEnabled());self.assertFalse(w._btn_salvar.isEnabled())
        self.assertEqual(w._tabela_alunos.rowCount(),0);self.assertEqual(w._total_alunos.text(),'—')
        self.assertFalse(hasattr(w,'_btn_turmas'));w.close()

    def test_cancelar_exportacao_preserva_entrada_e_resultado(self):
        path=Path(self.tmp.name)/'entrada.xlsx';fixture(path,alunos());h=controller.digest(path)
        w=gui.MainWindow();w._abrir(str(path));controller.resolver();w._resultado_disponivel=True;w._atualizar_dados()
        with patch.object(QFileDialog,'getSaveFileName',return_value=('','')) as dialogo,patch.object(controller,'salvar') as exportar:
            w._salvar();exportar.assert_not_called()
            self.assertTrue(str(dialogo.call_args.args[2]).endswith('entrada_ensalamento.xlsx'))
        self.assertEqual(controller.digest(path),h);self.assertTrue(w._btn_salvar.isEnabled());w.close()
    def test_escolher_entrada_na_exportacao_nao_sobrescreve(self):
        path=Path(self.tmp.name)/'entrada.xlsx';fixture(path,alunos());h=controller.digest(path)
        w=gui.MainWindow();w._abrir(str(path));controller.resolver();w._resultado_disponivel=True;w._atualizar_dados()
        with patch.object(QFileDialog,'getSaveFileName',return_value=(str(path),'')),patch.object(QMessageBox,'critical') as aviso:
            w._salvar();aviso.assert_called_once()
        self.assertEqual(controller.digest(path),h);self.assertTrue(w._btn_salvar.isEnabled());w.close()

    def test_menu_salvar_original_confirmacao_e_resumo(self):
        path=Path(self.tmp.name)/'entrada.xlsx';fixture(path,alunos());h=controller.digest(path)
        w=gui.MainWindow();w._abrir(str(path));controller.resolver();w._resultado_disponivel=True;w._atualizar_dados()
        self.assertEqual([a.text() for a in w._btn_salvar.menu().actions()],['Exportar novo Excel','Salvar no Excel de entrada'])
        self.assertTrue(w._btn_resumo.isEnabled());self.assertEqual(w._tabela_resumo.rowCount(),3)
        with patch.object(QMessageBox,'question',return_value=QMessageBox.StandardButton.No):w._salvar_na_entrada()
        self.assertEqual(controller.digest(path),h)
        with patch.object(QMessageBox,'question',return_value=QMessageBox.StandardButton.Yes):w._salvar_na_entrada()
        wb=load_workbook(path);self.assertIn('Ensalamento',wb.sheetnames);wb.close()
        dest=path.with_name('resumo.txt')
        with patch.object(QFileDialog,'getSaveFileName',return_value=(str(dest),'')):w._exportar_resumo()
        self.assertIn('Total alocado: 11',dest.read_text(encoding='utf-8'))
        w.show();w._abas.setCurrentIndex(4);self.app.processEvents();w.grab().save(str(Path(self.tmp.name)/'resumo-distribuicao.png'));w.close()
    def test_exportar_inconsistencias_habilitado_com_erro(self):
        path=Path(self.tmp.name)/'pendente.xlsx';fixture(path,[('Ana',None,'Manhã')]);h=controller.digest(path)
        w=gui.MainWindow();w._abrir(str(path));self.assertFalse(w._btn_resolver.isEnabled());self.assertTrue(w._btn_inconsistencias.isEnabled());self.assertFalse(w._btn_resumo.isEnabled())
        dest=path.with_name('inconsistencias.xlsx')
        with patch.object(QFileDialog,'getSaveFileName',return_value=(str(dest),'')):w._exportar_inconsistencias()
        wb=load_workbook(dest);self.assertEqual(wb.sheetnames,['Inconsistências']);wb.close();self.assertEqual(controller.digest(path),h);w.close()
    def test_tratamentos_em_visualizacao_separada(self):
        path=Path(self.tmp.name)/'tratado.xlsx';fixture(path,[(' Ana ','f','manha')])
        w=gui.MainWindow();w._abrir(str(path));self.assertEqual(w._total_pendencias.text(),'0');self.assertEqual(w._tabela_pendencias.rowCount(),0)
        w._tipo_ocorrencia.setCurrentIndex(1);self.assertEqual(w._tabela_pendencias.rowCount(),1);self.assertIn('Tratamentos aplicados: 1',w._aviso_pendencias.text())
        w._abrir(str(path));self.assertEqual(w._tipo_ocorrencia.currentIndex(),0);w.close()

if __name__=='__main__': unittest.main(verbosity=2)
