import unittest
from copy import deepcopy
import pandas as pd
from openpyxl import load_workbook
import controller as c
from pathlib import Path
import tempfile
from test_sistema import fixture, alunos, WORK


class OcorrenciasTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(dir=WORK)
        self.path = Path(self.tmp.name)/'teste.xlsx'
        self.saida = Path(self.tmp.name)/'resultado.xlsx'
        c._estado = c.Estado()

    def tearDown(self):
        self.tmp.cleanup()

    def abrir(self, rows):
        fixture(self.path, rows)
        c.abrir(self.path)
        return c._estado.problema

    # Usa apenas fixtures e diretórios temporários do teste base.
    def registrar_historico(self, rows):
        wb=load_workbook(self.path)
        ws=wb.create_sheet('Inconsistências')
        ws.append(c.OCORRENCIAS)
        for row in rows: ws.append(row)
        wb.save(self.path);wb.close()
        c.abrir(self.path)
        return c._estado.problema

    def test_historico_nao_repete_aviso_do_mesmo_nome(self):
        self.abrir([('Ana','Feminino','Manhã'),('ANA','Feminino','Manhã')])
        p=self.registrar_historico([('Ana',None,None,'Nome duplicado')])
        self.assertEqual(len(p.dados['Inconsistências']),2)
        self.assertEqual(len(p.ocorrencias_atuais),1)
        self.assertFalse(p.erros)
        c.resolver()
        self.assertEqual(len(p.dados['Ensalamento']),2)
        view=c.painel_ocorrencias()
        self.assertEqual(view.Origem.tolist(),['Aviso atual'])

    def test_grupo_novo_gera_um_aviso_preservando_todos(self):
        p=self.abrir([('Bia','Feminino','Manhã')]*3)
        self.assertEqual(len(p.ocorrencias_atuais),1)
        self.assertIn('2, 3, 4',p.ocorrencias_atuais[0]['Motivo'])
        c.resolver()
        self.assertEqual(len(p.dados['Ensalamento']),3)

    def test_outro_nome_no_historico_nao_suprime_aviso_novo(self):
        self.abrir([('Bia','Feminino','Manhã')]*2)
        p=self.registrar_historico([('Ana',None,None,'Nome duplicado')])
        self.assertEqual(len(p.dados['Inconsistências']),2)
        self.assertEqual(len(p.ocorrencias_atuais),1)
        self.assertEqual(set(c.painel_ocorrencias().Origem),{'Aviso atual'})

    def test_salvar_reabrir_nao_acumula_ocorrencias(self):
        p=self.abrir([('  Ana  ','f','manha'),('Ana','Feminino','Manhã')])
        c.resolver();c.salvar(self.saida)
        expected=len(p.dados['Inconsistências'])
        for _ in range(3):
            c.abrir(self.saida);c.resolver()
            self.saida=self.saida.with_name(self.saida.stem+'-nova.xlsx')
            c.salvar(self.saida)
            self.assertEqual(len(c._estado.problema.dados['Inconsistências']),expected)
            self.assertEqual(len(c._estado.problema.dados['Ensalamento']),2)

    def test_historico_nao_torna_erro_atual_um_aviso(self):
        self.abrir([('Ana',None,'Manhã')])
        p=self.registrar_historico([('Ana',None,'Manhã','Dados incompletos')])
        self.assertTrue(p.erros)
        self.assertIn('Impedimento atual',c.painel_ocorrencias().Origem.tolist())
        with self.assertRaises(ValueError): c.resolver()

    def test_revalidacao_com_turmas_nao_acumula(self):
        p=self.abrir([('Ana','Feminino','Manhã')]*2)
        before=p.dados['Inconsistências'].copy()
        c.validar(p);c.validar(p)
        pd.testing.assert_frame_equal(before,p.dados['Inconsistências'])

    def test_historico_integral_preservado_painel_agrupado(self):
        self.abrir([('Ana','Feminino','Manhã')]*2)
        history=[('Ana','Feminino','Manhã','Nome duplicado'),('Ana','Feminino','Manhã','AVISO: possível nome duplicado')]
        p=self.registrar_historico(history)
        self.assertEqual(len(p.historico),2)
        self.assertEqual(len(p.dados['Inconsistências']),3)
        self.assertEqual(len(c.painel_ocorrencias()),1)
        c.resolver();c.salvar(self.saida)
        saved=pd.read_excel(self.saida,sheet_name='Inconsistências').fillna('')
        self.assertEqual(saved.values.tolist()[:2],[list(row) for row in history])

if __name__=='__main__': unittest.main(verbosity=2)
