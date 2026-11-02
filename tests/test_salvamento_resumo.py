import unittest,tempfile
from pathlib import Path
from unittest.mock import patch
from openpyxl import load_workbook
import pandas as pd
import controller as c
from test_sistema import fixture,alunos,WORK

class SalvamentoResumoTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(dir=WORK);self.path=Path(self.tmp.name)/'entrada.xlsx';c._estado=c.Estado()
    def tearDown(self):self.tmp.cleanup()
    def abrir(self,rows=None,recursos=(2,1,10,10)):
        fixture(self.path,alunos() if rows is None else rows,recursos);c.abrir(self.path);return c._estado.problema
    def test_salvar_no_original_so_com_opcao_explicita(self):
        p=self.abrir();w=load_workbook(self.path);valores=list(w['Alunos'].values);w.close();h=c.digest(self.path);c.resolver()
        with self.assertRaises(ValueError):c.salvar(self.path)
        c.salvar(self.path,sobrescrever_entrada=True)
        self.assertNotEqual(c.digest(self.path),h);self.assertEqual(c.digest(self.path),p.hash_original)
        w=load_workbook(self.path);self.assertEqual(list(w['Alunos'].values),valores);self.assertEqual(w['Notas']['A1'].value,'=1+2');self.assertIn('Ensalamento',w.sheetnames);w.close()
        c.salvar(self.path.with_name('novo.xlsx'))
    def test_falha_no_original_preserva_bytes_e_hash(self):
        p=self.abrir();c.resolver();h=c.digest(self.path)
        with patch('controller.os.replace',side_effect=PermissionError('em uso')):
            with self.assertRaises(PermissionError):c.salvar(self.path,sobrescrever_entrada=True)
        self.assertEqual(c.digest(self.path),h);self.assertEqual(p.hash_original,h);self.assertFalse(list(self.path.parent.glob('.unisoma-*')))
    def test_inconsistencias_exportaveis_com_erro_sem_ensalamento(self):
        p=self.abrir([('Ana',None,'Manhã'),('Bia','F','Manhã')]);h=c.digest(self.path);dest=self.path.with_name('pendencias.xlsx')
        c.exportar_inconsistencias(dest);w=load_workbook(dest)
        self.assertEqual(w.sheetnames,['Inconsistências']);self.assertIn('Sexo',w['Inconsistências']['D2'].value);w.close()
        self.assertEqual(c.digest(self.path),h);self.assertIsNone(c._estado.resultado)
        with self.assertRaises(RuntimeError):c.salvar(self.path,sobrescrever_entrada=True)
    def test_exportar_pendencias_nao_pode_sobrescrever_entrada(self):
        self.abrir([('Ana',None,'Manhã')]);h=c.digest(self.path)
        with self.assertRaises(ValueError):c.exportar_inconsistencias(self.path)
        self.assertEqual(c.digest(self.path),h)
    def test_exportar_pendencias_rejeita_entrada_alterada(self):
        self.abrir([('Ana',None,'Manhã')]);w=load_workbook(self.path);w['Alunos']['B2']='F';w.save(self.path);w.close()
        dest=self.path.with_name('pendencias.xlsx')
        with self.assertRaises(RuntimeError):c.exportar_inconsistencias(dest)
        self.assertFalse(dest.exists())
    def test_tratamentos_nao_sao_erros_nem_avisos(self):
        p=self.abrir([('  Ana   Silva ',' f ','manha')]);r=c.resumo_ocorrencias()
        self.assertFalse(p.erros);self.assertEqual(r['tratamentos'],1);self.assertEqual(r['ocorrencias'],0);self.assertTrue(c.painel_ocorrencias().empty)
        self.assertEqual(len(c.painel_tratamentos()),1);self.assertIn('TRATADO:',p.dados['Inconsistências'].iloc[0].Motivo)
        self.assertIn('Nome:',p.dados['Inconsistências'].iloc[0].Motivo);self.assertIn('Disponibilidade:',p.dados['Inconsistências'].iloc[0].Motivo)
    def test_formato_M_F_valido_nao_gera_problema(self):
        p=self.abrir([('Ana','F','Manhã')]);self.assertEqual(len(p.tratamentos_atuais),0);self.assertFalse(p.erros)
    def test_tratamentos_estaveis_ao_salvar_e_reabrir_original(self):
        self.abrir([(' Ana ',' f ','manha')]);c.resolver();c.salvar(self.path,sobrescrever_entrada=True)
        for _ in range(3):
            c.abrir(self.path);p=c._estado.problema;c.validar(p);self.assertEqual(len(p.dados['Inconsistências']),1)
            c.resolver();c.salvar(self.path,sobrescrever_entrada=True)
    def test_resumo_por_turma_conferencia_independente(self):
        p=self.abrir();c.resolver();df=c.resumo_distribuicao();self.assertEqual(df.Alocados.sum(),len(p.alunos))
        for row in df.to_dict('records'):
            grupo=p.dados['Ensalamento'].loc[p.dados['Ensalamento'].Turma.eq(row['Turma'])];n=len(grupo);f=int(grupo.Sexo.eq('Feminino').sum())
            self.assertEqual(row['Alocados'],n);self.assertEqual(row['Feminino'],f);self.assertEqual(row['Masculino'],n-f)
            self.assertEqual(row['Feminino (%)'],round(100*f/n,2) if n else 0)
    def test_resumo_turmas_vazias_e_exportacao_txt(self):
        self.abrir([('Ana','F','Manhã')],(3,1,10,10));c.resolver();df=c.resumo_distribuicao();self.assertEqual(len(df),4);self.assertEqual(df.Alocados.sum(),1)
        self.assertTrue((df.loc[df.Alocados.eq(0),'Feminino (%)']==0).all());h=c.digest(self.path);dest=self.path.with_name('resumo.txt');c.exportar_resumo(dest)
        texto=dest.read_text(encoding='utf-8');self.assertIn('Total alocado: 1',texto);self.assertIn('100.00%',texto);self.assertEqual(c.digest(self.path),h)
    def test_nova_entrada_invalida_limpa_resumo_e_bloqueia_exportacao(self):
        self.abrir();c.resolver();self.assertFalse(c.resumo_distribuicao().empty);self.abrir([('Ana',None,'Manhã')])
        self.assertTrue(c.resumo_distribuicao().empty)
        with self.assertRaises(RuntimeError):c.exportar_resumo(self.path.with_name('resumo.txt'))

if __name__=='__main__':unittest.main(verbosity=2)
