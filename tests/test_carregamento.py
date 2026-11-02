import unittest,tempfile
from pathlib import Path
from unittest.mock import patch
import pandas as pd
from openpyxl import load_workbook
import controller as c
from test_sistema import fixture,alunos,WORK,ROOT

class CarregamentoTests(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(dir=WORK);self.path=Path(self.tmp.name)/'teste.xlsx';c._estado=c.Estado()
    def tearDown(self):self.tmp.cleanup()
    def abrir(self,rows=None,recursos=(2,1,10,10),include=True):
        fixture(self.path,alunos() if rows is None else rows,recursos,include);c.abrir(self.path);return c._estado.problema
    def test_nome_conhecido_nao_busca_configuracao(self):
        self.path=self.path.with_name('instancia_3_alunos_tratados.xlsx')
        p=self.abrir(include=False)
        self.assertTrue(any('Turmas ausente' in e for e in p.erros))
        with self.assertRaises(ValueError):c.resolver()
    def test_120_para_100_vagas_excedente_20(self):
        p=self.abrir(alunos(120),(4,0,25,0))
        self.assertTrue(any('Excedente: 20' in e for e in p.erros))
        with patch.object(c.cp_model.CpSolver,'solve') as solver:
            with self.assertRaises(ValueError):c.resolver()
            solver.assert_not_called()
    def test_cada_campo_obrigatorio_bloqueia_todo_solver(self):
        for coluna in range(3):
            row=['Ana','Feminino','Manhã'];row[coluna]=None
            p=self.abrir([('Bia','Feminino','Manhã'),tuple(row)])
            self.assertEqual(p.indices_pendentes,[1]);self.assertTrue(p.erros)
            with patch.object(c.cp_model.CpSolver,'solve') as solver:
                with self.assertRaises(ValueError):c.resolver()
                solver.assert_not_called()
            self.assertNotIn('Ensalamento',p.dados)
            with self.assertRaises(RuntimeError):c.salvar(self.path)
    def test_varios_campos_mesma_linha_uma_ocorrencia(self):
        p=self.abrir([('Ana',None,None),('Bia','F','Manhã')])
        self.assertEqual(len(p.ocorrencias_atuais),1)
        self.assertEqual(c.resumo_ocorrencias()['registros_afetados'],1)
        self.assertIn('Sexo, Disponibilidade',p.ocorrencias_atuais[0]['Campos'])
    def test_vazias_preservam_numero_real_excel(self):
        p=self.abrir([('Ana','f','manha'),(None,None,None),('Bia',None,'Tarde'),('  ',' ',None),('Caio','M','Tarde')])
        self.assertEqual(p.indices_vazios,[1,3]);self.assertEqual(p.indices_pendentes,[2])
        self.assertEqual(p.ocorrencias_atuais[0]['Linha'],4)
        self.assertEqual(len(p.alunos_importados),3)
        wb=load_workbook(self.path);self.assertEqual(wb['Alunos']['A4'].value,'Bia');self.assertIsNone(wb['Alunos']['B4'].value);wb.close()
        before=p.dados['Inconsistências'].copy()
        for _ in range(4):c.validar(p);pd.testing.assert_frame_equal(before,p.dados['Inconsistências'])
    def test_cabecalho_normalizado_sem_aviso_ou_mudanca_excel(self):
        self.abrir([('Ana','f','manha')]);wb=load_workbook(self.path);wb['Alunos']['A1']='Recebi as instâncias ';wb.save(self.path);wb.close()
        initial=c.digest(self.path);c.abrir(self.path)
        self.assertFalse(c._estado.problema.erros);self.assertEqual(len(c.painel_ocorrencias()),0)
        c.resolver();c.salvar(self.path.with_name("resultado.xlsx"))
        wb=load_workbook(self.path);self.assertEqual(wb['Alunos']['A1'].value,'Recebi as instâncias ');wb.close()
    def test_corrigir_reimportar_e_limpar_solucao_anterior(self):
        self.abrir([('Ana',None,'Manhã')]);self.assertTrue(c._estado.problema.erros)
        wb=load_workbook(self.path);wb['Alunos']['B2']='F';wb.save(self.path);wb.close()
        c.abrir(self.path);self.assertFalse(c._estado.problema.erros);c.resolver();self.assertEqual(c._estado.resultado,'resolvido')
        self.abrir([('Bia',None,'Tarde')]);self.assertIsNone(c._estado.resultado);self.assertNotIn('Ensalamento',c._estado.problema.dados)
        self.abrir();self.assertFalse(c._estado.problema.erros);self.assertIsNone(c._estado.resultado)
    def test_quantidade_obrigatoria_ausente(self):
        self.abrir();wb=load_workbook(self.path);wb['Turmas']['B4']=None;wb.save(self.path);wb.close();c.abrir(self.path)
        self.assertTrue(any('vagas_manha' in e for e in c._estado.problema.erros))
    def test_historico_nao_e_problema_atual(self):
        self.abrir([('Ana','F','Manhã')]);wb=load_workbook(self.path);ws=wb.create_sheet('Inconsistências');ws.append(c.OCORRENCIAS);ws.append(['Ana',None,None,'Dados incompletos']);wb.save(self.path);wb.close();c.abrir(self.path)
        self.assertTrue(c.painel_ocorrencias().empty);self.assertEqual(c.resumo_ocorrencias()['historico'],1);c.resolver()
    def test_duas_turnos_sem_vagas_informa_ambos(self):
        p=self.abrir([('Ana','F','Manhã'),('Bia','F','Tarde')],(1,1,0,0))
        msg=c.orientacao_validacao();self.assertIn('Turno: Manhã',msg);self.assertIn('Turno: Tarde',msg)
    def test_aba_alunos_ausente(self):
        self.abrir();wb=load_workbook(self.path);del wb['Alunos'];wb.save(self.path);wb.close();c.abrir(self.path)
        self.assertTrue(any('Aba Alunos ausente' in e for e in c._estado.problema.erros))

    @unittest.skipUnless((ROOT/'dados/instancia_4.xlsx').is_file(), 'Instância privada não distribuída.')
    def test_instancia_4_detalhes_e_revalidacao_estaveis(self):
        source=ROOT/'dados/instancia_4.xlsx';initial=c.digest(source);c.abrir(source);p=c._estado.problema
        self.assertEqual(len(p.indices_vazios),3);self.assertEqual(len(p.indices_pendentes),8)
        self.assertEqual(len(p.ocorrencias_atuais),14)
        wb=load_workbook(source,read_only=True)
        for r in p.ocorrencias_atuais:
            if r['Motivo'].startswith('ERRO:'):
                values=[wb['Alunos'].cell(r['Linha'],col).value for col in range(1,4)]
                self.assertEqual(['' if v is None else v for v in values],[r[k] for k in c.ALUNOS])
        wb.close();expected=p.dados['Inconsistências'].copy()
        for _ in range(5):c.validar(p);pd.testing.assert_frame_equal(expected,p.dados['Inconsistências'])
        self.assertEqual(c.digest(source),initial)
    def test_primeira_coluna_ambigua_nao_e_inventada(self):
        self.abrir();wb=load_workbook(self.path);wb['Alunos']['A1']='Identificador';wb.save(self.path);wb.close();c.abrir(self.path)
        self.assertTrue(c._estado.problema.erros)

if __name__=='__main__':unittest.main(verbosity=2)
