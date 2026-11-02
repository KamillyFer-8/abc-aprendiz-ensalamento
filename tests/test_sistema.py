import itertools
from copy import copy
import json
import os
import shutil
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
import controller as c

ROOT = Path(__file__).resolve().parents[1]
WORK = Path(tempfile.gettempdir())/'unisoma-testes'
WORK.mkdir(parents=True, exist_ok=True)


def fixture(path, alunos, recursos=(2, 1, 10, 10), include=True):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Alunos'
    ws.append(c.ALUNOS)
    for row in alunos:
        ws.append(row)
    ws['A1'].font = Font(bold=True, color='123456')
    ws.column_dimensions['A'].width = 38
    if include:
        ws = wb.create_sheet('Turmas')
        ws.append(['Recursos','Quantidade'])
        for k,v in zip(c.RECURSOS,recursos):
            ws.append([k,v])
    ws = wb.create_sheet('Notas')
    ws['A1'] = '=1+2'
    ws['A2'] = 'preservar'
    wb.save(path)
    wb.close()


def alunos(n=11):
    return [(f'Aluno {i}', 'Feminino' if i<8 else 'Masculino', 'Manhã') for i in range(n)]


class SistemaTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(dir=WORK)
        self.path = Path(self.tmp.name)/'teste.xlsx'
        self.saida = Path(self.tmp.name)/'resultado.xlsx'
        c._estado = c.Estado()

    def tearDown(self):
        self.tmp.cleanup()

    def abrir(self, rows=None, recursos=(2,1,10,10), include=True):
        fixture(self.path, alunos() if rows is None else rows, recursos, include)
        c.abrir(self.path)
        return c._estado.problema

    def test_equilibrio_genero_e_capacidade(self):
        p = self.abrir()
        c.resolver()
        self.assertEqual(c._estado.status_solver,'OPTIMAL')
        result = p.dados['Ensalamento']
        self.assertEqual(sorted(result.groupby('Turma').size()),[5,6])
        self.assertEqual(len(result),11)
        c.verificar_resultado(p,result)
        # 8F/11 total: 4F em cada turma, sem limite artificial de 50%.
        self.assertEqual(sorted(result[result.Sexo.eq('Feminino')].groupby('Turma').size()),[4,4])

    def test_turnos_e_letras_continuas(self):
        rows = alunos(5) + [('Tarde '+str(i),'Masculino','Tarde') for i in range(9)]
        p = self.abrir(rows,(2,3,10,10))
        c.resolver()
        self.assertEqual([x['Turma'] for x in p.turmas],['MANHÃ - A','MANHÃ - B','TARDE - C','TARDE - D','TARDE - E'])
        c.verificar_resultado(p,p.dados['Ensalamento'])
        self.assertEqual(c.letra(26),'AA')

    def test_homonimos_preservados(self):
        p = self.abrir([('Ana','Feminino','Manhã'),('Ana','Feminino','Manhã')])
        self.assertFalse(p.erros)
        self.assertEqual(len(p.dados['Inconsistências']),1)
        c.resolver()
        self.assertEqual(len(p.dados['Ensalamento']),2)

    def test_normalizacao_e_trilha(self):
        p=self.abrir([('  Ana   Silva  ','f','manha'),('Beto',' M ','TARDE')])
        self.assertFalse(p.erros)
        self.assertEqual(p.alunos.iloc[0].tolist(),['Ana Silva','Feminino','Manhã'])
        self.assertEqual(len(p.dados['Inconsistências']),2)
        self.assertEqual(len(p.ocorrencias_atuais),0)
        self.assertEqual(len(p.tratamentos_atuais),2)
        c.resolver()

    def test_erros_acumulados_sem_exclusao(self):
        p=self.abrir([(None,None,None),('Ana','X','Ambos'),('Bia','F','Noite')])
        self.assertTrue(p.erros)
        self.assertEqual(len(p.alunos_importados),2)
        self.assertEqual(len(p.indices_pendentes),2)
        self.assertEqual(len(p.indices_vazios),1)
        with self.assertRaises(ValueError): c.resolver()
        self.assertIsNone(c._estado.resultado)
        with self.assertRaises(RuntimeError): c.salvar(self.saida)

    def test_capacidade_insuficiente(self):
        p=self.abrir(alunos(),(1,0,3,0))
        self.assertTrue(any('8 estudante(s) excedente' in x for x in p.erros))
        with self.assertRaises(ValueError): c.resolver()

    def test_sem_turmas_e_capacidade_invalida(self):
        for config in [(0,0,1,1),(2,1,0,3),(2.5,1,10,10),(2,1,-1,10)]:
            p=self.abrir(recursos=config)
            self.assertTrue(p.erros)

    def test_sem_turmas_bloqueia_sem_busca_externa(self):
        p=self.abrir(include=False)
        self.assertTrue(p.erros)
        with self.assertRaises(ValueError):c.resolver()
        self.assertFalse(hasattr(c,'carregar_turmas'))

    def test_exportacao_preserva_entrada_e_formulas(self):
        p=self.abrir()
        original=c.digest(self.path)
        before=load_workbook(self.path)
        input_values=list(before['Alunos'].values)
        c.resolver()
        c.salvar(self.saida)
        after=load_workbook(self.saida)
        self.assertEqual(list(after['Alunos'].values),input_values)
        self.assertEqual(copy(after['Alunos']['A1'].font),copy(before['Alunos']['A1'].font))
        self.assertEqual(after['Alunos'].column_dimensions['A'].width,38)
        self.assertEqual(after['Notas']['A1'].value,'=1+2')
        self.assertEqual(list(next(after['Ensalamento'].values)),c.SAIDA)
        self.assertEqual(list(next(after['Inconsistências'].values)),c.OCORRENCIAS)
        self.assertEqual(after['Ensalamento'].max_row,12)
        c.verificar_resultado(p,pd.read_excel(self.saida,sheet_name='Ensalamento'))
        before.close(); after.close()
        self.assertEqual(c.digest(self.path),original)

    def test_gravacao_bloqueada_preserva_original(self):
        self.abrir(); c.resolver()
        original=c.digest(self.path)
        with patch('controller.os.replace',side_effect=PermissionError('bloqueado')):
            with self.assertRaises(PermissionError): c.salvar(self.saida)
        self.assertEqual(original,c.digest(self.path))
        self.assertFalse(list(self.path.parent.glob('.unisoma-*')))

    def test_mudanca_externa_bloqueia_gravacao(self):
        self.abrir(); c.resolver()
        wb=load_workbook(self.path); wb['Notas']['A3']='alterado'; wb.save(self.path); wb.close()
        with self.assertRaisesRegex(RuntimeError,'externamente'): c.salvar(self.saida)

    def test_estado_antigo_nao_reutilizado(self):
        self.abrir(); c.resolver()
        with self.assertRaises(FileNotFoundError): c.abrir(self.path.with_name('ausente.xlsx'))
        self.assertIsNone(c._estado.problema)
        self.assertIsNone(c._estado.resultado)

    def test_verificador_rejeita_adulteracao(self):
        p=self.abrir(); c.resolver()
        result=p.dados['Ensalamento'].copy()
        result.loc[0,'Turma']='TARDE - C'
        with self.assertRaises(ValueError): c.verificar_resultado(p,result)
        result=p.dados['Ensalamento'].copy()
        result['Turma']='MANHÃ - A'
        with self.assertRaises(ValueError): c.verificar_resultado(p,result)

    def test_solver_timeout_sem_sucesso(self):
        self.abrir()
        with patch.object(c.cp_model.CpSolver,'solve',return_value=c.cp_model.UNKNOWN):
            with self.assertRaisesRegex(RuntimeError,'Limite'): c.resolver()
        self.assertIsNone(c._estado.resultado)
        self.assertNotIn('Ensalamento',c._estado.problema.dados)

    def test_solver_inviavel_modelo_invalido_erro(self):
        for status in (c.cp_model.INFEASIBLE,c.cp_model.MODEL_INVALID):
            self.abrir()
            with patch.object(c.cp_model.CpSolver,'solve',return_value=status):
                with self.assertRaises(RuntimeError): c.resolver()
            self.assertIsNone(c._estado.resultado)
        with patch.object(c.cp_model.CpSolver,'solve',side_effect=RuntimeError('erro')):
            with self.assertRaises(RuntimeError): c.resolver()
        self.assertEqual(c._estado.status_solver,'ERRO_EXECUCAO')

    def test_solucao_viavel_nao_e_otima(self):
        self.abrir()
        actual=c.cp_model.CpSolver.solve
        def feasible(solver, model):
            actual(solver,model)
            return c.cp_model.FEASIBLE
        with patch.object(c.cp_model.CpSolver,'solve',feasible):
            message=c.resolver()
        self.assertIn('sem prova de otimalidade',message)
        self.assertEqual(c._estado.status_solver,'FEASIBLE')

    def test_otimalidade_com_enumeracao_independente(self):
        for n in range(1,9):
            for f in range(n+1):
                rows=[('Aluno '+str(i),'Feminino' if i<f else 'Masculino','Manhã') for i in range(n)]
                self.abrir(rows,(3,0,10,0))
                c.resolver()
                best=None
                for assignments in itertools.product(range(3),repeat=n):
                    occup=[assignments.count(j) for j in range(3)]
                    if max(occup)-min(occup)>1: continue
                    obj=sum(abs(n*sum(assignments[i]==j for i in range(f))-f*occup[j]) for j in range(3))
                    best=obj if best is None else min(best,obj)
                self.assertEqual(c._estado.metricas['objetivo'],best,(n,f))

    def test_lista_vazia(self):
        self.abrir([], (1,1,10,10))
        with self.assertRaisesRegex(ValueError,'Nenhum estudante'):
            c.resolver()
        self.assertIsNone(c._estado.resultado)


    def test_exportacao_rejeita_entrada_e_caminho_equivalente(self):
        self.abrir();c.resolver();h=c.digest(self.path)
        with self.assertRaisesRegex(ValueError,'entrada deve ser preservada'):c.salvar(self.path)
        (self.path.parent/'sub').mkdir()
        alias=self.path.parent/'sub/../teste.xlsx'
        with self.assertRaisesRegex(ValueError,'entrada deve ser preservada'):c.salvar(alias)
        self.assertEqual(c.digest(self.path),h)
    def test_exportacao_repetida_para_destinos_diferentes_preserva_entrada(self):
        self.abrir();c.resolver();h=c.digest(self.path)
        for nome in ('primeiro.xlsx','segundo.xlsx'):
            destino=self.path.with_name(nome);c.salvar(destino)
            c.verificar_resultado(c._estado.problema,pd.read_excel(destino,sheet_name='Ensalamento'))
            self.assertEqual(c.digest(self.path),h)
        self.assertEqual(c._estado.problema.caminho,self.path.resolve())
    def test_falha_de_exportacao_preserva_destino_anterior(self):
        self.abrir();c.resolver();self.saida.write_bytes(b'destino anterior');h=c.digest(self.path)
        with patch('controller.os.replace',side_effect=PermissionError('bloqueado')):
            with self.assertRaises(PermissionError):c.salvar(self.saida)
        self.assertEqual(self.saida.read_bytes(),b'destino anterior');self.assertEqual(c.digest(self.path),h)
    def test_exportacao_rejeita_extensao_invalida(self):
        self.abrir();c.resolver()
        with self.assertRaisesRegex(ValueError,'extensão .xlsx'):c.salvar(self.path.with_suffix('.csv'))

if __name__=='__main__': unittest.main(verbosity=2)
