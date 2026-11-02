"""Lógica local de validação, otimização CP-SAT e exportação, sem UI."""
from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from collections import Counter
import hashlib
import math
import os
import tempfile
import time
import unicodedata
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill
from ortools.sat.python import cp_model

ALUNOS = ['Nome', 'Sexo', 'Disponibilidade']
SAIDA = ALUNOS + ['Turma']
OCORRENCIAS = ALUNOS + ['Motivo']
RECURSOS = ('turmas_manha', 'turmas_tarde', 'vagas_manha', 'vagas_tarde')


def chave(value):
    if pd.isna(value):
        return ''
    text = ' '.join(str(value).strip().casefold().split())
    return ''.join(c for c in unicodedata.normalize('NFKD', text) if not unicodedata.combining(c))


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


@dataclass
class Problema:
    caminho: Path
    dados: dict[str, pd.DataFrame]
    hash_original: str
    alunos: pd.DataFrame = field(default_factory=lambda: pd.DataFrame(columns=ALUNOS))
    turmas: list[dict] = field(default_factory=list)
    erros: list[str] = field(default_factory=list)
    historico: pd.DataFrame = field(default_factory=lambda: pd.DataFrame(columns=OCORRENCIAS))
    indices_validos: list[int] = field(default_factory=list)
    ocorrencias_atuais: list[dict] = field(default_factory=list)
    alunos_importados: pd.DataFrame = field(default_factory=lambda: pd.DataFrame(columns=ALUNOS))
    indices_pendentes: list[int] = field(default_factory=list)
    linhas_excel: dict[int, int] = field(default_factory=dict)
    indices_vazios: list[int] = field(default_factory=list)
    tratamentos_atuais: list[dict] = field(default_factory=list)


@dataclass
class Estado:
    problema: Problema | None = None
    resultado: str | None = None
    status_solver: str = 'NAO_EXECUTADO'
    metricas: dict = field(default_factory=dict)


_estado = Estado()


def letra(index):
    result = ''
    index += 1
    while index:
        index, rest = divmod(index - 1, 26)
        result = chr(65 + rest) + result
    return result


def validar(p):
    """Recalcula diagnósticos atuais sem alterar registros ou posições de origem."""
    p.erros, p.turmas, p.indices_validos = [], [], []
    p.ocorrencias_atuais, p.indices_pendentes, p.indices_vazios = [], [], []
    p.linhas_excel = {}
    p.tratamentos_atuais = []
    historico = p.historico.reindex(columns=OCORRENCIAS).astype(object).fillna('')
    ocorrencias = historico.to_dict('records')
    assinaturas = {tuple(str(r[c]) for c in OCORRENCIAS) for r in ocorrencias}

    def registrar(motivo, row=None, impeditivo=False, linha=None, campos=(), aba='Alunos', tratado=False):
        registro = {c: '' for c in OCORRENCIAS}
        if row is not None:
            registro.update({c: '' if pd.isna(row.get(c, '')) else row.get(c, '') for c in ALUNOS})
        registro['Motivo'] = ('TRATADO: ' if tratado else 'ERRO: ' if impeditivo else 'AVISO: ') + motivo
        assinatura = tuple(str(registro[c]) for c in OCORRENCIAS)
        if assinatura not in assinaturas:
            ocorrencias.append(registro)
            assinaturas.add(assinatura)
        # O histórico nunca suprime nem multiplica o diagnóstico atual.
        destino=p.tratamentos_atuais if tratado else p.ocorrencias_atuais
        destino.append({**registro, 'Aba': aba, 'Linha': linha, 'Campos': ', '.join(campos)})
        if impeditivo:
            p.erros.append(motivo)

    raw = p.dados.get('Alunos')
    p.alunos = pd.DataFrame(columns=ALUNOS)
    p.alunos_importados = pd.DataFrame(columns=ALUNOS)
    if raw is None:
        registrar('Aba Alunos ausente. Inclua Alunos e importe novamente.', impeditivo=True)
    else:
        raw = raw.copy()
        mapa = {'nome':'Nome', 'nome do aluno':'Nome', 'sexo':'Sexo', 'disponibilidade':'Disponibilidade'}
        raw.columns = [mapa.get(chave(col), col) for col in raw.columns]
        if len(set(raw.columns)) != len(raw.columns):
            registrar('Alunos contém cabeçalhos ambíguos ou repetidos.', impeditivo=True)
        else:
            # O primeiro campo do template de três colunas representa Nome.
            if (len(raw.columns)==3 and list(raw.columns[1:])==['Sexo','Disponibilidade']
                    and chave(raw.columns[0])=='recebi as instancias'):
                raw=raw.rename(columns={raw.columns[0]:'Nome'})
                registrar('Alunos linha 1: cabeçalho conhecido corrigido para Nome na leitura.',
                          linha=1,campos=('Nome',),tratado=True)
            if not set(ALUNOS).issubset(raw.columns):
                registrar('Colunas obrigatórias em Alunos: '+', '.join(ALUNOS), impeditivo=True)
            else:
                # Mapeamento criado ANTES de descartar linhas vazias; nunca renumera filtros.
                p.linhas_excel = {i:pos for pos,i in enumerate(raw.index,start=2)}
                normalizados=raw[ALUNOS].copy().astype(object)
                sexos={'f':'Feminino','feminino':'Feminino','m':'Masculino','masculino':'Masculino'}
                turnos={'manha':'Manhã','tarde':'Tarde'}
                for i,row in raw[ALUNOS].iterrows():
                    if all(not chave(row[c]) for c in ALUNOS):
                        p.indices_vazios.append(i)
                        continue
                    campos=[]
                    if not isinstance(row['Nome'],str) or not row['Nome'].strip():
                        campos.append('Nome')
                    else:
                        normalizados.at[i,'Nome']=' '.join(row['Nome'].split())
                    for coluna,mapa_valores in [('Sexo',sexos),('Disponibilidade',turnos)]:
                        valor=mapa_valores.get(chave(row[coluna]))
                        if valor is None: campos.append(coluna)
                        else: normalizados.at[i,coluna]=valor
                    alteracoes=[]
                    for coluna in ALUNOS:
                        original=row[coluna];novo=normalizados.at[i,coluna]
                        if pd.isna(original) or pd.isna(novo) or original==novo:continue
                        if coluna=='Sexo' and original in ('F','M','Feminino','Masculino'):continue
                        alteracoes.append(f'{coluna}: {original!r} → {novo!r}')
                    if alteracoes:
                        registrar(f'Alunos linha {p.linhas_excel[i]}: normalização aplicada: '+ '; '.join(alteracoes),
                                  row,linha=p.linhas_excel[i],campos=tuple(a.split(':',1)[0] for a in alteracoes),tratado=True)
                    if campos:
                        p.indices_pendentes.append(i)
                        linha=p.linhas_excel[i]
                        registrar(f'Alunos linha {linha}: campos ausentes ou inválidos: {", ".join(campos)}. '
                                  'Corrija no Excel e importe novamente.',row,True,linha,campos)
                    else: p.indices_validos.append(i)
                p.alunos_importados=normalizados.drop(index=p.indices_vazios)
                p.alunos=normalizados.loc[p.indices_validos].copy()
                keys=p.alunos_importados['Nome'].map(lambda v:chave(v) if isinstance(v,str) else '')
                repetidos=keys[keys.ne('') & keys.duplicated(keep=False)]
                for nome,grupo in repetidos.groupby(repetidos,sort=False):
                    linhas=', '.join(str(p.linhas_excel[i]) for i in grupo.index)
                    registrar(f'Possível nome duplicado nas linhas {linhas} de Alunos; '
                              'possíveis homônimos, todos preservados. Revise os cadastros.',
                              raw.loc[grupo.index[0]],linha=p.linhas_excel[grupo.index[0]],campos=('Nome',))
    if not len(p.alunos_importados):
        registrar('Nenhum estudante cadastrado na aba Alunos.',impeditivo=True)

    config=p.dados.get('Turmas')
    if config is None:
        registrar('Aba Turmas ausente. Inclua Recursos e Quantidade no mesmo Excel e importe novamente.',impeditivo=True,aba='Turmas')
    elif not {'Recursos','Quantidade'}.issubset(config.columns):
        registrar('Turmas exige as colunas Recursos e Quantidade.',impeditivo=True,aba='Turmas')
    else:
        valores={}
        for linha,(_,row) in enumerate(config.iterrows(),start=2):
            if all(not chave(row[c]) for c in ('Recursos','Quantidade')): continue
            key=chave(row['Recursos'])
            if key not in RECURSOS:
                registrar(f'Turmas linha {linha}: recurso desconhecido ({row["Recursos"]}).',impeditivo=True,linha=linha,aba='Turmas');continue
            if key in valores:
                registrar(f'Turmas linha {linha}: recurso duplicado {key}.',impeditivo=True,linha=linha,aba='Turmas');continue
            value=row['Quantidade']
            try:
                if isinstance(value,bool) or pd.isna(value) or not math.isfinite(float(value)) or float(value)<0 or not float(value).is_integer():raise ValueError()
                valores[key]=int(value)
            except (TypeError,ValueError,OverflowError):
                registrar(f'Turmas linha {linha}: {key} deve ser inteiro não negativo ({value}).',impeditivo=True,linha=linha,aba='Turmas')
        for key in RECURSOS:
            if key not in valores:registrar(f'Turmas: recurso obrigatório ausente ou inválido: {key}.',impeditivo=True,aba='Turmas')
        if len(valores)==4:
            for turno,sufixo in [('Manhã','manha'),('Tarde','tarde')]:
                qtd,cap=valores[f'turmas_{sufixo}'],valores[f'vagas_{sufixo}']
                if qtd and not cap:registrar(f'{turno}: capacidade deve ser positiva quando há turmas.',impeditivo=True,aba='Turmas')
                total=int(p.alunos_importados['Disponibilidade'].map(chave).eq(chave(turno)).sum())
                if total and not qtd:registrar(f'{turno}: nenhuma turma para {total} estudantes.',impeditivo=True,aba='Turmas')
                if total>qtd*cap:registrar(f'Turno: {turno}. Estudantes: {total}. Vagas disponíveis: {qtd*cap}. Excedente: {total-qtd*cap} estudante(s) excedente(s).',impeditivo=True,aba='Turmas')
                for _ in range(qtd):p.turmas.append({'Turma':f'{turno.upper()} - {letra(len(p.turmas))}','Turno':turno,'Vagas':cap})
    p.dados['Inconsistências']=pd.DataFrame(ocorrencias,columns=OCORRENCIAS)
    p.dados.pop('Ensalamento',None)
    if _estado.problema is p:
        _estado.resultado=None
        _estado.status_solver='NAO_EXECUTADO'
        _estado.metricas={}


def painel_ocorrencias():
    """Apenas diagnósticos atuais; histórico permanece separado na exportação."""
    columns=['Origem','Aba','Linha','Campos']+OCORRENCIAS
    p=_estado.problema
    if p is None:return pd.DataFrame(columns=columns)
    return pd.DataFrame([{'Origem':'Impedimento atual' if r['Motivo'].startswith('ERRO:') else 'Aviso atual',**r}
                         for r in p.ocorrencias_atuais],columns=columns)


def resumo_ocorrencias():
    p=_estado.problema
    if p is None:return dict(registros_afetados=0,ocorrencias=0,avisos=0,impedimentos=0,historico=0,vazias=0,tratamentos=0)
    return dict(registros_afetados=len(p.indices_pendentes),ocorrencias=len(p.ocorrencias_atuais),
                avisos=sum(r['Motivo'].startswith('AVISO:') for r in p.ocorrencias_atuais),
                impedimentos=len(p.erros),historico=len(p.historico),vazias=len(p.indices_vazios),
                tratamentos=len(p.tratamentos_atuais))


def orientacao_validacao():
    p=_estado.problema
    if p is None:return 'Nenhuma planilha selecionada.'
    linhas=[f'{len(p.indices_pendentes)} registro(s) com dados obrigatórios inválidos.' ] if p.indices_pendentes else []
    linhas+=p.erros[:12]
    if len(p.erros)>12:linhas.append('Consulte os demais detalhes na aba Inconsistências.')
    return '\n'.join(linhas)+ '\n\nCorrija as abas Alunos e Turmas no mesmo Excel e importe o arquivo novamente. Nenhum estudante foi ensalado.'


def abrir(caminho):
    _estado.problema, _estado.resultado = None, None
    _estado.status_solver, _estado.metricas = 'NAO_EXECUTADO', {}
    path = Path(caminho).resolve()
    if path.suffix.lower() != '.xlsx':
        raise ValueError('Selecione um arquivo .xlsx.')
    initial = digest(path)
    dados = pd.read_excel(path, sheet_name=None, engine='openpyxl')
    if digest(path) != initial:
        raise RuntimeError('Arquivo alterado durante a leitura. Abra novamente.')
    p = Problema(path, dados, initial)
    p.historico = dados.get('Inconsistências', pd.DataFrame(columns=OCORRENCIAS)).copy()
    validar(p)
    _estado.problema = p
    return (f'Arquivo carregado: {path.name}. {len(p.alunos_importados)} linhas lidas; '
            f'{len(p.alunos)} alunos aptos; {len(p.indices_pendentes)} linhas pendentes; '
            f'{len(p.erros)} impedimento(s) de execução. {len(p.indices_vazios)} linhas vazias ignoradas.')


def planilha_inicial():
    return None if _estado.problema is None else _estado.problema.alunos_importados


def estrutura_turmas():
    return pd.DataFrame(_estado.problema.turmas if _estado.problema else [], columns=['Turma', 'Turno', 'Vagas'])


def verificar_resultado(p, resultado):
    """Verificação independente por linha de entrada, preservando homônimos."""
    if p.erros or p.indices_pendentes:
        raise ValueError("A entrada contém impedimentos; nenhum resultado pode ser considerado válido.")
    if list(resultado.columns) != SAIDA or len(resultado) != len(p.alunos):
        raise ValueError('Resultado incompleto ou colunas inválidas.')
    if not resultado[ALUNOS].astype(object).reset_index(drop=True).equals(p.alunos.astype(object).reset_index(drop=True)):
        raise ValueError('Identidades/ordem dos estudantes foram alteradas.')
    classes = {c['Turma']: c for c in p.turmas}
    counts = Counter(resultado['Turma'])
    for row in resultado.to_dict('records'):
        turma = classes.get(row['Turma'])
        if turma is None or turma['Turno'] != row['Disponibilidade']:
            raise ValueError('Turma inexistente ou turno incompatível.')
    for c in p.turmas:
        if counts[c['Turma']] > c['Vagas']:
            raise ValueError('Capacidade excedida.')
    for turno in ('Manhã', 'Tarde'):
        cs = [c for c in p.turmas if c['Turno'] == turno]
        if cs and len({c['Vagas'] for c in cs}) == 1:
            occup = [counts[c['Turma']] for c in cs]
            if max(occup)-min(occup) > 1:
                raise ValueError('Desequilíbrio quantitativo superior a um estudante.')


def resolver(limite_segundos=60.0):
    p = _estado.problema
    _estado.resultado, _estado.metricas = None, {}
    _estado.status_solver = 'NAO_EXECUTADO'
    if p is None:
        raise RuntimeError('Abra uma planilha primeiro.')
    p.dados.pop('Ensalamento', None)
    if p.erros or p.indices_pendentes:
        _estado.status_solver = 'DADOS_INVALIDOS'
        raise ValueError('Corrija os impedimentos:\n' + '\n'.join(p.erros))
    if not math.isfinite(limite_segundos) or limite_segundos <= 0:
        raise ValueError('Limite de tempo deve ser positivo e finito.')
    started = time.monotonic()
    model = cp_model.CpModel()
    counts, gender_devs = {}, []
    for turno in ('Manhã', 'Tarde'):
        ids = p.alunos.index[p.alunos['Disponibilidade'].eq(turno)].tolist()
        n = len(ids)
        f = int(sum(p.alunos.loc[ids, 'Sexo'].eq('Feminino')))
        classes = [j for j,c in enumerate(p.turmas) if c['Turno'] == turno]
        if not classes:
            continue
        lo, remainder = divmod(n, len(classes))
        for j in classes:
            cap = p.turmas[j]['Vagas']
            nf = model.new_int_var(0, min(f,cap), f'f_{j}')
            nm = model.new_int_var(0, min(n-f,cap), f'm_{j}')
            occ = model.new_int_var(lo, lo + bool(remainder), f'n_{j}')
            model.add(occ == nf + nm)
            model.add(occ <= cap)
            counts[j] = (nf,nm)
            deviation = model.new_int_var(0, n*n, f'd_{j}')
            model.add_abs_equality(deviation, n*nf - f*occ)
            gender_devs.append(deviation)
        model.add(sum(counts[j][0] for j in classes) == f)
        model.add(sum(counts[j][1] for j in classes) == n-f)
    model.minimize(sum(gender_devs))
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = limite_segundos
    solver.parameters.num_search_workers = 1
    solver.parameters.random_seed = 0
    try:
        status = solver.solve(model)
    except Exception:
        _estado.status_solver = 'ERRO_EXECUCAO'
        raise
    _estado.status_solver = solver.status_name(status)
    _estado.metricas = {'status': _estado.status_solver, 'segundos': time.monotonic()-started}
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        descriptions = {cp_model.INFEASIBLE:'Problema inviável.', cp_model.MODEL_INVALID:'Modelo matemático inválido.', cp_model.UNKNOWN:'Limite de tempo/interrupção sem solução comprovada.'}
        raise RuntimeError(descriptions.get(status,'Erro de execução do solver.'))
    resultado = p.alunos.copy()
    resultado['Turma'] = ''
    for turno in ('Manhã','Tarde'):
        for sex, pos in [('Feminino',0),('Masculino',1)]:
            ids = p.alunos.index[p.alunos['Disponibilidade'].eq(turno) & p.alunos['Sexo'].eq(sex)].tolist()
            offset = 0
            for j,c in enumerate(p.turmas):
                if c['Turno'] != turno:
                    continue
                amount = solver.value(counts[j][pos])
                resultado.loc[ids[offset:offset+amount], 'Turma'] = c['Turma']
                offset += amount
            if offset != len(ids):
                raise RuntimeError('Falha de desagregação da solução.')
    try:
        verificar_resultado(p, resultado)
    except Exception:
        _estado.status_solver = 'ERRO_VERIFICACAO'
        raise
    p.dados['Ensalamento'] = resultado
    _estado.resultado = 'resolvido'
    _estado.metricas.update({'objetivo': solver.objective_value, 'limite_inferior': solver.best_objective_bound, 'alocados':len(resultado), 'linhas_lidas':len(p.alunos_importados), 'linhas_pendentes':len(p.indices_pendentes)})
    return ('Solução ótima' if status == cp_model.OPTIMAL else 'Solução viável, sem prova de otimalidade (limite/interrupção)') + f': {len(resultado)} estudantes alocados e verificados.'


def _destino(caminho, extensao, sobrescrever_entrada=False):
    p=_estado.problema
    if p is None:raise RuntimeError('Selecione uma planilha primeiro.')
    path=Path(caminho).resolve()
    if path.suffix.lower()!=extensao:raise ValueError(f'Escolha um arquivo com extensão {extensao}.')
    mesma=path==p.caminho or (path.exists() and os.path.samefile(path,p.caminho))
    if mesma and not sobrescrever_entrada:
        raise ValueError('Escolha outro nome ou pasta: a planilha de entrada deve ser preservada. Use a opção Salvar no Excel de entrada para atualizá-la.')
    if sobrescrever_entrada and path!=p.caminho:
        raise ValueError('A opção de salvar na entrada exige o caminho original selecionado.')
    if digest(p.caminho)!=p.hash_original:
        raise RuntimeError('O Excel foi alterado externamente. Importe novamente antes de salvar ou exportar.')
    return p,path


def _gravar_excel(p,path,quadros,preservar_entrada=False):
    temp=None
    wb=load_workbook(p.caminho) if preservar_entrada else Workbook()
    try:
        if not preservar_entrada:wb.remove(wb.active)
        for nome,df in quadros.items():
            if nome in wb:del wb[nome]
            ws=wb.create_sheet(nome);ws.append(list(df.columns))
            for record in df.astype(object).fillna('').itertuples(index=False,name=None):
                ws.append(list(record))
                for cell in ws[ws.max_row]:
                    if isinstance(cell.value,str):cell.data_type='s'
            ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
            for cell in ws[1]:
                cell.font=Font(bold=True,color='FFFFFF');cell.fill=PatternFill('solid',fgColor='31529C')
            for column in ws.columns:
                ws.column_dimensions[column[0].column_letter].width=min(80,max(18,max(len(str(c.value or '')) for c in column)+2))
        fd,temp=tempfile.mkstemp(suffix='.xlsx',prefix='.unisoma-',dir=path.parent);os.close(fd)
        wb.save(temp);wb.close()
        check=load_workbook(temp,read_only=True)
        try:
            for nome,df in quadros.items():
                if check[nome].max_row!=len(df)+1 or list(next(check[nome].values))!=list(df.columns):
                    raise RuntimeError('Falha na conferência do Excel temporário.')
        finally:check.close()
        if digest(p.caminho)!=p.hash_original:raise RuntimeError('Arquivo alterado durante a gravação. Operação cancelada.')
        os.replace(temp,path)
        if path==p.caminho:p.hash_original=digest(path)
    except PermissionError as exc:
        raise PermissionError('Não foi possível gravar. Feche o Excel e confira a permissão da pasta.') from exc
    finally:
        wb.close()
        if temp and os.path.exists(temp):os.unlink(temp)


def salvar(caminho, sobrescrever_entrada=False):
    p=_estado.problema
    if p is None or p.erros or p.indices_pendentes or _estado.resultado!='resolvido':
        raise RuntimeError('Não há resultado válido para salvar.')
    p,path=_destino(caminho,'.xlsx',sobrescrever_entrada)
    verificar_resultado(p,p.dados['Ensalamento'])
    _gravar_excel(p,path,{'Ensalamento':p.dados['Ensalamento'][SAIDA],
                        'Inconsistências':p.dados['Inconsistências'][OCORRENCIAS]},True)
    if sobrescrever_entrada:return f'Ensalamento e Inconsistências salvos no Excel de entrada: {path}.'
    return f'Resultado exportado para: {path}. A planilha de entrada foi preservada.'


def exportar_inconsistencias(caminho):
    p,path=_destino(caminho,'.xlsx')
    _gravar_excel(p,path,{'Inconsistências':p.dados['Inconsistências'].reindex(columns=OCORRENCIAS)})
    return f'Inconsistências exportadas para: {path}. Nenhum ensalamento foi criado por esta operação.'


def painel_tratamentos():
    colunas=['Aba','Linha','Campos']+OCORRENCIAS
    p=_estado.problema
    return pd.DataFrame(p.tratamentos_atuais if p else [],columns=colunas)


def resumo_distribuicao():
    colunas=['Turno','Turma','Vagas','Alocados','Feminino','Masculino','Feminino (%)','Masculino (%)','Ocupação (%)']
    p=_estado.problema
    if p is None or _estado.resultado!='resolvido':return pd.DataFrame(columns=colunas)
    verificar_resultado(p,p.dados['Ensalamento']);rows=[]
    for turma in p.turmas:
        alunos=p.dados['Ensalamento'].loc[p.dados['Ensalamento']['Turma'].eq(turma['Turma'])]
        n=len(alunos);f=int(alunos['Sexo'].eq('Feminino').sum());m=n-f;cap=turma['Vagas']
        rows.append({'Turno':turma['Turno'],'Turma':turma['Turma'],'Vagas':cap,'Alocados':n,
                     'Feminino':f,'Masculino':m,'Feminino (%)':round(100*f/n,2) if n else 0.0,
                     'Masculino (%)':round(100*m/n,2) if n else 0.0,'Ocupação (%)':round(100*n/cap,2) if cap else 0.0})
    return pd.DataFrame(rows,columns=colunas)


def texto_resumo():
    df=resumo_distribuicao()
    if _estado.resultado!='resolvido':raise RuntimeError('Execute um ensalamento válido antes de exportar o resumo.')
    total=int(df['Alocados'].sum());f=int(df['Feminino'].sum());m=int(df['Masculino'].sum())
    linhas=['ABC Aprendiz — Resumo da distribuição',f'Solver: {_estado.status_solver}',
            f'Total alocado: {total} | Feminino: {f} | Masculino: {m}',
            'Percentuais de gênero referem-se aos estudantes de cada turma; turmas vazias têm percentuais zero.','']
    for turno,grupo in df.groupby('Turno',sort=False):
        n=int(grupo['Alocados'].sum());nf=int(grupo['Feminino'].sum());nm=int(grupo['Masculino'].sum())
        linhas.append(f'{turno}: {n} estudantes; Feminino {nf} ({100*nf/n if n else 0:.2f}%); Masculino {nm} ({100*nm/n if n else 0:.2f}%).')
        for row in grupo.to_dict('records'):
            linhas.append(f"  {row['Turma']}: {row['Alocados']}/{row['Vagas']} vagas; Feminino {row['Feminino']} ({row['Feminino (%)']:.2f}%); Masculino {row['Masculino']} ({row['Masculino (%)']:.2f}%).")
        linhas.append('')
    return '\n'.join(linhas)


def exportar_resumo(caminho):
    texto=texto_resumo();p,path=_destino(caminho,'.txt');temp=None
    try:
        fd,temp=tempfile.mkstemp(suffix='.txt',prefix='.unisoma-',dir=path.parent)
        with os.fdopen(fd,'w',encoding='utf-8') as stream:stream.write(texto)
        if digest(p.caminho)!=p.hash_original:raise RuntimeError('Entrada alterada externamente; importe novamente.')
        os.replace(temp,path)
    finally:
        if temp and os.path.exists(temp):os.unlink(temp)
    return f'Resumo exportado para: {path}.'
