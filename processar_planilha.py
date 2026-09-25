"""
PROCESSADOR DO PACOTE DE PREJUÍZO
==================================
Como usar:
1. Coloque este arquivo na mesma pasta que a planilha
2. Instale dependências: pip install openpyxl requests
3. Execute: python processar_planilha.py
4. Os dados serão enviados para a nuvem automaticamente

O dashboard em https://guilherme150497.github.io/dashboard-prejuizo
será atualizado para todos os usuários.
"""

import re, json, sys
from openpyxl import load_workbook
from collections import defaultdict

# ── CONFIGURAÇÃO ──
SUPABASE_URL = "https://mbteychnzitqvakqloes.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im1idGV5Y2hueml0cXZha3Fsb2VzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI1MDYyNDEsImV4cCI6MjA5ODA4MjI0MX0.0uoMcN0uU8lS0svEvI-iC8XH-Gge-unaunELoGrwi2M"
import glob

# Encontra automaticamente qualquer arquivo .xlsx na pasta
arquivos = glob.glob("*.xlsx")
if not arquivos:
    print("❌ Nenhum arquivo .xlsx encontrado nesta pasta!")
    print("   Coloque a planilha na mesma pasta que este script.")
    sys.exit(1)

ARQUIVO = arquivos[0]
if len(arquivos) > 1:
    print(f"⚠️  Múltiplos arquivos encontrados. Usando: {ARQUIVO}")

CODES = {
    (108,7000):'Quebra Entrega (Retorno de Rota)',(108,7008):'Acidente',
    (108,7015):'Refugo',(108,7013):'Assalto BO',(108,6996):'AmassadosEntr',
    (108,7002):'Quebra Armazém - Avaria',(108,7004):'Acidente Armazém',
    (108,6999):'Amassados Armazém',(108,7003):'Shelf Armazém',
    (108,7011):'Diferença de PA (Inventário)',(108,9100):'Erro de Programação',
    (108,7019):'Refugo Puxada',(108,7006):'Acidente Puxada',
    (108,7001):'Avarias no Recebimento (Puxada)',
}
# Lookup só pelo código (sem filtrar operação)
COD_MAP = {cod: nome for (op, cod), nome in CODES.items()}
WQI_CODES = {7002,6999,7004}
ALL_LINHAS = list(set(CODES.values())) + ['Diferença de AG','Reposição Entrega','Trocas Mercado (RN)']

# A partir de Set/2026 Reposição Entrega e Trocas Mercado (RN) saem de 03.18.05
NOVA_FONTE_ANO  = 2026
NOVA_FONTE_MES  = 9   # setembro

def nm(s):
    s=str(s).upper().strip()
    if 'LAGOA' in s or 'COBEB LP' in s: return 'COBEB LP'
    if 'PARA DE MINAS' in s or 'COBEB PM' in s: return 'COBEB PM'
    if 'RDC' in s or 'ABAET' in s: return 'RDC ABAETÉ'
    return None

def sf(v):
    try: return float(v) if isinstance(v,(int,float)) else 0.0
    except: return 0.0

def add(d,nome,base,ano,mes,brl,hl=0):
    for b in [base,'ALL']:
        d.setdefault(nome,{}).setdefault(b,{}).setdefault(ano,{}).setdefault(mes,{'brl':0.0,'hl':0.0})
        d[nome][b][ano][mes]['brl']+=brl; d[nome][b][ano][mes]['hl']+=hl

print(f"📂 Abrindo {ARQUIVO}...")
try:
    wb  = load_workbook(ARQUIVO, read_only=True, data_only=True)
    wb2 = load_workbook(ARQUIVO, read_only=True)
except FileNotFoundError:
    print(f"❌ Arquivo '{ARQUIVO}' não encontrado!")
    print("   Coloque este script na mesma pasta da planilha.")
    sys.exit(1)

# Lookup 01.11
print("🔍 Lendo cadastro de produtos (01.11)...")
ws11=wb['01.11']; lookup={}
for row in ws11.iter_rows(min_row=2,values_only=True):
    cod=row[0]
    if not isinstance(cod,(int,float)): continue
    marca=re.sub(r'^\d{3}\s*-\s*','',str(row[4] or '').strip())
    embal=re.sub(r'^\d{3}\s*-\s*','',str(row[6] or '').strip())
    lookup[int(cod)]={'marca':marca or 'Não Identificado','embal':embal or 'Não Identificado'}
print(f"   {len(lookup):,} produtos")

# Mensal
print("📊 Lendo 03.05.17 mensal...")
data={}; rawProd=[]; wqiMensal={}; volEntregue={}
wqiMetaAcc=defaultdict(lambda:{'hl_perd':0.0,'hl_vend':0.0})
n=0
for row in wb['03.05.17'].iter_rows(min_row=2,values_only=True):
    dt=row[3]; base=nm(row[2])
    if not hasattr(dt,'month') or not base: continue
    op=int(row[6]) if isinstance(row[6],(int,float)) else 0
    cod=int(row[8]) if isinstance(row[8],(int,float)) else 0
    hl=sf(row[20]); vl=sf(row[32])
    prod=str(row[18]).strip() if row[18] else ''
    cod_prod=int(row[16]) if isinstance(row[16],(int,float)) else None
    # WQI e Volume Entregue: ANTES do filtro de nome (ops 1 e 2 não têm nome)
    ym=f"{dt.year}-{dt.month:02d}"
    wqiMensal.setdefault(base,{}).setdefault(ym,{'hl_perd':0.0,'hl_vend':0.0})
    if cod in WQI_CODES: wqiMensal[base][ym]['hl_perd']+=hl
    if op in [1,2]: wqiMensal[base][ym]['hl_vend']+=hl
    # Volume entregue real (ops 1+2) — usado como denominador do R$/HL
    volEntregue.setdefault(base,{}).setdefault(ym,0.0)
    if op in [1,2]: volEntregue[base][ym]+=hl
    if dt.year==2025:
        if cod in WQI_CODES: wqiMetaAcc[base]['hl_perd']+=hl
        if op in [1,2]: wqiMetaAcc[base]['hl_vend']+=hl

    nome=COD_MAP.get(cod) or ('Diferença de AG' if op==107 else None)
    if not nome: continue
    add(data,nome,base,dt.year,dt.month,vl,hl)
    if prod and vl:
        info=lookup.get(cod_prod,{}) if cod_prod else {}
        rawProd.append({'base':base,'ano':dt.year,'mes':dt.month,'linha':nome,'prod':prod,
            'brl':round(vl,2),'hl':round(hl,4),'marca':info.get('marca',''),'embal':info.get('embal','')})
    n+=1
    if n%50000==0: print(f"   {n:,} linhas...")
print(f"   ✓ {n:,} linhas processadas, {len(rawProd):,} produtos")

# ── 03.05.28.01 — Reposição Entrega e Trocas Mercado (RN) até Ago/2026 ──
# A partir de Set/2026 essas linhas saem de 03.18.05 (ver abaixo)
print("📊 Lendo 03.05.28.01 (até Ago/2026)...")
n_rep=0
for row in wb['03.05.28.01'].iter_rows(min_row=2,values_only=True):
    if str(row[2] or '').strip()!='Total Gerente': continue
    dt=row[1]; base=nm(row[0])
    if not hasattr(dt,'month') or not base: continue
    # Somente meses ANTES de Set/2026 — a partir de Set/2026 usa 03.18.05
    if dt.year > NOVA_FONTE_ANO: continue
    if dt.year == NOVA_FONTE_ANO and dt.month >= NOVA_FONTE_MES: continue
    vR=sf(row[4]); vT=sum(sf(row[i]) for i in [3,5,6,7,8,9,10])
    if vR: add(data,'Reposição Entrega',base,dt.year,dt.month,vR); n_rep+=1
    if vT: add(data,'Trocas Mercado (RN)',base,dt.year,dt.month,vT); n_rep+=1
print(f"   ✓ {n_rep} registros (somente até Ago/2026)")

# ── 03.18.05 — Reposição Entrega e Trocas Mercado (RN) a partir de Set/2026 ──
# Regras:
#   Col A (idx 0)  = Base (filial)
#   Col H (idx 7)  = Data da solicitação
#   Col J (idx 9)  = Status — somente "Aprovada"
#   Col O (idx 14) = Status NF — somente "E" ou VAZIO (excluir "D")
#   Col P (idx 15) = Código produto
#   Col R (idx 17) = Descrição produto
#   Col U (idx 20) = Valor (prejuízo)
#   Col X (idx 23) = Justificativa — EXCLUIR se contiver "Falta de Produto"
#   Col BL (idx 63)= Sistema Origem — "Promax" → Reposição Entrega; "Force" → Trocas Mercado (RN)
print("📊 Lendo 03.18.05 (a partir de Set/2026)...")
n_nova=0
for row in wb['03.18.05'].iter_rows(min_row=2,values_only=True):
    base = nm(row[0])
    if not base: continue

    dt = row[7]   # Col H — data
    if not hasattr(dt,'month'): continue

    # Somente SET/2026 em diante
    if dt.year < NOVA_FONTE_ANO: continue
    if dt.year == NOVA_FONTE_ANO and dt.month < NOVA_FONTE_MES: continue

    # Col J — Status Solicitação: somente "Aprovada"
    status_sol = str(row[9] or '').strip()
    if status_sol != 'Aprovada': continue

    # Col O — Status NF: somente "E" ou VAZIO (excluir "D")
    status_nf = str(row[14] or '').strip()
    if status_nf == 'D': continue

    # Col X — Justificativa: excluir se contiver "Falta de Produto"
    justificativa = str(row[23] or '').strip()
    if 'Falta de Produto' in justificativa: continue

    # Col BL (idx 63) — Sistema Origem
    sistema = str(row[63] or '').strip()
    if sistema == 'Promax':
        linha = 'Reposição Entrega'
    elif sistema == 'Force':
        linha = 'Trocas Mercado (RN)'
    else:
        continue  # ignora linhas sem sistema reconhecido

    vl = sf(row[20])   # Col U — valor/custo
    if vl <= 0: continue

    # Col P (idx 15) — código produto, Col R (idx 17) — descrição produto
    cod_prod = int(row[15]) if isinstance(row[15],(int,float)) else None
    prod     = str(row[17] or '').strip()

    add(data, linha, base, dt.year, dt.month, vl)
    if prod and vl:
        info = lookup.get(cod_prod, {}) if cod_prod else {}
        rawProd.append({
            'base': base, 'ano': dt.year, 'mes': dt.month,
            'linha': linha, 'prod': prod,
            'brl': round(vl, 2), 'hl': 0.0,
            'marca': info.get('marca',''), 'embal': info.get('embal','')
        })
    n_nova += 1
    if n_nova % 10000 == 0: print(f"   {n_nova:,} linhas 03.18.05...")
print(f"   ✓ {n_nova:,} linhas 03.18.05 processadas")

# ── Diário — 03.05.17 ──
print("📊 Lendo 03.05.17 diária...")
rawDiario=[]; wqiDiario={}; n=0
for row in wb['03.05.17 diária'].iter_rows(min_row=2,values_only=True):
    dt=row[1]; base=nm(row[0])
    if not hasattr(dt,'date') or not base: continue
    op=int(row[3]) if isinstance(row[3],(int,float)) else 0
    cod=int(row[5]) if isinstance(row[5],(int,float)) else 0
    vol=sf(row[17]); vl=sf(row[29])
    prod=str(row[15]).strip() if row[15] else ''
    cod_prod=int(row[13]) if isinstance(row[13],(int,float)) else None
    ds=dt.strftime('%Y-%m-%d')
    wqiDiario.setdefault(base,{}).setdefault(ds,{'hl_perd':0.0,'hl_vend':0.0})
    if cod in WQI_CODES: wqiDiario[base][ds]['hl_perd']+=vol
    if op in [1,2]: wqiDiario[base][ds]['hl_vend']+=vol
    nome=COD_MAP.get(cod) or ('Diferença de AG' if op==107 else None)
    if not nome: continue
    info=lookup.get(cod_prod,{}) if cod_prod else {}
    rawDiario.append({'base':base,'data':ds,'linha':nome,'brl':round(vl,2),'hl':round(vol,4),
        'prod':prod,'marca':info.get('marca',''),'embal':info.get('embal',''),
        'wqi':cod in WQI_CODES,'op':op})
    n+=1
    if n%100000==0: print(f"   {n:,} linhas...")
print(f"   ✓ {n:,} linhas, {len(rawDiario):,} registros pacote")

# ── Diário — 03.05.28.01 (somente até Ago/2026) ──
print("📊 Lendo 03.05.28.01 diário (até Ago/2026)...")
n_drep=0
for row in wb['03.05.28.01 diário'].iter_rows(min_row=2,values_only=True):
    if str(row[2] or '').strip()!='Total Gerente': continue
    dt=row[1]; base=nm(row[0])
    if not hasattr(dt,'date') or not base: continue
    # Somente meses ANTES de Set/2026
    if dt.year > NOVA_FONTE_ANO: continue
    if dt.year == NOVA_FONTE_ANO and dt.month >= NOVA_FONTE_MES: continue
    ds=dt.strftime('%Y-%m-%d')
    vR=sf(row[4]); vT=sum(sf(row[i]) for i in [3,5,6,7,8,9,10])
    if vR:
        rawDiario.append({'base':base,'data':ds,'linha':'Reposição Entrega','brl':round(vR,2),'hl':0,'prod':'','marca':'','embal':'','wqi':False,'op':0})
        n_drep+=1
    if vT:
        rawDiario.append({'base':base,'data':ds,'linha':'Trocas Mercado (RN)','brl':round(vT,2),'hl':0,'prod':'','marca':'','embal':'','wqi':False,'op':0})
        n_drep+=1
print(f"   ✓ {n_drep} registros diários (somente até Ago/2026)")

# ── Diário — 03.18.05 (a partir de Set/2026) ──
print("📊 Lendo 03.18.05 diário (a partir de Set/2026)...")
n_dnova=0
for row in wb['03.18.05'].iter_rows(min_row=2,values_only=True):
    base = nm(row[0])
    if not base: continue

    dt = row[7]   # Col H
    if not hasattr(dt,'date'): continue

    # Somente SET/2026 em diante
    if dt.year < NOVA_FONTE_ANO: continue
    if dt.year == NOVA_FONTE_ANO and dt.month < NOVA_FONTE_MES: continue

    # Filtros
    if str(row[9] or '').strip() != 'Aprovada': continue
    if str(row[14] or '').strip() == 'D': continue
    if 'Falta de Produto' in str(row[23] or ''): continue

    sistema = str(row[63] or '').strip()
    if sistema == 'Promax':
        linha = 'Reposição Entrega'
    elif sistema == 'Force':
        linha = 'Trocas Mercado (RN)'
    else:
        continue

    vl = sf(row[20])
    if vl <= 0: continue

    cod_prod = int(row[15]) if isinstance(row[15],(int,float)) else None
    prod     = str(row[17] or '').strip()
    ds       = dt.strftime('%Y-%m-%d')
    info     = lookup.get(cod_prod, {}) if cod_prod else {}

    rawDiario.append({
        'base': base, 'data': ds, 'linha': linha,
        'brl': round(vl, 2), 'hl': 0.0,
        'prod': prod, 'marca': info.get('marca',''), 'embal': info.get('embal',''),
        'wqi': False, 'op': 0
    })
    n_dnova += 1
print(f"   ✓ {n_dnova} registros diários 03.18.05")

# Metas
print("📊 Lendo BASE_METAS...")
SKIP={'entrega','armazém','armazem','puxada','cobeb pm','cobeb lp','rdc abaeté','meta','janeiro',''}
rows_bm=list(wb2['BASE_METAS'].iter_rows(min_row=1,max_row=35,values_only=True))
vol_pm=[float(rows_bm[2][6+m] or 0) for m in range(12)]
vol_lp=[float(rows_bm[3][6+m] or 0) for m in range(12)]
vol_ab=[float(rows_bm[4][6+m] or 0) for m in range(12)]
metas={}
for ri in range(6,len(rows_bm)):
    row=rows_bm[ri]; nome=str(row[4] or '').strip()
    # Na BASE_METAS, 'Diferença de AG' representa 'Erro de Programação'
    if nome == 'Diferença de AG': nome = 'Erro de Programação'
    if not nome or nome.lower() in SKIP or nome not in ALL_LINHAS: continue
    for m in range(12):
        for bk,taxa,vol in [('COBEB PM',float(row[5] or 0),vol_pm),
                            ('COBEB LP',float(row[6] or 0),vol_lp),
                            ('RDC ABAETÉ',float(row[7] or 0),vol_ab)]:
            v=round(taxa*vol[m],2)
            if v<=0: continue
            for ano in [2025,2026]:
                metas.setdefault(nome,{}).setdefault(bk,{}).setdefault(ano,{})[m+1]=v
                metas.setdefault(nome,{}).setdefault('ALL',{}).setdefault(ano,{})
                metas[nome]['ALL'][ano][m+1]=round(metas[nome]['ALL'][ano].get(m+1,0)+v,2)

# Volume Real (Venda) e Volume Puxado — aba VOLUME_REAL
# Localiza as linhas de cada base dinamicamente (ignora cabeçalhos com texto como 'JANEIRO')
rows_vr=list(wb2['VOLUME_REAL'].iter_rows(min_row=1,max_row=30,values_only=True))

def sf_cell(v):
    try: return float(v or 0)
    except: return 0.0

# Encontra índices de cada base (pode aparecer 2x: tabela venda + tabela puxado)
idx_pm,idx_lp,idx_ab=[],[],[]
for i,row in enumerate(rows_vr):
    if not row or len(row)<3: continue
    bv=str(row[0] or '').upper().strip()
    if not bv: continue
    if 'PARA DE MINAS' in bv or 'COBEB PM' in bv: idx_pm.append(i)
    elif 'LAGOA' in bv or 'COBEB LP' in bv:       idx_lp.append(i)
    elif 'RDC' in bv or 'ABAET' in bv:            idx_ab.append(i)
print(f"   VOLUME_REAL idx: PM={idx_pm}, LP={idx_lp}, AB={idx_ab}")

# Tabela 1 — Volume Venda (primeira ocorrência de cada base)
vol_real={'ALL':{}}
for base,idxs in [('COBEB PM',idx_pm),('COBEB LP',idx_lp),('RDC ABAETÉ',idx_ab)]:
    if not idxs: continue
    vol_real[base]={}
    row=rows_vr[idxs[0]]
    for m in range(12):
        v=sf_cell(row[2+m] if 2+m<len(row) else None)
        if v:
            for ano in [2025,2026]:
                vol_real[base].setdefault(ano,{})[m+1]=v
                vol_real['ALL'].setdefault(ano,{})[m+1]=vol_real['ALL'].get(ano,{}).get(m+1,0)+v

# Tabela 2 — Volume Puxado (segunda ocorrência de cada base)
vol_puxado={'ALL':{}}
for base,idxs in [('COBEB PM',idx_pm),('COBEB LP',idx_lp),('RDC ABAETÉ',idx_ab)]:
    if len(idxs)<2: print(f"   ⚠️  {base}: só 1 tabela encontrada, volPuxado não preenchido"); continue
    vol_puxado[base]={}
    row=rows_vr[idxs[1]]
    for m in range(12):
        v=sf_cell(row[2+m] if 2+m<len(row) else None)
        if v:
            for ano in [2025,2026]:
                vol_puxado[base].setdefault(ano,{})[m+1]=v
                vol_puxado['ALL'].setdefault(ano,{})[m+1]=vol_puxado['ALL'].get(ano,{}).get(m+1,0)+v
print(f"   volReal LP Jan/26: {vol_real.get('COBEB LP',{}).get(2026,{}).get(1,'—')}")
print(f"   volPuxado LP Jan/26: {vol_puxado.get('COBEB LP',{}).get(2026,{}).get(1,'—')}")

# WQI Meta
wqiMeta={}
for b,v in wqiMetaAcc.items():
    wqiMeta[b]=round(v['hl_perd']/v['hl_vend']*1e6,2) if v['hl_vend']>0 else 0
tp=sum(v['hl_perd'] for v in wqiMetaAcc.values())
tv=sum(v['hl_vend'] for v in wqiMetaAcc.values())
wqiMeta['ALL']=round(tp/tv*1e6,2) if tv>0 else 0

def fk(d):
    if isinstance(d,dict): return {str(k):fk(v) for k,v in d.items()}
    return d

payload = {'data':fk(data),'metas':fk(metas),'vol':fk(vol_real),'volPuxado':fk(vol_puxado),'volEntregue':volEntregue,
           'rawProd':rawProd,'rawDiario':rawDiario,
           'wqiMensal':wqiMensal,'wqiDiario':wqiDiario,'wqiMeta':wqiMeta}

# Validação
tot=sum(sum(data.get(n,{}).get('COBEB LP',{}).get(2026,{}).get(m,{}).get('brl',0) for m in range(1,13)) for n in ALL_LINHAS)
print(f"\n✅ Validação: Total COBEB LP 2026 = R$ {tot:,.2f}")
print(f"   rawProd: {len(rawProd):,} | rawDiario: {len(rawDiario):,}")
print(f"   WQI Meta: {wqiMeta}")

# Salvar JSON local
json_str = json.dumps(payload, ensure_ascii=False)
print(f"\n💾 Tamanho dos dados: {len(json_str)/1024/1024:.1f} MB")

with open('dados_dashboard.json','w',encoding='utf-8') as f:
    f.write(json_str)
print("   Salvo em dados_dashboard.json")

# Enviar para Supabase
print("\n☁️  Enviando para a nuvem...")
try:
    import requests
    res = requests.patch(
        f"{SUPABASE_URL}/rest/v1/dados_dashboard?id=eq.dados",
        headers={
            'Content-Type': 'application/json',
            'apikey': SUPABASE_KEY,
            'Authorization': f'Bearer {SUPABASE_KEY}',
            'Prefer': 'return=minimal'
        },
        json={'payload': payload}
    )
    if res.status_code in [200,204]:
        print("✅ Dados enviados com sucesso!")
        print("   Todos os usuários verão os dados atualizados ao abrir o dashboard.")
    else:
        print(f"❌ Erro ao enviar: {res.status_code} - {res.text[:200]}")
except ImportError:
    print("⚠️  Instale requests: pip install requests")
    print("   Os dados foram salvos em dados_dashboard.json")
except Exception as e:
    print(f"❌ Erro de conexão: {e}")
    print("   Os dados foram salvos em dados_dashboard.json")

print("\n✅ Processamento concluído!")
