# -*- coding: utf-8 -*-
"""
Baixa a planilha oficial de faturamento (Google Drive, link publico),
agrega as notas por SIGLA/mes/tipo de servico (P/C/E) e gera o HTML
final do dashboard a partir do template.

Rodar de novo a qualquer momento reflete o estado atual da planilha.
"""
import json
import os
import re
import sys
import urllib.request
from datetime import datetime, timezone, timedelta

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_URL = "https://docs.google.com/spreadsheets/d/12w6lJOdcIggtQpymebiHp5nbJijpdOQQ/export?format=xlsx"
XLSX_PATH = os.path.join(BASE_DIR, "_faturamento_cache.xlsx")
TEMPLATE_PATH = os.path.join(BASE_DIR, "template.html")
OUTPUT_PATH = os.path.join(BASE_DIR, "index.html")

MESES_ORDEM = ['JAN-26', 'FEV-26', 'MAR-26', 'ABR-26', 'MAI-26', 'JUN-26',
               'JUL-26', 'AGO-26', 'SET-26', 'OUT-26', 'NOV-26', 'DEZ-26']

# aliases conhecidos de SIGLA (variam entre meses na planilha)
ALIAS = {
    "SAM'S": "SAMS", "GBRABOSA": "GBARBOSA", "BM": "BANCO MERCANTIL", "BCO MERC": "BANCO MERCANTIL",
    "BCO MERCANTIL": "BANCO MERCANTIL", "CARREFOUR": "CRF", "BOTICARIO": "BOT",
    "BRADESCO": "BRAD", "IMOPAR": "OUTROS", "OUTLET": "OUTROS", "NUC BRAD": "IN HAUS",
    "SHOPPING PAMPLONA": "CRF", "SHOP PAMPLONA": "CRF", "VIVARA": "OUTROS", "LG": "OUTROS",
    "GM": "OUTROS", "OKEAN": "OUTROS", "MERCANTIL MERCADO": "MERCANTIL",
    "ATAKARE": "ATAKAREJO", "BAUDUCCO2": "BAUDUCCO",
    # descobertos lendo a planilha real em 2026-08-11:
    "BRAD AG": "AG BRAD", "BRADESCO AGENCIAS": "AG BRAD",
    "FUND BRADESCO": "FUND BRAD", "FUND. BRAD": "FUND BRAD",
    "BK": "ZAMP",
}

# SIGLA -> {nome, obj_key (chave em OBJ_CLIENTE, ou None se sem meta), grupo, icon, cor, bg}
SIGLA_MAP = {
    'AG BRAD': {'nome': 'Agências Bradesco', 'obj_key': None, 'grupo': 'AG BRAD', 'icon': 'B', 'cor': '#0A7B8A', 'bg': '#E6F6F8'},
    'BOT': {'nome': 'O Boticário', 'obj_key': 'O BOTICARIO', 'grupo': 'BOT', 'icon': 'O', 'cor': '#5B35B0', 'bg': '#F0ECFC'},
    'ZAMP': {'nome': 'ZAMP', 'obj_key': None, 'grupo': 'ZAMP', 'icon': 'Z', 'cor': '#C0200E', 'bg': '#FEF0EE'},
    'CRF': {'nome': 'Carrefour', 'obj_key': 'CARREFOUR', 'grupo': 'normal', 'icon': 'C', 'cor': '#1B3A6B', 'bg': '#EEF2FA'},
    'FUND BRAD': {'nome': 'Bradesco Fundação', 'obj_key': 'BRADESCO FUNDAÇÃO', 'grupo': 'normal', 'icon': 'BF', 'cor': '#1B3A6B', 'bg': '#EEF2FA'},
    'IN HAUS': {'nome': 'Bradesco In Haus', 'obj_key': 'BRADESCO IN HAUS', 'grupo': 'normal', 'icon': 'BI', 'cor': '#1B3A6B', 'bg': '#EEF2FA'},
    'BRAD': {'nome': 'Bradesco Eng.', 'obj_key': None, 'grupo': 'normal', 'icon': 'BE', 'cor': '#1B3A6B', 'bg': '#EEF2FA'},
    'LM': {'nome': 'Leroy Merlin', 'obj_key': 'LEROY MERLIN', 'grupo': 'normal', 'icon': 'L', 'cor': '#0E8A5A', 'bg': '#EAF7F1'},
    'LM_RESIDENTES': {'nome': 'Leroy Merlin (Residentes)', 'obj_key': 'LEROY MERLIN RESIDENTES', 'grupo': 'normal', 'icon': 'LR', 'cor': '#0E8A5A', 'bg': '#EAF7F1'},
    'SMTF': {'nome': 'Smart Fit', 'obj_key': 'SMART FIT', 'grupo': 'normal', 'icon': 'SF', 'cor': '#E84B1A', 'bg': '#FEF3EE'},
    'GBARBOSA': {'nome': 'GBarbosa', 'obj_key': 'GBARBOSA', 'grupo': 'normal', 'icon': 'G', 'cor': '#B07000', 'bg': '#FFF8E6'},
    'ASSAI': {'nome': 'Assaí', 'obj_key': 'ASSAI', 'grupo': 'normal', 'icon': 'A', 'cor': '#C0200E', 'bg': '#FEF0EE'},
    'SAMS': {'nome': "Sam's Club", 'obj_key': 'SAMS', 'grupo': 'normal', 'icon': 'S', 'cor': '#1B3A6B', 'bg': '#EEF2FA'},
    'ATAKAREJO': {'nome': 'Atakarejo', 'obj_key': 'ATAKAREJO', 'grupo': 'normal', 'icon': 'AT', 'cor': '#E84B1A', 'bg': '#FEF3EE'},
    'AUTOZONE': {'nome': 'AutoZone', 'obj_key': None, 'grupo': 'normal', 'icon': 'AZ', 'cor': '#C0200E', 'bg': '#FEF0EE'},
    'OBRAMAX': {'nome': 'Obramax', 'obj_key': 'OBRAMAX', 'grupo': 'normal', 'icon': 'OB', 'cor': '#0A7B8A', 'bg': '#E6F6F8'},
    'TIM': {'nome': 'TIM', 'obj_key': None, 'grupo': 'normal', 'icon': 'T', 'cor': '#1B3A6B', 'bg': '#EEF2FA'},
    'GPA': {'nome': 'GPA', 'obj_key': None, 'grupo': 'normal', 'icon': 'GP', 'cor': '#5B35B0', 'bg': '#F0ECFC'},
    'JLL': {'nome': 'JLL', 'obj_key': 'JLL', 'grupo': 'normal', 'icon': 'J', 'cor': '#0E8A5A', 'bg': '#EAF7F1'},
    'TENDA': {'nome': 'Tenda', 'obj_key': 'TENDA', 'grupo': 'normal', 'icon': 'TE', 'cor': '#B07000', 'bg': '#FFF8E6'},
    'MERCANTIL': {'nome': 'Mercantil (Mercado)', 'obj_key': 'MERCANTIL', 'grupo': 'normal', 'icon': 'ME', 'cor': '#1B3A6B', 'bg': '#EEF2FA'},
    'BANCO MERCANTIL': {'nome': 'Banco Mercantil', 'obj_key': 'BANCO MERCANTIL', 'grupo': 'normal', 'icon': 'BM', 'cor': '#0A7B8A', 'bg': '#E6F6F8'},
    'ACCENTURE': {'nome': 'Accenture', 'obj_key': 'ACCENTURE', 'grupo': 'normal', 'icon': 'AC', 'cor': '#5B35B0', 'bg': '#F0ECFC'},
    'ATENTO': {'nome': 'Atento', 'obj_key': 'ATENTO', 'grupo': 'normal', 'icon': 'AT', 'cor': '#0A7B8A', 'bg': '#E6F6F8'},
    'DIA': {'nome': 'DIA', 'obj_key': None, 'grupo': 'normal', 'icon': 'D', 'cor': '#E84B1A', 'bg': '#FEF3EE'},
    'BAUDUCCO': {'nome': 'Bauducco', 'obj_key': None, 'grupo': 'normal', 'icon': 'BA', 'cor': '#B07000', 'bg': '#FFF8E6'},
    'SIEMENS': {'nome': 'Siemens', 'obj_key': 'SIEMENS', 'grupo': 'normal', 'icon': 'SI', 'cor': '#0A7B8A', 'bg': '#E6F6F8'},
    'SINDILOJAS': {'nome': 'Sindilojas', 'obj_key': 'SINDILOJAS', 'grupo': 'normal', 'icon': 'SN', 'cor': '#0E8A5A', 'bg': '#EAF7F1'},
    'SUMERBOL': {'nome': 'Sumerbol', 'obj_key': 'SUMERBOL', 'grupo': 'normal', 'icon': 'SU', 'cor': '#5B35B0', 'bg': '#F0ECFC'},
    'SAPATARIA': {'nome': 'Sapataria Nova', 'obj_key': 'SAPATARIA', 'grupo': 'normal', 'icon': 'SA', 'cor': '#B07000', 'bg': '#FFF8E6'},
    'PORTO': {'nome': 'Porto Seguro', 'obj_key': 'PORTO', 'grupo': 'normal', 'icon': 'PS', 'cor': '#C0200E', 'bg': '#FEF0EE'},
    'ROVERI': {'nome': 'Roveri', 'obj_key': 'ROVERI', 'grupo': 'normal', 'icon': 'RO', 'cor': '#0E8A5A', 'bg': '#EAF7F1'},
    'GIGA': {'nome': 'Cencosud Giga', 'obj_key': 'GIGA', 'grupo': 'normal', 'icon': 'GI', 'cor': '#E84B1A', 'bg': '#FEF3EE'},
    'COBASI': {'nome': 'Cobasi', 'obj_key': None, 'grupo': 'normal', 'icon': 'CO', 'cor': '#0E8A5A', 'bg': '#EAF7F1'},
    'BIG': {'nome': 'BIG', 'obj_key': None, 'grupo': 'normal', 'icon': 'BI', 'cor': '#5B35B0', 'bg': '#F0ECFC'},
    'SMS': {'nome': 'SMS', 'obj_key': None, 'grupo': 'normal', 'icon': 'SM', 'cor': '#B07000', 'bg': '#FFF8E6'},
    'OUTROS': {'nome': 'Outros clientes', 'obj_key': None, 'grupo': 'normal', 'icon': 'OU', 'cor': '#8A96B0', 'bg': '#F0F2F7'},
    # clientes novos, encontrados na leitura da planilha real em 2026-08-11:
    'BRETAS': {'nome': 'Bretas', 'obj_key': None, 'grupo': 'normal', 'icon': 'BR', 'cor': '#0E8A5A', 'bg': '#EAF7F1'},
    'COGNA': {'nome': 'Cogna', 'obj_key': None, 'grupo': 'normal', 'icon': 'CG', 'cor': '#5B35B0', 'bg': '#F0ECFC'},
    'ENGEMON': {'nome': 'Engemon', 'obj_key': None, 'grupo': 'normal', 'icon': 'EG', 'cor': '#0A7B8A', 'bg': '#E6F6F8'},
    'RD': {'nome': 'RD (Raia/Drogasil)', 'obj_key': None, 'grupo': 'normal', 'icon': 'RD', 'cor': '#E84B1A', 'bg': '#FEF3EE'},
    'SENAC': {'nome': 'Senac', 'obj_key': 'SENAC', 'grupo': 'normal', 'icon': 'SC', 'cor': '#B07000', 'bg': '#FFF8E6'},
}

# tradução dos rótulos do bloco "Modelo / Metas Valores" da aba Objetivo -> chaves usadas no dashboard
GRUPO_KEY_MAP = {
    'PREVENTIVA TOTAL': 'PREVENTIVA TOTAL',
    'CORRETIVA': 'CORRETIVA',
    'ENGENHARIA': 'ENGENHARIA',
    'ZAMP - PREVENTIVA': 'ZAMP_P',
    'ZAMP - CORRETIVA': 'ZAMP_C',
    'BRADESCO AGENCIAS': 'BRADESCO AGENCIAS',
}


def log(msg):
    print(msg, file=sys.stderr)


def download_xlsx():
    log(f"Baixando planilha de {XLSX_URL} ...")
    req = urllib.request.Request(XLSX_URL, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = resp.read()
    if not data.startswith(b'PK'):
        raise RuntimeError("Download não retornou um .xlsx válido (sem login/permissão?)")
    with open(XLSX_PATH, 'wb') as f:
        f.write(data)
    log(f"OK, {len(data)} bytes salvos em {XLSX_PATH}")


def parse_valor(v):
    if isinstance(v, (int, float)):
        return float(v)
    if v is None:
        return None
    s = str(v).strip()
    s = re.sub(r'[Rr]\$', '', s).strip()
    if re.search(r',\d{2}$', s) and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif re.search(r',\d{2}$', s):
        s = s.replace(',', '.')
    else:
        s = s.replace(',', '')
    try:
        return float(s)
    except ValueError:
        return None


def build():
    import openpyxl
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True, read_only=True)
    meses_disponiveis = [m for m in MESES_ORDEM if m in wb.sheetnames]
    log(f"Abas de mês encontradas: {meses_disponiveis}")

    agregado = {}
    warnings = []
    siglas_nao_mapeadas = set()

    for mes in meses_disponiveis:
        ws = wb[mes]
        agregado[mes] = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row is None or len(row) < 8:
                continue
            valor, sigla, serv = row[2], row[6], row[7]
            obs = row[5] if len(row) > 5 else None
            if serv is None or sigla is None:
                continue
            serv = str(serv).strip().upper()
            if serv not in ('P', 'C', 'E', 'P-ENG'):
                continue
            sigla = str(sigla).strip()
            sigla = ALIAS.get(sigla, sigla)
            if sigla == 'LM' and obs is not None and 'RESIDENTE' in str(obs).upper():
                sigla = 'LM_RESIDENTES'
            cat = 'E' if serv == 'P-ENG' else serv
            val = parse_valor(valor)
            if val is None:
                warnings.append(f"{mes}: VALOR não numérico em linha SIGLA={sigla} SERV={serv}: {valor!r}")
                continue
            d = agregado[mes].setdefault(sigla, {'P': 0.0, 'C': 0.0, 'E': 0.0})
            d[cat] = round(d[cat] + val, 2)
            if sigla not in SIGLA_MAP:
                siglas_nao_mapeadas.add(sigla)

    # aba Objetivo: metas por cliente + metas por grupo
    ws = wb['Objetivo']
    obj_cliente = {}
    obj_grupo = {}
    modo_grupo = False
    for row in ws.iter_rows(min_row=1, values_only=True):
        if not row or row[0] is None:
            continue
        chave = str(row[0]).strip()
        if chave.lower() == 'cliente' or chave.lower() == 'modelo':
            modo_grupo = (chave.lower() == 'modelo')
            continue
        if chave.lower() == 'total geral':
            continue
        valor = row[1]
        if valor is None:
            continue
        try:
            valor = float(valor)
        except (TypeError, ValueError):
            continue
        if modo_grupo:
            gk = GRUPO_KEY_MAP.get(chave)
            if gk is None:
                warnings.append(f"Objetivo (grupo): rótulo não reconhecido '{chave}' ignorado")
                continue
            obj_grupo[gk] = valor
        else:
            if chave == 'SHOP. PAMPLONA':
                chave = 'CARREFOUR'
                valor += obj_cliente.get('CARREFOUR', 0.0)
            obj_cliente[chave] = valor

    for k in ('PREVENTIVA TOTAL', 'CORRETIVA', 'ENGENHARIA', 'ZAMP_P', 'ZAMP_C', 'BRADESCO AGENCIAS'):
        if k not in obj_grupo:
            warnings.append(f"Meta de grupo '{k}' não encontrada na aba Objetivo")

    # checagem: toda meta de cliente deveria estar referenciada por algum SIGLA_MAP.obj_key
    obj_keys_usadas = {v['obj_key'] for v in SIGLA_MAP.values() if v['obj_key']}
    metas_orfas = sorted(set(obj_cliente.keys()) - obj_keys_usadas)

    if siglas_nao_mapeadas:
        warnings.append(f"SIGLAs com faturamento real mas SEM entrada em SIGLA_MAP (não aparecem no painel!): {sorted(siglas_nao_mapeadas)}")
    if metas_orfas:
        warnings.append(f"Metas na aba Objetivo sem nenhum cliente do painel apontando para elas: {metas_orfas}")

    # mês padrão = último mês (na ordem) que tem faturamento real
    mes_atual = meses_disponiveis[0]
    for m in meses_disponiveis:
        total_mes = sum(sum(v.values()) for v in agregado[m].values())
        if total_mes > 0:
            mes_atual = m

    tz = timezone(timedelta(hours=-3))
    agora = datetime.now(tz).strftime('%d/%m/%Y %H:%M')

    with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
        html = f.read()

    def inject(marker, value_json):
        nonlocal html
        if marker not in html:
            raise RuntimeError(f"Marcador {marker} não encontrado no template")
        html = html.replace(marker, value_json, 1)

    inject('/*ALIAS_JSON*/', json.dumps(ALIAS, ensure_ascii=False))
    inject('/*MESES_JSON*/', json.dumps(meses_disponiveis, ensure_ascii=False))
    inject('/*MES_ATUAL_JSON*/', json.dumps(mes_atual, ensure_ascii=False))
    inject('/*AGREGADO_JSON*/', json.dumps(agregado, ensure_ascii=False))
    inject('/*OBJ_CLIENTE_JSON*/', json.dumps(obj_cliente, ensure_ascii=False))
    inject('/*OBJ_GRUPO_JSON*/', json.dumps(obj_grupo, ensure_ascii=False))
    inject('/*SIGLA_MAP_JSON*/', json.dumps(SIGLA_MAP, ensure_ascii=False))
    html = html.replace('/*GENERATED_AT*/', f'gerado em {agora} (America/Sao_Paulo)')
    html = html.replace('/*ATUALIZADO_EM*/', agora)

    with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
        f.write(html)

    log(f"\nHTML final escrito em {OUTPUT_PATH}")
    log(f"Mês padrão selecionado: {mes_atual}")
    log("\nTotal faturado por mês (P+C+E, todas as siglas):")
    for m in meses_disponiveis:
        tot = sum(sum(v.values()) for v in agregado[m].values())
        log(f"  {m}: R$ {tot:,.2f}")

    if warnings:
        log("\n=== AVISOS (revisar) ===")
        for w in warnings:
            log(f"  - {w}")
    else:
        log("\nNenhum aviso — todas as SIGLAs e metas foram mapeadas corretamente.")

    return {
        'mes_atual': mes_atual,
        'meses': meses_disponiveis,
        'warnings': warnings,
        'output_path': OUTPUT_PATH,
    }


if __name__ == '__main__':
    download_xlsx()
    build()
