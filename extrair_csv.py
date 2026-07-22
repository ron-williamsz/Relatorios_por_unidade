# -*- coding: utf-8 -*-
"""
Extrai leituras "Para Cobrança" exportadas em CSV (mesma tabela por lote dos
.xlsx, mas em CSV com números no formato pt-BR).

Trata as particularidades do CSV:
- Números pt-BR: "47,60" -> 47.6 ; "50.231,00" -> 50231 ; "7,00" -> 7
- Metadados do topo podem estar DESATUALIZADOS (copiados do mês anterior).
  Por isso o período é derivado da DATA REAL da leitura (no nome do arquivo e
  no cabeçalho da coluna atual), não das linhas 3-5.
- Colunas mapeadas por NOME do cabeçalho (com ou sem "Releitura").

Saída: json/<slug>.json  +  mescla em json/todos_para_cobranca.json
(depois rode gerar_dados_relatorio.py para atualizar o relatório).

Uso:
    python extrair_csv.py                # processa todo *.csv de leitura na raiz
    python extrair_csv.py "arquivo.csv"  # processa um específico
"""

import csv
import glob
import json
import os
import re
import sys

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_DIR = os.path.join(BASE_DIR, "json")
COMBINADO = os.path.join(JSON_DIR, "todos_para_cobranca.json")

MESES = {"JANEIRO": 1, "FEVEREIRO": 2, "MARÇO": 3, "MARCO": 3, "ABRIL": 4, "MAIO": 5,
         "JUNHO": 6, "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10,
         "NOVEMBRO": 11, "DEZEMBRO": 12}
NOME = ["", "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO",
        "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]

CHAVES = ["bloco", "unidade", "leitura_anterior", "leitura_atual", "releitura",
          "consumo_m3", "valor", "condicao_hidrometro", "estado_cavalete", "numero_lacre"]


def limpar(v):
    if isinstance(v, str):
        v = v.strip()
        return v if v else None
    return v


def num_br(v):
    """Coage número pt-BR -> int/float. Texto/erro -> a própria string (ou None)."""
    if v is None:
        return None
    s = str(v).strip()
    if s == "":
        return None
    if re.fullmatch(r"-?[\d.]*\d,\d+|-?\d[\d.]*", s):  # tem cara de número pt-BR
        try:
            f = float(s.replace(".", "").replace(",", "."))
            return int(f) if f.is_integer() else round(f, 2)
        except ValueError:
            pass
    return s  # ex.: '#VALUE!', 'trancado'


def classificar(h):
    if not isinstance(h, str):
        return None
    t = h.strip().lower()
    if t == "bloco": return "bloco"
    if t == "unidade": return "unidade"
    if t == "anterior": return "leitura_anterior"
    if t.startswith("releitura"): return "releitura"
    if t in ("m³", "m3") or t.startswith("m³") or t.startswith("m3"): return "consumo_m3"
    if t == "valor": return "valor"
    if t.startswith("condi"): return "condicao_hidrometro"
    if t.startswith("estado"): return "estado_cavalete"
    if "lacre" in t: return "numero_lacre"
    return None


def meta_do_nome(nome_arquivo):
    """Deriva (data_leitura_atual, apuracao) do nome do arquivo.

    Ex.: 'Leitura de água - 29.06.2026.xlsx - AGOSTO 2026.csv'
         -> data '29/06/2026', emissão 'AGOSTO/2026', apuração mês = junho.
    """
    dt = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", nome_arquivo)
    emis = re.search(r"-\s*([A-Za-zÇç]+)\s+(\d{4})\.csv$", nome_arquivo)
    data_atual = f"{dt.group(1)}/{dt.group(2)}/{dt.group(3)}" if dt else None
    mes = int(dt.group(2)) if dt else None
    ano = dt.group(3) if dt else None
    apur = None
    if mes and ano:
        emis_txt = f"{emis.group(1).upper()}/{emis.group(2)}" if emis else "?"
        apur = f"Apuração {NOME[mes]}/{ano} - Emissão em {emis_txt}"
    return data_atual, apur


def achar_header(linhas):
    for i, row in enumerate(linhas):
        if any(isinstance(c, str) and c.strip() == "Bloco" for c in row):
            return i
    return None


def processar_csv(caminho, data_anterior=None):
    with open(caminho, encoding="utf-8-sig", newline="") as fp:
        linhas = [[c.strip() if isinstance(c, str) else c for c in row]
                  for row in csv.reader(fp)]

    hdr = achar_header(linhas)
    if hdr is None:
        return None  # não é uma planilha de leitura

    cab = linhas[hdr]
    mapa = {}
    for c, h in enumerate(cab):
        k = classificar(h)
        if k and k not in mapa:
            mapa[k] = c
    # leitura_atual = coluna logo após "Anterior"
    rotulo_atual = None
    if "leitura_anterior" in mapa:
        ca = mapa["leitura_anterior"] + 1
        if ca not in mapa.values():
            mapa["leitura_atual"] = ca
            rotulo_atual = cab[ca] if ca < len(cab) else None

    data_atual, apuracao = meta_do_nome(os.path.basename(caminho))

    registros = []
    for row in linhas[hdr + 1:]:
        def cel(k):
            i = mapa.get(k)
            return row[i] if (i is not None and i < len(row)) else None
        bloco = limpar(cel("bloco"))
        unidade = limpar(cel("unidade"))
        if not bloco or not unidade:
            continue  # pula vazias e a linha "total"
        reg = {}
        for k in CHAVES:
            v = cel(k)
            if k in ("bloco", "unidade", "condicao_hidrometro", "estado_cavalete"):
                reg[k] = limpar(v)
            else:
                reg[k] = num_br(v)
        registros.append(reg)

    meta = {
        "arquivo_origem": os.path.relpath(caminho, BASE_DIR).replace("\\", "/"),
        "pasta": "(raiz)",
        "planilha": None,
        "associacao": limpar(linhas[0][0]) if linhas and linhas[0] else None,
        "titulo": limpar(linhas[1][0]) if len(linhas) > 1 and linhas[1] else None,
        "apuracao": apuracao,
        "descricao_leitura_anterior": f"L.Anterior = Leitura feita em {data_anterior}" if data_anterior else None,
        "descricao_leitura_atual": f"L.Atual = Leitura feita em {data_atual}" if data_atual else None,
        "data_leitura_anterior": data_anterior,
        "data_leitura_atual": data_atual,
        "rotulo_coluna_atual": rotulo_atual,
        "rotulo_coluna_releitura": None,
        "possui_releitura": "releitura" in mapa,
        "total_registros": len(registros),
    }
    return {"metadados": meta, "leituras": registros}


def slug(caminho):
    base = os.path.splitext(os.path.basename(caminho))[0]
    return re.sub(r"[^0-9A-Za-zÀ-ÿ]+", "_", base).strip("_") + ".json"


def data_anterior_do_combinado(combinado):
    """A leitura 'anterior' do novo período = a leitura ATUAL mais recente já existente."""
    if not combinado:
        return None
    def chave(s):
        d = s["metadados"].get("data_leitura_atual") or ""
        m = re.match(r"(\d{2})/(\d{2})/(\d{4})", d)
        return (int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else (0, 0, 0)
    return max(combinado, key=chave)["metadados"].get("data_leitura_atual")


def main():
    alvos = sys.argv[1:] or [f for f in glob.glob(os.path.join(BASE_DIR, "*.csv"))]
    combinado = json.load(open(COMBINADO, encoding="utf-8")) if os.path.exists(COMBINADO) else []

    processados = 0
    for caminho in alvos:
        data_ant = data_anterior_do_combinado(combinado)
        dados = processar_csv(caminho, data_anterior=data_ant)
        if dados is None:
            print(f"IGNORADO (sem cabeçalho 'Bloco'): {os.path.basename(caminho)}")
            continue

        # grava individual
        saida = os.path.join(JSON_DIR, slug(caminho))
        os.makedirs(JSON_DIR, exist_ok=True)
        with open(saida, "w", encoding="utf-8") as fp:
            json.dump(dados, fp, ensure_ascii=False, indent=2)

        # mescla no combinado (substitui se já houver o mesmo data_leitura_atual)
        alvo_data = dados["metadados"]["data_leitura_atual"]
        combinado = [s for s in combinado if s["metadados"].get("data_leitura_atual") != alvo_data]
        combinado.append(dados)
        processados += 1
        m = dados["metadados"]
        print(f"OK  {m['arquivo_origem']}")
        print(f"    período: {m['apuracao']}")
        print(f"    leitura: {m['data_leitura_anterior']} -> {m['data_leitura_atual']}  "
              f"| {m['total_registros']} unidades | releitura={m['possui_releitura']}")

    if processados:
        with open(COMBINADO, "w", encoding="utf-8") as fp:
            json.dump(combinado, fp, ensure_ascii=False, indent=2)
        print(f"\nCombinado atualizado: {len(combinado)} períodos -> json/todos_para_cobranca.json")
        print("Agora rode:  python gerar_dados_relatorio.py")


if __name__ == "__main__":
    main()
