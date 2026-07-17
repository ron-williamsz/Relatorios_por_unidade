# -*- coding: utf-8 -*-
"""
Prepara os dados para o relatorio individual de consumo de agua por unidade.

Le os 5 snapshots mensais em json/todos_para_cobranca.json (gerados por
converter_para_json.py), faz o "pivot" por unidade (serie temporal mes a mes),
trata os problemas de qualidade dos dados e calcula deltas e resumos.

Saida: relatorio/dados_unidades.js  ->  window.DADOS = {...}
(carregado via <script src> para funcionar offline com duplo-clique, sem CORS).

Uso:
    python gerar_dados_relatorio.py
"""

import json
import os
import re
import unicodedata
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENTRADA = os.path.join(BASE_DIR, "json", "todos_para_cobranca.json")
SAIDA = os.path.join(BASE_DIR, "relatorio", "dados_unidades.js")

MESES_PT = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARCO": 3, "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
    "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12,
}
MESES_CURTO = ["", "Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
               "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

# Unidade renomeada apos troca de hidrometro (mesma unidade fisica).
RENOMEACOES = {"B0193B": "B0-193"}


def sem_acento(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def mes_ano_apuracao(apuracao):
    """'Apuracao NOVEMBRO/2025 - Emissao em ...' -> (2025, 11)."""
    m = re.search(r"([A-Za-zçÇ]+)\s*/\s*(\d{4})", sem_acento(apuracao or ""))
    if not m:
        return None
    mes = MESES_PT.get(m.group(1).upper())
    ano = int(m.group(2))
    return (ano, mes) if mes else None


def eh_numero(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def num(v):
    """Coage para numero; retorna (valor|None, tem_texto_bool).

    Strings de erro do Excel (#VALUE!, #REF!) ou texto livre viram None + flag.
    """
    if eh_numero(v):
        return v, False
    if v is None:
        return None, False
    # string: pode ser um numero em texto, um erro ou texto livre
    s = str(v).strip().replace(",", ".")
    if re.fullmatch(r"-?\d+(\.\d+)?", s):
        f = float(s)
        return (int(f) if f.is_integer() else f), False
    return None, True  # #VALUE!, #REF!, "trancado", "FECHADO", etc.


def arred(v, casas):
    if v is None:
        return None
    return round(v, casas) if casas else int(round(v))


def limpar_txt(v):
    if isinstance(v, str):
        v = v.strip()
        return v if v else None
    if v is None:
        return None
    return str(v)


def pct(atual, anterior):
    """Variacao percentual; None se base ausente ou <= 0 (base negativa distorce)."""
    if atual is None or anterior is None or anterior <= 0:
        return None
    return round((atual - anterior) / anterior * 100, 1)


def carregar_periodos(snapshots):
    """Ordena os snapshots e monta os metadados de periodo."""
    lista = []
    for idx, snap in enumerate(snapshots):
        meta = snap["metadados"]
        ma = mes_ano_apuracao(meta.get("apuracao"))
        if ma:
            ano, mes = ma
            chave = f"{ano}-{mes:02d}"
            curto = f"{MESES_CURTO[mes]}/{str(ano)[2:]}"
            completo = f"Apuração {MESES_CURTO[mes]}/{ano}"
        else:  # fallback improvavel
            chave, curto, completo, ano, mes = f"idx-{idx}", f"P{idx+1}", meta.get("apuracao", ""), 0, 0
        lista.append({
            "_idx": idx,
            "chave": chave,
            "ordem": ano * 100 + mes,
            "curto": curto,
            "completo": completo,
            "apuracao": meta.get("apuracao"),
            "data_leitura_anterior": meta.get("data_leitura_anterior"),
            "data_leitura_atual": meta.get("data_leitura_atual"),
            "possui_releitura": meta.get("possui_releitura"),
        })
    lista.sort(key=lambda p: p["ordem"])
    return lista


def chave_unidade(bloco, unidade):
    u = RENOMEACOES.get(unidade, unidade)
    return u, (u != unidade)


def main():
    with open(ENTRADA, encoding="utf-8") as fp:
        snapshots = json.load(fp)

    periodos = carregar_periodos(snapshots)
    ordem_periodo = {p["_idx"]: pos for pos, p in enumerate(periodos)}

    # unidades[chave] = {bloco, unidade, pontos_por_periodo{pos: ponto}, hidrometro}
    unidades = {}
    cond_serie = {p["chave"]: {"periodo": p["chave"], "consumo_total": 0.0,
                               "valor_total": 0.0, "unidades_com_consumo": 0}
                  for p in periodos}

    for snap in snapshots:
        idx = snapshots.index(snap)
        pos = ordem_periodo[idx]
        pchave = periodos[pos]["chave"]
        for r in snap["leituras"]:
            bloco = limpar_txt(r.get("bloco"))
            unidade_raw = limpar_txt(r.get("unidade"))
            if not bloco or not unidade_raw:
                continue
            unidade, trocou = chave_unidade(bloco, unidade_raw)

            consumo, c_txt = num(r.get("consumo_m3"))
            valor, v_txt = num(r.get("valor"))
            l_ant, la_txt = num(r.get("leitura_anterior"))
            l_atu, lu_txt = num(r.get("leitura_atual"))
            rele, re_txt = num(r.get("releitura"))

            consumo = arred(consumo, 0)
            valor = arred(valor, 2)

            flags = []
            if c_txt or v_txt:
                flags.append("erro_planilha")
            if la_txt or lu_txt or re_txt:
                flags.append("leitura_texto")
            if consumo is not None and consumo < 0:
                flags.append("consumo_negativo")
            if consumo == 0 and l_ant is not None and l_atu is not None and l_ant == l_atu:
                flags.append("medidor_parado")

            ponto = {
                "periodo": pchave,
                "leitura_anterior": arred(l_ant, 0),
                "leitura_atual": arred(l_atu, 0),
                "releitura": arred(rele, 0) if not re_txt else None,
                "consumo_m3": consumo,
                "valor": valor,
                "flags": flags,
                "_renomeado": trocou,  # True se este registro veio do id antigo (pre-troca)
            }

            u = unidades.setdefault(unidade, {
                "bloco": bloco, "unidade": unidade, "_pontos": {},
                "condicao_hidrometro": None, "estado_cavalete": None, "numero_lacre": None,
            })
            u["_pontos"][pos] = ponto
            # hidrometro: mantem o mais recente nao-vazio
            for campo in ("condicao_hidrometro", "estado_cavalete", "numero_lacre"):
                val = limpar_txt(r.get(campo))
                if val is not None:
                    u[campo] = val

            if consumo is not None:
                cond_serie[pchave]["consumo_total"] += consumo
                cond_serie[pchave]["valor_total"] += (valor or 0)
                if consumo != 0:
                    cond_serie[pchave]["unidades_com_consumo"] += 1

    # Monta serie ordenada + deltas + resumo por unidade.
    saida_unidades = {}
    for chave, u in unidades.items():
        serie = []
        anterior = None
        renomeado_antes = False
        troca_hidrometro = False
        for pos in range(len(periodos)):
            p = periodos[pos]
            ponto = u["_pontos"].get(pos)
            if ponto is None:
                # unidade sem registro naquele mes (raro): ponto vazio
                ponto = {"periodo": p["chave"], "leitura_anterior": None,
                         "leitura_atual": None, "releitura": None,
                         "consumo_m3": None, "valor": None, "flags": ["sem_registro"],
                         "_renomeado": False}
            # Troca de hidrometro = registro atual usa o id novo apos periodos com id antigo.
            if renomeado_antes and not ponto.pop("_renomeado", False):
                ponto["flags"] = ponto.get("flags", []) + ["troca_hidrometro"]
                troca_hidrometro = True
            else:
                renomeado_antes = renomeado_antes or ponto.pop("_renomeado", False)
            c = ponto["consumo_m3"]
            v = ponto["valor"]
            ca = anterior["consumo_m3"] if anterior else None
            va = anterior["valor"] if anterior else None
            ponto["delta_consumo_abs"] = (c - ca) if (c is not None and ca is not None) else None
            ponto["delta_consumo_pct"] = pct(c, ca)
            ponto["delta_valor_abs"] = round(v - va, 2) if (v is not None and va is not None) else None
            ponto["delta_valor_pct"] = pct(v, va)
            serie.append(ponto)
            if c is not None:
                anterior = ponto

        consumos = [p["consumo_m3"] for p in serie if p["consumo_m3"] is not None]
        valores = [p["valor"] for p in serie if p["valor"] is not None]
        meses_com_dado = len(consumos)
        consumo_total = sum(consumos) if consumos else 0
        valor_total = round(sum(valores), 2) if valores else 0
        # mes de pico (maior consumo)
        mes_pico = None
        if consumos:
            pico = max((p for p in serie if p["consumo_m3"] is not None),
                       key=lambda p: p["consumo_m3"])
            mes_pico = pico["periodo"]
        # variacao do ultimo mes com dado vs o anterior com dado
        variacao_ultimo_pct = None
        pontos_com_c = [p for p in serie if p["consumo_m3"] is not None]
        if len(pontos_com_c) >= 2:
            variacao_ultimo_pct = pct(pontos_com_c[-1]["consumo_m3"], pontos_com_c[-2]["consumo_m3"])

        saida_unidades[chave] = {
            "bloco": u["bloco"],
            "unidade": u["unidade"],
            "hidrometro": {
                "condicao": u["condicao_hidrometro"],
                "estado_cavalete": u["estado_cavalete"],
                "numero_lacre": u["numero_lacre"],
            },
            "serie": serie,
            "resumo": {
                "consumo_total": consumo_total,
                "valor_total": valor_total,
                "media_consumo": round(consumo_total / meses_com_dado, 1) if meses_com_dado else 0,
                "media_valor": round(valor_total / meses_com_dado, 2) if meses_com_dado else 0,
                "consumo_min": min(consumos) if consumos else None,
                "consumo_max": max(consumos) if consumos else None,
                "mes_pico": mes_pico,
                "variacao_ultimo_pct": variacao_ultimo_pct,
                "meses_com_dado": meses_com_dado,
                "troca_hidrometro": troca_hidrometro,
            },
        }

    # Blocos (contagem de unidades).
    blocos = {}
    for u in saida_unidades.values():
        blocos[u["bloco"]] = blocos.get(u["bloco"], 0) + 1
    blocos_lista = [{"codigo": b, "qtd": q} for b, q in sorted(blocos.items())]

    cond_meta = snapshots[0]["metadados"]
    saida = {
        "gerado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "condominio": {
            "associacao": cond_meta.get("associacao"),
            "titulo": cond_meta.get("titulo"),
        },
        "periodos": [{k: p[k] for k in p if not k.startswith("_")} for p in periodos],
        "blocos": blocos_lista,
        "condominio_serie": [
            {**cond_serie[p["chave"]],
             "valor_total": round(cond_serie[p["chave"]]["valor_total"], 2),
             "consumo_total": int(cond_serie[p["chave"]]["consumo_total"])}
            for p in periodos
        ],
        "unidades": saida_unidades,
    }

    os.makedirs(os.path.dirname(SAIDA), exist_ok=True)
    with open(SAIDA, "w", encoding="utf-8") as fp:
        fp.write("window.DADOS = ")
        json.dump(saida, fp, ensure_ascii=False, separators=(",", ":"))
        fp.write(";\n")

    print(f"OK -> {os.path.relpath(SAIDA, BASE_DIR)}")
    print(f"  periodos: {[p['curto'] for p in periodos]}")
    print(f"  unidades: {len(saida_unidades)}  |  blocos: {blocos}")
    print(f"  tamanho: {os.path.getsize(SAIDA)/1024:.0f} KB")


if __name__ == "__main__":
    main()
