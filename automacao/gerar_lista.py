"""Gera jogos.json (a lista de jogos da app de contagem) a partir do varrimento do site da Santa Casa.

Entradas:
  --site   Raspadinhas.xlsx gerado por "atualizar_raspadinhas v3.py" (jogo, nome, preço dos jogos ativos)
  jogos_extra.csv (nesta pasta)   jogos que ainda não estão no site (ex.: 718): jogo;preco;por_maco;nome
Bilhetes por maço fechado, pelo preço do bilhete (conferido com os talões): 1€ 100, 2€ 50,
3€ e 5€ 35, 10€ e 15€ 20. Um preço novo sem regra aqui fica com 0 (a app pede o valor à mão).

Uso:  python gerar_lista.py --site ../python/Raspadinhas.xlsx
"""
import argparse
import csv
import json
from datetime import datetime
from pathlib import Path

import openpyxl

AQUI = Path(__file__).resolve().parent
POR_MACO = {1: 100, 2: 50, 3: 35, 5: 35, 10: 20, 15: 20}


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--site", default=str(AQUI.parent / "python" / "Raspadinhas.xlsx"))
    p.add_argument("--extra", default=str(AQUI / "jogos_extra.csv"))
    p.add_argument("--saida", default=str(AQUI / "jogos.json"))
    p.add_argument("--minimo", type=int, default=0, help="falha se o site tiver menos jogos do que isto (proteção)")
    args = p.parse_args()
    ws = openpyxl.load_workbook(args.site, read_only=True).active
    jogos = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0]:
            jogos[int(r[0])] = {"jogo": int(r[0]), "nome": r[1], "preco": int(r[2])}
    if len(jogos) < args.minimo:
        raise SystemExit(f"Só {len(jogos)} jogos no site (mínimo {args.minimo}): não atualizo a lista.")
    extra = Path(args.extra)
    if extra.is_file():
        with open(extra, newline="", encoding="utf-8-sig") as f:
            for r in csv.DictReader(f, delimiter=";"):
                jogos.setdefault(int(r["jogo"]), {"jogo": int(r["jogo"]), "nome": r["nome"], "preco": int(r["preco"])})
    lista = sorted(jogos.values(), key=lambda j: (j["preco"], j["jogo"]))
    for j in lista:
        j["por_maco"] = POR_MACO.get(j["preco"], 0)
    saida = {"gerado": datetime.now().strftime("%Y-%m-%d %H:%M"), "jogos": lista}
    Path(args.saida).write_text(json.dumps(saida, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"{len(lista)} jogos em jogos.json")


if __name__ == "__main__":
    main()
