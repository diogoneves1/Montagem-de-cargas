"""
Leitura da aba PLANO a partir de diferentes formatos de arquivo
(.xlsx, .xlsb, .xlsm, .csv), identificando automaticamente a linha
de cabeçalho e as colunas relevantes (SKU, Produto, Família, Qtd,
Peso, Pallets, Lote).
"""

from __future__ import annotations

import io
import re

import pandas as pd


COLUNAS_ESPERADAS = {
    "sku": ["SKU"],
    "produto": ["PRODUTO"],
    "familia": ["FAMÍLIA", "FAMILIA"],
    "qtd": ["QTD", "QTD.", "QUANTIDADE"],
    "peso": ["PESO BRUTO", "PESO"],
    "pallets": ["PALLETS", "PALLET"],
    "lote": ["LOTE"],
}


def _normalizar(texto: str) -> str:
    """Remove acentos/quebras de linha e deixa em maiúsculas para
    comparação tolerante de nomes de coluna."""
    if texto is None:
        return ""
    texto = str(texto).upper()
    texto = texto.replace("\n", " ").replace("\r", " ")
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def _mapear_colunas(header_vals: list) -> dict[str, int]:
    """Dado os valores de uma linha de cabeçalho, retorna um dict
    {nome_padronizado: índice_da_coluna}."""
    normalizados = [_normalizar(v) for v in header_vals]
    mapa = {}
    for padrao, alternativas in COLUNAS_ESPERADAS.items():
        for idx, val in enumerate(normalizados):
            if any(val.startswith(_normalizar(alt)) for alt in alternativas):
                mapa[padrao] = idx
                break
    return mapa


def _localizar_header_e_mapa(matriz: list[list]) -> tuple[int, dict[str, int]]:
    """Procura, nas primeiras ~40 linhas, a linha que contém pelo menos
    as colunas SKU + PRODUTO + uma medida de quantidade/peso — essa é
    a linha de cabeçalho da tabela de itens."""
    for i, row in enumerate(matriz[:40]):
        mapa = _mapear_colunas(row)
        if "sku" in mapa and "produto" in mapa and ("qtd" in mapa or "peso" in mapa):
            return i, mapa
    raise ValueError(
        "Não foi possível localizar a linha de cabeçalho da tabela "
        "(esperado pelo menos as colunas SKU e PRODUTO)."
    )


def _matriz_de_xlsb(file_bytes: bytes, aba: str) -> list[list]:
    from pyxlsb import open_workbook

    with open_workbook(io.BytesIO(file_bytes)) as wb:
        nomes = wb.sheets
        nome_real = _resolver_nome_aba(nomes, aba)
        with wb.get_sheet(nome_real) as ws:
            return [[c.v for c in row] for row in ws.rows()]


def _matriz_de_excel(file_bytes: bytes, aba: str, engine: str | None = None) -> list[list]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine=engine)
    nome_real = _resolver_nome_aba(xls.sheet_names, aba)
    df = pd.read_excel(xls, sheet_name=nome_real, header=None)
    return df.values.tolist()


def _matriz_de_csv(file_bytes: bytes) -> list[list]:
    df = pd.read_csv(io.BytesIO(file_bytes), header=None, sep=None, engine="python")
    return df.values.tolist()


def _resolver_nome_aba(nomes_disponiveis: list[str], preferido: str) -> str:
    """Tenta achar a aba 'PLANO' (ou variação de maiúsculas/espaços);
    se não encontrar, usa a primeira aba disponível."""
    for nome in nomes_disponiveis:
        if _normalizar(nome) == _normalizar(preferido):
            return nome
    for nome in nomes_disponiveis:
        if _normalizar(preferido) in _normalizar(nome):
            return nome
    return nomes_disponiveis[0]


def ler_plano(file_bytes: bytes, nome_arquivo: str, aba_preferida: str = "PLANO") -> list[dict]:
    """Lê o arquivo (detectando o formato pela extensão) e retorna a
    lista de itens da aba de plano como dicts com as chaves:
    sku, produto, familia, qtd, peso, pallets, lote."""
    ext = nome_arquivo.lower().rsplit(".", 1)[-1]

    if ext == "xlsb":
        matriz = _matriz_de_xlsb(file_bytes, aba_preferida)
    elif ext in ("xlsx", "xlsm"):
        matriz = _matriz_de_excel(file_bytes, aba_preferida)
    elif ext == "csv":
        matriz = _matriz_de_csv(file_bytes)
    else:
        raise ValueError(f"Formato de arquivo não suportado: .{ext}")

    header_idx, mapa = _localizar_header_e_mapa(matriz)

    itens = []
    for row in matriz[header_idx + 1:]:
        if mapa["sku"] >= len(row):
            continue
        sku_val = row[mapa["sku"]]
        if not isinstance(sku_val, (int, float)) or pd.isna(sku_val):
            continue

        def get(col_key, default=None):
            idx = mapa.get(col_key)
            if idx is None or idx >= len(row):
                return default
            val = row[idx]
            return default if (val is None or (isinstance(val, float) and pd.isna(val))) else val

        qtd = get("qtd", 0)
        peso = get("peso", 0)
        pallets = get("pallets", 0)
        if not qtd or not peso or not pallets:
            continue

        lote_val = get("lote", "")
        lote_str = str(lote_val).strip() if lote_val not in (None, "") else ""

        itens.append(dict(
            sku=int(sku_val),
            produto=str(get("produto", "")).strip(),
            familia=str(get("familia", "")).strip(),
            qtd=float(qtd),
            peso=float(peso),
            pallets=float(pallets),
            lote=lote_str,
        ))

    if not itens:
        raise ValueError(
            "Nenhum item válido foi encontrado na aba de plano. "
            "Verifique se o arquivo contém a tabela esperada."
        )

    return itens
