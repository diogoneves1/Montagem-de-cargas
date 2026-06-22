"""
Núcleo da lógica de montagem de containers.

ESTRATÉGIA (v10 — fatiamento proporcional global):
Em vez de processar container por container sequencialmente (greedy),
o que causava desbalanceamento quando categorias com kg/pallet muito
diferentes competiam pelo mesmo espaço (ex: sardinha densa monopolizando
o peso e empurrando ATUM/AZEITE/frágil para os últimos containers),
esta versão:

  1. Calcula N = número de containers necessários (pelo peso total).
  2. Para cada SKU, fationa sua quantidade total em N partes iguais
     (proporcional), preservando peso/pallets/caixas corretamente.
  3. Distribui essas fatias entre os N containers, garantindo que cada
     container receba uma "mistura representativa" de TODAS as
     categorias (SARDINHA, ATUM, AZEITE, frágil) desde o início — não
     apenas as mais densas primeiro.
  4. Dentro de cada container, aplica as regras físicas de POSIÇÃO
     (BASE / TOPO empilhado / PISO) apenas para organizar onde cada
     pallet fica, respeitando:
       - 1 pallet de BASE sustenta no máximo 1 pallet de frágil (TOPO).
       - AZEITE nunca sustenta nada em cima (BASE sem topo).
       - POUCH precisa dividir em 2 ao empilhar (consome 2 posições
         de base por pallet físico).
       - Máximo MAX_POSICOES_PISO posições de piso por container.
  5. Sobras que não couberam por limite de piso são redistribuídas
     entre containers vizinhos com espaço de piso disponível.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field


PESO_MAX = 27800
MAX_POSICOES_PISO = 21

# SKUs cujo produto é frágil (não pode receber peso por cima).
SKUS_FRAGEIS = {
    1165, 1166, 1167, 1168, 1178, 1184, 1185, 1186, 1205, 1206,
    1207, 1213, 1400, 1401, 1402, 1403, 1501, 1502, 1509, 1510,
    2201, 2202, 2203, 2205, 2206, 3601, 3602, 3603, 3623, 3630,
    51172, 51173, 51174, 51175, 51177, 51184, 51201, 51400,
    82201, 82203,
}

# >>> Famílias cujo pallet, mesmo não sendo frágil, NÃO sustenta nada
# empilhado em cima (ex: AZEITE — o pallet não ocupa toda a área do
# piso, então não há apoio estável para outro pallet por cima).
FAMILIAS_SEM_SUSTENTACAO = {"AZEITE"}

FAM_ORDER = ["SARDINHA", "ATUM", "AZEITE", "OUTROS"]


@dataclass
class Item:
    sku: int
    produto: str
    familia: str
    lote: str | None
    rQtd: float
    rPeso: float
    rPal: float
    kgPal: float = field(init=False)
    qtdPal: float = field(init=False)
    eFragil: bool = field(init=False)
    eSustenta: bool = field(init=False)
    eDivideAoEmpilhar: bool = field(init=False)

    def __post_init__(self):
        self.kgPal = self.rPeso / self.rPal if self.rPal else 0.0
        self.qtdPal = self.rQtd / self.rPal if self.rPal else 0.0
        self.eFragil = self.sku in SKUS_FRAGEIS
        fam_str = str(self.familia).strip() if self.familia else ""
        if fam_str.upper() not in FAM_ORDER:
            fam_str = "OUTROS"
        self.familia = fam_str
        self.eSustenta = (not self.eFragil) and (self.familia not in FAMILIAS_SEM_SUSTENTACAO)
        self.eDivideAoEmpilhar = "POUCH" in str(self.produto).upper()


@dataclass
class Linha:
    sku: int
    produto: str
    familia: str
    lote: str | None
    posicao: str
    qtd: float
    pallets: float
    peso: float


@dataclass
class Container:
    numero: int
    linhas: list = field(default_factory=list)
    peso: float = 0.0
    pPiso: float = 0.0
    pTopo: float = 0.0
    pBase: float = 0.0

    @property
    def pallets_total(self) -> float:
        return sum(l.pallets for l in self.linhas)

    @property
    def skus_unicos(self) -> int:
        return len(set(l.sku for l in self.linhas))

    @property
    def lotes(self) -> list:
        return sorted({l.lote for l in self.linhas if l.lote})


def _fatiar_item_em_n(item, n, peso_medio_container):
    """Divide a quantidade/peso/pallets de um item em fatias.

    Para reduzir fracionamento desnecessário: se o peso TOTAL do item
    for pequeno em relação ao peso médio de um container (menos de
    ~15%), ele é colocado em poucas fatias "concentradas" (1 a 3
    containers) em vez de espalhado igualmente em todos os N — isso
    mantém SKUs pequenos o mais inteiros possível. Itens grandes (ex:
    sardinha, pouches volumosos) continuam sendo fatiados em todos os
    N containers, pois são justamente os que precisam ser distribuídos
    para equilibrar o peso total."""
    if n <= 0:
        return []

    limiar_pequeno = peso_medio_container * 0.15
    if item.rPeso < limiar_pequeno and item.rPeso > 0:
        # Quantos containers "concentrados" bastam para acomodar este
        # item sem ficar fracionado em frações irrelevantes.
        n_concentrado = max(1, min(n, math.ceil(item.rPeso / max(limiar_pequeno, 1e-6))))
        fatias = []
        for i in range(n):
            if i < n_concentrado:
                fatias.append(dict(
                    qtd=item.rQtd / n_concentrado,
                    peso=item.rPeso / n_concentrado,
                    pallets=item.rPal / n_concentrado,
                ))
            else:
                fatias.append(dict(qtd=0.0, peso=0.0, pallets=0.0))
        return fatias

    # Item grande: fatiamento proporcional normal em todos os N.
    fatias = []
    for i in range(n):
        fatias.append(dict(
            qtd=item.rQtd / n,
            peso=item.rPeso / n,
            pallets=item.rPal / n,
        ))
    return fatias


def _construir_fatias_globais(items, n):
    peso_total = sum(it.rPeso for it in items)
    peso_medio_container = peso_total / n if n > 0 else PESO_MAX
    fatias_por_container = [[] for _ in range(n)]
    peso_acumulado = [0.0] * n  # rastreia peso já reservado por container

    items_ordenados = sorted(items, key=lambda it: -it.rPeso)

    for item in items_ordenados:
        fatias = _fatiar_item_em_n(item, n, peso_medio_container)
        indices_nao_nulos = [i for i, f in enumerate(fatias) if f["pallets"] > 1e-9]

        if 0 < len(indices_nao_nulos) < n:
            n_concentrado = len(indices_nao_nulos)
            valores_nao_nulos = [fatias[i] for i in indices_nao_nulos]
            fatias = [dict(qtd=0.0, peso=0.0, pallets=0.0) for _ in range(n)]

            # >>> Escolhe os N containers com MENOS peso acumulado até
            # agora (não rotação cega), garantindo que itens pequenos
            # vão para onde realmente cabe sem ultrapassar PESO_MAX.
            candidatos_ordenados = sorted(range(n), key=lambda i: peso_acumulado[i])
            destinos = candidatos_ordenados[:n_concentrado]
            for k, valor in enumerate(valores_nao_nulos):
                idx_destino = destinos[k]
                fatias[idx_destino] = valor

        for c_idx, fatia in enumerate(fatias):
            if fatia["pallets"] > 1e-9:
                fatias_por_container[c_idx].append((item, fatia))
                peso_acumulado[c_idx] += fatia["peso"]

    return fatias_por_container


def _organizar_posicoes_container(c_numero, fatias):
    """Decide a POSIÇÃO (BASE/TOPO/PISO) de cada fatia já decidida
    para este container, respeitando os limites físicos."""
    c = Container(numero=c_numero)

    base_sustenta = sorted(
        [(it, f) for it, f in fatias if it.eSustenta],
        key=lambda x: -x[0].kgPal,
    )
    base_sem_sustentar = [(it, f) for it, f in fatias if (not it.eFragil) and not it.eSustenta]
    fragil = sorted(
        [(it, f) for it, f in fatias if it.eFragil],
        key=lambda x: -x[0].kgPal,
    )

    for item, f in base_sustenta:
        if f["pallets"] <= 1e-9:
            continue
        pa = f["pallets"]
        esp_piso = MAX_POSICOES_PISO - c.pPiso
        pa_usar = min(pa, max(0.0, esp_piso))
        if pa_usar <= 1e-9:
            continue
        proporcao = pa_usar / pa
        c.linhas.append(Linha(
            sku=item.sku, produto=item.produto, familia=item.familia,
            lote=item.lote, posicao="BASE",
            qtd=round(f["qtd"] * proporcao, 2),
            pallets=round(pa_usar, 4),
            peso=round(f["peso"] * proporcao, 3),
        ))
        c.peso += f["peso"] * proporcao
        c.pBase += pa_usar
        c.pPiso += pa_usar

    for item, f in base_sem_sustentar:
        if f["pallets"] <= 1e-9:
            continue
        pa = f["pallets"]
        esp_piso = MAX_POSICOES_PISO - c.pPiso
        pa_usar = min(pa, max(0.0, esp_piso))
        if pa_usar <= 1e-9:
            continue
        proporcao = pa_usar / pa
        c.linhas.append(Linha(
            sku=item.sku, produto=item.produto, familia=item.familia,
            lote=item.lote, posicao="BASE (sem topo)",
            qtd=round(f["qtd"] * proporcao, 2),
            pallets=round(pa_usar, 4),
            peso=round(f["peso"] * proporcao, 3),
        ))
        c.peso += f["peso"] * proporcao
        c.pPiso += pa_usar

    for item, f in fragil:
        if f["pallets"] <= 1e-9:
            continue
        pa_total = f["pallets"]
        fator = 2.0 if item.eDivideAoEmpilhar else 1.0

        esp_emp = max(0.0, c.pBase - c.pTopo)
        max_pa_emp = esp_emp / fator if fator > 0 else 0.0
        esp_piso = max(0.0, MAX_POSICOES_PISO - c.pPiso)

        pa_emp = min(pa_total, max_pa_emp)
        pa_piso_possivel = pa_total - pa_emp
        pa_piso = min(pa_piso_possivel, esp_piso)

        pa_alocado = pa_emp + pa_piso
        if pa_alocado <= 1e-9:
            continue

        if pa_emp > 1e-9:
            posicao = "TOPO (dividido em 2)" if item.eDivideAoEmpilhar else "TOPO"
            proporcao_emp = pa_emp / pa_total
            c.linhas.append(Linha(
                sku=item.sku, produto=item.produto, familia=item.familia,
                lote=item.lote, posicao=posicao,
                qtd=round(f["qtd"] * proporcao_emp, 2),
                pallets=round(pa_emp, 4),
                peso=round(f["peso"] * proporcao_emp, 3),
            ))
            c.peso += f["peso"] * proporcao_emp
            c.pTopo += pa_emp * fator

        if pa_piso > 1e-9:
            proporcao_piso = pa_piso / pa_total
            c.linhas.append(Linha(
                sku=item.sku, produto=item.produto, familia=item.familia,
                lote=item.lote, posicao="PISO",
                qtd=round(f["qtd"] * proporcao_piso, 2),
                pallets=round(pa_piso, 4),
                peso=round(f["peso"] * proporcao_piso, 3),
            ))
            c.peso += f["peso"] * proporcao_piso
            c.pPiso += pa_piso

    return c


def _calcular_sobras(fatias_por_container, containers):
    sobras = []
    for c_idx, fatias in enumerate(fatias_por_container):
        c = containers[c_idx]
        alocado_por_sku = {}
        for l in c.linhas:
            alocado_por_sku[l.sku] = alocado_por_sku.get(l.sku, 0.0) + l.pallets

        for item, f in fatias:
            alocado = alocado_por_sku.get(item.sku, 0.0)
            if alocado >= f["pallets"] - 1e-9:
                alocado_por_sku[item.sku] = alocado - f["pallets"]
                continue
            residual_pal = f["pallets"] - max(0.0, alocado)
            if residual_pal > 1e-9:
                proporcao = residual_pal / f["pallets"] if f["pallets"] > 0 else 0
                sobras.append((item, {
                    "qtd": f["qtd"] * proporcao,
                    "peso": f["peso"] * proporcao,
                    "pallets": residual_pal,
                }, c_idx))
            alocado_por_sku[item.sku] = 0.0
    return sobras


def _tentar_distribuir_com_n(items, n):
    items_copia = [
        Item(sku=it.sku, produto=it.produto, familia=it.familia, lote=it.lote,
             rQtd=it.rQtd, rPeso=it.rPeso, rPal=it.rPal)
        for it in items
    ]

    fatias_por_container = _construir_fatias_globais(items_copia, n)
    containers = [
        _organizar_posicoes_container(i + 1, fatias_por_container[i])
        for i in range(n)
    ]

    for _ in range(5):
        sobras = _calcular_sobras(fatias_por_container, containers)
        if not sobras:
            break

        progresso = False
        for item, fatia_residual, origem_idx in sobras:
            candidatos = sorted(
                range(len(containers)),
                key=lambda i: -(MAX_POSICOES_PISO - containers[i].pPiso),
            )
            for c_idx in candidatos:
                if c_idx == origem_idx:
                    continue
                c = containers[c_idx]
                esp_piso = MAX_POSICOES_PISO - c.pPiso
                if esp_piso <= 1e-9:
                    continue
                extra = _organizar_posicoes_container(c.numero, [(item, fatia_residual)])
                if not extra.linhas:
                    continue
                for l in extra.linhas:
                    c.linhas.append(l)
                c.peso += extra.peso
                c.pPiso += extra.pPiso
                c.pBase += extra.pBase
                c.pTopo += extra.pTopo
                progresso = True
                break

        if not progresso:
            break

        # Recalcula fatias_por_container a partir do estado consolidado,
        # para a próxima rodada de _calcular_sobras comparar correto.
        novas_fatias = [[] for _ in range(n)]
        for c_idx, c in enumerate(containers):
            agregado = {}
            for l in c.linhas:
                key = l.sku
                if key not in agregado:
                    item_ref = next(it for it in items_copia if it.sku == l.sku)
                    agregado[key] = [item_ref, dict(qtd=0.0, peso=0.0, pallets=0.0)]
                agregado[key][1]["qtd"] += l.qtd
                agregado[key][1]["peso"] += l.peso
                agregado[key][1]["pallets"] += l.pallets
            novas_fatias[c_idx] = [(v[0], v[1]) for v in agregado.values()]
        fatias_por_container = novas_fatias

    peso_alocado = sum(c.peso for c in containers)
    peso_original = sum(it.rPeso for it in items_copia)
    sobra_total = max(0.0, peso_original - peso_alocado)
    return containers, sobra_total


def distribuir_containers(items_raw):
    """Recebe a lista de itens do PLANO (dicts) e retorna a lista de
    containers já montados via fatiamento proporcional global."""
    items = [
        Item(
            sku=int(it["sku"]), produto=it["produto"], familia=it["familia"],
            lote=it.get("lote") or None,
            rQtd=float(it["qtd"]), rPeso=float(it["peso"]), rPal=float(it["pallets"]),
        )
        for it in items_raw
        if it["qtd"] and it["peso"] and it["pallets"]
    ]
    if not items:
        return []

    peso_total = sum(it.rPeso for it in items)
    n_min_peso = max(1, math.ceil((peso_total - 50.0) / PESO_MAX))

    melhor_resultado = None
    for n in range(n_min_peso, n_min_peso + 10):
        containers, sobra_total = _tentar_distribuir_com_n(items, n)
        if sobra_total < 1.0:
            melhor_resultado = containers
            break

    if melhor_resultado is None:
        melhor_resultado, _ = _tentar_distribuir_com_n(items, n_min_peso + 10)

    # >>> Remove linhas com quantidade praticamente nula (resíduo de
    # arredondamento de ponto flutuante das fatias proporcionais, ex:
    # 0.0003 pallets) e containers que ficaram totalmente vazios.
    for c in melhor_resultado:
        c.linhas = [l for l in c.linhas if l.pallets > 0.005 and l.peso > 0.5]

    return [c for c in melhor_resultado if c.linhas]
