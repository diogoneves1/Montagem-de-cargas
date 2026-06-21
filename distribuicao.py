"""
Núcleo da lógica de montagem de containers.
Replica fielmente a lógica validada da macro VBA v9.1:
  - FASE 1: SKUs não-frágeis (BASE) ocupam posição de piso.
  - FASE 2: SKUs frágeis empilham (TOPO) sobre BASE livre (1-para-1) ou
            ocupam posição própria no piso (PISO) quando não há base.
  - FASE 1B: se não há mais frágil disponível, mas ainda sobra peso e
             espaço de piso, completa com mais BASE ignorando a cota
             (evita containers fechando incompletos).
  - Cota de BASE distribuída igualmente entre os containers estimados,
    liberada no último container.

Regras físicas:
  - PESO_MAX kg por container (pode ultrapassar levemente no último).
  - MAX_POSICOES_PISO posições de piso (BASE + frágil solto).
  - Cada posição de BASE sustenta no máximo 1 pallet de frágil empilhado.
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
# piso, então não há apoio estável para outro pallet por cima). Esses
# itens ocupam posição de piso normalmente, mas nunca contam para
# liberar espaço de empilhamento (pBase).
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

    def __post_init__(self):
        self.kgPal = self.rPeso / self.rPal if self.rPal else 0.0
        self.qtdPal = self.rQtd / self.rPal if self.rPal else 0.0
        self.eFragil = self.sku in SKUS_FRAGEIS
        # >>> Robustez: alguns SKUs no plano vêm com a coluna FAMÍLIA
        # vazia/zerada (erro de digitação na planilha de origem). Sem
        # uma família válida, o item ficava invisível para a lógica de
        # alocação (que busca por família dentro de FAM_ORDER) e sobrava
        # sem ser alocado. Normaliza para string e usa "OUTROS" como
        # fallback, garantindo que o item sempre seja processável.
        fam_str = str(self.familia).strip() if self.familia else ""
        if fam_str.upper() not in FAM_ORDER:
            fam_str = "OUTROS"
        self.familia = fam_str
        # >>> Item "sustenta" empilhamento se NÃO for frágil E a família
        # não estiver na lista de famílias sem sustentação (ex: AZEITE).
        self.eSustenta = (not self.eFragil) and (self.familia not in FAMILIAS_SEM_SUSTENTACAO)


@dataclass
class Linha:
    sku: int
    produto: str
    familia: str
    lote: str | None
    posicao: str  # "BASE", "TOPO", "PISO"
    qtd: float
    pallets: float
    peso: float


@dataclass
class Container:
    numero: int
    linhas: list[Linha] = field(default_factory=list)
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
    def lotes(self) -> list[str]:
        return sorted({l.lote for l in self.linhas if l.lote})


def _melhor_disponivel(items: list[Item], fragil: bool,
                        sustenta: bool | None = None) -> Item | None:
    """Retorna o item disponível (rPeso > 0.01) com maior kg/pallet,
    respeitando a ordem de prioridade de família (SARDINHA > ATUM >
    AZEITE > OUTROS). Se `sustenta` for informado, filtra também por
    essa propriedade (True = pode sustentar empilhamento, False = não)."""
    for fam in FAM_ORDER:
        candidatos = [
            it for it in items
            if it.eFragil == fragil and it.familia == fam and it.rPeso > 0.01
            and (sustenta is None or it.eSustenta == sustenta)
        ]
        if candidatos:
            return max(candidatos, key=lambda x: x.kgPal)
    return None


def _gravar(container: Container, item: Item, qtd: float, pal: float,
            peso: float, posicao: str) -> Linha:
    linha = Linha(
        sku=item.sku, produto=item.produto, familia=item.familia,
        lote=item.lote, posicao=posicao,
        qtd=round(qtd, 2), pallets=round(pal, 4), peso=round(peso, 3),
    )
    container.linhas.append(linha)
    container.peso += peso
    item.rPal -= pal
    item.rPeso -= peso
    item.rQtd -= qtd
    return linha


def _fase1_base(items: list[Item], c: Container, limite_base: float,
                 peso_max_este: float = PESO_MAX) -> None:
    """Preenche o piso com SKUs não-frágeis. Dois sub-tipos:
      - eSustenta=True (ex: SARDINHA, ATUM): conta em pBase, liberando
        espaço de empilhamento para frágeis em cima (TOPO).
      - eSustenta=False (ex: AZEITE): ocupa posição de piso normalmente,
        mas NÃO conta em pBase (pallet não oferece apoio estável para
        nada em cima dele).
    A cota de base (limite_base) se aplica apenas ao que sustenta."""
    did_add = True
    while did_add and c.peso < peso_max_este - 0.5 and c.pPiso < MAX_POSICOES_PISO - 1e-6:
        # Prioriza itens que sustentam enquanto a cota permitir; quando
        # a cota se esgota, ainda processa os que NÃO sustentam (AZEITE)
        # livremente, já que eles não competem pela cota de empilhamento.
        sustenta_disponivel = c.pBase < limite_base - 1e-6
        item = None
        if sustenta_disponivel:
            item = _melhor_disponivel(items, fragil=False, sustenta=True)
        if item is None:
            item = _melhor_disponivel(items, fragil=False, sustenta=False)
        if item is None:
            break
        did_add = False

        sp = peso_max_este - c.peso
        if sp < 0.5:
            break
        pc = sp / item.kgPal if item.kgPal > 0 else item.rPal
        pa = min(pc, item.rPal)

        esp_piso = MAX_POSICOES_PISO - c.pPiso
        pa = min(pa, esp_piso)
        if item.eSustenta:
            esp_cota = limite_base - c.pBase
            pa = min(pa, esp_cota)
        if pa < 1e-6:
            break

        pe = pa * item.kgPal
        qa = pa * item.qtdPal
        if c.peso + pe > peso_max_este + 0.01:
            pa = (peso_max_este - c.peso) / item.kgPal
            pe = pa * item.kgPal
            qa = pa * item.qtdPal
        if pe < 0.01:
            continue

        posicao_gravar = "BASE" if item.eSustenta else "BASE (sem topo)"
        _gravar(c, item, qa, pa, pe, posicao_gravar)
        if item.eSustenta:
            c.pBase += pa
        c.pPiso += pa
        did_add = True


def _fase2_fragil(items: list[Item], c: Container, peso_max_este: float = PESO_MAX) -> None:
    """Distribui SKUs frágeis entre TOPO (empilhado sobre base livre)
    e PISO (posição própria), respeitando os dois limites físicos."""
    while True:
        if not any(it.eFragil and it.rPeso > 0.01 for it in items):
            break
        if c.peso >= peso_max_este - 0.5:
            break

        esp_emp = max(0.0, c.pBase - c.pTopo)
        esp_piso = max(0.0, MAX_POSICOES_PISO - c.pPiso)
        if esp_emp < 1e-6 and esp_piso < 1e-6:
            break

        item = _melhor_disponivel(items, fragil=True)
        if item is None:
            break

        sp = peso_max_este - c.peso
        if sp < 0.5:
            break
        pc = sp / item.kgPal if item.kgPal > 0 else item.rPal
        pa = min(pc, item.rPal)

        capac_total = esp_emp + esp_piso
        if capac_total >= pa - 1e-6:
            pa_emp = min(esp_emp, pa)
            pa_piso = pa - pa_emp
            pa_piso = min(pa_piso, esp_piso)
        else:
            pa_emp = esp_emp
            pa_piso = esp_piso
            pa = pa_emp + pa_piso

        if pa < 1e-6:
            break

        if pa_emp > 1e-6:
            pe_emp = pa_emp * item.kgPal
            qa_emp = pa_emp * item.qtdPal
            if c.peso + pe_emp > peso_max_este + 0.01:
                pa_emp = (peso_max_este - c.peso) / item.kgPal
                pe_emp = pa_emp * item.kgPal
                qa_emp = pa_emp * item.qtdPal
            if pe_emp >= 0.01:
                _gravar(c, item, qa_emp, pa_emp, pe_emp, "TOPO")
                c.pTopo += pa_emp
            else:
                pa_emp = 0.0

        if pa_piso > 1e-6 and c.peso < peso_max_este - 0.5:
            pe_piso = pa_piso * item.kgPal
            qa_piso = pa_piso * item.qtdPal
            if c.peso + pe_piso > peso_max_este + 0.01:
                pa_piso = (peso_max_este - c.peso) / item.kgPal
                pe_piso = pa_piso * item.kgPal
                qa_piso = pa_piso * item.qtdPal
            if pe_piso >= 0.01:
                _gravar(c, item, qa_piso, pa_piso, pe_piso, "PISO")
                c.pPiso += pa_piso
            else:
                pa_piso = 0.0

        if pa_emp < 1e-6 and pa_piso < 1e-6:
            break


def distribuir_containers(items_raw: list[dict]) -> list[Container]:
    """Recebe a lista de itens do PLANO (dicts) e retorna a lista de
    containers já montados, seguindo a lógica v9.1 validada."""
    items = [
        Item(
            sku=int(it["sku"]), produto=it["produto"], familia=it["familia"],
            lote=it.get("lote") or None,
            rQtd=float(it["qtd"]), rPeso=float(it["peso"]), rPal=float(it["pallets"]),
        )
        for it in items_raw
        if it["qtd"] and it["peso"] and it["pallets"]
    ]

    # >>> pal_base_total considera apenas itens que SUSTENTAM empilhamento
    # (ex: SARDINHA, ATUM). Itens não-frágeis sem sustentação (AZEITE)
    # ocupam piso mas não entram nessa cota, pois não competem pelo
    # espaço de empilhamento de TOPO.
    pal_base_total = sum(it.rPal for it in items if it.eSustenta)
    pal_piso_nao_sustenta = sum(it.rPal for it in items if not it.eFragil and not it.eSustenta)
    peso_total = sum(it.rPeso for it in items)

    # >>> Tolerância: o peso total pode passar de um múltiplo exato de
    # PESO_MAX por uma fração mínima (ex: 8,0007 containers), o que
    # forçaria ceil() a arredondar para 9 mesmo sendo essencialmente 8.
    # Uma tolerância de até 50kg (~0.2% do limite) evita esse efeito.
    TOLERANCIA_KG = 50.0
    n_cont_peso = math.ceil((peso_total - TOLERANCIA_KG) / PESO_MAX) if peso_total > 0 else 1
    n_cont_peso = max(n_cont_peso, 1)

    # >>> Para o piso, considera TODAS as posições ocupadas (base que
    # sustenta + base que não sustenta, ex: AZEITE), já que ambas
    # disputam o mesmo limite de MAX_POSICOES_PISO por container.
    pal_piso_total = pal_base_total + pal_piso_nao_sustenta
    n_cont_piso = math.ceil(pal_piso_total / MAX_POSICOES_PISO) if pal_piso_total > 0 else 1
    n_cont_estim = max(n_cont_peso, n_cont_piso, 1)

    cota_base = pal_base_total / n_cont_estim if pal_base_total > 0.01 else 0.0

    containers: list[Container] = []
    c_num = 0

    while any(it.rPeso > 0.01 for it in items):
        c_num += 1
        c = Container(numero=c_num)

        limite_base = cota_base if c_num < n_cont_estim else float("inf")
        # >>> No último container, libera uma margem extra de peso para
        # absorver qualquer resíduo da tolerância usada na estimativa
        # de n_cont_estim (evita sobrar peso sem container).
        peso_max_este = PESO_MAX if c_num < n_cont_estim else PESO_MAX + TOLERANCIA_KG + 50

        # Roda Fase1 -> Fase2 -> Fase1B (sem cota) em rodadas até estabilizar
        for _ in range(20):
            peso_antes = c.peso
            _fase1_base(items, c, limite_base, peso_max_este)
            _fase2_fragil(items, c, peso_max_este)

            ainda_tem_fragil = any(it.eFragil and it.rPeso > 0.01 for it in items)
            if (not ainda_tem_fragil and c.peso < peso_max_este - 0.5
                    and c.pPiso < MAX_POSICOES_PISO - 1e-6):
                _fase1_base(items, c, float("inf"), peso_max_este)

            if abs(c.peso - peso_antes) < 1e-3:
                break

        # Segurança: se nada foi alocado neste container (não deveria
        # acontecer), evita loop infinito.
        if not c.linhas:
            break

        containers.append(c)
        if c_num > 100:
            break

    return containers
