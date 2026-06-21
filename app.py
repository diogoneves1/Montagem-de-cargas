"""
Agente de Montagem de Cargas
=========================================
App Streamlit para substituir a macro VBA "MontarContainers".

Fluxo:
  1. Usuário faz upload do plano de embarque (.xlsx, .xlsb, .xlsm ou .csv).
  2. O app lê a aba PLANO (ou equivalente), identifica SKUs, pesos,
     pallets e lotes.
  3. Roda a lógica de distribuição (BASE/TOPO/PISO, cota de base, etc.)
  4. Mostra um resumo visual na tela.
  5. Gera um Excel para download, com abas RESUMO + uma por container +
     PLANO_MONTADO (pronta para colar de volta no plano original).

Para rodar localmente:
    streamlit run app.py
"""

import traceback

import pandas as pd
import streamlit as st

from distribuicao import PESO_MAX, MAX_POSICOES_PISO, distribuir_containers
from geracao_excel import gerar_excel
from leitura import ler_plano


st.set_page_config(
    page_title="Montagem de Cargas",
    page_icon="📦",
    layout="wide",
)

# ─────────────────────────────────────────────────────────────────
# Identidade visual da marca (mesma paleta da planilha Excel gerada):
#   azul marinho  #0D2C6B  — cabeçalhos e elementos primários
#   azul claro    #D6E4F7  — fundos secundários / cards
#   verde         #D9EAD3  — família SARDINHA
#   azul médio    #CFE2F3  — família ATUM
#   laranja       #FFF2CC  — lote / AZEITE
# ─────────────────────────────────────────────────────────────────
AZUL_MARINHO = "#0D2C6B"
AZUL_CLARO = "#D6E4F7"
VERDE_SARDINHA = "#D9EAD3"
AZUL_ATUM = "#CFE2F3"
LARANJA_LOTE = "#FFF2CC"
CINZA_TEXTO = "#1A1A1A"

st.markdown(f"""
<style>
    /* Tipografia e fundo geral */
    html, body, [class*="css"] {{
        font-family: 'Arial', 'Helvetica Neue', sans-serif;
    }}

    /* Cabeçalho principal estilo "banner" do Excel */
    .cdpb-header {{
        background-color: {AZUL_MARINHO};
        color: #FFFFFF;
        padding: 1.4rem 1.8rem;
        border-radius: 6px;
        margin-bottom: 0.6rem;
    }}
    .cdpb-header h1 {{
        color: #FFFFFF !important;
        margin: 0;
        font-size: 1.6rem;
        font-weight: 700;
    }}
    .cdpb-subheader {{
        background-color: {AZUL_CLARO};
        color: {AZUL_MARINHO};
        padding: 0.55rem 1.2rem;
        border-radius: 4px;
        font-size: 0.85rem;
        font-style: italic;
        margin-bottom: 1.2rem;
    }}

    /* Botões primários na cor da marca */
    div.stButton > button[kind="primary"],
    div.stDownloadButton > button[kind="primary"] {{
        background-color: {AZUL_MARINHO};
        border-color: {AZUL_MARINHO};
    }}
    div.stButton > button[kind="primary"]:hover,
    div.stDownloadButton > button[kind="primary"]:hover {{
        background-color: #163a8c;
        border-color: #163a8c;
    }}

    /* Métricas (cards de número) com leve destaque */
    div[data-testid="stMetric"] {{
        background-color: {AZUL_CLARO};
        padding: 0.8rem 1rem;
        border-radius: 6px;
        border-left: 4px solid {AZUL_MARINHO};
    }}
    div[data-testid="stMetricLabel"] {{
        color: {AZUL_MARINHO};
        font-weight: 600;
    }}

    /* Barra lateral com fundo levemente diferenciado */
    section[data-testid="stSidebar"] {{
        background-color: #F4F7FC;
    }}
    section[data-testid="stSidebar"] h2 {{
        color: {AZUL_MARINHO};
    }}

    /* Expander e divisores mais discretos */
    div[data-testid="stExpander"] summary {{
        font-weight: 600;
        color: {AZUL_MARINHO};
    }}

    /* Tabelas: cabeçalho com a cor da marca (aplica-se ao dataframe nativo) */
    [data-testid="stDataFrame"] thead tr th {{
        background-color: {AZUL_MARINHO} !important;
        color: #FFFFFF !important;
    }}
</style>
""", unsafe_allow_html=True)


def _resumo_dataframe(containers) -> pd.DataFrame:
    linhas = []
    for c in containers:
        pct = c.peso / PESO_MAX * 100
        linhas.append({
            "Container": f"CNTR {c.numero:02d}",
            "Peso (kg)": round(c.peso),
            "% peso": f"{pct:.1f}%",
            "Pallets totais": round(c.pallets_total, 2),
            "Posições de piso": f"{c.pPiso:.2f} / {MAX_POSICOES_PISO}",
            "SKUs": c.skus_unicos,
            "Lotes": ", ".join(c.lotes) or "-",
        })
    return pd.DataFrame(linhas)


def _detalhe_dataframe(container) -> pd.DataFrame:
    linhas = []
    for l in sorted(container.linhas, key=lambda x: (x.familia, x.sku)):
        linhas.append({
            "SKU": l.sku,
            "Produto": l.produto,
            "Família": l.familia,
            "Posição": l.posicao,
            "Caixas": round(l.qtd),
            "Peso (kg)": round(l.peso),
            "Pallets": round(l.pallets, 3),
            "Lote": l.lote or "-",
        })
    return pd.DataFrame(linhas)


# Cores por família/posição — mesma paleta usada na planilha Excel.
_COR_FAMILIA = {
    "SARDINHA": "#D9EAD3",
    "ATUM": "#CFE2F3",
    "AZEITE": "#FFF2CC",
}
_COR_POSICAO_TOPO = "#FCE4D6"
_COR_POSICAO_PISO = "#FFE5CC"
_COR_LOTE = "#FFE082"


def _estilizar_linha_detalhe(row: pd.Series) -> list[str]:
    """Pinta a linha conforme posição/família, espelhando o Excel."""
    if row.get("Lote", "-") != "-":
        cor = _COR_LOTE
    elif row.get("Posição") == "TOPO":
        cor = _COR_POSICAO_TOPO
    elif row.get("Posição") == "PISO":
        cor = _COR_POSICAO_PISO
    else:
        cor = _COR_FAMILIA.get(row.get("Família"), "#FFFFFF")
    return [f"background-color: {cor}"] * len(row)


def main():
    st.markdown(
        """
        <div class="cdpb-header">
            <h1>📦 Montagem de Cargas</h1>
        </div>
        <div class="cdpb-subheader">
            Sobe o plano de embarque, o app monta os containers respeitando
            peso máximo, posições de piso e a regra de empilhamento
            (frágil só empilha sobre base estável) — e devolve a planilha pronta.
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.header("⚙️ Parâmetros")
        st.metric("Peso máximo por container", f"{PESO_MAX:,.0f} kg".replace(",", "."))
        st.metric("Posições de piso por container", MAX_POSICOES_PISO)
        st.markdown(
            "**Regra de empilhamento:**\n"
            "- SKUs não-frágeis vão na BASE (ocupam posição de piso).\n"
            "- SKUs frágeis empilham (TOPO) sobre uma BASE livre, "
            "1 pallet de base sustenta 1 pallet de topo.\n"
            "- Frágil sem base disponível ocupa posição própria no "
            "piso (PISO), sem empilhar nem sustentar nada."
        )

    arquivo = st.file_uploader(
        "Envie o plano de embarque",
        type=["xlsx", "xlsb", "xlsm", "csv"],
        help="O app procura automaticamente a aba 'PLANO' (ou similar) "
             "dentro do arquivo.",
    )

    if arquivo is None:
        st.info("⬆️ Envie um arquivo para começar.")
        return

    try:
        with st.spinner("Lendo o plano de embarque..."):
            itens = ler_plano(arquivo.getvalue(), arquivo.name)
    except Exception as e:
        st.error(f"Não foi possível ler o arquivo: {e}")
        with st.expander("Detalhes técnicos"):
            st.code(traceback.format_exc())
        return

    st.success(f"✅ {len(itens)} SKUs lidos do plano de embarque.")

    with st.expander("Ver itens lidos (antes da montagem)"):
        st.dataframe(pd.DataFrame(itens), use_container_width=True)

    if st.button("🚀 Montar containers", type="primary"):
        try:
            with st.spinner("Montando containers..."):
                containers = distribuir_containers(itens)
        except Exception as e:
            st.error(f"Erro ao montar os containers: {e}")
            with st.expander("Detalhes técnicos"):
                st.code(traceback.format_exc())
            return

        st.session_state["containers"] = containers
        st.session_state["nome_arquivo"] = arquivo.name

    if "containers" not in st.session_state:
        return

    containers = st.session_state["containers"]

    st.divider()
    st.subheader(f"📊 Resultado: {len(containers)} containers")

    peso_total = sum(c.peso for c in containers)
    pallets_total = sum(c.pallets_total for c in containers)
    col1, col2, col3 = st.columns(3)
    col1.metric("Total de containers", len(containers))
    col2.metric("Peso total alocado", f"{peso_total:,.0f} kg".replace(",", "."))
    col3.metric("Pallets totais", f"{pallets_total:,.1f}".replace(",", "."))

    df_resumo = _resumo_dataframe(containers)
    st.dataframe(df_resumo, use_container_width=True, hide_index=True)

    # Alerta visual se algum container ficou muito abaixo do esperado
    abaixo = [c for c in containers if c.peso < PESO_MAX * 0.9 and c is not containers[-1]]
    if abaixo:
        st.warning(
            f"⚠️ {len(abaixo)} container(s) intermediário(s) ficaram com menos "
            f"de 90% do peso máximo. Isso pode indicar um desbalanceamento "
            f"entre BASE e FRÁGIL disponíveis no plano."
        )

    st.divider()
    st.subheader("🔍 Detalhe por container")
    cntr_selecionado = st.selectbox(
        "Selecione um container para ver o detalhamento",
        options=[c.numero for c in containers],
        format_func=lambda n: f"CNTR {n:02d}",
    )
    container_atual = next(c for c in containers if c.numero == cntr_selecionado)
    df_detalhe = _detalhe_dataframe(container_atual)
    st.dataframe(
        df_detalhe.style.apply(_estilizar_linha_detalhe, axis=1),
        use_container_width=True,
        hide_index=True,
    )

    if container_atual.lotes:
        st.info(f"📋 Lotes presentes neste container: **{', '.join(container_atual.lotes)}**")

    st.divider()
    st.subheader("⬇️ Download")

    with st.spinner("Gerando planilha..."):
        nome_base = st.session_state.get("nome_arquivo", "embarque")
        titulo = f"PLANO DE EMBARQUE — {nome_base}"
        excel_bytes = gerar_excel(containers, titulo=titulo)

    st.download_button(
        label="📥 Baixar planilha de montagem (.xlsx)",
        data=excel_bytes,
        file_name="Montagem_Containers.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )


if __name__ == "__main__":
    main()
