"""
Painel — Trabalho de cuidado não remunerado
Fonte única: IBGE/PNAD Contínua, Tabela 1.1.1.
"""

import io
import math
import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


BASE = Path(__file__).parent
ARQUIVO_PADRAO = BASE / "Tabela 1.1.1.xls"

SIGLAS = {
    "Rondônia": "RO", "Acre": "AC", "Amazonas": "AM", "Roraima": "RR",
    "Pará": "PA", "Amapá": "AP", "Tocantins": "TO", "Maranhão": "MA",
    "Piauí": "PI", "Ceará": "CE", "Rio Grande do Norte": "RN",
    "Paraíba": "PB", "Pernambuco": "PE", "Alagoas": "AL", "Sergipe": "SE",
    "Bahia": "BA", "Minas Gerais": "MG", "Espírito Santo": "ES",
    "Rio de Janeiro": "RJ", "São Paulo": "SP", "Paraná": "PR",
    "Santa Catarina": "SC", "Rio Grande do Sul": "RS",
    "Mato Grosso do Sul": "MS", "Mato Grosso": "MT", "Goiás": "GO",
    "Distrito Federal": "DF",
}
LOCALIDADES = {"Brasil", "Norte", "Nordeste", "Sudeste", "Sul", "Centro-Oeste", *SIGLAS}
COLUNAS = {
    4: ("Homens", "Branca"),
    5: ("Homens", "Preta ou parda"),
    6: ("Mulheres", "Branca"),
    7: ("Mulheres", "Preta ou parda"),
}
GRUPOS = {
    ("Homens", "Branca"): "Homens brancos",
    ("Homens", "Preta ou parda"): "Homens pretos ou pardos",
    ("Mulheres", "Branca"): "Mulheres brancas",
    ("Mulheres", "Preta ou parda"): "Mulheres pretas ou pardas",
}
CORES_SEXO = {"Homens": "#2878B5", "Mulheres": "#E87535"}
CORES_GRUPO = {
    "Homens brancos": "#2878B5",
    "Homens pretos ou pardos": "#55A6D9",
    "Mulheres brancas": "#E87535",
    "Mulheres pretas ou pardas": "#B84D27",
}
FUNDO = "rgba(0,0,0,0)"
TEXTO = "#30343B"
GRADE = "#E7E9EC"


class FormatoInvalido(ValueError):
    pass


def sigla(localidade):
    return SIGLAS.get(localidade, localidade)


def formatar(valor, casas=1):
    if pd.isna(valor):
        return "—"
    return f"{valor:.{casas}f}".replace(".", ",")


@st.cache_data
def ler_tabela(arquivo_bytes):
    """Lê somente as abas anuais da Tabela 1.1.1 e seus respectivos CVs."""
    buffer = io.BytesIO(arquivo_bytes)
    planilha = pd.ExcelFile(buffer)
    medidas = {}

    for aba in planilha.sheet_names:
        match = re.fullmatch(r"(\d{4})( \(CV\))?", str(aba).strip())
        if not match:
            continue

        ano = int(match.group(1))
        tipo = "cv" if match.group(2) else "horas"
        bruto = pd.read_excel(planilha, sheet_name=aba, header=None)

        if bruto.shape[1] < 8:
            raise FormatoInvalido(
                f"A aba {aba!r} não possui as oito colunas esperadas."
            )

        for _, linha in bruto[bruto[0].astype(str).str.strip().isin(LOCALIDADES)].iterrows():
            localidade = str(linha[0]).strip()
            for coluna, (sexo, raca) in COLUNAS.items():
                chave = (localidade, ano, sexo, raca)
                medidas.setdefault(chave, {})[tipo] = pd.to_numeric(
                    linha.iloc[coluna], errors="coerce"
                )

    if not medidas:
        raise FormatoInvalido(
            "Não encontrei abas anuais no padrão da Tabela 1.1.1."
        )

    dados = pd.DataFrame([
        {
            "localidade": chave[0],
            "ano": chave[1],
            "sexo": chave[2],
            "raca": chave[3],
            **valores,
        }
        for chave, valores in medidas.items()
    ])

    if "cv" not in dados:
        dados["cv"] = float("nan")

    dados["grupo"] = [
        GRUPOS[(sexo, raca)] for sexo, raca in zip(dados["sexo"], dados["raca"])
    ]
    return dados.dropna(subset=["horas"]).sort_values(
        ["ano", "localidade", "sexo", "raca"]
    )


def por_sexo(dados):
    quadro = dados.pivot_table(
        index="localidade", columns="sexo", values="horas", aggfunc="first"
    )
    quadro = quadro.reindex(columns=["Homens", "Mulheres"])
    quadro["Diferença"] = quadro["Mulheres"] - quadro["Homens"]
    return quadro


def por_grupo(dados):
    return dados.pivot_table(
        index="localidade", columns="grupo", values="horas", aggfunc="first"
    ).reindex(columns=list(GRUPOS.values()))


def estilo(fig, altura=420):
    fig.update_layout(
        height=altura,
        separators=",.",
        margin=dict(l=55, r=25, t=65, b=55),
        font=dict(family="Arial", size=12, color=TEXTO),
        paper_bgcolor=FUNDO,
        plot_bgcolor=FUNDO,
        legend=dict(
            orientation="h", y=1.08, yanchor="bottom",
            x=0, title_text=""
        ),
        hoverlabel=dict(bgcolor="white", font_color=TEXTO),
    )
    fig.update_xaxes(title=None, showgrid=False, linecolor="#B9BEC5", automargin=True)
    fig.update_yaxes(
        title=None, gridcolor=GRADE, zeroline=False, automargin=True
    )
    return fig


def mostrar(fig, chave):
    st.plotly_chart(
        fig, use_container_width=True, theme=None, key=chave,
        config={"displayModeBar": True, "displaylogo": False}
    )


st.set_page_config(
    page_title="Trabalho de cuidado não remunerado",
    page_icon="📊",
    layout="wide"
)

st.title("Trabalho de cuidado não remunerado")
st.markdown(
    "Média de horas semanais dedicadas a cuidados de pessoas e/ou afazeres "
    "domésticos pelas pessoas de 14 anos ou mais que realizaram essas atividades. "
    "Fonte: IBGE, PNAD Contínua, Tabela 1.1.1."
)

# Entrada dos dados: arquivo padrão ou upload de uma nova versão da mesma tabela.
upload = st.sidebar.file_uploader(
    "Carregar Tabela 1.1.1 (.xls ou .xlsx)",
    type=["xls", "xlsx"],
    help="O arquivo deve conter abas por ano e abas correspondentes com sufixo (CV)."
)

if upload is not None:
    conteudo = upload.getvalue()
elif ARQUIVO_PADRAO.exists():
    conteudo = ARQUIVO_PADRAO.read_bytes()
else:
    st.warning("Envie o arquivo da Tabela 1.1.1 para iniciar o painel.")
    st.stop()

try:
    dados = ler_tabela(conteudo)
except Exception as erro:
    st.error(f"Não foi possível ler a Tabela 1.1.1: {erro}")
    st.stop()

anos = sorted(dados["ano"].unique(), reverse=True)
ufs_disponiveis = sorted(
    [loc for loc in dados["localidade"].unique() if loc in SIGLAS],
    key=sigla
)
ano = st.sidebar.selectbox("Ano", anos, index=0)
ufs_padrao = [u for u in ["Bahia", "Pernambuco", "Amazonas", "Ceará"] if u in ufs_disponiveis]
ufs = st.sidebar.multiselect(
    "Estados para comparação",
    options=ufs_disponiveis,
    default=ufs_padrao or ufs_disponiveis[:4],
    format_func=lambda x: f"{sigla(x)} — {x}"
)

atual = dados[dados["ano"] == ano]
sexo = por_sexo(atual)
grupo = por_grupo(atual)

selecionados = [u for u in ufs if u in sexo.index]
sexo_sel = sexo.reindex(selecionados)
grupo_sel = grupo.reindex(selecionados)

if not sexo_sel.empty:
    media_mulheres = sexo_sel["Mulheres"].mean()
    media_homens = sexo_sel["Homens"].mean()
    media_diferenca = sexo_sel["Diferença"].mean()
    k1, k2, k3 = st.columns(3)
    k1.metric("Mulheres — h/semana", formatar(media_mulheres))
    k2.metric("Homens — h/semana", formatar(media_homens))
    k3.metric("Diferença média — h/semana", formatar(media_diferenca))
    st.caption(
        f"Média simples dos {len(sexo_sel)} estados selecionados; "
        "não ponderada pela população."
    )

tab_comparacao, tab_raca, tab_evolucao, tab_dados = st.tabs(
    ["Sexo", "Sexo e raça", "Evolução", "Dados"]
)

with tab_comparacao:
    st.subheader(f"Horas semanais dedicadas ao cuidado por sexo — {ano}")
    if sexo_sel.empty:
        st.info("Selecione ao menos um estado.")
    else:
        ordem = sexo_sel.sort_values("Diferença", ascending=False)
        comparacao = pd.concat([sexo.reindex(["Brasil"]), ordem])
        comparacao = comparacao.dropna(how="all")
        comparacao.index.name = "Localidade"
        longo = (
            comparacao[["Homens", "Mulheres"]]
            .rename(index=sigla)
            .reset_index()
            .melt(
                id_vars="Localidade", var_name="Sexo",
                value_name="Horas por semana"
            )
        )
        fig = px.bar(
            longo, x="Localidade", y="Horas por semana", color="Sexo",
            barmode="group", text_auto=".1f",
            color_discrete_map=CORES_SEXO,
            category_orders={"Sexo": ["Homens", "Mulheres"]}
        )
        fig.update_traces(
            textposition="outside", cliponaxis=False,
            hovertemplate="%{x} · %{fullData.name}<br>%{y:.1f} h/semana<extra></extra>"
        )
        fig.update_layout(bargroupgap=0.12)
        mostrar(estilo(fig, max(420, 34 * len(comparacao))), "grafico_sexo")

        st.subheader("Diferença entre mulheres e homens")
        diferencas = comparacao["Diferença"].rename(index=sigla).sort_values()
        fig = go.Figure(go.Bar(
            x=diferencas.values,
            y=diferencas.index,
            orientation="h",
            marker_color=["#A7ADB4" if x == "Brasil" else "#E87535"
                          for x in diferencas.index],
            text=[f"{v:.1f} h" for v in diferencas.values],
            textposition="outside",
            cliponaxis=False,
            hovertemplate="%{y}<br>%{x:.1f} h a mais para mulheres<extra></extra>"
        ))
        fig.update_layout(showlegend=False, margin=dict(l=65, r=75, t=35, b=45))
        mostrar(estilo(fig, max(360, 30 * len(diferencas))), "grafico_diferenca")

with tab_raca:
    st.subheader(f"Horas semanais por sexo e cor ou raça — {ano}")
    st.caption(
        "A tabela apresenta os grupos de homens e mulheres brancos e pretos ou pardos. "
        "O recorte por raça não é combinado com outras características nesta visualização."
    )
    if grupo_sel.empty:
        st.info("Selecione ao menos um estado.")
    else:
        comparacao_raca = pd.concat([grupo.reindex(["Brasil"]), grupo_sel])
        comparacao_raca = comparacao_raca.dropna(how="all")
        longo = (
            comparacao_raca.rename(index=sigla)
            .reset_index(names="Localidade")
            .melt(
                id_vars="Localidade", value_vars=list(GRUPOS.values()),
                var_name="Grupo", value_name="Horas por semana"
            )
        )
        fig = px.bar(
            longo, x="Localidade", y="Horas por semana", color="Grupo",
            barmode="group", color_discrete_map=CORES_GRUPO
        )
        fig.update_traces(
            hovertemplate="%{x} · %{fullData.name}<br>%{y:.1f} h/semana<extra></extra>"
        )
        fig.update_layout(bargroupgap=0.08)
        mostrar(estilo(fig, max(440, 38 * len(comparacao_raca))), "grafico_raca")

        st.subheader("Diferenças entre grupos")
        contrastes = pd.DataFrame(index=grupo_sel.index)
        contrastes["Mulheres − homens brancos"] = (
            grupo_sel["Mulheres brancas"] - grupo_sel["Homens brancos"]
        )
        contrastes["Mulheres − homens pretos ou pardos"] = (
            grupo_sel["Mulheres pretas ou pardas"] - grupo_sel["Homens pretos ou pardos"]
        )
        contrastes["Mulheres pretas ou pardas − mulheres brancas"] = (
            grupo_sel["Mulheres pretas ou pardas"] - grupo_sel["Mulheres brancas"]
        )
        contrastes.index = contrastes.index.map(sigla)
        st.dataframe(
            contrastes.style.format("{:.1f}", decimal=","),
            use_container_width=True
        )

with tab_evolucao:
    st.subheader("Evolução da diferença entre mulheres e homens")
    st.caption(
        "Diferença calculada como horas das mulheres menos horas dos homens. "
        "A série considera somente os anos disponíveis na Tabela 1.1.1."
    )
    locais = ["Brasil"] + selecionados
    serie = (
        dados[dados["localidade"].isin(locais)]
        .pivot_table(
            index=["ano", "localidade"], columns="sexo",
            values="horas", aggfunc="first"
        )
        .reset_index()
    )
    serie["Horas"] = serie["Mulheres"] - serie["Homens"]
    serie["Localidade"] = serie["localidade"].map(sigla)
    fig = px.line(
        serie, x="ano", y="Horas", color="Localidade",
        markers=True,
        color_discrete_sequence=px.colors.qualitative.Safe
    )
    fig.update_traces(
        line_width=2.5, marker_size=8,
        hovertemplate="%{fullData.name}<br>%{x}: %{y:.1f} h/semana<extra></extra>"
    )
    fig.update_layout(yaxis_title="Horas a mais por semana")
    mostrar(estilo(fig, 500), "grafico_evolucao")

with tab_dados:
    st.subheader("Dados utilizados")
    st.write(
        "Esta versão utiliza exclusivamente a Tabela 1.1.1 do IBGE. "
        "Foram removidas as fontes auxiliares de faixa etária e ocupação, "
        "bem como as seções que dependiam dessas bases."
    )
    tabela = dados[
        (dados["ano"] == ano)
        & (dados["localidade"].isin(["Brasil", *selecionados]))
    ].copy()
    tabela["sigla"] = tabela["localidade"].map(sigla)
    tabela = tabela[
        ["localidade", "sigla", "ano", "sexo", "raca", "grupo", "horas", "cv"]
    ].sort_values(["localidade", "sexo", "raca"])
    st.dataframe(tabela, use_container_width=True, hide_index=True)
    st.download_button(
        "Baixar dados tratados (CSV)",
        data=dados.to_csv(index=False).encode("utf-8-sig"),
        file_name="tabela_1_1_1_tratada.csv",
        mime="text/csv"
    )

st.caption(
    "Nota metodológica: as médias dos estados selecionados são médias simples. "
    "Não representam uma média estadual ponderada pela população."
)
