"""Painel sobre a divisão do trabalho de cuidado não remunerado entre os estados."""

import io
import math
import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

BASE = Path(__file__).parent
ARQ_RACA = BASE / "Tabela 1.1.1.xlsx"
ARQ_IDADE = BASE / "dados" / "sidra_9378_idade.csv"
ARQ_OCUPACAO = BASE / "dados" / "sidra_7013_ocupacao.csv"

SIGLAS = {
    "Rondônia": "RO", "Acre": "AC", "Amazonas": "AM", "Roraima": "RR", "Pará": "PA",
    "Amapá": "AP", "Tocantins": "TO", "Maranhão": "MA", "Piauí": "PI", "Ceará": "CE",
    "Rio Grande do Norte": "RN", "Paraíba": "PB", "Pernambuco": "PE", "Alagoas": "AL",
    "Sergipe": "SE", "Bahia": "BA", "Minas Gerais": "MG", "Espírito Santo": "ES",
    "Rio de Janeiro": "RJ", "São Paulo": "SP", "Paraná": "PR", "Santa Catarina": "SC",
    "Rio Grande do Sul": "RS", "Mato Grosso do Sul": "MS", "Mato Grosso": "MT",
    "Goiás": "GO", "Distrito Federal": "DF",
}
LOCALIDADES = {"Brasil", "Norte", "Nordeste", "Sudeste", "Sul", "Centro-Oeste", *SIGLAS}
PADRAO_UFS = ["Bahia", "Pernambuco", "Amazonas", "Ceará"]

# Colunas da Tabela 1.1.1 do IBGE, na ordem em que aparecem na planilha.
COLUNAS_RACA = {
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

# Paleta validada com o script do guia de visualização (fundo claro). A cor segue a
# categoria: filtrar estados nunca muda a cor de um grupo.
COR_SEXO = {"Homens": "#2a78d6", "Mulheres": "#eb6834"}
COR_GRUPO = dict(zip(GRUPOS.values(), ["#2a78d6", "#eb6834", "#1baf7a", "#eda100"]))
COR_OCUPACAO = {"Ocupadas": "#1baf7a", "Não ocupadas": "#6250d6"}
TINTA = "#0b0b0b"
TINTA_2 = "#52514e"
NEUTRO = "#a8a79f"
GRADE = "#e6e5e0"

REPORTAGENS = {
    "Carga mental feminina": {
        "imagem": BASE / "assets" / "carga_mental.svg",
        "itens": [
            {
                "titulo": "Mulheres com emprego dedicam 6h48 a mais que homens ao trabalho doméstico",
                "fonte": "Gênero e Número · Schirlei Alves · 22/08/2023",
                "resumo": "Com trabalho remunerado, elas somam 17h48 semanais em afazeres e cuidados, "
                          "contra 11h deles. Entre 2019 e 2022, a diferença geral caiu de 10h36 para 9h36.",
                "url": "https://www.generonumero.media/reportagens/trabalho-domestico-mulheres/",
            },
            {
                "titulo": "Pós-pandemia: 45% das mulheres mostram algum tipo de transtorno mental",
                "fonte": "Agência Brasil · Ana Cristina Campos · 30/08/2023",
                "resumo": "O relatório Esgotadas, da ONG Think Olga, aponta transtornos mentais em cerca de "
                          "45% das brasileiras, e 86% das entrevistadas relatam excesso de responsabilidades.",
                "url": "https://agenciabrasil.ebc.com.br/saude/noticia/2023-08/pos-pandemia-45-das-mulheres-mostram-algum-tipo-de-transtorno-mental",
            },
            {
                "titulo": "Invisível e não remunerado: como o trabalho doméstico provoca carga mental e afeta as mulheres",
                "fonte": "ND Mais · Maria Fernanda Salinet · 27/03/2022",
                "resumo": "A reportagem relaciona o trabalho doméstico à saúde mental e cita estudo de 2021 "
                          "liderado pelo Instituto de Psiquiatria da USP: 45,5% das mulheres relataram "
                          "sintomas de depressão.",
                "url": "https://ndmais.com.br/cidadania/invisivel-e-nao-remunerado-como-o-trabalho-domestico-provoca-carga-mental-e-afeta-as-mulheres/",
            },
        ],
    },
    "Economia do cuidado": {
        "imagem": BASE / "assets" / "economia_cuidado.svg",
        "itens": [
            {
                "titulo": "Economia do Cuidado: a invisibilidade do trabalho feito por mulheres",
                "fonte": "IREE · Samantha Maia e Juliana Pithon · 24/04/2024",
                "resumo": "Com base no FGV Ibre, o texto estima o trabalho de cuidado em cerca de 13% do PIB, "
                          "com aproximadamente 65% dele feito por mulheres, sobretudo mulheres negras.",
                "url": "https://iree.org.br/economia-do-cuidado-a-invisibilidade-do-trabalho-feito-por-mulheres/",
            },
            {
                "titulo": "Comissão aprova entrada da economia do cuidado no cálculo do PIB",
                "fonte": "Agência Câmara · 15/07/2024",
                "resumo": "Comissão aprovou substitutivo ao PL 638/19 que prevê uma conta-satélite da economia "
                          "do cuidado. Na data da reportagem, o texto ainda passaria por outras três comissões.",
                "url": "https://www.camara.leg.br/noticias/1083448-comissao-aprova-entrada-da-economia-do-cuidado-no-calculo-do-pib/",
            },
            {
                "titulo": "Política Nacional de Cuidados vai combater desigualdades no mercado de trabalho, diz governo",
                "fonte": "Agência Câmara · 22/05/2024",
                "resumo": "Dados citados na audiência: 31% das mulheres que deixaram de procurar trabalho "
                          "apontaram a necessidade de cuidar de alguém como motivo, contra 3% dos homens.",
                "url": "https://www.camara.leg.br/noticias/1065368-politica-nacional-de-cuidados-vai-combater-desigualdades-no-mercado-de-trabalho-diz-governo/",
            },
        ],
    },
}


# ---------------------------------------------------------------- leitura dos dados

class FormatoInvalido(ValueError):
    """O arquivo foi lido, mas não segue um formato aceito pelo painel."""


def normalizar_sexo(valor):
    texto = str(valor).strip().lower()
    if texto in {"homem", "homens", "masculino"}:
        return "Homens"
    if texto in {"mulher", "mulheres", "feminino"}:
        return "Mulheres"
    return "Total" if texto == "total" else None


def ler_xlsx_ibge(origem):
    """Lê as abas de uma planilha no formato da Tabela 1.1.1 do IBGE (uma por ano)."""
    planilha = pd.ExcelFile(origem)
    medidas = {}
    for aba in planilha.sheet_names:
        achado = re.fullmatch(r"(\d{4})( \(CV\))?", aba.strip())
        if not achado:
            continue
        bruto = pd.read_excel(planilha, aba, header=None)
        if bruto.shape[1] < 8:
            raise FormatoInvalido(f"A aba '{aba}' não tem as 8 colunas esperadas.")
        medida = "cv" if achado.group(2) else "horas"
        for _, linha in bruto[bruto[0].isin(LOCALIDADES)].iterrows():
            for coluna, (sexo, raca) in COLUNAS_RACA.items():
                chave = (linha[0], int(achado.group(1)), sexo, raca)
                medidas.setdefault(chave, {})[medida] = pd.to_numeric(linha[coluna], errors="coerce")
    if not medidas:
        raise FormatoInvalido("Nenhuma aba com nome de ano e linhas de UF foi encontrada.")
    linhas = [
        {"localidade": k[0], "ano": k[1], "sexo": k[2], "raca": k[3], **v}
        for k, v in medidas.items()
    ]
    tabela = pd.DataFrame(linhas)
    if "cv" not in tabela:
        tabela["cv"] = float("nan")
    return tabela.dropna(subset=["horas"])


def limpar_tabela(tabela):
    """Padroniza uma tabela no formato longo e descarta linhas não reconhecidas."""
    tabela = tabela.copy()
    if "cv" not in tabela:
        tabela["cv"] = float("nan")
    for coluna in ("horas", "cv"):
        tabela[coluna] = pd.to_numeric(
            tabela[coluna].astype(str).str.replace(",", ".", regex=False), errors="coerce"
        )
    tabela["localidade"] = tabela["localidade"].astype(str).str.strip()
    tabela["sexo"] = tabela["sexo"].map(normalizar_sexo)
    tabela["ano"] = pd.to_numeric(tabela["ano"], errors="coerce")
    tabela = tabela[tabela["localidade"].isin(LOCALIDADES)].dropna(subset=["sexo", "ano", "horas"])
    return tabela.assign(ano=tabela["ano"].astype(int))


@st.cache_data
def carregar_padrao():
    return {
        "raca": ler_xlsx_ibge(ARQ_RACA),
        "idade": limpar_tabela(pd.read_csv(ARQ_IDADE, encoding="utf-8-sig")),
        "ocupacao": limpar_tabela(pd.read_csv(ARQ_OCUPACAO, encoding="utf-8-sig")),
    }


@st.cache_data
def ler_upload(nome, conteudo):
    """Retorna (tipo da tabela, dados). Levanta FormatoInvalido se o formato não for aceito."""
    buffer = io.BytesIO(conteudo)
    if nome.lower().endswith(".xlsx"):
        abas = pd.ExcelFile(buffer).sheet_names
        buffer.seek(0)
        if any(re.fullmatch(r"\d{4}( \(CV\))?", a.strip()) for a in abas):
            return "raca", ler_xlsx_ibge(buffer)
        tabela = pd.read_excel(buffer)
    else:
        try:
            tabela = pd.read_csv(buffer, sep=None, engine="python", encoding="utf-8-sig")
        except UnicodeDecodeError:
            buffer.seek(0)
            tabela = pd.read_csv(buffer, sep=None, engine="python", encoding="latin-1")

    tabela.columns = [str(c).strip().lower() for c in tabela.columns]
    faltando = {"localidade", "ano", "sexo", "horas"} - set(tabela.columns)
    if faltando:
        raise FormatoInvalido("Faltam as colunas: " + ", ".join(sorted(faltando)) + ".")
    tipos = {"raca": "raca", "faixa_etaria": "idade", "ocupacao": "ocupacao"}
    presentes = [c for c in tipos if c in tabela.columns]
    if len(presentes) != 1:
        raise FormatoInvalido("Inclua exatamente uma destas colunas: raca, faixa_etaria ou ocupacao.")

    limpa = limpar_tabela(tabela)
    if limpa.empty:
        raise FormatoInvalido("Nenhuma linha válida. Confira os nomes das localidades e os valores de horas.")
    return tipos[presentes[0]], limpa


# ---------------------------------------------------------------- cálculos

def fmt(valor, casas=1):
    if valor is None or pd.isna(valor):
        return "—"
    return f"{valor:.{casas}f}".replace(".", ",")


def fmt_sinal(valor):
    if valor is None or pd.isna(valor):
        return "—"
    if round(valor, 1) == 0:
        return "0,0"
    return f"{valor:+.1f}".replace(".", ",")


def sigla(localidade):
    return SIGLAS.get(localidade, localidade)


def significativo(dif, cv_pares):
    """Teste z (95%) de uma diferença, com erro-padrão obtido de horas × CV.

    `cv_pares` lista (horas, cv) de cada estimativa. Assume estimativas independentes
    e usa CVs publicados com uma casa decimal, portanto é uma aproximação.
    """
    variancia = sum((h * c / 100) ** 2 for h, c in cv_pares)
    if pd.isna(variancia) or variancia <= 0:
        return "—"
    p = math.erfc(abs(dif) / math.sqrt(variancia) / math.sqrt(2))
    return "Sim" if p < 0.05 else "Não"


def por_sexo(tabela):
    """Horas e CV de homens e mulheres por localidade, com a diferença M − H."""
    horas = tabela.pivot_table(index="localidade", columns="sexo", values="horas", aggfunc="first")
    cvs = tabela.pivot_table(index="localidade", columns="sexo", values="cv", aggfunc="first")
    horas = horas.reindex(columns=["Homens", "Mulheres"])
    cvs = cvs.reindex(columns=["Homens", "Mulheres"]).add_prefix("cv_")
    saida = horas.join(cvs)
    saida["Diferença"] = saida["Mulheres"] - saida["Homens"]
    return saida


def por_grupo(tabela):
    """Horas por localidade nos quatro grupos de sexo e raça."""
    dados = tabela.assign(grupo=[GRUPOS.get(p) for p in zip(tabela["sexo"], tabela["raca"])])
    dados = dados.dropna(subset=["grupo"])
    return dados.pivot_table(index="localidade", columns="grupo", values="horas", aggfunc="first").reindex(
        columns=list(GRUPOS.values())
    )


def valor_brasil(quadro, coluna):
    return quadro.at["Brasil", coluna] if "Brasil" in quadro.index else float("nan")


def estilo(fig, altura=380):
    fig.update_layout(
        height=altura,
        separators=",.",
        margin=dict(l=64, r=16, t=48, b=56),
        font=dict(color=TINTA_2, size=13),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=1.02, yanchor="bottom", x=0, title_text=""),
        hoverlabel=dict(bgcolor="white", font_color=TINTA),
        barcornerradius=4,
    )
    fig.update_xaxes(title=None, showgrid=False, linecolor=NEUTRO, automargin=True)
    fig.update_yaxes(gridcolor=GRADE, zeroline=False, automargin=True)
    return fig


def mostrar(fig, chave):
    st.plotly_chart(fig, width="stretch", theme=None, key=chave, config={"displayModeBar": False})


def tabela_br(quadro, **kwargs):
    """Mostra uma tabela com vírgula decimal e travessão no lugar de valores ausentes."""
    quadro = quadro.copy()
    for coluna in quadro.select_dtypes("number").columns[quadro.select_dtypes("number").isna().any()]:
        quadro[coluna] = quadro[coluna].map(fmt)
    numericas = quadro.select_dtypes("number").columns
    st.dataframe(
        quadro.style.format("{:.1f}", subset=numericas, decimal=",", na_rep="—"),
        width="stretch",
        **kwargs,
    )


# ---------------------------------------------------------------- estrutura da página

st.set_page_config(page_title="Trabalho de cuidado não remunerado", page_icon="🏠", layout="wide")

padrao = carregar_padrao()
dados = dict(padrao)
origem = {"raca": "Tabela 1.1.1.xlsx", "idade": "SIDRA 9378", "ocupacao": "SIDRA 7013"}

st.sidebar.header("Filtros")
area_filtros = st.sidebar.container()
st.sidebar.divider()
st.sidebar.header("Seus dados")
enviados = st.sidebar.file_uploader(
    "Enviar arquivos (.xlsx ou .csv)",
    type=["xlsx", "csv"],
    accept_multiple_files=True,
    help="Aceita planilhas no formato da Tabela 1.1.1 do IBGE, com uma aba por ano, ou tabelas com as "
         "colunas localidade, ano, sexo, horas e uma destas: raca, faixa_etaria ou ocupacao. "
         "A coluna cv é opcional. Os arquivos da aba Dados servem de modelo.",
)
for arquivo in enviados or []:
    try:
        tipo, tabela = ler_upload(arquivo.name, arquivo.getvalue())
    except FormatoInvalido as erro:
        st.sidebar.error(f"{arquivo.name}: {erro}")
        continue
    except Exception:
        st.sidebar.error(f"{arquivo.name}: não foi possível ler o arquivo. Confira se é um .xlsx ou .csv válido.")
        continue
    dados[tipo] = tabela
    origem[tipo] = arquivo.name
    st.sidebar.success(f"{arquivo.name}: {len(tabela)} linhas carregadas.")

ufs_disponiveis = sorted(
    {loc for t in dados.values() for loc in t["localidade"].unique() if loc in SIGLAS},
    key=sigla,
)
anos = sorted({int(a) for t in dados.values() for a in t["ano"].unique()}, reverse=True)
faixas = sorted(dados["idade"]["faixa_etaria"].astype(str).unique(), key=lambda f: (f != "Total", f))

with area_filtros:
    ufs = st.multiselect(
        "Estados",
        ufs_disponiveis,
        default=[u for u in PADRAO_UFS if u in ufs_disponiveis],
        format_func=lambda u: f"{sigla(u)} – {u}",
    )
    faixa = st.selectbox(
        "Faixa etária",
        faixas,
        format_func=lambda f: "Todas (14 anos ou mais)" if f == "Total" else f,
        help="O IBGE não cruza idade com raça nem com ocupação. Por isso, este filtro altera apenas "
             "os indicadores e gráficos por sexo, inclusive a evolução.",
    )
    ano = st.selectbox(
        "Ano",
        anos,
        help="Não há dados publicados para 2020 e 2021. A aba Evolução mostra sempre todos os anos.",
    )

idade_faixa = dados["idade"][dados["idade"]["faixa_etaria"].astype(str) == faixa]
sexo_ano = por_sexo(idade_faixa[idade_faixa["ano"] == ano])
grupo_ano = por_grupo(dados["raca"][dados["raca"]["ano"] == ano])
sexo_sel = sexo_ano.reindex([u for u in ufs if u in sexo_ano.index])
grupo_sel = grupo_ano.reindex([u for u in ufs if u in grupo_ano.index])

st.title("Trabalho de cuidado não remunerado nos estados")
st.markdown(
    "Média de horas semanais dedicadas a afazeres domésticos e ao cuidado de pessoas, entre quem tem "
    "14 anos ou mais e realizou essas atividades. Fonte: IBGE, PNAD Contínua."
)

# Indicadores: média simples dos estados, pois as tabelas não trazem a população de cada grupo.
if ufs and not sexo_sel.empty:
    mulheres, homens = sexo_sel["Mulheres"].mean(), sexo_sel["Homens"].mean()
    diferenca = sexo_sel["Diferença"].mean()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric(
        "Mulheres (h/semana)", fmt(mulheres),
        f"{fmt_sinal(mulheres - valor_brasil(sexo_ano, 'Mulheres'))} h vs Brasil", delta_color="off",
    )
    k2.metric(
        "Homens (h/semana)", fmt(homens),
        f"{fmt_sinal(homens - valor_brasil(sexo_ano, 'Homens'))} h vs Brasil", delta_color="off",
    )
    k3.metric(
        "Diferença mulheres − homens", f"{fmt(diferenca)} h",
        f"{fmt_sinal(diferenca - valor_brasil(sexo_ano, 'Diferença'))} h vs Brasil", delta_color="inverse",
    )
    if not grupo_sel.empty:
        extremos = grupo_sel["Mulheres pretas ou pardas"] - grupo_sel["Homens brancos"]
        k4.metric(
            "Pretas ou pardas − homens brancos", f"{fmt(extremos.mean())} h",
            help="Diferença entre mulheres pretas ou pardas e homens brancos. Não muda com o filtro de faixa etária, pois o IBGE não cruza raça e idade.",
        )
    st.caption(f"Ano de {ano}. Média simples dos estados selecionados, sem ponderar pela população de cada um.")

aba_sexo, aba_evolucao, aba_ocupacao, aba_noticias, aba_dados = st.tabs(
    ["Sexo e raça", "Evolução", "Ocupação", "Reportagens", "Dados"]
)
SEM_UF = "Selecione ao menos um estado na barra lateral."

# ---------------------------------------------------------------- pergunta 1

with aba_sexo:
    st.subheader("Qual é a diferença de horas entre homens e mulheres por estado?")
    if not ufs:
        st.info(SEM_UF)
    elif sexo_sel.empty:
        st.info(f"Não há dados por sexo para {ano} nos estados selecionados.")
    else:
        maior, menor = sexo_sel["Diferença"].idxmax(), sexo_sel["Diferença"].idxmin()
        if len(sexo_sel) > 1:
            st.markdown(
                f"Em {ano}, a maior diferença entre os estados selecionados está em **{maior}** "
                f"({fmt(sexo_sel.at[maior, 'Diferença'])} h por semana) e a menor em **{menor}** "
                f"({fmt(sexo_sel.at[menor, 'Diferença'])} h). "
                f"A média do Brasil é {fmt(valor_brasil(sexo_ano, 'Diferença'))} h."
            )
        else:
            st.markdown(
                f"Em {ano}, em **{maior}**, as mulheres dedicam {fmt(sexo_sel.at[maior, 'Diferença'])} h "
                f"por semana a mais que os homens. A média do Brasil é "
                f"{fmt(valor_brasil(sexo_ano, 'Diferença'))} h."
            )

        ordem = sexo_sel.sort_values("Diferença", ascending=False)
        comparacao = pd.concat([sexo_ano.reindex(["Brasil"]).dropna(how="all"), ordem])
        comparacao = comparacao.assign(rotulo=[sigla(i) for i in comparacao.index])

        col_a, col_b = st.columns([3, 2])
        with col_a:
            st.markdown("**Horas semanais por sexo**")
            longo = comparacao.melt(
                id_vars="rotulo", value_vars=["Homens", "Mulheres"], var_name="Sexo", value_name="Horas"
            )
            fig = px.bar(
                longo, x="rotulo", y="Horas", color="Sexo", barmode="group",
                color_discrete_map=COR_SEXO, category_orders={"Sexo": ["Homens", "Mulheres"]},
                text_auto=".1f" if len(comparacao) <= 9 else False,
            )
            fig.update_traces(
                hovertemplate="%{x} · %{fullData.name}<br>%{y:.1f} h por semana<extra></extra>",
                textposition="outside", cliponaxis=False, textfont_color=TINTA_2,
            )
            estilo(fig, 400).update_layout(bargroupgap=0.06, yaxis_title="Horas por semana")
            mostrar(fig, "barras_sexo")
        with col_b:
            st.markdown("**Diferença mulheres − homens**")
            cores = [NEUTRO if i == "Brasil" else TINTA_2 for i in comparacao.index]
            fig = go.Figure(
                go.Bar(
                    x=comparacao["Diferença"], y=comparacao["rotulo"], orientation="h",
                    marker_color=cores, texttemplate="%{x:.1f} h", textposition="outside",
                    cliponaxis=False, textfont_color=TINTA_2,
                    hovertemplate="%{y}<br>%{x:.1f} h a mais para as mulheres<extra></extra>",
                )
            )
            estilo(fig, max(400, 26 * len(comparacao)))
            fig.update_layout(yaxis=dict(autorange="reversed", showgrid=False), margin=dict(r=56))
            fig.update_xaxes(showgrid=True, gridcolor=GRADE, rangemode="tozero")
            mostrar(fig, "barras_diferenca")

    st.divider()
    st.subheader("E por sexo e raça?")
    if not ufs:
        st.info(SEM_UF)
    elif grupo_sel.empty:
        st.info(f"Não há dados por raça para {ano} nos estados selecionados.")
    else:
        st.caption("Dados para todas as idades (14 anos ou mais). O IBGE não publica o cruzamento de raça e idade.")
        ordem_raca = grupo_sel.assign(rotulo=[sigla(i) for i in grupo_sel.index]).sort_values(
            "Mulheres pretas ou pardas", ascending=False
        )
        brasil_raca = grupo_ano.reindex(["Brasil"]).dropna(how="all").assign(rotulo="Brasil")
        barras = pd.concat([brasil_raca, ordem_raca]).melt(
            id_vars="rotulo", value_vars=list(GRUPOS.values()), var_name="Grupo", value_name="Horas"
        )
        fig = px.bar(
            barras, x="rotulo", y="Horas", color="Grupo", barmode="group",
            color_discrete_map=COR_GRUPO, category_orders={"Grupo": list(GRUPOS.values())},
        )
        fig.update_traces(hovertemplate="%{x} · %{fullData.name}<br>%{y:.1f} h por semana<extra></extra>")
        estilo(fig, 400).update_layout(bargroupgap=0.06, yaxis_title="Horas por semana")
        mostrar(fig, "barras_raca")

        contrastes = pd.DataFrame(
            {
                "Mulheres − homens, brancos": grupo_sel["Mulheres brancas"] - grupo_sel["Homens brancos"],
                "Mulheres − homens, pretos ou pardos":
                    grupo_sel["Mulheres pretas ou pardas"] - grupo_sel["Homens pretos ou pardos"],
                "Pretas ou pardas − brancas (mulheres)":
                    grupo_sel["Mulheres pretas ou pardas"] - grupo_sel["Mulheres brancas"],
                "Pretos ou pardos − brancos (homens)":
                    grupo_sel["Homens pretos ou pardos"] - grupo_sel["Homens brancos"],
            }
        ).rename(index=sigla).rename_axis("Estado")
        st.markdown("**Diferenças em horas por semana**")
        tabela_br(contrastes)

    st.divider()
    st.subheader("Distribuição de horas entre os estados")
    distribuicao = grupo_ano[grupo_ano.index.isin(SIGLAS)]
    if distribuicao.empty:
        st.info(f"Não há dados por raça para {ano}.")
    else:
        st.caption(
            "Cada caixa resume as médias dos estados disponíveis em cada grupo. Ela não mostra a distribuição "
            "entre pessoas, que exigiria os microdados da PNAD. Os pontos escuros são os estados selecionados."
        )
        fig = go.Figure()
        for grupo in GRUPOS.values():
            fig.add_trace(
                go.Box(
                    y=distribuicao[grupo], x=[grupo] * len(distribuicao), name=grupo,
                    marker_color=COR_GRUPO[grupo], boxpoints=False, hoverinfo="skip", showlegend=False,
                )
            )
        selecionados = distribuicao[distribuicao.index.isin(ufs)]
        if not selecionados.empty:
            pontos = selecionados.rename_axis("uf").reset_index().melt(
                id_vars="uf", value_vars=list(GRUPOS.values()), var_name="Grupo", value_name="Horas"
            )
            fig.add_trace(
                go.Scatter(
                    x=pontos["Grupo"], y=pontos["Horas"], mode="markers", name="Estados selecionados",
                    marker=dict(size=10, color=TINTA, line=dict(color="white", width=2)),
                    customdata=pontos["uf"],
                    hovertemplate="%{customdata}<br>%{x}: %{y:.1f} h por semana<extra></extra>",
                )
            )
        estilo(fig, 420).update_layout(yaxis_title="Horas por semana", boxgap=0.45)
        mostrar(fig, "boxplot")

# ---------------------------------------------------------------- pergunta 2

with aba_evolucao:
    st.subheader("Como a divisão evoluiu nos estados e na média nacional?")
    st.caption(
        "Diferença entre as horas de mulheres e homens. Esta aba mostra todos os anos e respeita a "
        "faixa etária escolhida. Não há dados publicados para 2020 e 2021, por isso o período está sombreado."
    )
    if not ufs:
        st.info(SEM_UF)
    else:
        series = {a: por_sexo(idade_faixa[idade_faixa["ano"] == a]) for a in sorted(idade_faixa["ano"].unique())}
        linhas = [
            {"ano": a, "localidade": loc, "Diferença": q.at[loc, "Diferença"]}
            for a, q in series.items() for loc in q.index
        ]
        evolucao = pd.DataFrame(linhas, columns=["ano", "localidade", "Diferença"])
        ufs_ok = [u for u in ufs if u in set(evolucao["localidade"])]

        if not ufs_ok:
            st.info("Não há série histórica para os estados selecionados.")
        else:
            paineis = []
            for uf in ufs_ok:
                paineis.append(evolucao[evolucao["localidade"] == uf].assign(painel=f"{sigla(uf)} – {uf}", Série="Estado"))
                paineis.append(evolucao[evolucao["localidade"] == "Brasil"].assign(painel=f"{sigla(uf)} – {uf}", Série="Brasil"))
            painel = pd.concat(paineis)
            colunas = min(len(ufs_ok), 4)
            n_linhas = math.ceil(len(ufs_ok) / colunas)
            fig = px.line(
                painel, x="ano", y="Diferença", color="Série", line_dash="Série",
                facet_col="painel", facet_col_wrap=colunas, markers=True,
                facet_row_spacing=0.28 / max(n_linhas, 1) if n_linhas > 1 else 0.05, facet_col_spacing=0.03,
                color_discrete_map={"Estado": COR_SEXO["Mulheres"], "Brasil": TINTA_2},
                line_dash_map={"Estado": "solid", "Brasil": "dash"},
                category_orders={"Série": ["Estado", "Brasil"], "painel": [f"{sigla(u)} – {u}" for u in ufs_ok]},
            )
            fig.update_traces(
                line_width=2, marker_size=8,
                hovertemplate="%{x}: %{y:.1f} h a mais para as mulheres<extra>%{fullData.name}</extra>",
            )
            fig.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1], font=dict(color=TINTA, size=13)))
            estilo(fig, 150 + 250 * n_linhas).update_layout(
                margin=dict(t=84), legend=dict(y=1, yanchor="top", yref="container")
            )
            anos_serie = sorted(series)
            fig.update_xaxes(title=None, tickvals=anos_serie, tickformat="d", tickangle=-45, matches=None)
            fig.update_yaxes(title=None)
            if {2019, 2022} <= set(anos_serie):
                fig.add_vrect(x0=2019, x1=2022, row="all", col="all", fillcolor=GRADE, opacity=0.45, line_width=0, layer="below")
            mostrar(fig, "evolucao")

            primeiro, ultimo = anos_serie[0], anos_serie[-1]
            q0, q1 = series[primeiro], series[ultimo]
            brasil_var = (
                q1.at["Brasil", "Diferença"] - q0.at["Brasil", "Diferença"]
                if "Brasil" in q0.index and "Brasil" in q1.index else float("nan")
            )
            resumo = []
            for loc in ufs_ok + ["Brasil"]:
                if loc not in q0.index or loc not in q1.index:
                    continue
                variacao = q1.at[loc, "Diferença"] - q0.at[loc, "Diferença"]
                pares = [
                    (q[s][loc], q[f"cv_{s}"][loc]) for q in (q0, q1) for s in ("Homens", "Mulheres")
                ]
                resumo.append(
                    {
                        "Local": loc,
                        f"Diferença em {primeiro} (h)": q0.at[loc, "Diferença"],
                        f"Diferença em {ultimo} (h)": q1.at[loc, "Diferença"],
                        "Variação (h)": variacao,
                        "Variação menos a do Brasil (h)": variacao - brasil_var if loc != "Brasil" else float("nan"),
                        "Variação significativa (95%)": significativo(variacao, pares),
                    }
                )
            so_ufs = [r for r in resumo if r["Local"] != "Brasil"]
            if so_ufs:
                subiu = sum(r["Variação (h)"] > 0 for r in so_ufs)
                caiu = sum(r["Variação (h)"] < 0 for r in so_ufs)
                signif = sum(r["Variação significativa (95%)"] == "Sim" for r in so_ufs)
                st.markdown(
                    f"Entre {primeiro} e {ultimo}, a diferença no Brasil variou **{fmt_sinal(brasil_var)} h**. "
                    f"Dos {len(so_ufs)} estados selecionados, ela aumentou em **{subiu}** e diminuiu em **{caiu}**. "
                    f"A variação é estatisticamente significativa em {signif} deles."
                )
            st.markdown(f"**Variação da diferença entre {primeiro} e {ultimo}**")
            tabela_br(pd.DataFrame(resumo).set_index("Local"))
            st.caption(
                "Variação negativa indica que a diferença diminuiu. O teste é aproximado, pois usa os coeficientes "
                "de variação publicados com uma casa decimal e assume que as estimativas são independentes."
            )

# ---------------------------------------------------------------- pergunta 3

with aba_ocupacao:
    st.subheader("Há diferença entre quem está ocupado e quem não está?")
    st.warning(
        "O IBGE compara **ocupadas** e **não ocupadas**. O segundo grupo inclui pessoas desempregadas, "
        "mas também aposentadas, estudantes e outras fora da força de trabalho. Portanto, os dados "
        "não permitem isolar as pessoas desempregadas."
    )
    if not ufs:
        st.info(SEM_UF)
    else:
        ocup = dados["ocupacao"]
        ocup = ocup[(ocup["ano"] == ano) & ocup["sexo"].isin(COR_SEXO) & ocup["ocupacao"].isin(COR_OCUPACAO)]
        ocup = ocup[ocup["localidade"].isin(["Brasil", *ufs])]
        if ocup[ocup["localidade"] != "Brasil"].empty:
            st.info(f"Não há dados de ocupação para {ano} nos estados selecionados.")
        else:
            ocup = ocup.assign(
                rotulo=ocup["localidade"].map(sigla),
                margem=1.96 * ocup["horas"] * ocup["cv"] / 100,
            )
            resultados = []
            for (loc, sexo), parte in ocup.groupby(["localidade", "sexo"]):
                linhas_ocup = parte.set_index("ocupacao")
                if not set(COR_OCUPACAO) <= set(linhas_ocup.index):
                    continue
                o, n = linhas_ocup.loc["Ocupadas"], linhas_ocup.loc["Não ocupadas"]
                dif = n["horas"] - o["horas"]
                resultados.append(
                    {
                        "Local": loc, "Sexo": sexo, "Ocupadas (h)": o["horas"], "Não ocupadas (h)": n["horas"],
                        "Diferença (h)": dif,
                        "Significativa (95%)": significativo(dif, [(o["horas"], o["cv"]), (n["horas"], n["cv"])]),
                    }
                )
            resultado = pd.DataFrame(resultados)
            estados = resultado[resultado["Local"] != "Brasil"]
            if not estados.empty:
                dif_sexo = estados.groupby("Sexo")["Diferença (h)"].mean()
                sim = int((estados["Significativa (95%)"] == "Sim").sum())
                st.markdown(
                    f"Em {ano}, nos estados selecionados, quem não está ocupado dedica em média "
                    f"**{fmt(dif_sexo.get('Mulheres'))} h a mais entre as mulheres** e "
                    f"**{fmt(dif_sexo.get('Homens'))} h a mais entre os homens**. "
                    f"A diferença é estatisticamente significativa em {sim} de {len(estados)} comparações. "
                    "Isso é uma associação e não indica causa."
                )
            ordem_ocup = ["Brasil"] + [sigla(u) for u in ufs]
            fig = px.bar(
                ocup, x="rotulo", y="horas", color="ocupacao", facet_col="sexo", barmode="group",
                error_y="margem", color_discrete_map=COR_OCUPACAO,
                category_orders={
                    "ocupacao": list(COR_OCUPACAO), "sexo": list(COR_SEXO), "rotulo": ordem_ocup,
                },
            )
            fig.update_traces(
                hovertemplate="%{x} · %{fullData.name}<br>%{y:.1f} h por semana<extra></extra>",
                error_y_color=TINTA_2, error_y_thickness=1.2, error_y_width=3,
            )
            fig.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1], font=dict(color=TINTA, size=13)))
            estilo(fig, 460).update_layout(
                bargroupgap=0.06, margin=dict(t=84), legend=dict(y=1, yanchor="top", yref="container")
            )
            fig.update_xaxes(title=None)
            fig.update_yaxes(title=None)
            fig.update_yaxes(title="Horas por semana", col=1)
            mostrar(fig, "barras_ocupacao")
            st.caption(
                "As linhas sobre as barras mostram o intervalo de confiança de 95%. A faixa etária não se aplica "
                "aqui, pois o IBGE não cruza ocupação e idade."
            )
            tabela_br(resultado.set_index(["Local", "Sexo"]))

# ---------------------------------------------------------------- reportagens

with aba_noticias:
    st.caption(
        "Seleção de textos publicados. As informações resumem cada reportagem na sua data de publicação "
        "e podem ter mudado desde então. As ilustrações foram criadas para este painel."
    )
    for tema, conteudo in REPORTAGENS.items():
        st.subheader(tema)
        col_img, col_txt = st.columns([1, 3])
        col_img.image(str(conteudo["imagem"]), width="stretch")
        with col_txt:
            for item in conteudo["itens"]:
                st.markdown(f"**[{item['titulo']}]({item['url']})**")
                st.caption(item["fonte"])
                st.markdown(item["resumo"])
        st.divider()

# ---------------------------------------------------------------- dados e método

with aba_dados:
    st.subheader("Origem dos dados")
    st.markdown(
        f"- **Sexo e raça:** {origem['raca']}\n"
        f"- **Sexo e faixa etária:** {origem['idade']}\n"
        f"- **Sexo e ocupação:** {origem['ocupacao']}\n\n"
        "As tabelas padrão são do IBGE, PNAD Contínua, com anos de 2016 a 2019 e 2022. As médias consideram "
        "apenas quem declarou ter feito essas atividades na semana de referência."
    )
    tabelas = {
        "Sexo e raça": ("raca", "modelo_raca.csv"),
        "Sexo e faixa etária": ("idade", "modelo_faixa_etaria.csv"),
        "Sexo e ocupação": ("ocupacao", "modelo_ocupacao.csv"),
    }
    for titulo, (tipo, arquivo) in tabelas.items():
        st.markdown(f"**{titulo}**")
        visivel = dados[tipo]
        visivel = visivel[(visivel["ano"] == ano) & visivel["localidade"].isin(["Brasil", *ufs])]
        visivel = visivel.drop(columns=["nivel"], errors="ignore").sort_values(["localidade", "sexo"])
        st.dataframe(visivel, width="stretch", hide_index=True)
        st.download_button(
            f"Baixar a tabela completa ({arquivo})",
            dados[tipo].drop(columns=["nivel"], errors="ignore").to_csv(index=False).encode("utf-8-sig"),
            file_name=arquivo,
            mime="text/csv",
            key=f"baixar_{tipo}",
        )
