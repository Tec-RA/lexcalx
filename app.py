import streamlit as st
from datetime import date
from io import BytesIO

from pypdf import PdfReader
from openai import OpenAI
from anthropic import Anthropic
import base64
from pathlib import Path


# =========================================================
# CONFIGURAÇÃO
# =========================================================
st.set_page_config(
    page_title="LexCalx - Cálculos jurídicos",
    page_icon="⚖️",
    layout="wide",
)

APP_DIR = Path(__file__).resolve().parent
CSS_PATH = APP_DIR / "style.css"
LOGO_PATH = APP_DIR / "ralogo.png"


def imagem_base64(caminho: Path) -> str:
    if not caminho.exists():
        st.error(f"Imagem não encontrada: {caminho}")
        return ""

    return base64.b64encode(caminho.read_bytes()).decode()


def carregar_css():
    if not CSS_PATH.exists():
        st.error(f"CSS não encontrado: {CSS_PATH}")
        return

    css = CSS_PATH.read_text(encoding="utf-8")
    logo_bg = imagem_base64(LOGO_PATH)

    css = css.replace("{{LOGO_BG}}", logo_bg)

    st.markdown(
        f"<style>{css}</style>",
        unsafe_allow_html=True,
    )

carregar_css()


def imagem_base64(caminho: Path) -> str:
    if not caminho.exists():
        return ""
    return base64.b64encode(caminho.read_bytes()).decode()


LOGO_BG = imagem_base64(LOGO_PATH)

# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================
def formatar_data_br(data_valor):
    if not data_valor:
        return ""
    return data_valor.strftime("%d/%m/%Y")


def extrair_texto_pdf(arquivo_pdf) -> str:
    if arquivo_pdf is None:
        return ""

    try:
        leitor = PdfReader(BytesIO(arquivo_pdf.read()))
        textos = []

        for pagina in leitor.pages:
            texto = pagina.extract_text() or ""
            textos.append(texto)

        return "\n\n".join(textos).strip()

    except Exception as erro:
        return f"ERRO AO LER PDF: {erro}"


def montar_prompt(dados: dict, texto_decisao: str) -> str:
    return f"""
Você é um assistente especializado em cálculos jurídicos cíveis brasileiros.

Analise a decisão judicial abaixo e elabore os cálculos líquidos ou, quando não for possível calcular exatamente, apresente as premissas, fórmulas, bases de cálculo e pontos pendentes.

IMPORTANTE:
- Use português do Brasil.
- Use datas no formato dd/mm/yyyy.
- Não invente valores que não estejam na decisão.
- Se faltar algum valor base, informe expressamente que o cálculo não pode ser fechado.
- Separe dano moral, dano material, juros, correção monetária, honorários, custas e eventual multa, se houver.
- Identifique o termo inicial dos juros e da correção monetária.
- Use as datas informadas pelo usuário quando forem necessárias.
- Apresente o resultado em formato objetivo, como uma memória de cálculo.
- Se a decisão não for líquida, diga quais dados faltam para liquidar.

DADOS INFORMADOS PELO USUÁRIO:

Número do processo: {dados["numero_processo"]}

Data da distribuição: {dados["data_distribuicao"]}
Data da citação: {dados["data_citacao"]}
Data do evento danoso: {dados["data_evento_danoso"]}

Data juros dano moral: {dados["data_juros_dano_moral"]}
Data correção monetária dano moral: {dados["data_correcao_dano_moral"]}

Data juros dano material: {dados["data_juros_dano_material"]}
Data correção monetária dano material: {dados["data_correcao_dano_material"]}

Data da decisão: {dados["data_decisao"]}
Data da publicação: {dados["data_publicacao"]}

TEXTO DA DECISÃO:

{texto_decisao}
"""


def chamar_gpt(prompt: str) -> str:
    try:
        api_key = st.secrets.get("OPENAI_API_KEY", "")
        modelo = st.secrets.get("OPENAI_MODEL", "gpt-5.2")

        if not api_key:
            return "Chave OPENAI_API_KEY não configurada."

        client = OpenAI(api_key=api_key)

        resposta = client.responses.create(
            model=modelo,
            instructions="Você é um calculista jurídico brasileiro especializado em decisões judiciais líquidas.",
            input=prompt,
        )

        return resposta.output_text

    except Exception as erro:
        return f"Erro ao consultar GPT: {erro}"


def chamar_claude(prompt: str) -> str:
    try:
        api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
        modelo = st.secrets.get("ANTHROPIC_MODEL", "claude-opus-4-7")

        if not api_key:
            return "Chave ANTHROPIC_API_KEY não configurada."

        client = Anthropic(api_key=api_key)

        resposta = client.messages.create(
            model=modelo,
            max_tokens=4096,
            system="Você é um calculista jurídico brasileiro especializado em decisões judiciais líquidas.",
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                }
            ],
        )

        partes = []
        for bloco in resposta.content:
            if getattr(bloco, "type", "") == "text":
                partes.append(bloco.text)

        return "\n".join(partes).strip()

    except Exception as erro:
        return f"Erro ao consultar Claude: {erro}"


def chamar_deepseek(prompt: str) -> str:
    try:
        api_key = st.secrets.get("DEEPSEEK_API_KEY", "")
        modelo = st.secrets.get("DEEPSEEK_MODEL", "deepseek-v4-pro")

        if not api_key:
            return "Chave DEEPSEEK_API_KEY não configurada."

        client = OpenAI(
            api_key=api_key,
            base_url="https://api.deepseek.com",
        )

        resposta = client.chat.completions.create(
            model=modelo,
            messages=[
                {
                    "role": "system",
                    "content": "Você é um calculista jurídico brasileiro especializado em decisões judiciais líquidas.",
                },
                {
                    "role": "user",
                    "content": prompt,
                },
            ],
            stream=False,
        )

        return resposta.choices[0].message.content

    except Exception as erro:
        return f"Erro ao consultar DeepSeek: {erro}"


# =========================================================
# INTERFACE
# =========================================================
st.markdown(
    """
    <div class="hero-title-box">
        <div class="main-title">LexCalx</div>
        <div class="subtitle">
            Cálculos jurídicos
            <span class="subtitle-helper">?</span>
            <span class="subtitle-tooltip">
                LEXCALX realiza análise de decisões judiciais líquidáveis ou passíveis de cálculos.<br>
                Para maior precisão, preencha assertivamente o máximo possível de campos de data.<br>
                Me envia o arquivo da decisão em pdf nativo ou cole no campo final o texto da decisão.
            </span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

numero_processo = st.text_input("Número do processo:")

col1, col2, col3 = st.columns(3)

with col1:
    data_distribuicao = st.date_input(
        "Data da distribuição",
        value=None,
        format="DD/MM/YYYY",
    )

    data_citacao = st.date_input(
        "Data da citação",
        value=None,
        format="DD/MM/YYYY",
    )

    data_evento_danoso = st.date_input(
        "Data do evento danoso",
        value=None,
        format="DD/MM/YYYY",
    )

with col2:
    data_juros_dano_moral = st.date_input(
        "Data juros dano moral",
        value=None,
        format="DD/MM/YYYY",
    )

    data_correcao_dano_moral = st.date_input(
        "Data correção monetária dano moral",
        value=None,
        format="DD/MM/YYYY",
    )

    data_juros_dano_material = st.date_input(
        "Data juros dano material",
        value=None,
        format="DD/MM/YYYY",
    )

with col3:
    data_correcao_dano_material = st.date_input(
        "Data correção monetária dano material",
        value=None,
        format="DD/MM/YYYY",
    )

    data_decisao = st.date_input(
        "Data da decisão",
        value=None,
        format="DD/MM/YYYY",
    )

    data_publicacao = st.date_input(
        "Data da publicação",
        value=None,
        format="DD/MM/YYYY",
    )

st.divider()

arquivo_pdf = st.file_uploader(
    "Me envia o arquivo PDF da decisão:",
    type=["pdf"],
)

texto_colado = st.text_area(
    "Ou cole aqui o texto da decisão:",
    height=260,
)

executar = st.button("Executar cálculo")

if executar:
    texto_pdf = extrair_texto_pdf(arquivo_pdf)
    texto_decisao = texto_pdf if texto_pdf else texto_colado.strip()

    if not texto_decisao:
        st.warning("Vincule um PDF ou cole o texto da decisão antes de executar o cálculo.")
        st.stop()

    dados = {
        "numero_processo": numero_processo,
        "data_distribuicao": formatar_data_br(data_distribuicao),
        "data_citacao": formatar_data_br(data_citacao),
        "data_evento_danoso": formatar_data_br(data_evento_danoso),
        "data_juros_dano_moral": formatar_data_br(data_juros_dano_moral),
        "data_correcao_dano_moral": formatar_data_br(data_correcao_dano_moral),
        "data_juros_dano_material": formatar_data_br(data_juros_dano_material),
        "data_correcao_dano_material": formatar_data_br(data_correcao_dano_material),
        "data_decisao": formatar_data_br(data_decisao),
        "data_publicacao": formatar_data_br(data_publicacao),
    }

    prompt = montar_prompt(dados, texto_decisao)

    st.subheader("Resultados")

    with st.spinner("Consultando GPT..."):
        resultado_gpt = chamar_gpt(prompt)

    with st.spinner("Consultando Claude..."):
        resultado_claude = chamar_claude(prompt)

    with st.spinner("Consultando DeepSeek..."):
        resultado_deepseek = chamar_deepseek(prompt)

    st.markdown("### Cálculo GPT:")
    st.markdown(
        f'<div class="resultado-box">{resultado_gpt}</div>',
        unsafe_allow_html=True,
    )

    st.markdown("### Cálculo Claude:")
    st.markdown(
        f'<div class="resultado-box">{resultado_claude}</div>',
        unsafe_allow_html=True,
    )

    st.markdown("### Cálculo DeepSeek:")
    st.markdown(
        f'<div class="resultado-box">{resultado_deepseek}</div>',
        unsafe_allow_html=True,
    )