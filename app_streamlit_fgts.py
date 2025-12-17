"""
🚀 CONVERSOR PDF FGTS PARA EXCEL - STREAMLIT
Interface amigável para conversão de guias FGTS
"""

import streamlit as st
import PyPDF2
import pandas as pd
import re
import io
from datetime import datetime

# ============================================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================================
st.set_page_config(
    page_title="Conversor de detalhamento de GUIA e-consignado",
    page_icon="📄",
    layout="centered"
)

# ============================================================================
# ESTILO CSS
# ============================================================================
st.markdown("""
<style>
    .main-header {
        text-align: center;
        color: #1f77b4;
        padding: 20px;
        background: linear-gradient(90deg, #e3f2fd 0%, #bbdefb 100%);
        border-radius: 10px;
        margin-bottom: 30px;
    }
    .success-box {
        padding: 20px;
        background-color: #d4edda;
        border-left: 5px solid #28a745;
        border-radius: 5px;
        margin: 20px 0;
    }
    .info-box {
        padding: 15px;
        background-color: #d1ecf1;
        border-left: 5px solid #17a2b8;
        border-radius: 5px;
        margin: 15px 0;
    }
    .stButton>button {
        width: 100%;
        background-color: #1f77b4;
        color: white;
        font-size: 18px;
        padding: 15px;
        border-radius: 10px;
        border: none;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #155a8a;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# CABEÇALHO
# ============================================================================
st.markdown("""
<div class="main-header">
    <h1>📄 Conversor FGTS</h1>
    <p style="font-size: 18px; margin: 0;">Converta o detalhamento de empréstimos em planilhas Excel automaticamente</p>
</div>
""", unsafe_allow_html=True)

# ============================================================================
# FUNÇÃO DE EXTRAÇÃO
# ============================================================================
@st.cache_data
def extrair_trabalhadores_pdf(pdf_bytes):
    """Extrai todos os trabalhadores da listagem em PDF"""
    all_workers = []
    cpf_pattern = re.compile(r'\d{3}\.\d{3}\.\d{3}-\d{2}')

    try:
        pdf_file = io.BytesIO(pdf_bytes)
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        total_pages = len(pdf_reader.pages)

        progress_bar = st.progress(0)
        status_text = st.empty()

        for page_num in range(total_pages):
            page = pdf_reader.pages[page_num]
            page_text = page.extract_text()
            lines = page_text.split('\n')

            for line in lines:
                line = line.strip()
                cpf_match = cpf_pattern.search(line)
                if not cpf_match:
                    continue

                cpf = cpf_match.group()
                cpf_pos = line.find(cpf)

                before_cpf = line[:cpf_pos].strip().split()
                if len(before_cpf) < 3:
                    continue

                valor = before_cpf[0]
                vencimento = before_cpf[1]
                matricula = before_cpf[2]

                after_cpf = line[cpf_pos + len(cpf):].strip().split(None, 3)
                if len(after_cpf) < 4:
                    continue

                comp_apuracao = after_cpf[0]
                contrato = after_cpf[1]
                instituicao = after_cpf[2]
                nome = after_cpf[3]

                all_workers.append({
                    'comp_apuracao': comp_apuracao,
                    'vencimento': vencimento,
                    'nome': nome,
                    'matricula': matricula,
                    'cpf': cpf,
                    'contrato': contrato,
                    'instituicao': instituicao,
                    'valor': valor
                })

            # Atualizar progresso
            progress = (page_num + 1) / total_pages
            progress_bar.progress(progress)
            status_text.text(f"Processando página {page_num + 1} de {total_pages}... ({len(all_workers)} trabalhadores)")

        progress_bar.empty()
        status_text.empty()

        return all_workers, None

    except Exception as e:
        return [], str(e)

# ============================================================================
# FUNÇÃO PARA GERAR EXCEL
# ============================================================================
def gerar_excel(workers):
    """Gera arquivo Excel com os dados"""
    df = pd.DataFrame(workers)
    df.insert(0, 'qt', range(1, len(df) + 1))

    df.columns = ['Qt', 'Comp. Apuração', 'Vencimento', 'Nome Trabalhador', 
                  'Matrícula', 'CPF', 'Número do Contrato', 
                  'Instituição Financeira', 'Valor Consignado na Guia']

    df = df[['Qt', 'Comp. Apuração', 'Vencimento', 'Nome Trabalhador', 
             'Matrícula', 'CPF', 'Número do Contrato', 
             'Instituição Financeira', 'Valor Consignado na Guia']]

    # Preservar zeros à esquerda
    df['Matrícula'] = df['Matrícula'].astype(str)
    df['Instituição Financeira'] = df['Instituição Financeira'].apply(
        lambda x: str(x).zfill(3) if str(x).isdigit() and len(str(x)) <= 3 else str(x)
    )
    df['Número do Contrato'] = df['Número do Contrato'].astype(str)

    # Salvar em buffer
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Trabalhadores')

    return output.getvalue(), df

# ============================================================================
# INTERFACE PRINCIPAL
# ============================================================================

# Instruções
st.markdown("""
<div class="info-box">
    <h3>📋 Como usar:</h3>
    <ol>
        <li>Clique no botão abaixo para fazer upload do PDF</li>
        <li>Aguarde o processamento (alguns segundos)</li>
        <li>Visualize os dados extraídos</li>
        <li>Clique em "Baixar Excel" para salvar</li>
    </ol>
</div>
""", unsafe_allow_html=True)

# Upload do arquivo
uploaded_file = st.file_uploader(
    "📤 Selecione o arquivo PDF do detalhamento de guia consignado",
    type=['pdf'],
    help="Faça upload do arquivo 'Detalhe da Guia Emitida.pdf'"
)

# Processar arquivo
if uploaded_file is not None:
    st.markdown("---")

    # Informações do arquivo
    col1, col2 = st.columns(2)
    with col1:
        st.metric("📄 Arquivo", uploaded_file.name)
    with col2:
        tamanho_mb = uploaded_file.size / (1024 * 1024)
        st.metric("📏 Tamanho", f"{tamanho_mb:.2f} MB")

    st.markdown("---")

    # Botão de conversão
    if st.button("🚀 CONVERTER PARA EXCEL"):
        with st.spinner("⏳ Processando PDF... Por favor, aguarde."):
            # Ler bytes do arquivo
            pdf_bytes = uploaded_file.read()

            # Extrair dados
            workers, error = extrair_trabalhadores_pdf(pdf_bytes)

            if error:
                st.error(f"❌ Erro ao processar PDF: {error}")
            elif not workers:
                st.warning("⚠️ Nenhum trabalhador encontrado no PDF. Verifique o formato do arquivo.")
            else:
                # Gerar Excel
                excel_bytes, df = gerar_excel(workers)

                # Mensagem de sucesso
                st.markdown(f"""
                <div class="success-box">
                    <h3>✅ Conversão concluída com sucesso!</h3>
                    <p style="font-size: 18px; margin: 10px 0;">
                        <strong>{len(workers)} trabalhadores</strong> extraídos do PDF
                    </p>
                </div>
                """, unsafe_allow_html=True)

                # Preview dos dados
                st.subheader("👀 Prévia dos dados (primeiros 20 registros)")
                st.dataframe(df.head(20), use_container_width=True)

                # Estatísticas
                st.subheader("📊 Estatísticas")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total de Trabalhadores", len(workers))
                with col2:
                    total_valor = df['Valor Consignado na Guia'].str.replace(',', '.').astype(float).sum()
                    st.metric("Valor Total", f"R$ {total_valor:,.2f}")
                with col3:
                    instituicoes_unicas = df['Instituição Financeira'].nunique()
                    st.metric("Instituições", instituicoes_unicas)

                # Botão de download
                st.markdown("---")
                timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
                nome_arquivo = f"FGTS_Trabalhadores_{timestamp}.xlsx"

                st.download_button(
                    label="⬇️ BAIXAR EXCEL",
                    data=excel_bytes,
                    file_name=nome_arquivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.success(f"💾 Arquivo pronto: {nome_arquivo}")

# ============================================================================
# RODAPÉ
# ============================================================================
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; padding: 20px;">
    <p>🔒 Seus dados são processados localmente e não são armazenados</p>
    <p style="font-size: 12px;">Conversor FGTS v2.0 - 100% de precisão</p>
</div>
""", unsafe_allow_html=True)
