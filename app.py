import streamlit as st
import pandas as pd
import qrcode
from PIL import Image, ImageDraw, ImageFont
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
import tempfile

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Gerador de QR Codes - Inventário",
    page_icon="🏥",
    layout="centered"
)

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def gerar_qr_code_bytes(texto, tamanho_qr=120):
    """Gera uma etiqueta com QR Code + TAG escrita embaixo."""

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2,
    )
    qr.add_data(str(texto))
    qr.make(fit=True)

    img_qr = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    img_qr = img_qr.resize((tamanho_qr, tamanho_qr), Image.LANCZOS)

    altura_texto = 30
    largura_final = tamanho_qr
    altura_final = tamanho_qr + altura_texto

    etiqueta = Image.new("RGB", (largura_final, altura_final), "white")
    etiqueta.paste(img_qr, (0, 0))

    draw = ImageDraw.Draw(etiqueta)

    try:
        fonte = ImageFont.truetype("arial.ttf", 14)
    except OSError:
        try:
            fonte = ImageFont.truetype("DejaVuSans.ttf", 14)
        except OSError:
            fonte = ImageFont.load_default()

    texto_str = str(texto)
    bbox = draw.textbbox((0, 0), texto_str, font=fonte)
    largura_texto = bbox[2] - bbox[0]
    x_texto = (largura_final - largura_texto) // 2
    y_texto = tamanho_qr + 5

    draw.text((x_texto, y_texto), texto_str, fill="black", font=fonte)

    buffer = BytesIO()
    etiqueta.save(buffer, format="PNG")
    buffer.seek(0)
    return buffer


def processar_planilha(arquivo_excel, coluna_alvo, nome_aba=None):
    """Lê a planilha, gera QR Codes e retorna o Excel com as etiquetas."""

    df = pd.read_excel(arquivo_excel, sheet_name=nome_aba)

    if coluna_alvo not in df.columns:
        raise ValueError(
            f"A coluna '{coluna_alvo}' não foi encontrada na planilha. "
            f"Colunas disponíveis: {list(df.columns)}"
        )

    arquivo_saida = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    arquivo_saida.close()

    wb = load_workbook(arquivo_excel)
    ws = wb[nome_aba] if nome_aba else wb.active

    coluna_qr = ws.max_column + 1
    letra_qr = get_column_letter(coluna_qr)
    ws.cell(row=1, column=coluna_qr, value="QR Code")

    header_row = [cell.value for cell in ws[1]]
    try:
        indice_coluna_alvo = header_row.index(coluna_alvo) + 1
    except ValueError:
        raise ValueError(f"Coluna '{coluna_alvo}' não encontrada no cabeçalho.")

    ws.column_dimensions[letra_qr].width = 22

    total_linhas = ws.max_row - 1
    barra = st.progress(0, text="Gerando QR Codes...")

    for i, row in enumerate(range(2, ws.max_row + 1), start=1):
        valor = ws.cell(row=row, column=indice_coluna_alvo).value

        if valor is None or str(valor).strip() == "":
            continue

        buffer = gerar_qr_code_bytes(valor, tamanho_qr=120)

        img = XLImage(buffer)
        img.width = 130
        img.height = 155
        ws.add_image(img, f"{letra_qr}{row}")

        ws.row_dimensions[row].height = 115

        progresso = i / total_linhas
        barra.progress(min(progresso, 1.0), text=f"Gerando QR Codes... {i}/{total_linhas}")

    barra.empty()
    wb.save(arquivo_saida.name)

    return arquivo_saida.name, len(df)


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

st.title("🏥 Gerador de QR Codes - Inventário")
st.markdown(
    "Suba a planilha de equipamentos desejada e gere automaticamente "
    "os QR Codes de cada item, com sua TAG para impressão em etiquetas."
)
st.divider()

arquivo_enviado = st.file_uploader(
    "📄 Faça o upload da planilha (.xlsx)",
    type=["xlsx"],
    help="A planilha deve ter a coluna 'TAG' com o código do equipamento."
)

if arquivo_enviado is not None:
    try:
        xls = pd.ExcelFile(arquivo_enviado)
        abas = xls.sheet_names

        if len(abas) == 1:
            aba_selecionada = abas[0]
        else:
            aba_selecionada = st.selectbox("📑 Selecione a aba da planilha:", abas)

        df_preview = pd.read_excel(arquivo_enviado, sheet_name=aba_selecionada)

        st.success(
            f"✅ Planilha carregada: **{len(df_preview)} linhas** "
            f"e **{len(df_preview.columns)} colunas**."
        )
        with st.expander("👀 Ver prévia dos dados"):
            st.dataframe(df_preview.head(10), use_container_width=True)

        colunas = list(df_preview.columns)
        indice_padrao = colunas.index("TAG") if "TAG" in colunas else 0

        coluna_alvo = st.selectbox(
            "🎯 Selecione a coluna que contém o Tombamento/ID:",
            options=colunas,
            index=indice_padrao,
            help="Por padrão, use a coluna 'TAG'. Esta é a informação que será codificada no QR Code."
        )

        if st.button("🚀 Gerar QR Codes", type="primary", use_container_width=True):
            with st.spinner("Processando planilha..."):
                try:
                    caminho_saida, total = processar_planilha(
                        arquivo_enviado,
                        coluna_alvo,
                        nome_aba=aba_selecionada
                    )

                    st.success(f"🎉 {total} QR Codes gerados com sucesso!")

                    with open(caminho_saida, "rb") as f:
                        st.download_button(
                            label="📥 Baixar planilha com QR Codes",
                            data=f,
                            file_name=f"equipamentos_com_qr_{aba_selecionada}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )

                except ValueError as ve:
                    st.error(f"❌ {ve}")
                except Exception as e:
                    st.error(f"❌ Erro ao processar: {e}")

    except Exception as e:
        st.error(f"❌ Não foi possível ler a planilha: {e}")

else:
    st.info("👆 Envie uma planilha `.xlsx` para começar.")

st.divider()
st.caption(
    "🔒 **Privacidade:** os arquivos são processados em memória e não são "
    "armazenados em disco. Nenhum dado é enviado para terceiros."
)