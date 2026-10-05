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
    page_title="Gerador de QR Codes - Inventário Hospitalar",
    page_icon="🏥",
    layout="centered"
)

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def gerar_qr_code_bytes(tag, equipamento, serie, tamanho_qr=120):
    """Gera uma etiqueta com QR Code (contendo TAG + Equipamento + Série)
    e, embaixo, TAG e Série visíveis."""

    # 1. Conteúdo do QR Code — 3 informações
    conteudo_qr = f"TAG: {tag}\nEQUIPAMENTO: {equipamento}\nSERIE: {serie}"

    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=2,
    )
    qr.add_data(conteudo_qr)
    qr.make(fit=True)

    img_qr = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    img_qr = img_qr.resize((tamanho_qr, tamanho_qr), Image.LANCZOS)

    # 2. Espaço embaixo para o texto
    altura_texto = 40
    largura_final = tamanho_qr
    altura_final = tamanho_qr + altura_texto

    etiqueta = Image.new("RGB", (largura_final, altura_final), "white")
    etiqueta.paste(img_qr, (0, 0))

    # 3. Escreve TAG e Série embaixo
    draw = ImageDraw.Draw(etiqueta)

    try:
        fonte = ImageFont.truetype("arial.ttf", 12)
    except OSError:
        try:
            fonte = ImageFont.truetype("DejaVuSans.ttf", 12)
        except OSError:
            fonte = ImageFont.load_default()

    # Linha 1: TAG
    tag_str = str(tag) if tag else "—"
    bbox_tag = draw.textbbox((0, 0), tag_str, font=fonte)
    x_tag = (largura_final - (bbox_tag[2] - bbox_tag[0])) // 2
    draw.text((x_tag, tamanho_qr + 4), tag_str, fill="black", font=fonte)

    # Linha 2: Série
    serie_str = str(serie) if serie else "—"
    bbox_serie = draw.textbbox((0, 0), serie_str, font=fonte)
    x_serie = (largura_final - (bbox_serie[2] - bbox_serie[0])) // 2
    draw.text((x_serie, tamanho_qr + 20), serie_str, fill="black", font=fonte)

    # 4. Retorna como BytesIO
    buffer = BytesIO()
    etiqueta.save(buffer, format="PNG")
    buffer.seek(0)
    return buffer


def processar_planilha(arquivo_excel, coluna_tag, coluna_equip, coluna_serie, nome_aba=None):
    """Lê a planilha, gera QR Codes com TAG + Equipamento + Série
    e retorna o Excel com as etiquetas inseridas."""

    df = pd.read_excel(arquivo_excel, sheet_name=nome_aba)

    # Validação das colunas
    for col, nome in [(coluna_tag, "TAG"), (coluna_equip, "Equipamento"), (coluna_serie, "Nº Série")]:
        if col not in df.columns:
            raise ValueError(
                f"A coluna '{col}' não foi encontrada. "
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
        idx_tag = header_row.index(coluna_tag) + 1
        idx_equip = header_row.index(coluna_equip) + 1
        idx_serie = header_row.index(coluna_serie) + 1
    except ValueError as e:
        raise ValueError(f"Erro localizando colunas no cabeçalho: {e}")

    ws.column_dimensions[letra_qr].width = 22

    total_linhas = ws.max_row - 1
    barra = st.progress(0, text="Gerando QR Codes...")

    for i, row in enumerate(range(2, ws.max_row + 1), start=1):
        tag = ws.cell(row=row, column=idx_tag).value
        equip = ws.cell(row=row, column=idx_equip).value
        serie = ws.cell(row=row, column=idx_serie).value

        if tag is None or str(tag).strip() == "":
            continue

        buffer = gerar_qr_code_bytes(tag, equip or "—", serie or "—", tamanho_qr=120)

        img = XLImage(buffer)
        img.width = 130
        img.height = 165
        ws.add_image(img, f"{letra_qr}{row}")

        ws.row_dimensions[row].height = 125

        progresso = i / total_linhas
        barra.progress(min(progresso, 1.0), text=f"Gerando QR Codes... {i}/{total_linhas}")

    barra.empty()
    wb.save(arquivo_saida.name)

    return arquivo_saida.name, len(df)


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

st.title("🏥 Gerador de QR Codes - Inventário Hospitalar")
st.markdown(
    "Suba a planilha de equipamentos do hospital e gere automaticamente "
    "os QR Codes de cada item, prontos para impressão em etiquetas."
)
st.divider()

arquivo_enviado = st.file_uploader(
    "📄 Faça o upload da planilha (.xlsx)",
    type=["xlsx"],
    help="A planilha deve conter as colunas: TAG, Equipamento e Nº Série."
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

        # --- Seleção das 3 colunas ---
        colunas = list(df_preview.columns)

        idx_tag = colunas.index("TAG") if "TAG" in colunas else 0
        coluna_tag = st.selectbox(
            "🎯 Coluna da TAG (vai dentro do QR):",
            options=colunas,
            index=idx_tag
        )

        idx_equip = colunas.index("Equipamento") if "Equipamento" in colunas else 0
        coluna_equip = st.selectbox(
            "🎯 Coluna do Equipamento (vai dentro do QR):",
            options=colunas,
            index=idx_equip
        )

        idx_serie = colunas.index("Nº Série") if "Nº Série" in colunas else 0
        coluna_serie = st.selectbox(
            "🎯 Coluna do Nº Série (vai dentro do QR e na etiqueta):",
            options=colunas,
            index=idx_serie
        )

        if st.button("🚀 Gerar QR Codes", type="primary", use_container_width=True):
            with st.spinner("Processando planilha..."):
                try:
                    caminho_saida, total = processar_planilha(
                        arquivo_enviado,
                        coluna_tag,
                        coluna_equip,
                        coluna_serie,
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