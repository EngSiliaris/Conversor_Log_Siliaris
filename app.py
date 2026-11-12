import io
import struct
from datetime import datetime, time, timedelta
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Conversor_Log_Siliaris",
    page_icon="🏭",
    layout="wide"
)

# Colunas na ordem oficial da planilha
COLUMNS_ORDER = [
    "AIT-1 Analisador",
    "FQI-1 Medidor",
    "TI-1 Temp 1º est",
    "TI-2 Temp 2º est",
    "TI-3 Ent AQ",
    "TI-4 Saída AQ",
    "TI-5 Ent. Cond.",
    "TI-6 Saída Cond.",
    "Vazao agua quent"
]


def parse_fieldlogger_memflash(raw_bytes: bytes):
    """
    Decodifica os blocos de 44 bytes puros e desembaralha o buffer circular (ring buffer).
    Lógica 100% estrutural, sem adivinhação de datas nos bytes.
    """
    marker = b'\x80\xff\x00\x00'
    records = []
    tags = []
    pos = 0
    total_len = len(raw_bytes)

    while True:
        idx = raw_bytes.find(marker, pos)
        if idx == -1 or idx + 44 > total_len:
            break

        chunk = raw_bytes[idx : idx + 44]
        tag_val = struct.unpack("<H", chunk[4:6])[0]
        float_bytes = chunk[6:42]

        if len(float_bytes) == 36:
            floats = struct.unpack("<9f", float_bytes)
            # Filtro para ignorar blocos vazios ou sujos da memória
            if 10.0 <= floats[0] <= 120.0 and abs(floats[8]) < 500.0:
                ti_1, ti_2, ti_3, ti_4, ti_5, ti_6, ait_1, fqi_1, vazao = floats
                records.append({
                    "AIT-1 Analisador": round(float(ait_1), 3),
                    "FQI-1 Medidor": round(float(fqi_1), 3),
                    "TI-1 Temp 1º est": round(float(ti_1), 3),
                    "TI-2 Temp 2º est": round(float(ti_2), 3),
                    "TI-3 Ent AQ": round(float(ti_3), 3),
                    "TI-4 Saída AQ": round(float(ti_4), 3),
                    "TI-5 Ent. Cond.": round(float(ti_5), 3),
                    "TI-6 Saída Cond.": round(float(ti_6), 3),
                    "Vazao agua quent": round(float(vazao), 3)
                })
                tags.append(tag_val)

        pos = idx + 44

    if not records:
        return []

    # Desembaralha a memória circular (caso tenha dado a volta)
    wrap_idx = -1
    for i in range(len(tags) - 1):
        if tags[i] > tags[i + 1] + 100:  
            wrap_idx = i + 1
            break

    if wrap_idx != -1:
        records = records[wrap_idx:] + records[:wrap_idx]

    return records


def build_strict_timeline(records, end_dt: datetime):
    """
    Constrói a linha do tempo com matemática exata.
    O último registro do arquivo recebe a Data e Hora Final definida pelo usuário.
    Os registros anteriores são recuados exatamente 1 minuto cada.
    """
    total = len(records)
    if total == 0:
        return pd.DataFrame()

    # Calcula de trás para frente (1 minuto por linha)
    timestamps = [end_dt - timedelta(minutes=(total - 1 - i)) for i in range(total)]

    df = pd.DataFrame(records)
    df.insert(0, "TIME", [t.strftime("%H:%M:%S") for t in timestamps])
    df.insert(0, "DATE", [t.strftime("%d/%m/%Y") for t in timestamps])
    df.insert(0, "_DT", timestamps)
    return df


def generate_excel_bytes(df_export: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dados de Processo"

    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    headers = list(df_export.columns)
    ws.append(headers)

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for _, row in df_export.iterrows():
        row_vals = []
        for col in headers:
            val = row[col]
            if col in ["DATE", "TIME"]:
                row_vals.append(str(val))
            else:
                try:
                    row_vals.append(float(val))
                except (ValueError, TypeError):
                    row_vals.append(val)
        ws.append(row_vals)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for col_idx, cell in enumerate(row, start=1):
            cell.font = data_font
            cell.border = border_thin
            col_name = headers[col_idx - 1]
            if col_name in ["DATE", "TIME"]:
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = '#,##0.000'

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 13)

    wb.save(output)
    return output.getvalue()


def main():
    st.title("🏭 Conversor_Log_Siliaris")
    st.markdown("Extração Matemática Exata. **Total precisão cronológica e fidelidade aos dados numéricos.**")

    col_m1, col_m2 = st.columns([1, 2])
    with col_m1:
        machine = st.selectbox(
            "Selecione o Equipamento:",
            ["UD-2-85-006", "UD-1-60-005", "UO-1-240-009"]
        )

    tag_maquina = "UD85" if "85" in machine else ("UD60" if "60" in machine else "UO240")

    uploaded_file = st.file_uploader(
        "Arraste e solte o arquivo MEMFLASH.fl aqui:",
        type=["fl", "bin"],
        key="main_file"
    )

    if uploaded_file is not None:
        raw_bytes = uploaded_file.read()

        with st.spinner("Lendo blocos de memória e organizando sensores..."):
            records = parse_fieldlogger_memflash(raw_bytes)

        if not records:
            st.error("❌ Nenhum dado válido encontrado no arquivo MEMFLASH.fl.")
            return
            
        st.divider()

        # ----------------------------------------------------------------------
        # PASSO 1: ÂNCORA FINAL ABSOLUTA (FIM DA COLETA)
        # ----------------------------------------------------------------------
        st.subheader("1️⃣ Sincronização de Data/Hora (Retirada do Pendrive)")
        st.warning(
            "**COMO FUNCIONA:** O software Novus original utiliza a Data de Modificação do Windows para saber quando o arquivo terminou. "
            "Como o navegador bloqueia essa informação, **informe abaixo a Data e Hora exata em que o arquivo terminou (ou em que o pendrive foi retirado).**\n\n"
            "*Dica: Olhe as Propriedades do arquivo no Windows se tiver dúvida.*"
        )

        col_fim1, col_fim2 = st.columns(2)
        with col_fim1:
            end_date = st.date_input(
                "Data Final da Coleta:",
                value=datetime(2026, 11, 10).date()
            )
        with col_fim2:
            end_time_str = st.text_input(
                "Hora Final da Coleta (HH:MM:SS):",
                value="22:52:56",
                help="Hora e segundo exatos em que a coleta encerrou no arquivo."
            )

        try:
            p_f = [int(p) for p in end_time_str.strip().split(":")]
            end_t = time(p_f[0], p_f[1], p_f[2] if len(p_f) > 2 else 0)
        except:
            end_t = time(22, 52, 56)

        end_dt = datetime.combine(end_date, end_t)

        # Monta a linha temporal cravada 1 minuto para trás a partir da âncora final
        df_all = build_strict_timeline(records, end_dt)

        st.divider()

        # ----------------------------------------------------------------------
        # PASSO 2: FILTRO DE CORTE E NOMEAÇÃO
        # ----------------------------------------------------------------------
        st.subheader("2️⃣ Filtro de Corte (Onde iniciar a planilha)")
        st.markdown("Informe a partir de que momento você quer extrair os dados. Eles casarão perfeitamente com sua última planilha.")

        col_ini1, col_ini2, col_nome = st.columns([1.2, 1.2, 1.6])

        with col_ini1:
            start_date = st.date_input(
                "Data Inicial do Corte:",
                value=datetime(2026, 11, 6).date()
            )

        with col_ini2:
            start_time_str = st.text_input(
                "Hora Inicial do Corte (HH:MM:SS):",
                value="01:05:42",
                help="Digite o tempo exato com os segundos para corte (ex: 01:05:42)."
            )

        with col_nome:
            sugestao_nome = start_date.strftime("%d.%m.%y")
            custom_data_str = st.text_input(
                "Data no Nome do Arquivo (DD.MM.AA):",
                value=sugestao_nome
            )

        try:
            p_i = [int(p) for p in start_time_str.strip().split(":")]
            start_t = time(p_i[0], p_i[1], p_i[2] if len(p_i) > 2 else 0)
        except:
            start_t = time(1, 5, 42)

        start_dt = datetime.combine(start_date, start_t)

        # O Filtro Mágico: Pega tudo do momento do corte até o fim!
        df_filtered = df_all[df_all["_DT"] >= start_dt].copy()

        if df_filtered.empty:
            st.warning("⚠️ O horário inicial solicitado é maior que o horário final. Verifique as datas informadas.")
            return

        df_final = df_filtered.drop(columns=["_DT"]).reset_index(drop=True)

        st.success(
            f"🎯 **Exatidão Atingida:** {len(df_final):,} registros compilados.\n\n"
            f"A primeira linha é pontualmente **{df_final['DATE'].iloc[0]} às {df_final['TIME'].iloc[0]}**. Pode conferir com o seu padrão!"
        )

        st.dataframe(df_final.head(100), use_container_width=True)

        nome_limpo = custom_data_str.strip() or sugestao_nome
        filename_excel = f"{tag_maquina} {nome_limpo}.xlsx"

        excel_bytes = generate_excel_bytes(df_final)

        st.download_button(
            label=f"📥 Baixar Arquivo Excel ({filename_excel})",
            data=excel_bytes,
            file_name=filename_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

if __name__ == "__main__":
    main()