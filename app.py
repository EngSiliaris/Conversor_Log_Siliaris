from datetime import datetime, timedelta, timezone
import io
import struct
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# ==============================================================================
# CONFIGURAÇÕES GERAIS E PERFIS PADRÃO
# ==============================================================================
st.set_page_config(
    page_title="Conversor_Log_Siliaris", page_icon="🏭", layout="wide"
)

# Perfis pré-configurados para as máquinas UD
PROFILES = {
    "UD-2-85-006": [
        "AIT-1 Analisador",
        "FQI-1 Medidor",
        "TI-1 Temp 1º est",
        "TI-2 Temp 2º est",
        "TI-3 Ent AQ",
        "TI-4 Saída AQ",
        "TI-5 Ent. Cond.",
        "TI-6 Saída Cond.",
        "Vazao agua quent",
    ],
    "UD-1-60-005": [
        "AIT-1 Analisador",
        "FQI-1 Medidor",
        "TI-1 Temp 1º est",
        "TI-2 Temp 2º est",
        "TI-3 Ent AQ",
        "TI-4 Saída AQ",
        "TI-5 Ent. Cond.",
        "TI-6 Saída Cond.",
        "Vazao agua quent",
    ],
}


# ==============================================================================
# PROCESSAMENTO BINÁRIO (.fl)
# ==============================================================================
def parse_fl_file(raw_bytes: bytes, num_channels: int, tz_offset_hours: int = -3):
    if not raw_bytes:
        return pd.DataFrame()

    # 1. Inversão dos pares de bytes (correção de endianness 16-bit)
    ba = bytearray(raw_bytes)
    if len(ba) % 2 != 0:
        ba = ba[:-1]

    ba[0::2], ba[1::2] = ba[1::2], ba[0::2]
    swapped = bytes(ba)

    # 2. Split pelo marcador de registro do FieldLogger (\xff\x80\x00\x00)
    marker = b"\xff\x80\x00\x00"
    chunks = swapped.split(marker)

    bytes_per_record = 4 + (num_channels * 4)
    tz = timezone(timedelta(hours=tz_offset_hours))
    records = []
    float_fmt = f"<{num_channels}f"

    for chunk in chunks[1:]:  # Pula o cabeçalho antes do primeiro marcador
        if len(chunk) >= bytes_per_record:
            ts = struct.unpack("<I", chunk[:4])[0]

            # Validação do timestamp Unix (anos entre 2000 e 2050)
            if 946684800 <= ts <= 2524608000:
                dt = datetime.fromtimestamp(ts, tz=tz)
                floats = struct.unpack(
                    float_fmt, chunk[4 : 4 + (num_channels * 4)]
                )
                floats_cleaned = [
                    round(float(v), 3) if abs(v) < 1e7 else 0.0 for v in floats
                ]

                records.append(
                    {
                        "DATETIME": dt,
                        "DATE": dt.strftime("%d/%m/%Y"),
                        "TIME": dt.strftime("%H:%M:%S"),
                        "values": floats_cleaned,
                    }
                )

    if not records:
        return pd.DataFrame()

    data_dict = {
        "DATETIME": [r["DATETIME"] for r in records],
        "DATE": [r["DATE"] for r in records],
        "TIME": [r["TIME"] for r in records],
    }

    for i in range(num_channels):
        data_dict[f"CH_{i}"] = [r["values"][i] for r in records]

    df = pd.DataFrame(data_dict)
    df = df.sort_values("DATETIME").reset_index(drop=True)
    return df


# ==============================================================================
# GERAÇÃO DO EXCEL (.xlsx) COM DADOS NUMÉRICOS
# ==============================================================================
def generate_excel_bytes(df_export: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dados de Processo"

    # Estilização padrão industrial
    header_fill = PatternFill(
        start_color="1F497D", end_color="1F497D", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    border_thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
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
                    row_vals.append(float(val))  # Salva como FLOAT real no Excel
                except (ValueError, TypeError):
                    row_vals.append(val)
        ws.append(row_vals)

    for row in ws.iter_rows(
        min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column
    ):
        for col_idx, cell in enumerate(row, start=1):
            cell.font = data_font
            cell.border = border_thin
            col_name = headers[col_idx - 1]
            if col_name in ["DATE", "TIME"]:
                cell.alignment = Alignment(horizontal="center")
            else:
                cell.alignment = Alignment(horizontal="right")
                cell.number_format = (
                    "#,##0.000"  # Padrão com 3 casas decimais
                )

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output)
    return output.getvalue()


# ==============================================================================
# INTERFACE PRINCIPAL
# ==============================================================================
def main():
    st.title("🏭 Conversor_Log_Siliaris")
    st.markdown(
        "Extração direta de dados do arquivo **MEMFLASH.fl** para planilha Excel (.xlsx)."
    )

    # 1. Seleção do equipamento
    selected_machine = st.selectbox(
        "Selecione o Equipamento:", list(PROFILES.keys())
    )
    active_channels = PROFILES[selected_machine]

    # 2. Upload do arquivo binário
    uploaded_file = st.file_uploader(
        "Arraste e solte o arquivo MEMFLASH.fl aqui:",
        type=["fl", "bin"],
    )

    if uploaded_file is not None:
        raw_bytes = uploaded_file.read()

        with st.spinner("Processando e decodificando o arquivo binário..."):
            df_raw = parse_fl_file(
                raw_bytes, len(active_channels), tz_offset_hours=-3
            )

        if df_raw.empty:
            st.error(
                "❌ Nenhum dado válido encontrado. Verifique se o arquivo MEMFLASH.fl está íntegro."
            )
            return

        rename_map = {
            f"CH_{i}": active_channels[i] for i in range(len(active_channels))
        }
        df_raw = df_raw.rename(columns=rename_map)

        min_dt = df_raw["DATETIME"].min()
        max_dt = df_raw["DATETIME"].max()

        st.divider()
        st.subheader("📅 Filtro de Período")

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            sel_start_date = st.date_input(
                "Data Inicial do Log:",
                value=min_dt.date(),
                min_value=min_dt.date(),
                max_value=max_dt.date(),
            )
        with c2:
            sel_start_time = st.time_input("Hora Inicial:", value=min_dt.time())
        with c3:
            st.date_input(
                "Data Final (automática):",
                value=max_dt.date(),
                disabled=True,
            )
        with c4:
            st.time_input(
                "Hora Final (automática):",
                value=max_dt.time(),
                disabled=True,
            )

        start_datetime = datetime.combine(
            sel_start_date, sel_start_time
        ).replace(tzinfo=min_dt.tzinfo)

        df_filtered = df_raw[
            (df_raw["DATETIME"] >= start_datetime)
            & (df_raw["DATETIME"] <= max_dt)
        ].copy()
        display_columns = ["DATE", "TIME"] + active_channels
        df_final = df_filtered[display_columns].reset_index(drop=True)

        if df_final.empty:
            st.warning(
                "⚠️ Nenhum registro encontrado para esse horário. Escolha um horário inicial anterior."
            )
            return

        st.success(
            f"✅ **{len(df_final):,} registros** prontos ({df_final['DATE'].iloc[0]} às {df_final['TIME'].iloc[0]} até {df_final['DATE'].iloc[-1]} às {df_final['TIME'].iloc[-1]})."
        )

        st.dataframe(df_final.head(50), use_container_width=True)

        # Nome automático do Excel no seu padrão (Ex: UD85 05.11.26.xlsx)
        tag_maquina = "UD85" if "85" in selected_machine else "UD60"
        data_str = sel_start_date.strftime("%d.%m.%y")
        filename_excel = f"{tag_maquina} {data_str}.xlsx"

        excel_bytes = generate_excel_bytes(df_final)

        st.download_button(
            label=f"📥 Baixar Arquivo Excel ({filename_excel})",
            data=excel_bytes,
            file_name=filename_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )


if __name__ == "__main__":
    main()