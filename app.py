from datetime import datetime, timedelta, timezone
import io
import struct
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Conversor_Log_Siliaris", page_icon="🏭", layout="wide"
)

# Colunas na ordem oficial da planilha da Siliaris
COLUMNS_ORDER = [
    "AIT-1 Analisador",
    "FQI-1 Medidor",
    "TI-1 Temp 1º est",
    "TI-2 Temp 2º est",
    "TI-3 Ent AQ",
    "TI-4 Saída AQ",
    "TI-5 Ent. Cond.",
    "TI-6 Saída Cond.",
    "Vazao agua quent",
]


def parse_fieldlogger_memflash(raw_bytes: bytes):
    """Decodifica os blocos de 44 bytes do MEMFLASH.fl da Novus:

    [4 bytes marcador: \x80\xff\x00\x00]
    [2 bytes tag/status]
    [36 bytes: 9 floats IEEE 754 little-endian]
    [2 bytes CRC16]
    """
    marker = b"\x80\xff\x00\x00"
    records = []

    pos = 0
    total_len = len(raw_bytes)

    while True:
        idx = raw_bytes.find(marker, pos)
        if idx == -1 or idx + 44 > total_len:
            break

        # Bloco de dados de 44 bytes
        chunk = raw_bytes[idx : idx + 44]

        # Os 9 floats começam no byte 6 e vão até o byte 42 (36 bytes)
        float_bytes = chunk[6:42]
        if len(float_bytes) == 36:
            floats = struct.unpack("<9f", float_bytes)

            # Validação física: descarta blocos corrompidos ou em branco
            # TI-1 geralmente está entre 20°C e 100°C
            if 10.0 <= floats[0] <= 120.0 and abs(floats[8]) < 500.0:
                # Ordem interna gravada no FieldLogger:
                # 0: TI-1, 1: TI-2, 2: TI-3, 3: TI-4, 4: TI-5, 5: TI-6, 6: AIT-1, 7: FQI-1, 8: Vazao
                ti_1, ti_2, ti_3, ti_4, ti_5, ti_6, ait_1, fqi_1, vazao = floats

                # Remapeia para a ordem visual exata do Excel modelo da Siliaris
                records.append(
                    {
                        "AIT-1 Analisador": round(float(ait_1), 3),
                        "FQI-1 Medidor": round(float(fqi_1), 3),
                        "TI-1 Temp 1º est": round(float(ti_1), 3),
                        "TI-2 Temp 2º est": round(float(ti_2), 3),
                        "TI-3 Ent AQ": round(float(ti_3), 3),
                        "TI-4 Saída AQ": round(float(ti_4), 3),
                        "TI-5 Ent. Cond.": round(float(ti_5), 3),
                        "TI-6 Saída Cond.": round(float(ti_6), 3),
                        "Vazao agua quent": round(float(vazao), 3),
                    }
                )

        pos = idx + 44

    return records


def generate_excel_bytes(df_export: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dados de Processo"

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
                    row_vals.append(float(val))
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
                cell.number_format = "#,##0.000"

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 13)

    wb.save(output)
    return output.getvalue()


def main():
    st.title("🏭 Conversor_Log_Siliaris")
    st.markdown(
        "Conversão direta de **MEMFLASH.fl** para planilha Excel (.xlsx) nos padrões de coleta da fábrica."
    )

    col_m1, col_m2 = st.columns([1, 2])
    with col_m1:
        machine = st.selectbox(
            "Equipamento:", ["UD-2-85-006", "UD-1-60-005", "UO-1-240-009"]
        )

    uploaded_file = st.file_uploader(
        "Arraste e solte o arquivo MEMFLASH.fl aqui:",
        type=["fl", "bin"],
        key="main_file",
    )

    if uploaded_file is not None:
        raw_bytes = uploaded_file.read()

        with st.spinner("Processando blocos binários de 44-bytes..."):
            records = parse_fieldlogger_memflash(raw_bytes)

        if not records:
            st.error(
                "❌ Não foi possível extrair dados válidos. O arquivo pode estar vazio ou corrompido."
            )
            return

        total_pts = len(records)

        st.success(
            f"✅ **{total_pts:,} leituras válidas de 1 minuto** encontradas na memória do FieldLogger."
        )

        st.divider()
        st.subheader("📅 Período do Log")

        # Configuração do intervalo de tempo
        # Por padrão, assume que o log finaliza na data/hora atual ou do dia anterior
        hoje = datetime.now()

        col_d1, col_d2, col_d3, col_d4 = st.columns(4)

        with col_d3:
            end_date = st.date_input(
                "Data Final do Arquivo:",
                value=hoje.date(),
                help="Data em que o arquivo foi coletado no FieldLogger.",
            )
        with col_d4:
            end_time = st.time_input(
                "Hora Final do Arquivo:",
                value=datetime.strptime("01:05:48", "%H:%M:%S").time(),
                help="Hora em que a coleta encerrou.",
            )

        end_dt = datetime.combine(end_date, end_time)

        # Calcula a linha do tempo retroativa (1 minuto por leitura)
        timestamps = [
            end_dt - timedelta(minutes=(total_pts - 1 - i))
            for i in range(total_pts)
        ]

        df_all = pd.DataFrame(records)
        df_all.insert(0, "TIME", [t.strftime("%H:%M:%S") for t in timestamps])
        df_all.insert(0, "DATE", [t.strftime("%d/%m/%Y") for t in timestamps])
        df_all.insert(0, "_DT", timestamps)

        with col_d1:
            start_date = st.date_input(
                "Data Inicial Desejada:",
                value=(end_dt - timedelta(days=1)).date(),
                help="Escolha o dia de início da sua leitura diária.",
            )
        with col_d2:
            start_time = st.time_input(
                "Hora Inicial Desejada:",
                value=datetime.strptime("01:05:00", "%H:%M:%S").time(),
                help="Hora de corte inicial (ex: 01:05:00 para repetir a última linha anterior).",
            )

        start_dt = datetime.combine(start_date, start_time)

        # Filtra o período selecionado
        df_filtered = df_all[
            (df_all["_DT"] >= start_dt) & (df_all["_DT"] <= end_dt)
        ].copy()
        df_final = df_filtered.drop(columns=["_DT"]).reset_index(drop=True)

        st.caption(
            f"Mostrando **{len(df_final):,} registros** entre **{start_dt.strftime('%d/%m/%Y %H:%M:%S')}** e **{end_dt.strftime('%d/%m/%Y %H:%M:%S')}**."
        )

        st.dataframe(df_final.head(100), use_container_width=True)

        # Nome automático idêntico ao modelo (ex: UD85 06.11.26.xlsx)
        tag_maquina = "UD85" if "85" in machine else "UD60"
        data_str = start_date.strftime("%d.%m.%y")
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