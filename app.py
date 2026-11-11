from datetime import datetime, time, timedelta, timezone
import io
import re
import struct
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Conversor_Log_Siliaris", page_icon="🏭", layout="wide"
)

# Colunas oficiais da planilha Siliaris
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


def detect_file_end_timestamp(raw_bytes: bytes) -> datetime:
    """Varre a memória do arquivo para detectar o carimbo final da coleta."""
    valid_ts = []
    MIN_TS = 1704067200  # 2024-01-01
    MAX_TS = 1893456000  # 2030-01-01
    total_len = len(raw_bytes)

    for i in range(0, min(total_len - 4, 300000), 2):
        chunk = raw_bytes[i : i + 4]
        val_le = struct.unpack("<I", chunk)[0]
        if MIN_TS <= val_le <= MAX_TS:
            valid_ts.append(val_le)
        val_be = struct.unpack(">I", chunk)[0]
        if MIN_TS <= val_be <= MAX_TS:
            valid_ts.append(val_be)

    if valid_ts:
        ts_final = max(valid_ts)
        return datetime.fromtimestamp(
            ts_final, tz=timezone(timedelta(hours=-3))
        ).replace(tzinfo=None)

    # Fallback seguro para data/hora atual
    return datetime.now().replace(microsecond=0)


def parse_fieldlogger_memflash(raw_bytes: bytes):
    """Extrai os 9 floats limpos de cada bloco de 44 bytes e desembaralha o buffer circular."""
    marker = b"\x80\xff\x00\x00"
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

            # Filtra blocos corrompidos ou em branco
            if 10.0 <= floats[0] <= 120.0 and abs(floats[8]) < 500.0:
                ti_1, ti_2, ti_3, ti_4, ti_5, ti_6, ait_1, fqi_1, vazao = floats
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
                tags.append(tag_val)

        pos = idx + 44

    if not records:
        return []

    # Se a memória circular deu a volta (wrap around), desembaralha
    wrap_idx = -1
    for i in range(len(tags) - 1):
        if tags[i] > tags[i + 1]:
            wrap_idx = i + 1
            break

    if wrap_idx != -1:
        records = records[wrap_idx:] + records[:wrap_idx]

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
                    row_vals.append(float(val))  # Salva número real no Excel
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
        "Conversão direta de **MEMFLASH.fl** para planilha Excel (.xlsx) mantendo a sincronia de segundos e formatação oficial."
    )

    col_m1, col_m2 = st.columns([1, 2])
    with col_m1:
        machine = st.selectbox(
            "Selecione o Equipamento:",
            ["UD-2-85-006", "UD-1-60-005", "UO-1-240-009"],
        )

    tag_maquina = "UD85" if "85" in machine else "UD60"

    uploaded_file = st.file_uploader(
        "Arraste e solte o arquivo MEMFLASH.fl aqui:",
        type=["fl", "bin"],
        key="main_file",
    )

    if uploaded_file is not None:
        raw_bytes = uploaded_file.read()

        with st.spinner("Processando blocos binários de 44-bytes..."):
            records = parse_fieldlogger_memflash(raw_bytes)
            auto_end_dt = detect_file_end_timestamp(raw_bytes)

        if not records:
            st.error("❌ Nenhum dado válido encontrado no arquivo MEMFLASH.fl.")
            return

        total_pts = len(records)

        # Monta a linha temporal preservando os segundos reais do encerramento
        timestamps = [
            auto_end_dt - timedelta(minutes=(total_pts - 1 - i))
            for i in range(total_pts)
        ]

        df_all = pd.DataFrame(records)
        df_all.insert(0, "TIME", [t.strftime("%H:%M:%S") for t in timestamps])
        df_all.insert(0, "DATE", [t.strftime("%d/%m/%Y") for t in timestamps])
        df_all.insert(0, "_DT", timestamps)

        st.divider()
        st.subheader("📅 Período do Log")

        # Exibição da Data Final Automática com segundos
        st.info(
            f"🔒 **Encerramento da Coleta (Automático do Arquivo):** "
            f"**{auto_end_dt.strftime('%d/%m/%Y')} às {auto_end_dt.strftime('%H:%M:%S')}**"
        )

        col_ini1, col_ini2, col_nome = st.columns([1.2, 1.2, 1.6])

        with col_ini1:
            start_date = st.date_input(
                "Data Inicial do Período:",
                value=(auto_end_dt - timedelta(days=1)).date(),
                help="Escolha o dia em que você quer que o novo log comece.",
            )

        with col_ini2:
            # Entrada de hora com segundos via texto (permite precisão máxima: HH:MM:SS)
            sugestao_hora = "01:05:00"
            hora_digitada = st.text_input(
                "Hora Inicial (HH:MM:SS):",
                value=sugestao_hora,
                help="Digite a hora de corte (ex: 01:05:00 ou 01:05:48). O sistema busca a leitura correspondente a partir desse momento.",
            )

        # Trata a hora digitada pelo usuário
        try:
            partes = [int(p) for p in hora_digitada.strip().split(":")]
            if len(partes) == 2:
                user_time = time(partes[0], partes[1], 0)
            elif len(partes) == 3:
                user_time = time(partes[0], partes[1], partes[2])
            else:
                user_time = time(1, 5, 0)
        except Exception:
            user_time = time(1, 5, 0)

        with col_nome:
            sugestao_data_nome = start_date.strftime("%d.%m.%y")
            custom_data_str = st.text_input(
                "Data no Nome do Arquivo (DD.MM.AA):",
                value=sugestao_data_nome,
                help="Você pode alterar para o dia de referência do turno (ex: 11.11.26 ou 10.11.26).",
            )

        start_dt = datetime.combine(start_date, user_time)

        # Filtro inteligente: pega a partir do momento solicitado
        df_filtered = df_all[
            (df_all["_DT"] >= start_dt) & (df_all["_DT"] <= auto_end_dt)
        ].copy()
        df_final = df_filtered.drop(columns=["_DT"]).reset_index(drop=True)

        if df_final.empty:
            st.warning(
                "⚠️ Nenhuma leitura encontrada para esse horário. Verifique a data/hora inicial."
            )
            return

        st.success(
            f"✅ **{len(df_final):,} leituras prontas** — Primeira linha: **{df_final['DATE'].iloc[0]} às {df_final['TIME'].iloc[0]}** | Última linha: **{df_final['DATE'].iloc[-1]} às {df_final['TIME'].iloc[-1]}**"
        )

        st.dataframe(df_final.head(100), use_container_width=True)

        nome_limpo = custom_data_str.strip() or sugestao_data_nome
        filename_excel = f"{tag_maquina} {nome_limpo}.xlsx"

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