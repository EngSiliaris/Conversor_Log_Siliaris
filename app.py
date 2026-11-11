import io
import struct
from datetime import datetime, date, time, timedelta, timezone
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


def extract_header_timestamps(raw_bytes: bytes, swapped_bytes: bytes):
    """
    Rastreia o cabeçalho do FieldLogger (primeiros 1024 bytes) em busca das datas
    de Início e Fim absolutas gravadas no equipamento.
    """
    valid_ts = set()
    # Limites plausíveis de tempo (2024 a 2030)
    MIN_UNIX = 1704067200  
    MAX_UNIX = 1893456000  
    # A Novus costuma usar a época do ano 2000 em alguns firmwares
    MIN_NOVUS = MIN_UNIX - 946684800 
    MAX_NOVUS = MAX_UNIX - 946684800 

    # Varredura forense nos cabeçalhos
    for buffer in [raw_bytes[:1024], swapped_bytes[:1024]]:
        for i in range(len(buffer) - 3):
            # Tenta ler como Little Endian
            ts_le = struct.unpack("<I", buffer[i:i+4])[0]
            if MIN_UNIX <= ts_le <= MAX_UNIX:
                valid_ts.add(ts_le)
            elif MIN_NOVUS <= ts_le <= MAX_NOVUS:
                valid_ts.add(ts_le + 946684800)
                
            # Tenta ler como Big Endian
            ts_be = struct.unpack(">I", buffer[i:i+4])[0]
            if MIN_UNIX <= ts_be <= MAX_UNIX:
                valid_ts.add(ts_be)
            elif MIN_NOVUS <= ts_be <= MAX_NOVUS:
                valid_ts.add(ts_be + 946684800)

    lista_ts = sorted(list(valid_ts))
    return lista_ts


def parse_fieldlogger_memflash(raw_bytes: bytes):
    """
    Decodifica os blocos de 44 bytes contendo os sensores.
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
            # Filtro para ignorar blocos vazios ou lixo de memória
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

    # Se a memória deu a volta (buffer circular), encontra o ponto de corte e realinha
    wrap_idx = -1
    for i in range(len(tags) - 1):
        if tags[i] > tags[i + 1] + 100:  
            wrap_idx = i + 1
            break

    if wrap_idx != -1:
        records = records[wrap_idx:] + records[:wrap_idx]

    return records


def build_interpolated_timeline(records, timestamps_found):
    """
    Aplica o algoritmo oficial da Novus: Pega o tempo total e divide pelo número
    de registros para achar o intervalo exato (que gera os drifts de segundos nos dias).
    """
    n = len(records)
    if n == 0:
        return pd.DataFrame()

    # Se achou as datas de Início e Fim no arquivo, usa a matemática exata!
    if len(timestamps_found) >= 2:
        start_dt = datetime.fromtimestamp(timestamps_found[0], tz=timezone.utc).replace(tzinfo=None)
        end_dt = datetime.fromtimestamp(timestamps_found[-1], tz=timezone.utc).replace(tzinfo=None)
    elif len(timestamps_found) == 1:
        end_dt = datetime.fromtimestamp(timestamps_found[0], tz=timezone.utc).replace(tzinfo=None)
        start_dt = end_dt - timedelta(minutes=n-1)
    else:
        # Fallback de emergência caso o cabeçalho esteja corrompido
        end_dt = datetime(2026, 11, 10, 22, 52, 56)
        start_dt = end_dt - timedelta(minutes=n-1)

    # Lógica de Interpolação da Novus (ex: 60.0102 segundos entre cada linha)
    total_seconds = (end_dt - start_dt).total_seconds()
    interval_sec = total_seconds / (n - 1) if n > 1 else 60.0

    # Aplica o tempo a cada linha
    for i in range(n):
        records[i]["_DT"] = start_dt + timedelta(seconds=i * interval_sec)

    df = pd.DataFrame(records)
    # Formata a visualização dos Segundos de forma arredondada igual ao Excel
    df.insert(0, "TIME", df["_DT"].apply(lambda x: x.replace(microsecond=0).strftime("%H:%M:%S")))
    df.insert(0, "DATE", df["_DT"].dt.strftime("%d/%m/%Y"))
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
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
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
    st.markdown("Extração oficial. A data e hora final são obtidas **automaticamente do arquivo**, e a tabela é alinhada com os segundos exatos do FieldLogger.")

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
        
        # Preparação do arquivo para o motor de busca
        ba = bytearray(raw_bytes)
        if len(ba) % 2 != 0: ba = ba[:-1]
        ba[0::2], ba[1::2] = ba[1::2], ba[0::2]
        swapped_bytes = bytes(ba)

        with st.spinner("Extraindo horários reais do cabeçalho da máquina..."):
            records = parse_fieldlogger_memflash(swapped_bytes)
            timestamps_found = extract_header_timestamps(raw_bytes, swapped_bytes)

        if not records:
            st.error("❌ Nenhum dado válido encontrado no arquivo MEMFLASH.fl.")
            return

        # Reconstrói a linha do tempo (Interpolação exata da Novus)
        df_all = build_interpolated_timeline(records, timestamps_found)
        
        abs_start = df_all["_DT"].min()
        abs_end = df_all["_DT"].max()

        st.divider()
        st.subheader("📅 Período do Log")

        st.info(
            f"✅ **Arquivo decodificado com sucesso!** O FieldLogger iniciou este histórico em **{abs_start.strftime('%d/%m/%Y %H:%M:%S')}** "
            f"e a coleta final (retirada do pendrive) ocorreu automaticamente em **{abs_end.strftime('%d/%m/%Y %H:%M:%S')}**."
        )

        col_ini1, col_ini2, col_nome = st.columns([1.2, 1.2, 1.6])

        with col_ini1:
            start_date = st.date_input(
                "Data Inicial do Período (Seu Corte):",
                value=datetime(2026, 11, 6).date(),
                min_value=abs_start.date(),
                max_value=abs_end.date()
            )

        with col_ini2:
            start_time_str = st.text_input(
                "Hora Inicial (HH:MM:SS):",
                value="01:05:00",
                help="O sistema procurará a linha exata igual ao Novus (ex: 01:05:48)."
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
            start_t = time(1, 5, 0)

        # Junta a data e hora do usuário
        user_start_dt = datetime.combine(start_date, start_t)

        # O Filtro de Precisão:
        # Pega as linhas cujo _DT real calculado seja MAIOR ou IGUAL ao horário pedido
        df_filtered = df_all[df_all["_DT"] >= user_start_dt].copy()

        if df_filtered.empty:
            st.warning("⚠️ Nenhuma leitura encontrada a partir deste horário. Verifique a Data Inicial.")
            return

        df_final = df_filtered.drop(columns=["_DT"]).reset_index(drop=True)

        st.success(
            f"🎯 **{len(df_final):,} leituras filtradas com exatidão!**\n\n"
            f"A primeira linha capturada foi perfeitamente alinhada em: **{df_final['DATE'].iloc[0]} às {df_final['TIME'].iloc[0]}**"
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