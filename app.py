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


def generate_diagnostic_report(raw_bytes: bytes) -> str:
    """Gera um relatório técnico em texto seguro contendo amostras hexadecimais

    do arquivo para identificação exata da estrutura do FieldLogger.
    """
    report = []
    total_len = len(raw_bytes)
    report.append(f"=== DIAGNÓSTICO DO ARQUIVO MEMFLASH.fl ===")
    report.append(f"Tamanho Total: {total_len:,} bytes ({total_len/1024:.2f} KB)")
    report.append("")

    # Primeiros 1024 bytes (Cabeçalho do FieldLogger)
    report.append("--- [BLOCO 1] CABEÇALHO (PRIMEIROS 1024 BYTES) ---")
    header_slice = raw_bytes[:1024]
    report.append(header_slice.hex())
    report.append("")

    # Últimos 1024 bytes
    report.append("--- [BLOCO 2] FIM DO ARQUIVO (ÚLTIMOS 1024 BYTES) ---")
    tail_slice = raw_bytes[-1024:] if total_len >= 1024 else raw_bytes
    report.append(tail_slice.hex())
    report.append("")

    # Varredura à procura de padrões de temperatura conhecidos (50.0 a 85.0 °C)
    report.append("--- [BLOCO 3] LOCALIZADOR DE VALORES REAIS (IEEE 754) ---")
    matches = 0
    # Teste Little-Endian padrão
    for offset in range(0, min(total_len - 4, 100000), 2):
        chunk = raw_bytes[offset : offset + 4]
        try:
            val = struct.unpack("<f", chunk)[0]
            # Procura por números típicos do seu processo (Ex: 57.5, 41.7, 82.1)
            if 56.0 <= val <= 59.0:
                snippet = raw_bytes[max(0, offset - 16) : offset + 48]
                report.append(
                    f"Possível valor de Temp (~{val:.3f}) no offset {offset} (0x{offset:X}):"
                )
                report.append(f"Hex vizinho: {snippet.hex()}")
                matches += 1
                if matches >= 5:
                    break
        except Exception:
            pass

    return "\n".join(report)


def main():
    st.title("🏭 Conversor_Log_Siliaris")

    tab_conversor, tab_diag = st.tabs(
        ["📄 Conversor de Log", "🔍 Diagnóstico & Amostra (.txt)"]
    )

    with tab_conversor:
        st.markdown(
            "Extração direta de dados do arquivo **MEMFLASH.fl** para planilha Excel (.xlsx)."
        )
        selected_machine = st.selectbox(
            "Selecione o Equipamento:", list(PROFILES.keys())
        )
        uploaded_file = st.file_uploader(
            "Arraste e solte o arquivo MEMFLASH.fl aqui:",
            type=["fl", "bin"],
            key="file_conv",
        )

        if uploaded_file is not None:
            st.info(
                "Arquivo carregado com sucesso. Use a aba **'🔍 Diagnóstico & Amostra (.txt)'** acima para extrairmos a estrutura correta!"
            )

    with tab_diag:
        st.subheader("🛠️ Extrator de Amostra para Ajuste de Precisão")
        st.markdown(
            "Como o arquivo `.fl` possui uma criptografia/empacotamento proprietário da Novus, "
            "esta ferramenta extrai uma amostra em formato texto seguro para desvendarmos a estrutura exata."
        )

        uploaded_diag = st.file_uploader(
            "Envie o MEMFLASH.fl para gerar o relatório de diagnóstico:",
            type=["fl", "bin"],
            key="file_diag",
        )

        if uploaded_diag is not None:
            raw_bytes = uploaded_diag.read()
            with st.spinner("Gerando diagnóstico em formato texto seguro..."):
                diag_text = generate_diagnostic_report(raw_bytes)

            st.success("Diagnóstico gerado com sucesso!")

            st.download_button(
                label="📥 Baixar Amostra de Diagnóstico (amostra_diagnostico.txt)",
                data=diag_text,
                file_name="amostra_diagnostico.txt",
                mime="text/plain",
                type="primary",
            )

            with st.expander("Ver prévia do diagnóstico"):
                st.text_area(
                    "Texto do diagnóstico:", diag_text[:2000] + "\n...", height=300
                )


if __name__ == "__main__":
    main()