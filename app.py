import streamlit as st

from motor_conciliacao import carregar_arquivo, conciliar, gerar_excel

st.set_page_config(page_title="Conciliação Vindi x HubSpot", layout="wide")
st.title("Conciliação de Faturas: Vindi x HubSpot")
st.write("Faça o upload dos relatórios para processar a conciliação e gerar a planilha de resultados.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Relatório Vindi")
    file_vindi = st.file_uploader("Envie o arquivo da Vindi (Excel ou CSV)", type=["xlsx", "xls", "csv"], key="vindi")

with col2:
    st.subheader("Relatório HubSpot")
    file_hub = st.file_uploader("Envie o arquivo do HubSpot (Excel ou CSV)", type=["xlsx", "xls", "csv"], key="hub")

if file_vindi and file_hub:
    if st.button("Processar Conciliação", type="primary"):
        with st.spinner("Lendo arquivos e cruzando dados..."):
            try:
                abas = conciliar(carregar_arquivo(file_vindi), carregar_arquivo(file_hub))

                # Guarda no session_state: o clique no download_button reexecuta o script
                # e, sem isso, o resultado sumiria antes do download.
                total = len(abas['Compilado Geral'])
                nok = len(abas['Apenas NOKs'])
                st.session_state['relatorio_bytes'] = gerar_excel(abas)
                st.session_state['relatorio_resumo'] = {'total': total, 'ok': total - nok, 'nok': nok}
                st.session_state['relatorio_preview'] = abas['Apenas NOKs'].head(100)

            except Exception as e:
                st.session_state.pop('relatorio_bytes', None)
                st.error(f"Erro ao processar a conciliação: {e}")

if 'relatorio_bytes' in st.session_state:
    resumo = st.session_state['relatorio_resumo']
    st.success("Conciliação concluída com sucesso!")

    m1, m2, m3 = st.columns(3)
    m1.metric("Total de faturas", f"{resumo['total']:,}".replace(',', '.'))
    m2.metric("Sem divergência", f"{resumo['ok']:,}".replace(',', '.'))
    m3.metric("Com divergência (NOK)", f"{resumo['nok']:,}".replace(',', '.'))

    st.download_button(
        "Baixar Relatório de Conciliação (.xlsx)",
        data=st.session_state['relatorio_bytes'],
        file_name="Relatorio_Conciliacao_Vindi_Hubspot.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.subheader("Prévia — primeiras 100 divergências")
    st.dataframe(st.session_state['relatorio_preview'])
