import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. FUNÇÕES AUXILIARES
# ==========================================

def tratar_valor_float(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    if ',' in val_str:
        val_str = val_str.replace('.', '').replace(',', '.')
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def validar_valor(row):
    if pd.isna(row['ID Fatura Vindi']) or pd.isna(row['ID Fatura HubSpot']):
        return 'NOK - Valor (Fatura Ausente)'
    diff = abs((row['Valor Vindi'] or 0) - (row['Valor HubSpot'] or 0))
    return 'OK - Valor' if diff < 0.01 else 'NOK - Valor'

def validar_status(row):
    if not row.get('_existe_vindi') or not row.get('_existe_hub'):
        return 'NOK - Status'
    st_vindi = str(row.get('status_normalizado_vindi', '')).strip()
    st_hub = str(row.get('status_normalizado_hub', '')).strip()
    return 'OK - Status' if st_vindi == st_hub else 'NOK - Status'

def eh_vazio(val):
    if val is None or pd.isna(val) or str(val).strip() in ['', 'nan', 'NaT', 'None', '<NaT>']:
        return True
    return False

def datas_sao_iguais(d1, d2):
    if eh_vazio(d1) and eh_vazio(d2):
        return True
    if eh_vazio(d1) or eh_vazio(d2):
        return False
    return d1 == d2

def validar_datas(row):
    if not row.get('_existe_vindi') or not row.get('_existe_hub'):
        return 'NOK - Fatura Ausente'

    erros = []
    if not datas_sao_iguais(row.get('Data de Inicio_dt_vindi'), row.get('Data de Inicio_dt_hub')):
        erros.append('Data de Inicio')
    if not datas_sao_iguais(row.get('Data Fim_dt_vindi'), row.get('Data Fim_dt_hub')):
        erros.append('Data Fim')
    if not datas_sao_iguais(row.get('Data de vencimento_dt_vindi'), row.get('Data de vencimento_dt_hub')):
        erros.append('Data de Vencimento')

    if not erros:
        return 'OK - Datas'
    return f"NOK - {', '.join(erros)}"

def mapear_etapa_nok(row):
    etapas = []
    if row['Status Conciliação Fatura'] != 'OK':
        etapas.append('Fatura Ausente')
    if row['Validação Valor'] != 'OK - Valor':
        etapas.append('Divergência de Valor')
    if row['Validação Status'] != 'OK - Status':
        etapas.append('Divergência de Status')
    if row['Validação Datas'] != 'OK - Datas':
        etapas.append('Divergência de Datas')
    return ' | '.join(etapas)

def carregar_arquivo(uploaded_file):
    if uploaded_file.name.endswith('.csv'):
        # Tenta carregar CSV (pode precisar ajustar o separador dependendo do seu padrão)
        return pd.read_csv(uploaded_file)
    else:
        return pd.read_excel(uploaded_file)

# ==========================================
# 2. INTERFACE STREAMLIT
# ==========================================

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
                df_vindi = carregar_arquivo(file_vindi)
                df_hub = carregar_arquivo(file_hub)

                # ==========================================
                # 3. TRATAMENTO E PREPARAÇÃO DOS DADOS
                # ==========================================
                id_col_hub = 'ID da Fatura Vindi' if 'ID da Fatura Vindi' in df_hub.columns else ('ID' if 'ID' in df_hub.columns else df_hub.columns[0])
                id_col_vindi = 'id' if 'id' in df_vindi.columns else ('ID' if 'ID' in df_vindi.columns else df_vindi.columns[0])

                df_hub['id_vindi_cross'] = df_hub[id_col_hub].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                df_vindi['id_vindi_cross'] = df_vindi[id_col_vindi].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

                df_hub['_existe_hub'] = True
                df_vindi['_existe_vindi'] = True

                val_col_vindi = 'amount' if 'amount' in df_vindi.columns else df_vindi.columns[3]
                df_vindi['amount_clean'] = df_vindi[val_col_vindi].apply(tratar_valor_float)

                val_col_hub = 'Valor cobrado' if 'Valor cobrado' in df_hub.columns else ('amount' if 'amount' in df_hub.columns else 'Subtotal')
                df_hub['amount_clean'] = df_hub[val_col_hub].apply(tratar_valor_float)

                de_para_status = {
                    'anulado': 'canceled',
                    'pago': 'paid',
                    'aberto': 'pending',
                    'rascunho': 'draft'
                }

                status_col_hub = 'Status da fatura' if 'Status da fatura' in df_hub.columns else ('status' if 'status' in df_hub.columns else 'Status')
                status_col_vindi = 'status' if 'status' in df_vindi.columns else 'Status'

                df_hub['status_normalizado'] = df_hub[status_col_hub].astype(str).str.strip().str.lower().map(de_para_status).fillna(df_hub[status_col_hub].astype(str).str.strip().str.lower())
                df_vindi['status_normalizado'] = df_vindi[status_col_vindi].astype(str).str.strip().str.lower()

                for col in ['Data de Inicio', 'Data Fim', 'Data de vencimento']:
                    df_hub[col + '_dt'] = pd.to_datetime(df_hub[col], errors='coerce').dt.date if col in df_hub.columns else None

                vindi_venc = 'Data de vencimento' if 'Data de vencimento' in df_vindi.columns else ('Data de' if 'Data de' in df_vindi.columns else None)
                df_vindi['Data de Inicio_dt'] = pd.to_datetime(df_vindi['Data de Inicio'], errors='coerce').dt.date if 'Data de Inicio' in df_vindi.columns else None
                df_vindi['Data Fim_dt'] = pd.to_datetime(df_vindi['Data Fim'], errors='coerce').dt.date if 'Data Fim' in df_vindi.columns else None
                df_vindi['Data de vencimento_dt'] = pd.to_datetime(df_vindi[vindi_venc], errors='coerce').dt.date if vindi_venc else None

                # ==========================================
                # 4. MERGE / CRUZAMENTO DAS BASES
                # ==========================================
                merged = pd.merge(
                    df_vindi, df_hub, on='id_vindi_cross', how='outer', suffixes=('_vindi', '_hub')
                )

                # ==========================================
                # 5. CONSTRUÇÃO DAS ABAS DO RELATÓRIO
                # ==========================================
                aba1 = pd.DataFrame()
                aba1['ID Fatura Vindi'] = merged.apply(lambda r: r['id_vindi_cross'] if r.get('_existe_vindi') == True else None, axis=1)
                aba1['ID Fatura HubSpot'] = merged.apply(lambda r: r['id_vindi_cross'] if r.get('_existe_hub') == True else None, axis=1)
                aba1['Status Conciliação'] = merged.apply(lambda r: 'OK' if r.get('_existe_vindi') == True and r.get('_existe_hub') == True else 'NOK', axis=1)

                aba2 = pd.DataFrame()
                aba2['ID Fatura Vindi'] = aba1['ID Fatura Vindi']
                aba2['ID Fatura HubSpot'] = aba1['ID Fatura HubSpot']
                aba2['Valor Vindi'] = merged['amount_clean_vindi']
                aba2['Valor HubSpot'] = merged['amount_clean_hub']
                aba2['Validação Valor'] = aba2.apply(validar_valor, axis=1)

                aba3 = pd.DataFrame()
                aba3['ID Fatura Vindi'] = aba1['ID Fatura Vindi']
                aba3['ID Fatura HubSpot'] = aba1['ID Fatura HubSpot']
                aba3['Status Validação'] = merged.apply(validar_status, axis=1)
                aba3['Status HubSpot'] = merged.apply(lambda r: r.get(f'{status_col_hub}_hub') if str(r.get('status_normalizado_vindi')) != str(r.get('status_normalizado_hub')) else None, axis=1)
                aba3['Status Vindi'] = merged.apply(lambda r: r.get(f'{status_col_vindi}_vindi') if str(r.get('status_normalizado_vindi')) != str(r.get('status_normalizado_hub')) else None, axis=1)

                aba4 = pd.DataFrame()
                aba4['ID Fatura Vindi'] = aba1['ID Fatura Vindi']
                aba4['ID Fatura HubSpot'] = aba1['ID Fatura HubSpot']
                aba4['Validação Datas'] = merged.apply(validar_datas, axis=1)

                aba5 = pd.DataFrame()
                aba5['hubspot_company_id'] = merged.get('hubspot_company_id', merged.get('Empresa ID', None))
                aba5['hubspot_contact_id'] = merged.get('hubspot_contact_id', merged.get('Contato ID', None))
                aba5['ID Fatura Vindi'] = aba1['ID Fatura Vindi']
                aba5['ID Fatura HubSpot'] = aba1['ID Fatura HubSpot']
                aba5['Status Conciliação Fatura'] = aba1['Status Conciliação']
                aba5['Validação Valor'] = aba2['Validação Valor']
                aba5['Validação Status'] = aba3['Status Validação']
                aba5['Validação Datas'] = aba4['Validação Datas']
                aba5['Valor Vindi'] = merged['amount_clean_vindi']
                aba5['Valor HubSpot'] = merged['amount_clean_hub']

                condicao_nok = (
                    (aba5['Status Conciliação Fatura'] != 'OK') |
                    (aba5['Validação Valor'] != 'OK - Valor') |
                    (aba5['Validação Status'] != 'OK - Status') |
                    (aba5['Validação Datas'] != 'OK - Datas')
                )
                aba6 = aba5[condicao_nok].copy()
                aba6['Etapas NOK'] = aba6.apply(mapear_etapa_nok, axis=1)

                # ==========================================
                # 6. EXPORTAÇÃO PARA O EXCEL EM MEMÓRIA
                # ==========================================
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    aba1.to_excel(writer, sheet_name='Relação de Faturas', index=False)
                    aba2.to_excel(writer, sheet_name='Relação Faturas e Valor', index=False)
                    aba3.to_excel(writer, sheet_name='Validação de Status', index=False)
                    aba4.to_excel(writer, sheet_name='Validação de Datas', index=False)
                    aba5.to_excel(writer, sheet_name='Compilado Geral', index=False)
                    aba6.to_excel(writer, sheet_name='Apenas NOKs', index=False)
                
                excel_data = output.getvalue()

                st.success("Processo concluído com sucesso!")
                
                # Exibe um preview rápido dos erros encontrados (Opcional)
                st.warning(f"Foram encontrados {len(aba6)} registros com divergências (NOKs).")
                with st.expander("Visualizar os 5 primeiros NOKs"):
                    st.dataframe(aba6.head(5))

                # Botão para baixar o arquivo
                st.download_button(
                    label="📥 Baixar Relatório de Conciliação em Excel",
                    data=excel_data,
                    file_name="Relatorio_Conciliacao_Vindi_Hubspot.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Ocorreu um erro ao processar os arquivos: {e}")