import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

st.set_page_config(layout="wide", page_title="Análise de Qualidade - Cpk")

st.title("📊 Monitorização de Altura, Decisão e Cpk")

# --- BARRA LATERAL ---
st.sidebar.header("📂 Configurações")
arquivo = st.sidebar.file_uploader("Carregue a planilha", type=["xlsx"])

if arquivo is not None:
    # Carregamento
    df_dados = pd.read_excel(arquivo, sheet_name="Dados")
    df_config = pd.read_excel(arquivo, sheet_name="Config")
    
    # --- LIMPEZA DE CABEÇALHOS (Resolve o erro das colunas não encontradas) ---
    df_dados.columns = df_dados.columns.str.strip()
    df_config.columns = df_config.columns.str.strip()
    
    # Limpeza de dados nas colunas principais
    for col in ['Cod_Item', 'Tipo_Item']:
        if col in df_dados.columns: df_dados[col] = df_dados[col].astype(str).str.strip()
    if 'Forma_Utlizada' in df_dados.columns: df_dados['Forma_Utlizada'] = df_dados['Forma_Utlizada'].astype(str).str.strip()
    if 'Cod_Item' in df_config.columns: df_config['Cod_Item'] = df_config['Cod_Item'].astype(str).str.strip()
    
    df_dados['Data'] = pd.to_datetime(df_dados['Data']).dt.normalize()
    df_config['Data_Inicio'] = pd.to_datetime(df_config['Data_Inicio']).dt.normalize()

    st.sidebar.markdown("---")
    st.sidebar.subheader("🎯 Filtros")
    produtos = df_dados['Cod_Item'].unique()
    produto_sel = st.sidebar.selectbox("Produto selecionado", produtos)
    
    datas = df_dados['Data'].min().date(), df_dados['Data'].max().date()
    periodo = st.sidebar.date_input("Período", [datas[0], datas[1]])

    # --- CONTROLOS DE VISIBILIDADE ---
    st.sidebar.markdown("---")
    st.sidebar.subheader("👁️ Mostrar/Ocultar")
    show_kpis = st.sidebar.checkbox("Mostrar Resumo e Indicadores", value=True)
    show_decision = st.sidebar.checkbox("Mostrar Painel de Decisão", value=True)
    show_sino = st.sidebar.checkbox("Mostrar Curva Normal (Sino)", value=True)
    show_box = st.sidebar.checkbox("Mostrar Boxplots", value=True)
    show_table = st.sidebar.checkbox("Mostrar Tabela de Dados", value=False)

    # --- PROCESSAMENTO ---
    if len(periodo) == 2:
        start, end = pd.to_datetime(periodo[0]), pd.to_datetime(periodo[1])
        mask = (df_dados['Data'] >= start) & (df_dados['Data'] <= end) & (df_dados['Cod_Item'] == produto_sel)
        df_filtrado = df_dados.loc[mask].copy()
        df_filtrado = df_filtrado.dropna(subset=['Altura_Medida'])

        if not df_filtrado.empty:
            def buscar_config(row):
                tipo = str(row['Tipo_Item']).upper()
                search_id = row['Cod_Item'] if tipo == "FX" else row['Forma_Utlizada']
                
                # Busca exata
                conf = df_config[(df_config['Cod_Item'] == search_id) & (df_config['Data_Inicio'] <= row['Data'])]
                
                # Busca Inteligente para FX (ignora se tem 'FX' no nome ou não)
                if conf.empty and tipo == "FX":
                    clean_id = search_id.replace("FX", "")
                    conf = df_config[(df_config['Cod_Item'].str.contains(clean_id, case=False, na=False)) & (df_config['Data_Inicio'] <= row['Data'])]
                
                if not conf.empty:
                    u = conf.sort_values('Data_Inicio', ascending=False).iloc[0]
                    nom = u['Valor_Nominal']
                    # Garante que pega Limite_Sup e Limite_Inf mesmo com nomes trocados
                    v1, v2 = nom + u['Limite_Sup'], nom + u['Limite_Inf']
                    return pd.Series([nom, min(v1, v2), max(v1, v2)], index=['Nominal', 'LEI', 'LES'])
                
                return pd.Series([None, None, None], index=['Nominal', 'LEI', 'LES'])

            df_filtrado[['Nominal', 'LEI', 'LES']] = df_filtrado.apply(buscar_config, axis=1)
            df_filtrado['Status'] = df_filtrado.apply(lambda r: "✅ OK" if pd.notna(r['LEI']) and r['LEI'] <= r['Altura_Medida'] <= r['LES'] else ("❌ Fora" if pd.notna(r['LEI']) else "⚠️ Sem Regra"), axis=1)
            df_filtrado['Desvio (mm)'] = df_filtrado['Altura_Medida'] - df_filtrado['Nominal']

            # --- EXIBIÇÃO: KPIs E CPK ---
            if show_kpis:
                st.markdown("### 📈 Resumo Estatístico e Capabilidade (Cpk)")
                resumo = df_filtrado.groupby("Tipo_Item")["Altura_Medida"].agg(['count', 'mean', 'std']).reset_index()
                resumo.columns = ["Item", "Qtd", "Média (mm)", "Desvio Padrão (σ)"]
                
                cpk_list = []
                for idx, row in resumo.iterrows():
                    df_t = df_filtrado[df_filtrado['Tipo_Item'] == row['Item']]
                    std, mean = row['Desvio Padrão (σ)'], row['Média (mm)']
                    les = df_t['LES'].dropna().iloc[-1] if not df_t['LES'].dropna().empty else None
                    lei = df_t['LEI'].dropna().iloc[-1] if not df_t['LEI'].dropna().empty else None
                    
                    if pd.notna(std) and std > 0 and pd.notna(les) and pd.notna(lei):
                        cpk_list.append(min((les - mean)/(3*std), (mean - lei)/(3*std)))
                    else:
                        cpk_list.append(None)
                
                resumo['Cpk'] = cpk_list
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Amostras", len(df_filtrado))
                c2.metric("Aprovação", f"{(len(df_filtrado[df_filtrado['Status'] == '✅ OK']) / len(df_filtrado)) * 100:.1f}%")
                c3.metric("Média Desvio", f"{df_filtrado['Desvio (mm)'].mean():.3f} mm")
                fx_cpk = resumo[resumo['Item'].str.upper() == 'FX']['Cpk']
                c4.metric("Cpk Geral (FX)", f"{fx_cpk.values[0]:.2f}" if not fx_cpk.empty and pd.notna(fx_cpk.values[0]) else "N/A")
                
                st.dataframe(resumo.style.format({"Média (mm)": "{:.2f}", "Desvio Padrão (σ)": "{:.3f}", "Cpk": "{:.2f}"}, na_rep="-"), use_container_width=True)

            # --- PAINEL DE DECISÃO ---
            if show_decision:
                st.markdown("---")
                st.subheader("🎯 Onde Ajustar?")
                desvio_resumo = df_filtrado.groupby("Tipo_Item")["Desvio (mm)"].mean().reset_index()
                fig_dec = px.bar(desvio_resumo, x="Tipo_Item", y="Desvio (mm)", color="Desvio (mm)", color_continuous_scale='RdBu_r')
                st.plotly_chart(fig_dec, use_container_width=True)

            # --- GRÁFICOS DETALHADOS ---
            for t in df_filtrado['Tipo_Item'].unique():
                if show_box or show_sino:
                    st.subheader(f"Análise: {t}")
                    df_t = df_filtrado[df_filtrado['Tipo_Item'] == t]
                    les = df_t['LES'].dropna().iloc[-1] if not df_t['LES'].dropna().empty else None
                    lei = df_t['LEI'].dropna().iloc[-1] if not df_t['LEI'].dropna().empty else None
                    nom = df_t['Nominal'].dropna().iloc[-1] if not df_t['Nominal'].dropna().empty else None
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        if show_sino:
                            m, s = df_t['Altura_Medida'].mean(), df_t['Altura_Medida'].std()
                            fig_sino = go.Figure()
                            if pd.notna(s) and s > 0:
                                x = np.linspace(m-4*s, m+4*s, 200)
                                y = (1/(s*np.sqrt(2*np.pi)))*np.exp(-0.5*((x-m)/s)**2)
                                fig_sino.add_trace(go.Scatter(x=x, y=y, mode='lines', name='Sino', line=dict(color='gray', width=1)))
                                fig_sino.add_trace(go.Scatter(x=df_t['Altura_Medida'], y=(1/(s*np.sqrt(2*np.pi)))*np.exp(-0.5*((df_t['Altura_Medida']-m)/s)**2), mode='markers', name='Peças'))
                            fig_sino.update_layout(title="Curva Normal", showlegend=False)
                            if pd.notna(les): fig_sino.add_vline(x=les, line_dash="dash", line_color="red")
                            if pd.notna(lei): fig_sino.add_vline(x=lei, line_dash="dash", line_color="red")
                            if pd.notna(nom): fig_sino.add_vline(x=nom, line_dash="dot", line_color="green")
                            st.plotly_chart(fig_sino, use_container_width=True)
                    with col2:
                        if show_box:
                            fig_b = px.box(df_t, y="Altura_Medida", points="all", title="Boxplot")
                            if pd.notna(les): fig_b.add_hline(y=les, line_dash="dash", line_color="red")
                            if pd.notna(lei): fig_b.add_hline(y=lei, line_dash="dash", line_color="red")
                            st.plotly_chart(fig_b, use_container_width=True)

            if show_table:
                st.markdown("---")
                st.subheader("📋 Dados Brutos")
                st.dataframe(df_filtrado, use_container_width=True)
else:
    st.info("Carregue a planilha para ver os gráficos.")
