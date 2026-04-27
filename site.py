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
    # Carregamento e limpeza de espaços invisíveis
    df_dados = pd.read_excel(arquivo, sheet_name="Dados")
    df_config = pd.read_excel(arquivo, sheet_name="Config")
    
    if 'Cod_Item' in df_config.columns:
        df_config['Cod_Item'] = df_config['Cod_Item'].astype(str).str.strip()
    if 'Cod_Item' in df_dados.columns:
        df_dados['Cod_Item'] = df_dados['Cod_Item'].astype(str).str.strip()
    if 'Forma_Utlizada' in df_dados.columns:
        df_dados['Forma_Utlizada'] = df_dados['Forma_Utlizada'].astype(str).str.strip()
    if 'Tipo_Item' in df_dados.columns:
        df_dados['Tipo_Item'] = df_dados['Tipo_Item'].astype(str).str.strip()
    
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
        
        # Remove linhas que não tenham altura medida para evitar quebra de cálculo
        df_filtrado = df_filtrado.dropna(subset=['Altura_Medida'])

        if not df_filtrado.empty:
            def buscar_config(row):
                ref_id = row['Cod_Item'] if str(row['Tipo_Item']).upper() == "FX" else row['Forma_Utlizada']
                conf = df_config[(df_config['Cod_Item'] == ref_id) & (df_config['Data_Inicio'] <= row['Data'])]
                
                if not conf.empty:
                    u = conf.sort_values('Data_Inicio', ascending=False).iloc[0]
                    nom = u['Valor_Nominal']
                    
                    # CORREÇÃO: Usa min() e max() para blindar contra valores trocados no Excel
                    v1 = nom + u['Limite_Sup']
                    v2 = nom + u['Limite_Inf']
                    lei = min(v1, v2)
                    les = max(v1, v2)
                    
                    # CORREÇÃO: Adiciona as etiquetas (index) para o Pandas saber onde guardar os números
                    return pd.Series([nom, lei, les], index=['Nominal', 'LEI', 'LES'])
                
                return pd.Series([None, None, None], index=['Nominal', 'LEI', 'LES'])

            # Atribuição direta com as colunas corretas
            df_filtrado[['Nominal', 'LEI', 'LES']] = df_filtrado.apply(buscar_config, axis=1)
            
            # Status seguro
            df_filtrado['Status'] = df_filtrado.apply(lambda r: "✅ OK" if pd.notna(r['LEI']) and r['LEI'] <= r['Altura_Medida'] <= r['LES'] else ("❌ Fora" if pd.notna(r['LEI']) else "⚠️ Sem Regra"), axis=1)
            df_filtrado['Desvio (mm)'] = df_filtrado['Altura_Medida'] - df_filtrado['Nominal']

            # --- EXIBIÇÃO: KPIs E CPK ---
            if show_kpis:
                st.markdown("### 📈 Resumo Estatístico e Capabilidade (Cpk)")
                
                resumo = df_filtrado.groupby("Tipo_Item")["Altura_Medida"].agg(['count', 'mean', 'std']).reset_index()
                resumo.columns = ["Item", "Qtd", "Média (mm)", "Desvio Padrão (σ)"]
                
                cpk_list = []
                for idx, row in resumo.iterrows():
                    t = row['Item']
                    df_t = df_filtrado[df_filtrado['Tipo_Item'] == t]
                    std = row['Desvio Padrão (σ)']
                    mean = row['Média (mm)']
                    
                    les = df_t['LES'].dropna().iloc[-1] if not df_t['LES'].dropna().empty else None
                    lei = df_t['LEI'].dropna().iloc[-1] if not df_t['LEI'].dropna().empty else None
                    
                    if pd.notna(std) and std > 0 and pd.notna(les) and pd.notna(lei):
                        cp_k = min((les - mean) / (3 * std), (mean - lei) / (3 * std))
                        cpk_list.append(cp_k)
                    else:
                        cpk_list.append(None)
                        
                resumo['Cpk'] = cpk_list

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Amostras", len(df_filtrado))
                c2.metric("Aprovação", f"{(len(df_filtrado[df_filtrado['Status'] == '✅ OK']) / len(df_filtrado)) * 100:.1f}%")
                c3.metric("Média Desvio", f"{df_filtrado['Desvio (mm)'].mean():.3f} mm")
                
                fx_cpk = resumo[resumo['Item'].str.upper() == 'FX']['Cpk']
                c4.metric("Cpk Geral (FX)", f"{fx_cpk.values[0]:.2f}" if not fx_cpk.empty and pd.notna(fx_cpk.values[0]) else "N/A")
                
                st.dataframe(resumo.style.format({
                    "Média (mm)": "{:.2f}", 
                    "Desvio Padrão (σ)": "{:.3f}",
                    "Cpk": "{:.2f}"
                }, na_rep="-"), use_container_width=True)

            # --- PAINEL DE DECISÃO ---
            if show_decision:
                st.markdown("---")
                st.subheader("🎯 Onde Ajustar?")
                desvio_resumo = df_filtrado.groupby("Tipo_Item")["Desvio (mm)"].mean().reset_index()
                fig_dec = px.bar(desvio_resumo, x="Tipo_Item", y="Desvio (mm)", color="Desvio (mm)",
                                 title="Componentes acima de zero estão empurrando o FX para cima; abaixo, para baixo.",
                                 color_continuous_scale='RdBu_r')
                st.plotly_chart(fig_dec, use_container_width=True)

            # --- GRÁFICOS DETALHADOS ---
            tipos = df_filtrado['Tipo_Item'].unique()
            for t in tipos:
                if show_box or show_sino:
                    st.subheader(f"Análise: {t}")
                    df_t = df_filtrado[df_filtrado['Tipo_Item'] == t]
                    
                    les = df_t['LES'].dropna().iloc[-1] if not df_t['LES'].dropna().empty else None
                    lei = df_t['LEI'].dropna().iloc[-1] if not df_t['LEI'].dropna().empty else None
                    nom = df_t['Nominal'].dropna().iloc[-1] if not df_t['Nominal'].dropna().empty else None
                    
                    col1, col2 = st.columns(2)
                    
                    # --- CURVA NORMAL (SINO) ---
                    with col1:
                        if show_sino:
                            mean_val = df_t['Altura_Medida'].mean()
                            std_val = df_t['Altura_Medida'].std()
                            fig_sino = go.Figure()

                            if pd.notna(std_val) and std_val > 0:
                                val_lei = lei if pd.notna(lei) else df_t['Altura_Medida'].min()
                                val_les = les if pd.notna(les) else df_t['Altura_Medida'].max()
                                min_x = min(val_lei, df_t['Altura_Medida'].min()) - std_val
                                max_x = max(val_les, df_t['Altura_Medida'].max()) + std_val
                                
                                x_curve = np.linspace(min_x, max_x, 200)
                                y_curve = (1 / (std_val * np.sqrt(2 * np.pi))) * np.exp(-0.5 * ((x_curve - mean_val) / std_val) ** 2)
                                
                                fig_sino.add_trace(go.Scatter(x=x_curve, y=y_curve, mode='lines', name='Curva Teórica', 
                                                              line=dict(color='rgba(150, 150, 150, 0.5)', width=2)))
                                
                                y_points = (1 / (std_val * np.sqrt(2 * np.pi))) * np.exp(-0.5 * ((df_t['Altura_Medida'] - mean_val) / std_val) ** 2)
                                hover_text = df_t.apply(lambda r: f"Amostra: {r.get('ID_Amostra', 'N/A')}<br>Data: {r['Data'].strftime('%d/%m/%Y')}<br>Altura: {r['Altura_Medida']}mm", axis=1)

                                fig_sino.add_trace(go.Scatter(x=df_t['Altura_Medida'], y=y_points, mode='markers', name='Peças Medidas', 
                                                              marker=dict(color='blue', size=8, line=dict(color='black', width=1)),
                                                              text=hover_text, hoverinfo="text"))
                            else:
                                fig_sino.add_trace(go.Scatter(x=df_t['Altura_Medida'], y=[1]*len(df_t), mode='markers', name='Peças Medidas'))

                            fig_sino.update_layout(title=f"Curva Normal (Dispersão)", xaxis_title="Altura (mm)", yaxis_title="Densidade", showlegend=False)
                            
                            if pd.notna(les): fig_sino.add_vline(x=les, line_dash="dash", line_color="red", annotation_text="Máx")
                            if pd.notna(lei): fig_sino.add_vline(x=lei, line_dash="dash", line_color="red", annotation_text="Mín")
                            if pd.notna(nom): fig_sino.add_vline(x=nom, line_dash="dot", line_color="green", annotation_text="Nominal")
                            
                            st.plotly_chart(fig_sino, use_container_width=True)
                    
                    # --- BOXPLOT ---
                    with col2:
                        if show_box:
                            fig_b = px.box(df_t, y="Altura_Medida", points="all", title=f"Boxplot")
                            if pd.notna(les): fig_b.add_hline(y=les, line_dash="dash", line_color="red", annotation_text="LES")
                            if pd.notna(lei): fig_b.add_hline(y=lei, line_dash="dash", line_color="red", annotation_text="LEI")
                            if pd.notna(nom): fig_b.add_hline(y=nom, line_dash="dot", line_color="green")
                            st.plotly_chart(fig_b, use_container_width=True)

            if show_table:
                st.markdown("---")
                st.subheader("📋 Dados Brutos")
                st.dataframe(df_filtrado, use_container_width=True)
else:
    st.info("Carregue a planilha para ver os gráficos.")
