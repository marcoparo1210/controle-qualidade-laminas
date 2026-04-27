import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import math

st.set_page_config(layout="wide", page_title="Análise de Qualidade - UTFPR")

st.title("📊 Monitorização de Altura e Diagnóstico de Limites")

# --- BARRA LATERAL ---
st.sidebar.header("📂 Configurações")
arquivo = st.sidebar.file_uploader("Carregue a planilha", type=["xlsx"])

if arquivo is not None:
    # Carregamento e Normalização Forçada de Colunas
    df_dados = pd.read_excel(arquivo, sheet_name="Dados")
    df_config = pd.read_excel(arquivo, sheet_name="Config")
    
    df_dados.columns = df_dados.columns.str.strip().str.upper()
    df_config.columns = df_config.columns.str.strip().str.upper()
    
    for df in [df_dados, df_config]:
        for col in ['COD_ITEM', 'TIPO_ITEM', 'FORMA_UTLIZADA']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.upper()

    df_dados['DATA'] = pd.to_datetime(df_dados['DATA']).dt.normalize()
    df_config['DATA_INICIO'] = pd.to_datetime(df_config['DATA_INICIO']).dt.normalize()

    st.sidebar.markdown("---")
    st.sidebar.subheader("🎯 Filtros")
    produto_sel = st.sidebar.selectbox("Produto em Análise", df_dados['COD_ITEM'].unique())
    
    datas = df_dados['DATA'].min().date(), df_dados['DATA'].max().date()
    periodo = st.sidebar.date_input("Período", [datas[0], datas[1]])
    
    st.sidebar.markdown("---")
    st.sidebar.info("📐 **Padrão Estatístico Ativo:** O cálculo de amostras ideais para 95% de confiança está fixado com uma Margem de Erro (E) equivalente a **10% da Tolerância Total** de cada peça.")

    st.sidebar.markdown("---")
    st.sidebar.subheader("👁️ Mostrar/Ocultar")
    show_kpis = st.sidebar.checkbox("Mostrar Resumo e Indicadores", value=True)
    show_decision = st.sidebar.checkbox("Mostrar Painel de Decisão", value=True)
    show_sino = st.sidebar.checkbox("Mostrar Curva Normal (Sino)", value=True)
    show_box = st.sidebar.checkbox("Mostrar Boxplots", value=True)

    if len(periodo) == 2:
        start, end = pd.to_datetime(periodo[0]), pd.to_datetime(periodo[1])
        df_filtrado = df_dados[(df_dados['DATA'] >= start) & (df_dados['DATA'] <= end) & (df_dados['COD_ITEM'] == produto_sel)].copy()
        df_filtrado = df_filtrado.dropna(subset=['ALTURA_MEDIDA'])

        if not df_filtrado.empty:
            def buscar_config(row):
                tipo = str(row['TIPO_ITEM']).upper()
                search_id = row['COD_ITEM']
                if "FX" not in search_id and tipo == "FX": search_id_alt = f"{search_id}FX"
                else: search_id_alt = search_id.replace("FX", "")
                forma = str(row['FORMA_UTLIZADA']).upper()

                mask_nome = ((df_config['COD_ITEM'] == search_id) | 
                             (df_config['COD_ITEM'] == search_id_alt) | 
                             (df_config['COD_ITEM'] == forma))
                
                conf = df_config[mask_nome & (df_config['DATA_INICIO'] <= row['DATA'])]
                if conf.empty: conf = df_config[mask_nome]
                
                if not conf.empty:
                    u = conf.sort_values('DATA_INICIO', ascending=False).iloc[0]
                    nom = u['VALOR_NOMINAL']
                    col_sup = [c for c in df_config.columns if "SUP" in c][0]
                    col_inf = [c for c in df_config.columns if "INF" in c][0]
                    v1, v2 = nom + u[col_sup], nom + u[col_inf]
                    return pd.Series([nom, min(v1, v2), max(v1, v2)], index=['NOMINAL', 'LEI', 'LES'])
                return pd.Series([None, None, None], index=['NOMINAL', 'LEI', 'LES'])

            df_filtrado[['NOMINAL', 'LEI', 'LES']] = df_filtrado.apply(buscar_config, axis=1)
            df_filtrado['DESVIO'] = df_filtrado['ALTURA_MEDIDA'] - df_filtrado['NOMINAL']
            df_filtrado['STATUS'] = df_filtrado.apply(lambda r: "✅ OK" if pd.notna(r['LEI']) and r['LEI'] <= r['ALTURA_MEDIDA'] <= r['LES'] else "⚠️ S/ Limite", axis=1)

            itens_sem_limite = df_filtrado[df_filtrado['NOMINAL'].isna()]['TIPO_ITEM'].unique()
            if len(itens_sem_limite) > 0:
                st.warning(f"Atenção: Limites não encontrados na aba Config para os itens: {', '.join(itens_sem_limite)}.")

            # --- MÉTRICAS, CPK E CONFIANÇA 95% PADRONIZADA ---
            if show_kpis:
                resumo = df_filtrado.groupby("TIPO_ITEM")["ALTURA_MEDIDA"].agg(['count', 'mean', 'std']).reset_index()
                resumo.columns = ["Item", "Qtd", "Média (mm)", "Desvio Padrão (σ)"]
                
                cpk_list = []
                amostras_status_list = []
                
                for idx, row in resumo.iterrows():
                    df_t = df_filtrado[df_filtrado['TIPO_ITEM'] == row['Item']]
                    std, mean, qtd = row['Desvio Padrão (σ)'], row['Média (mm)'], row['Qtd']
                    les = df_t['LES'].dropna().iloc[-1] if not df_t['LES'].dropna().empty else None
                    lei = df_t['LEI'].dropna().iloc[-1] if not df_t['LEI'].dropna().empty else None
                    
                    if pd.notna(std) and std > 0 and pd.notna(les) and pd.notna(lei):
                        cpk_list.append(min((les - mean)/(3*std), (mean - lei)/(3*std)))
                    else:
                        cpk_list.append(None)
                        
                    if pd.notna(std) and std > 0 and pd.notna(les) and pd.notna(lei) and (les > lei):
                        tolerancia_total = les - lei
                        margem_erro = 0.10 * tolerancia_total
                        n_ideal = math.ceil(((1.96 * std) / margem_erro) ** 2)
                        
                        if qtd >= n_ideal:
                            amostras_status_list.append("✅ OK")
                        else:
                            amostras_status_list.append(f"Faltam {n_ideal - qtd}")
                    else:
                        amostras_status_list.append("N/A")

                resumo['Cpk'] = cpk_list
                resumo['Confiança 95%'] = amostras_status_list

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Amostras", len(df_filtrado))
                c2.metric("Aprovação", f"{(len(df_filtrado[df_filtrado['STATUS'] == '✅ OK']) / len(df_filtrado)) * 100:.1f}%")
                c3.metric("Desvio Médio", f"{df_filtrado['DESVIO'].mean():.3f} mm")
                fx_cpk = resumo[resumo['Item'].str.upper() == 'FX']['Cpk']
                c4.metric("Cpk (FX)", f"{fx_cpk.values[0]:.2f}" if not fx_cpk.empty and pd.notna(fx_cpk.values[0]) else "N/A")
                
                st.dataframe(resumo.style.format({"Média (mm)": "{:.2f}", "Desvio Padrão (σ)": "{:.3f}", "Cpk": "{:.2f}"}, na_rep="-"), use_container_width=True)

            # --- PAINEL DE DECISÃO (NOVO: FOCADO NO ALVO FX) ---
            if show_decision:
                st.markdown("---")
                st.subheader("🎯 Decisão de Ajuste do Feixe (Alvo: Nominal + 2)")
                
                df_fx = df_filtrado[df_filtrado['TIPO_ITEM'] == 'FX']
                
                if not df_fx.empty:
                    media_fx = df_fx['ALTURA_MEDIDA'].mean()
                    nominal_fx = df_fx['NOMINAL'].dropna().iloc[-1] if not df_fx['NOMINAL'].dropna().empty else None
                    
                    if pd.notna(nominal_fx):
                        alvo_fx = nominal_fx + 2
                        ajuste = alvo_fx - media_fx
                        
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Média Atual do Feixe", f"{media_fx:.2f} mm")
                        col2.metric("Alvo Ideal (Nominal + 2)", f"{alvo_fx:.2f} mm")
                        
                        if ajuste > 0:
                            col3.metric("Ajuste Necessário", f"⬆️ Subir {ajuste:.2f} mm", delta=f"+{ajuste:.2f} mm")
                        elif ajuste < 0:
                            col3.metric("Ajuste Necessário", f"⬇️ Descer {abs(ajuste):.2f} mm", delta=f"{ajuste:.2f} mm", delta_color="inverse")
                        else:
                            col3.metric("Ajuste Necessário", "✅ Perfeito", delta="0.00 mm", delta_color="off")
                            
                        st.info(f"Para colocar a média do feixe exatamente no alvo de **{alvo_fx:.2f} mm**, aplique uma alteração de **{ajuste:+.2f} mm** no conjunto.")
                    else:
                        st.warning("Valor nominal do Feixe (FX) não encontrado para calcular o alvo.")
                else:
                    st.info("Não existem amostras do Feixe (FX) no período selecionado para calcular o ajuste.")

            # --- GRÁFICOS ---
            for t in df_filtrado['TIPO_ITEM'].unique():
                if show_box or show_sino:
                    st.subheader(f"Análise: {t}")
                    df_t = df_filtrado[df_filtrado['TIPO_ITEM'] == t]
                    les = df_t['LES'].dropna().iloc[-1] if not df_t['LES'].dropna().empty else None
                    lei = df_t['LEI'].dropna().iloc[-1] if not df_t['LEI'].dropna().empty else None
                    nom = df_t['NOMINAL'].dropna().iloc[-1] if not df_t['NOMINAL'].dropna().empty else None
                    
                    col_a, col_b = st.columns(2)
                    with col_a:
                        if show_sino:
                            fig_sino = go.Figure()
                            m, s = df_t['ALTURA_MEDIDA'].mean(), df_t['ALTURA_MEDIDA'].std()
                            if pd.notna(s) and s > 0:
                                x = np.linspace(m-4*s, m+4*s, 100)
                                y = (1/(s*np.sqrt(2*np.pi)))*np.exp(-0.5*((x-m)/s)**2)
                                fig_sino.add_trace(go.Scatter(x=x, y=y, mode='lines', name='Sino', line=dict(color='gray')))
                                fig_sino.add_trace(go.Scatter(x=df_t['ALTURA_MEDIDA'], y=(1/(s*np.sqrt(2*np.pi)))*np.exp(-0.5*((df_t['ALTURA_MEDIDA']-m)/s)**2), mode='markers', name='Peças'))
                            
                            if pd.notna(les): fig_sino.add_vline(x=les, line_dash="dash", line_color="red", annotation_text="LES")
                            if pd.notna(lei): fig_sino.add_vline(x=lei, line_dash="dash", line_color="red", annotation_text="LEI")
                            if pd.notna(nom): fig_sino.add_vline(x=nom, line_dash="dot", line_color="green", annotation_text="NOM")
                            st.plotly_chart(fig_sino, use_container_width=True)
                    
                    with col_b:
                        if show_box:
                            fig_box = px.box(df_t, y="ALTURA_MEDIDA", points="all", title="Boxplot")
                            if pd.notna(les): fig_box.add_hline(y=les, line_dash="dash", line_color="red")
                            if pd.notna(lei): fig_box.add_hline(y=lei, line_dash="dash", line_color="red")
                            st.plotly_chart(fig_box, use_container_width=True)
else:
    st.info("A aguardar carregamento da planilha...")
