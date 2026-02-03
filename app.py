import streamlit as st
import pandas as pd
import io
from src.processador_orcamento import ProcessadorOrcamento

def main():
    st.set_page_config(
        page_title="Remanejamento Orçamentário - SEFAZ",
        page_icon="💰",
        layout="wide"
    )

    st.title("💰 Sistema de Remanejamento Orçamentário Automatizado")
    st.markdown("""
    Esta aplicação processa planilhas de orçamento e realiza o remanejamento automatizado
    seguindo as regras estabelecidas pela SEFAZ.
    """)

    # Upload do arquivo
    st.header("1. Upload da Planilha")
    uploaded_file = st.file_uploader(
        "Envie o arquivo Excel (.xlsx ou .xls)",
        type=['xlsx', 'xls'],
        help="Selecione a planilha orçamentária que deseja processar"
    )

    if uploaded_file is not None:
        st.success(f"✅ Arquivo carregado: {uploaded_file.name}")

        # Configurações
        st.header("2. Configurações")

        with st.expander("⚙️ Configurar Fonte e Naturezas Proibidas", expanded=True):
            col1, col2 = st.columns(2)

            with col1:
                st.subheader("Fonte Proibida")
                fonte_proibida_input = st.text_input(
                    "Digite o código da fonte que NÃO deve participar de remanejamentos:",
                    value="761",
                    help="Exemplo: 761. Deixe em branco se não houver fonte proibida.",
                    placeholder="Ex: 761"
                )

                # Converter para int ou None
                fonte_proibida = None
                if fonte_proibida_input.strip():
                    try:
                        fonte_proibida = int(fonte_proibida_input.strip())
                    except ValueError:
                        st.error("Fonte deve ser um número inteiro!")

            with col2:
                st.subheader("Naturezas Proibidas")
                naturezas_input = st.text_area(
                    "Digite os códigos das naturezas que NÃO devem participar de remanejamentos (uma por linha):",
                    value="339018\n339092\n319092\n339047\n339048\n319096\n339093\n339091",
                    height=200,
                    help="Digite cada código de natureza em uma linha separada. Deixe em branco se não houver naturezas proibidas.",
                    placeholder="339018\n339092\n..."
                )

                # Processar naturezas
                naturezas_proibidas = set()
                if naturezas_input.strip():
                    for linha in naturezas_input.strip().split('\n'):
                        codigo = linha.strip()
                        if codigo:
                            # Remover pontos e espaços
                            codigo_limpo = codigo.replace('.', '').replace(' ', '')
                            naturezas_proibidas.add(codigo_limpo)

                if naturezas_proibidas:
                    st.info(f"📋 {len(naturezas_proibidas)} natureza(s) configurada(s) como proibida(s)")

        # Botão para processar
        st.header("3. Processamento")

        col1, col2 = st.columns([1, 3])
        with col1:
            processar = st.button("🔄 Calcular Remanejamento", type="primary", use_container_width=True)

        if processar:
            with st.spinner("Processando planilha... Por favor aguarde."):
                try:
                    # Inicializar processador com as configurações
                    processador = ProcessadorOrcamento(
                        fonte_proibida=fonte_proibida,
                        naturezas_proibidas=naturezas_proibidas
                    )

                    # Processar arquivo
                    resultado = processador.processar_arquivo(uploaded_file)

                    # Armazenar no session_state
                    st.session_state['resultado'] = resultado
                    st.session_state['processado'] = True

                    st.success("✅ Processamento concluído com sucesso!")

                except Exception as e:
                    st.error(f"❌ Erro ao processar arquivo: {str(e)}")
                    st.exception(e)
                    return

        # Exibir resultados se já processado
        if st.session_state.get('processado', False):
            resultado = st.session_state['resultado']

            st.header("4. Análise dos Resultados")

            # Métricas resumidas
            col1, col2, col3, col4 = st.columns(4)

            with col1:
                st.metric(
                    "UGs Analisadas",
                    resultado['estatisticas']['total_ugs']
                )

            with col2:
                st.metric(
                    "Déficits Encontrados",
                    resultado['estatisticas']['total_deficits']
                )

            with col3:
                st.metric(
                    "Remanejamentos Internos",
                    resultado['estatisticas']['remanejamentos_internos']
                )

            with col4:
                st.metric(
                    "Remanejamentos Externos",
                    resultado['estatisticas']['remanejamentos_externos']
                )

            # Exibir déficits encontrados
            if resultado['deficits']:
                with st.expander("📊 Déficits Identificados", expanded=True):
                    df_deficits = pd.DataFrame(resultado['deficits'])
                    st.dataframe(
                        df_deficits,
                        use_container_width=True,
                        hide_index=True
                    )

            # Exibir remanejamentos
            if resultado['remanejamentos']:
                with st.expander("🔄 Remanejamentos Realizados", expanded=False):
                    df_remanejamentos = pd.DataFrame(resultado['remanejamentos'])
                    st.dataframe(
                        df_remanejamentos,
                        use_container_width=True,
                        hide_index=True
                    )

            # Exibir diagnósticos detalhados
            with st.expander("🔍 Diagnósticos Detalhados (Log de Processamento)", expanded=False):
                st.code(resultado.get('diagnosticos', 'Nenhum diagnóstico disponível'), language='text')

            # Validações
            st.header("5. Validações")

            col1, col2 = st.columns(2)

            with col1:
                if resultado['validacoes']['nenhum_saldo_negativo']:
                    st.success("✅ Nenhuma UG ficou com saldo negativo")
                else:
                    st.error("❌ Ainda existem saldos negativos!")

            with col2:
                if resultado['validacoes']['somas_conferem']:
                    st.success("✅ Somas das transferências conferem")
                else:
                    st.warning("⚠️ Inconsistência nas somas")

            # Download do arquivo
            st.header("6. Download do Arquivo Ajustado")

            st.download_button(
                label="📥 Baixar Planilha Ajustada",
                data=resultado['arquivo_excel'],
                file_name=f"orcamento_ajustado_{uploaded_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

            st.info("""
            📋 **O arquivo contém duas abas:**
            - **Aba 1**: Saldos Ajustados (mesma estrutura da planilha original, com valores corrigidos)
            - **Aba 2**: Quadro de Remanejamento (detalhamento de todas as transferências realizadas)
            """)

    else:
        st.info("👆 Por favor, faça o upload de uma planilha Excel para começar.")

        st.markdown("""
        ### 📖 Como usar:

        1. **Upload**: Selecione o arquivo Excel com os dados orçamentários
        2. **Processamento**: Clique no botão "Calcular Remanejamento"
        3. **Análise**: Revise os déficits e remanejamentos realizados
        4. **Download**: Baixe a planilha ajustada com duas abas:
           - Aba 1: Saldos corrigidos
           - Aba 2: Detalhamento dos remanejamentos

        ### ⚙️ Regras de Remanejamento:

        - **Primeiro**: Remanejamento interno (dentro da mesma UG)
        - **Segundo**: Remanejamento externo (entre UGs diferentes)
        - **Garantia**: Nenhuma UG ficará com saldo negativo
        - **Rastreabilidade**: Todas as transferências são documentadas
        """)

if __name__ == "__main__":
    main()
