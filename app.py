import streamlit as st
import pandas as pd
import io
import os
from src.processador_orcamento import ProcessadorOrcamento
from src.gerador_lote import GeradorLote
from datetime import date

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets')
REGRA41_PATH = os.path.join(ASSETS_DIR, 'Itens da Regra de Mapeamento 41 (2).xls')
REGRA100_PATH = os.path.join(ASSETS_DIR, 'Itens da Regra de Mapeamento 100 (1).xls')


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

    st.header("1. Upload da Planilha")
    uploaded_file = st.file_uploader(
        "Envie o arquivo Excel (.xlsx ou .xls)",
        type=['xlsx', 'xls'],
        help="Selecione a planilha orçamentária que deseja processar"
    )

    if uploaded_file is not None:
        st.success(f"✅ Arquivo carregado: {uploaded_file.name}")

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

            # Gerador
            st.header("7. Gerar Arquivo SIAFE")

            st.markdown("""
            Gere o arquivo no formato de importação em lote do SIAFE a partir dos
            remanejamentos calculados. Você precisa fornecer os arquivos de Regras
            de Mapeamento e preencher os dados obrigatórios.
            """)

            with st.expander("📤 Configurar e Gerar Arquivo SIAFE", expanded=False):
                st.subheader("Dados Obrigatórios")

                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    data_emissao = st.date_input(
                        "Data de Emissão",
                        value=date.today(),
                        format="DD/MM/YYYY"
                    )
                with col_d2:
                    processo = st.text_input(
                        "Número do Processo",
                        value="",
                        placeholder="Ex: 2026/00001",
                        help="Número do processo administrativo"
                    )

                observacao = st.text_input(
                    "Observação",
                    value="REMANEJAMENTO FOLHA DE PESSOAL",
                    help="Texto descritivo para o campo Observação do SIAFE"
                )

                st.caption("ℹ️ A UG Emitente é preenchida automaticamente com a UG Acrescida (UG Destino) de cada linha.")

                # Botão para gerar
                gerar_siafe = st.button(
                    "📄 Gerar Arquivo SIAFE",
                    type="secondary",
                    use_container_width=True,
                )

                if gerar_siafe:
                    with st.spinner("Gerando arquivo SIAFE..."):
                        try:
                            gerador = GeradorLote()

                            # Carregar Regra 41 do diretório assets/
                            gerador.carregar_regra41(REGRA41_PATH)
                            st.info(f"📋 Regra 41 carregada: {len(gerador.mapa_ug)} UGs mapeadas")

                            # Carregar Regra 100 do diretório assets/
                            gerador.carregar_regra100(REGRA100_PATH)
                            st.info("📋 Regra 100 carregada")

                            # Preparar DataFrame de remanejamentos
                            df_rem = pd.DataFrame(resultado['remanejamentos'])

                            # Formatar data
                            data_fmt = data_emissao.strftime("%d/%m/%Y")

                            # Gerar lote
                            arquivo_siafe, erros = gerador.gerar_lote(
                                df_remanejamentos=df_rem,
                                data_emissao=data_fmt,
                                observacao=observacao,
                                processo=processo,
                            )

                            # Exibir erros (se houver)
                            if erros:
                                st.markdown(f"**⚠️ {len(erros)} aviso(s) durante a geração:**")
                                for erro in erros:
                                    st.warning(erro)

                            # Contar linhas
                            n_remanejamentos = len(df_rem)
                            n_linhas_siafe = n_remanejamentos * 2

                            st.success(
                                f"✅ Arquivo SIAFE gerado com sucesso! "
                                f"{n_remanejamentos} remanejamentos → {n_linhas_siafe} linhas SIAFE "
                                f"(Redução + Acréscimo para cada)"
                            )

                            # Download
                            st.download_button(
                                label="📥 Baixar Arquivo SIAFE",
                                data=arquivo_siafe,
                                file_name=f"siafe_importacao_{data_emissao.strftime('%Y%m%d')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary",
                                use_container_width=True
                            )

                        except Exception as e:
                            st.error(f"❌ Erro ao gerar arquivo SIAFE: {str(e)}")
                            st.exception(e)

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
        5. **SIAFE**: Gere o arquivo de importação em lote para o SIAFE

        ### ⚙️ Regras de Remanejamento:

        - **Primeiro**: Remanejamento interno (dentro da mesma UG)
        - **Segundo**: Remanejamento externo (entre UGs diferentes)
        - **Garantia**: Nenhuma UG ficará com saldo negativo
        - **Rastreabilidade**: Todas as transferências são documentadas
        """)

if __name__ == "__main__":
    main()
