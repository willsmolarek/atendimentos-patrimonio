"""
app.py
------
Este é o arquivo PRINCIPAL do sistema — é ele que você executa para abrir
a aplicação no navegador (com o comando "streamlit run app.py").

Responsabilidades deste arquivo:
    - Montar toda a interface visual (telas, botões, filtros, formulários)
      usando a biblioteca Streamlit.
    - Chamar as funções prontas de database.py (para ler/gravar no banco)
      e de excel_handler.py (para importar/exportar planilhas).

Este arquivo NÃO contém regras de negócio nem comandos SQL diretamente —
toda essa lógica fica nos outros dois arquivos. Aqui só "montamos a tela".
"""

import streamlit as st
import pandas as pd
from datetime import date

import database
import excel_handler

# ----------------------------------------------------------------------
# CONFIGURAÇÃO INICIAL DA PÁGINA
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="Gestão de Atendimentos e Patrimônio",
    layout="wide",
)

# Garante que a tabela do banco de dados exista antes de qualquer coisa.
database.criar_tabela()

TIPOS_ATENDIMENTO = ["Máquina Nova", "Remanejamento"]
OPCOES_SIM_NAO = ["Sim", "Não"]
OPCOES_STATUS = ["Pendente", "Feito"]


# ----------------------------------------------------------------------
# FUNÇÕES AUXILIARES DE INTERFACE
# ----------------------------------------------------------------------

def linha_para_dict_formulario(prefixo_vazio=True) -> dict:
    """Retorna um dicionário 'em branco' usado para inicializar o formulário."""
    return {col: "" for col in database.COLUNAS_SEM_ID}


def exibir_metricas():
    """Desenha o painel de métricas no topo da página."""
    metricas = database.buscar_metricas()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total de Registros", metricas["total"])
    col2.metric("Pendentes", metricas["pendentes"])
    col3.metric("Feitos", metricas["feitos"])
    col4.metric("Prioridades de Hoje (pendentes)", metricas["prioridades_hoje"])


def estilizar_tabela(df: pd.DataFrame):
    """
    Aplica destaque visual e garante alto contraste para o texto das linhas:
    - Linhas com Status = "Feito" ficam esverdeadas com texto verde escuro.
    - Linhas com Prioridade do Dia = "Sim" ficam avermelhadas com texto vermelho escuro.
    """
    def cor_da_linha(linha):
        status = linha.get("Status") or linha.get("status")
        prioridade = linha.get("Prioridade do Dia") or linha.get("prioridade_dia")

        if status == "Feito":
            return ["background-color: #d4edda; color: #155724; font-weight: bold;"] * len(linha)
        if prioridade == "Sim":
            return ["background-color: #f8d7da; color: #721c24; font-weight: bold;"] * len(linha)
        return [""] * len(linha)

    return df.style.apply(cor_da_linha, axis=1)


def renomear_para_exibicao(df: pd.DataFrame) -> pd.DataFrame:
    """Troca os nomes internos das colunas pelos nomes amigáveis, só para exibição."""
    mapa = {"id": "ID", **excel_handler.MAPA_COLUNAS}
    return df.rename(columns=mapa)


# ----------------------------------------------------------------------
# BARRA LATERAL: FILTROS + IMPORTAÇÃO DE EXCEL
# ----------------------------------------------------------------------

st.sidebar.title("Filtros")

filtro_status = st.sidebar.selectbox("Status", ["Todos"] + OPCOES_STATUS)
filtro_prioridade = st.sidebar.selectbox("Prioridade do Dia", ["Todos"] + OPCOES_SIM_NAO)

predios_disponiveis = ["Todos"] + database.buscar_valores_unicos("numero_predio")
filtro_predio = st.sidebar.selectbox("Número do Prédio", predios_disponiveis)

ruas_disponiveis = ["Todos"] + database.buscar_valores_unicos("rua_dsi")
filtro_rua = st.sidebar.selectbox("Rua de Organização no DSI", ruas_disponiveis)

st.sidebar.markdown("---")
st.sidebar.title("Importar Excel")

arquivo_upload = st.sidebar.file_uploader(
    "Selecione um arquivo .xlsx",
    type=["xlsx"],
    help="A planilha pode estar praticamente do jeito que você já tem: "
    "colunas em qualquer ordem, com ou sem acento, siglas ou sinônimos "
    "(ex: 'Colaborador' no lugar de 'Nome do Usuário'). "
    "O sistema tenta reconhecer tudo automaticamente.",
)

if arquivo_upload is not None:
    if st.sidebar.button("Importar registros da planilha", use_container_width=True):
        try:
            registros, diagnostico = excel_handler.ler_planilha_para_dicionarios(
                arquivo_upload, retornar_diagnostico=True
            )
            quantidade = database.inserir_varios_atendimentos(registros)
            st.sidebar.success(f"{quantidade} registro(s) importado(s) com sucesso!")

            # Mostra um resumo de como cada coluna da planilha foi interpretada,
            # para o usuário conferir se o reconhecimento automático acertou.
            with st.sidebar.expander("Ver como as colunas foram reconhecidas"):
                st.write("**Colunas identificadas:**")
                for original, campo in diagnostico["colunas_reconhecidas"].items():
                    st.caption(f"'{original}' -> {excel_handler.MAPA_COLUNAS.get(campo, campo)}")

                if diagnostico["colunas_ignoradas"]:
                    st.write("**Colunas da planilha que foram ignoradas** (não reconhecidas):")
                    st.caption(", ".join(diagnostico["colunas_ignoradas"]))

                if diagnostico["campos_nao_encontrados"]:
                    st.write("**Campos do sistema que ficaram em branco** (não encontrados na planilha):")
                    st.caption(", ".join(
                        excel_handler.MAPA_COLUNAS.get(c, c) for c in diagnostico["campos_nao_encontrados"]
                    ))
            st.rerun()
        except Exception as erro:
            st.sidebar.error(f"Erro ao importar a planilha: {erro}")

st.sidebar.download_button(
    label="Baixar planilha modelo (vazia)",
    data=excel_handler.gerar_modelo_planilha(),
    file_name="modelo_atendimentos.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)


# ----------------------------------------------------------------------
# CABEÇALHO E MÉTRICAS
# ----------------------------------------------------------------------

st.title("Sistema de Gerenciamento de Atendimentos e Controle de Patrimônio")
st.caption("Gerencie trocas de máquina, remanejamentos e controle de patrimônio de forma simples.")

exibir_metricas()
st.markdown("---")


# ----------------------------------------------------------------------
# ABAS PRINCIPAIS
# ----------------------------------------------------------------------

aba_tabela, aba_cadastro = st.tabs(["Painel de Atendimentos", "Cadastro Manual"])


# ========================================================================
# ABA 1: PAINEL / TABELA INTERATIVA
# ========================================================================
with aba_tabela:

    df = database.buscar_atendimentos(
        status=filtro_status,
        prioridade_dia=filtro_prioridade,
        numero_predio=filtro_predio,
        rua_dsi=filtro_rua,
    )

    col_titulo, col_download = st.columns([3, 1])
    with col_titulo:
        st.subheader(f"Registros encontrados: {len(df)}")
    with col_download:
        if len(df) > 0:
            st.download_button(
                label="Exportar para Excel",
                data=excel_handler.gerar_excel_para_download(df),
                file_name="atendimentos_exportados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    if df.empty:
        st.info("Nenhum atendimento encontrado com os filtros selecionados.")
    else:
        st.markdown(
            "**Vermelho** = Prioridade do dia (pendente)  |  "
            "**Verde** = Já feito"
        )
        
        # Renomeia as colunas para exibição amigável antes de estilizar
        df_exibicao = renomear_para_exibicao(df)
        
        st.dataframe(
            estilizar_tabela(df_exibicao),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("---")
        st.subheader("Ação Rápida: Atualizar Status e Técnico")

        col_a, col_b, col_c, col_d = st.columns([1, 1, 1, 1])

        with col_a:
            ids_disponiveis = df["id"].tolist()
            id_selecionado = st.selectbox(
                "Selecione o ID do atendimento",
                ids_disponiveis,
                format_func=lambda x: f"ID {x} — {df.loc[df['id'] == x, 'nome_usuario'].values[0]}",
            )

        linha_selecionada = df.loc[df["id"] == id_selecionado].iloc[0]

        with col_b:
            novo_status = st.selectbox(
                "Novo status",
                OPCOES_STATUS,
                index=OPCOES_STATUS.index(linha_selecionada["status"])
                if linha_selecionada["status"] in OPCOES_STATUS
                else 0,
            )

        with col_c:
            novo_tecnico = st.text_input(
                "Técnico responsável",
                value=linha_selecionada["tecnico_responsavel"],
            )

        with col_d:
            st.write("")
            st.write("")
            if st.button("Salvar alteração", use_container_width=True, type="primary"):
                database.atualizar_status(int(id_selecionado), novo_status, novo_tecnico)
                st.success(f"Atendimento ID {id_selecionado} atualizado com sucesso!")
                st.rerun()

        with st.expander("Excluir um atendimento"):
            id_excluir = st.selectbox(
                "Selecione o ID a excluir",
                ids_disponiveis,
                key="id_excluir",
            )
            if st.button("Excluir definitivamente", type="secondary"):
                database.excluir_atendimento(int(id_excluir))
                st.warning(f"Atendimento ID {id_excluir} excluído.")
                st.rerun()


# ========================================================================
# ABA 2: CADASTRO MANUAL
# ========================================================================
with aba_cadastro:
    st.subheader("Cadastrar novo atendimento")

    with st.form("form_cadastro", clear_on_submit=True):
        col1, col2 = st.columns(2)

        with col1:
            patrimonio_novo = st.text_input("Patrimônio Novo")
            serial_number = st.text_input("Serial Number")
            nome_usuario = st.text_input("Nome do Usuário")
            nome_gestor = st.text_input("Nome do Gestor")
            tipo_atendimento = st.selectbox("Tipo de Atendimento", TIPOS_ATENDIMENTO)
            rua_dsi = st.text_input("Rua de Organização no DSI")
            prioridade_dia = st.selectbox("Prioridade do Dia", OPCOES_SIM_NAO, index=1)

        with col2:
            patrimonio_antigo = st.text_input("Patrimônio Antigo")
            secao_usuario = st.text_input("Seção do Usuário")
            tecnico_responsavel = st.text_input("Técnico Responsável")
            numero_predio = st.text_input("Número do Prédio", placeholder="Ex: Prédio 12")
            data_agendada = st.date_input("Data Agendada", value=date.today())
            status = st.selectbox("Status", OPCOES_STATUS)

        enviado = st.form_submit_button("Salvar Atendimento", use_container_width=True, type="primary")

        if enviado:
            if not nome_usuario.strip():
                st.error("O campo 'Nome do Usuário' é obrigatório.")
            else:
                novo_registro = {
                    "patrimonio_novo": patrimonio_novo,
                    "patrimonio_antigo": patrimonio_antigo,
                    "serial_number": serial_number,
                    "nome_usuario": nome_usuario,
                    "secao_usuario": secao_usuario,
                    "nome_gestor": nome_gestor,
                    "tecnico_responsavel": tecnico_responsavel,
                    "tipo_atendimento": tipo_atendimento,
                    "numero_predio": numero_predio,
                    "rua_dsi": rua_dsi,
                    "data_agendada": data_agendada.strftime("%Y-%m-%d"),
                    "prioridade_dia": prioridade_dia,
                    "status": status,
                }
                database.inserir_atendimento(novo_registro)
                st.success(f"Atendimento de '{nome_usuario}' cadastrado com sucesso!")
                st.rerun()