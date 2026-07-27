"""
excel_handler.py
----------------
Este arquivo cuida de toda a "conversa" entre o sistema e o Excel.

Responsabilidades:
    1. Ler um arquivo .xlsx enviado pelo usuário (upload) e transformar
       cada linha da planilha em um dicionário pronto para ser salvo no
       banco de dados (usando database.inserir_varios_atendimentos).
    2. Pegar dados que já estão no banco (um DataFrame do Pandas) e
       transformar em um arquivo .xlsx para o usuário baixar.

RECONHECIMENTO INTELIGENTE DE COLUNAS
--------------------------------------
A planilha do usuário pode vir com o cabeçalho praticamente de qualquer jeito:
    - Nomes "bonitos" (ex: "Patrimônio Novo") ou nomes internos (ex: "patrimonio_novo")
    - Com ou sem acento ("Patrimonio Novo")
    - Maiúsculas, minúsculas ou misturado ("PATRIMONIO NOVO")
    - Abreviado ou com sinônimos comuns ("Colaborador" em vez de "Nome do Usuário",
      "Setor" em vez de "Seção do Usuário", "SN" em vez de "Serial Number")
    - Em qualquer ordem de colunas (a ordem nunca importa, só o nome do cabeçalho)

Para lidar com isso, o reconhecimento acontece em 3 etapas, por coluna:
    1. Normalização do texto do cabeçalho (tira acento, deixa minúsculo,
       remove espaços/pontuação extra).
    2. Comparação exata com uma lista de "apelidos" conhecidos para cada
       campo do sistema (dicionário SINONIMOS).
    3. Se não encontrar exatamente, tenta uma comparação por semelhança
       (fuzzy matching) — útil para pequenos erros de digitação ou
       variações que não estão na lista de sinônimos.

Se mesmo assim uma coluna não for reconhecida, ela é apenas ignorada (e o
usuário é avisado na tela, para poder conferir se não esqueceu de nada).
"""

import io
import re
import unicodedata
import difflib

import pandas as pd
from database import COLUNAS_SEM_ID

# ------------------------------------------------------------------------
# NOMES "BONITOS" (usados na exportação e como referência principal)
# ------------------------------------------------------------------------
MAPA_COLUNAS = {
    "patrimonio_novo": "Patrimônio Novo",
    "patrimonio_antigo": "Patrimônio Antigo",
    "serial_number": "Serial Number",
    "nome_usuario": "Nome do Usuário",
    "secao_usuario": "Seção do Usuário",
    "nome_gestor": "Nome do Gestor",
    "tecnico_responsavel": "Técnico Responsável",
    "tipo_atendimento": "Tipo de Atendimento",
    "numero_predio": "Número do Prédio",
    "rua_dsi": "Rua de Organização no DSI",
    "data_agendada": "Data Agendada",
    "prioridade_dia": "Prioridade do Dia",
    "status": "Status",
}

# ------------------------------------------------------------------------
# SINÔNIMOS: para cada campo interno, uma lista de formas alternativas que
# o cabeçalho da planilha do usuário pode usar. Quanto mais variações você
# adicionar aqui, mais "esperto" fica o reconhecimento automático.
#
# Não precisa se preocupar com acento, maiúscula/minúscula ou espaços
# extras — tudo isso é normalizado antes da comparação (veja
# _normalizar_texto). Ou seja, "Seção", "secao", "SEÇÃO" e "Seçao" caem
# todos no mesmo texto normalizado.
# ------------------------------------------------------------------------
SINONIMOS = {
    "patrimonio_novo": [
        "patrimonio novo", "patrimonio", "novo patrimonio", "tombamento novo",
        "tombamento", "patrimonio atual", "num patrimonio novo", "pat novo",
        "codigo patrimonio", "cod patrimonio",
    ],
    "patrimonio_antigo": [
        "patrimonio antigo", "antigo patrimonio", "tombamento antigo",
        "pat antigo", "patrimonio anterior",
    ],
    "serial_number": [
        "serial number", "numero de serie", "numero serie", "serial",
        "sn", "n serie", "nº serie", "n de serie", "numero de serial",
    ],
    "nome_usuario": [
        "nome do usuario", "usuario", "nome usuario", "colaborador",
        "nome do colaborador", "funcionario", "nome do funcionario",
        "nome completo", "nome", "cliente", "nome do cliente",
    ],
    "secao_usuario": [
        "secao do usuario", "secao", "setor", "departamento",
        "secao usuario", "area", "divisao", "diretoria",
    ],
    "nome_gestor": [
        "nome do gestor", "gestor", "gerente", "responsavel gestor",
        "supervisor", "chefe", "coordenador", "lider",
    ],
    "tecnico_responsavel": [
        "tecnico responsavel", "tecnico", "responsavel tecnico",
        "atendente", "responsavel", "executor",
    ],
    "tipo_atendimento": [
        "tipo de atendimento", "tipo atendimento", "tipo", "categoria",
        "motivo", "motivo do atendimento",
    ],
    "numero_predio": [
        "numero do predio", "predio", "numero predio", "edificio",
        "bloco", "n predio", "nº predio", "unidade",
    ],
    "rua_dsi": [
        "rua de organizacao no dsi", "rua dsi", "rua",
        "organizacao dsi", "dsi", "rua de organizacao",
        "localizacao dsi", "rua no dsi",
    ],
    "data_agendada": [
        "data agendada", "data", "data do agendamento", "agendamento",
        "data prevista", "data marcada", "dia agendado",
    ],
    "prioridade_dia": [
        "prioridade do dia", "prioridade", "prioridade dia", "urgente",
        "prioritario", "e prioridade",
    ],
    "status": [
        "status", "situacao", "estado", "andamento",
    ],
}


def _normalizar_texto(texto: str) -> str:
    """
    Deixa um texto "limpo" para comparação, removendo o que costuma variar
    de uma planilha para outra sem mudar o significado:
        - Acentos (ç, ã, é, etc.)
        - Maiúsculas/minúsculas
        - Espaços duplicados, pontos, underlines, hífens

    Exemplos:
        "Seção do Usuário"  -> "secao do usuario"
        "SEÇÃO_USUARIO"     -> "secao usuario"
        "Nº Série"          -> "n serie"
    """
    texto = str(texto).strip().lower()
    # Remove acentos: decompõe (á -> a + acento) e descarta os acentos
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    # Troca qualquer caractere que não seja letra/número por espaço
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    # Colapsa espaços múltiplos em um só
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def _montar_mapa_normalizado() -> dict:
    """
    Monta, uma única vez, um dicionário "texto normalizado -> campo do
    banco" juntando o nome bonito, o nome interno e todos os sinônimos de
    cada campo.
    """
    mapa = {}
    for campo in COLUNAS_SEM_ID:
        candidatos = [campo, MAPA_COLUNAS.get(campo, "")] + SINONIMOS.get(campo, [])
        for candidato in candidatos:
            chave = _normalizar_texto(candidato)
            if chave:
                mapa[chave] = campo
    return mapa


MAPA_NORMALIZADO = _montar_mapa_normalizado()

# Mapa inverso "nome bonito -> nome do banco", mantido por compatibilidade
# com quem já usava essa constante.
MAPA_COLUNAS_INVERSO = {v: k for k, v in MAPA_COLUNAS.items()}

# Ponto de corte da comparação por semelhança (0 a 1). Quanto mais perto
# de 1, mais rigorosa a comparação (menos "chutes"). 0.72 tolera pequenos
# erros de digitação e abreviações leves, sem confundir campos diferentes.
LIMIAR_SEMELHANCA = 0.72


def identificar_coluna(cabecalho: str):
    """
    Descobre a qual campo do banco de dados um cabeçalho de planilha
    corresponde, usando (nessa ordem):
        1. Correspondência exata (após normalizar o texto)
        2. Correspondência por semelhança (fuzzy matching), para tolerar
           pequenas variações que não estão na lista de sinônimos

    Retorna:
        str: o nome do campo no banco (ex: "nome_usuario"), ou
        None: se não for possível identificar com confiança suficiente
    """
    texto_normalizado = _normalizar_texto(cabecalho)
    if not texto_normalizado:
        return None

    if texto_normalizado in MAPA_NORMALIZADO:
        return MAPA_NORMALIZADO[texto_normalizado]

    candidatos = list(MAPA_NORMALIZADO.keys())
    melhores = difflib.get_close_matches(
        texto_normalizado, candidatos, n=1, cutoff=LIMIAR_SEMELHANCA
    )
    if melhores:
        return MAPA_NORMALIZADO[melhores[0]]

    return None


def ler_planilha_para_dicionarios(arquivo_excel, retornar_diagnostico: bool = False):
    """
    Lê um arquivo Excel (.xlsx) enviado pelo usuário e devolve uma lista
    de dicionários, um por linha da planilha, já no formato aceito pelo
    banco de dados.

    A planilha pode ter as colunas em QUALQUER ordem e com nomes de
    cabeçalho variados (ver "RECONHECIMENTO INTELIGENTE DE COLUNAS" no
    topo deste arquivo) — o sistema se adapta automaticamente.

    Parâmetros:
        arquivo_excel: objeto de arquivo recebido do st.file_uploader()
                        do Streamlit (ou qualquer objeto compatível com
                        pandas.read_excel).
        retornar_diagnostico (bool): se True, além da lista de registros,
                        retorna também um dicionário com informações sobre
                        o que foi ou não reconhecido (útil para mostrar
                        um aviso na tela).

    Retorna:
        list[dict] — se retornar_diagnostico=False (padrão)
        (list[dict], dict) — se retornar_diagnostico=True, onde o dict tem:
            "colunas_reconhecidas": {cabecalho_original: campo_do_banco}
            "colunas_ignoradas": [cabecalhos que não foram reconhecidos]
            "campos_nao_encontrados": [campos do sistema que não vieram
                                        em nenhuma coluna da planilha]
    """
    df = pd.read_excel(arquivo_excel, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    renomear = {}
    colunas_reconhecidas = {}
    colunas_ignoradas = []
    campos_ja_usados = set()

    for coluna_planilha in df.columns:
        campo = identificar_coluna(coluna_planilha)

        # Evita que duas colunas da planilha "briguem" pelo mesmo campo do
        # banco: a primeira que for reconhecida para um campo vale; as
        # próximas colunas que apontarem para o mesmo campo são ignoradas
        # (e reportadas no diagnóstico, para o usuário conferir).
        if campo and campo not in campos_ja_usados:
            renomear[coluna_planilha] = campo
            colunas_reconhecidas[coluna_planilha] = campo
            campos_ja_usados.add(campo)
        else:
            colunas_ignoradas.append(coluna_planilha)

    df = df.rename(columns=renomear)

    colunas_presentes = [c for c in COLUNAS_SEM_ID if c in df.columns]
    df = df[colunas_presentes]

    campos_nao_encontrados = [c for c in COLUNAS_SEM_ID if c not in df.columns]

    for col in COLUNAS_SEM_ID:
        if col not in df.columns:
            df[col] = ""

    df = df[COLUNAS_SEM_ID]

    df = df.fillna("")
    df = df.astype(str)
    df = df.replace("nan", "")

    df["prioridade_dia"] = df["prioridade_dia"].apply(_normalizar_sim_nao)
    df["status"] = df["status"].apply(_normalizar_status)
    df["tipo_atendimento"] = df["tipo_atendimento"].apply(_normalizar_tipo_atendimento)

    registros = df.to_dict(orient="records")

    if retornar_diagnostico:
        diagnostico = {
            "colunas_reconhecidas": colunas_reconhecidas,
            "colunas_ignoradas": colunas_ignoradas,
            "campos_nao_encontrados": campos_nao_encontrados,
        }
        return registros, diagnostico

    return registros


def _normalizar_sim_nao(valor: str) -> str:
    """Converte variações de texto (sim, SIM, s, yes, 1, true, x) para 'Sim' ou 'Não'."""
    valor_limpo = _normalizar_texto(valor)
    if valor_limpo in ("sim", "s", "yes", "y", "1", "true", "verdadeiro", "x", "prioritario"):
        return "Sim"
    return "Não"


def _normalizar_status(valor: str) -> str:
    """Converte variações de texto para 'Pendente' ou 'Feito'."""
    valor_limpo = _normalizar_texto(valor)
    if valor_limpo in (
        "feito", "concluido", "done", "ok", "finalizado", "completo",
        "concluida", "realizado", "resolvido",
    ):
        return "Feito"
    return "Pendente"


def _normalizar_tipo_atendimento(valor: str) -> str:
    """Converte variações de texto para 'Máquina Nova' ou 'Remanejamento'."""
    valor_limpo = _normalizar_texto(valor)
    if "remanej" in valor_limpo:
        return "Remanejamento"
    if "nova" in valor_limpo or "maquina" in valor_limpo:
        return "Máquina Nova"
    return valor if valor in ("Máquina Nova", "Remanejamento") else "Máquina Nova"


def gerar_excel_para_download(df: pd.DataFrame) -> bytes:
    """
    Recebe um DataFrame do Pandas (por exemplo, o resultado de
    database.buscar_atendimentos()) e devolve os bytes de um arquivo
    .xlsx pronto para ser oferecido em um botão de download no Streamlit.
    """
    df_export = df.copy()

    colunas_renomeadas = {
        col: MAPA_COLUNAS.get(col, col) for col in df_export.columns if col != "id"
    }
    df_export = df_export.rename(columns=colunas_renomeadas)

    if "id" in df.columns:
        df_export = df_export.rename(columns={"id": "ID"})
        colunas_ordenadas = ["ID"] + [c for c in df_export.columns if c != "ID"]
        df_export = df_export[colunas_ordenadas]

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Atendimentos")

        planilha = writer.sheets["Atendimentos"]
        for i, coluna in enumerate(df_export.columns, start=1):
            maior_valor = df_export[coluna].astype(str).map(len).max() if len(df_export) else 0
            largura = max(len(str(coluna)), int(maior_valor)) + 4
            planilha.column_dimensions[planilha.cell(row=1, column=i).column_letter].width = min(largura, 45)

    buffer.seek(0)
    return buffer.getvalue()


def gerar_modelo_planilha() -> bytes:
    """
    Gera um arquivo .xlsx "modelo" vazio, apenas com os cabeçalhos corretos,
    para o usuário baixar, preencher e depois importar de volta no sistema.
    """
    colunas_bonitas = [MAPA_COLUNAS[col] for col in COLUNAS_SEM_ID]
    df_modelo = pd.DataFrame(columns=colunas_bonitas)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_modelo.to_excel(writer, index=False, sheet_name="Modelo")
        planilha = writer.sheets["Modelo"]
        for i, coluna in enumerate(colunas_bonitas, start=1):
            planilha.column_dimensions[planilha.cell(row=1, column=i).column_letter].width = max(len(coluna) + 4, 18)

    buffer.seek(0)
    return buffer.getvalue()
