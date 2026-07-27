"""
database.py
-----------
Este arquivo é o "coração" de dados do sistema.

Responsabilidades deste arquivo:
    1. Criar o banco de dados SQLite e a tabela "atendimentos" automaticamente,
       caso ainda não existam (na primeira vez que o app rodar).
    2. Fornecer funções prontas para:
       - Inserir um novo atendimento (cadastro manual ou importado do Excel)
       - Buscar todos os atendimentos (com ou sem filtros)
       - Atualizar o status de um atendimento (Pendente -> Feito) e o técnico
       - Excluir um atendimento
       - Buscar métricas (totais de pendentes, feitos, prioridades do dia)

Nenhuma outra parte do sistema deve "conversar" diretamente com o SQLite.
Toda a lógica de banco fica concentrada aqui. Isso é uma boa prática:
se um dia você trocar o SQLite por outro banco, só precisa mexer neste arquivo.
"""

import sqlite3
import pandas as pd
from contextlib import contextmanager

# Nome do arquivo do banco de dados. Ele será criado na mesma pasta do projeto.
DB_NAME = "patrimonio.db"

# Lista oficial das colunas da tabela, na ordem que usaremos em todo o sistema.
# Mantemos essa lista em um único lugar para evitar erros de digitação espalhados
# pelo código (todo o resto do sistema importa essa lista).
COLUNAS = [
    "id",
    "patrimonio_novo",
    "patrimonio_antigo",
    "serial_number",
    "nome_usuario",
    "secao_usuario",
    "nome_gestor",
    "tecnico_responsavel",
    "tipo_atendimento",
    "numero_predio",
    "rua_dsi",
    "data_agendada",
    "prioridade_dia",
    "status",
]

# Colunas que o usuário preenche (todas, exceto o "id", que é automático)
COLUNAS_SEM_ID = COLUNAS[1:]


@contextmanager
def get_connection():
    """
    Cria e devolve uma conexão com o banco de dados SQLite.

    Usamos um "context manager" (o decorador @contextmanager) para garantir
    que a conexão sempre seja fechada corretamente, mesmo se der algum erro
    no meio do caminho. É como usar "with open(...) as f:" para arquivos.

    Uso típico:
        with get_connection() as conn:
            conn.execute("SELECT ...")
    """
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    # Faz com que os resultados das consultas venham como dicionários
    # (na prática, sqlite3.Row), o que facilita muito o trabalho com Pandas.
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()  # Salva (confirma) as alterações no banco
    except Exception:
        conn.rollback()  # Se der erro, desfaz qualquer alteração parcial
        raise
    finally:
        conn.close()


def criar_tabela():
    """
    Cria a tabela "atendimentos" caso ela ainda não exista.

    Esta função deve ser chamada logo no início do app (app.py chama ela
    automaticamente), então você nunca precisa criar o banco manualmente.

    Tipos de dados usados:
    - INTEGER PRIMARY KEY AUTOINCREMENT -> gera o ID sozinho, sempre único
    - TEXT -> qualquer texto (nomes, datas em formato texto, etc.)
    """
    with get_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS atendimentos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                patrimonio_novo TEXT,
                patrimonio_antigo TEXT,
                serial_number TEXT,
                nome_usuario TEXT,
                secao_usuario TEXT,
                nome_gestor TEXT,
                tecnico_responsavel TEXT,
                tipo_atendimento TEXT,
                numero_predio TEXT,
                rua_dsi TEXT,
                data_agendada TEXT,
                prioridade_dia TEXT,
                status TEXT
            )
            """
        )


def inserir_atendimento(dados: dict):
    """
    Insere um único atendimento no banco de dados.

    Parâmetro:
        dados (dict): dicionário cujas chaves devem ser (pelo menos parte de)
                       COLUNAS_SEM_ID. Chaves ausentes serão gravadas como
                       string vazia.

    Exemplo de uso:
        inserir_atendimento({
            "patrimonio_novo": "12345",
            "patrimonio_antigo": "",
            "serial_number": "SN-001",
            "nome_usuario": "Maria Silva",
            "secao_usuario": "Financeiro",
            "nome_gestor": "João Souza",
            "tecnico_responsavel": "Carlos",
            "tipo_atendimento": "Máquina Nova",
            "numero_predio": "Prédio 12",
            "rua_dsi": "Rua A",
            "data_agendada": "2026-08-01",
            "prioridade_dia": "Sim",
            "status": "Pendente",
        })
    """
    valores = [str(dados.get(col, "") or "") for col in COLUNAS_SEM_ID]
    placeholders = ", ".join(["?"] * len(COLUNAS_SEM_ID))
    colunas_sql = ", ".join(COLUNAS_SEM_ID)

    with get_connection() as conn:
        conn.execute(
            f"INSERT INTO atendimentos ({colunas_sql}) VALUES ({placeholders})",
            valores,
        )


def inserir_varios_atendimentos(lista_de_dicionarios: list):
    """
    Insere vários atendimentos de uma vez (usado na importação de Excel).

    Parâmetro:
        lista_de_dicionarios (list): lista onde cada item é um dicionário
                                      no mesmo formato aceito por
                                      inserir_atendimento().

    Retorna:
        int: quantidade de registros inseridos com sucesso.
    """
    if not lista_de_dicionarios:
        return 0

    colunas_sql = ", ".join(COLUNAS_SEM_ID)
    placeholders = ", ".join(["?"] * len(COLUNAS_SEM_ID))

    linhas = []
    for dados in lista_de_dicionarios:
        linha = [str(dados.get(col, "") or "") for col in COLUNAS_SEM_ID]
        linhas.append(linha)

    with get_connection() as conn:
        conn.executemany(
            f"INSERT INTO atendimentos ({colunas_sql}) VALUES ({placeholders})",
            linhas,
        )
    return len(linhas)


def buscar_atendimentos(
    status: str = None,
    prioridade_dia: str = None,
    numero_predio: str = None,
    rua_dsi: str = None,
) -> pd.DataFrame:
    """
    Busca atendimentos no banco, aplicando filtros opcionais.

    Todos os parâmetros são opcionais. Quando um parâmetro não é informado
    (fica como None), o filtro correspondente não é aplicado.

    Parâmetros:
        status (str): "Pendente", "Feito" ou None (todos)
        prioridade_dia (str): "Sim", "Não" ou None (todos)
        numero_predio (str): valor exato do prédio ou None (todos)
        rua_dsi (str): valor exato da rua DSI ou None (todos)

    Retorna:
        pandas.DataFrame: tabela com os resultados, já pronta para ser
                           exibida no Streamlit ou exportada para Excel.
    """
    condicoes = []
    parametros = []

    if status and status != "Todos":
        condicoes.append("status = ?")
        parametros.append(status)

    if prioridade_dia and prioridade_dia != "Todos":
        condicoes.append("prioridade_dia = ?")
        parametros.append(prioridade_dia)

    if numero_predio and numero_predio != "Todos":
        condicoes.append("numero_predio = ?")
        parametros.append(numero_predio)

    if rua_dsi and rua_dsi != "Todos":
        condicoes.append("rua_dsi = ?")
        parametros.append(rua_dsi)

    query = f"SELECT {', '.join(COLUNAS)} FROM atendimentos"
    if condicoes:
        query += " WHERE " + " AND ".join(condicoes)
    query += " ORDER BY prioridade_dia DESC, data_agendada ASC, id ASC"

    with get_connection() as conn:
        df = pd.read_sql_query(query, conn, params=parametros)
    return df


def buscar_valores_unicos(coluna: str) -> list:
    """
    Retorna a lista de valores distintos já cadastrados em uma coluna.

    Usado para popular as caixas de seleção dos filtros (ex: lista de
    prédios já cadastrados, lista de ruas DSI já cadastradas), assim o
    usuário não precisa digitar, só escolher.
    """
    if coluna not in COLUNAS:
        raise ValueError(f"Coluna inválida: {coluna}")

    with get_connection() as conn:
        cursor = conn.execute(
            f"SELECT DISTINCT {coluna} FROM atendimentos "
            f"WHERE {coluna} IS NOT NULL AND {coluna} != '' "
            f"ORDER BY {coluna} ASC"
        )
        valores = [linha[0] for linha in cursor.fetchall()]
    return valores


def atualizar_status(id_atendimento: int, novo_status: str, tecnico_responsavel: str = None):
    """
    Atualiza o status de um atendimento (ex: de "Pendente" para "Feito")
    e, opcionalmente, atribui/atualiza o técnico responsável ao mesmo tempo.

    Parâmetros:
        id_atendimento (int): o ID do registro a ser atualizado
        novo_status (str): "Pendente" ou "Feito"
        tecnico_responsavel (str): nome do técnico (opcional). Se None,
                                    o técnico atual não é alterado.
    """
    with get_connection() as conn:
        if tecnico_responsavel is not None and tecnico_responsavel != "":
            conn.execute(
                "UPDATE atendimentos SET status = ?, tecnico_responsavel = ? WHERE id = ?",
                (novo_status, tecnico_responsavel, id_atendimento),
            )
        else:
            conn.execute(
                "UPDATE atendimentos SET status = ? WHERE id = ?",
                (novo_status, id_atendimento),
            )


def atualizar_atendimento(id_atendimento: int, dados: dict):
    """
    Atualiza todos os campos de um atendimento existente (edição completa).

    Parâmetros:
        id_atendimento (int): ID do registro a atualizar
        dados (dict): dicionário com os novos valores (mesmo formato de
                      inserir_atendimento)
    """
    colunas_sql = ", ".join([f"{col} = ?" for col in COLUNAS_SEM_ID])
    valores = [str(dados.get(col, "") or "") for col in COLUNAS_SEM_ID]
    valores.append(id_atendimento)

    with get_connection() as conn:
        conn.execute(
            f"UPDATE atendimentos SET {colunas_sql} WHERE id = ?",
            valores,
        )


def excluir_atendimento(id_atendimento: int):
    """Remove definitivamente um atendimento do banco de dados, pelo ID."""
    with get_connection() as conn:
        conn.execute("DELETE FROM atendimentos WHERE id = ?", (id_atendimento,))


def buscar_metricas() -> dict:
    """
    Calcula os números usados no painel de métricas do topo da tela:
    - total de atendimentos pendentes
    - total de atendimentos já feitos
    - total de atendimentos marcados como prioridade do dia (e ainda pendentes)
    - total geral de registros

    Retorna:
        dict: {"pendentes": int, "feitos": int, "prioridades_hoje": int, "total": int}
    """
    with get_connection() as conn:
        pendentes = conn.execute(
            "SELECT COUNT(*) FROM atendimentos WHERE status = 'Pendente'"
        ).fetchone()[0]

        feitos = conn.execute(
            "SELECT COUNT(*) FROM atendimentos WHERE status = 'Feito'"
        ).fetchone()[0]

        prioridades_hoje = conn.execute(
            "SELECT COUNT(*) FROM atendimentos WHERE prioridade_dia = 'Sim' AND status = 'Pendente'"
        ).fetchone()[0]

        total = conn.execute("SELECT COUNT(*) FROM atendimentos").fetchone()[0]

    return {
        "pendentes": pendentes,
        "feitos": feitos,
        "prioridades_hoje": prioridades_hoje,
        "total": total,
    }
