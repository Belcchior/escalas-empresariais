from datetime import datetime, timedelta, date, time  # importa tipos de data/hora
import calendar  # importa calendar para trabalhar com meses e semanas
import re  # importa re para tratar texto (gerar IDs de funcionários)

from estado import ESTADO_JSON, gravar_estado, carregar_estado
from engine import gerar_escala_semana
from export_excel import gerar_e_exportar_excel_4_semanas
from export_relatorio import (
    contar_horas_turno,
    analisar_necessidades_e_equipa,
    analisar_4_semanas,
    gerar_relatorio_docx,
    DOCX_DISPONIVEL,
)

# =========================
# CONSTANTES GERAIS
# =========================

DIAS_LONGOS = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]  # nomes completos dos dias
DIAS_CURTOS = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]  # nomes curtos dos dias

# nomes dos meses em português (para usar no cabeçalho da semana)
MESES_PT = [
    "janeiro", "fevereiro", "março", "abril",
    "maio", "junho", "julho", "agosto",
    "setembro", "outubro", "novembro", "dezembro"
]

def _mes_pt(d):
    """Devolve o nome do mês em português para uma data 'd'."""
    return MESES_PT[d.month - 1]

RELATORIO_DOCX = "relatorio_operacional.docx"  # nome do ficheiro de relatório em Word

# regras simplificadas do Código do Trabalho de Portugal
RULES_PT = {  # dicionário com regras trabalhistas básicas
    "max_horas_dia": 10.0,  # máximo de horas por dia por trabalhador
    "max_horas_semana": 48.0,  # máximo de horas por semana por trabalhador
    "min_descanso_horas": 11.0,  # descanso mínimo entre dias de trabalho
}

def clean_role(role: str) -> str:  # define função para "normalizar" o nome da função (role)
    role = role.strip()           # remove espaços em branco no início e no fim da string
    role = role.lower()           # converte todo o texto para minúsculas (para comparar de forma consistente)
    return role                   # devolve a string normalizada


# =========================
# FUNÇÕES DE UTILIDADE GERAL
# =========================

def ask_int(prompt, default=None, min_val=None, max_val=None):
    """
    Pergunta ao utilizador um número inteiro, com opções de valor por defeito e limites.
    """
    while True:  # repete até obter valor válido
        txt = input(prompt).strip()  # lê input e tira espaços
        if not txt:  # se o utilizador só carregou ENTER
            if default is not None:  # se temos valor por defeito
                return default  # devolve valor por defeito
            print("Por favor, introduza um número.")  # avisa que precisa de número
            continue  # volta ao início do while
        try:
            val = int(txt)  # tenta converter para inteiro
        except ValueError:
            print("Valor inválido. Introduza um número inteiro.")  # se não der, avisa
            continue
        if min_val is not None and val < min_val:  # verifica mínimo
            print(f"O valor mínimo é {min_val}.")  # avisa
            continue
        if max_val is not None and val > max_val:  # verifica máximo
            print(f"O valor máximo é {max_val}.")  # avisa
            continue
        return val  # devolve valor válido

def ask_yesno(prompt, default="n"):
    """
    Pergunta ao utilizador uma resposta sim/não.
    """
    default = default.lower()  # normaliza valor por defeito
    while True:
        txt = input(prompt + " [s/n]: ").strip().lower()  # lê input em minúsculas
        if not txt:
            txt = default  # se vazio, usa valor por defeito
        if txt in ("s", "sim"):
            return True
        if txt in ("n", "nao", "não"):
            return False
        print("Por favor, responda com 's' ou 'n'.")  # mensagem de erro

def proxima_segunda(a_partir_de: date) -> date:
    """
    Devolve a próxima segunda-feira a partir de uma data dada (pode ser a própria).
    """
    delta = (0 - a_partir_de.weekday()) % 7  # weekday: 0=segunda, 6=domingo
    if delta == 0:  # se já é segunda
        return a_partir_de  # devolve a própria
    return a_partir_de + timedelta(days=delta)  # devolve próxima segunda

def semana_datas(segunda: date):
    """
    Devolve lista de 7 datas (segunda a domingo) a partir de uma segunda-feira.
    """
    return [segunda + timedelta(days=i) for i in range(7)]  # cria lista de 7 dias

def _norm_hhmm(s: str) -> str:
    """
    Normaliza uma string no formato HH:MM, garantindo 2 dígitos.
    """
    s = s.strip()  # remove espaços
    if not s:  # se vazio
        raise ValueError("hora vazia")  # erro
    if s == "24:00":  # trata 24:00 como 23:59 para evitar problemas
        return "23:59"  # devolve 23:59 para evitar problemas
    h, m = map(int, s.split(":"))  # separa horas e minutos
    if not (0 <= h <= 23 and 0 <= m <= 59):  # verifica se está num intervalo válido
        raise ValueError("horário inválido")  # lança erro se for inválido
    return f"{h:02d}:{m:02d}"  # devolve string formatada com dois dígitos

def _parse_interval(txt: str):
    """Recebe 'HH:MM-HH:MM' (ou com '–') e devolve (ini, fim) normalizados."""
    txt = txt.strip().replace("–", "-")  # remove espaços e troca travessão por hífen
    if "-" not in txt:  # se não houver "-"
        raise ValueError("formato deve ser HH:MM-HH:MM")  # lança erro de formato
    ini, fim = txt.split("-", 1)  # separa no primeiro "-"
    return _norm_hhmm(ini), _norm_hhmm(fim)  # normaliza as duas partes

def _duration_ok(ini: str, fim: str) -> bool:
    """Verifica se um intervalo tem duração positiva."""
    h1, m1 = map(int, ini.split(":"))  # converte início
    h2, m2 = map(int, fim.split(":"))  # converte fim
    t1 = h1 * 60 + m1  # minutos totais início
    t2 = h2 * 60 + m2  # minutos totais fim
    return t2 > t1  # só é válido se fim for depois do início

def horas_entre(ini: str, fim: str) -> float:
    """Calcula duração em horas entre duas strings HH:MM."""
    h1, m1 = map(int, ini.split(":"))
    h2, m2 = map(int, fim.split(":"))
    return (h2 * 60 + m2 - (h1 * 60 + m1)) / 60.0  # diferença em minutos / 60

def dias_entre(d1: date, d2: date) -> int:
    """Número de dias entre duas datas."""
    return (d2 - d1).days

# =========================
# DEMO – PADARIA
# =========================

def estado_demo():
    """
    Devolve um estado DEMO (padaria), com:
      • turnos fixos 07:00–11:00 e 15:00–19:00
      • necessidades iguais todos os dias
      • 8 funcionários com restrições simuladas
    """
    horario = {
        "Segunda":  [("07:00", "11:00"), ("15:00", "19:00")],
        "Terça":    [("07:00", "11:00"), ("15:00", "19:00")],
        "Quarta":   [("07:00", "11:00"), ("15:00", "19:00")],
        "Quinta":   [("07:00", "11:00"), ("15:00", "19:00")],
        "Sexta":    [("07:00", "11:00"), ("15:00", "19:00")],
        "Sábado":   [("07:00", "11:00"), ("15:00", "19:00")],
        "Domingo":  [("07:00", "11:00"), ("15:00", "19:00")],
    }

    # funções da padaria
    roles = ["padeiro", "balconista", "garçom"]

    # necessidades fixas: manhã = [2,1,1], tarde = [1,1,2]
    necessidades = {
        dia: {
            "07:00–11:00": {"padeiro": 2, "balconista": 1, "garçom": 1},
            "15:00–19:00": {"padeiro": 1, "balconista": 1, "garçom": 2},
        }
        for dia in DIAS_LONGOS
    }

    # funcionários DEMO
    employees = {
        "ana": {
            "nome": "Ana Paula",
            "roles": ["padeiro", "garçom"],
            "folga_fixa": "Segunda",
            "restricoes": {"Quinta": ["15:00-19:00"]},
            "pedidos_folga": []
        },
        "leo": {
            "nome": "Leonardo",
            "roles": ["garçom"],
            "folga_fixa": "Sábado",
            "restricoes": {},
            "pedidos_folga": []
        },
        "elisa": {
            "nome": "Elisa",
            "roles": ["balconista"],
            "folga_fixa": "",
            "restricoes": {},
            "pedidos_folga": []
        },
        "lucas": {
            "nome": "Lucas",
            "roles": ["garçom", "balconista"],
            "folga_fixa": "Domingo",
            "restricoes": {"Terça": ["07:00-11:00"]},
            "pedidos_folga": []
        },
        "marcelo": {
            "nome": "Marcelo",
            "roles": ["padeiro"],
            "folga_fixa": "",
            "restricoes": {"Sexta": ["15:00-19:00"]},
            "pedidos_folga": []
        },
        "antonio": {
            "nome": "Antônio",
            "roles": ["garçom"],
            "folga_fixa": "",
            "restricoes": {},
            "pedidos_folga": []
        },
        "maria": {
            "nome": "Maria",
            "roles": ["balconista"],
            "folga_fixa": "",
            "restricoes": {},
            "pedidos_folga": []
        },
        "julio": {
            "nome": "Júlio",
            "roles": ["padeiro"],
            "folga_fixa": "",
            "restricoes": {},
            "pedidos_folga": []
        },
    }

    return {
        "legislacao": "PT",
        "horario": horario,
        "roles": roles,
        "necessidades": necessidades,
        "employees": employees,
        "rules": RULES_PT,
    }


# =========================
# ONBOARDING (configuração real do utilizador)
# =========================

def onboarding():
    """
    Pergunta ao utilizador:
      • legislação (por enquanto só PT)
      • horário da empresa
      • lista de funções
      • necessidades por turno
      • funcionários
    """
    print("\n=== CONFIGURAÇÃO INICIAL ===\n")

    print("Legislação disponível:")
    print("  [1] Portugal – Código do Trabalho\n")
    escolha_lex = ask_int("Escolha a legislação [1]: ", default=1)
    if escolha_lex != 1:
        escolha_lex = 1

    # horario semanal
    print("\n--- HORÁRIO DE FUNCIONAMENTO DA EMPRESA ---")
    print("Insira janelas no formato: 07:00-12:00, 15:00-20:00")
    print("Deixe vazio para indicar que a empresa não abre nesse dia.")
    print()

    horario = {}
    for dia in DIAS_LONGOS:
        txt = input(f"{dia} (ex.: 07:00-12:00, 15:00-20:00): ").strip()
        if not txt:
            horario[dia] = []
            continue
        janelas = []
        partes = txt.split(",")
        for p in partes:
            try:
                ini, fim = _parse_interval(p)
                if not _duration_ok(ini, fim):
                    raise ValueError
                janelas.append((ini, fim))
            except Exception:
                print(f"⚠️  Intervalo ignorado: {p}")
        horario[dia] = janelas

    # funções
    print("\n--- FUNÇÕES DA EMPRESA ---")
    roles = []
    while True:
        txt = input("Função (ENTER para terminar): ").strip()
        if not txt:
            break
        roles.append(txt.lower())

    # necessidades
    print("\n--- NECESSIDADES POR TURNO ---")
    necessidades = {dia: {} for dia in DIAS_LONGOS}
    for dia in DIAS_LONGOS:
        if not horario[dia]:
            continue
        print(f"\nDia: {dia}")
        for ini, fim in horario[dia]:
            turno = f"{ini}–{fim}"
            print(f"  Turno {turno}:")
            needs = {}
            for r in roles:
                v = ask_int(f"    Quantos '{r}'? (ENTER=0): ", default=0)
                needs[r] = v
            necessidades[dia][turno] = needs

    # funcionários
    print("\n--- FUNCIONÁRIOS ---")
    employees = {}
    while True:
        nome = input("Nome do funcionário (ENTER para terminar): ").strip()
        if not nome:
            break
        emp_id = nome.lower().replace(" ", "_")
        roles_f = input("Funções (separe por vírgula): ").strip().lower().split(",")
        roles_f = [r.strip() for r in roles_f if r.strip()]

        folga = input("Folga fixa semanal (Segunda/Domingo ou vazio): ").strip().title()

        restr = {}
        print("Restrições por dia (mesmo formato dos turnos ou ENTER):")
        for d in DIAS_LONGOS:
            rr = input(f"  {d}: ").strip()
            if rr:
                restr[d] = [p.strip() for p in rr.split(",")]

        pedidos = []
        print("Pedidos de folga (YYYY-MM-DD, ENTER p/ terminar):")
        while True:
            pf = input("  Pedido: ").strip()
            if not pf:
                break
            tipo = input("  Tipo (HARD/SOFT, default=SOFT): ").strip().upper()
            if tipo not in ("HARD", "SOFT"):
                tipo = "SOFT"
            pedidos.append({"data": pf, "tipo": tipo})

        employees[emp_id] = {
            "nome": nome,
            "roles": roles_f,
            "folga_fixa": folga,
            "restricoes": restr,
            "pedidos_folga": pedidos,
        }

    estado = {
        "legislacao": "PT",
        "horario": horario,
        "roles": roles,
        "necessidades": necessidades,
        "employees": employees,
        "rules": RULES_PT,
    }

    gravar_estado(estado)
    print("\nConfiguração gravada com sucesso!\n")
    return estado

def gerar_id_unico_funcionario(estado, nome):  # define função para gerar um ID único para o funcionário
    """
    Gera um emp_id único, baseado no nome do funcionário,
    evitando conflitos com IDs já existentes.
    """  # docstring explicativa
    employees = estado.get("employees", {})  # obtém o dicionário de funcionários já existentes
    base = re.sub(r"[^a-z0-9]", "", nome.lower())  # cria um "slug" minúsculo, removendo tudo que não for letra/número
    if not base:  # se depois da limpeza o nome ficou vazio
        base = "emp"  # usa um prefixo genérico "emp"
    emp_id = base  # começa tentando usar o slug diretamente como ID
    sufixo = 2  # inicia o contador de sufixo a partir de 2
    while emp_id in employees:  # enquanto o ID já existir na estrutura
        emp_id = f"{base}{sufixo}"  # cria um novo ID com sufixo numérico
        sufixo += 1  # incrementa o sufixo para a próxima tentativa
    return emp_id  # devolve um ID que não entra em conflito com os existentes


def listar_funcionarios(estado):  # define função para listar os funcionários da configuração atual
    """
    Lista os funcionários atuais da configuração, mostrando ID, nome,
    funções, folga fixa e eventuais limites semanais individuais.
    """  # docstring explicativa
    employees = estado.get("employees", {})  # obtém o dicionário de funcionários do estado
    if not employees:  # se não existir nenhum funcionário
        print("\n⚠️  Ainda não existem funcionários registados nesta configuração.\n")  # avisa o utilizador
        return  # termina a função sem listar nada

    print("\n=== FUNCIONÁRIOS ATUAIS ===")  # cabeçalho visual da listagem
    for idx, (emp_id, emp) in enumerate(employees.items(), start=1):  # percorre funcionários com um índice (1,2,3,...)
        nome = emp.get("nome", emp_id)  # obtém o nome do funcionário ou, em último caso, o próprio ID
        roles = emp.get("roles", [])  # obtém a lista de funções associadas ao funcionário
        folga = emp.get("folga_fixa", "")  # obtém a folga fixa semanal, se existir
        max_sem = emp.get("max_horas_semana_emp")  # obtém limite semanal individual, se definido
        roles_txt = ", ".join(roles) if roles else "(sem funções definidas)"  # formata texto das funções
        folga_txt = folga if folga else "(sem folga fixa)"  # formata texto da folga fixa
        if max_sem is not None:  # se existir um limite semanal individual
            limite_txt = f"{max_sem} h/semana"  # monta texto com limite em horas por semana
        else:  # caso não exista limite individual
            limite_txt = "usa limite geral das regras"  # indica que usa o limite definido em RULES_PT
        print(f"\n[{idx}] ID: {emp_id}")  # imprime índice e ID do funcionário
        print(f"     Nome: {nome}")  # imprime o nome do funcionário
        print(f"     Funções: {roles_txt}")  # imprime as funções do funcionário
        print(f"     Folga fixa: {folga_txt}")  # imprime a folga fixa
        print(f"     Limite semanal: {limite_txt}")  # imprime o limite de horas semanais
    print()  # imprime uma linha em branco no final da listagem para separar visualmente


def adicionar_funcionario(estado):  # define função para adicionar um novo funcionário à configuração
    """
    Adiciona um novo funcionário à configuração atual,
    pedindo nome, funções e folga fixa.
    """  # docstring explicativa
    employees = estado.setdefault("employees", {})  # garante que existe o dicionário "employees" no estado

    print("\n=== ADICIONAR NOVO FUNCIONÁRIO ===\n")  # título da seção de adição
    nome = input("Nome do funcionário (ENTER para cancelar): ").strip()  # lê o nome do funcionário e remove espaços
    if not nome:  # se o utilizador não escreveu nada
        print("Operação cancelada.\n")  # avisa que a operação foi cancelada
        return  # termina a função sem adicionar ninguém

    roles_disponiveis = estado.get("roles", [])  # obtém a lista de funções existentes na empresa
    if not roles_disponiveis:  # se não houver nenhuma função registada
        print("⚠️  Não existem funções definidas na configuração. Crie funções primeiro no onboarding.\n")  # avisa
        return  # termina a função, pois não faz sentido criar funcionário sem funções possíveis

    print("\nFunções disponíveis na empresa:")  # título para mostrar funções disponíveis
    print("  " + ", ".join(roles_disponiveis))  # imprime a lista de funções separadas por vírgulas

    txt_roles = input("Funções deste funcionário (separe por vírgula): ").strip().lower()  # lê funções do funcionário
    roles_f = [r.strip() for r in txt_roles.split(",") if r.strip()]  # cria lista de funções sem espaços vazios
    roles_f_validas = [r for r in roles_f if r in roles_disponiveis]  # filtra apenas as funções que existem na empresa
    if not roles_f_validas:  # se nenhuma das funções inseridas for válida
        print("⚠️  Nenhuma função válida foi indicada. Funcionário não será criado.\n")  # avisa o utilizador
        return  # termina a função sem criar o funcionário

    print("\nDias da semana:")  # cabeçalho para explicar as opções de folga fixa
    for d in DIAS_LONGOS:  # percorre os nomes completos dos dias
        print(f"  - {d}")  # imprime cada dia em uma linha
    folga = input("Folga fixa semanal (digite o dia exatamente como acima ou deixe vazio): ").strip().title()  # lê folga fixa
    if folga and folga not in DIAS_LONGOS:  # se o utilizador escreveu algo que não corresponde a um dia válido
        print("⚠️  Dia de folga inválido. Não será registada folga fixa.\n")  # avisa que a folga é inválida
        folga = ""  # limpa o valor da folga fixa

    emp_id = gerar_id_unico_funcionario(estado, nome)  # gera um ID único para o novo funcionário
    employees[emp_id] = {  # cria a entrada do novo funcionário no dicionário de funcionários
        "nome": nome,  # guarda o nome do funcionário
        "roles": roles_f_validas,  # guarda apenas as funções válidas associadas ao funcionário
        "folga_fixa": folga,  # guarda a folga fixa semanal (ou vazio se não houver)
        "restricoes": {},  # inicia sem restrições de horário específicas
        "pedidos_folga": [],  # inicia sem pedidos de folga registados
    }

    print(f"\n✅ Funcionário '{nome}' criado com ID '{emp_id}'.\n")  # confirma ao utilizador a criação do funcionário

def escolher_funcionario(estado):  # define função para escolher um funcionário da lista atual
    """
    Permite ao utilizador escolher um funcionário existente,
    devolvendo o emp_id escolhido ou None se não houver escolha.
    """  # docstring explicativa
    employees = estado.get("employees", {})  # obtém o dicionário de funcionários existentes
    if not employees:  # se não existir nenhum funcionário registado
        print("\n⚠️  Ainda não existem funcionários registados nesta configuração.\n")  # informa o utilizador
        return None  # devolve None indicando que não há quem escolher

    lista_ids = list(employees.keys())  # cria uma lista com os IDs dos funcionários para indexação
    print("\n=== SELECIONAR FUNCIONÁRIO ===")  # cabeçalho visual da seleção
    for idx, emp_id in enumerate(lista_ids, start=1):  # percorre cada ID com um índice (1,2,3,...)
        emp = employees[emp_id]  # obtém o dicionário de dados desse funcionário
        nome = emp.get("nome", emp_id)  # obtém o nome ou, em último caso, o próprio ID
        print(f"  [{idx}] {nome} (ID: {emp_id})")  # imprime opção com índice, nome e ID

    escolha = ask_int("Escolha o número do funcionário (0 para cancelar): ", default=0, min_val=0, max_val=len(lista_ids))  # lê a escolha do utilizador
    if escolha == 0:  # se o utilizador escolher 0
        print("Operação cancelada.\n")  # informa que a operação foi cancelada
        return None  # devolve None
    emp_id_escolhido = lista_ids[escolha - 1]  # calcula o emp_id correspondente ao índice escolhido
    return emp_id_escolhido  # devolve o emp_id escolhido


def editar_funcionario(estado):  # define função para editar os dados de um funcionário existente
    """
    Permite editar nome, funções, folga fixa e limite semanal de um funcionário.
    """  # docstring explicativa
    employees = estado.get("employees", {})  # obtém o dicionário de funcionários do estado
    emp_id = escolher_funcionario(estado)  # usa a função auxiliar para escolher um funcionário
    if emp_id is None:  # se o utilizador cancelou a seleção
        return  # termina a função sem alterações

    emp = employees[emp_id]  # obtém o dicionário de dados do funcionário escolhido
    while True:  # laço para permitir múltiplas edições antes de sair
        nome = emp.get("nome", emp_id)  # obtém o nome atual do funcionário
        roles = emp.get("roles", [])  # obtém a lista de funções atuais do funcionário
        folga = emp.get("folga_fixa", "")  # obtém a folga fixa atual
        max_sem = emp.get("max_horas_semana_emp")  # obtém limite semanal individual, se existir

        roles_txt = ", ".join(roles) if roles else "(sem funções definidas)"  # texto amigável com as funções
        folga_txt = folga if folga else "(sem folga fixa)"  # texto amigável com a folga fixa
        if max_sem is not None:  # se existe limite semanal individual
            limite_txt = f"{max_sem} h/semana"  # texto com limite
        else:  # se não existe limite individual
            limite_txt = "usa limite geral das regras"  # texto indicando uso do limite geral

        print("\n=== EDITAR FUNCIONÁRIO ===")  # cabeçalho do menu de edição
        print(f"Funcionário selecionado: {nome} (ID: {emp_id})")  # mostra quem está a ser editado
        print(f"  Funções atuais: {roles_txt}")  # mostra funções atuais
        print(f"  Folga fixa atual: {folga_txt}")  # mostra folga fixa atual
        print(f"  Limite semanal atual: {limite_txt}")  # mostra limite semanal atual
        print("\nO que deseja editar?")  # pergunta próxima ação
        print("  [1] Nome")  # opção para editar o nome
        print("  [2] Funções")  # opção para editar funções
        print("  [3] Folga fixa")  # opção para editar folga fixa
        print("  [4] Definir/alterar limite semanal individual")  # opção para definir ou alterar limite semanal
        print("  [5] Remover limite semanal individual (usar regra geral)")  # opção para limpar limite individual
        print("  [6] Voltar")  # opção para sair do menu de edição deste funcionário

        escolha = ask_int("Escolha uma opção [6]: ", default=6, min_val=1, max_val=6)  # lê a escolha do utilizador

        if escolha == 1:  # editar nome
            novo_nome = input("Novo nome (ENTER para manter o atual): ").strip()  # lê o novo nome
            if novo_nome:  # se o utilizador escreveu algo
                emp["nome"] = novo_nome  # atualiza o nome no dicionário
                print("✅ Nome atualizado.\n")  # confirma alteração
        elif escolha == 2:  # editar funções
            roles_disponiveis = estado.get("roles", [])  # obtém funções possíveis da empresa
            if not roles_disponiveis:  # se não houver funções registadas
                print("⚠️  Não existem funções definidas na configuração. Não é possível atualizar funções.\n")  # avisa
            else:  # se houver funções
                print("\nFunções disponíveis na empresa:")  # título da lista de funções
                print("  " + ", ".join(roles_disponiveis))  # imprime as funções disponíveis
                txt_roles = input("Novas funções deste funcionário (separe por vírgula, ENTER para cancelar): ").strip().lower()  # lê as novas funções
                if txt_roles:  # se o utilizador escreveu algo
                    roles_f = [r.strip() for r in txt_roles.split(",") if r.strip()]  # separa por vírgula e limpa espaços
                    roles_f_validas = [r for r in roles_f if r in roles_disponiveis]  # filtra só funções válidas
                    if roles_f_validas:  # se existe pelo menos uma função válida
                        emp["roles"] = roles_f_validas  # atualiza a lista de funções do funcionário
                        print("✅ Funções atualizadas.\n")  # confirma alteração
                    else:  # se nenhuma função indicada for válida
                        print("⚠️  Nenhuma função válida foi indicada. As funções não foram alteradas.\n")  # avisa
        elif escolha == 3:  # editar folga fixa
            print("\nDias da semana:")  # cabeçalho para os dias válidos
            for d in DIAS_LONGOS:  # percorre os dias da semana
                print(f"  - {d}")  # imprime cada dia
            nova_folga = input("Nova folga fixa semanal (digite o dia exatamente como acima ou ENTER para remover): ").strip().title()  # lê nova folga
            if not nova_folga:  # se o utilizador só deu ENTER
                emp["folga_fixa"] = ""  # remove a folga fixa
                print("✅ Folga fixa removida.\n")  # confirma alteração
            elif nova_folga in DIAS_LONGOS:  # se o valor corresponde a um dia válido
                emp["folga_fixa"] = nova_folga  # atualiza folga fixa
                print("✅ Folga fixa atualizada.\n")  # confirma alteração
            else:  # se o valor não corresponder a um dia válido
                print("⚠️  Dia de folga inválido. A folga fixa não foi alterada.\n")  # avisa
        elif escolha == 4:  # definir ou alterar limite semanal individual
            max_h = ask_int("Novo limite semanal de horas para este funcionário (por exemplo 40): ", min_val=1, max_val=80)  # lê um novo limite
            emp["max_horas_semana_emp"] = max_h  # atualiza o limite semanal individual
            print("✅ Limite semanal individual atualizado.\n")  # confirma alteração
        elif escolha == 5:  # remover limite semanal individual
            if "max_horas_semana_emp" in emp:  # se existir essa chave
                del emp["max_horas_semana_emp"]  # remove a chave do dicionário
            print("✅ Limite semanal individual removido. Será usado o limite geral das regras.\n")  # confirma
        else:  # opção 6: voltar
            break  # sai do laço de edição para este funcionário

    employees[emp_id] = emp  # garante que o dicionário de funcionários esteja atualizado com as alterações

def remover_funcionario(estado):  # define função para remover um funcionário da configuração
    """
    Remove um funcionário da configuração atual,
    após confirmação do utilizador.
    """  # docstring explicativa
    employees = estado.get("employees", {})  # obtém o dicionário de funcionários do estado
    if not employees:  # se não existir nenhum funcionário registado
        print("\n⚠️  Ainda não existem funcionários registados nesta configuração.\n")  # informa o utilizador
        return  # termina a função

    emp_id = escolher_funcionario(estado)  # usa função auxiliar para o utilizador escolher quem remover
    if emp_id is None:  # se o utilizador cancelou a operação
        return  # termina a função sem remover ninguém

    emp = employees.get(emp_id, {})  # obtém os dados do funcionário escolhido
    nome = emp.get("nome", emp_id)  # obtém o nome do funcionário ou o próprio ID
    confirmar = ask_yesno(f"Tem a certeza que deseja remover o funcionário '{nome}' (ID: {emp_id})? [s/n]: ")  # pede confirmação
    if not confirmar:  # se o utilizador não confirmar
        print("Operação cancelada. O funcionário não foi removido.\n")  # informa que não houve remoção
        return  # termina a função
    del employees[emp_id]  # remove a entrada correspondente a esse funcionário no dicionário
    print(f"✅ Funcionário '{nome}' removido com sucesso.\n")  # confirma a remoção ao utilizador

def menu_editar_funcionarios(estado):  # define menu específico para gerir funcionários
    """
    Menu de gestão de funcionários: listar, adicionar, editar e remover.
    """  # docstring explicativa
    while True:  # laço para manter o menu de funcionários até o utilizador escolher sair
        print("\n=== GESTÃO DE FUNCIONÁRIOS ===")  # título do menu de funcionários
        print("  [1] Listar funcionários")  # opção para listar todos os funcionários
        print("  [2] Adicionar novo funcionário")  # opção para adicionar um novo funcionário
        print("  [3] Editar funcionário existente")  # opção para editar um funcionário
        print("  [4] Remover funcionário")  # opção para remover um funcionário
        print("  [5] Voltar ao menu anterior")  # opção para sair do menu de funcionários
        escolha = ask_int("Escolha uma opção [5]: ", default=5, min_val=1, max_val=5)  # lê a opção escolhida

        if escolha == 1:  # se a opção for listar
            listar_funcionarios(estado)  # chama a função que lista os funcionários
        elif escolha == 2:  # se a opção for adicionar
            adicionar_funcionario(estado)  # chama a função que adiciona um novo funcionário
        elif escolha == 3:  # se a opção for editar
            editar_funcionario(estado)  # chama a função que edita um funcionário existente
        elif escolha == 4:  # se a opção for remover
            remover_funcionario(estado)  # chama a função que remove um funcionário
        else:  # qualquer outra opção (aqui será o 5) volta ao menu anterior
            break  # sai do laço e retorna ao menu de edição de configuração

def menu_editar_configuracao(estado):  # define menu geral de edição da configuração
    """
    Menu principal de edição da configuração atual.
    Por enquanto, apenas gestão de funcionários.
    """  # docstring explicativa
    while True:  # laço para manter o menu de edição até o utilizador escolher sair
        print("\n=== EDIÇÃO DA CONFIGURAÇÃO ATUAL ===")  # título do menu de edição
        print("  [1] Gerir funcionários")  # opção para entrar no menu de funcionários
        print("  [2] Voltar (terminar edição)")  # opção para sair da edição
        escolha = ask_int("Escolha uma opção [2]: ", default=2, min_val=1, max_val=2)  # lê escolha do utilizador

        if escolha == 1:  # se o utilizador escolher gerir funcionários
            menu_editar_funcionarios(estado)  # entra no submenu de gestão de funcionários
        else:  # se escolher voltar
            break  # sai do laço e termina o menu de edição

    gravar_estado(estado)  # após possíveis alterações, grava o estado atualizado no ficheiro JSON
    print("\n✅ Configuração atualizada e gravada com sucesso.\n")  # informa que a configuração foi gravada
    return estado  # devolve o estado (já potencialmente alterado)

# =========================================================
# ========  EXECUÇÃO COMPLETA – 4 SEMANAS  =================
# =========================================================

def executar_4_semanas_com_relatorios(estado):
    """
    Gera a escala para um MÊS específico (modo calendário mensal),
    construindo tantas semanas (segunda–domingo) quanto forem necessárias,
    e produzindo Excel + relatório DOCX para esse período.
    """
    hoje = date.today()

    print("\n➡️  Configuração do período (modo calendário mensal)")

    # Pergunta ao utilizador para que ano/mês quer gerar
    ano = ask_int(
        "Ano (YYYY) [Enter para ano atual]: ",
        default=hoje.year,
        min_val=2000,
        max_val=2100,
    )
    mes = ask_int(
        "Mês [1-12] (Enter para mês atual]: ",
        default=hoje.month,
        min_val=1,
        max_val=12,
    )

    # Primeiro dia do mês escolhido
    primeiro_dia_mes = date(ano, mes, 1)

    # Segunda-feira da semana que contém o dia 1
    offset_inicio = primeiro_dia_mes.weekday()  # 0 = segunda, 6 = domingo
    primeira_segunda = primeiro_dia_mes - timedelta(days=offset_inicio)

    print(
        f"\n➡️  A gerar escala para {mes:02d}/{ano} "
        f"(semanas completas de {primeira_segunda.strftime('%d/%m/%Y')} em diante)."
    )

    # Determinar até onde o calendário desse mês vai (último domingo)
    _, last_day_num = calendar.monthrange(ano, mes)
    ultimo_dia_mes = date(ano, mes, last_day_num)
    offset_fim = 6 - ultimo_dia_mes.weekday()  # 6 = domingo
    ultimo_domingo_mes = ultimo_dia_mes + timedelta(days=offset_fim)

    num_dias = (ultimo_domingo_mes - primeira_segunda).days + 1
    num_semanas = num_dias // 7  # 4 ou 5 semanas, consoante o mês

    # 1) Gera TODAS as semanas do período em memória (para análise)
    escala_4_semanas = []
    for i in range(num_semanas):
        segunda_semana = primeira_segunda + timedelta(days=7 * i)
        esc_sem = gerar_escala_semana(estado, segunda_semana)
        escala_4_semanas.append(esc_sem)

    # 2) Gera o Excel (usa a função que escreve a planilha toda)
    gerar_e_exportar_excel_4_semanas(estado, primeira_segunda, ano, mes)

    # 3) Faz análise geral da operação (horas, défices, etc.)
    analise = analisar_necessidades_e_equipa(estado, escala_4_semanas)

    # 4) Gera relatório (DOCX) usando a mesma escala gerada
    semanas, metrics = analisar_4_semanas(estado, escala_4_semanas, primeira_segunda, ano, mes)

    if DOCX_DISPONIVEL:
        gerar_relatorio_docx(estado, semanas, metrics)
    else:
        print("❗ python-docx não está instalado; relatório DOCX não foi gerado.")

    # 5) Informação final de intervalo (mês lógico da escala)
    _, last_day_num = calendar.monthrange(ano, mes)
    inicio_geral = date(ano, mes, 1)
    fim_geral = date(ano, mes, last_day_num)

    print(f"\nPeríodo total gerado: {inicio_geral.strftime('%d/%m/%Y')} até {fim_geral.strftime('%d/%m/%Y')}")
    print("✅ Processo concluído com sucesso!\n")

# =========================================================
# =========================  MAIN  ========================
# =========================================================

if __name__ == "__main__":
    print("ESCALA – legislação → horários → funções → necessidades → funcionários → Excel + Relatórios")
    print("========================================================================\n")

    # tenta carregar config gravada (caso o utilizador já tenha feito onboarding antes)
    estado_salvo = carregar_estado()

    # se não existir configuração gravada, tratamos como primeira utilização
    if estado_salvo is None:
        print("Parece ser a primeira utilização deste programa.\n")
        print("Como deseja gerar a escala?")
        print("  [1] Usar modo DEMO (padaria de exemplo)")
        print("  [2] Fazer configuração completa da sua empresa (onboarding)")
        escolha = ask_int("Escolha uma opção [1]: ", default=1, min_val=1, max_val=2)

        if escolha == 1:
            print("\n➡️ A usar modo DEMO (padaria).\n")
            estado = estado_demo()
        else:
            print("\n➡️ A configurar empresa pela primeira vez.\n")
            estado = onboarding()
            # depois de concluir o onboarding inicial, guardamos a configuração
            gravar_estado(estado)
    else:
        # já existe uma configuração guardada – damos opções ao utilizador
        print("Foi encontrada uma configuração gravada anteriormente.\n")  # informa que há config guardada
        print("O que deseja fazer agora?")  # pergunta pela ação desejada
        print("  [1] Usar a configuração gravada e gerar a escala mensal")  # opção para usar config como está
        print("  [2] Editar a configuração gravada (funcionários, etc.)")  # opção para entrar no menu de edição
        print("  [3] Refazer a configuração (novo onboarding)")  # opção para descartar e refazer tudo
        print("  [4] Ignorar configuração gravada e usar modo DEMO")  # opção para usar apenas o modo DEMO
        escolha = ask_int("Escolha uma opção [1]: ", default=1, min_val=1, max_val=4)  # lê a escolha do utilizador

        if escolha == 1:  # se escolher usar a configuração como está
            print("\n➡️ A usar configuração gravada.\n")  # mensagem de confirmação
            estado = estado_salvo  # usa o estado carregado do ficheiro
        elif escolha == 2:  # se escolher editar a configuração atual
            print("\n➡️ A editar configuração gravada.\n")  # mensagem de confirmação
            estado = menu_editar_configuracao(estado_salvo)  # entra no menu de edição e obtém o estado possivelmente atualizado
        elif escolha == 3:  # se escolher refazer completamente a configuração
            print("\n➡️ A reconfigurar empresa (novo onboarding).\n")  # mensagem de confirmação
            estado = onboarding()  # executa o onboarding completo novamente
            gravar_estado(estado)  # grava a nova configuração por cima da anterior
        else:  # qualquer outra opção (aqui será o 4) usa o modo DEMO
            print("\n➡️ A usar modo DEMO (padaria), sem alterar a configuração gravada.\n")  # mensagem de confirmação
            estado = estado_demo()  # carrega o estado DEMO

    executar_4_semanas_com_relatorios(estado)
