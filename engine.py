from datetime import date, timedelta  # importa tipos de data e o delta de dias para cálculos de datas
from collections import defaultdict   # importa defaultdict para facilitar estruturas com valores padrão

# Lista de nomes completos dos dias da semana em português (índice 0 = Segunda, 6 = Domingo)
DIAS_LONGOS = [                      # começa a lista com os dias da semana por extenso
    "Segunda",                       # índice 0 → segunda-feira
    "Terça",                         # índice 1 → terça-feira
    "Quarta",                        # índice 2 → quarta-feira
    "Quinta",                        # índice 3 → quinta-feira
    "Sexta",                         # índice 4 → sexta-feira
    "Sábado",                        # índice 5 → sábado
    "Domingo",                       # índice 6 → domingo
]

# Regras simplificadas do Código do Trabalho de Portugal (usadas como fallback)
RULES_PT = {                         # dicionário com valores padrão de regras trabalhistas
    "max_horas_dia": 10.0,          # máximo de horas que um colaborador pode trabalhar num único dia
    "max_horas_semana": 48.0,       # máximo de horas que um colaborador pode trabalhar por semana
    "min_descanso_horas": 11.0,     # descanso mínimo (em horas) entre o fim de um dia de trabalho e o início do seguinte
}

def semana_datas(segunda: date):    # define função que recebe uma data que é segunda-feira
    """
    Devolve lista de 7 datas (segunda a domingo) a partir de uma segunda-feira.
    """                               # docstring explicando o objetivo da função
    return [                          # devolve uma lista de datas
        segunda + timedelta(days=i)   # soma i dias à segunda-feira, gerando segunda..domingo
        for i in range(7)             # i vai de 0 a 6 (7 dias)
    ]

def horas_entre(ini: str, fim: str) -> float:
    """Calcula duração em horas entre duas strings HH:MM."""  # docstring explicando a função
    h1, m1 = map(int, ini.split(":"))  # separa a hora e minuto da hora inicial e converte para inteiros
    h2, m2 = map(int, fim.split(":"))  # separa a hora e minuto da hora final e converte para inteiros
    minutos_ini = h1 * 60 + m1         # converte a hora inicial para minutos totais
    minutos_fim = h2 * 60 + m2         # converte a hora final para minutos totais
    diff_min = minutos_fim - minutos_ini  # calcula a diferença de minutos entre fim e início
    return diff_min / 60.0            # devolve a diferença convertida para horas (float)

def formatar_turno(ini: str, fim: str) -> str:
    """Monta a chave padrao do turno no formato HH:MM–HH:MM."""
    return f"{ini}–{fim}"

def separar_turno(turno: str):
    """Aceita turnos com travessao ou hifen e devolve inicio/fim."""
    if "–" in turno:
        return turno.split("–", 1)
    return turno.split("-", 1)

def buscar_necessidades_turno(necessidades_dia, turno: str):
    """Busca necessidades aceitando HH:MM–HH:MM e HH:MM-HH:MM."""
    if turno in necessidades_dia:
        return necessidades_dia[turno]
    turno_com_hifen = turno.replace("–", "-")
    if turno_com_hifen in necessidades_dia:
        return necessidades_dia[turno_com_hifen]
    return {}

def funcionario_disponivel(emp, dia_semana, turno):  # define função que verifica se o colaborador pode trabalhar num dia/turno
    """
    Verifica se o funcionário pode trabalhar no dia+turno:
      • não está de folga fixa
      • não pediu folga HARD
      • não tem restrição para aquele dia/turno
    """                                   # docstring explicando as regras de disponibilidade
    folga = emp.get("folga_fixa", "")     # lê a folga fixa do colaborador (se existir)
    if folga and folga == dia_semana:     # se o colaborador tem folga fixa e coincide com o dia em questão
        return False                      # não está disponível

    for pf in emp.get("pedidos_folga", []):       # percorre a lista de pedidos de folga do colaborador
        if pf["tipo"] == "HARD" and pf["data"] == str(dia_semana):  # se for um pedido HARD no dia em questão
            return False                      # não está disponível

    restr = emp.get("restricoes", {})        # obtém o dicionário de restrições de horário por dia da semana
    restr_dia = restr.get(dia_semana, [])    # obtém lista de intervalos de restrição para o dia em questão
    if restr_dia:                             # se existirem restrições naquele dia
        ini_turno, fim_turno = separar_turno(turno)  # separa o turno "HH:MM–HH:MM" em início e fim
        h1, m1 = map(int, ini_turno.split(":"))  # separa hora e minuto do início e converte para inteiros
        h2, m2 = map(int, fim_turno.split(":"))  # separa hora e minuto do fim e converte para inteiros
        t1 = h1 * 60 + m1                        # converte início para minutos totais
        t2 = h2 * 60 + m2                        # converte fim para minutos totais

        for intervalo in restr_dia:              # percorre cada intervalo restrito daquele dia
            ri, rf = intervalo.split("-")        # separa o intervalo "HH:MM-HH:MM" em hora inicial e final
            rh1, rm1 = map(int, ri.split(":"))   # separa hora e minuto inicial da restrição
            rh2, rm2 = map(int, rf.split(":"))   # separa hora e minuto final da restrição
            rt1 = rh1 * 60 + rm1                 # converte início da restrição para minutos totais
            rt2 = rh2 * 60 + rm2                 # converte fim da restrição para minutos totais

            if not (t2 <= rt1 or t1 >= rt2):     # verifica se o turno sobrepõe o intervalo restrito
                return False                     # se houver sobreposição, o colaborador não está disponível

    return True                                  # se passou por todas as verificações, está disponível

def gerar_escala_semana(estado, segunda):  # função principal do motor que monta a escala de uma semana
    """
    Constrói a escala (pré-formatada) para uma semana inteira (segunda–domingo).
    Devolve um dicionário:
        { date: {turno: [ (emp_id, role), ... ] } }
    Preenche VAGA ABERTA quando não houver funcionários suficientes.

    Respeita as regras legais básicas (horas máximas por dia/semana
    e descanso mínimo entre turnos) definidas em estado["rules"],
    e tenta distribuir as horas de forma equilibrada entre os colaboradores.
    """                                           # docstring explicando o comportamento geral
    horario = estado["horario"]                  # lê do estado os horários de funcionamento (janelas de turnos por dia)
    necessidades = estado["necessidades"]        # lê do estado as necessidades por função/turno/dia
    employees = estado["employees"]              # lê do estado o dicionário de colaboradores
    rules = estado.get("rules", RULES_PT)        # lê as regras legais do estado ou usa as regras padrão PT

    max_horas_dia = rules.get("max_horas_dia", 9999.0)         # obtém limite de horas por dia (fallback grande se não existir)
    max_horas_semana = rules.get("max_horas_semana", 9999.0)   # obtém limite de horas por semana (fallback grande se não existir)
    min_descanso_horas = rules.get("min_descanso_horas", 0.0)  # obtém descanso mínimo entre dias (em horas)

    datas = semana_datas(segunda)                              # gera lista de datas (segunda..domingo) para a semana
    escala = {d: {} for d in datas}                           # cria estrutura base da escala: para cada dia, um dicionário de turnos

    horas_semana = {emp_id: 0.0 for emp_id in employees.keys()}        # inicializa acumulador de horas semanais por colaborador
    horas_dia = {d: {emp_id: 0.0 for emp_id in employees.keys()}       # inicializa acumulador de horas diárias por colaborador
                 for d in datas}
    ultimo_fim = {emp_id: None for emp_id in employees.keys()}        # guarda o último fim de turno (data, "HH:MM") de cada colaborador

    def _diff_horas_entre_turnos(emp_id, data_turno, ini_str):        # função interna para calcular descanso desde o último turno
        """
        Devolve, em horas, o intervalo desde o último fim de turno deste colaborador
        até ao início do turno atual. Se não houver histórico, devolve None.
        """                                       # docstring explicando a função interna
        last = ultimo_fim.get(emp_id)            # obtém o último registo de fim de turno do colaborador
        if not last:                             # se ainda não há nenhum turno registado
            return None                          # devolve None (não aplica regra de descanso)
        last_date, last_fim = last               # desempacota data e hora final do último turno
        if data_turno == last_date:             # se o último turno foi no mesmo dia deste turno
            return None                          # não aplicamos regra de descanso mínimo entre turnos do mesmo dia

        def _to_minutes(d, hhmm):               # função interna para converter data + "HH:MM" em minutos absolutos
            h, m = map(int, hhmm.split(":"))     # separa hora e minuto da string "HH:MM"
            return d.toordinal() * 24 * 60 + h * 60 + m  # converte dia para minutos e soma hora/minuto em minutos

        inicio_atual_min = _to_minutes(data_turno, ini_str)  # calcula minutos totais do início do turno atual
        fim_anterior_min = _to_minutes(last_date, last_fim)  # calcula minutos totais do fim do turno anterior
        return (inicio_atual_min - fim_anterior_min) / 60.0  # devolve diferença em horas entre os dois momentos

    for dt in datas:                                 # percorre cada dia (data) da semana
        dia_nome = DIAS_LONGOS[dt.weekday()]         # calcula o nome do dia (Segunda, Terça, ...) a partir do índice
        janelas = horario.get(dia_nome, [])          # obtém a lista de turnos configurados para esse dia
        for ini, fim in janelas:                     # percorre cada turno configurado (início, fim)
            turno = formatar_turno(ini, fim)         # monta a string do turno no formato "HH:MM–HH:MM"
            escala[dt][turno] = []                   # inicializa lista de alocações para este dia/turno

            dur_turno = horas_entre(ini, fim)        # calcula quantas horas este turno representa

            need = buscar_necessidades_turno(necessidades.get(dia_nome, {}), turno)  # obtém necessidades do turno
            for role, qt in need.items():            # percorre cada função e a quantidade necessária
                if qt <= 0:                          # se a quantidade for zero ou negativa
                    continue                         # não há nada a alocar para esta função

                candidatos = []                      # lista que vai guardar os colaboradores elegíveis para este slot
                for emp_id, emp in employees.items():  # percorre todos os colaboradores registados
                    if role not in emp.get("roles", []):             # se o colaborador não tem esta função
                        continue                                     # ignora este colaborador
                    if not funcionario_disponivel(emp, dia_nome, turno):  # verifica folgas fixas, HARD e restrições
                        continue                                     # se não estiver disponível, ignora

                    max_sem_emp = float(emp.get("max_horas_semana_emp", max_horas_semana))  # lê limite semanal individual (ou o geral)
                    horas_sem_emp = horas_semana.get(emp_id, 0.0)                            # lê horas já trabalhadas na semana
                    if horas_sem_emp + dur_turno > max_sem_emp:                              # se ultrapassar o limite semanal
                        continue                                                             # não pode trabalhar mais

                    horas_dia_emp = horas_dia[dt].get(emp_id, 0.0)                           # lê horas no dia atual
                    if horas_dia_emp + dur_turno > max_horas_dia:                            # se ultrapassar limite diário
                        continue                                                             # não pode trabalhar este turno

                    diff_descanso = _diff_horas_entre_turnos(emp_id, dt, ini)               # calcula descanso desde último fim de turno
                    if diff_descanso is not None and diff_descanso < min_descanso_horas:    # se existe histórico e descanso insuficiente
                        continue                                                             # não pode ser escalado neste turno

                    candidatos.append(emp_id)                                               # se passou por todas as regras, é candidato

                candidatos.sort(                                                            # ordena candidatos para favorecer quem trabalhou menos
                    key=lambda cid: (horas_semana.get(cid, 0.0), cid)                       # primeiro por horas semanais, depois por id
                )

                alocados = min(len(candidatos), qt)                                         # determina quantos colaboradores consegue alocar
                for i in range(alocados):                                                   # para cada colaborador a ser alocado
                    emp_id = candidatos[i]                                                  # pega o id do colaborador escolhido
                    escala[dt][turno].append((emp_id, role))                                # regista (emp_id, função) no turno

                    horas_semana[emp_id] += dur_turno                                       # atualiza horas semanais do colaborador
                    horas_dia[dt][emp_id] += dur_turno                                      # atualiza horas diárias do colaborador
                    ultimo_fim[emp_id] = (dt, fim)                                          # regista data e hora de fim deste turno

                if alocados < qt:                                                           # se ainda faltam vagas para este role
                    falta = qt - alocados                                                   # calcula quantas vagas ficaram por preencher
                    for _ in range(falta):                                                  # repete para cada vaga em falta
                        escala[dt][turno].append(f"VAGA ABERTA (...{role})")               # marca vaga aberta na escala com role entre parênteses

    return escala                                                                          # devolve o dicionário completo da escala semanal
