import json
from pathlib import Path

# Nome do ficheiro onde o estado (onboarding/configuração) vai ser guardado
ESTADO_JSON = "estado_escala.json"


def gravar_estado(estado):
    """
    Recebe o dicionário 'estado' (roles, employees, horários, regras, etc.)
    e grava num ficheiro JSON para uso futuro.
    """
    try:
        with open(ESTADO_JSON, "w", encoding="utf-8") as f:
            json.dump(estado, f, ensure_ascii=False, indent=2)
        print(f"💾 Estado gravado em {ESTADO_JSON}")
    except Exception as e:
        print(f"⚠️ Erro ao gravar estado em {ESTADO_JSON}: {e}")


def carregar_estado():
    """
    Tenta carregar o ficheiro de estado.
    - Se existir e for válido, devolve o dicionário com a configuração.
    - Se não existir ou der erro, devolve None.
    """
    caminho = Path(ESTADO_JSON)

    if not caminho.exists():
        return None

    try:
        with open(caminho, "r", encoding="utf-8") as f:
            estado = json.load(f)
        # Pequena validação básica
        if not isinstance(estado, dict):
            print("⚠️ Conteúdo de estado inválido, a ignorar.")
            return None
        return estado
    except Exception as e:
        print(f"⚠️ Erro ao ler {ESTADO_JSON}: {e}")
        return None
