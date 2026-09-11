import json
import os

from path_utils import caminho_arquivo_local

ARQUIVO_CONFIG = caminho_arquivo_local('config_coordenadas.json')

COORDENADAS_PADRAO = {
    'grupo': {'x': 273, 'y': 367},
    'mensagem': {'x': 798, 'y': 1044},
    'anexo': {'x': 674, 'y': 1041},
    'fotos': {'x': 684, 'y': 777}
}


def carregar_coordenadas():
    if os.path.exists(ARQUIVO_CONFIG):
        try:
            with open(ARQUIVO_CONFIG, 'r') as f:
                return json.load(f)
        except Exception:
            pass
    return COORDENADAS_PADRAO.copy()


def salvar_coordenadas(coordenadas):
    with open(ARQUIVO_CONFIG, 'w') as f:
        json.dump(coordenadas, f, indent=2)


def obter_coordenada(nome):
    config = carregar_coordenadas()
    return config.get(nome, COORDENADAS_PADRAO.get(nome, {'x': 0, 'y': 0}))
