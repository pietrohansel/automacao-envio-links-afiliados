from pathlib import Path
import sys


def _base_dir_do_projeto() -> Path:
    if getattr(sys, "_MEIPASS", None):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def caminho_arquivo_local(nome_arquivo: str) -> str:
    return str((_base_dir_do_projeto() / nome_arquivo).resolve())


def caminho_recurso(relativo: str) -> str:
    base = Path(sys._MEIPASS) if getattr(sys, "_MEIPASS", None) else _base_dir_do_projeto()

    candidatos = [
        base / relativo,
        base / "imagens" / relativo,
        base / "images" / relativo,
        _base_dir_do_projeto() / relativo,
        _base_dir_do_projeto() / "imagens" / relativo,
        _base_dir_do_projeto() / "images" / relativo,
    ]

    for candidato in candidatos:
        if candidato.exists():
            return str(candidato.resolve())

    return str((_base_dir_do_projeto() / relativo).resolve())
