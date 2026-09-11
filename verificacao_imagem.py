import re
from urllib.request import urlopen, Request
from urllib.error import URLError


EXTENSOES_IMAGEM = {'.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp', '.svg'}


def eh_url_formato_valido(url):
    """Verifica se a URL tem formato válido"""
    if not url or not isinstance(url, str):
        return False

    padrao = re.compile(
        r'^https?://'
        r'\S+'
    )

    return bool(padrao.match(url.strip()))


def servidor_entrega_imagem(url):
    """Faz request HEAD/GET e verifica se o Content-Type é imagem"""
    try:
        requisicao = Request(url, method='GET')
        requisicao.add_header('User-Agent', 'Mozilla/5.0')
        resposta = urlopen(requisicao, timeout=10)

        content_type = resposta.headers.get('Content-Type', '')
        return content_type.startswith('image/')

    except (URLError, OSError, ValueError):
        return False


def eh_link_imagem_valido(url):
    """Validação completa: formato + servidor"""
    if not eh_url_formato_valido(url):
        return False, 'URL com formato inválido'

    if not servidor_entrega_imagem(url):
        return False, 'O servidor não retornou uma imagem válida'

    return True, 'Imagem válida'
