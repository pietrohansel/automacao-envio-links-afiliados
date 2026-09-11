import webbrowser
import pyautogui
import time
import subprocess
import tempfile
import os
import shutil
from urllib.request import urlopen
from coordenadas import obter_coordenada


def copiar_para_clipboard(texto):
    """Copia o texto fornecido para a área de transferência do sistema."""
    if texto is None:
        return False

    try:
        try:
            import pyperclip
            pyperclip.copy(texto)
            return True
        except Exception:
            pass

        if os.name == 'nt':
            processo = subprocess.run(
                ['powershell', '-NoProfile', '-Command', f"Set-Clipboard -Value '{texto}'"],
                capture_output=True,
                text=True,
                check=False
            )
            return processo.returncode == 0

        # Linux: tenta os mecanismos mais comuns de clipboard em X11 e Wayland.
        candidatos = [
            ['wl-copy'],
            ['xclip', '-selection', 'clipboard'],
            ['xsel', '--clipboard', '--input'],
        ]

        for comando in candidatos:
            if shutil.which(comando[0]) is None:
                continue

            try:
                subprocess.run(comando, input=texto.encode('utf-8'), check=True)
                return True
            except Exception:
                continue

        return False
    except Exception:
        return False


def garantir_caminho_imagem_local(link_ou_caminho):
    """
    Garante que o parâmetro seja um caminho de arquivo local válido.
    Aceita tanto caminhos do sistema de arquivos quanto URLs remotas.
    """
    if not link_ou_caminho:
        return None

    caminho_limpo = str(link_ou_caminho).strip()

    if os.path.exists(caminho_limpo):
        return os.path.abspath(caminho_limpo)

    if caminho_limpo.startswith('http://') or caminho_limpo.startswith('https://'):
        try:
            req = urlopen(caminho_limpo, timeout=10)
            dados = req.read()
            ext = ".jpg"
            if ".png" in caminho_limpo.lower():
                ext = ".png"
            elif ".webp" in caminho_limpo.lower():
                ext = ".webp"

            caminho_temp = os.path.join(tempfile.gettempdir(), f"temp_wp_img{ext}")
            with open(caminho_temp, "wb") as f:
                f.write(dados)
            return caminho_temp
        except Exception as e:
            print(f"Erro ao baixar imagem: {e}")
            return None

    return None


def enviar_para_grupo(nome_grupo, mensagem, link_imagem=None, cancel_event=None):
    if cancel_event and cancel_event.is_set():
        return False, 'Envio cancelado pelo usuário.'

    if not nome_grupo:
        return False, 'Digite o nome do grupo do WhatsApp!'

    if mensagem == 'Sem descricao' or mensagem is None:
        mensagem = ''

    caminho_imagem = garantir_caminho_imagem_local(link_imagem)

    if not mensagem.strip() and not caminho_imagem:
        return False, 'Preencha a descrição ou informe uma imagem antes de enviar!'

    try:
        webbrowser.open('https://web.whatsapp.com/')
        time.sleep(12)

        if cancel_event and cancel_event.is_set():
            return False, 'Envio cancelado pelo usuário.'

        # Buscar o grupo/contato
        pyautogui.hotkey('ctrl', 'f')
        time.sleep(1)

        copiar_para_clipboard(nome_grupo)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(2)

        coord_grupo = obter_coordenada('grupo')
        pyautogui.click(x=coord_grupo['x'], y=coord_grupo['y'])
        time.sleep(2)

        if cancel_event and cancel_event.is_set():
            return False, 'Envio cancelado pelo usuário.'

        if caminho_imagem:
            # 1. Clicar no botão Anexo
            coord_anexo = obter_coordenada('anexo')
            pyautogui.click(x=coord_anexo['x'], y=coord_anexo['y'])
            time.sleep(1.5)

            # 2. Clicar em Fotos e Vídeos
            coord_fotos = obter_coordenada('fotos')
            pyautogui.click(x=coord_fotos['x'], y=coord_fotos['y'])
            time.sleep(2.5)

            # 3. Colar o caminho da imagem no campo de texto padrão da janela do SO
            copiar_para_clipboard(caminho_imagem)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
            pyautogui.press('enter')

            # 4. Aguarda carregar a pré-visualização da mídia no WhatsApp Web
            time.sleep(4)

            # 5. Colar a descrição/legenda da imagem na pré-visualização (se houver)
            if mensagem.strip():
                copiar_para_clipboard(mensagem)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(1)

            # 6. Pressionar Enter para enviar a imagem (com ou sem legenda)
            pyautogui.press('enter')
            time.sleep(2)
        else:
            # Envio simples apenas de texto caso não haja imagem
            coord_msg = obter_coordenada('mensagem')
            pyautogui.click(x=coord_msg['x'], y=coord_msg['y'])
            time.sleep(1)

            if not copiar_para_clipboard(mensagem):
                return False, 'Não foi possível acessar o clipboard.'

            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)

            pyautogui.press('enter')
            time.sleep(2)

        return True, 'Mensagem enviada com sucesso!'

    except Exception as e:
        return False, f'Erro ao enviar: {str(e)}'