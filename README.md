# automacao-envio-links-afiliados

Aplicativo desktop em Python para gerenciamento de produtos, geração de mensagens, envio via WhatsApp Web e automação de divulgação com suporte a templates, imagens e agendamento.

## Visão geral

Este projeto inclui:

- interface desktop com `customtkinter`
- integração com WhatsApp Web para envio de mensagens
- cadastro e edição de produtos em planilha Excel
- geração de templates de texto
- suporte a coordenadas para posicionamento em interface do WhatsApp
- build dos executáveis Linux e Windows
- instalação local em Linux com ícone e launcher no menu de aplicativos

## Requisitos

### Linux

- Python 3.10+
- `xclip`, `xsel` ou `wl-copy` para melhor compatibilidade com clipboard
- `sudo` para o script de instalação do executável em sistema

### Windows

- Windows 10 ou 11
- Python 3.10+
- Google Chrome ou Microsoft Edge
- WhatsApp Web autenticado

## Instalação do código-fonte

### 1) Clonar o projeto

```bash
git clone <url-do-repositorio>
cd automacao_afiliados-main
```

### 2) Criar ambiente virtual

#### Linux/macOS

```bash
python -m venv .venv
source .venv/bin/activate
```

#### Windows (PowerShell)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3) Instalar dependências

#### Linux/macOS

```bash
pip install -r requirements.txt -r requirements-linux.txt
pip install pyinstaller
```

#### Windows

```powershell
pip install -r requirements.txt
pip install pyinstaller
```

## Executar diretamente

```bash
python main.py
```

## Gerar executável

### Linux

```bash
./build_executable.sh
```

### Windows

```powershell
.\build_executable_windows.ps1
```

Executáveis gerados:

- Linux: `dist/automacao_afiliados`
- Windows: `dist/automacao_afiliados.exe`

## Instalar o executável no Linux

Este projeto também inclui um script para instalar o binário e registrar o launcher no sistema Linux:

```bash
sudo ./install_linux.sh
```

### O que o script faz

- copia o executável para `/opt/automacao_afiliados/automacao_afiliados`
- copia o ícone do software (`imagens/icone.png`) para `/usr/share/icons/automacao_afiliados.png`
- instala o arquivo `.desktop` em `/usr/share/applications/automacao_afiliados.desktop`
- atualiza o cache de desktop entries

Depois disso, o aplicativo pode ser aberto pelo menu de aplicativos ou executado diretamente com:

```bash
/opt/automacao_afiliados/automacao_afiliados
```

## Estrutura principal do projeto

- `main.py`: aplicação principal
- `whatsapp.py`: envio de mensagens e manipulação do clipboard
- `coordenadas.py`: armazenamento de presets e coordenadas
- `templates.py`: gestão de templates de mensagens
- `verificacao_imagem.py`: validação de links de imagem
- `path_utils.py`: resolução de caminhos em desenvolvimento e em executável
- `build_executable.sh`: build do executável Linux
- `build_executable_windows.ps1`: build do executável Windows
- `install_linux.sh`: instalação do executável e launcher no Linux
- `automacao_afiliados.desktop`: arquivo de launcher para Linux
- `requirements.txt`: dependências comuns
- `requirements-linux.txt`: dependências específicas do Linux
- `imagens/icone.png`: ícone do software usado pelo launcher Linux

## Observações importantes

- O App usa `PyAutoGUI`, `customtkinter`, `openpyxl`, `Pillow`, `pyperclip` e `PyInstaller`.
- Para uso real no WhatsApp Web, o navegador deve estar aberto e autenticado.
- Em Linux, a compatibilidade com clipboard pode variar entre X11 e Wayland; o código tenta usar `wl-copy`, `xclip` e `xsel` automaticamente.
- O build do Windows deve ser feito em ambiente Windows, pois o executável `.exe` não é gerado nativamente em Linux.
- O arquivo `termo_aceito.json` é armazenado localmente e não deve ser enviado para o repositório.
