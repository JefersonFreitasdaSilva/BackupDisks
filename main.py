import os
import json
import psutil
import time

import requests
import win32api
import tkinter as tk
from tkinter import ttk, Text, Scrollbar, simpledialog, messagebox
import threading
from pystray import Icon, MenuItem, Menu
from PIL import Image, ImageDraw, ImageTk  # Importar ImageTk
import sys
# Variáveis globais
monitorando = False
REGISTRO_ARQUIVOS = "registro_hd.json"
CONFIG_ARQUIVO = "config.json"

# Funções para salvar e carregar configurações
def salvar_configuracoes(url, discos):
    config = {"url": url, "discos": discos}
    with open(CONFIG_ARQUIVO, "w") as f:
        json.dump(config, f)

def carregar_configuracoes():
    try:
        with open(CONFIG_ARQUIVO, "r") as f:
            config = json.load(f)
            return config["url"], config.get("discos", [])
    except FileNotFoundError:
        return "",  ["C","D"]  # Valores padrão

# Carregar configurações no início do programa
GOOGLE_SHEETS_URL, discos_ignorados = carregar_configuracoes()


# Função para alterar a URL do Google Sheets
def alterar_url():
    global GOOGLE_SHEETS_URL  # Mova a declaração global para o início da função
    nova_url = simpledialog.askstring("Alterar URL", "Digite a nova URL do Google Sheets:", initialvalue=GOOGLE_SHEETS_URL)
    if nova_url:
        GOOGLE_SHEETS_URL = nova_url
        salvar_configuracoes(GOOGLE_SHEETS_URL, discos_ignorados)
        messagebox.showinfo("URL Alterada", "A URL do Google Sheets foi alterada.")

# Função para alterar os discos ignorados
def alterar_discos_ignorados():
    global discos_ignorados  # Mova a declaração global para o início da função
    discos_str = ",".join(discos_ignorados)
    novos_discos = simpledialog.askstring("Alterar Discos Ignorados", "Digite os discos ignorados (separados por vírgulas):", initialvalue=discos_str)
    if novos_discos:
        discos_ignorados = novos_discos.split(",")
        salvar_configuracoes(GOOGLE_SHEETS_URL, discos_ignorados)
        messagebox.showinfo("Discos Ignorados Alterados", "Os discos ignorados foram alterados.")



# Função para carregar o ícone para a barra de título
def carregar_icone_janela():
    # Caminho para o ícone personalizado dentro da pasta "image"
    if getattr(sys, 'frozen', False):
        # Estamos rodando como um executável
        base_path = sys._MEIPASS
    else:
        # Estamos rodando como um script
        base_path = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(base_path, 'image', 'bandeja_ico.png')
    if os.path.exists(icon_path):
        icon_image = Image.open(icon_path)
        return ImageTk.PhotoImage(icon_image)
    else:
        print(f"Erro: o ícone '{icon_path}' não foi encontrado.")
        sys.exit(1)

def create_icon():
    # Carregar a imagem personalizada (substitua o caminho abaixo pela sua imagem)
    if getattr(sys, 'frozen', False):
        # Estamos rodando como um executável
        base_path = sys._MEIPASS
    else:
        # Estamos rodando como um script
        base_path = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(base_path, 'image', 'bandeja_ico.png')
    icon_image = Image.open(icon_path)
    icon = Icon("HD Monitor", icon_image, menu=create_menu())
    return icon
# Função para criar o menu do ícone
def create_menu():
    return Menu(MenuItem('Abrir', abrir_janela), MenuItem('Sair', sair))

# Função para restaurar a janela
def abrir_janela(icon, item):
    janela.after(0, lambda: janela.deiconify())  # Abertura da janela agendada no thread principal
    icon.stop()  # Para o ícone da bandeja

# Função para sair do programa
# Função para sair do programa
def sair(icon, item):
    global monitorando
    monitorando = False
    janela.quit()  # Fecha a janela principal
    icon.stop()  # Remove o ícone da bandeja

# Função para ocultar a janela quando o X for clicado
def on_closing():
    janela.withdraw()  # Minimiza a janela
def carregar_registro():
    """Carrega os registros locais ou cria um novo se não existir."""
    if getattr(sys, 'frozen', False):
        # Estamos rodando como um executável
        base_path = sys._MEIPASS
    else:
        # Estamos rodando como um script
        base_path = os.path.dirname(os.path.abspath(__file__))
    registro_path = os.path.join(base_path, REGISTRO_ARQUIVOS)
    if os.path.exists(registro_path):
        try:
            with open(registro_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            return {}
    else:
        return {}


def salvar_registro(disk_id, arquivos, memoria_livre):
    if disk_id.lower() in [disco.lower() for disco in discos_ignorados]:
        return

    dados = carregar_registro()
    arquivos_novos = []

    if disk_id in dados:
        arquivos_existentes = set(dados[disk_id]["arquivos"])
        arquivos_novos = [arq for arq in arquivos if arq not in arquivos_existentes]
        dados[disk_id]["arquivos"].extend(arquivos_novos)
    else:
        dados[disk_id] = {"arquivos": arquivos, "memoria_livre": memoria_livre}

    dados[disk_id]["memoria_livre"] = memoria_livre

    with open(REGISTRO_ARQUIVOS, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)


def encontrar_hd():
    global discos_ignorados
    unidades = []
    for letra in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        caminho = f"{letra}:\\"
        if os.path.exists(caminho) and letra not in discos_ignorados:
            unidades.append(caminho)
    return unidades

def sincronizar_com_planilha():
    """Sincroniza os dados entre a planilha do Google e o arquivo JSON."""
    if not GOOGLE_SHEETS_URL:
        print("⚠️ URL do Google Sheets não configurada. Sincronização não disponível.")
        return

    try:
        # Obter dados da planilha
        response = requests.get(GOOGLE_SHEETS_URL)
        if response.status_code != 200:
            print(f"⚠️ Falha ao obter dados da planilha: {response.status_code}")
            return

        try:
            dados_planilha = response.json()
        except json.JSONDecodeError:
            print(f"⚠️ Erro: Resposta da planilha não é um JSON válido: {response.text}")
            return

        # Obter dados do arquivo JSON
        dados_locais = carregar_registro()

        # Comparar e sincronizar
        discos_planilha = set(dados_planilha.keys())
        discos_locais = set(dados_locais.keys())

        # Adicionar discos da planilha ao JSON (se necessário)
        for disk_id in discos_planilha - discos_locais:
            dados_locais[disk_id] = dados_planilha[disk_id]

        # Adicionar discos do JSON à planilha (se necessário)
        for disk_id in discos_locais - discos_planilha:
            salvar_registro(disk_id, dados_locais[disk_id]['arquivos'], dados_locais[disk_id]['memoria_livre'])

        # Atualizar arquivos no JSON (se necessário)
        for disk_id in discos_planilha & discos_locais:
            if dados_planilha[disk_id]['arquivos'] != dados_locais[disk_id]['arquivos']:
                dados_locais[disk_id]['arquivos'] = dados_planilha[disk_id]['arquivos']

        # Atualizar arquivo JSON
        with open(REGISTRO_ARQUIVOS, "w", encoding="utf-8") as f:
            json.dump(dados_locais, f, indent=4, ensure_ascii=False)

        print("✅ Dados sincronizados com sucesso.")
    except Exception as e:
        print(f"⚠️ Erro ao sincronizar dados: {e}")

    # Atualizar a lista de discos após a sincronização
    atualizar_lista_discos()
    texto_json.delete("1.0", tk.END)
    texto_json.insert(tk.END, "Dados sincronizados com sucesso.\n", "info")


def listar_arquivos(diretorio):
    arquivos_e_pastas = []
    try:
        for item in os.listdir(diretorio):
            caminho_completo = os.path.join(diretorio, item)
            arquivos_e_pastas.append(caminho_completo)
    except Exception as e:
        print(f"Erro ao listar arquivos/pastas em {diretorio}: {e}")
    return arquivos_e_pastas


def encontrar_hd():
    global discos_ignorados
    unidades = []
    for letra in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        caminho = f"{letra}:\\"
        if os.path.exists(caminho) and caminho.lower() not in [disco.lower() for disco in discos_ignorados]: # adicionado o .lower() para evitar erros de case sensitive
            unidades.append(caminho)
    return unidades

def get_disk_id(drive):
    try:
        return drive[0]
    except:
        return drive

def atualizar_lista_discos():
    lista_discos.delete(0, tk.END)
    dados = carregar_registro()
    for disk_id, info in dados.items():
        try:
            memoria_livre = float(info['memoria_livre'])
            lista_discos.insert(tk.END, f"HD: {disk_id} - Memória Livre: {memoria_livre:.2f} MB")
        except ValueError:
            lista_discos.insert(tk.END, f"HD: {disk_id} - Memória Livre: {info['memoria_livre']} MB")

def buscar_arquivos(event=None):
    termo_busca = barra_busca.get().lower()
    dados = carregar_registro()
    texto_json.delete("1.0", tk.END)  # Limpa a área de detalhes de arquivos
    lista_discos.delete(0, tk.END)  # Limpa a lista de discos superior

    # Se a barra de busca estiver vazia, exibe todos os arquivos e discos normalmente
    if not termo_busca:
        atualizar_lista_discos()  # Exibe todos os discos
        texto_json.insert(tk.END, "Digite algo para buscar arquivos...\n", "info")
        return

    # Filtra os discos que possuem arquivos que contêm o termo de busca
    discos_filtrados = []
    for disk_id, info in dados.items():
        arquivos = info['arquivos']
        arquivos_filtrados = [arq for arq in arquivos if termo_busca in arq.lower()]

        if arquivos_filtrados:
            discos_filtrados.append((disk_id, arquivos_filtrados, info['memoria_livre']))

    # Atualiza a lista de discos para mostrar apenas os discos filtrados
    for disk_id, arquivos_filtrados, memoria_livre in discos_filtrados:
        lista_discos.insert(tk.END, f"HD: {disk_id} - Memória Livre: {memoria_livre:.2f} MB")

    # Exibe os arquivos encontrados na parte inferior (área de detalhes)
    if discos_filtrados:
        texto_json.insert(tk.END, "Arquivos encontrados:\n", "subtitulo")
        for disk_id, arquivos_filtrados, _ in discos_filtrados:
            texto_json.insert(tk.END, f"HD: {disk_id}\n", "titulo")
            for arq in arquivos_filtrados:
                texto_json.insert(tk.END, f"- {arq}\n", "item")
    else:
        texto_json.insert(tk.END, "Nenhum arquivo encontrado para o termo de busca.\n", "info")

# Função para exibir todos os arquivos de um disco quando ele for clicado
def mostrar_arquivos(event=None):
    selecionado = lista_discos.curselection()
    if not selecionado:
        return
    disk_info = lista_discos.get(selecionado[0])
    disk_id = disk_info.split(" -")[0].replace("HD: ", "")  # Extraindo o ID do disco
    dados = carregar_registro()

    if disk_id in dados:
        arquivos = dados[disk_id]['arquivos']
        memoria_livre = dados[disk_id]['memoria_livre']
        texto_json.delete("1.0", tk.END)  # Limpa a área de detalhes de arquivos
        texto_json.insert(tk.END, f"HD: {disk_id}\n", "titulo")
        texto_json.insert(tk.END, f"Memória Livre: {memoria_livre:.2f} MB\n", "info")
        texto_json.insert(tk.END, "Arquivos:\n", "subtitulo")

        for arq in arquivos:
            texto_json.insert(tk.END, f"- {arq}\n", "item")


def obter_memoria_livre():
    memoria = psutil.virtual_memory()
    return memoria.available / (1024 ** 2)  # Retorna a memória livre em MB

def forcar_busca_atualizacao():
    atualizar_lista_discos()
    texto_json.delete("1.0", tk.END)
    texto_json.insert(tk.END, "Registro atualizado!\n", "info")


# Função de monitoramento contínuo
def monitorar_hd():
    hd_detectados = set()
    hd_previos = set()

    while monitorando:
        unidades_conectadas = set(encontrar_hd())

        novos_hd = unidades_conectadas - hd_previos
        for caminho in novos_hd:
            disk_id = get_disk_id(caminho)
            if disk_id.lower() in [disco.lower() for disco in discos_ignorados]:
                continue
            if disk_id and disk_id not in hd_detectados:
                arquivos_total = listar_arquivos(caminho)
                memoria_livre = obter_memoria_livre()

                if arquivos_total:
                    salvar_registro(disk_id, arquivos_total, memoria_livre)
                    print(f"✅ Registro atualizado para o HD '{disk_id}'")
                    print(f"📊 Memória livre no momento da conexão: {memoria_livre:.2f} MB")
                    hd_detectados.add(disk_id)

        hd_desconectados = hd_previos - unidades_conectadas
        for caminho in hd_desconectados:
            disk_id = get_disk_id(caminho)
            if disk_id in hd_detectados:
                print(f"⚠️ HD desconectado '{disk_id}'")
                hd_detectados.remove(disk_id)

        hd_previos = unidades_conectadas

        time.sleep(5)


def iniciar_monitoramento():
    global monitorando
    monitorando = True
    monitoramento_thread = threading.Thread(target=monitorar_hd)
    monitoramento_thread.daemon = True
    monitoramento_thread.start()

def parar_monitoramento():
    global monitorando
    monitorando = False


# Configuração da janela principal
janela = tk.Tk()
janela.title("Exibição de Registros de Discos")
janela.geometry("500x1000")
janela.configure(bg="#2e2e2e")  # Cor de fundo escura da janela principal

icone_janela = carregar_icone_janela()
janela.iconphoto(True, icone_janela)  # Define o ícone na barra de título da janel

# Configuração de cores dos widgets
# Definindo as cores
cor_fundo = "#1e1e1e"
cor_texto = "#ffffff"
cor_lista = "#333333"
cor_destaque = "#007acc"
cor_config ="#00ac47"

# Configuração da barra de pesquisa
barra_busca = tk.Entry(janela, width=40, bg=cor_lista, fg=cor_texto, borderwidth=0)
barra_busca.pack(pady=10)
barra_busca.bind("<KeyRelease>", buscar_arquivos)  # Atualiza a lista ao digitar

# Lista de discos
lista_discos = tk.Listbox(janela, width=50, height=15, bg=cor_lista, fg=cor_texto, selectmode=tk.SINGLE, activestyle="none", highlightthickness=0, bd=0)
lista_discos.pack(pady=10)
lista_discos.bind("<ButtonRelease-1>", mostrar_arquivos)

# Área de texto para exibir arquivos e detalhes
texto_json = Text(janela, width=70, height=15,  bg=cor_lista, fg=cor_texto, borderwidth=0, highlightthickness=0)
texto_json.pack(pady=10)

# Botões de controle de monitoramento
btn_iniciar = tk.Button(janela, text="Iniciar Monitoramento", command=iniciar_monitoramento, bg=cor_destaque, fg=cor_texto, relief="flat")
btn_iniciar.pack(pady=5)

btn_parar = tk.Button(janela, text="Parar Monitoramento", command=parar_monitoramento, bg=cor_destaque, fg=cor_texto, relief="flat")
btn_parar.pack(pady=5)

# Botão de sincronização manual
btn_sincronizar = tk.Button(janela, text="Sincronizar Dados", command=sincronizar_com_planilha, bg=cor_destaque, fg=cor_texto, relief="flat")
btn_sincronizar.pack(pady=5)

# Botão para alterar a URL do Google Sheets
btn_alterar_url = tk.Button(janela, text="Alterar URL do Google Sheets", command=alterar_url, bg=cor_config, fg=cor_texto, relief="flat")
btn_alterar_url.pack(pady=5)

# Botão para alterar os discos ignorados
btn_alterar_discos = tk.Button(janela, text="Alterar Discos Ignorados", command=alterar_discos_ignorados, bg=cor_config, fg=cor_texto, relief="flat")
btn_alterar_discos.pack(pady=5)

# Após a inicialização da interface, chame buscar_arquivos para garantir que todos os discos sejam carregados
buscar_arquivos()

# Atualiza a lista de discos
iniciar_monitoramento()
sincronizar_com_planilha()

# Inicia o ícone da bandeja
icone_bandeja = create_icon()
icone_bandeja.run_detached()

# Inicia a interface gráfica
janela.protocol("WM_DELETE_WINDOW", on_closing)
janela.mainloop()

