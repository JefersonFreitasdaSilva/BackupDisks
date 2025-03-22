import os
import json
import psutil
import time

import requests
import win32api
import tkinter as tk
from tkinter import ttk, Text, Scrollbar
import threading
from pystray import Icon, MenuItem, Menu
from PIL import Image, ImageDraw, ImageTk  # Importar ImageTk
import sys

# Variável global para controlar o monitoramento
monitorando = False
REGISTRO_ARQUIVOS = "registro_hd.json"
GOOGLE_SHEETS_URL = ""

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
    """Salva os registros localmente e envia para o Google Sheets via API."""
    dados = carregar_registro()
    dados[disk_id] = {"arquivos": arquivos, "memoria_livre": memoria_livre}

    # Salva localmente
    with open(REGISTRO_ARQUIVOS, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)

    # Prepara os dados para envio ao Google Sheets
    dados_para_enviar = []
    for arquivo in arquivos:
        dados_para_enviar.append({
            "disk_id": disk_id,
            "arquivo": arquivo,
            "memoria_livre": memoria_livre
        })

    # Envia para Google Sheets
    try:
        response = requests.post(GOOGLE_SHEETS_URL, json=dados_para_enviar)
        print(f"Resposta do Google Sheets: {response.status_code} - {response.text}")
        if response.status_code == 200:
            print("✅ Dados enviados para o Google Sheets com sucesso.")
        else:
            print(f"⚠️ Erro ao enviar para o Google Sheets: {response.status_code}")
            print("Resposta detalhada:", response.json())
    except Exception as e:
        print(f"⚠️ Erro ao se conectar com o Google Sheets: {e}")

def sincronizar_com_planilha():
    """Sincroniza os dados entre a planilha do Google e o arquivo JSON."""
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
    unidades = []
    discos_ignorados = ["C:\\", "D:\\", "M:\\"]
    for letra in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        caminho = f"{letra}:\\"  # Verifica cada letra de A-Z
        if os.path.exists(caminho) and caminho not in discos_ignorados:
            unidades.append(caminho)
    return unidades

def get_disk_id(drive):
    try:
        return win32api.GetVolumeInformation(drive)[0] or drive
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
    """Monitora as mudanças de discos conectados e atualiza o registro automaticamente."""
    hd_detectados = set()
    hd_previos = set()

    while monitorando:  # Verifica a variável global 'monitorando'
        # Verifica as unidades conectadas
        unidades_conectadas = set(encontrar_hd())

        # Detecta discos conectados novos
        novos_hd = unidades_conectadas - hd_previos
        for caminho in novos_hd:
            disk_id = get_disk_id(caminho)
            if disk_id and disk_id not in hd_detectados:
                arquivos_total = listar_arquivos(caminho)
                memoria_livre = obter_memoria_livre()

                if arquivos_total:
                    salvar_registro(disk_id, arquivos_total, memoria_livre)
                    print(f"✅ Registro atualizado para o HD '{disk_id}'")
                    print(f"📊 Memória livre no momento da conexão: {memoria_livre:.2f} MB")
                    hd_detectados.add(disk_id)

        # Detecta discos desconectados
        hd_desconectados = hd_previos - unidades_conectadas
        for caminho in hd_desconectados:
            disk_id = get_disk_id(caminho)
            if disk_id in hd_detectados:
                print(f"⚠️ HD desconectado '{disk_id}'")
                hd_detectados.remove(disk_id)

        # Atualiza a lista de discos
        hd_previos = unidades_conectadas

        time.sleep(5)  # Aguarda 5 segundos antes de verificar novamente


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

# Após a inicialização da interface, chame buscar_arquivos para garantir que todos os discos sejam carregados
buscar_arquivos()

# Atualiza a lista de discos
iniciar_monitoramento()
sincronizar_com_planilha()

# Inicia o ícone da bandeja
icone_bandeja = create_icon()
icone_bandeja.run_detached()

# Inicia a interface gráfica
janela.protocol("WM_DELETE_WINDOW", on_closing)  # Trata o evento de fechar a janela
janela.mainloop()

