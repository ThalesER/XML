import xmltodict
import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

def pegar_infos(nome_arquivo, valores):
    with open(f'{nome_arquivo}', "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)

        # Verifica a estrutura do XML
        if "NFe" in dic_arquivo:
            infos_nf = dic_arquivo["NFe"]['infNFe']
        else:
            infos_nf = dic_arquivo['nfeProc']["NFe"]['infNFe']

        # Extração das informações
        numero_nota = infos_nf["@Id"]
        chave_acesso = infos_nf["@Id"]  # A chave de acesso é o ID completo da nota
        empresa_emissora = infos_nf['emit']['xNome']
        cnpj_emissora = infos_nf['emit']['CNPJ']
        
        # Verifica se a data de emissão está no formato dhEmi ou dEmi
        data_emissao = infos_nf['ide'].get('dhEmi', infos_nf['ide'].get('dEmi', 'Não informado'))
        
        nome_cliente = infos_nf["dest"]["xNome"]
        endereco_completo_cliente = f"{infos_nf['dest']['enderDest']['xLgr']}, {infos_nf['dest']['enderDest']['nro']} - {infos_nf['dest']['enderDest']['xMun']} - {infos_nf['dest']['enderDest']['CEP']}"
        valor_total = infos_nf['total']['ICMSTot']['vNF']
        
        if "vol" in infos_nf["transp"]:
            peso = infos_nf["transp"]["vol"]["pesoB"]
        else:
            peso = "Não informado"
        
        # Adiciona todas as informações à lista
        valores.append([numero_nota, chave_acesso, empresa_emissora, cnpj_emissora, data_emissao, nome_cliente, endereco_completo_cliente, valor_total, peso])

def processar_arquivos(pasta_origem, arquivo_destino):
    try:
        lista_arquivos = os.listdir(pasta_origem)
        colunas = ["numero_nota", "chave_acesso", "empresa_emissora", "cnpj_emissora", "data_emissao", "nome_cliente", "endereco_completo_cliente", "valor_total", "peso"]
        valores = []

        for arquivo in lista_arquivos:
            if arquivo.endswith('.xml'):
                pegar_infos(os.path.join(pasta_origem, arquivo), valores)

        tabela = pd.DataFrame(columns=colunas, data=valores)
        tabela.to_excel(arquivo_destino, index=False)

        messagebox.showinfo("Sucesso", f"O arquivo {arquivo_destino} foi gerado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

def selecionar_pasta_origem():
    pasta_origem = filedialog.askdirectory(title="Selecione a pasta com os arquivos XML")
    if pasta_origem:
        selecionar_pasta_destino(pasta_origem)

def selecionar_pasta_destino(pasta_origem):
    arquivo_destino = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                   filetypes=[("Excel files", "*.xlsx")],
                                                   title="Escolha o local e o nome do arquivo Excel",
                                                   initialfile="NotasFiscais.xlsx")
    if arquivo_destino:
        processar_arquivos(pasta_origem, arquivo_destino)

# Configuração da interface gráfica
root = tk.Tk()
root.title("Gerador de Notas Fiscais")
root.geometry("300x150")

label = tk.Label(root, text="Selecione a pasta com os arquivos XML")
label.pack(pady=20)

btn_selecionar = tk.Button(root, text="Selecionar Pasta", command=selecionar_pasta_origem)
btn_selecionar.pack(pady=10)

root.mainloop()
