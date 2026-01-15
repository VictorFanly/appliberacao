import tkinter as tk
import json
import sys
from tkinter import filedialog
from sheets import registrar_liberacao
from tkinter import ttk, messagebox
from datetime import datetime
from docx import Document
import traceback
import re
import os

# ================= FUNÇÕES DE DOCUMENTO ================= #
CONFIG_ARQ = "config.json"

def carregar_config():
    if os.path.exists(CONFIG_ARQ):
        with open(CONFIG_ARQ, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def salvar_config(config):
    with open(CONFIG_ARQ, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

def calcular_dias(data_apreensao, data_liberacao):
    formato = "%d/%m/%Y"
    d1 = datetime.strptime(data_apreensao, formato)
    d2 = datetime.strptime(data_liberacao, formato)
    return (d2 - d1).days + 1

def gerar_documento(dados, pasta_saida):
    doc = Document(caminho_recurso("LIBERACAO.docx"))
    substituir_placeholders(doc, dados)

    nome_arquivo = f"TERMO_{dados['{{PLACA}}']}.docx"
    caminho = os.path.join(pasta_saida, nome_arquivo)

    doc.save(caminho)
    os.startfile(caminho)


def substituir_placeholders(doc, dados):
    for p in doc.paragraphs:
        texto = "".join(r.text for r in p.runs)
        alterou = False
        for chave, valor in dados.items():
            if chave in texto:
                texto = texto.replace(chave, valor)
                alterou = True
        if alterou and p.runs:
            p.runs[0].text = texto
            for r in p.runs[1:]:
                r.text = ""

def caminho_recurso(nome):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, nome)
    return nome

# ================= VALIDAÇÕES ================= #

def cpf_valido(cpf):
    cpf = re.sub(r'\D', '', cpf)
    if len(cpf) != 11 or cpf == cpf[0] * 11:
        return False
    for i in range(9, 11):
        soma = sum(int(cpf[j]) * ((i + 1) - j) for j in range(i))
        digito = (soma * 10) % 11
        digito = 0 if digito == 10 else digito
        if digito != int(cpf[i]):
            return False
    return True

def data_valida(data):
    try:
        d = datetime.strptime(data, "%d/%m/%Y")
        return d <= datetime.today()
    except ValueError:
        return False

# ================= FORMATAÇÕES ================= #

def formatar_placa(event):
    e = event.widget
    pos = e.index(tk.INSERT)
    texto = re.sub(r'[^A-Z0-9]', '', e.get().upper())
    novo = ""
    for i, c in enumerate(texto):
        if i == 3:
            novo += "-"
            if pos > i: pos += 1
        novo += c
    e.delete(0, tk.END)
    e.insert(0, novo[:8])
    e.icursor(pos)

def formatar_data(event):
    e = event.widget
    pos = e.index(tk.INSERT)
    texto = re.sub(r'\D', '', e.get())[:8]
    novo = ""
    for i, c in enumerate(texto):
        if i in (2, 4):
            novo += "/"
            if pos > i: pos += 1
        novo += c
    e.delete(0, tk.END)
    e.insert(0, novo)
    e.icursor(pos)

def formatar_cpf(event):
    e = event.widget
    pos = e.index(tk.INSERT)
    texto = re.sub(r'\D', '', e.get())[:11]
    novo = ""
    for i, c in enumerate(texto):
        if i in (3, 6):
            novo += "."
            if pos > i: pos += 1
        if i == 9:
            novo += "-"
            if pos > i: pos += 1
        novo += c
    e.delete(0, tk.END)
    e.insert(0, novo)
    e.icursor(pos)

def formatar_telefone(event):
    e = event.widget
    pos = e.index(tk.INSERT)
    texto = re.sub(r'\D', '', e.get())[:11]
    novo = ""
    for i, c in enumerate(texto):
        if i == 2 or i == 3:
            novo += " "
            if pos > i: pos += 1
        if i == 7:
            novo += "-"
            if pos > i: pos += 1
        novo += c
    e.delete(0, tk.END)
    e.insert(0, novo)
    e.icursor(pos)

def forcar_maiusculo(var):
    if var.get() != var.get().upper():
        var.set(var.get().upper())

# ================= UI ================= #

def iniciar_app():
    root = tk.Tk()
    root.title("Gerador de Termo de Liberação")

    config = carregar_config()

    # Proporção 9:16
    root.geometry("540x820")

    # Impede abrir menor que isso
    root.minsize(540, 820)

    # (Opcional) permitir redimensionar
    root.resizable(True, True)

    container = ttk.Frame(root)
    container.pack(fill="both", expand=True, padx=20, pady=20)

    style = ttk.Style()
    style.theme_use("clam")

    def escolher_pasta():
        pasta = filedialog.askdirectory(title="Selecione a pasta para salvar os termos")
        if pasta:
            config["pasta_saida"] = pasta
            salvar_config(config)
            label_pasta.config(text=pasta)

    # Variáveis
    marca_var = tk.StringVar(); marca_var.trace_add("write", lambda *_: forcar_maiusculo(marca_var))
    modelo_var = tk.StringVar(); modelo_var.trace_add("write", lambda *_: forcar_maiusculo(modelo_var))
    cor_var = tk.StringVar(); cor_var.trace_add("write", lambda *_: forcar_maiusculo(cor_var))
    chassi_var = tk.StringVar(); chassi_var.trace_add("write", lambda *_: forcar_maiusculo(chassi_var))
    lote_var = tk.StringVar()

    tipo_agente_var = tk.StringVar()
    codigo_agente_var = tk.StringVar(); codigo_agente_var.trace_add("write", lambda *_: forcar_maiusculo(codigo_agente_var))
    chave_var = tk.StringVar(); chave_var.trace_add("write", lambda *_: forcar_maiusculo(chave_var))

    nome_var = tk.StringVar(); nome_var.trace_add("write", lambda *_: forcar_maiusculo(nome_var))
    logradouro_var = tk.StringVar(); logradouro_var.trace_add("write", lambda *_: forcar_maiusculo(logradouro_var))
    bairro_var = tk.StringVar(); bairro_var.trace_add("write", lambda *_: forcar_maiusculo(bairro_var))
    cidade_var = tk.StringVar(value="ITAQUAQUECETUBA")
    uf_var = tk.StringVar(value="SP")

    trafego_var = tk.StringVar(value="SIM")
    assinatura_var = tk.StringVar(value="VICTOR LOIOLA")

    # Containers
    f_veiculo = ttk.LabelFrame(container, text="DADOS DO VEÍCULO")
    f_veiculo.pack(fill="x", padx=15, pady=6)

    f_agente = ttk.LabelFrame(container, text="APREENSÃO / AGENTE")
    f_agente.pack(fill="x", padx=15, pady=6)

    f_mun = ttk.LabelFrame(container, text="DADOS DO MUNÍCIPE")
    f_mun.pack(fill="x", padx=15, pady=6)

    f_op = ttk.LabelFrame(container, text="OPÇÕES")
    f_op.pack(fill="x", padx=15, pady=6)

    def campo(parent, label, row):
        ttk.Label(parent, text=label).grid(row=row, column=0, padx=10, pady=5, sticky="w")
        e = ttk.Entry(parent)
        e.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        return e

    def atualizar_chave_por_agente(*_):
        if tipo_agente_var.get() == "DFP":
            chave_var.set("MEIO AMBIENTE")
        else:
            chave_var.set("")

    tipo_agente_var.trace_add("write", atualizar_chave_por_agente)

    # Veículo
    e_placa = campo(f_veiculo, "Placa", 0); e_placa.bind("<KeyRelease>", formatar_placa)
    campo(f_veiculo, "Marca", 1).config(textvariable=marca_var)
    campo(f_veiculo, "Modelo", 2).config(textvariable=modelo_var)
    campo(f_veiculo, "Cor", 3).config(textvariable=cor_var)
    campo(f_veiculo, "Chassi", 4).config(textvariable=chassi_var)
    campo(f_veiculo, "Lote", 5).config(textvariable=lote_var)

    # Apreensão
    def atalho_tipo_agente(event, var):
        tecla = event.char.upper()

        mapa = {
            "P": "PM",
            "G": "GCM",
            "D": "DFP"
        }

        if tecla in mapa:
            var.set(mapa[tecla])
            return "break"  # impede outros comportamentos

    e_data = campo(f_agente, "Data Apreensão", 0); e_data.bind("<KeyRelease>", formatar_data)

    frame_ag = ttk.Frame(f_agente)
    frame_ag.grid(row=1, column=1, sticky="ew", padx=10)
    ttk.Label(f_agente, text="Agente").grid(row=1, column=0, sticky="w", padx=10)

    combo_tipo = ttk.Combobox(
        frame_ag,
        values=["PM", "GCM", "DFP"],
        textvariable=tipo_agente_var,
        width=6,
        state="readonly"
    )
    combo_tipo.pack(side="left")

    combo_tipo.bind(
        "<KeyPress>",
        lambda e: atalho_tipo_agente(e, tipo_agente_var)
    )

    ttk.Entry(frame_ag, textvariable=codigo_agente_var, width=25).pack(side="left", padx=6)

    campo(f_agente, "Chave de Liberação", 2).config(textvariable=chave_var)

    # Munícipe
    campo(f_mun, "Nome", 0).config(textvariable=nome_var)
    e_cpf = campo(f_mun, "CPF", 1); e_cpf.bind("<KeyRelease>", formatar_cpf)
    e_tel = campo(f_mun, "Telefone", 2); e_tel.bind("<KeyRelease>", formatar_telefone)
    campo(f_mun, "Logradouro", 3).config(textvariable=logradouro_var)
    campo(f_mun, "Bairro", 4).config(textvariable=bairro_var)
    campo(f_mun, "Cidade", 5).config(textvariable=cidade_var)

    estados = ["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA",
               "PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"]

    ttk.Combobox(f_mun, values=estados, textvariable=uf_var, width=5, state="readonly")\
        .grid(row=5, column=2, padx=5)

    # Opções
    ttk.Label(f_op, text="Tráfego").grid(row=0, column=0, padx=10)
    tk.Radiobutton(f_op, text="SIM", variable=trafego_var, value="SIM").grid(row=0, column=1)
    tk.Radiobutton(f_op, text="NÃO", variable=trafego_var, value="NAO").grid(row=0, column=2)

    ttk.Label(f_op, text="Atendente").grid(row=1, column=0, padx=10)
    ttk.Combobox(
        f_op,
        values=["VICTOR LOIOLA", "CRISTIANA", "ALINE"],
        textvariable=assinatura_var,
        state="readonly",
        width=30
    ).grid(row=1, column=1, columnspan=2, sticky="w")

    def gerar():
        if not e_placa.get() or not nome_var.get():
            messagebox.showerror("Erro", "Campos obrigatórios não preenchidos.")
            return
        if not data_valida(e_data.get()):
            messagebox.showerror("Erro", "Data inválida.")
            return
        if not cpf_valido(e_cpf.get()):
            messagebox.showerror("Erro", "CPF inválido.")
            return

        if not tipo_agente_var.get():
            messagebox.showerror("Erro", "Selecione o tipo de agente.")
            return

        hoje = datetime.today().strftime("%d/%m/%Y")
        dias = calcular_dias(e_data.get(), hoje)

        dados = {
            "{{PLACA}}": e_placa.get(),
            "{{MARCA}}": marca_var.get(),
            "{{MODELO}}": modelo_var.get(),
            "{{COR}}": cor_var.get(),
            "{{CHASSI}}": chassi_var.get(),
            "{{LOTE}}": lote_var.get(),
            "{{APREENSAO}}": e_data.get(),
            "{{TIPO}}": tipo_agente_var.get(),
            "{{CODIGO}}": codigo_agente_var.get(),
            "{{CHAVE}}": chave_var.get(),
            "{{NOME}}": nome_var.get(),
            "{{CPF}}": e_cpf.get(),
            "{{TELEFONE}}": e_tel.get(),
            "{{LOGRADOURO}}": logradouro_var.get(),
            "{{BAIRRO}}": bairro_var.get(),
            "{{CIDADE}}": cidade_var.get(),
            "{{UF}}": uf_var.get(),
            "{{DATA}}": hoje,
            "{{DIAS}}": str(dias),
            "{{TRAFEGO}}": "" if trafego_var.get() == "SIM" else "LIBERADO SEM DIREITO A TRÁFEGO.",
            "{{ASSINATURA}}": assinatura_var.get()
        }

        registro = {
            "LOTE": lote_var.get(),
            "DATA_APREENSAO": e_data.get(),
            "TIPO_AGENTE": tipo_agente_var.get(),
            "RECOLHA": codigo_agente_var.get(),
            "DIAS": str(dias),
            "PLACA": e_placa.get(),
            "MODELO": modelo_var.get(),
            "ATENDENTE": assinatura_var.get(),
            "DATA_LIBERACAO": hoje
        }

        if not lote_var.get():
            messagebox.showerror("Erro", "O campo LOTE é obrigatório.")
            return

        pasta = config.get("pasta_saida")

        if not pasta:
            messagebox.showerror("Erro", "Selecione uma pasta para salvar os documentos.")
            return

        #Validação de campos para o Sheets
        campos_sheets = ["LOTE", "TIPO_AGENTE", "RECOLHA"]

        for c in campos_sheets:
            if not registro[c]:
                messagebox.showerror(
                    "Erro",
                    f"O campo '{c}' é obrigatório para o registro no Sheets."
                )
                return

        gerar_documento(dados, pasta)

        try:
            registrar_liberacao(registro)
            messagebox.showinfo(
                "Sucesso",
                "Termo gerado, salvo e registrado com sucesso."
            )
        except Exception as e:
           with open("erro_sheets.log", "a", encoding="utf-8") as f:
              f.write("\n" + "=" * 50 + "\n")
              f.write(datetime.now().strftime("%d/%m/%Y %H:%M:%S") + "\n")
              f.write(traceback.format_exc())

           messagebox.showwarning(
                "Aviso",
                "Documento gerado, mas não foi possível registrar no Google Sheets.\n"
                "O erro foi salvo em 'erro_sheets.log'."
            )
    def limpar_campos():
        e_placa.delete(0, tk.END)
        e_data.delete(0, tk.END)
        e_cpf.delete(0, tk.END)
        e_tel.delete(0, tk.END)

        for var in [
            marca_var, modelo_var, cor_var, chassi_var, lote_var,
            tipo_agente_var, codigo_agente_var, chave_var,
            nome_var, logradouro_var, bairro_var
        ]:
            var.set("")

        # Mantém padrões
        cidade_var.set("ITAQUAQUECETUBA")
        uf_var.set("SP")
        trafego_var.set("SIM")
        # ❗ NÃO limpa assinatura_var (atendente)

    frame_pasta = ttk.LabelFrame(container, text="PASTA DE SALVAMENTO")
    frame_pasta.pack(fill="x", padx=15, pady=6)

    label_pasta = ttk.Label(
        frame_pasta,
        text=config.get("pasta_saida", "Nenhuma pasta selecionada"),
        wraplength=400
    )
    label_pasta.pack(side="left", padx=10)

    ttk.Button(
        frame_pasta,
        text="Escolher Pasta",
        command=escolher_pasta
    ).pack(side="right", padx=10)

    frame_botoes = ttk.Frame(container)
    frame_botoes.pack(fill="x", pady=20)

    ttk.Button(
        frame_botoes,
        text="GERAR TERMO",
        command=gerar,
        width=20
    ).pack(side="left", padx=10)

    ttk.Button(
        frame_botoes,
        text="LIMPAR",
        command=limpar_campos,
        width=12
    ).pack(side="left")

    root.mainloop()
