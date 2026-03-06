import customtkinter as ctk
from PIL import Image
from gerador import GeradorKit
import threading
import os
import re
from tkinter import filedialog

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("RMKit - Gerador de Documentos Jurídicos")
app.geometry("560x720")
app.minsize(500, 580)

# Container principal
container = ctk.CTkFrame(app, corner_radius=15)
container.pack(pady=15, padx=15, fill="both", expand=True)

# Scrollable frame
scroll = ctk.CTkScrollableFrame(container, scrollbar_fg_color="transparent")
scroll.pack(fill="both", expand=True, padx=10, pady=10)

# Logo
try:
    logo_img = ctk.CTkImage(Image.open("RMKIT.png"), size=(280, 130))
    logo_label = ctk.CTkLabel(scroll, image=logo_img, text="")
    logo_label.pack(pady=(20, 5))
except:
    ctk.CTkLabel(scroll, text="RMKit", font=("Segoe UI", 30, "bold")).pack(pady=(30, 10))

# Título
titulo = ctk.CTkLabel(scroll, text="Gerador de Documentos Jurídicos",
                      font=("Segoe UI", 18, "bold"))
titulo.pack(pady=(0, 25))

# Função para criar campo com label em cima
def campo_com_label(texto_label):
    frame = ctk.CTkFrame(scroll, fg_color="transparent")
    frame.pack(fill="x", pady=8, padx=5)
    
    label = ctk.CTkLabel(frame, text=texto_label, anchor="w",
                         font=("Segoe UI", 12))
    label.pack(fill="x", padx=3, pady=(0, 2))
    
    entry = ctk.CTkEntry(frame, height=38, corner_radius=8,
                         fg_color="#2b2b2b", border_color="#3b3b3b")
    entry.pack(fill="x")
    
    return entry

# Dicionário para armazenar os campos
entries = {}

# Campos comuns
entries["nome"] = campo_com_label("Nome Completo")
entries["nacionalidade"] = campo_com_label("Nacionalidade")
entries["estado_civil"] = campo_com_label("Estado Civil")
entries["profissao"] = campo_com_label("Profissão")

# RG com formatação automática
frame_rg = ctk.CTkFrame(scroll, fg_color="transparent")
frame_rg.pack(fill="x", pady=8, padx=5)
ctk.CTkLabel(frame_rg, text="RG", anchor="w", font=("Segoe UI", 12)).pack(fill="x", padx=3, pady=(0, 2))
rg_entry = ctk.CTkEntry(frame_rg, height=38, corner_radius=8, fg_color="#2b2b2b", border_color="#3b3b3b")
rg_entry.pack(fill="x")
entries["rg"] = rg_entry

# CPF com formatação automática
frame_cpf = ctk.CTkFrame(scroll, fg_color="transparent")
frame_cpf.pack(fill="x", pady=8, padx=5)
ctk.CTkLabel(frame_cpf, text="CPF", anchor="w", font=("Segoe UI", 12)).pack(fill="x", padx=3, pady=(0, 2))
cpf_entry = ctk.CTkEntry(frame_cpf, height=38, corner_radius=8, fg_color="#2b2b2b", border_color="#3b3b3b")
cpf_entry.pack(fill="x")
entries["cpf"] = cpf_entry

# Demais campos
entries["rua"] = campo_com_label("Rua")
entries["numero"] = campo_com_label("Número")
entries["bairro"] = campo_com_label("Bairro")
entries["cep"] = campo_com_label("CEP")
entries["cidade"] = campo_com_label("Cidade")
entries["estado"] = campo_com_label("Estado (UF)")
entries["tipo_acao"] = campo_com_label("Tipo de Ação")
entries["honorarios_fixos"] = campo_com_label("Honorários Fixos (ex: 03 salários)")
entries["honorarios_exito"] = campo_com_label("Honorários de Êxito (ex: 30%)")
entries["multa"] = campo_com_label("Multa (cláusula 8)")

# Funções de formatação
def formatar_rg(event=None):
    texto = rg_entry.get()
    # Remove tudo que não for dígito
    numeros = re.sub(r'\D', '', texto)
    
    if len(numeros) > 9:
        numeros = numeros[:9]
    
    # Aplica máscara XX.XXX.XXX
    if len(numeros) <= 2:
        novo = numeros
    elif len(numeros) <= 5:
        novo = f"{numeros[:2]}.{numeros[2:]}"
    elif len(numeros) <= 8:
        novo = f"{numeros[:2]}.{numeros[2:5]}.{numeros[5:]}"
    else:
        novo = f"{numeros[:2]}.{numeros[2:5]}.{numeros[5:8]}"
    
    # Atualiza só se mudou (evita loop infinito)
    if novo != texto:
        rg_entry.delete(0, 'end')
        rg_entry.insert(0, novo)

def formatar_cpf(event=None):
    texto = cpf_entry.get()
    numeros = re.sub(r'\D', '', texto)
    
    if len(numeros) > 11:
        numeros = numeros[:11]
    
    # Aplica máscara XXX.XXX.XXX-XX
    if len(numeros) <= 3:
        novo = numeros
    elif len(numeros) <= 6:
        novo = f"{numeros[:3]}.{numeros[3:]}"
    elif len(numeros) <= 9:
        novo = f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:]}"
    else:
        novo = f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:9]}-{numeros[9:]}"
    
    if novo != texto:
        cpf_entry.delete(0, 'end')
        cpf_entry.insert(0, novo)

# Bind dos eventos
rg_entry.bind('<KeyRelease>', formatar_rg)
cpf_entry.bind('<KeyRelease>', formatar_cpf)

# Frame dos botões
botoes_frame = ctk.CTkFrame(scroll, fg_color="transparent")
botoes_frame.pack(pady=25)

gerar_btn = ctk.CTkButton(botoes_frame, text="⚡ GERAR KIT", width=130, height=40,
                          corner_radius=12, fg_color="#1f6eaa", hover_color="#144870")
gerar_btn.grid(row=0, column=0, padx=6)

limpar_btn = ctk.CTkButton(botoes_frame, text="🗑️ LIMPAR", width=130, height=40,
                           corner_radius=12, fg_color="#3b3b3b", hover_color="#505050")
limpar_btn.grid(row=0, column=1, padx=6)

sair_btn = ctk.CTkButton(botoes_frame, text="🚪 SAIR", width=130, height=40,
                         corner_radius=12, fg_color="#aa3333", hover_color="#882222")
sair_btn.grid(row=0, column=2, padx=6)

# Botão ABRIR PDF (inicialmente desabilitado)
abrir_btn = ctk.CTkButton(botoes_frame, text="📂 ABRIR PDF", width=130, height=40,
                          corner_radius=12, fg_color="#2b5e2b", hover_color="#1e451e",
                          state="disabled", command=lambda: abrir_pdf())
abrir_btn.grid(row=1, column=0, columnspan=3, pady=(15, 0), sticky="ew")

# Barra de status
status = ctk.CTkLabel(app, text="✅ Pronto", anchor="w", height=20)
status.pack(side="bottom", fill="x", padx=15, pady=(0, 10))

# Progress bar
progress = ctk.CTkProgressBar(app, mode="indeterminate", height=4)
progress.pack(side="bottom", fill="x", padx=15, pady=(0, 5))
progress.pack_forget()

# Variável para armazenar último PDF gerado
ultimo_pdf = None

def abrir_pdf():
    if ultimo_pdf and os.path.exists(ultimo_pdf):
        os.startfile(ultimo_pdf)
    else:
        status.configure(text="❌ Nenhum PDF disponível")

# ========== LÓGICA ==========
gerador = GeradorKit()

def coletar_dados():
    return {
        '{{NOME}}': entries["nome"].get().strip(),
        '{{NACIONALIDADE}}': entries["nacionalidade"].get().strip(),
        '{{ESTADO_CIVIL}}': entries["estado_civil"].get().strip(),
        '{{PROFISSAO}}': entries["profissao"].get().strip(),
        '{{RG}}': entries["rg"].get().strip(),
        '{{CPF}}': entries["cpf"].get().strip(),
        '{{RUA}}': entries["rua"].get().strip(),
        '{{NUMERO}}': entries["numero"].get().strip(),
        '{{BAIRRO}}': entries["bairro"].get().strip(),
        '{{CEP}}': entries["cep"].get().strip(),
        '{{CIDADE}}': entries["cidade"].get().strip(),
        '{{ESTADO}}': entries["estado"].get().strip(),
        '{{TIPO_ACAO}}': entries["tipo_acao"].get().strip(),
        '{{HONORARIOS_FIXOS}}': entries["honorarios_fixos"].get().strip() or "03 (três) salários-mínimos",
        '{{HONORARIOS_EXITO}}': entries["honorarios_exito"].get().strip() or "30% (trinta por cento)",
        '{{MULTA_RESCISAO}}': entries["multa"].get().strip() or "01 (um) salário-mínimo"
    }

def gerar_thread():
    global ultimo_pdf
    dados = coletar_dados()
    
    if not dados['{{NOME}}'] or not dados['{{CPF}}'] or not dados['{{RG}}']:
        status.configure(text="❌ Preencha Nome, CPF e RG!")
        progress.pack_forget()
        gerar_btn.configure(state="normal")
        return

    template_proc = "procuração.docx"
    template_cont = "contrato.docx"
    if not os.path.exists(template_proc) or not os.path.exists(template_cont):
        status.configure(text="❌ Templates não encontrados!")
        progress.pack_forget()
        gerar_btn.configure(state="normal")
        return

    nome_padrao = f"Kit_{dados['{{NOME}}'].replace(' ', '_')}.pdf"
    arquivo = filedialog.asksaveasfilename(
        title="Salvar kit como",
        defaultextension=".pdf",
        initialfile=nome_padrao,
        filetypes=[("PDF", "*.pdf")]
    )
    if not arquivo:
        status.configure(text="✅ Pronto")
        progress.pack_forget()
        gerar_btn.configure(state="normal")
        return

    try:
        gerador.gerar_kit(template_proc, template_cont, dados, arquivo)
        ultimo_pdf = arquivo
        status.configure(text=f"✅ PDF salvo: {os.path.basename(arquivo)}")
        # Ativar botão ABRIR PDF e deixar verde
        abrir_btn.configure(state="normal", fg_color="#2e7d32", hover_color="#1e5622")
    except Exception as e:
        status.configure(text=f"❌ Erro: {str(e)[:50]}...")
    finally:
        progress.pack_forget()
        gerar_btn.configure(state="normal")

def iniciar_geracao():
    gerar_btn.configure(state="disabled")
    abrir_btn.configure(state="disabled", fg_color="#2b5e2b")  # volta cor original
    status.configure(text="⏳ Gerando...")
    progress.pack(side="bottom", fill="x", padx=15, pady=(0, 5))
    progress.start()
    threading.Thread(target=gerar_thread, daemon=True).start()

def limpar_campos():
    for entry in entries.values():
        entry.delete(0, 'end')
    status.configure(text="✅ Campos limpos")
    abrir_btn.configure(state="disabled", fg_color="#2b5e2b")

gerar_btn.configure(command=iniciar_geracao)
limpar_btn.configure(command=limpar_campos)
sair_btn.configure(command=app.quit)

app.mainloop()