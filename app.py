import os
import math
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from pypdf import PdfReader, PdfWriter

# Configurações do Design Moderno (Claro / Light Mode)
ctk.set_appearance_mode("Light")  
ctk.set_default_color_theme("blue")

class AppPDFCompleto:
    def __init__(self, root):
        self.root = root
        self.root.title("15º Promotoria Criminal - Ferramenta Avançada de PDF")
        self.root.geometry("750x800")
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        self.root.configure(fg_color="#f1f5f9") 

        # --- ABA PRINCIPAL ---
        self.tabview_principal = ctk.CTkTabview(
            root, width=700, height=750, 
            fg_color="#ffffff", 
            segmented_button_selected_color="#2563eb", 
            segmented_button_selected_hover_color="#1d4ed8"
        )
        self.tabview_principal.pack(fill="both", expand=True, padx=20, pady=10)

        self.aba_dividir = self.tabview_principal.add("✂️ Dividir PDF")
        self.aba_juntar = self.tabview_principal.add("🔗 Juntar PDFs")
        self.aba_extrair = self.tabview_principal.add("📄 Extrair Páginas") # <--- NOVA ABA

        # Variáveis Dividir
        self.caminho_pdf_div = tk.StringVar()
        self.total_paginas_div = 0
        self.paginas_restantes = 0
        self.blocos_personalizados = []

        # Variáveis Extrair
        self.caminho_pdf_ext = tk.StringVar()
        self.total_paginas_ext = 0

        self.construir_aba_dividir()
        self.construir_aba_juntar()
        self.construir_aba_extrair()

    # ==========================================
    # MÓDULO 1: DIVIDIR PDF
    # ==========================================
    def construir_aba_dividir(self):
        ctk.CTkLabel(self.aba_dividir, text="1. Selecione o PDF:", font=ctk.CTkFont(size=15, weight="bold"), text_color="#0f172a").pack(anchor="w", padx=10, pady=(15, 0))
        
        frame_arquivo = ctk.CTkFrame(self.aba_dividir, fg_color="transparent")
        frame_arquivo.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkEntry(frame_arquivo, textvariable=self.caminho_pdf_div, state='readonly', width=450, height=40, border_color="#cbd5e1", fg_color="#f8fafc", text_color="#333333").pack(side="left", padx=(0, 10))
        ctk.CTkButton(frame_arquivo, text="Procurar PDF", command=self.carregar_pdf_dividir, height=40, corner_radius=6, fg_color="#2563eb", hover_color="#1d4ed8").pack(side="left")

        self.lbl_info_pdf_div = ctk.CTkLabel(self.aba_dividir, text="Nenhum PDF carregado.", text_color="#64748b", font=ctk.CTkFont(size=13))
        self.lbl_info_pdf_div.pack(anchor="w", padx=10)

        ctk.CTkFrame(self.aba_dividir, height=1, fg_color="#e2e8f0").pack(fill="x", padx=10, pady=20)

        ctk.CTkLabel(self.aba_dividir, text="2. Escolha como dividir:", font=ctk.CTkFont(size=15, weight="bold"), text_color="#0f172a").pack(anchor="w", padx=10)

        self.tabview_divisao = ctk.CTkTabview(self.aba_dividir, height=320, fg_color="#f8fafc", segmented_button_selected_color="#475569")
        self.tabview_divisao.pack(fill="both", expand=True, padx=10, pady=10)

        aba_auto = self.tabview_divisao.add("Automático (2 a 9 partes)")
        aba_custom = self.tabview_divisao.add("Personalizado (Por saldo)")

        self.combo_auto = ctk.CTkComboBox(aba_auto, state="readonly", width=400, height=40, border_color="#cbd5e1", dropdown_fg_color="#ffffff", dropdown_text_color="#0f172a", text_color="#0f172a")
        self.combo_auto.pack(pady=50)
        self.combo_auto.set("Carregue um PDF primeiro...")

        self.lbl_saldo = ctk.CTkLabel(aba_custom, text="Páginas restantes: 0", font=ctk.CTkFont(size=18, weight="bold"), text_color="#10b981")
        self.lbl_saldo.pack(pady=10)

        frame_add_bloco = ctk.CTkFrame(aba_custom, fg_color="transparent")
        frame_add_bloco.pack(pady=5)
        
        ctk.CTkLabel(frame_add_bloco, text="Páginas neste bloco:", text_color="#333333").pack(side="left", padx=5)
        self.entry_qtd_bloco = ctk.CTkEntry(frame_add_bloco, width=100, height=35, border_color="#cbd5e1", text_color="#0f172a")
        self.entry_qtd_bloco.pack(side="left", padx=10)
        ctk.CTkButton(frame_add_bloco, text="+ Adicionar Bloco", command=self.adicionar_bloco, fg_color="#10b981", hover_color="#059669", height=35, corner_radius=6).pack(side="left")

        self.lista_blocos = tk.Listbox(aba_custom, height=5, bg="#ffffff", fg="#0f172a", selectbackground="#e2e8f0", selectforeground="#0f172a", highlightthickness=1, highlightcolor="#cbd5e1", highlightbackground="#cbd5e1", borderwidth=0, font=("Segoe UI", 12))
        self.lista_blocos.pack(fill="x", pady=15, padx=20)
        
        ctk.CTkButton(aba_custom, text="Desfazer Último Bloco", command=self.remover_ultimo_bloco, fg_color="#ef4444", hover_color="#dc2626", height=30, corner_radius=6).pack(pady=0)

        ctk.CTkButton(self.aba_dividir, text="⚡ GERAR PDFs DIVIDIDOS", font=ctk.CTkFont(size=15, weight="bold"), height=50, corner_radius=8, fg_color="#2563eb", hover_color="#1d4ed8", command=self.executar_divisao).pack(pady=25)

    def carregar_pdf_dividir(self):
        arquivo = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
        if arquivo:
            try:
                leitor = PdfReader(arquivo)
                self.total_paginas_div = len(leitor.pages)
                self.caminho_pdf_div.set(arquivo)
                self.lbl_info_pdf_div.configure(text=f"✅ PDF carregado! Total de páginas: {self.total_paginas_div}", text_color="#10b981")
                self.atualizar_opcoes_automaticas()
                self.reiniciar_personalizado()
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível ler o PDF:\n{e}")

    def atualizar_opcoes_automaticas(self):
        opcoes = [f"Dividir em {i} partes (aprox. {math.ceil(self.total_paginas_div / i)} páginas cada)" for i in range(2, 10)]
        self.combo_auto.configure(values=opcoes)
        self.combo_auto.set(opcoes[0])

    def reiniciar_personalizado(self):
        self.paginas_restantes = self.total_paginas_div
        self.blocos_personalizados = []
        self.lista_blocos.delete(0, tk.END)
        self.atualizar_saldo()

    def atualizar_saldo(self):
        cor = "#10b981" if self.paginas_restantes > 0 else "#ef4444"
        texto = f"Páginas restantes para alocar: {self.paginas_restantes}" if self.paginas_restantes > 0 else "Todas as páginas alocadas! ✅"
        self.lbl_saldo.configure(text=texto, text_color=cor)

    def adicionar_bloco(self):
        if self.total_paginas_div == 0: return messagebox.showwarning("Aviso", "Carregue um PDF primeiro.")
        if self.paginas_restantes <= 0: return messagebox.showinfo("Aviso", "Todas as páginas já foram alocadas!")
        
        try:
            qtd = int(self.entry_qtd_bloco.get())
            if qtd <= 0: raise ValueError
        except ValueError:
            return messagebox.showerror("Erro", "Digite um número válido e maior que zero.")
            
        if qtd > self.paginas_restantes: return messagebox.showwarning("Aviso", f"Você só tem {self.paginas_restantes} páginas restantes!")

        self.blocos_personalizados.append(qtd)
        self.lista_blocos.insert(tk.END, f"  Bloco {len(self.blocos_personalizados)}: {qtd} páginas")
        self.paginas_restantes -= qtd
        self.atualizar_saldo()
        self.entry_qtd_bloco.delete(0, tk.END)

    def remover_ultimo_bloco(self):
        if not self.blocos_personalizados: return
        self.paginas_restantes += self.blocos_personalizados.pop()
        self.lista_blocos.delete(tk.END)
        self.atualizar_saldo()

    def executar_divisao(self):
        arquivo = self.caminho_pdf_div.get()
        if not arquivo or self.total_paginas_div == 0: return messagebox.showwarning("Aviso", "Selecione um PDF primeiro.")

        tamanhos_blocos = []
        if self.tabview_divisao.get() == "Automático (2 a 9 partes)":
            num_partes = int(self.combo_auto.get().split()[2]) 
            paginas_por_parte = math.ceil(self.total_paginas_div / num_partes)
            paginas_alocadas = 0
            for i in range(num_partes):
                qtd = min(paginas_por_parte, self.total_paginas_div - paginas_alocadas)
                if qtd > 0:
                    tamanhos_blocos.append(qtd)
                    paginas_alocadas += qtd
        else:
            if self.paginas_restantes > 0 and not messagebox.askyesno("Aviso", f"Ainda faltam {self.paginas_restantes} páginas. Continuar ignorando elas?"): return
            if not self.blocos_personalizados: return messagebox.showwarning("Aviso", "Crie pelo menos um bloco.")
            tamanhos_blocos = self.blocos_personalizados

        pasta_saida = filedialog.askdirectory(title="Onde salvar os PDFs divididos?")
        if not pasta_saida: return

        try:
            leitor = PdfReader(arquivo)
            pagina_atual = 0
            for i, tamanho in enumerate(tamanhos_blocos):
                escritor = PdfWriter()
                fim = pagina_atual + tamanho
                for num in range(pagina_atual, fim): escritor.add_page(leitor.pages[num])
                
                with open(os.path.join(pasta_saida, f"Parte_{i+1}_{tamanho}pags.pdf"), "wb") as saida:
                    escritor.write(saida)
                pagina_atual = fim
            messagebox.showinfo("Sucesso!", f"PDF dividido com sucesso!\nSalvos em: {pasta_saida}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao dividir: {e}")

    # ==========================================
    # MÓDULO 2: JUNTAR PDFs
    # ==========================================
    def construir_aba_juntar(self):
        ctk.CTkLabel(self.aba_juntar, text="1. Adicione os PDFs que deseja juntar:", font=ctk.CTkFont(size=15, weight="bold"), text_color="#0f172a").pack(anchor="w", padx=20, pady=(25, 5))
        ctk.CTkButton(self.aba_juntar, text="+ Selecionar PDFs", command=self.adicionar_pdfs_juntar, height=40, corner_radius=6, fg_color="#2563eb", hover_color="#1d4ed8").pack(anchor="w", padx=20, pady=10)
        ctk.CTkLabel(self.aba_juntar, text="2. Ordene a lista abaixo (Selecione um item e use os botões):", font=ctk.CTkFont(size=13), text_color="#64748b").pack(anchor="w", padx=20, pady=(15, 0))
        
        frame_lista = ctk.CTkFrame(self.aba_juntar, fg_color="transparent")
        frame_lista.pack(pady=10, padx=20, fill="both", expand=True)

        self.listbox_pdfs = tk.Listbox(frame_lista, selectmode=tk.SINGLE, bg="#ffffff", fg="#0f172a", selectbackground="#e2e8f0", selectforeground="#0f172a", highlightthickness=1, highlightcolor="#cbd5e1", highlightbackground="#cbd5e1", borderwidth=0, font=("Segoe UI", 11))
        self.listbox_pdfs.pack(side="left", fill="both", expand=True, padx=(0, 15))

        frame_botoes = ctk.CTkFrame(frame_lista, fg_color="transparent")
        frame_botoes.pack(side="right", fill="y")
        
        ctk.CTkButton(frame_botoes, text="↑ Subir", width=120, height=35, corner_radius=6, text_color="#0f172a", fg_color="#f1f5f9", hover_color="#e2e8f0", border_width=1, border_color="#cbd5e1", command=self.mover_cima).pack(pady=5)
        ctk.CTkButton(frame_botoes, text="↓ Descer", width=120, height=35, corner_radius=6, text_color="#0f172a", fg_color="#f1f5f9", hover_color="#e2e8f0", border_width=1, border_color="#cbd5e1", command=self.mover_baixo).pack(pady=5)
        ctk.CTkButton(frame_botoes, text="🗑️ Remover", width=120, height=35, corner_radius=6, fg_color="#ef4444", hover_color="#dc2626", command=self.remover_selecionado).pack(pady=30)

        ctk.CTkButton(self.aba_juntar, text="🔗 JUNTAR PDFs AGORA", font=ctk.CTkFont(size=15, weight="bold"), height=50, corner_radius=8, fg_color="#10b981", hover_color="#059669", command=self.executar_juncao).pack(pady=25)

    def adicionar_pdfs_juntar(self):
        for arq in filedialog.askopenfilenames(filetypes=[("Arquivos PDF", "*.pdf")]): self.listbox_pdfs.insert(tk.END, f"  {arq}")

    def mover_cima(self):
        sel = self.listbox_pdfs.curselection()
        if sel and sel[0] > 0:
            idx, texto = sel[0], self.listbox_pdfs.get(sel[0])
            self.listbox_pdfs.delete(idx); self.listbox_pdfs.insert(idx - 1, texto); self.listbox_pdfs.select_set(idx - 1)

    def mover_baixo(self):
        sel = self.listbox_pdfs.curselection()
        if sel and sel[0] < self.listbox_pdfs.size() - 1:
            idx, texto = sel[0], self.listbox_pdfs.get(sel[0])
            self.listbox_pdfs.delete(idx); self.listbox_pdfs.insert(idx + 1, texto); self.listbox_pdfs.select_set(idx + 1)

    def remover_selecionado(self):
        sel = self.listbox_pdfs.curselection()
        if sel: self.listbox_pdfs.delete(sel[0])

    def executar_juncao(self):
        arquivos = self.listbox_pdfs.get(0, tk.END)
        if len(arquivos) < 2: return messagebox.showwarning("Aviso", "Adicione pelo menos 2 PDFs para juntar.")
        arquivo_final = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Arquivo PDF", "*.pdf")], initialfile="pdf_unificado.pdf")
        if not arquivo_final: return

        try:
            escritor = PdfWriter()
            for pdf in arquivos:
                for pagina in PdfReader(pdf.strip()).pages: escritor.add_page(pagina)
            with open(arquivo_final, "wb") as saida: escritor.write(saida)
            messagebox.showinfo("Sucesso", f"PDFs unidos com sucesso!\nSalvo em: {arquivo_final}")
        except Exception as e: messagebox.showerror("Erro", f"Erro ao juntar os PDFs: {e}")

    # ==========================================
    # MÓDULO 3: EXTRAIR PÁGINAS (NOVO)
    # ==========================================
    def construir_aba_extrair(self):
        ctk.CTkLabel(self.aba_extrair, text="1. Selecione o PDF de origem:", font=ctk.CTkFont(size=15, weight="bold"), text_color="#0f172a").pack(anchor="w", padx=10, pady=(15, 0))
        
        frame_arquivo = ctk.CTkFrame(self.aba_extrair, fg_color="transparent")
        frame_arquivo.pack(fill="x", padx=10, pady=5)
        
        ctk.CTkEntry(frame_arquivo, textvariable=self.caminho_pdf_ext, state='readonly', width=450, height=40, border_color="#cbd5e1", fg_color="#f8fafc", text_color="#333333").pack(side="left", padx=(0, 10))
        ctk.CTkButton(frame_arquivo, text="Procurar PDF", command=self.carregar_pdf_extrair, height=40, corner_radius=6, fg_color="#2563eb", hover_color="#1d4ed8").pack(side="left")

        self.lbl_info_pdf_ext = ctk.CTkLabel(self.aba_extrair, text="Nenhum PDF carregado.", text_color="#64748b", font=ctk.CTkFont(size=13))
        self.lbl_info_pdf_ext.pack(anchor="w", padx=10)

        ctk.CTkFrame(self.aba_extrair, height=1, fg_color="#e2e8f0").pack(fill="x", padx=10, pady=20)

        ctk.CTkLabel(self.aba_extrair, text="2. Informe as páginas que deseja extrair:", font=ctk.CTkFont(size=15, weight="bold"), text_color="#0f172a").pack(anchor="w", padx=10)
        ctk.CTkLabel(self.aba_extrair, text="Dica: Use vírgulas para separar e traços para intervalos. (Ex: 1, 5, 10-15, 20)", font=ctk.CTkFont(size=12), text_color="#64748b").pack(anchor="w", padx=10, pady=(0, 10))

        self.entry_paginas = ctk.CTkEntry(self.aba_extrair, placeholder_text="Ex: 1, 5, 8, 10-20", height=45, border_color="#cbd5e1", fg_color="#ffffff", text_color="#0f172a", font=ctk.CTkFont(size=14))
        self.entry_paginas.pack(fill="x", padx=10, pady=5)

        ctk.CTkButton(self.aba_extrair, text="📄 EXTRAIR E GERAR NOVO PDF", font=ctk.CTkFont(size=15, weight="bold"), height=50, corner_radius=8, fg_color="#10b981", hover_color="#059669", command=self.executar_extracao).pack(pady=40)

    def carregar_pdf_extrair(self):
        arquivo = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
        if arquivo:
            try:
                leitor = PdfReader(arquivo)
                self.total_paginas_ext = len(leitor.pages)
                self.caminho_pdf_ext.set(arquivo)
                self.lbl_info_pdf_ext.configure(text=f"✅ PDF carregado! Total de páginas: {self.total_paginas_ext}", text_color="#10b981")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível ler o PDF:\n{e}")

    def executar_extracao(self):
        arquivo = self.caminho_pdf_ext.get()
        if not arquivo or self.total_paginas_ext == 0: return messagebox.showwarning("Aviso", "Selecione o PDF de origem primeiro.")
        
        texto_paginas = self.entry_paginas.get().strip()
        if not texto_paginas: return messagebox.showwarning("Aviso", "Digite as páginas que deseja extrair.")

        paginas_para_extrair = set()
        try:
            # Lógica para interpretar vírgulas e traços (ranges)
            partes = texto_paginas.split(',')
            for parte in partes:
                parte = parte.strip()
                if not parte: continue
                if '-' in parte:
                    limites = parte.split('-')
                    if len(limites) != 2: raise ValueError
                    inicio, fim = int(limites[0]), int(limites[1])
                    if inicio > fim: inicio, fim = fim, inicio
                    for p in range(inicio, fim + 1):
                        if 1 <= p <= self.total_paginas_ext: paginas_para_extrair.add(p - 1) # -1 porque o pypdf conta do zero
                else:
                    p = int(parte)
                    if 1 <= p <= self.total_paginas_ext: paginas_para_extrair.add(p - 1)
        except ValueError:
            return messagebox.showerror("Erro", "Formato inválido. Use apenas números, vírgulas e traços. Ex: 1, 5, 10-15")

        if not paginas_para_extrair: return messagebox.showwarning("Aviso", "Nenhuma página válida foi selecionada para extração.")

        # Ordenar as páginas para manter a sequência lógica no novo PDF
        lista_final_paginas = sorted(list(paginas_para_extrair))

        arquivo_final = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("Arquivo PDF", "*.pdf")], initialfile="paginas_extraidas.pdf")
        if not arquivo_final: return

        try:
            leitor = PdfReader(arquivo)
            escritor = PdfWriter()
            
            for num_pagina in lista_final_paginas:
                escritor.add_page(leitor.pages[num_pagina])
                
            with open(arquivo_final, "wb") as saida:
                escritor.write(saida)
            
            messagebox.showinfo("Sucesso", f"Foram extraídas {len(lista_final_paginas)} páginas com sucesso!\nSalvo em: {arquivo_final}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao extrair as páginas: {e}")

if __name__ == "__main__":
    root = ctk.CTk()
    app = AppPDFCompleto(root)
    root.mainloop()