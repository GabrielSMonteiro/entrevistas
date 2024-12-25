import tkinter as tk
from tkinter import ttk, Text
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
import os

class FormularioVisitaMetlife:
    def __init__(self, root):
        self.root = root
        self.root.title("Ficha de Visita - MetLife")
        
        # Criar notebook para organizar as seções
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(pady=10, expand=True)
        
        # Abas principais
        self.aba_dados = ttk.Frame(self.notebook)
        self.aba_mapeamento = ttk.Frame(self.notebook)
        self.aba_checklist = ttk.Frame(self.notebook)
        
        self.notebook.add(self.aba_dados, text="Dados Principais")
        self.notebook.add(self.aba_mapeamento, text="Mapeamento")
        self.notebook.add(self.aba_checklist, text="Checklist")
        
        self.criar_campos_principais()
        self.criar_campos_mapeamento()
        self.criar_campos_checklist()
        self.criar_botoes()

    def criar_campos_principais(self):

        # Dados básicos
        tk.Label(self.aba_dados, text="Data:").grid(row=0, column=0, padx=5, pady=5)
        self.data = tk.Entry(self.aba_dados)
        self.data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.data.grid(row=0, column=1, padx=5, pady=5)
        
        tk.Label(self.aba_dados, text="Corretora:").grid(row=1, column=0, padx=5, pady=5)
        self.corretora = tk.Entry(self.aba_dados, width=40)
        self.corretora.grid(row=1, column=1, padx=5, pady=5)
        
        tk.Label(self.aba_dados, text="Classificação Be Metlife:").grid(row=2, column=0, padx=5, pady=5)
        self.classificacao = tk.Entry(self.aba_dados, width=40)
        self.classificacao.grid(row=2, column=1, padx=5, pady=5)
        
        # Tipo de visita
        tk.Label(self.aba_dados, text="Tipo de visita:").grid(row=3, column=0, padx=5, pady=5)
        self.tipo_visita_frame = tk.Frame(self.aba_dados)
        self.tipo_visita_frame.grid(row=3, column=1)
        self.tipo_visita = tk.StringVar(value="")  # Inicializa vazio
        self.tipo_visita = tk.StringVar()
        tk.Radiobutton(self.tipo_visita_frame, text="Virtual", variable=self.tipo_visita, value="Virtual").pack(side=tk.LEFT)
        tk.Radiobutton(self.tipo_visita_frame, text="Presencial", variable=self.tipo_visita, value="Presencial").pack(side=tk.LEFT)
        
        tk.Label(self.aba_dados, text="Participantes:").grid(row=4, column=0, padx=5, pady=5)
        self.participantes = tk.Entry(self.aba_dados, width=40)
        self.participantes.grid(row=4, column=1, padx=5, pady=5)
        
        # Objetivo da visita
        tk.Label(self.aba_dados, text="Objetivo da visita:").grid(row=5, column=0, padx=5, pady=5)
        self.objetivo = tk.Text(self.aba_dados, height=3, width=40)
        self.objetivo.grid(row=5, column=1, padx=5, pady=5)
        
        # Checkboxes para assuntos
        self.assuntos_frame = tk.LabelFrame(self.aba_dados, text="Assuntos")
        self.assuntos_frame.grid(row=6, column=0, columnspan=2, pady=10)
        
        self.assuntos = {
            "Operacionais e Sinistros": tk.BooleanVar(),
            "Vendas Novas": tk.BooleanVar(),
            "Renovações": tk.BooleanVar(),
            "Treinamento e Suporte": tk.BooleanVar()
        }
        
        for i, (texto, var) in enumerate(self.assuntos.items()):
            tk.Checkbutton(self.assuntos_frame, text=texto, variable=var).grid(row=0, column=i, padx=5)

    def criar_campos_mapeamento(self):

        

        # Grupo corretor ativo
        tk.Label(self.aba_mapeamento, text="Grupo corretor ativo em vendas?").grid(row=0, column=0, padx=5, pady=5)
        self.grupo_ativo_frame = tk.Frame(self.aba_mapeamento)
        self.grupo_ativo_frame.grid(row=0, column=1)
        self.grupo_ativo = tk.StringVar(value="")  # Inicializa vazio
        self.grupo_ativo = tk.StringVar()
        tk.Radiobutton(self.grupo_ativo_frame, text="Sim", variable=self.grupo_ativo, value="Sim").pack(side=tk.LEFT)
        tk.Radiobutton(self.grupo_ativo_frame, text="Não", variable=self.grupo_ativo, value="Não").pack(side=tk.LEFT)
        self.valor_rs = tk.Entry(self.grupo_ativo_frame, width=10)
        self.valor_rs.pack(side=tk.LEFT)
        tk.Label(self.grupo_ativo_frame, text="mil").pack(side=tk.LEFT)
        
        # Campos de cotação
        campos_cotacao = [
            "Há alguma cotação potencial Corporate Dental? Quais?",
            "Há alguma cotação potencial Life Corporate? Quais?",
            "Há algum Negócio Estruturado? Quais?",
            "Tem relacionamento direto com algum Sindicato Laboral ou Patronal? Quais?"
        ]
        
        self.cotacoes = {}
        for i, campo in enumerate(campos_cotacao):
            tk.Label(self.aba_mapeamento, text=campo).grid(row=i+1, column=0, padx=5, pady=5)
            self.cotacoes[campo] = tk.Text(self.aba_mapeamento, height=2, width=40)
            self.cotacoes[campo].grid(row=i+1, column=1, padx=5, pady=5)

    def criar_campos_checklist(self):
        # Checklist
        self.checklist_items = [
            "PME", "Mapeamento", "Dental", "Cotações em Aberto",
            "Projetos (Distribuição, Facility)", "Prestamista", "Prospecções",
            "Ações Metlife (Be Metlife)", "Problemas Operacionais",
            "Oportunidades Estruturadas", "Renovações", "Campanhas"
        ]
        
        self.checklist_vars = {}
        for i, item in enumerate(self.checklist_items):
            self.checklist_vars[item] = tk.BooleanVar()
            tk.Checkbutton(self.aba_checklist, text=item, variable=self.checklist_vars[item]).grid(
                row=i//3, column=i%3, sticky="w", padx=5, pady=5
            )
        
        # Compromissos e Devolutivas
        tk.Label(self.aba_checklist, text="Compromissos / Devolutivas:").grid(
            row=len(self.checklist_items)//3 + 1, column=0, columnspan=3, pady=10
        )
        self.compromissos = tk.Text(self.aba_checklist, height=4, width=60)
        self.compromissos.grid(row=len(self.checklist_items)//3 + 2, column=0, columnspan=3, pady=5)

    def criar_botoes(self):
        tk.Button(self.root, text="Gerar PDF", command=self.gerar_pdf).pack(side=tk.LEFT, padx=5, pady=10)
        tk.Button(self.root, text="Limpar", command=self.limpar_campos).pack(side=tk.LEFT, padx=5, pady=10)

    def gerar_pdf(self):
        filename = f"visita_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        doc = SimpleDocTemplate(filename, pagesize=letter)
        elementos = []
        styles = getSampleStyleSheet()
        
        # Dados principais
        dados_principais = [
            ["FICHA DE VISITA METLIFE"],
            ["Data", self.data.get()],
            ["Corretora", self.corretora.get()],
            ["Classificação Be Metlife", self.classificacao.get()],
            ["Tipo de Visita", self.tipo_visita.get()],
            ["Participantes", self.participantes.get()],
            ["Objetivo da Visita", self.objetivo.get("1.0", tk.END).strip()]
        ]
        
        # Assuntos abordados
        assuntos_marcados = [k for k, v in self.assuntos.items() if v.get()]
        dados_principais.append(["Assuntos Abordados", ", ".join(assuntos_marcados)])
        
        # Mapeamento
        dados_mapeamento = [
            ["MAPEAMENTO NA VISITA"],
            ["Grupo corretor ativo em vendas?", 
            f"{'Sim' if self.grupo_ativo.get() == 'Sim' else 'Não'} {self.valor_rs.get() if self.grupo_ativo.get() == 'Sim' else ''}"],
        ]
        
        # Adicionar cotações
        for campo, texto in self.cotacoes.items():
            dados_mapeamento.append([campo, texto.get("1.0", tk.END).strip()])
        
        # Potencial de Distribuição
        potencial_distribuicao = [
            ["POTENCIAL DE DISTRIBUIÇÃO"],
            ["Associações, Federações", ""],
            ["Rede de franquias e varejistas", ""],
            ["Cooperativas de crédito, consórcio", ""],
            ["Sindicatos", ""],
            ["Outros", ""]
        ]
        
        # Checklist
        checklist_marcados = [
            ["CHECKLIST - O QUE FOI ABORDADO?"]
        ]
        for item, var in self.checklist_vars.items():
            if var.get():
                checklist_marcados.append([item, "✓"])
            else:
                checklist_marcados.append([item, ""])
        
        # Compromissos
        compromissos = [
            ["COMPROMISSOS E DEVOLUTIVAS"],
            ["", self.compromissos.get("1.0", tk.END).strip()]
        ]
        
        # Estilo das tabelas
        estilo_tabela = TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
        ])
        
        # Criar e adicionar todas as tabelas
        for dados in [dados_principais, dados_mapeamento, potencial_distribuicao, checklist_marcados, compromissos]:
            tabela = Table(dados)
            tabela.setStyle(estilo_tabela)
            elementos.append(tabela)
            elementos.append(Spacer(1, 20))  # Espaço entre tabelas
        
        # Construir o PDF
        doc.build(elementos)
        os.startfile(filename)


    def limpar_campos(self):
        # Implementar limpeza de todos os campos
        pass

if __name__ == "__main__":
    root = tk.Tk()
    app = FormularioVisitaMetlife(root)
    root.mainloop()
