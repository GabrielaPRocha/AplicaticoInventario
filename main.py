import pandas as pd
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox

class ControleEstoqueApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Controle de Estoque")
        self.root.geometry("1500x400")

        self.df = pd.DataFrame(columns=["Nome", "OC/OG", "OLD USER", "FABRICANTE", "MODELO", "S/N", "ATIVO",
                                        "CARREGADOR", "VENDOR", "NF", "DATA", "Valor Unitario",
                                        "Valor Residual", "Status de Venda"])

        self.criar_interface()

    def criar_interface(self):
        self.tree = ttk.Treeview(self.root, selectmode="browse", show="headings", height=15)
        self.tree["columns"] = tuple(self.df.columns)

        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)

        self.tree.pack(expand=True, fill="both")

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Carregar Dados", command=self.carregar_dados).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Adicionar Dados", command=self.adicionar_dados).pack(side=tk.LEFT, padx=10)
    #    tk.Button(btn_frame, text="Editar Dados", command=self.preparar_editar).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Excluir Dados", command=self.excluir_dados).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Salvar Alterações", command=self.salvar_alteracoes).pack(side=tk.LEFT, padx=10)

        self.tree.bind("<ButtonRelease-1>", self.preparar_editar)

    def carregar_dados(self):
        try:
            self.df = pd.read_excel("teste1912.xlsx")
            self.atualizar_treeview()
        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo não encontrado.")

    def adicionar_dados(self):
        novo_registro = self.obter_dados_adicao()
        if novo_registro:
            novo_df = pd.DataFrame([novo_registro], columns=self.df.columns)
            self.df = pd.concat([self.df, novo_df], ignore_index=True)
            self.atualizar_treeview()

    def preparar_editar(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            coluna_clicada = self.tree.identify_column(event.x)
            coluna_nome = self.tree.heading(coluna_clicada, "text")
            linha_clicada = self.tree.index(selected_item[0])

            # Chame a função para editar dados
            self.editar_dados(coluna_nome, linha_clicada)

    def editar_dados(self, coluna_nome, linha_clicada):
        if coluna_nome is not None:
            # Obtenha o valor atual na célula
            valor_atual = self.df.at[linha_clicada, coluna_nome]

            # Crie uma janela de diálogo para editar o valor
            novo_valor = simpledialog.askstring("Editar Valor", f"Digite o novo valor para {coluna_nome}:",
                                                initialvalue=valor_atual)

            if novo_valor is not None:
                self.df.at[linha_clicada, coluna_nome] = novo_valor
                self.atualizar_treeview()

    def excluir_dados(self):
        selected_item = self.tree.selection()
        if selected_item:
            linha_clicada = self.tree.index(selected_item[0])
            self.df = self.df.drop(index=linha_clicada)
            self.atualizar_treeview()

    def salvar_alteracoes(self):
        self.df.to_excel("teste1912.xlsx", index=False)

    def atualizar_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for _, row in self.df.iterrows():
            self.tree.insert("", "end", values=tuple(row))

    def obter_dados_adicao(self):
        novo_registro = {}
        for coluna in self.df.columns:
            while True:
                novo_valor = simpledialog.askstring("Adicionar Dados", f"Digite o valor para {coluna}:")
                if novo_valor is None:
                    return None  # O usuário cancelou a adição
                elif novo_valor.strip():  # Verifica se o valor não está vazio ou contém apenas espaços
                    novo_registro[coluna] = novo_valor
                    break
                else:
                    messagebox.showwarning("Aviso", "Por favor, insira um valor válido.")

        return novo_registro

if __name__ == "__main__":
    root = tk.Tk()
    app = ControleEstoqueApp(root)
    root.mainloop()
