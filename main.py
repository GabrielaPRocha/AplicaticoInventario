import pandas as pd
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
from datetime import datetime

class ControleEstoqueApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Controle de Estoque")
        self.root.geometry("1500x400")
        self.entry_pesquisa = tk.Entry(self.root)
        self.entry_pesquisa.pack(pady=10, padx=10, side="top")

        self.df = pd.DataFrame(columns=["Nome", "OC/OG", "OLD USER", "FABRICANTE", "MODELO", "S/N", "ATIVO",
                                        "CARREGADOR", "VENDEDOR", "NF", "DATA", "Valor Unitario",
                                        "Valor Residual", "Status de Venda"])

        # Adiciona coluna para rastrear a última data de alteração
        self.df["Data Alteracao"] = ""

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
        tk.Button(btn_frame, text="Excluir Dados", command=self.excluir_dados).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Salvar Alterações", command=self.perguntar_salvar_alteracoes).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="Pesquisar", command=self.pesquisar).pack(side="top", pady=5)

        self.tree.bind("<ButtonRelease-1>", self.preparar_editar)

    def pesquisar(self):
        termo_pesquisa = self.entry_pesquisa.get().strip()
        if termo_pesquisa:
            # Cria um novo DataFrame apenas com as linhas filtradas
            df_filtrado = self.df[
                self.df.apply(lambda row: termo_pesquisa.lower() in row.astype(str).str.lower().values.any(), axis=1)]
            # Atualiza o Treeview com o novo DataFrame
            self.atualizar_treeview(df_filtrado)
        else:
            # Se a pesquisa estiver vazia, atualize com o DataFrame original
            self.atualizar_treeview()



    def carregar_dados(self):
        try:
            self.df = pd.read_excel("teste1912.xlsx")
            # Adiciona a coluna "Data Alteracao" para o DataFrame carregado
            self.df["Data Alteracao"] = ""
            self.atualizar_treeview()
        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo não encontrado.")

    def adicionar_dados(self):
        novo_registro = self.obter_dados_adicao()
        if novo_registro:
            novo_df = pd.DataFrame([novo_registro], columns=self.df.columns)
            novo_df["Data Alteracao"] = ""  # Adiciona a coluna "Data Alteracao" para o novo registro
            self.df = pd.concat([self.df, novo_df], ignore_index=True)
            self.atualizar_treeview()

    def preparar_editar(self, event):
        # Remova a marcação da célula como alterada ao editar
        self.tree.tag_configure("red", background="white")

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
                # Atualiza o valor na célula
                self.df.at[linha_clicada, coluna_nome] = novo_valor
                # Atualiza a data da alteração
                self.df.at[linha_clicada, "Data Alteracao"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.atualizar_treeview()

    def excluir_dados(self):
        selected_item = self.tree.selection()
        if selected_item:
            linha_clicada = self.tree.index(selected_item[0])
            self.df = self.df.drop(index=linha_clicada)
            self.atualizar_treeview()

    def perguntar_salvar_alteracoes(self):
        # Pergunta se realmente deseja salvar as alterações
        resposta = messagebox.askyesno("Salvar Alterações", "Deseja realmente salvar as alterações?")
        if resposta:
            self.salvar_alteracoes()

    def salvar_alteracoes(self):
        # Salva apenas as linhas alteradas
        df_alterado = self.df[self.df["Data Alteracao"].notnull()]
        if not df_alterado.empty:
            df_alterado.to_excel("teste1912.xlsx", index=False)
            # Reseta a data da alteração
            self.df["Data Alteracao"] = ""
            self.atualizar_treeview()
        else:
            messagebox.showinfo("Informação", "Nenhuma alteração a ser salva.")

    def atualizar_treeview(self, df=None):
        for item in self.tree.get_children():
            self.tree.delete(item)

        if df is None:
            df = self.df

        for _, row in df.iterrows():
            # Define a cor de fundo da célula como vermelha se a data de alteração estiver presente
            bg_color = "red" if row["Data Alteracao"] else "white"
            self.tree.insert("", "end", values=tuple(row), tags=(bg_color,))
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
