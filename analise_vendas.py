import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import matplotlib.pyplot as plt
import os
from tabulate import tabulate
from PIL import Image, ImageTk

# Função para ler o arquivo Excel e fazer a análise temporal
def ler_excel(arquivo_excel, diretorio_saida, data_inicio, data_fim):
    try:
        df = pd.read_excel(arquivo_excel)
        
        # Verificar se as colunas necessárias estão presentes
        required_columns = ['dhEmi', 'vProd', 'Chave_de_Acesso', 'xProd']
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Coluna necessária não encontrada: {col}")
        
        # Processar a coluna "dhEmi"
        df['dhEmi'] = pd.to_datetime(df['dhEmi'])
         
        # Filtrar os dados com base nas datas fornecidas
        if data_inicio and data_fim:
            data_inicio = pd.to_datetime(data_inicio, dayfirst=True).tz_localize(df['dhEmi'].dt.tz)
            data_fim = pd.to_datetime(data_fim, dayfirst=True).tz_localize(df['dhEmi'].dt.tz)
            df = df[(df['dhEmi'] >= data_inicio) & (df['dhEmi'] <= data_fim)]
        
        if df.empty:
            raise ValueError("Não há dados no intervalo de datas fornecido.")
        
        df['Data'] = df['dhEmi'].dt.date
        df['Hora'] = df['dhEmi'].dt.hour
        df['Mes'] = df['dhEmi'].dt.to_period('M')

        # Análise de vendas por dia
        vendas_por_dia = df.groupby('Data').size()

        # Análise de vendas por hora (média mensal)
        vendas_por_hora = df.groupby('Hora').size() / df['dhEmi'].dt.month.nunique()

        # Calcular o ticket médio por dia e quantidade de clientes por dia
        ticket_medio_por_dia = df.groupby('Data')['vProd'].mean()
        clientes_por_dia = df.groupby('Data')['Chave_de_Acesso'].nunique()

        # Calcular o ticket médio por mês e quantidade de clientes por mês
        ticket_medio_por_mes = df.groupby('Mes')['vProd'].mean()
        clientes_por_mes = df.groupby('Mes')['Chave_de_Acesso'].nunique()

        # Lista dos 30 produtos mais vendidos por mês por valor
        top_produtos_mes = df.groupby(['Mes', 'xProd'])['vProd'].sum().reset_index()
        top_produtos_mes = top_produtos_mes.sort_values(by=['Mes', 'vProd'], ascending=[True, False])
        top_30_produtos_mes = top_produtos_mes.groupby('Mes').head(30)

        # Percentual do faturamento dos produtos mais vendidos
        total_faturamento = df['vProd'].sum()
        produtos_faturamento = df.groupby('xProd')['vProd'].sum().sort_values(ascending=False).reset_index()
        produtos_faturamento['Percentual'] = (produtos_faturamento['vProd'] / total_faturamento) * 100
        
        # Lista com percentuais dos produtos mais vendidos
        percentuais = [5, 10, 15, 20, 25, 30, 50]
        resumo_percentual = {}
        for p in percentuais:
            percentual_faturamento = produtos_faturamento.head(p)['Percentual'].sum()
            resumo_percentual[f'Top {p} Produtos'] = percentual_faturamento
        
        # Análise dos 10 produtos mais vendidos
        top_10_produtos = produtos_faturamento.head(10)['xProd']
        
        produtos_comprados_juntos = {}
        
        for produto in top_10_produtos:
            chaves_acesso = df[df['xProd'] == produto]['Chave_de_Acesso'].unique()
            df_juntos = df[df['Chave_de_Acesso'].isin(chaves_acesso)]
            produtos_contagem = df_juntos[df_juntos['xProd'] != produto]['xProd'].value_counts().head(10)
            produtos_comprados_juntos[produto] = produtos_contagem

        # Criar pasta com o nome do arquivo Excel (sem extensão) no diretório selecionado
        pasta_nome = os.path.splitext(os.path.basename(arquivo_excel))[0]
        caminho_pasta = os.path.join(diretorio_saida, pasta_nome)
        if not os.path.exists(caminho_pasta):
            os.makedirs(caminho_pasta)

        # Plotar gráfico de vendas por dia e salvar
        plt.figure(figsize=(12, 6))
        vendas_por_dia.plot(kind='bar')
        plt.title('Vendas por Dia')
        plt.xlabel('Data')
        plt.ylabel('Número de Vendas')
        plt.xticks(rotation=45)
        plt.tight_layout()
        caminho_imagem_dia = os.path.join(caminho_pasta, 'vendas_por_dia.png')
        plt.savefig(caminho_imagem_dia)
        plt.close()

        # Plotar gráfico de vendas por hora e salvar
        plt.figure(figsize=(12, 6))
        vendas_por_hora.plot(kind='bar')
        plt.title('Vendas por Hora (Média Mensal)')
        plt.xlabel('Hora')
        plt.ylabel('Número Médio de Vendas')
        plt.xticks(rotation=0)
        plt.tight_layout()
        caminho_imagem_hora = os.path.join(caminho_pasta, 'vendas_por_hora.png')
        plt.savefig(caminho_imagem_hora)
        plt.close()

        # Criar resumo detalhado diário
        total_vendas = vendas_por_dia.sum()
        resumo = pd.DataFrame({
            'Data': vendas_por_dia.index,
            'Número de Vendas': vendas_por_dia.values,
            'Percentual': (vendas_por_dia.values / total_vendas) * 100,
            'Ticket Médio': ticket_medio_por_dia.values,
            'Quantidade de Clientes': clientes_por_dia.values
        })
        resumo['Percentual'] = resumo['Percentual'].map('{:.2f}%'.format)

        # Exibir tabela de resumo no chat
        print("Resumo das Vendas por Dia:")
        print(tabulate(resumo, headers='keys', tablefmt='grid'))

        # Criar resumo mensal
        resumo_mensal = pd.DataFrame({
            'Mês': ticket_medio_por_mes.index.astype(str),
            'Ticket Médio': ticket_medio_por_mes.values,
            'Quantidade de Clientes': clientes_por_mes.values
        })

        print("\nResumo das Vendas por Mês:")
        print(tabulate(resumo_mensal, headers='keys', tablefmt='grid'))

        # Exibir os 30 produtos mais vendidos por mês
        print("\nTop 30 Produtos Mais Vendidos por Mês:")
        for mes, group in top_30_produtos_mes.groupby('Mes'):
            print(f'\nMês: {mes}')
            print(tabulate(group[['xProd', 'vProd']], headers=['Produto', 'Valor'], tablefmt='grid'))

        # Exibir percentuais dos produtos mais vendidos
        print("\nPercentual do Faturamento dos Produtos Mais Vendidos:")
        for key, value in resumo_percentual.items():
            print(f'{key}: {value:.2f}%')

        # Exibir produtos comprados juntos com os 10 itens mais vendidos
        print("\nProdutos Comprados Juntos com os 10 Itens Mais Vendidos:")
        for produto, contagem in produtos_comprados_juntos.items():
            print(f'\nProduto: {produto}')
            print(tabulate(contagem.reset_index().rename(columns={'index': 'Produto', 'xProd': 'Quantidade'}), headers='keys', tablefmt='grid'))

        # Salvar resumo em formato Markdown
        caminho_resumo = os.path.join(caminho_pasta, 'resumo_vendas.md')
        with open(caminho_resumo, 'w') as f:
            f.write('# Resumo das Vendas\n\n')
            f.write('## Gráficos\n')
            f.write(f'![Vendas por Dia]({caminho_imagem_dia})\n')
            f.write(f'![Vendas por Hora]({caminho_imagem_hora})\n\n')
            f.write('## Tabela Resumo por Dia\n\n')
            f.write(resumo.to_markdown(index=False))
            f.write('\n\n## Tabela Resumo por Mês\n\n')
            f.write(resumo_mensal.to_markdown(index=False))
            f.write('\n\n## Top 30 Produtos Mais Vendidos por Mês\n\n')
            for mes, group in top_30_produtos_mes.groupby('Mes'):
                f.write(f'\n### Mês: {mes}\n')
                f.write(group[['xProd', 'vProd']].to_markdown(index=False))
            f.write('\n\n## Percentual do Faturamento dos Produtos Mais Vendidos\n\n')
            for key, value in resumo_percentual.items():
                f.write(f'- {key}: {value:.2f}%\n')
            f.write('\n\n## Produtos Comprados Juntos com os 10 Itens Mais Vendidos\n\n')
            for produto, contagem in produtos_comprados_juntos.items():
                f.write(f'\n### Produto: {produto}\n')
                f.write(contagem.reset_index().rename(columns={'index': 'Produto', 'xProd': 'Quantidade'}).to_markdown(index=False))
        
        messagebox.showinfo("Sucesso", f"Análise concluída e arquivos salvos em '{caminho_pasta}'!")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao ler o arquivo Excel ou processar os dados:\n{str(e)}")

# Função para abrir o diálogo de seleção de arquivo e diretório
def selecionar_arquivo():
    arquivo_excel = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
    )
    if arquivo_excel:
        diretorio_saida = filedialog.askdirectory(title="Selecione o diretório para salvar a pasta")
        if diretorio_saida:
            data_inicio = entrada_data_inicio.get()
            data_fim = entrada_data_fim.get()
            ler_excel(arquivo_excel, diretorio_saida, data_inicio, data_fim)

# Criar a janela principal
root = tk.Tk()
root.title("Leitor de Arquivo Excel")

# Definir o tamanho da janela
window_width = 800
window_height = 600
root.geometry(f"{window_width}x{window_height}")

# Centralizar a janela na tela
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_top = int(screen_height / 2 - window_height / 2)
position_right = int(screen_width / 2 - window_width / 2)
root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")

# Estilizar a interface
style = ttk.Style(root)
style.theme_use('clam')  # Temas disponíveis: 'clam', 'alt', 'default', 'classic'

# Configurar a fonte padrão
style.configure('TLabel', font=('Helvetica', 12))
style.configure('TButton', font=('Helvetica', 12))

# Criar frames
left_frame = ttk.Frame(root, width=int(window_width * 0.5), height=window_height, relief=tk.SUNKEN)
left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

right_frame = ttk.Frame(root, width=int(window_width * 0.5), height=window_height, relief=tk.RAISED)
right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

# Configurar o fundo brranco no fame direito
right_frame.configure(style="Right.TFrame")
style.configure("Right.TFrame", background="white")

# Adicionar imagem no frame esquerdo
image_path = "../assets/imgs/logo.png"
if os.path.exists(image_path):
    image = Image.open(image_path)

    # Definir tamanho máximo para a imagem
    max_width = int(window_width * 0.5)  # Ajuste aqui para adicionar padding
    max_height = window_height - 40  # Ajuste aqui para adicionar padding vertical

    # Redimensionar a imagem mantendo a proporção
    image.thumbnail((max_width, max_height), Image.LANCZOS)
    photo = ImageTk.PhotoImage(image)

    # Calcular margens para centralizar a imagem com padding
    x_offset = (int(window_width * 0.6) - image.width) // 2
    y_offset = (window_height - image.height) // 2

    # Criar label para a imagem e centralizá-la com padding
    label_image = ttk.Label(left_frame, image=photo)
    label_image.image = photo  # Manter uma referência da imagem
    label_image.place(x=x_offset, y=y_offset)
else:
    label_image = ttk.Label(left_frame, text="Imagem não encontrada", font=('Helvetica', 16))
    label_image.pack(fill=tk.BOTH, expand=True)

# Criar um frame principal dentro do frame direito
frame = ttk.Frame(right_frame, padding="20", style="Right.TFrame")
frame.pack(expand=True)

# Adicionar um título
titulo = ttk.Label(frame, text="Análise de Vendas", font=('Helvetica', 16, 'bold'), background="white")
titulo.pack(pady=10)

# Entrada para data de início
ttk.Label(frame, text="Data de Início (dd/mm/yyyy):", background="white").pack(pady=5)
entrada_data_inicio = ttk.Entry(frame, width= 30)
entrada_data_inicio.pack(pady=5)

# Entrada para data de fim
ttk.Label(frame, text="Data de Fim (dd/mm/yyyy):", background="white").pack(pady=5)
entrada_data_fim = ttk.Entry(frame, width= 30)
entrada_data_fim.pack(pady=5)

# Criar um botão para abrir o diálogo de seleção de arquivo
btn_selecionar = ttk.Button(frame, text="Adicionar Arquivo Excel", command=selecionar_arquivo)
btn_selecionar.pack(pady=20)

# Iniciar o loop da interface gráfica
root.mainloop()