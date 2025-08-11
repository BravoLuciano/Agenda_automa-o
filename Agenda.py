from threading import Timer
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import mysql.connector
from datetime import datetime
import pywhatkit
import os
from PIL import Image, ImageTk
import pandas as pd
from fpdf import FPDF
import sys
from tkcalendar import DateEntry, Calendar

# Conexão com o banco de dados
conexao = mysql.connector.connect(
    host='localhost',
    port=3306,
    user='root',
    password='2712',
    database='agenda_treinamento'
)

# Criar a tabela automaticamente se não existir
cursor = conexao.cursor()
criar_tabela = """
CREATE TABLE IF NOT EXISTS agenda_treinamento (
    id INT AUTO_INCREMENT PRIMARY KEY,
    data DATE,
    treinamento VARCHAR(255),
    nome VARCHAR(255),
    horario VARCHAR(20),
    status VARCHAR(50),
    observacoes TEXT
)
"""
cursor.execute(criar_tabela)
conexao.commit()
cursor.close()

# Função de mensagem personalizada


def caixa_info(titulo, mensagem, texto_botao="Entendi"):
    win = tk.Toplevel(root)
    win.title(titulo)
    win.geometry("")
    win.resizable(False, False)
    win.grab_set()
    frame = tk.Frame(win, padx=20, pady=20)
    frame.pack()
    tk.Label(frame, text=mensagem, wraplength=300, pady=10).pack()
    btn = tk.Button(frame, text=texto_botao, width=12,
                    command=win.destroy, bg='#0078D7', fg='white')
    btn.pack(pady=(10, 0))
    btn.focus_set()
    win.transient(root)
    win.wait_window()

# Funções de banco


def cadastrar_treinamento():
    data = data_var.get().strip()
    treinamento = treinamento_var.get().strip()
    nome = nome_var.get().strip()
    horario = horario_var.get().strip()
    status = status_var.get().strip()  # Agora pode ser vazio
    observacoes = obs_var.get().strip()
    # Validação básica
    if not data or not treinamento or not nome or not horario:
        messagebox.showerror('Erro', 'Preencha todos os campos obrigatórios!')
        return
    try:
        # Valida formato da data
        data_sql = datetime.strptime(data, '%d/%m/%Y').date()
    except ValueError:
        messagebox.showerror('Erro', 'Data deve estar no formato DD/MM/AAAA!')
        return
    cursor = conexao.cursor()
    comando = "INSERT INTO agenda_treinamento (data, treinamento, nome, horario, status, observacoes) VALUES (%s, %s, %s, %s, %s, %s)"
    valores = (data_sql, treinamento, nome, horario, status, observacoes)
    cursor.execute(comando, valores)
    conexao.commit()
    cursor.close()
    caixa_info('Sucesso', 'Treinamento cadastrado com sucesso!')
    data_var.set('')
    treinamento_var.set('')
    nome_var.set('')
    horario_var.set('')
    status_var.set('')
    obs_var.set('')
    listar_treinamentos()


def listar_treinamentos():
    for row in tree.get_children():
        tree.delete(row)
    cursor = conexao.cursor()
    cursor.execute(
        "SELECT id, data, treinamento, nome, horario, status, observacoes FROM agenda_treinamento")
    resultados = cursor.fetchall()
    for row in resultados:
        # Formatar data para DD/MM/AAAA
        row = list(row)
        row[1] = row[1].strftime('%d/%m/%Y') if row[1] else ''
        tree.insert('', 'end', values=row)
    cursor.close()


def buscar_treinamento():
    termo = simpledialog.askstring('Buscar', 'Digite o nome ou treinamento:')
    if termo:
        for row in tree.get_children():
            tree.delete(row)
        cursor = conexao.cursor()
        cursor.execute(
            "SELECT id, data, treinamento, nome, horario, status, observacoes FROM agenda_treinamento WHERE nome LIKE %s OR treinamento LIKE %s", (f'%{termo}%', f'%{termo}%'))
        resultados = cursor.fetchall()
        if resultados:
            for row in resultados:
                row = list(row)
                row[1] = row[1].strftime('%d/%m/%Y') if row[1] else ''
                tree.insert('', 'end', values=row)
        else:
            caixa_info('Busca', 'Nenhum treinamento encontrado.')
        cursor.close()


def editar_treinamento():
    selected = tree.focus()
    if not selected:
        messagebox.showwarning('Editar', 'Selecione um treinamento na lista!')
        return
    values = tree.item(selected, 'values')
    id_treinamento = values[0]

    # Nova janela de edição
    edit_win = tk.Toplevel(root)
    edit_win.title('Editar Treinamento')
    edit_win.geometry('600x300')
    edit_win.resizable(False, False)
    edit_win.grab_set()

    # Variáveis para edição
    data_edit = tk.StringVar(value=values[1])
    treinamento_edit = tk.StringVar(value=values[2])
    nome_edit = tk.StringVar(value=values[3])
    horario_edit = tk.StringVar(value=values[4])
    status_edit = tk.StringVar(value=values[5])
    obs_edit = tk.StringVar(value=values[6])

    # Layout dos campos
    tk.Label(edit_win, text='Data (DD/MM/AAAA):').grid(row=0,
                                                       column=0, sticky='e', padx=(15, 5), pady=10)
    entry_data_edit = tk.Entry(edit_win, textvariable=data_edit, width=16)
    entry_data_edit.grid(row=0, column=1, padx=(5, 15))

    def abrir_calendario_edicao():
        top = tk.Toplevel(edit_win)
        top.title('Selecionar Data')
        cal = Calendar(top, date_pattern='dd/MM/yyyy',
                       selectmode='day', locale='pt_BR')
        cal.pack(padx=10, pady=10)

        def selecionar():
            data_edit.set(cal.get_date())
            top.destroy()
        tk.Button(top, text='OK', command=selecionar,
                  bg='#0078D7', fg='white').pack(pady=5)
        top.grab_set()
    btn_calendario_edit = tk.Button(
        edit_win, text='Selecionar Data', command=abrir_calendario_edicao, width=14)
    btn_calendario_edit.grid(row=0, column=2, padx=(5, 15))
    entry_data_edit.bind('<<DateEntrySelected>>',
                         lambda event: edit_win.focus_set())
    tk.Label(edit_win, text='Treinamento:').grid(
        row=0, column=2, sticky='e', padx=5)
    tk.Entry(edit_win, textvariable=treinamento_edit,
             width=20).grid(row=0, column=3, padx=5)

    tk.Label(edit_win, text='Nome:').grid(
        row=1, column=0, sticky='e', padx=5, pady=10)
    tk.Entry(edit_win, textvariable=nome_edit,
             width=30).grid(row=1, column=3, padx=5)
    tk.Label(edit_win, text='Horário:').grid(
        row=1, column=2, sticky='e', padx=5)
    tk.Entry(edit_win, textvariable=horario_edit,
             width=15).grid(row=1, column=3, padx=5)

    tk.Label(edit_win, text='Status:').grid(
        row=2, column=0, sticky='e', padx=5, pady=10)

    tk.Entry(edit_win, textvariable=status_edit,
             width=20).grid(row=2, column=1, padx=5)
    tk.Label(edit_win, text='Observações:').grid(
        row=2, column=2, sticky='e', padx=5)
    tk.Entry(edit_win, textvariable=obs_edit,
             width=30).grid(row=2, column=3, padx=5)

    def salvar_edicao():
        nova_data = data_edit.get().strip()
        novo_treinamento = treinamento_edit.get().strip()
        novo_nome = nome_edit.get().strip()
        novo_horario = horario_edit.get().strip()
        novo_status = status_edit.get().strip()  # Agora pode ser vazio
        novas_obs = obs_edit.get().strip()
        if not nova_data or not novo_treinamento or not novo_nome or not novo_horario:
            messagebox.showerror(
                'Erro', 'Preencha todos os campos obrigatórios!')
            return
        try:
            data_sql = datetime.strptime(nova_data, '%d/%m/%Y').date()
        except ValueError:
            messagebox.showerror(
                'Erro', 'Data deve estar no formato DD/MM/AAAA!')
            return
        cursor = conexao.cursor()
        comando = "UPDATE agenda_treinamento SET data = %s, treinamento = %s, nome = %s, horario = %s, status = %s, observacoes = %s WHERE id = %s"
        valores = (data_sql, novo_treinamento, novo_nome,
                   novo_horario, novo_status, novas_obs, id_treinamento)
        cursor.execute(comando, valores)
        conexao.commit()
        cursor.close()
        caixa_info('Sucesso', 'Treinamento atualizado!')
        listar_treinamentos()
        edit_win.destroy()

    btn_salvar = tk.Button(edit_win, text='Salvar',
                           command=salvar_edicao, bg='#0078D7', fg='white', width=14)
    btn_salvar.grid(row=3, column=0, columnspan=4, pady=20)


def remover_treinamento():
    selected = tree.focus()
    if not selected:
        messagebox.showwarning('Remover', 'Selecione um treinamento na lista!')
        return
    values = tree.item(selected, 'values')
    id_treinamento = values[0]
    confirm = messagebox.askyesno(
        'Remover', f'Tem certeza que deseja remover o treinamento de ID {id_treinamento}?')
    if confirm:
        cursor = conexao.cursor()
        cursor.execute(
            "DELETE FROM agenda_treinamento WHERE id = %s", (id_treinamento,))
        conexao.commit()
        cursor.close()
        caixa_info('Remover', 'Treinamento removido com sucesso!')
        listar_treinamentos()

# Remover funções de exportação para Excel e PDF


# DASHBOARD


def abrir_dashboard():
    # from calendar import month_name  # Não usar mais
    import matplotlib
    matplotlib.use('TkAgg')
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    import matplotlib.pyplot as plt
    dash = tk.Toplevel(root)
    dash.title('Dashboard de Treinamentos')
    dash.geometry('650x650')
    dash.resizable(False, False)
    dash.grab_set()

    # Filtros de mês e ano (em português)
    frame_filtro = tk.Frame(dash)
    frame_filtro.pack(pady=10)
    tk.Label(frame_filtro, text='Mês:').pack(side='left')
    mes_var = tk.StringVar()
    ano_var = tk.StringVar()
    meses_pt = [
        "01 - Janeiro", "02 - Fevereiro", "03 - Março", "04 - Abril", "05 - Maio", "06 - Junho",
        "07 - Julho", "08 - Agosto", "09 - Setembro", "10 - Outubro", "11 - Novembro", "12 - Dezembro"
    ]
    mes_var.set(meses_pt[datetime.now().month-1])
    mes_menu = ttk.Combobox(frame_filtro, textvariable=mes_var,
                            values=meses_pt, width=15, state='readonly')
    mes_menu.pack(side='left', padx=5)
    tk.Label(frame_filtro, text='Ano:').pack(side='left', padx=(10, 0))
    ano_atual = datetime.now().year
    anos = [str(ano) for ano in range(ano_atual-5, ano_atual+2)]
    ano_var.set(str(ano_atual))
    ano_menu = ttk.Combobox(
        frame_filtro, textvariable=ano_var, values=anos, width=7, state='readonly')
    ano_menu.pack(side='left', padx=5)

    frame_result = tk.Frame(dash)
    frame_result.pack(pady=10, fill='x')

    frame_grafico = tk.Frame(dash)
    frame_grafico.pack(pady=10)

    def atualizar_dash():
        for widget in frame_result.winfo_children():
            widget.destroy()
        for widget in frame_grafico.winfo_children():
            widget.destroy()
        mes_num = int(mes_var.get().split(' - ')[0])
        ano_num = int(ano_var.get())
        # Datas para o filtro
        data_ini = datetime(ano_num, mes_num, 1).date()
        if mes_num == 12:
            data_fim = datetime(ano_num+1, 1, 1).date()
        else:
            data_fim = datetime(ano_num, mes_num+1, 1).date()
        cursor = conexao.cursor()
        # Total de treinamentos
        cursor.execute(
            'SELECT COUNT(*) FROM agenda_treinamento WHERE data >= %s AND data < %s', (data_ini, data_fim))
        total = cursor.fetchone()[0]
        # Realizados (status preenchido)
        cursor.execute(
            "SELECT COUNT(*) FROM agenda_treinamento WHERE (status IS NOT NULL AND status != '') AND data >= %s AND data < %s", (data_ini, data_fim))
        realizados = cursor.fetchone()[0]
        # Pendentes (status vazio)
        cursor.execute(
            "SELECT COUNT(*) FROM agenda_treinamento WHERE (status IS NULL OR status = '') AND data >= %s AND data < %s", (data_ini, data_fim))
        pendentes = cursor.fetchone()[0]
        # Próximos treinamentos (data >= hoje, dentro do mês)
        hoje = datetime.now().date()
        cursor.execute("SELECT data, treinamento, nome, horario FROM agenda_treinamento WHERE data >= %s AND data >= %s AND data < %s ORDER BY data ASC LIMIT 5", (hoje, data_ini, data_fim))
        proximos = cursor.fetchall()
        cursor.close()

        tk.Label(frame_result, text=f'Total de treinamentos cadastrados: {total}', font=(
            'Arial', 12, 'bold')).pack(pady=10)
        tk.Label(frame_result, text=f'Treinamentos realizados: {realizados}', font=(
            'Arial', 12)).pack(pady=5)
        tk.Label(frame_result, text=f'Treinamentos pendentes: {pendentes}', font=(
            'Arial', 12)).pack(pady=5)
        tk.Label(frame_result, text='Próximos treinamentos:',
                 font=('Arial', 12, 'bold')).pack(pady=15)

        frame_prox = tk.Frame(frame_result)
        frame_prox.pack()
        cols = ('Data', 'Treinamento', 'Nome', 'Horário')
        tree_dash = ttk.Treeview(
            frame_prox, columns=cols, show='headings', height=5)
        for col in cols:
            tree_dash.heading(col, text=col)
            tree_dash.column(col, anchor='center', width=100)
        tree_dash.pack()
        for row in proximos:
            row = list(row)
            row[0] = row[0].strftime('%d/%m/%Y') if row[0] else ''
            tree_dash.insert('', 'end', values=row)

        # Gráfico de barras
        fig, ax = plt.subplots(figsize=(5, 2.5))
        categorias = ['Total', 'Realizados', 'Pendentes']
        valores = [total, realizados, pendentes]
        cores = ['#0078D7', '#28a745', '#ffc107']
        ax.bar(categorias, valores, color=cores)
        ax.set_ylabel('Quantidade')
        ax.set_title(f'Treinamentos em {mes_var.get()} de {ano_var.get()}')
        for i, v in enumerate(valores):
            ax.text(i, v + 0.1, str(v), ha='center',
                    va='bottom', fontweight='bold')
        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, master=frame_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack()
        plt.close(fig)

    atualizar_dash()

    btn_filtrar = tk.Button(frame_filtro, text='Filtrar',
                            command=atualizar_dash, bg='#0078D7', fg='white', width=10)
    btn_filtrar.pack(side='left', padx=10)

    tk.Button(dash, text='Fechar', command=dash.destroy,
              width=12, bg='#0078D7', fg='white').pack(pady=20)


def enviar_whatsapp_treinamentos():
    from datetime import datetime
    hoje = datetime.now().date()
    cursor = conexao.cursor()
    cursor.execute(
        "SELECT treinamento, nome, horario FROM agenda_treinamento WHERE data = %s", (hoje,))
    treinamentos = cursor.fetchall()
    cursor.close()
    if treinamentos:
        mensagem = "Segue os treinamentos do dia e seus horarios:\n"
        for t in treinamentos:
            mensagem += f"- {t[0]} (Responsável: {t[1]}): {t[2]}\n"
    else:
        mensagem = "Não há treinamentos agendados para hoje."
    numero = "+5514998864173"
    pywhatkit.sendwhatmsg_instantly(numero, mensagem)


# Envia WhatsApp ao abrir a agenda
enviar_whatsapp_treinamentos()

# GUI
root = tk.Tk()
root.title('Agenda de Treinamentos')
root.geometry('900x500')
root.resizable(False, False)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# Exibir o logo no topo
try:
    top_logo_frame = tk.Frame(root)
    top_logo_frame.pack(fill='x')
    # Redimensiona a imagem para uma largura menor para evitar desfoque
    img = Image.open(resource_path('logo.png'))
    img = img.resize((700, 90))  # 700px de largura, 90px de altura
    img_tk = ImageTk.PhotoImage(img)
    logo_label = tk.Label(top_logo_frame, image=img_tk)
    logo_label.image = img_tk  # Manter referência
    logo_label.pack(fill='x')
except Exception as e:
    print(f'Erro ao carregar logo: {e}')

# Frame de cadastro
frame_cad = tk.LabelFrame(root, text='Cadastrar Treinamento', padx=10, pady=10)
frame_cad.pack(fill='x', padx=10, pady=5)

data_var = tk.StringVar()
treinamento_var = tk.StringVar()
nome_var = tk.StringVar()
horario_var = tk.StringVar()
status_var = tk.StringVar()
obs_var = tk.StringVar()

# --- Cadastro ---
lbl_data = tk.Label(frame_cad, text='Data:')
lbl_data.grid(row=0, column=0, sticky='e', padx=(10, 3))
entry_data = tk.Entry(frame_cad, textvariable=data_var, width=12)
entry_data.grid(row=0, column=1, padx=3)


def abrir_calendario_cadastro():
    top = tk.Toplevel(root)
    top.title('Selecionar Data')
    cal = Calendar(top, date_pattern='dd/MM/yyyy',
                   selectmode='day', locale='pt_BR')
    cal.pack(padx=10, pady=10)

    def selecionar():
        data_var.set(cal.get_date())
        top.destroy()
    tk.Button(top, text='OK', command=selecionar,
              bg='#0078D7', fg='white').pack(pady=5)
    top.grab_set()


btn_calendario = tk.Button(
    frame_cad, text='Selecionar Data', command=abrir_calendario_cadastro, width=12)
btn_calendario.grid(row=0, column=2, padx=3)

lbl_treinamento = tk.Label(frame_cad, text='Treinamento:')
lbl_treinamento.grid(row=0, column=3, sticky='e', padx=(10, 3))
entry_treinamento = tk.Entry(frame_cad, textvariable=treinamento_var, width=20)
entry_treinamento.grid(row=0, column=4, padx=3)

lbl_horario = tk.Label(frame_cad, text='Horário:')
lbl_horario.grid(row=0, column=5, sticky='e', padx=(10, 3))
entry_horario = tk.Entry(frame_cad, textvariable=horario_var, width=10)
entry_horario.grid(row=0, column=6, padx=(3, 10))

lbl_nome = tk.Label(frame_cad, text='Nome:')
lbl_nome.grid(row=1, column=0, sticky='e', padx=(10, 3))
entry_nome = tk.Entry(frame_cad, textvariable=nome_var, width=25)
entry_nome.grid(row=1, column=1, padx=3)

lbl_status = tk.Label(frame_cad, text='Status:')
lbl_status.grid(row=1, column=2, sticky='e', padx=(10, 3))
entry_status = tk.Entry(frame_cad, textvariable=status_var, width=15)
entry_status.grid(row=1, column=3, padx=3)

lbl_obs = tk.Label(frame_cad, text='Observações:')
lbl_obs.grid(row=1, column=4, sticky='e', padx=(10, 3))
entry_obs = tk.Entry(frame_cad, textvariable=obs_var, width=25)
entry_obs.grid(row=1, column=5, padx=(3, 10))

# Botão Cadastrar em nova linha, centralizado
btn_cadastrar = tk.Button(frame_cad, text='Cadastrar',
                          command=cadastrar_treinamento, bg='#0078D7', fg='white', width=16)
btn_cadastrar.grid(row=2, column=0, columnspan=7, pady=(10, 0))

# Frame de ações
frame_acao = tk.Frame(root)
frame_acao.pack(fill='x', padx=10, pady=5)

btn_listar = tk.Button(frame_acao, text='Listar Todos',
                       command=listar_treinamentos, width=15)
btn_listar.pack(side='left', padx=5)
btn_buscar = tk.Button(frame_acao, text='Buscar por Nome/Treinamento',
                       command=buscar_treinamento, width=22)
btn_buscar.pack(side='left', padx=5)
btn_editar = tk.Button(frame_acao, text='Editar Selecionado',
                       command=editar_treinamento, width=18)
btn_editar.pack(side='left', padx=5)
btn_remover = tk.Button(frame_acao, text='Remover Selecionado',
                        command=remover_treinamento, width=18)
btn_remover.pack(side='left', padx=5)
btn_dash = tk.Button(frame_acao, text='Dashboard',
                     command=abrir_dashboard, width=15, bg='#28a745', fg='white')
btn_dash.pack(side='left', padx=5)
btn_sair = tk.Button(frame_acao, text='Sair', command=root.destroy, width=10)
btn_sair.pack(side='right', padx=5)

# Frame de filtro por data
frame_filtro_data = tk.Frame(root)
frame_filtro_data.pack(fill='x', padx=10, pady=(5, 0))

filtro_data_var = tk.StringVar()
entry_filtro_data = tk.Entry(
    frame_filtro_data, textvariable=filtro_data_var, width=16)
entry_filtro_data.pack(side='left', padx=5)


def abrir_calendario_filtro():
    top = tk.Toplevel(root)
    top.title('Selecionar Data para Filtro')
    cal = Calendar(top, date_pattern='dd/MM/yyyy',
                   selectmode='day', locale='pt_BR')
    cal.pack(padx=10, pady=10)

    def selecionar():
        filtro_data_var.set(cal.get_date())
        top.destroy()
    tk.Button(top, text='OK', command=selecionar,
              bg='#0078D7', fg='white').pack(pady=5)
    top.grab_set()


btn_calendario_filtro = tk.Button(
    frame_filtro_data, text='Selecionar Data', command=abrir_calendario_filtro, width=14)
btn_calendario_filtro.pack(side='left', padx=2)


def filtrar_por_data():
    data_filtro = filtro_data_var.get().strip()
    for row in tree.get_children():
        tree.delete(row)
    if not data_filtro:
        listar_treinamentos()
        return
    try:
        data_sql = datetime.strptime(data_filtro, '%d/%m/%Y').date()
    except ValueError:
        messagebox.showerror('Erro', 'Data deve estar no formato DD/MM/AAAA!')
        return
    cursor = conexao.cursor()
    cursor.execute(
        "SELECT id, data, treinamento, nome, horario, status, observacoes FROM agenda_treinamento WHERE data = %s", (data_sql,))
    resultados = cursor.fetchall()
    for row in resultados:
        row = list(row)
        row[1] = row[1].strftime('%d/%m/%Y') if row[1] else ''
        tree.insert('', 'end', values=row)
    cursor.close()


btn_filtrar_data = tk.Button(frame_filtro_data, text='Filtrar por Data',
                             command=filtrar_por_data, width=16, bg='#0078D7', fg='white')
btn_filtrar_data.pack(side='left', padx=5)

# Frame da lista
frame_lista = tk.LabelFrame(
    root, text='Treinamentos Cadastrados', padx=10, pady=10)
frame_lista.pack(fill='both', expand=True, padx=10, pady=5)

colunas = ('ID', 'Data', 'Treinamento', 'Nome',
           'Horário', 'Status', 'Observações')
tree = ttk.Treeview(frame_lista, columns=colunas, show='headings', height=12)
tree.heading('ID', text='ID')
tree.column('ID', anchor='center', width=40, minwidth=40)
tree.heading('Data', text='Data')
tree.column('Data', anchor='center', width=90, minwidth=80)
tree.heading('Treinamento', text='Treinamento')
tree.column('Treinamento', anchor='center', width=150, minwidth=100)
tree.heading('Nome', text='Nome')
tree.column('Nome', anchor='center', width=150, minwidth=100)
tree.heading('Horário', text='Horário')
tree.column('Horário', anchor='center', width=80, minwidth=60)
tree.heading('Status', text='Status')
tree.column('Status', anchor='center', width=100, minwidth=80)
tree.heading('Observações', text='Observações')
tree.column('Observações', anchor='center', width=180, minwidth=100)
tree.pack(fill='both', expand=True)

listar_treinamentos()

root.mainloop()
