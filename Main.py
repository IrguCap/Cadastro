import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
import sqlite3

from openpyxl import Workbook
from datetime import datetime

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("dark-blue")

id_atual = None


def exportar_para_excel():
    conn = sqlite3.connect('cadastros.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM pessoas')
    rows = cursor.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = 'Cadastros'
    ws.append(['ID', 'Nome', 'CPF', 'Data Nascimento', 'Cód. Empresa', 'Razão Social', 'CNPJ', 'Status', 'Admissão',
               'Função', 'Salário'])

    for row in rows:
        ws.append(row)

    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

    if file_path:
        wb.save(file_path)
        messagebox.showinfo('Sucesso', 'Dados exportados para Excel com sucesso!')
    else:
        messagebox.showwarning('Aviso', 'Exportação cancelada.')


def buscar_cadastros():
    nome = entry_busca.get()

    conn = sqlite3.connect('cadastros.db')
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM pessoas WHERE nome LIKE ?', (f'%{nome}%',))
    rows = cursor.fetchall()
    conn.close()

    for row in treeview.get_children():
        treeview.delete(row)

    for row in rows:
        treeview.insert('', ctk.END, values=row)


def selecionar_cadastro(event):
    try:
        global id_atual
        item = treeview.selection()[0]
        id_atual = treeview.item(item, 'values')[0]

        selected_item = treeview.selection()[0]
        selected_data = treeview.item(selected_item, 'values')

        entry_nome_edit.delete(0, ctk.END)
        entry_nome_edit.insert(ctk.END, selected_data[1])

        entry_cpf_edit.configure(validate='none')
        entry_cpf_edit.delete(0, ctk.END)
        entry_cpf_edit.insert(ctk.END, selected_data[2])
        entry_cpf_edit.configure(validate='key')

        entry_nascimento_edit.delete(0, ctk.END)
        entry_nascimento_edit.insert(ctk.END, selected_data[3])

        entry_codempresa_edit.delete(0, ctk.END)
        entry_codempresa_edit.insert(ctk.END, selected_data[4])

        buscar_empresa(None)

        # entry_razao_edit.delete(0, ctk.END)
        # entry_razao_edit.insert(ctk.END, selected_data[5])

        entry_status_edit.delete(0, ctk.END)
        entry_status_edit.insert(ctk.END, selected_data[6])

        entry_admissao_edit.delete(0, ctk.END)
        entry_admissao_edit.insert(ctk.END, selected_data[7])

        entry_funcao_edit.delete(0, ctk.END)
        entry_funcao_edit.insert(ctk.END, selected_data[8])

        entry_salario_edit.delete(0, ctk.END)
        entry_salario_edit.insert(ctk.END, selected_data[9])

    except IndexError:
        pass


def editar_cadastro():
    global id_atual
    try:
        if id_atual is None:
            messagebox.showwarning('Aviso', 'Nenhum registro selecionado para edição.')
            return

        novo_nome = entry_nome_edit.get()
        novo_cpf = entry_cpf_edit.get()
        novo_nascimento = entry_nascimento_edit.get()
        novo_codempresa = entry_codempresa_edit.get()
        novo_razao = entry_razao_edit.get()
        novo_status = entry_status_edit.get()
        novo_admissao = entry_admissao_edit.get()
        novo_funcao = entry_funcao_edit.get()
        novo_salario = entry_salario_edit.get()


        conn = sqlite3.connect('cadastros.db')
        cursor = conn.cursor()

        cursor.execute('SELECT * FROM pessoas WHERE id=?', (id_atual,))
        registro = cursor.fetchone()

        if registro:
            if (novo_nome != registro[1] or
                    novo_cpf != registro[2] or
                    novo_nascimento != registro[3] or
                    novo_codempresa != registro[4] or
                    novo_razao != registro[5] or
                    novo_status != registro[6] or
                    novo_admissao != registro[7] or
                    novo_funcao != registro[8] or
                    novo_salario != registro[9]):


                cursor.execute(
                    'UPDATE pessoas SET nome=?, cpf=?, nascimento=?, codempresa=?, razao=?, cnpj=?, status=?, admissao=?, funcao=?, salario=? WHERE id=?',
                    (novo_nome, novo_cpf, novo_nascimento, novo_codempresa, novo_razao, buscar_cnpj(novo_codempresa) if novo_codempresa != registro[4] else registro[6], novo_status, novo_admissao,
                     novo_funcao, novo_salario, id_atual))
                conn.commit()

                if cursor.rowcount > 0:
                    messagebox.showinfo('Sucesso', 'Cadastro atualizado com sucesso!')
                else:
                    messagebox.showwarning('Aviso', 'Nenhuma alteração foi feita.')
            else:
                messagebox.showwarning('Aviso', 'Nenhuma alteração foi feita.')

            buscar_cadastros()

            entry_nome_edit.delete(0, ctk.END)
            entry_cpf_edit.delete(0, ctk.END)
            entry_nascimento_edit.delete(0, ctk.END)
            entry_codempresa_edit.delete(0, ctk.END)
            entry_razao_edit.delete(0, ctk.END)
            entry_status_edit.delete(0, ctk.END)
            entry_admissao_edit.delete(0, ctk.END)
            entry_funcao_edit.delete(0, ctk.END)
            entry_salario_edit.delete(0, ctk.END)

        else:
            messagebox.showwarning('Aviso', 'Registro não encontrado.')

        conn.close()

    except Exception as e:
        messagebox.showerror('Erro', f'Ocorreu um erro ao editar o cadastro: {str(e)}')


def excluir_cadastro():
    try:
        selected_item = treeview.selection()[0]
        selected_data = treeview.item(selected_item, 'values')

        nome = selected_data[1]
        cpf = selected_data[2]
        nascimento = selected_data[3]
        codempresa = selected_data[4]
        razao = selected_data[5]
        status = selected_data[6]
        admissao = selected_data[7]
        funcao = selected_data[8]
        salario = selected_data[9]

        resposta = messagebox.askyesno('Confirmação', 'Tem certeza de que deseja excluir o cadastro?')
        if resposta:
            conn = sqlite3.connect('cadastros.db')
            cursor = conn.cursor()

            cursor.execute(
                'DELETE FROM pessoas WHERE nome=? AND cpf=? AND nascimento AND codempresa=? AND razao=? AND status AND admissao=? AND funcao=? AND salario=?',
                (
                nome, cpf, nascimento, codempresa, razao, status, admissao, funcao, salario))
            conn.commit()
            conn.close()

            messagebox.showinfo('Sucesso', 'Cadastro excluído com sucesso!')

            buscar_cadastros()

            entry_nome_edit.delete(0, ctk.END)
            entry_cpf_edit.delete(0, ctk.END)
            entry_nascimento_edit.delete(0, ctk.END)
            entry_codempresa_edit.delete(0, ctk.END)
            entry_razao_edit.delete(0, ctk.END)
            entry_status_edit.delete(0, ctk.END)
            entry_admissao_edit.delete(0, ctk.END)
            entry_funcao_edit.delete(0, ctk.END)
            entry_salario_edit.delete(0, ctk.END)

        else:
            messagebox.showinfo('Informação', 'Exclusão abortada!')

    except IndexError:
        pass


def inserir_cadastro():
    nome = entry_nome.get()
    cpf = entry_cpf.get()
    nascimento = entry_nascimento.get()
    codempresa = entry_codempresa.get()
    razao = entry_razao.get()
    status = entry_status.get()
    admissao = entry_admissao.get()
    funcao = entry_funcao.get()
    salario = entry_salario.get()

    cnpj_empresa = buscar_cnpj(codempresa)

    conn = sqlite3.connect('cadastros.db')
    cursor = conn.cursor()


    cursor.execute('''
        CREATE TABLE IF NOT EXISTS pessoas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT,
            cpf TEXT,
            nascimento TEXT,
            codempresa TEXT,
            razao TEXT,
            cnpj TEXT,
            status TEXT,
            admissao TEXT,
            funcao TEXT,
            salario TEXT
        )
    ''')

    cursor.execute("PRAGMA table_info(pessoas)")
    colunas = [info[1] for info in cursor.fetchall()]

    # if 'cnpj_empresa' not in colunas:
    #     cursor.execute('ALTER TABLE pessoas ADD COLUMN cnpj TEXT')
    #     conn.commit()

    cursor.execute(
        'INSERT INTO pessoas (nome, cpf, nascimento, codempresa, razao, cnpj, status, admissao, funcao, salario) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
        (nome, cpf, nascimento, codempresa, razao, cnpj_empresa, status, admissao, funcao, salario))
    conn.commit()
    conn.close()


    messagebox.showinfo('Sucesso', 'Cadastro incluído com sucesso!')

    entry_nome.delete(0, 'end')
    entry_cpf.delete(0, 'end')
    entry_nascimento.delete(0, 'end')
    entry_codempresa.delete(0, 'end')
    entry_razao.delete(0, 'end')
    entry_status.delete(0, 'end')
    entry_admissao.delete(0, 'end')
    entry_funcao.delete(0, 'end')
    entry_salario.delete(0, 'end')

    entry_nome.focus_set()

def buscar_cnpj(codempresa):
    # Conectar ao banco de dados de empresas
    conn_empresas = sqlite3.connect('empresas.db')
    cursor_empresas = conn_empresas.cursor()

    # Consultar o CNPJ da empresa
    cursor_empresas.execute("SELECT cnpj FROM empresas WHERE codigo=?", (codempresa,))
    resultado = cursor_empresas.fetchone()

    conn_empresas.close()

    if resultado:
        return resultado[0]
    else:
        return None

def foco_proximo_entry(event, entry):
    entry.focus_set()


def converter_maiusculas(entry, *args):
    entry_var = entry.get()
    entry.set(entry_var.upper())


def validar_inteiro(char, entry):
    return char.isdigit() or char == ""


def validar_data(entry):
    data = entry.get()
    if not data:  # Se a entrada estiver vazia, retorna True
        return True
    try:
        datetime.strptime(data, '%d/%m/%Y')
        return True
    except ValueError:
        messagebox.showerror("Erro de Validação", "Por favor, insira uma data válida no formato DD/MM/AAAA.")
        entry.focus_set()
        return False


def formatar_data(entry, event=None):
    entrada = entry.get()
    nova_entrada = ''.join(filter(str.isdigit, entrada))

    if len(nova_entrada) > 8:
        nova_entrada = nova_entrada[:8]

    if len(nova_entrada) >= 5:
        nova_entrada = f"{nova_entrada[:2]}/{nova_entrada[2:4]}/{nova_entrada[4:8]}"
    elif len(nova_entrada) >= 3:
        nova_entrada = f"{nova_entrada[:2]}/{nova_entrada[2:4]}/{nova_entrada[4:]}"
    elif len(nova_entrada) >= 1:
        nova_entrada = f"{nova_entrada[:2]}/{nova_entrada[2:]}"

    entry.delete(0, ctk.END)
    entry.insert(0, nova_entrada)


def validar_float(event):
    entry = event.widget
    text = entry.get()

    if event.keysym == "BackSpace":
        return

    if event.char.isdigit():
        return

    # Permite ponto ou vírgula, mas apenas um de cada
    if event.char in ('.', ','):
        if '.' in text or ',' in text:
            return "break"  # Impede a inserção de mais de um ponto ou vírgula
        return

    # Impede outros caracteres
    return "break"


def validar_status(entry_var, name, index, mode):
    valor = entry_var.get().upper()
    if len(valor) > 1 or (valor not in ("", "A", "I")):
        entry_var.set(valor[0] if valor[0] in ("A", "I") else "")


def validar_cpf(cpf):
    cpf = ''.join(filter(str.isdigit, cpf))

    if len(cpf) != 11:
        return False

    # Calculando o primeiro dígito verificador
    soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
    digito1 = (soma * 10 % 11) % 10

    # Calculando o segundo dígito verificador
    soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
    digito2 = (soma * 10 % 11) % 10

    return cpf[-2:] == f'{digito1}{digito2}'


def validar_cpf_entry(char, entry_value):
    # Adicionando o caractere à string atual e mantendo apenas os dígitos
    entry_value = ''.join(filter(str.isdigit, entry_value + char))

    # Permitir apenas até 11 dígitos
    return len(entry_value) <= 12 and char.isdigit()


def validar_cnpj(cnpj):
    cnpj = ''.join(filter(str.isdigit, cnpj))

    if len(cnpj) != 14:
        return False

        # # Calculando o primeiro dígito verificador
        # soma = sum(int(cnpj[i]) * (13 - i) for i in range(12))
        # digito1 = (soma * 13 % 14) % 13
        #
        # # Calculando o segundo dígito verificador
        # soma = sum(int(cnpj[i]) * (14 - i) for i in range(13))
        # digito2 = (soma * 13 % 14) % 13
        #
        # return cnpj[-2:] == f'{digito1}{digito2}'


    # Pesos para o primeiro dígito verificador
    pesos_primeiro_digito = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    soma = sum(int(cnpj[i]) * pesos_primeiro_digito[i] for i in range(12))
    resto = soma % 11
    digito1 = 0 if resto < 2 else 11 - resto

    # Pesos para o segundo dígito verificador
    pesos_segundo_digito = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    soma = sum(int(cnpj[i]) * pesos_segundo_digito[i] for i in range(13))
    resto = soma % 11
    digito2 = 0 if resto < 2 else 11 - resto

    # Verifica se os dígitos calculados são iguais aos últimos dois dígitos do CNPJ
    return cnpj[-2:] == f'{digito1}{digito2}'

def validar_cnpj_entry(char, entry_value):
    # Adicionando o caractere à string atual e mantendo apenas os dígitos
    entry_value = ''.join(filter(str.isdigit, entry_value + char))

    # Permitir apenas até 14 dígitos
    return len(entry_value) <= 15 and char.isdigit()

def on_focus_out_cnpj(event):
    cnpj = event.widget.get()
    if cnpj and not validar_cnpj(cnpj):
        messagebox.showwarning("CNPJ Inválido", "O CNPJ digitado é inválido.")
        event.widget.focus_set()

def on_focus_out(event):
    cpf = event.widget.get()
    if cpf and not validar_cpf(cpf):
        messagebox.showwarning("CPF Inválido", "O CPF digitado é inválido.")
        event.widget.focus_set()  # Mantém o foco na Entry incorreta


def inserir_empresa():
    codigo = entry_empcode.get()
    nome = entry_empnome.get()
    cnpj = entry_cnpj.get()

    if codigo and nome and cnpj:
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()

            cursor.execute('''
                            CREATE TABLE IF NOT EXISTS empresas (
                                codigo INTEGER PRIMARY KEY,
                                nome TEXT NOT NULL,
                                cnpj TEXT
                            )
                        ''')

            cursor.execute('INSERT INTO empresas (codigo, nome, cnpj) VALUES (?, ?, ?)', (codigo, nome, cnpj))

            conn.commit()
            conn.close()

            messagebox.showinfo('Sucesso', 'Empresa incluída com sucesso!')

            entry_empcode.delete(0, 'end')
            entry_empnome.delete(0, 'end')
            entry_cnpj.delete(0, 'end')
            entry_empcode.focus_set()

        except sqlite3.Error as e:
            messagebox.showerror('Erro', f'Erro ao inserir empresa: {e}')
    else:
        messagebox.showwarning('Atenção', 'Preencha todos os campos!')

def buscar_empresa_include(event=None):
    cod_empresa = entry_codempresa.get()

    entry_razao_var.set('')

    if cod_empresa:
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()

            cursor.execute("SELECT nome FROM empresas WHERE codigo = ?", (cod_empresa,))
            resultado = cursor.fetchone()

            if resultado:
                nome_empresa = resultado[0]
                entry_razao_var.set(nome_empresa)
            else:
                entry_razao_var.set('Nome não encontrado')

            conn.close()
        except sqlite3.Error as e:
            print(f"Erro ao conectar ao banco de dados: {e}")

        entry_codempresa.bind("<KeyRelease>", buscar_empresa_include)

def buscar_empresa(event):
    codigo_empresa = entry_codempresa_edit.get()
    if codigo_empresa:
        try:
            conn = sqlite3.connect('empresas.db')
            cursor = conn.cursor()
            cursor.execute('SELECT nome FROM empresas WHERE codigo=?', (codigo_empresa,))
            resultado = cursor.fetchone()
            conn.close()
            if resultado:
                entry_razao_edit_var.set(resultado[0])
            else:
                entry_razao_edit_var.set('')
        except sqlite3.Error as e:
            messagebox.showerror('Erro', f'Erro ao buscar empresa: {e}')

def toggle_dark_mode():
    global dark_mode
    dark_mode = not dark_mode

    ctk.set_appearance_mode("dark" if dark_mode else "light")

    if dark_mode:
        style.configure("TNotebook.Tab", background="#666666", foreground="#000000",)
        treestyle.configure("Treeview", background="#333333", foreground="#FFFFFF", rowheight=20,
                            fieldbackground="#333333")
        treestyle.map('Treeview', background=[('selected', '#1a1a1a')], foreground=[('selected', 'white')])
    else:
        style.configure("TNotebook.Tab", background="#cbcbcb", foreground="#000000")
        treestyle.configure("Treeview", background="#white", foreground="black", rowheight=20, fieldbackground="white")
        treestyle.map('Treeview', background=[('selected', '#3470EB')], foreground=[('selected', 'white')])


# Criando a interface do programa
janela = ctk.CTk()
janela.title('Cadastro de Funcionários')
janela.geometry("1350x650")
janela.minsize(1350, 650)
janela.maxsize(1350, 650)

frame_borda = ctk.CTkFrame(janela, border_color="black", border_width=2)
frame_borda.pack(padx=10, pady=10, fill="both", expand=True)

notebook = ttk.Notebook(janela)
notebook.pack(expand=True, fill="both")
notebook.configure(style="TNotebook")

aba1 = ctk.CTkFrame(notebook)  # Aba do cadastro de funcionários (já existente)
aba2 = ctk.CTkFrame(notebook)
aba3 = ctk.CTkFrame(notebook)

style = ttk.Style()

notebook.add(aba1, text="Cadastro")
notebook.add(aba2, text="Edição")
notebook.add(aba3, text="Cadastro de Empresas")

button = ctk.CTkButton(frame_borda, text="", fg_color="#788d8d", font=("arial", 9, "bold"),
                       hover_color="#2F4F4F", border_width=2, corner_radius=15, command=toggle_dark_mode)
button.place(relx=0.92, rely=0.04, relwidth=0.06, relheight=0.08)

dark_mode = False

# Frames para organização
frame_borda1 = ctk.CTkFrame(master=aba1, border_color="black", border_width=2)
frame_borda1.pack(padx=10, pady=10, fill="both", expand=True)

frame_cadastro = ctk.CTkFrame(frame_borda1, border_color="black", border_width=1)
frame_cadastro.place(relx=0.15, rely=0.05, relwidth=0.7, relheight=0.8)

frame_busca = ctk.CTkFrame(frame_borda)
frame_busca.place(relx=0.05, rely=0.2, relwidth=0.9, relheight=0.12)

frame_resultados = ctk.CTkFrame(frame_borda)
frame_resultados.place(relx=0.05, rely=0.32, relwidth=0.9, relheight=0.5)

frame_borda2 = ctk.CTkFrame(master=aba2, border_color="black", border_width=2)
frame_borda2.pack(padx=10, pady=10, fill="both", expand=True)

frame_edicao = ctk.CTkFrame(frame_borda2, border_color="black", border_width=1)
frame_edicao.place(relx=0.15, rely=0.05, relwidth=0.7, relheight=0.8)

frame_borda3 = ctk.CTkFrame(master=aba3, border_color="black", border_width=2)
frame_borda3.pack(padx=10, pady=10, fill="both", expand=True)

frame_CadEmp = ctk.CTkFrame(frame_borda3, border_color="black", border_width=1)
frame_CadEmp.place(relx=0.25, rely=0.05, relwidth=0.5, relheight=0.5)


# Criação dos widgets de entrada
label_titulo1 = ctk.CTkLabel(frame_cadastro, text='INCLUSÃO DE FUNCIONÁRIOS', font=("arial", 14, "bold", "underline"))
label_titulo1.place(relx=0.37, rely=0.02)

label_nome = ctk.CTkLabel(frame_cadastro, text='Nome:', justify='right', font=("calibri", 11, "bold"))
label_nome.place(relx=0.15, rely=0.2)
entry_nome_var = ctk.StringVar()
entry_nome_var.trace_add("write", lambda name, index, mode, sv=entry_nome_var: converter_maiusculas(entry_nome_var))
entry_nome = ctk.CTkEntry(frame_cadastro, font=("calibri", 11), textvariable=entry_nome_var, width=240)
entry_nome.place(relx=0.19, rely=0.2)
entry_nome.bind('<Return>', lambda event: foco_proximo_entry(event, entry_cpf))

label_cpf = ctk.CTkLabel(frame_cadastro, text='CPF:', justify='right', font=("calibri", 11, "bold"))
label_cpf.place(relx=0.48, rely=0.2)
entry_cpf = ctk.CTkEntry(frame_cadastro, font=("calibri", 11), width=100, validate="key",
                     validatecommand=(janela.register(validar_cpf_entry), '%S', '%P'))
entry_cpf.place(relx=0.51, rely=0.2)
entry_cpf.bind('<Return>', lambda event: foco_proximo_entry(event, entry_nascimento))
entry_cpf.bind('<FocusOut>', on_focus_out)  # Verifica o CPF ao perder o foco

label_nascimento = ctk.CTkLabel(frame_cadastro, text='Data Nascimento:', font=("calibri", 11, "bold"))
label_nascimento.place(relx=0.64, rely=0.2)
entry_nascimento = ctk.CTkEntry(frame_cadastro, font=("calibri", 11), width=75)
entry_nascimento.place(relx=0.735, rely=0.2)
entry_nascimento.bind("<KeyRelease>", lambda event: formatar_data(entry_nascimento, event))
entry_nascimento.bind('<FocusOut>', lambda event: validar_data(entry_nascimento))
entry_nascimento.bind('<Return>', lambda event: foco_proximo_entry(event, entry_codempresa))

label_codempresa = ctk.CTkLabel(frame_cadastro, text='Cód. Empresa:', justify='right', font=("calibri", 11, "bold"))
label_codempresa.place(relx=0.14, rely=0.375)
entry_codempresa = ctk.CTkEntry(frame_cadastro, font=("calibri", 11), width=90, validate="key",
                            validatecommand=(janela.register(validar_inteiro), '%S', '%P'))
entry_codempresa.place(relx=0.22, rely=0.375)
entry_codempresa.bind("<KeyRelease>", buscar_empresa_include)
entry_codempresa.bind('<Return>', lambda event: foco_proximo_entry(event, entry_status))

label_razao = ctk.CTkLabel(frame_cadastro, text='Razão Social:', justify='right', font=("calibri", 11, "bold"))
label_razao.place(relx=0.34, rely=0.375)
entry_razao_var = ctk.StringVar()
entry_razao_var.trace("w", lambda name, index, mode, sv=entry_razao_var: converter_maiusculas(entry_razao_var))
entry_razao = ctk.CTkEntry(frame_cadastro, font=("calibri", 11), textvariable=entry_razao_var, state='readonly', width=270)
entry_razao.place(relx=0.41, rely=0.375)

label_status = ctk.CTkLabel(frame_cadastro, text='Status (Ativo/Inativo):', justify='right', font=("calibri", 11, "bold"))
label_status.place(relx=0.72, rely=0.375)
entry_status_var = ctk.StringVar()
entry_status_var.trace_add("write", lambda name, index, mode, sv=entry_status_var: (
converter_maiusculas(entry_status_var), validar_status(entry_status_var, name, index, mode)))
entry_status = ctk.CTkEntry(frame_cadastro, font=("calibri", 11), textvariable=entry_status_var, width=6)
entry_status.place(relx=0.835, rely=0.375)
entry_status.bind('<Return>', lambda event: foco_proximo_entry(event, entry_admissao))

label_admissao = ctk.CTkLabel(frame_cadastro, text='Admissão:', font=("calibri", 11, "bold"))
label_admissao.place(relx=0.24, rely=0.55)
entry_admissao = ctk.CTkEntry(frame_cadastro, font=("calibri", 11), width=75)
entry_admissao.place(relx=0.3, rely=0.55)
entry_admissao.bind("<KeyRelease>", lambda event: formatar_data(entry_admissao, event))
entry_admissao.bind('<FocusOut>', lambda event: validar_data(entry_admissao))
entry_admissao.bind('<Return>', lambda event: foco_proximo_entry(event, entry_funcao))

label_funcao = ctk.CTkLabel(frame_cadastro, text='Função:', font=("calibri", 11, "bold"))
label_funcao.place(relx=0.4, rely=0.55)
entry_funcao_var = ctk.StringVar()
entry_funcao_var.trace_add("write", lambda name, index, mode, sv=entry_funcao_var: converter_maiusculas(entry_funcao_var))
entry_funcao = ctk.CTkEntry(frame_cadastro, font=("calibri", 11), textvariable=entry_funcao_var, width=150)
entry_funcao.place(relx=0.45, rely=0.55)
entry_funcao.bind('<Return>', lambda event: foco_proximo_entry(event, entry_salario))

label_salario = ctk.CTkLabel(frame_cadastro, text='Salário:', font=("calibri", 11, "bold"))
label_salario.place(relx=0.635, rely=0.55)
entry_salario = ctk.CTkEntry(frame_cadastro, font=("calibri", 11), width=70)
entry_salario.place(relx=0.68, rely=0.55)
entry_salario.bind("<KeyPress>", validar_float)
entry_salario.bind('<Return>', lambda event: foco_proximo_entry(event, botao_inserir))

# Botão para inserir cadastro
botao_inserir = ctk.CTkButton(frame_cadastro, text='Incluir', command=inserir_cadastro, fg_color="#5F9EA0",
                          font=("arial", 9, "bold"), corner_radius=10, hover_color="#2F4F4F")
botao_inserir.place(relx=0.43, rely=0.8, relwidth=0.1, relheight=0.14)

# Campo de busca
label_busca = ctk.CTkLabel(frame_busca, text='Buscar por Nome:')
label_busca.place(relx=0.265, rely=0.2)
entry_busca_var = ctk.StringVar()
entry_busca_var.trace_add("write", lambda name, index, mode, sv=entry_busca_var: converter_maiusculas(entry_busca_var))
entry_busca = ctk.CTkEntry(frame_busca, textvariable=entry_busca_var)
entry_busca.place(relx=0.355, rely=0.2, relwidth=0.3)
entry_busca.bind('<Return>', lambda event: foco_proximo_entry(event, botao_buscar))

# Botão para buscar cadastro
botao_buscar = ctk.CTkButton(frame_busca, text='Buscar', command=buscar_cadastros, fg_color="#6495ED",
                             font=("arial", 9, "bold"), corner_radius=8, hover_color="#4682B4")
botao_buscar.place(relx=0.67, rely=0.12, relwidth=0.13, relheight=0.6)


treestyle = ttk.Style()
treestyle.theme_use('classic')
ttk.Style().configure("Treeview",
    background="#white",
    foreground="black",
    rowheight=20,
    fieldbackground="white")

ttk.Style().map('Treeview',
    background=[('selected', '#3470EB')],
    foreground=[('selected', 'white')])

# Treeview para exibir os cadastros
treeview = ttk.Treeview(frame_resultados, columns=(
'ID', 'Nome', 'CPF', 'Nascimento', 'CodEmpresa', 'Razao', 'CNPJ', 'Status', 'Admissao',
'Funcao', 'Salario'), show='headings')
treeview.pack(fill="both", expand=True)
treeview.heading('ID', text='ID')
treeview.heading('Nome', text='Nome')
treeview.heading('CPF', text='CPF')
treeview.heading('Nascimento', text='Nascimento')
treeview.heading('CodEmpresa', text='Cód. Empresa')
treeview.heading('Razao', text='Razão Social')
treeview.heading('CNPJ', text='')
treeview.heading('Status', text='Status')
treeview.heading('Admissao', text='Admissão')
treeview.heading('Funcao', text='Função')
treeview.heading('Salario', text='Salário')
treeview.place(relx=0.02, rely=0.02, relwidth=0.95, relheight=0.9)
treeview.bind('<<TreeviewSelect>>', selecionar_cadastro)

treeview.column("#0", width=0, stretch=0)  # Coluna de índice
treeview.column("ID", width=30, anchor="center")
treeview.column("Nome", width=240, anchor="center")
treeview.column("CPF", width=80, anchor="center")
treeview.column("Nascimento", width=65, anchor="center")
treeview.column("CodEmpresa", width=75, anchor="center")
treeview.column("Razao", width=200, anchor="center")
treeview.column("CNPJ", width=0, stretch=0,)
treeview.column("Status", width=35, anchor="center")
treeview.column("Admissao", width=65, anchor="center")
treeview.column("Funcao", width=175, anchor="center")
treeview.column("Salario", width=55, anchor="center")

# Botão para editar cadastro
botao_editar = ctk.CTkButton(frame_edicao, text='Alterar', command=editar_cadastro, fg_color="#DAA520",
                             font=("arial", 10, "bold"), corner_radius=10, hover_color="#c4a036")
botao_editar.place(relx=0.35, rely=0.78, relwidth=0.1, relheight=0.14)

# Botão para excluir cadastro
botao_excluir = ctk.CTkButton(frame_edicao, text='Excluir', command=excluir_cadastro, fg_color="#800000",
                              font=("arial", 10, "bold"), corner_radius=10, hover_color="#72202d")
botao_excluir.place(relx=0.58, rely=0.78, relwidth=0.1, relheight=0.14)

# Botão para exportar cadastros para Excel
botao_exportar = ctk.CTkButton(frame_borda, text='Exportar para Excel', command=exportar_para_excel, border_width=0, fg_color="#006400",
                           font=("arial", 12, "bold"), corner_radius=10, hover_color="#556B2F")
botao_exportar.place(relx=0.43, rely=0.84, relwidth=0.16, relheight=0.12)

# Campos para editar cadastro
label_titulo2 = ctk.CTkLabel(frame_edicao, text='ALTERAÇÕES DE CADASTRO', font=("arial", 14, "bold", "underline"))
label_titulo2.place(relx=0.37, rely=0.02)

label_nome_edit = ctk.CTkLabel(frame_edicao, text='Nome:', font=("calibri", 11, "bold"))
label_nome_edit.place(relx=0.15, rely=0.2)
entry_nome_edit_var = ctk.StringVar()
entry_nome_edit_var.trace_add("write", lambda name, index, mode, sv=entry_nome_edit_var: converter_maiusculas(
    entry_nome_edit_var))
entry_nome_edit = ctk.CTkEntry(frame_edicao, font=("calibri", 11), textvariable=entry_nome_edit_var, width=240)
entry_nome_edit.place(relx=0.19, rely=0.2)
entry_nome_edit.bind('<Return>', lambda event: foco_proximo_entry(event, entry_cpf_edit))

label_cpf_edit = ctk.CTkLabel(frame_edicao, text='CPF:', font=("calibri", 11, "bold"))
label_cpf_edit.place(relx=0.48, rely=0.2)
entry_cpf_edit = ctk.CTkEntry(frame_edicao, font=("calibri", 11), width=100, validate="key", validatecommand=(janela.register(validar_cpf_entry), '%S', '%P'))
entry_cpf_edit.place(relx=0.51, rely=0.2)
entry_cpf_edit.bind('<Return>', lambda event: foco_proximo_entry(event, entry_nascimento_edit))
entry_cpf_edit.bind('<FocusOut>', on_focus_out)  # Verifica o CPF ao perder o foco

label_nascimento_edit = ctk.CTkLabel(frame_edicao, text='Data Nascimento:', font=("calibri", 11, "bold"))
label_nascimento_edit.place(relx=0.64, rely=0.2)
entry_nascimento_edit = ctk.CTkEntry(frame_edicao, font=("calibri", 11), width=75)
entry_nascimento_edit.place(relx=0.735, rely=0.2)
entry_nascimento_edit.bind("<KeyRelease>", lambda event: formatar_data(entry_nascimento_edit, event))
entry_nascimento_edit.bind('<FocusOut>', lambda event: validar_data(entry_nascimento_edit))
entry_nascimento_edit.bind('<Return>', lambda event: foco_proximo_entry(event, entry_codempresa_edit))

label_codempresa_edit = ctk.CTkLabel(frame_edicao, text='Cód. Empresa:', font=("calibri", 11, "bold"))
label_codempresa_edit.place(relx=0.14, rely=0.375)
entry_codempresa_edit = ctk.CTkEntry(frame_edicao, font=("calibri", 11), width=90, validate="key",
                                     validatecommand=(janela.register(validar_inteiro), '%S', '%P'))
entry_codempresa_edit.place(relx=0.22, rely=0.375)
entry_codempresa_edit.bind("<KeyRelease>", buscar_empresa)
entry_codempresa_edit.bind('<Return>', lambda event: foco_proximo_entry(event, entry_status_edit))

label_razao_edit = ctk.CTkLabel(frame_edicao, text='Razão Social:', font=("calibri", 11, "bold"))
label_razao_edit.place(relx=0.34, rely=0.375)
entry_razao_edit_var = ctk.StringVar()
entry_razao_edit_var.trace_add("write", lambda name, index, mode, sv=entry_razao_edit_var: converter_maiusculas(
    entry_razao_edit_var))
entry_razao_edit = ctk.CTkEntry(frame_edicao, font=("calibri", 11), textvariable=entry_razao_edit_var, state='readonly', width=270)
entry_razao_edit.place(relx=0.41, rely=0.375)

label_status_edit = ctk.CTkLabel(frame_edicao, text='Status (Ativo/Inativo):', font=("calibri", 11, "bold"))
label_status_edit.place(relx=0.72, rely=0.375)
entry_status_edit_var = ctk.StringVar()
entry_status_edit_var.trace_add("write", lambda name, index, mode, sv=entry_status_edit_var: (
converter_maiusculas(entry_status_edit_var), validar_status(entry_status_edit_var, name, index, mode)))
entry_status_edit = ctk.CTkEntry(frame_edicao, font=("calibri", 11), textvariable=entry_status_edit_var, width=6)
entry_status_edit.place(relx=0.835, rely=0.375)
entry_status_edit.bind('<Return>', lambda event: foco_proximo_entry(event, entry_admissao_edit))

label_admissao_edit = ctk.CTkLabel(frame_edicao, text='Admissão:', font=("calibri", 11, "bold"))
label_admissao_edit.place(relx=0.24, rely=0.55)
entry_admissao_edit = ctk.CTkEntry(frame_edicao, font=("calibri", 11), width=75)
entry_admissao_edit.place(relx=0.3, rely=0.55)
entry_admissao_edit.bind("<KeyRelease>", lambda event: formatar_data(entry_admissao_edit, event))
entry_admissao_edit.bind('<FocusOut>', lambda event: validar_data(entry_admissao_edit))
entry_admissao_edit.bind('<Return>', lambda event: foco_proximo_entry(event, entry_funcao_edit))

label_funcao_edit = ctk.CTkLabel(frame_edicao, text='Função:', font=("calibri", 11, "bold"))
label_funcao_edit.place(relx=0.4, rely=0.55)
entry_funcao_edit_var = ctk.StringVar()
entry_funcao_edit_var.trace_add("write", lambda name, index, mode, sv=entry_funcao_edit_var: converter_maiusculas(
    entry_funcao_edit_var))
entry_funcao_edit = ctk.CTkEntry(frame_edicao, font=("calibri", 11), textvariable=entry_funcao_edit_var, width=150)
entry_funcao_edit.place(relx=0.45, rely=0.55)
entry_funcao_edit.bind('<Return>', lambda event: foco_proximo_entry(event, entry_salario_edit))

label_salario_edit = ctk.CTkLabel(frame_edicao, text='Salário:', font=("calibri", 11, "bold"))
label_salario_edit.place(relx=0.635, rely=0.55)
entry_salario_edit = ctk.CTkEntry(frame_edicao, font=("calibri", 11), width=70)
entry_salario_edit.place(relx=0.68, rely=0.55)
entry_salario_edit.bind('<Return>', lambda event: foco_proximo_entry(event, botao_editar))
entry_salario_edit.bind("<KeyPress>", validar_float)


def chamar_insercao(event):
    inserir_cadastro()


def chamar_edicao(event):
    editar_cadastro()


def chamar_busca(event):
    buscar_cadastros()


def chamar_inserirempresa(event):
    inserir_empresa()

# Campos para inserir empresa
label_titulo4 = ctk.CTkLabel(frame_CadEmp, text='EMPRESAS', font=("arial", 14, "bold", "underline"))
label_titulo4.place(relx=0.37, rely=0.02)

label_empcode = ctk.CTkLabel(frame_CadEmp, text='Código:', font=("calibri", 11, "bold"))
label_empcode.place(relx=0.15, rely=0.3)
entry_empcode = ctk.CTkEntry(frame_CadEmp, font=("calibri", 11), width=40, validate="key", validatecommand=(janela.register(validar_inteiro), '%S', '%P'))
entry_empcode.place(relx=0.22, rely=0.3)
entry_empcode.bind('<Return>', lambda event: foco_proximo_entry(event, entry_empnome))

label_empnome = ctk.CTkLabel(frame_CadEmp, text='Nome da empresa:', font=("calibri", 11, "bold"))
label_empnome.place(relx=0.32, rely=0.3)
entry_empnome_var = ctk.StringVar()
entry_empnome_var.trace("w", lambda name, index, mode, sv=entry_empnome_var: converter_maiusculas(entry_empnome_var))
entry_empnome = ctk.CTkEntry(frame_CadEmp, font=("calibri", 11), width=270, textvariable=entry_empnome_var)
entry_empnome.place(relx=0.46, rely=0.3)
entry_empnome.bind('<Return>', lambda event: foco_proximo_entry(event, entry_cnpj))

label_cnpj = ctk.CTkLabel(frame_CadEmp, text='CNPJ:', justify='right', font=("calibri", 11, "bold"))
label_cnpj.place(relx=0.2, rely=0.6)
entry_cnpj = ctk.CTkEntry(frame_CadEmp, font=("calibri", 11), width=110, validate="key", validatecommand=(janela.register(validar_cnpj_entry), '%S', '%P'))
entry_cnpj.place(relx=0.27, rely=0.6)
entry_cnpj.bind('<Return>', lambda event: foco_proximo_entry(event, botao_insemp))
entry_cnpj.bind('<FocusOut>', on_focus_out_cnpj)  # Verifica o CNPJ ao perder o foco


botao_insemp = ctk.CTkButton(frame_CadEmp, text='Inserir Empresa', command=inserir_empresa, fg_color="#CD853F", font=("arial", 9, "bold"))
botao_insemp.place(relx=0.65, rely=0.6, relwidth=0.2, relheight=0.2)


# Vincular a tecla ENTER aos botões
botao_inserir.bind("<Return>", chamar_insercao)
botao_editar.bind("<Return>", chamar_edicao)
botao_buscar.bind("<Return>", chamar_busca)
botao_insemp.bind("<Return>", chamar_inserirempresa)

# Loop principal da interface gráfica
janela.mainloop()