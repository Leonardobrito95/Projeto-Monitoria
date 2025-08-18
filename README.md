import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import customtkinter as ctk
from tkcalendar import DateEntry
import pandas as pd
import sqlite3
import os
from datetime import datetime
import matplotlib
# Definir backend ANTES de importar pyplot para integração com Tkinter
matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from PIL import Image

# --- CONFIGURAÇÕES DA APLICAÇÃO ---
DB_FILE = 'monitoria.db'
EXCEL_FILE = 'Métricas de Atendimento.xlsx'
# Diretório e arquivos de imagem (logo e ícone)
ASSETS_DIR = 'assets'
APP_LOGO_FILE = os.path.join(ASSETS_DIR, 'logo_canaa.png')
APP_ICON_FILE = os.path.join(ASSETS_DIR, 'icon_canaa.png')
# NOTA DE SEGURANÇA: A senha está codificada. Para um ambiente de produção,
# considere usar um método mais seguro para armazená-la, como variáveis de ambiente ou um arquivo de configuração seguro.
ADMIN_PASSWORD = os.getenv("MONITORIA_ADMIN_PASSWORD", "admin123")  # Senha administrativa padrão
COLUNAS = [
    'Motivo do Atendimento',
    'Monitoria Zero',
    'Protocolo', 'Data M', 'Nome do Agente', 'Equipe', 'Script inicial/final', 'Sondagem',
    'Conhecimento técnico', 'Vícios de linguagem', 'Tom de voz', 'Cordialidade',
    'Controle de Objeção', 'Ofensa Verbal', 'Retorno ao cliente', 'Ação de retenção',
    'Confirmação de dados', 'Transferencia Indevida', 'Uso do Mute', 'Erro de procedimento',
    'Negociação e venda', 'Inf. Protocolo?', 'Agilidade', 'Prontidão', 'Tabulação',
    'Resolução do conflito', 'Personalização', 'Omissão de atendimento', 'Avaliação ATD.',
    'Erro Crítico?', 'Itens Aplicáveis', 'Pontuação', 'Observações'
]
BOOLEAN_FIELDS = ['Erro Crítico?', 'Inf. Protocolo?']
YES_NO_FIELDS = [
    'Script inicial/final', 'Sondagem', 'Conhecimento técnico', 'Vícios de linguagem',
    'Tom de voz', 'Cordialidade', 'Controle de Objeção', 'Ofensa Verbal',
    'Retorno ao cliente', 'Ação de retenção', 'Confirmação de dados',
    'Transferencia Indevida', 'Uso do Mute', 'Erro de procedimento',
    'Negociação e venda', 'Agilidade', 'Prontidão', 'Tabulação',
    'Resolução do conflito', 'Personalização', 'Omissão de atendimento'
]
NUMERIC_FIELDS = ['Avaliação ATD.', 'Itens Aplicáveis', 'Pontuação']
CRITICAL_ERRORS = {
    'Omissão de atendimento': 'Não Conforme',
    'Ofensa Verbal': 'Não Conforme',
    'Erro de procedimento': 'Não Conforme',
    'Confirmação de dados': 'Não Conforme',
    'Inf. Protocolo?': 'Não Conforme'
}

# Penalizações para cada critério (quando "Não Conforme")
PENALIZACOES = {
    'Script inicial/final': 0.50,
    'Sondagem': 0.50,
    'Conhecimento técnico': 0.50,
    'Vícios de linguagem': 0.50,
    'Transferencia Indevida': 0.50,
    'Ofensa Verbal' : 1.00,
    'Controle de Objeção': 0.50,
    'Retorno ao cliente': 0.50,
    'Ação de retenção': 0.50,
    'Confirmação de dados': 1.00,
    'Tom de voz': 0.50,
    'Uso do Mute': 0.50,
    'Erro de procedimento': 1.00,
    'Negociação e venda': 0.50,
    'Inf. Protocolo?': 1.00,
    'Agilidade': 0.50,
    'Prontidão': 0.50,
    'Tabulação': 0.50,
    'Resolução do conflito': 0.50,
    'Cordialidade': 0.50,
    'Personalização': 0.50,
    'Omissão de atendimento': 0.00
}

# Agentes e suas equipes
AGENTES_EQUIPE = {
    'Sarah Couto': 'SAC',
    'Matheus Henrique': 'SAC',
    'Larissa Santos': 'SAC',
    'Manoel Junior': 'SAC',
    'Andressa Costa': 'SAC',
    'Priscilla Rodrigues': 'SAC',
    'Matheus Ferreira': 'SAC',
    'Hyrum Castro': 'SAC',
    'Rafael Vieira': 'SAC',
    'Aline Dias': 'SAC',
    'Livia Reis': 'SAC',
    'Walison Rodrigues': 'SAC',
    'Caique Abreu': 'SAC',
    'Rafael Brito': 'N2',
    'Ubiratan Sobrinho': 'N2',
    'Karoliny Lira': 'Retenção',
    'Anna Santos': 'SAC',
    'Clara Vieira':'SAC',
    'Camila':'SAC'
}

# Estado para modo de edição
edit_mode = False
edit_id = None
botao_salvar = None  # Referência global para o botão de salvar monitoria
edit_agente_mode = False  # Estado para edição de agente
agente_em_edicao = None  # Agente sendo editado
canvas_bar = None  # Canvas para gráfico de barras
canvas_pie = None  # Canvas para gráfico de pizza
app_logo_img = None  # Referência global para a logo exibida no topo
current_theme = 'dark'  # Tema atual: 'dark' ou 'light'
chat_window = None  # Janela flutuante de chat
chat_fab = None  # Botão flutuante para abrir chat

def get_theme_colors(mode: str) -> dict:
    """Retorna paleta de cores básica por tema."""
    if mode == 'light':
        return {
            'text': '#0F172A',
            'placeholder_text': '#64748B',
            'input_bg': '#EEF2F6',
            'app_bg': '#E9EDF3',
            'tab_bg': '#a6aaad',
            'tree_bg': '#FFFFFF',
            'tree_fg': '#0F172A',
            'tree_sel': '#4A90E2',
            'tree_head_bg': '#E1E8F0',
            'tree_head_fg': '#0F172A',
            'fig_bg': '#FFFFFF',
            'fig_text': '#0F172A',
            'bar_color': '#4A90E2',
            'pie_ok': '#4A90E2',
            'pie_crit': '#FF4C4C',
            'list_bg': '#FFFFFF',
            'list_fg': '#0F172A',
            'seg_unselected_bg': '#E1E8F0',
            'seg_selected_bg': '#4A90E2'
        }
    # default dark
    return {
        'text': '#FFFFFF',
        'input_bg': '#2E5A88',
        'app_bg': '#1C2526',
        'tab_bg': '#1C2526',
        'tree_bg': '#2a2d2e',
        'tree_fg': '#FFFFFF',
        'tree_sel': '#4A90E2',
        'tree_head_bg': '#2E5A88',
        'tree_head_fg': '#FFFFFF',
        'fig_bg': '#2a2d2e',
        'fig_text': '#FFFFFF',
        'bar_color': '#4A90E2',
        'pie_ok': '#4A90E2',
        'pie_crit': '#FF4C4C',
        'list_bg': '#2a2d2e',
        'list_fg': '#FFFFFF',
        'seg_unselected_bg': '#4A90E2',
        'seg_selected_bg': '#2E5A88'
    }

def configure_ttk_style_for_theme(mode: str):
    colors = get_theme_colors(mode)
    style = ttk.Style()
    # força o tema 'default' e aplica cores
    try:
        style.theme_use('default')
    except Exception:
        pass
    style.configure(
        'Treeview',
        background=colors['tree_bg'],
        foreground=colors['tree_fg'],
        fieldbackground=colors['tree_bg'],
        borderwidth=0,
        rowheight=25
    )
    style.map('Treeview', background=[('selected', colors['tree_sel'])])
    style.configure(
        'Treeview.Heading',
        background=colors['tree_head_bg'],
        foreground=colors['tree_head_fg'],
        font=('Arial', 10, 'bold'),
        relief='flat'
    )
    style.map('Treeview.Heading', background=[('active', '#3B6EA8')])

def _iter_widgets(widget):
    yield widget
    for child in widget.winfo_children():
        yield from _iter_widgets(child)

def apply_theme_colors(mode: str):
    """Aplica cores básicas nos widgets conhecidos."""
    colors = get_theme_colors(mode)
    # Fundo da aplicação e Tabview
    try:
        app.configure(fg_color=colors['app_bg'])
        tabview.configure(
            fg_color=colors['tab_bg'],
            text_color=colors['text'],
            segmented_button_fg_color=colors['seg_unselected_bg'],
            segmented_button_selected_color=colors['seg_selected_bg']
        )
        # Abas/frames principais
        for f in [tab_form, tab_table, tab_dashboard, tab_agentes]:
            try:
                f.configure(fg_color=colors['tab_bg'])
            except Exception:
                pass
    except Exception:
        pass
    # Campos do formulário
    for col, w in widgets.items():
        if not w:
            continue
        try:
            if isinstance(w, ctk.CTkEntry):
                w.configure(fg_color=colors['input_bg'], text_color=colors['text'],
                           placeholder_text_color=colors.get('placeholder_text', colors['text']))
            elif isinstance(w, ctk.CTkComboBox):
                w.configure(fg_color=colors['input_bg'], text_color=colors['text'],
                           button_color=colors['seg_selected_bg'], button_hover_color='#3B6EA8')
            elif isinstance(w, ctk.CTkTextbox):
                w.configure(fg_color=colors['input_bg'], text_color=colors['text'])
            elif isinstance(w, DateEntry):
                if mode == 'light':
                    w.configure(background='#E1E8F0', foreground=colors['text'])
                else:
                    w.configure(background='#4A90E2', foreground='#FFFFFF')
            if col == 'Erro Crítico?':
                w.configure(text_color=colors['text'])
        except Exception:
            pass
    # Rótulos (inclui títulos). Mantém vermelho para os críticos.
    try:
        critical_keys = set(CRITICAL_ERRORS.keys())
        for wid in _iter_widgets(app):
            if isinstance(wid, ctk.CTkLabel):
                try:
                    txt = wid.cget('text') or ''
                except Exception:
                    txt = ''
                if txt not in critical_keys:
                    try:
                        wid.configure(text_color=colors['text'])
                    except Exception:
                        pass
    except Exception:
        pass
    # Filtros da tabela
    try:
        combo_filtro_agente.configure(fg_color=colors['input_bg'], text_color=colors['text'],
                                     button_color=colors['seg_selected_bg'], button_hover_color='#3B6EA8')
        entry_filtro_protocolo.configure(fg_color=colors['input_bg'], text_color=colors['text'])
    except Exception:
        pass
    # Filtros do dashboard
    try:
        combo_filtro_agente_dashboard.configure(fg_color=colors['input_bg'], text_color=colors['text'],
                                               button_color=colors['seg_selected_bg'], button_hover_color='#3B6EA8')
        combo_filtro_equipe_dashboard.configure(fg_color=colors['input_bg'], text_color=colors['text'],
                                               button_color=colors['seg_selected_bg'], button_hover_color='#3B6EA8')
        entry_filtro_avaliacao_dashboard.configure(fg_color=colors['input_bg'], text_color=colors['text'])
        entry_filtro_pontuacao_dashboard.configure(fg_color=colors['input_bg'], text_color=colors['text'])
    except Exception:
        pass
    # Listbox de agentes
    try:
        listbox_agentes.configure(bg=colors['list_bg'], fg=colors['list_fg'], selectbackground='#4A90E2', selectforeground=colors['tree_fg'])
    except Exception:
        pass

def set_theme(mode: str):
    """Altera o tema claro/escuro da aplicação e atualiza estilos e gráficos."""
    global current_theme
    if mode not in ('dark', 'light'):
        return
    current_theme = mode
    try:
        ctk.set_appearance_mode('dark' if mode == 'dark' else 'light')
    except Exception:
        pass
    configure_ttk_style_for_theme(mode)
    apply_theme_colors(mode)
    # Recarrega tabelas e dashboard para refletir cores e estilos
    try:
        atualizar_ultimos_lancamentos()
    except Exception:
        pass
    try:
        aplicar_filtros_dashboard()
    except Exception:
        try:
            atualizar_dashboard()
        except Exception:
            pass

# --- FUNÇÕES DO BANCO DE DADOS ---
def init_db():
    """Inicializa o banco de dados SQLite."""
    with sqlite3.connect(DB_FILE) as conn:
        cursor = conn.cursor()
        
        # Cria a tabela com um ID se ela não existir
        cursor.execute(f'''
            CREATE TABLE IF NOT EXISTS monitoria (
                id INTEGER PRIMARY KEY AUTOINCREMENT
            )
        ''')
        
        # Obtém a lista de colunas existentes
        cursor.execute(f"PRAGMA table_info(monitoria)")
        existing_columns = [info[1] for info in cursor.fetchall()]

        # Adiciona cada coluna da lista COLUNAS se ela não existir
        for col in COLUNAS:
            if col not in existing_columns:
                try:
                    cursor.execute(f'ALTER TABLE monitoria ADD COLUMN "{col}" TEXT')
                except sqlite3.OperationalError:
                    # Pode ocorrer se a coluna já existir por algum motivo, ignoramos
                    pass
        conn.commit()

        # Cria a tabela de agentes para persistir agentes/equipes
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS agentes (
                nome TEXT PRIMARY KEY,
                equipe TEXT NOT NULL
            )
        ''')

        # Seed inicial de agentes se a tabela estiver vazia
        cursor.execute('SELECT COUNT(*) FROM agentes')
        total_agentes = cursor.fetchone()[0]
        if total_agentes == 0:
            for nome, equipe in AGENTES_EQUIPE.items():
                cursor.execute('INSERT OR IGNORE INTO agentes (nome, equipe) VALUES (?, ?)', (nome, equipe))
        conn.commit()

def carregar_dados_iniciais():
    """Carrega agentes e equipes da tabela 'agentes'. Fallback para dicionário."""
    try:
        with sqlite3.connect(DB_FILE) as conn:
            df = pd.read_sql_query('SELECT nome, equipe FROM agentes', conn)
        if not df.empty:
            agentes = sorted(df['nome'].astype(str).tolist())
            equipes = sorted(df['equipe'].astype(str).unique().tolist())
            return equipes, agentes
    except Exception:
        pass
    agentes = sorted(AGENTES_EQUIPE.keys())
    equipes = sorted(set(AGENTES_EQUIPE.values()))
    return equipes, agentes

def verificar_protocolo_duplicado(protocolo, exclude_id=None):
    """Verifica se o protocolo já existe no banco, exceto para o ID em edição."""
    with sqlite3.connect(DB_FILE) as conn:
        cursor = conn.cursor()
        if exclude_id:
            cursor.execute("SELECT COUNT(*) FROM monitoria WHERE Protocolo = ? AND id != ?", (protocolo, exclude_id))
        else:
            cursor.execute("SELECT COUNT(*) FROM monitoria WHERE Protocolo = ?", (protocolo,))
        count = cursor.fetchone()[0]
    return count > 0

def calcular_pontuacao(dados):
    """Calcula a pontuação e itens aplicáveis com base nos dados do formulário."""
    pontuacao = 10.0  # Nota inicial
    itens_aplicaveis = 0
    erro_critico = False

    # Verificar erros críticos
    for campo, valor_critico in CRITICAL_ERRORS.items():
        if dados.get(campo) == valor_critico:
            erro_critico = True
            pontuacao = 0.0
            break

    # Calcular pontuação apenas se não houver erro crítico
    if not erro_critico:
        for campo in YES_NO_FIELDS:
            if dados.get(campo) == 'Não Conforme':
                pontuacao -= PENALIZACOES.get(campo, 0)
            if dados.get(campo) in ['Conforme', 'Não Conforme']:
                itens_aplicaveis += 1

    return max(0, pontuacao), itens_aplicaveis, 'Sim' if erro_critico else 'Não'

def salvar_monitoria():
    """Salva ou atualiza uma monitoria no banco de dados."""
    global edit_mode, edit_id
    dados = {}
    
    # CORREÇÃO: Removida a captura redundante de 'Motivo do Atendimento'.
    # A captura agora é feita de forma unificada no loop abaixo.
    for col in COLUNAS:
        if col in widgets and widgets[col]:
            if col in BOOLEAN_FIELDS or col in YES_NO_FIELDS or col in ['Nome do Agente', 'Equipe', 'Motivo do Atendimento']:
                dados[col] = widgets[col].get()
            elif col in ['Protocolo', 'Avaliação ATD.']:
                dados[col] = widgets[col].get()
            elif col == 'Data M':
                dados[col] = widgets[col].get_date().strftime('%d/%m/%Y')
            elif col == 'Observações':
                dados[col] = widgets[col].get("1.0", tk.END).strip()
            elif col == 'Monitoria Zero':
                # Salva em branco se for "Nenhum"
                dados[col] = widgets[col].get() if widgets[col].get() != 'Nenhum' else ''

    # Validação
    required_fields = ['Protocolo', 'Nome do Agente', 'Equipe']
    if not all(dados.get(field) for field in required_fields):
        messagebox.showwarning("Campos Obrigatórios", "Preencha Protocolo, Nome do Agente e Equipe.")
        return

    if dados.get('Avaliação ATD.'):
        try:
            valor_avaliacao = float(dados['Avaliação ATD.'])
        except (ValueError, TypeError):
            messagebox.showwarning("Entrada Inválida", "O campo Avaliação ATD. deve ser numérico.")
            return
        if not (0.0 <= valor_avaliacao <= 10.0):
            messagebox.showwarning("Entrada Inválida", "O campo Avaliação ATD. deve estar entre 0 e 10.")
            return

    # Verificar duplicidade de protocolo
    if verificar_protocolo_duplicado(dados['Protocolo'], exclude_id=edit_id if edit_mode else None):
        messagebox.showwarning("Protocolo Duplicado", "Este número de protocolo já está registrado.")
        return

    # Calcular pontuação, itens aplicáveis e erro crítico
    pontuacao, itens_aplicaveis, erro_critico = calcular_pontuacao(dados)
    dados['Pontuação'] = f"{pontuacao:.2f}"
    dados['Itens Aplicáveis'] = str(itens_aplicaveis)
    dados['Erro Crítico?'] = erro_critico

    # Preencher campos restantes que podem não estar no formulário
    for col in COLUNAS:
        if col not in dados:
            dados[col] = ''

    try:
        with sqlite3.connect(DB_FILE) as conn:
            cursor = conn.cursor()
            if edit_mode:
                # Atualizar registro existente
                columns = ', '.join([f'"{col}" = ?' for col in COLUNAS])
                values = [dados.get(col, '') for col in COLUNAS] + [edit_id]
                cursor.execute(f'UPDATE monitoria SET {columns} WHERE id = ?', values)
            else:
                # Inserir novo registro
                columns = ', '.join([f'"{col}"' for col in COLUNAS])
                placeholders = ', '.join(['?' for _ in COLUNAS])
                values = [dados.get(col, '') for col in COLUNAS]
                cursor.execute(f'INSERT INTO monitoria ({columns}) VALUES ({placeholders})', values)
            conn.commit()

        update_excel()
        messagebox.showinfo("Sucesso", "Monitoria salva com sucesso!" if not edit_mode else "Monitoria atualizada com sucesso!")
        limpar_formulario()
        aplicar_filtros()
        aplicar_filtros_dashboard()
        if edit_mode:
            edit_mode = False
            edit_id = None
            botao_salvar.configure(text="Salvar Monitoria")
            tabview.set("Nova Monitoria")
    except Exception as e:
        messagebox.showerror("Erro ao Salvar", f"Erro ao salvar dados: {e}")

def update_excel():
    """Atualiza a aba 'Base de dados da Monitoria' no arquivo Excel."""
    try:
        from openpyxl import load_workbook
        from openpyxl.utils.dataframe import dataframe_to_rows
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

        with sqlite3.connect(DB_FILE) as conn:
            df = pd.read_sql_query("SELECT * FROM monitoria", conn)
        if 'id' in df.columns:
            df = df.drop(columns=['id'])

        df['Data M'] = pd.to_datetime(df['Data M'], format='%d/%m/%Y', errors='coerce').dt.date
        for col in ['Avaliação ATD.', 'Itens Aplicáveis', 'Pontuação']:
            df[col] = pd.to_numeric(df[col], errors='coerce').round(2).fillna(0)

        if os.path.exists(EXCEL_FILE):
            wb = load_workbook(EXCEL_FILE)
        else:
            from openpyxl import Workbook
            wb = Workbook()
            if 'Sheet' in wb.sheetnames: # Remove a folha padrão
                wb.remove(wb['Sheet'])

        if 'Base de dados da Monitoria' in wb.sheetnames:
            wb.remove(wb['Base de dados da Monitoria'])

        ws = wb.create_sheet('Base de dados da Monitoria')

        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                if r_idx == 1:
                    cell.fill = PatternFill(start_color='4A90E2', end_color='4A90E2', fill_type='solid')
                    cell.font = Font(bold=True, color='FFFFFF')
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        for col_cells in ws.columns:
            max_length = max(len(str(cell.value or "")) for cell in col_cells)
            ws.column_dimensions[col_cells[0].column_letter].width = max(max_length + 2, 12)

        wb.save(EXCEL_FILE)

    except Exception as e:
        messagebox.showerror("Erro ao Atualizar Excel", f"Erro ao atualizar Excel: {e}")

def excluir_registro():
    """Exclui o registro selecionado após confirmação."""
    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("Nenhum Registro Selecionado", "Selecione um registro para excluir.")
        return
    
    selected_item = selected_items[0]
    values = tree.item(selected_item, 'values')
    protocolo_selecionado = values[COLUNAS.index('Protocolo')]
    try:
        registro_id = int(selected_item)
    except ValueError:
        registro_id = None

    if not messagebox.askyesno("Confirmar Exclusão", f"Deseja realmente excluir a monitoria com protocolo {protocolo_selecionado}?"):
        return

    try:
        with sqlite3.connect(DB_FILE) as conn:
            cursor = conn.cursor()
            if registro_id is not None:
                cursor.execute("DELETE FROM monitoria WHERE id = ?", (registro_id,))
            else:
            cursor.execute("DELETE FROM monitoria WHERE Protocolo = ?", (protocolo_selecionado,))
            conn.commit()

        update_excel()
        aplicar_filtros()
        aplicar_filtros_dashboard()
        messagebox.showinfo("Sucesso", f"Monitoria com protocolo {protocolo_selecionado} excluída com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro ao Excluir", f"Erro ao excluir registro: {e}")

def limpar_formulario():
    """Limpa todos os campos do formulário."""
    global edit_mode, edit_id
    for col, widget in widgets.items():
        if not widget:
            continue
        if col in YES_NO_FIELDS:
            widget.set('Conforme')
            if col in CRITICAL_ERRORS:
                widget.configure(text_color="#FFFFFF")
        elif col == 'Erro Crítico?':
            widget.set('Não')
            widget.configure(text_color="#FFFFFF")
        elif col == 'Data M':
            widget.set_date(datetime.now())
        elif col in ['Protocolo', 'Avaliação ATD.']:
            widget.delete(0, 'end')
        elif col == 'Observações':
            widget.delete("1.0", tk.END)
        # CORREÇÃO: Acessava 'dados', que não existe neste escopo.
        # Agora, define corretamente o valor padrão dos ComboBoxes.
        elif col == 'Motivo do Atendimento':
            widget.set('')
        elif col == 'Monitoria Zero':
            widget.set('Nenhum')
        elif col in ['Nome do Agente', 'Equipe']:
            widget.set('')

    widgets['Protocolo'].focus()
    if edit_mode:
        edit_mode = False
        edit_id = None
        botao_salvar.configure(text="Salvar Monitoria")

def atualizar_ultimos_lancamentos(filtro_agente=None, filtro_protocolo=None):
    """Atualiza a tabela com registros do banco de dados, aplicando filtros."""
    for i in tree.get_children():
        tree.delete(i)
    try:
        with sqlite3.connect(DB_FILE) as conn:
            query = "SELECT * FROM monitoria"
            conditions, params = [], []

            if filtro_agente and filtro_agente != "Todos":
                conditions.append('"Nome do Agente" = ?')
                params.append(filtro_agente)
            if filtro_protocolo:
                conditions.append('Protocolo LIKE ?')
                params.append(f'%{filtro_protocolo}%')

            if conditions:
                query += " WHERE " + " AND ".join(conditions)
            
            query += " ORDER BY id DESC"
            df = pd.read_sql_query(query, conn, params=params)
            
            for col in COLUNAS:
                if col not in df.columns:
                    df[col] = ''
            df = df.fillna('')

        tree['columns'] = COLUNAS
        for col in COLUNAS:
            tree.heading(col, text=col)
            tree.column(col, width=120 if col != 'Observações' else 200, anchor='center', stretch=tk.NO)

        for _, row in df.iterrows():
            # Usa o id como iid para operações seguras
            iid = None
            if 'id' in df.columns:
                try:
                    iid_val = row['id']
                    if pd.notna(iid_val):
                        iid = str(int(iid_val))
                except Exception:
                    iid = None
            valores = [row[col] if col in row else '' for col in COLUNAS]
            tree.insert("", "end", iid=iid, values=valores)
    except Exception as e:
        messagebox.showerror("Erro de Leitura", f"Erro ao carregar lançamentos: {e}")

def aplicar_filtros():
    """Aplica filtros de agente e protocolo à tabela."""
    agente = combo_filtro_agente.get()
    protocolo = entry_filtro_protocolo.get().strip()
    atualizar_ultimos_lancamentos(filtro_agente=agente if agente != "Todos" else None, filtro_protocolo=protocolo)

def limpar_filtros():
    """Limpa os campos de filtro e recarrega todos os registros."""
    combo_filtro_agente.set("Todos")
    entry_filtro_protocolo.delete(0, tk.END)
    atualizar_ultimos_lancamentos()

def atualizar_equipe(*args):
    """Atualiza o campo Equipe com base no agente selecionado."""
    agente = widgets['Nome do Agente'].get()
    equipe = AGENTES_EQUIPE.get(agente, '')
    try:
        with sqlite3.connect(DB_FILE) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT equipe FROM agentes WHERE nome = ?', (agente,))
            row = cursor.fetchone()
            if row:
                equipe = row[0]
    except Exception:
        pass
    widgets['Equipe'].set(equipe)

def atualizar_cor_critica(widget, campo):
    """Atualiza a cor do texto do ComboBox com base na seleção crítica."""
    valor = widget.get()
    if campo in CRITICAL_ERRORS and valor == CRITICAL_ERRORS[campo]:
        widget.configure(text_color="#FF0000")
    else:
        widget.configure(text_color="#FFFFFF")

def verificar_senha_admin():
    """Solicita e verifica a senha administrativa."""
    dialog = ctk.CTkInputDialog(title="Autenticação Administrativa", text="Digite a senha administrativa:")
    senha = dialog.get_input()
    return senha == ADMIN_PASSWORD if senha is not None else False

# --- AUXILIARES DE DATA (filtros) ---
def _parse_date_str(date_str):
    """Converte dd/mm/YYYY em objeto date. Retorna None se vazio ou inválido."""
    if not date_str:
        return None
    try:
        return datetime.strptime(date_str, '%d/%m/%Y').date()
    except Exception:
        return None

def _to_ymd(date_obj):
    """Retorna YYYYMMDD para comparações lexicográficas no SQLite."""
    if not date_obj:
        return None
    return date_obj.strftime('%Y%m%d')

# CORREÇÃO: Função auxiliar para centralizar a atualização dos comboboxes
def _atualizar_comboboxes_agentes():
    """Atualiza todos os comboboxes de agentes e equipes na UI."""
    equipes, agentes = carregar_dados_iniciais()
    widgets['Nome do Agente'].configure(values=agentes)
    combo_filtro_agente.configure(values=["Todos"] + agentes)
    combo_filtro_agente_dashboard.configure(values=["Todos"] + agentes)
    combo_filtro_equipe_dashboard.configure(values=["Todas"] + equipes)
    combo_equipe_novo_agente.configure(values=equipes)

def adicionar_agente():
    """Adiciona ou edita um agente no dicionário e atualiza a interface."""
    global edit_agente_mode, agente_em_edicao
    if not verificar_senha_admin():
        messagebox.showerror("Erro de Autenticação", "Senha administrativa incorreta.")
        return

    nome_agente = entry_novo_agente.get().strip()
    equipe = combo_equipe_novo_agente.get()

    if not nome_agente or not equipe:
        messagebox.showwarning("Campos Vazios", "Por favor, insira o nome do agente e selecione uma equipe.")
        return
    try:
        with sqlite3.connect(DB_FILE) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT COUNT(*) FROM agentes WHERE nome = ?', (nome_agente,))
            existe = cursor.fetchone()[0] > 0
    except Exception:
        existe = nome_agente in AGENTES_EQUIPE
    if existe and (not edit_agente_mode or nome_agente != agente_em_edicao):
        messagebox.showwarning("Agente Duplicado", f"O agente {nome_agente} já existe.")
        return

    message = ""
    if edit_agente_mode:
        try:
            with sqlite3.connect(DB_FILE) as conn:
                cursor = conn.cursor()
        if agente_em_edicao != nome_agente:
                    cursor.execute('DELETE FROM agentes WHERE nome = ?', (agente_em_edicao,))
                    cursor.execute('INSERT OR REPLACE INTO agentes (nome, equipe) VALUES (?, ?)', (nome_agente, equipe))
                else:
                    cursor.execute('UPDATE agentes SET equipe = ? WHERE nome = ?', (equipe, nome_agente))
                conn.commit()
        except Exception:
            AGENTES_EQUIPE.pop(agente_em_edicao, None)
        AGENTES_EQUIPE[nome_agente] = equipe
        message = f"Agente {nome_agente} atualizado com sucesso!"
    else:
        try:
            with sqlite3.connect(DB_FILE) as conn:
                cursor = conn.cursor()
                cursor.execute('INSERT OR IGNORE INTO agentes (nome, equipe) VALUES (?, ?)', (nome_agente, equipe))
                conn.commit()
        except Exception:
        AGENTES_EQUIPE[nome_agente] = equipe
        message = f"Agente {nome_agente} adicionado com sucesso!"

    _atualizar_comboboxes_agentes() # Centraliza a atualização

    listbox_agentes.delete(0, tk.END)
    try:
        with sqlite3.connect(DB_FILE) as conn:
            df = pd.read_sql_query('SELECT nome, equipe FROM agentes ORDER BY nome COLLATE NOCASE', conn)
        for _, r in df.iterrows():
            listbox_agentes.insert(tk.END, f"{r['nome']} ({r['equipe']})")
    except Exception:
    for ag in sorted(AGENTES_EQUIPE.keys()):
        listbox_agentes.insert(tk.END, f"{ag} ({AGENTES_EQUIPE[ag]})")

    entry_novo_agente.delete(0, tk.END)
    combo_equipe_novo_agente.set('')
    if edit_agente_mode:
        edit_agente_mode = False
        agente_em_edicao = None
        botao_adicionar_agente.configure(text="Adicionar Agente")

    messagebox.showinfo("Sucesso", message)

def excluir_agente():
    """Exclui o agente selecionado do dicionário e atualiza a interface."""
    if not verificar_senha_admin():
        messagebox.showerror("Erro de Autenticação", "Senha administrativa incorreta.")
        return

    try:
        selected = listbox_agentes.get(listbox_agentes.curselection())
        agente = selected.split(' (')[0]
    except tk.TclError:
        messagebox.showwarning("Nenhum Agente Selecionado", "Selecione um agente para excluir.")
        return

    if not messagebox.askyesno("Confirmar Exclusão", f"Deseja realmente excluir o agente {agente}?"):
        return

    with sqlite3.connect(DB_FILE) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM monitoria WHERE \"Nome do Agente\" = ?", (agente,))
        count = cursor.fetchone()[0]
        if count > 0:
            messagebox.showwarning("Agente em Uso", f"O agente {agente} está vinculado a {count} monitoria(s) e não pode ser excluído.")
            return

    try:
        with sqlite3.connect(DB_FILE) as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM agentes WHERE nome = ?', (agente,))
            conn.commit()
    except Exception:
        AGENTES_EQUIPE.pop(agente, None)
    _atualizar_comboboxes_agentes() # Centraliza a atualização

    listbox_agentes.delete(0, tk.END)
    try:
        with sqlite3.connect(DB_FILE) as conn:
            df = pd.read_sql_query('SELECT nome, equipe FROM agentes ORDER BY nome COLLATE NOCASE', conn)
        for _, r in df.iterrows():
            listbox_agentes.insert(tk.END, f"{r['nome']} ({r['equipe']})")
    except Exception:
    for ag in sorted(AGENTES_EQUIPE.keys()):
        listbox_agentes.insert(tk.END, f"{ag} ({AGENTES_EQUIPE[ag]})")

    messagebox.showinfo("Sucesso", f"Agente {agente} excluído com sucesso!")

def editar_agente():
    """Carrega o agente selecionado para edição."""
    global edit_agente_mode, agente_em_edicao
    if not verificar_senha_admin():
        messagebox.showerror("Erro de Autenticação", "Senha administrativa incorreta.")
        return

    try:
        selected = listbox_agentes.get(listbox_agentes.curselection())
        agente = selected.split(' (')[0]
    except tk.TclError:
        messagebox.showwarning("Nenhum Agente Selecionado", "Selecione um agente para editar.")
        return

    entry_novo_agente.delete(0, tk.END)
    entry_novo_agente.insert(0, agente)
    try:
        with sqlite3.connect(DB_FILE) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT equipe FROM agentes WHERE nome = ?', (agente,))
            row = cursor.fetchone()
            combo_equipe_novo_agente.set(row[0] if row else '')
    except Exception:
        combo_equipe_novo_agente.set(AGENTES_EQUIPE.get(agente, ''))

    edit_agente_mode = True
    agente_em_edicao = agente
    botao_adicionar_agente.configure(text="Salvar Alterações")

def alterar_senha_admin():
    """Altera a senha administrativa após validação."""
    global ADMIN_PASSWORD
    if not verificar_senha_admin():
        messagebox.showerror("Erro", "Senha administrativa atual incorreta.")
        return

    dialog_nova = ctk.CTkInputDialog(title="Alterar Senha", text="Digite a nova senha administrativa:")
    nova_senha = dialog_nova.get_input()

    if nova_senha:
        ADMIN_PASSWORD = nova_senha
        messagebox.showinfo("Sucesso", "Senha administrativa alterada com sucesso!")
    else:
        messagebox.showwarning("Ação Cancelada", "A nova senha não pode estar vazia.")

def limpar_lancamentos():
    """Limpa todos os lançamentos do banco de dados após validação administrativa."""
    if not verificar_senha_admin():
        messagebox.showerror("Erro de Autenticação", "Senha administrativa incorreta.")
        return

    if not messagebox.askyesno("Confirmar Limpeza", "Deseja realmente excluir TODOS os lançamentos? Esta ação não pode ser desfeita."):
        return

    try:
        with sqlite3.connect(DB_FILE) as conn:
            conn.execute("DELETE FROM monitoria")
            conn.commit()

        update_excel()
        aplicar_filtros()
        aplicar_filtros_dashboard()
        messagebox.showinfo("Sucesso", "Todos os lançamentos foram excluídos.")
    except Exception as e:
        messagebox.showerror("Erro ao Limpar", f"Erro ao limpar lançamentos: {e}")

def atualizar_dashboard(filtro_agente=None, filtro_equipe=None, filtro_avaliacao=None, filtro_pontuacao=None, data_ini=None, data_fim=None):
    """Atualiza a aba Dashboard com métricas."""
    global canvas_bar, canvas_pie
    
    for i in dashboard_tree.get_children():
        dashboard_tree.delete(i)
    
    default_dashboard_columns = ['Agente', 'Média Pontuação', 'Total Monitorias', 'Erros Críticos', 'Média Erro Crítico (%)', 'Média Conforme (%)', 'Média Não Conforme (%)']
    dashboard_tree['columns'] = default_dashboard_columns
    for col in default_dashboard_columns:
        dashboard_tree.heading(col, text=col)
        dashboard_tree.column(col, width=150, anchor='center', stretch=tk.NO)

    try:
        with sqlite3.connect(DB_FILE) as conn:
            query = "SELECT * FROM monitoria"
            conditions, params = [], []

            if filtro_agente and filtro_agente != "Todos":
                conditions.append('"Nome do Agente" = ?')
                params.append(filtro_agente)
            if filtro_equipe and filtro_equipe != "Todas":
                conditions.append('Equipe = ?')
                params.append(filtro_equipe)
            if filtro_avaliacao:
                conditions.append('"Avaliação ATD." = ?')
                params.append(str(filtro_avaliacao))
            if filtro_pontuacao:
                conditions.append('Pontuação = ?')
                params.append(str(filtro_pontuacao))

            # intervalo de datas (Data M é armazenada como dd/mm/YYYY)
            # Convertendo para YYYYMMDD para comparação no SQLite
            expr_ymd = "substr(\"Data M\",7,4) || substr(\"Data M\",4,2) || substr(\"Data M\",1,2)"
            if data_ini:
                conditions.append(f"{expr_ymd} >= ?")
                params.append(_to_ymd(data_ini))
            if data_fim:
                conditions.append(f"{expr_ymd} <= ?")
                params.append(_to_ymd(data_fim))

            if conditions:
                query += " WHERE " + " AND ".join(conditions)
            
            df = pd.read_sql_query(query, conn, params=params)
        
        if df.empty:
            if canvas_bar: canvas_bar.get_tk_widget().destroy()
            if canvas_pie: canvas_pie.get_tk_widget().destroy()
            canvas_bar, canvas_pie = None, None
            return

        df['Pontuação'] = pd.to_numeric(df['Pontuação'], errors='coerce')
        df['Itens Aplicáveis'] = pd.to_numeric(df['Itens Aplicáveis'], errors='coerce')
        
        df['Total Conforme'] = df[YES_NO_FIELDS].eq('Conforme').sum(axis=1)
        df['Total Não Conforme'] = df[YES_NO_FIELDS].eq('Não Conforme').sum(axis=1)
        df['Total Itens Validos'] = df[YES_NO_FIELDS].isin(['Conforme', 'Não Conforme']).sum(axis=1)
        
        grouped = df.groupby('Nome do Agente')
        metrics = grouped.agg({
            'Pontuação': 'mean',
            'Protocolo': 'count',
            'Erro Crítico?': lambda x: (x == 'Sim').sum(),
            'Total Conforme': 'sum',
            'Total Não Conforme': 'sum',
            'Total Itens Validos': 'sum'
        }).reset_index()
        metrics.columns = ['Agente', 'Média Pontuação', 'Total Monitorias', 'Erros Críticos', 'Total Conforme', 'Total Não Conforme', 'Total Itens Validos']

        # CORREÇÃO: Tratamento de divisão por zero
        metrics['Média Conforme (%)'] = (metrics['Total Conforme'] / metrics['Total Itens Validos'] * 100).fillna(0).round(2)
        metrics['Média Não Conforme (%)'] = (metrics['Total Não Conforme'] / metrics['Total Itens Validos'] * 100).fillna(0).round(2)
        metrics['Média Erro Crítico (%)'] = (metrics['Erros Críticos'] / metrics['Total Monitorias'] * 100).fillna(0).round(2)
        
        metrics['Média Pontuação'] = pd.to_numeric(metrics['Média Pontuação'], errors='coerce').fillna(0).round(2)
        
        display_metrics = metrics[['Agente', 'Média Pontuação', 'Total Monitorias', 'Erros Críticos', 'Média Erro Crítico (%)', 'Média Conforme (%)', 'Média Não Conforme (%)']]
        
        for _, row in display_metrics.iterrows():
            row_values = list(row)
            row_values[1] = f"{row['Média Pontuação']:.2f}"
            dashboard_tree.insert("", "end", values=row_values)
        
        update_charts(metrics, filtro_agente, filtro_equipe, data_ini, data_fim)

    except Exception as e:
        messagebox.showerror("Erro ao Carregar Dashboard", f"Erro ao carregar dashboard: {e}")

def update_charts(df, filtro_agente, filtro_equipe, data_ini=None, data_fim=None):
    """Atualiza os gráficos de barras e pizza na aba Dashboard."""
    global canvas_bar, canvas_pie
    
    if canvas_bar: canvas_bar.get_tk_widget().destroy()
    if canvas_pie: canvas_pie.get_tk_widget().destroy()
    canvas_bar, canvas_pie = None, None
    
    if df.empty: return
    
    # Gráfico de barras
    fig_bar, ax_bar = plt.subplots(figsize=(6, 4))
    ax_bar.bar(df['Agente'], pd.to_numeric(df['Média Pontuação'], errors='coerce').fillna(0), color='#4A90E2')
    ax_bar.set_title('Média de Pontuação por Agente', color='white')
    ax_bar.set_xlabel('Agente', color='white')
    ax_bar.set_ylabel('Pontuação', color='white')
    ax_bar.set_ylim(0, 10)
    ax_bar.tick_params(axis='x', rotation=45, colors='white')
    ax_bar.tick_params(axis='y', colors='white')
    fig_bar.patch.set_facecolor('#2a2d2e')
    ax_bar.set_facecolor('#2a2d2e')
    plt.tight_layout()
    
    canvas_bar = FigureCanvasTkAgg(fig_bar, master=charts_frame)
    canvas_bar.draw()
    canvas_bar.get_tk_widget().pack(side='left', padx=5, pady=5, fill='both', expand=True)
    plt.close(fig_bar)
    
    # Gráfico de pizza
    with sqlite3.connect(DB_FILE) as conn:
        query = 'SELECT "Erro Crítico?", COUNT(*) as count FROM monitoria'
        conditions, params = [], []
        if filtro_agente and filtro_agente != "Todos":
            conditions.append('"Nome do Agente" = ?')
            params.append(filtro_agente)
        if filtro_equipe and filtro_equipe != "Todas":
            conditions.append('Equipe = ?')
            params.append(filtro_equipe)
        expr_ymd = "substr(\"Data M\",7,4) || substr(\"Data M\",4,2) || substr(\"Data M\",1,2)"
        if data_ini:
            conditions.append(f"{expr_ymd} >= ?")
            params.append(_to_ymd(data_ini))
        if data_fim:
            conditions.append(f"{expr_ymd} <= ?")
            params.append(_to_ymd(data_fim))
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        query += ' GROUP BY "Erro Crítico?"'
        df_pie = pd.read_sql_query(query, conn, params=params)
    
    if not df_pie.empty:
        labels = df_pie['Erro Crítico?']
        sizes = df_pie['count']
        colors = ['#FF4C4C' if label == 'Sim' else '#4A90E2' for label in labels]
        explode = [0.1 if label == 'Sim' else 0 for label in labels]
        
        fig_pie, ax_pie = plt.subplots(figsize=(4, 4))
        ax_pie.pie(sizes, explode=explode, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90, textprops={'color':"w"})
        ax_pie.set_title('Proporção de Erros Críticos', color='white')
        fig_pie.patch.set_facecolor('#2a2d2e')
        plt.tight_layout()
        
        canvas_pie = FigureCanvasTkAgg(fig_pie, master=charts_frame)
        canvas_pie.draw()
        canvas_pie.get_tk_widget().pack(side='left', padx=5, pady=5, fill='both', expand=True)
        plt.close(fig_pie)

def aplicar_filtros_dashboard():
    """Aplica filtros ao Dashboard."""
    agente = combo_filtro_agente_dashboard.get()
    equipe = combo_filtro_equipe_dashboard.get()
    avaliacao = entry_filtro_avaliacao_dashboard.get().strip()
    pontuacao = entry_filtro_pontuacao_dashboard.get().strip()
    data_ini = entry_data_ini_dashboard.get_date() if entry_data_ini_dashboard.get() else None
    data_fim = entry_data_fim_dashboard.get_date() if entry_data_fim_dashboard.get() else None
    if data_ini and data_fim and data_ini > data_fim:
        messagebox.showwarning("Período inválido", "A data inicial não pode ser maior que a data final.")
        return
    atualizar_dashboard(
        filtro_agente=agente if agente != "Todos" else None,
        filtro_equipe=equipe if equipe != "Todas" else None,
        filtro_avaliacao=avaliacao,
        filtro_pontuacao=pontuacao,
        data_ini=data_ini,
        data_fim=data_fim
    )

def limpar_filtros_dashboard():
    """Limpa os filtros do Dashboard."""
    combo_filtro_agente_dashboard.set("Todos")
    combo_filtro_equipe_dashboard.set("Todas")
    entry_filtro_avaliacao_dashboard.delete(0, tk.END)
    entry_filtro_pontuacao_dashboard.delete(0, tk.END)
    try:
        entry_data_ini_dashboard.set_date(datetime.now().replace(day=1))
        entry_data_fim_dashboard.set_date(datetime.now())
    except Exception:
        pass
    atualizar_dashboard()

def gerar_relatorio():
    """Gera um relatório em Excel com os dados atuais."""
    default_dashboard_columns = ['Agente', 'Média Pontuação', 'Total Monitorias', 'Erros Críticos', 'Média Erro Crítico (%)', 'Média Conforme (%)', 'Média Não Conforme (%)']
    try:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_filename = f"Relatorio_Monitoria_{timestamp}.xlsx"
        output_excel = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_filename,
            title="Salvar Relatório Como"
        )
        if not output_excel: return

        # Obter dados do dashboard da Treeview (reflete filtros)
        dashboard_data = [dashboard_tree.item(item, 'values') for item in dashboard_tree.get_children()]
        df_dashboard = pd.DataFrame(dashboard_data, columns=default_dashboard_columns)
        df_dashboard['Média Pontuação'] = pd.to_numeric(df_dashboard['Média Pontuação'], errors='coerce').fillna(0)

        # Obter todos os lançamentos do banco de dados respeitando os filtros atuais do dashboard
        agente = combo_filtro_agente_dashboard.get()
        equipe = combo_filtro_equipe_dashboard.get()
        avaliacao = entry_filtro_avaliacao_dashboard.get().strip()
        pontuacao = entry_filtro_pontuacao_dashboard.get().strip()
        data_ini = entry_data_ini_dashboard.get_date() if entry_data_ini_dashboard.get() else None
        data_fim = entry_data_fim_dashboard.get_date() if entry_data_fim_dashboard.get() else None

        with sqlite3.connect(DB_FILE) as conn:
            query = "SELECT * FROM monitoria"
            conditions, params = [], []
            if agente and agente != "Todos":
                conditions.append('"Nome do Agente" = ?')
                params.append(agente)
            if equipe and equipe != "Todas":
                conditions.append('Equipe = ?')
                params.append(equipe)
            if avaliacao:
                conditions.append('"Avaliação ATD." = ?')
                params.append(str(avaliacao))
            if pontuacao:
                conditions.append('Pontuação = ?')
                params.append(str(pontuacao))
            expr_ymd = "substr(\"Data M\",7,4) || substr(\"Data M\",4,2) || substr(\"Data M\",1,2)"
            if data_ini:
                conditions.append(f"{expr_ymd} >= ?")
                params.append(data_ini.strftime('%Y%m%d'))
            if data_fim:
                conditions.append(f"{expr_ymd} <= ?")
                params.append(data_fim.strftime('%Y%m%d'))
            if conditions:
                query += " WHERE " + " AND ".join(conditions)
            df_lancamentos = pd.read_sql_query(query, conn, params=params)
        if 'id' in df_lancamentos.columns:
            df_lancamentos = df_lancamentos.drop(columns=['id'])
        # Garante a ordem e a presença de todas as colunas
        df_lancamentos = df_lancamentos.reindex(columns=COLUNAS).fillna('')

        # Calcular métricas gerais
        total_monitorias = pd.to_numeric(df_dashboard['Total Monitorias'], errors='coerce').sum()
        media_pontuacao = df_dashboard['Média Pontuação'].mean()
        media_erro_critico = pd.to_numeric(df_dashboard['Média Erro Crítico (%)'], errors='coerce').mean()

        resumo_data = {
            'Métrica': ['Total de Monitorias', 'Média Geral de Pontuação', 'Média de Erros Críticos por Agente (%)'],
            'Valor': [f"{total_monitorias:.0f}", f"{media_pontuacao:.2f}", f"{media_erro_critico:.2f}"]
        }
        df_resumo = pd.DataFrame(resumo_data)

        # Salvar gráficos temporariamente
        temp_dir = os.getcwd()
        bar_path = os.path.join(temp_dir, f"temp_bar_{timestamp}.png")
        pie_path = os.path.join(temp_dir, f"temp_pie_{timestamp}.png")

        # Gráfico de barras (com dados do dashboard filtrado)
        fig_bar, ax_bar = plt.subplots(figsize=(8, 5))
        ax_bar.bar(df_dashboard['Agente'], df_dashboard['Média Pontuação'], color='#4A90E2')
        ax_bar.set_title('Média de Pontuação por Agente')
        ax_bar.set_xlabel('Agente')
        ax_bar.set_ylabel('Pontuação')
        ax_bar.set_ylim(0, 10)
        ax_bar.tick_params(axis='x', rotation=45)
        plt.tight_layout()
        fig_bar.savefig(bar_path)
        plt.close(fig_bar)

        # Gráfico de pizza (com dados do dashboard filtrado)
        total_erros_criticos = pd.to_numeric(df_dashboard['Erros Críticos'], errors='coerce').sum()
        total_sem_erros = total_monitorias - total_erros_criticos
        
        if total_monitorias > 0:
            labels = ['Com Erro Crítico', 'Sem Erro Crítico']
            sizes = [total_erros_criticos, total_sem_erros]
            colors = ['#FF4C4C', '#4A90E2']
            explode = [0.1, 0]

            fig_pie, ax_pie = plt.subplots(figsize=(5, 5))
            ax_pie.pie(sizes, explode=explode, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
            ax_pie.set_title('Proporção de Erros Críticos')
            plt.tight_layout()
            fig_pie.savefig(pie_path)
            plt.close(fig_pie)
        
        # Escrever no Excel
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            df_resumo.to_excel(writer, sheet_name='Resumo', index=False)
            df_dashboard.to_excel(writer, sheet_name='Dashboard', index=False)
            
            # CORREÇÃO: Converte 'Pontuação' para numérico antes de comparar
            df_lancamentos['Pontuação'] = pd.to_numeric(df_lancamentos['Pontuação'], errors='coerce')
            ranking_zero = df_lancamentos[df_lancamentos['Pontuação'] == 0]['Monitoria Zero'].value_counts().reset_index()
            ranking_zero.columns = ['Motivo', 'Quantidade']
            ranking_zero['Motivo'] = ranking_zero['Motivo'].replace('', 'Não especificado').astype(str)
            ranking_zero.to_excel(writer, sheet_name='Ranking Zeros', index=False)

            df_lancamentos['Pontuação'] = df_lancamentos['Pontuação'].apply(lambda x: f"{x:.2f}") # Reverte para string para exibição
            df_lancamentos.to_excel(writer, sheet_name='Lançamentos Completos', index=False)

        # Adicionar filtros aplicados ao Excel (aba Resumo)
        from openpyxl import load_workbook
        from openpyxl.drawing.image import Image
        wb = load_workbook(output_excel)
        ws_resumo = wb['Resumo']
        try:
            periodo_texto = ""
            if entry_data_ini_dashboard.get() and entry_data_fim_dashboard.get():
                periodo_texto = f"Período: {entry_data_ini_dashboard.get()} a {entry_data_fim_dashboard.get()}"
            elif entry_data_ini_dashboard.get():
                periodo_texto = f"Período: a partir de {entry_data_ini_dashboard.get()}"
            elif entry_data_fim_dashboard.get():
                periodo_texto = f"Período: até {entry_data_fim_dashboard.get()}"
            if periodo_texto:
                ws_resumo.cell(row=5, column=1, value=periodo_texto)
        except Exception:
            pass
        
        img_bar = Image(bar_path)
        img_bar.width, img_bar.height = 600, 375
        ws_resumo.add_image(img_bar, 'A10')

        if total_monitorias > 0 and os.path.exists(pie_path):
            img_pie = Image(pie_path)
            img_pie.width, img_pie.height = 375, 375
            ws_resumo.add_image(img_pie, 'J10')
        
        wb.save(output_excel)

        # Remover arquivos temporários
        for file in [bar_path, pie_path]:
            if os.path.exists(file): os.remove(file)

        messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso em: {output_excel}")

    except Exception as e:
        messagebox.showerror("Erro ao Gerar Relatório", f"Ocorreu um erro: {e}\n\nVerifique se o arquivo Excel não está aberto.")
        # Limpa arquivos temporários em caso de falha
        for file in [bar_path, pie_path]:
            if 'path' in locals() and os.path.exists(file): os.remove(file)


def editar_registro():
    """Carrega o registro selecionado para edição no formulário."""
    global edit_mode, edit_id
    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("Nenhum Registro Selecionado", "Selecione um registro para editar.")
        return
    
    item = selected_items[0]
    values = tree.item(item, 'values')
    try:
        edit_id = int(item)
    except ValueError:
    protocolo_selecionado = values[COLUNAS.index('Protocolo')]
    try:
        with sqlite3.connect(DB_FILE) as conn:
            result = conn.execute("SELECT id FROM monitoria WHERE Protocolo = ?", (protocolo_selecionado,)).fetchone()
            if not result:
                messagebox.showerror("Erro", "Não foi possível encontrar o ID do registro.")
                return
            edit_id = result[0]
    except Exception as e:
        messagebox.showerror("Erro de Banco de Dados", f"Não foi possível buscar o ID do registro: {e}")
        return

    limpar_formulario() # Limpa o formulário antes de preencher
    for i, col in enumerate(COLUNAS):
        value = values[i] if i < len(values) else ""
        if col in widgets and widgets[col]:
            if col in YES_NO_FIELDS or col == 'Erro Crítico?' or col in ['Nome do Agente', 'Equipe']:
                widgets[col].set(value)
                if col in CRITICAL_ERRORS:
                    atualizar_cor_critica(widgets[col], col)
            elif col == 'Data M':
                try:
                    widgets[col].set_date(datetime.strptime(value, '%d/%m/%Y'))
                except (ValueError, TypeError):
                    widgets[col].set_date(datetime.now())
            elif col in ['Protocolo', 'Avaliação ATD.']:
                widgets[col].insert(0, value)
            elif col == 'Observações':
                widgets[col].insert("1.0", value)
            # CORREÇÃO: Acessava 'dados', que não existe neste escopo.
            # Agora, preenche corretamente os ComboBoxes.
            elif col == 'Motivo do Atendimento':
                widgets[col].set(value)
            elif col == 'Monitoria Zero':
                widgets[col].set(value if value else 'Nenhum')

    edit_mode = True
    botao_salvar.configure(text="Atualizar Monitoria")
    tabview.set("Nova Monitoria")

# --- CONFIGURAÇÃO DA JANELA PRINCIPAL ---
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Sistema de Monitoria")

# Define ícone da aplicação, se disponível
app_icon_img = None
try:
    if os.path.exists(APP_ICON_FILE):
        app_icon_img = tk.PhotoImage(file=APP_ICON_FILE)
        app.iconphoto(True, app_icon_img)
except Exception:
    pass

screen_width, screen_height = app.winfo_screenwidth(), app.winfo_screenheight()
window_width, window_height = 1200, 750
x, y = (screen_width - window_width) // 2, (screen_height - window_height) // 2
app.geometry(f"{window_width}x{window_height}+{x}+{y}")
app.minsize(1100, 700)

# Exibe a logo no topo, se disponível; caso contrário, mostra texto padrão
label_canaa = None
try:
    if os.path.exists(APP_LOGO_FILE):
        pil_img = Image.open(APP_LOGO_FILE)
        # Redimensiona preservando proporção para caber bem no topo
        max_w, max_h = 300, 60
        w, h = pil_img.size
        scale = min(max_w / float(w), max_h / float(h))
        new_size = (max(1, int(w * scale)), max(1, int(h * scale)))
        ctk_img = ctk.CTkImage(light_image=pil_img.resize(new_size, Image.LANCZOS), size=new_size)
        app_logo_img = ctk_img  # manter referência global
        label_canaa = ctk.CTkLabel(app, image=ctk_img, text="")
        label_canaa.pack(pady=(10, 5))
    else:
        label_canaa = ctk.CTkLabel(app, text="Canaã Telecom", font=ctk.CTkFont(family="Helvetica", size=28, weight="bold"), text_color="#FFFFFF")
        label_canaa.pack(pady=(10, 5))
except Exception:
label_canaa = ctk.CTkLabel(app, text="Canaã Telecom", font=ctk.CTkFont(family="Helvetica", size=28, weight="bold"), text_color="#FFFFFF")
label_canaa.pack(pady=(10, 5))

tabview = ctk.CTkTabview(app, fg_color="#1C2526", segmented_button_fg_color="#4A90E2", segmented_button_selected_color="#2E5A88", text_color="#FFFFFF")
tabview.pack(pady=5, padx=10, fill="both", expand=True)

tab_form = tabview.add("Nova Monitoria")
tab_table = tabview.add("Últimos Lançamentos")
tab_dashboard = tabview.add("Dashboard")
tab_agentes = tabview.add("Configurações")

# --- FORMULÁRIO ---
# Garante que o banco esteja inicializado antes de carregar comboboxes
init_db()
form_frame = ctk.CTkScrollableFrame(tab_form, fg_color="transparent")
form_frame.pack(pady=10, padx=10, fill="both", expand=True)

num_columns = 5
for i in range(num_columns):
    form_frame.grid_columnconfigure(i, weight=1)

# Fontes padrão para a aba 'Nova Monitoria' (Poppins)
FONT_POPPINS_12 = ctk.CTkFont(family="Poppins", size=12)
FONT_POPPINS_14_BOLD = ctk.CTkFont(family="Poppins", size=14, weight="bold")
FONT_POPPINS_20_BOLD = ctk.CTkFont(family="Poppins", size=20, weight="bold")

label_titulo = ctk.CTkLabel(form_frame, text="Nova Monitoria", font=FONT_POPPINS_20_BOLD, text_color="#FFFFFF")
label_titulo.grid(row=0, column=0, columnspan=num_columns, pady=(5, 15), sticky="n")

widgets = {col: None for col in COLUNAS}
equipes, agentes = carregar_dados_iniciais()

def create_widget_frame(parent, label_text, widget_class, widget_kwargs, row, col, colspan=1):
    frame = ctk.CTkFrame(parent, fg_color="transparent")
    frame.grid(row=row, column=col, columnspan=colspan, padx=5, pady=(2, 8), sticky="ew")
    
    text_color = "#FF0000" if label_text in CRITICAL_ERRORS else "#FFFFFF"
    label = ctk.CTkLabel(frame, text=label_text, font=FONT_POPPINS_12, text_color=text_color)
    label.pack(anchor="w", padx=5)

    if widget_class == DateEntry:
        widget = widget_class(frame, **widget_kwargs)
    else:
        widget = widget_class(frame, **widget_kwargs)

    widget.pack(fill="x", expand=True, padx=5)
    # Tenta aplicar fonte padrão Poppins aos widgets suportados
    try:
        widget.configure(font=FONT_POPPINS_12)
    except Exception:
        pass
    return widget

main_fields_layout = [
    ('Protocolo', ctk.CTkEntry, {'fg_color': '#2E5A88', 'text_color': '#FFFFFF', 'font': FONT_POPPINS_12}),
    ('Data M', DateEntry, {'width': 14, 'background': '#4A90E2', 'foreground': 'white', 'borderwidth': 2, 'date_pattern': 'dd/mm/yyyy'}),
    ('Nome do Agente', ctk.CTkComboBox, {'values': agentes, 'fg_color': '#2E5A88', 'text_color': '#FFFFFF', 'font': FONT_POPPINS_12}),
    ('Equipe', ctk.CTkComboBox, {'values': equipes, 'fg_color': '#2E5A88', 'text_color': '#FFFFFF', 'font': FONT_POPPINS_12}),
]
current_row = 1
for idx, (col, widget_type, kwargs) in enumerate(main_fields_layout):
    widgets[col] = create_widget_frame(form_frame, col, widget_type, kwargs, current_row, idx)
widgets['Data M'].set_date(datetime.now())
widgets['Nome do Agente'].set('')
widgets['Equipe'].set('')
form_frame.grid_rowconfigure(current_row, minsize=80)
widgets['Nome do Agente'].bind('<<ComboboxSelected>>', atualizar_equipe)
current_row += 1

for idx, col in enumerate(YES_NO_FIELDS):
    col_idx = idx % num_columns
    row_idx = current_row + (idx // num_columns)
    form_frame.grid_rowconfigure(row_idx, minsize=80)
    widgets[col] = create_widget_frame(form_frame, col, ctk.CTkComboBox, {'values': ['Conforme', 'Não Conforme', 'Não se aplica'], 'fg_color': '#2E5A88', 'text_color': '#FFFFFF', 'font': FONT_POPPINS_12}, row_idx, col_idx)
    widgets[col].set('Conforme')
    if col in CRITICAL_ERRORS:
        widgets[col].bind('<<ComboboxSelected>>', lambda e, w=widgets[col], c=col: atualizar_cor_critica(w, c))
current_row += (len(YES_NO_FIELDS) + num_columns - 1) // num_columns

form_frame.grid_rowconfigure(current_row, minsize=80)
widgets['Motivo do Atendimento'] = create_widget_frame(form_frame, 'Motivo do Atendimento', ctk.CTkComboBox, {'values': ['Conexão', 'Cancelamento', 'Reclamação', 'Financeiro', 'Outros Assuntos'], 'fg_color': '#2E5A88', 'text_color': '#FFFFFF', 'font': FONT_POPPINS_12}, current_row, 0)
widgets['Monitoria Zero'] = create_widget_frame(form_frame, 'Monitoria Zero', ctk.CTkComboBox, {'values': ['Nenhum', 'Ofensa Verbal', 'Erro de Procedimento', 'Confirmação de dados', 'Omissão de Atendimento', 'Inf. Protocolo'], 'fg_color': '#2E5A88', 'text_color': '#FFFFFF', 'font': FONT_POPPINS_12}, current_row, 1)
widgets['Motivo do Atendimento'].set('')
widgets['Monitoria Zero'].set('Nenhum')
current_row += 1

other_fields_layout = [
    ('Avaliação ATD.', ctk.CTkEntry, {'fg_color': '#2E5A88', 'text_color': '#FFFFFF', 'font': FONT_POPPINS_12}),
    ('Erro Crítico?', ctk.CTkComboBox, {'values': ['Sim', 'Não'], 'fg_color': '#2E5A88', 'text_color': '#FFFFFF', 'font': FONT_POPPINS_12}),
    ('Inf. Protocolo?', ctk.CTkComboBox, {'values': ['Conforme', 'Não Conforme', 'Não se aplica'], 'fg_color': '#2E5A88', 'text_color': '#FFFFFF', 'font': FONT_POPPINS_12}),
]
form_frame.grid_rowconfigure(current_row, minsize=80)
for idx, (col, widget_type, kwargs) in enumerate(other_fields_layout):
    widgets[col] = create_widget_frame(form_frame, col, widget_type, kwargs, current_row, idx)
widgets['Erro Crítico?'].set('Não')
widgets['Inf. Protocolo?'].set('Conforme')
current_row += 1

form_frame.grid_rowconfigure(current_row, minsize=100)
widgets['Observações'] = create_widget_frame(form_frame, 'Observações', ctk.CTkTextbox, {'height': 100, 'fg_color': '#2E5A88', 'text_color': '#FFFFFF', 'font': FONT_POPPINS_12}, current_row, 0, colspan=num_columns)
current_row += 1

button_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
button_frame.grid(row=current_row, column=0, columnspan=num_columns, pady=20)
botao_salvar = ctk.CTkButton(button_frame, text="Salvar Monitoria", command=salvar_monitoria, font=ctk.CTkFont(family="Arial", size=14, weight="bold"), fg_color="#4A90E2", hover_color="#2E5A88")
botao_salvar.pack()

# --- TABELA ---
table_frame = ctk.CTkFrame(tab_table, fg_color="transparent")
table_frame.pack(pady=10, padx=10, fill="both", expand=True)

filter_frame = ctk.CTkFrame(table_frame, fg_color="transparent")
filter_frame.pack(pady=5, padx=10, fill="x")

ctk.CTkLabel(filter_frame, text="Filtrar por Agente:", font=ctk.CTkFont(family="Arial", size=12)).pack(side="left", padx=5)
combo_filtro_agente = ctk.CTkComboBox(filter_frame, values=["Todos"] + agentes, width=200, fg_color="#2E5A88", text_color="#FFFFFF")
combo_filtro_agente.pack(side="left", padx=5)
combo_filtro_agente.set("Todos")

ctk.CTkLabel(filter_frame, text="Filtrar por Protocolo:", font=ctk.CTkFont(family="Arial", size=12)).pack(side="left", padx=5)
entry_filtro_protocolo = ctk.CTkEntry(filter_frame, width=150, fg_color="#2E5A88", text_color="#FFFFFF")
entry_filtro_protocolo.pack(side="left", padx=5)

ctk.CTkButton(filter_frame, text="Filtrar", command=aplicar_filtros, font=ctk.CTkFont(family="Arial", size=14, weight="bold"), fg_color="#4A90E2", hover_color="#2E5A88").pack(side="left", padx=5)
ctk.CTkButton(filter_frame, text="Limpar Filtros", command=limpar_filtros, font=ctk.CTkFont(family="Arial", size=14, weight="bold"), fg_color="#4A90E2", hover_color="#2E5A88").pack(side="left", padx=5)

ctk.CTkLabel(table_frame, text="Últimos Lançamentos", font=ctk.CTkFont(family="Arial", size=16, weight="bold")).pack(pady=(5, 10))

style = ttk.Style()
style.theme_use("default")
style.configure("Treeview", background="#2a2d2e", foreground="#FFFFFF", fieldbackground="#2a2d2e", borderwidth=0, rowheight=25)
style.map('Treeview', background=[('selected', '#4A90E2')])
style.configure("Treeview.Heading", background="#2E5A88", foreground="#FFFFFF", font=('Arial', 10, 'bold'), relief="flat")
style.map("Treeview.Heading", background=[('active', '#3B6EA8')])

tree_container = ctk.CTkFrame(table_frame, fg_color="transparent")
tree_container.pack(fill="both", expand=True)

tree = ttk.Treeview(tree_container, columns=COLUNAS, show='headings')
vsb = ctk.CTkScrollbar(tree_container, orientation="vertical", command=tree.yview)
vsb.pack(side='right', fill='y')
hsb = ctk.CTkScrollbar(tree_container, orientation="horizontal", command=tree.xview)
hsb.pack(side='bottom', fill='x')
tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
tree.pack(fill="both", expand=True)

# Botão flutuante de chat (FAB)
def toggle_chat_window():
    global chat_window
    if chat_window and tk.Toplevel.winfo_exists(chat_window):
        chat_window.destroy()
        chat_window = None
        return
    chat_window = tk.Toplevel(app)
    chat_window.title("Assistente (pré-IA)")
    chat_window.geometry("360x480")
    chat_window.transient(app)
    chat_window.attributes('-topmost', True)
    # Área de mensagens
    chat_display = tk.Text(chat_window, state='disabled', wrap='word', bg='#1E1E1E', fg='#FFFFFF')
    chat_display.pack(fill='both', expand=True, padx=8, pady=(8,4))
    # Área de envio
    entry_frame = ctk.CTkFrame(chat_window)
    entry_frame.pack(fill='x', padx=8, pady=(0,8))
    chat_entry = ctk.CTkEntry(entry_frame, placeholder_text='Digite sua mensagem...')
    chat_entry.pack(side='left', fill='x', expand=True, padx=(0,8))
    send_btn = ctk.CTkButton(entry_frame, text='Enviar', command=lambda: None)
    send_btn.pack(side='left')
    # Notas: integração de IA será adicionada futuramente; UI já preparada.

# Cria um FAB no canto inferior direito do app
fab_frame = ctk.CTkFrame(app, corner_radius=30, fg_color='transparent')
fab_frame.place(relx=1.0, rely=1.0, x=-20, y=-20, anchor='se')
chat_fab = ctk.CTkButton(fab_frame, text="💬", width=50, height=50, corner_radius=25, font=ctk.CTkFont(size=20, weight='bold'),
                        command=toggle_chat_window, fg_color="#4A90E2", hover_color="#2E5A88")
chat_fab.pack()

button_table_frame = ctk.CTkFrame(table_frame, fg_color="transparent")
button_table_frame.pack(pady=(10, 5))
ctk.CTkButton(button_table_frame, text="Editar Selecionado", command=editar_registro, font=ctk.CTkFont(family="Arial", size=14, weight="bold"), fg_color="#4A90E2", hover_color="#2E5A88").pack(side="left", padx=5)
ctk.CTkButton(button_table_frame, text="Excluir Selecionado", command=excluir_registro, font=ctk.CTkFont(family="Arial", size=14, weight="bold"), fg_color="#FF4C4C", hover_color="#CC3333").pack(side="left", padx=5)

# --- DASHBOARD ---
dashboard_frame = ctk.CTkScrollableFrame(tab_dashboard, fg_color="transparent")
dashboard_frame.pack(pady=10, padx=10, fill="both", expand=True)

filter_dashboard_frame = ctk.CTkFrame(dashboard_frame, fg_color="transparent")
filter_dashboard_frame.pack(pady=5, padx=10, fill="x")

ctk.CTkLabel(filter_dashboard_frame, text="Agente:").pack(side="left", padx=(5,2))
combo_filtro_agente_dashboard = ctk.CTkComboBox(filter_dashboard_frame, values=["Todos"] + agentes, width=180, fg_color="#2E5A88")
combo_filtro_agente_dashboard.pack(side="left", padx=5)
combo_filtro_agente_dashboard.set("Todos")

ctk.CTkLabel(filter_dashboard_frame, text="Equipe:").pack(side="left", padx=(5,2))
combo_filtro_equipe_dashboard = ctk.CTkComboBox(filter_dashboard_frame, values=["Todas"] + equipes, width=140, fg_color="#2E5A88")
combo_filtro_equipe_dashboard.pack(side="left", padx=5)
combo_filtro_equipe_dashboard.set("Todas")

ctk.CTkLabel(filter_dashboard_frame, text="Avaliação ATD.:").pack(side="left", padx=(5,2))
entry_filtro_avaliacao_dashboard = ctk.CTkEntry(filter_dashboard_frame, width=80, fg_color="#2E5A88")
entry_filtro_avaliacao_dashboard.pack(side="left", padx=5)

ctk.CTkLabel(filter_dashboard_frame, text="Pontuação:").pack(side="left", padx=(5,2))
entry_filtro_pontuacao_dashboard = ctk.CTkEntry(filter_dashboard_frame, width=80, fg_color="#2E5A88")
entry_filtro_pontuacao_dashboard.pack(side="left", padx=5)

# Filtros de período (Data M)
ctk.CTkLabel(filter_dashboard_frame, text="De:").pack(side="left", padx=(10,2))
entry_data_ini_dashboard = DateEntry(filter_dashboard_frame, width=12, background="#4A90E2", foreground="white", borderwidth=2, date_pattern='dd/mm/yyyy')
entry_data_ini_dashboard.pack(side="left", padx=5)
entry_data_ini_dashboard.set_date(datetime.now().replace(day=1))

ctk.CTkLabel(filter_dashboard_frame, text="Até:").pack(side="left", padx=(5,2))
entry_data_fim_dashboard = DateEntry(filter_dashboard_frame, width=12, background="#4A90E2", foreground="white", borderwidth=2, date_pattern='dd/mm/yyyy')
entry_data_fim_dashboard.pack(side="left", padx=5)
entry_data_fim_dashboard.set_date(datetime.now())

ctk.CTkButton(filter_dashboard_frame, text="Filtrar", command=aplicar_filtros_dashboard, font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
ctk.CTkButton(filter_dashboard_frame, text="Limpar", command=limpar_filtros_dashboard, font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)

ctk.CTkLabel(dashboard_frame, text="Dashboard", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(5, 10))

dashboard_container = ctk.CTkFrame(dashboard_frame, fg_color="transparent")
dashboard_container.pack(fill="both", expand=True)
dashboard_tree = ttk.Treeview(dashboard_container, show='headings')
dashboard_vsb = ctk.CTkScrollbar(dashboard_container, orientation="vertical", command=dashboard_tree.yview)
dashboard_vsb.pack(side='right', fill='y')
dashboard_hsb = ctk.CTkScrollbar(dashboard_container, orientation="horizontal", command=dashboard_tree.xview)
dashboard_hsb.pack(side='bottom', fill='x')
dashboard_tree.configure(yscrollcommand=dashboard_vsb.set, xscrollcommand=dashboard_hsb.set)
dashboard_tree.pack(fill="both", expand=True)

charts_frame = ctk.CTkFrame(dashboard_frame, fg_color="transparent")
charts_frame.pack(pady=10, fill="x", expand=True)

button_dashboard_frame = ctk.CTkFrame(dashboard_frame, fg_color="transparent")
button_dashboard_frame.pack(pady=10, fill="x")
ctk.CTkButton(button_dashboard_frame, text="Exportar Relatório", command=gerar_relatorio, font=ctk.CTkFont(size=16, weight="bold"), fg_color="#28A745", hover_color="#218838", height=40).pack(pady=10)

# --- CONFIGURAÇÕES ---
agentes_frame = ctk.CTkFrame(tab_agentes, fg_color="transparent")
agentes_frame.pack(pady=10, padx=10, fill="both", expand=True)

ctk.CTkLabel(agentes_frame, text="Configurações", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=(5, 15))

adicionar_frame = ctk.CTkFrame(agentes_frame, fg_color="transparent")
adicionar_frame.pack(pady=10, fill="x")

ctk.CTkLabel(adicionar_frame, text="Nome do Agente:").pack(side="left", padx=5)
entry_novo_agente = ctk.CTkEntry(adicionar_frame, width=200, fg_color="#2E5A88")
entry_novo_agente.pack(side="left", padx=5)

ctk.CTkLabel(adicionar_frame, text="Equipe:").pack(side="left", padx=5)
combo_equipe_novo_agente = ctk.CTkComboBox(adicionar_frame, values=equipes, width=150, fg_color="#2E5A88")
combo_equipe_novo_agente.pack(side="left", padx=5)
combo_equipe_novo_agente.set('')

botao_adicionar_agente = ctk.CTkButton(adicionar_frame, text="Adicionar Agente", command=adicionar_agente, font=ctk.CTkFont(weight="bold"))
botao_adicionar_agente.pack(side="left", padx=10)
ctk.CTkButton(adicionar_frame, text="Alterar Senha Admin", command=alterar_senha_admin, font=ctk.CTkFont(weight="bold")).pack(side="left", padx=10)

# Seletor de Tema (Claro/Escuro)
tema_frame = ctk.CTkFrame(agentes_frame, fg_color="transparent")
tema_frame.pack(pady=10, fill="x")
ctk.CTkLabel(tema_frame, text="Tema:").pack(side="left", padx=5)
tema_switch = ctk.CTkSwitch(tema_frame, text="Modo Escuro", onvalue=1, offvalue=0)
tema_switch.select()  # padrão escuro
tema_switch.pack(side="left", padx=10)

def _on_tema_toggle():
    mode = 'dark' if tema_switch.get() == 1 else 'light'
    set_theme(mode)
tema_switch.configure(command=_on_tema_toggle)

list_frame = ctk.CTkFrame(agentes_frame, fg_color="transparent")
list_frame.pack(pady=10, fill="both", expand=True)
ctk.CTkLabel(list_frame, text="Agentes Cadastrados:", anchor="w").pack(fill="x", padx=5)
listbox_agentes = tk.Listbox(list_frame, height=10, bg="#2a2d2e", fg="#FFFFFF", font=("Arial", 12), selectbackground="#4A90E2", selectforeground="#FFFFFF")
listbox_agentes.pack(fill="both", expand=True, padx=5, pady=5)
listbox_scroll = ctk.CTkScrollbar(list_frame, orientation="vertical", command=listbox_agentes.yview)
listbox_scroll.pack(side="right", fill="y")
listbox_agentes.configure(yscrollcommand=listbox_scroll.set)
try:
	with sqlite3.connect(DB_FILE) as conn:
		df_init = pd.read_sql_query('SELECT nome, equipe FROM agentes ORDER BY nome COLLATE NOCASE', conn)
	for _, r in df_init.iterrows():
		listbox_agentes.insert(tk.END, f"{r['nome']} ({r['equipe']})")
except Exception:
for agente in sorted(AGENTES_EQUIPE.keys()):
    listbox_agentes.insert(tk.END, f"{agente} ({AGENTES_EQUIPE[agente]})")

button_agentes_frame = ctk.CTkFrame(list_frame, fg_color="transparent")
button_agentes_frame.pack(pady=5)
ctk.CTkButton(button_agentes_frame, text="Excluir Agente", command=excluir_agente, font=ctk.CTkFont(weight="bold"), fg_color="#FF4C4C", hover_color="#CC3333").pack(side="left", padx=5)
ctk.CTkButton(button_agentes_frame, text="Editar Agente", command=editar_agente, font=ctk.CTkFont(weight="bold")).pack(side="left", padx=5)
ctk.CTkButton(button_agentes_frame, text="Limpar Lançamentos", command=limpar_lancamentos, font=ctk.CTkFont(weight="bold"), fg_color="#FF4C4C", hover_color="#CC3333").pack(side="left", padx=5)

# --- INICIALIZAÇÃO ---
if __name__ == "__main__":
    # init_db já foi chamado antes de carregar o formulário, mas chamamos novamente por segurança idempotente
    init_db()
    atualizar_ultimos_lancamentos()
    atualizar_dashboard()
    app.mainloop()
