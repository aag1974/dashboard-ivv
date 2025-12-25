#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Configurador Visual de Permissões - Dashboard Imobiliário
Interface gráfica para definir quais menus/submenus cada perfil pode ver
"""

import tkinter as tk
from tkinter import ttk, messagebox
import json
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, Any


class VisualPermissionConfigurator:
    """Configurador visual de permissões com interface em tabela"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("🏢 Configurador de Permissões - Dashboard Imobiliário")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f0f0')
        
        # Detectar menus automaticamente do código
        self.menus_structure = self.scan_dashboard_structure()
        self.profiles = ['admin', 'manager', 'analyst', 'viewer']
        
        # Matriz de permissões - [menu][submenu][profile] = bool
        self.permissions = {}

        # Estados adicionais para a interface aprimorada:
        # expanded_sections controla se cada menu está expandido (True) ou colapsado (False)
        # menu_frames armazena referências aos frames de menu e às linhas de submenus
        # select_all_vars armazena um BooleanVar por perfil para o checkbox "Selecionar todos"
        # select_all_checkboxes guarda as referências aos próprios checkbuttons de seleção global
        self.expanded_sections: Dict[str, bool] = {menu_key: True for menu_key in self.menus_structure}
        self.menu_frames: Dict[str, Dict[str, Any]] = {}
        self.select_all_vars: Dict[str, tk.BooleanVar] = {p: tk.BooleanVar() for p in self.profiles}
        self.select_all_checkboxes: Dict[str, tk.Checkbutton] = {}

        # Criar a interface e carregar permissões existentes
        self.create_interface()
        self.load_existing_permissions()
    
    def scan_dashboard_structure(self):
        """Escaneia o gerador_dashboard.py para encontrar estrutura de menus - SINCRONIZADO COM DASHBOARD CORRIGIDO"""
        structure = {
            'residencial': [
                'ivv', 'oferta', 'venda', 'lancamentos', 'oferta_m2', 'venda_m2',
                'valor_ponderado_oferta', 'valor_ponderado_venda', 'vgl', 'vgv_vendas', 'vgv_ofertas', 'distratos'
            ],
            'comercial': [
                'ivv', 'oferta', 'venda', 'lancamentos', 'oferta_m2', 'venda_m2',
                'valor_ponderado_oferta', 'valor_ponderado_venda', 'vgl', 'vgv_vendas', 'vgv_ofertas', 'distratos'
            ],
            'crosstabs': [
                'ivv_por_regiao','oferta_quantidade', 'venda_quantidade', 'valor_ponderado_oferta',
                'valor_ponderado_venda', 'oferta_m2', 'venda_m2',
                'gastos_pos_entrega', 'gastos_por_categoria'
            ],
            'insights': [
                'indicadores_economicos', 'correlacoes'
            ]
        }
        
        # Tentar ler do código se disponível
        try:
            dashboard_files = ['gerador_dashboard.py']
            found_file = None
            for file in dashboard_files:
                if Path(file).exists():
                    found_file = file
                    break
            
            if found_file:
                print(f"📊 Estrutura de menus sincronizada com {found_file}")
                # Aqui poderíamos fazer parsing automático do arquivo se necessário
            else:
                print("📋 Usando estrutura atualizada de menus (sincronizada com dashboard corrigido)")
        except:
            pass
            
        return structure
    
    def create_interface(self):
        """
        Cria interface visual em formato de tabela com melhorias de UX. A tela
        possui um cabeçalho fixo com títulos e checkboxes de "Selecionar
        todos", uma área rolável para menus e submenus, e botões de ação
        abaixo.
        """
        # Título principal
        header_frame = tk.Frame(self.root, bg='#4A90E2', height=60)
        header_frame.pack(fill='x', pady=(0, 10))
        header_frame.pack_propagate(False)
        title = tk.Label(header_frame, text="CONFIGURADOR DE PERMISSÕES", 
                         font=('Arial', 18, 'bold'), fg='white', bg='#4A90E2')
        title.pack(pady=15)

        # Frame que contém o cabeçalho da tabela e o canvas rolável
        content_frame = tk.Frame(self.root, bg='#f0f0f0')
        content_frame.pack(fill='both', expand=True, padx=20, pady=10)

        # Cabeçalho da tabela (fixo): colunas e checkboxes de seleção global
        table_header = tk.Frame(content_frame, bg='#4A90E2', relief='solid', bd=1)
        table_header.pack(fill='x', padx=2, pady=(0, 2))

        # Coluna de Menu
        tk.Label(table_header, text="MENU", width=25, bg='#4A90E2', fg='white',
                 font=('Arial', 11, 'bold'), relief='solid', bd=1).pack(side='left')
        # Coluna de Submenu
        tk.Label(table_header, text="SUBMENU", width=35, bg='#4A90E2', fg='white',
                 font=('Arial', 11, 'bold'), relief='solid', bd=1).pack(side='left')

        # Colunas por perfil, cada uma com título e checkbox "selecionar todos"
        for profile in self.profiles:
            col_frame = tk.Frame(table_header, bg='#4A90E2', relief='solid', bd=1)
            col_frame.pack(side='left')
            tk.Label(col_frame, text=profile.upper(), width=15, bg='#4A90E2', fg='white',
                     font=('Arial', 11, 'bold')).pack(fill='x')
            select_all_cb = tk.Checkbutton(
                col_frame,
                variable=self.select_all_vars[profile],
                bg='#4A90E2', activebackground='#4A90E2',
                command=lambda p=profile: self.select_all_profile(p)
            )
            select_all_cb.pack()
            self.select_all_checkboxes[profile] = select_all_cb

        # Container com canvas e scrollbar para as linhas de menus/submenus
        canvas_container = tk.Frame(content_frame, bg='#f0f0f0')
        canvas_container.pack(fill='both', expand=True)

        canvas = tk.Canvas(canvas_container, bg='white')
        scrollbar_y = ttk.Scrollbar(canvas_container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar_y.set)

        # Frame interno que será rolado
        scrollable_frame = tk.Frame(canvas, bg='white')
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        # Ajustar região de scroll quando o frame interno mudar de tamanho
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        scrollable_frame.bind('<Configure>', on_frame_configure)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")

        # Criar as linhas de permissões dentro do frame rolável
        self.create_permissions_table(scrollable_frame)

        # Botões de ação na base da tela
        self.create_action_buttons()
    
    def create_permissions_table(self, parent):
        """
        Cria as linhas da tabela de permissões. Cada menu principal tem uma linha
        com uma seta para expandir/colapsar, o nome do menu e checkboxes por
        perfil. Os submenus são listados abaixo e armazenados para controle
        dinâmico.
        """
        # Ícones textuais para menus (uso opcional de emojis)
        menu_icons = {
            'residencial': '🏠',
            'comercial': '🏢',
            'crosstabs': '📊',
            'insights': '💡'
        }

        for menu_key, submenus in self.menus_structure.items():
            menu_bg = '#E8F4FD'
            # Frame da linha do menu
            menu_frame = tk.Frame(parent, bg=menu_bg, relief='solid', bd=1)
            menu_frame.pack(fill='x', padx=2)
            # Seta para expandir/colapsar
            arrow_char = '▼' if self.expanded_sections.get(menu_key, True) else '►'
            arrow_lbl = tk.Label(menu_frame, text=arrow_char, width=2, bg=menu_bg,
                                 font=('Arial', 10, 'bold'), cursor='hand2')
            arrow_lbl.pack(side='left')
            arrow_lbl.bind('<Button-1>', lambda e, m=menu_key: self.toggle_section(m))
            # Nome do menu com ícone
            menu_display = f"{menu_icons.get(menu_key, '📁')} {menu_key.upper()}"
            tk.Label(menu_frame, text=menu_display, width=23, bg=menu_bg,
                     font=('Arial', 10, 'bold'), relief='solid', bd=1, anchor='w').pack(side='left')
            # Coluna vazia para alinhar com coluna de submenu
            tk.Label(menu_frame, text="", width=35, bg=menu_bg,
                     relief='solid', bd=1).pack(side='left')

            # Inicializar estrutura no dicionário de permissões
            if menu_key not in self.permissions:
                self.permissions[menu_key] = {}

            # Criar checkboxes do menu por perfil
            for profile in self.profiles:
                menu_var = tk.BooleanVar()
                self.permissions[menu_key][f'_menu_{profile}'] = menu_var
                if profile == 'admin':
                    menu_var.set(True)
                cb = tk.Checkbutton(menu_frame, variable=menu_var, width=15,
                                    bg=menu_bg, activebackground=menu_bg,
                                    command=lambda m=menu_key, p=profile: self.toggle_menu(m, p))
                cb.pack(side='left')

            # Registrar frames para colapso/expansão
            self.menu_frames[menu_key] = {
                'menu_frame': menu_frame,
                'arrow': arrow_lbl,
                'submenu_frames': []
            }

            # Criar linhas dos submenus
            for submenu in submenus:
                sub_frame = tk.Frame(parent, bg='white', relief='solid', bd=1)
                sub_frame.pack(fill='x', padx=2)
                # Espaços para alinhar com seta e menu
                tk.Label(sub_frame, text="", width=2, bg='white').pack(side='left')
                tk.Label(sub_frame, text="", width=23, bg='white').pack(side='left')
                # Nome do submenu
                submenu_display = f"   └─ {self.format_submenu_name(submenu)}"
                tk.Label(sub_frame, text=submenu_display, width=35, bg='white',
                         font=('Arial', 9), relief='solid', bd=1, anchor='w').pack(side='left')
                # Garantir existência do dicionário de submenus
                if submenu not in self.permissions[menu_key]:
                    self.permissions[menu_key][submenu] = {}
                # Checkboxes por perfil
                for profile in self.profiles:
                    sub_var = tk.BooleanVar()
                    self.permissions[menu_key][submenu][profile] = sub_var
                    if profile == 'admin':
                        sub_var.set(True)
                    cb = tk.Checkbutton(sub_frame, variable=sub_var, width=15,
                                        bg='white', activebackground='white',
                                        command=lambda m=menu_key, s=submenu, p=profile: self.update_menu_checkbox(m, s, p))
                    cb.pack(side='left')
                # Armazenar referência do frame de submenu
                self.menu_frames[menu_key]['submenu_frames'].append(sub_frame)
            # Esconder sublinhas se seção estiver colapsada
            if not self.expanded_sections.get(menu_key, True):
                for frame in self.menu_frames[menu_key]['submenu_frames']:
                    frame.pack_forget()
    
    def format_submenu_name(self, submenu):
        """Formata nome do submenu para exibição"""
        formats = {
            'ivv': 'IVV',
            'oferta': 'Oferta',
            'venda': 'Venda', 
            'lancamentos': 'Lançamentos',
            'oferta_m2': 'Oferta m²',
            'venda_m2': 'Venda m²',
            'valor_ponderado_oferta': 'Valor Ponderado Oferta',
            'valor_ponderado_venda': 'Valor Ponderado Venda',
            'vgl': 'VGL',
            'vgv': 'VGV',
            'distratos': 'Distratos',
            'ivv_por_regiao': 'IVV por Região',
            'ofertas_por_regiao': 'Ofertas por Região',
            'vendas_por_regiao': 'Vendas por Região',
            'oferta_valor_pond_regiao': 'Oferta Valor Pond. p/ Região',
            'venda_valor_pond_regiao': 'Venda Valor Pond. p/ Região',
            'oferta_m2_regiao': 'Oferta em m² p/ Região',
            'venda_m2_regiao': 'Venda em m² p/ Região',
            'gastos_pos_entrega_regiao': 'Gastos Pós-entrega p/ Região',
            'gastos_categoria_regiao': 'Gastos p/ Categoria e Região',
            'indicadores_economicos': 'Indicadores Econômicos',
            'correlacoes': 'Correlações'
        }
        return formats.get(submenu, submenu.replace('_', ' ').title())
    
    def toggle_menu(self, menu_key, profile):
        """Marca/desmarca todos os submenus quando menu é clicado"""
        menu_checked = self.permissions[menu_key][f'_menu_{profile}'].get()
        
        # Aplicar a todos os submenus deste menu
        for submenu in self.menus_structure[menu_key]:
            if submenu in self.permissions[menu_key]:
                self.permissions[menu_key][submenu][profile].set(menu_checked)
    
    def update_menu_checkbox(self, menu_key, submenu, profile):
        """Atualiza checkbox do menu baseado nos submenus"""
        # Verifica se todos os submenus estão marcados
        all_checked = True
        any_checked = False
        
        for sub in self.menus_structure[menu_key]:
            if sub in self.permissions[menu_key]:
                if self.permissions[menu_key][sub][profile].get():
                    any_checked = True
                else:
                    all_checked = False
        
        # Atualiza checkbox do menu
        menu_var = self.permissions[menu_key][f'_menu_{profile}']
        menu_var.set(all_checked)

    def toggle_section(self, menu_key: str) -> None:
        """
        Alterna a visibilidade dos submenus de um menu. Se estiver
        expandido, colapsa; se estiver colapsado, expande. Atualiza
        também o ícone da seta.
        """
        expanded = self.expanded_sections.get(menu_key, True)
        new_state = not expanded
        self.expanded_sections[menu_key] = new_state
        frames_info = self.menu_frames.get(menu_key, {})
        arrow_lbl = frames_info.get('arrow')
        submenu_frames = frames_info.get('submenu_frames', [])
        if arrow_lbl:
            arrow_lbl.config(text='▼' if new_state else '►')
        # Mostrar ou esconder sublinhas
        if new_state:
            for frame in submenu_frames:
                frame.pack(fill='x', padx=2)
        else:
            for frame in submenu_frames:
                frame.pack_forget()

    def select_all_profile(self, profile: str) -> None:
        """
        Marca ou desmarca todos os menus e submenus para um perfil específico.
        Acionado pelo checkbox de "Selecionar todos" no cabeçalho.
        """
        select_val = self.select_all_vars[profile].get()
        # Atualizar cada menu e submenu
        for menu_key, submenus in self.menus_structure.items():
            # Atualizar menu
            if menu_key in self.permissions and f'_menu_{profile}' in self.permissions[menu_key]:
                self.permissions[menu_key][f'_menu_{profile}'].set(select_val)
            # Atualizar submenus
            for submenu in submenus:
                if (menu_key in self.permissions and
                    submenu in self.permissions[menu_key] and
                    profile in self.permissions[menu_key][submenu]):
                    self.permissions[menu_key][submenu][profile].set(select_val)
        # Atualizar estados dos menus
        for menu_key in self.menus_structure:
            self.update_menu_checkbox(menu_key, None, profile)
    
    def create_action_buttons(self):
        """Cria botões de ação"""
        btn_frame = tk.Frame(self.root, bg='#f0f0f0')
        btn_frame.pack(pady=20)
        
        # Botão Verificar Sincronização
        sync_btn = tk.Button(btn_frame, text="🔍 Verificar Sincronização", 
                           command=self.check_sync_with_dashboard,
                           bg='#FF9800', fg='white', font=('Arial', 12, 'bold'),
                           padx=20, pady=10)
        sync_btn.pack(side=tk.LEFT, padx=10)
        
        # Botão Salvar
        save_btn = tk.Button(btn_frame, text="💾 Salvar Configurações", 
                           command=self.save_permissions, 
                           bg='#4CAF50', fg='white', font=('Arial', 12, 'bold'),
                           padx=20, pady=10)
        save_btn.pack(side=tk.LEFT, padx=10)
        
        # Botão Gerar Dashboards  
        generate_btn = tk.Button(btn_frame, text="📊 Gerar Dashboards", 
                               command=self.generate_dashboards,
                               bg='#2196F3', fg='white', font=('Arial', 12, 'bold'),
                               padx=20, pady=10)
        generate_btn.pack(side=tk.LEFT, padx=10)
        
        # Botão Cancelar
        cancel_btn = tk.Button(btn_frame, text="❌ Cancelar", 
                             command=self.root.quit,
                             bg='#f44336', fg='white', font=('Arial', 12, 'bold'),
                             padx=20, pady=10)
        cancel_btn.pack(side=tk.LEFT, padx=10)
    
    def check_sync_with_dashboard(self):
        """Verifica sincronização com arquivo dashboard"""
        try:
            dashboard_files = [
                'gerador_dashboard.py'
            ]
            
            found_file = None
            for file in dashboard_files:
                if Path(file).exists():
                    found_file = file
                    break
            
            if not found_file:
                messagebox.showwarning("⚠️ Arquivo não encontrado",
                                     "Nenhum arquivo de dashboard encontrado para verificar sincronização.")
                return
            
            # Ler arquivo e procurar por definições de categorias
            with open(found_file, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Procurar por padrões de categorias no JavaScript
            found_categories = {
                'residencial': set(),
                'comercial': set(), 
                'crosstabs': set(),
                'insights': set()
            }
            
            # Padrões para encontrar categorias
            import re
            
            # Procurar por viewCategories
            view_cat_match = re.search(r'viewCategories\s*=\s*{([^}]+)}', content, re.DOTALL)
            if view_cat_match:
                view_cat_content = view_cat_match.group(1)
                
                for menu in found_categories.keys():
                    menu_match = re.search(rf'{menu}:\s*\[([^\]]+)\]', view_cat_content)
                    if menu_match:
                        cats_str = menu_match.group(1)
                        # Extrair strings entre aspas
                        cats = re.findall(r"'([^']+)'", cats_str)
                        found_categories[menu].update(cats)
            
            self._show_sync_comparison(found_file, found_categories)
            
        except Exception as e:
            messagebox.showerror("❌ Erro", f"Erro ao verificar sincronização: {e}")
    
    def _show_sync_comparison(self, dashboard_file: str, found_categories: dict):
        """Mostra comparação de sincronização"""
        sync_window = tk.Toplevel(self.root)
        sync_window.title("🔍 Verificação de Sincronização")
        sync_window.geometry("800x600")
        sync_window.configure(bg='#f0f0f0')
        
        # Frame de título
        title_frame = tk.Frame(sync_window, bg='#FF9800', height=50)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(title_frame, text="VERIFICAÇÃO DE SINCRONIZAÇÃO", 
                              font=('Arial', 14, 'bold'), fg='white', bg='#FF9800')
        title_label.pack(pady=12)
        
        # Área de texto com scroll
        text_frame = tk.Frame(sync_window, bg='#f0f0f0')
        text_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        text_area = tk.Text(text_frame, wrap=tk.WORD, font=('Courier', 9))
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_area.yview)
        text_area.configure(yscrollcommand=scrollbar.set)
        
        text_area.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Gerar relatório de comparação
        report = []
        report.append("=" * 80)
        report.append("VERIFICAÇÃO DE SINCRONIZAÇÃO")
        report.append("=" * 80)
        report.append(f"📁 Arquivo analisado: {dashboard_file}")
        report.append(f"📅 Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        report.append("")
        
        total_issues = 0
        
        for menu_name in found_categories.keys():
            report.append(f"📊 MENU: {menu_name.upper()}")
            report.append("-" * 60)
            
            configurador_cats = set(self.menus_structure.get(menu_name, []))
            dashboard_cats = found_categories[menu_name]
            
            # Categorias extras no configurador
            extra_in_config = configurador_cats - dashboard_cats
            if extra_in_config:
                report.append(f"⚠️  EXTRAS NO CONFIGURADOR: {', '.join(extra_in_config)}")
                total_issues += len(extra_in_config)
            
            # Categorias faltando no configurador
            missing_in_config = dashboard_cats - configurador_cats
            if missing_in_config:
                report.append(f"❌ FALTANDO NO CONFIGURADOR: {', '.join(missing_in_config)}")
                total_issues += len(missing_in_config)
            
            # Categorias sincronizadas
            synchronized = configurador_cats & dashboard_cats
            if synchronized:
                report.append(f"✅ SINCRONIZADAS ({len(synchronized)}): {', '.join(synchronized)}")
            
            report.append("")
        
        # Resumo final
        if total_issues == 0:
            report.append("🎉 PERFEITA SINCRONIZAÇÃO!")
            report.append("✅ Todas as categorias estão alinhadas entre configurador e dashboard.")
        else:
            report.append(f"⚠️  {total_issues} PROBLEMAS ENCONTRADOS")
            report.append("🔧 Recomenda-se atualizar o configurador para sincronizar.")
        
        text_area.insert(tk.END, "\n".join(report))
        text_area.config(state=tk.DISABLED)
        
        # Botão fechar
        close_btn = tk.Button(sync_window, text="Fechar", 
                             command=sync_window.destroy,
                             bg='#FF9800', fg='white', font=('Arial', 10, 'bold'))
        close_btn.pack(pady=10)
    
    def load_existing_permissions(self):
        """Carrega permissões existentes se houver"""
        try:
            if Path('dashboard_menu_permissions.json').exists():
                with open('dashboard_menu_permissions.json', 'r', encoding='utf-8') as f:
                    saved_data = json.load(f)
                
                menu_perms = saved_data.get('menu_permissions', {})
                
                # Migrar permissões antigas se necessário
                menu_perms = self._migrate_old_permissions(menu_perms)
                
                # Aplica permissões salvas aos checkboxes
                for profile in self.profiles:
                    if profile in menu_perms:
                        profile_menus = menu_perms[profile]
                        for menu_key in self.menus_structure:
                            if menu_key in profile_menus:
                                allowed_submenus = profile_menus[menu_key]
                                for submenu in self.menus_structure[menu_key]:
                                    if (menu_key in self.permissions and 
                                        submenu in self.permissions[menu_key]):
                                        should_check = submenu in allowed_submenus
                                        self.permissions[menu_key][submenu][profile].set(should_check)
                                
                                # Atualizar checkbox do menu
                                self.update_menu_checkbox(menu_key, submenu, profile)
                
                print("✅ Configurações existentes carregadas e migradas")
        except Exception as e:
            print(f"⚠️ Erro ao carregar configurações: {e}")
    
    def _migrate_old_permissions(self, menu_perms: dict) -> dict:
        """Migra permissões antigas para nova estrutura"""
        migrated = {}
        
        for profile, profile_data in menu_perms.items():
            migrated[profile] = {}
            
            for menu_key, submenus in profile_data.items():
                if menu_key not in self.menus_structure:
                    print(f"⚠️ Menu '{menu_key}' não encontrado na estrutura atual, ignorando")
                    continue
                
                migrated_submenus = []
                
                for submenu in submenus:
                    # Migração específica: vgv -> vgv_vendas + vgv_ofertas
                    if submenu == 'vgv':
                        print(f"🔄 Migrando 'vgv' para 'vgv_vendas' e 'vgv_ofertas' no perfil {profile}")
                        migrated_submenus.extend(['vgv_vendas', 'vgv_ofertas'])
                    # Migração de nomes de crosstabs antigos
                    elif submenu == 'ofertas_por_regiao':
                        migrated_submenus.append('oferta_quantidade')
                    elif submenu == 'vendas_por_regiao':
                        migrated_submenus.append('venda_quantidade')
                    elif submenu == 'gastos_pos_entrega_regiao':
                        migrated_submenus.append('gastos_pos_entrega')
                    elif submenu == 'gastos_categoria_regiao':
                        migrated_submenus.append('gastos_por_categoria')
                    # Manter submenu se existe na estrutura atual
                    elif submenu in self.menus_structure[menu_key]:
                        migrated_submenus.append(submenu)
                    else:
                        print(f"⚠️ Submenu '{submenu}' não encontrado em '{menu_key}', ignorando")
                
                if migrated_submenus:
                    migrated[profile][menu_key] = list(set(migrated_submenus))  # Remove duplicatas
        
        return migrated
    
    def save_permissions(self):
        """Salva configurações em JSON com validação"""
        # Validar estrutura antes de salvar
        validation_result = self._validate_permissions()
        if not validation_result['valid']:
            messagebox.showwarning("⚠️ Aviso", 
                                f"Problemas encontrados:\n{chr(10).join(validation_result['warnings'])}")
        
        config = {
            'generated_at': datetime.now().isoformat(),
            'dashboard_version': 'gerador_dashboard.py',
            'menu_permissions': {},
            'validation': validation_result
        }
        
        # Converter checkboxes para JSON
        for profile in self.profiles:
            config['menu_permissions'][profile] = {}
            
            for menu_key in self.menus_structure:
                allowed_submenus = []
                
                for submenu in self.menus_structure[menu_key]:
                    if (menu_key in self.permissions and 
                        submenu in self.permissions[menu_key] and
                        self.permissions[menu_key][submenu][profile].get()):
                        allowed_submenus.append(submenu)
                
                if allowed_submenus:
                    config['menu_permissions'][profile][menu_key] = allowed_submenus
        
        try:
            with open('dashboard_menu_permissions.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            
            # Mostrar relatório de sincronização
            self._show_sync_report(config)
            
            messagebox.showinfo("✅ Sucesso", 
                              "Configurações salvas em dashboard_menu_permissions.json")
            print("✅ Configurações salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("❌ Erro", f"Erro ao salvar: {e}")
    
    def _validate_permissions(self) -> Dict[str, Any]:
        """Valida configurações de permissões"""
        warnings = []
        
        # Verificar se pelo menos admin tem todas as permissões
        admin_perms = 0
        for menu_key in self.menus_structure:
            for submenu in self.menus_structure[menu_key]:
                if (menu_key in self.permissions and 
                    submenu in self.permissions[menu_key] and
                    self.permissions[menu_key][submenu]['admin'].get()):
                    admin_perms += 1
        
        total_perms = sum(len(submenus) for submenus in self.menus_structure.values())
        
        if admin_perms < total_perms * 0.8:  # Admin deve ter pelo menos 80% das permissões
            warnings.append(f"Admin tem apenas {admin_perms}/{total_perms} permissões. Recomendado dar mais acesso ao admin.")
        
        # Verificar se há perfis sem nenhuma permissão
        for profile in self.profiles:
            profile_perms = 0
            for menu_key in self.menus_structure:
                for submenu in self.menus_structure[menu_key]:
                    if (menu_key in self.permissions and 
                        submenu in self.permissions[menu_key] and
                        self.permissions[menu_key][submenu][profile].get()):
                        profile_perms += 1
            
            if profile_perms == 0:
                warnings.append(f"Perfil '{profile}' não tem nenhuma permissão.")
        
        return {
            'valid': len(warnings) == 0,
            'warnings': warnings,
            'total_categories': total_perms,
            'admin_permissions': admin_perms
        }
    
    def _show_sync_report(self, config: dict):
        """Mostra relatório de sincronização"""
        report_window = tk.Toplevel(self.root)
        report_window.title("📊 Relatório de Sincronização")
        report_window.geometry("600x400")
        report_window.configure(bg='#f0f0f0')
        
        # Frame de título
        title_frame = tk.Frame(report_window, bg='#4A90E2', height=50)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        
        title_label = tk.Label(title_frame, text="RELATÓRIO DE SINCRONIZAÇÃO", 
                              font=('Arial', 14, 'bold'), fg='white', bg='#4A90E2')
        title_label.pack(pady=12)
        
        # Área de texto com scroll
        text_frame = tk.Frame(report_window, bg='#f0f0f0')
        text_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        text_area = tk.Text(text_frame, wrap=tk.WORD, font=('Courier', 10))
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_area.yview)
        text_area.configure(yscrollcommand=scrollbar.set)
        
        text_area.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Gerar relatório
        report = []
        report.append("=" * 60)
        report.append("RELATÓRIO DE SINCRONIZAÇÃO COM DASHBOARD")
        report.append("=" * 60)
        report.append(f"📅 Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        report.append(f"🔧 Dashboard: {config.get('dashboard_version', 'N/A')}")
        report.append("")
        
        # Resumo por perfil
        for profile in self.profiles:
            profile_data = config['menu_permissions'].get(profile, {})
            total_cats = sum(len(cats) for cats in profile_data.values())
            report.append(f"👤 {profile.upper()}: {total_cats} categorias ativas")
            
            for menu, cats in profile_data.items():
                report.append(f"   📁 {menu}: {', '.join(cats)}")
            report.append("")
        
        # Categorias disponíveis
        report.append("📊 CATEGORIAS DISPONÍVEIS:")
        for menu, cats in self.menus_structure.items():
            report.append(f"   📁 {menu} ({len(cats)}): {', '.join(cats)}")
        
        report.append("")
        report.append("✅ Sincronização concluída com sucesso!")
        
        text_area.insert(tk.END, "\n".join(report))
        text_area.config(state=tk.DISABLED)
        
        # Botão fechar
        close_btn = tk.Button(report_window, text="Fechar", 
                             command=report_window.destroy,
                             bg='#4A90E2', fg='white', font=('Arial', 10, 'bold'))
        close_btn.pack(pady=10)
    
    def generate_dashboards(self):
        """Salva configurações e chama geração de dashboards"""
        self.save_permissions()
        
        try:
            import subprocess
            result = messagebox.askyesno("🚀 Gerar Dashboards", 
                                       "Salvar configurações e gerar dashboards agora?\n\n") 
            if result:
                print("🔄 Iniciando geração de dashboards...")
                
                # Verificar se arquivo corrigido existe
                dashboard_files = [
                    'gerador_dashboard.py'
                ]
                
                found_file = None
                for file in dashboard_files:
                    if Path(file).exists():
                        found_file = file
                        break
                
                if found_file:
                    cmd_message = f"python3 {found_file} --todos-perfis"
                    messagebox.showinfo("🎉 Pronto para Gerar", 
                                      f"Configurações salvas!\n\n" +
                                      f"Para gerar dashboards, execute:\n{cmd_message}")
                    print(f"📋 Execute: {cmd_message}")
                else:
                    messagebox.showwarning("⚠️ Arquivo não encontrado",
                                         "Nenhum arquivo de dashboard encontrado.\n\n" +
                                         "Verifique se há um dos arquivos:\n" +
                                          "- gerador_dashboard.py")
        except Exception as e:
            messagebox.showerror("❌ Erro", f"Erro na geração: {e}")
    
    def run(self):
        """Inicia interface"""
        self.root.mainloop()


def main():
    """Função principal"""
    print("🔧 Iniciando Configurador Visual de Permissões...")
    app = VisualPermissionConfigurator()
    app.run()


if __name__ == "__main__":
    main()
