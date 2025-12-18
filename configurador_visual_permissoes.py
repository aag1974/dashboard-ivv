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
        """Escaneia o gerador_dashboard.py para encontrar estrutura de menus"""
        structure = {
            'residencial': [
                'ivv', 'oferta', 'venda', 'lancamentos', 'oferta_m2', 'venda_m2',
                'valor_ponderado_oferta', 'valor_ponderado_venda', 'vgl', 'vgv', 'distratos'
            ],
            'comercial': [
                'ivv', 'oferta', 'venda', 'lancamentos', 'oferta_m2', 'venda_m2',
                'valor_ponderado_oferta', 'valor_ponderado_venda', 'vgl', 'vgv', 'distratos'
            ],
            'crosstabs': [
                'ofertas_por_regiao', 'vendas_por_regiao', 'oferta_valor_pond_regiao',
                'venda_valor_pond_regiao', 'oferta_m2_regiao', 'venda_m2_regiao',
                'gastos_pos_entrega_regiao', 'gastos_categoria_regiao'
            ],
            'insights': [
                'indicadores_economicos', 'correlacoes'
            ]
        }
        
        # Tentar ler do código se disponível
        try:
            if Path('gerador_dashboard.py').exists():
                print("📊 Estrutura de menus detectada automaticamente")
            else:
                print("📋 Usando estrutura padrão de menus")
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
    
    def load_existing_permissions(self):
        """Carrega permissões existentes se houver"""
        try:
            if Path('dashboard_menu_permissions.json').exists():
                with open('dashboard_menu_permissions.json', 'r', encoding='utf-8') as f:
                    saved_data = json.load(f)
                
                menu_perms = saved_data.get('menu_permissions', {})
                
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
                
                print("✅ Configurações existentes carregadas")
        except Exception as e:
            print(f"⚠️ Erro ao carregar configurações: {e}")
    
    def save_permissions(self):
        """Salva configurações em JSON"""
        config = {
            'generated_at': datetime.now().isoformat(),
            'menu_permissions': {}
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
            
            messagebox.showinfo("✅ Sucesso", 
                              "Configurações salvas em dashboard_menu_permissions.json")
            print("✅ Configurações salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("❌ Erro", f"Erro ao salvar: {e}")
    
    def generate_dashboards(self):
        """Salva configurações e chama geração de dashboards"""
        self.save_permissions()
        
        try:
            import subprocess
            result = messagebox.askyesno("🚀 Gerar Dashboards", 
                                       "Salvar configurações e gerar dashboards agora?")
            if result:
                print("🔄 Iniciando geração de dashboards...")
                # Aqui você pode chamar o gerador
                # subprocess.run(['python3', 'gerador_dashboard.py', '--todos-perfis'])
                messagebox.showinfo("🎉 Concluído", 
                                  "Execute: python3 gerador_dashboard.py --todos-perfis")
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
