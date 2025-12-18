#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sistema de Controle de Acesso e Permissões para Dashboard Imobiliário
"""

import json
import os
from datetime import datetime
from typing import Dict, List, Any, Optional


class UserPermissionManager:
    """
    Gerenciador de permissões e controle de acesso por perfil de usuário.
    
    Esta classe gerencia diferentes níveis de acesso ao dashboard imobiliário,
    permitindo ocultar informações sensíveis baseado no perfil do usuário.
    """
    
    def __init__(self, config_file: str = "user_profiles.json"):
        """
        Inicializa o gerenciador de permissões.
        
        Args:
            config_file: Caminho para o arquivo de configuração JSON
        """
        self.config_file = config_file
        self.config = self._load_config()
        self.current_user = None
        self.current_permissions = None
    
    def _load_config(self) -> Dict[str, Any]:
        """Carrega configuração de usuários e perfis."""
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"⚠️ Arquivo de configuração {self.config_file} não encontrado.")
            return self._create_default_config()
        except json.JSONDecodeError as e:
            print(f"❌ Erro ao ler configuração: {e}")
            return self._create_default_config()
    
    def _create_default_config(self) -> Dict[str, Any]:
        """Cria configuração padrão caso não exista."""
        default_config = {
            "profiles": {
                "admin": {
                    "name": "Administrador",
                    "permissions": {
                        "view_company_details": True,
                        "view_project_names": True,
                        "view_financial_data": True,
                        "view_detailed_metrics": True,
                        "view_sensitive_reports": True,
                        "download_private_txt": True,
                        "view_competitor_analysis": True,
                        "access_raw_data": True
                    }
                }
            },
            "users": {}
        }
        
        # Salva configuração padrão
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, ensure_ascii=False, indent=2)
            print(f"✅ Configuração padrão criada: {self.config_file}")
        except Exception as e:
            print(f"❌ Erro ao criar configuração padrão: {e}")
        
        return default_config
    
    def authenticate_user(self, email: str) -> bool:
        """
        Autentica usuário e carrega suas permissões.
        
        Args:
            email: Email do usuário (obtido do OAuth do Google)
            
        Returns:
            bool: True se usuário autorizado, False caso contrário
        """
        users = self.config.get("users", {})
        
        if email not in users:
            print(f"❌ Usuário {email} não autorizado")
            return False
        
        user_data = users[email]
        if not user_data.get("active", False):
            print(f"❌ Usuário {email} está desativado")
            return False
        
        # Atualiza último acesso
        user_data["last_access"] = datetime.now().isoformat()
        self._save_config()
        
        # Carrega perfil e permissões
        profile_name = user_data.get("profile", "viewer")
        profile = self.config.get("profiles", {}).get(profile_name, {})
        
        self.current_user = {
            "email": email,
            "name": user_data.get("name", email),
            "profile": profile_name,
            "profile_name": profile.get("name", profile_name.title())
        }
        
        self.current_permissions = profile.get("permissions", {})
        
        print(f"✅ Usuário autenticado: {self.current_user['name']} ({self.current_user['profile_name']})")
        return True
    
    def has_permission(self, permission: str) -> bool:
        """
        Verifica se usuário atual tem determinada permissão.
        
        Args:
            permission: Nome da permissão a verificar
            
        Returns:
            bool: True se tem permissão, False caso contrário
        """
        if not self.current_permissions:
            return False
        
        return self.current_permissions.get(permission, False)
    
    def get_filtered_data_config(self) -> Dict[str, bool]:
        """
        Retorna configuração de filtragem baseada nas permissões do usuário.
        
        Returns:
            Dict com flags de exibição para cada tipo de dado
        """
        if not self.current_permissions:
            # Usuário não autenticado = apenas dados públicos básicos
            return {
                "show_company_names": False,
                "show_project_details": False,
                "show_financial_details": False,
                "show_detailed_metrics": False,
                "show_sensitive_charts": False,
                "enable_txt_download": False,
                "show_competitor_data": False,
                "show_raw_data_tables": False
            }
        
        return {
            "show_company_names": self.has_permission("view_company_details"),
            "show_project_details": self.has_permission("view_project_names"),
            "show_financial_details": self.has_permission("view_financial_data"),
            "show_detailed_metrics": self.has_permission("view_detailed_metrics"),
            "show_sensitive_charts": self.has_permission("view_sensitive_reports"),
            "enable_txt_download": self.has_permission("download_private_txt"),
            "show_competitor_data": self.has_permission("view_competitor_analysis"),
            "show_raw_data_tables": self.has_permission("access_raw_data")
        }
    
    def _save_config(self):
        """Salva configuração atualizada."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"❌ Erro ao salvar configuração: {e}")
    
    def add_user(self, email: str, name: str, profile: str = "viewer") -> bool:
        """
        Adiciona novo usuário ao sistema.
        
        Args:
            email: Email do usuário
            name: Nome do usuário
            profile: Perfil de acesso (admin, manager, analyst, viewer)
            
        Returns:
            bool: True se usuário adicionado com sucesso
        """
        if profile not in self.config.get("profiles", {}):
            print(f"❌ Perfil '{profile}' não existe")
            return False
        
        users = self.config.setdefault("users", {})
        users[email] = {
            "name": name,
            "profile": profile,
            "active": True,
            "created_at": datetime.now().isoformat(),
            "last_access": None
        }
        
        self._save_config()
        print(f"✅ Usuário {email} adicionado com perfil '{profile}'")
        return True
    
    def update_user_profile(self, email: str, new_profile: str) -> bool:
        """
        Atualiza perfil de um usuário.
        
        Args:
            email: Email do usuário
            new_profile: Novo perfil de acesso
            
        Returns:
            bool: True se atualizado com sucesso
        """
        users = self.config.get("users", {})
        
        if email not in users:
            print(f"❌ Usuário {email} não encontrado")
            return False
        
        if new_profile not in self.config.get("profiles", {}):
            print(f"❌ Perfil '{new_profile}' não existe")
            return False
        
        users[email]["profile"] = new_profile
        self._save_config()
        print(f"✅ Usuário {email} atualizado para perfil '{new_profile}'")
        return True
    
    def deactivate_user(self, email: str) -> bool:
        """
        Desativa um usuário.
        
        Args:
            email: Email do usuário
            
        Returns:
            bool: True se desativado com sucesso
        """
        users = self.config.get("users", {})
        
        if email not in users:
            print(f"❌ Usuário {email} não encontrado")
            return False
        
        users[email]["active"] = False
        self._save_config()
        print(f"✅ Usuário {email} desativado")
        return True
    
    def list_users(self) -> List[Dict[str, Any]]:
        """
        Lista todos os usuários cadastrados.
        
        Returns:
            Lista de dicionários com dados dos usuários
        """
        users = self.config.get("users", {})
        user_list = []
        
        for email, data in users.items():
            profile_name = self.config.get("profiles", {}).get(
                data.get("profile", ""), {}
            ).get("name", data.get("profile", "Desconhecido"))
            
            user_list.append({
                "email": email,
                "name": data.get("name", ""),
                "profile": data.get("profile", ""),
                "profile_name": profile_name,
                "active": data.get("active", False),
                "last_access": data.get("last_access")
            })
        
        return user_list
    
    def get_user_info(self) -> Optional[Dict[str, Any]]:
        """
        Retorna informações do usuário atual.
        
        Returns:
            Dict com informações do usuário ou None se não autenticado
        """
        return self.current_user


class DataSanitizer:
    """
    Classe responsável por limpar/mascarar dados sensíveis baseado nas permissões.
    """
    
    def __init__(self, permission_manager: UserPermissionManager):
        self.permission_manager = permission_manager
        self.filter_config = permission_manager.get_filtered_data_config()
    
    def sanitize_dataframe(self, df, data_type: str = "residential"):
        """
        Sanitiza DataFrame removendo/mascarando dados sensíveis.
        
        Args:
            df: DataFrame pandas com dados
            data_type: Tipo de dados (residential, commercial, etc.)
            
        Returns:
            DataFrame sanitizado
        """
        if df is None or df.empty:
            return df
        
        sanitized_df = df.copy()
        
        # Remove detalhes de empresas se não tem permissão
        if not self.filter_config["show_company_names"]:
            if 'EMPRESA' in sanitized_df.columns:
                sanitized_df['EMPRESA'] = 'EMPRESA PRIVADA'
            if 'INCORPORADORA' in sanitized_df.columns:
                sanitized_df['INCORPORADORA'] = 'INCORPORADORA PRIVADA'
        
        # Remove nomes de projetos se não tem permissão
        if not self.filter_config["show_project_details"]:
            if 'EMPREENDIMENTO' in sanitized_df.columns:
                sanitized_df['EMPREENDIMENTO'] = f'PROJETO {hash(sanitized_df["EMPREENDIMENTO"]) % 1000:03d}'
        
        # Remove dados financeiros detalhados se não tem permissão
        if not self.filter_config["show_financial_details"]:
            financial_cols = ['VALOR_TOTAL', 'FINANCIAMENTO_VALOR', 'ENTRADA_VALOR']
            for col in financial_cols:
                if col in sanitized_df.columns:
                    # Mantém apenas faixas de valores em vez de valores exatos
                    sanitized_df[col] = self._categorize_values(sanitized_df[col])
        
        return sanitized_df
    
    def _categorize_values(self, series):
        """Categoriza valores em faixas em vez de mostrar valores exatos."""
        def get_value_range(value):
            if pd.isna(value) or value == 0:
                return "N/A"
            elif value < 500000:
                return "< 500K"
            elif value < 1000000:
                return "500K - 1M"
            elif value < 2000000:
                return "1M - 2M"
            else:
                return "> 2M"
        
        return series.apply(get_value_range)
    
    def sanitize_json_data(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Sanitiza dados JSON para exibição baseado nas permissões.
        
        Args:
            data: Dicionário com dados para sanitizar
            
        Returns:
            Dicionário sanitizado
        """
        if not isinstance(data, dict):
            return data
        
        sanitized_data = data.copy()
        
        # Remove seções inteiras baseado nas permissões
        if not self.filter_config["show_sensitive_charts"]:
            # Remove gráficos sensíveis
            keys_to_remove = [
                'competitor_analysis', 'detailed_financial_metrics',
                'company_performance', 'project_details'
            ]
            for key in keys_to_remove:
                sanitized_data.pop(key, None)
        
        if not self.filter_config["show_detailed_metrics"]:
            # Simplifica métricas
            if 'metrics' in sanitized_data:
                basic_metrics = ['total_units', 'average_price_m2', 'total_projects']
                sanitized_data['metrics'] = {
                    k: v for k, v in sanitized_data['metrics'].items() 
                    if k in basic_metrics
                }
        
        return sanitized_data
    
    def should_show_download_buttons(self) -> bool:
        """Verifica se deve exibir botões de download de arquivos privados."""
        return self.filter_config["enable_txt_download"]
    
    def get_visible_sections(self) -> List[str]:
        """
        Retorna lista de seções que devem ser visíveis no HTML.
        
        Returns:
            Lista de IDs de seções a exibir
        """
        visible_sections = ['summary', 'charts', 'basic_metrics']
        
        if self.filter_config["show_detailed_metrics"]:
            visible_sections.extend(['detailed_metrics', 'advanced_charts'])
        
        if self.filter_config["show_competitor_data"]:
            visible_sections.append('competitor_analysis')
        
        if self.filter_config["show_raw_data_tables"]:
            visible_sections.append('raw_data_tables')
        
        if self.filter_config["enable_txt_download"]:
            visible_sections.append('download_section')
        
        return visible_sections


# Exemplo de uso
if __name__ == "__main__":
    # Inicializar gerenciador de permissões
    pm = UserPermissionManager("user_profiles.json")
    
    # Adicionar usuários de exemplo
    pm.add_user("admin@empresa.com", "Administrador", "admin")
    pm.add_user("gerente@empresa.com", "Gerente Regional", "manager")
    pm.add_user("analista@empresa.com", "Analista Junior", "analyst")
    pm.add_user("visualizador@empresa.com", "Usuário Básico", "viewer")
    
    # Autenticar usuário
    if pm.authenticate_user("analista@empresa.com"):
        print(f"Usuário autenticado: {pm.get_user_info()}")
        print(f"Configuração de filtros: {pm.get_filtered_data_config()}")
    
    # Listar todos os usuários
    print("\nUsuários cadastrados:")
    for user in pm.list_users():
        status = "✅ Ativo" if user["active"] else "❌ Inativo"
        print(f"  {user['email']} - {user['name']} ({user['profile_name']}) {status}")
