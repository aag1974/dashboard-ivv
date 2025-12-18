#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Utilitário de Gerenciamento de Usuários e Permissões

Este script fornece uma interface de linha de comando para gerenciar
usuários e permissões do sistema de dashboard imobiliário.
"""

import sys
import json
from datetime import datetime
from typing import Optional

try:
    from user_permission_manager import UserPermissionManager
    PERMISSIONS_AVAILABLE = True
except ImportError:
    print("❌ Módulo user_permission_manager não encontrado!")
    print("   Certifique-se de que o arquivo user_permission_manager.py está no mesmo diretório.")
    PERMISSIONS_AVAILABLE = False
    sys.exit(1)


class UserManagementCLI:
    """Interface de linha de comando para gerenciar usuários."""
    
    def __init__(self, config_file: str = "user_profiles.json"):
        self.pm = UserPermissionManager(config_file)
        self.config_file = config_file
    
    def show_menu(self):
        """Exibe o menu principal."""
        print("\n" + "="*60)
        print("🏢 GERENCIADOR DE USUÁRIOS - DASHBOARD IMOBILIÁRIO")
        print("="*60)
        print("1. 👥 Listar usuários")
        print("2. ➕ Adicionar usuário")
        print("3. ✏️  Editar perfil de usuário")
        print("4. ❌ Desativar usuário")
        print("5. ✅ Ativar usuário")
        print("6. 📊 Mostrar perfis disponíveis")
        print("7. 🔍 Testar autenticação de usuário")
        print("8. 📋 Exportar configuração")
        print("9. 📥 Migrar de allowed_users.json")
        print("0. 🚪 Sair")
        print("-"*60)
    
    def list_users(self):
        """Lista todos os usuários."""
        users = self.pm.list_users()
        
        if not users:
            print("\n📭 Nenhum usuário cadastrado.")
            return
        
        print(f"\n👥 USUÁRIOS CADASTRADOS ({len(users)} total):")
        print("-"*80)
        print(f"{'EMAIL':<30} {'NOME':<20} {'PERFIL':<15} {'STATUS':<10} {'ÚLTIMO ACESSO':<15}")
        print("-"*80)
        
        for user in sorted(users, key=lambda x: x['email']):
            status = "✅ Ativo" if user['active'] else "❌ Inativo"
            last_access = user['last_access']
            if last_access:
                # Formatar data
                try:
                    dt = datetime.fromisoformat(last_access.replace('Z', '+00:00'))
                    last_access = dt.strftime("%d/%m/%y %H:%M")
                except:
                    last_access = "Data inválida"
            else:
                last_access = "Nunca"
            
            print(f"{user['email']:<30} {user['name']:<20} {user['profile']:<15} {status:<10} {last_access:<15}")
    
    def add_user(self):
        """Adiciona um novo usuário."""
        print("\n➕ ADICIONAR NOVO USUÁRIO")
        print("-"*40)
        
        email = input("📧 Email do usuário: ").strip().lower()
        if not email:
            print("❌ Email é obrigatório!")
            return
        
        if not "@" in email:
            print("❌ Email inválido!")
            return
        
        name = input("👤 Nome do usuário: ").strip()
        if not name:
            name = email.split('@')[0].title()
        
        print("\n📋 Perfis disponíveis:")
        profiles = self.pm.config.get("profiles", {})
        for i, (profile_key, profile_data) in enumerate(profiles.items(), 1):
            print(f"  {i}. {profile_key} - {profile_data.get('name', profile_key)}")
            print(f"     {profile_data.get('description', 'Sem descrição')}")
        
        profile_choice = input(f"\n🎯 Escolha o perfil (1-{len(profiles)} ou nome): ").strip()
        
        # Tentar converter para índice
        try:
            profile_index = int(profile_choice) - 1
            profile_key = list(profiles.keys())[profile_index]
        except (ValueError, IndexError):
            # Tentar usar como nome do perfil
            if profile_choice in profiles:
                profile_key = profile_choice
            else:
                print(f"❌ Perfil '{profile_choice}' inválido!")
                return
        
        # Confirmar
        profile_name = profiles[profile_key].get('name', profile_key)
        print(f"\n📄 RESUMO:")
        print(f"   Email: {email}")
        print(f"   Nome: {name}")
        print(f"   Perfil: {profile_name} ({profile_key})")
        
        confirm = input("\n✅ Confirmar adição? (s/N): ").strip().lower()
        if confirm == 's':
            if self.pm.add_user(email, name, profile_key):
                print(f"✅ Usuário {email} adicionado com sucesso!")
            else:
                print(f"❌ Erro ao adicionar usuário!")
        else:
            print("❌ Operação cancelada.")
    
    def edit_user_profile(self):
        """Edita o perfil de um usuário."""
        print("\n✏️ EDITAR PERFIL DE USUÁRIO")
        print("-"*40)
        
        email = input("📧 Email do usuário: ").strip().lower()
        if not email:
            print("❌ Email é obrigatório!")
            return
        
        # Verificar se usuário existe
        users = self.pm.config.get("users", {})
        if email not in users:
            print(f"❌ Usuário {email} não encontrado!")
            return
        
        current_profile = users[email].get("profile", "viewer")
        print(f"📋 Perfil atual: {current_profile}")
        
        print("\n📋 Perfis disponíveis:")
        profiles = self.pm.config.get("profiles", {})
        for i, (profile_key, profile_data) in enumerate(profiles.items(), 1):
            current = " (atual)" if profile_key == current_profile else ""
            print(f"  {i}. {profile_key} - {profile_data.get('name', profile_key)}{current}")
        
        profile_choice = input(f"\n🎯 Novo perfil (1-{len(profiles)} ou nome): ").strip()
        
        # Tentar converter para índice
        try:
            profile_index = int(profile_choice) - 1
            profile_key = list(profiles.keys())[profile_index]
        except (ValueError, IndexError):
            if profile_choice in profiles:
                profile_key = profile_choice
            else:
                print(f"❌ Perfil '{profile_choice}' inválido!")
                return
        
        if profile_key == current_profile:
            print("ℹ️ Perfil selecionado é o mesmo atual.")
            return
        
        # Confirmar
        profile_name = profiles[profile_key].get('name', profile_key)
        print(f"\n📄 ALTERAÇÃO:")
        print(f"   Usuário: {email}")
        print(f"   De: {current_profile}")
        print(f"   Para: {profile_name} ({profile_key})")
        
        confirm = input("\n✅ Confirmar alteração? (s/N): ").strip().lower()
        if confirm == 's':
            if self.pm.update_user_profile(email, profile_key):
                print(f"✅ Perfil do usuário {email} atualizado!")
            else:
                print(f"❌ Erro ao atualizar perfil!")
        else:
            print("❌ Operação cancelada.")
    
    def deactivate_user(self):
        """Desativa um usuário."""
        print("\n❌ DESATIVAR USUÁRIO")
        print("-"*40)
        
        email = input("📧 Email do usuário: ").strip().lower()
        if not email:
            print("❌ Email é obrigatório!")
            return
        
        users = self.pm.config.get("users", {})
        if email not in users:
            print(f"❌ Usuário {email} não encontrado!")
            return
        
        if not users[email].get("active", True):
            print(f"ℹ️ Usuário {email} já está desativado.")
            return
        
        confirm = input(f"❌ Confirmar desativação de {email}? (s/N): ").strip().lower()
        if confirm == 's':
            if self.pm.deactivate_user(email):
                print(f"✅ Usuário {email} desativado!")
            else:
                print(f"❌ Erro ao desativar usuário!")
        else:
            print("❌ Operação cancelada.")
    
    def activate_user(self):
        """Ativa um usuário."""
        print("\n✅ ATIVAR USUÁRIO")
        print("-"*40)
        
        email = input("📧 Email do usuário: ").strip().lower()
        if not email:
            print("❌ Email é obrigatório!")
            return
        
        users = self.pm.config.get("users", {})
        if email not in users:
            print(f"❌ Usuário {email} não encontrado!")
            return
        
        if users[email].get("active", True):
            print(f"ℹ️ Usuário {email} já está ativo.")
            return
        
        users[email]["active"] = True
        self.pm._save_config()
        print(f"✅ Usuário {email} ativado!")
    
    def show_profiles(self):
        """Mostra os perfis disponíveis e suas permissões."""
        print("\n📊 PERFIS DISPONÍVEIS")
        print("-"*60)
        
        profiles = self.pm.config.get("profiles", {})
        
        for profile_key, profile_data in profiles.items():
            name = profile_data.get('name', profile_key)
            description = profile_data.get('description', 'Sem descrição')
            permissions = profile_data.get('permissions', {})
            
            print(f"\n🎯 {name.upper()} ({profile_key})")
            print(f"   📝 {description}")
            print(f"   🔐 Permissões:")
            
            for perm, enabled in permissions.items():
                status = "✅" if enabled else "❌"
                perm_readable = perm.replace('_', ' ').title()
                print(f"     {status} {perm_readable}")
    
    def test_authentication(self):
        """Testa a autenticação de um usuário."""
        print("\n🔍 TESTAR AUTENTICAÇÃO")
        print("-"*40)
        
        email = input("📧 Email do usuário: ").strip().lower()
        if not email:
            print("❌ Email é obrigatório!")
            return
        
        # Criar novo manager para teste
        test_pm = UserPermissionManager(self.config_file)
        
        if test_pm.authenticate_user(email):
            user_info = test_pm.get_user_info()
            filter_config = test_pm.get_filtered_data_config()
            
            print(f"✅ Autenticação bem-sucedida!")
            print(f"   👤 Nome: {user_info['name']}")
            print(f"   🎯 Perfil: {user_info['profile_name']}")
            print(f"   📧 Email: {user_info['email']}")
            
            print(f"\n🔐 Configuração de acesso:")
            for config_key, enabled in filter_config.items():
                status = "✅" if enabled else "❌"
                readable_key = config_key.replace('_', ' ').replace('show ', '').title()
                print(f"   {status} {readable_key}")
        else:
            print(f"❌ Falha na autenticação!")
            print(f"   Verifique se o usuário existe e está ativo.")
    
    def export_config(self):
        """Exporta a configuração atual."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        export_file = f"user_profiles_backup_{timestamp}.json"
        
        try:
            with open(export_file, 'w', encoding='utf-8') as f:
                json.dump(self.pm.config, f, ensure_ascii=False, indent=2)
            
            print(f"✅ Configuração exportada para: {export_file}")
        except Exception as e:
            print(f"❌ Erro ao exportar: {e}")
    
    def migrate_from_allowed_users(self):
        """Migra usuários de um arquivo allowed_users.json."""
        print("\n📥 MIGRAR DE ALLOWED_USERS.JSON")
        print("-"*40)
        
        allowed_file = input("📁 Caminho para allowed_users.json (Enter para 'allowed_users.json'): ").strip()
        if not allowed_file:
            allowed_file = "allowed_users.json"
        
        try:
            with open(allowed_file, 'r', encoding='utf-8') as f:
                allowed_users = json.load(f)
            
            if isinstance(allowed_users, dict) and "users" in allowed_users:
                # Formato com estrutura
                user_emails = list(allowed_users["users"].keys())
            elif isinstance(allowed_users, list):
                # Lista simples de emails
                user_emails = allowed_users
            else:
                print(f"❌ Formato do arquivo não reconhecido!")
                return
            
            print(f"\n📋 Encontrados {len(user_emails)} usuários:")
            for email in user_emails[:5]:  # Mostra primeiros 5
                print(f"   - {email}")
            if len(user_emails) > 5:
                print(f"   ... e mais {len(user_emails) - 5} usuários")
            
            # Escolher perfil padrão
            print("\n🎯 Perfil padrão para usuários migrados:")
            profiles = self.pm.config.get("profiles", {})
            for i, (profile_key, profile_data) in enumerate(profiles.items(), 1):
                print(f"  {i}. {profile_key} - {profile_data.get('name', profile_key)}")
            
            profile_choice = input(f"\nEscolha (1-{len(profiles)}): ").strip()
            try:
                profile_index = int(profile_choice) - 1
                default_profile = list(profiles.keys())[profile_index]
            except (ValueError, IndexError):
                print("❌ Escolha inválida!")
                return
            
            confirm = input(f"\n✅ Migrar {len(user_emails)} usuários com perfil '{default_profile}'? (s/N): ").strip().lower()
            
            if confirm == 's':
                migrated = 0
                for email in user_emails:
                    # Nome baseado no email
                    name = email.split('@')[0].title()
                    try:
                        if self.pm.add_user(email, name, default_profile):
                            migrated += 1
                    except:
                        # Usuário já existe
                        pass
                
                print(f"✅ Migração concluída!")
                print(f"   📊 {migrated} usuários migrados")
                print(f"   📋 {len(user_emails) - migrated} usuários já existiam")
            else:
                print("❌ Migração cancelada.")
        
        except FileNotFoundError:
            print(f"❌ Arquivo {allowed_file} não encontrado!")
        except json.JSONDecodeError:
            print(f"❌ Arquivo {allowed_file} não é um JSON válido!")
        except Exception as e:
            print(f"❌ Erro na migração: {e}")
    
    def run(self):
        """Executa o CLI."""
        while True:
            self.show_menu()
            
            try:
                choice = input("\n🎯 Escolha uma opção (0-9): ").strip()
                
                if choice == '0':
                    print("\n👋 Até logo!")
                    break
                elif choice == '1':
                    self.list_users()
                elif choice == '2':
                    self.add_user()
                elif choice == '3':
                    self.edit_user_profile()
                elif choice == '4':
                    self.deactivate_user()
                elif choice == '5':
                    self.activate_user()
                elif choice == '6':
                    self.show_profiles()
                elif choice == '7':
                    self.test_authentication()
                elif choice == '8':
                    self.export_config()
                elif choice == '9':
                    self.migrate_from_allowed_users()
                else:
                    print("❌ Opção inválida! Digite um número de 0 a 9.")
                
                if choice != '0':
                    input("\n📌 Pressione Enter para continuar...")
            
            except KeyboardInterrupt:
                print("\n\n👋 Saindo...")
                break
            except EOFError:
                print("\n\n👋 Saindo...")
                break


def main():
    """Função principal."""
    if not PERMISSIONS_AVAILABLE:
        return
    
    config_file = "user_profiles.json"
    
    # Permitir especificar arquivo de configuração
    if len(sys.argv) > 1:
        config_file = sys.argv[1]
    
    cli = UserManagementCLI(config_file)
    
    print(f"📁 Usando arquivo de configuração: {config_file}")
    
    try:
        cli.run()
    except Exception as e:
        print(f"\n❌ Erro inesperado: {e}")
        print("   Por favor, reporte este erro ao desenvolvedor.")


if __name__ == "__main__":
    main()
