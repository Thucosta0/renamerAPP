#!/usr/bin/env python3
"""
Script de preparação para build do renamerPRO© ThTweaks
Otimiza e prepara todos os componentes para empacotamento
"""

import os
import shutil
import subprocess
import json

def verificar_dependencias():
    """Verifica se todas as dependências estão instaladas"""
    print("🔍 Verificando dependências...")
    
    # Verificar Python
    try:
        import customtkinter
        print("✅ CustomTkinter encontrado")
    except ImportError:
        print("❌ CustomTkinter não encontrado!")
        print("💡 Execute: pip install customtkinter")
        return False
    
    try:
        import pandas
        print("✅ Pandas encontrado")
    except ImportError:
        print("❌ Pandas não encontrado!")
        print("💡 Execute: pip install pandas")
        return False
    
    try:
        import openpyxl
        print("✅ OpenPyXL encontrado")
    except ImportError:
        print("❌ OpenPyXL não encontrado!")
        print("💡 Execute: pip install openpyxl")
        return False
    
    # Verificar PHP
    # Verificar se estamos no diretório correto ou no pai
    if os.path.exists("php/php.exe"):
        php_path = "php/php.exe"
    elif os.path.exists("renamerPRO/php/php.exe"):
        php_path = "renamerPRO/php/php.exe"
        # Mudar para o diretório correto
        os.chdir("renamerPRO")
    else:
        print("❌ PHP não encontrado!")
        print("💡 Certifique-se que a pasta php/ existe com php.exe")
        print(f"💡 Diretório atual: {os.getcwd()}")
        return False
    
    print("✅ PHP encontrado")
    
    # Verificar Composer
    if not os.path.exists("vendor"):
        print("❌ Dependências PHP não instaladas!")
        print("💡 Execute: composer install")
        return False
    else:
        print("✅ Dependências PHP encontradas")
    
    return True

def otimizar_php():
    """Otimiza instalação PHP removendo arquivos desnecessários"""
    print("🔧 Otimizando instalação PHP...")
    
    # Arquivos/pastas que podem ser removidos para reduzir tamanho
    remover = [
        "php/dev",
        "php/extras/ssl",
        "php/php.ini-development",
        "php/php.ini-production", 
        "php/license.txt",
        "php/news.txt",
        "php/readme-redist-bins.txt",
        "php/snapshot.txt"
    ]
    
    removidos = 0
    for item in remover:
        if os.path.exists(item):
            try:
                if os.path.isdir(item):
                    shutil.rmtree(item)
                else:
                    os.remove(item)
                removidos += 1
                print(f"  🗑️ Removido: {item}")
            except Exception as e:
                print(f"  ⚠️ Não foi possível remover {item}: {e}")
    
    print(f"✅ {removidos} itens desnecessários removidos")

def otimizar_vendor():
    """Otimiza pasta vendor removendo arquivos de desenvolvimento"""
    print("🔧 Otimizando dependências PHP...")
    
    # Padrões de arquivos/pastas para remover
    padroes_remover = [
        "*/tests",
        "*/test", 
        "*/.git",
        "*/.github",
        "*/docs",
        "*/examples",
        "*/sample*",
        "*/*.md",
        "*/README*",
        "*/CHANGELOG*",
        "*/LICENSE*"
    ]
    
    removidos = 0
    for root, dirs, files in os.walk("vendor"):
        # Remover diretórios desnecessários
        dirs_to_remove = []
        for d in dirs:
            if d in ['tests', 'test', '.git', '.github', 'docs', 'examples']:
                dirs_to_remove.append(d)
        
        for d in dirs_to_remove:
            path = os.path.join(root, d)
            try:
                shutil.rmtree(path)
                removidos += 1
                dirs.remove(d)  # Não continuar explorando
            except:
                pass
        
        # Remover arquivos desnecessários
        for f in files:
            if any(f.lower().endswith(ext) for ext in ['.md', '.txt']) and \
               any(name in f.lower() for name in ['readme', 'changelog', 'license', 'contributing']):
                try:
                    os.remove(os.path.join(root, f))
                    removidos += 1
                except:
                    pass
    
    print(f"✅ {removidos} arquivos de desenvolvimento removidos")

def criar_manifest():
    """Cria manifest com informações do build"""
    manifest = {
        "name": "renamerPRO© ThTweaks",
        "version": "1.0.0",
        "description": "Sistema de Processamento e Renomeação de DANFEs",
        "features": [
            "Processamento em massa de XMLs para PDFs",
            "Renomeação inteligente de XMLs e PDFs",
            "Associação automática XML-PDF",
            "Interface profissional CustomTkinter",
            "Processamento paralelo otimizado",
            "Documentação completa incluída (LEIA-ME.txt e README.md)",
            "Script de inicialização automática incluso"
        ],
        "components": {
            "python": "3.x",
            "php": "8.4.8",
            "customtkinter": "5.2.0+",
            "nfephp": "5.1+"
        },
        "build_info": {
            "target": "Windows x64",
            "type": "Standalone Executable",
            "includes": ["PHP Runtime", "Python Runtime", "All Dependencies"]
        }
    }
    
    with open("manifest.json", "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
    
    print("✅ Manifest criado")

def verificar_tamanho():
    """Verifica tamanho estimado do build"""
    print("📊 Calculando tamanho estimado...")
    
    total_size = 0
    componentes = {
        "PHP": "php",
        "Vendor": "vendor", 
        "Python Scripts": "."
    }
    
    for nome, path in componentes.items():
        if os.path.exists(path):
            size = 0
            if os.path.isfile(path):
                size = os.path.getsize(path)
            else:
                for root, dirs, files in os.walk(path):
                    for file in files:
                        try:
                            size += os.path.getsize(os.path.join(root, file))
                        except:
                            pass
            
            size_mb = size / (1024 * 1024)
            total_size += size_mb
            print(f"  📁 {nome}: {size_mb:.1f} MB")
    
    print(f"📊 Tamanho total estimado: {total_size:.1f} MB")
    print(f"💾 Executável final estimado: {total_size * 1.5:.1f} MB (com compressão)")

def main():
    """Função principal de preparação"""
    print("🏥 renamerPRO©- Preparação para Build")
    print("=" * 60)
    
    if not verificar_dependencias():
        print("\n❌ Dependências não atendidas!")
        return
    
    print("\n🔧 Iniciando otimizações...")
    
    # Fazer backup antes das otimizações
    print("💾 Criando backup...")
    backup_dir = "backup"  # Inicializar variável
    
    if os.path.exists("backup"):
        try:
            shutil.rmtree("backup")
        except PermissionError as e:
            print(f"⚠️ Não foi possível remover backup anterior: {e}")
            print("💡 Tentando criar backup com nome único...")
            import time
            backup_name = f"backup_{int(time.time())}"
            if os.path.exists(backup_name):
                shutil.rmtree(backup_name)
            backup_dir = backup_name  # Atualizar para o novo nome
        except Exception as e:
            print(f"❌ Erro ao remover backup: {e}")
            return
    
    if not os.path.exists(backup_dir):
        os.makedirs(backup_dir, exist_ok=True)
    
    # Backup dos componentes principais
    if os.path.exists("php"):
        shutil.copytree("php", f"{backup_dir}/php")
    if os.path.exists("vendor"):
        shutil.copytree("vendor", f"{backup_dir}/vendor")
    
    print("✅ Backup criado")
    
    # Executar otimizações
    otimizar_php()
    otimizar_vendor()
    criar_manifest()
    verificar_tamanho()
    
    print("\n🎉 PREPARAÇÃO CONCLUÍDA!")
    print("=" * 60)
    print("✅ Projeto otimizado para build")
    print(f"📁 Backup salvo em: {backup_dir}/")
    print("🚀 Execute agora: python build_exe.py")
    print(f"\n💡 Para restaurar backup: copie {backup_dir}/* para ./")

if __name__ == "__main__":
    main()
