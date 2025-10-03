"""
Script: upgrade_pacotes.py
Objetivo: Atualizar todos os pacotes Python do ambiente atual.
Uso: python upgrade_pacotes.py
"""

import subprocess
import pkg_resources 

def listar_pacotes_desatualizados():
    """Retorna uma lista de pacotes que estão desatualizados"""
    pacotes_desatualizados = []
    for dist in pkg_resources.working_set:
        try:
            resultado = subprocess.run(
                ["pip", "install", "--upgrade", dist.project_name, "--dry-run"],
                capture_output=True,
                text=True
            )
            if "Requirement already satisfied" not in resultado.stdout:
                pacotes_desatualizados.append(dist.project_name)
        except Exception as e:
            print(f"Erro verificando {dist.project_name}: {e}")
    return pacotes_desatualizados

def atualizar_pacotes(pacotes):
    """Atualiza a lista de pacotes fornecida"""
    for pacote in pacotes:
        print(f"Atualizando {pacote}...")
        try:
            subprocess.run(["pip", "install", "--upgrade", pacote], check=True)
            print(f"{pacote} atualizado com sucesso!\n")
        except subprocess.CalledProcessError:
            print(f"Falha ao atualizar {pacote}\n")

def main():
    print("Verificando pacotes desatualizados...")
    pacotes = listar_pacotes_desatualizados()
    if not pacotes:
        print("Todos os pacotes já estão atualizados!")
        return
    print(f"{len(pacotes)} pacotes serão atualizados: {', '.join(pacotes)}\n")
    atualizar_pacotes(pacotes)
    print("Atualização finalizada!")

if __name__ == "__main__":
    main()
