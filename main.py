import sys
import time
from funcoes import (
    criar_db, conectar_db, limpar_terminal, buscar_cliente,
    buscar_processo, cadastrar_cliente, cadastrar_processo, preencher_docx,
    ver_andamento, selecionar_salvar, backup_db, hoje,
    VERMELHO, VERDE, AZUL, AMARELO, CIANO, MAGENTA, NEGRITO, RESET
)

# Caminho do Banco de Dados
CAMINHO_DB = "/home/rodrigo/Gdrive/DADOS/procvit.db"


def main():
    criar_db(CAMINHO_DB)
    conn = conectar_db(CAMINHO_DB)

    while True:
        limpar_terminal()
        print(f"{VERMELHO}{NEGRITO}" + "--------------------- PROCVIT LEX ---------------------".center(50) + RESET)
        print(f"{VERDE}[1] 👤 CLIENTES{RESET}")
        print(f"{AZUL}[2] ⚖️  PROCESSOS{RESET}")
        print(f"{AMARELO}[3] ➕ CADASTRO{RESET}")
        print(f"{CIANO}[4] 📑 PETIÇÕES/RECURSOS{RESET}")
        print(f"{MAGENTA}[5] ⚠️  PROCESSOS +25 DIAS SEM VERIFICAÇÃO{RESET}")
        print(f"{AZUL}[6] 💾 BACKUP DO BANCO DE DADOS{RESET}")
        print(f"{VERMELHO}[Q] SAIR ❌{RESET}")
        print(f"{VERMELHO}{NEGRITO}" + "-------------------------------------------------------".center(50) + RESET)

        opcao = input("\nEscolha uma opção: ").strip().lower()

        if opcao == "1":
            limpar_terminal()
            buscar_cliente(conn)
        elif opcao == "2":
            limpar_terminal()
            buscar_processo(conn)
        elif opcao == "3":
            limpar_terminal()
            print(f"{AZUL}----------- ÁREA DE CADASTRO -----------{RESET}")
            print(f"{VERDE}[1] CLIENTE{RESET} | {AMARELO}[2] PROCESSO{RESET} | {VERMELHO}[Q] VOLTAR{RESET}")
            sub = input("\nSeleção: ").strip().lower()
            if sub == "1":
                cadastrar_cliente(conn)
            elif sub == "2":
                cadastrar_processo(conn)
        elif opcao == "4":
            limpar_terminal()
            preencher_docx(conn)
        elif opcao == "5":
            ver_andamento(conn)
        elif opcao == "6":
            nome_sugestao = f"procvit_{hoje()}.db"
            backup_caminho = selecionar_salvar(nome_sugestao, [("Banco de Dados", "*.db")])
            if backup_caminho:
                limpar_terminal()
                backup_db(CAMINHO_DB, backup_caminho)
                time.sleep(2)
            else:
                print(f"\n{AMARELO}Operação de backup cancelada.{RESET}")
                time.sleep(1.5)
        elif opcao == "q":
            limpar_terminal()
            break
        else:
            print(f"\n{AMARELO}Opção inválida! Tente novamente.{RESET}")
            time.sleep(1)

    conn.close()


if __name__ == "__main__":
    main()