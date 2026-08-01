import os
import re
import shutil
import sqlite3
import sys
import time
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
from typing import Dict, List, Optional, Tuple

from colorama import Fore, Style, init
from docx import Document

# Inicializa o Colorama para suporte a cores ANSI (Windows/Linux/Mac)
init(autoreset=True)

# Definição de Cores
RESET = Style.RESET_ALL
VERMELHO = Fore.RED + Style.BRIGHT
VERDE = Fore.GREEN + Style.BRIGHT
AMARELO = Fore.YELLOW + Style.BRIGHT
CIANO = Fore.CYAN + Style.BRIGHT
NEGRITO = Style.BRIGHT
AZUL = Fore.BLUE + Style.BRIGHT
MAGENTA = Fore.MAGENTA + Style.BRIGHT
CINZA_ESCURO = Fore.BLACK + Style.BRIGHT


def limpar_terminal() -> None:
    """Limpa a tela do terminal e o buffer de rolagem."""
    os.system("cls" if os.name == "nt" else "clear")


def menu_pausa() -> None:
    input(f"\n{AMARELO}Pressione Enter para voltar ao menu...{RESET}")


def input_cancelavel(prompt: str) -> Tuple[str, bool]:
    """Captura a entrada do usuário e permite cancelar com 'q'."""
    entrada = input(f"{prompt} | [Q] Cancelar: ").strip()
    if entrada.lower() == "q":
        return "", True
    return entrada, False


def hoje() -> str:
    """Retorna a data atual formatada (YYYY-MM-DD)."""
    return datetime.now().strftime("%Y-%m-%d")


def formatar_data(data_str: str) -> str:
    """Valida e converte sequências de dígitos de datas em YYYY-MM-DD."""
    numeros = "".join(filter(str.isdigit, data_str))
    if len(numeros) != 8:
        return ""

    for fmt in ("%Y%m%d", "%d%m%Y"):
        try:
            dt = datetime.strptime(numeros, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


def validar_formatar_doc(doc: str) -> str:
    """Valida e formata CPF (11 dígitos) ou CNPJ (14 dígitos)."""
    n = "".join(filter(str.isdigit, doc))
    if len(n) not in (11, 14) or len(set(n)) == 1:
        return ""

    def calc_digito(s: str, pesos: List[int]) -> int:
        soma = sum(int(a) * b for a, b in zip(s, pesos))
        resto = soma % 11
        return 0 if resto < 2 else 11 - resto

    if len(n) == 11:
        d1 = calc_digito(n[:9], list(range(10, 1, -1)))
        d2 = calc_digito(n[:9] + str(d1), list(range(11, 1, -1)))
        if n[9:] != f"{d1}{d2}":
            return ""
        return f"{n[:3]}.{n[3:6]}.{n[6:9]}-{n[9:]}"
    else:
        pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
        pesos2 = [6] + pesos1
        d1 = calc_digito(n[:12], pesos1)
        d2 = calc_digito(n[:12] + str(d1), pesos2)
        if n[12:] != f"{d1}{d2}":
            return ""
        return f"{n[:2]}.{n[2:5]}.{n[5:8]}/{n[8:12]}-{n[12:]}"


def formatar_telefone(telefone: str) -> str:
    fone = "".join(filter(str.isdigit, telefone))
    if not fone.startswith("55"):
        fone = "55" + fone
    return f"+{fone}" if len(fone) >= 12 else ""


def formatar_endereco(endereco: str) -> str:
    minusculas = {"de", "do", "da", "dos", "das", "e", "com", "em", "apto", "apt", "sala", "bloco", "quadra", "lote", "lt.", "qd.", "lt", "qd"}
    maiusculas = {"rj", "sp", "mg", "es", "pr", "sc", "rs", "ms", "mt", "go", "df", "to", "pa", "am", "ro", "ac", "ap", "rr", "ma", "pi", "ce", "rn", "pb", "pe", "al", "se", "ba", "cep", "bpc", "loas"}
    
    palavras = endereco.lower().split()
    resultado = []
    for p in palavras:
        if p in maiusculas:
            resultado.append(p.upper())
        elif p in minusculas:
            resultado.append(p)
        else:
            resultado.append(p.capitalize())
    return " ".join(resultado)


def formatar_processo(processo: str) -> str:
    n = "".join(filter(str.isdigit, processo))
    n = n[-20:].zfill(20)
    return f"{n[0:7]}-{n[7:9]}.{n[9:13]}.{n[13]}.{n[14:16]}.{n[16:20]}"


def selecionar_arquivo(titulo: str = "Selecione o arquivo", extensoes: List[Tuple[str, str]] = [("Documentos Word", "*.docx")]) -> str:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    caminho = filedialog.askopenfilename(title=titulo, filetypes=extensoes)
    root.destroy()
    return caminho


def selecionar_salvar(nome_padrao: str = "arquivo.docx", extensoes: List[Tuple[str, str]] = [("Documentos Word", "*.docx")]) -> str:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    caminho = filedialog.asksaveasfilename(title="Escolha onde salvar", initialfile=nome_padrao, filetypes=extensoes)
    root.destroy()
    return caminho


# Database Operations
def criar_db(caminho: str) -> None:
    try:
        with sqlite3.connect(caminho) as conn:
            cursor = conn.cursor()
            cursor.execute("PRAGMA foreign_keys = ON;")
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS clientes (
                    CPF_CNPJ TEXT PRIMARY KEY,
                    NOME TEXT NOT NULL,
                    DATA_NASCIMENTO TEXT,
                    IDENTIDADE TEXT,
                    ESTADO_CIVIL TEXT,
                    PROFISSAO TEXT,
                    ENDERECO TEXT,
                    TELEFONE TEXT,
                    EMAIL TEXT,
                    SENHA_GOV TEXT,
                    OBS TEXT
                )
            """)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS processos (
                    PROCESSO TEXT PRIMARY KEY,
                    CPF_CNPJ TEXT NOT NULL,
                    CLIENTE TEXT,
                    PARTE_CONTRARIA TEXT,
                    CARTORIO TEXT,
                    SITUACAO TEXT,
                    ULTIMA_VERIFICACAO TEXT,
                    DISTRIBUICAO TEXT,
                    OBS TEXT,
                    FOREIGN KEY (CPF_CNPJ) REFERENCES clientes (CPF_CNPJ)
                )
            """)
    except Exception as e:
        print(f"{VERMELHO}[ERRO] Falha ao criar tabelas: {e}{RESET}")


def conectar_db(caminho: str) -> sqlite3.Connection:
    conn = sqlite3.connect(caminho)
    conn.row_factory = sqlite3.Row
    return conn


def backup_db(origem: str, destino: str) -> None:
    if not os.path.exists(origem):
        print(f"{VERMELHO}[Erro] Arquivo de origem não encontrado.{RESET}")
        return
    try:
        shutil.copy2(origem, destino)
        print(f"{VERDE}BACKUP CONCLUÍDO!\nLocal: {destino}{RESET}")
    except Exception as e:
        print(f"{VERMELHO}[Erro] Falha ao realizar backup: {e}{RESET}")


# Word Processing using python-docx
def substituir_texto_docx(doc: Document, substituicoes: Dict[str, str]) -> None:
    for p in doc.paragraphs:
        for ch, vl in substituicoes.items():
            if ch in p.text:
                p.text = p.text.replace(ch, vl)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for ch, vl in substituicoes.items():
                        if ch in p.text:
                            p.text = p.text.replace(ch, vl)


def preparar_dados_documento(conn: sqlite3.Connection, busca: str) -> Dict[str, str]:
    dados = {}
    cursor = conn.cursor()
    cursor.execute("""
        SELECT PROCESSO, CPF_CNPJ, CLIENTE, PARTE_CONTRARIA, CARTORIO 
        FROM processos WHERE PROCESSO = ? OR CPF_CNPJ = ? LIMIT 1
    """, (busca, busca))
    row = cursor.fetchone()

    if row:
        dt = datetime.now()
        meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", 
                 "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

        dados["<<PROCESSO>>"] = row["PROCESSO"] or ""
        dados["<<CPF_CNPJ>>"] = row["CPF_CNPJ"] or ""
        dados["<<CLIENTE>>"] = (row["CLIENTE"] or "").upper()
        dados["<<REU>>"] = (row["PARTE_CONTRARIA"] or "").upper()
        dados["<<CARTÓRIO>>"] = (row["CARTORIO"] or "").upper()
        
        dados["<<DD>>"] = dt.strftime("%d")
        dados["<<MM>>"] = meses[dt.month - 1]
        dados["<<AAAA>>"] = dt.strftime("%Y")

        cursor.execute("SELECT ENDERECO, PROFISSAO, ESTADO_CIVIL, IDENTIDADE FROM clientes WHERE CPF_CNPJ = ?", (row["CPF_CNPJ"],))
        row_c = cursor.fetchone()
        if row_c:
            dados["<<ENDEREÇO>>"] = row_c["ENDERECO"] or ""
            dados["<<PROFISSÃO>>"] = (row_c["PROFISSAO"] or "").lower()
            dados["<<ESTADO CIVIL>>"] = (row_c["ESTADO_CIVIL"] or "").lower()
            dados["<<RG>>"] = row_c["IDENTIDADE"] or ""

    return dados


def preencher_docx(conn: sqlite3.Connection) -> None:
    print(f"{AZUL}=========== GERADOR DE DOCUMENTOS PROCVIT ===========")
    busca, cancelou = input_cancelavel(f"{AZUL}Digite o Processo ou CPF/CNPJ{RESET}")
    if cancelou or not busca:
        return

    dados = preparar_dados_documento(conn, busca)
    if not dados:
        print(f"{VERMELHO}[!] Nenhum registro encontrado.{RESET}")
        time.sleep(1.5)
        return

    print(f"{AMARELO}\n> Escolha o modelo .docx...{RESET}")
    modelo = selecionar_arquivo("Selecione o Modelo de Documento")
    if not modelo:
        return

    sugestao = f"Gerado_{dados.get('<<CLIENTE>>', 'doc')}.docx"
    print(f"{AMARELO}> Onde deseja salvar o documento preenchido?{RESET}")
    destino = selecionar_salvar(sugestao)

    if destino:
        try:
            doc = Document(modelo)
            substituir_texto_docx(doc, dados)
            doc.save(destino)
            print(f"\n{VERDE}[OK] Documento gerado com sucesso!{RESET}")
            os.startfile(destino) if os.name == 'nt' else os.system(f'xdg-open "{destino}"')
        except Exception as e:
            print(f"{VERMELHO}[!] Erro ao gerar/abrir documento: {e}{RESET}")
        time.sleep(2)


# Features: Cadastros, Edições e Buscas
def cadastrar_cliente(conn: sqlite3.Connection) -> None:
    while True:
        limpar_terminal()
        print("\n----------- NOVO CADASTRO DE CLIENTE -----------")
        doc_raw, cancel = input_cancelavel("🪪 CPF/CNPJ")
        if cancel: return
        cpf = validar_formatar_doc(doc_raw)
        if not cpf:
            print(f"{VERMELHO}[!] Documento inválido.{RESET}")
            time.sleep(1.5); continue

        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM clientes WHERE CPF_CNPJ = ?", (cpf,))
        if cursor.fetchone():
            print(f"{AMARELO}[!] Cliente já cadastrado.{RESET}")
            time.sleep(1.5); continue

        nome, c = input_cancelavel("👤 Nome Completo"); 
        if c: return
        nasc, c = input_cancelavel("📅 Data Nasc (DDMMAAAA)"); 
        if c: return
        rg, c = input_cancelavel("🪪 RG"); 
        if c: return
        end, c = input_cancelavel("📍 Endereço"); 
        if c: return
        tel, c = input_cancelavel("📞 Telefone"); 
        if c: return
        email, c = input_cancelavel("📧 Email"); 
        if c: return
        est_civil, c = input_cancelavel("📋 Estado Civil"); 
        if c: return
        prof, c = input_cancelavel("💼 Profissão"); 
        if c: return
        senha, c = input_cancelavel("🔑 Senha GOV"); 
        if c: return
        obs, c = input_cancelavel("📝 Observações"); 
        if c: return

        try:
            cursor.execute("""
                INSERT INTO clientes VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (cpf, nome.upper().strip(), formatar_data(nasc), rg.upper().strip(), 
                  est_civil.upper().strip(), prof.upper().strip(), formatar_endereco(end), 
                  formatar_telefone(tel), email.lower().strip(), senha, obs.strip()))
            conn.commit()
            print(f"\n{VERDE}[OK] Cliente cadastrado com sucesso!{RESET}")
        except Exception as e:
            print(f"{VERMELHO}[ERRO]: {e}{RESET}")
        
        time.sleep(2)
        cont, _ = input_cancelavel("\nCadastrar outro? [Enter] Sim")
        if cont.lower() == "q": break


def buscar_cliente(conn: sqlite3.Connection) -> None:
    while True:
        termo_input = input("\n 👤 Buscar Cliente | [Q] Voltar: ").strip()
        if termo_input.lower() == "q": return

        termo = f"%{termo_input}%"
        pagina = 0
        limite = 5

        while True:
            limpar_terminal()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT COUNT(*) FROM clientes 
                WHERE NOME LIKE ? OR CPF_CNPJ LIKE ? OR PROFISSAO LIKE ? OR EMAIL LIKE ?
            """, (termo, termo, termo, termo))
            total = cursor.fetchone()[0]
            total_paginas = max(1, (total + limite - 1) // limite)

            cursor.execute("""
                SELECT * FROM clientes 
                WHERE NOME LIKE ? OR CPF_CNPJ LIKE ? OR PROFISSAO LIKE ? OR EMAIL LIKE ?
                ORDER BY NOME ASC LIMIT ? OFFSET ?
            """, (termo, termo, termo, termo, limite, pagina * limite))
            rows = cursor.fetchall()

            if not rows and pagina == 0:
                print(f"\n{AMARELO}[!] Nenhum registro encontrado.{RESET}")
                time.sleep(1.5)
                break

            print(f"\n{AZUL}--- BUSCA CLIENTES | Pág: {pagina + 1}/{total_paginas} ---{RESET}")
            for i, r in enumerate(rows):
                print("-" * 80)
                print(f"{VERDE}{i + 1}.{RESET} 👤 {r['NOME']} | RG: {r['IDENTIDADE']} | CPF/CNPJ: {r['CPF_CNPJ']}")
                print(f"   End: {r['ENDERECO']} | Tel: {r['TELEFONE']} | Email: {r['EMAIL']}")
                print(f"   Obs: {r['OBS']}")
            print("-" * 80)

            acao = input("[<] Ant | [>] Próx | [E+Nº] Editar | [G] Gerar Doc | [Q] Voltar: ").strip().lower()
            if acao == "q": break
            elif acao == ">" and pagina + 1 < total_paginas: pagina += 1
            elif acao == "<" and pagina > 0: pagina -= 1
            elif acao == "g": preencher_docx(conn)
            elif acao.startswith("e"):
                try:
                    idx = int(acao.replace("e", "")) - 1
                    if 0 <= idx < len(rows):
                        editar_cliente(conn, rows[idx]['CPF_CNPJ'])
                except ValueError: pass


def editar_cliente(conn: sqlite3.Connection, cpf: str) -> None:
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM clientes WHERE CPF_CNPJ = ?", (cpf,))
    row = cursor.fetchone()
    if not row: return

    dados = dict(row)
    limpar_terminal()
    print(f"\n{AZUL}Editando Cliente: {dados['NOME']}{RESET}")
    novo_nome, _ = input_cancelavel(f"Nome [{dados['NOME']}]")
    if novo_nome: dados['NOME'] = novo_nome.upper()

    novo_tel, _ = input_cancelavel(f"Telefone [{dados['TELEFONE']}]")
    if novo_tel: dados['TELEFONE'] = formatar_telefone(novo_tel)

    nova_obs, _ = input_cancelavel(f"Obs [{dados['OBS']}]")
    if nova_obs: dados['OBS'] = nova_obs

    cursor.execute("""
        UPDATE clientes SET NOME=?, TELEFONE=?, OBS=? WHERE CPF_CNPJ=?
    """, (dados['NOME'], dados['TELEFONE'], dados['OBS'], cpf))
    conn.commit()
    print(f"{VERDE}[OK] Cliente atualizado!{RESET}")
    time.sleep(1.5)


def cadastrar_processo(conn: sqlite3.Connection) -> None:
    while True:
        limpar_terminal()
        print("------- NOVO CADASTRO DE PROCESSO -------")
        proc_raw, c = input_cancelavel("⚖️ Número do Processo")
        if c: return
        num_proc = formatar_processo(proc_raw)

        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM processos WHERE PROCESSO = ?", (num_proc,))
        if cursor.fetchone():
            print(f"{AMARELO}[!] Processo já cadastrado.{RESET}")
            time.sleep(1.5); continue

        doc_cli, c = input_cancelavel("🪪 CPF/CNPJ do Cliente")
        if c: return
        cpf_cli = validar_formatar_doc(doc_cli)
        
        cursor.execute("SELECT NOME FROM clientes WHERE CPF_CNPJ = ?", (cpf_cli,))
        cli_row = cursor.fetchone()
        if not cli_row:
            print(f"{VERMELHO}[!] Cliente não cadastrado.{RESET}")
            time.sleep(2); return

        parte, c = input_cancelavel("⚔️ Parte Contrária")
        if c: return
        cartorio, c = input_cancelavel("🏛️ Cartório/Vara")
        if c: return
        obs, c = input_cancelavel("📝 Observações")
        if c: return

        try:
            cursor.execute("""
                INSERT INTO processos VALUES (?, ?, ?, ?, ?, 'ATIVO', ?, ?, ?)
            """, (num_proc, cpf_cli, cli_row['NOME'], parte.upper(), cartorio.upper(), hoje(), hoje(), obs.upper()))
            conn.commit()
            print(f"\n{VERDE}[OK] Processo cadastrado com sucesso!{RESET}")
        except Exception as e:
            print(f"{VERMELHO}[ERRO]: {e}{RESET}")
            
        time.sleep(2)
        break


def buscar_processo(conn: sqlite3.Connection) -> None:
    while True:
        termo_input = input("\n 📄 Buscar Processo | [Q] Voltar: ").strip()
        if termo_input.lower() == "q": return

        termo = f"%{termo_input}%"
        pagina = 0
        limite = 5

        while True:
            limpar_terminal()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT COUNT(*) FROM processos 
                WHERE PROCESSO LIKE ? OR CLIENTE LIKE ? OR PARTE_CONTRARIA LIKE ? OR CPF_CNPJ LIKE ?
            """, (termo, termo, termo, termo))
            total = cursor.fetchone()[0]
            total_paginas = max(1, (total + limite - 1) // limite)

            cursor.execute("""
                SELECT * FROM processos 
                WHERE PROCESSO LIKE ? OR CLIENTE LIKE ? OR PARTE_CONTRARIA LIKE ? OR CPF_CNPJ LIKE ?
                ORDER BY DISTRIBUICAO DESC LIMIT ? OFFSET ?
            """, (termo, termo, termo, termo, limite, pagina * limite))
            rows = cursor.fetchall()

            print(f"\n{AZUL}--- PROCESSOS | Pág: {pagina + 1}/{total_paginas} ---{RESET}")
            for i, r in enumerate(rows):
                print("-" * 80)
                print(f"{VERDE}{i + 1}.{RESET} 📄 PROC: {r['PROCESSO']} | SITUAÇÃO: {r['SITUACAO']}")
                print(f"   CLIENTE: {r['CLIENTE']} | CONTRÁRIO: {r['PARTE_CONTRARIA']}")
                print(f"   CARTÓRIO: {r['CARTORIO']} | ÚLT. VERIF: {r['ULTIMA_VERIFICACAO']}")
            print("-" * 80)

            acao = input("[<] Ant | [>] Próx | [A+Nº] Atualizar Visto | [Q] Voltar: ").strip().lower()
            if acao == "q": break
            elif acao == ">" and pagina + 1 < total_paginas: pagina += 1
            elif acao == "<" and pagina > 0: pagina -= 1
            elif acao.startswith("a"):
                try:
                    idx = int(acao.replace("a", "")) - 1
                    if 0 <= idx < len(rows):
                        cursor.execute("UPDATE processos SET ULTIMA_VERIFICACAO = ? WHERE PROCESSO = ?", (hoje(), rows[idx]['PROCESSO']))
                        conn.commit()
                except ValueError: pass


def ver_andamento(conn: sqlite3.Connection) -> None:
    limpar_terminal()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT * FROM processos 
        WHERE SITUACAO = 'ATIVO' AND ULTIMA_VERIFICACAO <= date('now', '-25 days')
        ORDER BY ULTIMA_VERIFICACAO ASC
    """)
    rows = cursor.fetchall()

    if not rows:
        print(f"\n{VERDE}[OK] Nenhum processo parado há mais de 25 dias!{RESET}")
        time.sleep(1.5)
        return

    print(f"\n{AMARELO}-------------- PROCESSOS PARADOS (HÁ 25+ DIAS) --------------{RESET}")
    for i, r in enumerate(rows):
        print(f"{VERDE}{i + 1}.{RESET} PROC: {r['PROCESSO']} | Client: {r['CLIENTE']} | ÚLTIMA VERIF: {r['ULTIMA_VERIFICACAO']}")
    
    menu_pausa()