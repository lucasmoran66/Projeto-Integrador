# app_avaliacao_cli.py
# Versão console (sem interface gráfica) do avaliador de trabalhos
# Requisitos: pandas (pip install pandas) e openpyxl (se for usar .xlsx)

import os                       # Verificações de existência de arquivo e caminhos
import re                       # Validação simples de e-mail
from datetime import datetime   # Timestamp das avaliações
import pandas as pd             # Leitura e escrita dos dados

# ===================== CONFIGURAÇÕES =====================

# Caminhos dos arquivos de entrada (banco) e saída (avaliações)
CSV_PATH = "banco_congresso_ficticio.csv"
XLSX_PATH = "banco_congresso_ficticio.xlsx"
AVALIACOES_PATH = "avaliacoes.csv"

# Critérios e notas aceitáveis
CRITERIOS = [
    "Qualidade do material apresentado",
    "Postura na apresentação",
    "Domínio teórico do conteúdo",
    "Relevância do projeto",
    "Uso adequado do tempo",
]
VALORES_NOTA = [0, 0.5, 1.0, 1.5, 2.0]  # Notas permitidas

# ===================== FUNÇÕES AUXILIARES =====================

def carregar_banco():
    """
    Carrega o banco de dados dos trabalhos.
    Tenta primeiro CSV; se não existir, tenta XLSX.
    """
    if os.path.exists(CSV_PATH):
        return pd.read_csv(CSV_PATH)
    if os.path.exists(XLSX_PATH):
        return pd.read_excel(XLSX_PATH)
    raise FileNotFoundError(
        f"Banco não encontrado. Coloque {CSV_PATH} ou {XLSX_PATH} na mesma pasta."
    )

def garantir_arquivo_avaliacoes():
    """
    Garante que o arquivo de avaliações exista com o cabeçalho correto.
    Se não existir, cria um CSV vazio com as colunas padronizadas.
    """
    if not os.path.exists(AVALIACOES_PATH):
        colunas = [
            "numero_protocolo",
            "titulo",
            "curso",
            "categoria",
            "modalidade",
            "avaliador_nome",
            "avaliador_email",
        ]
        # Acrescenta uma coluna por critério (ex.: nota_1_qualidade_do_material_apresentado)
        colunas.extend([f"nota_{i+1}_{slug(c)}" for i, c in enumerate(CRITERIOS)])
        colunas.extend(["soma_notas", "timestamp"])
        pd.DataFrame(columns=colunas).to_csv(AVALIACOES_PATH, index=False, encoding="utf-8")

def slug(texto: str) -> str:
    """Converte um texto em identificador curto para nome de coluna."""
    return re.sub(r"[^a-z0-9]+", "", texto.lower()).strip("")

def validar_email(email: str) -> bool:
    """
    Validação simples de e-mail.
    Ajuste a regex para restringir a domínios específicos (ex.: .edu.br) se quiser.
    """
    padrao = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
    return re.match(padrao, email) is not None

def contar_avaliacoes_existentes(numero_protocolo: str) -> int:
    """
    Conta quantas avaliações já existem para determinado protocolo (máximo = 2).
    """
    if not os.path.exists(AVALIACOES_PATH):
        return 0
    df = pd.read_csv(AVALIACOES_PATH)
    if "numero_protocolo" not in df.columns:
        return 0
    return int((df["numero_protocolo"].astype(str) == str(numero_protocolo)).sum())

def selecionar_registro_por_protocolo(banco: pd.DataFrame, numero: str) -> dict:
    """
    Filtra o DataFrame do banco pela coluna 'Número de protocolo' e retorna um dict da linha.
    Levanta ValueError se não encontrar.
    """
    col_proto = "Número de protocolo"  # ajuste se seu CSV tiver outro nome de coluna
    if col_proto not in banco.columns:
        raise KeyError(f"Coluna '{col_proto}' não encontrada no banco.")
    sel = banco[banco[col_proto].astype(str) == str(numero)]
    if sel.empty:
        raise ValueError(f"Nenhum trabalho com protocolo {numero}.")
    return sel.iloc[0].to_dict()

def imprimir_resumo_trabalho(registro: dict):
    """Imprime no console os principais dados do trabalho para confirmação."""
    print("\n=== Dados do trabalho ===")
    print(f"Número de protocolo: {registro.get('Número de protocolo', '')}")
    print(f"Título: {registro.get('Título', '')}")
    print(f"Autores: {registro.get('Autores (nome — curso/universidade)', '')}")
    print(f"Categoria: {registro.get('Categoria', '')}")
    print(f"Curso: {registro.get('Curso', '')}")
    print(f"Modalidade: {registro.get('Modalidade', '')}\n")

def ler_nota(criterio: str) -> float:
    """
    Lê uma nota do console garantindo que esteja entre os valores permitidos.
    Aceita vírgula ou ponto como separador decimal.
    """
    while True:
        entrada = input(f"Nota para '{criterio}' (0, 0,5, 1, 1,5, 2): ").strip()
        entrada = entrada.replace(",", ".")
        try:
            valor = float(entrada)
        except ValueError:
            print("Entrada inválida. Tente novamente.")
            continue
        if valor in VALORES_NOTA:
            return valor
        print("Valor fora dos permitidos. Tente novamente.")

def salvar_avaliacao(registro_trabalho: dict, nome: str, email: str, notas: list[float]):
    """
    Acrescenta uma avaliação no arquivo AVALIACOES_PATH.
    """
    garantir_arquivo_avaliacoes()

    linha = {
        "numero_protocolo": str(registro_trabalho.get("Número de protocolo", "")),
        "titulo": str(registro_trabalho.get("Título", "")),
        "curso": str(registro_trabalho.get("Curso", "")),
        "categoria": str(registro_trabalho.get("Categoria", "")),
        "modalidade": str(registro_trabalho.get("Modalidade", "")),
        "avaliador_nome": nome.strip(),
        "avaliador_email": email.strip(),
    }

    # Notas individuais por critério
    for i, crit in enumerate(CRITERIOS):
        linha[f"nota_{i+1}_{slug(crit)}"] = float(notas[i])

    # Soma e timestamp
    soma = float(sum(notas))
    linha["soma_notas"] = soma
    linha["timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Lê, concatena e salva
    df = pd.read_csv(AVALIACOES_PATH)
    df = pd.concat([df, pd.DataFrame([linha])], ignore_index=True)
    df.to_csv(AVALIACOES_PATH, index=False, encoding="utf-8")

# ===================== FLUXO PRINCIPAL (CLI) =====================

def main():
    print("=== Avaliador de Trabalhos — Modo Console ===")
    print("Para sair a qualquer momento, pressione Ctrl+C.\n")

    # Carrega banco e garante arquivo de avaliações
    try:
        banco = carregar_banco()
        garantir_arquivo_avaliacoes()
    except Exception as e:
        print(f"ERRO: {e}")
        return

    while True:
        try:
            # 1) Receber número de protocolo
            numero = input("Digite o número de protocolo: ").strip()
            if not numero:
                print("Informe um número de protocolo válido.\n")
                continue

            # 2) Impedir > 2 avaliações
            ja = contar_avaliacoes_existentes(numero)
            if ja >= 2:
                print(f"Limite atingido: o protocolo {numero} já possui 2 avaliações.\n")
                continue

            # 3) Buscar registro e mostrar ao avaliador
            try:
                registro = selecionar_registro_por_protocolo(banco, numero)
            except Exception as e:
                print(f"{e}\n")
                continue

            imprimir_resumo_trabalho(registro)

            # 4) Confirmação mais tolerante:
            #    - Enter vazio = SIM
            #    - s / sim = SIM
            #    - n / nao / não = NÃO (volta pro início)
            while True:
                confirma = input(
                    "Os dados estão corretos? (Enter/s = sim, n = não): "
                ).strip().lower()

                if confirma in ("", "s", "sim"):
                    # Prossegue para as notas
                    break
                elif confirma in ("n", "nao", "não"):
                    print("Voltando ao início.\n")
                    # Volta para o começo do while True (pede novo protocolo)
                    numero = None
                    break
                else:
                    print("Resposta inválida. Digite 's' para sim, 'n' para não, ou apenas Enter para sim.")

            # Se a confirmação foi negativa, volta ao início do laço principal
            if confirma in ("n", "nao", "não"):
                continue

            # 5) Coleta das notas (um critério por vez)
            print("\nAtribuição de notas (permitidas: 0; 0,5; 1; 1,5; 2):")
            notas = []
            for crit in CRITERIOS:
                nota = ler_nota(crit)
                notas.append(nota)

            # 6) Apresenta soma e pede confirmação com nome e e-mail
            soma = sum(notas)
            print("\n=== Resumo da avaliação ===")
            for i, (crit, nota) in enumerate(zip(CRITERIOS, notas), start=1):
                print(f"{i}) {crit}: {nota:.1f}")
            print(f"Soma das notas: {soma:.1f}\n")

            nome = input("Digite seu NOME COMPLETO para confirmar: ").strip()
            if len(nome.split()) < 2:
                print("Nome completo inválido. Operação cancelada.\n")
                continue

            email = input("Digite seu E-MAIL INSTITUCIONAL: ").strip()
            if not validar_email(email):
                print("E-mail inválido. Operação cancelada.\n")
                continue

            # 7) Rechecar limite de 2 avaliações (caso outra pessoa tenha avaliado no meio tempo)
            ja = contar_avaliacoes_existentes(numero)
            if ja >= 2:
                print(f"Limite atingido: o protocolo {numero} acabou de atingir 2 avaliações.\n")
                continue

            # 8) Salvar avaliação
            salvar_avaliacao(registro, nome, email, notas)
            print("Avaliação registrada com sucesso!\n")

            # 9) Pergunta se deseja avaliar outro
            outro = input("Deseja avaliar outro trabalho? (s/n): ").strip().lower()
            if outro not in ("s", "sim"):
                print("Encerrando. Obrigado!")
                break

        except KeyboardInterrupt:
            print("\nEncerrado pelo usuário. Até mais!")
            break
        except Exception as e:
            print(f"Erro inesperado: {e}\n")

if __name__ == "__main__":
    main()