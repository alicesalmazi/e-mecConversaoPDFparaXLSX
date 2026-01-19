# IMPORTS
import pdfplumber
import re
import json
import pandas as pd
from pathlib import Path

# 1. PDF -> TEXTO
def pdf_para_texto(caminho_pdf: Path) -> str:
    textos = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            conteudo = pagina.extract_text()
            if conteudo:
                textos.append(conteudo)
    return " ".join(textos)

# 2. LIMPEZA DO TEXTO
def limpar_texto(texto: str) -> str:
    padroes_remover = [
        r'about:blank',
        r'\n\s*\d+\s*\n',
        r'\n\s*NSA\s*\n',
        r'Firefox.*?\d{2}:\d{2}:\d{2}',
        r'Firefox.*?\d{2}/\d{2}/\d{4}',
        r'Data\s*\d{2}/\d{2}/\d{4}',
        r'Hora\s*\d{2}:\d{2}(:\d{2})?',
        r'Página\s*\d+\s*de\s*\d+',
        r'\bblank\b',
    ]

    for p in padroes_remover:
        texto = re.sub(p, ' ', texto, flags=re.IGNORECASE)

    texto = texto.replace("\n", " ").replace("\r", " ")
    texto = re.sub(r'\s+', ' ', texto)
    texto = texto.replace(" :", ":").replace(" .", ".")

    return texto.strip()

# 3. EXTRAIR TODOS OS ITENS (MESMO SEM JUSTIFICATIVA)
def extrair_todos_itens(texto: str) -> dict:
    padrao = re.compile(
        r'(\d+\.\d+)\.\s+([^.]+?\.)',
        re.IGNORECASE
    )

    itens = {}

    for num, titulo in padrao.findall(texto):
        chave = f"{num}. {titulo.strip()}"
        itens[chave] = {
            "Nota": "",
            "Justificativa": ""
        }

    return itens

# 4. EXTRAÇÃO DE NOTA + JUSTIFICATIVA (QUANDO EXISTIR)
def extrair_notas_justificativas(texto: str) -> dict:
    padrao = re.compile(
        r'(?P<titulo>\d+\.\d+\.\s+[^.]+?\.)\s*'
        r'(?:\d|NSA)?\s*'
        r'Justificativa\s+para\s+conceito\s+(?P<conceito>\d|NSA)\s*:'
        r'(?P<justificativa>.*?)(?=\s\d+\.\d+\.|\Z)',
        re.IGNORECASE | re.DOTALL
    )

    resultado = {}

    for m in padrao.finditer(texto):
        titulo = m.group("titulo").strip()
        nota = m.group("conceito").strip()
        justificativa = m.group("justificativa")

        # limpeza pesada
        justificativa = re.sub(r'\b\d+\s+of\s+\d+\b', ' ', justificativa, flags=re.IGNORECASE)
        justificativa = re.sub(r'\d{2}/\d{2}/\d{4},?\s*\d{2}:\d{2}', ' ', justificativa)
        justificativa = re.sub(r'Firefox', ' ', justificativa, flags=re.IGNORECASE)
        justificativa = re.sub(r'Dimensão\s+\d+\s*:.*$', '', justificativa, flags=re.IGNORECASE)

        justificativa = re.sub(r'\s+', ' ', justificativa).strip()

        if len(justificativa) < 10:
            justificativa = ""

        resultado[titulo] = {
            "Nota": nota,
            "Justificativa": justificativa
        }

    return resultado

# 5. INFORMAÇÕES DO CURSO
def extrair_informacoes_curso(texto: str) -> dict:
    info = {
        "Nome": "",
        "Campus": "",
        "Ano da avaliação": "",
        "CONCEITO FINAL CONTÍNUO": "",
        "CONCEITO FINAL FAIXA": ""
    }

    # Nome do curso (curto e limpo)
    m = re.search(
        r'Curso\(s\).*?avaliado\(s\)\s*:\s*([A-ZÀ-Ú\s]{5,80})',
        texto
    )
    if m:
        info["Nome"] = m.group(1).strip()

    # Campus
    if re.search(r'Ato Regulatório.*EAD', texto, flags=re.IGNORECASE):
        info["Campus"] = "EAD"
    else:
        m = re.search(
            r'Endereço da IES\s*:?\s*\d+\s*-\s*(.+?)\s*-',
            texto,
            flags=re.IGNORECASE
        )
        if m:
            info["Campus"] = m.group(1).strip()

    # Ano da avaliação
    m = re.search(
        r'Data\s+de\s+\d{2}/\d{2}/(\d{4})',
        texto,
        flags=re.IGNORECASE
    )
    if m:
        info["Ano da avaliação"] = m.group(1)

    # Conceito final
    m = re.search(
        r'CONCEITO FINAL CONT[IÍ]NUO\s*CONCEITO FINAL FAIXA\s*([\d,]+)\s*(\d)',
        texto,
        flags=re.IGNORECASE
    )
    if m:
        info["CONCEITO FINAL CONTÍNUO"] = m.group(1)
        info["CONCEITO FINAL FAIXA"] = m.group(2)

    return info

# 6. ESTRUTURA BASE
def criar_estrutura_base() -> dict:
    return {
        "Informações curso": {
            "Nome": "",
            "Campus": "",
            "Ano da avaliação": "",
            "CONCEITO FINAL CONTÍNUO": "",
            "CONCEITO FINAL FAIXA": ""
        },
        "Dimensões": {
            "ORGANIZAÇÃO DIDÁTICO-PEDAGÓGICA": [],
            "CORPO DOCENTE E TUTORIAL": [],
            "INFRAESTRUTURA": []
        }
    }

# 7. INSERÇÃO NAS DIMENSÕES
def inserir_dados(estrutura: dict, itens: dict) -> None:
    for titulo, dados in sorted(itens.items()):
        if titulo.startswith("1."):
            estrutura["Dimensões"]["ORGANIZAÇÃO DIDÁTICO-PEDAGÓGICA"].append({titulo: dados})
        elif titulo.startswith("2."):
            estrutura["Dimensões"]["CORPO DOCENTE E TUTORIAL"].append({titulo: dados})
        elif titulo.startswith("3."):
            estrutura["Dimensões"]["INFRAESTRUTURA"].append({titulo: dados})

# 8. PIPELINE PDF -> JSON
def pdf_para_json(pdf_path: Path, json_path: Path) -> dict:
    texto = limpar_texto(pdf_para_texto(pdf_path))

    estrutura = criar_estrutura_base()
    estrutura["Informações curso"].update(extrair_informacoes_curso(texto))

    todos_itens = extrair_todos_itens(texto)
    itens_avaliados = extrair_notas_justificativas(texto)

    # sobrescreve quando existir nota/justificativa
    for k, v in itens_avaliados.items():
        todos_itens[k] = v

    # 🔥 REGRA FINAL: sem justificativa → Nota 6 + NSA
    for item, dados in todos_itens.items():
        if not dados["Justificativa"]:
            dados["Nota"] = "6"
            dados["Justificativa"] = "NSA. Não se aplica."

    inserir_dados(estrutura, todos_itens)

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(estrutura, f, ensure_ascii=False, indent=4)

    print(f"✅ JSON gerado: {json_path}")
    return estrutura

# 9. JSON -> EXCEL
def json_para_excel(json_dados: dict, caminho_excel: Path) -> None:
    linhas = []
    info = json_dados["Informações curso"]

    for dimensao, itens in json_dados["Dimensões"].items():
        for item in itens:
            for titulo, dados in item.items():
                linhas.append({
                    "Curso": info["Nome"],
                    "Campus": info["Campus"],
                    "Ano da avaliação": info["Ano da avaliação"],
                    "Conceito Final Contínuo": info["CONCEITO FINAL CONTÍNUO"],
                    "Conceito Final Faixa": info["CONCEITO FINAL FAIXA"],
                    "Dimensão": dimensao,
                    "Item": titulo,
                    "Nota": dados["Nota"],
                    "Justificativa": dados["Justificativa"]
                })

    pd.DataFrame(linhas).to_excel(caminho_excel, index=False, engine="openpyxl")
    print(f"✅ Excel gerado: {caminho_excel}")

# 10. EXECUÇÃO
if __name__ == "__main__":
    pdf_entrada = Path("Testes//teste.pdf")
    json_saida = Path("Testes//avaliacao.json")
    excel_saida = Path("Testes//avaliacao.xlsx")

    dados = pdf_para_json(pdf_entrada, json_saida)
    json_para_excel(dados, excel_saida)