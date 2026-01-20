# IMPORTS
import pdfplumber
import re
import json
import pandas as pd
from pathlib import Path

# def salvar_txt_debug(texto: str, caminho_txt: Path) -> None:
#     caminho_txt.write_text(texto, encoding="utf-8")
#     print(f"📝 TXT gerado para inspeção: {caminho_txt}")

# 1. PDF -> TEXTO
def pdf_para_texto(caminho_pdf: Path) -> str:
    textos = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            conteudo = pagina.extract_text()
            if conteudo:
                textos.append(conteudo)
    return " ".join(textos)

def pdf_para_texto_bruto(caminho_pdf: Path) -> str:
    textos = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            conteudo = pagina.extract_text()
            if conteudo:
                textos.append(conteudo)
    return "\n".join(textos)

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
        r'(?:[\d,]+\s*)?' 
        r'.*?'
        r'Justificativa\s+para\s+conceito\s+(?P<conceito>\d|NSA)\s*:'
        r'(?P<justificativa>.*?)'
        r'(?=\s+\d+\.\d+\.\s+[A-Z]|\s+Dimensão\s+\d+|\Z)', # O segredo está no [A-Z]
        re.IGNORECASE | re.DOTALL
    )

    resultado = {}

    for m in padrao.finditer(texto):
        titulo = m.group("titulo").strip()
        nota = m.group("conceito").strip()
        justificativa = m.group("justificativa")

        # limpeza
        padrao_lixo = (
            r'(?:\d{1,2}/\d{1,2}/\d{2,4},?\s*\d{1,2}:\d{1,2}(?::\d{1,2})?\s*(?:AM|PM)?)' # Datas e Horas flexíveis
            r'|(?:\d+\s+of\s+\d+)' # Paginação "1 of 10"
            r'|Firefox|about:blank' # Marcas do navegador
        )

        justificativa = re.sub(
            padrao_lixo,
            ' ',
            justificativa,
            flags=re.IGNORECASE
        )
        
        justificativa = re.sub(r'\s+', ' ', justificativa).strip()

        resultado[titulo] = {
            "Nota": nota,
            "Justificativa": justificativa
        }

    return resultado

# def extrair_docentes(texto: str, ato_regulatorio: str, info_curso: dict) -> list:
#     docentes = []

#     bloco = re.search(
#         r'\bDOCENTES\b(.*?)(CATEGORIAS AVALIADAS)',
#         texto,
#         flags=re.IGNORECASE | re.DOTALL
#     )

#     if not bloco:
#         return docentes

#     texto_docentes = bloco.group(1)

#     # remove cabeçalhos da tabela
#     texto_docentes = re.sub(
#         r'Nome do Docente|Titulação|Regime|Vínculo|Tempo de vínculo.*?meses\)',
#         '',
#         texto_docentes,
#         flags=re.IGNORECASE | re.DOTALL
#     )

#     # cada docente termina em Mês(es)
#     registros = re.findall(
#         r'.*?\d+\s*Mês\(es\)',
#         texto_docentes,
#         flags=re.IGNORECASE | re.DOTALL
#     )

#     for reg in registros:
#         reg = re.sub(r'\s+', ' ', reg).strip()

#         m = re.search(
#             r'(?P<nome>.*?)\s+'
#             r'(?P<titulacao>Doutorado|Mestrado|Especialização)\s+'
#             r'(?P<regime>Integral|Parcial|Horista)\s+'
#             r'(?P<vinculo>CLT|Outro)\s+'
#             r'(?P<meses>\d+)\s*Mês\(es\)',
#             reg,
#             flags=re.IGNORECASE
#         )

#         if not m:
#             continue

#         nome = re.sub(r'\s+', ' ', m.group("nome")).strip()

#         docente = {
#             "Nome do Docente": nome,
#             "Titulação": m.group("titulacao").capitalize(),
#             "Regime de Trabalho": m.group("regime").capitalize(),
#             "Vínculo Empregatício": m.group("vinculo").upper(),
#             "Curso": info_curso["Nome"],
#             "Campus": info_curso["Campus"],
#             "Ano da avaliação": info_curso["Ano da avaliação"],
#             "Ato Regulatório": info_curso["Ato Regulatório"]
#         }

#         if ato_regulatorio.lower() == "reconhecimento":
#             docente["Tempo de vínculo (meses)"] = m.group("meses")

#         docentes.append(docente)

#     return docentes

def extrair_protocolo(texto: str) -> str:
    m = re.search(
        r'Protocolo\s*:\s*(\d+)',
        texto,
        flags=re.IGNORECASE
    )
    return m.group(1) if m else ""

def protocolo_ja_processado(
    protocolo: str,
    pasta_excel: Path
) -> bool:
    if not protocolo:
        return False

    for arquivo in pasta_excel.glob("*.xlsx"):
        if protocolo in arquivo.stem:
            return True

    return False

# 5. INFORMAÇÕES DO CURSO
def extrair_informacoes_curso(texto: str) -> dict:
    info = {
        "Nome": "",
        "Campus": "",
        "Ano da avaliação": "",
        "Ato Regulatório": "",
        "CONCEITO FINAL CONTÍNUO": "",
        "CONCEITO FINAL FAIXA": ""
    }

    # =========================
    # NOME DO CURSO
    # =========================
    m = re.search(
        r'Curso\(s\).*?avaliado\(s\)\s*:\s*(.*?)\s*Informações da comissão',
        texto,
        flags=re.IGNORECASE | re.DOTALL
    )
    if m:
        nome = m.group(1)
        nome = re.sub(r'\s+', ' ', nome).strip()

        # remove apenas " I", " II", " III" no final
        nome = re.sub(r'\s+\bI{1,3}\b$', '', nome)

        info["Nome"] = nome


    # =========================
    # CAMPUS (MODALIDADE + NOME)
    # =========================

    inicio_texto = texto[:1500]

    # Modalidade
    if re.search(r'\(EAD\)|\(EaD\)', inicio_texto):
        info["Campus"] = "EAD"
    else:
        info["Campus"] = "Presencial"

    # Nome do campus físico
    m = re.search(
        r'Endereço da IES\s*:?\s*\d+\s*-\s*(UNASP campus [A-Za-zÀ-ÿ\s]+?)\s*-',
        texto,
        flags=re.IGNORECASE
    )

    if m:
        campus_fisico = m.group(1).strip()
        info["Campus"] = campus_fisico

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

    # Ato Regulatório
    m = re.search(
        r'Ato Regulatório\s*:\s*(Reconhecimento|Autorização)',
        texto,
        flags=re.IGNORECASE
    )
    if m:
        info["Ato Regulatório"] = m.group(1).capitalize()

    return info

# 6. ESTRUTURA BASE
def criar_estrutura_base() -> dict:
    return {
        "Informações curso": {
            "Nome": "",
            "Campus": "",
            "Ano da avaliação": "",
            "Ato Regulatório": "",
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
                    "Ato Regulatório": info["Ato Regulatório"],
                    "Conceito Final Contínuo": info["CONCEITO FINAL CONTÍNUO"],
                    "Conceito Final Faixa": info["CONCEITO FINAL FAIXA"],
                    "Dimensão": dimensao,
                    "Item": titulo,
                    "Nota": dados["Nota"],
                    "Justificativa": dados["Justificativa"]
                })

    pd.DataFrame(linhas).to_excel(caminho_excel, index=False, engine="openpyxl")
    print(f"✅ Excel gerado: {caminho_excel}")

# def docentes_para_excel(
#     docentes: list,
#     caminho_excel: Path
# ) -> None:
#     if not docentes:
#         print("⚠️ Nenhum docente encontrado.")
#         return

#     df = pd.DataFrame(docentes)
#     df.to_excel(caminho_excel, index=False, engine="openpyxl")
#     print(f"✅ Excel de docentes gerado: {caminho_excel}")


# 10. PROCESSAR PASTA DE PDFs
def processar_pasta_pdfs(
    pasta_pdfs: Path,
    pasta_saida_json: Path,
    pasta_saida_excel: Path
) -> None:
    pasta_saida_json.mkdir(parents=True, exist_ok=True)
    pasta_saida_excel.mkdir(parents=True, exist_ok=True)

    pdfs = list(pasta_pdfs.glob("*.pdf"))

    if not pdfs:
        print("⚠️ Nenhum PDF encontrado na pasta.")
        return

    for pdf in pdfs:
        try:
            print(f"📄 Analisando: {pdf.name}")

            # 🔥 leitura mínima só para pegar protocolo
            texto_bruto = pdf_para_texto_bruto(pdf)
            protocolo = extrair_protocolo(texto_bruto)

            if protocolo_ja_processado(protocolo, pasta_saida_excel):
                print(f"⏭️ Protocolo {protocolo} já processado. Pulando...")
                continue

            print(f"📄 Processando: {pdf.name}")

            texto_bruto = pdf_para_texto_bruto(pdf)
            protocolo = extrair_protocolo(texto_bruto)

            if not protocolo:
                print(f"⚠️ Protocolo não encontrado em {pdf.name}")
                continue

            excel_saida = pasta_saida_excel / f"{protocolo}.xlsx"
            excel_docentes = pasta_saida_excel / f"{protocolo}_docentes.xlsx"
            json_saida = pasta_saida_json / f"{protocolo}.json"

            # se já existe, pula
            if excel_saida.exists() and excel_docentes.exists():
                print(f"⏭️ Protocolo {protocolo} já processado. Pulando.")
                continue

            dados = pdf_para_json(pdf, json_saida)
            json_para_excel(dados, excel_saida)

            ato = dados["Informações curso"]["Ato Regulatório"]

            texto_bruto = pdf_para_texto_bruto(pdf)
            # txt_debug = pasta_saida_excel / f"{nome_base}_debug.txt"
            # salvar_txt_debug(texto_bruto, txt_debug)


            info_curso = dados["Informações curso"]

            # docentes = extrair_docentes(
            #     texto_bruto,
            #     ato,
            #     info_curso
            # )

            # excel_docentes = pasta_saida_excel / f"{nome_base}_docentes.xlsx"
            # docentes_para_excel(docentes, excel_docentes)

        except Exception as e:
            print(f"❌ Erro ao processar {pdf.name}: {e}")

# 11. EXECUÇÃO
if __name__ == "__main__":
    pasta_pdfs = Path("Testes/pdfs")
    pasta_json = Path("Testes/json")
    pasta_excel = Path("Testes/excel")

    processar_pasta_pdfs(
        pasta_pdfs,
        pasta_json,
        pasta_excel
    )