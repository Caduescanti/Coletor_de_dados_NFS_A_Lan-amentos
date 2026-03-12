
import os
import re
import json
import pandas as pd
import shutil
import tempfile
import time
import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from google import genai


CHAVE_API = ""
client = genai.Client(api_key=CHAVE_API)
MODELO = 'gemini-2.5-flash' 

PASTA_RAIZ = r'\\files\VPFinan\Seguranca\Lançamentos\NOTAS ANO 2026\2-FEVEREIRO'
DATA_VENCIMENTO = '25/02/2026'

CAMINHO_POPPLER = r'C:\Users\143644\Downloads\Release-25.12.0-0\poppler-25.12.0\Library\bin'
CAMINHO_TESSERACT = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

if os.path.exists(CAMINHO_TESSERACT):
    pytesseract.pytesseract.tesseract_cmd = CAMINHO_TESSERACT


def limpar_valor_dinheiro(valor_str):
    if not valor_str: return "0,00"
    v = str(valor_str).upper().replace('R$', '').replace(' ', '').strip()
    if '.' in v and ',' not in v:
        parts = v.split('.')
        if len(parts[-1]) == 2: v = "".join(parts[:-1]) + "," + parts[-1]
    if ',' in v: v = v.replace('.', '')
    return re.sub(r'[^\d,]', '', v)

def limpar_valor_float(valor_str):
    """Garante que qualquer lixo (R$, espaços) seja removido antes da matemática"""
    if not valor_str: return 0.0
    v = str(valor_str).upper().replace('R$', '').replace(' ', '').strip()
    try: return float(v.replace('.', '').replace(',', '.'))
    except: return 0.0

def ler_pdf_como_imagem(caminho_pdf):
    try:
        imagens = convert_from_path(caminho_pdf, first_page=1, last_page=1, dpi=300, poppler_path=CAMINHO_POPPLER)
        if imagens: return pytesseract.image_to_string(imagens[0], lang='por', config='--psm 6')
    except: return ""
    return ""

def ler_texto_nativo(caminho_pdf):
    texto = ""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for p in pdf.pages: texto += p.extract_text() or ""
    except: pass
    
    if len(texto.strip()) < 50 or not re.search(r'\d+', texto):
        texto_ocr = ler_pdf_como_imagem(caminho_pdf)
        if texto_ocr: texto = texto_ocr
    return texto

def definir_bandeira_global(texto_pdf, nome_arquivo):
    t_upper = texto_pdf.upper()
    n_upper = nome_arquivo.upper()
    if "04.899.316" in t_upper or "04899316" in t_upper or "IMIFARMA" in t_upper or "EXTRAFARMA" in t_upper or "EXTRAFARMA" in n_upper or "EF" in n_upper:
        return "EF"
    if "06.626.253" in t_upper or "06626253" in t_upper or "PAGUE MENOS" in t_upper or "PAGUEMENOS" in t_upper or "PGMN" in n_upper:
        return "PGMN"
    return "OUTROS"

def extrair_numero_do_nome(nome_arquivo):
    if "NOTAFISCAL_V" in nome_arquivo.upper(): return "NÃO ACHOU"
    m = re.search(r'(?:NF|NFS-E|NFS|NOTA)\s*[-_]?\s*(\d+)', nome_arquivo, re.IGNORECASE)
    if m: return m.group(1)
    m_puro = re.search(r'^(\d+)', nome_arquivo)
    if m_puro: return m_puro.group(1)
    return "NÃO ACHOU"

# ====================================================================
# --- 3. MOTORES DE EXTRAÇÃO LOCAIS ---
# ====================================================================
def motor_sekron(caminho_pdf):
    nome_arq = os.path.basename(caminho_pdf)
    texto_flat = re.sub(r'\s+', ' ', ler_texto_nativo(caminho_pdf).replace('\n', ' ').replace('\r', ' ')).upper()
    numero = "NÃO ACHOU"
    m_nfse = re.search(r'N[UÚ]MERO\s*DA\s*NFS.*?(\d{1,10})\b', texto_flat)
    if m_nfse: numero = m_nfse.group(1)
    valor = "0,00"
    m_serv = re.search(r'VALOR\s*DO\s*SERVIÇO.*?R\$\s*([\d\.,]+)', texto_flat)
    if m_serv: valor = limpar_valor_dinheiro(m_serv.group(1))
    bandeira = definir_bandeira_global(texto_flat, nome_arq)
    return {"Número NF": numero, "Valor": valor, "BANDEIRA": bandeira, "FORNECEDOR": "SEKRON", "Assunto": f"serviço de vigilancia eletronica NF {numero} R$ {valor} - SEKRON"}

def motor_padrao_uf(caminho_pdf, fornecedor_nome):
    nome_arq = os.path.basename(caminho_pdf)
    texto_flat = re.sub(r'\s+', ' ', ler_texto_nativo(caminho_pdf).replace('\n', ' ').replace('\r', ' ')).upper()
    numero = "NÃO ACHOU"
    m_nf = re.search(r'N[UÚ]MERO\s*DA\s*NFS.*?(\d{1,10})\b', texto_flat)
    if not m_nf: m_nf = re.search(r'NOTA\s*FISCAL.*?N[UÚ]MERO.*?(\d{1,10})\b', texto_flat)
    if m_nf: numero = m_nf.group(1)
    valor = "0,00"
    m_val = re.search(r'VALOR\s*DO\s*SERVIÇO.*?R\$\s*([\d\.,]+)', texto_flat)
    if m_val: valor = limpar_valor_dinheiro(m_val.group(1))
    bandeira = definir_bandeira_global(texto_flat, nome_arq)
    return {"Número NF": numero, "Valor": valor, "BANDEIRA": bandeira, "FORNECEDOR": fornecedor_nome, "Assunto": f"serviço NF {numero} R$ {valor} - {fornecedor_nome}"}

def motor_ogenio(caminho_pdf):
    nome_arq = os.path.basename(caminho_pdf)
    texto_limpo = re.sub(r'\s+', ' ', ler_texto_nativo(caminho_pdf))
    numero = "NÃO ACHOU"
    m_num = re.search(r'Número\s*da\s*NFS-e\s*(\d{1,10})\b', texto_limpo, re.IGNORECASE)
    if not m_num: m_num = re.search(r'Número\s*(?:da\s*)?Nota\s*Fiscal[:\s]*(\d{1,10})\b', texto_limpo, re.IGNORECASE)
    if m_num: numero = m_num.group(1)
    else: numero = extrair_numero_do_nome(nome_arq)
    valor = "0,00"
    m_val = re.search(r'VALOR\s*(?:DO\s*|DOS\s*)?SERVIÇO(?:S)?.*?(\d{1,3}(?:\.\d{3})*,\d{2})', texto_limpo, re.IGNORECASE)
    if m_val: valor = m_val.group(1)
    else:
        m_val2 = re.search(r'Valor\s*Total\s*da\s*NFS-e.*?(\d{1,3}(?:\.\d{3})*,\d{2})', texto_limpo, re.IGNORECASE)
        if m_val2: valor = m_val2.group(1)
    bandeira = definir_bandeira_global(texto_limpo, nome_arq)
    return {"Número NF": numero, "Valor": valor, "BANDEIRA": bandeira, "FORNECEDOR": "OGENIO", "Assunto": f"serviço NF {numero} R$ {valor} - OGENIO"}

def motor_grenke(caminho_pdf):
    nome_arq = os.path.basename(caminho_pdf)
    texto = ler_texto_nativo(caminho_pdf)
    fatura = "NÃO ACHOU"
    m_fat = re.search(r'Fatura/Recibo\s*nº\s*([\d/]+)', texto, re.IGNORECASE)
    if m_fat: fatura = m_fat.group(1).strip()
    montante = "0,00"
    m_val = re.search(r'Montante\s*\n.*?R\$\s*([\d\.,]+)', texto, re.IGNORECASE | re.DOTALL)
    if not m_val: m_val = re.search(r'Total\s*Documento.*?R\$\s*([\d\.,]+)', texto.replace('\n', ' '), re.IGNORECASE)
    if m_val: montante = m_val.group(1).strip()
    return {"Número NF": fatura, "Valor": montante, "BANDEIRA": "PGMN", "FORNECEDOR": "GRENKE", "Assunto": f"Locação Grenke - Fatura {fatura} - R$ {montante}"}

def motor_life_khronos(caminho_pdf, fornecedor_nome):
    nome_arq = os.path.basename(caminho_pdf)
    texto = " ".join(ler_texto_nativo(caminho_pdf).split())
    numero = "NÃO ACHOU"
    m_num = re.search(r'N(?:úmero|º)\s*(?:da\s*)?N(?:ota|FS-e).*?(\d{1,10})\b', texto, re.IGNORECASE)
    if m_num: numero = m_num.group(1)
    valor = "0,00"
    pos = texto.lower().find("total dos serviços")
    if pos != -1:
        m_val = re.search(r'R\$\s*([\d\.,]+)', texto[pos:pos+300])
        if m_val: valor = m_val.group(1)
    bandeira = definir_bandeira_global(texto, nome_arq)
    return {"Número NF": numero, "Valor": valor, "BANDEIRA": bandeira, "FORNECEDOR": fornecedor_nome, "Assunto": f"Nota Fiscal {numero} - R$ {valor} - {fornecedor_nome}"}

def motor_tws_perony(caminho_pdf, fornecedor_nome):
    nome_arq = os.path.basename(caminho_pdf)
    texto = " ".join(ler_texto_nativo(caminho_pdf).split())
    numero = "NÃO ACHOU"
    m_num = re.search(r'NFS-e\s*(\d{1,5})\b', texto, re.IGNORECASE)
    if m_num: numero = m_num.group(1)
    valor = "0,00"
    m_val = re.search(r'Valor do Serviço.*?(\d{1,3}(?:\.\d{3})*,\d{2})', texto, re.IGNORECASE)
    if m_val: valor = m_val.group(1)
    bandeira = definir_bandeira_global(texto, nome_arq)
    return {"Número NF": numero, "Valor": valor, "BANDEIRA": bandeira, "FORNECEDOR": fornecedor_nome, "Assunto": f"Nota Fiscal {numero} - R$ {valor} - {fornecedor_nome}"}

def motor_staff(caminho_pdf):
    nome_arq = os.path.basename(caminho_pdf)
    texto = re.sub(r'\s+', ' ', ler_texto_nativo(caminho_pdf))
    numero = "NÃO ACHOU"
    m_num = re.search(r'N[úu]mero\s*da\s*N(?:FS-e|ota).*?(\d{1,10})\b', texto, re.IGNORECASE)
    if not m_num: m_num = re.search(r'Nº\s*da\s*Nota.*?(\d{1,10})\b', texto, re.IGNORECASE)
    if m_num: numero = str(int(m_num.group(1)))
    else: numero = extrair_numero_do_nome(nome_arq)
    valor = "0,00"
    m_val = re.search(r'VALOR\s*TOTAL\s*DA\s*N(?:FS-e|OTA).*?(\d{1,3}(?:\.\d{3})*,\d{2})', texto, re.IGNORECASE)
    if not m_val: m_val = re.search(r'Valor\s*do\s*Serviço.*?(\d{1,3}(?:\.\d{3})*,\d{2})', texto, re.IGNORECASE)
    if m_val: valor = m_val.group(1)
    else:
        valores = re.findall(r'(?:\b\d{1,3}(?:\.\d{3})*,\d{2}\b)', texto)
        if valores: valor = max(valores, key=limpar_valor_float)

    tipo_servico = "SERVIÇO GERAL"
    check_str = (nome_arq + " " + texto).upper()
    if any(t in check_str for t in ["EXPANSÃO", "EXPANSAO", "OBRA", "REFORMA"]): tipo_servico = "OBRA/EXPANSÃO"
    elif any(t in check_str for t in ["RONDA", "VIGILANCIA", "FISCAL", "FISCAIS"]): tipo_servico = "RONDA/VIGILÂNCIA"
    bandeira = definir_bandeira_global(texto, nome_arq)
    return {"Número NF": numero, "Valor": valor, "BANDEIRA": bandeira, "FORNECEDOR": "STAFF", "Assunto": f"{tipo_servico} NF {numero} R$ {valor} - STAFF"}

def motor_hold_rondespe(caminho_pdf, fornecedor_nome):
    nome_arq = os.path.basename(caminho_pdf)
    texto = " ".join(ler_texto_nativo(caminho_pdf).split())
    numero = "NÃO ACHOU"
    m_num_sp = re.search(r'N[úu]mero\s*da\s*Nota\s*(\d+)', texto, re.IGNORECASE)
    m_num_sorocaba = re.search(r'(\d{1,10})\s*/\s*U', texto)
    if m_num_sp: numero = str(int(m_num_sp.group(1)))
    elif m_num_sorocaba: numero = m_num_sorocaba.group(1)
    else: numero = extrair_numero_do_nome(nome_arq)
    valor = "0,00"
    m_val_sp = re.search(r'VALOR\s*TOTAL\s*DO\s*SERVIÇO\s*=\s*R\$?\s*([\d\.,]+)', texto, re.IGNORECASE)
    m_val_sorocaba = re.search(r'Valor\s*Serviço\s*\(?R\$\)?.*?(\d{1,3}(?:\.\d{3})*,\d{2})', texto, re.IGNORECASE)
    if m_val_sp: valor = m_val_sp.group(1)
    elif m_val_sorocaba: valor = m_val_sorocaba.group(1)
    else:
        valores = re.findall(r'(?:\b\d{1,3}(?:\.\d{3})*,\d{2}\b)', texto)
        if valores: valor = max(valores, key=limpar_valor_float)
    bandeira = definir_bandeira_global(texto, nome_arq)
    return {"Número NF": numero, "Valor": valor, "BANDEIRA": bandeira, "FORNECEDOR": fornecedor_nome, "Assunto": f"Nota Fiscal {numero} - R$ {valor} - {fornecedor_nome}"}

def motor_checklist(caminho_pdf):
    nome_arq = os.path.basename(caminho_pdf)
    texto_f = " ".join(ler_texto_nativo(caminho_pdf).split())
    numero = "NÃO ACHOU"
    m_num = re.search(r'Número\s+da\s+NFS-e\s*(\d{1,10})\b', texto_f, re.IGNORECASE)
    if m_num: numero = m_num.group(1)
    if numero == "NÃO ACHOU": numero = extrair_numero_do_nome(nome_arq)
    valor = "0,00"
    m_val = re.search(r'Valor\s+do\s+Serviço.*?(\d{1,3}(?:\.\d{3})*,\d{2})', texto_f, re.IGNORECASE)
    if not m_val: m_val = re.search(r'VALOR\s+TOTAL\s+DA\s+NOTA.*?(\d{1,3}(?:\.\d{3})*,\d{2})', texto_f, re.IGNORECASE)
    if m_val: valor = m_val.group(1)
    bandeira = definir_bandeira_global(texto_f, nome_arq)
    return {"Número NF": numero, "Valor": valor, "BANDEIRA": bandeira, "FORNECEDOR": "CHECKLIST", "Assunto": f"Nota Fiscal {numero} - R$ {valor} - CHECKLIST"}

def motor_inviolavel(caminho_pdf):
    nome_arq = os.path.basename(caminho_pdf)
    texto_full = ler_texto_nativo(caminho_pdf)
    texto_flat = texto_full.replace('\n', ' ').replace('  ', ' ')
    numero, valor = "NÃO ACHOU", "0,00"
    
    if "DANFE" in texto_full or "NOTA FISCAL ELETRÔNICA" in texto_full:
        m_num = re.search(r'N[º°\.]\s*0*(\d{1,3}(?:\.\d{3})+|\d{4,15})\b', texto_flat)
        if m_num:
            numero = str(int(m_num.group(1).replace('.', '').strip()))
            
        pos_chave = texto_flat.upper().find("VALOR TOTAL DA NOTA")
        if pos_chave != -1:
            valores = re.findall(r'([\d\.]+,\d{2})', texto_flat[pos_chave : pos_chave + 200])
            if valores: valor = max(valores, key=limpar_valor_float)
    else: 
        # 1ª PRIORIDADE (A que você descobriu): Número / Série do RPS
        m_num_rps = re.search(r'N[uú]mero\s*/\s*S[eé]rie\s*do\s*RPS\s*(\d{1,10})', texto_flat, re.IGNORECASE)
        
        # 2ª PRIORIDADE: Padrão Belém (Ex: "1512 / E" ou "1512 / RPS")
        m_num_belem = re.search(r'(\d{3,10})\s*/\s*(?:E|RPS|[A-Z])\b', texto_flat, re.IGNORECASE)
        
        if m_num_rps:
            numero = m_num_rps.group(1).strip()
        elif m_num_belem: 
            numero = m_num_belem.group(1).strip()
        else:
            m_num_gen = re.search(r'N[uú]mero\s*/\s*S[eé]rie.*?(\d{3,10})\b', texto_flat, re.IGNORECASE)
            if m_num_gen: numero = m_num_gen.group(1).strip()
            
        m_val = re.search(r'Valor\s*L[íi]quido\s*da\s*NFSe.*?([\d\.,]+)', texto_flat, re.IGNORECASE)
        if m_val: valor = m_val.group(1).strip()

    # BLOQUEIO ANTI-"04": Se o número tiver 2 dígitos ou menos, é lixo da data. Rejeita!
    if numero != "NÃO ACHOU" and len(numero) <= 2:
        numero = "NÃO ACHOU"

    # Fallback Nome do Arquivo
    if numero == "NÃO ACHOU":
        num_nome = extrair_numero_do_nome(nome_arq)
        if num_nome != "NÃO ACHOU": numero = num_nome

    bandeira = definir_bandeira_global(texto_full, nome_arq)
    return {"Número NF": numero, "Valor": valor, "BANDEIRA": bandeira, "FORNECEDOR": "INVIOLAVEL", "Assunto": f"Nota Fiscal {numero} - R$ {valor} - INVIOLAVEL"}
def motor_hp_ph(caminho_pdf):
    nome_arq = os.path.basename(caminho_pdf)
    texto = re.sub(r'\s+', ' ', ler_texto_nativo(caminho_pdf))
    numero = "NÃO ACHOU"
    m_num = re.search(r'N(?:úmer|u)o\s*da\s*NFS-e\s*[:\s]*(\d{1,10})\b', texto, re.IGNORECASE)
    if not m_num: m_num = re.search(r'NFS-e\s*[:\s]*(\d{1,10})\b', texto, re.IGNORECASE)
    if m_num: numero = str(int(m_num.group(1)))
    else: numero = extrair_numero_do_nome(nome_arq)
    valor = "0,00"
    valores = re.findall(r'(?:\b\d{1,3}(?:\.\d{3})*,\d{2}\b)', texto)
    if valores: valor = max(valores, key=limpar_valor_float)
    bandeira = definir_bandeira_global(texto, nome_arq)
    return {"Número NF": numero, "Valor": valor, "BANDEIRA": bandeira, "FORNECEDOR": "HP-PH", "Assunto": f"Nota Fiscal {numero} - R$ {valor} - HP-PH"}

# ====================================================================
# --- 4. O AGENTE DE IA (A REDE DE SEGURANÇA) ---
# ====================================================================
def ler_nota_com_ia(caminho_pdf):
    caminho_temp = os.path.join(tempfile.gettempdir(), "arquivo_leitura_ia.pdf")
    try: shutil.copy2(caminho_pdf, caminho_temp)
    except: return None
    
    # Aumentado para 4 tentativas
    for tentativa in range(4):
        arquivo_pdf = None
        try:
            arquivo_pdf = client.files.upload(file=caminho_temp)
            prompt = """
            Você é um auditor fiscal especialista em notas do Brasil. Extraia as informações desta Nota Fiscal ESTRITAMENTE em JSON. Se não encontrar algo, retorne null.
            REGRAS:
            1. NÚMERO DA NOTA: Procure por "Número da Nota", "Número da NFS-e", "Número Nota Fiscal:" (ex: 756) ou "Número / Série do RPS". Nunca pegue a Chave de Acesso gigante.
            2. VALOR: Procure as âncoras exatas: "Valor Total do Serviço", "VALOR TOTAL DA NOTA", "Valor dos Serviços R$" ou "VALOR TOTAL".
            3. PRESTADOR: Procure por "Nome Fantasia" ou Razão Social abaixo de "Prestador de Serviços".
            4. TOMADOR: Procure a Razão Social em "Tomador de Serviços" (Ex: EMPREENDIMENTOS PAGUE MENOS S/A ou IMIFARMA...).
            SAÍDA JSON ESPERADA:
            { "numero_nf": "Número da nota", "valor_total": "Valor final", "fornecedor": "Nome do prestador", "tomador": "Nome do tomador" }
            """
            resposta = client.models.generate_content(model=MODELO, contents=[arquivo_pdf, prompt])
            texto_limpo = resposta.text.strip().replace("```json", "").replace("```", "")
            dados_ia = json.loads(texto_limpo)
            
            # Tenta apagar o ficheiro no Google para poupar espaço
            try: client.files.delete(name=arquivo_pdf.name) 
            except: pass
            
            # 🛡️ O SEGREDO ANTI-BLOQUEIO: Pausa de 5 segundos após o sucesso!
            time.sleep(5)
            
            return dados_ia
            
        except Exception as e:
            if arquivo_pdf:
                try: client.files.delete(name=arquivo_pdf.name)
                except: pass
                
            erro_str = str(e).lower()
            # Se o erro for de limite (429, quota, exhausted)
            if '429' in erro_str or 'quota' in erro_str or 'exhausted' in erro_str or '503' in erro_str:
                print(" [Limite da API atingido! Pausando 45s...] ", end="", flush=True)
                time.sleep(45)
                continue 
            elif tentativa < 3:
                # Se for uma pequena falha de internet, pausa 10s e tenta de novo
                time.sleep(10)
                continue
            else:
                return None
    return None
def processar_ia_para_planilha(dados_ia, nome_arquivo, fornecedor_base):
    numero = str(dados_ia.get('numero_nf', 'S/N'))
    if len(numero) > 15: numero = extrair_numero_do_nome(nome_arquivo) 
    
    # Limpa o valor para o padrão Brasileiro antes de ir para o Excel
    valor_bruto = str(dados_ia.get('valor_total', '0,00'))
    valor = limpar_valor_dinheiro(valor_bruto)
    
    forn_ia = str(dados_ia.get('fornecedor', 'DESCONHECIDO')).upper()
    tomador = str(dados_ia.get('tomador', '')).upper()
    bandeira = definir_bandeira_global(tomador, nome_arquivo)
    
    # UNIFICAÇÃO: Força o nome da pasta base para não fatiar o Resumo Financeiro
    forn_final = fornecedor_base if fornecedor_base != "DESCONHECIDO" else forn_ia

    return {
        "Arquivo": nome_arquivo, 
        "BANDEIRA": bandeira, 
        "FORNECEDOR": forn_final,
        "Número NF": numero, 
        "Valor": valor, 
        "Vencimento": DATA_VENCIMENTO,
        "Assunto": f"NF {numero} R$ {valor} - {forn_ia} (IA)"
    }

# ====================================================================
# --- 5. O ROTEADOR INTELIGENTE ---
# ====================================================================
def resultado_valido(dados):
    if not dados: return False
    num = str(dados.get("Número NF", "NÃO ACHOU"))
    if num == "NÃO ACHOU" or len(num) > 15: return False 
    v = str(dados.get("Valor", "0,00"))
    if v in ["0,00", "0.00", "", None]: return False
    return True

def processar_documento_hibrido(caminho_pdf):
    pasta_origem = os.path.dirname(caminho_pdf).upper()
    nome_ficheiro = os.path.basename(caminho_pdf)
    dados_locais = None
    
    fornecedor_base = "DESCONHECIDO"
    if "SEKRON" in pasta_origem: fornecedor_base = "SEKRON"
    elif "CRA" in pasta_origem: fornecedor_base = "CRA"
    elif "HERC" in pasta_origem: fornecedor_base = "HERC"
    elif "SOUZA LIMA" in pasta_origem or "SOUZALIMA" in pasta_origem: fornecedor_base = "SOUZA LIMA"
    elif "OGENIO" in pasta_origem: fornecedor_base = "OGENIO"
    elif "GRENKE" in pasta_origem or "GREENKE" in pasta_origem: fornecedor_base = "GRENKE"
    elif "LIFE" in pasta_origem: fornecedor_base = "LIFE DEFENSE"
    elif "KHONOS" in pasta_origem or "KHRONOS" in pasta_origem: fornecedor_base = "KHRONOS"
    elif "STAFF" in pasta_origem: fornecedor_base = "STAFF"
    elif "TWS" in pasta_origem: fornecedor_base = "TWS"
    elif "PERONY" in pasta_origem: fornecedor_base = "PERONY"
    elif "HOLD" in pasta_origem: fornecedor_base = "HOLD"
    elif "RONDESPE" in pasta_origem: fornecedor_base = "RONDESPE"
    elif "CHECKLIST" in pasta_origem or "CHECLIST" in pasta_origem: fornecedor_base = "CHECKLIST"
    elif "INVIOLAVEL" in pasta_origem or "INVIOLÁVEL" in pasta_origem: fornecedor_base = "INVIOLAVEL"
    elif "HP" in pasta_origem or "PH" in pasta_origem: fornecedor_base = "HP-PH"
    
    try:
        if fornecedor_base == "SEKRON": dados_locais = motor_sekron(caminho_pdf)
        elif fornecedor_base in ["CRA", "HERC", "SOUZA LIMA"]: dados_locais = motor_padrao_uf(caminho_pdf, fornecedor_base)
        elif fornecedor_base == "OGENIO": dados_locais = motor_ogenio(caminho_pdf)
        elif fornecedor_base == "GRENKE": dados_locais = motor_grenke(caminho_pdf)
        elif fornecedor_base in ["LIFE DEFENSE", "KHRONOS"]: dados_locais = motor_life_khronos(caminho_pdf, fornecedor_base)
        elif fornecedor_base == "STAFF": dados_locais = motor_staff(caminho_pdf)
        elif fornecedor_base in ["TWS", "PERONY"]: dados_locais = motor_tws_perony(caminho_pdf, fornecedor_base)
        elif fornecedor_base in ["HOLD", "RONDESPE"]: dados_locais = motor_hold_rondespe(caminho_pdf, fornecedor_base)
        elif fornecedor_base == "CHECKLIST": dados_locais = motor_checklist(caminho_pdf)
        elif fornecedor_base == "INVIOLAVEL": dados_locais = motor_inviolavel(caminho_pdf)
        elif fornecedor_base == "HP-PH": dados_locais = motor_hp_ph(caminho_pdf)
    except: pass
        
    if not resultado_valido(dados_locais):
        print(" [Mudando para Resgate...] ", end="")
        dados_temp = motor_hp_ph(caminho_pdf)
        if resultado_valido(dados_temp):
            dados_locais = dados_temp
            dados_locais["FORNECEDOR"] = fornecedor_base 
            dados_locais["Assunto"] = f"Nota Fiscal {dados_locais['Número NF']} - R$ {dados_locais['Valor']} - {fornecedor_base}"
        else:
            dados_temp2 = motor_padrao_uf(caminho_pdf, fornecedor_base)
            if resultado_valido(dados_temp2): dados_locais = dados_temp2

    if resultado_valido(dados_locais):
        dados_locais["Arquivo"] = nome_ficheiro
        dados_locais["Caminho Completo"] = caminho_pdf 
        if "Vencimento" not in dados_locais: dados_locais["Vencimento"] = DATA_VENCIMENTO
        print(" [Script Local: SUCESSO]")
        return dados_locais
        
    print(" [Tudo Falhou -> Acionando IA...] ", end="", flush=True)
    dados_ia = ler_nota_com_ia(caminho_pdf)
    
    if dados_ia:
        print(" [IA: SUCESSO]")
        # Passa o fornecedor base para não desorganizar a planilha
        res_ia = processar_ia_para_planilha(dados_ia, nome_ficheiro, fornecedor_base)
        res_ia["Caminho Completo"] = caminho_pdf
        return res_ia
        
    print(" [FALHA TOTAL - SALVANDO COMO ERRO]")
    return {
        "Arquivo": nome_ficheiro, "Caminho Completo": caminho_pdf,
        "BANDEIRA": "FALHA", "FORNECEDOR": fornecedor_base,
        "Número NF": "ERRO", "Valor": "0,00", "Vencimento": DATA_VENCIMENTO,
        "Assunto": "FALHA NA EXTRAÇÃO - VERIFICAR PDF"
    }

# ====================================================================
# --- 6. MOTOR PRINCIPAL ---
# ====================================================================
def main():
    print("=== SUPER EXTRATOR HÍBRIDO DEFINITIVO (VERSÃO FINAL) ===")
    
    if not os.path.exists(PASTA_RAIZ):
        print(f"ERRO: Pasta não encontrada -> {PASTA_RAIZ}")
        return

    caminho_excel = os.path.join(PASTA_RAIZ, "Relatorio_Mestre_Hibrido_Final.xlsx")
    
    dados_finais = []
    arquivos_processados = set()
    
    if os.path.exists(caminho_excel):
        try:
            df_existente = pd.read_excel(caminho_excel, sheet_name=0) 
            if 'Caminho Completo' in df_existente.columns:
                arquivos_processados = set(df_existente['Caminho Completo'].dropna().tolist())
                dados_finais = df_existente.to_dict('records')
                print(f">> Relatório anterior encontrado! {len(arquivos_processados)} notas puladas.\n")
            else:
                print(">> Planilha antiga detetada. Ignorando memória para aplicar as correções...\n")
        except: pass

    arquivos_pdf = []
    for raiz, diretorios, arquivos in os.walk(PASTA_RAIZ):
        for arquivo in arquivos:
            if arquivo.lower().endswith('.pdf'):
                arquivos_pdf.append(os.path.join(raiz, arquivo))

    total = len(arquivos_pdf)
    print(f">> Encontrados {total} PDFs. Iniciando motor...\n")

    for i, caminho_pdf in enumerate(arquivos_pdf, 1):
        nome_ficheiro = os.path.basename(caminho_pdf)
        
        if caminho_pdf in arquivos_processados:
            print(f"[{i}/{total}] {nome_ficheiro}... [PULADO]")
            continue

        print(f"[{i}/{total}] Lendo: {nome_ficheiro}...", end="", flush=True)
        
        resultado = processar_documento_hibrido(caminho_pdf)
        
        if resultado:
            dados_finais.append(resultado)
            arquivos_processados.add(caminho_pdf)
            pd.DataFrame(dados_finais).to_excel(caminho_excel, index=False)
            
    # --- CÁLCULO DO RESUMO FINANCEIRO ---
    if dados_finais:
        df_principal = pd.DataFrame(dados_finais)
        df_principal['Valor_Calculo'] = df_principal['Valor'].apply(limpar_valor_float)
        resumo = df_principal.groupby('FORNECEDOR')['Valor_Calculo'].sum().reset_index()
        resumo = resumo.rename(columns={'Valor_Calculo': 'Subtotal (R$)'})
        total_geral = df_principal['Valor_Calculo'].sum()
        linha_total = pd.DataFrame([{'FORNECEDOR': 'TOTAL GERAL', 'Subtotal (R$)': total_geral}])
        resumo = pd.concat([resumo, linha_total], ignore_index=True)
        
        try:
            with pd.ExcelWriter(caminho_excel, engine='openpyxl') as writer:
                df_principal.drop(columns=['Valor_Calculo']).to_excel(writer, sheet_name='Notas Extraídas', index=False)
                resumo.to_excel(writer, sheet_name='Resumo Financeiro', index=False)
            
            print(f"\n[FINALIZADO] Ficheiro gerado com sucesso em: {caminho_excel}")
            
            print("\n" + "="*50)
            print("  RESUMO FINANCEIRO POR FORNECEDOR ")
            print("="*50)
            for index, row in resumo.iterrows():
                if row['FORNECEDOR'] != 'TOTAL GERAL':
                    valor_br = f"{row['Subtotal (R$)']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    print(f" {row['FORNECEDOR']:<25}: R$ {valor_br:>15}")
            print("-" * 50)
            total_br = f"{total_geral:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            print(f" TOTAL GERAL               : R$ {total_br:>15}")
            print("="*50 + "\n")
        except Exception as e:
            print(f"\n[ERRO] Feche o arquivo Excel se ele estiver aberto: {e}")

if __name__ == "__main__":
    main()