#!/usr/bin/env python3
"""
ROTINA COMPLETA DE CT-e — consulta SEFAZ, extrai dados, filtra por data,
atualiza a coluna "Frete" na planilha
==============================================================================
Faz tudo em um script só, pensado pra rodar sozinho (agendado no Agendador
de Tarefas do Windows, por exemplo):

  1. Carrega o certificado digital (.pfx)
  2. Consulta a SEFAZ (CTeDistribuicaoDFe) a partir do último NSU salvo
  3. Salva os XMLs recebidos, nomeados por NF + transportadora
  4. Extrai os dados de cada CT-e (pulando eventos e CT-e cancelados)
  5. Filtra por data de emissão do CT-e (opcional — configure abaixo)
  6. Atualiza a coluna "Frete" da planilha do Google Sheets, casando pela NF
     (soma quando uma NF tiver mais de um CT-e vinculado)

INSTALAÇÃO:
    pip install cryptography requests signxml lxml gspread google-auth

USO MANUAL:
    python3 rotina_cte.py

USO AGENDADO (sem digitar senha toda vez): defina a variável de ambiente
CERT_SENHA antes de rodar (ex., no Agendador de Tarefas do Windows, na
aba "Ação" → "Adicionar argumentos", ou via `set CERT_SENHA=...` antes
do comando). Se a variável não existir, o script pede a senha na hora.
"""

import getpass
import gzip
import json
import os
import re
import time
from base64 import b64decode
from datetime import datetime, timedelta

import gspread
import requests
from cryptography.hazmat.primitives.serialization import (
    Encoding,
    NoEncryption,
    PrivateFormat,
    pkcs12,
)
from google.oauth2.service_account import Credentials
from lxml import etree
from signxml import XMLSigner, methods
from brazilfiscalreport.dacte import Dacte

# Pasta onde estão o script, certificado, credenciais e arquivos de controle.
# Mantém os arquivos relativos funcionando mesmo ao abrir por atalho/clique duplo.
BASE_DIR = os.environ.get("CTE_BASE_DIR") or os.path.dirname(os.path.abspath(__file__))
if os.path.isdir(BASE_DIR):
    os.chdir(BASE_DIR)


def _carregar_env(caminho=".env"):
    """Lê um .env simples (CHAVE=valor) sem depender de bibliotecas extras."""
    if not os.path.exists(caminho):
        return
    with open(caminho, encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if linha and not linha.startswith("#") and "=" in linha:
                chave, valor = linha.split("=", 1)
                os.environ.setdefault(chave.strip(), valor.strip())


_carregar_env()


# ═══════════════════════════════════════════════════════════════════════════
# CONFIGURAÇÃO — ajuste aqui, sem precisar digitar nada na hora de rodar
# ═══════════════════════════════════════════════════════════════════════════

# --- Certificado e CNPJ ---
CAMINHO_PFX = os.path.join(BASE_DIR, os.environ.get("CERT_ARQUIVO", "certificado.pfx"))
CNPJ_CONSULTADO = os.environ.get("CNPJ_CONSULTADO", "")  # CNPJ do tomador, só números

# --- Ambiente e UF ---
AMBIENTE = "producao"   # "homologacao" ou "producao"
TP_AMB = "2" if AMBIENTE == "homologacao" else "1"
CUF_AUTOR = os.environ.get("CUF_AUTOR", "41")  # código IBGE da UF: 41=PR, 35=SP

# --- Filtro de DATA DE EMISSÃO do CT-e (opcional) ---
# Esse filtro é sobre a data que o CT-e foi EMITIDO pela transportadora —
# diferente do NSU, que controla o que já foi baixado.
#
# FILTRAR_MES_ATUAL = True: calcula sozinho o intervalo do mês corrente
# (não precisa atualizar todo mês). Deixe False e preencha
# DATA_EMISSAO_INICIO/FIM manualmente (formato "dd/mm/aaaa") se quiser
# um período específico, ou deixe os dois None pra não filtrar nada.
FILTRAR_MES_ATUAL = True

DATA_EMISSAO_INICIO = None   # usado só se FILTRAR_MES_ATUAL = False. ex.: "01/07/2026"
DATA_EMISSAO_FIM = None      # usado só se FILTRAR_MES_ATUAL = False. ex.: "31/07/2026"

# --- Google Sheets ---
CAMINHO_CREDENCIAIS_GOOGLE = os.environ.get("GOOGLE_CREDENCIAIS", "credenciais_google.json")
PLANILHA_URL_OU_ID = os.environ.get("PLANILHA_URL_OU_ID", "")  # URL ou ID da planilha de destino

# --- Arquivos/pastas de trabalho ---
ARQUIVO_ESTADO_NSU = "ultimo_nsu.json"
ARQUIVO_CANCELADOS = "ctes_cancelados.json"
ARQUIVO_CTES_SALVOS = "ctes_salvos.json"
ARQUIVO_NFS_ATUALIZADAS = "nfs_frete_atualizado.json"
PASTA_SAIDA_XML = "cte_recebidos"
ARQUIVO_SAIDA_CSV = "cte_extraido.csv"

ABA_SAIDAS = os.environ.get("ABA_SAIDAS", "Detalhado Saídas")
COLUNA_NF = 1       # A
COLUNA_FRETE = 17   # Q — "Frete" (mudou de P para Q depois que "Valor sem NF" foi inserido)
COLUNA_CTE_VINCULADO = 31   # AE — "CT-e Vinculado" (última coluna da planilha)

# --- Constantes técnicas (não precisa mexer) ---
URL_HOMOLOGACAO = "https://hom1.cte.fazenda.gov.br/CTeDistribuicaoDFe/CTeDistribuicaoDFe.asmx"
URL_PRODUCAO = "https://www1.cte.fazenda.gov.br/CTeDistribuicaoDFe/CTeDistribuicaoDFe.asmx"
SOAP_ACTION = "http://www.portalfiscal.inf.br/cte/wsdl/CTeDistribuicaoDFe/cteDistDFeInteresse"
CTE_NS = "http://www.portalfiscal.inf.br/cte"
ASSINAR_XML = False
TP_EVENTO_CANCELAMENTO = "110111"
TIPOS_EVENTO = {"110111": "Cancelamento", "110110": "CartaCorrecao", "110140": "EPEC"}
ESCOPOS_GOOGLE = ["https://www.googleapis.com/auth/spreadsheets"]

# Data "de hoje" confirmada pela própria SEFAZ (preenchida em consultar_sefaz).
# Usada no lugar do relógio do computador pro filtro de mês atual — assim o
# filtro não quebra se o relógio local estiver errado.
_HOJE_SEFAZ = None


# ═══════════════════════════════════════════════════════════════════════════
# 1) CERTIFICADO
# ═══════════════════════════════════════════════════════════════════════════
def carregar_certificado(caminho_pfx, senha):
    with open(caminho_pfx, "rb") as f:
        dados_pfx = f.read()
    chave_privada, certificado, _ = pkcs12.load_key_and_certificates(dados_pfx, senha.encode("utf-8"))
    if chave_privada is None or certificado is None:
        raise ValueError("Não consegui extrair chave/certificado do .pfx. Confira a senha.")

    import tempfile
    tmp_dir = tempfile.mkdtemp(prefix="cte_cert_")
    caminho_cert = os.path.join(tmp_dir, "cert.pem")
    caminho_key = os.path.join(tmp_dir, "key.pem")
    with open(caminho_cert, "wb") as f:
        f.write(certificado.public_bytes(Encoding.PEM))
    with open(caminho_key, "wb") as f:
        f.write(chave_privada.private_bytes(Encoding.PEM, PrivateFormat.TraditionalOpenSSL, NoEncryption()))

    titular = certificado.subject.rfc4514_string()
    print(f"✓ Certificado carregado — titular: {titular}")
    return caminho_cert, caminho_key


# ═══════════════════════════════════════════════════════════════════════════
# 2) CONSULTA À SEFAZ
# ═══════════════════════════════════════════════════════════════════════════
def carregar_ultimo_nsu(cnpj):
    chave = f"{AMBIENTE}:{cnpj}"
    if os.path.exists(ARQUIVO_ESTADO_NSU):
        with open(ARQUIVO_ESTADO_NSU, encoding="utf-8") as f:
            return json.load(f).get(chave, "0")
    return "0"


def salvar_ultimo_nsu(cnpj, nsu):
    chave = f"{AMBIENTE}:{cnpj}"
    estado = {}
    if os.path.exists(ARQUIVO_ESTADO_NSU):
        with open(ARQUIVO_ESTADO_NSU, encoding="utf-8") as f:
            estado = json.load(f)
    estado[chave] = nsu
    with open(ARQUIVO_ESTADO_NSU, "w", encoding="utf-8") as f:
        json.dump(estado, f, indent=2, ensure_ascii=False)


# ── Lista leve de CT-e cancelados (só a chave, não o XML do evento inteiro) ─
def carregar_cancelados():
    if os.path.exists(ARQUIVO_CANCELADOS):
        with open(ARQUIVO_CANCELADOS, encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def adicionar_cancelado(chave_cte):
    cancelados = carregar_cancelados()
    if chave_cte in cancelados:
        return
    cancelados.add(chave_cte)
    with open(ARQUIVO_CANCELADOS, "w", encoding="utf-8") as f:
        json.dump(sorted(cancelados), f, indent=2, ensure_ascii=False)


# ── Lista leve de CT-e já salvos em disco (evita duplicar se o NSU for
#       resetado/apagado e o script reprocessar o backlog inteiro) ─────────
def carregar_ctes_salvos():
    if os.path.exists(ARQUIVO_CTES_SALVOS):
        with open(ARQUIVO_CTES_SALVOS, encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def adicionar_cte_salvo(chave_cte):
    salvos = carregar_ctes_salvos()
    if chave_cte in salvos:
        return
    salvos.add(chave_cte)
    with open(ARQUIVO_CTES_SALVOS, "w", encoding="utf-8") as f:
        json.dump(sorted(salvos), f, indent=2, ensure_ascii=False)


# ── Lista leve de NFs cujo "Frete" já foi escrito na planilha pela rotina
#       — depois da primeira vez, NUNCA mais mexe nessa célula, mesmo que
#       apareça outro CT-e pra mesma NF. Isso preserva ajustes manuais que
#       você faça direto na planilha (ex.: valor combinado ≠ valor do CT-e),
#       sem a rotina sobrescrever na próxima execução. ─────────────────────
def carregar_nfs_atualizadas():
    if os.path.exists(ARQUIVO_NFS_ATUALIZADAS):
        with open(ARQUIVO_NFS_ATUALIZADAS, encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def marcar_nf_atualizada(nf):
    atualizadas = carregar_nfs_atualizadas()
    if nf in atualizadas:
        return
    atualizadas.add(nf)
    with open(ARQUIVO_NFS_ATUALIZADAS, "w", encoding="utf-8") as f:
        json.dump(sorted(atualizadas), f, indent=2, ensure_ascii=False)


MESES_PT = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]


def pasta_do_mes(data_emissao_iso):
    """Nome da subpasta pro mês de emissão do CT-e (ex.: '2026-07 Julho').
    Se não conseguir ler a data, devolve None (o arquivo cai na pasta raiz)."""
    if not data_emissao_iso:
        return None
    try:
        data = datetime.fromisoformat(data_emissao_iso).date()
    except ValueError:
        try:
            data = datetime.strptime(data_emissao_iso[:10], "%Y-%m-%d").date()
        except Exception:
            return None
    return f"{data.year:04d}-{data.month:02d} {MESES_PT[data.month - 1]}"


def montar_dist_dfe_int(cnpj, ult_nsu):
    xml_texto = (
        f'<distDFeInt xmlns="{CTE_NS}" versao="1.00">'
        f"<tpAmb>{TP_AMB}</tpAmb><cUFAutor>{CUF_AUTOR}</cUFAutor><CNPJ>{cnpj}</CNPJ>"
        f"<distNSU><ultNSU>{str(ult_nsu).zfill(15)}</ultNSU></distNSU></distDFeInt>"
    )
    return etree.fromstring(xml_texto.encode("utf-8"))


def montar_envelope_soap(dist):
    dist_str = etree.tostring(dist, encoding="unicode")
    return (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        'xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
        'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
        "<soap:Body>"
        '<cteDistDFeInteresse xmlns="http://www.portalfiscal.inf.br/cte/wsdl/CTeDistribuicaoDFe">'
        f"<cteDadosMsg>{dist_str}</cteDadosMsg>"
        "</cteDistDFeInteresse></soap:Body></soap:Envelope>"
    )


def consultar_sefaz(cnpj, caminho_cert, caminho_key):
    ult_nsu = carregar_ultimo_nsu(cnpj)
    print(f"Consultando CT-e a partir do NSU {ult_nsu} (ambiente: {AMBIENTE})...")

    dist = montar_dist_dfe_int(cnpj, ult_nsu)
    envelope = montar_envelope_soap(dist)
    url = URL_HOMOLOGACAO if AMBIENTE == "homologacao" else URL_PRODUCAO
    headers = {"Content-Type": "text/xml; charset=utf-8", "SOAPAction": SOAP_ACTION}

    resp = requests.post(url, data=envelope.encode("utf-8"), headers=headers,
                          cert=(caminho_cert, caminho_key), timeout=60)
    if resp.status_code != 200:
        raise RuntimeError(f"SEFAZ respondeu HTTP {resp.status_code}: {resp.text[:1000]}")

    raiz = etree.fromstring(resp.text.encode("utf-8"))
    ns = {"cte": CTE_NS}
    cstat = raiz.find(".//cte:cStat", ns)
    xmotivo = raiz.find(".//cte:xMotivo", ns)
    ult_nsu_resp = raiz.find(".//cte:ultNSU", ns)
    max_nsu_resp = raiz.find(".//cte:maxNSU", ns)
    doc_zips = raiz.findall(".//cte:docZip", ns)

    # Guarda a data/hora que a PRÓPRIA SEFAZ devolveu (dhResp) como fonte
    # confiável de "hoje" — não depende do relógio do computador estar
    # certo, o que é importante pra um script que roda sozinho, sem
    # ninguém conferindo.
    global _HOJE_SEFAZ
    dh_resp = raiz.find(".//cte:dhResp", ns)
    if dh_resp is not None and dh_resp.text:
        try:
            _HOJE_SEFAZ = datetime.fromisoformat(dh_resp.text).date()
            print(f"  Data/hora confirmada pela SEFAZ: {dh_resp.text}")
        except ValueError:
            pass

    print(f"  cStat: {cstat.text if cstat is not None else '?'} — {xmotivo.text if xmotivo is not None else ''}")
    print(f"  Documentos recebidos: {len(doc_zips)}")

    arquivos_salvos = []
    nomes_usados_por_pasta = {}  # cache: pasta -> set de nomes de arquivo já usados nela
    ctes_ja_salvos = carregar_ctes_salvos()
    ctes_nao_autorizados = 0
    ctes_com_erro_leitura = 0
    ctes_ja_existentes_pulados = 0
    eventos_nao_salvos = 0
    cancelamentos_novos = 0

    if doc_zips:
        os.makedirs(PASTA_SAIDA_XML, exist_ok=True)
        for doc in doc_zips:
            nsu_doc = doc.get("NSU")
            schema_doc = doc.get("schema", "desconhecido")
            try:
                xml_bytes = gzip.decompress(b64decode(doc.text))
            except Exception as e:
                print(f"  ✗ Falha ao descompactar NSU {nsu_doc}: {e}")
                continue

            if "Evento" in schema_doc:
                # Eventos NÃO são salvos como arquivo — só processados em
                # memória. Se for cancelamento, guarda a chave numa lista
                # leve (ctes_cancelados.json), sem precisar do XML inteiro.
                tp_evento, ch_cte = extrair_evento_de_bytes(xml_bytes)
                if tp_evento == TP_EVENTO_CANCELAMENTO and ch_cte:
                    adicionar_cancelado(ch_cte)
                    cancelamentos_novos += 1
                eventos_nao_salvos += 1
                continue

            dados = extrair_dados_cte_de_bytes(xml_bytes)

            # Já salvamos esse CT-e antes (ex.: NSU foi resetado e está
            # reprocessando o backlog)? Pula, sem duplicar arquivo.
            if not dados.get("erro") and dados["chave_cte"] in ctes_ja_salvos:
                ctes_ja_existentes_pulados += 1
                continue

            if dados.get("erro"):
                pasta_destino = PASTA_SAIDA_XML  # sem data legível — vai pra raiz
                nome_arquivo = f"NSU{nsu_doc}_{schema_doc.replace('.xsd', '')}.xml"
                ctes_com_erro_leitura += 1
                print(f"  ⚠ NSU {nsu_doc} ({schema_doc}): não consegui ler os dados pra nomear "
                      f"(motivo: {dados['erro']}) — salvo com nome genérico. Me manda esse "
                      f"arquivo que eu ajusto a leitura.")
            elif not (dados.get("situacao") or "").startswith("100"):
                # Não autorizado (rejeitado, denegado etc.) — não salva
                ctes_nao_autorizados += 1
                print(f"  ⊘ NSU {nsu_doc}: não autorizado ({dados.get('situacao')}) — não salvo")
                continue
            else:
                subpasta = pasta_do_mes(dados["data_emissao"])
                pasta_destino = os.path.join(PASTA_SAIDA_XML, subpasta) if subpasta else PASTA_SAIDA_XML
                os.makedirs(pasta_destino, exist_ok=True)
                if pasta_destino not in nomes_usados_por_pasta:
                    nomes_usados_por_pasta[pasta_destino] = set(os.listdir(pasta_destino))
                nome_arquivo = montar_nome_cte(dados, nomes_usados_por_pasta[pasta_destino])

            nomes_usados_por_pasta.setdefault(pasta_destino, set()).add(nome_arquivo)
            caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
            with open(caminho_arquivo, "wb") as f:
                f.write(xml_bytes)
            arquivos_salvos.append(caminho_arquivo)
            if not dados.get("erro"):
                adicionar_cte_salvo(dados["chave_cte"])

            # Gera o PDF (DACTE) na mesma pasta, com o mesmo nome do XML —
            # só pros CT-e autorizados de verdade (não pros de nome genérico
            # com erro de leitura, já que a lib provavelmente vai falhar
            # igual nesses casos)
            if not dados.get("erro"):
                caminho_pdf = caminho_arquivo.rsplit(".xml", 1)[0] + ".pdf"
                try:
                    dacte = Dacte(xml=xml_bytes.decode("utf-8"))
                    dacte.output(caminho_pdf)
                except Exception as e:
                    print(f"  ⚠ Não consegui gerar o PDF de '{nome_arquivo}': {e}")

    if ctes_nao_autorizados:
        print(f"  ⓘ {ctes_nao_autorizados} CT-e não autorizado(s) foram ignorados (não salvos)")
    if ctes_ja_existentes_pulados:
        print(f"  ⓘ {ctes_ja_existentes_pulados} CT-e já estava(m) salvo(s) de antes — pulado(s) sem duplicar")
    if ctes_com_erro_leitura:
        print(f"  ⚠ {ctes_com_erro_leitura} CT-e salvo(s) com nome genérico (NSU...) por erro de leitura — "
              f"veja os avisos acima")
    if eventos_nao_salvos:
        print(f"  ⓘ {eventos_nao_salvos} evento(s) processado(s) em memória (não salvos como arquivo)"
              + (f", {cancelamentos_novos} cancelamento(s) novo(s) registrado(s)" if cancelamentos_novos else ""))

    if ult_nsu_resp is not None and ult_nsu_resp.text:
        salvar_ultimo_nsu(cnpj, ult_nsu_resp.text)
        print(f"  NSU salvo para a próxima execução: {ult_nsu_resp.text}")

    ult_nsu_int = int(ult_nsu_resp.text) if ult_nsu_resp is not None and ult_nsu_resp.text else None
    max_nsu_int = int(max_nsu_resp.text) if max_nsu_resp is not None and max_nsu_resp.text else None
    return arquivos_salvos, ult_nsu_int, max_nsu_int


def consultar_sefaz_ate_atualizar(cnpj, caminho_cert, caminho_key, max_consultas=200):
    """Chama consultar_sefaz em loop, avançando o NSU sozinho, até não ter
    mais documentos novos na fila da SEFAZ (ultNSU alcança maxNSU) — assim
    não precisa rodar o script várias vezes na mão só pra 'passar' pelo
    backlog antigo até chegar nos documentos recentes."""
    for i in range(1, max_consultas + 1):
        if i > 1:
            print(f"\n  ── Consulta {i}/{max_consultas} (ainda tem documento novo na fila) ──")
            time.sleep(1)  # educado com a SEFAZ, evita parecer abuso de consulta

        _, ult_nsu, max_nsu = consultar_sefaz(cnpj, caminho_cert, caminho_key)

        if ult_nsu is None or max_nsu is None:
            break
        if ult_nsu >= max_nsu:
            print("  ✓ Alcançou o fim da fila — não há mais documentos novos na SEFAZ.")
            break
    else:
        print(f"  ⚠ Parou após {max_consultas} consultas (limite de segurança) — ainda pode ter mais "
              f"documentos. Roda a rotina de novo pra continuar de onde parou.")


# ═══════════════════════════════════════════════════════════════════════════
# 3) NOMEAÇÃO E EXTRAÇÃO DOS XML
# ═══════════════════════════════════════════════════════════════════════════
def texto(elemento, xpath_local_names):
    caminho = "".join(f'/*[local-name()="{nome}"]' for nome in xpath_local_names)
    resultado = elemento.xpath("." + caminho)
    return resultado[0].text if resultado else None


def nome_seguro(txt, tamanho_maximo=30):
    if not txt:
        return ""
    limpo = re.sub(r'[<>:"/\\|?*]', "", txt).strip()
    return re.sub(r"\s+", "_", limpo)[:tamanho_maximo]


def extrair_evento_de_bytes(conteudo):
    try:
        raiz = etree.fromstring(conteudo)
    except Exception:
        return None, None
    tp = raiz.xpath('//*[local-name()="tpEvento"]/text()')
    ch = raiz.xpath('//*[local-name()="chCTe"]/text()')
    return (tp[0].strip() if tp else None), (ch[0].strip() if ch else None)


def extrair_dados_cte_de_bytes(conteudo, nome_referencia="(memória)"):
    try:
        raiz = etree.fromstring(conteudo)
    except Exception as e:
        return {"arquivo": nome_referencia, "erro": f"XML inválido: {e}"}

    inf_cte = raiz.xpath('//*[local-name()="infCte"]')
    if not inf_cte:
        return {"arquivo": nome_referencia, "erro": "Não encontrei <infCte> no XML"}
    inf_cte = inf_cte[0]

    chave = inf_cte.get("Id", "")
    if chave.upper().startswith("CTE"):
        chave = chave[3:]
    if not chave:
        ch_prot = raiz.xpath('//*[local-name()="infProt"]/*[local-name()="chCTe"]')
        chave = ch_prot[0].text if ch_prot else ""

    dh_emi = texto(inf_cte, ["ide", "dhEmi"])
    emit_nome = texto(inf_cte, ["emit", "xNome"])
    emit_cnpj = texto(inf_cte, ["emit", "CNPJ"])
    valor_total = texto(inf_cte, ["vPrest", "vTPrest"])

    cstat = raiz.xpath('//*[local-name()="infProt"]/*[local-name()="cStat"]')
    xmotivo = raiz.xpath('//*[local-name()="infProt"]/*[local-name()="xMotivo"]')
    situacao = f"{cstat[0].text} - {xmotivo[0].text}" if cstat and xmotivo else ""

    chaves_nfe = raiz.xpath('//*[local-name()="infNFe"]/*[local-name()="chave"]/text()')
    nfs = []
    for chave_nfe in chaves_nfe:
        chave_nfe = chave_nfe.strip()
        if chave_nfe and len(chave_nfe) >= 34:
            numero_nf = chave_nfe[25:34].lstrip("0") or "0"
            nfs.append({"chave_nfe": chave_nfe, "numero_nf": numero_nf})

    return {
        "arquivo": nome_referencia, "chave_cte": chave, "data_emissao": dh_emi,
        "transportadora": emit_nome, "cnpj_transportadora": emit_cnpj,
        "valor_total_cte": valor_total, "situacao": situacao,
        "nfs_vinculadas": nfs, "erro": None,
    }


def montar_nome_cte(dados, nomes_usados):
    """Nome no formato 'NF - TRANSPORTADORA.xml'. Se já existir um arquivo
    com esse nome (mesma NF + transportadora, de um CT-e diferente),
    adiciona ' (2)', ' (3)' etc. pra não sobrescrever."""
    transp = nome_seguro(dados.get("transportadora")) or "SemTransportadora"
    nfs = dados.get("nfs_vinculadas") or []
    numeros_nf = "+".join(nf["numero_nf"] for nf in nfs) if nfs else "SemNF"

    nome_base = f"{numeros_nf} - {transp}"
    nome_arquivo = f"{nome_base}.xml"
    contador = 2
    while nome_arquivo in nomes_usados:
        nome_arquivo = f"{nome_base} ({contador}).xml"
        contador += 1
    return nome_arquivo


def eh_evento(nome_arquivo):
    return "evento" in nome_arquivo.lower()


def obter_intervalo_do_filtro():
    """Devolve (data_inicio, data_fim) como objetos date, ou (None, None)
    se não houver filtro configurado. Calcula o mês atual automaticamente
    quando FILTRAR_MES_ATUAL = True; senão usa DATA_EMISSAO_INICIO/FIM."""
    if FILTRAR_MES_ATUAL:
        # Prioriza a data confirmada pela SEFAZ (mais confiável); só usa o
        # relógio do computador se ainda não tivermos consultado a SEFAZ
        # nesta execução (ex.: chamando essa função isoladamente).
        if _HOJE_SEFAZ is not None:
            hoje = _HOJE_SEFAZ
        else:
            hoje = datetime.now().date()
        primeiro_dia = hoje.replace(day=1)
        if hoje.month == 12:
            ultimo_dia = hoje.replace(year=hoje.year + 1, month=1, day=1) - timedelta(days=1)
        else:
            ultimo_dia = hoje.replace(month=hoje.month + 1, day=1) - timedelta(days=1)
        return primeiro_dia, ultimo_dia

    di = datetime.strptime(DATA_EMISSAO_INICIO, "%d/%m/%Y").date() if DATA_EMISSAO_INICIO else None
    df = datetime.strptime(DATA_EMISSAO_FIM, "%d/%m/%Y").date() if DATA_EMISSAO_FIM else None
    return di, df


def data_dentro_do_filtro(data_emissao_iso):
    """Confere se a data de emissão do CT-e cai dentro do filtro
    configurado (mês atual, ou DATA_EMISSAO_INICIO/FIM manual). Sem
    filtro configurado, sempre deixa passar."""
    data_inicio, data_fim = obter_intervalo_do_filtro()
    if not data_inicio and not data_fim:
        return True
    if not data_emissao_iso:
        return True  # não conseguiu ler a data — não filtra silenciosamente, deixa passar e avisa em outro lugar

    try:
        data_cte = datetime.fromisoformat(data_emissao_iso).date()
    except ValueError:
        try:
            data_cte = datetime.strptime(data_emissao_iso[:10], "%Y-%m-%d").date()
        except Exception:
            return True

    if data_inicio and data_cte < data_inicio:
        return False
    if data_fim and data_cte > data_fim:
        return False
    return True


def processar_xmls_baixados():
    """Lê TODOS os XMLs de CT-e já salvos em PASTA_SAIDA_XML — inclusive
    dentro das subpastas por mês (ex.: '2026-07 Julho/') — pula os
    cancelados (usando a lista leve salva em ARQUIVO_CANCELADOS — os
    eventos em si não ficam salvos como arquivo), aplica o filtro de
    data, e devolve as linhas prontas pra planilha."""
    if not os.path.isdir(PASTA_SAIDA_XML):
        return []

    caminhos_xml = []
    for pasta_atual, _subpastas, arquivos_na_pasta in os.walk(PASTA_SAIDA_XML):
        for nome in arquivos_na_pasta:
            if nome.lower().endswith(".xml"):
                caminhos_xml.append(os.path.join(pasta_atual, nome))

    chaves_canceladas = carregar_cancelados()

    if chaves_canceladas:
        print(f"  ⚠ {len(chaves_canceladas)} CT-e cancelado(s) na lista — serão ignorados se aparecerem")

    linhas = []
    ignorados_por_data = 0
    for caminho in caminhos_xml:
        nome = os.path.basename(caminho)
        if eh_evento(nome):
            continue  # arquivo antigo de uma versão anterior do script — ignora
        with open(caminho, "rb") as f:
            dados = extrair_dados_cte_de_bytes(f.read(), nome)
        if dados.get("erro") or dados["chave_cte"] in chaves_canceladas:
            continue
        if not data_dentro_do_filtro(dados["data_emissao"]):
            ignorados_por_data += 1
            continue
        for nf in dados["nfs_vinculadas"]:
            linhas.append({
                "numero_nf": nf["numero_nf"], "chave_cte": dados["chave_cte"],
                "transportadora": dados["transportadora"], "data_emissao": dados["data_emissao"],
                "valor_total_cte": dados["valor_total_cte"], "situacao": dados["situacao"],
            })

    if ignorados_por_data:
        print(f"  ⓘ {ignorados_por_data} CT-e fora do filtro de data ignorado(s)")

    return linhas


# ═══════════════════════════════════════════════════════════════════════════
# 4) ATUALIZAÇÃO DA PLANILHA
# ═══════════════════════════════════════════════════════════════════════════
def normalizar_nf(valor):
    return str(valor or "").strip().lstrip("0") or "0"


def agrupar_por_nf(linhas):
    valores_por_nf, ctes_por_nf = {}, {}
    for linha in linhas:
        nf = normalizar_nf(linha["numero_nf"])
        if nf == "0":
            continue
        try:
            valor = float((linha["valor_total_cte"] or "0").replace(",", "."))
        except ValueError:
            continue
        valores_por_nf[nf] = valores_por_nf.get(nf, 0.0) + valor
        ctes_por_nf.setdefault(nf, []).append(linha["chave_cte"])
    return valores_por_nf, ctes_por_nf


def extrair_id_planilha(url_ou_id):
    if "/d/" in url_ou_id:
        return url_ou_id.split("/d/")[1].split("/")[0]
    return url_ou_id


def atualizar_planilha(valores_por_nf, ctes_por_nf):
    if not valores_por_nf:
        print("  Nenhuma NF pra atualizar.")
        return

    credenciais = Credentials.from_service_account_file(CAMINHO_CREDENCIAIS_GOOGLE, scopes=ESCOPOS_GOOGLE)
    cliente = gspread.authorize(credenciais)
    planilha = cliente.open_by_key(extrair_id_planilha(PLANILHA_URL_OU_ID))
    sheet = planilha.worksheet(ABA_SAIDAS)

    todos_valores = sheet.get_all_values()
    atualizacoes_frete = []
    atualizacoes_cte = []
    nfs_encontradas = set()
    nfs_para_marcar = []
    nfs_ja_atualizadas = carregar_nfs_atualizadas()
    nfs_puladas_por_edicao_manual = 0

    for idx, linha in enumerate(todos_valores[1:], start=2):
        nf_planilha = normalizar_nf(linha[COLUNA_NF - 1] if len(linha) >= COLUNA_NF else "")
        if nf_planilha in valores_por_nf:
            nfs_encontradas.add(nf_planilha)

            # Já escrevemos o Frete dessa NF antes — não mexe mais, pra não
            # sobrescrever um ajuste manual que você tenha feito na planilha
            # (ex.: o valor combinado com a transportadora é diferente do
            # valor que veio no CT-e).
            if nf_planilha in nfs_ja_atualizadas:
                nfs_puladas_por_edicao_manual += 1
                continue

            novo_valor = valores_por_nf[nf_planilha]
            chaves_cte = ctes_por_nf[nf_planilha]
            referencia_cte = " + ".join(chaves_cte)

            atualizacoes_frete.append((idx, novo_valor))
            atualizacoes_cte.append((idx, referencia_cte))
            nfs_para_marcar.append(nf_planilha)
            print(f"    NF {nf_planilha} (linha {idx}): Frete → R$ {novo_valor:.2f} "
                  f"({len(chaves_cte)} CT-e)")

    if nfs_puladas_por_edicao_manual:
        print(f"  ⓘ {nfs_puladas_por_edicao_manual} NF(s) já tinham sido atualizadas antes — "
              f"não mexidas de novo (protege edição manual do Frete)")

    if atualizacoes_frete:
        # Garante que a planilha tem colunas suficientes — se ela foi criada
        # antes da coluna "CT-e Vinculado" existir, o grid físico pode ser
        # menor do que o necessário, e tanto o format() quanto o
        # update_cells() dão erro "exceeds grid limits" nesse caso.
        if sheet.col_count < COLUNA_CTE_VINCULADO:
            sheet.add_cols(COLUNA_CTE_VINCULADO - sheet.col_count)
            sheet.update_cell(1, COLUNA_CTE_VINCULADO, "CT-e Vinculado")
            print(f"  ⓘ Planilha só tinha {sheet.col_count} coluna(s) — expandida pra caber "
                  f"a coluna 'CT-e Vinculado'.")

        # Trava a coluna "CT-e Vinculado" como texto puro ANTES de escrever —
        # a chave do CT-e tem 44 dígitos, e sem isso o Sheets tenta
        # converter pra número e perde precisão (aparece em notação
        # científica ou cortado). Usa um range limitado (não aberto) pra
        # não estourar o limite de linhas da planilha.
        col_letra = "".join(filter(str.isalpha, gspread.utils.rowcol_to_a1(1, COLUNA_CTE_VINCULADO)))
        ultima_linha = max(sheet.row_count, len(todos_valores) + 10)
        sheet.format(f"{col_letra}2:{col_letra}{ultima_linha}", {"numberFormat": {"type": "TEXT"}})

        celulas = (
            [gspread.Cell(row=l, col=COLUNA_FRETE, value=v) for l, v in atualizacoes_frete]
            + [gspread.Cell(row=l, col=COLUNA_CTE_VINCULADO, value=v) for l, v in atualizacoes_cte]
        )
        sheet.update_cells(celulas, value_input_option="USER_ENTERED")
        print(f"  ✓ {len(atualizacoes_frete)} nota(s) atualizada(s) na planilha (Frete + CT-e Vinculado).")

        for nf in nfs_para_marcar:
            marcar_nf_atualizada(nf)

    faltando = set(valores_por_nf) - nfs_encontradas
    if faltando:
        print(f"  ⚠ {len(faltando)} NF(s) não encontradas na planilha: {', '.join(sorted(faltando))}")


# ═══════════════════════════════════════════════════════════════════════════
# ROTINA PRINCIPAL + INTERFACE GRÁFICA
# ═══════════════════════════════════════════════════════════════════════════
def executar_rotina(senha):
    """Executa exatamente o mesmo fluxo da rotina original, usando a senha informada."""
    print("=" * 70)
    print("ROTINA COMPLETA DE CT-e")
    print("=" * 70)

    if not senha:
        raise ValueError("Informe a senha do certificado digital.")
    if not CNPJ_CONSULTADO:
        raise ValueError("Defina CNPJ_CONSULTADO no .env (veja .env.example).")

    print("\n[1/4] Carregando certificado...")
    caminho_cert, caminho_key = carregar_certificado(CAMINHO_PFX, senha)

    print("\n[2/4] Consultando a SEFAZ (avança sozinho até alcançar os documentos mais recentes)...")
    consultar_sefaz_ate_atualizar(CNPJ_CONSULTADO, caminho_cert, caminho_key)

    print("\n[3/4] Extraindo dados dos CT-e (aplicando filtro de data, se configurado)...")
    _di, _df = obter_intervalo_do_filtro()
    if _di or _df:
        origem = "mês atual" if FILTRAR_MES_ATUAL else "configurado manualmente"
        print(f"  Filtro de data de emissão ({origem}): "
              f"{_di.strftime('%d/%m/%Y') if _di else '(sem início)'} "
              f"até {_df.strftime('%d/%m/%Y') if _df else '(sem fim)'}")
    linhas = processar_xmls_baixados()
    print(f"  {len(linhas)} linha(s) de CT-e válidas dentro do filtro.")

    if linhas:
        with open(ARQUIVO_SAIDA_CSV, "w", newline="", encoding="utf-8-sig") as f:
            import csv
            w = csv.DictWriter(f, fieldnames=list(linhas[0].keys()), delimiter=";")
            w.writeheader()
            w.writerows(linhas)
        print(f"  ✓ Salvo em {ARQUIVO_SAIDA_CSV}")

    print("\n[4/4] Atualizando planilha do Google Sheets...")
    if not PLANILHA_URL_OU_ID:
        print("  ✗ Configure PLANILHA_URL_OU_ID no topo do script antes de rodar essa etapa.")
    elif not os.path.isfile(CAMINHO_CREDENCIAIS_GOOGLE):
        print(f"  ✗ Não encontrei o arquivo de credenciais: {CAMINHO_CREDENCIAIS_GOOGLE}")
    else:
        valores_por_nf, ctes_por_nf = agrupar_por_nf(linhas)
        atualizar_planilha(valores_por_nf, ctes_por_nf)

    print("\n" + "=" * 70)
    print("Rotina concluída.")
    print("=" * 70)


def main():
    """Modo terminal antigo, preservado para uso agendado ou manual."""
    senha = os.environ.get("CERT_SENHA") or getpass.getpass("Senha do certificado: ")
    executar_rotina(senha)


class _ConsoleParaTela:
    def __init__(self, fila):
        self.fila = fila

    def write(self, texto):
        if texto:
            self.fila.put(texto)

    def flush(self):
        pass


def abrir_interface():
    import queue
    import subprocess
    import sys
    import threading
    import tkinter as tk
    from tkinter import messagebox
    from tkinter import ttk

    AZUL = "#0F5FA8"
    AZUL_ESCURO = "#0B3F73"
    AZUL_CLARO = "#EAF4FF"
    BRANCO = "#FFFFFF"
    CINZA = "#F5F7FA"
    TEXTO = "#1F2937"

    janela = tk.Tk()
    janela.title("Rotina CT-e | SEFAZ")
    janela.geometry("920x640")
    janela.minsize(820, 560)
    janela.configure(bg=CINZA)

    style = ttk.Style(janela)
    style.theme_use("clam")
    style.configure("TFrame", background=CINZA)
    style.configure("Card.TFrame", background=BRANCO)
    style.configure("TLabel", background=CINZA, foreground=TEXTO, font=("Segoe UI", 10))
    style.configure("Card.TLabel", background=BRANCO, foreground=TEXTO, font=("Segoe UI", 10))
    style.configure("Title.TLabel", background=AZUL, foreground=BRANCO, font=("Segoe UI", 18, "bold"))
    style.configure("Subtitle.TLabel", background=AZUL, foreground="#DCEEFF", font=("Segoe UI", 10))
    style.configure("Primary.TButton", font=("Segoe UI", 11, "bold"), padding=(18, 10))
    style.map("Primary.TButton", background=[("active", AZUL_ESCURO)], foreground=[("active", BRANCO)])

    fila_logs = queue.Queue()
    executando = {"valor": False}

    header = tk.Frame(janela, bg=AZUL, height=92)
    header.pack(fill="x")
    header.pack_propagate(False)

    ttk.Label(header, text="Rotina completa de CT-e", style="Title.TLabel").pack(anchor="w", padx=28, pady=(18, 0))
    ttk.Label(
        header,
        text="Consulta SEFAZ, salva XML/PDF, extrai CT-e e atualiza o Google Sheets — mesmas funcionalidades do script original.",
        style="Subtitle.TLabel",
    ).pack(anchor="w", padx=30, pady=(4, 0))

    corpo = ttk.Frame(janela, padding=22)
    corpo.pack(fill="both", expand=True)

    card = ttk.Frame(corpo, style="Card.TFrame", padding=18)
    card.pack(fill="x")

    ttk.Label(card, text="Senha do certificado digital", style="Card.TLabel", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
    senha_var = tk.StringVar(value=os.environ.get("CERT_SENHA", ""))
    entrada_senha = ttk.Entry(card, textvariable=senha_var, show="•", width=34, font=("Segoe UI", 11))
    entrada_senha.grid(row=1, column=0, sticky="ew", pady=(6, 0))

    mostrar_var = tk.BooleanVar(value=False)
    def alternar_senha():
        entrada_senha.configure(show="" if mostrar_var.get() else "•")
    ttk.Checkbutton(card, text="Mostrar senha", variable=mostrar_var, command=alternar_senha).grid(row=1, column=1, padx=(12, 0), sticky="w")

    info = (
        f"Pasta: {BASE_DIR}\n"
        f"Certificado: {CAMINHO_PFX}\n"
        f"Credenciais Google: {CAMINHO_CREDENCIAIS_GOOGLE}\n"
        f"Ambiente: {AMBIENTE} | CNPJ: {CNPJ_CONSULTADO}"
    )
    ttk.Label(card, text=info, style="Card.TLabel", justify="left").grid(row=2, column=0, columnspan=3, sticky="w", pady=(12, 0))
    card.columnconfigure(0, weight=1)

    botoes = ttk.Frame(card, style="Card.TFrame")
    botoes.grid(row=1, column=2, rowspan=2, sticky="e", padx=(18, 0))

    btn_executar = tk.Button(
        botoes,
        text="▶ Executar rotina",
        bg=AZUL,
        fg=BRANCO,
        activebackground=AZUL_ESCURO,
        activeforeground=BRANCO,
        relief="flat",
        padx=22,
        pady=11,
        font=("Segoe UI", 11, "bold"),
        cursor="hand2",
    )
    btn_executar.pack(fill="x")

    def abrir_pasta():
        if os.path.isdir(BASE_DIR):
            subprocess.Popen(f'explorer "{BASE_DIR}"')
        else:
            messagebox.showwarning("Pasta não encontrada", f"Não encontrei a pasta:\n{BASE_DIR}")

    tk.Button(
        botoes,
        text="Abrir pasta",
        bg=AZUL_CLARO,
        fg=AZUL_ESCURO,
        activebackground="#D9ECFF",
        activeforeground=AZUL_ESCURO,
        relief="flat",
        padx=18,
        pady=8,
        font=("Segoe UI", 10, "bold"),
        cursor="hand2",
        command=abrir_pasta,
    ).pack(fill="x", pady=(10, 0))

    status_var = tk.StringVar(value="Pronto para executar")
    ttk.Label(corpo, textvariable=status_var, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(16, 6))

    area_logs_frame = tk.Frame(corpo, bg=BRANCO, highlightthickness=1, highlightbackground="#D8E2EE")
    area_logs_frame.pack(fill="both", expand=True)

    area_logs = tk.Text(
        area_logs_frame,
        bg=BRANCO,
        fg=TEXTO,
        insertbackground=TEXTO,
        relief="flat",
        wrap="word",
        font=("Consolas", 10),
        padx=14,
        pady=12,
    )
    scroll = ttk.Scrollbar(area_logs_frame, orient="vertical", command=area_logs.yview)
    area_logs.configure(yscrollcommand=scroll.set)
    area_logs.pack(side="left", fill="both", expand=True)
    scroll.pack(side="right", fill="y")

    def log(texto):
        area_logs.configure(state="normal")
        area_logs.insert("end", texto)
        area_logs.see("end")
        area_logs.configure(state="disabled")

    area_logs.configure(state="disabled")
    log("Aguardando execução...\n")

    def consumir_logs():
        try:
            while True:
                log(fila_logs.get_nowait())
        except queue.Empty:
            pass
        janela.after(120, consumir_logs)

    def finalizar_ok():
        executando["valor"] = False
        btn_executar.configure(state="normal", text="▶ Executar rotina")
        status_var.set("Concluído")
        messagebox.showinfo("Rotina concluída", "A rotina de CT-e foi concluída.")

    def finalizar_erro(erro):
        executando["valor"] = False
        btn_executar.configure(state="normal", text="▶ Executar rotina")
        status_var.set("Erro na execução")
        messagebox.showerror("Erro na rotina", str(erro))

    def executar_em_thread():
        senha = senha_var.get().strip()
        if not senha:
            messagebox.showwarning("Senha obrigatória", "Informe a senha do certificado digital.")
            entrada_senha.focus_set()
            return
        if executando["valor"]:
            return

        executando["valor"] = True
        btn_executar.configure(state="disabled", text="Executando...")
        status_var.set("Executando rotina... acompanhe os logs abaixo")
        area_logs.configure(state="normal")
        area_logs.delete("1.0", "end")
        area_logs.configure(state="disabled")

        def alvo():
            stdout_original = sys.stdout
            stderr_original = sys.stderr
            sys.stdout = _ConsoleParaTela(fila_logs)
            sys.stderr = _ConsoleParaTela(fila_logs)
            try:
                executar_rotina(senha)
            except Exception as e:
                janela.after(0, lambda erro=e: finalizar_erro(erro))
            else:
                janela.after(0, finalizar_ok)
            finally:
                sys.stdout = stdout_original
                sys.stderr = stderr_original

        threading.Thread(target=alvo, daemon=True).start()

    btn_executar.configure(command=executar_em_thread)
    janela.after(120, consumir_logs)
    entrada_senha.focus_set()
    janela.mainloop()


if __name__ == "__main__":
    # Use --cli para manter o modo antigo no terminal/agendador.
    import sys
    if "--cli" in sys.argv:
        main()
    else:
        abrir_interface()
