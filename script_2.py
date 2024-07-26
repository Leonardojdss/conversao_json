import pandas as pd
import json
from datetime import datetime

# Função para ler cada planilha e convertê-la num dictionary
def read_excel_to_dict(file_path):
    excel_data = pd.ExcelFile(file_path)
    data = {}
    for sheet_name in excel_data.sheet_names:
        data[sheet_name] = excel_data.parse(sheet_name).to_dict(orient='records')
    return data

# Função para converter colunas em uma lista de dicionários
def convert_columns_to_list(offer, prefixes):
    items = []
    columns = {prefix: [col for col in offer.keys() if col.startswith(prefix)] for prefix in prefixes}
    for col_set in zip(*columns.values()):
        item = {prefix: offer[col] for prefix, col in zip(prefixes, col_set) if offer.get(col)}
        if len(item) == len(prefixes):
            items.append(item)
    return items

# Funções criar lista dinamica
def convert_formas_pagamento(offer):
    return convert_columns_to_list(offer, ["formaPagamento", "descontoPagamento"])

def convert_lista_de_promocoes(offer):
    return convert_columns_to_list(offer, ["descricaoPromocao", "tempoDesconto", "descontoPromocao"])

def convert_modalidades_recarga(offer):
    return convert_columns_to_list(offer, ["valorRecarga", "validadeRecarga", "beneficioRecarga"])

def convert_pontos(offer):
    return convert_columns_to_list(offer, ["tipo", "numeroPontos", "pontoAdicional"])

# Função principal para leitura do Excel e escrita do JSON
def generate_json_from_excel(excel_path, json_path):
    # Lendo os dados do Excel
    data = read_excel_to_dict(excel_path)

    # Extraindo CNPJ
    for offer in data['offers']:
        # Construindo os dicionários
        cnpj = {    
        "cnpj": str(offer.get("cnpj", "")),
        }
    
    # Construção do JSON
    json_data = {
        "dataUltimaAtualizacaoArquivo": datetime.now().strftime("%d/%m/%Y"),
        "cnpj": cnpj,
        "ofertas": []
    }

    for offer in data['offers']:
        # Construindo os dicionários
        custo_inicial = {
            "adesao": str(offer.get("adesao", "")),
            "instalacao": str(offer.get("instalacao", "")),
            "equipamento": str(offer.get("equipamento", ""))
        }
        
        fidelizacao = {
            "tempoFidelizacao": offer.get("tempoFidelizacao", ""),
            "descontoFidelizacao": offer.get("descontoFidelizacao", ""),
            "tempoDesconto": offer.get("tempoDesconto", ""),
            "beneficioFidelizacao": offer.get("beneficioFidelizacao", ""),
            "multaFidelizacao": str(offer.get("multaFidelizacao", ""))
        }
        
        franquia_voz = {
            "localFixoOnNet": offer.get("localFixoOnNet", ""),
            "localFixoOffNet": offer.get("localFixoOffNet", ""),
            "localMovelOnNet": offer.get("localMovelOnNet", ""),
            "localMovelOffNet": offer.get("localMovelOffNet", ""),
            "fixoLdnOnNet": offer.get("fixoLdnOnNet", ""),
            "fixoLdnOffNet": offer.get("fixoLdnOffNet", ""),
            "movelLdnOnNet": offer.get("movelLdnOnNet", ""),
            "movelLdnOffNet": offer.get("movelLdnOffNet", ""),
            "ldi": offer.get("ldi", "")
        }
        
        # Convertendo listaPUC para lista de strings
        lista_PUC = offer.get("listaPUC", "")
        if isinstance(lista_PUC, str):
            lista_PUC = [puc.strip() for puc in lista_PUC.split(",")] if lista_PUC else []
        
        STFC = {
            "listaPUC": lista_PUC,
            "franquiaVoz": franquia_voz, 
            "condicoesAposConsumoFranquia": offer.get("condicoesAposConsumoFranquia", "")
        }

        # Convertendo listaAppIsentos para lista de strings
        listaAppsFranquiaEspecial = offer.get("listaAppsFranquiaEspecial", "")
        if isinstance(listaAppsFranquiaEspecial, str):
            listaAppsFranquiaEspecial = [puc.strip() for puc in listaAppsFranquiaEspecial.split(",")] if listaAppsFranquiaEspecial else []
        
        franquia_dados = {
            "unidadeFranquia": offer.get("unidadeFranquia", ""),
            "franquia": offer.get("franquia", ""),
            "listaAppsFranquiaEspecial": listaAppsFranquiaEspecial,
            "unidadeFranquiaEspecial": offer.get("unidadeFranquiaEspecial", ""),
            "franquiaEspecial": offer.get("franquiaEspecial", "")
        }

        cobranca_tipo = {
            "tipoCobranca": offer.get("tipoCobranca", ""),
            "detalhesCobranca": offer.get("detalhesCobranca", "")
        }
        
        franquia_SMS = {
            "onNet": offer.get("onNet", ""),
            "offNet": offer.get("offNet", "")
        }
        
        dependentes = {
            "quantidade": offer.get("quantidade", ""),
            "valor": offer.get("valor", ""),
            "compartilhamento": offer.get("compartilhamento", "")
        }

        # Convertendo listaAppIsentos para lista de strings
        listaAppIsentos = offer.get("SMP_listaAppIsentos", "")
        if isinstance(listaAppIsentos, str):
            listaAppIsentos = [puc.strip() for puc in listaAppIsentos.split(",")] if listaAppIsentos else []

        # Convertendo listaAppIsentos para lista de strings
        SMP_listaSVA = offer.get("SMP_listaSVA", "")
        if isinstance(SMP_listaSVA, str):
            SMP_listaSVA = [puc.strip() for puc in SMP_listaSVA.split(",")] if SMP_listaSVA else []

        SMP = {
            "modalidadePagamento": offer.get("modalidadePagamento", ""),
            "validadePacote": offer.get("validadePacote", ""),
            "franquiaDados": franquia_dados,
            "listaAppIsentos": listaAppIsentos,
            "listaSVA": SMP_listaSVA,
            "cobrancaTipo": cobranca_tipo,
            "franquiaVoz": franquia_voz,
            "franquiaSMS": franquia_SMS,
            "condicoesAposConsumoFranquia": offer.get("condicoesAposConsumoFranquia", ""),
            "condicoesAposValidadePacote": offer.get("condicoesAposValidadePacote", ""),
            "modalidadesRecarga": convert_modalidades_recarga(offer),
            "roamingNacional": offer.get("roamingNacional", ""),
            "roamingInternacional": offer.get("roamingInternacional", ""),
            "dependentes": dependentes
        }
        
        velocidade = {
            "download": offer.get("download", ""),
            "unidadeDownload": offer.get("unidadeDownload", ""),
            "downloadMinGarantida": offer.get("downloadMinGarantida", ""),
            "unidadeDownloadMinGarantida": offer.get("unidadeDownloadMinGarantida", ""),
            "upload": offer.get("upload", ""),
            "unidadeUpload": offer.get("unidadeUpload", "")
        }

        # Convertendo listaTecnologia para lista de strings
        SCM_listaTecnologia = offer.get("SCM_listaTecnologia", "")
        if isinstance(SCM_listaTecnologia, str):
            SCM_listaTecnologia = [puc.strip() for puc in SCM_listaTecnologia.split(",")] if SCM_listaTecnologia else []

        # Convertendo listaSVA para lista de strings
        SCM_listaSVA = offer.get("SCM_listaSVA", "")
        if isinstance(SCM_listaSVA, str):
            SCM_listaSVA = [puc.strip() for puc in SCM_listaSVA.split(",")] if SCM_listaSVA else []        

        SCM = {
            "wifiIncluso": offer.get("wifiIncluso", ""),
            "listaTecnologia": SCM_listaTecnologia,
            "velocidade": velocidade,
            "listaSVA": SCM_listaSVA
        }

        # Convertendo listaTecnologia para lista de strings
        areasAbrangencia = offer.get("areasAbrangencia", "")
        if isinstance(areasAbrangencia, str):
            areasAbrangencia = [puc.strip() for puc in areasAbrangencia.split(",")] if areasAbrangencia else []

        # Convertendo listaTecnologia para lista de strings
        SEAC_listaTecnologia = offer.get("SEAC_listaTecnologia", "")
        if isinstance(SEAC_listaTecnologia, str):
            SEAC_listaTecnologia = [puc.strip() for puc in SEAC_listaTecnologia.split(",")] if SEAC_listaTecnologia else []

        # Convertendo listaTecnologia para lista de strings
        SEAC_listaCanais = offer.get("SEAC_listaCanais", "")
        if isinstance(SEAC_listaCanais, str):
            SEAC_listaCanais = [puc.strip() for puc in SEAC_listaCanais.split(",")] if SEAC_listaCanais else []

        # Convertendo listaTecnologia para lista de strings
        SEAC_listaCanaisAvulsos = offer.get("SEAC_listaCanaisAvulsos", "")
        if isinstance(SEAC_listaCanaisAvulsos, str):
            SEAC_listaCanaisAvulsos = [puc.strip() for puc in SEAC_listaCanaisAvulsos.split(",")] if SEAC_listaCanaisAvulsos else []      

        # Convertendo listaTecnologia para lista de strings
        SEAC_listaSVA = offer.get("SEAC_listaSVA", "")
        if isinstance(SEAC_listaSVA, str):
            SEAC_listaSVA = [puc.strip() for puc in SEAC_listaSVA.split(",")] if SEAC_listaSVA else [] 
        
        SEAC = {
            "listaTecnologia": SEAC_listaTecnologia,
            "multiPlataforma": offer.get("multiPlataforma", ""),
            "dvr": offer.get("dvr", ""),
            "pontos": convert_pontos(offer),
            "listaCanais": SEAC_listaCanais,
            "listaCanaisAvulsos": SEAC_listaCanaisAvulsos,
            "listaSVA": SEAC_listaSVA
        }
        
        offer_json = {
            "identificadorUnico": offer.get("identificadorUnico", ""),
            "tipoOferta": offer.get("tipoOferta", ""),
            "nomeOferta": offer.get("nomeOferta", ""),
            "codigoOferta": str(offer.get("codigoOferta", "")),
            "custoInicial": custo_inicial,
            "etiquetaOferta": offer.get("etiquetaOferta", ""),
            "linkSite": offer.get("linkSite", ""),
            "dataInicioOferta": offer.get("dataInicioOferta", ""),
            "dataFimOferta": offer.get("dataFimOferta", ""),
            "fidelizacao": fidelizacao,
            "formasPagamento": convert_formas_pagamento(offer),
            "destaqueOferta": offer.get("areasAbrangencia", ""),
            "areasAbrangencia": areasAbrangencia,
            "notasExtras": offer.get("notasExtras", ""),
            "focoVenda": offer.get("focoVenda", ""),
            "regOferta": offer.get("regOferta", ""),
            "modoEquipamento": offer.get("modoEquipamento", ""),
            "precoSemDescontos": offer.get("precoSemDescontos", ""),
            "listaPromocoes": convert_lista_de_promocoes(offer),
            "beneficiosOfertaConjunta": offer.get("beneficiosOfertaConjunta", ""),
            "STFC": STFC,
            "SMP": SMP,
            "SCM": SCM,
            "SEAC": SEAC 
        }
        json_data["ofertas"].append(offer_json)
    
    # Salvando o JSON em arquivo
    with open(json_path, 'w') as json_file:
        json.dump(json_data, json_file, indent=4, ensure_ascii=False)

# Caminhos dos arquivos (Excel de entrada e JSON de saída)
excel_path = 'anatel_ofertas.xlsx'
json_path = 'output_file.json'

# Chamando a função principal para gerar o JSON
generate_json_from_excel(excel_path, json_path)
print("JSON gerado com sucesso!")
