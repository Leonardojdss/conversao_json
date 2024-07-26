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

# Função para converter formas de pagamento
def convert_formas_pagamento(offer):
    formas_pagamento = []
    col_prefix = "formaPagamento"
    desconto_prefix = "descontoPagamento"

    # Identificar todas as colunas de formas de pagamento e descontos
    forma_columns = [col for col in offer.keys() if col.startswith(col_prefix)]
    desconto_columns = [col for col in offer.keys() if col.startswith(desconto_prefix)]
    
    # Combinar as colunas de formaPagamento com descontoPagamento
    for forma_col, desconto_col in zip(forma_columns, desconto_columns):
        forma_pagamento = offer.get(forma_col)
        desconto_pagamento = offer.get(desconto_col)
        if forma_pagamento and desconto_pagamento:
            formas_pagamento.append({
                "formaPagamento": forma_pagamento,
                "descontoPagamento": desconto_pagamento
            })
    
    return formas_pagamento

# Função para converter lista de promoções
def convert_lista_de_promocoes(offer):
    lista_promocoes = []
    col_descricaoPromocao = "descricaoPromocao"
    col_tempoDesconto = "tempoDesconto"
    col_descontoPromocao = "descontoPromocao"

    # Identificar todas as colunas de lista de promoções
    descricaoPromocao_columns = [col for col in offer.keys() if col.startswith(col_descricaoPromocao)]
    tempoDesconto_columns = [col for col in offer.keys() if col.startswith(col_tempoDesconto)]
    descontoPromocao_columns = [col for col in offer.keys() if col.startswith(col_descontoPromocao)]
    
    # Combinar as colunas de descricaoPromocao, tempoDesconto e descontoPromocao
    for descri_col, tempo_col, desco_col in zip(descricaoPromocao_columns, tempoDesconto_columns, descontoPromocao_columns):
        descr_promocao = offer.get(descri_col)
        tempo_desconto = offer.get(tempo_col)
        desconto_promocao = offer.get(desco_col)
        if descr_promocao and tempo_desconto and desconto_promocao:
            lista_promocoes.append({
                "descricaoPromocao": descr_promocao,
                "tempoDesconto": tempo_desconto,
                "descontoPromocao": desconto_promocao
            })
    
    return lista_promocoes

# Função para converter modalidades Recarga
def convert_modalidades_Recarga(offer):
    modalidades_Recarga = []
    col_valorRecarga = "valorRecarga"
    col_validadeRecarga = "validadeRecarga"
    col_beneficioRecarga = "beneficioRecarga"

    # Identificar todas as colunas de lista de promoções
    valorRecarga_columns = [col for col in offer.keys() if col.startswith(col_valorRecarga)]
    validadeRecarga_columns = [col for col in offer.keys() if col.startswith(col_validadeRecarga)]
    beneficioRecarga_columns = [col for col in offer.keys() if col.startswith(col_beneficioRecarga)]
        
    # Combinar as colunas de descricaoPromocao, tempoDesconto e descontoPromocao
    for valorRecarga_col, validadeRecarga_col, validadeRecarga_col in zip(valorRecarga_columns, validadeRecarga_columns, beneficioRecarga_columns):
        valorRecarga = offer.get(valorRecarga_col)
        validadeRecarga = offer.get(validadeRecarga_col)
        beneficioRecarga = offer.get(validadeRecarga_col)
        if valorRecarga and validadeRecarga and beneficioRecarga:
            modalidades_Recarga.append({
                "valorRecarga": valorRecarga,
                "validadeRecarga": validadeRecarga,
                "beneficioRecarga": beneficioRecarga
            })
        
    return modalidades_Recarga

# Função para converter lista de pontos
def convert_pontos(offer):
    pontos = []
    col_tipo = "tipo"
    numeroPontos = "numeroPontos"
    pontoAdicional = "pontoAdicional"

    # Identificar todas as colunas de lista de promoções
    tipo_columns = [col for col in offer.keys() if col.startswith(col_tipo)]
    numeroPontos_columns = [col for col in offer.keys() if col.startswith(numeroPontos)]
    pontoAdicional_columns = [col for col in offer.keys() if col.startswith(pontoAdicional)]
    
    # Combinar as colunas de descricaoPromocao, tempoDesconto e descontoPromocao
    for tipo_col, numeroPontos_col, pontoAdicional_col in zip(tipo_columns, numeroPontos_columns, pontoAdicional_columns):
        tipo = offer.get(tipo_col)
        numeroPontos = offer.get(numeroPontos_col)
        pontoAdicional = offer.get(pontoAdicional_col)
        if tipo and numeroPontos and pontoAdicional:
            pontos.append({
                "tipo": tipo,
                "numeroPontos": numeroPontos,
                "pontoAdicional": pontoAdicional
            })
    
    return pontos

# Função principal para leitura do Excel e escrita do JSON
def generate_json_from_excel(excel_path, json_path):
    # Lendo os dados do Excel
    data = read_excel_to_dict(excel_path)
    
    # Construção do JSON conforme o manual
    json_data = {
        "dataUltimaAtualizacaoArquivo": datetime.now().strftime("%d/%m/%Y"),  # Atualize a data conforme necessário
        "cnpj": "11111111111111",  # Exemplo de CNPJ, substitua pelo valor lido do Excel
        "ofertas": []
    }

    for offer in data['offers']:  # Supondo que a planilha de ofertas se chama "offers"
        # Construindo o dicionário para custoInicial
        custo_inicial = {
            "adesao": offer.get("adesao", ""),
            "instalacao": offer.get("instalacao", ""),
            "equipamento": offer.get("equipamento", "")
        }
        
        # Construindo o dicionário para fidelizacao
        fidelizacao = {
            "tempoFidelizacao": offer.get("tempoFidelizacao", ""),
            "descontoFidelizacao": offer.get("descontoFidelizacao", ""),
            "tempoDesconto": offer.get("tempoDesconto", ""),
            "beneficioFidelizacao": offer.get("beneficioFidelizacao", ""),
            "multaFidelizacao": offer.get("multaFidelizacao", ""),
        }
        
        # Construindo a lista de formas de pagamento
        formas_pagamento = convert_formas_pagamento(offer)

        # Construindo a lista de promoções
        lista_promocoes = convert_lista_de_promocoes(offer)

        # Construindo a lista de promoções
        modalidades_Recarga = convert_modalidades_Recarga(offer)

        # Construindo a lista de promoções
        lista_pontos = convert_pontos(offer)

    for offer in data['offers']:  # Supondo que a planilha de ofertas se chama "offers"
        # Construindo o dicionário para franquiaVoz
        franquiaVoz = {
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
    
    for offer in data['offers']:  # Supondo que a planilha de ofertas se chama "offers"
        # Construindo o dicionário para STFC
        STFC = {
            "listaPUC": offer.get("listaPUC", ""),
            "franquiaVoz": franquiaVoz, 
            "condicoesAposConsumoFranquia": offer.get("condicoesAposConsumoFranquia", "")
        }

    for offer in data['offers']:  # Supondo que a planilha de ofertas se chama "offers"
        # Construindo o dicionário para franquiaDados
        franquiaDados = {
            "unidadeFranquia": offer.get("unidadeFranquia", ""),
            "franquia": offer.get("franquia", ""),
            "listaAppsFranquiaEspecial": offer.get("listaAppsFranquiaEspecial", ""),
            "unidadeFranquiaEspecial": offer.get("unidadeFranquiaEspecial", ""),
            "franquiaEspecial": offer.get("franquiaEspecial", "")
        }

    for offer in data['offers']:  # Supondo que a planilha de ofertas se chama "offers"
        # Construindo o dicionário para cobrancaTipo
        cobrancaTipo = {
            "tipoCobranca": offer.get("tipoCobranca", ""),
            "detalhesCobrança": offer.get("detalhesCobrança", "")
        }

    for offer in data['offers']:  # Supondo que a planilha de ofertas se chama "offers"
        # Construindo o dicionário para franquiaSMS
        franquiaSMS = {
            "onNet": offer.get("onNet", ""),
            "offNet": offer.get("offNet", "")
        }
    
    for offer in data['offers']:  # Supondo que a planilha de ofertas se chama "offers"
        # Construindo o dicionário para dependentes
        dependentes = {
            "quantidade": offer.get("quantidade", ""),
            "valor": offer.get("valor", ""),
            "compartilhamento": offer.get("compartilhamento", "")
        }

    for offer in data['offers']:  # Supondo que a planilha de ofertas se chama "offers"
        # Construindo o dicionário para SMP
        SMP = {
            "modalidadePagamento": offer.get("modalidadePagamento", ""),
            "validadePacote": offer.get("validadePacote", ""),
            "franquiaDados": franquiaDados,
            "listaAppIsentos": offer.get("listaAppIsentos", ""),
            "listaSVA": offer.get("listaSVA", ""),
            "cobrancaTipo": cobrancaTipo,
            "franquiaVoz": franquiaVoz,
            "franquiaSMS": franquiaSMS,
            "condicoesAposConsumoFranquia": offer.get("condicoesAposConsumoFranquia", ""),
            "condicoesAposValidadePacote": offer.get("condicoesAposValidadePacote", ""),
            "modalidadesRecarga ": modalidades_Recarga,
            "roamingNacional": offer.get("roamingNacional", ""),
            "roamingInternacional": offer.get("roamingInternacional", ""),
            "dependentes": dependentes
        }
    
    for offer in data['offers']:  # Supondo que a planilha de ofertas se chama "offers"
        # Construindo o dicionário para velocidade
        velocidade = {
            "download": offer.get("download", ""),
            "unidadeDownload": offer.get("unidadeDownload", ""),
            "downloadMinGarantida": offer.get("downloadMinGarantida", ""),
            "unidadeDownloadMinGarantida": offer.get("unidadeDownloadMinGarantida", ""),
            "upload": offer.get("upload", ""),
            "unidadeUpload": offer.get("unidadeUpload", "")
        }    
           
    
    for offer in data['offers']:  # Supondo que a planilha de ofertas se chama "offers"
        # Construindo o dicionário para SCM
        SCM = {
            "wifiIncluso": offer.get("wifiIncluso", ""),
            "listaTecnologia": offer.get("listaTecnologia", ""),
            "velocidade": velocidade,
            "listaSVA": offer.get("listaSVA", "")
        }

    for offer in data['offers']:  # Supondo que a planilha de ofertas se chama "offers"
        # Construindo o dicionário para SEAC
        SEAC = {
            "listaTecnologia": offer.get("listaTecnologia", ""),
            "multiPlataforma": offer.get("multiPlataforma", ""),
            "dvr ": offer.get("dvr ", ""),
            "pontos ":  lista_pontos,
            "listaCanais ": offer.get("listaCanais ", ""),
            "listaCanaisAvulsos ": offer.get("listaCanaisAvulsos ", ""),
            "listaSVA ": offer.get("listaSVA ", "")
        }    
        
        offer_json = {
            "identificadorUnico": offer["identificadorUnico"],
            "tipoOferta": offer["tipoOferta"],
            "nomeOferta": offer["nomeOferta"],
            "codigoOferta": offer["codigoOferta"],
            "custoInicial": custo_inicial,  # Adiciona o dicionário custoInicial
            "etiquetaOferta": offer["etiquetaOferta"],
            "linkSite": offer["linkSite"],
            "dataInicioOferta": offer["dataInicioOferta"],
            "dataFimOferta": offer["dataFimOferta"],
            "fidelizacao": fidelizacao,
            "formasPagamento": formas_pagamento,  # Adiciona a lista de formas de pagamento
            "areasAbrangencia": offer["areasAbrangencia"],
            "focoVenda": offer["focoVenda"],
            "regOferta": offer["regOferta"],
            "modoEquipamento": offer["modoEquipamento"],
            "precoSemDescontos": offer["precoSemDescontos"],
            "listaPromocoes": lista_promocoes,  # Adiciona a lista de promoções
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
