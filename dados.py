import os
import xml.etree.ElementTree as ET
import sys




def extrair_dados(xml_path):
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Inicializa campos padrão
    seguradora = ""
    marca = ""
    veiculo = ""
    placa = ""
    funilaria = 0.0
    pintura = 0.0
    mecanica = 0.0
    montagem_desmontagem = 0.0
    valor_mao_de_obra = 0.0
    valor_pecas = 0.0
    franquia = 0.0
    valor_total_liquido_geral = 0.0
    servicos_terceirizados = []

    # Detecta estrutura do XML
    if root.tag == "IFX":  # Novo tipo de XML
        seguradora = "porto/Azul"
        dados_orc = root.find("dadosOrcamento")
        
        
        veiculo = dados_orc.findtext("descricaoModelo", "")
        
        veiculo_base = dados_orc.findtext("descricaoModelo", "")
        palavras = veiculo_base.split()
        veiculo = " ".join(palavras[1:3])
        
        
        placa = dados_orc.findtext("licencaDoVeiculo", "")
        marca = veiculo.split()[0] if veiculo else ""
        

        def parse_float(text):
            return float(text.replace(".", "").replace(",", ".")) if text else 0.0

        mao_obra = root.find("maoDeObra")
        if mao_obra is not None:
            pintura = parse_float(mao_obra.findtext("valorMaoObraPintura", "0"))
            funilaria = parse_float(mao_obra.findtext("valorMaoObraFunilaria", "0"))
            mecanica = parse_float(mao_obra.findtext("valorMaoObraMecanica", "0"))
            montagem_desmontagem = (
                parse_float(mao_obra.findtext("valorMaoObraAcabamento", "0")) +
                parse_float(mao_obra.findtext("valorMaoObraTapecaria", "0")) +
                parse_float(mao_obra.findtext("valorMaoObraEletrica", "0"))
            )
            valor_mao_de_obra = parse_float(mao_obra.findtext("valorTotalServicos", "0")) 
            franquia = parse_float(mao_obra.findtext("valorFranquia", "0"))

        faturamento = root.find("faturamento")
        valor_total_liquido_geral = parse_float(faturamento.findtext("valorFaturar", "0"))

        pecas = root.find("pecasTrocadas")
        if pecas is not None:
         for peca in pecas.findall("peca"):
             if peca.findtext("pecaFornecida", "") == "false":
              preco = parse_float(peca.findtext("precoLiquido", "0"))
              valor_pecas += preco

            
        servicos_node = root.find("servicosTerceiros")
        if servicos_node is not None:
            for servico in servicos_node.findall("servico"):
                nome_servico = servico.findtext("descricaoServico", "").strip()
                preco = parse_float(servico.findtext("valorLiquido", "0"))
                servicos_terceirizados.append(f"{nome_servico} \n R$ {preco:.2f} \n\n")

    else:  # Estrutura antiga
        seguradoraSup = root.findtext("seguradora/nome", default="")
        seguradora = seguradoraSup.split()[0] if seguradoraSup else ""
        marca = root.findtext("veiculo/marca", default="")
        veiculo_int = root.findtext("veiculo/nome_veiculo", default="")
        veiculo = veiculo_int.split()[0] if veiculo_int else ""
        placa = root.findtext("veiculo/placa", default="")

        mao_de_obra_node = root.find("resumo_geral/totais_em_impacto")
        if mao_de_obra_node is not None:
            pintura = float(mao_de_obra_node.findtext("pintura/valor", "0"))
            funilaria = float(mao_de_obra_node.findtext("funilaria/valor", "0")) + float(
                mao_de_obra_node.findtext("reparacao/valor", "0"))
            montagem_desmontagem = (
                float(mao_de_obra_node.findtext("vidracaria/valor", "0")) +
                float(mao_de_obra_node.findtext("tapecaria/valor", "0")) +
                float(mao_de_obra_node.findtext("eletrica/valor", "0")))
            mecanica = float(mao_de_obra_node.findtext("mecanica/valor", "0"))

        for item in root.findall("itens_orcamento/item"):
            tipo_item = item.findtext("tipo_item", "")
            fornecimento = item.findtext("fornecimento", "")
            if tipo_item == "Serviço" and fornecimento == "Oficina":
                nome_servico = item.findtext("nome", "")
                preco = float(item.findtext("preco_liquido", "0") or 0)
                servicos_terceirizados.append(f"{nome_servico} \n R$ {preco:.2f} \n\n")

        total_node = root.find("total_do_orcamento")
        if total_node is not None:
            valor_pecas = float(total_node.findtext("valor_pecas_pela_oficina", "0"))
            valor_mao_de_obra = float(total_node.findtext("valor_liquido_mao_de_obra", "0"))
            franquia = float(total_node.findtext("valor_franquia", "0"))
            valor_total_liquido_geral = float(total_node.findtext("valor_total_liquido_geral", "0"))

    return {
        "seguradora:": seguradora,
        "marca:": marca,
        "veiculo:": veiculo,
        "placa:": placa,
        "funilaria:": funilaria,
        "pintura:": pintura,
        "mecanica:": mecanica,
        "montagem desmontagem:": montagem_desmontagem,
        "Serviços tercerizados: \n": "".join(servicos_terceirizados),
        "Total mão de obra:": valor_mao_de_obra,
        "pecas oficina:": valor_pecas,
        "franquia:": franquia,
        "total_liquido_geral:": valor_total_liquido_geral,
    }


pasta = "./orcamentos"
arquivos_xml = [arquivo for arquivo in os.listdir(pasta) if arquivo.endswith(".xml")]
