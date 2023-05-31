from xml.dom import minidom
import zipfile
import csv
from datetime import datetime

notas_zip = "nfse.zip"

# dicionário para armazenar os TOTAIS por competência
totalCompetencia = {}

with zipfile.ZipFile(notas_zip, 'r') as zip_ref:
    # verifica todos os arquivos dentro do zip
    for notas_xml in zip_ref.namelist():
        with zip_ref.open(notas_xml) as xml_file:
            xml = minidom.parse(xml_file)
            competencia = xml.getElementsByTagName("DataEmissaoRps")[0].firstChild.data
            valorServicos = float(xml.getElementsByTagName("ValorServicos")[0].firstChild.data)
            valorIss = float(xml.getElementsByTagName("ValorIss")[0].firstChild.data)
            issRetido = float(xml.getElementsByTagName("IssRetido")[0].firstChild.data)

            # verifica se a tag outrasRetencoes existe
            if xml.getElementsByTagName("OutrasRetencoes"):
                outrasRetencoes = float(xml.getElementsByTagName("OutrasRetencoes")[0].firstChild.data)
            else:
                outrasRetencoes = 0.0

            baseCalculo = float(xml.getElementsByTagName("BaseCalculo")[0].firstChild.data)
            valorLiquidoNfse = float(xml.getElementsByTagName("ValorLiquidoNfse")[0].firstChild.data)

            #separa as notas por data
            competencia_date = datetime.strptime(competencia, "%Y-%m-%d")
            competencia_mes = competencia_date.strftime("%Y-%m")

            # Verifica se a competência já está adicionada
            if competencia_mes in totalCompetencia:
                # Atualiza os totais da competência que já existe
                totalCompetencia[competencia_mes]["ValorServicos"] += valorServicos
                totalCompetencia[competencia_mes]["ValorIss"] += valorIss
                totalCompetencia[competencia_mes]["ValorIssRetido"] += issRetido
                totalCompetencia[competencia_mes]["OutrasRetencoes"] += outrasRetencoes
                totalCompetencia[competencia_mes]["BaseCalculo"] += baseCalculo
                totalCompetencia[competencia_mes]["ValorLiquidoNfse"] += valorLiquidoNfse
            else:
                # add a competência ao dicionário com os totais iniciais
                totalCompetencia[competencia_mes] = {
                    "ValorServicos": valorServicos,
                    "ValorIss": valorIss,
                    "ValorIssRetido": issRetido,
                    "OutrasRetencoes": outrasRetencoes,
                    "BaseCalculo": baseCalculo,
                    "ValorLiquidoNfse": valorLiquidoNfse
                }
                
# Gerando o arquivo CSV 
with open("notas_fiscais.csv", 'w', encoding='utf-8', newline='') as f:
    writer = csv.writer(f)
    field = ["Competencia", "ValorServicos", "ValorIss", "ValorIssRetido", "OutrasRetencoes", "BaseCalculo", "ValorLiquidoNfse"]
    writer.writerow(field)

    for competencia_mes, totais in totalCompetencia.items():
        row = [
            competencia_mes,
            totais["ValorServicos"],
            totais["ValorIss"],
            totais["ValorIssRetido"],
            totais["OutrasRetencoes"],
            totais["BaseCalculo"],
            totais["ValorLiquidoNfse"]
        ]
        writer.writerow(row)
