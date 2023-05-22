
import os
import csv
from docx import Document


if __name__ == "__main__":

    documento = Document("modelo_informacao.docx")
    imagem = open(os.path.join(os.getcwd(),"logo.png"), "r")


    with open(os.path.join(os.getcwd(), "Organização das salas.csv"), "r") as arquivo:
        arquivo_csv = csv.reader(arquivo, delimiter = ";")

        for linha in arquivo_csv:

            sala = str(linha[1])
            projetor = str(linha[2])
            patrimonio = str(linha[3])
            capacidade = str(linha[4])

            for paragrafo in documento.paragraphs:

        
                referencia = {
                "XXXX" : sala ,
                "XXXXXXXXX" : projetor,
                "XXXXXXXXXXX" : patrimonio,
                "XX" : capacidade,
            }
        
                for paragrafo in documento.paragraphs:
                    for codigo in referencia:
                        valor = referencia[codigo]
                        paragrafo.text = paragrafo.text.replace(codigo, valor)

                    documento.save(f"Sala - {sala}.docx")
