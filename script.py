import os
import pandas as pd
from datetime import datetime
from tkinter import filedialog, messagebox
import tkinter as tk

def converter_excel_txt(pasta):

    for arquivo in os.listdir(pasta):
        if arquivo.startswith("RE_Química") and arquivo.endswith(('.xls', '.xlsx')):
            caminho_arquivo = os.path.join(pasta, arquivo)
            nome_base = os.path.splitext(arquivo)[0]

            # Carregar os dados do Excel
            df = pd.read_excel(caminho_arquivo, sheet_name='Plan1', header=None)

            # Obter colunas 
            colunas_resultados= {
                "M.O.": (6, "01MO", "g.dm-3"), "pH CaCl²": (5, "01PHCACl2", ""), "P": (7, "01PRES", "mg.dm-3"),
                "K": (12, "01K", "mmolc.dm-3"), "Ca": (9, "01CA", "mmolc.dm-3"), "Mg": (10, "01MG", "mmolc.dm-3"),
                "Al": (13, "01AL", "mmolc.dm-3"), "H+Al SMP": (14, "01H+AL", "mmolc.dm-3"), "S": (8, "01S", "mg.dm-3"),
                "S.B": (15, "01SB", "mmolc.dm-3"), "CTC": (16, "01CTC", "mmolc.dm-3"), 
                "Sat. Bases V%": (17, "01V", "%"), "Sat. Al m%": (18, "01M", "%")
            }

            # Nome do arquivo TXT
            data_atual = datetime.now().strftime("%d_%m_%Y")
            nome_txt = f"{nome_base}_{data_atual}.txt"
            caminho_txt = os.path.join(pasta, nome_txt)

            with open(caminho_txt, "w") as txt:
                txt.write(f"p\t\tRelatório Técnico\n")

                
                for index, row in df.iterrows():
                    if pd.notna(row[1]) and any(keyword in str(row[1]).upper() for keyword in ["FAZENDA", "FAZEBDA"]): # Identificar linha com fazenda
                        nome_fazenda = row[1]  
                        numero_amostra = row[0]  
                        data = datetime.now().strftime("%d%m%Y")

                        # Escrever dados da fazenda
                        txt.write(f"f\t{nome_fazenda}\n")
                        txt.write(f"a\t{numero_amostra}\t2025\t{data}\n")

                        
                        valores = {}
                        for nome_resultado, (coluna, identificador, unidade) in colunas_resultados.items():
                            valor = row[coluna] if not pd.isnull(row[coluna]) else 0
                            try:
                                valor = float(str(valor).replace(",", "."))
                            except ValueError:
                                valor = 0
                            valores[identificador] =valor
                            valor_txt = f"{valor:.2f}".replace(".", ",")
                            txt.write(f"r\t{identificador}\t{unidade}\t{valor_txt}\n")

                        # Cálculos derivados
                        valores["01H"] = valores["01H+AL"] - valores["01AL"]
                        valores["01KCTC"] = (valores["01K"] / valores["01CTC"] * 100) if valores["01CTC"] else 0
                        valores["01CACTC"] = (valores["01CA"] / valores["01CTC"] * 100) if valores["01CTC"] else 0
                        valores["01MGCTC"] = (valores["01MG"] / valores["01CTC"] * 100) if valores["01CTC"] else 0
                        valores["01CAMG"] = (valores["01CA"] / valores["01MG"] * 100) if valores["01MG"] else 0

                       
                        derivados = {
                            "01H": "mmolc.dm-3", "01KCTC": "%", "01CACTC": "%", 
                            "01MGCTC": "%", "01CAMG": "%"
                        }
                        for identificador, unidade in derivados.items():
                            valor = valores[identificador]
                            valor_txt = f"{valor:.2f}".replace(".", ",")
                            txt.write(f"r\t{identificador}\t{unidade}\t{valor_txt}\n")

            print(f"Arquivo {nome_txt} gerado com sucesso!")

def escolher_pasta():
    pasta = filedialog.askdirectory(title="Selecione a pasta onde estão os arquivos Excel")
    if pasta:
        if os.path.isdir(pasta):
            converter_excel_txt(pasta)
            messagebox.showinfo("Sucesso", f"Processamento concluído! Os arquivos foram salvos na pasta {pasta}.")
        else:
            messagebox.showerror("Erro", "Caminho inválido. Por favor, tente novamente.")
    else:
        messagebox.showwarning("Aviso", "Nenhuma pasta foi selecionada.")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  
    escolher_pasta()


