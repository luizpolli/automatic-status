from operator import ne
from numpy import float64, int64
import pandas as pd
import os, re
import colorama
from colorama import Fore, Back, Style

colorama.deinit()
colorama.init(autoreset=True)

def get_filename():
  while True:
    excel_akad_name = input("Nome da planilha AKAD: ")
    if excel_akad_name not in os.listdir():
      print(Fore.RED + Style.BRIGHT + f"\n*** Arquivo \"{excel_akad_name}\" não existe dentro da pasta {os.getcwd()}. Valide o nome do arquivo com a extensão. ***\n")
      continue
    break

  while True:
    excel_partner_name = input("Nome da planilha parceiro: ")
    if excel_partner_name not in os.listdir():
      print(Fore.RED + Style.BRIGHT + f"\n*** Arquivo \"{excel_partner_name}\" não existe dentro da pasta {os.getcwd()}. Valide o nome do arquivo com a extensão. ***\n")
      continue
    break

  while True:
    regex_new_filename = re.compile(r"^[\w\-\s]+\.(\w+)$")
    new_filename = input("Nome da nova planilha que será salva: ")
    if regex_new_filename.match(new_filename):
      if regex_new_filename.match(new_filename).group(1) != "xlsx":
        print(Fore.RED + Style.BRIGHT + "\n*** Arquivo não é em formato excel XLSX. Formato <anyname.xlsx> ***\n")
        continue
    else:
      print(Fore.RED + Style.BRIGHT + "\n*** Nome da nova planilha não é padrão com extensão xlsx. Formato <anyname.xlsx> ***\n")
      continue
    break

  return excel_akad_name, excel_partner_name, new_filename


excel_akad_name, excel_partner_name, new_filename = get_filename()



akad_table = pd.read_excel(excel_akad_name, sheet_name="PSL")
partner_table = pd.read_excel(excel_partner_name)

for sin_p in partner_table["Número do Sinistro"]:
  for sin_a in akad_table["Número do Sinistro"]:
    if (sin_p == sin_a):
      value_akad_au = float(akad_table.loc[akad_table.index[akad_table["Número do Sinistro"]==sin_p][0], "Reserva Indenização ATUAL"])
      value_partner_ax = float(partner_table.loc[partner_table.index[partner_table["Número do Sinistro"]==sin_p][0], "RESERVA ATUAL"])
      if (value_partner_ax == 0):
        akad_table.loc[akad_table.index[akad_table["Número do Sinistro"]==sin_p][0], "STATUS"] = "Caso encerrado"
        akad_table.loc[akad_table.index[akad_table["Número do Sinistro"]==sin_p][0], "RESERVA ATUAL"] = value_partner_ax     
      if (value_partner_ax != 0):
        akad_table.loc[akad_table.index[akad_table["Número do Sinistro"]==sin_p][0], "STATUS"] = "Em regulação"
        akad_table.loc[akad_table.index[akad_table["Número do Sinistro"]==sin_p][0], "RESERVA ATUAL"] = value_partner_ax

akad_table.to_excel(new_filename)

print(Fore.CYAN + Style.BRIGHT + f"\nPronto. Checar o novo arquivo criado \"{new_filename}\".\n")

colorama.deinit()