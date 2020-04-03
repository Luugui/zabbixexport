#!/usr/bin/python3

from pyzabbix import ZabbixAPI
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from pyfiglet import Figlet
from tqdm import tqdm
import datetime, argparse


parser = argparse.ArgumentParser(description="Extracao de dados do Zabbix")
parser.add_argument('-u', '--user', required=True, help="Usuario do Zabbix (Admin)")
parser.add_argument('-p', '--password', required=True, help="Senha do Zabbix (zabbix)")
parser.add_argument('-s', '--server', required=True, help="Endereco Zabbix (http://localhost/zabbix)")
parser.add_argument('-g', '--group', action="append", help="ID de grupos de Host")
parser.add_argument('-n', '--name', default="Zabbix", help="Nome do relatorio gerado ex: Name_DATAHOJE.xlsx")
args = vars(parser.parse_args())

f = Figlet(font='slant')
print(f.renderText("Zabbix Export All Hosts"))


# CONECTANDO AO ZABBIX
zapi = ZabbixAPI(args['server'])
zapi.login(args['user'],args['password'])
print("--> Conectado com sucesso!\n")

if args['group']:
    for g in zapi.hostgroup.get(output="extend", groupids=args['group']):
     print("--> Grupo selecionado: "+g['name'])

# CRIANDO PLANILHA
wb = Workbook()
sheet = wb.active
sheet.title = "RESUMO"
sheet['A1'] = "GRUPO"
sheet['A1'].font = Font(sz=12, bold=True)
sheet['B1'] = "HOSTS"
sheet['B1'].font = Font(sz=12, bold=True)

row=3

# SEVERIDADE
SEV = {
    "0": "Not classified",
    "1": "Information",
    "2": "Warning",
    "3": "Average",
    "4": "High",
    "5": "Disaster"
}

# LISTANDO GRUPOS
print("\n--> Listando grupos e quantidade de hosts\n")
for g in tqdm(zapi.hostgroup.get(output="extend", groupids=args['group'])):
    if "Templates" not in g['name']:
        sheet.cell(row=row, column=1).value = g['name']
        h = zapi.host.get(output="extend",groupids=g['groupid'])
        sheet.cell(row=row, column=2).value = len(h)
        #print("G: ",g['name']," N:",len(h))
        row+=1

print("\n--> Gerando relatorio!\n")
for g in tqdm(zapi.hostgroup.get(output="extend", groupids=args['group'])):
    if "Template" not in g['name']:
        GRUPO = g['name']
        if "/" in GRUPO:
            GRUPO = GRUPO.replace("/",".")
        #print('g: ',GRUPO)
        if len(GRUPO) >= 16:
            GRUPO_NAME = GRUPO.split(".")
            GRUPO = GRUPO_NAME[-1]
        wb.create_sheet(GRUPO)
        S_GRUPO = wb[GRUPO]
        S_GRUPO['A1'] = "HOST"
        S_GRUPO['A1'].font = Font(sz=12, bold=True)
        S_GRUPO['B1'] = "IP"
        S_GRUPO['B1'].font = Font(sz=12, bold=True)
        S_GRUPO['C1'] = "ITEM"
        S_GRUPO['C1'].font = Font(sz=12, bold=True)
        row=2
        for h in zapi.host.get(output="extend", groupids=g["groupid"]):
            for i in zapi.hostinterface.get(output="extend",hostids=h['hostid']):
                for item in zapi.item.get(output="extend",hostids=h['hostid']):
                        ativos = {"HOST":h['host'], "IP":i['ip'], "ITEM":item['name']}
                        #print('id: ',row," ",ativos)
                        S_GRUPO.cell(row=row, column=1).value = ativos['HOST']
                        S_GRUPO.cell(row=row, column=2).value = ativos['IP']
                        S_GRUPO.cell(row=row, column=3).value = ativos['ITEM']
                        row+=1


DATE = datetime.date.today()
NOME = args['name']+'_'+str(DATE)+'.xlsx'

wb.save(NOME)
print("\n--> Relatorio {0} Gerado!".format(NOME))
