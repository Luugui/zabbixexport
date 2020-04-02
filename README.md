# Zabbix Export All Hosts

Exporta os itens monitorados de cada Servidor separando por grupo em abas numa pasta do Excel.

### Requisitos

É necessário ter os pacotes pyzabbix, openpyxl, argparse, tqdm e pyfiglet

    pip install pyzabbix openpyxl argparse pyfiglet tqdm

### Exemplo

É obrigatório passar o servidor, login e senha do qual os dados serão extraídos. 

    

> python .\resumo_ambiente.py -s http://localhost/zabbix -u Admin -p  zabbix

 O parâmetro -n [Name] é o nome que será salvo o relatório junto com a data atual
 

> python .\resumo_ambiente.py -s http://localhost/zabbix -u Admin -p zabbix -n CLIENT

O parâmetro -g [Group] serve para especificar um ou mais grupos dos quais os hosts serão extraídos. Na ausência desse parâmetro ele irá extrair de todos os grupos que tiver permissão de leitura.

> python .\resumo_ambiente.py -s http://localhost/zabbix -u Admin -p zabbix -n CLIENT -g 1 -g 2 -g 3

Usando o -h [Help] irá apresentar uma ajuda com uma breve explicação de cada parâmetro

    PS C:\Users\luisg\Documents\Scripts> python .\resumo_ambiente.py --help
    usage: resumo_ambiente.py [-h] -u USER -p PASSWORD -s SERVER [-g GROUP] [-n NAME]
    
    Extracao de dados do Zabbix
    
    optional arguments:
      -h, --help            show this help message and exit
      -u USER, --user USER  Usuario do Zabbix (Admin)
      -p PASSWORD, --password PASSWORD
                            Senha do Zabbix (zabbix)
      -s SERVER, --server SERVER
                            Endereco Zabbix (http://localhost/zabbix)
      -g GROUP, --group GROUP
                            ID de grupos de Host
      -n NAME, --name NAME  Nome do relatorio gerado ex: Name_DATAHOJE.xlsx
