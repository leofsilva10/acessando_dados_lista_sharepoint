# -*- coding: utf-8 -*-

import sys
# Bibliotecas Sharepoint: "pip install shareplum"
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
# Autenticação Sharepoint (On Premise)
from requests_ntlm import HttpNtlmAuth

#############################################
# CONECTANDO PYTHON COM LISTA DO SHAREPOINT #
#############################################

# Obter informações de autenticação
username = '' # Usuário com acesso ao sharepoint.
password = '' # Senha do usuário com acesso ao sharepoint.
site_url = "" # Site do sharepoint
list_name = "" # Nome da lista do sharepoint
if len(sys.argv) < 4:
    print('Processo interrompido! Informe os parâmetros corretamente, por favor.')
    sys.exit()
else:    
    username = sys.argv[1]
    password = sys.argv[2]
    site_url = sys.argv[3]
    list_name = sys.argv[4]

# Autenticação Sharepoint (On Premise)
cred = HttpNtlmAuth(username, password)
site = Site(site_url, version=Version.v365, auth=cred)
sp_list = site.List(list_name)

# Atualizar item na lista
m_data = data=[{'ID':'1','Título': 'Título atualizado'}]    
sp_list.UpdateListItems(data=m_data, kind='Update')

# Adicionar item à lista
m_data = data=[{'Título': 'Novo título'}]      
sp_list.UpdateListItems(data=m_data, kind='New')

# Excluir itens da lista
m_data = data=['12','13']
sp_list.UpdateListItems(data=m_data, kind='Delete')

# Listar itens da lista
for item in sp_list.GetListItems():
    print(item["ID"], item["Título"])

