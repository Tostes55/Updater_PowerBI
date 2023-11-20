import pandas as pd
import datetime
import requests
import tkinter as tk
from tkinter import filedialog, Text
import os



#https://api.movidesk.com/public/v1/tickets?
# $select=id,ownerTeam,serviceFirstLevel,serviceSecondLevel,serviceThirdLevel,status,origin,resolvedIn&
# $orderby=resolvedIn&$expand=Clients ($expand=Organization),owner&token=4913d642-488b-4743-a817-b09a3166541a&
# $filter=resolvedIn ge 2023-10-01T00:00:00.00z and resolvedIn le 2023-10-31T23:59:00.00z and ownerTeam eq 'ATENDIMENTO EXTERNO'

#data_inicio = '2023-10-01T00:00:00.00z'
#data_fim = '2023-10-01T23:59:00.00z'
#token = '4913d642-488b-4743-a817-b09a3166541a'
#ownerTeam = 'ATENDIMENTO EXTERNO'

def get_data():
    
# Definindo as variáveis de filtro
    data_inicio = simpledialog.askstring(title="Data Inicio", prompt="Insira a data inicial (dd/mm/aaaa):")
    data_final = simpledialog.askstring(title="Data Final", prompt="Insira a data final (dd/mm/aaaa):")
    token = '4913d642-488b-4743-a817-b09a3166541a'
    ownerTeam = 'ATENDIMENTO EXTERNO'

    # Executando o GET da Movidesk
    response = requests.get(f'https://api.movidesk.com/public/v1/tickets?$select=id,ownerTeam,serviceFirstLevel,serviceSecondLevel,serviceThirdLevel,status,origin,resolvedIn&$orderby=resolvedIn&$expand=Clients ($expand=Organization),owner&token={token}&$filter=resolvedIn ge ({data_inicio}) and resolvedIn le ({data_fim}) and ownerTeam eq ({ownerTeam})')    
    data = response.json()

    if get_data.response.status_code == 200:
        tk.messagebox(response.json())
    else:
        tk.MessageBox(f"Error: {response.status_code}")

    # Transformando os dados em um Data Frame do Pandas
    df = pd.DataFrame(data)

    # Salvando o Dataframe em XLSX
    df.to_excel('BD.xlsx')

# Criando uma janela
root = tk.Tk()

canvas = tk.Canvas(root,height=600,width=400, bg="#263d42")
canvas.pack()
frame =tk.Frame(root,bg="#3e646c")
frame.place(relwidth=0.8, relheight=0.8, rely=0.1, relx=0.1)

# Adicionando os campos de datas

start_date_label = tk.Label(root, text='Data Inicial:')
start_date_label.pack()
start_date_entry = tk.Entry(get_data.data_inicio)
start_date_entry.pack()

end_date_label = tk.Label(root, text='Data Final:')
end_date_label.pack()
end_date_entry = tk.Entry(get_data.data_final)
end_date_entry.pack()

# Adicione um botão chamado Atualizar
update_button = tk.Button(root, text='Atualizar', command=get_data)
update_button.pack()

# Exiba a janela e aguarde eventos
root.mainloop()


# Imprimindo o resultado da requisição
