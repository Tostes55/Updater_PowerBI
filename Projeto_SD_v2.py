import tkinter as tk
import requests
import logging
import locale
import pandas as pd
from datetime import datetime
from tkinter import *
from tkinter import messagebox
from tkcalendar import *
from datetime import datetime
from urllib.parse import quote

locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')

# Configura o logging
logging.basicConfig(filename='log_app.txt', filemode='w', format='%(name)s - %(levelname)s - %(message)s', level=logging.INFO)


class Ambiente:
    def __init__(self):
        self.data_inicio = None
        self.data_final = None
        self.token = '4913d642-488b-4743-a817-b09a3166541a'
        self.ownerTeam = 'ATENDIMENTO EXTERNO'

    def get_data(self):
        
        # Executando o GET da Movidesk
        url = 'https://api.movidesk.com/public/v1/tickets'

        # Construindo a query string manualmente
        query_params = [
            '$select=id,ownerTeam,serviceFirstLevel,serviceSecondLevel,serviceThirdLevel,status,origin,resolvedIn',
            '$orderby=resolvedIn',
            '$expand=Clients($expand=Organization),owner',
            f'token={self.token}',
            f"$filter=resolvedIn ge {self.data_inicio} and resolvedIn le {self.data_final} and ownerTeam eq '{self.ownerTeam}'"
            #f"$filter=resolvedIn ge '{self.data_inicio}' and resolvedIn le '{self.data_final}' and ownerTeam eq '{self.ownerTeam}'"
        ]

        #data_inicio = '2023-10-01T00:00:00.00z'
        #data_fim = '2023-10-01T23:59:00.00z'


        # Unindo os parâmetros em uma única string
        query_string = '&'.join(query_params)

        full_url = f"{url}?{query_string}"

        logging.info(f'Making GET request to URL: {full_url}')

        try:
            response = requests.get(full_url)
            response.raise_for_status()  # Lança uma exceção para códigos de status de erro
            data = response.json()
            return data
        except requests.exceptions.RequestException as e:
            logging.error(f'Erro na solicitação HTTP: {e}')
            return None
            

# Cria uma instância da classe Ambiente
ambiente = Ambiente()

# Cria uma janela raiz
root = tk.Tk()
root.title('Updater Power BI')
root.geometry("400x400")
root.configure(bg='light sea green')


start_date_label = tk.Label(root, text='Updater Power BI', bg='light sea green', font=("Arial",18))
start_date_label.place(x=110, y=100)

# Adicionando os campos de datas
start_date_label = tk.Label(root, text='Data Inicial:', bg='light sea green')
start_date_label.place(x=100, y=160)

start_date_entry = DateEntry(root)
start_date_entry.place(x=180, y=160)
# Obtendo a data selecionada
selected_date_entry = start_date_entry.get_date()
# Formatando a data no formato desejado (ano-mês-diaT00:00:00.00z)
formatted_date_entry = selected_date_entry.strftime("%Y-%m-%dT%H:%M:%S.00z")


end_date_label = tk.Label(root, text='Data Final:', bg='light sea green')
end_date_label.place(x=100, y=190)

end_date_entry = DateEntry(root) 
# Obtendo a data selecionada
selected_date_final = start_date_entry.get_date()
# Formatando a data no formato desejado (ano-mês-diaT00:00:00.00z)
formatted_date_final = selected_date_final.strftime("%Y-%m-%dT%H:%M:%S.00z")
end_date_entry.place(x=180, y=190)


# Adicione um botão chamado Atualizar
def update_data():
    selected_date_entry = start_date_entry.get_date()
    formatted_date_entry = selected_date_entry.strftime("%Y-%m-%dT%H:%M:%S.00z")

    selected_date_final = end_date_entry.get_date()
    formatted_date_final = selected_date_final.strftime("%Y-%m-%dT%H:%M:%S.00z")

    ambiente.data_inicio = formatted_date_entry
    ambiente.data_final = formatted_date_final

    # Registra o comando get antes de executá-lo
    logging.info('Comando get pronto para ser executado')

    data = ambiente.get_data()
    if data:
        messagebox.showinfo("Sucesso", "Processo finalizado com sucesso!")
        print(f"Os dados obtidos foram: {data}")
    else:
        messagebox.showerror("Erro", "Falha ao obter os dados. Verifique o log para mais informações.")

def save_as_xlsx(self, data, filename='output.xlsx'):
        if data is not None:
            df = pd.DataFrame(data['result'])  # Supondo que 'result' seja a chave dos dados retornados
            
            # Salvar como XLSX
            df.to_excel(filename, index=False)
            logging.info(f'Data saved to {filename}')
            ambiente.save_as_xlsx(data)
        else:
            logging.error('Nenhum dado para salvar')

update_button = tk.Button(root, text='Atualizar', command=update_data)
#update_button.grid(row=30, column=1, columnspan=1, pady=(10, 0))
update_button.place(x=170, y=250)

# Exiba a janela e aguarde eventos
root.mainloop()
