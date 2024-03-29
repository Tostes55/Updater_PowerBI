import tkinter as tk
import requests
import logging
import locale
import pandas as pd
from token_movidesk import token_api
from tkinter import *
from tkinter import messagebox
from tkcalendar import *


locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')

# Configura o logging
logging.basicConfig(filename='log_app.txt', filemode='w', format='%(name)s - %(levelname)s - %(message)s', level=logging.INFO)


class Ambiente:
    def __init__(self):
        self.data_inicio = None
        self.data_final = None
        self.token = token_api
        self.ownerTeam = 'ATENDIMENTO EXTERNO'

    def get_data(self):
        url = 'https://api.movidesk.com/public/v1/tickets'

        query_params = [
            '$select=id,ownerTeam,serviceFirstLevel,serviceSecondLevel,serviceThirdLevel,status,origin,resolvedIn',
            '$orderby=resolvedIn',
            '$expand=Clients($expand=Organization),owner',
            f'token={self.token}',
            f"$filter=resolvedIn ge {self.data_inicio} and resolvedIn le {self.data_final} and ownerTeam eq '{self.ownerTeam}'"
        ]

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

    def save_as_xlsx(self, data, filename='output.xlsx'): 
        if data is not None:
            try:
                if isinstance(data, list) and len(data) > 0:
                    # Supondo que cada item na lista seja um dicionário
                    df = pd.DataFrame(data)  
                    
                    # Salvar como XLSX
                    df.to_excel(filename, index=False)
                    logging.info(f'Data saved to {filename}')
                    return True  # Indica que a operação de salvamento foi bem-sucedida
                else:
                    logging.error('Nenhum dado válido para salvar')
            except Exception as ex:
                logging.error(f'Error saving data to {filename}: {ex}')
        else:
            logging.error('Nenhum dado para salvar')
        return False  # Indica que não houve dados para salvar ou ocorreu um erro

    def update_data(self):
        selected_date_entry = start_date_entry.get_date()
        formatted_date_entry = selected_date_entry.strftime("%Y-%m-%dT%H:%M:%S.00z")

        selected_date_final = end_date_entry.get_date()
        formatted_date_final = selected_date_final.strftime("%Y-%m-%dT%H:%M:%S.00z")

        self.data_inicio = formatted_date_entry
        self.data_final = formatted_date_final

        logging.info('Comando get pronto para ser executado')

        data = self.get_data()
        if data:
            messagebox.showinfo("Sucesso", "Processo finalizado com sucesso!")
            print(f"Os dados obtidos foram: {data}")

            # Chama o método save_as_xlsx para salvar os dados
            saved_successfully = self.save_as_xlsx(data)
            if saved_successfully:
                messagebox.showinfo("Sucesso", "Dados salvos com sucesso!")
            else:
                messagebox.showerror("Erro", "Falha ao salvar os dados. Verifique o log para mais informações.")
        else:
            messagebox.showerror("Erro", "Falha ao obter os dados. Verifique o log para mais informações.")

def update_excel(data):
    try:
        # Carregar o arquivo XLSX existente em um DataFrame
        df_existing = pd.read_excel('output.xlsx')
        
        # Criar um DataFrame com os novos dados
        df_new = pd.DataFrame(data['result'])  # Supondo que 'result' seja a chave dos dados retornados

        # Verificar as colunas disponíveis nos DataFrames
        print("Colunas em df_existing:", df_existing.columns)
        print("Colunas em df_new:", df_new.columns)

        # Verificar os conjuntos de IDs
        existing_ids = set(df_existing['id']) if 'id' in df_existing.columns else set()
        new_ids = set(df_new['id']) if 'id' in df_new.columns else set()

        print("Existing IDs:", existing_ids)
        print("New IDs:", new_ids)

        # Adicionar apenas os novos dados ao DataFrame existente
        df_to_append = df_new[~df_new['id'].isin(existing_ids)]

        # Verificar se existem novos dados para adicionar
        if not df_to_append.empty:
            # Adicionar os novos dados ao DataFrame existente
            df_updated = pd.concat([df_existing, df_to_append], ignore_index=True)

            # Salvar o DataFrame atualizado de volta no arquivo XLSX
            df_updated.to_excel('output.xlsx', index=False)
            print("Dados atualizados com sucesso no arquivo XLSX!")
        else:
            print("Não há novos dados para adicionar.")
    except Exception as e:
        print(f"Erro ao atualizar dados no arquivo XLSX: {e}")

# Chamada da função para atualizar o arquivo XLSX com os novos dados
        update_excel(data)

# Cria uma janela raiz
root = tk.Tk()
root.title('Updater Power BI')
root.geometry("400x400")
root.configure(bg='light sea green')

start_date_label = tk.Label(root, text='Updater Power BI', bg='light sea green', font=("Arial",18))
start_date_label.place(x=110, y=100)

start_date_label = tk.Label(root, text='Data Inicial:', bg='light sea green')
start_date_label.place(x=100, y=160)

start_date_entry = DateEntry(root,date_pattern='dd/mm/yyyy')
start_date_entry.place(x=180, y=160)

end_date_label = tk.Label(root, text='Data Final:', bg='light sea green')
end_date_label.place(x=100, y=190)

end_date_entry = DateEntry(root,date_pattern='dd/mm/yyyy') 
end_date_entry.place(x=180, y=190)


ambiente = Ambiente()

def update_data():
    selected_date_entry = start_date_entry.get_date()
    formatted_date_entry = selected_date_entry.strftime("%Y-%m-%dT%H:%M:%S.00z")

    selected_date_final = end_date_entry.get_date()
    formatted_date_final = selected_date_final.strftime("%Y-%m-%dT%H:%M:%S.00z")

    ambiente.data_inicio = formatted_date_entry
    ambiente.data_final = formatted_date_final

    logging.info('Comando get pronto para ser executado')

    ambiente.update_data()

update_button = tk.Button(root, text='Atualizar', command=update_data)
update_button.place(x=170, y=250)

root.mainloop()
