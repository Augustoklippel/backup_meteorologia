import os
import time 
from datetime import datetime
from datetime import timedelta 
from os.path import exists, isfile
from pathlib import Path
import tkinter as tk
from tkinter import TclError, ttk, filedialog, simpledialog
from tkinter.messagebox import showinfo, showerror, askyesno
import shutil
from tkcalendar import DateEntry
from openpyxl import Workbook, load_workbook
import pandas as pd
import zipfile

list_db_files = ["/Torre-A1/WIFI/TorreA1_15M_BD.dat", "/Torre-A1/GPRS/TorreA1_15M_BD.dat", 
                 "/Torre-A2/WIFI/TorreA2_15M_BD.dat", "/Torre-A2/GPRS/TorreA2_15M_BD.dat",
                 "/Torre-B/WIFI/TorreB_15M_BD.dat", "/Torre-B/GPRS/TorreB_15M_BD.dat",
                 "/Torre-C/WIFI/TorreC_15M_BD.dat", "/Torre-C/GPRS/TorreC_15M_BD.dat",
                 "/Torre-D/WIFI/TorreD_15M_BD.dat", "/Torre-D/GPRS/TorreD_15M_BD.dat",
                 "/Torre-E1/WIFI/TorreE1_15M_BD.dat", "/Torre-E1/GPRS/TorreE1_15M_BD.dat",
                 "/Torre-E2/WIFI/TorreE2_15M_BD.dat", "/Torre-E2/GPRS/TorreE2_15M_BD.dat",
                 "/Torre-F/WIFI/TorreF_15M_BD.dat", "/Torre-F/GPRS/TorreF_15M_BD.dat",
                 "/Torre-G/WIFI/TorreG_15M_BD.dat", "/Torre-G/GPRS/TorreG_15M_BD.dat"
                 ]

list_file_name = []

is_updated = False

def update_files():
    if is_updated == True:
        if dtInicio.get()!="" and dtFinal.get()!="":
            a = datetime.strptime(dtInicio.get(), "%d/%m/%Y")
            b = datetime.strptime(dtFinal.get(), "%d/%m/%Y")
            delta = b - a
            if(delta.days < 1):
                #print("Data Inválida!")
                showerror("Erro", "As datas selecionadas são inválida!") 
            else:
                multiplicador = delta.days + 1
                nrDias.set(multiplicador)
                nr_registros = 96 * multiplicador;
                nrRegistros.set(nr_registros)
                #print("O total de dias é", nrDias.get()) 
                #print("O total de registros esperados é", nrRegistros.get()) 
                
                #print(origem.get()+list_db_files[0])
                #print(date_format(dtInicio.get(),True))
                #print(date_format(dtFinal.get(),False))
                root.config(cursor="exchange")
                root.update()
                if selTorreA1.get() == 1:
                    count = output_files(origem.get()+list_db_files[0], origem.get()+list_db_files[1], date_format(dtInicio.get(),True), date_format(dtFinal.get(), False), True)
                    totTorreA1.set(nrRegistros.get()+" / "+str(count))
                    resTorreA1.set(str(round(((int(count) / int(nrRegistros.get())) * 100),2))+" % ")
                if selTorreA2.get() == 1:
                    count = output_files(origem.get()+list_db_files[2], origem.get()+list_db_files[3], date_format(dtInicio.get(),True), date_format(dtFinal.get(), False), True)
                    totTorreA2.set(nrRegistros.get()+" / "+str(count))
                    resTorreA2.set(str(round(((int(count) / int(nrRegistros.get())) * 100),2))+" % ")
                if selTorreB.get() == 1:
                    count = output_files(origem.get()+list_db_files[4], origem.get()+list_db_files[5], date_format(dtInicio.get(),True), date_format(dtFinal.get(), False), False)
                    totTorreB.set(nrRegistros.get()+" / "+str(count))
                    resTorreB.set(str(round(((int(count) / int(nrRegistros.get())) * 100),2))+" % ")
                if selTorreC.get() == 1:
                    count = output_files(origem.get()+list_db_files[6], origem.get()+list_db_files[7], date_format(dtInicio.get(),True), date_format(dtFinal.get(), False), False)
                    totTorreC.set(nrRegistros.get()+" / "+str(count))
                    resTorreC.set(str(round(((int(count) / int(nrRegistros.get())) * 100),2))+" % ")
                if selTorreD.get() == 1:
                    count = output_files(origem.get()+list_db_files[8], origem.get()+list_db_files[9], date_format(dtInicio.get(),True), date_format(dtFinal.get(), False), False)
                    totTorreD.set(nrRegistros.get()+" / "+str(count))
                    resTorreD.set(str(round(((int(count) / int(nrRegistros.get())) * 100),2))+" % ")
                if selTorreE1.get() == 1:
                    count = output_files(origem.get()+list_db_files[10], origem.get()+list_db_files[11], date_format(dtInicio.get(),True), date_format(dtFinal.get(), False), False)
                    totTorreE1.set(nrRegistros.get()+" / "+str(count))
                    resTorreE1.set(str(round(((int(count) / int(nrRegistros.get())) * 100),2))+" % ")
                if selTorreE2.get() == 1:
                    count = output_files(origem.get()+list_db_files[12], origem.get()+list_db_files[13], date_format(dtInicio.get(),True), date_format(dtFinal.get(), False), False)
                    totTorreE2.set(nrRegistros.get()+" / "+str(count))
                    resTorreE2.set(str(round(((int(count) / int(nrRegistros.get())) * 100),2))+" % ")
                if selTorreF.get() == 1:
                    count = output_files(origem.get()+list_db_files[14], origem.get()+list_db_files[15], date_format(dtInicio.get(),True), date_format(dtFinal.get(), False), False)
                    totTorreF.set(nrRegistros.get()+" / "+str(count))
                    resTorreF.set(str(round(((int(count) / int(nrRegistros.get())) * 100),2))+" % ")
                if selTorreG.get() == 1:
                    count = output_files(origem.get()+list_db_files[16], origem.get()+list_db_files[17], date_format(dtInicio.get(),True), date_format(dtFinal.get(), False), False)
                    totTorreG.set(nrRegistros.get()+" / "+str(count))
                    resTorreG.set(str(round(((int(count) / int(nrRegistros.get())) * 100),2))+" % ")
                root.config(cursor="")
                #showinfo("Sucesso","As planilhas foram geradas com sucesso !")
                result = askyesno("Sucesso", "As planilhas foram geradas com sucesso !\n Deseja criar um arquivo compactado?")
                if result:
                    file_name = simpledialog.askstring("Arquivo", "Informe o nome do arquivo:", initialvalue="")
                    global list_file_name
                    with zipfile.ZipFile(file_name+'.zip', 'a') as myzip:
                        for file in list_file_name:
                            #print("Comprimindo o arquivo", file)
                            myzip.write(file, arcname=os.path.basename(file))
                            #myzip.write(file)
                        showinfo("Sucesso","As planilhas foram compactadas com sucesso !")
                        
                list_file_name = []
                
    else:
        showerror("Selecione", "Selecione a localização dos dados!")
     
def output_files(file_wifi, file_gprs, data_ini, data_fim, is_Torre_A = True):
    dados_torre_wifi = []
    dados_torre_gprs = []
    date_error = []
    date_error_index = []
    count_data = 0
    count_error = 0
    total_count = 0
    file_name_ext = os.path.basename(file_wifi)
    file_name = file_name_ext.split('.', 1)[0] + "_" + data_ini[0:10] + "-" + data_fim[0:10] + ".xlsx"
    data_final = increment_date(data_fim)
    #pb.configure(maximum=int(nrRegistros.get()))
    # Busca os dados do perido selecionado na tabela WIFI
    with open(file_wifi) as f:
        line = f.readline()
        data_atual = data_ini
        while line:
            #print(line[1:5],"=>",line[6:8],"=>",line[9:11])
            #print(is_date(line[1:5],line[6:8],line[9:11]))
            if datetime.strptime(data_final, "%Y-%m-%d %H:%M:%S") > datetime.strptime(data_atual, "%Y-%m-%d %H:%M:%S"):
                if is_date(line[1:5],line[6:8],line[9:11]):
                    #print(line[1:5], "=>", line[6:8], "=>", line[9:11], "=>", data_atual)
                    if data_atual == format_midnight(line[1:20]):
                        count_data += 1
                        dados_torre_wifi.append(line + ", "+ str(count_data))
                        total_count+=1
                        #print(line)
                        data_atual = increment_date(data_atual)
                    elif datetime.strptime(format_midnight(line[1:20]), "%Y-%m-%d %H:%M:%S") > datetime.strptime(data_atual, "%Y-%m-%d %H:%M:%S"):
                        count_data += 2
                        dados_torre_wifi.append(line+ ", "+ str(count_data))
                        total_count+=1
                        count_error += 1
                        date_error.append(data_atual)
                        date_error_index.append((count_data - 1))
                        count_data += 1
                        #print(line)
                        data_atual = increment_date(format_midnight(line[1:20]))
                    else:
                        line = f.readline()
                else:                  
                    line = f.readline()
            else:
                #print("Foram encontrados",count_data, "registros e tiveram", count_error, "dados faltantes!")
                break
            
            #var_bar.set(total_count)
            #root.update()
    # Busca os dados faltantes na tabela GPRS
    if len(date_error) > 0:
        with open(file_gprs) as f:
            line = f.readline()
            i = 0
            data_atual = date_error[i]
            while line:
                if datetime.strptime(data_final, "%Y-%m-%d %H:%M:%S") > datetime.strptime(data_atual, "%Y-%m-%d %H:%M:%S"):
                    if is_date(line[1:5],line[6:8],line[9:11]):
                        if data_atual == format_midnight(line[1:20]):
                            dados_torre_gprs.append(line+ ", "+ str(date_error_index[i]))
                            total_count+=1
                            count_data += 1
                            count_error -= 1
                            #print(line)
                            i += 1
                            
                            if i  < len(date_error):
                                data_atual = date_error[i]
                            else:
                                break
                        elif datetime.strptime(format_midnight(line[1:20]), "%Y-%m-%d %H:%M:%S") > datetime.strptime(data_atual, "%Y-%m-%d %H:%M:%S"):
                            i += 1
                            #print(i, "===>", len(date_error))
                            if i  < len(date_error):
                                data_atual = date_error[i]
                            else:
                                break
                        line = f.readline()
                    else:    
                        line = f.readline()
                else:
                    break
                
                #var_bar.set(total_count)
                #root.update()
                
        print("Foram encontrados",total_count, "registros e tiveram", count_error, "dados faltantes!")
        #print(date_error)
        #print("Nome_arquivo", file_name)
        create_spreadsheet(file_name, dados_torre_wifi, dados_torre_gprs, is_Torre_A)
        return total_count

def create_spreadsheet(name: str, lista1: list, lista2: list, is_tower_a: bool):
    wb = Workbook()
    ws = wb.active
    ws.title = 'dados'
    if is_tower_a == True:
        title = ["TIMESTAMP","RECORD","WS100","WD100","SIG100","WS60","WD60","SIG60","WS10","WD10","SIG10","TAr_100m_Avg","TAr_60m_Avg","TAr_10m_Avg","RH_Avg","SR_Avg","DT1",
                 "DT2","DT3","WG_Max","WG_2_Max","WG_3_Max","TAr_10m_Min","TAr_60m_Min","TAr_100m_Min","TAr_10m_Max","TAr_60m_Max","TAr_100m_Max","DT1_Min","DT2_Min","DT3_Min",
                 "DT1_Max","DT2_Max","DT3_Max","RH_Min","RH_Max","SR_Min","SR_Max","RNTOT15","PASQUILH","PASQUILV","Bateria_URA2","Bateria_SA3","Bateria_SA4","Recarga_SA3",
                 "Recarga_SA4","UR_Int","Porta_URA2","Porta_SA","Index"]
    else:
        title = ['TIMESTAMP','RECORD', 'WS' ,'WD' ,'SIG' ,'WG_Max' ,'PASQUILH' ,'Bat_UR' ,'Bat_SA1' ,'Bat_SA2' ,'Recarga_SA1' ,'Recarga_SA2' ,'UR_Int' ,'Porta_UR','Porta_SA', 'Index']
    
    ws.append(title)
    
    for line1 in lista1:
        list_field1 =  line1.split(',')
        date_format = list_field1[0] 
        list_field1[0] = date_format[1:20]
        ws.append(list_field1)
    
    for line2 in lista2:
        list_field2 = line2.split(',')
        date_format = list_field2[0] 
        list_field2[0] = date_format[1:20]
        ws.append(list_field2)
    
    name = os.path.dirname(os.path.abspath(__file__)) + "\\" + name
    list_file_name.append(name)
         
    wb.save(name)
    df = pd.read_excel(name)
    df = df.sort_values('Index')
    df.drop(['Index','RECORD'], axis=1, inplace=True)
    df.to_excel(name, index=False)
    #df.save()

def date_format(date_unformated, is_initial = True):
    resultado = date_unformated[6:11]+"-"+ date_unformated[3:5]+"-"+ date_unformated[0:2]
    if is_initial:
        return resultado + " 00:00:00"
    else:
        return resultado + " 23:45:00"

def increment_date(date):
    if date[11:16]  == "23:45":
        return add_day(date[0:10])  + " 00:00:00"
    else:
        if date[14:16] == "00":
            return date[0:14]+"15:00"
        elif date[14:16] == "15":
            return date[0:14]+"30:00"
        elif date[14:16] == "30":
            return date[0:14]+"45:00"
        elif  date[14:16] == "45":
            add_hour = int(date[11:13]) + 1
            if add_hour < 10:
                return date[0:11]+"0"+str(add_hour) +":00:00"
            else:
                return date[0:11]+str(add_hour) +":00:00"

def add_day(date):
    actual_day = datetime.strptime(date, "%Y-%m-%d") 
    return_day = str(actual_day + timedelta(days=1))
    return return_day[0:10]

def format_midnight(date):
    if date[11:13] == "24":
        return add_day(date[0:10])+" 00:00:00"
    else:
        return date
        

def is_date(year, month, day):
    if year.isnumeric() and month.isnumeric() and day.isnumeric():
        correctDate = None
        try:
            newDate = datetime(int(year), int(month), int(day))
            correctDate = True
        except ValueError:
            correctDate = False
        return correctDate
    else:
        return False
    
def open_select_origem():
    retorno = filedialog.askdirectory()
    global is_updated
    for x in list_db_files:
        path_file = Path(retorno + x)
        if path_file.is_file() == False:
            #print(path_file)
            showerror("Erro", "O arquivo " + str(path_file) + " não foi localizado!")
            is_updated = False
            break
    origem.set(retorno)
    is_updated = True
            
def get_timestamp_file(path):
    ti_m = os.path.getmtime(path) 
    return ti_m

if __name__ == "__main__":
    root = tk.Tk()
    root.title('Backup Meteorologia')
    root.resizable(0, 0)
    
    try:
        # windows only (remove the minimize/maximize button)
        root.attributes('-toolwindow', True)
    except TclError:
        print('Not supported on your platform')

    # layout on the root window
    root.columnconfigure(0, weight=1)
    #root.columnconfigure(1, weight=1)
    origem = tk.StringVar()
    dtInicio = tk.StringVar()
    dtFinal = tk.StringVar()
    nrDias = tk.StringVar()
    nrRegistros = tk.StringVar()
    totTorreA1 = tk.StringVar()
    resTorreA1 = tk.StringVar()
    selTorreA1 = tk.IntVar()
    selTorreA1.set(1)
    totTorreA2 = tk.StringVar()
    resTorreA2 = tk.StringVar()
    selTorreA2 = tk.IntVar()
    selTorreA2.set(1)
    totTorreB = tk.StringVar()
    resTorreB = tk.StringVar()
    selTorreB = tk.IntVar()
    selTorreB.set(1)
    totTorreC = tk.StringVar()
    resTorreC = tk.StringVar()
    selTorreC = tk.IntVar()
    selTorreC.set(1)
    totTorreD = tk.StringVar()
    resTorreD = tk.StringVar()
    selTorreD = tk.IntVar()
    selTorreD.set(1)
    totTorreE1 = tk.StringVar()
    resTorreE1 = tk.StringVar()
    selTorreE1 = tk.IntVar()
    selTorreE1.set(1)
    totTorreE2 = tk.StringVar()
    resTorreE2 = tk.StringVar()
    selTorreE2 = tk.IntVar()
    selTorreE2.set(1)
    totTorreF = tk.StringVar()
    resTorreF = tk.StringVar()
    selTorreF = tk.IntVar()
    selTorreF.set(1)
    totTorreG = tk.StringVar()
    resTorreG = tk.StringVar()
    selTorreG = tk.IntVar()
    selTorreG.set(1)
    
    root.columnconfigure(0, weight=1)
    ttk.Label(root, text='Localização dos dados:').grid(column=0, row=0, sticky=tk.W)
    txt_origem = ttk.Entry(root, width=52, textvariable=origem)
    txt_origem.grid(column=1, row=0, sticky=tk.W, columnspan=3)

    root.columnconfigure(0, weight=1)
    ttk.Button(root, text='Localizar', command=open_select_origem).grid(column=4, row=0, sticky=tk.W)
    
    ttk.Label(root, text='Período de aquisição:').grid(column=0, row=1, sticky=tk.W)
    #dt_inicio = ttk.Entry(root, width=18, textvariable=dtInicio).grid(column=1, row=1, sticky=tk.W)
    dt_inicio = DateEntry(root, width=18, date_pattern="dd/mm/yyyy", textvariable=dtInicio).grid(column=1, row=1, sticky=tk.W)
    
    ttk.Label(root, text='à', width=2, anchor="nw").grid(column=2, row=1)
    #dt_final = ttk.Entry(root, width=18, textvariable=dtFinal)
    #dt_final.grid(column=3, row=1, sticky=tk.W)
    dt_final = DateEntry(root, width=18, date_pattern="dd/mm/yyyy", textvariable=dtFinal).grid(column=3, row=1, sticky=tk.W)
        
    root.columnconfigure(0, weight=1)
    btn_update = ttk.Button(root, text='Atualizar', command=update_files).grid(column=4, row=1, sticky=tk.E)
    
    # Título
    ttk.Label(root, text='Torre', borderwidth=2, relief="ridge", width=22, anchor="center").grid(column=0, row=5, sticky=tk.W)
    ttk.Label(root, text='Esperados / Encontrados', borderwidth=2, relief="ridge", width=30, anchor="center").grid(column=1, row=5, sticky=tk.W, columnspan=2)
    ttk.Label(root, text='Resultado', borderwidth=2, relief="ridge", width=15, anchor="center").grid(column=3, row=5, sticky=tk.W)
    # Dados Torre A-1
    lbl_torre_a1 = ttk.Label(root, text='Torre A1', borderwidth=2, relief="ridge", width=22, anchor="center").grid(column=0, row=6, sticky=tk.W)
    tot_torre_a1 = ttk.Entry(root, width=30, textvariable=totTorreA1).grid(column=1, row=6, sticky=tk.W, columnspan=2)
    res_torre_a1 = ttk.Entry(root, width=15, textvariable=resTorreA1).grid(column=3, row=6, sticky=tk.W)
    tk.Checkbutton(root, variable=selTorreA1).grid(column=4, row=6, sticky=tk.W)
    # Dados Torre A-2
    ttk.Label(root, text='Torre A2', borderwidth=2, relief="ridge", width=22, anchor="center").grid(column=0, row=7, sticky=tk.W)
    tot_torre_a2 = ttk.Entry(root, width=30, textvariable=totTorreA2).grid(column=1, row=7, sticky=tk.W, columnspan=2)
    res_torre_a2 = ttk.Entry(root, width=15, textvariable=resTorreA2).grid(column=3, row=7, sticky=tk.W)
    tk.Checkbutton(root, variable=selTorreA2).grid(column=4, row=7, sticky=tk.W)
    # Dados Torre B
    ttk.Label(root, text='Torre B', borderwidth=2, relief="ridge", width=22, anchor="center").grid(column=0, row=8, sticky=tk.W)
    tot_torre_b = ttk.Entry(root, width=30, textvariable=totTorreB).grid(column=1, row=8, sticky=tk.W, columnspan=2)
    res_torre_b = ttk.Entry(root, width=15, textvariable=resTorreB).grid(column=3, row=8, sticky=tk.W)
    tk.Checkbutton(root, variable=selTorreB).grid(column=4, row=8, sticky=tk.W)
    # Dados Torre C
    ttk.Label(root, text='Torre C', borderwidth=2, relief="ridge", width=22, anchor="center").grid(column=0, row=9, sticky=tk.W)
    tot_torre_c = ttk.Entry(root, width=30, textvariable=totTorreC).grid(column=1, row=9, sticky=tk.W, columnspan=2)
    res_torre_c = ttk.Entry(root, width=15, textvariable=resTorreC).grid(column=3, row=9, sticky=tk.W)
    tk.Checkbutton(root, variable=selTorreC).grid(column=4, row=9, sticky=tk.W)
    # Dados Torre D
    ttk.Label(root, text='Torre D', borderwidth=2, relief="ridge", width=22, anchor="center").grid(column=0, row=10, sticky=tk.W)
    tot_torre_d = ttk.Entry(root, width=30, textvariable=totTorreD).grid(column=1, row=10, sticky=tk.W, columnspan=2)
    res_torre_d = ttk.Entry(root, width=15, textvariable=resTorreD).grid(column=3, row=10, sticky=tk.W)
    tk.Checkbutton(root, variable=selTorreD).grid(column=4, row=10, sticky=tk.W)
    # Dados Torre E1
    ttk.Label(root, text='Torre E1', borderwidth=2, relief="ridge", width=22, anchor="center").grid(column=0, row=11, sticky=tk.W)
    tot_torre_e1 = ttk.Entry(root, width=30, textvariable=totTorreE1).grid(column=1, row=11, sticky=tk.W, columnspan=2)
    res_torre_e1 = ttk.Entry(root, width=15, textvariable=resTorreE1).grid(column=3, row=11, sticky=tk.W)
    tk.Checkbutton(root, variable=selTorreE1).grid(column=4, row=11, sticky=tk.W)
    # Dados Torre E2
    ttk.Label(root, text='Torre E2', borderwidth=2, relief="ridge", width=22, anchor="center").grid(column=0, row=12, sticky=tk.W)
    tot_torre_e2 = ttk.Entry(root, width=30, textvariable=totTorreE2).grid(column=1, row=12, sticky=tk.W, columnspan=2)
    res_torre_e2 = ttk.Entry(root, width=15, textvariable=resTorreE2).grid(column=3, row=12, sticky=tk.W)
    tk.Checkbutton(root, variable=selTorreE2).grid(column=4, row=12, sticky=tk.W)
    # Dados Torre F
    ttk.Label(root, text='Torre F', borderwidth=2, relief="ridge", width=22, anchor="center").grid(column=0, row=13, sticky=tk.W)
    tot_torre_f = ttk.Entry(root, width=30, textvariable=totTorreF).grid(column=1, row=13, sticky=tk.W, columnspan=2)
    res_torre_f = ttk.Entry(root, width=15, textvariable=resTorreF).grid(column=3, row=13, sticky=tk.W)
    tk.Checkbutton(root, variable=selTorreF).grid(column=4, row=13, sticky=tk.W)
    # Dados Torre G
    ttk.Label(root, text='Torre G', borderwidth=2, relief="ridge", width=22, anchor="center").grid(column=0, row=14, sticky=tk.W)
    tot_torre_g = ttk.Entry(root, width=30, textvariable=totTorreG).grid(column=1, row=14, sticky=tk.W, columnspan=2)
    res_torre_g = ttk.Entry(root, width=15, textvariable=resTorreG).grid(column=3, row=14, sticky=tk.W)
    tk.Checkbutton(root, variable=selTorreG).grid(column=4, row=14, sticky=tk.W)
    # label
    value_label = ttk.Label(root, text="")
    value_label.grid(column=0, row=18, columnspan=2)
      
    root.mainloop()