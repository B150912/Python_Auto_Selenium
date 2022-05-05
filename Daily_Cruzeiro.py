# -*- coding: utf-8 -*-
"""
Created on Wed Mar  2 16:05:17 2022

@author: bruno.andriolli
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager #Para automação no processo do formulario
from time import sleep
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pyodbc
import pandas as pd
#import numpy as np
from s_email import s_email

#DECLARAÇÃO DAS VARIÁVEIS


dt_atu = datetime.now().date()
#dt_atu = datetime(2022,2,25).date()
dt_fim = datetime(2022,4,1).date()
dt_reg = datetime(2022,3,1).date()
#
#if dt_atu.weekday() == 0:
#    dt_atu = dt_atu - relativedelta(days=2)
#else:
#dt_reg = dt_atu - relativedelta(days=1)
    
dt = dt_atu.strftime("%d/%m/%Y")

#navegador = webdriver.Chrome(ChromeDriverManager().install())


while dt_reg < dt_fim :
    
    server = '192.168.200.182\db2014'
    database = 'DW_JAREZENDE'
    username = 'MIS'
    password = 'm1s'
    cnn = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
    cursor = cnn.cursor()
     
    sql ="SELECT * FROM DW_JAREZENDE.DBO.DAILY_CRUZEIRO WHERE DATA = '{}'".format(dt_reg)
    
    #EXECUÇÃO DO SQL
    
    
    df = pd.read_sql(sql, cnn)
    
    if len(df)>0:
        
        navegador = webdriver.Chrome(ChromeDriverManager().install())
        try:
            #TRATATIVA DO RESULTADO E PREENCHIMENTO DO FORMULÁRIO
            for i, resp in enumerate(df["RESPONSAVEL"]):
                dta = df.loc[i, "DATA"].strftime("%d/%m/%Y")
                capacity = df.loc[i, "CAPACITY"]
                estoque = df.loc[i, "ESTOQUE"]
                mailing = df.loc[i, "MAILING"]
                discados = df.loc[i, "DISCADOS"]
                discagens = df.loc[i, "DISCAGENS"]
                spin = df.loc[i, "SPIN"]
                alo = df.loc[i, "ALO"]
                cpc = df.loc[i, "CPC"]
                alega_pgto = df.loc[i, "ALEGA"]
                promessa_pgto = df.loc[i, "PROMESSA"]
                vl_promessa_pgto = df.loc[i, "VL_PROMESSA"]
                volume_pgto = df.loc[i, "QTD_PGTO"]
                vl_pgto = df.loc[i, "VL_PAGTO"]
                dtt = df.loc[i, "DT_ATUALIZACAO"].strftime("%d/%m/%Y")
                
                lista_emails = 'planejamento-mis@jarezende.com.br; renan.betioli@jarezende.com.br; camila.boaro@jarezende.com.br'
                assunto = 'Formulario Daily Cruzeiro'
                msg = 'Bom dia! \n \nFormulário preenchido e disparado com sucesso. \n \nAtenciosamente \nMIS JAREZENDE'
                  
            
                navegador.get('https://forms.office.com/Pages/ResponsePage.aspx?id=F2EAcdgOgECWq-H3QZe37jMexdXEtFVMmRF_MIPVSfhUQ1NYRkpHQTlaS0I4UjU2UjBQNzc3QU43Vy4u&wdLOR=c69C56923-F188-4325-A26E-989838D10CC3')
                
                sleep(15)
                
            #SELECIONAR ESCRITÓRIO
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[1]/div/div[2]/div/div/div/label/input').click()
            
            #RESPONSÁVEL PELO PREENCHIMENTO
            
                navegador.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[2]/div/div[2]/div/div/input').send_keys(resp)
            
            #PERÍODO INFORME DATA
            
                navegador.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[3]/div/div[2]/div/div/input[1]').send_keys(dta)
            
            #CAPACITY
            
                navegador.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[4]/div/div[2]/div/div/input').send_keys(capacity)
            
            #ESTOQUE
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[5]/div/div[2]/div/div/input').send_keys(estoque)
            
            #MAILING
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[6]/div/div[2]/div/div/input').send_keys(mailing)
            
            #DISCADOS
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[7]/div/div[2]/div/div/input').send_keys(discados)
            
            #DISCAGENS
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[8]/div/div[2]/div/div/input').send_keys(discagens)
            
            #SPIN
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[9]/div/div[2]/div/div/input').send_keys(spin)
            
            #ALO
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[10]/div/div[2]/div/div/input').send_keys(alo)
            
            #CPC
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[11]/div/div[2]/div/div/input').send_keys(cpc)
            
            #ALEGA PAGAMENTO
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[12]/div/div[2]/div/div/input').send_keys(alega_pgto)
            
            #PROMESSA DE PAGAMENTO QUANTIDADE
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[13]/div/div[2]/div/div/input').send_keys(promessa_pgto)
            
            #PROMESSA DE PAGAMENTO VALOR
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[14]/div/div[2]/div/div/input').send_keys(vl_promessa_pgto)
            
            #VOLUME TOTAL DE PAGAMENTOS
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[15]/div/div[2]/div/div/input').send_keys(volume_pgto)
            
            #VALOR PAGO
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[2]/div[16]/div/div[2]/div/div/input').send_keys(str(vl_pgto))
            
            #ENVIAR
            
                navegador.find_element(By.XPATH, '//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[2]/div[3]/div[1]/button/div').click()
                
                sleep(5)  
                   
            #FINALIZAR
                navegador.close()
            
                cnn.commit()
                cursor.close()
                
               
        
        except Exception as e:
            print(str(e))
            s_email(lista_emails,'Erro - Formulário Daily Crizeiro','Bom dia! \n \nHouve um problema bi envio do Daily! \n \nErro:{}'.format(e))
            
    dt_reg = dt_reg + relativedelta(days=1)
else:
    raise SystemExit()