import random
import time
import pandas as pd
import urllib
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By

# Saída de emergência, caso de erro
# Encostar o mouse na topo superior esquerdo
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3


# Arquivo da planilha de agendamento, SHEET_NAME = nome da aba que será lida da planilha
p = pd.read_excel("JUNHO.xlsx", sheet_name="dados", skiprows=2)

# caminho do navegador do Selenium
navegador = webdriver.Chrome(executable_path="chromedriver.exe")

# Loop que itera nas linhas da planilha de agendamento
for indice, linha in p.iterrows():
  # Variáveis de cada linha da planilha
  especialidade = linha['ESPECIALIDADE']
  data_da_consulta = linha['DATA'].strftime("%d/%m/%Y")
  horario_da_consulta = linha['HORÁRIO'].strftime("%H:%M")
  nome_do_usuario = linha['NOME']
  celular_do_usuario = int(linha['TELEFONE 1'])

  # Mensagem que é enviada
  mensagem = urllib.parse.quote (f"""

Sr(a) *{nome_do_usuario}* SUA CONSULTA *{especialidade}* no *AMA ESPECIALIDADES CAPÃO REDONDO* está agendado para o dia *{data_da_consulta}* às *{horario_da_consulta}h* chegar com (30 minutos) de antecedência.

*AVENIDA COMENDADOR SANT'ANNA, 542 - VILA FAZZEONI.* 
Ponto de Referência, a 4 pontos de ônibus saindo do metrô Capão Redondo

Para *CONFIRMAR* responda *1*
Para *CANCELAR* responda *2*

*Por favor, confirme esta mensagem em até 24 horas após o recebimento. OU SERÁ CANCELADO AUTOMATICAMENTE*

Caso apresente sintomas gripais (tosse, dor de garganta, falta de ar ou febre) que indiquem suspeita de COVID-19, ou tenha tido contato recentemente com pessoas que testaram positivo para COVID-19, informe para que seja realizado o cancelamento e posteriormente a remarcação. Obrigatório o uso de máscara.

*POR FAVOR NÃO DENUNCIE ESTE MEIO DE COMINUCAÇÃO AO WHATSAPP, POIS OUTROS PACIENTES FICARÃO IMPOSSIBILITADOS DE RECEBER NOTÍCIAS DE SUAS CONSULTAS. CASO NÃO CONHEÇA O PACIENTE POR FAVOR NOS INFORME PARA QUE POSSAMOS EXCLUIR O CONTATO DO CADASTRO.*

*Qualquer dúvida, por favor entre em contato pelo telefone:* 
*(11) 5874-9200*

*ESTE CANAL É EXCLUSIVO PARA CONFIRMAÇÃO E CANCELAMENTO DE CONSULTA. PARA OUTRAS INFORMAÇÕES, LIGUE NO NÚMERO ACIMA OU DIRIJA-SE A UNIDADE.*""")

  # Link que é aberto no navegador
  link_whatsapp = f"https://api.whatsapp.com/send?phone=55{celular_do_usuario}&text={mensagem}"

  # navegador abre o link
  navegador.get(link_whatsapp)
  time.sleep(3)

  # Localiza o botão de iniciar a conversa, que é o que abre o WhatsAPP WEB
  iniciar_conversa = navegador.find_element(By.ID, "action-button")
  time.sleep(2)
  iniciar_conversa.click()

  # Localizar o botão de enviar a mensagem e envia a mensagem
  enviar = pyautogui.locateCenterOnScreen("send.png", confidence=0.9)
  time.sleep(2)
  pyautogui.click(enviar)
  
  # Tempo que ele aguarda entra uma iteração e outra.
  tempo = random.randint(5, 10)
  print(tempo)
  time.sleep(tempo)

pyautogui.alert("Fim da Execução")
