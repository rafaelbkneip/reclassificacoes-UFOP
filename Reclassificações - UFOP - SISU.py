import requests
import xlsxwriter  
from selenium import webdriver
import selenium.webdriver.support.expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from time import sleep
from datetime import date

options = Options()
options.add_experimental_option("detach", True)

navegador = webdriver.Chrome(ChromeDriverManager().install(), options=options)

navegador.get("https://vestibular.ufop.br/resultvest/2023_1/Chamada02Matricula/2023-1Chamada02.html")

#Garantir que a página carregou por completo
sleep(3)

#Variavéis auxiliares para o loop dos alunos por curso
cursos = []
cont = 2

while(True):
    #O loop se encerrará quando não existir mais nenhum link de nenhum curso no HTML da página. Todos os existentes são salvos na lista 'cursos'
    try:
        cursos.append(navegador.find_element(By.XPATH, '/html/body/table[1]/tbody/tr['+ str(cont)+ ']/td/font/font/a'))
        cont = cont + 1
    except:
        break

#Lista com os cursos
print(cursos)

#Variavéis auxiliares para o loop dos alunos por curso
aux = []
lista = []
cont_alunos = 2

#Acessa a página para cada um dos cursos
for i in range(len(cursos)):
    cursos[i].click()
    sleep(5)
    
    while(True):
        try:
            #O loop se encerrará quando não existir mais nenhum aluno para o curso acessado. Todos as informações são salvos na lsita 'cursos'
            navegador.find_element(By.XPATH, '/html/body/center/center/table[1]/tbody/tr[' +str (cont_alunos)+ ']/td[1]')
            #Quando não gouver uam respsota positiva, significa que o programa chegou aoo último aluno apra aquele curso e com isso o loop pode ser finalziado
        except:
            break;

        for j in range (1,15):
            aux.append(navegador.find_element(By.XPATH, '/html/body/center/center/table[1]/tbody/tr[' +str (cont_alunos)+ ']/td[' + str(j) + ']').text)
            print(aux)
            
        lista.append(aux)
        aux = []

        cont_alunos = cont_alunos + 1

    #Volta à página com os links dos cursos
    navegador.back()
    cont_alunos = 2

print(lista)

#Definir o caminho para salvar o arquivo .xlsx      
book = xlsxwriter.Workbook('')     
sheet = book.add_worksheet()  

#Cabeçalho do arquivo
sheet.write(0, 0, 'Nome')
sheet.write(0, 1, 'Inscrito Cota Escola Pública')
sheet.write(0, 2, 'Usou Cota')
sheet.write(0, 3, 'Classificação Geral')
sheet.write(0, 4, 'AC')
sheet.write(0, 5, 'L1')
sheet.write(0, 6, 'L2')
sheet.write(0, 7, 'L5')
sheet.write(0, 8, 'L6')
sheet.write(0, 9, 'L9')
sheet.write(0, 10, 'L10')
sheet.write(0, 11, 'L13')
sheet.write(0, 12, 'L14')
sheet.write(0, 13, 'Status')

#Todas as listas possuem o mesmo número de elementos
for i in range(len(lista)):
    sheet.write(i+1, 0, lista[i][0])
    sheet.write(i+1, 1, lista[i][1])
    sheet.write(i+1, 2, lista[i][2])
    sheet.write(i+1, 3, lista[i][3])
    sheet.write(i+1, 4, lista[i][4])
    sheet.write(i+1, 5, lista[i][5])
    sheet.write(i+1, 6, lista[i][6])
    sheet.write(i+1, 7, lista[i][7])
    sheet.write(i+1, 8, lista[i][8])
    sheet.write(i+1, 9, lista[i][9])
    sheet.write(i+1, 10, lista[i][10]) 
    sheet.write(i+1, 11, lista[i][11])
    sheet.write(i+1, 12, lista[i][12])
    sheet.write(i+1, 13, lista[i][13])

book.close()

