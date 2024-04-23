from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import openpyxl
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


driver = webdriver.Chrome()
# 1 - Navegar até o site https://contabilidade-devaprender.netlify.app/
driver.get('https://contabilidade-devaprender.netlify.app/')
sleep(5)

# 2 - Digitar e-mail
email = driver.find_element(By.XPATH, "//input[@id='email']")
sleep(2)
email.send_keys('teste@gmail.com')
# 3 - Digitar senha
senha = driver.find_element(By.XPATH, "//input[@id='senha']")
sleep(2)
senha.send_keys('123456jm')
# 4 - cliclar em entrar
botao = driver.find_element(By.XPATH, "//button[@id='Entrar']")
sleep(2)
botao.click()
# 5 - extrair informaçoes da planilha
empresas = openpyxl.load_workbook('./empresas.xlsx')
pagina_empresa = empresas['dados empresas']

for linha in pagina_empresa.iter_rows(min_row=2,values_only=True):
    nome_da_empresa, email, telefone, endereco, cnpj, area_atuacao, quantidade_de_funcionarios, data_fundacao = linha

    # Esperar até que o elemento de nome da empresa esteja visível
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, 'nomeEmpresa'))).send_keys(nome_da_empresa)
    driver.find_element(By.ID, 'emailEmpresa').send_keys(email)
    sleep(1)
    driver.find_element(By.ID, 'telefoneEmpresa').send_keys(telefone)
    sleep(1)
    driver.find_element(By.ID, 'enderecoEmpresa').send_keys(endereco)
    sleep(1)
    driver.find_element(By.ID, 'cnpj').send_keys(cnpj)
    sleep(1)
    driver.find_element(By.ID, 'areaAtuacao').send_keys(area_atuacao)
    sleep(1)
    driver.find_element(By.ID, 'numeroFuncionarios').send_keys(quantidade_de_funcionarios)
    sleep(1)
    driver.find_element(By.ID, 'dataFundacao').send_keys(data_fundacao)
    sleep(1)

    driver.find_element(By.ID, 'Cadastrar').click()
    sleep(3)
# 6 - cliclar em cada campo e preencher com informação 
# 7 - cliclar em cadastrar
# 8 - Repitir os passos 5 e 6