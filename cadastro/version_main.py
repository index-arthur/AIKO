import time
import random
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

grupos = ["reserva", "teste", "inativos"]
model = ["feller", "basculante"]

def pausa(min_s=0.5, max_s=3.5):
    time.sleep(random.uniform(min_s, max_s))

def fechar_emergencias(driver):
    while True:
        botoes = driver.find_elements(
            By.XPATH, '//*[@id="legacy-global-modals"]//button'
        )
        if not botoes:
            break
        for b in botoes:
            try:
                b.click()
                time.sleep(0.5)
            except:
                pass

def clicar_salvar_ultimo_visivel(driver, timeout=8):
    wait_local = WebDriverWait(driver, timeout)
    wait_local.until(EC.visibility_of_element_located((By.ID, "name")))

    spans = driver.find_elements(By.XPATH, "//*[@id='app']//span[normalize-space()='Salvar']")
    spans = [s for s in spans if s.is_displayed()]

    span_salvar = spans[-1]
    clicavel = span_salvar.find_element(By.XPATH, "./ancestor::button[1] | ./ancestor::a[1]")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", clicavel)
    driver.execute_script("arguments[0].click();", clicavel)

while True:
    usuario = input('Digite seu usuario: ') + "@aiko.digital"
    senha = input('Digite sua senha: ')
    empresa = input('Digite qual empresa: ').upper()
    equipamento = input('Digite qual o modelo de equipamento: ').upper()
    ticket = int(input("Digite o ticket: "))
    zendesk = input("Ticket Zendesk (N se não tiver): ").strip()
    parou = int(input("Se parou em algum bordo, digite o número (0 se não parou): "))
    limite = int(input("Digite quantos bordos: "))

    print("\n--- CONFIRA OS DADOS ---")
    print(f"Usuário: {usuario}")
    print(f"Senha: {senha}")
    print(f"Empresa: {empresa}")
    print(f"Equipamento: {equipamento}")
    print(f"Ticket: {ticket}")
    print(f"T.Zendesk: {zendesk}")
    print(f"Parou no bordo: {parou}")
    print(f"Q.Bordos: {limite}")
    print("------------------------")

    confirma = input("Os dados estão corretos? (S/N): ").strip().upper()

    if confirma == "S":
        break

driver = webdriver.Edge()
driver.maximize_window()
driver.get(f"https://{empresa}.br.trackit.host/")

# LOGIN
driver.find_element("id", "UserName").send_keys(usuario)
driver.find_element("id", "Password").send_keys(senha, Keys.ENTER)

wait = WebDriverWait(driver, 15)

fechar_emergencias(driver)

# ENTRANDO NO CADASTRO DE BORDO
wait.until(EC.presence_of_element_located((By.ID, "nav-menu")))
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="nav-menu"]/div[1]/div[2]/a'))).click()
wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="nav-menu"]/div[1]/div[2]/div/div/section[1]/article/ul/li[4]/a'))).click()

#LOOP PRINCIPAL
user_zendesk = zendesk.lower() != 'n'
zendesk_txt = f" | #{zendesk}" if user_zendesk else ""

inicio = parou + 1
fim = limite + 1

for k in range(inicio, fim):
    numero = f"{k:02d}"
    bordo = f"{empresa} | {equipamento} | HWS-{ticket}{zendesk_txt} | {numero}"

    #Click no botao azul
    wait.until(EC.presence_of_element_located((By.ID, "app")))
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="app"]/div[1]/div[5]/div/section/div/div[2]/a/span'))).click()
    pausa(1, 1.3)

    #Escreve o nome do bordo
    wait.until(EC.element_to_be_clickable((By.ID, "name"))).send_keys(bordo)
    pausa(1, 1.3)

    #Seleciona o Modelo
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="equipmentModel"]/div[1]')
    )).click()
    wait.until(EC.visibility_of_element_located(
        (By.XPATH, '//*[@id="equipmentModel"]/div[3]')
    ))
    opcoes1 = wait.until(EC.presence_of_all_elements_located(
        (By.XPATH, '//*[@id="equipmentModel"]/div[3]//li')
    ))
    for opcao1 in opcoes1:
        texto1 = opcao1.text.strip().lower()
        if any(p in texto1 for p in model):
            opcao1.click()
            break
    pausa(0.5, 0.8)

    #Seleciona o Perfil
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="networkProfiles"]/div[1]')
    )).click()
    wait.until(EC.visibility_of_element_located(
        (By.XPATH, '//*[@id="networkProfiles"]/div[3]')
    ))
    opcoes2 = wait.until(EC.presence_of_all_elements_located(
        (By.XPATH, '//*[@id="networkProfiles"]/div[3]//li')
    ))
    for opcao2 in opcoes2:
        texto2 = opcao2.text.strip().lower()
        if "equipamento" in texto2:
            opcao2.click()
            break
    pausa(0.5, 0.8)

    #Seleciona Grupo
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="equipmentGroup"]/div[1]')
    )).click()
    wait.until(EC.visibility_of_element_located(
        (By.XPATH, '//*[@id="equipmentGroup"]/div[3]')
    ))
    opcoes3 = wait.until(EC.presence_of_all_elements_located(
        (By.XPATH, '//*[@id="equipmentGroup"]/div[3]//li')
    ))
    for opcao3 in opcoes3:
        texto3 = opcao3.text.strip().lower()
        if any(p in texto3 for p in grupos):
            opcao3.click()
            break
    pausa(0.5, 0.8)

    #Salva
    clicar_salvar_ultimo_visivel(driver, timeout=8)
    wait.until(EC.invisibility_of_element_located((By.ID, "name")))
    pausa(0.5, 1)
