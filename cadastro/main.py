import time
import random
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
import tkinter as tk
from tkinter import messagebox

TUTORIAL = """
╔══════════════════════════════════════════════════════════════╗
║                     📋 TUTORIAL DE USO                       ║
╠══════════════════════════════════════════════════════════════╣
║                                                              ║
║  INFORMAÇÕES NECESSÁRIAS:                                    ║
║                                                              ║
║  [1] Usuário    → Seu usuário sem o @aiko.digital            ║
║                   Ex: acosta                                 ║
║                                                              ║
║  [2] Senha      → Sua senha de acesso                        ║
║                                                              ║
║  [3] Empresa    → Sigla da empresa no sistema                ║
║                   Ex: BRC, RAI, QA.BRC                       ║
║                                                              ║
║  [4] Equipamento→ Tipo do equipamento                        ║
║                   Ex: COMODATO, SERVICO DE CAMPO             ║
║                                                              ║
║  [5] Ticket     → Número do ticket (só os números)           ║
║                   Ex: 123456                                 ║
║                                                              ║
║  [6] T.Zendesk  → Número do ticket Zendesk                   ║
║                   Se não tiver, digite N                     ║
║                                                              ║
║  [7] Parou no bordo → Se a automação parou em algum bordo,   ║
║                      digite o número. Se não parou, digite 0 ║
║                                                              ║
║  [8] Q.Bordos   → Total de bordos a cadastrar                ║
║                   Ex: 10                                     ║
║                                                              ║
║  [9] Grupos     → Grupos do site. Padrão já vem preenchido.  ║
║                   Se o site tiver grupos diferentes, digite N║
║                   e informe                                  ║
║                   Ex: colheita sp                            ║
║                                                              ║
║  [10] Modelos   → Modelos de equipamento. Mesmo esquema      ║
║                   dos grupos acima                           ║
║                                                              ║
║  [11] Perfil    → Perfil de rede. Padrão: equipamentos       ║
║                   Altere só se o site usar outro perfil      ║
║                                                              ║
╚══════════════════════════════════════════════════════════════╝
"""

while True:
    print("\n╔══════════════════════════════╗")
    print("║    AUTOMAÇÃO CRIAÇÃO BORDO   ║")
    print("╠══════════════════════════════╣")
    print("║  [1] Tutorial                ║")
    print("║  [2] Iniciar automação       ║")
    print("║  [0] Sair                    ║")
    print("╚══════════════════════════════╝")

    opcao = input("\nEscolha uma opção: ").strip()

    if opcao == "1":
        print(TUTORIAL)
        input("Pressione ENTER para voltar ao menu...")
    elif opcao == "2":
        break
    elif opcao == "0":
        print("Saindo...")
        exit()
    else:
        print("Opção inválida, tente novamente.")

grupos_default = ["BAR-VN-000", "reserva", "teste", "inativos"]
model_default = ["AXOR 3344", "feller", "basculante", "forwarder"]
perfil_default = ["equipamentos"]

def clicar_dropdown(driver, xpaths_botao, xpaths_opcoes, condicao):
    wait_curto = WebDriverWait(driver, 3)
    wait_normal = WebDriverWait(driver, 15)

    # Tenta abrir o dropdown
    for xpath in xpaths_botao:
        try:
            wait_curto.until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
            break
        except TimeoutException:
            continue

    # Espera as opções aparecerem e pega elas
    opcoes = []
    for xpath in xpaths_opcoes:
        try:
            wait_normal.until(EC.visibility_of_element_located((By.XPATH, xpath)))
            opcoes = driver.find_elements(By.XPATH, xpath + "//li")
            if opcoes:
                break
        except TimeoutException:
            continue

    # Clica na opção certa
    for opcao in opcoes:
        texto = opcao.text.strip().lower()
        if condicao(texto):
            opcao.click()
            return True
    return False

def coletar_lista(nome, default):
    print(f"\n{nome} padrão: {default}")
    escolha = input(f"Usar padrão? (S = sim / N = digitar os seus): ").strip().upper()
    if escolha == "S":
        return default
    print(f"Digite os termos separados por vírgula:")
    entrada = input(">>> ").strip()
    return [x.strip() for x in entrada.split(",") if x.strip()]

def aviso_ok(titulo: str, mensagem: str):
    root = tk.Tk()
    root.withdraw()              
    root.attributes("-topmost", True) 
    messagebox.showinfo(titulo, mensagem) 
    root.destroy()

def pausa(min_s=0.5, max_s=3.5):
    time.sleep(random.uniform(min_s, max_s))

def login(driver, usuario, senha):
    wait_curto = WebDriverWait(driver, 3)
    wait_normal = WebDriverWait(driver, 15)

    try:
        botao_entrada = wait_curto.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//*[@id='social-azuread-aiko']/span",
            ))
        )
        driver.execute_script("arguments[0].click();", botao_entrada)
        return
    except TimeoutException:
        wait_normal.until(EC.visibility_of_element_located((By.ID, "UserName"))).send_keys(usuario)
        driver.find_element(By.ID, "Password").send_keys(senha, Keys.ENTER)

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

def click_forcado(driver, el):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        ActionChains(driver).move_to_element(el).pause(0.1).click(el).perform()
        return
    except Exception:
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

def confirmar_sim_se_existir(driver, timeout=3):
    wait_local = WebDriverWait(driver, timeout)

    seletores = [
        (By.CSS_SELECTOR, "button[data-test-ak-confirm-dialog-btn-confirm]"),
        (By.XPATH, "//button[normalize-space()='Sim']"),
        (By.XPATH, '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[18]/div[2]/div/div[2]/button[2]'),
        (By.XPATH, '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[19]/div[2]/div/div[2]/button[2]'),
    ]

    for by, sel in seletores:
        try:
            sim = wait_local.until(EC.element_to_be_clickable((by, sel)))
            click_forcado(driver, sim)
            time.sleep(0.3)
            return True
        except TimeoutException:
            pass

    return False

def esperar_loading(driver, timeout=15):
    wait_local = WebDriverWait(driver, timeout)
    try:
        wait_local.until(
            EC.invisibility_of_element_located((By.ID, "waiting-update"))
        )
    except TimeoutException:
        pass 

while True:
    usuario = input('Digite seu usuario (Sem a parte do @): ') + "@aiko.digital"
    senha = input('Digite sua senha: ')
    empresa = input('Digite qual empresa: ').upper()
    equipamento = input('Digite qual o modelo de equipamento (comodato, servico de campo, etc): ').upper()
    ticket = int(input("Digite o ticket (Somente o numero): "))
    zendesk = input("Ticket Zendesk (N se não tiver, somente o numero): ").strip()
    parou = int(input("Se parou em algum bordo, digite o número (0 se não parou): "))
    limite = int(input("Digite quantos bordos no total: "))
    grupos = coletar_lista("Grupos", grupos_default)  
    model  = coletar_lista("Modelos", model_default)   
    perfil = coletar_lista("Perfil", perfil_default)  

    while True:
        print("\n--- CONFIRA OS DADOS ---")
        print(f"  [1] Usuário:        {usuario}")
        print(f"  [2] Senha:          {senha}")
        print(f"  [3] Empresa:        {empresa}")
        print(f"  [4] Equipamento:    {equipamento}")
        print(f"  [5] Ticket:         {ticket}")
        print(f"  [6] T.Zendesk:      {zendesk}")
        print(f"  [7] Parou no bordo: {parou}")
        print(f"  [8] Q.Bordos:       {limite}")
        print(f"  [9] Grupos:         {grupos}")
        print(f"  [10] Modelos:       {model}")
        print(f"  [11] Perfil:         {perfil}")
        print("------------------------")

        confirma = input("Os dados estão corretos? (S/N): ").strip().upper()

        if confirma == "S":
            break

        print("\nQuais campos deseja alterar? (ex: 5  ou  1 3 5 9)")
        print("  [0] Refazer tudo")
        escolha = input(">>> ").strip()

        if escolha == "0":
            break

        for n in escolha.split():
            if n == "1":
                usuario = input('Digite seu usuario (Sem a parte do @): ') + "@aiko.digital"
            elif n == "2":
                senha = input('Digite sua senha: ')
            elif n == "3":
                empresa = input('Digite qual empresa: ').upper()
            elif n == "4":
                equipamento = input('Digite qual o modelo de equipamento (comodato, servico de campo, etc): ').upper()
            elif n == "5":
                ticket = int(input("Digite o ticket (Somente o numero): "))
            elif n == "6":
                zendesk = input("Ticket Zendesk (N se não tiver, somente o numero): ").strip()
            elif n == "7":
                parou = int(input("Se parou em algum bordo, digite o número (0 se não parou): "))
            elif n == "8":
                limite = int(input("Digite quantos bordos no total: "))
            elif n == "9":
                grupos = coletar_lista("Grupos", grupos_default)  
            elif n == "10":
                model = coletar_lista("Modelos", model_default)   
            elif n == "11":
                perfil = coletar_lista("Perfil", perfil_default)
                break

    if confirma == "S":
        break

print("\nDados confirmados! Iniciando...")

driver = webdriver.Edge()
driver.maximize_window()
driver.get(f"https://{empresa}.br.trackit.host/")

#LOGIN
login(driver, usuario, senha)

wait = WebDriverWait(driver, 20)
wait_curto = WebDriverWait(driver, 3)  

fechar_emergencias(driver)

# ENTRANDO NO CADASTRO DE BORDO
fechar_emergencias(driver)
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
    fechar_emergencias(driver)
    wait.until(EC.presence_of_element_located((By.ID, "app")))
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//*[@id="app"]/div[1]/div[5]/div/section/div/div[2]/a/span'))).click()
    pausa(1, 1.3)

    #Escreve o nome do bordo
    fechar_emergencias(driver)
    wait.until(EC.element_to_be_clickable((By.ID, "name"))).send_keys(bordo)
    pausa(1, 1.3)

    #Seleciona o Modelo
    fechar_emergencias(driver)
    esperar_loading(driver)
    clicar_dropdown(
        driver,
        xpaths_botao=[                          # Adiciona novos XPaths aqui
            '//*[@id="equipmentModel"]/div[1]',
            '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[4]/div[1]/div/div/div[1]',
        ],
        xpaths_opcoes=[                         # E aqui
            '//*[@id="equipmentModel"]/div[3]',
            '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[4]/div[1]/div/div/div[3]',
        ],
        condicao=lambda texto: any(p.lower() in texto for p in model)
    )
    pausa(0.5, 0.8)

    #Seleciona o Perfil
    fechar_emergencias(driver)
    esperar_loading(driver)
    clicar_dropdown(
        driver,
        xpaths_botao=[
            '//*[@id="networkProfiles"]/div[1]',
            '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[4]/div[3]/div/div/div[1]',
        ],
        xpaths_opcoes=[
            '//*[@id="networkProfiles"]/div[3]',
            '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[4]/div[3]/div/div/div[3]',
        ],
        condicao=lambda texto: any(p.lower() in texto for p in perfil)
    )
    pausa(0.5, 0.8)

    #Seleciona Grupo
    fechar_emergencias(driver)
    esperar_loading(driver)
    clicar_dropdown(
        driver,
        xpaths_botao=[
            '//*[@id="equipmentGroup"]/div[1]',
            '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[7]/div[1]/div/div/div/div[1]',
        ],
        xpaths_opcoes=[
            '//*[@id="equipmentGroup"]/div[3]',
            '//*[@id="app"]/div[1]/div[5]/div/section/div/div/div[2]/div/div/div[7]/div[1]/div/div/div/div[3]',
        ],
        condicao=lambda texto: any(p.lower() in texto for p in grupos)
    )
    pausa(0.5, 0.8)

    #Salva
    fechar_emergencias(driver)
    clicar_salvar_ultimo_visivel(driver, timeout=8)
    confirmar_sim_se_existir(driver, timeout=3)
    wait.until(EC.invisibility_of_element_located((By.ID, "name")))
    pausa(0.5, 1)

total_criados = limite - parou 
aviso_ok(
    "Automação finalizada",
    f"Foram criados {total_criados} bordos com sucesso.\nEmpresa: {empresa}\nTicket: {ticket}"
)
