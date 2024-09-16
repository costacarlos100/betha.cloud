import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from openpyxl import Workbook
import time
import threading

class AutomatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatizador Betha.Cloud")

        self.status_var = tk.StringVar()
        self.status_var.set("Clique no botão para iniciar a automação.")

        self.progress_var = tk.DoubleVar()
        self.progress_var.set(0)

        # Layout da interface gráfica
        tk.Label(self.root, text="Status:").pack(padx=10, pady=5)
        self.status_label = tk.Label(self.root, textvariable=self.status_var)
        self.status_label.pack(padx=10, pady=5)

        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=400, mode="determinate", variable=self.progress_var)
        self.progress_bar.pack(padx=10, pady=20)

        self.start_button = tk.Button(self.root, text="Iniciar Automação", command=self.on_start_click)
        self.start_button.pack(padx=10, pady=10)

    def update_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()

    def update_progress(self, value):
        self.progress_var.set(value)
        self.root.update_idletasks()

    def get_credentials(self):
        class CredentialsWindow:
            def __init__(self, root):
                self.root = root
                self.login = None
                self.senha = None
                self.descricao = None
                self.create_window()

            def create_window(self):
                self.window = tk.Toplevel(self.root)
                self.window.title("Faça seu Login")

                tk.Label(self.window, text="Login:").pack(padx=10, pady=5)
                self.entry_login = tk.Entry(self.window)
                self.entry_login.pack(padx=10, pady=5)

                tk.Label(self.window, text="Senha:").pack(padx=10, pady=5)
                self.entry_senha = tk.Entry(self.window, show='*')
                self.entry_senha.pack(padx=10, pady=5)

                tk.Label(self.window, text="Descrição:").pack(padx=10, pady=5)
                self.entry_descricao = tk.Entry(self.window)
                self.entry_descricao.pack(padx=10, pady=5)

                tk.Button(self.window, text="OK", command=self.submit).pack(padx=10, pady=10)

            def submit(self):
                self.login = self.entry_login.get()
                self.senha = self.entry_senha.get()
                self.descricao = self.entry_descricao.get()
                self.window.destroy()

        cred_window = CredentialsWindow(self.root)
        self.root.wait_window(cred_window.window)
        return cred_window.login, cred_window.senha, cred_window.descricao

    def start_automation(self, login, senha, descricao):
        def automation_task():
            options = Options()
            options.headless = True

            service = Service(GeckoDriverManager().install())
            driver = webdriver.Firefox(service=service, options=options)

            try:
                driver.get("https://betha.cloud")
                wdw = WebDriverWait(driver, 10)

                def wait_until(xpath, by=By.XPATH):
                    return wdw.until(EC.presence_of_element_located((by, xpath)))

                def click_element(xpath):
                    element = wait_until(xpath)
                    element.click()

                def send_keys_to_element(xpath, keys):
                    element = wait_until(xpath)
                    element.send_keys(keys)

                self.update_status("Realizando login...")
                send_keys_to_element('//*[@id="login:iUsuarios"]', login)
                send_keys_to_element('//*[@id="login:senha"]', senha)
                click_element('//*[@id="login:btAcessar"]')

                self.update_status("Navegando para a página de compras...")
                click_element('/html/body/div[3]/div/div/div/div[2]/ul/li[3]/div/div[2]/div[2]/a')
                click_element('/html/body/div[3]/div/div/div/div[2]/ul/li[1]/div/div/div[2]/div')
                click_element('//*[@id="hrefSelecaoEntidadeCtrlSelecionar-1"]')
                click_element('//*[@id="hrefvmSelecionar-2024"]')
                click_element('/html/body/div[3]/div/div/div[2]/div/div/ul/li[1]/div')
                click_element('/html/body/div[1]/header/div[1]/div/ul/li[4]/a')

                self.update_status("Selecionando fornecedores...")
                wdw.until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/header/div[1]/div/ul/li[4]/div/div/div[2]/ul[1]/li[2]/a')))
                click_element('/html/body/div[1]/header/div[1]/div/ul/li[4]/div/div/div[2]/ul[1]/li[2]/a')
                
                time.sleep(3)
                click_element('/html/body/div[3]/div/div/div[2]/div/div[1]/ui-list-grid/div[1]/div[2]/div/div/ui-search/div/input')
                send_keys_to_element('/html/body/div[3]/div/div/div[2]/div/div[1]/ui-list-grid/div[1]/div[2]/div/div/ui-search/div/input', descricao)
                click_element('/html/body/div[3]/div/div/div[2]/div/div[1]/ui-list-grid/div[1]/div[2]/div/div/ui-search/div/div/button[1]')

                self.update_status("Selecionando itens...")
                click_element('/html/body/div[3]/div/div/div[2]/div/div[1]/ui-list-grid/div[3]/div[2]/div[3]/div/ui-pagination/div[1]/form/div/a')
                click_element('/html/body/div[7]/ul/li[3]')

                texto = wait_until('/html/body/div[3]/div/div/div[2]/div/div[1]/ui-list-grid/div[3]/div[2]/div[3]/div/ui-pagination/div[1]/form/span').text
                ultimo_item = int(texto.split()[-1])
                
                workbook = Workbook()
                sheet = workbook.active
                sheet.title = "Emails"
                sheet.append(["Email"])

                unique_emails = set()

                for n1 in range(1, ultimo_item + 1):
                    self.update_status(f"Processando item {n1}/{ultimo_item}...")
                    self.update_progress(n1 / ultimo_item * 100)
                    time.sleep(3)
                    click_element(f'/html/body/div[3]/div/div/div[2]/div/div[1]/ui-list-grid/div[3]/div[2]/div[2]/div/table/tbody/tr[{n1}]/td[3]/a')
                    email_element = wait_until(f'/html/body/div[9]/div/div/div/fieldset/div[2]/div/div[2]/div[2]/div/div/div[2]/div[3]/div[2]/h5')
                    email = email_element.text
                    
                    if email and len(email) >= 3 and email not in unique_emails:
                        unique_emails.add(email)
                        sheet.append([email])
                    
                    click_element(f'/html/body/div[9]/div/div/div/fieldset/div[3]/button[2]')

                workbook.save("emails_encontrados.xlsx")
                
                self.update_status("Processo concluído!")
            finally:
                driver.quit()

        threading.Thread(target=automation_task).start()

    def on_start_click(self):
        login, senha, descricao = self.get_credentials()
        if login and senha and descricao:
            self.start_automation(login, senha, descricao)
        else:
            self.update_status("Por favor, preencha todas as informações.")

if __name__ == "__main__":
    root = tk.Tk()
    app = AutomatorApp(root)
    root.mainloop()
