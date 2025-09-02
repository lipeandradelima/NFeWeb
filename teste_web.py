import os
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

NOME_PASTA_SAIDA = "XMLs_Baixados"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAIDA_DIR = os.path.join(BASE_DIR, NOME_PASTA_SAIDA)
FSIST_URL = "https://www.fsist.com.br/"

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Automação NF-e FSist")
        self.root.geometry("800x550")

        self.chaves = []
        self.driver = None
        self.processing = False

        self.setup_ui()

    def setup_ui(self):
        self.btn_load = ttk.Button(self.root, text="Selecionar arquivo Excel", command=self.load_excel)
        self.btn_load.pack(pady=10)

        columns = ("chave", "status")
        self.tree = ttk.Treeview(self.root, columns=columns, show="headings")
        self.tree.heading("chave", text="Chave NF-e (44 caracteres)")
        self.tree.heading("status", text="Status")
        self.tree.column("chave", width=550)
        self.tree.column("status", width=150, anchor=tk.CENTER)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tree.tag_configure('baixado', foreground='green')

        self.btn_start = ttk.Button(self.root, text="Iniciar Automação", command=self.start_automation, state=tk.DISABLED)
        self.btn_start.pack(pady=10)

    def load_excel(self):
        filepath = filedialog.askopenfilename(
            title="Selecione o arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if not filepath:
            return

        try:
            df = pd.read_excel(filepath, dtype=str)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler o arquivo Excel:\n{e}")
            return

        self.chaves.clear()
        self.tree.delete(*self.tree.get_children())

        found = False
        for col in df.columns:
            valid_chaves = []
            for val in df[col].dropna():
                if isinstance(val, str):
                    val_str = val.strip()
                    if len(val_str) == 44:
                        valid_chaves.append(val_str)
            if valid_chaves:
                self.chaves = valid_chaves
                found = True
                break

        if not found:
            messagebox.showwarning("Aviso", "Nenhuma coluna com chaves de 44 caracteres encontrada!")
            self.btn_start.config(state=tk.DISABLED)
            return

        for chave in self.chaves:
            if chave:
                self.tree.insert("", tk.END, values=(chave, "Espera"))
        self.btn_start.config(state=tk.NORMAL)

    def start_automation(self):
        if self.processing:
            return
        self.processing = True
        self.btn_start.config(state=tk.DISABLED)
        threading.Thread(target=self.automate, daemon=True).start()

    def update_status(self, index, status):
        def update():
            item_id = self.tree.get_children()[index]
            if status == "Baixado":
                display_status = "✔ Baixado"
                self.tree.item(item_id, values=(self.chaves[index], display_status), tags=('baixado',))
            else:
                self.tree.item(item_id, values=(self.chaves[index], status), tags=())
        self.root.after(0, update)

    def automate(self):
        os.makedirs(SAIDA_DIR, exist_ok=True)

        options = webdriver.ChromeOptions()
        prefs = {"download.default_directory": os.path.abspath(SAIDA_DIR)}
        options.add_experimental_option("prefs", prefs)
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        options.add_argument("--disable-blink-features=AutomationControlled")

        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=options)
        wait = WebDriverWait(self.driver, 120)

        try:
            for idx, chave in enumerate(self.chaves):
                self.update_status(idx, "Processando")
                self.driver.get(FSIST_URL)

                campo_chave = wait.until(EC.presence_of_element_located((By.ID, "chave")))
                campo_chave.clear()
                campo_chave.send_keys(chave)

                botao_consulta = self.driver.find_element(By.ID, "butconsulta")
                botao_consulta.click()

                botao_clicado = False
                start_time = time.time()
                timeout = 200
                check_interval = 0.25

                while time.time() - start_time < timeout:
                    try:
                        btn = self.driver.find_element(By.ID, "butComCertificado")
                        if btn.is_displayed() and btn.is_enabled():
                            self.driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                            self.driver.execute_script("arguments[0].click();", btn)
                            botao_clicado = True
                            break
                    except Exception:
                        pass
                    time.sleep(check_interval)

                if botao_clicado:
                    self.update_status(idx, "Baixado")
                else:
                    self.update_status(idx, "Erro: botão não encontrado")

                time.sleep(1)

        except Exception as e:
            def show_error():
                messagebox.showerror("Erro", f"Erro durante automação:\n{e}")
            self.root.after(0, show_error)
        finally:
            self.driver.quit()
            self.processing = False
            self.root.after(0, lambda: self.btn_start.config(state=tk.NORMAL))
            self.root.after(0, lambda: messagebox.showinfo("Info", "Automação finalizada!"))

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
