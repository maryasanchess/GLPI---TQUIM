# -*- coding: utf-8 -*-
"""App de um clique pra gerar os relatórios de TQUIM > Ocorrências.

Dá dois cliques neste arquivo (não precisa abrir terminal). Cada botão roda
o script correspondente e, quando termina, já abre o arquivo gerado
automaticamente.

Se quiser rodar sem interface (linha de comando), os scripts continuam
funcionando normalmente: exportar_validacao_qualidade.py e
exportar_relatorio_ocorrencias.py.
"""
import os
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox, scrolledtext

PASTA = Path(__file__).resolve().parent
PYTHON = sys.executable
DOWNLOADS = Path.home() / "Downloads"  # pasta Downloads do usuário que estiver logado, em qualquer PC

BOTOES = [
    {
        "titulo": "Validação Código 4",
        "descricao": "Ocorrências pendentes de decisão da Qualidade (culpabilidade do motorista)",
        "script": "exportar_validacao_qualidade.py",
        "nome_arquivo": lambda: f"Validacao_Qualidade_{datetime.now():%Y-%m}.xlsx",
    },
    {
        "titulo": "Relatório Anual / Mensal",
        "descricao": "Todos os chamados do ano, com uma aba por mês",
        "script": "exportar_relatorio_ocorrencias.py",
        "nome_arquivo": lambda: f"Relatorio_Ocorrencias_{datetime.now():%Y}.xlsx",
    },
]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TQUIM > Ocorrências — Relatórios")
        self.geometry("560x400")
        self.configure(bg="#f2f4f8")
        self.resizable(False, False)

        tk.Label(
            self, text="Relatórios de Ocorrências", font=("Segoe UI", 16, "bold"),
            bg="#f2f4f8", fg="#1f4e79",
        ).pack(pady=(18, 2))
        tk.Label(
            self, text="Clique num relatório pra gerar e abrir o arquivo automaticamente.",
            font=("Segoe UI", 9), bg="#f2f4f8", fg="#555",
        ).pack(pady=(0, 14))

        for item in BOTOES:
            self._criar_bloco(item)

        tk.Label(self, text="Andamento", font=("Segoe UI", 9, "bold"), bg="#f2f4f8", fg="#333").pack(
            anchor="w", padx=20, pady=(10, 2)
        )
        self.log = scrolledtext.ScrolledText(self, height=8, font=("Consolas", 9), state="disabled")
        self.log.pack(fill="both", expand=True, padx=20, pady=(0, 16))

    def _criar_bloco(self, item):
        frame = tk.Frame(self, bg="#ffffff", highlightbackground="#d4d9e2", highlightthickness=1)
        frame.pack(fill="x", padx=20, pady=6)

        textos = tk.Frame(frame, bg="#ffffff")
        textos.pack(side="left", fill="x", expand=True, padx=12, pady=10)
        tk.Label(textos, text=item["titulo"], font=("Segoe UI", 11, "bold"), bg="#ffffff", fg="#1f4e79").pack(
            anchor="w"
        )
        tk.Label(
            textos, text=item["descricao"], font=("Segoe UI", 8), bg="#ffffff", fg="#666", wraplength=360,
            justify="left",
        ).pack(anchor="w")

        btn = tk.Button(
            frame, text="Gerar", font=("Segoe UI", 10, "bold"), bg="#1f4e79", fg="white",
            activebackground="#16385a", activeforeground="white", relief="flat", padx=14, pady=6,
            command=lambda i=item: self._rodar(i),
        )
        btn.pack(side="right", padx=12)

    def _escrever(self, texto):
        self.log.configure(state="normal")
        self.log.insert("end", texto + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _rodar(self, item):
        self._escrever(f"\n--- {item['titulo']} ---")
        thread = threading.Thread(target=self._executar_script, args=(item,), daemon=True)
        thread.start()

    def _executar_script(self, item):
        caminho_script = PASTA / item["script"]
        DOWNLOADS.mkdir(parents=True, exist_ok=True)
        caminho_saida = DOWNLOADS / item["nome_arquivo"]()
        try:
            resultado = subprocess.run(
                [PYTHON, str(caminho_script), str(caminho_saida)],
                cwd=str(PASTA), capture_output=True, text=True, encoding="utf-8", errors="replace",
            )
        except Exception as e:
            self._escrever(f"ERRO ao rodar: {e}")
            return

        saida = (resultado.stdout or "") + (resultado.stderr or "")
        for linha in saida.strip().splitlines():
            self._escrever("  " + linha)

        if resultado.returncode != 0:
            self._escrever("Terminou com erro.")
            self.after(0, lambda: messagebox.showerror(item["titulo"], "O relatório não foi gerado. Veja o andamento pra detalhes."))
            return

        if caminho_saida.exists():
            self._escrever(f"Pronto: {caminho_saida}")
            try:
                os.startfile(str(caminho_saida))
            except Exception as e:
                self._escrever(f"(não consegui abrir automaticamente: {e})")
        else:
            self._escrever("Concluído sem gerar arquivo novo (pode ser que não havia nada pra exportar).")


if __name__ == "__main__":
    App().mainloop()
