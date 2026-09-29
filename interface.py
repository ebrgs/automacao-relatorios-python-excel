"""Interface gráfica para consolidar relatórios sem usar o terminal."""

from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from consolidar import BASE, Resultado, consolidar_pasta


AZUL = "#142C3F"
VERDE = "#218A7C"
FUNDO = "#F3F7F9"
CINZA = "#5D7180"


class Aplicativo(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Automação de Relatórios | Excel + Python")
        self.geometry("820x560")
        self.minsize(760, 520)
        self.configure(bg=FUNDO)
        self.pasta = tk.StringVar(value=str(BASE / "relatorios"))
        self.mensagens: queue.Queue[tuple[str, object]] = queue.Queue()
        self.processando = False
        self._montar()
        self.after(150, self._verificar_mensagens)

    def _montar(self) -> None:
        cabecalho = tk.Frame(self, bg=AZUL, height=155)
        cabecalho.pack(fill="x")
        cabecalho.pack_propagate(False)
        tk.Frame(cabecalho, bg=VERDE, width=8).pack(side="left", fill="y")
        bloco = tk.Frame(cabecalho, bg=AZUL)
        bloco.pack(side="left", padx=32, pady=27)
        tk.Label(bloco, text="AUTOMAÇÃO DE RELATÓRIOS", bg=AZUL, fg="white",
                 font=("Segoe UI Semibold", 22), anchor="w").pack(anchor="w")
        tk.Label(bloco, text="Excel + Python  •  consolidação e qualidade de dados",
                 bg=AZUL, fg="#B6CBD6", font=("Segoe UI", 11)).pack(anchor="w", pady=(7, 0))

        corpo = tk.Frame(self, bg=FUNDO)
        corpo.pack(fill="both", expand=True, padx=30, pady=26)
        tk.Label(corpo, text="01   SELECIONE A PASTA", bg=FUNDO, fg=VERDE,
                 font=("Segoe UI Semibold", 10)).pack(anchor="w")
        tk.Label(corpo, text="Escolha a pasta com as planilhas .xlsx das equipes.",
                 bg=FUNDO, fg=CINZA, font=("Segoe UI", 11)).pack(anchor="w", pady=(6, 16))

        selecao = tk.Frame(corpo, bg=FUNDO)
        selecao.pack(fill="x")
        self.campo = tk.Entry(selecao, textvariable=self.pasta, font=("Segoe UI", 11),
                              bg="white", fg=AZUL, relief="solid", bd=1)
        self.campo.pack(side="left", fill="x", expand=True, ipady=10)
        self.botao_pasta = tk.Button(selecao, text="Selecionar pasta", command=self._selecionar,
                                     bg="#DCE9EC", fg=AZUL, activebackground="#C9DFE2",
                                     relief="flat", cursor="hand2", font=("Segoe UI Semibold", 10),
                                     padx=18, pady=9)
        self.botao_pasta.pack(side="left", padx=(10, 0))

        tk.Label(corpo, text="02   PROCESSE OS ARQUIVOS", bg=FUNDO, fg=VERDE,
                 font=("Segoe UI Semibold", 10)).pack(anchor="w", pady=(32, 7))
        tk.Label(corpo, text="Os resultados serão salvos na pasta projeto2/resultado.",
                 bg=FUNDO, fg=CINZA, font=("Segoe UI", 11)).pack(anchor="w")
        linha_botoes = tk.Frame(corpo, bg=FUNDO)
        linha_botoes.pack(fill="x", pady=(18, 15))
        self.botao_processar = tk.Button(linha_botoes, text="Processar relatórios  →",
                                         command=self._processar, bg=VERDE, fg="white",
                                         activebackground="#176D63", activeforeground="white",
                                         relief="flat", cursor="hand2",
                                         font=("Segoe UI Semibold", 12), padx=22, pady=12)
        self.botao_processar.pack(side="left")
        self.botao_abrir = tk.Button(linha_botoes, text="Abrir resultados", command=self._abrir_resultados,
                                     bg="#DCE9EC", fg=AZUL, relief="flat", cursor="hand2",
                                     font=("Segoe UI Semibold", 10), padx=17, pady=13,
                                     state="disabled")
        self.botao_abrir.pack(side="left", padx=(12, 0))

        self.progresso = ttk.Progressbar(corpo, mode="indeterminate")
        self.progresso.pack(fill="x", pady=(0, 19))
        cartao = tk.Frame(corpo, bg="white", highlightbackground="#DDE8ED", highlightthickness=1)
        cartao.pack(fill="x")
        tk.Label(cartao, text="RESULTADO", bg="white", fg=VERDE,
                 font=("Segoe UI Semibold", 9)).pack(anchor="w", padx=20, pady=(17, 5))
        self.status = tk.Label(cartao, text="Pronto para processar as planilhas.",
                               bg="white", fg=AZUL, justify="left", anchor="w",
                               font=("Segoe UI", 11), wraplength=700)
        self.status.pack(fill="x", padx=20, pady=(0, 18))

    def _selecionar(self) -> None:
        selecionada = filedialog.askdirectory(initialdir=self.pasta.get() if Path(self.pasta.get()).is_dir() else BASE)
        if selecionada:
            self.pasta.set(selecionada)

    def _processar(self) -> None:
        entrada = Path(self.pasta.get().strip())
        if not entrada.is_dir():
            messagebox.showerror("Pasta inválida", "Selecione uma pasta existente com arquivos .xlsx.")
            return
        self.processando = True
        self.botao_processar.configure(state="disabled")
        self.botao_pasta.configure(state="disabled")
        self.botao_abrir.configure(state="disabled")
        self.progresso.start(12)
        self.status.configure(text="Lendo, validando e consolidando os arquivos...", fg=AZUL)
        threading.Thread(target=self._executar, args=(entrada,), daemon=True).start()

    def _executar(self, entrada: Path) -> None:
        try:
            resultado = consolidar_pasta(entrada, BASE / "resultado")
            self.mensagens.put(("sucesso", resultado))
        except Exception as exc:
            self.mensagens.put(("erro", str(exc)))

    def _verificar_mensagens(self) -> None:
        try:
            tipo, conteudo = self.mensagens.get_nowait()
        except queue.Empty:
            self.after(150, self._verificar_mensagens)
            return
        self.progresso.stop()
        self.botao_processar.configure(state="normal")
        self.botao_pasta.configure(state="normal")
        self.processando = False
        if tipo == "sucesso":
            resultado: Resultado = conteudo  # type: ignore[assignment]
            self.status.configure(
                text=(f"Concluído: {resultado.arquivos_lidos} arquivo(s), "
                      f"{resultado.linhas_validas} registro(s) válido(s), "
                      f"{resultado.linhas_descartadas} linha(s) descartada(s) e "
                      f"{resultado.inconsistencias} inconsistência(s).\n"
                      "Arquivos criados: consolidado.xlsx, inconsistencias.xlsx e resumo.xlsx."),
                fg=AZUL,
            )
            self.botao_abrir.configure(state="normal")
        else:
            self.status.configure(text=f"Não foi possível processar: {conteudo}", fg="#B04B30")
        self.after(150, self._verificar_mensagens)

    def _abrir_resultados(self) -> None:
        pasta = BASE / "resultado"
        if pasta.is_dir():
            os.startfile(pasta)  # type: ignore[attr-defined]  # Interface destinada ao Windows.


if __name__ == "__main__":
    Aplicativo().mainloop()
