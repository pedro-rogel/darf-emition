"""Prepara as dependências na primeira abertura, sem pedir comandos ao usuário."""

from __future__ import annotations

from pathlib import Path
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk


BASE_DIR = Path(__file__).resolve().parent
PYTHON_VENV = BASE_DIR / "venv" / "Scripts" / "python.exe"
PYTHONW_VENV = BASE_DIR / "venv" / "Scripts" / "pythonw.exe"
LOG = BASE_DIR / "instalacao.log"
SEM_JANELA = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def executar(comando: list[str]) -> None:
    resultado = subprocess.run(
        comando, cwd=BASE_DIR, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace", creationflags=SEM_JANELA,
    )
    with LOG.open("a", encoding="utf-8") as arquivo:
        arquivo.write(f"\n$ {' '.join(comando)}\n{resultado.stdout}\n")
    if resultado.returncode:
        raise RuntimeError(
            f"A instalação falhou (código {resultado.returncode}). "
            f"Consulte o arquivo {LOG.name} na pasta do programa."
        )


def main() -> None:
    janela = tk.Tk()
    janela.title("Preparando a emissão de DARF")
    janela.geometry("480x160")
    janela.resizable(False, False)
    janela.configure(padx=24, pady=20)
    status = tk.StringVar(value="Preparando o programa para o primeiro uso...")
    ttk.Label(janela, textvariable=status, wraplength=430,
              font=("Segoe UI", 11)).pack(anchor="w", pady=(0, 18))
    barra = ttk.Progressbar(janela, mode="indeterminate")
    barra.pack(fill="x")
    barra.start(12)
    eventos: list[tuple[str, str]] = []

    def preparar() -> None:
        try:
            if not PYTHON_VENV.is_file():
                executar([sys.executable, "-m", "venv", str(BASE_DIR / "venv")])
            executar([str(PYTHON_VENV), "-m", "pip", "install", "-r",
                      str(BASE_DIR / "requirements.txt")])
            eventos.append(("ok", ""))
        except Exception as exc:
            eventos.append(("erro", str(exc)))

    def verificar() -> None:
        if not eventos:
            janela.after(150, verificar)
            return
        tipo, detalhe = eventos.pop(0)
        barra.stop()
        if tipo == "erro":
            status.set("Não foi possível preparar o programa.")
            messagebox.showerror("Instalação", detalhe, parent=janela)
        else:
            janela.destroy()
            subprocess.Popen([str(PYTHONW_VENV), str(BASE_DIR / "gui.pyw")],
                             cwd=BASE_DIR, creationflags=SEM_JANELA)

    threading.Thread(target=preparar, daemon=True).start()
    janela.after(150, verificar)
    janela.mainloop()


if __name__ == "__main__":
    main()
