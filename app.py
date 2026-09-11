"""Simple GUI to drive the CNPJ-checking automation against the Citrix-hosted app."""
import queue
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pyautogui

import auto_detect
import automation
import calibration
import io_utils

pyautogui.FAILSAFE = True  # jogue o mouse pro canto superior-esquerdo da tela para abortar na hora
pyautogui.PAUSE = 0.05

OUTPUT_DIR = Path(__file__).parent / "output"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automação Vivo Qualidade - Checagem de CNPJ")
        self.geometry("760x560")

        self.log_queue = queue.Queue()
        self.runner = None
        self.runner_thread = None
        self.records = []
        self.input_path = None
        self.is_paused = False

        self._build_ui()
        self.after(150, self._drain_log_queue)

    # ---------------- UI ----------------
    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Button(top, text="Testar calibração automática", command=self._on_auto_calibrate).grid(
            row=0, column=0, padx=4)
        ttk.Button(top, text="1) Calibrar manualmente", command=self._on_calibrate).grid(row=0, column=1, padx=4)
        ttk.Button(top, text="2) Upar lista (Excel)", command=self._on_upload).grid(row=0, column=2, padx=4)
        self.start_btn = ttk.Button(top, text="3) Iniciar", command=self._on_start, state="disabled")
        self.start_btn.grid(row=0, column=3, padx=4)
        self.pause_btn = ttk.Button(top, text="Pausar", command=self._on_pause_resume, state="disabled")
        self.pause_btn.grid(row=0, column=4, padx=4)
        self.stop_btn = ttk.Button(top, text="Parar", command=self._on_stop, state="disabled")
        self.stop_btn.grid(row=0, column=5, padx=4)
        ttk.Button(top, text="Calibrar aviso de tabela vazia (opcional)", command=self._on_calibrate_banners).grid(
            row=1, column=0, columnspan=2, padx=4, pady=(4, 0), sticky="w")

        top2 = ttk.Frame(self, padding=(10, 0))
        top2.pack(fill="x")
        self.dual_window_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            top2, text="Usar 2 janelas do Citrix lado a lado (~2x mais rápido)",
            variable=self.dual_window_var,
        ).pack(anchor="w")

        info = ttk.Frame(self, padding=(10, 0))
        info.pack(fill="x")
        self.file_label = ttk.Label(info, text="Nenhum arquivo carregado.")
        self.file_label.pack(anchor="w")

        prog_frame = ttk.Frame(self, padding=10)
        prog_frame.pack(fill="x")
        self.progress = ttk.Progressbar(prog_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x")
        self.progress_label = ttk.Label(prog_frame, text="0 / 0")
        self.progress_label.pack(anchor="e")

        console_frame = ttk.Frame(self, padding=10)
        console_frame.pack(fill="both", expand=True)
        ttk.Label(console_frame, text="Console:").pack(anchor="w")
        self.console = scrolledtext.ScrolledText(console_frame, state="disabled", wrap="word", height=20)
        self.console.pack(fill="both", expand=True)

        hint = ("Dica: jogue o mouse para o canto superior-esquerdo da tela a qualquer momento "
                "para abortar imediatamente (failsafe do pyautogui).")
        ttk.Label(self, text=hint, foreground="gray", padding=(10, 0, 10, 10)).pack(anchor="w")

    # ---------------- logging ----------------
    def log(self, msg: str):
        self.log_queue.put(msg)

    def _drain_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.console.configure(state="normal")
                ts = datetime.now().strftime("%H:%M:%S")
                self.console.insert("end", f"[{ts}] {msg}\n")
                self.console.see("end")
                self.console.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(150, self._drain_log_queue)

    # ---------------- actions ----------------
    def _on_calibrate(self):
        if self.runner_thread and self.runner_thread.is_alive():
            messagebox.showwarning("Aviso", "Pare a automação antes de calibrar novamente.")
            return

        def worker():
            ok = calibration.run_calibration(log=self.log)
            if ok:
                self.log("Calibração concluída com sucesso.")
            else:
                self.log("Calibração não foi concluída.")

        threading.Thread(target=worker, daemon=True).start()

    def _on_calibrate_banners(self):
        if self.runner_thread and self.runner_thread.is_alive():
            messagebox.showwarning("Aviso", "Pare a automação antes de calibrar novamente.")
            return

        def worker():
            ok = calibration.calibrate_result_banners(log=self.log)
            if ok:
                self.log("Calibração dos avisos concluída com sucesso.")
            else:
                self.log("Calibração dos avisos não foi concluída.")

        threading.Thread(target=worker, daemon=True).start()

    def _on_auto_calibrate(self):
        if self.runner_thread and self.runner_thread.is_alive():
            messagebox.showwarning("Aviso", "Pare a automação antes de recalibrar.")
            return

        def worker():
            ok = auto_detect.try_auto_calibrate(log=self.log)
            if ok:
                self.log("Calibração automática funcionou! Já pode iniciar.")
            else:
                self.log("Calibração automática não encontrou tudo. Rode a calibração manual (passo 1).")

        threading.Thread(target=worker, daemon=True).start()

    def _on_upload(self):
        path = filedialog.askopenfilename(
            title="Selecione a lista (Excel com colunas CEP, Número e CNPJ)",
            filetypes=[("Excel files", "*.xlsx *.xls")],
        )
        if not path:
            return
        try:
            records, columns_used = io_utils.load_search_list_from_excel(path)
        except Exception as exc:
            messagebox.showerror("Erro ao ler arquivo", str(exc))
            return

        self.records = records
        self.input_path = path
        cnpj_col = columns_used["cnpj"] or "(nenhuma -- usando CEP+Número como identificador)"
        self.file_label.config(
            text=f"{Path(path).name} — CEP: '{columns_used['cep']}', Número: '{columns_used['numero']}', "
                 f"CNPJ: '{cnpj_col}' — {len(records)} registros carregados."
        )
        self.log(f"Lista carregada: {len(records)} registros "
                 f"(CEP: '{columns_used['cep']}', Número: '{columns_used['numero']}', CNPJ: '{cnpj_col}').")
        self.progress.configure(maximum=len(records), value=0)
        self.progress_label.config(text=f"0 / {len(records)}")
        self.start_btn.configure(state="normal")

    def _on_start(self):
        config = calibration.load_config()
        if config is None:
            messagebox.showwarning("Calibração pendente", "Rode a calibração (passo 1) antes de iniciar.")
            return
        if not self.records:
            messagebox.showwarning("Lista pendente", "Upe uma lista (passo 2) antes de iniciar.")
            return

        use_dual = self.dual_window_var.get()
        win_a = win_b = None
        if use_dual:
            windows = auto_detect.locate_all_windows(log=self.log)
            if len(windows) < 2:
                self.log(f"Modo 2 janelas: só encontrei {len(windows)} janela(s) na tela. "
                         f"Seguindo no modo normal (1 janela) em vez de travar.")
                use_dual = False
            else:
                win_a, win_b = windows[0], windows[1]
                self.log(f"Modo 2 janelas: janela A em ({win_a.left},{win_a.top}), "
                         f"janela B em ({win_b.left},{win_b.top}).")

        signature = io_utils.file_signature(self.input_path)
        self.runner = automation.AutomationRunner(
            config=config,
            records=self.records,
            log=self.log,
            on_progress=self._on_progress,
            on_logout=self._on_logout_detected,
            input_signature=signature,
        )

        self.start_btn.configure(state="disabled")
        self.pause_btn.configure(state="normal", text="Pausar")
        self.stop_btn.configure(state="normal")
        self.is_paused = False

        def worker():
            if use_dual:
                results = self.runner.run_dual(win_a, win_b)
            else:
                results = self.runner.run()
            self._finish_run(results)

        self.runner_thread = threading.Thread(target=worker, daemon=True)
        self.runner_thread.start()

    def _on_pause_resume(self):
        if not self.runner:
            return
        if self.is_paused:
            self.runner.resume()
            self.pause_btn.configure(text="Pausar")
            self.is_paused = False
        else:
            self.runner.pause()
            self.pause_btn.configure(text="Retomar")
            self.is_paused = True

    def _on_stop(self):
        if not self.runner:
            return
        if messagebox.askyesno("Confirmar", "Parar a automação? O progresso fica salvo e você pode retomar depois."):
            self.runner.stop()
            self.stop_btn.configure(state="disabled")

    def _on_logout_detected(self):
        self.pause_btn.configure(text="Retomar")
        self.is_paused = True
        self.log("*** Ação necessária: verifique/relogue no Citrix e clique em 'Retomar'. ***")
        try:
            self.bell()
        except Exception:
            pass

    def _on_progress(self, index, total, cnpj, status):
        self.progress.configure(value=index)
        self.progress_label.config(text=f"{index} / {total}  (último: {cnpj} -> {status})")

    def _finish_run(self, results):
        self.start_btn.configure(state="normal")
        self.pause_btn.configure(state="disabled")
        self.stop_btn.configure(state="disabled")

        OUTPUT_DIR.mkdir(exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = OUTPUT_DIR / f"resultado_{ts}.xlsx"
        io_utils.write_output_excel(results, str(out_path))
        self.log(f"Arquivo de saída gerado: {out_path}")

        errors_path = OUTPUT_DIR / f"erros_{ts}.xlsx"
        if io_utils.write_errors_excel(results, str(errors_path)):
            self.log(f"Planilha de erros gerada: {errors_path}")

        kept = sum(1 for s in results.values() if s == "mantido")
        eliminated = sum(1 for s in results.values() if s == "eliminado")
        errors = sum(1 for s in results.values() if s == "erro")
        invalid = sum(1 for s in results.values() if s == "cep_ou_numero_invalido")
        self.log(f"Resumo: {kept} mantidos, {eliminated} eliminados, {errors} com erro (após nova tentativa), "
                 f"{invalid} com CEP/Número inválido ou vazio na planilha.")

        if self.runner and self.runner.index >= len(self.runner.records):
            self.runner.clear_checkpoint()


if __name__ == "__main__":
    App().mainloop()
