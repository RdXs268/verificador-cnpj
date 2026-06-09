"""
Verificador de CNPJ - Inapto / Baixado
Consulta a API pública da Receita Federal via receitaws.com.br
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import requests
import re
import time
import threading
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ── Paleta de cores ─────────────────────────────────────────────────────────
BG        = "#1e1e2e"
SURFACE   = "#2a2a3e"
ACCENT    = "#7c5cbf"
ACCENT2   = "#9b77e0"
TEXT      = "#e0e0f0"
TEXT_DIM  = "#888aaa"
RED       = "#e05555"
GREEN     = "#55c987"
YELLOW    = "#e0b855"
BORDER    = "#3a3a55"

# ── Helpers ──────────────────────────────────────────────────────────────────

def limpar_cnpj(cnpj: str) -> str:
    return re.sub(r"\D", "", cnpj)

def formatar_cnpj(cnpj: str) -> str:
    c = limpar_cnpj(cnpj)
    if len(c) == 14:
        return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
    return cnpj

def validar_cnpj(cnpj: str) -> bool:
    c = limpar_cnpj(cnpj)
    if len(c) != 14 or len(set(c)) == 1:
        return False
    def calc(c, n):
        s = sum(int(c[i]) * ((n - i) % 8 + 2) for i in range(n - 1))
        r = 11 - s % 11
        return 0 if r >= 10 else r
    return int(c[12]) == calc(c, 13) and int(c[13]) == calc(c, 14)

def consultar_cnpj(cnpj: str) -> dict:
    """Consulta receitaws.com.br (sem autenticação)."""
    c = limpar_cnpj(cnpj)
    url = f"https://receitaws.com.br/v1/cnpj/{c}"
    headers = {"User-Agent": "VerificadorCNPJ/1.0"}
    try:
        r = requests.get(url, headers=headers, timeout=15)
        if r.status_code == 429:
            return {"erro": "rate_limit"}
        r.raise_for_status()
        return r.json()
    except requests.exceptions.ConnectionError:
        return {"erro": "sem_conexao"}
    except Exception as e:
        return {"erro": str(e)}

def classificar_situacao(data: dict):
    """Retorna (situacao_label, cor, situacao_raw)."""
    if "erro" in data:
        e = data["erro"]
        if e == "rate_limit":
            return "Aguardando (limite API)", YELLOW, "rate_limit"
        if e == "sem_conexao":
            return "Sem conexão", RED, "sem_conexao"
        if "CNPJ inválido" in e or "cnpj" in e.lower():
            return "CNPJ inválido", RED, "invalido"
        return f"Erro: {e}", RED, "erro"

    situacao = data.get("situacao", "").upper()
    if situacao == "INAPTA":
        return "INAPTO", RED, situacao
    elif situacao == "BAIXADA":
        return "BAIXADO", RED, situacao
    elif situacao == "ATIVA":
        return "ATIVA", GREEN, situacao
    elif situacao == "SUSPENSA":
        return "SUSPENSA", YELLOW, situacao
    elif situacao == "NULA":
        return "NULA", RED, situacao
    else:
        return situacao or "Desconhecida", TEXT_DIM, situacao

# ── Janela principal ──────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Verificador de CNPJ")
        self.geometry("860x640")
        self.minsize(720, 500)
        self.configure(bg=BG)
        self.resultados = []          # lista de dicts para exportação
        self._build_ui()

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        # Título
        hdr = tk.Frame(self, bg=ACCENT, height=52)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🔍  Verificador de CNPJ", bg=ACCENT,
                 fg="white", font=("Segoe UI", 15, "bold")).pack(side="left", padx=20, pady=10)

        # Área de entrada
        entry_frame = tk.Frame(self, bg=SURFACE, padx=14, pady=12)
        entry_frame.pack(fill="x", padx=12, pady=(12, 4))

        tk.Label(entry_frame, text="Cole os CNPJs abaixo (um por linha ou separados por vírgula / espaço):",
                 bg=SURFACE, fg=TEXT, font=("Segoe UI", 9)).pack(anchor="w")

        txt_wrap = tk.Frame(entry_frame, bg=BORDER, bd=1)
        txt_wrap.pack(fill="both", expand=True, pady=(6, 0))

        self.txt_input = tk.Text(txt_wrap, height=5, bg="#12121e", fg=TEXT,
                                 insertbackground=TEXT, font=("Courier New", 10),
                                 relief="flat", bd=6, wrap="word")
        self.txt_input.pack(fill="both", expand=True)

        # Botões
        btn_row = tk.Frame(self, bg=BG)
        btn_row.pack(fill="x", padx=12, pady=6)

        self._btn(btn_row, "▶  Verificar", self._iniciar_verificacao, ACCENT).pack(side="left", padx=(0, 6))
        self._btn(btn_row, "🗑  Limpar", self._limpar_tudo, SURFACE).pack(side="left", padx=(0, 6))
        self._btn(btn_row, "📥  Exportar Excel", self._exportar_excel, "#2e6e45").pack(side="left")

        # Barra de progresso
        self.progress_var = tk.DoubleVar()
        self.progress_lbl = tk.Label(self, text="", bg=BG, fg=TEXT_DIM, font=("Segoe UI", 8))
        self.progress_lbl.pack(anchor="w", padx=14)
        self.progressbar = ttk.Progressbar(self, variable=self.progress_var,
                                           maximum=100, length=200)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TProgressbar", troughcolor=SURFACE, background=ACCENT2,
                        thickness=6, bordercolor=BG, lightcolor=ACCENT2, darkcolor=ACCENT2)
        self.progressbar.pack(fill="x", padx=12, pady=(2, 6))

        # Tabela de resultados
        cols = ("CNPJ", "Razão Social", "Situação", "Motivo", "UF")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=14)

        style.configure("Treeview", background=SURFACE, foreground=TEXT,
                        fieldbackground=SURFACE, rowheight=26, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", background=ACCENT, foreground="white",
                        font=("Segoe UI", 9, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", ACCENT)])

        widths = {"CNPJ": 160, "Razão Social": 280, "Situação": 100, "Motivo": 180, "UF": 50}
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=widths.get(c, 120), anchor="w")

        self.tree.tag_configure("red",    foreground=RED)
        self.tree.tag_configure("green",  foreground=GREEN)
        self.tree.tag_configure("yellow", foreground=YELLOW)
        self.tree.tag_configure("dim",    foreground=TEXT_DIM)

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(fill="both", expand=True, padx=12, pady=(0, 4), side="left")
        vsb.pack(fill="y", padx=(0, 12), pady=(0, 4), side="right")

        # Rodapé
        foot = tk.Frame(self, bg=SURFACE, pady=5)
        foot.pack(fill="x", side="bottom")
        tk.Label(foot, text="Fonte: receitaws.com.br  |  Dados da Receita Federal",
                 bg=SURFACE, fg=TEXT_DIM, font=("Segoe UI", 8)).pack()

    def _btn(self, parent, text, cmd, bg):
        return tk.Button(parent, text=text, command=cmd, bg=bg, fg="white",
                         font=("Segoe UI", 9, "bold"), relief="flat", bd=0,
                         padx=14, pady=7, cursor="hand2",
                         activebackground=ACCENT2, activeforeground="white")

    # ── Verificação ───────────────────────────────────────────────────────────

    def _iniciar_verificacao(self):
        raw = self.txt_input.get("1.0", "end").strip()
        if not raw:
            messagebox.showwarning("Atenção", "Informe ao menos um CNPJ.")
            return

        tokens = re.split(r"[\s,;]+", raw)
        cnpjs = [limpar_cnpj(t) for t in tokens if limpar_cnpj(t)]
        cnpjs = list(dict.fromkeys(cnpjs))          # remove duplicatas mantendo ordem

        if not cnpjs:
            messagebox.showwarning("Atenção", "Nenhum CNPJ encontrado no texto.")
            return

        # Limpa tabela e resultados anteriores
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.resultados.clear()
        self.progress_var.set(0)

        threading.Thread(target=self._verificar_thread, args=(cnpjs,), daemon=True).start()

    def _verificar_thread(self, cnpjs):
        total = len(cnpjs)
        for i, cnpj in enumerate(cnpjs, 1):
            self.progress_lbl.config(text=f"Consultando {i}/{total}: {formatar_cnpj(cnpj)}")

            if not validar_cnpj(cnpj):
                data = {"erro": "CNPJ inválido"}
            else:
                data = consultar_cnpj(cnpj)

                # Retry automático se rate limit
                retries = 0
                while data.get("erro") == "rate_limit" and retries < 5:
                    wait = 20 + retries * 10
                    self.progress_lbl.config(text=f"Limite de requisições — aguardando {wait}s...")
                    time.sleep(wait)
                    data = consultar_cnpj(cnpj)
                    retries += 1

            label, cor, situacao_raw = classificar_situacao(data)
            razao   = data.get("nome", "—")
            motivo  = data.get("motivo_situacao", "—") or "—"
            uf      = data.get("uf", "—") or "—"

            tag = {RED: "red", GREEN: "green", YELLOW: "yellow"}.get(cor, "dim")

            self.tree.insert("", "end",
                             values=(formatar_cnpj(cnpj), razao, label, motivo, uf),
                             tags=(tag,))

            self.resultados.append({
                "cnpj":     formatar_cnpj(cnpj),
                "razao":    razao,
                "situacao": label,
                "motivo":   motivo,
                "uf":       uf,
                "raw":      situacao_raw,
            })

            self.progress_var.set(i / total * 100)

            # Respeita o rate-limit da API (máx 3 req/min no plano gratuito)
            if i < total:
                time.sleep(21)

        self.progress_lbl.config(text=f"✅  Concluído — {total} CNPJ(s) verificado(s).")

    # ── Export ────────────────────────────────────────────────────────────────

    def _exportar_excel(self):
        if not self.resultados:
            messagebox.showinfo("Atenção", "Nenhum resultado para exportar.")
            return

        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"cnpjs_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        )
        if not path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Verificação CNPJ"

        # Estilos
        hdr_fill  = PatternFill("solid", fgColor="5A3E9B")
        hdr_font  = Font(bold=True, color="FFFFFF", name="Segoe UI", size=10)
        red_fill  = PatternFill("solid", fgColor="F5CCCC")
        grn_fill  = PatternFill("solid", fgColor="C6EFCE")
        yel_fill  = PatternFill("solid", fgColor="FFEB9C")
        thin_side = Side(style="thin", color="CCCCCC")
        border    = Border(thin_side, thin_side, thin_side, thin_side)
        center    = Alignment(horizontal="center", vertical="center")

        headers = ["CNPJ", "Razão Social", "Situação", "Motivo da Situação", "UF"]
        ws.append(headers)
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = center
            cell.border = border
        ws.row_dimensions[1].height = 22

        for r in self.resultados:
            row = [r["cnpj"], r["razao"], r["situacao"], r["motivo"], r["uf"]]
            ws.append(row)
            last = ws.max_row
            raw = r.get("raw", "")
            fill = None
            if raw in ("INAPTA", "BAIXADA", "NULA", "invalido", "erro", "sem_conexao"):
                fill = red_fill
            elif raw == "ATIVA":
                fill = grn_fill
            elif raw in ("SUSPENSA", "rate_limit"):
                fill = yel_fill
            for cell in ws[last]:
                if fill:
                    cell.fill = fill
                cell.border = border
                cell.alignment = Alignment(vertical="center")
            ws.row_dimensions[last].height = 18

        # Larguras
        for col, w in zip("ABCDE", [22, 45, 14, 35, 6]):
            ws.column_dimensions[col].width = w

        # Filtro automático
        ws.auto_filter.ref = ws.dimensions

        # Aba Resumo
        ws2 = wb.create_sheet("Resumo")
        total   = len(self.resultados)
        ativos  = sum(1 for r in self.resultados if r["raw"] == "ATIVA")
        inapts  = sum(1 for r in self.resultados if r["raw"] == "INAPTA")
        baixads = sum(1 for r in self.resultados if r["raw"] == "BAIXADA")
        outros  = total - ativos - inapts - baixads

        ws2.append(["Resumo da Verificação"])
        ws2["A1"].font = Font(bold=True, size=13, name="Segoe UI")
        ws2.append([])
        ws2.append(["Total verificado",   total])
        ws2.append(["Ativos",             ativos])
        ws2.append(["Inaptos",            inapts])
        ws2.append(["Baixados",           baixads])
        ws2.append(["Outros / Erros",     outros])
        ws2.column_dimensions["A"].width = 22
        ws2.column_dimensions["B"].width = 12

        wb.save(path)
        messagebox.showinfo("Exportado!", f"Arquivo salvo em:\n{path}")

    # ── Limpar ────────────────────────────────────────────────────────────────

    def _limpar_tudo(self):
        self.txt_input.delete("1.0", "end")
        for row in self.tree.get_children():
            self.tree.delete(row)
        self.resultados.clear()
        self.progress_var.set(0)
        self.progress_lbl.config(text="")


if __name__ == "__main__":
    app = App()
    app.mainloop()
