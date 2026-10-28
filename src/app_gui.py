"""Interface do Gerador de PGR — Marine Seguros."""
import json
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont

from app_version import APP_NOME, APP_VERSAO
import pgr_core
from windows_visual import preparar_windows, aplicar_icone
from PIL import Image, ImageTk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND = True
except Exception:  # app continua funcionando sem arrastar e soltar
    _DND = False

# ---------------------------------------------------------------- paleta Marine
AZUL = "#073B6F"
AZUL_HOVER = "#052F59"
AZUL_CLARO = "#E8F0F8"
AQUA = "#2BB3C0"
FUNDO = "#EEF3F7"
CARTAO = "#FFFFFF"
BORDA = "#D8E2EA"
TEXTO = "#082E55"
TEXTO_2 = "#52697D"
VERDE = "#194FA8"
AMBAR = "#7A4F00"
VERMELHO = "#B01E29"
C = {"bg": FUNDO, "border": BORDA, "surface": CARTAO, "surface_alt": "#F3F7FA", "surface_hi": "#E4EDF3", "border_hi": "#B7C7D3", "border_forte": "#91A6B6", "text": TEXTO}
F_UI = "Segoe UI"

FONTE = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"


def recurso(*partes):
    base = getattr(sys, "_MEIPASS", os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
    return os.path.join(base, "resources", *partes)


def pasta_config():
    base = os.environ.get("APPDATA") or os.path.expanduser("~/.config")
    p = os.path.join(base, "PGR Marine")
    os.makedirs(p, exist_ok=True)
    return p


def carregar_config():
    try:
        with open(os.path.join(pasta_config(), "config.json"), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def salvar_config(cfg):
    try:
        with open(os.path.join(pasta_config(), "config.json"), "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def abrir(caminho):
    if sys.platform == "win32":
        os.startfile(caminho)  # noqa
    elif sys.platform == "darwin":
        subprocess.Popen(["open", caminho])
    else:
        subprocess.Popen(["xdg-open", caminho])


def nome_saida(pasta, segurado):
    nome = (segurado or "SEGURADO").upper()
    nome = re.sub(r"\b(LTDA|S/?A|EIRELI|EPP|ME)\b\.?", "", nome)
    nome = re.sub(r'[\\/:*?"<>|]', "", nome).strip()
    primeiro = nome.split()[0] if nome.split() else "SEGURADO"
    base = f"PGR {primeiro} - {date.today():%m_%Y}"
    destino = os.path.join(pasta, base + ".docx")
    n = 1
    while os.path.exists(destino):  # nunca sobrescreve
        sufixo = " (Cópia)" if n == 1 else f" (Cópia {n})"
        destino = os.path.join(pasta, base + sufixo + ".docx")
        n += 1
    return destino


class CartaoSuave(tk.Frame):
    """Container com borda realmente arredondada no Tk clássico.

    A borda retangular nativa do Frame não é usada. Em vez disso, as quatro
    quinas e os quatro trechos retos da borda são desenhados por overlays,
    evitando os pequenos "restígios" quadrados que apareciam nos cantos.
    """
    def __init__(self, master, raio=8, cor_borda=None, **kwargs):
        self._cor_externa = master.cget("bg") if hasattr(master, "cget") else C["bg"]
        self._cor_borda = cor_borda or C["border"]
        # Nunca deixe o highlight retangular nativo participar do desenho.
        kwargs.pop("highlightbackground", None)
        kwargs.pop("highlightcolor", None)
        kwargs.pop("highlightthickness", None)
        kwargs.setdefault("bd", 0)
        super().__init__(master, highlightthickness=0, **kwargs)
        self._raio = max(5, int(raio))
        self._mascaras = []
        self._bordas = []
        self.bind("<Configure>", self._reposicionar_decoracao, add="+")
        self.after_idle(self._criar_decoracao)

    def _imagem_canto(self, canto):
        r = self._raio
        img = tk.PhotoImage(width=r, height=r)
        cx = (r - 1) if "r" not in canto else 0
        cy = (r - 1) if "b" not in canto else 0
        fundo = self.cget("bg")
        # Anti-alias simples em 1 px: fora = fundo do pai; arco = borda;
        # dentro = fundo do próprio card.
        for y in range(r):
            for x in range(r):
                d = ((x - cx) ** 2 + (y - cy) ** 2) ** 0.5
                if d > r - 0.35:
                    cor = self._cor_externa
                elif d > r - 1.45:
                    cor = self._cor_borda
                else:
                    cor = fundo
                img.put(cor, (x, y))
        return img

    def _criar_decoracao(self):
        if self._mascaras:
            return
        for canto in ("tl", "tr", "bl", "br"):
            imagem = self._imagem_canto(canto)
            lbl = tk.Label(self, image=imagem, bg=self.cget("bg"), bd=0,
                           highlightthickness=0)
            lbl.image = imagem
            self._mascaras.append((canto, lbl))
        # Borda reta desenhada somente entre os arcos — nunca nos cantos.
        for lado in ("top", "bottom", "left", "right"):
            barra = tk.Frame(self, bg=self._cor_borda, bd=0, highlightthickness=0)
            self._bordas.append((lado, barra))
        self._reposicionar_decoracao()

    def _reposicionar_decoracao(self, _evento=None):
        if not self._mascaras:
            return
        r = self._raio
        w, h = self.winfo_width(), self.winfo_height()
        pos = {"tl": (0, 0), "tr": (max(0, w-r), 0),
               "bl": (0, max(0, h-r)), "br": (max(0, w-r), max(0, h-r))}
        for canto, lbl in self._mascaras:
            x, y = pos[canto]
            lbl.place(x=x, y=y, width=r, height=r)
            lbl.lift()
        for lado, barra in self._bordas:
            if lado == "top":
                barra.place(x=r, y=0, width=max(0, w-2*r), height=1)
            elif lado == "bottom":
                barra.place(x=r, y=max(0, h-1), width=max(0, w-2*r), height=1)
            elif lado == "left":
                barra.place(x=0, y=r, width=1, height=max(0, h-2*r))
            else:
                barra.place(x=max(0, w-1), y=r, width=1, height=max(0, h-2*r))
            barra.lift()



# ---------------------------------------------------------------- widgets
class Botao(tk.Canvas):
    """Botão arredondado desenhado em canvas."""

    def __init__(self, master, texto, comando, primario=True, largura=150, altura=36, bg=CARTAO):
        super().__init__(master, width=largura, height=altura, bg=bg, highlightthickness=0, cursor="hand2")
        self.cmd, self.primario, self.ativo = comando, primario, True
        self.l, self.a, self.texto = largura, altura, texto
        self._desenhar(False)
        self.bind("<Enter>", lambda e: self._desenhar(True))
        self.bind("<Leave>", lambda e: self._desenhar(False))
        self.bind("<Button-1>", lambda e: self.ativo and self.cmd())

    def _rr(self, x0, y0, x1, y1, r, **kw):
        pts = [x0 + r, y0, x1 - r, y0, x1, y0, x1, y0 + r, x1, y1 - r, x1, y1,
               x1 - r, y1, x0 + r, y1, x0, y1, x0, y1 - r, x0, y0 + r, x0, y0]
        return self.create_polygon(pts, smooth=True, **kw)

    def _desenhar(self, hover):
        self.delete("all")
        if not self.ativo:
            fill, fg, out = "#C9D2E0", "#FFFFFF", "#C9D2E0"
        elif self.primario:
            fill, fg, out = (AZUL_HOVER if hover else AZUL), "#FFFFFF", (AZUL_HOVER if hover else AZUL)
        else:
            fill, fg, out = (AZUL_CLARO if hover else CARTAO), AZUL, BORDA
        self._rr(1, 1, self.l - 1, self.a - 1, 12, fill=fill, outline=out)
        self.create_text(self.l / 2, self.a / 2, text=self.texto, fill=fg, font=(FONTE, 10, "bold"))

    def habilitar(self, sim):
        self.ativo = sim
        self.config(cursor="hand2" if sim else "arrow")
        self._desenhar(False)


# ---------------------------------------------------------------- app
class MenuSuave(tk.Toplevel):
    """Dropdown com cantos arredondados, no lugar do tk.Menu retangular."""

    def __init__(self, app, ancora, itens, fundo=None):
        super().__init__(app)
        self.app = app
        self.ancora = ancora
        # Cor que fica atrás do menu: os cantos arredondados são compostos
        # sobre ela, então precisa ser a do que está embaixo na tela.
        fundo = fundo or C["bg"]
        # Some antes de existir na tela: no Windows, uma janela sem moldura
        # ignora o geometry aplicado antes do mapeamento e aparece no canto.
        self.withdraw()
        self.overrideredirect(True)
        self.configure(bg=fundo)

        fonte = tkfont.Font(family=F_UI, size=10)
        pad_x, alt_item = app.px(18), app.px(32)
        pad_y, alt_traco = app.px(8), app.px(11)
        rotulos = [t for t, _ in itens if t != "-"]
        largura = max(fonte.measure(t) for t in rotulos) + pad_x * 2 + app.px(16)
        altura = sum(alt_traco if t == "-" else alt_item for t, _ in itens) + pad_y * 2

        tela = tk.Canvas(self, width=largura, height=altura, bg=fundo,
                         highlightthickness=0, bd=0)
        tela.pack()
        imagem_fundo = app.imagem_pilula(largura, altura, C["surface"], fundo,
                                         app.px(12), C["border_forte"])
        if imagem_fundo is not None:
            self._fundo = imagem_fundo
            tela.create_image(0, 0, anchor="nw", image=imagem_fundo)
        else:
            tela.create_rectangle(0, 0, largura, altura, fill=C["surface_alt"],
                                  outline=C["border_hi"])

        # O realce nasce antes dos textos, então fica sempre por baixo deles.
        self._realce = app.imagem_pilula(largura - app.px(12), alt_item - app.px(2),
                                         C["surface_hi"], C["surface"], app.px(8))
        self._id_realce = (tela.create_image(app.px(6), -999, anchor="nw", image=self._realce)
                           if self._realce is not None else None)

        y = pad_y
        for texto, acao in itens:
            if texto == "-":
                meio = y + alt_traco / 2
                tela.create_line(pad_x, meio, largura - pad_x, meio, fill=C["border_hi"])
                y += alt_traco
                continue
            etiqueta = f"item{y}"
            tela.create_rectangle(0, y, largura, y + alt_item, fill="", outline="",
                                  tags=etiqueta)
            tela.create_text(pad_x, y + alt_item / 2, text=texto, anchor="w",
                             fill=C["text"], font=fonte, tags=etiqueta)
            tela.tag_bind(etiqueta, "<Enter>", lambda _e, t=y: self._destacar(t))
            tela.tag_bind(etiqueta, "<Leave>", lambda _e: self._destacar(None))
            tela.tag_bind(etiqueta, "<Button-1>", lambda _e, a=acao: self._executar(a))
            y += alt_item

        self.tela = tela
        self.largura, self.altura = largura, altura
        self.bind("<Escape>", lambda _e: self.fechar())
        # Com o grab ativo, cliques em qualquer ponto do app chegam aqui com
        # coordenadas relativas ao menu — negativas ou além do tamanho quando
        # o clique foi fora dele.
        tela.bind("<Button-1>", self._clique, add="+")
        self.bind("<Button-1>", self._clique, add="+")
        self.bind("<Button-2>", self._clique, add="+")
        self.bind("<Button-3>", self._clique, add="+")

        self.update_idletasks()
        self._posicionar(ancora)
        try:
            self.attributes("-topmost", True)
        except tk.TclError:
            pass
        self.deiconify()
        self.lift()
        # Alguns gerenciadores reposicionam ao mapear: reafirma o lugar.
        self.after_idle(lambda: self._posicionar(ancora))
        try:
            self.grab_set()
            self.focus_set()
        except tk.TclError:
            pass
        self.after(250, self._vigiar_foco)

    def _posicionar(self, ancora):
        """Ancora o menu no botão e o mantém dentro da janela: os cantos
        arredondados são compostos sobre o fundo do app, então vazar para fora
        deixaria quinas visíveis."""
        app = self.app
        folga = app.px(6)
        largura = max(self.largura, self.winfo_reqwidth())
        altura = max(self.altura, self.winfo_reqheight())

        try:
            bx, by = ancora.winfo_rootx(), ancora.winfo_rooty()
            bh = ancora.winfo_height()
            ax, ay = app.winfo_rootx(), app.winfo_rooty()
            aw, ah = app.winfo_width(), app.winfo_height()
            valido = ancora.winfo_ismapped() and aw > 1 and ah > 1
        except tk.TclError:
            valido = False

        if not valido:                       # último recurso: junto do cursor
            bx, by, bh = app.winfo_pointerx(), app.winfo_pointery(), 0
            ax, ay = bx - largura // 2, by
            aw, ah = largura + folga * 2, altura + folga * 2

        abaixo = by + bh + folga
        acima = by - altura - folga
        if abaixo + altura <= ay + ah - folga:
            y = abaixo
        elif acima >= ay + folga:
            y = acima                        # abre para cima (caso do rodapé)
        else:
            y = max(ay + folga, ay + ah - altura - folga)

        # Alinha pela esquerda do botão; se não couber, alinha pela direita
        # dele — fica claro de qual botão o menu saiu.
        x = bx
        if x + largura > ax + aw - folga:
            x = bx + (ancora.winfo_width() if valido else 0) - largura
        x = max(ax + folga, min(x, ax + aw - largura - folga))

        # Nunca deixa o menu sair da tela, mesmo com a janela na borda.
        x = max(0, min(int(x), self.winfo_screenwidth() - largura))
        y = max(0, min(int(y), self.winfo_screenheight() - altura))
        self.geometry(f"{largura}x{altura}+{x}+{y}")

    def _vigiar_foco(self):
        """Fecha ao perder o foco — ligado só depois do menu já estar na tela."""
        if self.winfo_exists():
            self.bind("<FocusOut>", lambda _e: self.fechar())

    def _destacar(self, y):
        if self._id_realce is None:
            return
        self.tela.coords(self._id_realce, self.app.px(6),
                         -999 if y is None else y + self.app.px(1))

    def _clique(self, evento):
        if not (0 <= evento.x < self.largura and 0 <= evento.y < self.altura):
            self.fechar()

    def _executar(self, acao):
        self.fechar()
        if acao:
            self.app.after(10, acao)

    def fechar(self):
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()



class App:
    def __init__(self):
        preparar_windows()
        self.root = TkinterDnD.Tk() if _DND else tk.Tk()
        self.root.title(f"{APP_NOME} · v{APP_VERSAO}")
        altura = min(864, max(650, self.root.winfo_screenheight() - 100))
        self.root.geometry(f"1080x{altura}")
        self.root.minsize(980, 650)
        self.root.configure(bg=FUNDO)
        try:
            if sys.platform == "win32":
                aplicar_icone(self.root, recurso("app.ico"))
            else:
                self.root.iconphoto(True, tk.PhotoImage(file=recurso("app.png")))
        except Exception:
            pass

        self.cfg = carregar_config()
        self.itens = {}          # iid -> dict(pdf, status, relatorio, saida)
        self.fila = queue.Queue()
        self.processando = False
        self._perguntados = set()
        self._estilos()
        self._montar()
        self.root.after(150, self._ler_fila)

    # ------------------------------------------------------------ layout
    def _estilos(self):
        st = ttk.Style(self.root)
        st.theme_use("clam")
        st.configure("Treeview", background=CARTAO, fieldbackground=CARTAO, foreground=TEXTO,
                     rowheight=38, font=(FONTE, 10), borderwidth=0)
        st.configure("Treeview.Heading", background=AZUL_CLARO, foreground=AZUL,
                     font=("Consolas", 8), relief="flat", padding=(8, 6))
        st.map("Treeview", background=[("selected", AZUL_CLARO)], foreground=[("selected", TEXTO)])
        st.map("Treeview.Heading", background=[("active", AZUL_CLARO)])
        st.configure("Vertical.TScrollbar", background=FUNDO, troughcolor=CARTAO, bordercolor=CARTAO,
                     arrowcolor=TEXTO_2)

    def _cartao(self, master, **kw):
        f = CartaoSuave(master, bg=CARTAO, cor_borda=BORDA, raio=9, **kw)
        return f

    def _montar(self):
        r = self.root
        header = tk.Frame(r, bg=FUNDO)
        header.pack(fill="x", padx=18, pady=(16, 12))
        try:
            lado = max(32, round(32 * self.root.winfo_fpixels("1i") / 96))
            with Image.open(recurso("logo_app_master.png")) as original:
                self._logo = ImageTk.PhotoImage(original.convert("RGBA").resize((lado, lado), Image.Resampling.LANCZOS), master=r)
            tk.Label(header, image=self._logo, bg=FUNDO).pack(side="left", padx=(0, 10))
        except tk.TclError:
            pass
        marca = tk.Frame(header, bg=FUNDO)
        marca.pack(side="left")
        linha = tk.Frame(marca, bg=FUNDO)
        linha.pack(anchor="w")
        tk.Label(linha, text="Marine PGR", font=(FONTE, 14, "bold"), bg=FUNDO, fg=TEXTO).pack(side="left")
        tk.Label(linha, text=f"v{APP_VERSAO}", font=("Consolas", 8), bg="#F3F7FA", fg=TEXTO_2, padx=7, pady=3).pack(side="left", padx=8)
        tk.Label(marca, text="Gera e confere o PGR a partir dos documentos da seguradora.", font=(FONTE, 9), bg=FUNDO, fg=TEXTO_2).pack(anchor="w", pady=(2, 0))
        self.b_cfg = Botao(header, "⋯", self._menu_mais, False, 42, bg=FUNDO)
        self.b_cfg.pack(side="right")

        corpo = tk.Frame(r, bg=FUNDO)
        corpo.pack(fill="both", expand=True, padx=18)
        corpo.columnconfigure(0, weight=1)
        corpo.columnconfigure(1, minsize=290)
        corpo.rowconfigure(2, weight=1)
        compat = self._cartao(corpo)
        compat.grid(row=0, column=0, sticky="ew", padx=(0, 14), pady=(0, 8))
        grade = tk.Frame(compat, bg=CARTAO)
        grade.pack(fill="x", padx=14, pady=10)
        for col, peso in enumerate((1, 1, 2)):
            grade.columnconfigure(col, weight=peso, uniform="compat")
        for col, titulo, texto, cor in (
            (0, "SEGURADORA", "Tokio Marine", AZUL),
            (1, "RAMO", "Utilize sempre RC-DC.", AMBAR),
            (2, "PDF ORIGINAL", "Use o original da seguradora.\nPDFs escaneados não são aceitos.", AMBAR),
        ):
            bloco = tk.Frame(grade, bg=CARTAO)
            bloco.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 12, 0))
            tk.Label(bloco, text=titulo, font=("Consolas", 8, "bold"), bg=CARTAO, fg=TEXTO_2).pack(anchor="w")
            rotulo = tk.Label(bloco, text=texto, font=(FONTE, 9), bg=CARTAO, fg=cor, justify="left", anchor="nw", wraplength=230)
            rotulo.pack(fill="x", pady=(4, 0))
            bloco.bind("<Configure>", lambda e, label=rotulo: label.configure(wraplength=max(80, e.width)), add="+")

        self.zona = self._cartao(corpo)
        self.zona.grid(row=1, column=0, sticky="ew", padx=(0, 14), pady=(0, 8))
        entrada = tk.Frame(self.zona, bg=CARTAO)
        entrada.pack(fill="x", padx=12, pady=10)
        textos = tk.Frame(entrada, bg=CARTAO)
        textos.pack(side="left", fill="x", expand=True)
        tk.Label(textos, text="Adicionar documentos", font=(FONTE, 9, "bold"), bg=CARTAO, fg=TEXTO).pack(anchor="w")
        self.drop_texto = tk.StringVar(value="Arraste PDFs para cá ou escolha uma opção" if _DND else "Escolha os PDFs ou uma pasta")
        dica = tk.Label(textos, textvariable=self.drop_texto, font=(FONTE, 9), bg=CARTAO, fg=TEXTO_2)
        dica.pack(anchor="w", pady=(1, 0))
        Botao(entrada, "Arquivos", self.selecionar, True, 92).pack(side="left", padx=(10, 6))
        Botao(entrada, "Pasta", self.selecionar_pasta, False, 82).pack(side="left")

        fila = self._cartao(corpo)
        fila.grid(row=2, column=0, sticky="nsew", padx=(0, 14))
        topo = tk.Frame(fila, bg=CARTAO)
        topo.pack(fill="x", padx=14, pady=(12, 8))
        tk.Label(topo, text="FILA DE PROCESSAMENTO", font=("Consolas", 8, "bold"), fg=TEXTO_2, bg=CARTAO).pack(side="left")
        self.contador = tk.Label(topo, text="0 arquivos", font=("Consolas", 8), bg=CARTAO, fg=TEXTO_2)
        self.contador.pack(side="right")
        quadro = tk.Frame(fila, bg=CARTAO)
        quadro.pack(fill="both", expand=True, padx=8, pady=(0, 10))
        self.tree = ttk.Treeview(quadro, columns=("arquivo", "segurado", "status"), show="headings", selectmode="browse", height=8)
        for c, t, w in (("arquivo", "ARQUIVO", 220), ("segurado", "SEGURADO", 145), ("status", "STATUS", 115)):
            self.tree.heading(c, text=t, anchor="w")
            self.tree.column(c, width=w, minwidth=80, anchor="w", stretch=(c != "status"))
        for tag, cor in (("ok", VERDE), ("revisar", AMBAR), ("erro", VERMELHO), ("pend", TEXTO_2)):
            self.tree.tag_configure(tag, foreground=cor)
        sb = ttk.Scrollbar(quadro, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._mostrar_detalhe())
        self.tree.bind("<Double-1>", lambda e: self.abrir_pgr())
        self.vazio = tk.Label(self.tree, text="Nenhum documento na fila\n\nAdicione os PDFs acima para começar.", bg=CARTAO, fg=TEXTO_2, font=(FONTE, 10), justify="center")
        self.vazio.place(relx=0.5, rely=0.45, anchor="center")
        if _DND:
            for alvo in (self.zona, entrada, textos, dica, fila, self.tree):
                alvo.drop_target_register(DND_FILES)
                alvo.dnd_bind("<<Drop>>", self._soltar)

        lateral = tk.Frame(corpo, bg=FUNDO, width=290)
        lateral.grid(row=0, column=1, rowspan=3, sticky="nsew")
        lateral.columnconfigure(0, weight=1)
        lateral.rowconfigure(1, weight=1)
        destino = self._cartao(lateral)
        destino.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        tk.Label(destino, text="DESTINO DO PGR", font=("Consolas", 8, "bold"), bg=CARTAO, fg=TEXTO_2).pack(anchor="w", padx=14, pady=(12, 8))
        tk.Label(destino, text="Documento Word · modelo timbrado", font=(FONTE, 9, "bold"), bg=CARTAO, fg=TEXTO).pack(anchor="w", padx=14)
        self.destino_texto = tk.StringVar()
        tk.Label(destino, textvariable=self.destino_texto, wraplength=255, justify="left", font=(FONTE, 9), bg=CARTAO, fg=TEXTO_2).pack(anchor="w", padx=14, pady=(4, 12))
        det = self._cartao(lateral)
        det.grid(row=1, column=0, sticky="nsew")
        tk.Label(det, text="CONFERÊNCIA DO DOCUMENTO", font=("Consolas", 8, "bold"), fg=TEXTO_2, bg=CARTAO).pack(anchor="w", padx=14, pady=(12, 8))
        area = tk.Frame(det, bg=CARTAO)
        area.pack(fill="both", expand=True, padx=10, pady=(0, 12))
        self.detalhe = tk.Text(area, wrap="word", width=30, height=10, bg=CARTAO, fg=TEXTO, relief="flat", font=(FONTE, 9), highlightthickness=0, cursor="arrow")
        ds = ttk.Scrollbar(area, orient="vertical", command=self.detalhe.yview)
        self.detalhe.configure(yscrollcommand=ds.set)
        self.detalhe.pack(side="left", fill="both", expand=True)
        ds.pack(side="right", fill="y")
        for tag, cor, peso in (("tit", TEXTO, "bold"), ("ok", VERDE, "bold"), ("rev", AMBAR, "bold"), ("err", VERMELHO, "bold"), ("sub", TEXTO_2, "normal")):
            self.detalhe.tag_configure(tag, foreground=cor, font=(FONTE, 9, peso))
        self._texto_detalhe([("Selecione um documento da fila para ver a conferência.\n\nUse sempre o documento mais recente do cliente. Se houve endosso, gere o PGR a partir dele.", "sub")])
        acoes = tk.Frame(r, bg=FUNDO)
        acoes.pack(fill="x", padx=18, pady=(12, 14))
        self.b_gerar = Botao(acoes, "Gerar PGR", self.gerar, True, 140, bg=FUNDO)
        self.b_gerar.pack(side="left", padx=(0, 8))
        self.b_pasta = Botao(acoes, "Abrir pasta", self.abrir_pasta, False, 130, bg=FUNDO)
        self.b_pasta.pack(side="left", padx=(0, 8))
        self.b_abrir = Botao(acoes, "Abrir PGR", self.abrir_pgr, False, 120, bg=FUNDO)
        self.b_abrir.pack(side="left", padx=(0, 14))
        self.status_texto = tk.StringVar(value="●  Pronto para começar")
        tk.Label(acoes, textvariable=self.status_texto, font=(FONTE, 9), bg=FUNDO, fg=TEXTO_2).pack(side="left")
        self.b_limpar = Botao(acoes, "Limpar fila", self.limpar, False, 120, bg=FUNDO)
        self.b_limpar.pack(side="right")
        self._atualizar_rodape()
        r.bind("<Control-o>", lambda e: self.selecionar())
        r.bind("<Control-Return>", lambda e: None if self.processando else self.gerar())

    def _centralizar(self, janela):
        self.root.update_idletasks()
        janela.update_idletasks()
        largura, altura = janela.winfo_width(), janela.winfo_height()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - largura) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - altura) // 2
        x = max(0, min(x, janela.winfo_screenwidth() - largura))
        y = max(0, min(y, janela.winfo_screenheight() - altura))
        janela.geometry(f"+{x}+{y}")

    def _menu_mais(self):
        self.root.px = lambda valor: int(valor)
        self.root.imagem_pilula = self._imagem_menu
        MenuSuave(self.root, self.b_cfg, [
            ("Selecionar PDFs…", self.selecionar),
            ("Selecionar pasta…", self.selecionar_pasta),
            ("Limpar fila", self.limpar),
            ("-", None),
            ("Configurações", self.configurar),
            ("Como usar", self._como_usar),
            ("Sobre o Marine PGR", self._sobre),
            ("-", None),
            ("Sair", self.root.destroy),
        ])

    def _imagem_menu(self, largura, altura, cor, fundo, raio, borda=None):
        import math
        imagem = tk.PhotoImage(master=self.root, width=largura, height=altura)
        imagem.put(fundo, to=(0, 0, largura, altura))
        for y in range(altura):
            dy = max(raio - y - 0.5, y - (altura - raio) + 0.5, 0)
            margem = max(0, round(raio - math.sqrt(max(0, raio * raio - dy * dy))))
            imagem.put(borda or cor, to=(margem, y, largura - margem, y + 1))
            if borda and 0 < y < altura - 1:
                imagem.put(cor, to=(margem + 1, y, largura - margem - 1, y + 1))
        return imagem

    def _como_usar(self):
        messagebox.showinfo("Como usar", "1. Adicione os PDFs originais da seguradora, sempre do ramo RC-DC.\n2. Use o documento mais recente: se houve endosso, utilize o endosso.\n3. Clique em Gerar PGR.\n4. Confira o resultado no painel lateral antes de enviar.\n\nPDFs escaneados não são aceitos. Em Configurações você pode escolher o modelo Word e a pasta de saída.", parent=self.root)

    def _sobre(self):
        messagebox.showinfo("Sobre o Marine PGR", f"Marine PGR · v{APP_VERSAO}\nGerador de Programa de Gerenciamento de Risco\n\nSeguradora: Tokio Marine · ramo RC-DC.\nCópia literal do PDF original para o modelo timbrado, com conferência do texto.", parent=self.root)

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta dos PDFs")
        if pasta:
            self._adicionar([os.path.join(pasta, n) for n in sorted(os.listdir(pasta)) if n.lower().endswith(".pdf")])

    def _resumo_fila(self):
        total = len(self.itens)
        self.contador.config(text=f"{total} arquivo(s)")
        if total:
            self.vazio.place_forget()
        else:
            self.vazio.place(relx=0.5, rely=0.45, anchor="center")
        self.drop_texto.set(f"{total} documento(s) na fila" if total else "Adicione PDFs para começar")
        self.status_texto.set("●  Gerando documentos…" if self.processando else f"●  {total} documento(s) na fila" if total else "●  Pronto para começar")

    # ------------------------------------------------------------ helpers
    def _modelo(self):
        m = self.cfg.get("modelo")
        return m if m and os.path.exists(m) else recurso("modelo_pgr.docx")

    def _atualizar_rodape(self):
        m = self.cfg.get("modelo")
        mod = os.path.basename(m) if m and os.path.exists(m) else "modelo timbrado padrão (embutido)"
        dest = self.cfg.get("pasta_saida") or "mesma pasta do PDF"
        self.destino_texto.set(dest)


    def _texto_detalhe(self, partes):
        self.detalhe.config(state="normal")
        self.detalhe.delete("1.0", "end")
        for t, tag in partes:
            self.detalhe.insert("end", t, tag)
        self.detalhe.config(state="disabled")

    def _adicionar(self, caminhos):
        novos = 0
        existentes = {v["pdf"] for v in self.itens.values()}
        for c in caminhos:
            if c.lower().endswith(".pdf") and os.path.isfile(c) and c not in existentes:
                iid = self.tree.insert("", "end", values=(os.path.basename(c), "—", "Na fila"), tags=("pend",))
                self.itens[iid] = {"pdf": c, "status": "pend", "rel": None}
                existentes.add(c)
                novos += 1
        self._resumo_fila()
        if novos and not self.tree.selection():
            self.tree.selection_set(self.tree.get_children()[0])

    def _soltar(self, ev):
        self._adicionar(self.root.tk.splitlist(ev.data))

    def _selecionado(self):
        s = self.tree.selection()
        return (s[0], self.itens[s[0]]) if s else (None, None)

    # ------------------------------------------------------------ ações
    def selecionar(self):
        arqs = filedialog.askopenfilenames(title="Selecione apólice / proposta / endosso",
                                           filetypes=[("PDF", "*.pdf")])
        self._adicionar(arqs)

    def limpar(self):
        if self.processando:
            return
        self.tree.delete(*self.tree.get_children())
        self.itens.clear()
        self._resumo_fila()
        self._texto_detalhe([("Fila vazia.", "sub")])

    def configurar(self):
        if self.processando:
            return
        win = tk.Toplevel(self.root, bg=CARTAO)
        win.title("Configurações")
        win.geometry("560x300")
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)
        try:
            aplicar_icone(win, recurso("app.ico"))
        except tk.TclError:
            pass
        tk.Label(win, text="Modelo timbrado (.docx)", font=(FONTE, 10, "bold"), bg=CARTAO, fg=TEXTO).pack(anchor="w", padx=18, pady=(16, 2))
        v_mod = tk.StringVar(value=self.cfg.get("modelo") or "")
        f1 = tk.Frame(win, bg=CARTAO); f1.pack(fill="x", padx=18)
        tk.Entry(f1, textvariable=v_mod, font=(FONTE, 9)).pack(side="left", fill="x", expand=True)
        tk.Button(f1, text="Escolher…", command=lambda: v_mod.set(
            filedialog.askopenfilename(filetypes=[("Word", "*.docx")]) or v_mod.get())).pack(side="left", padx=6)
        tk.Label(win, text="Vazio = usa o modelo padrão que vem dentro do app.", font=(FONTE, 8), bg=CARTAO, fg=TEXTO_2).pack(anchor="w", padx=18)

        tk.Label(win, text="Pasta onde salvar os PGRs", font=(FONTE, 10, "bold"), bg=CARTAO, fg=TEXTO).pack(anchor="w", padx=18, pady=(14, 2))
        v_pas = tk.StringVar(value=self.cfg.get("pasta_saida") or "")
        f2 = tk.Frame(win, bg=CARTAO); f2.pack(fill="x", padx=18)
        tk.Entry(f2, textvariable=v_pas, font=(FONTE, 9)).pack(side="left", fill="x", expand=True)
        tk.Button(f2, text="Escolher…", command=lambda: v_pas.set(
            filedialog.askdirectory() or v_pas.get())).pack(side="left", padx=6)
        tk.Label(win, text="Vazio = salva na mesma pasta do PDF.", font=(FONTE, 8), bg=CARTAO, fg=TEXTO_2).pack(anchor="w", padx=18)

        n_apr = sum(len(v) for v in (self.cfg.get("mapa_secoes") or {}).values())
        f3 = tk.Frame(win, bg=CARTAO); f3.pack(fill="x", padx=18, pady=(14, 0))
        l_apr = tk.Label(f3, text=f"Títulos de seção aprendidos: {n_apr}", font=(FONTE, 9), bg=CARTAO, fg=TEXTO)
        l_apr.pack(side="left")

        def esquecer():
            self.cfg["mapa_secoes"] = {}
            salvar_config(self.cfg)
            l_apr.config(text="Títulos de seção aprendidos: 0")
        tk.Button(f3, text="Esquecer", command=esquecer).pack(side="left", padx=8)

        def ok():
            self.cfg["modelo"] = v_mod.get().strip() or None
            self.cfg["pasta_saida"] = v_pas.get().strip() or None
            salvar_config(self.cfg)
            self._atualizar_rodape()
            win.destroy()
        Botao(win, "Salvar", ok, True, 110).pack(anchor="e", padx=18, pady=16)
        self._centralizar(win)

    def abrir_pgr(self):
        _, it = self._selecionado()
        if it and it.get("rel") and it["rel"].get("saida") and os.path.exists(it["rel"]["saida"]):
            abrir(it["rel"]["saida"])

    def abrir_pasta(self):
        _, it = self._selecionado()
        if it:
            alvo = it["rel"]["saida"] if it.get("rel") and it["rel"].get("saida") else it["pdf"]
            abrir(os.path.dirname(alvo))

    def gerar(self):
        if self.processando:
            return
        pend = [i for i, v in self.itens.items() if v["status"] in ("pend", "erro", "revisar")]
        if not pend:
            if not self.itens:
                self.selecionar()
            return
        self.processando = True
        self._resumo_fila()
        for b in (self.b_gerar, self.b_limpar, self.b_cfg):
            b.habilitar(False)
        self._perguntados = set()
        self._iniciar(pend, {})

    def _iniciar(self, iids, mapas_item):
        self.processando = True
        self._resumo_fila()
        for b in (self.b_gerar, self.b_limpar, self.b_cfg):
            b.habilitar(False)
        aprendido = dict(self.cfg.get("mapa_secoes") or {})
        mapas = {}
        for iid in iids:
            m = {k: list(v) for k, v in aprendido.items()}
            for k, v in (mapas_item.get(iid) or {}).items():
                m[k] = v if v == "__PULAR__" else [v] + m.get(k, [])
            mapas[iid] = m
        modelo, pasta_cfg = self._modelo(), self.cfg.get("pasta_saida")
        threading.Thread(target=self._trabalho, args=(iids, modelo, pasta_cfg, mapas), daemon=True).start()

    def _trabalho(self, iids, modelo, pasta_cfg, mapas):
        for iid in iids:
            pdf = self.itens[iid]["pdf"]
            anterior = (self.itens[iid].get("rel") or {}).get("saida")
            self.fila.put(("andamento", iid))
            try:
                pasta = pasta_cfg if pasta_cfg and os.path.isdir(pasta_cfg) else os.path.dirname(pdf)
                tmp = os.path.join(pasta, f"~pgr_{os.getpid()}_{iid}.docx")
                rel = pgr_core.gerar(pdf, modelo, tmp, mapa=mapas.get(iid))
                if rel.get("erro"):
                    if os.path.exists(tmp):
                        os.remove(tmp)
                    rel["saida"] = None
                else:
                    # ao gerar de novo o mesmo item, substitui o arquivo que este app acabou de criar
                    final = anterior if anterior and os.path.dirname(anterior) == pasta else \
                        nome_saida(pasta, rel["cabecalho"].get("SEGURADO"))
                    os.replace(tmp, final)
                    rel["saida"] = final
                self.fila.put(("fim", iid, rel))
            except Exception as e:
                self.fila.put(("fim", iid, {"ok": False, "erro": f"Falha ao processar: {e}", "saida": None,
                                            "cabecalho": {}, "problemas": []}))
        self.fila.put(("lote",))

    def _ler_fila(self):
        try:
            while True:
                msg = self.fila.get_nowait()
                if msg[0] == "andamento":
                    iid = msg[1]
                    v = self.tree.item(iid, "values")
                    self.tree.item(iid, values=(v[0], v[1], "Gerando…"), tags=("pend",))
                elif msg[0] == "fim":
                    iid, rel = msg[1], msg[2]
                    it = self.itens[iid]
                    it["rel"] = rel
                    if rel.get("erro"):
                        it["status"], st, tag = "erro", "✖ Erro", "erro"
                    elif rel["ok"]:
                        it["status"], st, tag = "ok", "✔ Conferido", "ok"
                    else:
                        it["status"], st, tag = "revisar", "⚠ Revisar", "revisar"
                    seg = (rel.get("cabecalho") or {}).get("SEGURADO") or "—"
                    self.tree.item(iid, values=(os.path.basename(it["pdf"]), seg, st), tags=(tag,))
                    if self.tree.selection() and self.tree.selection()[0] == iid:
                        self._mostrar_detalhe()
                elif msg[0] == "lote":
                    self._fim_lote()
        except queue.Empty:
            pass
        self.root.after(150, self._ler_fila)

    def _fim_lote(self):
        self.processando = False
        self._resumo_fila()
        for b in (self.b_gerar, self.b_limpar, self.b_cfg):
            b.habilitar(True)
        # seções não encontradas: pergunta o título exato, um arquivo por vez
        ajustes = {}
        for iid, it in self.itens.items():
            rel = it.get("rel") or {}
            if rel.get("faltando") and not rel.get("erro") and iid not in self._perguntados:
                self._perguntados.add(iid)
                self.tree.selection_set(iid)
                self._mostrar_detalhe()
                m = self._perguntar_secoes(it["pdf"], rel)
                if m:
                    ajustes[iid] = m
        if ajustes:
            return self._iniciar(list(ajustes), ajustes)
        vals = [v["status"] for v in self.itens.values()]
        ok, rev, err = vals.count("ok"), vals.count("revisar"), vals.count("erro")
        msg = f"{ok} PGR conferido(s)"
        if rev:
            msg += f"\n{rev} para revisar"
        if err:
            msg += f"\n{err} com erro"
        msg += "\n\nSelecione um item na fila para ver a conferência."
        (messagebox.showinfo if not (rev or err) else messagebox.showwarning)("Lote concluído", msg)
        self._mostrar_detalhe()

    def _perguntar_secoes(self, pdf, rel):
        """Janela para o usuário indicar o título exato de cada seção não encontrada."""
        PULAR = "— Não incluir esta seção —"
        titulos = rel.get("secoes_pdf") or []
        opcoes = [t for t in rel.get("fora", [])] + [PULAR]
        win = tk.Toplevel(self.root, bg=CARTAO)
        win.title("Seção não encontrada")
        win.transient(self.root)
        win.resizable(False, False)
        tk.Label(win, text="Seção não encontrada", font=(FONTE, 12, "bold"), fg=AMBAR, bg=CARTAO).pack(anchor="w", padx=18, pady=(16, 2))
        tk.Label(win, text=os.path.basename(pdf), font=(FONTE, 9, "bold"), fg=TEXTO, bg=CARTAO).pack(anchor="w", padx=18)
        tk.Label(win, text="Nesta proposta o título está escrito de outro jeito. Escolha na lista o título\n"
                           "como aparece na faixa cinza do PDF, ou digite exatamente igual.",
                 font=(FONTE, 9), fg=TEXTO_2, bg=CARTAO, justify="left").pack(anchor="w", padx=18, pady=(6, 10))
        campos = {}
        for nome in rel["faltando"]:
            f = tk.Frame(win, bg=CARTAO)
            f.pack(fill="x", padx=18, pady=4)
            tk.Label(f, text=nome, font=(FONTE, 9, "bold"), fg=TEXTO, bg=CARTAO, width=34, anchor="w").pack(side="left")
            v = tk.StringVar()
            # sugere o título que contém o nome procurado (ex.: "LIMITES PARA O PLANO DE GERENCIAMENTO DE RISCO")
            sug = next((t for t in rel.get("fora", []) if pgr_core.norm(nome) in pgr_core.norm(t)), "")
            v.set(sug)
            cb = ttk.Combobox(f, textvariable=v, values=opcoes, width=52, font=(FONTE, 9))
            cb.pack(side="left", padx=(6, 0))
            campos[nome] = v
        v_lembrar = tk.BooleanVar(value=True)
        tk.Checkbutton(win, text="Lembrar para as próximas propostas", variable=v_lembrar, bg=CARTAO,
                       font=(FONTE, 9), fg=TEXTO, activebackground=CARTAO).pack(anchor="w", padx=14, pady=(10, 0))
        aviso = tk.Label(win, text="", font=(FONTE, 9), fg=VERMELHO, bg=CARTAO, justify="left", wraplength=560)
        aviso.pack(anchor="w", padx=18)
        resultado = {}

        def aplicar():
            mapa, erros = {}, []
            for nome, v in campos.items():
                t = v.get().strip()
                if not t:
                    continue
                if t == PULAR:
                    mapa[nome] = "__PULAR__"
                    continue
                real = next((x for x in titulos if pgr_core.norm(x) == pgr_core.norm(t)), None)
                if not real:
                    erros.append(f"“{t}” não existe neste PDF. Confira a grafia ou escolha na lista.")
                else:
                    mapa[nome] = real
            if erros:
                aviso.config(text="\n".join(erros))
                return
            if v_lembrar.get():
                aprendido = self.cfg.setdefault("mapa_secoes", {})
                for nome, t in mapa.items():
                    if t != "__PULAR__" and t not in aprendido.setdefault(nome, []):
                        aprendido[nome].append(t)
                salvar_config(self.cfg)
            resultado.update(mapa)
            win.destroy()

        bts = tk.Frame(win, bg=CARTAO)
        bts.pack(fill="x", padx=18, pady=16)
        Botao(bts, "Aplicar e gerar de novo", aplicar, True, 200).pack(side="right")
        Botao(bts, "Manter sem a seção", win.destroy, False, 170).pack(side="right", padx=8)
        win.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - win.winfo_width()) // 2
        y = self.root.winfo_rooty() + 120
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        win.grab_set()
        self.root.wait_window(win)
        return resultado

    def _mostrar_detalhe(self):
        _, it = self._selecionado()
        if not it:
            return
        rel = it.get("rel")
        p = [(os.path.basename(it["pdf"]) + "\n\n", "tit")]
        if not rel:
            p.append(("Ainda não gerado. Clique em Gerar PGR.", "sub"))
            return self._texto_detalhe(p)
        if rel.get("erro"):
            p.append(("✖ " + rel["erro"] + "\n", "err"))
            return self._texto_detalhe(p)
        p.append(("✔ Pronto para enviar\n\n" if rel["ok"] else "⚠ Revise antes de enviar\n\n", "ok" if rel["ok"] else "rev"))
        p.append(("Cabeçalho\n", "tit"))
        for k, v in rel["cabecalho"].items():
            p.append((f"{k}: ", "sub"))
            p.append(((v or "NÃO ENCONTRADO — preencher no Word") + "\n", "" if v else "err"))
        p.append(("\nConferência literal\n", "tit"))
        n, probs = rel["palavras"], rel["problemas"]
        p.append((f"{n} palavras comparadas · {len(probs)} divergência(s)\n", "ok" if not probs else "err"))
        for x in probs[:15]:
            p.append(("• " + x + "\n", "sub"))
        p.append(("\nSeções incluídas\n", "tit"))
        p.append(("\n".join("• " + s for s in rel["incluidas"]) + "\n", ""))
        if rel["faltando"]:
            p.append(("\nNão encontradas no PDF\n", "rev"))
            p.append(("\n".join("• " + s for s in rel["faltando"]) + "\n", ""))
        if rel.get("puladas"):
            p.append(("\nDeixadas de fora pelo usuário\n", "rev"))
            p.append(("\n".join("• " + s for s in rel["puladas"]) + "\n", ""))
        if rel.get("saida"):
            p.append(("\nArquivo\n", "tit"))
            p.append((rel["saida"] + "\n", "sub"))
        self._texto_detalhe(p)

    def executar(self):
        self.root.mainloop()
