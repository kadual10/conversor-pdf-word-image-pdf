"""Conversor de arquivos: PDF para Word e imagem para PDF, com interface gráfica."""

import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import font as tkfont

from pdf2docx import Converter
from PIL import Image, ImageOps

# Deixa a janela nítida em telas com escala acima de 100% no Windows.
# Precisa rodar antes de criar a janela.
try:
    import ctypes

    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# ---------------------------------------------------------------- Paleta
FUNDO = "#F3F4F6"
FOLHA = "#FFFFFF"
TINTA = "#1C2635"
TINTA_HOVER = "#2D3B52"
TEXTO_SUAVE = "#667085"
LINHA = "#D5D9E0"
LINHA_FOLHA = "#E3E6EB"
DOBRA = "#E9ECF0"
TRILHO = "#E4E7EC"
PDF = "#D2433D"
WORD = "#2A5DB0"
IMAGEM = "#2E8B6A"
ERRO = "#B3261E"
INATIVO = "#8F98A6"
HOVER_CLARO = "#E9ECF1"

# ------------------------------------------------- Ajustes da imagem para PDF
DPI_PDF = 150                      # qualidade das páginas geradas
PAGINA_A4 = (round(8.27 * DPI_PDF), round(11.69 * DPI_PDF))  # largura x altura em pixels
MARGEM_PDF = round(0.3 * DPI_PDF)  # margem de cerca de 0,8 cm

# ------------------------------------------------------------------- Modos
MODOS = {
    "word": {
        "rotulo": "PDF para Word",
        "subtitulo": "Escolha o PDF e a pasta onde o Word será salvo.",
        "origem": ("PDF", PDF, "docs"),
        "destino": ("DOCX", WORD, "docs"),
        "ext": ".docx",
        "acao": "Converter para Word",
        "dialogo": "Escolher PDF",
        "tipos": [("Arquivos PDF", "*.pdf")],
        "multiplos": False,
        "escolher": "Escolher PDF",
        "trocar": "Trocar PDF",
        "vazio": "PDF não escolhido",
        "plural": "PDFs",
        "msg_inicio": "Comece escolhendo o PDF que você quer converter.",
        "msg_falta": "Agora escolha o PDF.",
        "msg_auto": "O Word será salvo na mesma pasta do PDF. Para mudar o destino, use Trocar pasta.",
        "msg_pronto": "Tudo pronto. Clique em Converter para Word.",
        "msg_andamento": "Convertendo. PDFs grandes podem levar alguns minutos.",
        "msg_erro": ("Não foi possível converter. Confira se o PDF não tem senha e se não é "
                     "um arquivo escaneado (só imagem)."),
        "ordem": "",
    },
    "pdf": {
        "rotulo": "Imagem para PDF",
        "subtitulo": "Escolha uma ou mais imagens e a pasta onde o PDF será salvo.",
        "origem": ("IMG", IMAGEM, "imagem"),
        "destino": ("PDF", PDF, "docs"),
        "ext": ".pdf",
        "acao": "Criar PDF",
        "dialogo": "Escolher imagens",
        "tipos": [("Imagens", "*.png *.jpg *.jpeg *.bmp *.gif *.webp *.tif *.tiff")],
        "multiplos": True,
        "escolher": "Escolher imagens",
        "trocar": "Trocar imagens",
        "vazio": "Nenhuma imagem escolhida",
        "plural": "imagens",
        "msg_inicio": "Comece escolhendo as imagens que você quer juntar em um PDF.",
        "msg_falta": "Agora escolha as imagens.",
        "msg_auto": "O PDF será salvo na mesma pasta das imagens. Para mudar o destino, use Trocar pasta.",
        "msg_pronto": "Tudo pronto. Clique em Criar PDF.",
        "msg_andamento": "Criando o PDF. Muitas imagens grandes podem levar alguns segundos.",
        "msg_erro": ("Não foi possível criar o PDF. Confira se os arquivos são imagens válidas "
                     "e se o PDF de destino não está aberto em outro programa."),
        "ordem": "As imagens entram em ordem de nome, uma por página.",
    },
}


# ------------------------------------------------------------- Utilitários
def misturar(cor_a, cor_b, t):
    """Cor intermediária entre duas cores hexadecimais (t de 0 a 1)."""
    a = [int(cor_a[i : i + 2], 16) for i in (1, 3, 5)]
    b = [int(cor_b[i : i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def encurtar(texto, fonte, largura):
    """Corta o nome do arquivo com reticências, preservando a extensão."""
    if fonte.measure(texto) <= largura:
        return texto
    base, ext = os.path.splitext(texto)
    while len(base) > 1:
        base = base[:-1]
        candidato = base + "…" + ext
        if fonte.measure(candidato) <= largura:
            return candidato
    return texto[:1] + "…"


def formatar_tamanho(total):
    if total < 1024:
        return f"{total} B"
    total /= 1024
    for unidade in ("KB", "MB", "GB"):
        if total < 1024 or unidade == "GB":
            return f"{total:.1f} {unidade}".replace(".", ",")
        total /= 1024


def tamanho_total(caminhos):
    total = 0
    for caminho in caminhos:
        try:
            total += os.path.getsize(caminho)
        except OSError:
            pass
    return total


def chave_natural(caminho):
    """Ordena foto2.jpg antes de foto10.jpg."""
    nome = os.path.basename(caminho).lower()
    return [int(parte) if parte.isdigit() else parte for parte in re.split(r"(\d+)", nome)]


def pontos_retangulo(x1, y1, x2, y2, raio):
    return [
        x1 + raio, y1, x2 - raio, y1, x2, y1, x2, y1 + raio,
        x2, y2 - raio, x2, y2, x2 - raio, y2, x1 + raio, y2,
        x1, y2, x1, y2 - raio, x1, y1 + raio, x1, y1,
    ]


def retangulo_arredondado(canvas, x1, y1, x2, y2, raio, **opcoes):
    return canvas.create_polygon(pontos_retangulo(x1, y1, x2, y2, raio), smooth=True, **opcoes)


# ------------------------------------------------------ Imagem para PDF
def imagens_para_pdf(caminhos, destino):
    """Junta as imagens em um PDF, uma por página A4 (retrato ou paisagem).

    Imagens pequenas não são ampliadas; ficam centralizadas na página.
    """
    paginas = []
    for caminho in caminhos:
        with Image.open(caminho) as original:
            imagem = ImageOps.exif_transpose(original)  # respeita a rotação das fotos de celular
            if imagem.mode in ("RGBA", "LA") or (imagem.mode == "P" and "transparency" in imagem.info):
                imagem = imagem.convert("RGBA")
                fundo = Image.new("RGB", imagem.size, "white")
                fundo.paste(imagem, mask=imagem.split()[-1])
                imagem = fundo
            else:
                imagem = imagem.convert("RGB")

        largura_pag, altura_pag = PAGINA_A4 if imagem.height >= imagem.width else PAGINA_A4[::-1]
        escala = min(
            (largura_pag - 2 * MARGEM_PDF) / imagem.width,
            (altura_pag - 2 * MARGEM_PDF) / imagem.height,
            1,
        )
        novo_tamanho = (max(1, round(imagem.width * escala)), max(1, round(imagem.height * escala)))
        if novo_tamanho != imagem.size:
            imagem = imagem.resize(novo_tamanho, Image.Resampling.LANCZOS)

        pagina = Image.new("RGB", (largura_pag, altura_pag), "white")
        pagina.paste(imagem, ((largura_pag - imagem.width) // 2, (altura_pag - imagem.height) // 2))
        paginas.append(pagina)

    paginas[0].save(destino, "PDF", resolution=DPI_PDF, quality=90,
                    save_all=True, append_images=paginas[1:])


# ------------------------------------------------------------------ Botão
class Botao(tk.Canvas):
    """Botão arredondado com hover, foco por teclado e estado desativado."""

    def __init__(self, master, texto, comando, largura, altura, raio, fonte,
                 fundo, fundo_hover, cor_texto, contorno=None):
        super().__init__(master, width=largura, height=altura, bg=master.cget("bg"),
                         highlightthickness=0, bd=0, takefocus=True)
        self.comando = comando
        self.fundo, self.fundo_hover, self.contorno = fundo, fundo_hover, contorno
        self.ativo, self.hover, self.foco = True, False, False

        self.forma = retangulo_arredondado(
            self, 2, 2, largura - 2, altura - 2, raio,
            fill=fundo, outline=contorno or fundo, width=2,
        )
        self.rotulo = self.create_text(largura / 2, altura / 2, text=texto,
                                       fill=cor_texto, font=fonte)

        self.bind("<Enter>", lambda e: self._estado(hover=True))
        self.bind("<Leave>", lambda e: self._estado(hover=False))
        self.bind("<FocusIn>", lambda e: self._estado(foco=True))
        self.bind("<FocusOut>", lambda e: self._estado(foco=False))
        self.bind("<Button-1>", self._clicar)
        self.bind("<Return>", self._clicar)
        self.bind("<space>", self._clicar)
        self._pintar()

    def _clicar(self, _evento=None):
        if self.ativo:
            self.comando()

    def _estado(self, hover=None, foco=None):
        if hover is not None:
            self.hover = hover
        if foco is not None:
            self.foco = foco
        self._pintar()

    def _pintar(self):
        if not self.ativo:
            fundo = borda = INATIVO
            cursor = "arrow"
        else:
            fundo = self.fundo_hover if self.hover else self.fundo
            borda = WORD if self.foco else (self.contorno or fundo)
            cursor = "hand2"
        self.itemconfigure(self.forma, fill=fundo, outline=borda)
        self.configure(cursor=cursor)

    def definir_texto(self, texto):
        self.itemconfigure(self.rotulo, text=texto)

    def definir_ativo(self, ativo):
        self.ativo = ativo
        self._pintar()


# ----------------------------------------------------------- Seletor de modo
class Seletor(tk.Canvas):
    """Controle segmentado para trocar de conversor (também funciona com as setas)."""

    def __init__(self, master, opcoes, atual, comando, largura, altura, raio, fonte):
        super().__init__(master, width=largura, height=altura, bg=master.cget("bg"),
                         highlightthickness=0, bd=0, takefocus=True)
        self.chaves = [chave for chave, _ in opcoes]
        self.comando = comando
        self.atual = atual
        self.ativo, self.foco = True, False
        self.altura, self.raio = altura, raio
        self.segmento = (largura - 8) / len(opcoes)

        self.trilho = retangulo_arredondado(self, 2, 2, largura - 2, altura - 2, raio,
                                            fill=TRILHO, outline=TRILHO, width=2)
        self.pilula = retangulo_arredondado(self, *self._caixa(self.chaves.index(atual)),
                                            raio - 2, fill=FOLHA, outline=LINHA, width=1)
        self.rotulos = {}
        for i, (chave, texto) in enumerate(opcoes):
            self.rotulos[chave] = self.create_text(
                4 + self.segmento * (i + 0.5), altura / 2, text=texto, font=fonte)

        self.bind("<Button-1>", self._clicar)
        self.bind("<Left>", lambda e: self._mover(-1))
        self.bind("<Right>", lambda e: self._mover(1))
        self.bind("<FocusIn>", lambda e: self._foco(True))
        self.bind("<FocusOut>", lambda e: self._foco(False))
        self._pintar()

    def _caixa(self, indice):
        x1 = 4 + self.segmento * indice + 1
        x2 = 4 + self.segmento * (indice + 1) - 1
        return x1, 5, x2, self.altura - 5

    def _foco(self, foco):
        self.foco = foco
        self._pintar()

    def _pintar(self):
        for chave, item in self.rotulos.items():
            if chave == self.atual:
                cor = TINTA
            else:
                cor = TEXTO_SUAVE if self.ativo else INATIVO
            self.itemconfigure(item, fill=cor)
        borda = WORD if (self.foco and self.ativo) else TRILHO
        self.itemconfigure(self.trilho, outline=borda)
        self.configure(cursor="hand2" if self.ativo else "arrow")

    def _clicar(self, evento):
        indice = int((evento.x - 4) // self.segmento)
        indice = min(max(indice, 0), len(self.chaves) - 1)
        self.selecionar(self.chaves[indice])

    def _mover(self, passo):
        indice = self.chaves.index(self.atual) + passo
        if 0 <= indice < len(self.chaves):
            self.selecionar(self.chaves[indice])

    def selecionar(self, chave):
        if not self.ativo or chave == self.atual:
            return
        self.atual = chave
        x1, y1, x2, y2 = self._caixa(self.chaves.index(chave))
        self.coords(self.pilula, *pontos_retangulo(x1, y1, x2, y2, self.raio - 2))
        self._pintar()
        self.comando(chave)

    def definir_ativo(self, ativo):
        self.ativo = ativo
        self._pintar()


# --------------------------------------------------- Ilustração: de A para B
class Fluxo(tk.Canvas):
    """Duas folhas ligadas por uma trilha que vai da cor de origem à de destino.

    As folhas também são clicáveis: a da esquerda escolhe o(s) arquivo(s) e a da
    direita, a pasta de destino.
    """

    def __init__(self, master, k, fontes, origem, destino, ao_clicar_a, ao_clicar_b):
        self.k = k
        s = self.s
        self.largura = s(640)
        self.fw, self.fh = s(158), s(200)  # largura e altura de cada folha
        self.y0 = s(2)
        super().__init__(master, width=self.largura, height=s(268), bg=FUNDO,
                         highlightthickness=0, bd=0)

        self.fontes = fontes
        self.interativo = True
        self.tem_a = self.tem_b = False
        self.sessao = 0
        self._pulso_id = None
        self._pulso_pos = 0.0
        self.linhas_b = []

        self.x_a = s(1)
        self.x_b = self.largura - s(1) - self.fw
        self.y_meio = self.y0 + self.fh // 2
        self.trilha_x0 = self.x_a + self.fw + s(26)
        self.trilha_x1 = self.x_b - s(26)

        self.create_line(self.trilha_x0, self.y_meio, self.trilha_x1, self.y_meio,
                         fill=LINHA, width=s(3), dash=(s(3), s(7)),
                         capstyle="round", tags="trilha_base")

        y_texto = self.y0 + self.fh + s(14)
        for lado, x in (("a", self.x_a), ("b", self.x_b)):
            self.create_text(x, y_texto, anchor="nw", font=fontes["nome"], tags=f"{lado}_l1")
            self.create_text(x, y_texto + s(24), anchor="nw", font=fontes["pequena"],
                             tags=f"{lado}_l2")

        self.tag_bind("folha_a", "<Button-1>", lambda e: self.interativo and ao_clicar_a())
        self.tag_bind("folha_b", "<Button-1>", lambda e: self.interativo and ao_clicar_b())
        for tag in ("folha_a", "folha_b"):
            self.tag_bind(tag, "<Enter>", self._entrar)
            self.tag_bind(tag, "<Leave>", self._sair)

        self.configurar(origem, destino)

    def s(self, valor):
        return int(round(valor * self.k))

    # ---- desenho
    def configurar(self, origem, destino):
        """Troca as duas folhas (selo, cor e tipo de conteúdo) e zera a animação."""
        self.selo_a, self.cor_a, self.tipo_a = origem
        self.selo_b, self.cor_b, self.tipo_b = destino
        self.delete("folha_a", "folha_b")
        self.linhas_b = []
        self._desenhar_folha(self.x_a, "a", *origem)
        self._desenhar_folha(self.x_b, "b", *destino)
        self.reiniciar()

    def _desenhar_folha(self, x, lado, selo, cor, tipo):
        s = self.s
        y, w, h, d = self.y0, self.fw, self.fh, s(30)
        tag = f"folha_{lado}"
        borda = (tag, tag + "_borda")

        self.create_polygon(x, y, x + w - d, y, x + w, y + d, x + w, y + h, x, y + h,
                            fill=FOLHA, outline=LINHA, width=2, joinstyle="miter", tags=borda)
        self.create_polygon(x + w - d, y, x + w - d, y + d, x + w, y + d,
                            fill=DOBRA, outline=LINHA, width=2, joinstyle="miter", tags=borda)

        sw, sh = s(58), s(24)
        retangulo_arredondado(self, x + s(16), y + s(18), x + s(16) + sw, y + s(18) + sh,
                              s(6), fill=cor, outline=cor, tags=(tag,))
        self.create_text(x + s(16) + sw / 2, y + s(18) + sh / 2, text=selo,
                         fill="#FFFFFF", font=self.fontes["selo"], tags=(tag,))

        if tipo == "imagem":
            self._conteudo_imagem(x, y, w, tag, cor)
            return
        for i, fracao in enumerate((1.0, 1.0, 0.82, 1.0, 0.5)):
            ly = y + s(66) + i * s(18)
            x_fim = x + s(16) + (w - s(32)) * fracao
            item = self.create_rectangle(x + s(16), ly, x_fim, ly + s(6),
                                         fill=LINHA_FOLHA, outline="", tags=(tag,))
            if lado == "b":
                self.linhas_b.append(item)

    def _conteudo_imagem(self, x, y, w, tag, cor):
        """Uma paisagem simples (sol e montanhas) no lugar das linhas de texto."""
        s = self.s
        x1, y1 = x + s(16), y + s(66)
        x2, y2 = x + w - s(16), y1 + s(84)
        claro, medio, forte = (misturar(cor, "#FFFFFF", t) for t in (0.88, 0.55, 0.4))
        self.create_rectangle(x1, y1, x2, y2, fill=claro, outline="", tags=(tag,))
        self.create_oval(x2 - s(34), y1 + s(10), x2 - s(16), y1 + s(28),
                         fill=medio, outline="", tags=(tag,))
        self.create_polygon(x1, y2, x1 + s(34), y1 + s(34), x1 + s(66), y2,
                            fill=medio, outline="", tags=(tag,))
        self.create_polygon(x1 + s(40), y2, x1 + s(80), y1 + s(46), x2, y2,
                            fill=forte, outline="", tags=(tag,))
        for i, fracao in enumerate((1.0, 0.55)):
            ly = y2 + s(16) + i * s(16)
            self.create_rectangle(x1, ly, x1 + (x2 - x1) * fracao, ly + s(6),
                                  fill=LINHA_FOLHA, outline="", tags=(tag,))

    def _ponto(self, t):
        return self.trilha_x0 + (self.trilha_x1 - self.trilha_x0) * t

    def _trecho(self, t0, t1, tag):
        """Desenha parte da trilha, com a cor indo da origem ao destino."""
        passos = max(1, int((t1 - t0) * 48))
        for i in range(passos):
            a = t0 + (t1 - t0) * i / passos
            b = t0 + (t1 - t0) * (i + 1) / passos
            self.create_line(self._ponto(a), self.y_meio, self._ponto(b) + 1, self.y_meio,
                             fill=misturar(self.cor_a, self.cor_b, (a + b) / 2),
                             width=self.s(3), capstyle="round", tags=tag)

    # ---- textos e estado
    def _texto(self, tag, texto, padrao, chave, limite):
        fonte = self.fontes[chave]
        if texto:
            cor = TINTA if chave == "nome" else TEXTO_SUAVE
            self.itemconfigure(tag, text=encurtar(texto, fonte, limite), fill=cor)
        else:
            self.itemconfigure(tag, text=padrao, fill=TEXTO_SUAVE)

    def definir_textos(self, a_nome, a_detalhe, b_nome, b_detalhe, vazio_a, vazio_b):
        self._texto("a_l1", a_nome, vazio_a, "nome", self.s(300))
        self._texto("a_l2", a_detalhe, "", "pequena", self.s(300))
        self._texto("b_l1", b_nome, vazio_b, "nome", self.fw)
        self._texto("b_l2", b_detalhe, "", "pequena", self.fw)

    def marcar(self, tem_a, tem_b):
        self.tem_a, self.tem_b = tem_a, tem_b
        self._cores_bordas()

    def _cores_bordas(self):
        self.itemconfigure("folha_a_borda", outline=self.cor_a if self.tem_a else LINHA)
        self.itemconfigure("folha_b_borda", outline=self.cor_b if self.tem_b else LINHA)

    def _entrar(self, _evento):
        if self.interativo:
            self.configure(cursor="hand2")

    def _sair(self, _evento):
        self.configure(cursor="")

    # ---- animação da conversão
    def iniciar_pulso(self):
        self.parar_pulso()
        self._pulso_pos = -0.3
        self._passo_pulso()

    def _passo_pulso(self):
        self.delete("pulso")
        inicio = max(0.0, self._pulso_pos)
        fim = min(1.0, self._pulso_pos + 0.3)
        if fim > inicio:
            self._trecho(inicio, fim, "pulso")
        self._pulso_pos += 0.02
        if self._pulso_pos > 1:
            self._pulso_pos = -0.3
        self._pulso_id = self.after(25, self._passo_pulso)

    def parar_pulso(self):
        if self._pulso_id is not None:
            self.after_cancel(self._pulso_id)
            self._pulso_id = None
        self.delete("pulso")

    def mostrar_sucesso(self):
        """Trilha completa, o conteúdo 'chega' na folha de destino e aparece o visto."""
        self.parar_pulso()
        self.itemconfigure("trilha_base", state="hidden")
        self._trecho(0, 1, "completo")
        self._escrever(0, self.sessao)

    def _escrever(self, i, sessao):
        if sessao != self.sessao:
            return
        if i < len(self.linhas_b):
            self.itemconfigure(self.linhas_b[i], fill=misturar(self.cor_b, "#FFFFFF", 0.7))
            self.after(110, self._escrever, i + 1, sessao)
        else:
            self._desenhar_visto()

    def _desenhar_visto(self):
        s = self.s
        cx = self.x_b + self.fw / 2
        cy = self.y0 + self.fh - s(32)
        r = s(19)
        self.create_oval(cx - r, cy - r, cx + r, cy + r, fill=self.cor_b,
                         outline=self.cor_b, tags="visto")
        self.create_line(cx - s(8), cy + s(1), cx - s(2), cy + s(7), cx + s(9), cy - s(6),
                         fill="#FFFFFF", width=s(3), capstyle="round", joinstyle="round",
                         tags="visto")

    def reiniciar(self):
        self.sessao += 1
        self.parar_pulso()
        self.delete("completo")
        self.delete("visto")
        self.itemconfigure("trilha_base", state="normal")
        for linha in self.linhas_b:
            self.itemconfigure(linha, fill=LINHA_FOLHA)
        self._cores_bordas()


# ---------------------------------------------------------------- Aplicativo
class App(tk.Tk):
    LARGURA, ALTURA = 720, 640

    def __init__(self):
        super().__init__()
        self.k = self.winfo_fpixels("1i") / 96 if sys.platform.startswith("win") else 1.0
        s = self.s

        self.modo = "word"
        self.arquivos = []
        self.pasta = ""
        self.pasta_automatica = False
        self.convertendo = False
        self.fila = queue.Queue()  # recebe o resultado da thread de conversão
        cfg = MODOS[self.modo]

        self.title("Conversor de arquivos")
        self.configure(bg=FUNDO)
        self.resizable(False, False)
        w, h = s(self.LARGURA), s(self.ALTURA)
        x = (self.winfo_screenwidth() - w) // 2
        y = max(0, (self.winfo_screenheight() - h) // 3)
        self.geometry(f"{w}x{h}+{x}+{y}")

        self.fontes = self._criar_fontes()

        corpo = tk.Frame(self, bg=FUNDO)
        corpo.pack(fill="both", expand=True, padx=s(40), pady=(s(36), s(28)))

        # Título e seletor de modo
        topo = tk.Frame(corpo, bg=FUNDO)
        topo.pack(fill="x")
        tk.Label(topo, text="Conversor", font=self.fontes["titulo"],
                 bg=FUNDO, fg=TINTA).pack(side="left")
        self.seletor = Seletor(
            topo, [(chave, MODOS[chave]["rotulo"]) for chave in MODOS], self.modo,
            self.trocar_modo, s(320), s(40), s(10), self.fontes["botao"])
        self.seletor.pack(side="right")

        self.subtitulo = tk.Label(corpo, text=cfg["subtitulo"], font=self.fontes["texto"],
                                  bg=FUNDO, fg=TEXTO_SUAVE)
        self.subtitulo.pack(anchor="w", pady=(s(4), s(26)))

        self.fluxo = Fluxo(corpo, self.k, self.fontes, cfg["origem"], cfg["destino"],
                           self.escolher_origem, self.escolher_pasta)
        self.fluxo.pack(anchor="w")

        # Botões de escolha, alinhados com cada folha
        escolhas = tk.Frame(corpo, bg=FUNDO)
        escolhas.pack(fill="x", pady=(s(2), s(26)))
        escolhas.columnconfigure(1, weight=1)
        self.btn_origem = self._botao_secundario(escolhas, cfg["escolher"], self.escolher_origem)
        self.btn_origem.grid(row=0, column=0, sticky="w")
        self.btn_pasta = self._botao_secundario(escolhas, "Escolher pasta", self.escolher_pasta)
        self.btn_pasta.grid(row=0, column=2, sticky="e")

        # Ação principal
        rodape = tk.Frame(corpo, bg=FUNDO)
        rodape.pack(fill="x")
        self.btn_converter = Botao(
            rodape, cfg["acao"], self.iniciar, s(240), s(48), s(10),
            self.fontes["botao"], TINTA, TINTA_HOVER, "#FFFFFF",
        )
        self.btn_converter.pack(side="left")
        self.btn_abrir = self._botao_secundario(rodape, "Abrir pasta", self.abrir_pasta, altura=48)

        self.status = tk.Label(corpo, text="", anchor="w", justify="left",
                               wraplength=s(640), font=self.fontes["texto"],
                               bg=FUNDO, fg=TEXTO_SUAVE)
        self.status.pack(fill="x", pady=(s(16), 0))

        self.atualizar()

    # ---- construção
    def s(self, valor):
        return int(round(valor * self.k))

    def _criar_fontes(self):
        familias = set(tkfont.families(self))
        base = next((f for f in ("Segoe UI", "SF Pro Text", "Helvetica Neue", "Ubuntu",
                                 "DejaVu Sans") if f in familias), "Helvetica")
        semi = "Segoe UI Semibold" if "Segoe UI Semibold" in familias else None

        def fonte(tamanho, forte=False):
            if forte and semi:
                return tkfont.Font(family=semi, size=tamanho)
            return tkfont.Font(family=base, size=tamanho, weight="bold" if forte else "normal")

        return {
            "titulo": fonte(24, True),
            "texto": fonte(11),
            "nome": fonte(10, True),
            "pequena": fonte(9),
            "selo": fonte(9, True),
            "botao": fonte(10, True),
        }

    def _botao_secundario(self, master, texto, comando, altura=40):
        s = self.s
        return Botao(master, texto, comando, s(160), s(altura), s(8), self.fontes["botao"],
                     FOLHA, HOVER_CLARO, TINTA, contorno=LINHA)

    # ---- estado da tela
    def definir_status(self, texto, cor=TEXTO_SUAVE):
        self.status.configure(text=texto, fg=cor)

    def nome_saida(self):
        base = os.path.splitext(os.path.basename(self.arquivos[0]))[0]
        if len(self.arquivos) > 1:
            base += "_combinado"
        return base + MODOS[self.modo]["ext"]

    def atualizar(self):
        """Redesenha textos e botões de acordo com o que já foi escolhido."""
        cfg = MODOS[self.modo]
        n = len(self.arquivos)

        if n == 0:
            nome = detalhe = ""
        elif n == 1:
            nome = os.path.basename(self.arquivos[0])
            detalhe = formatar_tamanho(tamanho_total(self.arquivos))
        else:
            nome = f"{n} {cfg['plural']}"
            detalhe = formatar_tamanho(tamanho_total(self.arquivos))

        pasta_nome = ""
        if self.pasta:
            pasta_nome = os.path.basename(os.path.normpath(self.pasta)) or self.pasta
        saida = self.nome_saida() if (n and self.pasta) else ""

        self.fluxo.definir_textos(nome, detalhe, pasta_nome, saida, cfg["vazio"], "Pasta não escolhida")
        self.fluxo.marcar(n > 0, bool(self.pasta))
        self.btn_origem.definir_texto(cfg["trocar"] if n else cfg["escolher"])
        self.btn_pasta.definir_texto("Trocar pasta" if self.pasta else "Escolher pasta")
        self.btn_converter.definir_ativo(bool(n and self.pasta))

        if not n and not self.pasta:
            texto = cfg["msg_inicio"]
        elif not n:
            texto = cfg["msg_falta"]
        elif self.pasta_automatica:
            texto = cfg["msg_auto"]
        else:
            texto = cfg["msg_pronto"]
        if n > 1 and cfg["ordem"]:
            texto += " " + cfg["ordem"]
        self.definir_status(texto)

    def _limpar_resultado(self):
        self.fluxo.reiniciar()
        self.btn_abrir.pack_forget()

    # ---- modo
    def trocar_modo(self, chave):
        if self.convertendo or chave == self.modo:
            return
        self.modo = chave
        cfg = MODOS[chave]
        self.arquivos = []
        if self.pasta_automatica:
            self.pasta, self.pasta_automatica = "", False
        self.subtitulo.configure(text=cfg["subtitulo"])
        self.btn_converter.definir_texto(cfg["acao"])
        self.btn_abrir.pack_forget()
        self.fluxo.configurar(cfg["origem"], cfg["destino"])
        self.atualizar()

    # ---- escolhas
    def escolher_origem(self):
        if self.convertendo:
            return
        cfg = MODOS[self.modo]
        if cfg["multiplos"]:
            escolhidos = self.tk.splitlist(
                filedialog.askopenfilenames(title=cfg["dialogo"], filetypes=cfg["tipos"]))
        else:
            unico = filedialog.askopenfilename(title=cfg["dialogo"], filetypes=cfg["tipos"])
            escolhidos = (unico,) if unico else ()
        if not escolhidos:
            return

        self.arquivos = sorted((os.path.normpath(c) for c in escolhidos), key=chave_natural)
        if not self.pasta or self.pasta_automatica:
            self.pasta = os.path.dirname(self.arquivos[0])
            self.pasta_automatica = True
        self._limpar_resultado()
        self.atualizar()

    def escolher_pasta(self):
        if self.convertendo:
            return
        pasta = filedialog.askdirectory(title="Escolher pasta de destino", mustexist=True)
        if not pasta:
            return
        self.pasta = os.path.normpath(pasta)
        self.pasta_automatica = False
        self._limpar_resultado()
        self.atualizar()

    # ---- conversão
    def iniciar(self):
        nome = self.nome_saida()
        destino = os.path.join(self.pasta, nome)

        if os.path.exists(destino) and not messagebox.askyesno(
                "Substituir arquivo?", f"Já existe {nome} nessa pasta. Deseja substituir?"):
            return

        self.convertendo = True
        self._limpar_resultado()
        self.fluxo.interativo = False
        for controle in (self.seletor, self.btn_origem, self.btn_pasta, self.btn_converter):
            controle.definir_ativo(False)
        self.btn_converter.definir_texto("Aguarde…")
        self.definir_status(MODOS[self.modo]["msg_andamento"])
        self.fluxo.iniciar_pulso()

        threading.Thread(target=self._converter,
                         args=(self.modo, list(self.arquivos), destino), daemon=True).start()
        self.after(100, self._verificar_fila)

    def _converter(self, modo, origens, destino):
        """Roda em outra thread; só devolve o resultado pela fila."""
        cv = None
        try:
            if modo == "word":
                cv = Converter(origens[0])
                cv.convert(destino)
            else:
                imagens_para_pdf(origens, destino)
            self.fila.put(("sucesso", destino))
        except Exception as erro:
            self.fila.put(("falha", modo, erro))
        finally:
            if cv is not None:
                try:
                    cv.close()
                except Exception:
                    pass

    def _verificar_fila(self):
        """Roda na thread principal: espera o resultado sem travar a janela."""
        try:
            resultado = self.fila.get_nowait()
        except queue.Empty:
            self.after(100, self._verificar_fila)
            return
        if resultado[0] == "sucesso":
            self._sucesso(resultado[1])
        else:
            self._falha(resultado[1], resultado[2])

    def _liberar(self):
        self.convertendo = False
        self.fluxo.interativo = True
        for controle in (self.seletor, self.btn_origem, self.btn_pasta, self.btn_converter):
            controle.definir_ativo(True)
        self.btn_converter.definir_texto(MODOS[self.modo]["acao"])

    def _sucesso(self, destino):
        self._liberar()
        self.fluxo.mostrar_sucesso()
        self.btn_abrir.pack(side="left", padx=(self.s(12), 0))
        self.definir_status(f"Pronto. {os.path.basename(destino)} foi salvo em {self.pasta}.", TINTA)

    def _falha(self, modo, erro):
        self._liberar()
        self._limpar_resultado()
        detalhe = str(erro).strip().replace("\n", " ")[:140]
        self.definir_status(
            MODOS[modo]["msg_erro"] + "\nDetalhe: " + (detalhe or type(erro).__name__), ERRO)

    def abrir_pasta(self):
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.pasta)
            elif sys.platform == "darwin":
                subprocess.run(["open", self.pasta])
            else:
                subprocess.run(["xdg-open", self.pasta])
        except Exception:
            pass


if __name__ == "__main__":
    App().mainloop()
