"""
Módulo isolado: geração do relatório PDF (Laranja + Vermelho) com logo centralizada no topo.
Visual de planilha compacta — linhas finas, células densas, igual ao Excel original.
"""

import io
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer, Image, HRFlowable
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT


# ── helpers de cor ───────────────────────────────────────────────────────────
def _hex(h):
    h = h.lstrip("#")
    return colors.Color(int(h[0:2],16)/255, int(h[2:4],16)/255, int(h[4:6],16)/255)

COR_LARANJA       = _hex("FF6B00")
COR_VERMELHO      = _hex("CC0000")
COR_LARANJA_FONT  = _hex("B35900")
COR_VERMELHO_FONT = _hex("CC0000")
COR_BG_PAGE       = colors.white
COR_COL_HEADER    = _hex("2A2A2A")
COR_ROW_LARANJA   = _hex("FFF3E0")
COR_ROW_VERMELHO  = _hex("FFEBEE")
COR_ROW_DARK      = _hex("F9F9F9")
COR_SUBTITULO_BG  = _hex("EEEEEE")
COR_SUBTITULO_TXT = _hex("444444")
COR_BORDER        = _hex("CCCCCC")
BRANCO            = colors.white
COR_TEXTO         = _hex("1A1A1A")


def _fmt(val):
    try:
        f = float(val)
        return str(int(f)) if f == int(f) else str(round(f, 2))
    except Exception:
        s = str(val) if val is not None else ""
        return "" if s in ("nan", "None", "NaN", "") else s


def _p(txt, size=7, bold=False, color=BRANCO, align=TA_LEFT):
    style = ParagraphStyle(
        "s",
        fontSize=size,
        fontName="Helvetica-Bold" if bold else "Helvetica",
        textColor=color,
        alignment=align,
        leading=size + 2,
        wordWrap="LTR",
        leftPadding=0, rightPadding=0,
        spaceBefore=0, spaceAfter=0,
    )
    return Paragraph(str(txt), style)


def _secao(df_sec, titulo, cor_header, cor_row, cor_font, col_widths, cor_bolinha=None):
    if df_sec.empty:
        return []

    df = df_sec.copy()
    if "SALDO TOTAL" in df.columns:
        df["__s"] = pd.to_numeric(df["SALDO TOTAL"], errors="coerce").fillna(0)
        df = df.sort_values("__s", ascending=False).drop(columns=["__s"])

    colunas = [c for c in df.columns if not c.startswith("_")]
    n_cols  = len(colunas)
    total_w = sum(col_widths)

    # ── bloco título + subtítulo ─────────────────────────────────────────
    tit_tbl = Table(
        [[_p(titulo.upper(), size=10, bold=True, color=BRANCO, align=TA_CENTER)],
         [_p(f"Total de itens: {len(df)}", size=7, color=COR_SUBTITULO_TXT, align=TA_CENTER)]],
        colWidths=[total_w]
    )
    tit_tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (-1,0), cor_header),
        ("BACKGROUND",   (0,1), (-1,1), COR_SUBTITULO_BG),
        ("TOPPADDING",   (0,0), (-1,0), 5),
        ("BOTTOMPADDING",(0,0), (-1,0), 5),
        ("TOPPADDING",   (0,1), (-1,1), 2),
        ("BOTTOMPADDING",(0,1), (-1,1), 2),
        ("LEFTPADDING",  (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("LINEBELOW",    (0,1), (-1,1), 0.8, cor_header),
    ]))

    # ── cabeçalho das colunas ────────────────────────────────────────────
    dot_w = [4*mm] if cor_bolinha else []
    col_widths = dot_w + list(col_widths)
    header = (([_p("", size=7, color=BRANCO, align=TA_CENTER)] if cor_bolinha else []) +
              [_p(c, size=7, bold=True, color=BRANCO, align=TA_CENTER) for c in colunas])
    rows   = [header]

    # ── linhas de dados ──────────────────────────────────────────────────
    for _, row in df.iterrows():
        linha = []
        if cor_bolinha:
            linha.append(_p("●", size=9, bold=True, color=cor_bolinha, align=TA_CENTER))
        for c in colunas:
            v   = row.get(c, "")
            txt = _fmt(v)
            al  = TA_CENTER if c in ("SALDO TOTAL", "ENTRADA", "SAIDA") else TA_LEFT
            linha.append(_p(txt, size=7, color=COR_TEXTO, align=al))
        rows.append(linha)

    # ── subtotal ─────────────────────────────────────────────────────────
    has_sub = "SALDO TOTAL" in colunas
    if has_sub:
        idx_st    = colunas.index("SALDO TOTAL")
        total_val = int(
            df["SALDO TOTAL"].apply(lambda x: pd.to_numeric(x, errors="coerce")).fillna(0).sum()
        )
        sub = []
        for i in range(n_cols):
            if i == idx_st - 1:
                sub.append(_p("SUBTOTAL", size=7, bold=True, color=BRANCO, align=TA_RIGHT))
            elif i == idx_st:
                sub.append(_p(str(total_val), size=8, bold=True, color=BRANCO, align=TA_CENTER))
            else:
                sub.append(_p(""))
        rows.append(sub)

    # ── monta tabela ─────────────────────────────────────────────────────
    tbl = Table(rows, colWidths=col_widths, repeatRows=1)
    cmds = [
        ("BACKGROUND",   (0,0), (-1,0), COR_COL_HEADER),
        ("ALIGN",        (0,0), (-1,0), "CENTER"),
        ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
        # padding mínimo = visual de planilha apertada
        ("TOPPADDING",   (0,0), (-1,0), 3),
        ("BOTTOMPADDING",(0,0), (-1,0), 3),
        ("TOPPADDING",   (0,1), (-1,-1), 2),
        ("BOTTOMPADDING",(0,1), (-1,-1), 2),
        ("LEFTPADDING",  (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        # grade fina como planilha
        ("GRID",         (0,0), (-1,-1), 0.25, COR_BORDER),
        ("LINEBELOW",    (0,0), (-1,0),  0.8,  cor_header),
    ]
    # cores alternadas nas linhas
    for r in range(len(df)):
        bg = cor_row if r % 2 == 0 else COR_ROW_DARK
        cmds.append(("BACKGROUND", (0, r+1), (-1, r+1), bg))
    # subtotal
    if has_sub:
        last = len(rows) - 1
        cmds += [
            ("BACKGROUND",   (0,last), (-1,last), cor_header),
            ("TOPPADDING",   (0,last), (-1,last), 4),
            ("BOTTOMPADDING",(0,last), (-1,last), 4),
            ("LINEABOVE",    (0,last), (-1,last), 0.8, cor_header),
        ]
    tbl.setStyle(TableStyle(cmds))

    return [tit_tbl, tbl, Spacer(1, 4*mm)]


def gerar_bytes_relatorio(df_alerta, df_zerado, logo_path="logo_luft.png"):
    """
    Gera PDF compacto (visual de planilha) em memória.
    df_alerta → saldo 1-3 (laranja) | df_zerado → saldo 0 (vermelho)
    """
    buf    = io.BytesIO()
    df_ref = df_alerta if not df_alerta.empty else df_zerado
    colunas = [c for c in df_ref.columns if not c.startswith("_")] if not df_ref.empty else []

    # Usa paisagem se tiver muitas colunas
    pagesize = landscape(A4) if len(colunas) > 7 else A4
    marg     = 10 * mm
    PAGE_W   = pagesize[0] - 2 * marg

    doc = SimpleDocTemplate(
        buf, pagesize=pagesize,
        leftMargin=marg, rightMargin=marg,
        topMargin=marg, bottomMargin=marg,
        title="Relatório de Alertas — Almoxarifado",
    )

    def fundo(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(COR_BG_PAGE)
        canvas.rect(0, 0, pagesize[0], pagesize[1], fill=1, stroke=0)
        canvas.restoreState()

    story = []

    # ── LOGO grande e centralizada ───────────────────────────────────────
    LOGO_MAX_H = 38 * mm
    LOGO_MAX_W = PAGE_W * 0.65

    try:
        img       = Image(logo_path)
        ratio     = min(LOGO_MAX_H / img.imageHeight, LOGO_MAX_W / img.imageWidth)
        img.drawWidth  = img.imageWidth  * ratio
        img.drawHeight = img.imageHeight * ratio
        img.hAlign = "CENTER"
        story.append(img)
    except Exception:
        story.append(_p(
            "ALMOXARIFADO — RELATÓRIO DE ALERTAS",
            size=15, bold=True, color=COR_LARANJA, align=TA_CENTER
        ))

    story.append(Spacer(1, 3*mm))
    story.append(HRFlowable(width="100%", thickness=1.2, color=COR_LARANJA, spaceAfter=4*mm))

    # ── Larguras das colunas ─────────────────────────────────────────────
    if colunas:
        FIXAS = {
            "SALDO TOTAL": 18*mm,
            "ENTRADA":     16*mm,
            "SAIDA":       16*mm,
            "CÓDIGO":      22*mm,
            "STATUS":      20*mm,
            "POSIÇÃO":     16*mm,
        }
        larguras  = [FIXAS.get(c) for c in colunas]
        w_fixo    = sum(w for w in larguras if w is not None)
        n_livre   = larguras.count(None)
        w_livre   = max((PAGE_W - w_fixo) / n_livre, 20*mm) if n_livre > 0 else 0
        larguras  = [w if w is not None else w_livre for w in larguras]
        # normaliza se ultrapassar
        total = sum(larguras)
        if total > PAGE_W:
            fator   = PAGE_W / total
            larguras = [w * fator for w in larguras]
    else:
        larguras = [PAGE_W]

    # ── seções ───────────────────────────────────────────────────────────
    story += _secao(df_alerta, "⚠  ALERTA — ESTOQUE BAIXO  (saldo 1–3)",
                    COR_LARANJA,  COR_ROW_LARANJA,  COR_LARANJA_FONT,  larguras,
                    cor_bolinha=COR_LARANJA)

    story += _secao(df_zerado, "🚨  SEM ESTOQUE — ITENS ZERADOS  (saldo 0)",
                    COR_VERMELHO, COR_ROW_VERMELHO, COR_VERMELHO_FONT, larguras,
                    cor_bolinha=COR_VERMELHO)

    doc.build(story, onFirstPage=fundo, onLaterPages=fundo)
    buf.seek(0)
    return buf.getvalue()
