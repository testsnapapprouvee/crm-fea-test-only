# =============================================================================
# pdf_generator.py  —  CRM Asset Management  —  Amundi Edition
# Page de garde : reproduit fidèlement le modèle Canva fourni
#   - Fond marine #001c4b plein page
#   - Logo Amundi en haut à droite (extrait du PDF Canva)
#   - "INTERNAL PURPOSE ONLY" en ciel haut gauche
#   - "Executive Report" (28pt bold) + "Pipeline & Reporting"
#   - "Amundi Asset Management" en bleu moyen
#   - Périmètre + Date en ciel
#   - Bande KPI fond marine légèrement différent avec séparateurs ciel
#   - Disclaimer italique en bas
#   - Footer "Executive Report - Internal Only" + "Page 1"
# =============================================================================

import io, os
from datetime import date
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import pandas as pd
import numpy as np

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image, PageBreak, KeepTogether,
)
from reportlab.platypus.flowables import Flowable
from reportlab.pdfgen import canvas as rl_canvas

# ---------------------------------------------------------------------------
# COULEURS
# ---------------------------------------------------------------------------
COL_MARINE   = HexColor("#001c4b")
COL_MARINE2  = HexColor("#012060")   # bande KPI légèrement plus claire
COL_CIEL     = HexColor("#019ee1")
COL_CIEL2    = HexColor("#089ee0")
COL_ORANGE   = HexColor("#f07d00")
COL_BLANC    = HexColor("#ffffff")
COL_GRIS     = HexColor("#e8e8e8")
COL_TEXTE    = HexColor("#444444")
COL_BLEU_MD  = HexColor("#7ab8d8")
COL_VERT     = HexColor("#1a7a3c")
COL_ROUGE    = HexColor("#8b2020")
COL_HEADER   = HexColor("#f0f5fa")

HX_MARINE="#001c4b"; HX_CIEL="#019ee1"; HX_BLANC="#ffffff"; HX_GRIS="#e8e8e8"; HX_TEXTE="#444444"

# ---------------------------------------------------------------------------
# PAGE
# ---------------------------------------------------------------------------
PAGE_W, PAGE_H = A4    # 595.3 x 841.9
MARGIN_H  = 2.0 * cm
MARGIN_V  = 2.0 * cm
USABLE_W  = PAGE_W - 2 * MARGIN_H

DONUT_W   = USABLE_W * 0.46
DONUT_H   = DONUT_W  * 0.95
TOP10_W   = USABLE_W
NAV_W     = USABLE_W
NAV_H     = NAV_W * 0.40

PALETTE = ["#1a5e8a","#001c4b","#4a8fbd","#003f7a","#2c7fb8","#004f8c","#6baed6","#08519c","#9ecae1","#003060"]

MPL_RC = {"font.family":"DejaVu Sans","font.size":8.5,"axes.spines.top":False,"axes.spines.right":False,
           "axes.edgecolor":HX_GRIS,"axes.linewidth":0.6,"axes.labelcolor":HX_MARINE,
           "xtick.color":HX_MARINE,"ytick.color":HX_MARINE,"figure.facecolor":HX_BLANC,
           "axes.facecolor":HX_BLANC,"grid.color":HX_GRIS,"grid.linewidth":0.4}

# ---------------------------------------------------------------------------
# FLOWABLE
# ---------------------------------------------------------------------------
class ColorRect(Flowable):
    def __init__(self,w,h,color): super().__init__(); self.width=w; self.height=h; self.color=color
    def draw(self): self.canv.setFillColor(self.color); self.canv.rect(0,0,self.width,self.height,fill=1,stroke=0)

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def fmt_aum(v):
    try: v=float(v)
    except: return "-"
    if v==0: return "0.0 M EUR"
    if v>=1_000_000_000: return "{:.1f} Md EUR".format(v/1e9)
    return "{:.1f} M EUR".format(v/1e6)

def _logo_path():
    """Chemin vers le logo Amundi marine (généré une fois)."""
    base = os.path.dirname(os.path.abspath(__file__))
    p = os.path.join(base, "amundi_logo_marine.png")
    return p if os.path.exists(p) else None

# ---------------------------------------------------------------------------
# PAGE DE GARDE — canvas direct, fidèle au modèle Canva
# Toutes les coordonnées sont converties depuis les % HTML Canva :
#   x_pt = PAGE_W * x_pct/100
#   y_pt = PAGE_H - PAGE_H * top_pct/100  (ReportLab: 0=bas)
# ---------------------------------------------------------------------------
def _draw_cover(c, mode_comex, kpis, fonds_perimetre):
    W = PAGE_W; H = PAGE_H

    # ---- Fond marine plein ----
    c.setFillColor(COL_MARINE)
    c.rect(0, 0, W, H, fill=1, stroke=0)

    # ---- Logo Amundi — haut droite ----
    logo = _logo_path()
    if logo:
        # Dans le Canva : logo environ 20% de largeur, positionné à droite
        logo_w = W * 0.22; logo_h = logo_w * 267/465
        logo_x = W - logo_w - W*0.05
        logo_y = H - logo_h - H*0.06
        c.drawImage(logo, logo_x, logo_y, width=logo_w, height=logo_h, mask="auto")

    # ---- "INTERNAL PURPOSE ONLY" — top: 2.88%, left: 9.15% ----
    c.setFont("Helvetica-Bold", 8)
    c.setFillColor(COL_CIEL2)
    c.drawString(W*0.0915, H - H*0.0288 - 8, "INTERNAL PURPOSE ONLY")
    if mode_comex:
        c.setFont("Helvetica", 7)
        c.drawString(W*0.0915, H - H*0.0288 - 20, "| MODE COMEX — Noms anonymises")

    # ---- "Executive Report" — top: 7.07%, bold 28pt ----
    c.setFont("Helvetica-Bold", 28)
    c.setFillColor(COL_BLANC)
    c.drawString(W*0.0915, H - H*0.0707 - 28, "Executive Report")

    # ---- "Pipeline & Reporting" — top: 10.20%, 18pt ----
    c.setFont("Helvetica", 18)
    c.setFillColor(COL_BLANC)
    c.drawString(W*0.0915, H - H*0.1020 - 18, "Pipeline & Reporting")

    # ---- Ligne séparatrice ciel ----
    c.setStrokeColor(COL_CIEL)
    c.setLineWidth(0.8)
    c.line(W*0.0915, H - H*0.135, W*0.60, H - H*0.135)

    # ---- "Amundi Asset Management" — top: 15.08%, bleu moyen 13pt ----
    c.setFont("Helvetica", 13)
    c.setFillColor(COL_BLEU_MD)
    c.drawString(W*0.0915, H - H*0.1508 - 13, "Amundi Asset Management")

    # ---- Périmètre — top: 17.22%, blanc 11pt ----
    perim_str = ", ".join(fonds_perimetre) if fonds_perimetre else "Tous les fonds"
    perim_display = "Perimetre : " + (perim_str[:78]+"..." if len(perim_str)>78 else perim_str)
    c.setFont("Helvetica", 10)
    c.setFillColor(COL_BLANC)
    c.drawString(W*0.0915, H - H*0.1722 - 10, perim_display)

    # ---- Date — top: 23.53%, ciel bold 15pt ----
    today_str = date.today().strftime("%d %B %Y")
    c.setFont("Helvetica-Bold", 15)
    c.setFillColor(COL_CIEL2)
    c.drawString(W*0.0915, H - H*0.2353 - 15, today_str)

    # ---- Bande KPI ----
    # Dans le Canva les labels sont à top: 35.52% et les valeurs à top: 38.43%
    # La bande KPI fait environ 8% de hauteur
    kpi_band_y   = H - H*0.4400   # bas de la bande
    kpi_band_h   = H * 0.090
    kpi_band_top = kpi_band_y + kpi_band_h

    # Fond bande KPI — légèrement plus clair que le fond page
    c.setFillColor(HexColor("#01224e"))
    c.rect(0, kpi_band_y, W, kpi_band_h, fill=1, stroke=0)

    # Données KPI
    kpi_items = [
        ("AUM Total",       fmt_aum(kpis.get("total_funded",0))),
        ("Pipeline Actif",  fmt_aum(kpis.get("pipeline_actif",0))),
        ("Taux Conversion", "{:.1f}%".format(kpis.get("taux_conversion",0))),
        ("Deals Actifs",    str(kpis.get("nb_deals_actifs",0))),
    ]
    # Positions X Canva (gauche de chaque bloc) : 9.87%, 32.75%, 55.14%, 81.09%
    kpi_xs = [W*0.0987, W*0.3275, W*0.5514, W*0.8109]

    for i, ((label, value), kx) in enumerate(zip(kpi_items, kpi_xs)):
        # Label — row top: 35.52%
        label_y = kpi_band_y + kpi_band_h * 0.62
        value_y = kpi_band_y + kpi_band_h * 0.20

        c.setFont("Helvetica", 7.5)
        c.setFillColor(COL_BLANC)
        c.drawString(kx, label_y, label)

        c.setFont("Helvetica-Bold", 14)
        c.setFillColor(COL_BLANC)
        c.drawString(kx, value_y, value)

        # Séparateur vertical ciel entre blocs
        if i < 3:
            sep_x = kpi_xs[i+1] - W*0.02
            c.setStrokeColor(COL_CIEL)
            c.setLineWidth(0.7)
            c.line(sep_x, kpi_band_y + kpi_band_h*0.12, sep_x, kpi_band_y + kpi_band_h*0.88)

    # ---- Disclaimer italique — juste sous la bande KPI ----
    disc_y = kpi_band_y - 32
    c.setFont("Helvetica-Oblique", 7)
    c.setFillColor(HexColor("#aaaaaa"))
    disc1 = "Document strictement confidentiel a usage interne exclusif. Reproduction et diffusion externe interdites."
    disc2 = "Les performances passees ne prejudgent pas des performances futures." + (" Mode confidentiel actif : noms clients anonymises." if mode_comex else "")
    c.drawCentredString(W/2, disc_y + 9, disc1)
    c.drawCentredString(W/2, disc_y - 2, disc2)

    # ---- Footer ----
    footer_y = H * 0.025
    c.setFont("Helvetica", 7)
    c.setFillColor(COL_BLEU_MD)
    c.drawString(W*0.0987, footer_y, "Executive Report - Internal Only")
    c.drawRightString(W - W*0.04, footer_y, "Page 1")


# ---------------------------------------------------------------------------
# GRAPHIQUES
# ---------------------------------------------------------------------------
def _make_donut_png(labels, values, title, fw, fh):
    plt.rcParams.update(MPL_RC)
    fig, ax = plt.subplots(figsize=(fw, fh))
    if not labels or not values or sum(values)==0:
        ax.text(0.5,0.5,"Aucune donnee",ha="center",va="center",fontsize=8,color=HX_MARINE,transform=ax.transAxes)
        ax.axis("off")
    else:
        colors=[PALETTE[i%len(PALETTE)] for i in range(len(labels))]
        _,_,autotexts=ax.pie(values,colors=colors,autopct="%1.0f%%",startangle=90,pctdistance=0.72,
                              wedgeprops={"width":0.52,"edgecolor":HX_BLANC,"linewidth":1.5})
        for at in autotexts: at.set_fontsize(7.5); at.set_color(HX_BLANC); at.set_fontweight("bold")
        ax.text(0,0.10,fmt_aum(sum(values)),ha="center",va="center",fontsize=8,fontweight="bold",color=HX_MARINE)
        ax.text(0,-0.15,"Finance",ha="center",va="center",fontsize=6.5,color=HX_TEXTE)
        patches=[mpatches.Patch(color=colors[i],label="{}: {}".format(labels[i],fmt_aum(values[i]))) for i in range(len(labels))]
        ax.legend(handles=patches,loc="lower center",bbox_to_anchor=(0.5,-0.32),ncol=2,fontsize=6.5,frameon=False,labelcolor=HX_MARINE)
        ax.set_title(title,fontsize=8.5,fontweight="bold",color=HX_MARINE,pad=7)
    fig.subplots_adjust(left=0.02,right=0.98,top=0.88,bottom=0.26)
    buf=io.BytesIO(); fig.savefig(buf,format="png",dpi=150,facecolor=HX_BLANC); plt.close(fig); buf.seek(0)
    return buf

def _make_top10_png(deals, mode_comex, fw, fh, title="Top 10 Inflows — AUM Finance"):
    plt.rcParams.update(MPL_RC)
    fig, ax = plt.subplots(figsize=(fw, fh))
    deals_10=deals[:10]
    if not deals_10:
        ax.text(0.5,0.5,"Aucun deal",ha="center",va="center",fontsize=9,color=HX_MARINE,transform=ax.transAxes); ax.axis("off")
    else:
        labels=["{} - {}".format(d.get("type_client",""),d.get("region","")) if mode_comex else str(d.get("nom_client","")) for d in deals_10]
        values=[float(d.get("funded_aum",0)) for d in deals_10]
        pairs=sorted(zip(values,labels),reverse=True); values=[v for v,_ in pairs]; labels=[l for _,l in pairs]
        max_v=max(values) if values else 1.0
        colors=[HX_CIEL if i==0 else HX_MARINE for i in range(len(deals_10))]
        bars=ax.barh(range(len(labels)),values,color=colors,edgecolor=HX_BLANC,height=0.65,linewidth=0.3)
        for bar,val in zip(bars,values):
            ax.text(bar.get_width()+max_v*0.015,bar.get_y()+bar.get_height()/2,fmt_aum(val),va="center",ha="left",fontsize=8,color=HX_MARINE,fontweight="bold")
        ax.set_yticks(range(len(labels))); ax.set_yticklabels([l[:28] for l in labels],fontsize=8,color=HX_MARINE); ax.invert_yaxis()
        ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda x,_: fmt_aum(x))); ax.tick_params(axis="x",labelsize=7.5,colors=HX_MARINE)
        ax.set_xlim(0,max_v*1.32); ax.grid(axis="x",alpha=0.22,color=HX_GRIS,linewidth=0.35)
        ax.spines["left"].set_color(HX_GRIS); ax.spines["bottom"].set_color(HX_GRIS)
        ax.set_title(title,fontsize=10,fontweight="bold",color=HX_MARINE,pad=10)
    fig.subplots_adjust(left=0.26,right=0.86,top=0.92,bottom=0.06)
    buf=io.BytesIO(); fig.savefig(buf,format="png",dpi=150,facecolor=HX_BLANC); plt.close(fig); buf.seek(0)
    return buf

def _make_nav_png(nav_df, fw, fh):
    plt.rcParams.update(MPL_RC)
    fig, ax = plt.subplots(figsize=(fw, fh))
    if nav_df is None or not hasattr(nav_df,"columns") or nav_df.empty:
        ax.text(0.5,0.5,"Donnees NAV non disponibles",ha="center",va="center",fontsize=9,color=HX_MARINE,transform=ax.transAxes); ax.axis("off")
    else:
        NCOLORS=[HX_MARINE,HX_CIEL,"#1a5e8a","#4a8fbd","#003f7a","#2c7fb8","#004f8c","#6baed6"]
        for i,col in enumerate(nav_df.columns):
            s=nav_df[col].dropna()
            if s.empty: continue
            ax.plot(s.index,s.values,label=col,color=NCOLORS[i%len(NCOLORS)],linewidth=1.5,alpha=0.92) if len(s)>=2 else ax.scatter(s.index,s.values,color=NCOLORS[i%len(NCOLORS)],s=50,label=col)
        ax.axhline(100,color=HX_GRIS,linewidth=0.7,linestyle="dotted")
        ax.set_ylabel("Base 100",fontsize=7.5,color=HX_MARINE); ax.tick_params(colors=HX_MARINE,labelsize=7)
        for sp in ["top","right"]: ax.spines[sp].set_visible(False)
        ax.spines["left"].set_color(HX_GRIS); ax.spines["bottom"].set_color(HX_GRIS)
        ax.grid(axis="y",alpha=0.18,color=HX_GRIS); ax.grid(axis="x",visible=False)
        ax.legend(fontsize=7,frameon=True,framealpha=0.90,edgecolor=HX_GRIS,labelcolor=HX_MARINE,loc="upper left")
        ax.set_title("Evolution NAV - Base 100",fontsize=9,fontweight="bold",color=HX_MARINE,pad=8)
        plt.xticks(rotation=12,ha="right",fontsize=7)
    fig.subplots_adjust(left=0.08,right=0.97,top=0.90,bottom=0.14)
    buf=io.BytesIO(); fig.savefig(buf,format="png",dpi=150,facecolor=HX_BLANC); plt.close(fig); buf.seek(0)
    return buf

# ---------------------------------------------------------------------------
# STYLES
# ---------------------------------------------------------------------------
def _styles():
    def ps(name,**kw): return ParagraphStyle(name,**kw)
    return {
        "section": ps("section",fontName="Helvetica-Bold",fontSize=11,textColor=COL_CIEL,spaceBefore=10,spaceAfter=4),
        "subsect": ps("subsect",fontName="Helvetica-Bold",fontSize=9,textColor=COL_MARINE,spaceBefore=6,spaceAfter=3),
        "body":    ps("body",   fontName="Helvetica",     fontSize=8.5,textColor=COL_TEXTE,spaceAfter=4,leading=13),
        "th":      ps("th",     fontName="Helvetica-Bold",fontSize=7.5,textColor=COL_BLANC,alignment=TA_LEFT),
        "th_r":    ps("th_r",   fontName="Helvetica-Bold",fontSize=7.5,textColor=COL_BLANC,alignment=TA_RIGHT),
        "td":      ps("td",     fontName="Helvetica",     fontSize=7.5,textColor=COL_MARINE,alignment=TA_LEFT),
        "td_r":    ps("td_r",   fontName="Helvetica",     fontSize=7.5,textColor=COL_MARINE,alignment=TA_RIGHT),
        "td_g":    ps("td_g",   fontName="Helvetica",     fontSize=7.5,textColor=COL_TEXTE, alignment=TA_LEFT),
        "caption": ps("caption",fontName="Helvetica-Oblique",fontSize=7,textColor=COL_TEXTE,alignment=TA_CENTER,spaceAfter=3),
        "alert":   ps("alert",  fontName="Helvetica-Bold",fontSize=7.5,textColor=COL_ORANGE,alignment=TA_LEFT),
        "pth_l":   ps("pth_l",  fontName="Helvetica-Bold",fontSize=7.5,textColor=COL_BLANC,alignment=TA_LEFT),
        "pth_r":   ps("pth_r",  fontName="Helvetica-Bold",fontSize=7.5,textColor=COL_BLANC,alignment=TA_RIGHT),
        "ptd_l":   ps("ptd_l",  fontName="Helvetica",     fontSize=7.5,textColor=COL_MARINE,alignment=TA_LEFT),
        "ptd":     ps("ptd",    fontName="Helvetica",     fontSize=7.5,textColor=COL_MARINE,alignment=TA_RIGHT),
        "ppos":    ps("ppos",   fontName="Helvetica-Bold",fontSize=7.5,textColor=COL_VERT,  alignment=TA_RIGHT),
        "pneg":    ps("pneg",   fontName="Helvetica-Bold",fontSize=7.5,textColor=COL_ROUGE, alignment=TA_RIGHT),
        "disc":    ps("disc",   fontName="Helvetica-Oblique",fontSize=6.5,textColor=COL_TEXTE,alignment=TA_CENTER,leading=10),
    }

# ---------------------------------------------------------------------------
# EN-TÊTE / PIED DE PAGE — pages 2+
# Fidèle au modèle Canva page 2 : ligne ciel en haut, footer marine
# ---------------------------------------------------------------------------
def _hf(canvas, doc):
    canvas.saveState()
    if doc.page == 1:
        canvas.restoreState(); return
    W=PAGE_W; H=PAGE_H
    # Ligne ciel sous l'en-tête
    canvas.setStrokeColor(COL_CIEL)
    canvas.setLineWidth(1.8)
    canvas.line(0, H-0.40*cm, W, H-0.40*cm)
    # CONFIDENTIEL haut droite
    canvas.setFont("Helvetica",6.5); canvas.setFillColor(COL_MARINE)
    canvas.drawRightString(W-MARGIN_H, H-0.26*cm, "CONFIDENTIEL")
    # Footer marine
    canvas.setFillColor(COL_MARINE)
    canvas.rect(0, 0, W, 0.85*cm, fill=1, stroke=0)
    canvas.setFont("Helvetica",7); canvas.setFillColor(COL_BLEU_MD)
    canvas.drawString(MARGIN_H, 0.28*cm, "Executive Report - Internal Only")
    canvas.drawRightString(W-MARGIN_H, 0.28*cm, "Page {}".format(doc.page))
    canvas.restoreState()

# ---------------------------------------------------------------------------
# SECTIONS
# ---------------------------------------------------------------------------
def _donuts(kpis, aum_by_region, styles):
    els=[]; abt=kpis.get("aum_by_type",{})
    els.append(Paragraph("Repartition AUM Finance", styles["section"]))
    els.append(ColorRect(USABLE_W,2,COL_CIEL)); els.append(Spacer(1,14))
    fw=DONUT_W/72; fh=DONUT_H/72
    b1=_make_donut_png(list(abt.keys()),list(abt.values()),"AUM par Type de Client",fw,fh)
    b2=_make_donut_png(list(aum_by_region.keys()),list(aum_by_region.values()),"AUM par Region Geographique",fw,fh)
    gap=USABLE_W-2*DONUT_W
    tbl=Table([[Image(b1,DONUT_W,DONUT_H),Spacer(gap,1),Image(b2,DONUT_W,DONUT_H)]],colWidths=[DONUT_W,gap,DONUT_W])
    tbl.setStyle(TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"TOP"),
                              ("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),0),
                              ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),0)]))
    els.append(KeepTogether([tbl,Paragraph("Figure 1 — Repartition AUM Finance par Type de Client (gauche) et par Region (droite)",styles["caption"])]))
    return els

def _top10(top_deals, outflows, styles, mode_comex, include_outflows):
    els=[PageBreak()]
    els.append(Paragraph("Top 10 Deals", styles["section"]))
    els.append(ColorRect(USABLE_W,2,COL_CIEL)); els.append(Spacer(1,16))
    els.append(Paragraph("Inflows — AUM Finance (Funded)", styles["subsect"])); els.append(Spacer(1,8))
    n=max(len(top_deals[:10]),1); h_pts=max(180,min(320,n*27+40)); fw=TOP10_W/72; fh=h_pts/72
    b=_make_top10_png(top_deals,mode_comex,fw,fh,"Top 10 Inflows — AUM Finance")
    els.append(KeepTogether([Image(b,TOP10_W,h_pts),Paragraph("Figure 2 — Top 10 Deals par AUM Finance (statut Funded uniquement)",styles["caption"])]))
    if include_outflows and outflows:
        els.append(Spacer(1,22)); els.append(Paragraph("Outflows — AUM Rachete (Redeemed)",styles["subsect"])); els.append(Spacer(1,8))
        n2=max(len(outflows[:10]),1); h2=max(120,min(260,n2*27+40))
        b2=_make_top10_png(outflows,mode_comex,fw,h2/72,"Top 10 Outflows — AUM Rachete")
        els.append(KeepTogether([Image(b2,TOP10_W,h2),Paragraph("Figure 3 — Top 10 Rachats (Redeemed)",styles["caption"])]))
    return els

def _pipeline(pipeline_df, styles, mode_comex):
    els=[PageBreak()]
    els.append(Paragraph("Pipeline Actif - Recapitulatif",styles["section"]))
    els.append(ColorRect(USABLE_W,2,COL_CIEL)); els.append(Spacer(1,10))
    actifs=pipeline_df[pipeline_df["statut"].isin(["Prospect","Initial Pitch","Due Diligence","Soft Commit"])].copy()
    if actifs.empty:
        els.append(Paragraph("Aucun deal actif.",styles["body"]))
    else:
        if mode_comex:
            actifs["nom_client"]=actifs.apply(lambda r:"{} - {}".format(r.get("type_client",""),r.get("region","")),axis=1)
        ratios=[0.22,0.16,0.12,0.13,0.13,0.13,0.11]; cw=[USABLE_W*r for r in ratios]
        headers=["Client","Fonds","Statut","AUM Cible","AUM Revise","Prochaine Action","Commercial"]
        rows=[[Paragraph(h,styles["th"]) for h in headers]]; today=date.today()
        for _,row in actifs.iterrows():
            nad=row.get("next_action_date")
            if isinstance(nad,date):
                nad_str="[!] {}".format(nad.isoformat()) if nad<today else nad.isoformat()
                ns=styles["alert"] if nad<today else styles["td"]
            else: nad_str="-"; ns=styles["td"]
            rows.append([Paragraph(str(row.get("nom_client",""))[:26],styles["td"]),
                         Paragraph(str(row.get("fonds","")),styles["td"]),
                         Paragraph(str(row.get("statut","")),styles["td"]),
                         Paragraph(fmt_aum(float(row.get("target_aum_initial",0) or 0)),styles["td_r"]),
                         Paragraph(fmt_aum(float(row.get("revised_aum",0) or 0)),styles["td_r"]),
                         Paragraph(nad_str,ns),
                         Paragraph(str(row.get("sales_owner",""))[:14],styles["td"])])
        tbl=Table(rows,colWidths=cw,repeatRows=1)
        tbl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),COL_MARINE),("LINEBELOW",(0,0),(-1,0),1.5,COL_CIEL),
                                  ("GRID",(0,0),(-1,-1),0.3,COL_GRIS),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                                  ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
                                  ("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),
                                  ("ROWBACKGROUNDS",(0,1),(-1,-1),[COL_HEADER,COL_BLANC])]))
        els.append(tbl)
    lp=pipeline_df[pipeline_df["statut"].isin(["Lost","Paused"])].copy()
    if not lp.empty:
        els.append(Spacer(1,16)); els.append(Paragraph("Deals Perdus / En Pause",styles["subsect"]))
        els.append(ColorRect(USABLE_W,1,COL_GRIS)); els.append(Spacer(1,8))
        if mode_comex: lp["nom_client"]=lp.apply(lambda r:"{} - {}".format(r.get("type_client",""),r.get("region","")),axis=1)
        lp_cw=[USABLE_W*r for r in [0.26,0.18,0.12,0.22,0.22]]
        lp_rows=[[Paragraph(h,styles["th"]) for h in ["Client","Fonds","Statut","Raison","Concurrent"]]]
        for _,row in lp.iterrows():
            lp_rows.append([Paragraph(str(row.get("nom_client",""))[:24],styles["td_g"]),
                             Paragraph(str(row.get("fonds","")),styles["td_g"]),
                             Paragraph(str(row.get("statut","")),styles["td_g"]),
                             Paragraph(str(row.get("raison_perte","") or "-"),styles["td_g"]),
                             Paragraph(str(row.get("concurrent_choisi","") or "-"),styles["td_g"])])
        lp_tbl=Table(lp_rows,colWidths=lp_cw,repeatRows=1)
        lp_tbl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),HexColor("#7a7a7a")),
                                     ("GRID",(0,0),(-1,-1),0.3,COL_GRIS),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                                     ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
                                     ("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5)]))
        els.append(lp_tbl)
    return els

def _performance(perf_data, nav_df, styles, fonds_perimetre):
    els=[PageBreak()]; pf=perf_data.copy() if perf_data is not None else pd.DataFrame(); nb=nav_df
    if fonds_perimetre:
        if not pf.empty and "Fonds" in pf.columns: pf=pf[pf["Fonds"].isin(fonds_perimetre)]
        if nb is not None and hasattr(nb,"columns"): cols_k=[c for c in nb.columns if c in fonds_perimetre]; nb=nb[cols_k] if cols_k else pd.DataFrame()
    els.append(Paragraph("Analyse de Performance - NAV",styles["section"]))
    els.append(ColorRect(USABLE_W,2,COL_CIEL)); els.append(Spacer(1,14))
    fw=NAV_W/72; fh=NAV_H/72
    buf=_make_nav_png(nb,fw,fh)
    els.append(KeepTogether([Image(buf,NAV_W,NAV_H),Paragraph("Figure — Evolution NAV Base 100",styles["caption"])]))
    els.append(Spacer(1,16))
    if pf.empty: els.append(Paragraph("Aucune donnee de performance.",styles["body"])); return els
    els.append(Paragraph("Tableau des Performances",styles["subsect"]))
    els.append(ColorRect(USABLE_W,1,COL_GRIS)); els.append(Spacer(1,8))
    col_order=["Fonds","NAV Derniere","Base 100 Actuel","Perf 1M (%)","Perf YTD (%)","Perf Periode (%)"]
    avail=[c for c in col_order if c in pf.columns]; n=len(avail)
    if n==0: els.append(Paragraph("Colonnes manquantes.",styles["body"])); return els
    cw=[USABLE_W*(0.28 if i==0 else 0.72/max(n-1,1)) for i in range(n)]
    perf_cols={"Perf 1M (%)","Perf YTD (%)","Perf Periode (%)"}
    hrow=[Paragraph(c,styles["pth_l"] if i==0 else styles["pth_r"]) for i,c in enumerate(avail)]
    rows=[hrow]
    for _,row in pf.iterrows():
        drow=[]
        for i,col in enumerate(avail):
            val=row.get(col)
            if col=="Fonds": drow.append(Paragraph(str(val),styles["ptd_l"]))
            elif col in perf_cols:
                try: fv=float(val); ok=not np.isnan(fv)
                except: ok=False; fv=0.0
                if not ok: drow.append(Paragraph("n.d.",styles["ptd"]))
                else:
                    sg="+" if fv>0 else ""; st2=styles["ppos"] if fv>=0 else styles["pneg"]
                    drow.append(Paragraph("{}{:.2f}%".format(sg,fv),st2))
            else:
                try: fv=float(val); txt="{:.4f}".format(fv) if "NAV" in col else "{:.2f}".format(fv)
                except: txt="-"
                drow.append(Paragraph(txt,styles["ptd"]))
        rows.append(drow)
    ptbl=Table(rows,colWidths=cw,repeatRows=1)
    ptbl.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),COL_MARINE),("LINEBELOW",(0,0),(-1,0),1.2,COL_CIEL),
                               ("GRID",(0,0),(-1,-1),0.3,COL_GRIS),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                               ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
                               ("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
                               ("ROWBACKGROUNDS",(0,1),(-1,-1),[COL_HEADER,COL_BLANC])]))
    els.append(ptbl); return els

# ---------------------------------------------------------------------------
# ENTREE PRINCIPALE
# ---------------------------------------------------------------------------
def generate_pdf(pipeline_df, kpis, aum_by_region=None, mode_comex=False,
                 perf_data=None, nav_base100_df=None, fonds_perimetre=None,
                 include_top10=True, include_outflows=False, include_perf=True):
    aum_by_region=aum_by_region or {}
    outflows=kpis.get("outflows",[])
    if mode_comex:
        pipeline_df=pipeline_df.copy()
        pipeline_df["nom_client"]=pipeline_df.apply(lambda r:"{} - {}".format(r.get("type_client",""),r.get("region","")),axis=1)
        top_deals_safe=[dict(d,nom_client="{} - {}".format(d.get("type_client",""),d.get("region",""))) for d in kpis.get("top_deals",[])]
        outflows_safe=[dict(d,nom_client="{} - {}".format(d.get("type_client",""),d.get("region",""))) for d in outflows]
        kpis=dict(kpis,top_deals=top_deals_safe)
    else:
        top_deals_safe=kpis.get("top_deals",[]); outflows_safe=outflows
    S=_styles()

    # --- Page de garde via canvas direct ---
    cover_buf=io.BytesIO()
    cv=rl_canvas.Canvas(cover_buf,pagesize=A4)
    _draw_cover(cv,mode_comex,kpis,fonds_perimetre)
    cv.showPage(); cv.save(); cover_buf.seek(0)

    # --- Pages 2+ via SimpleDocTemplate ---
    content_buf=io.BytesIO()
    doc=SimpleDocTemplate(content_buf,pagesize=A4,
                          leftMargin=MARGIN_H,rightMargin=MARGIN_H,
                          topMargin=MARGIN_V,bottomMargin=1.2*cm,
                          title="Executive Report",author="AM Division")
    els=[]
    els+=_donuts(kpis,aum_by_region,S)
    if include_top10: els+=_top10(top_deals_safe,outflows_safe,S,mode_comex,include_outflows)
    els+=_pipeline(pipeline_df,S,mode_comex)
    if include_perf and perf_data is not None and hasattr(perf_data,"empty") and not perf_data.empty:
        els+=_performance(perf_data,nav_base100_df,S,fonds_perimetre)
    doc.build(els,onFirstPage=_hf,onLaterPages=_hf)
    content_buf.seek(0)

    # --- Fusion page de garde + contenu ---
    from pypdf import PdfReader, PdfWriter
    writer=PdfWriter()
    writer.add_page(PdfReader(cover_buf).pages[0])
    for page in PdfReader(content_buf).pages: writer.add_page(page)
    out=io.BytesIO(); writer.write(out); out.seek(0)
    return out.getvalue()
