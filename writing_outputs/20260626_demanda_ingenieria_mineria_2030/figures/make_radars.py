#!/usr/bin/env python3
"""Genera los radares comparativos por país del informe (small multiples con realce).

En lugar de superponer las ocho geografías en un solo radar (ilegible), se produce
un panel por país: el país destacado se dibuja en color y las siete geografías
restantes quedan en gris como referencia ("small multiples" con realce).

Salidas:
  - radar_riesgo_por_pais.png   -> Entorno / riesgo sistémico (6 dimensiones)
  - radar_madurez_por_pais.png  -> Madurez de la demanda de ingeniería (6 dimensiones)

Escala 1 (bajo/desfavorable) a 5 (alto/favorable), homogeneizada para que
'mayor = mejor para la firma de ingeniería' en ambos radares.
"""
import os
import numpy as np
import matplotlib.pyplot as plt

# Escribir siempre junto a este script (figures/), sea cual sea el cwd.
OUTDIR = os.path.dirname(os.path.abspath(__file__))

plt.rcParams.update({
    "font.size": 9,
    "axes.titlesize": 11,
    "figure.dpi": 200,
    "savefig.dpi": 200,
})

# Paleta accesible (Okabe-Ito subset). El orden fija el layout de los paneles.
COLORS = {
    "Chile":        "#0072B2",
    "Perú":         "#882255",
    "Brasil":       "#009E73",
    "Argentina":    "#56B4E9",
    "Canadá":       "#E69F00",
    "Europa Este":  "#CC79A7",
    "Rusia":        "#D55E00",
    "EE.UU.":       "#000000",
}
GREY = "#BBBBBB"  # geografías de contexto

# ---- Datos: Perfil de entorno / riesgo sistémico (5 = MÁS favorable / menor riesgo) ----
cats_risk = ["Estab.\npolítica", "Velocidad\npermisología",
             "Atractivo\ngeológico", "Competitiv.\nde costos",
             "Talento\ntécnico", "Estabilidad\nCAPEX/FX"]
risk = {
    "Chile":       [4.2, 2.6, 4.8, 3.0, 3.4, 3.6],
    "Perú":        [2.4, 2.6, 4.7, 3.6, 3.6, 3.0],
    "Brasil":      [3.4, 3.0, 4.4, 3.6, 3.2, 3.0],
    "Argentina":   [2.6, 3.2, 4.6, 3.8, 2.8, 1.8],
    "Canadá":      [4.8, 2.2, 4.5, 2.6, 3.8, 4.4],
    "Europa Este": [3.0, 2.8, 3.2, 3.4, 3.0, 3.0],
    "Rusia":       [1.6, 2.4, 4.7, 4.0, 2.6, 1.4],
    "EE.UU.":      [4.0, 2.0, 4.0, 2.4, 3.6, 4.2],
}

# ---- Datos: Madurez de la demanda de ingeniería (5 = mayor tracción/sofisticación) ----
cats_mat = ["Pipeline de\nproyectos", "Intensidad\nEPCM", "Demanda\noptimización",
            "Minerales\ncríticos", "Madurez\nlocal",
            "Apertura a\nconsultoría"]
mat = {
    "Chile":       [4.6, 4.4, 4.7, 4.2, 4.5, 3.8],
    "Perú":        [4.5, 4.2, 4.3, 4.2, 4.0, 4.0],
    "Brasil":      [4.0, 3.8, 3.6, 3.8, 3.6, 3.4],
    "Argentina":   [4.2, 3.4, 3.0, 4.6, 2.8, 4.2],
    "Canadá":      [4.0, 4.2, 4.0, 4.4, 4.3, 3.6],
    "Europa Este": [3.4, 3.0, 3.2, 3.6, 2.8, 4.0],
    "Rusia":       [2.6, 3.4, 3.6, 3.2, 3.2, 1.4],
    "EE.UU.":      [4.4, 4.5, 4.4, 4.5, 4.4, 3.0],
}


def _mini_radar(ax, angles, categories, data, target):
    """Un panel: 'target' en color, el resto en gris de referencia."""
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontsize=6.5)
    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels(["1", "2", "3", "4", "5"], fontsize=6, color="0.55")
    ax.tick_params(axis="x", pad=-2)
    ax.grid(color="0.85", linewidth=0.5)
    # 1) resto de países en gris
    for label, values in data.items():
        if label == target:
            continue
        v = values + values[:1]
        ax.plot(angles, v, linewidth=0.7, color=GREY, alpha=0.75, zorder=1)
    # 2) país destacado en color, encima
    v = data[target] + data[target][:1]
    ax.plot(angles, v, linewidth=2.2, color=COLORS[target], zorder=3)
    ax.fill(angles, v, alpha=0.18, color=COLORS[target], zorder=2)
    ax.set_title(target, pad=10, weight="bold", color=COLORS[target], fontsize=11)


def build_panel(categories, data, fname, suptitle):
    N = len(categories)
    angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
    angles += angles[:1]
    countries = list(data.keys())  # 8 -> rejilla 2 x 4
    fig, axes = plt.subplots(2, 4, figsize=(15, 8.4),
                             subplot_kw=dict(polar=True))
    for ax, country in zip(axes.ravel(), countries):
        _mini_radar(ax, angles, categories, data, country)
    fig.suptitle(suptitle, fontsize=14, weight="bold", y=1.005)
    fig.text(0.5, 0.015,
             "En cada panel: color = país destacado; gris = las otras siete geografías "
             "(referencia). Escala 1–5, mayor = más favorable para la firma de ingeniería.",
             ha="center", fontsize=8.5, color="0.35")
    fig.tight_layout(rect=[0, 0.03, 1, 0.98])
    fig.subplots_adjust(hspace=0.45, wspace=0.35)
    outpath = os.path.join(OUTDIR, fname)
    fig.savefig(outpath, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("OK", outpath)


if __name__ == "__main__":
    build_panel(cats_risk, risk, "radar_riesgo_por_pais.png",
                "Radares por país — Entorno / riesgo sistémico (al 2030)")
    build_panel(cats_mat, mat, "radar_madurez_por_pais.png",
                "Radares por país — Madurez de la demanda de ingeniería (al 2030)")
