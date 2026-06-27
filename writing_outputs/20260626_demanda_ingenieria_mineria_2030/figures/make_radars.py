#!/usr/bin/env python3
"""Genera los dos gráficos de radar (spider charts) del informe.
Fig 1: Perfil de riesgo sistémico comparado por país.
Fig 2: Madurez de la demanda de ingeniería minera por país.
Escala 1 (bajo/desfavorable) a 5 (alto/favorable), homogeneizada para que
'mayor = mejor para la firma de ingeniería' en ambos radares.
"""
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import font_manager

plt.rcParams.update({
    "font.size": 9,
    "axes.titlesize": 12,
    "figure.dpi": 300,
    "savefig.dpi": 300,
})

# Paleta accesible (Okabe-Ito subset)
COLORS = {
    "Chile":        "#0072B2",
    "Brasil":       "#009E73",
    "Argentina":    "#56B4E9",
    "Canadá":       "#E69F00",
    "Europa Este":  "#CC79A7",
    "Rusia":        "#D55E00",
    "EE.UU.":       "#000000",
}

def radar(ax, categories, data, title):
    N = len(categories)
    angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
    angles += angles[:1]
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontsize=8)
    ax.set_ylim(0, 5)
    ax.set_yticks([1, 2, 3, 4, 5])
    ax.set_yticklabels(["1", "2", "3", "4", "5"], fontsize=7, color="grey")
    ax.set_title(title, pad=24, weight="bold")
    for label, values in data.items():
        v = values + values[:1]
        ax.plot(angles, v, linewidth=1.6, label=label, color=COLORS[label])
        ax.fill(angles, v, alpha=0.06, color=COLORS[label])

# ---- Fig 1: Perfil de riesgo sistémico (5 = entorno MÁS favorable / menor riesgo) ----
cats_risk = ["Estabilidad\npolítica", "Velocidad\npermisología",
             "Atractivo\ngeológico", "Competitividad\nde costos",
             "Disponibilidad\nde talento", "Estabilidad\nde CAPEX/FX"]
risk = {
    "Chile":       [4.2, 2.6, 4.8, 3.0, 3.4, 3.6],
    "Brasil":      [3.4, 3.0, 4.4, 3.6, 3.2, 3.0],
    "Argentina":   [2.6, 3.2, 4.6, 3.8, 2.8, 1.8],
    "Canadá":      [4.8, 2.2, 4.5, 2.6, 3.8, 4.4],
    "Europa Este": [3.0, 2.8, 3.2, 3.4, 3.0, 3.0],
    "Rusia":       [1.6, 2.4, 4.7, 4.0, 2.6, 1.4],
    "EE.UU.":      [4.0, 2.0, 4.0, 2.4, 3.6, 4.2],
}

# ---- Fig 2: Madurez de la demanda de ingeniería (5 = mayor tracción/sofisticación) ----
cats_mat = ["Pipeline de\nproyectos", "Intensidad\nEPCM", "Demanda de\noptimización",
            "Penetración\nminerales críticos", "Madurez\nlocal de ingeniería",
            "Apertura a\nconsultoría externa"]
mat = {
    "Chile":       [4.6, 4.4, 4.7, 4.2, 4.5, 3.8],
    "Brasil":      [4.0, 3.8, 3.6, 3.8, 3.6, 3.4],
    "Argentina":   [4.2, 3.4, 3.0, 4.6, 2.8, 4.2],
    "Canadá":      [4.0, 4.2, 4.0, 4.4, 4.3, 3.6],
    "Europa Este": [3.4, 3.0, 3.2, 3.6, 2.8, 4.0],
    "Rusia":       [2.6, 3.4, 3.6, 3.2, 3.2, 1.4],
    "EE.UU.":      [4.4, 4.5, 4.4, 4.5, 4.4, 3.0],
}

fig, axes = plt.subplots(1, 2, figsize=(13.5, 6.4), subplot_kw=dict(polar=True))
radar(axes[0], cats_risk, risk,
      "Fig. A — Perfil de entorno / riesgo sistémico\n(5 = entorno más favorable, menor riesgo)")
radar(axes[1], cats_mat, mat,
      "Fig. B — Madurez de la demanda de ingeniería\n(5 = mayor tracción y sofisticación)")
handles, labels = axes[0].get_legend_handles_labels()
fig.legend(handles, labels, loc="lower center", ncol=7, fontsize=8,
           frameon=False, bbox_to_anchor=(0.5, -0.02))
fig.suptitle("Comparativa multivariable de mercados de la gran minería (al 2030)",
             fontsize=14, weight="bold", y=1.02)
fig.tight_layout(rect=[0, 0.04, 1, 1])
fig.savefig("radar_paises.png", bbox_inches="tight", facecolor="white")
print("OK radar_paises.png")
