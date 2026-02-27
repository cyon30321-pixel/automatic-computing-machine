"""
SutamMaker v4 — core/graph_engine.py
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Matplotlib-based graph engine for economics diagrams.
No Tkinter / docx dependencies.
"""

import re
import platform

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Korean font setup (called once at import time)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _setup_korean_font():
    try:
        plt.style.use("seaborn-v0_8-whitegrid")
    except Exception:
        pass
    names = {f.name for f in fm.fontManager.ttflist}
    system = platform.system()
    if system == "Windows":
        for c in ("Malgun Gothic", "NanumGothic", "Gulim"):
            if c in names:
                plt.rcParams["font.family"] = c
                break
    elif system == "Darwin":
        plt.rcParams["font.family"] = "AppleGothic"
    else:
        plt.rcParams["font.family"] = "NanumGothic"
    plt.rcParams["axes.unicode_minus"] = False


_setup_korean_font()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Helper: base axes
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _draw_base_axes(ax, xlabel="Q", ylabel="P"):
    ax.grid(True, linestyle="--", color="lightgray", alpha=0.7, zorder=0)
    ax.spines["right"].set_visible(False)
    ax.spines["top"].set_visible(False)
    ax.set_xlabel(xlabel, loc="right", fontsize=11, fontweight="bold", fontstyle="italic")
    ax.set_ylabel(ylabel, loc="top", fontsize=11, fontweight="bold", fontstyle="italic", rotation=0)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Individual graph drawing functions
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _draw_sd_subplot(ax, shift_match):
    _draw_base_axes(ax)
    q = [0, 10]
    ax.plot(q, [0, 10], "k-", lw=2, zorder=3)
    ax.plot(q, [10, 0], "k-", lw=2, zorder=3)
    ax.text(9.5, 9.5, "S", fontsize=12, fontweight="bold")
    ax.text(9.5, 0.5, "D", fontsize=12, fontweight="bold")
    ax.set_xticks([])
    ax.set_yticks([])
    if not shift_match:
        return
    curve = shift_match.group(1).upper()
    direction = shift_match.group(2).upper()
    color_map = {"S": "darkred", "D": "blue"}
    color = color_map.get(curve, "gray")
    offset = 2 if direction == "RIGHT" else -2
    arrow_dx = 2 if direction == "RIGHT" else -2
    if curve == "S":
        new_line = ([offset, 10 + offset], [0, 10])
    else:
        new_line = ([offset, 10 + offset], [10, 0])
    ax.plot(new_line[0], new_line[1], color=color, linestyle="--", lw=2, zorder=3)
    label_x = new_line[0][-1] - 0.5
    label_y = new_line[1][-1] + (0.5 if curve == "S" else -0.5)
    ax.text(label_x, label_y, f"{curve}'", fontsize=12, fontweight="bold", color=color)
    ax.annotate("", xy=(5 + arrow_dx / 2, 5), xytext=(5 - arrow_dx / 2, 5),
                arrowprops=dict(arrowstyle="->", color="gray", lw=1.5))


def _draw_ppf(ax, tag_str):
    _draw_base_axes(ax, xlabel="재화 X", ylabel="재화 Y")
    ax.set_xticks([])
    ax.set_yticks([])
    t = np.linspace(0, 1, 100)
    x = 10 * (1 - t ** 2) ** 0.5
    y = 10 * t
    ax.plot(x, y, "b-", lw=2.5, zorder=3, label="PPF")
    ax.fill_between(x, y, alpha=0.05, color="blue")
    ax.text(x[0] + 0.3, y[0] + 0.3, "A", fontsize=11, fontweight="bold")
    ax.text(x[-1] + 0.3, y[-1] - 0.5, "B", fontsize=11, fontweight="bold")
    shift = re.search(r"PPF_SHIFT\s*=\s*(OUT|IN)", tag_str, re.IGNORECASE)
    if shift:
        factor = 1.3 if shift.group(1).upper() == "OUT" else 0.7
        x2 = factor * 10 * (1 - t ** 2) ** 0.5
        y2 = factor * 10 * t
        ax.plot(x2, y2, "r--", lw=2, zorder=3, label=f"PPF' ({shift.group(1)})")
        mid = len(t) // 2
        ax.annotate("", xy=(x2[mid], y2[mid]), xytext=(x[mid], y[mid]),
                    arrowprops=dict(arrowstyle="->", color="gray", lw=1.5))
    ax.legend(fontsize=9)


def _draw_asad(ax, tag_str):
    _draw_base_axes(ax, xlabel="Y (실질GDP)", ylabel="P (물가)")
    ax.set_xticks([])
    ax.set_yticks([])
    q = [0, 10]
    ax.plot(q, [10, 0], "b-", lw=2, zorder=3)
    ax.text(9.5, 0.5, "AD", fontsize=11, fontweight="bold", color="blue")
    ax.plot(q, [2, 8], "r-", lw=2, zorder=3)
    ax.text(9.5, 8.3, "SRAS", fontsize=10, fontweight="bold", color="red")
    ax.axvline(x=6, color="green", lw=2, linestyle="-", zorder=3)
    ax.text(6.2, 9.5, "LRAS", fontsize=10, fontweight="bold", color="green")
    ad_shift = re.search(r"Shift_AD\s*=\s*(Right|Left)", tag_str, re.IGNORECASE)
    if ad_shift:
        d = ad_shift.group(1).upper()
        offset = 2 if d == "RIGHT" else -2
        ax.plot([offset, 10 + offset], [10, 0], "b--", lw=2, zorder=3)
        ax.text(9.5 + offset, 0.5, "AD'", fontsize=11, fontweight="bold", color="darkblue")
    sras_shift = re.search(r"Shift_SRAS\s*=\s*(Right|Left)", tag_str, re.IGNORECASE)
    if sras_shift:
        d = sras_shift.group(1).upper()
        offset = 2 if d == "RIGHT" else -2
        ax.plot([offset, 10 + offset], [2, 8], "r--", lw=2, zorder=3)
        ax.text(9.5 + offset, 8.3, "SRAS'", fontsize=10, fontweight="bold", color="darkred")


def _draw_lorenz(ax, tag_str):
    _draw_base_axes(ax, xlabel="인구 누적 비율(%)", ylabel="소득 누적 비율(%)")
    x = np.linspace(0, 1, 100)
    ax.plot(x, x, "k--", lw=1.5, label="완전평등선", zorder=3)
    gini = re.search(r"GINI\s*=\s*([0-9.]+)", tag_str, re.IGNORECASE)
    exp = 2.0
    if gini:
        g = float(gini.group(1))
        exp = max(1.1, 1 / (1 - g + 0.01))
    y = x ** exp
    ax.plot(x, y, "b-", lw=2.5, label="로렌츠 곡선", zorder=3)
    ax.fill_between(x, x, y, alpha=0.15, color="blue")
    ax.text(0.6, 0.3, "불평등\n면적", fontsize=9, ha="center", color="blue")
    compare = re.search(r"COMPARE_GINI\s*=\s*([0-9.]+)", tag_str, re.IGNORECASE)
    if compare:
        g2 = float(compare.group(1))
        exp2 = max(1.1, 1 / (1 - g2 + 0.01))
        y2 = x ** exp2
        ax.plot(x, y2, "r--", lw=2, label=f"비교 (Gini={g2})", zorder=3)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.legend(fontsize=8, loc="upper left")


def _draw_phillips(ax, tag_str):
    _draw_base_axes(ax, xlabel="실업률(%)", ylabel="인플레이션(%)")
    ax.set_xticks([])
    ax.set_yticks([])
    x = np.linspace(1, 10, 100)
    y = 8 / x + 0.5
    ax.plot(x, y, "b-", lw=2.5, zorder=3, label="단기 필립스")
    shift = re.search(r"PHILLIPS_SHIFT\s*=\s*(UP|DOWN)", tag_str, re.IGNORECASE)
    if shift:
        d = shift.group(1).upper()
        offset = 2 if d == "UP" else -2
        y2 = 8 / x + 0.5 + offset
        ax.plot(x, y2, "r--", lw=2, zorder=3, label=f"이동 ({d})")
    ax.legend(fontsize=9)


def _draw_tax(ax, tag_str):
    """Tax incidence graph."""
    _draw_base_axes(ax, xlabel="Q", ylabel="P")
    q = np.linspace(0, 10, 100)
    d = 10 - q
    s = q
    tax = 2
    s_tax = q + tax
    ax.plot(q, d, "b-", lw=2, label="D", zorder=3)
    ax.plot(q, s, "r-", lw=2, label="S", zorder=3)
    ax.plot(q, s_tax, "r--", lw=2, label="S + Tax", zorder=3)
    q_eq1, p_eq1 = 5, 5
    q_eq2, p_eq2, p_seller = 4, 6, 4
    ax.plot(q_eq1, p_eq1, "ko", ms=5, zorder=4)
    ax.plot(q_eq2, p_eq2, "ko", ms=5, zorder=4)
    ax.plot([0, q_eq2, q_eq2], [p_eq2, p_eq2, 0], "k:", lw=1, zorder=2)
    ax.plot([0, q_eq2], [p_seller, p_seller], "k:", lw=1, zorder=2)
    ax.fill_between([0, q_eq2], p_seller, p_eq2, alpha=0.1, color="orange")
    ax.text(q_eq2 + 0.15, p_eq2 + 0.3, "E'", fontsize=10, fontweight="bold")
    ax.text(q_eq1 + 0.15, p_eq1 + 0.3, "E", fontsize=10, fontweight="bold")
    ax.text(0.3, (p_eq2 + p_seller) / 2, "Tax\nRevenue", fontsize=8, ha="left", va="center", color="orange")
    ax.legend(fontsize=9, loc="upper right")


def _draw_tariff(ax, tag_str):
    """Tariff graph."""
    _draw_base_axes(ax, xlabel="Q", ylabel="P")
    q = np.linspace(0, 10, 100)
    d = 10 - q
    s = q
    p_w = 3
    tariff = 2
    p_t = p_w + tariff
    ax.plot(q, d, "b-", lw=2, label="D (Domestic)", zorder=3)
    ax.plot(q, s, "r-", lw=2, label="S (Domestic)", zorder=3)
    ax.axhline(p_w, color="green", linestyle="-", lw=1.5, label="Pw (World Price)", zorder=3)
    ax.axhline(p_t, color="purple", linestyle="--", lw=1.5, label="Pw + Tariff", zorder=3)
    qd_w = 10 - p_w
    qs_w = p_w
    qd_t = 10 - p_t
    qs_t = p_t
    ax.fill_between([qs_w, qs_t], p_w, p_t, alpha=0.08, color="red")
    ax.fill_between([qd_t, qd_w], p_w, p_t, alpha=0.08, color="red")
    ax.fill_between([qs_t, qd_t], p_w, p_t, alpha=0.15, color="purple")
    ax.text((qs_t + qd_t) / 2, (p_w + p_t) / 2, "Tariff\nRevenue", fontsize=8,
            ha="center", va="center", color="purple")
    ax.legend(fontsize=8, loc="upper right")


def _draw_externality(ax, tag_str):
    """Negative production externality graph."""
    _draw_base_axes(ax, xlabel="Q", ylabel="P / Cost / Benefit")
    q = np.linspace(0, 10, 100)
    pmb = 10 - q
    pmc = q + 1
    smc = q + 4
    ax.plot(q, pmb, "b-", lw=2, label="PMB = SMB", zorder=3)
    ax.plot(q, pmc, "r-", lw=2, label="PMC", zorder=3)
    ax.plot(q, smc, "r--", lw=2, label="SMC", zorder=3)
    q_market = 4.5
    q_social = 3.0
    ax.fill_between(q, pmc, smc, where=(q >= q_social) & (q <= q_market),
                    color="gray", alpha=0.3, zorder=2)
    ax.text((q_social + q_market) / 2, (pmc[50] + smc[50]) / 2 - 0.3,
            "DWL", fontsize=10, fontweight="bold", ha="center")
    ax.plot([q_market, q_market], [0, 10 - q_market], "k:", lw=1)
    ax.plot([q_social, q_social], [0, 10 - q_social], "k:", lw=1)
    ax.text(q_market, -0.5, "Qm", fontsize=9, ha="center")
    ax.text(q_social, -0.5, "Qs", fontsize=9, ha="center")
    ax.legend(fontsize=8, loc="upper right")


def _draw_indifference(ax, tag_str):
    """Indifference curves + budget line graph."""
    _draw_base_axes(ax, xlabel="Good X", ylabel="Good Y")
    ax.set_xticks([])
    ax.set_yticks([])
    x = np.linspace(0.5, 10, 100)
    u1 = 12 / x
    u2 = 20 / x
    budget = 10 - x
    ax.plot(x, u1, "b-", lw=2, label="U1", zorder=3)
    ax.plot(x, u2, "b--", lw=2, label="U2", zorder=3)
    ax.plot(x[x <= 10], np.maximum(budget[x <= 10], 0), "k-", lw=1.5, label="Budget Line", zorder=3)
    ax.plot(3.46, 3.46, "ro", ms=6, zorder=5)
    ax.text(3.8, 3.6, "Optimal", fontsize=9, fontweight="bold", color="red")
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 11)
    ax.legend(fontsize=9, loc="upper right")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Main dispatcher
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def process_and_draw_graph(tag_str, save_path):
    """Parse the graph tag string and render the appropriate chart to save_path.
    Returns save_path on success, None on failure.
    """
    clean = re.sub(r"[\$\\\[\]\|]", " ", tag_str)
    upper = clean.upper()

    # v4 extended graphs
    if "TAX" in upper and "TARIFF" not in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_tax(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path

    if "TARIFF" in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_tariff(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path

    if "EXTERNALITY" in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_externality(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path

    if "INDIFFERENCE" in upper or ("CURVE" in upper and "PPF" not in upper
                                    and "LORENZ" not in upper and "PHILLIPS" not in upper):
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_indifference(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path

    # v3 graphs
    if "PPF" in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_ppf(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path

    if "ASAD" in upper or "AS_AD" in upper or "AS-AD" in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_asad(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path

    if "LORENZ" in upper or "GINI" in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_lorenz(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path

    if "PHILLIPS" in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_phillips(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path

    if "DUAL" in upper or "LEFT:" in upper:
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(7, 3.5))
        left_m = re.search(r"Left:\s*Shift_([SD])\s*=\s*(Right|Left)", clean, re.IGNORECASE)
        right_m = re.search(r"Right:\s*Shift_([SD])\s*=\s*(Right|Left)", clean, re.IGNORECASE)
        _draw_sd_subplot(ax1, left_m)
        _draw_sd_subplot(ax2, right_m)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path

    # Numeric S-D graph
    qd_m = re.search(r"Qd\s*=\s*([0-9, ]+)", clean, re.IGNORECASE)
    qs_m = re.search(r"Qs\s*=\s*([0-9, ]+)", clean, re.IGNORECASE)
    p_m = re.search(r"P\s*=\s*([0-9, ]+)", clean, re.IGNORECASE)
    shift_s_m = re.search(r"Shift_S\s*=\s*([-0-9]+)", clean, re.IGNORECASE)
    shift_d_m = re.search(r"Shift_D\s*=\s*([-0-9]+)", clean, re.IGNORECASE)
    ceiling_m = re.search(r"Price_Ceiling\s*=\s*([0-9]+)", clean, re.IGNORECASE)
    floor_m = re.search(r"Price_Floor\s*=\s*([0-9]+)", clean, re.IGNORECASE)

    if qd_m and qs_m:
        qd = list(map(int, qd_m.group(1).split(",")))
        qs = list(map(int, qs_m.group(1).split(",")))
        p_vals = list(map(int, p_m.group(1).split(","))) if p_m else [10 * (i + 1) for i in range(len(qd))]

        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_base_axes(ax)
        ax.plot(qd, p_vals, color="blue", lw=2, zorder=3)
        ax.plot(qs, p_vals, color="red", lw=2, zorder=3)
        ax.text(qd[-1], p_vals[-1] - max(p_vals) * 0.05, "D", fontsize=12, color="blue", fontweight="bold")
        ax.text(qs[-1], p_vals[-1] + max(p_vals) * 0.05, "S", fontsize=12, color="red", fontweight="bold")

        if shift_s_m:
            s = int(shift_s_m.group(1))
            sp = [p + s for p in p_vals]
            ax.plot(qs, sp, color="darkred", linestyle="--", lw=2, zorder=3)
            ax.text(qs[-1], sp[-1] + max(p_vals) * 0.05, "S'", fontsize=12, color="darkred", fontweight="bold")
        if shift_d_m:
            s = int(shift_d_m.group(1))
            sp = [p + s for p in p_vals]
            ax.plot(qd, sp, color="darkblue", linestyle="--", lw=2, zorder=3)
            ax.text(qd[-1], sp[-1] + max(p_vals) * 0.05, "D'", fontsize=12, color="darkblue", fontweight="bold")
        if ceiling_m:
            c = int(ceiling_m.group(1))
            ax.axhline(y=c, color="purple", linestyle="-.", lw=2, zorder=4)
            ax.text(min(qd), c + max(p_vals) * 0.03, "Price Ceiling", color="purple", fontweight="bold", fontstyle="italic")
        if floor_m:
            f = int(floor_m.group(1))
            ax.axhline(y=f, color="green", linestyle="-.", lw=2, zorder=4)
            ax.text(min(qd), f + max(p_vals) * 0.03, "Price Floor", color="green", fontweight="bold", fontstyle="italic")

        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path

    return None
