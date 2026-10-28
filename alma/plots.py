"""Colour-coded matrix figures (red = unpublished ... green = >= high threshold)."""
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402
from matplotlib.colors import LinearSegmentedColormap  # noqa: E402

ALMA_CMAP = LinearSegmentedColormap.from_list("alma", ["#F8696B", "#FFEB84", "#63BE7B"])


def heatmap(df, path, title="", vmax=20, annotate=None, fmt="{:.0f}", max_labels=80):
    """df: rows x cols numbers.  vmax: value mapped to full green (high threshold)."""
    nr, nc = df.shape
    annotate = (nr * nc <= 600) if annotate is None else annotate
    w = min(24, 2.5 + 0.45 * nc + max(len(str(r)) for r in df.index) * 0.07)
    h = min(30, 1.8 + 0.28 * nr)
    fig, ax = plt.subplots(figsize=(max(5, w), max(3, h)))
    ax.imshow(np.clip(df.values.astype(float), 0, vmax), cmap=ALMA_CMAP, vmin=0, vmax=vmax, aspect="auto")
    if nc <= max_labels:
        ax.set_xticks(range(nc))
        ax.set_xticklabels([str(c) for c in df.columns], rotation=60, ha="right", fontsize=8)
    else:
        ax.set_xticks([])
    if nr <= max_labels * 2:
        ax.set_yticks(range(nr))
        ax.set_yticklabels([str(r) for r in df.index], fontsize=7 if nr > 40 else 8)
    else:
        ax.set_yticks([])
    if annotate:
        for i in range(nr):
            for j in range(nc):
                ax.text(j, i, fmt.format(df.iat[i, j]), ha="center", va="center", fontsize=6)
    ax.set_title(title, fontsize=10)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)


def lines(series_by_label, path, title="", ylabel="NOP"):
    fig, ax = plt.subplots(figsize=(7, 4))
    for label, s in series_by_label.items():
        ax.plot([int(x) for x in s.index], s.values, marker="o", label=label)
    ax.set_xlabel("year")
    ax.set_ylabel(ylabel)
    ax.set_title(title, fontsize=10)
    ax.legend(fontsize=7)
    fig.tight_layout()
    fig.savefig(path, dpi=150)
    plt.close(fig)
