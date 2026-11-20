# -*- coding: utf-8 -*-
"""Общие утилиты лабораторного практикума.

Используется всеми курсовыми модулями (kioe.py, vipe.py): настройка
шрифтов matplotlib, кодирование графиков PNG и книг Excel в base64,
кусочно-линейная аппроксимация с одной точкой излома.
"""

import io
import base64

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams["font.sans-serif"] = ["DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False


# ===========================================================================
#  Общие функции
# ===========================================================================


def fig_to_base64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("ascii")


def wb_to_base64(wb):
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("ascii")


# ===========================================================================
#  Кусочно-линейная аппроксимация с одной точкой излома
# ===========================================================================


def piecewise_linear_fit(x, y, split=None):
    """Кусочно-линейная аппроксимация с одной точкой излома.

    Возвращает коэффициенты двух прямых (светодиодный/лазерный режимы)
    и абсциссу точки их пересечения (пороговый ток I_пор).
    """
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    n = len(x)
    if n < 4:
        return {"error": "Недостаточно точек для аппроксимации (нужно минимум 4)."}
    if n < 6 and split is None:
        # при малом числе точек авто-поиск менее устойчив — оставляем, но предупреждаем
        pass

    def seg_fit(a, b):
        k, b0 = np.polyfit(a, b, 1)
        return float(k), float(b0)

    if split is None:
        best = None
        for s in range(2, n - 1):
            k1, c1 = seg_fit(x[:s], y[:s])
            k2, c2 = seg_fit(x[s:], y[s:])
            sse = (np.sum((np.polyval([k1, c1], x[:s]) - y[:s]) ** 2)
                   + np.sum((np.polyval([k2, c2], x[s:]) - y[s:]) ** 2))
            if best is None or sse < best["sse"]:
                best = {"s": s, "k1": k1, "c1": c1, "k2": k2, "c2": c2, "sse": float(sse)}
        split = best["s"]
        k1, c1, k2, c2 = best["k1"], best["c1"], best["k2"], best["c2"]
    else:
        if split < 2 or split > n - 2:
            return {"error": "Некорректная точка раздела режимов (нужно от 2 до n-2)."}
        k1, c1 = seg_fit(x[:split], y[:split])
        k2, c2 = seg_fit(x[split:], y[split:])

    if abs(k2 - k1) < 1e-12:
        i_por = None
    else:
        i_por = float((c1 - c2) / (k2 - k1))

    return {
        "split": int(split),
        "k_sv": float(k1), "b_sv": float(c1),
        "k_laz": float(k2), "b_laz": float(c2),
        "i_por": i_por,
        "n": int(n),
    }


def fmt_n(v, nd=4):
    """Число для подстановки в «Расчетную часть»: фиксированная точность,
    хвостовые нули убраны; None → «—»."""
    if v is None:
        return "—"
    v = float(v)
    s = f"{v:.{nd}f}".rstrip("0").rstrip(".")
    if s in ("", "-0"):
        s = "0"
    return s


def fmt_g(v):
    """Компактное число для сумм МНК: до 2 знаков для больших, 4 для малых."""
    v = float(v)
    return fmt_n(v, 2) if abs(v) >= 1000 else fmt_n(v, 4)


def work_block(rows, note=None):
    """HTML-блок «Расчетная часть»: таблица «формула = подстановка = ответ».

    rows — [(формула, подстановка, результат), ...]; значения — HTML-строки
    (разрешены <sub>/<sup>/<b>). Блок вставляется внутрь .calc-panel."""
    h = '<p class="work-title">Расчетная часть</p>'
    h += '<div class="work-scroll"><table class="work">'
    h += '<tr><th>№</th><th>Формула</th><th>Подстановка</th><th>Результат</th></tr>'
    for i, (f, s, r) in enumerate(rows, 1):
        h += (f'<tr><td class="wn">{i}</td><td class="weq">{f}</td>'
              f'<td class="weq">= {s}</td><td class="wres">= {r}</td></tr>')
    h += '</table></div>'
    if note:
        h += f'<p class="note">{note}</p>'
    return h
