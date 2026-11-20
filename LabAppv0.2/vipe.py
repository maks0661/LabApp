# -*- coding: utf-8 -*-
"""Курс «Вакуумная и плазменная электроника» — лабораторные работы.

Лабораторная работа № 1 — «Исследование термоэлектронной эмиссии»:
семейство вольтамперных характеристик вакуумного диода при разных токах
накала, определение токов эмиссии (точка перегиба ВАХ), плотности тока
j = I_Э/S, температуры катода по удельной мощности накала
P_уд = I_н·U_н/(π·l_к·d_к), прямая Ричардсона ln(j/T²) = f(1/T)
(константы A и eφ), прямые Шоттки ln j = f(√E) (температура по наклону),
зависимость j = f(T).

Лабораторные работы № 2 («Исследование электронно-лучевой трубки
с магнитным управлением») и № 3 («Исследование фотоэлектронных приборов»)
— заглушки.

Маршруты: /vipe/lab/<n>, /api/vipe/tee/process.
"""

import math

import numpy as np
import matplotlib.pyplot as plt
from flask import Blueprint, request, jsonify
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.utils import get_column_letter

from common import fig_to_base64, wb_to_base64, piecewise_linear_fit

bp = Blueprint("vipe", __name__)

# --- физические константы ---
K_BOLTZ = 1.380649e-23        # постоянная Больцмана, Дж/К
E_CHARGE = 1.602176634e-19    # заряд электрона, Кл
EPS0 = 8.8541878128e-12       # диэлектрическая постоянная, Ф/м
SIGMA_SB = 5.67e-8            # постоянная Стефана-Больцмана, Вт/(м²·К⁴)
A0_THEOR = 1.2e6              # теоретическая константа Ричардсона, А/(м²·К²)

# Константа наклона прямой Шоттки B = (1/kT)·sqrt(e³/(4πε₀)); T = (B·k⁻¹·sqrt(...))⁻¹
C_SCHOTTKY = math.sqrt(E_CHARGE ** 3 / (4 * math.pi * EPS0)) / K_BOLTZ  # К·√м/В


def tungsten_eps(T):
    """Интегральная степень черноты вольфрама (модельная аппроксимация)."""
    if T < 2000:
        return 0.24
    if T > 3000:
        return 0.33
    return 0.25 + 0.07 * (T - 2000.0) / 1000.0


def tungsten_T_from_P(p_wcm2):
    """Температура катода (К) по удельной мощности накала P (Вт/см²).

    Модельная калибровка для вольфрама: P = ε(T)·σ·T⁴ (лучистый теплообмен
    доминирует над теплопроводностью вводов). Решается обратная задача
    делением отрезка; вне диапазона 500-3600 К значение ограничивается.
    """
    lo, hi = 500.0, 3600.0

    def pud(T):
        return tungsten_eps(T) * SIGMA_SB * T ** 4 / 1e4  # Вт/см²

    if p_wcm2 <= pud(lo):
        return lo
    if p_wcm2 >= pud(hi):
        return hi
    for _ in range(80):
        mid = 0.5 * (lo + hi)
        if pud(mid) < p_wcm2:
            lo = mid
        else:
            hi = mid
    return 0.5 * (lo + hi)


# ===========================================================================
#  HTML-страницы лабораторных работ
# ===========================================================================

LAB_TEE_HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Лабораторная работа № 1 — Исследование термоэлектронной эмиссии</title>
<style>
  :root { --accent:#1f4e79; --bg:#f4f6f9; --panel:#fff; --border:#d8dee9; --ok:#2e7d32; --warn:#c62828; }
  * { box-sizing:border-box; }
  body { margin:0; font-family:'Segoe UI', system-ui, Arial, sans-serif; background:var(--bg); color:#222; line-height:1.5; }
  header { background:var(--accent); color:#fff; padding:22px 20px; }
  header .wrap { max-width:1200px; margin:0 auto; display:flex; justify-content:space-between; align-items:center; gap:16px; flex-wrap:wrap; }
  header h1 { margin:0; font-size:19px; }
  header .crs { font-size:12.5px; opacity:.85; display:block; margin-top:4px; }
  header a { color:#fff; text-decoration:none; font-size:13.5px; border:1px solid rgba(255,255,255,.5); padding:6px 14px; border-radius:6px; }
  header a:hover { background:rgba(255,255,255,.15); }
  main { max-width:1240px; margin:0 auto; padding:24px 16px 60px; }
  .lead { font-size:14.5px; color:#444; background:#fff; border:1px solid var(--border); border-left:4px solid var(--accent); padding:14px 18px; border-radius:6px; margin-bottom:22px; }
  .lead b { color:var(--accent); }
  h2.sec { font-size:19px; color:var(--accent); border-bottom:2px solid var(--accent); padding-bottom:6px; margin:30px 0 14px; }
  .params { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:16px 18px; margin-bottom:22px; }
  .params-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(190px,1fr)); gap:14px 18px; align-items:end; }
  .params-grid label { font-weight:600; font-size:13.5px; display:flex; flex-direction:column; gap:5px; }
  .params-grid input { font-size:15px; padding:7px 10px; border:1px solid #b9c2d0; border-radius:5px; width:100%; }
  .params .hint { color:#777; font-size:12.5px; margin-top:10px; }
  .toolbar { display:flex; gap:10px; flex-wrap:wrap; margin:14px 0; align-items:center; }
  .toolbar .mini { display:flex; align-items:center; gap:8px; font-size:13.5px; color:#444; font-weight:600; }
  .toolbar .mini input { width:70px; font-size:14px; padding:6px 8px; border:1px solid #b9c2d0; border-radius:5px; }
  button { font-family:inherit; font-size:14px; padding:9px 16px; border:none; border-radius:6px; cursor:pointer; background:var(--accent); color:#fff; font-weight:600; transition:background .12s, transform .05s; }
  button:hover { background:#173d61; }
  button:active { transform:translateY(1px); }
  button.ghost { background:#fff; color:var(--accent); border:1px solid var(--accent); }
  button.ghost:hover { background:#eef3f9; }
  .heats-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(300px,1fr)); gap:16px; }
  .heat { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:10px 12px 14px; }
  .heat-title { font-weight:600; font-size:13.5px; color:var(--accent); margin:2px 0 8px; display:flex; gap:8px; flex-wrap:wrap; align-items:center; }
  .heat-title .lbl { color:#555; font-size:12.5px; }
  .heat-title input { width:76px; font-size:13.5px; padding:5px 7px; border:1px solid #b9c2d0; border-radius:4px; }
  .heat table { border-collapse:collapse; width:100%; font-size:12.5px; }
  .heat th, .heat td { border:1px solid var(--border); padding:4px 6px; text-align:center; }
  .heat th { background:#f0f4f9; font-weight:600; }
  .heat td.num-col { color:#888; width:34px; background:#fafbfc; }
  .heat input { width:100%; border:1px solid transparent; background:transparent; text-align:center; font-size:12.5px; padding:4px 3px; border-radius:4px; font-family:inherit; }
  .heat input:focus { border-color:var(--accent); background:#fff; outline:none; }
  .calc-panel { background:#fffbea; border:1px solid #e6d98a; border-left:4px solid #d4a72c; border-radius:6px; padding:14px 18px; margin:14px 0; display:none; }
  .calc-panel .formula { font-weight:600; margin:0 0 6px; color:#6b5400; }
  .calc-panel .formula-eq { font-family:'Cambria Math', 'Times New Roman', serif; font-size:16px; margin:6px 0; }
  .calc-panel .note { color:#777; font-size:13px; margin:8px 0 0; }
  .result-panel { background:#eef7ee; border:1px solid #bfe3bf; border-left:4px solid var(--ok); border-radius:6px; padding:14px 18px; margin:14px 0; display:none; }
  .result-panel h4 { margin:0 0 10px; color:var(--ok); font-size:15px; }
  table.res { border-collapse:collapse; font-size:12.5px; margin:6px 0; }
  table.res th, table.res td { border:1px solid #cfe6cf; padding:5px 8px; text-align:left; }
  table.res th { background:#e3f3e3; }
  .plot-wrap { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:14px; margin:14px 0; display:none; text-align:center; }
  .plot-wrap .ptitle { font-weight:600; font-size:14px; margin-bottom:6px; color:#333; }
  .plot-wrap img { max-width:100%; height:auto; border:1px solid #e0e0e0; border-radius:4px; }
  .plot-wrap .dl { display:inline-block; margin-top:8px; font-size:13px; color:var(--accent); }
  .plot-wrap .dl-stack { margin-top:4px; }
  .muted { color:#777; font-size:13px; }
  sub, sup { font-size:75%; }
</style>
</head>
<body>
<header>
  <div class="wrap">
    <div>
      <h1>Лабораторная работа № 1. Исследование термоэлектронной эмиссии</h1>
      <span class="crs">Курс «Вакуумная и плазменная электроника»</span>
    </div>
    <a href="/" target="_blank">&larr; к списку работ</a>
  </div>
</header>
<main>
  <div class="lead">
    Исследуется вакуумный диод с вольфрамовым катодом и охранными кольцами.
    Снимается <b>семейство вольтамперных характеристик</b> I<sub>а</sub>(U<sub>а</sub>)
    при 4–5 значениях тока накала I<sub>н</sub> (и напряжения накала U<sub>н</sub>).
    По точке перегиба каждой ВАХ определяется <b>ток эмиссии I<sub>Э</sub></b> и
    плотность тока j = I<sub>Э</sub>/S (S = π·d<sub>к</sub>·l<sub>а</sub> — поверхность
    катода внутри анода). Температура катода T находится по удельной мощности накала
    P = I<sub>н</sub>·U<sub>н</sub>/(π·l<sub>к</sub>·d<sub>к</sub>). Строятся
    <b>прямая Ричардсона</b> ln(j/T²) = f(1/T) — константы A и e&phi;, и
    <b>прямые Шоттки</b> ln j = f(√E) для области насыщения — температура по наклону.
    Напряжённость поля у катода коаксиального диода: E = U<sub>а</sub>/(r<sub>к</sub>·ln(r<sub>а</sub>/r<sub>к</sub>)).
  </div>

  <h2 class="sec">Паспортные данные диода</h2>
  <div class="params">
    <div class="params-grid">
      <label>Диаметр катода d<sub>к</sub>, мм
        <input type="number" id="p-dk" step="any" min="0.0001" value="0.5">
      </label>
      <label>Длина катода l<sub>к</sub>, мм
        <input type="number" id="p-lk" step="any" min="0.0001" value="20">
      </label>
      <label>Длина анода l<sub>а</sub>, мм
        <input type="number" id="p-la" step="any" min="0.0001" value="15">
      </label>
      <label>Радиус катода r<sub>к</sub>, мм
        <input type="number" id="p-rk" step="any" min="0.0001" value="0.25">
      </label>
      <label>Радиус анода r<sub>а</sub>, мм
        <input type="number" id="p-ra" step="any" min="0.0001" value="5">
      </label>
    </div>
    <p class="hint">Значения по умолчанию — типовые; замените их паспортными данными исследуемого диода. Температура катода рассчитывается по модельной калибровке вольфрама P = ε(T)·σ·T⁴ (ε = 0,25 при 2000 К … 0,32 при 3000 К); при наличии графика T(P) из приложения к методичке значения могут несколько отличаться.</p>
  </div>

  <h2 class="sec">Вольтамперные характеристики (Таблица 3.1)</h2>
  <div class="toolbar">
    <span class="mini">Накалов: <input type="number" id="n-heats" min="2" max="5" value="5"></span>
    <button class="ghost" onclick="buildHeats()">Изменить число накалов</button>
    <span class="mini">Строк в таблицах: <input type="number" id="n-rows" min="4" max="60" value="21"></span>
    <button class="ghost" onclick="rebuildRows()">Изменить число строк</button>
    <button class="ghost" onclick="fillDemo()">Заполнить демонстрационными данными</button>
  </div>
  <div class="heats-grid" id="heats"></div>

  <div class="toolbar">
    <button onclick="processAll()">Рассчитать и построить графики</button>
  </div>

  <div class="calc-panel" id="calc-panel"></div>
  <div class="result-panel" id="res-panel"></div>

  <div class="plot-wrap" id="plot1">
    <div class="ptitle">График 1. Семейство вольтамперных характеристик диода</div>
    <img id="plot1-img" alt="График 1">
    <div class="dl-stack"><a class="dl" id="plot1-dl-png" download="VAX_family.png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot1-dl-xlsx" download="VAX_family.xlsx">Скачать график EXCEL</a></div>
  </div>
  <div class="plot-wrap" id="plot2">
    <div class="ptitle">График 2. Прямая Ричардсона ln(j/T²) = f(1/T)</div>
    <img id="plot2-img" alt="График 2">
    <div class="dl-stack"><a class="dl" id="plot2-dl-png" download="Richardson_line.png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot2-dl-xlsx" download="Richardson_line.xlsx">Скачать график EXCEL</a></div>
  </div>
  <div class="plot-wrap" id="plot3">
    <div class="ptitle">График 3. Прямые Шоттки ln j = f(√E) в области насыщения</div>
    <img id="plot3-img" alt="График 3">
    <div class="dl-stack"><a class="dl" id="plot3-dl-png" download="Schottky_lines.png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot3-dl-xlsx" download="Schottky_lines.xlsx">Скачать график EXCEL</a></div>
  </div>
  <div class="plot-wrap" id="plot4">
    <div class="ptitle">График 4. Зависимость плотности тока эмиссии от температуры катода j = f(T)</div>
    <img id="plot4-img" alt="График 4">
    <div class="dl-stack"><a class="dl" id="plot4-dl-png" download="j_of_T.png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot4-dl-xlsx" download="j_of_T.xlsx">Скачать график EXCEL</a></div>
  </div>

  <p class="muted">Совет: заполните I<sub>н</sub>, U<sub>н</sub> и таблицы (U<sub>а</sub>, I<sub>а</sub>) для каждого накала (или нажмите «Заполнить демонстрационными данными»), затем «Рассчитать и построить графики». Точка перегиба ВАХ (ток эмиссии) определяется кусочно-линейной аппроксимацией; для прямых Шоттки используются точки ВАХ правее перегиба.</p>
</main>

<script>
const HEAT_COLORS = ['#1f4e79', '#c62828', '#2e7d32', '#6a1b9a', '#ef6c00'];

// Демонстрационные данные: вольфрамовый катод (d=0.5 мм, l=20 мм, l_а=15 мм,
// r_к=0.25 мм, r_а=5 мм), j = A·T²·exp(-eφ/kT) при eφ=4.54 эВ, A=6·10⁵ А/(м²·К²),
// рост насыщения по Шоттки. Сетка U_а = 0…60 В с шагом 3 В.
const DEMO = [
  { i_n: 0.45, u_n: 19.8, ia: [0.0, 0.099, 0.28, 0.515, 0.793, 0.796, 0.798, 0.8, 0.802, 0.804, 0.805, 0.807, 0.809, 0.81, 0.812, 0.813, 0.814, 0.816, 0.817, 0.818, 0.819] },
  { i_n: 0.5, u_n: 22.0, ia: [0.0, 0.243, 0.689, 1.265, 1.948, 2.722, 2.73, 2.736, 2.743, 2.749, 2.754, 2.76, 2.765, 2.77, 2.775, 2.779, 2.784, 2.788, 2.792, 2.797, 2.801] },
  { i_n: 0.55, u_n: 24.6, ia: [0.0, 0.573, 1.622, 2.98, 4.588, 6.412, 8.428, 8.448, 8.467, 8.485, 8.501, 8.517, 8.532, 8.547, 8.561, 8.575, 8.588, 8.601, 8.613, 8.625, 8.637] },
  { i_n: 0.6, u_n: 27.4, ia: [0.0, 1.2, 3.395, 6.236, 9.602, 13.419, 17.639, 22.228, 23.868, 23.915, 23.96, 24.003, 24.045, 24.084, 24.122, 24.159, 24.195, 24.229, 24.263, 24.296, 24.328] },
  { i_n: 0.65, u_n: 30.5, ia: [0.0, 2.439, 6.899, 12.674, 19.513, 27.27, 35.847, 45.173, 55.19, 62.27, 62.382, 62.49, 62.593, 62.691, 62.787, 62.879, 62.968, 63.054, 63.138, 63.22, 63.3] }
];
const DEMO_UA = Array.from({length: 21}, (_, k) => 3 * k);

function num(id) {
  const v = parseFloat(document.getElementById(id).value);
  return isNaN(v) ? null : v;
}

function heatRow(idx) {
  return '<tr>' +
    '<td class="num-col">' + idx + '</td>' +
    '<td><input type="number" step="any" class="ua" placeholder="—"></td>' +
    '<td><input type="number" step="any" class="ia" placeholder="—"></td>' +
    '</tr>';
}

function buildHeatBlock(idx) {
  const div = document.createElement('div');
  div.className = 'heat';
  div.innerHTML =
    '<div class="heat-title">Накал № ' + idx +
    ' &nbsp;<span class="lbl">I_н, А:</span><input type="number" step="any" class="in" placeholder="—">' +
    '&nbsp;<span class="lbl">U_н, В:</span><input type="number" step="any" class="un" placeholder="—">' +
    '</div>' +
    '<table><thead><tr><th>№</th><th>U<sub>а</sub>, В</th><th>I<sub>а</sub>, мА</th></tr></thead>' +
    '<tbody></tbody></table>';
  return div;
}

function buildHeats() {
  const nH = Math.max(2, Math.min(5, parseInt(document.getElementById('n-heats').value) || 5));
  const nR = Math.max(4, Math.min(60, parseInt(document.getElementById('n-rows').value) || 21));
  const box = document.getElementById('heats');
  box.innerHTML = '';
  for (let h = 1; h <= nH; h++) {
    const div = buildHeatBlock(h);
    const tb = div.querySelector('tbody');
    for (let i = 1; i <= nR; i++) tb.insertAdjacentHTML('beforeend', heatRow(i));
    box.appendChild(div);
  }
}

function rebuildRows() {
  const nH = document.querySelectorAll('.heat').length;
  if (nH < 1) { buildHeats(); return; }
  document.getElementById('n-heats').value = nH;
  buildHeats();
}

function fillDemo() {
  document.getElementById('p-dk').value = 0.5;
  document.getElementById('p-lk').value = 20;
  document.getElementById('p-la').value = 15;
  document.getElementById('p-rk').value = 0.25;
  document.getElementById('p-ra').value = 5;
  document.getElementById('n-heats').value = DEMO.length;
  document.getElementById('n-rows').value = DEMO_UA.length;
  buildHeats();
  const blocks = document.querySelectorAll('.heat');
  blocks.forEach((h, i) => {
    h.querySelector('.in').value = DEMO[i].i_n;
    h.querySelector('.un').value = DEMO[i].u_n;
    const rows = h.querySelectorAll('tbody tr');
    rows.forEach((tr, k) => {
      tr.querySelector('.ua').value = DEMO_UA[k];
      tr.querySelector('.ia').value = DEMO[i].ia[k];
    });
  });
}

async function processAll() {
  const p = { d_k: num('p-dk'), l_k: num('p-lk'), l_a: num('p-la'), r_k: num('p-rk'), r_a: num('p-ra') };
  for (const [k, name] of [['d_k', 'd_к'], ['l_k', 'l_к'], ['l_a', 'l_а'], ['r_k', 'r_к'], ['r_a', 'r_а']]) {
    if (p[k] === null || p[k] <= 0) { alert('Введите положительное значение ' + name + ' (мм).'); return; }
  }
  const heats = [];
  let anyRows = false;
  document.querySelectorAll('.heat').forEach(h => {
    const rows = [...h.querySelectorAll('tbody tr')].map(tr =>
      [tr.querySelector('.ua').value, tr.querySelector('.ia').value]);
    if (rows.some(r => r[0] !== '' && r[1] !== '')) anyRows = true;
    heats.push({ i_n: h.querySelector('.in').value, u_n: h.querySelector('.un').value, rows: rows });
  });
  if (!anyRows) { alert('Заполните хотя бы одну таблицу накала (U_а, I_а).'); return; }
  let res;
  try {
    res = await fetch('/api/vipe/tee/process', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(Object.assign(p, { heats: heats }))
    });
  } catch (e) { alert('Ошибка сети: ' + e); return; }
  const data = await res.json();
  if (data.error) { alert(data.error); return; }
  const cp = document.getElementById('calc-panel');
  cp.innerHTML = '<p class="formula">Формулы расчёта:</p>' +
    '<p class="formula-eq">P = I<sub>н</sub>·U<sub>н</sub> / (π·l<sub>к</sub>·d<sub>к</sub>) &nbsp;— удельная мощность накала, Вт/см² (I<sub>н</sub> в А, U<sub>н</sub> в В, l<sub>к</sub>, d<sub>к</sub> в см)</p>' +
    '<p class="formula-eq">j = I<sub>Э</sub> / S, &nbsp; S = π·d<sub>к</sub>·l<sub>а</sub> &nbsp;— плотность тока эмиссии</p>' +
    '<p class="formula-eq">ln(j/T²) = ln A − (e&phi;/k)·(1/T) &nbsp;— прямая Ричардсона</p>' +
    '<p class="formula-eq">ln j = ln j<sub>0</sub> + B·√E, &nbsp; B = (1/kT)·(e³/4πε₀)<sup>1/2</sup> &nbsp;— прямая Шоттки, T = (e³/4πε₀)<sup>1/2</sup>/(k·B)</p>' +
    '<p class="formula-eq">E = U<sub>а</sub> / (r<sub>к</sub>·ln(r<sub>а</sub>/r<sub>к</sub>)) &nbsp;— поле у катода коаксиального диода</p>';
  cp.style.display = 'block';
  const rp = document.getElementById('res-panel');
  rp.innerHTML = data.analysis_html;
  rp.style.display = 'block';
  for (const g of ['g1', 'g2', 'g3', 'g4']) {
    const idx = g.slice(1);
    const img = data[g + '_image'];
    const xlsx = data[g + '_xlsx'];
    const wrap = document.getElementById('plot' + idx);
    if (img) {
      const src = 'data:image/png;base64,' + img;
      document.getElementById('plot' + idx + '-img').src = src;
      document.getElementById('plot' + idx + '-dl-png').href = src;
    }
    if (xlsx) document.getElementById('plot' + idx + '-dl-xlsx').href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + xlsx;
    wrap.style.display = img ? 'block' : 'none';
  }
  rp.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

buildHeats();
</script>
</body>
</html>
"""


LAB_ELT_HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Лабораторная работа № 2 — Исследование электронно-лучевой трубки с магнитным управлением</title>
<style>
  :root { --accent:#1f4e79; --bg:#f4f6f9; --panel:#fff; --border:#d8dee9; --ok:#2e7d32; }
  * { box-sizing:border-box; }
  body { margin:0; font-family:'Segoe UI', system-ui, Arial, sans-serif; background:var(--bg); color:#222; line-height:1.5; }
  header { background:var(--accent); color:#fff; padding:22px 20px; }
  header .wrap { max-width:1240px; margin:0 auto; display:flex; justify-content:space-between; align-items:center; gap:16px; flex-wrap:wrap; }
  header h1 { margin:0; font-size:19px; }
  header .crs { font-size:12.5px; opacity:.85; display:block; margin-top:4px; }
  header a { color:#fff; text-decoration:none; font-size:13.5px; border:1px solid rgba(255,255,255,.5); padding:6px 14px; border-radius:6px; }
  header a:hover { background:rgba(255,255,255,.15); }
  main { max-width:1240px; margin:0 auto; padding:24px 16px 60px; }
  .lead { font-size:14.5px; color:#444; background:#fff; border:1px solid var(--border); border-left:4px solid var(--accent); padding:14px 18px; border-radius:6px; margin-bottom:22px; }
  .lead b { color:var(--accent); }
  h2.sec { font-size:19px; color:var(--accent); border-bottom:2px solid var(--accent); padding-bottom:6px; margin:30px 0 14px; }
  .params { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:16px 18px; margin-bottom:22px; }
  .params-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(200px,1fr)); gap:14px 18px; align-items:end; }
  .params-grid label { font-weight:600; font-size:13px; display:flex; flex-direction:column; gap:5px; }
  .params-grid input { font-size:15px; padding:7px 10px; border:1px solid #b9c2d0; border-radius:5px; width:100%; }
  .params .hint { color:#777; font-size:12.5px; margin-top:10px; }
  .toolbar { display:flex; gap:10px; flex-wrap:wrap; margin:14px 0; align-items:center; }
  .toolbar .mini { display:flex; align-items:center; gap:8px; font-size:13.5px; color:#444; font-weight:600; }
  .toolbar .mini input { width:70px; font-size:14px; padding:6px 8px; border:1px solid #b9c2d0; border-radius:5px; }
  button { font-family:inherit; font-size:14px; padding:9px 16px; border:none; border-radius:6px; cursor:pointer; background:var(--accent); color:#fff; font-weight:600; transition:background .12s, transform .05s; }
  button:hover { background:#173d61; }
  button:active { transform:translateY(1px); }
  button.ghost { background:#fff; color:var(--accent); border:1px solid var(--accent); }
  button.ghost:hover { background:#eef3f9; }
  .mods-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(280px,1fr)); gap:16px; }
  .modbox { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:10px 12px 14px; }
  .modbox .ttl { font-weight:600; font-size:13.5px; color:var(--accent); margin:2px 0 8px; text-align:center; }
  .data-wrap { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:10px 14px 14px; overflow-x:auto; }
  table.data { border-collapse:collapse; font-size:13px; }
  table.data th, table.data td { border:1px solid var(--border); padding:5px 8px; text-align:center; }
  table.data th { background:#f0f4f9; font-weight:600; }
  td.num-col { color:#888; width:36px; background:#fafbfc; }
  table.data input { width:100%; min-width:78px; border:1px solid transparent; background:transparent; text-align:center; font-size:13px; padding:4px 4px; border-radius:4px; font-family:inherit; }
  table.data input:focus { border-color:var(--accent); background:#fff; outline:none; }
  .two-cols { display:grid; grid-template-columns:repeat(auto-fit, minmax(300px,1fr)); gap:16px; }
  .calc-panel { background:#fffbea; border:1px solid #e6d98a; border-left:4px solid #d4a72c; border-radius:6px; padding:14px 18px; margin:14px 0; display:none; }
  .calc-panel .formula { font-weight:600; margin:0 0 6px; color:#6b5400; }
  .calc-panel .formula-eq { font-family:'Cambria Math', 'Times New Roman', serif; font-size:16px; margin:6px 0; }
  .result-panel { background:#eef7ee; border:1px solid #bfe3bf; border-left:4px solid var(--ok); border-radius:6px; padding:14px 18px; margin:14px 0; display:none; }
  .result-panel h4 { margin:0 0 10px; color:var(--ok); font-size:15px; }
  table.res { border-collapse:collapse; font-size:12.5px; margin:6px 0; }
  table.res th, table.res td { border:1px solid #cfe6cf; padding:5px 8px; text-align:left; }
  table.res th { background:#e3f3e3; }
  .plot-wrap { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:14px; margin:14px 0; display:none; text-align:center; }
  .plot-wrap .ptitle { font-weight:600; font-size:14px; margin-bottom:6px; color:#333; }
  .plot-wrap img { max-width:100%; height:auto; border:1px solid #e0e0e0; border-radius:4px; }
  .plot-wrap .dl { display:inline-block; margin-top:8px; font-size:13px; color:var(--accent); }
  .plot-wrap .dl-stack { margin-top:4px; }
  .muted { color:#777; font-size:13px; }
  .err { color:#c62828; }
  sub, sup { font-size:75%; }
</style>
</head>
<body>
<header>
  <div class="wrap">
    <div>
      <h1>Лабораторная работа № 2. Исследование электронно-лучевой трубки с магнитным управлением</h1>
      <span class="crs">Курс «Вакуумная и плазменная электроника»</span>
    </div>
    <a href="/" target="_blank">&larr; к списку работ</a>
  </div>
</header>
<main>
  <div class="lead">
    Исследуется электронно-лучевая трубка с <b>электромагнитной фокусировкой</b>
    (магнитная линза, N<sub>ф</sub> витков) и <b>электромагнитным отклонением</b>
    (две пары катушек, по N<sub>о</sub> витков). Снимаются: семейство
    <b>модуляционных характеристик</b> I<sub>к</sub>(U<sub>м</sub>) при трёх значениях
    U<sub>а</sub> (крутизна S и проницаемость модулятора &Delta; = &minus;&Delta;U<sub>м</sub>/&Delta;U<sub>а</sub>
    при I<sub>к</sub> = const); зависимость тока <b>фокусирующей катушки</b> от U<sub>а</sub>
    (фокусное расстояние f = 10&radic;(U<sub>а</sub>·R)/(N<sub>ф</sub>·I<sub>ф</sub>));
    <b>отклонение пятна</b> D от тока катушек (чувствительность &epsilon; = D/(N<sub>о</sub>·I<sub>отк</sub>),
    мм/ампервиток) и от U<sub>а</sub> (D &prop; 1/&radic;U<sub>а</sub>); <b>яркость экрана</b>
    (фототок фотоэлемента) от тока луча и от U<sub>а</sub>.
  </div>

  <h2 class="sec">Параметры установки</h2>
  <div class="params">
    <div class="params-grid">
      <label>U<sub>а</sub> № 1 (характеристики), В
        <input type="number" id="p-ua1" step="any" min="1" value="750">
      </label>
      <label>U<sub>а</sub> № 2, В
        <input type="number" id="p-ua2" step="any" min="1" value="1000">
      </label>
      <label>U<sub>а</sub> № 3, В
        <input type="number" id="p-ua3" step="any" min="1" value="1250">
      </label>
      <label>U<sub>а</sub> номинальное, В
        <input type="number" id="p-unom" step="any" min="1" value="1000">
      </label>
      <label>I<sub>к</sub> для &Delta; (I<sub>реф</sub>), мкА
        <input type="number" id="p-iref" step="any" min="1" value="150">
      </label>
      <label>Витки фокус. катушки N<sub>ф</sub>
        <input type="number" id="p-nf" step="any" min="1" value="2500">
      </label>
      <label>Средний диаметр фокус. катушки, мм
        <input type="number" id="p-df" step="any" min="0.1" value="70">
      </label>
      <label>Витки отклон. катушек N<sub>о</sub>
        <input type="number" id="p-no" step="any" min="1" value="200">
      </label>
      <label>Протяжённость поля a, мм
        <input type="number" id="p-a" step="any" min="0.1" value="20">
      </label>
      <label>Диаметр горловины d, мм
        <input type="number" id="p-d" step="any" min="0.1" value="30">
      </label>
      <label>Линза — экран L, мм
        <input type="number" id="p-l" step="any" min="0.1" value="200">
      </label>
    </div>
    <p class="hint">Значения по умолчанию соответствуют методическим указаниям (N<sub>ф</sub> = 2500 витков, средний диаметр катушки 70 мм, N<sub>о</sub> = 200 витков). Формулы методички: N·Y = 10&radic;(U<sub>а</sub>·R)/f — ампервитки линзы (R в см, f в см); N·Y = 2.7&radic;U<sub>а</sub>·D·d/(a·L) — ампервитки отклоняющих катушек.</p>
  </div>

  <div class="toolbar">
    <span class="mini">Строк в таблицах: <input type="number" id="n-rows" min="4" max="60" value="21"></span>
    <button class="ghost" onclick="buildAll()">Изменить число строк</button>
    <button class="ghost" onclick="fillDemo()">Заполнить демонстрационными данными</button>
  </div>

  <h2 class="sec">Таблица 1. Модуляционные характеристики: I<sub>к</sub> от U<sub>м</sub> (задание 3.1.4)</h2>
  <div class="mods-grid">
    <div class="modbox">
      <div class="ttl">U<sub>а</sub> = <span class="ua-label">750</span> В</div>
      <table class="data"><thead><tr><th>№</th><th>U<sub>м</sub>, В</th><th>I<sub>к</sub>, мкА</th></tr></thead>
      <tbody id="mod-1"></tbody></table>
    </div>
    <div class="modbox">
      <div class="ttl">U<sub>а</sub> = <span class="ua-label">1000</span> В</div>
      <table class="data"><thead><tr><th>№</th><th>U<sub>м</sub>, В</th><th>I<sub>к</sub>, мкА</th></tr></thead>
      <tbody id="mod-2"></tbody></table>
    </div>
    <div class="modbox">
      <div class="ttl">U<sub>а</sub> = <span class="ua-label">1250</span> В</div>
      <table class="data"><thead><tr><th>№</th><th>U<sub>м</sub>, В</th><th>I<sub>к</sub>, мкА</th></tr></thead>
      <tbody id="mod-3"></tbody></table>
    </div>
  </div>

  <h2 class="sec">Таблица 2. Ток фокусирующей катушки (задание 3.1.5)</h2>
  <div class="data-wrap">
    <table class="data"><thead><tr><th>№</th><th>U<sub>а</sub>, В</th><th>I<sub>ф</sub>, мА</th></tr></thead>
    <tbody id="focus-body"></tbody></table>
  </div>

  <h2 class="sec">Таблица 3. Отклонение пятна от тока катушек (задание 3.1.6)</h2>
  <div class="data-wrap">
    <table class="data"><thead><tr><th>№</th><th>I<sub>отк</sub>, мА</th><th>D<sub>верт</sub>, мм</th><th>D<sub>гор</sub>, мм</th></tr></thead>
    <tbody id="di-body"></tbody></table>
  </div>

  <h2 class="sec">Таблица 4. Отклонение пятна от потенциала анода (задание 3.1.7)</h2>
  <div class="data-wrap">
    <table class="data"><thead><tr><th>№</th><th>U<sub>а</sub>, В</th><th>D<sub>верт</sub>, мм</th><th>D<sub>гор</sub>, мм</th></tr></thead>
    <tbody id="du-body"></tbody></table>
  </div>

  <h2 class="sec">Таблица 5. Яркость экрана (фототок фотоэлемента)</h2>
  <div class="two-cols">
    <div class="data-wrap">
      <div style="font-weight:600;font-size:13.5px;color:var(--accent);margin-bottom:6px;">а) от тока луча (U<sub>а</sub> — номинальное)</div>
      <table class="data"><thead><tr><th>№</th><th>I<sub>л</sub>, мкА</th><th>J<sub>ф</sub>, мкА</th></tr></thead>
      <tbody id="bi-body"></tbody></table>
    </div>
    <div class="data-wrap">
      <div style="font-weight:600;font-size:13.5px;color:var(--accent);margin-bottom:6px;">б) от потенциала анода (I<sub>л</sub> — постоянный)</div>
      <table class="data"><thead><tr><th>№</th><th>U<sub>а</sub>, В</th><th>J<sub>ф</sub>, мкА</th></tr></thead>
      <tbody id="bu-body"></tbody></table>
    </div>
  </div>

  <div class="toolbar">
    <button onclick="processAll()">Рассчитать и построить графики</button>
  </div>

  <div class="calc-panel" id="calc-panel"></div>
  <div class="result-panel" id="res-panel"></div>

  <div class="plot-wrap" id="plot1">
    <div class="ptitle">График 1. Семейство модуляционных характеристик I<sub>к</sub>(U<sub>м</sub>)</div>
    <img id="plot1-img" alt="График 1">
    <div class="dl-stack"><a class="dl" id="plot1-dl-png" download="Mod_chars.png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot1-dl-xlsx" download="Mod_chars.xlsx">Скачать график EXCEL</a></div>
  </div>
  <div class="plot-wrap" id="plot2">
    <div class="ptitle">График 2. Ток фокусирующей катушки I<sub>ф</sub>(U<sub>а</sub>) и модель N·Y = 10&radic;(U·R)/f</div>
    <img id="plot2-img" alt="График 2">
    <div class="dl-stack"><a class="dl" id="plot2-dl-png" download="Focus.png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot2-dl-xlsx" download="Focus.xlsx">Скачать график EXCEL</a></div>
  </div>
  <div class="plot-wrap" id="plot3">
    <div class="ptitle">График 3. Отклонение пятна от тока отклоняющих катушек D(I<sub>отк</sub>)</div>
    <img id="plot3-img" alt="График 3">
    <div class="dl-stack"><a class="dl" id="plot3-dl-png" download="Defl_current.png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot3-dl-xlsx" download="Defl_current.xlsx">Скачать график EXCEL</a></div>
  </div>
  <div class="plot-wrap" id="plot4">
    <div class="ptitle">График 4. Отклонение пятна от потенциала анода D(U<sub>а</sub>)</div>
    <img id="plot4-img" alt="График 4">
    <div class="dl-stack"><a class="dl" id="plot4-dl-png" download="Defl_anode.png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot4-dl-xlsx" download="Defl_anode.xlsx">Скачать график EXCEL</a></div>
  </div>
  <div class="plot-wrap" id="plot5">
    <div class="ptitle">График 5. Яркость экрана: от тока луча и от потенциала анода</div>
    <img id="plot5-img" alt="График 5">
    <div class="dl-stack"><a class="dl" id="plot5-dl-png" download="Brightness.png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot5-dl-xlsx" download="Brightness.xlsx">Скачать график EXCEL</a></div>
  </div>

  <p class="muted">Совет: заполните таблицы (или нажмите «Заполнить демонстрационными данными»), затем «Рассчитать и построить графики». Крутизна модуляционной характеристики определяется линейной аппроксимацией рабочего участка кривой при U<sub>а</sub> = U<sub>а ном</sub>; проницаемость — по сдвигу U<sub>м</sub> при постоянном I<sub>к</sub> = I<sub>реф</sub>; чувствительность отклонения — по наклону D(I<sub>отк</sub>) в мм/ампервиток.</p>
</main>

<script>
const HEAT_COLORS = ['#1f4e79', '#c62828', '#2e7d32', '#6a1b9a', '#ef6c00'];

// Демонстрационные данные: закон степени 3/2 I_к = G·(U_м + Δ·U_а)^1.5 (Δ = 0.08),
// N·Y = 10√(U·R)/f при f = 6 см (N_ф = 2500, R = 3.5 см),
// D = ε·N_о·I_отк при ε_V = 1.45 и ε_H = 1.65 мм/ав (теория 1.56 мм/ав).
const DEMO_MOD_UA = [750, 1000, 1250];
const DEMO_MOD = [
  { um: [-56, -52, -48, -44, -40, -36, -32, -28, -24, -20, -16, -12, -8, -4, 0], ik: [4.2, 11.8, 21.6, 33.3, 46.6, 61.2, 77.1, 94.2, 112.4, 131.7, 151.9, 173.1, 195.2, 218.1, 241.9] },
  { um: [-76, -72, -68, -64, -60, -56, -52, -48, -44, -40, -36, -32, -28, -24, -20, -16, -12, -8, -4, 0], ik: [4.2, 11.8, 21.6, 33.3, 46.6, 61.2, 77.1, 94.2, 112.4, 131.7, 151.9, 173.1, 195.2, 218.1, 241.9, 266.5, 291.9, 318.0, 344.9, 372.5] },
  { um: [-96, -92, -88, -84, -80, -76, -72, -68, -64, -60, -56, -52, -48, -44, -40, -36, -32, -28, -24, -20, -16, -12, -8, -4, 0], ik: [4.2, 11.8, 21.6, 33.3, 46.6, 61.2, 77.1, 94.2, 112.4, 131.7, 151.9, 173.1, 195.2, 218.1, 241.9, 266.5, 291.9, 318.0, 344.9, 372.5, 400.7, 429.7, 459.3, 489.6, 520.5] }
];
const DEMO_FOCUS = [
  [400, 24.94], [500, 27.89], [600, 30.55], [700, 33.0], [800, 35.28], [900, 37.42],
  [1000, 39.44], [1100, 41.37], [1200, 43.2], [1300, 44.97], [1400, 46.67], [1500, 48.3], [1600, 49.89]
];
const DEMO_DEFL_I = [
  [0, 0.0, 0.0], [20, 5.8, 6.6], [40, 11.6, 13.2], [60, 17.4, 19.8], [80, 23.2, 26.4],
  [100, 29.0, 33.0], [120, 34.8, 39.6], [140, 40.6, 46.2], [160, 46.4, 52.8], [180, 52.2, 59.4],
  [200, 58.0, 66.0], [220, 63.8, 72.6], [240, 69.6, 79.2], [260, 75.4, 85.8], [280, 81.2, 92.4]
];
const DEMO_DEFL_U = [
  [500, 61.5, 70.0], [600, 56.2, 63.9], [700, 52.0, 59.2], [800, 48.6, 55.3], [900, 45.9, 52.2],
  [1000, 43.5, 49.5], [1100, 41.5, 47.2], [1200, 39.7, 45.2], [1300, 38.2, 43.4], [1400, 36.8, 41.8], [1500, 35.5, 40.4]
];
const DEMO_BR_I = [
  [0, 0.0], [25, 0.5], [50, 1.0], [75, 1.5], [100, 2.0], [125, 2.5], [150, 3.0],
  [175, 3.5], [200, 4.0], [225, 4.5], [250, 5.0], [275, 5.5], [300, 6.0]
];
const DEMO_BR_U = [
  [500, 1.06], [600, 1.39], [700, 1.76], [800, 2.15], [900, 2.56], [1000, 3.0],
  [1100, 3.46], [1200, 3.94], [1300, 4.45], [1400, 4.97], [1500, 5.51]
];

function num(id) {
  const v = parseFloat(document.getElementById(id).value);
  return isNaN(v) ? null : v;
}

function fillTbody(tbodyId, nRows, classes) {
  const tb = document.getElementById(tbodyId);
  tb.innerHTML = '';
  for (let i = 1; i <= nRows; i++) {
    let h = '<tr><td class="num-col">' + i + '</td>';
    classes.forEach(c => h += '<td><input type="number" step="any" class="' + c + '" placeholder="—"></td>');
    h += '</tr>';
    tb.insertAdjacentHTML('beforeend', h);
  }
}

function buildAll() {
  const nR = Math.max(4, Math.min(60, parseInt(document.getElementById('n-rows').value) || 21));
  ['mod-1', 'mod-2', 'mod-3'].forEach(id => fillTbody(id, nR, ['um', 'ik']));
  fillTbody('focus-body', nR, ['ua', 'if']);
  fillTbody('di-body', nR, ['io', 'dv', 'dh']);
  fillTbody('du-body', nR, ['ua2', 'dv2', 'dh2']);
  fillTbody('bi-body', nR, ['il', 'jf']);
  fillTbody('bu-body', nR, ['ua3', 'jf2']);
}

function updateUaLabels() {
  document.querySelectorAll('.ua-label').forEach((el, k) => {
    el.textContent = document.getElementById('p-ua' + (k + 1)).value;
  });
}

function fillDemo() {
  document.getElementById('p-ua1').value = DEMO_MOD_UA[0];
  document.getElementById('p-ua2').value = DEMO_MOD_UA[1];
  document.getElementById('p-ua3').value = DEMO_MOD_UA[2];
  document.getElementById('p-unom').value = 1000;
  document.getElementById('p-iref').value = 150;
  document.getElementById('p-nf').value = 2500;
  document.getElementById('p-df').value = 70;
  document.getElementById('p-no').value = 200;
  document.getElementById('p-a').value = 20;
  document.getElementById('p-d').value = 30;
  document.getElementById('p-l').value = 200;
  updateUaLabels();
  DEMO_MOD.forEach((g, k) => {
    fillTbody('mod-' + (k + 1), g.um.length, ['um', 'ik']);
    const tb = document.getElementById('mod-' + (k + 1));
    [...tb.querySelectorAll('tr')].forEach((tr, i) => {
      tr.querySelector('.um').value = g.um[i];
      tr.querySelector('.ik').value = g.ik[i];
    });
  });
  fillTbody('focus-body', DEMO_FOCUS.length, ['ua', 'if']);
  [...document.getElementById('focus-body').querySelectorAll('tr')].forEach((tr, i) => {
    tr.querySelector('.ua').value = DEMO_FOCUS[i][0];
    tr.querySelector('.if').value = DEMO_FOCUS[i][1];
  });
  fillTbody('di-body', DEMO_DEFL_I.length, ['io', 'dv', 'dh']);
  [...document.getElementById('di-body').querySelectorAll('tr')].forEach((tr, i) => {
    tr.querySelector('.io').value = DEMO_DEFL_I[i][0];
    tr.querySelector('.dv').value = DEMO_DEFL_I[i][1];
    tr.querySelector('.dh').value = DEMO_DEFL_I[i][2];
  });
  fillTbody('du-body', DEMO_DEFL_U.length, ['ua2', 'dv2', 'dh2']);
  [...document.getElementById('du-body').querySelectorAll('tr')].forEach((tr, i) => {
    tr.querySelector('.ua2').value = DEMO_DEFL_U[i][0];
    tr.querySelector('.dv2').value = DEMO_DEFL_U[i][1];
    tr.querySelector('.dh2').value = DEMO_DEFL_U[i][2];
  });
  fillTbody('bi-body', DEMO_BR_I.length, ['il', 'jf']);
  [...document.getElementById('bi-body').querySelectorAll('tr')].forEach((tr, i) => {
    tr.querySelector('.il').value = DEMO_BR_I[i][0];
    tr.querySelector('.jf').value = DEMO_BR_I[i][1];
  });
  fillTbody('bu-body', DEMO_BR_U.length, ['ua3', 'jf2']);
  [...document.getElementById('bu-body').querySelectorAll('tr')].forEach((tr, i) => {
    tr.querySelector('.ua3').value = DEMO_BR_U[i][0];
    tr.querySelector('.jf2').value = DEMO_BR_U[i][1];
  });
  document.getElementById('n-rows').value = DEMO_MOD[1].um.length;
}

function collectGroup(k) {
  const tb = document.getElementById('mod-' + k);
  return {
    um: [...tb.querySelectorAll('.um')].map(x => x.value),
    ik: [...tb.querySelectorAll('.ik')].map(x => x.value)
  };
}

function collectRows(tbodyId, classes) {
  return [...document.getElementById(tbodyId).querySelectorAll('tr')].map(tr =>
    classes.map(c => tr.querySelector('.' + c).value));
}

async function processAll() {
  const p = { u_a: [num('p-ua1'), num('p-ua2'), num('p-ua3')],
              u_nom: num('p-unom'), i_ref: num('p-iref'),
              n_f: num('p-nf'), d_f: num('p-df'), n_o: num('p-no'),
              a_mm: num('p-a'), d_mm: num('p-d'), l_mm: num('p-l') };
  for (const [k, name] of [['u_nom', 'U_а ном'], ['n_f', 'N_ф'], ['d_f', 'диаметр катушки'],
                           ['n_o', 'N_о'], ['a_mm', 'a'], ['d_mm', 'd'], ['l_mm', 'L']]) {
    if (p[k] === null || p[k] <= 0) { alert('Введите положительное значение: ' + name + '.'); return; }
  }
  for (let k = 0; k < 3; k++) {
    if (p.u_a[k] === null || p.u_a[k] <= 0) { alert('Введите U_а № ' + (k + 1) + ' (В).'); return; }
  }
  const payload = Object.assign(p, {
    mod: [collectGroup(1), collectGroup(2), collectGroup(3)],
    focus: collectRows('focus-body', ['ua', 'if']),
    defl_i: collectRows('di-body', ['io', 'dv', 'dh']),
    defl_u: collectRows('du-body', ['ua2', 'dv2', 'dh2']),
    br_i: collectRows('bi-body', ['il', 'jf']),
    br_u: collectRows('bu-body', ['ua3', 'jf2'])
  });
  let res;
  try {
    res = await fetch('/api/vipe/elt/process', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
  } catch (e) { alert('Ошибка сети: ' + e); return; }
  const data = await res.json();
  if (data.error) { alert(data.error); return; }
  const cp = document.getElementById('calc-panel');
  cp.innerHTML = '<p class="formula">Формулы обработки:</p>' +
    '<p class="formula-eq">S = &Delta;I<sub>к</sub>/&Delta;U<sub>м</sub> — крутизна модуляционной характеристики (рабочий участок кривой при U<sub>а ном</sub>)</p>' +
    '<p class="formula-eq">&Delta; = &minus;&Delta;U<sub>м</sub>/&Delta;U<sub>а</sub> при I<sub>к</sub> = I<sub>реф</sub> — проницаемость модулятора</p>' +
    '<p class="formula-eq">N·Y = 10&radic;(U<sub>а</sub>·R)/f &nbsp;&rArr;&nbsp; f = 10&radic;(U<sub>а</sub>·R)/(N<sub>ф</sub>·I<sub>ф</sub>) — фокусное расстояние линзы (R в см, f в см)</p>' +
    '<p class="formula-eq">&epsilon; = D/(N<sub>о</sub>·I<sub>отк</sub>) — чувствительность отклонения, мм/ампервиток</p>' +
    '<p class="formula-eq">N·Y = 2.7&radic;U<sub>а</sub>·D·d/(a·L) &nbsp;&rArr;&nbsp; &epsilon;<sub>теор</sub> = a·L/(2.7&radic;U<sub>а</sub>·d) &nbsp;&rArr;&nbsp; D &prop; 1/&radic;U<sub>а</sub></p>';
  cp.style.display = 'block';
  const rp = document.getElementById('res-panel');
  rp.innerHTML = data.analysis_html;
  rp.style.display = 'block';
  for (const g of ['g1', 'g2', 'g3', 'g4', 'g5']) {
    const idx = g.slice(1);
    const img = data[g + '_image'];
    const xlsx = data[g + '_xlsx'];
    const wrap = document.getElementById('plot' + idx);
    if (img) {
      const src = 'data:image/png;base64,' + img;
      document.getElementById('plot' + idx + '-img').src = src;
      document.getElementById('plot' + idx + '-dl-png').href = src;
    }
    if (xlsx) document.getElementById('plot' + idx + '-dl-xlsx').href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + xlsx;
    wrap.style.display = img ? 'block' : 'none';
  }
  rp.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

document.getElementById('p-ua1').addEventListener('input', updateUaLabels);
document.getElementById('p-ua2').addEventListener('input', updateUaLabels);
document.getElementById('p-ua3').addEventListener('input', updateUaLabels);
buildAll();
</script>
</body>
</html>
"""


STUB_HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__ — Вакуумная и плазменная электроника</title>
<style>
  :root { --accent:#1f4e79; --bg:#f4f6f9; --border:#d8dee9; }
  * { box-sizing:border-box; }
  body { margin:0; font-family:'Segoe UI', system-ui, Arial, sans-serif; background:var(--bg); color:#222; }
  header { background:var(--accent); color:#fff; padding:22px 20px; }
  header .wrap { max-width:1000px; margin:0 auto; display:flex; justify-content:space-between; align-items:center; gap:16px; flex-wrap:wrap; }
  header h1 { margin:0; font-size:19px; }
  header a { color:#fff; text-decoration:none; font-size:13.5px; border:1px solid rgba(255,255,255,.5); padding:6px 14px; border-radius:6px; }
  header a:hover { background:rgba(255,255,255,.15); }
  main { max-width:1000px; margin:40px auto; padding:0 16px; }
  .box { background:#fff; border:1px solid var(--border); border-left:4px solid #d4a72c; border-radius:8px; padding:26px 30px; }
  .box h2 { margin:0 0 10px; font-size:20px; color:var(--accent); }
  .box p { color:#555; font-size:14.5px; margin:0; }
  .badge { display:inline-block; margin-top:16px; font-size:12.5px; padding:4px 12px; border-radius:12px; background:#f0f0f0; color:#888; }
</style>
</head>
<body>
<header>
  <div class="wrap">
    <h1>Вакуумная и плазменная электроника</h1>
    <a href="/" target="_blank">&larr; к списку работ</a>
  </div>
</header>
<main>
  <div class="box">
    <h2>__TITLE__</h2>
    <p>Лабораторная работа находится в разработке. Таблицы ввода данных, автоматический расчёт и построение графиков будут добавлены позже.</p>
    <span class="badge">Заглушка</span>
  </div>
</main>
</body>
</html>
"""

# ===========================================================================
#  Обработка данных лабораторной работы № 1 (термоэлектронная эмиссия)
# ===========================================================================

HEAT_COLORS = ["#1f4e79", "#c62828", "#2e7d32", "#6a1b9a", "#ef6c00"]


def _f(v):
    """Безопасный парсинг числа из входа JSON."""
    s = str(v).strip() if v is not None else ""
    if s == "" or s in ("None", "—"):
        return None
    try:
        return float(s)
    except (TypeError, ValueError):
        return None


def _r2_of(y, yfit):
    sst = float(np.sum((np.asarray(y) - np.mean(y)) ** 2))
    sse = float(np.sum((np.asarray(yfit) - np.asarray(y)) ** 2))
    return 1.0 - sse / sst if sst > 0 else 1.0


def make_vax_plot(heats):
    """Семейство ВАХ с отмеченными точками перегиба (токами эмиссии)."""
    fig, ax = plt.subplots(figsize=(9, 6), constrained_layout=True)
    for k, h in enumerate(heats):
        c = HEAT_COLORS[k % len(HEAT_COLORS)]
        ax.plot(h["u"], h["i"], "-o", color=c, markersize=4.5, linewidth=1.4,
                label=h["label"])
        ax.plot([h["u_knee"]], [h["i_knee"]], marker="*", markersize=16,
                color=c, markeredgecolor="#333", zorder=6)
    ax.plot([], [], marker="*", linestyle="none", color="#555", markersize=14,
            label="Точка перегиба (I$_{Э}$)")
    ax.set_xlabel("U$_{а}$, В", fontsize=12)
    ax.set_ylabel("I$_{а}$, мА", fontsize=12)
    ax.set_title("Семейство вольтамперных характеристик диода", fontsize=12)
    ax.grid(True, alpha=0.3)
    ax.legend(loc="best", fontsize=9.5)
    return fig_to_base64(fig)


def make_richardson_plot(pts, rich):
    """Прямая Ричардсона ln(j/T²) = f(1/T) с аппроксимирующей прямой."""
    xs = np.array([1.0 / t for t, _, _ in pts])
    ys = np.array([math.log(j / (t * t)) for t, j, _ in pts])
    fig, ax = plt.subplots(figsize=(9, 5.8), constrained_layout=True)
    ax.plot(xs, ys, "o", color="#1f4e79", markersize=7, zorder=5,
            label="Экспериментальные точки")
    xl = np.linspace(float(xs.min()) * 0.97, float(xs.max()) * 1.03, 50)
    ax.plot(xl, rich["slope"] * xl + rich["intercept"], "--", color="#c62828",
            linewidth=1.7, label="МНК-аппроксимация")
    sign = "+" if rich["intercept"] >= 0 else "−"
    ax.set_xlabel("1/T, К$^{-1}$", fontsize=12)
    ax.set_ylabel("ln(j/T$^2$)", fontsize=12)
    ax.set_title("Прямая Ричардсона", fontsize=12)
    ax.grid(True, alpha=0.3)
    ax.text(0.04, 0.95,
            (f"eφ = {rich['phi']:.3f} эВ\n"
             f"A = {rich['A']:.3e} А/(м²·К²)\n"
             f"R² = {rich['r2']:.5f}\n"
             f"ln(j/T²) = {rich['slope']:.1f}·(1/T) {sign} {abs(rich['intercept']):.2f}"),
            transform=ax.transAxes, fontsize=9.5, va="top",
            bbox=dict(boxstyle="round,pad=0.4", fc="white", ec="#999", alpha=0.9))
    ax.legend(loc="lower left", fontsize=9.5)
    return fig_to_base64(fig)


def make_schottky_plot(heats):
    """Прямые Шоттки ln j = f(√E) в области насыщения для всех накалов."""
    fig, ax = plt.subplots(figsize=(9, 6), constrained_layout=True)
    for k, h in enumerate(heats):
        c = HEAT_COLORS[k % len(HEAT_COLORS)]
        sch = h["sch"]
        ax.plot(sch["x"], sch["y"], "o", color=c, markersize=5.5,
                label=f"{h['label']}, T ≈ {h['T']:.0f} К" if h.get("T") else h["label"])
        if "beta" in sch:
            xf = np.array([min(sch["x"]), max(sch["x"])])
            ax.plot(xf, sch["beta"] * xf + sch["b0"], "--", color=c, linewidth=1.5)
    ax.set_xlabel("√E, (В/м)$^{1/2}$", fontsize=12)
    ax.set_ylabel("ln j, j в А/м²", fontsize=12)
    ax.set_title("Прямые Шоттки в области насыщения ВАХ", fontsize=12)
    ax.grid(True, alpha=0.3)
    ax.legend(loc="best", fontsize=9)
    return fig_to_base64(fig)


def make_jT_plot(pts, rich):
    """Зависимость j = f(T): экспериментальные точки и модель Ричардсона-Дешмана."""
    ts = np.array([t for t, _, _ in pts])
    js = np.array([j for _, j, _ in pts])
    fig, ax = plt.subplots(figsize=(9, 5.8), constrained_layout=True)
    ax.plot(ts, js, "o", color="#1f4e79", markersize=8, zorder=5,
            label="Эксперимент (токи эмиссии)")
    tm = np.linspace(float(ts.min()) * 0.95, float(ts.max()) * 1.05, 200)
    jm = rich["A"] * tm ** 2 * np.exp(-rich["phi"] * E_CHARGE / (K_BOLTZ * tm))
    ax.plot(tm, jm, "--", color="#c62828", linewidth=1.7,
            label=(f"Модель j = A·T²·exp(−eφ/kT)\n"
                   f"(eφ = {rich['phi']:.3f} эВ, A = {rich['A']:.2e})"))
    ax.set_yscale("log")
    ax.set_xlabel("T, К", fontsize=12)
    ax.set_ylabel("j, А/м²", fontsize=12)
    ax.set_title("Зависимость плотности тока эмиссии от температуры катода", fontsize=12)
    ax.grid(True, which="both", alpha=0.3)
    ax.legend(loc="lower right", fontsize=9.5)
    return fig_to_base64(fig)


def _chart_common(title, x_title, y_title, width=22, height=11):
    ch = ScatterChart()
    ch.title = title
    ch.x_axis.title = x_title
    ch.y_axis.title = y_title
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.height = height
    ch.width = width
    return ch


def make_vax_xlsx(heats):
    """Excel: ВАХ всех накалов (пары столбцов U_а/I_а) + параметры перегиба."""
    wb = Workbook()
    ws = wb.active
    ws.title = "ВАХ"
    headers = ["№"]
    nmax = 0
    for k, h in enumerate(heats, 1):
        headers += [f"U_а {k} ({h['label']}), В", f"I_а {k}, мА"]
        nmax = max(nmax, len(h["u"]))
    ws.append(headers)
    for i in range(nmax):
        row = [i + 1]
        for h in heats:
            if i < len(h["u"]):
                row += [round(h["u"][i], 4), round(h["i"][i], 4)]
            else:
                row += ["", ""]
        ws.append(row)
    prow = nmax + 3
    ws.cell(row=prow, column=1, value="Накал")
    ws.cell(row=prow, column=2, value="U_перегиба, В")
    ws.cell(row=prow, column=3, value="I_Э (ток эмиссии), мА")
    ws.cell(row=prow, column=4, value="j = I_Э/S, А/м²")
    for k, h in enumerate(heats, 1):
        ws.cell(row=prow + k, column=1, value=h["label"])
        ws.cell(row=prow + k, column=2, value=round(h["u_knee"], 3))
        ws.cell(row=prow + k, column=3, value=round(h["i_knee"], 4))
        ws.cell(row=prow + k, column=4, value=round(h["j"], 2))
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 20
    chart = _chart_common("Семейство ВАХ диода", "U_а, В", "I_а, мА")
    for k, h in enumerate(heats, 1):
        xcol = 2 + (k - 1) * 2
        ycol = xcol + 1
        xref = Reference(ws, min_col=xcol, min_row=2, max_row=nmax + 1)
        yref = Reference(ws, min_col=ycol, min_row=2, max_row=nmax + 1)
        s = Series(yref, xref, title=f"I_а {k} ({h['label']})")
        s.marker = Marker(symbol="circle", size=6)
        s.graphicalProperties = GraphicalProperties()
        s.graphicalProperties.line = LineProperties(solidFill=HEAT_COLORS[(k - 1) % len(HEAT_COLORS)][1:], w=18000)
        chart.series.append(s)
    ws.add_chart(chart, "B" + str(prow + len(heats) + 3))
    return wb_to_base64(wb)


def make_richardson_xlsx(pts, rich):
    """Excel: прямая Ричардсона — точки, аппроксимация и константы эмиссии."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Ричардсон"
    ws.append(["№", "T, К", "1/T, К⁻¹", "j, А/м²", "ln(j/T²)",
               "ln(j/T²) аппрокс."])
    n = len(pts)
    for i, (t, j, _) in enumerate(pts, 1):
        x = 1.0 / t
        y = math.log(j / (t * t))
        ws.append([i, round(t, 1), round(x, 8), round(j, 3), round(y, 5),
                   round(rich["slope"] * x + rich["intercept"], 5)])
    prow = n + 3
    items = [("Наклон прямой (−eφ/k), К", round(rich["slope"], 1)),
             ("Отрезок ln A", round(rich["intercept"], 4)),
             ("Работа выхода eφ, эВ", round(rich["phi"], 4)),
             ("Потенциал выхода φ, В", round(rich["phi"], 4)),
             ("Константа Ричардсона A, А/(м²·К²)", f"{rich['A']:.4e}"),
             ("Теоретическое A₀, А/(м²·К²)", f"{A0_THEOR:.2e}"),
             ("R²", round(rich["r2"], 5))]
    ws.cell(row=prow, column=1, value="Параметр")
    ws.cell(row=prow, column=2, value="Значение")
    for k, (key, val) in enumerate(items, 1):
        ws.cell(row=prow + k, column=1, value=key)
        ws.cell(row=prow + k, column=2, value=val)
    ws.column_dimensions["A"].width = 38
    for col in range(2, 7):
        ws.column_dimensions[get_column_letter(col)].width = 18
    chart = _chart_common("Прямая Ричардсона", "1/T, К⁻¹", "ln(j/T²)")
    xref = Reference(ws, min_col=3, min_row=2, max_row=n + 1)
    s_exp = Series(Reference(ws, min_col=5, min_row=2, max_row=n + 1), xref, title="Эксперимент")
    s_exp.marker = Marker(symbol="circle", size=7)
    s_exp.graphicalProperties = GraphicalProperties()
    s_exp.graphicalProperties.line.noFill = True
    chart.series.append(s_exp)
    s_fit = Series(Reference(ws, min_col=6, min_row=2, max_row=n + 1), xref, title="МНК")
    s_fit.marker = Marker(symbol="none")
    s_fit.graphicalProperties = GraphicalProperties()
    s_fit.graphicalProperties.line = LineProperties(solidFill="C62828", w=20000)
    chart.series.append(s_fit)
    ws.add_chart(chart, "H2")
    return wb_to_base64(wb)


def make_schottky_xlsx(heats):
    """Excel: прямые Шоттки — (√E, ln j) по накалам + наклоны и температуры."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Шоттки"
    headers = ["№"]
    nmax = 0
    for k, h in enumerate(heats, 1):
        headers += [f"√E {k}, (В/м)⁰·⁵", f"ln j {k}", f"E {k}, кВ/м", f"j {k}, А/м²"]
        nmax = max(nmax, len(h["sch"]["x"]))
    ws.append(headers)
    for i in range(nmax):
        row = [i + 1]
        for h in heats:
            sch = h["sch"]
            if i < len(sch["x"]):
                row += [round(sch["x"][i], 2), round(sch["y"][i], 5),
                        round(sch["E"][i] / 1e3, 3), round(sch["j"][i], 3)]
            else:
                row += ["", "", "", ""]
        ws.append(row)
    prow = nmax + 3
    ws.cell(row=prow, column=1, value="Накал")
    ws.cell(row=prow, column=2, value="Наклон B, (А/м²)⁻¹·(В/м)⁻⁰·⁵")
    ws.cell(row=prow, column=3, value="T_Шоттки, К")
    ws.cell(row=prow, column=4, value="T_калибр., К")
    ws.cell(row=prow, column=5, value="R²")
    for k, h in enumerate(heats, 1):
        sch = h["sch"]
        ws.cell(row=prow + k, column=1, value=h["label"])
        if "beta" in sch:
            ws.cell(row=prow + k, column=2, value=f"{sch['beta']:.3e}")
            ws.cell(row=prow + k, column=3, value=round(sch["T_sh"], 1))
            ws.cell(row=prow + k, column=4, value=round(h["T"], 1) if h.get("T") else "—")
            ws.cell(row=prow + k, column=5, value=round(sch["r2"], 5))
        else:
            ws.cell(row=prow + k, column=2, value="—")
            ws.cell(row=prow + k, column=3, value="—")
            ws.cell(row=prow + k, column=4, value=round(h["T"], 1) if h.get("T") else "—")
            ws.cell(row=prow + k, column=5, value="—")
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 20
    chart = _chart_common("Прямые Шоттки", "√E, (В/м)⁰·⁵", "ln j")
    for k, h in enumerate(heats, 1):
        sch = h["sch"]
        if len(sch["x"]) < 2:
            continue
        xcol = 2 + (k - 1) * 4
        xref = Reference(ws, min_col=xcol, min_row=2, max_row=len(sch["x"]) + 1)
        s = Series(Reference(ws, min_col=xcol + 1, min_row=2, max_row=len(sch["x"]) + 1),
                   xref, title=f"ln j {k}")
        s.marker = Marker(symbol="circle", size=6)
        s.graphicalProperties = GraphicalProperties()
        s.graphicalProperties.line = LineProperties(solidFill=HEAT_COLORS[(k - 1) % len(HEAT_COLORS)][1:], w=18000)
        chart.series.append(s)
    ws.add_chart(chart, "G" + str(prow + len(heats) + 3))
    return wb_to_base64(wb)


def make_jT_xlsx(pts, rich):
    """Excel: зависимость j = f(T) с модельной кривой."""
    wb = Workbook()
    ws = wb.active
    ws.title = "j(T)"
    ws.append(["№", "T, К", "j эксп., А/м²", "j модель, А/м²"])
    n = len(pts)
    for i, (t, j, _) in enumerate(pts, 1):
        jm = rich["A"] * t ** 2 * math.exp(-rich["phi"] * E_CHARGE / (K_BOLTZ * t))
        ws.append([i, round(t, 1), round(j, 3), round(jm, 3)])
    prow = n + 3
    items = [("Работа выхода eφ, эВ", round(rich["phi"], 4)),
             ("Константа Ричардсона A, А/(м²·К²)", f"{rich['A']:.4e}"),
             ("R² (прямая Ричардсона)", round(rich["r2"], 5))]
    ws.cell(row=prow, column=1, value="Параметр")
    ws.cell(row=prow, column=2, value="Значение")
    for k, (key, val) in enumerate(items, 1):
        ws.cell(row=prow + k, column=1, value=key)
        ws.cell(row=prow + k, column=2, value=val)
    for col in range(1, 5):
        ws.column_dimensions[get_column_letter(col)].width = 26
    chart = _chart_common("Зависимость j = f(T)", "T, К", "j, А/м²")
    xref = Reference(ws, min_col=2, min_row=2, max_row=n + 1)
    s_exp = Series(Reference(ws, min_col=3, min_row=2, max_row=n + 1), xref, title="j экспер.")
    s_exp.marker = Marker(symbol="circle", size=7)
    s_exp.graphicalProperties = GraphicalProperties()
    s_exp.graphicalProperties.line.noFill = True
    chart.series.append(s_exp)
    s_mod = Series(Reference(ws, min_col=4, min_row=2, max_row=n + 1), xref, title="j модель")
    s_mod.marker = Marker(symbol="none")
    s_mod.graphicalProperties = GraphicalProperties()
    s_mod.graphicalProperties.line = LineProperties(solidFill="C62828", w=20000)
    chart.series.append(s_mod)
    ws.add_chart(chart, "F2")
    return wb_to_base64(wb)


def build_tee_analysis(proc, rich, rich_pts):
    """HTML-панель результатов: таблица обработки + константы эмиссии + выводы."""
    html = '<h4>Обработка результатов (Таблица 3.2)</h4>'
    html += ('<table class="res"><tr><th>Накал</th><th>I<sub>н</sub>, А</th>'
             '<th>U<sub>н</sub>, В</th><th>P, Вт/см²</th><th>I′<sub>н</sub>, А'
             '</th><th>T, К</th><th>U<sub>пер</sub>, В</th><th>I<sub>Э</sub>, мА</th>'
             '<th>j, А/м²</th><th>Точек Шоттки</th><th>T<sub>Ш</sub>, К</th></tr>')
    for r in proc:
        if "error" in r:
            html += f'<tr><td>№ {r["num"]}</td><td colspan="10" class="err">{r["error"]}</td></tr>'
            continue
        sch = r.get("sch", {})
        t_sh = f"{sch['T_sh']:.0f}" if "T_sh" in sch else "—"
        html += (f'<tr><td>№ {r["num"]}</td><td>{r["i_n"]:g}</td><td>{r["u_n"]:g}</td>'
                 f'<td>{r["P_ud"]:.1f}</td><td>{r["I_unit"]:.3f}</td>'
                 f'<td>{r["T"]:.0f}</td><td>{r["u_knee"]:.1f}</td>'
                 f'<td>{r["i_knee"]:.3f}</td><td>{r["j"]:.1f}</td>'
                 f'<td>{len(sch.get("x", []))}</td><td>{t_sh}</td></tr>')
    html += '</table>'
    if rich:
        sign = "+" if rich["intercept"] >= 0 else "−"
        html += (f'<p style="font-size:13.5px;margin:10px 0 4px;"><b>Прямая Ричардсона</b> '
                 f'(по {rich["n"]} накалам): ln(j/T²) = {rich["slope"]:.1f}·(1/T) {sign} '
                 f'{abs(rich["intercept"]):.2f}, R² = {rich["r2"]:.5f}.</p>')
        html += (f'<p style="font-size:13.5px;margin:4px 0;"><b>Константы термоэмиссии:</b> '
                 f'работа выхода e&phi; = <b>{rich["phi"]:.3f} эВ</b> (φ = {rich["phi"]:.3f} В), '
                 f'константа Ричардсона A = <b>{rich["A"]:.3e} А/(м²·К²)</b> '
                 f'(теоретическое значение A₀ = {A0_THEOR:.1e} А/(м²·К²)).</p>')
    html += ('<p class="note" style="color:#555;margin-top:8px;">Ток эмиссии I_Э определяется '
             'по точке пересечения прямых кусочно-линейной аппроксимации ВАХ (точка перегиба «a»). '
             'Для прямых Шоттки используются точки ВАХ правее перегиба; наклон B даёт '
             'T = (e³/4πε₀)½/(k·B). I′_н = I_н/d_к^1.5 — ток единичного катода (метод сравнения, '
             'd_к в см); температура T рассчитана по удельной мощности P = I_н·U_н/(π·l_к·d_к).</p>')

    html += '<h4 style="margin-top:14px;">Выводы</h4><ul style="font-size:13.5px;line-height:1.6;">'
    html += ('<li>Ток эмиссии (и плотность тока j) сильно (экспоненциально) растёт с температурой '
             'катода: прямая Ричардсона ln(j/T²) = f(1/T) линейна, что подтверждает закон '
             'Ричардсона–Дешмана j = A·T²·exp(−e&phi;/kT).</li>')
    if rich:
        if abs(rich["phi"] - 4.5) < 0.5:
            html += (f'<li>Полученная работа выхода e&phi; = {rich["phi"]:.2f} эВ близка к '
                     f'табличному значению для вольфрама (4,5 эВ).</li>')
        else:
            html += (f'<li>Полученная работа выхода e&phi; = {rich["phi"]:.2f} эВ отличается '
                     f'от табличного значения для вольфрама (4,5 эВ) — возможные причины: '
                     f'погрешности определения перегиба ВАХ, температуры и площади катода.</li>')
        if rich["A"] < A0_THEOR:
            html += (f'<li>Константа Ричардсона A = {rich["A"]:.2e} А/(м²·К²) меньше теоретической '
                     f'A₀ = {A0_THEOR:.1e}: коэффициент прозрачности барьера D &lt; 1 и отличие '
                     f'реальной эмиттирующей поверхности от геометрической.</li>')
    sch_ok = [r for r in proc if "error" not in r and "T_sh" in r.get("sch", {}) and r.get("T")]
    if sch_ok:
        devs = [abs(r["sch"]["T_sh"] - r["T"]) / r["T"] * 100 for r in sch_ok]
        html += (f'<li>Температуры, найденные по наклону прямых Шоттки '
                 f'(T_Ш = {min(r["sch"]["T_sh"] for r in sch_ok):.0f}…'
                 f'{max(r["sch"]["T_sh"] for r in sch_ok):.0f} К), согласуются с температурами '
                 f'по удельной мощности накала (отклонение {min(devs):.1f}–{max(devs):.1f}%), '
                 f'что подтверждает закон Шоттки ln j = ln j₀ + B·√E.</li>')
    html += ('<li>В области насыщения анодный ток медленно растёт с U_а из-за снижения '
             'потенциального барьера внешним полем (эффект Шоттки); для измерения тока эмиссии '
             'следует брать ток в точке перегиба ВАХ.</li>')
    html += '</ul>'
    return html


# ===========================================================================
#  Обработка данных лабораторной работы № 2 (ЭЛТ с магнитным управлением)
# ===========================================================================

def _um_at_current(pts_sorted, i_ref):
    """U_м на кривой I_к(U_м) при заданном токе (линейная интерполяция)."""
    for k in range(1, len(pts_sorted)):
        i1, i2 = pts_sorted[k - 1][1], pts_sorted[k][1]
        if i1 <= i_ref <= i2 and i2 > i1:
            u1, u2 = pts_sorted[k - 1][0], pts_sorted[k][0]
            return u1 + (i_ref - i1) * (u2 - u1) / (i2 - i1)
    return None


def _steepness(pts_sorted):
    """Крутизна модуляционной характеристики: МНК по рабочему участку
    (ток от 15 % до 85 % максимального)."""
    pts = [(u, i) for u, i in pts_sorted if i > 0]
    if len(pts) < 2:
        return None
    imax = max(i for _, i in pts)
    work = [(u, i) for u, i in pts if 0.15 * imax <= i <= 0.85 * imax]
    if len(work) < 3:
        work = pts
    xs = np.array([p[0] for p in work])
    ys = np.array([p[1] for p in work])
    s, b = np.polyfit(xs, ys, 1)
    return {"S": float(s), "r2": _r2_of(ys, s * xs + b), "n": len(work)}


def make_mod_plot(groups, i_ref, um_refs):
    """Семейство модуляционных характеристик I_к(U_м) для трёх U_а."""
    fig, ax = plt.subplots(figsize=(9, 6), constrained_layout=True)
    for k, g in enumerate(groups):
        c = HEAT_COLORS[k % len(HEAT_COLORS)]
        ax.plot(g["um"], g["ik"], "-o", color=c, markersize=4.5, linewidth=1.4,
                label=f"U$_{{а}}$ = {g['u_a']:g} В")
        if um_refs[k] is not None:
            ax.plot([um_refs[k]], [i_ref], marker="*", markersize=15, color=c,
                    markeredgecolor="#333", zorder=6)
    if i_ref:
        ax.axhline(i_ref, color="#888", linestyle=":", linewidth=1.2)
        ax.annotate(f"I$_{{реф}}$ = {i_ref:g} мкА", xy=(0.02, i_ref),
                    xycoords=("axes fraction", "data"), fontsize=9,
                    color="#555", va="bottom")
    ax.set_xlabel("U$_{м}$, В", fontsize=12)
    ax.set_ylabel("I$_{к}$, мкА", fontsize=12)
    ax.set_title("Семейство модуляционных характеристик трубки", fontsize=12)
    ax.grid(True, alpha=0.3)
    ax.legend(loc="upper left", fontsize=9.5)
    return fig_to_base64(fig)


def make_focus_plot(pts, f_avg, n_f, r_sm):
    """Ток фокусирующей катушки I_ф(U_а) и модель I_ф = 10·sqrt(U·R)/(N·f)."""
    ua = np.array([p[0] for p in pts])
    if_ = np.array([p[1] for p in pts])
    fig, ax = plt.subplots(figsize=(9, 5.8), constrained_layout=True)
    ax.plot(ua, if_, "o", color="#1f4e79", markersize=7, zorder=5,
            label="Эксперимент")
    um = np.linspace(float(ua.min()) * 0.95, float(ua.max()) * 1.05, 100)
    im = 1000 * 10 * np.sqrt(um * r_sm) / (n_f * f_avg)
    ax.plot(um, im, "--", color="#c62828", linewidth=1.7,
            label=f"Модель N·Y = 10·√(U·R)/f (f = {f_avg:.2f} см)")
    ax.set_xlabel("U$_{а}$, В", fontsize=12)
    ax.set_ylabel("I$_{ф}$, мА", fontsize=12)
    ax.set_title("Ток фокусирующей катушки оптимальной фокусировки", fontsize=12)
    ax.grid(True, alpha=0.3)
    ax.legend(loc="lower right", fontsize=9.5)
    return fig_to_base64(fig)


def make_defl_i_plot(di_v, di_h, b_v, b_h, eps_v, eps_h, eps_t):
    """Отклонение пятна от тока катушек: две пары + прямые МНК."""
    fig, ax = plt.subplots(figsize=(9, 6), constrained_layout=True)
    for pts, b, col, name, eps in (
            (di_v, b_v, HEAT_COLORS[0], "Вертикальные катушки", eps_v),
            (di_h, b_h, HEAT_COLORS[1], "Горизонтальные катушки", eps_h)):
        xs = np.array([p[0] for p in pts])
        ys = np.array([p[1] for p in pts])
        ax.plot(xs, ys, "o", color=col, markersize=6, label=name)
        xl = np.linspace(float(xs.min()), float(xs.max()), 50)
        ax.plot(xl, b * xl, "--", color=col, linewidth=1.6,
                label=f"МНК: {eps:.3f} мм/ав")
    ax.set_xlabel("I$_{отк}$, мА", fontsize=12)
    ax.set_ylabel("D, мм", fontsize=12)
    ax.set_title("Отклонение пятна на экране от тока отклоняющих катушек", fontsize=12)
    ax.grid(True, alpha=0.3)
    ax.text(0.03, 0.96, f"ε$_{{теор}}$ ≈ {eps_t:.3f} мм/ав (при U$_{{а ном}}$)",
            transform=ax.transAxes, fontsize=9.5, va="top",
            bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#999", alpha=0.9))
    ax.legend(loc="lower right", fontsize=9)
    return fig_to_base64(fig)


def make_defl_u_plot(du_v, du_h, c_v, c_h):
    """Отклонение пятна от потенциала анода: две пары + модель c/√U."""
    fig, ax = plt.subplots(figsize=(9, 6), constrained_layout=True)
    for pts, c, col, name in (
            (du_v, c_v, HEAT_COLORS[0], "Вертикальные катушки"),
            (du_h, c_h, HEAT_COLORS[1], "Горизонтальные катушки")):
        xs = np.array([p[0] for p in pts])
        ys = np.array([p[1] for p in pts])
        ax.plot(xs, ys, "o", color=col, markersize=6, label=name)
        um = np.linspace(float(xs.min()) * 0.95, float(xs.max()) * 1.05, 100)
        ax.plot(um, c / np.sqrt(um), "--", color=col, linewidth=1.6,
                label=f"Модель D = {c:.1f}/√U$_{{а}}$")
    ax.set_xlabel("U$_{а}$, В", fontsize=12)
    ax.set_ylabel("D, мм", fontsize=12)
    ax.set_title("Отклонение пятна от потенциала анода (I$_{отк}$ = const)", fontsize=12)
    ax.grid(True, alpha=0.3)
    ax.legend(loc="upper right", fontsize=9)
    return fig_to_base64(fig)


def make_bright_plot(br_i, fit_i, br_u, fit_u):
    """Яркость экрана: график(и) от тока луча и/или от потенциала анода."""
    have_i = bool(br_i) and fit_i is not None and fit_i.get("r2") is not None
    have_u = bool(br_u) and fit_u is not None and fit_u.get("r2") is not None
    if have_i and have_u:
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11.5, 5.2), constrained_layout=True)
    else:
        fig, ax1 = plt.subplots(figsize=(9, 5.4), constrained_layout=True)
        ax2 = None
    if have_i:
        xs = np.array([p[0] for p in br_i])
        ys = np.array([p[1] for p in br_i])
        ax1.plot(xs, ys, "o", color="#1f4e79", markersize=6, label="Эксперимент")
        xl = np.linspace(float(xs.min()), float(xs.max()), 50)
        ax1.plot(xl, fit_i["k"] * xl + fit_i["b"], "--", color="#c62828", linewidth=1.6,
                 label=f"МНК: J = {fit_i['k']:.4f}·I$_{{л}}$ + {fit_i['b']:.2f}")
        ax1.set_xlabel("I$_{л}$, мкА", fontsize=12)
        ax1.set_ylabel("J$_{ф}$, мкА", fontsize=12)
        ax1.set_title("Яркость от тока луча (U$_{а}$ — ном.)", fontsize=11.5)
        ax1.grid(True, alpha=0.3)
        ax1.legend(loc="upper left", fontsize=9)
    if have_u:
        xs = np.array([p[0] for p in br_u])
        ys = np.array([p[1] for p in br_u])
        target = ax2 if ax2 is not None else ax1
        target.plot(xs, ys, "o", color="#2e7d32", markersize=6, label="Эксперимент")
        um = np.linspace(float(xs.min()) * 0.95, float(xs.max()) * 1.05, 100)
        target.plot(um, fit_u["c"] * um ** fit_u["p"], "--", color="#c62828", linewidth=1.6,
                    label=(f"Модель J = {fit_u['c']:.3g}·U$^{{{fit_u['p']:.2f}}}$"))
        target.set_xlabel("U$_{а}$, В", fontsize=12)
        target.set_ylabel("J$_{ф}$, мкА", fontsize=12)
        target.set_title("Яркость от потенциала анода (I$_{л}$ = const)", fontsize=11.5)
        target.grid(True, alpha=0.3)
        target.legend(loc="lower right", fontsize=9)
    return fig_to_base64(fig)


def _pairs(rows, idx_x, idx_y, positive_x=False):
    """Пары (x, y) из строк таблицы с валидными числами."""
    out = []
    for row in rows:
        x = _f(row[idx_x]) if len(row) > idx_x else None
        y = _f(row[idx_y]) if len(row) > idx_y else None
        if x is None or y is None:
            continue
        if positive_x and x <= 0:
            continue
        out.append((x, y))
    out.sort(key=lambda t: t[0])
    return out


def _fit_c_over_sqrtu(pts):
    """МНК для модели D = c/√U (через начало координат)."""
    xs = np.array([1.0 / math.sqrt(u) for u, _ in pts])
    ys = np.array([d for _, d in pts])
    c = float(np.sum(xs * ys) / np.sum(xs * xs))
    return {"c": c, "r2": _r2_of(ys, c * xs)}


def _fit_power(pts):
    """МНК для модели J = c·U^p (логарифмирование)."""
    xs = np.log(np.array([u for u, _ in pts]))
    ys = np.log(np.array([j for _, j in pts]))
    p, b = np.polyfit(xs, ys, 1)
    return {"p": float(p), "c": float(math.exp(float(b))), "r2": _r2_of(ys, p * xs + b)}


def _chart_common(title, x_title, y_title, width=22, height=11):
    ch = ScatterChart()
    ch.title = title
    ch.x_axis.title = x_title
    ch.y_axis.title = y_title
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.height = height
    ch.width = width
    return ch


def _add_series(ws, chart, xcol, ycol, n_pts, title, color, marker=True):
    xref = Reference(ws, min_col=xcol, min_row=2, max_row=n_pts + 1)
    yref = Reference(ws, min_col=ycol, min_row=2, max_row=n_pts + 1)
    s = Series(yref, xref, title=title)
    s.marker = Marker(symbol="circle" if marker else "none", size=6)
    s.graphicalProperties = GraphicalProperties()
    if marker:
        s.graphicalProperties.line.noFill = True
    else:
        s.graphicalProperties.line = LineProperties(solidFill=color, w=20000)
    chart.series.append(s)


def make_mod_xlsx(groups, steep, um_refs, delta):
    """Excel: модуляционные характеристики (пары столбцов по U_а) + S и Δ."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Модуляционные характеристики"
    headers = ["№"]
    nmax = 0
    for k, g in enumerate(groups, 1):
        headers += [f"U_м {k} (U_а = {g['u_a']:g} В), В", f"I_к {k}, мкА"]
        nmax = max(nmax, len(g["um"]))
    ws.append(headers)
    for i in range(nmax):
        row = [i + 1]
        for g in groups:
            if i < len(g["um"]):
                row += [round(g["um"][i], 3), round(g["ik"][i], 2)]
            else:
                row += ["", ""]
        ws.append(row)
    prow = nmax + 3
    ws.cell(row=prow, column=1, value="U_а, В")
    ws.cell(row=prow, column=2, value="U_м при I_реф, В")
    ws.cell(row=prow, column=3, value="I_к макс, мкА")
    ws.cell(row=prow, column=4, value="Крутизна S, мкА/В")
    for k, (g, um_r) in enumerate(zip(groups, um_refs), 1):
        ws.cell(row=prow + k, column=1, value=g["u_a"])
        ws.cell(row=prow + k, column=2, value=round(um_r, 2) if um_r is not None else "—")
        ws.cell(row=prow + k, column=3, value=round(max(g["ik"]), 1))
        if steep is not None and abs(g["u_a"] - steep["u_nom"]) < 1e-9:
            ws.cell(row=prow + k, column=4, value=round(steep["S"], 3))
    k2 = prow + len(groups) + 2
    ws.cell(row=k2, column=1, value="Проницаемость модулятора Δ = −ΔU_м/ΔU_а при I_к = I_реф")
    ws.cell(row=k2, column=2, value=round(delta["delta"], 4) if delta else "—")
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 22
    chart = _chart_common("Модуляционные характеристики", "U_м, В", "I_к, мкА")
    for k, g in enumerate(groups, 1):
        xcol = 2 + (k - 1) * 2
        _add_series(ws, chart, xcol, xcol + 1, len(g["um"]),
                    f"I_к {k} (U_а = {g['u_a']:g} В)", HEAT_COLORS[(k - 1) % 5][1:])
    ws.add_chart(chart, "G" + str(prow))
    return wb_to_base64(wb)


def make_focus_xlsx(pts, f_list, f_avg, n_f, r_sm):
    """Excel: ток фокусирующей катушки и фокусные расстояния."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Фокусировка"
    ws.append(["№", "U_а, В", "I_ф, мА", "N·Y, ампервиток", "f, см", "I_ф модель, мА"])
    n = len(pts)
    for i, ((ua, ifp), f_i) in enumerate(zip(pts, f_list), 1):
        im = 1000 * 10 * math.sqrt(ua * r_sm) / (n_f * f_avg)
        ws.append([i, ua, ifp, round(n_f * ifp / 1000, 1), round(f_i, 2), round(im, 2)])
    prow = n + 3
    items = [("Средний радиус катушки R, см", round(r_sm, 3)),
             ("Витки фокусирующей катушки N_ф", int(n_f)),
             ("Фокусное расстояние f (среднее), см", round(f_avg, 2))]
    ws.cell(row=prow, column=1, value="Параметр")
    ws.cell(row=prow, column=2, value="Значение")
    for k, (key, val) in enumerate(items, 1):
        ws.cell(row=prow + k, column=1, value=key)
        ws.cell(row=prow + k, column=2, value=val)
    for col in range(1, 7):
        ws.column_dimensions[get_column_letter(col)].width = 20
    chart = _chart_common("Ток фокусирующей катушки", "U_а, В", "I_ф, мА")
    _add_series(ws, chart, 2, 3, n, "I_ф экспер.", "1F4E79")
    _add_series(ws, chart, 2, 6, n, "I_ф модель", "C62828", marker=False)
    ws.add_chart(chart, "H2")
    return wb_to_base64(wb)


def make_defl_i_xlsx(rows, b_v, b_h, eps_v, eps_h, eps_t):
    """Excel: отклонение от тока катушек + чувствительности."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Отклонение-ток"
    ws.append(["№", "I_отк, мА", "D_верт, мм", "D_гор, мм"])
    valid = [r for r in rows
             if _f(r[0] if len(r) > 0 else None) is not None
             and (_f(r[1] if len(r) > 1 else None) is not None
                  or _f(r[2] if len(r) > 2 else None) is not None)]
    n = len(valid)
    for i, r in enumerate(valid, 1):
        ws.append([i, round(_f(r[0]), 3),
                   round(_f(r[1]), 2) if len(r) > 1 and _f(r[1]) is not None else "",
                   round(_f(r[2]), 2) if len(r) > 2 and _f(r[2]) is not None else ""])
    prow = n + 3
    items = [("Наклон D_верт(I), мм/мА", round(b_v, 4) if b_v is not None else "—"),
             ("Наклон D_гор(I), мм/мА", round(b_h, 4) if b_h is not None else "—"),
             ("Чувствительность ε_верт, мм/ампервиток", round(eps_v, 4) if eps_v is not None else "—"),
             ("Чувствительность ε_гор, мм/ампервиток", round(eps_h, 4) if eps_h is not None else "—"),
             ("Теоретическая ε_теор = a·L/(2.7·√U_а·d), мм/ав", round(eps_t, 4))]
    ws.cell(row=prow, column=1, value="Параметр")
    ws.cell(row=prow, column=2, value="Значение")
    for k, (key, val) in enumerate(items, 1):
        ws.cell(row=prow + k, column=1, value=key)
        ws.cell(row=prow + k, column=2, value=val)
    ws.column_dimensions["A"].width = 44
    for col in range(2, 5):
        ws.column_dimensions[get_column_letter(col)].width = 16
    chart = _chart_common("Отклонение от тока катушек", "I_отк, мА", "D, мм")
    _add_series(ws, chart, 2, 3, n, "D_верт", "1F4E79")
    _add_series(ws, chart, 2, 4, n, "D_гор", "C62828")
    ws.add_chart(chart, "F2")
    return wb_to_base64(wb)


def make_defl_u_xlsx(rows, c_v, c_h):
    """Excel: отклонение от потенциала анода + коэффициенты модели."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Отклонение-анод"
    ws.append(["№", "U_а, В", "D_верт, мм", "D_гор, мм",
               "D_верт модель, мм", "D_гор модель, мм"])
    valid = [r for r in rows
             if _f(r[0] if len(r) > 0 else None) is not None
             and (_f(r[1] if len(r) > 1 else None) is not None
                  or _f(r[2] if len(r) > 2 else None) is not None)]
    n = len(valid)
    for i, r in enumerate(valid, 1):
        ua = _f(r[0])
        dv = _f(r[1]) if len(r) > 1 else None
        dh = _f(r[2]) if len(r) > 2 else None
        ws.append([i, ua,
                   round(dv, 2) if dv is not None else "",
                   round(dh, 2) if dh is not None else "",
                   round(c_v / math.sqrt(ua), 2) if dv is not None else "",
                   round(c_h / math.sqrt(ua), 2) if dh is not None else ""])
    prow = n + 3
    items = [("Коэффициент c_верт (D = c/√U_а), мм·В^0.5", round(c_v, 1)),
             ("Коэффициент c_гор (D = c/√U_а), мм·В^0.5", round(c_h, 1))]
    ws.cell(row=prow, column=1, value="Параметр")
    ws.cell(row=prow, column=2, value="Значение")
    for k, (key, val) in enumerate(items, 1):
        ws.cell(row=prow + k, column=1, value=key)
        ws.cell(row=prow + k, column=2, value=val)
    for col in range(1, 7):
        ws.column_dimensions[get_column_letter(col)].width = 20
    chart = _chart_common("Отклонение от потенциала анода", "U_а, В", "D, мм")
    _add_series(ws, chart, 2, 3, n, "D_верт", "1F4E79")
    _add_series(ws, chart, 2, 4, n, "D_гор", "C62828")
    ws.add_chart(chart, "F2")
    return wb_to_base64(wb)


def make_bright_xlsx(br_i, fit_i, br_u, fit_u):
    """Excel: яркость экрана (два листа) + параметры аппроксимаций."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Яркость-ток луча"
    n1 = len(br_i)
    ws1.append(["№", "I_л, мкА", "J_ф, мкА"])
    for i, (il, jf) in enumerate(br_i, 1):
        ws1.append([i, il, jf])
    prow = n1 + 3
    if fit_i is not None:
        ws1.cell(row=prow, column=1, value="Наклон J(I_л), мкА/мкА")
        ws1.cell(row=prow, column=2, value=round(fit_i["k"], 4))
        ws1.cell(row=prow + 1, column=1, value="R²")
        ws1.cell(row=prow + 1, column=2, value=round(fit_i["r2"], 5))
    for col in range(1, 4):
        ws1.column_dimensions[get_column_letter(col)].width = 18

    ws2 = wb.create_sheet("Яркость-анод")
    n2 = len(br_u)
    ws2.append(["№", "U_а, В", "J_ф, мкА", "J_ф модель, мкА"])
    for i, (ua, jf) in enumerate(br_u, 1):
        ws2.append([i, ua, jf,
                    round(fit_u["c"] * ua ** fit_u["p"], 3) if fit_u is not None else ""])
    prow = n2 + 3
    if fit_u is not None:
        items = [("Показатель степени p (J = c·U^p)", round(fit_u["p"], 3)),
                 ("Коэффициент c", round(fit_u["c"], 5)),
                 ("R²", round(fit_u["r2"], 5))]
        ws2.cell(row=prow, column=1, value="Параметр")
        ws2.cell(row=prow, column=2, value="Значение")
        for k, (key, val) in enumerate(items, 1):
            ws2.cell(row=prow + k, column=1, value=key)
            ws2.cell(row=prow + k, column=2, value=val)
    for col in range(1, 5):
        ws2.column_dimensions[get_column_letter(col)].width = 18
    if n1:
        chart = _chart_common("Яркость от тока луча", "I_л, мкА", "J_ф, мкА")
        _add_series(ws1, chart, 2, 3, n1, "J_ф", "1F4E79")
        ws1.add_chart(chart, "E2")
    if n2:
        chart = _chart_common("Яркость от потенциала анода", "U_а, В", "J_ф, мкА")
        _add_series(ws2, chart, 2, 3, n2, "J_ф", "2E7D32")
        if fit_u is not None:
            _add_series(ws2, chart, 2, 4, n2, "J_ф модель", "C62828", marker=False)
        ws2.add_chart(chart, "E2")
    return wb_to_base64(wb)


def build_elt_analysis(mod_out, steep, um_refs, delta, focus_out, f_avg,
                       n_f, r_sm, eps_v, eps_h, eps_t, defl_u_out,
                       br_i_fit, br_u_fit):
    """HTML-панель результатов лабораторной работы № 2."""
    html = '<h4>Модуляционные характеристики (крутизна и проницаемость)</h4>'
    html += ('<table class="res"><tr><th>U<sub>а</sub>, В</th><th>I<sub>к</sub> макс, мкА</th>'
             '<th>U<sub>м</sub> при I<sub>реф</sub>, В</th><th>Крутизна S, мкА/В</th></tr>')
    for g, um_r in zip(mod_out, um_refs):
        s_cell = "—"
        if steep is not None and abs(g["u_a"] - steep["u_nom"]) < 1e-9:
            s_cell = f"{steep['S']:.3f} (R² = {steep['r2']:.4f})"
        html += (f'<tr><td>{g["u_a"]:g}</td><td>{max(g["ik"]):.1f}</td>'
                 f'<td>{"—" if um_r is None else f"{um_r:.2f}"}</td><td>{s_cell}</td></tr>')
    html += '</table>'
    if delta:
        html += (f'<p style="font-size:13.5px;margin:6px 0;"><b>Проницаемость модулятора</b> '
                 f'Δ = −ΔU<sub>м</sub>/ΔU<sub>а</sub> при I<sub>к</sub> = I<sub>реф</sub>: '
                 f'<b>{delta["delta"]:.3f}</b> (R² = {delta["r2"]:.5f}).</p>')
    if steep is not None:
        html += (f'<p style="font-size:13.5px;margin:6px 0;"><b>Крутизна модуляционной '
                 f'характеристики</b> при U<sub>а</sub> = {steep["u_nom"]:g} В: '
                 f'<b>S = {steep["S"]:.2f} мкА/В</b>.</p>')

    html += '<h4 style="margin-top:14px;">Магнитная фокусирующая линза</h4>'
    if focus_out and f_avg:
        html += ('<table class="res"><tr><th>U<sub>а</sub>, В</th><th>I<sub>ф</sub>, мА</th>'
                 '<th>N·Y, ампервиток</th><th>f, см</th></tr>')
        for (ua, ifp), f_i in focus_out:
            html += (f'<tr><td>{ua:g}</td><td>{ifp:g}</td>'
                     f'<td>{n_f * ifp / 1000:.1f}</td>'
                     f'<td>{f_i:.2f}</td></tr>')
        html += '</table>'
        html += (f'<p style="font-size:13.5px;margin:6px 0;"><b>Фокусное расстояние линзы</b> '
                 f'(среднее): <b>f = {f_avg:.2f} см</b>. Ток фокусировки растёт пропорционально '
                 f'√U<sub>а</sub>, что согласуется с формулой N·Y = 10·√(U<sub>а</sub>·R)/f.</p>')

    html += '<h4 style="margin-top:14px;">Чувствительность отклоняющей системы</h4>'
    sens = []
    if eps_v is not None:
        sens.append(f'верт. <b>ε = {eps_v:.3f} мм/ампервиток</b>')
    if eps_h is not None:
        sens.append(f'гор. <b>ε = {eps_h:.3f} мм/ампервиток</b>')
    if sens:
        html += (f'<p style="font-size:13.5px;margin:6px 0;">Чувствительность по наклону '
                 f'D(I<sub>отк</sub>): {", ".join(sens)}; теоретическая оценка по ф. (2.5): '
                 f'ε<sub>теор</sub> = a·L/(2.7·√U<sub>а</sub>·d) = <b>{eps_t:.3f} мм/ав</b> '
                 f'при U<sub>а</sub> = номинальном.</p>')
    if defl_u_out:
        html += (f'<p style="font-size:13.5px;margin:6px 0;">Зависимость D(U<sub>а</sub>) при '
                 f'постоянном токе катушек аппроксимируется моделью D = c/√U<sub>а</sub>: '
                 f'c<sub>верт</sub> = {defl_u_out["c_v"]:.1f}, c<sub>гор</sub> = {defl_u_out["c_h"]:.1f} '
                 f'мм·В<sup>0.5</sup> (R² = {defl_u_out["r2_v"]:.4f} и {defl_u_out["r2_h"]:.4f}) — '
                 f'отклонение уменьшается с ростом U<sub>а</sub> как 1/√U<sub>а</sub>.</p>')

    html += '<h4 style="margin-top:14px;">Яркость экрана</h4>'
    if br_i_fit:
        html += (f'<p style="font-size:13.5px;margin:6px 0;">Яркость растёт с током луча '
                 f'практически линейно: J = {br_i_fit["k"]:.4f}·I<sub>л</sub> + {br_i_fit["b"]:.2f} '
                 f'(R² = {br_i_fit["r2"]:.4f}).</p>')
    if br_u_fit:
        html += (f'<p style="font-size:13.5px;margin:6px 0;">Яркость от потенциала анода: '
                 f'J = {br_u_fit["c"]:.3g}·U<sub>а</sub><sup>{br_u_fit["p"]:.2f}</sup> '
                 f'(R² = {br_u_fit["r2"]:.4f}) — рост заметно быстрее линейного, так как '
                 f'увеличивается и энергия, и число электронов пучка.</p>')

    html += '<h4 style="margin-top:14px;">Выводы</h4><ul style="font-size:13.5px;line-height:1.6;">'
    if mod_out:
        html += ('<li>Модуляционные характеристики I<sub>к</sub>(U<sub>м</sub>) при разных U<sub>а</sub> '
                 'смещаются: с ростом U<sub>а</sub> увеличивается и максимальный ток катода, '
                 'и запирающее напряжение (по модулю), поскольку действующее напряжение '
                 'U<sub>д</sub> = U<sub>м</sub> + Δ·U<sub>а</sub>.</li>')
    if delta:
        html += (f'<li>Проницаемость модулятора Δ = {delta["delta"]:.2f} показывает, во сколько раз '
                 f'действие U<sub>м</sub> на ток сильнее действия U<sub>а</sub>; поле анода проникает '
                 f'к катоду ослабленным (через отверстие модулятора).</li>')
    if f_avg:
        html += (f'<li>Магнитная линза (N<sub>ф</sub> = {n_f:g} витков, средний диаметр катушки '
                 f'{2 * r_sm:g} мм) при номинальном режиме имеет фокусное расстояние '
                 f'f ≈ {f_avg:.1f} см; ток фокусировки ∝ √U<sub>а</sub>.</li>')
    if eps_v is not None or eps_h is not None:
        html += ('<li>Отклонение пятна линейно по току катушек; чувствительность ε в мм/ампервиток '
                 'удобна для сравнения пар катушек и согласуется с оценкой по ф. (2.5); '
                 'с ростом U<sub>а</sub> отклонение падает как 1/√U<sub>а</sub> — та же '
                 'зависимость, что и в ф. (2.5).</li>')
    html += ('<li>Яркость свечения экрана растёт с током луча и с U<sub>а</sub>; '
             'магнитная отклоняющая система имеет более низкий частотный диапазон, '
             'чем электростатическая, из-за индуктивности катушек.</li>')
    html += '</ul>'
    return html




Vipe_STUB_TITLES = {
    3: "Исследование фотоэлектронных приборов",
}


@bp.route("/vipe/lab/<int:lab_id>")
def vipe_lab(lab_id):
    if lab_id == 1:
        return LAB_TEE_HTML
    if lab_id == 2:
        return LAB_ELT_HTML
    title = Vipe_STUB_TITLES.get(lab_id, f"Лабораторная работа № {lab_id}")
    return STUB_HTML.replace("__TITLE__", title)


@bp.route("/api/vipe/tee/process", methods=["POST"])
def api_vipe_tee_process():
    try:
        d = request.get_json(force=True)
        d_k = float(d["d_k"]) / 1000.0    # мм → м
        l_k = float(d["l_k"]) / 1000.0
        l_a = float(d["l_a"]) / 1000.0
        r_k = float(d["r_k"]) / 1000.0
        r_a = float(d["r_a"]) / 1000.0
        if min(d_k, l_k, l_a, r_k) <= 0 or r_a <= r_k:
            return jsonify({"error": "Проверьте геометрию: все размеры должны быть "
                                     "положительными, r_а > r_к."}), 400
        s_emit = math.pi * d_k * l_a
        ln_ratio = math.log(r_a / r_k)

        heats_in = d.get("heats", [])
        if not heats_in:
            return jsonify({"error": "Нет данных: заполните хотя бы один накал."}), 400

        proc = []
        rich_pts = []
        for idx, h in enumerate(heats_in, 1):
            i_n = _f(h.get("i_n"))
            u_n = _f(h.get("u_n"))
            pts = []
            for row in h.get("rows", []):
                ua = _f(row[0] if len(row) > 0 else None)
                ia = _f(row[1] if len(row) > 1 else None)
                if ua is not None and ia is not None:
                    pts.append((ua, ia))
            label = f"I_н = {i_n:g} А" if i_n is not None else f"накал {idx}"
            rec = {"num": idx, "i_n": i_n, "u_n": u_n, "label": label}
            if len(pts) < 4:
                rec["error"] = "недостаточно точек ВАХ (нужно ≥ 4)"
                proc.append(rec)
                continue
            pts.sort(key=lambda t: t[0])
            u_arr = [p[0] for p in pts]
            i_arr = [p[1] for p in pts]
            fit = piecewise_linear_fit(u_arr, i_arr)
            if "error" in fit:
                rec["error"] = fit["error"]
                proc.append(rec)
                continue
            x_int = fit.get("i_por")
            if x_int is None:
                rec["error"] = "не удалось определить перегиб (участки ВАХ параллельны)"
                proc.append(rec)
                continue
            u_knee = min(max(x_int, float(u_arr[0])), float(u_arr[-1]))
            i_knee = fit["k_sv"] * u_knee + fit["b_sv"]
            rec.update({"u": u_arr, "i": i_arr,
                        "u_knee": u_knee, "i_knee": i_knee,
                        "out_of_range": bool(x_int < u_arr[0] or x_int > u_arr[-1])})
            if i_knee <= 0:
                rec["error"] = "неположительный ток эмиссии — проверьте данные"
                proc.append(rec)
                continue
            rec["j"] = i_knee * 1e-3 / s_emit          # А/м²
            if i_n is not None and u_n is not None and i_n > 0 and u_n > 0:
                p_ud = i_n * u_n / (math.pi * l_k * d_k) / 1e4   # Вт/см²
                rec["P_ud"] = p_ud
                rec["I_unit"] = i_n / (d_k * 100.0) ** 1.5       # ток единичного катода
                rec["T"] = tungsten_T_from_P(p_ud)
                rich_pts.append((rec["T"], rec["j"], label))
            # прямая Шоттки: точки правее перегиба
            sx, sy, sE, sj = [], [], [], []
            for ua, ia in zip(u_arr, i_arr):
                if ua > u_knee and ia > 0:
                    E = ua / (r_k * ln_ratio)
                    j_pt = ia * 1e-3 / s_emit
                    sx.append(math.sqrt(E))
                    sy.append(math.log(j_pt))
                    sE.append(E)
                    sj.append(j_pt)
            rec["sch"] = {"x": sx, "y": sy, "E": sE, "j": sj}
            if len(sx) >= 3:
                beta, b0 = np.polyfit(np.array(sx), np.array(sy), 1)
                yfit = float(beta) * np.array(sx) + float(b0)
                rec["sch"].update({"beta": float(beta), "b0": float(b0),
                                   "r2": _r2_of(sy, yfit),
                                   "T_sh": C_SCHOTTKY / float(beta)})
            proc.append(rec)

        rich = None
        if len(rich_pts) >= 2:
            xs = np.array([1.0 / t for t, _, _ in rich_pts])
            ys = np.array([math.log(j / (t * t)) for t, j, _ in rich_pts])
            slope, inter = np.polyfit(xs, ys, 1)
            rich = {"slope": float(slope), "intercept": float(inter),
                    "r2": _r2_of(ys, slope * xs + inter),
                    "phi": abs(float(slope)) * K_BOLTZ / E_CHARGE,  # эВ; наклон = −eφ/k
                    "A": math.exp(float(inter)),
                    "n": len(rich_pts)}

        resp = {"analysis_html": build_tee_analysis(proc, rich, rich_pts)}
        ok = [r for r in proc if "error" not in r]
        if ok:
            resp["g1_image"] = make_vax_plot(ok)
            resp["g1_xlsx"] = make_vax_xlsx(ok)
        if rich:
            resp["g2_image"] = make_richardson_plot(rich_pts, rich)
            resp["g2_xlsx"] = make_richardson_xlsx(rich_pts, rich)
        g3 = [r for r in proc if "error" not in r and len(r.get("sch", {}).get("x", [])) >= 3]
        if g3:
            resp["g3_image"] = make_schottky_plot(g3)
            resp["g3_xlsx"] = make_schottky_xlsx(g3)
        if ok and rich:
            resp["g4_image"] = make_jT_plot(rich_pts, rich)
            resp["g4_xlsx"] = make_jT_xlsx(rich_pts, rich)
        return jsonify(resp)
    except (KeyError, ValueError, TypeError) as e:
        return jsonify({"error": f"Ошибка входных данных: {e}. Проверьте параметры "
                                 f"и таблицы."}), 400
    except Exception as e:
        return jsonify({"error": f"Ошибка: {e}"}), 400


@bp.route("/api/vipe/elt/process", methods=["POST"])
def api_vipe_elt_process():
    try:
        d = request.get_json(force=True)
        ua_list = [_f(d["u_a"][k]) for k in range(3)]
        u_nom = _f(d["u_nom"])
        i_ref = _f(d["i_ref"])
        n_f = _f(d["n_f"])
        r_sm = _f(d["d_f"]) / 20.0            # диаметр в мм → средний радиус в см
        n_o = _f(d["n_o"])
        a_mm = _f(d["a_mm"])
        d_mm = _f(d["d_mm"])
        l_mm = _f(d["l_mm"])
        if min(u_nom, n_f, r_sm, n_o, a_mm, d_mm, l_mm) <= 0:
            return jsonify({"error": "Все параметры установки должны быть положительными."}), 400
        if any(u is None or u <= 0 for u in ua_list):
            return jsonify({"error": "Задайте три положительных значения U_а для "
                                     "модуляционных характеристик."}), 400

        resp = {}

        # --- 1. Модуляционные характеристики ---
        mod_out = []
        um_refs = []
        for k in range(3):
            pairs = []
            for um, ik in zip(d["mod"][k].get("um", []), d["mod"][k].get("ik", [])):
                um_v, ik_v = _f(um), _f(ik)
                if um_v is not None and ik_v is not None:
                    pairs.append((um_v, ik_v))
            pairs.sort(key=lambda t: t[0])
            if len(pairs) < 3:
                mod_out.append({"u_a": ua_list[k], "um": [], "ik": [],
                                "error": "недостаточно точек (нужно ≥ 3)"})
                um_refs.append(None)
                continue
            mod_out.append({"u_a": ua_list[k],
                            "um": [p[0] for p in pairs],
                            "ik": [p[1] for p in pairs]})
            um_refs.append(_um_at_current(pairs, i_ref) if i_ref else None)

        steep = None
        if i_ref:
            k_nom = min(range(3), key=lambda k: abs(ua_list[k] - u_nom))
            g_nom = mod_out[k_nom]
            if "error" not in g_nom:
                steep = _steepness(list(zip(g_nom["um"], g_nom["ik"])))
                if steep is not None:
                    steep["u_nom"] = ua_list[k_nom]

        delta = None
        um_pairs = [(ua_list[k], um_refs[k]) for k in range(3) if um_refs[k] is not None]
        if len(um_pairs) >= 2:
            xs = np.array([p[0] for p in um_pairs])
            ys = np.array([p[1] for p in um_pairs])
            slope, inter = np.polyfit(xs, ys, 1)
            delta = {"delta": -float(slope), "r2": _r2_of(ys, slope * xs + inter),
                     "intercept": float(inter)}

        if any("error" not in g for g in mod_out):
            ok_groups = [g for g in mod_out if "error" not in g]
            resp["g1_image"] = make_mod_plot(ok_groups, i_ref, um_refs)
            resp["g1_xlsx"] = make_mod_xlsx(
                ok_groups, steep,
                [um_refs[k] for k in range(3) if "error" not in mod_out[k]],
                delta)

        # --- 2. Фокусировка ---
        focus_pts = _pairs(d.get("focus", []), 0, 1)
        focus_out, f_list, f_avg = [], [], None
        if len(focus_pts) >= 2:
            for ua, ifp in focus_pts:
                if ifp > 0:
                    f_i = 10 * math.sqrt(ua * r_sm) / (n_f * ifp / 1000.0)  # см, ф. (2.3)
                    f_list.append(f_i)
                    focus_out.append(((ua, ifp), f_i))
            if f_list:
                f_avg = sum(f_list) / len(f_list)
                resp["g2_image"] = make_focus_plot(focus_pts, f_avg, n_f, r_sm)
                resp["g2_xlsx"] = make_focus_xlsx(focus_pts, f_list, f_avg, n_f, r_sm)

        # --- 3. Отклонение от тока катушек ---
        di_v = _pairs(d.get("defl_i", []), 0, 1)
        di_h = _pairs(d.get("defl_i", []), 0, 2)
        b_v, b_h, eps_v, eps_h = None, None, None, None
        if len(di_v) >= 2:
            xs = np.array([p[0] for p in di_v])
            ys = np.array([p[1] for p in di_v])
            b_v, bv0 = np.polyfit(xs, ys, 1)
            b_v = float(b_v)
            eps_v = b_v * 1000.0 / n_o
        if len(di_h) >= 2:
            xs = np.array([p[0] for p in di_h])
            ys = np.array([p[1] for p in di_h])
            b_h, bh0 = np.polyfit(xs, ys, 1)
            b_h = float(b_h)
            eps_h = b_h * 1000.0 / n_o
        eps_t = a_mm * l_mm / (2.7 * math.sqrt(u_nom) * d_mm)
        if len(di_v) >= 2 or len(di_h) >= 2:
            resp["g3_image"] = make_defl_i_plot(di_v, di_h, b_v, b_h,
                                                eps_v, eps_h, eps_t)
            resp["g3_xlsx"] = make_defl_i_xlsx(d.get("defl_i", []), b_v, b_h,
                                               eps_v, eps_h, eps_t)

        # --- 4. Отклонение от потенциала анода ---
        du_v = _pairs(d.get("defl_u", []), 0, 1, positive_x=True)
        du_h = _pairs(d.get("defl_u", []), 0, 2, positive_x=True)
        fit_v = _fit_c_over_sqrtu(du_v) if len(du_v) >= 2 else None
        fit_h = _fit_c_over_sqrtu(du_h) if len(du_h) >= 2 else None
        defl_u_out = None
        if fit_v is not None or fit_h is not None:
            c_v = fit_v["c"] if fit_v else 0.0
            c_h = fit_h["c"] if fit_h else 0.0
            resp["g4_image"] = make_defl_u_plot(du_v, du_h, c_v, c_h)
            resp["g4_xlsx"] = make_defl_u_xlsx(d.get("defl_u", []), c_v, c_h)
            defl_u_out = {"c_v": c_v, "c_h": c_h,
                          "r2_v": fit_v["r2"] if fit_v else None,
                          "r2_h": fit_h["r2"] if fit_h else None}

        # --- 5. Яркость экрана ---
        br_i_pts = _pairs(d.get("br_i", []), 0, 1)
        br_u_pts = _pairs(d.get("br_u", []), 0, 1, positive_x=True)
        br_i_fit, br_u_fit = None, None
        if len(br_i_pts) >= 2:
            xs = np.array([p[0] for p in br_i_pts])
            ys = np.array([p[1] for p in br_i_pts])
            k_il, b_il = np.polyfit(xs, ys, 1)
            br_i_fit = {"k": float(k_il), "b": float(b_il),
                        "r2": _r2_of(ys, k_il * xs + b_il)}
        br_u_valid = [(u, j) for u, j in br_u_pts if j > 0]
        if len(br_u_valid) >= 2:
            br_u_fit = _fit_power(br_u_valid)
        if br_i_fit is not None or br_u_fit is not None:
            resp["g5_image"] = make_bright_plot(br_i_pts, br_i_fit, br_u_pts, br_u_fit)
            resp["g5_xlsx"] = make_bright_xlsx(br_i_pts, br_i_fit, br_u_pts, br_u_fit)

        resp["analysis_html"] = build_elt_analysis(
            mod_out, steep, um_refs, delta, focus_out, f_avg, n_f, r_sm,
            eps_v, eps_h, eps_t, defl_u_out, br_i_fit, br_u_fit)
        return jsonify(resp)
    except (KeyError, ValueError, TypeError) as e:
        return jsonify({"error": f"Ошибка входных данных: {e}. Проверьте параметры "
                                 f"и таблицы."}), 400
    except Exception as e:
        return jsonify({"error": f"Ошибка: {e}"}), 400
