# -*- coding: utf-8 -*-
"""Курс «Квантовая и оптическая электроника» — лабораторные работы № 1-3.

Лабораторная работа № 1 — «Исследование основных параметров
полупроводникового лазера»;
лабораторная работа № 2 — «Определение электрооптических параметров
анизотропных кристаллов» (эффект Поккельса, LiNbO3);
лабораторная работа № 3 — «Полупроводниковые детекторы оптического
излучения» (фотодиод ФД-24К).

Маршруты: /kioe/lab/<n>, /api/kioe/lab1/..., /api/kioe/lab2/process,
/api/kioe/lab3/process.
"""

from flask import Blueprint, request, jsonify

import numpy as np
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.utils import get_column_letter

from common import fig_to_base64, wb_to_base64, piecewise_linear_fit, work_block, fmt_n, fmt_g

bp = Blueprint("kioe", __name__)


# ===========================================================================
#  HTML-страницы лабораторных работ
# ===========================================================================


LAB1_HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Лабораторная работа № 1 — Полупроводниковый лазер</title>
<style>
  :root { --accent:#1f4e79; --bg:#f4f6f9; --panel:#fff; --border:#d8dee9; --ok:#2e7d32; --warn:#c62828; }
  * { box-sizing:border-box; }
  body { margin:0; font-family:'Segoe UI', system-ui, Arial, sans-serif; background:var(--bg); color:#222; line-height:1.5; }
  header { background:var(--accent); color:#fff; padding:22px 20px; }
  header .wrap { max-width:1100px; margin:0 auto; display:flex; justify-content:space-between; align-items:center; gap:16px; flex-wrap:wrap; }
  header h1 { margin:0; font-size:21px; }
  header a { color:#fff; text-decoration:none; font-size:13.5px; border:1px solid rgba(255,255,255,.5); padding:6px 14px; border-radius:6px; }
  header a:hover { background:rgba(255,255,255,.15); }
  main { max-width:1100px; margin:0 auto; padding:24px 16px 60px; }
  .lead { font-size:14.5px; color:#444; background:#fff; border:1px solid var(--border); border-left:4px solid var(--accent); padding:14px 18px; border-radius:6px; margin-bottom:22px; }
  .lead b { color:var(--accent); }
  h2.sec { font-size:19px; color:var(--accent); border-bottom:2px solid var(--accent); padding-bottom:6px; margin:30px 0 14px; }
  h3.sub { font-size:15.5px; margin:18px 0 10px; color:#333; }
  .r1-box { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:16px 18px; margin-bottom:22px; display:flex; align-items:center; gap:14px; flex-wrap:wrap; }
  .r1-box label { font-weight:600; font-size:14.5px; }
  .r1-box input { font-size:15px; padding:7px 10px; border:1px solid #b9c2d0; border-radius:5px; width:120px; }
  .r1-box .hint { color:#777; font-size:13px; }
  .table-wrap { overflow-x:auto; }
  table.data { border-collapse:collapse; width:100%; background:var(--panel); font-size:13.5px; min-width:520px; }
  table.data caption { caption-side:top; text-align:left; font-weight:600; font-size:14px; padding:8px 10px; color:#333; background:#eef2f7; border:1px solid var(--border); border-bottom:none; border-radius:6px 6px 0 0; }
  table.data th, table.data td { border:1px solid var(--border); padding:6px 8px; text-align:center; }
  table.data th { background:#f0f4f9; font-weight:600; }
  table.data td.num-col { color:#888; width:42px; background:#fafbfc; }
  table.data td.calc { background:#fffdf3; font-weight:600; color:#7a5d00; }
  table.data input { width:100%; border:1px solid transparent; background:transparent; text-align:center; font-size:13.5px; padding:5px 4px; border-radius:4px; font-family:inherit; }
  table.data input:focus { border-color:var(--accent); background:#fff; outline:none; }
  .toolbar { display:flex; gap:10px; flex-wrap:wrap; margin:14px 0; align-items:center; }
  button { font-family:inherit; font-size:14px; padding:9px 16px; border:none; border-radius:6px; cursor:pointer; background:var(--accent); color:#fff; font-weight:600; transition:background .12s, transform .05s; }
  button:hover { background:#173d61; }
  button:active { transform:translateY(1px); }
  button.sec { background:#5a6b80; }
  button.sec:hover { background:#475566; }
  button.ghost { background:#fff; color:var(--accent); border:1px solid var(--accent); }
  button.ghost:hover { background:#eef3f9; }
  .split-box { display:flex; align-items:center; gap:6px; font-size:13px; color:#555; margin-left:auto; }
  .split-box input { padding:6px 8px; border:1px solid #b9c2d0; border-radius:5px; font-size:13px; }
  .calc-panel { background:#fffbea; border:1px solid #e6d98a; border-left:4px solid #d4a72c; border-radius:6px; padding:14px 18px; margin:14px 0; display:none; }
  .calc-panel .formula { font-weight:600; margin:0 0 6px; color:#6b5400; }
  .calc-panel .formula-eq { font-family:'Cambria Math', 'Times New Roman', serif; font-size:16px; margin:6px 0; }
  .work-title { font-weight:600; margin:12px 0 6px; color:#6b5400; }
  .work { border-collapse:collapse; width:100%; font-size:13.5px; background:#fff; }
  .work th, .work td { border:1px solid #e2d9b0; padding:5px 8px; text-align:left; vertical-align:middle; }
  .work th { background:#f3ecd2; font-weight:600; }
  .work td.wn { width:30px; text-align:center; color:#555; }
  .work td.weq { font-family:'Cambria Math', 'Times New Roman', serif; font-size:14.5px; }
  .work td.wres { font-weight:600; white-space:nowrap; font-family:'Cambria Math', 'Times New Roman', serif; }
  .work-scroll { max-height:430px; overflow:auto; border:1px solid #e6d98a; }
  .calc-panel ol { margin:6px 0 6px 22px; padding:0; }
  .calc-panel ol li { margin:3px 0; font-family:'Cambria Math', 'Times New Roman', serif; font-size:14.5px; }
  .calc-panel .note { color:#777; font-size:13px; margin:8px 0 0; }
  .result-panel { background:#eef7ee; border:1px solid #bfe3bf; border-left:4px solid var(--ok); border-radius:6px; padding:14px 18px; margin:14px 0; display:none; }
  .result-panel h4 { margin:0 0 10px; color:var(--ok); font-size:15px; }
  table.res { border-collapse:collapse; font-size:13.5px; margin:6px 0; }
  table.res th, table.res td { border:1px solid #cfe6cf; padding:6px 12px; text-align:left; }
  table.res th { background:#e3f3e3; }
  .plot-wrap { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:14px; margin:14px 0; display:none; text-align:center; }
  .plot-wrap img { max-width:100%; height:auto; border:1px solid #e0e0e0; border-radius:4px; }
  .plot-wrap .dl { display:inline-block; margin-top:8px; font-size:13px; color:var(--accent); }
  .plot-wrap .dl-stack { margin-top:4px; }
  .err { color:var(--warn); font-weight:600; }
  .muted { color:#777; font-size:13px; }
  sub { font-size:75%; }
</style>
</head>
<body>
<header>
  <div class="wrap">
    <h1>Лабораторная работа № 1. Исследование основных параметров полупроводникового лазера</h1>
    <a href="/" target="_blank">&larr; к списку работ</a>
  </div>
</header>
<main>
  <div class="lead">
    Объект исследования — полупроводниковый инжекционный лазер (модуль KLM&#8209;650). В работе измеряются
    зависимости <b>интенсивности излучения</b> и <b>степени линейной поляризации</b> от тока через p&#8209;n переход.
    Ток через p&#8209;n переход измеряется косвенно: вольтметром М1 фиксируется падение напряжения
    <b>U<sub>ПЛ</sub></b> на резисторе <b>R<sub>1</sub></b>, после чего ток рассчитывается по закону Ома.
  </div>

  <div class="r1-box">
    <label for="r1">Сопротивление R<sub>1</sub>, Ом:</label>
    <input type="number" id="r1" step="any" min="0" placeholder="напр. 100" value="100">
    <span class="hint">Измеряется мультиметром М1 (шаг 8 методики). Используется в обеих таблицах.</span>
  </div>

  <!-- =================== ЧАСТЬ 1. ИНТЕНСИВНОСТЬ =================== -->
  <h2 class="sec">Часть 1. Зависимость интенсивности излучения от тока через p&#8209;n переход</h2>

  <div class="table-wrap">
    <table class="data" id="t1">
      <caption>Таблица 1.1 — Данные для построения зависимости интенсивности излучения лазера от тока, протекающего через p&#8209;n переход</caption>
      <thead>
        <tr><th>№</th><th>Прибор М1, U<sub>ПЛ</sub>, мВ</th><th>I<sub>pn</sub>, мА (расч.)</th><th>Прибор М2, I<sub>ФП</sub>, мкА</th></tr>
      </thead>
      <tbody id="t1-body"></tbody>
    </table>
  </div>

  <div class="toolbar">
    <button onclick="calcIpn1()">Рассчитать I<sub>pn</sub>, мА</button>
    <button class="ghost" onclick="buildTable1(prompt('Количество точек в таблице 1.1?', 18) || 18)">Изменить число строк</button>
  </div>

  <div class="calc-panel" id="calc1"></div>
  <div class="calc-panel" id="work1"></div>

  <div class="toolbar">
    <button onclick="plotIntensity()">Построить зависимость I<sub>ФП</sub>(I<sub>pn</sub>)</button>
    <div class="split-box">
      <label for="laser1">Ввести значения — точки красной линии (№ строк):</label>
      <input type="text" id="laser1" style="width:210px;" placeholder="напр. 10 или 10-18 или 10,12,14-18">
    </div>
  </div>
  <p class="hint" style="margin:-4px 0 16px;font-size:12.5px;color:#777;">Формат: «10» — точки с № 10 до конца таблицы (первая точка лазерного режима); «10-18» — диапазон; «10,12,14-18» — список строк. Можно задать точки координатами парами «x;y» через пробел: <i>2.45;70 2.6;209 2.8;486</i>. Пусто — автоматический поиск точки раздела. Красная линия строится по МНК из выбранных точек, параметры пересчитываются; на графике подписаны значения каждой точки.</p>

  <div class="result-panel" id="res1"></div>
  <div class="calc-panel" id="work1-res"></div>
  <div class="plot-wrap" id="plot1">
    <img id="plot1-img" alt="График I_ФП(I_pn)">
    <div class="dl-stack">
      <a class="dl" id="plot1-dl-png" download="I_FP(I_pn).png">Скачать график PNG</a>
    </div>
    <div class="dl-stack">
      <a class="dl" id="plot1-dl-xlsx" download="I_FP(I_pn).xlsx">Скачать график EXCEL</a>
    </div>
  </div>

  <!-- =================== ЧАСТЬ 2. ПОЛЯРИЗАЦИЯ =================== -->
  <h2 class="sec">Часть 2. Зависимость степени поляризации излучения от тока через p&#8209;n переход</h2>

  <div class="table-wrap">
    <table class="data" id="t2">
      <caption>Таблица 1.2 — Данные для получения зависимости степени линейной поляризации P от тока через p&#8209;n переход</caption>
      <thead>
        <tr>
          <th rowspan="2">№</th>
          <th rowspan="2">Прибор М1, U<sub>ПЛ</sub>, мВ</th>
          <th rowspan="2">I<sub>pn</sub>, мА (расч.)</th>
          <th colspan="2">Прибор М2, I<sub>ФП</sub>, мкА</th>
          <th rowspan="2">P = (I<sub>max</sub>−I<sub>min</sub>) / (I<sub>max</sub>+I<sub>min</sub>) (расч.)</th>
        </tr>
        <tr><th>I<sub>max</sub></th><th>I<sub>min</sub></th></tr>
      </thead>
      <tbody id="t2-body"></tbody>
    </table>
  </div>

  <div class="toolbar">
    <button onclick="calcIpn2()">Рассчитать I<sub>pn</sub>, мА</button>
    <button onclick="calcP2()">Рассчитать P = (I<sub>max</sub> − I<sub>min</sub>) / (I<sub>max</sub> + I<sub>min</sub>)</button>
    <button class="ghost" onclick="buildTable2(prompt('Количество точек в таблице 1.2?', 10) || 10)">Изменить число строк</button>
  </div>

  <div class="calc-panel" id="calc2-ipn"></div>
  <div class="calc-panel" id="work2-ipn"></div>
  <div class="calc-panel" id="calc2-p"></div>
  <div class="calc-panel" id="work2-p"></div>

  <div class="toolbar">
    <button onclick="plotPolar()">Построить зависимость P(I<sub>pn</sub>)</button>
  </div>
  <p class="hint" style="margin:-4px 0 16px;font-size:12.5px;color:#777;">На графике отмечаются положение I<sub>пор</sub>, область светодиодного режима (P мал) и область лазерного режима (P возрастает). I<sub>пор</sub> берётся из последнего построения части 1 — сначала постройте зависимость I<sub>ФП</sub>(I<sub>pn</sub>); если она ещё не построена, график будет без разметки режимов.</p>

  <div class="plot-wrap" id="plot2">
    <img id="plot2-img" alt="График P(I_pn)">
    <div class="dl-stack">
      <a class="dl" id="plot2-dl-png" download="P(I_pn).png">Скачать график PNG</a>
    </div>
    <div class="dl-stack">
      <a class="dl" id="plot2-dl-xlsx" download="P(I_pn).xlsx">Скачать график EXCEL</a>
    </div>
  </div>

  <p class="muted">Совет: значение R<sub>1</sub> задаётся один раз вверху страницы и используется в обеих частях работы.</p>
</main>

<script>
const R1_ID = 'r1';

function getR1() {
  const v = parseFloat(document.getElementById(R1_ID).value);
  return (isNaN(v) || v <= 0) ? null : v;
}

// ---- Таблица 1.1 ----
function buildTable1(n) {
  n = Math.max(1, Math.min(60, parseInt(n) || 18));
  const tbody = document.getElementById('t1-body');
  tbody.innerHTML = '';
  for (let i = 1; i <= n; i++) {
    const tr = document.createElement('tr');
    tr.innerHTML =
      '<td class="num-col">' + i + '</td>' +
      '<td><input type="number" step="any" class="u" placeholder="—"></td>' +
      '<td class="calc ipn">—</td>' +
      '<td><input type="number" step="any" class="ifp" placeholder="—"></td>';
    tbody.appendChild(tr);
  }
}

function readCol1(cls) {
  const out = [];
  document.querySelectorAll('#t1-body tr').forEach(tr => {
    const v = tr.querySelector('.' + cls).value;
    out.push(v);
  });
  return out;
}

async function calcIpn1() {
  const r1 = getR1();
  if (r1 === null) { alert('Введите положительное значение R1 (Ом).'); return; }
  const uInputs = document.querySelectorAll('#t1-body .u');
  const us = [];
  uInputs.forEach(inp => { if (inp.value !== '') us.push(inp.value); });
  if (us.length === 0) { alert('Введите хотя бы одно значение U_ПЛ (мВ).'); return; }
  let res;
  try {
    res = await fetch('/api/kioe/lab1/calc_ipn', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({r1: r1, u: us})
    });
  } catch(e) { alert('Ошибка сети: ' + e); return; }
  const data = await res.json();
  if (data.error) { alert(data.error); return; }
  let idx = 0;
  document.querySelectorAll('#t1-body tr').forEach(tr => {
    const uInp = tr.querySelector('.u');
    const cell = tr.querySelector('.ipn');
    if (uInp.value !== '') { cell.textContent = data.i_pn[idx++]; }
    else { cell.textContent = '—'; }
  });
  // панель расчёта
  let html = '<p class="formula">Формула расчёта (закон Ома):</p>';
  html += '<p class="formula-eq">I<sub>pn</sub> = U<sub>ПЛ</sub> / R<sub>1</sub>, &nbsp;&nbsp; R<sub>1</sub> = ' + data.r1 + ' Ом</p>';
  html += '<p class="note">U<sub>ПЛ</sub> — в мВ, R<sub>1</sub> — в Ом, тогда I<sub>pn</sub> — в мА (мВ/Ом = мА).</p>';
  html += '<p class="note">Значения занесены в столбец «I<sub>pn</sub>, мА». Все расчёты по каждой точке — в блоке «Расчетная часть» ниже.</p>';
  const panel = document.getElementById('calc1');
  panel.innerHTML = html;
  panel.style.display = 'block';
  const wp1 = document.getElementById('work1');
  wp1.innerHTML = data.work_html;
  wp1.style.display = 'block';
}

let lastIPor = null;  // пороговый ток из последнего построения части 1

async function plotIntensity() {
  const r1 = getR1();
  if (r1 === null) { alert('Введите положительное значение R1 (Ом).'); return; }
  const pairs = [];
  document.querySelectorAll('#t1-body tr').forEach((tr, i) => {
    const u = tr.querySelector('.u').value;
    const ifp = tr.querySelector('.ifp').value;
    const ipn = tr.querySelector('.ipn').textContent;
    if (u !== '' && ifp !== '' && ipn !== '—' && ipn !== '') {
      pairs.push({row: i + 1, ipn: parseFloat(ipn), ifp: parseFloat(ifp)});
    }
  });
  if (pairs.length < 4) {
    alert('Для построения графика и аппроксимации нужно минимум 4 точки с заполненными U_ПЛ, I_pn и I_ФП. Сначала нажмите «Рассчитать I_pn».');
    return;
  }
  const laser = document.getElementById('laser1').value.trim();
  let res;
  try {
    res = await fetch('/api/kioe/lab1/plot_intensity', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({i_pn: pairs.map(p=>p.ipn), i_fp: pairs.map(p=>p.ifp),
                            rows: pairs.map(p=>p.row), laser: laser})
    });
  } catch(e) { alert('Ошибка сети: ' + e); return; }
  const data = await res.json();
  if (data.error) { alert(data.error); return; }
  lastIPor = (data.i_por !== undefined && data.i_por !== null) ? data.i_por : lastIPor;
  const img1 = 'data:image/png;base64,' + data.image;
  document.getElementById('plot1-img').src = img1;
  document.getElementById('plot1-dl-png').href = img1;
  if (data.xlsx) document.getElementById('plot1-dl-xlsx').href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + data.xlsx;
  document.getElementById('plot1').style.display = 'block';
  let r = '<h4>Результаты обработки</h4>';
  r += '<table class="res"><tr><th>Параметр</th><th>Значение</th></tr>';
  r += '<tr><td>k<sub>св</sub> — коэффициент (светодиодный режим)</td><td>' + data.k_sv + ' мкА/мА</td></tr>';
  r += '<tr><td>b<sub>св</sub> — смещение (светодиодный режим)</td><td>' + data.b_sv + ' мкА</td></tr>';
  r += '<tr><td>k<sub>лаз</sub> — коэффициент (лазерный режим)</td><td>' + data.k_laz + ' мкА/мА</td></tr>';
  r += '<tr><td>b<sub>лаз</sub> — смещение (лазерный режим)</td><td>' + data.b_laz + ' мкА</td></tr>';
  r += '<tr><td>R² (светодиодный режим)</td><td>' + data.r2_sv + '</td></tr>';
  r += '<tr><td>R² (лазерный режим)</td><td>' + data.r2_laz + '</td></tr>';
  r += '<tr><td><b>Пороговый ток I<sub>пор</sub></b></td><td><b>' + data.i_por + ' мА</b></td></tr>';
  r += '<tr><td>Точки лазерного режима (красная линия)</td><td>' + data.mode_desc + '</td></tr>';
  r += '</table>';
  r += '<p class="note" style="color:#555;margin-top:8px;">Аппроксимация: I<sub>ФП</sub> = k·I<sub>pn</sub> + b для каждого режима. ' +
       'I<sub>пор</sub> определён как точка пересечения двух прямых (граница светодиодного и лазерного режимов). ' +
       'Красная линия лазерного режима строится по точкам из поля «Ввести значения» (МНК): подберите точки так, ' +
       'чтобы линия прошла по колену характеристики, и нажмите кнопку построения ещё раз — график и параметры пересчитаются.</p>';
  const rp = document.getElementById('res1');
  rp.innerHTML = r;
  rp.style.display = 'block';
  const wk1 = document.getElementById('work1-res');
  wk1.innerHTML = data.work_html;
  wk1.style.display = 'block';
}

// ---- Таблица 1.2 ----
function buildTable2(n) {
  n = Math.max(1, Math.min(40, parseInt(n) || 10));
  const tbody = document.getElementById('t2-body');
  tbody.innerHTML = '';
  for (let i = 1; i <= n; i++) {
    const tr = document.createElement('tr');
    tr.innerHTML =
      '<td class="num-col">' + i + '</td>' +
      '<td><input type="number" step="any" class="u" placeholder="—"></td>' +
      '<td class="calc ipn">—</td>' +
      '<td><input type="number" step="any" class="imax" placeholder="—"></td>' +
      '<td><input type="number" step="any" class="imin" placeholder="—"></td>' +
      '<td class="calc p">—</td>';
    tbody.appendChild(tr);
  }
}

function readRows2() {
  const rows = [];
  document.querySelectorAll('#t2-body tr').forEach(tr => {
    rows.push({
      u: tr.querySelector('.u').value,
      imax: tr.querySelector('.imax').value,
      imin: tr.querySelector('.imin').value
    });
  });
  return rows;
}

async function calcIpn2() {
  const r1 = getR1();
  if (r1 === null) { alert('Введите положительное значение R1 (Ом).'); return; }
  const rows = readRows2().filter(r => r.u !== '');
  if (rows.length === 0) { alert('Введите хотя бы одно значение U_ПЛ (мВ).'); return; }
  let res;
  try {
    res = await fetch('/api/kioe/lab1/calc_polar', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({r1: r1, rows: rows})
    });
  } catch(e) { alert('Ошибка сети: ' + e); return; }
  const data = await res.json();
  if (data.error) { alert(data.error); return; }
  let i = 0;
  document.querySelectorAll('#t2-body tr').forEach(tr => {
    const uInp = tr.querySelector('.u');
    const cell = tr.querySelector('.ipn');
    if (uInp.value !== '') { cell.textContent = data.rows[i++].ipn; }
    else { cell.textContent = '—'; }
  });
  // панель
  let html = '<p class="formula">Формула расчёта (закон Ома):</p>';
  html += '<p class="formula-eq">I<sub>pn</sub> = U<sub>ПЛ</sub> / R<sub>1</sub>, &nbsp;&nbsp; R<sub>1</sub> = ' + data.r1 + ' Ом</p>';
  html += '<p class="note">Значения занесены в столбец «I<sub>pn</sub>, мА». Все расчёты по каждой точке — в блоке «Расчетная часть» ниже.</p>';
  const panel = document.getElementById('calc2-ipn');
  panel.innerHTML = html;
  panel.style.display = 'block';
  const wp2 = document.getElementById('work2-ipn');
  wp2.innerHTML = data.work_html_ipn;
  wp2.style.display = 'block';
}

async function calcP2() {
  const r1 = getR1();
  if (r1 === null) { alert('Введите положительное значение R1 (Ом).'); return; }
  const rows = readRows2().filter(r => r.imax !== '' && r.imin !== '');
  if (rows.length === 0) { alert('Введите пары значений I_max и I_min (мкА).'); return; }
  let res;
  try {
    res = await fetch('/api/kioe/lab1/calc_polar', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({r1: r1, rows: rows})
    });
  } catch(e) { alert('Ошибка сети: ' + e); return; }
  const data = await res.json();
  if (data.error) { alert(data.error); return; }
  // заполняем столбец P по соответствию imax/imin
  let i = 0;
  document.querySelectorAll('#t2-body tr').forEach(tr => {
    const imaxInp = tr.querySelector('.imax');
    const iminInp = tr.querySelector('.imin');
    const cell = tr.querySelector('.p');
    if (imaxInp.value !== '' && iminInp.value !== '') {
      cell.textContent = (data.rows[i].p === null) ? '—' : data.rows[i].p;
      i++;
    } else { cell.textContent = '—'; }
  });
  // панель
  let html = '<p class="formula">Формула расчёта степени поляризации:</p>';
  html += '<p class="formula-eq">P = (I<sub>max</sub> − I<sub>min</sub>) / (I<sub>max</sub> + I<sub>min</sub>)</p>';
  html += '<p class="note">I<sub>max</sub>, I<sub>min</sub> — максимальная и минимальная интенсивности излучения, прошедшего поляризационный светофильтр (прибор М2, мкА).</p>';
  html += '<p class="note">Значения занесены в столбец «P». Все расчёты по каждой точке — в блоке «Расчетная часть» ниже.</p>';
  const panel = document.getElementById('calc2-p');
  panel.innerHTML = html;
  panel.style.display = 'block';
  const wp3 = document.getElementById('work2-p');
  wp3.innerHTML = data.work_html_p;
  wp3.style.display = 'block';
}

async function plotPolar() {
  const r1 = getR1();
  if (r1 === null) { alert('Введите положительное значение R1 (Ом).'); return; }
  const rows = [];
  document.querySelectorAll('#t2-body tr').forEach(tr => {
    const u = tr.querySelector('.u').value;
    const imax = tr.querySelector('.imax').value;
    const imin = tr.querySelector('.imin').value;
    if (u !== '' && imax !== '' && imin !== '') {
      rows.push({u: parseFloat(u), imax: parseFloat(imax), imin: parseFloat(imin)});
    }
  });
  if (rows.length < 2) { alert('Нужно минимум 2 полностью заполненные строки (U_ПЛ, I_max, I_min).'); return; }
  let res;
  try {
    res = await fetch('/api/kioe/lab1/plot_polar', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({r1: r1, rows: rows, i_por: lastIPor})
    });
  } catch(e) { alert('Ошибка сети: ' + e); return; }
  const data = await res.json();
  if (data.error) { alert(data.error); return; }
  const img2 = 'data:image/png;base64,' + data.image;
  document.getElementById('plot2-img').src = img2;
  document.getElementById('plot2-dl-png').href = img2;
  if (data.xlsx) document.getElementById('plot2-dl-xlsx').href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + data.xlsx;
  document.getElementById('plot2').style.display = 'block';
}

// инициализация
buildTable1(18);
buildTable2(10);
</script>
</body>
</html>
"""


LAB2_HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Лабораторная работа № 2 — Электрооптический эффект</title>
<style>
  :root { --accent:#1f4e79; --bg:#f4f6f9; --panel:#fff; --border:#d8dee9; --ok:#2e7d32; --warn:#c62828; }
  * { box-sizing:border-box; }
  body { margin:0; font-family:'Segoe UI', system-ui, Arial, sans-serif; background:var(--bg); color:#222; line-height:1.5; }
  header { background:var(--accent); color:#fff; padding:22px 20px; }
  header .wrap { max-width:1100px; margin:0 auto; display:flex; justify-content:space-between; align-items:center; gap:16px; flex-wrap:wrap; }
  header h1 { margin:0; font-size:21px; }
  header a { color:#fff; text-decoration:none; font-size:13.5px; border:1px solid rgba(255,255,255,.5); padding:6px 14px; border-radius:6px; }
  header a:hover { background:rgba(255,255,255,.15); }
  main { max-width:1100px; margin:0 auto; padding:24px 16px 60px; }
  .lead { font-size:14.5px; color:#444; background:#fff; border:1px solid var(--border); border-left:4px solid var(--accent); padding:14px 18px; border-radius:6px; margin-bottom:22px; }
  .lead b { color:var(--accent); }
  h2.sec { font-size:19px; color:var(--accent); border-bottom:2px solid var(--accent); padding-bottom:6px; margin:30px 0 14px; }
  h3.sub { font-size:15.5px; margin:18px 0 10px; color:#333; }
  .params { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:16px 18px; margin-bottom:22px; }
  .params-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(220px,1fr)); gap:14px 18px; align-items:end; }
  .params-grid label { font-weight:600; font-size:13.5px; display:flex; flex-direction:column; gap:5px; }
  .params-grid input { font-size:15px; padding:7px 10px; border:1px solid #b9c2d0; border-radius:5px; width:100%; }
  .params .hint { color:#777; font-size:12.5px; margin-top:10px; }
  .table-wrap { overflow-x:auto; }
  table.data { border-collapse:collapse; width:100%; background:var(--panel); font-size:13.5px; min-width:560px; }
  table.data caption { caption-side:top; text-align:left; font-weight:600; font-size:14px; padding:8px 10px; color:#333; background:#eef2f7; border:1px solid var(--border); border-bottom:none; border-radius:6px 6px 0 0; }
  table.data th, table.data td { border:1px solid var(--border); padding:6px 8px; text-align:center; }
  table.data th { background:#f0f4f9; font-weight:600; }
  table.data td.num-col { color:#888; width:42px; background:#fafbfc; }
  table.data td.calc { background:#fffdf3; font-weight:600; color:#7a5d00; }
  table.data input { width:100%; border:1px solid transparent; background:transparent; text-align:center; font-size:13.5px; padding:5px 4px; border-radius:4px; font-family:inherit; }
  table.data input:focus { border-color:var(--accent); background:#fff; outline:none; }
  .toolbar { display:flex; gap:10px; flex-wrap:wrap; margin:14px 0; align-items:center; }
  button { font-family:inherit; font-size:14px; padding:9px 16px; border:none; border-radius:6px; cursor:pointer; background:var(--accent); color:#fff; font-weight:600; transition:background .12s, transform .05s; }
  button:hover { background:#173d61; }
  button:active { transform:translateY(1px); }
  button.ghost { background:#fff; color:var(--accent); border:1px solid var(--accent); }
  button.ghost:hover { background:#eef3f9; }
  .calc-panel { background:#fffbea; border:1px solid #e6d98a; border-left:4px solid #d4a72c; border-radius:6px; padding:14px 18px; margin:14px 0; display:none; }
  .calc-panel .formula { font-weight:600; margin:0 0 6px; color:#6b5400; }
  .calc-panel .formula-eq { font-family:'Cambria Math', 'Times New Roman', serif; font-size:16px; margin:6px 0; }
  .work-title { font-weight:600; margin:12px 0 6px; color:#6b5400; }
  .work { border-collapse:collapse; width:100%; font-size:13.5px; background:#fff; }
  .work th, .work td { border:1px solid #e2d9b0; padding:5px 8px; text-align:left; vertical-align:middle; }
  .work th { background:#f3ecd2; font-weight:600; }
  .work td.wn { width:30px; text-align:center; color:#555; }
  .work td.weq { font-family:'Cambria Math', 'Times New Roman', serif; font-size:14.5px; }
  .work td.wres { font-weight:600; white-space:nowrap; font-family:'Cambria Math', 'Times New Roman', serif; }
  .work-scroll { max-height:430px; overflow:auto; border:1px solid #e6d98a; }
  .calc-panel .note { color:#777; font-size:13px; margin:8px 0 0; }
  .result-panel { background:#eef7ee; border:1px solid #bfe3bf; border-left:4px solid var(--ok); border-radius:6px; padding:14px 18px; margin:14px 0; display:none; }
  .result-panel h4 { margin:0 0 10px; color:var(--ok); font-size:15px; }
  table.res { border-collapse:collapse; font-size:13.5px; margin:6px 0; }
  table.res th, table.res td { border:1px solid #cfe6cf; padding:6px 12px; text-align:left; }
  table.res th { background:#e3f3e3; }
  .plot-wrap { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:14px; margin:14px 0; display:none; text-align:center; }
  .plot-wrap .ptitle { font-weight:600; font-size:14px; margin-bottom:6px; color:#333; }
  .plot-wrap img { max-width:100%; height:auto; border:1px solid #e0e0e0; border-radius:4px; }
  .plot-wrap .dl { display:inline-block; margin-top:8px; font-size:13px; color:var(--accent); }
  .plot-wrap .dl-stack { margin-top:4px; }
  .err { color:var(--warn); font-weight:600; }
  .muted { color:#777; font-size:13px; }
  sub { font-size:75%; }
</style>
</head>
<body>
<header>
  <div class="wrap">
    <h1>Лабораторная работа № 2. Определение электрооптических параметров анизотропных кристаллов</h1>
    <a href="/" target="_blank">&larr; к списку работ</a>
  </div>
</header>
<main>
  <div class="lead">
    Объект исследования — кристалл <b>ниобата лития (LiNbO<sub>3</sub>)</b>. Измеряется
    зависимость <b>интенсивности I</b> прошедшего света от <b>напряжения U</b>,
    приложенного к кристаллу (поперечный электрооптический эффект Поккельса).
    Зависимость описывается формулой I = I<sub>0</sub>·sin²(π·l·r·n³·U / (2·λ·d)).
    После линеаризации arcsin√(I/I<sub>0</sub>) = m·U определяются полуволновое
    напряжение U<sub>λ/2</sub> = π/(2m) и электрооптический коэффициент r.
  </div>

  <h2 class="sec">Параметры установки и кристалла</h2>
  <div class="params">
    <div class="params-grid">
      <label>Длина волны λ, нм
        <input type="number" id="p-lambda" step="any" min="1" value="632.8">
      </label>
      <label>Расстояние d между электродами, мм
        <input type="number" id="p-d" step="any" min="0.0001" value="2.55">
      </label>
      <label>Длина кристалла l, мм
        <input type="number" id="p-l" step="any" min="0.0001" value="27.5">
      </label>
      <label>Показатель преломления n<sub>0</sub>
        <input type="number" id="p-n0" step="any" min="0.0001" value="2.286">
      </label>
      <label>Показатель преломления n<sub>e</sub>
        <input type="number" id="p-ne" step="any" min="0.0001" value="2.196">
      </label>
      <label>I<sub>0</sub> (фототок при макс. пропускании), мкА
        <input type="number" id="p-i0" step="any" min="0" placeholder="напр. 100">
      </label>
      <label>Табличное r<sub>33</sub>, 10⁻¹² м/В (справочно)
        <input type="number" id="p-rtable" step="any" min="0" value="30.8">
      </label>
    </div>
    <p class="hint">Значения по умолчанию — из методички (LiNbO<sub>3</sub>, размеры 27,5×11,4×2,55 мм; n<sub>0</sub>=2,286, n<sub>e</sub>=2,196). λ=632,8 нм соответствует He-Ne лазеру. I<sub>0</sub> введите по результату настройки анализатора (максимальное показание фотодиода при U=0).</p>
  </div>

  <h2 class="sec">Таблица экспериментальных данных</h2>
  <div class="table-wrap">
    <table class="data" id="t1">
      <caption>Таблица 2.1 — Данные зависимости фототока I от напряжения U на кристалле</caption>
      <thead>
        <tr>
          <th>№</th>
          <th>U, В</th>
          <th>I, мкА</th>
          <th>I/I<sub>0</sub> (расч.)</th>
          <th>arcsin√(I/I<sub>0</sub>), рад (расч.)</th>
        </tr>
      </thead>
      <tbody id="t1-body"></tbody>
    </table>
  </div>

  <div class="toolbar">
    <button onclick="buildTable(prompt('Количество точек в таблице 2.1?', 15) || 15)">Изменить число строк</button>
    <button class="ghost" onclick="fillDemo()">Заполнить демонстрационными данными</button>
  </div>

  <div class="toolbar">
    <button onclick="processAll()">Рассчитать и построить графики</button>
  </div>

  <div class="calc-panel" id="calc-panel"></div>
  <div class="calc-panel" id="work-panel"></div>
  <div class="result-panel" id="res-panel"></div>

  <div class="plot-wrap" id="plot1">
    <div class="ptitle">График 1. Зависимость I(U)</div>
    <img id="plot1-img" alt="График I(U)">
    <div class="dl-stack"><a class="dl" id="plot1-dl-png" download="I(U).png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot1-dl-xlsx" download="I(U).xlsx">Скачать график EXCEL</a></div>
  </div>

  <div class="plot-wrap" id="plot2">
    <div class="ptitle">График 2. Линеаризация arcsin√(I/I<sub>0</sub>) = f(U)</div>
    <img id="plot2-img" alt="График arcsin(I/I0)">
    <div class="dl-stack"><a class="dl" id="plot2-dl-png" download="arcsin(I_I0)(U).png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot2-dl-xlsx" download="arcsin(I_I0)(U).xlsx">Скачать график EXCEL</a></div>
  </div>

  <p class="muted">Совет: сначала задайте I<sub>0</sub> в параметрах (максимальное показание фотодиода при настройке анализатора). Значения I, превышающие I<sub>0</sub>, будут исключены из линеаризации (помечены «—» в столбце arcsin).</p>
</main>

<script>
const P = { lambda:'p-lambda', d:'p-d', l:'p-l', n0:'p-n0', ne:'p-ne', i0:'p-i0', rtable:'p-rtable' };

function num(id) {
  const v = parseFloat(document.getElementById(id).value);
  return (isNaN(v)) ? null : v;
}

function getParams() {
  return {
    lambda_nm: num(P.lambda),
    d_mm: num(P.d),
    l_mm: num(P.l),
    n0: num(P.n0),
    ne: num(P.ne),
    i0: num(P.i0),
    rtable: num(P.rtable)
  };
}

function buildTable(n) {
  n = Math.max(1, Math.min(60, parseInt(n) || 15));
  const tbody = document.getElementById('t1-body');
  tbody.innerHTML = '';
  for (let i = 1; i <= n; i++) {
    const tr = document.createElement('tr');
    tr.innerHTML =
      '<td class="num-col">' + i + '</td>' +
      '<td><input type="number" step="any" class="u" placeholder="—"></td>' +
      '<td><input type="number" step="any" class="i" placeholder="—"></td>' +
      '<td class="calc ratio">—</td>' +
      '<td class="calc arcsin">—</td>';
    tbody.appendChild(tr);
  }
}

function readRows() {
  const rows = [];
  document.querySelectorAll('#t1-body tr').forEach(tr => {
    rows.push({ u: tr.querySelector('.u').value, i: tr.querySelector('.i').value });
  });
  return rows;
}

function fillDemo() {
  // Демонстрационные данные (по форме sin², характерной для LiNbO3 ~0-300 В)
  const demo = [
    [0, 0.3], [20, 1.2], [40, 3.8], [60, 8.5], [80, 16.2], [100, 26.5],
    [120, 39.0], [140, 52.5], [160, 66.0], [180, 78.5], [200, 88.5],
    [220, 95.0], [240, 97.5], [260, 95.5], [280, 88.0], [300, 76.0]
  ];
  buildTable(demo.length);
  document.getElementById('p-i0').value = 100;
  const trs = document.querySelectorAll('#t1-body tr');
  demo.forEach((d, idx) => {
    trs[idx].querySelector('.u').value = d[0];
    trs[idx].querySelector('.i').value = d[1];
  });
}

async function processAll() {
  const p = getParams();
  if (p.lambda_nm === null || p.d_mm === null || p.l_mm === null || p.n0 === null || p.ne === null || p.i0 === null) {
    alert('Заполните все параметры установки (λ, d, l, n0, ne, I0).'); return;
  }
  if (p.i0 <= 0) { alert('I0 должно быть положительным.'); return; }
  const rows = readRows().filter(r => r.u !== '' && r.i !== '');
  if (rows.length < 3) { alert('Введите минимум 3 точки (U, I).'); return; }
  let res;
  try {
    res = await fetch('/api/kioe/lab2/process', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({
        lambda_nm: p.lambda_nm, d_mm: p.d_mm, l_mm: p.l_mm,
        n0: p.n0, ne: p.ne, i0: p.i0, rtable: p.rtable,
        rows: rows
      })
    });
  } catch(e) { alert('Ошибка сети: ' + e); return; }
  const data = await res.json();
  if (data.error) { alert(data.error); return; }
  // заполнить таблицу: ratio и arcsin
  let idx = 0;
  document.querySelectorAll('#t1-body tr').forEach(tr => {
    const uInp = tr.querySelector('.u');
    const iInp = tr.querySelector('.i');
    if (uInp.value !== '' && iInp.value !== '') {
      const r = data.table[idx++];
      tr.querySelector('.ratio').textContent = r.ratio;
      tr.querySelector('.arcsin').textContent = r.arcsin;
    } else {
      tr.querySelector('.ratio').textContent = '—';
      tr.querySelector('.arcsin').textContent = '—';
    }
  });
  // панель расчёта
  const cp = document.getElementById('calc-panel');
  cp.innerHTML = data.calc_html;
  cp.style.display = 'block';
  const wp = document.getElementById('work-panel');
  wp.innerHTML = data.work_html;
  wp.style.display = 'block';
  // панель результатов
  const rp = document.getElementById('res-panel');
  rp.innerHTML = data.res_html;
  rp.style.display = 'block';
  // график I(U)
  const img1 = 'data:image/png;base64,' + data.iu_image;
  document.getElementById('plot1-img').src = img1;
  document.getElementById('plot1-dl-png').href = img1;
  document.getElementById('plot1-dl-xlsx').href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + data.iu_xlsx;
  document.getElementById('plot1').style.display = 'block';
  // график arcsin
  const img2 = 'data:image/png;base64,' + data.arcsin_image;
  document.getElementById('plot2-img').src = img2;
  document.getElementById('plot2-dl-png').href = img2;
  document.getElementById('plot2-dl-xlsx').href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + data.arcsin_xlsx;
  document.getElementById('plot2').style.display = 'block';
  document.getElementById('plot1').scrollIntoView({behavior:'smooth', block:'start'});
}

buildTable(15);
</script>
</body>
</html>
"""


LAB3_HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Лабораторная работа № 3 — Полупроводниковые детекторы</title>
<style>
  :root { --accent:#1f4e79; --bg:#f4f6f9; --panel:#fff; --border:#d8dee9; --ok:#2e7d32; --warn:#c62828; }
  * { box-sizing:border-box; }
  body { margin:0; font-family:'Segoe UI', system-ui, Arial, sans-serif; background:var(--bg); color:#222; line-height:1.5; }
  header { background:var(--accent); color:#fff; padding:22px 20px; }
  header .wrap { max-width:1100px; margin:0 auto; display:flex; justify-content:space-between; align-items:center; gap:16px; flex-wrap:wrap; }
  header h1 { margin:0; font-size:20px; }
  header a { color:#fff; text-decoration:none; font-size:13.5px; border:1px solid rgba(255,255,255,.5); padding:6px 14px; border-radius:6px; }
  header a:hover { background:rgba(255,255,255,.15); }
  main { max-width:1120px; margin:0 auto; padding:24px 16px 60px; }
  .lead { font-size:14.5px; color:#444; background:#fff; border:1px solid var(--border); border-left:4px solid var(--accent); padding:14px 18px; border-radius:6px; margin-bottom:22px; }
  .lead b { color:var(--accent); }
  h2.sec { font-size:19px; color:var(--accent); border-bottom:2px solid var(--accent); padding-bottom:6px; margin:30px 0 14px; }
  .params { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:16px 18px; margin-bottom:22px; }
  .params-grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(200px,1fr)); gap:14px 18px; align-items:end; }
  .params-grid label { font-weight:600; font-size:13.5px; display:flex; flex-direction:column; gap:5px; }
  .params-grid input { font-size:15px; padding:7px 10px; border:1px solid #b9c2d0; border-radius:5px; width:100%; }
  .params .hint { color:#777; font-size:12.5px; margin-top:10px; }
  .table-wrap { overflow-x:auto; }
  table.data { border-collapse:collapse; width:100%; background:var(--panel); font-size:13px; min-width:560px; }
  table.data caption { caption-side:top; text-align:left; font-weight:600; font-size:14px; padding:8px 10px; color:#333; background:#eef2f7; border:1px solid var(--border); border-bottom:none; border-radius:6px 6px 0 0; }
  table.data th, table.data td { border:1px solid var(--border); padding:6px 8px; text-align:center; }
  table.data th { background:#f0f4f9; font-weight:600; }
  table.data td.num-col { color:#888; width:42px; background:#fafbfc; }
  table.data td.calc { background:#fffdf3; font-weight:600; color:#7a5d00; }
  table.data input { width:100%; border:1px solid transparent; background:transparent; text-align:center; font-size:13px; padding:5px 4px; border-radius:4px; font-family:inherit; }
  table.data input:focus { border-color:var(--accent); background:#fff; outline:none; }
  .toolbar { display:flex; gap:10px; flex-wrap:wrap; margin:14px 0; align-items:center; }
  button { font-family:inherit; font-size:14px; padding:9px 16px; border:none; border-radius:6px; cursor:pointer; background:var(--accent); color:#fff; font-weight:600; transition:background .12s, transform .05s; }
  button:hover { background:#173d61; }
  button:active { transform:translateY(1px); }
  button.ghost { background:#fff; color:var(--accent); border:1px solid var(--accent); }
  button.ghost:hover { background:#eef3f9; }
  .calc-panel { background:#fffbea; border:1px solid #e6d98a; border-left:4px solid #d4a72c; border-radius:6px; padding:14px 18px; margin:14px 0; display:none; }
  .calc-panel .formula { font-weight:600; margin:0 0 6px; color:#6b5400; }
  .calc-panel .formula-eq { font-family:'Cambria Math', 'Times New Roman', serif; font-size:16px; margin:6px 0; }
  .work-title { font-weight:600; margin:12px 0 6px; color:#6b5400; }
  .work { border-collapse:collapse; width:100%; font-size:13.5px; background:#fff; }
  .work th, .work td { border:1px solid #e2d9b0; padding:5px 8px; text-align:left; vertical-align:middle; }
  .work th { background:#f3ecd2; font-weight:600; }
  .work td.wn { width:30px; text-align:center; color:#555; }
  .work td.weq { font-family:'Cambria Math', 'Times New Roman', serif; font-size:14.5px; }
  .work td.wres { font-weight:600; white-space:nowrap; font-family:'Cambria Math', 'Times New Roman', serif; }
  .work-scroll { max-height:430px; overflow:auto; border:1px solid #e6d98a; }
  .calc-panel .note { color:#777; font-size:13px; margin:8px 0 0; }
  .result-panel { background:#eef7ee; border:1px solid #bfe3bf; border-left:4px solid var(--ok); border-radius:6px; padding:14px 18px; margin:14px 0; display:none; }
  .result-panel h4 { margin:0 0 10px; color:var(--ok); font-size:15px; }
  table.res { border-collapse:collapse; font-size:13.5px; margin:6px 0; }
  table.res th, table.res td { border:1px solid #cfe6cf; padding:6px 10px; text-align:left; }
  table.res th { background:#e3f3e3; }
  .plot-wrap { background:var(--panel); border:1px solid var(--border); border-radius:8px; padding:14px; margin:14px 0; display:none; text-align:center; }
  .plot-wrap .ptitle { font-weight:600; font-size:14px; margin-bottom:6px; color:#333; }
  .plot-wrap img { max-width:100%; height:auto; border:1px solid #e0e0e0; border-radius:4px; }
  .plot-wrap .dl { display:inline-block; margin-top:8px; font-size:13px; color:var(--accent); }
  .plot-wrap .dl-stack { margin-top:4px; }
  .err { color:var(--warn); font-weight:600; }
  .muted { color:#777; font-size:13px; }
  sub { font-size:75%; }
</style>
</head>
<body>
<header>
  <div class="wrap">
    <h1>Лабораторная работа № 3. Полупроводниковые детекторы оптического излучения</h1>
    <a href="/" target="_blank">&larr; к списку работ</a>
  </div>
</header>
<main>
  <div class="lead">
    Объект исследования — полупроводниковый фотодиод (ФД-24К). Измеряются
    <b>U<sub>ПЛ</sub></b> (падение напряжения на R<sub>ПЛ</sub> лазера),
    <b>U<sub>ФД</sub></b> (без светофильтра) и <b>U*<sub>ФД</sub></b> (с нейтральным
    светофильтром) в двух режимах: <b>фотогальваническом</b> (без смещения,
    нагрузки R₁ и R₂) и <b>фотодиодном</b> (обратное смещение U<sub>см</sub>).
    Рассчитываются I<sub>pn</sub> = U<sub>ПЛ</sub>/R<sub>ПЛ</sub> и
    τ = U*<sub>ФД</sub>/U<sub>ФД</sub>. По постоянству τ определяется диапазон
    линейности фотодетектора.
  </div>

  <h2 class="sec">Параметры установки</h2>
  <div class="params">
    <div class="params-grid">
      <label>Сопротивление R<sub>ПЛ</sub>, Ом
        <input type="number" id="p-rpl" step="any" min="0.0001" value="100">
      </label>
      <label>Нагрузка R₁, Ом (фотогальв.)
        <input type="number" id="p-r1" step="any" min="0.0001" value="180">
      </label>
      <label>Нагрузка R₂, Ом (фотогальв.)
        <input type="number" id="p-r2" step="any" min="0.0001" value="1500">
      </label>
      <label>Смещение U<sub>см</sub>, В (фотодиодный)
        <input type="number" id="p-usm" step="any" min="0" value="1.5">
      </label>
      <label>Допуск на τ, %
        <input type="number" id="p-tol" step="any" min="0.1" value="5">
      </label>
    </div>
    <p class="hint">I<sub>pn</sub> = U<sub>ПЛ</sub>/R<sub>ПЛ</sub> (мВ/Ом = мА). Допуск на τ задаёт критерий линейности: точка считается линейной, если τ отклоняется от опорного значения не более чем на указанный процент.</p>
  </div>

  <h2 class="sec">Таблица 3.1 — Фотогальванический режим (без смещения)</h2>
  <div class="table-wrap">
    <table class="data" id="t1">
      <caption>Зависимость U_ФД и τ от I_pn при нагрузках R₁ и R₂</caption>
      <thead>
        <tr>
          <th rowspan="2">№</th>
          <th rowspan="2">U<sub>ПЛ</sub>, мВ</th>
          <th rowspan="2">I<sub>pn</sub>, мА</th>
          <th colspan="3">Нагрузка R₁</th>
          <th colspan="3">Нагрузка R₂</th>
        </tr>
        <tr>
          <th>U<sub>ФД</sub>, мВ</th><th>U*<sub>ФД</sub>, мВ</th><th>τ = U*<sub>ФД</sub>/U<sub>ФД</sub></th>
          <th>U<sub>ФД</sub>, мВ</th><th>U*<sub>ФД</sub>, мВ</th><th>τ = U*<sub>ФД</sub>/U<sub>ФД</sub></th>
        </tr>
      </thead>
      <tbody id="t1-body"></tbody>
    </table>
  </div>
  <div class="toolbar">
    <button class="ghost" onclick="buildTable1(prompt('Количество строк в таблице 3.1?', 12) || 12)">Изменить число строк</button>
  </div>

  <h2 class="sec">Таблица 3.2 — Фотодиодный режим (обратное смещение)</h2>
  <div class="table-wrap">
    <table class="data" id="t2">
      <caption>Зависимость U_ФД и τ от I_pn в фотодиодном режиме</caption>
      <thead>
        <tr>
          <th>№</th><th>U<sub>ПЛ</sub>, мВ</th><th>I<sub>pn</sub>, мА</th>
          <th>U<sub>ФД</sub>, мВ</th><th>U*<sub>ФД</sub>, мВ</th><th>τ = U*<sub>ФД</sub>/U<sub>ФД</sub></th>
        </tr>
      </thead>
      <tbody id="t2-body"></tbody>
    </table>
  </div>
  <div class="toolbar">
    <button class="ghost" onclick="buildTable2(prompt('Количество строк в таблице 3.2?', 12) || 12)">Изменить число строк</button>
    <button class="ghost" onclick="fillDemo()">Заполнить демонстрационными данными</button>
  </div>

  <div class="toolbar">
    <button onclick="processAll()">Рассчитать I_pn, τ и построить графики</button>
  </div>

  <div class="calc-panel" id="calc-panel"></div>
  <div class="calc-panel" id="work-panel"></div>
  <div class="result-panel" id="res-panel"></div>

  <div class="plot-wrap" id="plot1">
    <div class="ptitle">График 1. Фотогальванический режим, нагрузка R₁</div>
    <img id="plot1-img" alt="График 1">
    <div class="dl-stack"><a class="dl" id="plot1-dl-png" download="UFD_tau_R1.png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot1-dl-xlsx" download="UFD_tau_R1.xlsx">Скачать график EXCEL</a></div>
  </div>
  <div class="plot-wrap" id="plot2">
    <div class="ptitle">График 2. Фотогальванический режим, нагрузка R₂</div>
    <img id="plot2-img" alt="График 2">
    <div class="dl-stack"><a class="dl" id="plot2-dl-png" download="UFD_tau_R2.png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot2-dl-xlsx" download="UFD_tau_R2.xlsx">Скачать график EXCEL</a></div>
  </div>
  <div class="plot-wrap" id="plot3">
    <div class="ptitle">График 3. Фотодиодный режим (обратное смещение)</div>
    <img id="plot3-img" alt="График 3">
    <div class="dl-stack"><a class="dl" id="plot3-dl-png" download="UFD_tau_photodiode.png">Скачать график PNG</a></div>
    <div class="dl-stack"><a class="dl" id="plot3-dl-xlsx" download="UFD_tau_photodiode.xlsx">Скачать график EXCEL</a></div>
  </div>

  <p class="muted">Совет: сначала заполните таблицы (или нажмите «Заполнить демонстрационными данными»), затем «Рассчитать I_pn, τ и построить графики». Линейный диапазон определяется по постоянству τ: рост τ при увеличении I<sub>pn</sub> означает насыщение фотодиода.</p>
</main>

<script>
const P = { rpl:'p-rpl', r1:'p-r1', r2:'p-r2', usm:'p-usm', tol:'p-tol' };

function num(id) {
  const v = parseFloat(document.getElementById(id).value);
  return (isNaN(v)) ? null : v;
}

function getParams() {
  return { r_pl: num(P.rpl), r1: num(P.r1), r2: num(P.r2), u_sm: num(P.usm), tol: num(P.tol) };
}

function buildTable1(n) {
  n = Math.max(1, Math.min(40, parseInt(n) || 12));
  const tbody = document.getElementById('t1-body');
  tbody.innerHTML = '';
  for (let i = 1; i <= n; i++) {
    const tr = document.createElement('tr');
    tr.innerHTML =
      '<td class="num-col">'+i+'</td>' +
      '<td><input type="number" step="any" class="upl" placeholder="—"></td>' +
      '<td class="calc ipn">—</td>' +
      '<td><input type="number" step="any" class="ufd1" placeholder="—"></td>' +
      '<td><input type="number" step="any" class="usfd1" placeholder="—"></td>' +
      '<td class="calc tau1">—</td>' +
      '<td><input type="number" step="any" class="ufd2" placeholder="—"></td>' +
      '<td><input type="number" step="any" class="usfd2" placeholder="—"></td>' +
      '<td class="calc tau2">—</td>';
    tbody.appendChild(tr);
  }
}

function buildTable2(n) {
  n = Math.max(1, Math.min(40, parseInt(n) || 12));
  const tbody = document.getElementById('t2-body');
  tbody.innerHTML = '';
  for (let i = 1; i <= n; i++) {
    const tr = document.createElement('tr');
    tr.innerHTML =
      '<td class="num-col">'+i+'</td>' +
      '<td><input type="number" step="any" class="upl" placeholder="—"></td>' +
      '<td class="calc ipn">—</td>' +
      '<td><input type="number" step="any" class="ufd" placeholder="—"></td>' +
      '<td><input type="number" step="any" class="usfd" placeholder="—"></td>' +
      '<td class="calc tau">—</td>';
    tbody.appendChild(tr);
  }
}

function fillDemo() {
  const N = 10;
  const upl = [2000,1800,1600,1400,1200,1000,800,600,400,200];
  // Фотогальванический R1=180 (лин. до I_pn~12 мА, k=15 мВ/мА), насыщение выше
  const ufd1 = [218,212,205,195,180,150,120,90,60,30];
  const usfd1 = [150,135,120,105,90,75,60,45,30,15];
  // Фотогальванический R2=1500 (лин. до I_pn~6 мА, k=80 мВ/мА), насыщение выше
  const ufd2 = [596,595,593,590,585,570,540,480,320,160];
  const usfd2 = [590,585,570,540,480,400,320,240,160,80];
  // Фотодиодный режим (лин. во всём диапазоне, k=20 мВ/мА)
  const ufd3 = [400,360,320,280,240,200,160,120,80,40];
  const usfd3 = [200,180,160,140,120,100,80,60,40,20];
  buildTable1(N);
  buildTable2(N);
  document.getElementById('p-rpl').value = 100;
  document.getElementById('p-r1').value = 180;
  document.getElementById('p-r2').value = 1500;
  document.getElementById('p-usm').value = 1.5;
  document.getElementById('p-tol').value = 5;
  const trs1 = document.querySelectorAll('#t1-body tr');
  trs1.forEach((tr, i) => {
    tr.querySelector('.upl').value = upl[i];
    tr.querySelector('.ufd1').value = ufd1[i];
    tr.querySelector('.usfd1').value = usfd1[i];
    tr.querySelector('.ufd2').value = ufd2[i];
    tr.querySelector('.usfd2').value = usfd2[i];
  });
  const trs2 = document.querySelectorAll('#t2-body tr');
  trs2.forEach((tr, i) => {
    tr.querySelector('.upl').value = upl[i];
    tr.querySelector('.ufd').value = ufd3[i];
    tr.querySelector('.usfd').value = usfd3[i];
  });
}

async function processAll() {
  const p = getParams();
  if (p.r_pl === null || p.r_pl <= 0) { alert('Введите положительное R_ПЛ (Ом).'); return; }
  if (p.tol === null || p.tol <= 0) { alert('Введите положительный допуск на τ (%).'); return; }
  // Таблица 3.1
  const t1 = [];
  document.querySelectorAll('#t1-body tr').forEach(tr => {
    t1.push({
      u_pl: tr.querySelector('.upl').value,
      u_fd_r1: tr.querySelector('.ufd1').value,
      us_fd_r1: tr.querySelector('.usfd1').value,
      u_fd_r2: tr.querySelector('.ufd2').value,
      us_fd_r2: tr.querySelector('.usfd2').value
    });
  });
  // Таблица 3.2
  const t2 = [];
  document.querySelectorAll('#t2-body tr').forEach(tr => {
    t2.push({
      u_pl: tr.querySelector('.upl').value,
      u_fd: tr.querySelector('.ufd').value,
      us_fd: tr.querySelector('.usfd').value
    });
  });
  const hasT1 = t1.some(r => r.u_pl !== '' && (r.u_fd_r1 !== '' || r.u_fd_r2 !== ''));
  const hasT2 = t2.some(r => r.u_pl !== '' && r.u_fd !== '');
  if (!hasT1 && !hasT2) { alert('Заполните хотя бы одну таблицу (U_ПЛ, U_ФД, U*_ФД).'); return; }
  let res;
  try {
    res = await fetch('/api/kioe/lab3/process', {
      method:'POST', headers:{'Content-Type':'application/json'},
      body: JSON.stringify({
        r_pl: p.r_pl, r1: p.r1, r2: p.r2, u_sm: p.u_sm, tol_pct: p.tol,
        table1: t1, table2: t2
      })
    });
  } catch(e) { alert('Ошибка сети: ' + e); return; }
  const data = await res.json();
  if (data.error) { alert(data.error); return; }
  // заполнить расчётные столбцы таблицы 3.1
  let i = 0;
  document.querySelectorAll('#t1-body tr').forEach(tr => {
    if (tr.querySelector('.upl').value !== '') {
      const r = data.table1[i++];
      tr.querySelector('.ipn').textContent = r.ipn;
      tr.querySelector('.tau1').textContent = r.tau1;
      tr.querySelector('.tau2').textContent = r.tau2;
    } else {
      tr.querySelector('.ipn').textContent = '—';
      tr.querySelector('.tau1').textContent = '—';
      tr.querySelector('.tau2').textContent = '—';
    }
  });
  // таблица 3.2
  let j = 0;
  document.querySelectorAll('#t2-body tr').forEach(tr => {
    if (tr.querySelector('.upl').value !== '') {
      const r = data.table2[j++];
      tr.querySelector('.ipn').textContent = r.ipn;
      tr.querySelector('.tau').textContent = r.tau;
    } else {
      tr.querySelector('.ipn').textContent = '—';
      tr.querySelector('.tau').textContent = '—';
    }
  });
  // панель формул
  const cp = document.getElementById('calc-panel');
  cp.innerHTML = '<p class="formula">Формулы расчёта:</p>' +
    '<p class="formula-eq">I<sub>pn</sub> = U<sub>ПЛ</sub> / R<sub>ПЛ</sub> &nbsp;&nbsp;(U<sub>ПЛ</sub> в мВ, R<sub>ПЛ</sub> в Ом → I<sub>pn</sub> в мА)</p>' +
    '<p class="formula-eq">τ = U*<sub>ФД</sub> / U<sub>ФД</sub> &nbsp;&nbsp;— пропускание нейтрального светофильтра</p>' +
    '<p class="note">В линейном режиме τ = const. Рост τ при увеличении I<sub>pn</sub> означает насыщение фотодиода (выход за пределы линейного диапазона): U<sub>ФД</sub> (без фильтра, полная интенсивность) подавляется, а U*<sub>ФД</sub> (с фильтром, меньшая интенсивность) остаётся в линейной области.</p>';
  cp.style.display = 'block';
  const wp = document.getElementById('work-panel');
  wp.innerHTML = data.work_html;
  wp.style.display = 'block';
  // панель результатов
  const rp = document.getElementById('res-panel');
  rp.innerHTML = data.analysis_html;
  rp.style.display = 'block';
  // графики
  for (const g of ['g1','g2','g3']) {
    const idx = g.slice(1);
    const img = data[g+'_image'];
    const xlsx = data[g+'_xlsx'];
    const wrap = document.getElementById('plot'+idx);
    if (img) {
      const src = 'data:image/png;base64,' + img;
      document.getElementById('plot'+idx+'-img').src = src;
      document.getElementById('plot'+idx+'-dl-png').href = src;
    }
    if (xlsx) document.getElementById('plot'+idx+'-dl-xlsx').href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + xlsx;
    wrap.style.display = img ? 'block' : 'none';
  }
  document.getElementById('res-panel').scrollIntoView({behavior:'smooth', block:'start'});
}

buildTable1(12);
buildTable2(12);
</script>
</body>
</html>
"""


# ===========================================================================
#  Расчёты: лабораторная работа № 1
# ===========================================================================


def _linfit_r2(xs, ys):
    """МНК-прямая и R² по набору точек."""
    x = np.asarray(xs, dtype=float)
    y = np.asarray(ys, dtype=float)
    if len(x) < 2:
        return None, None, None
    k, b0 = np.polyfit(x, y, 1)
    yfit = k * x + b0
    sst = float(np.sum((y - np.mean(y)) ** 2))
    r2 = 1.0 - float(np.sum((yfit - y) ** 2)) / sst if sst > 0 else 1.0
    return float(k), float(b0), float(r2)


def _ranges_str(nums):
    """Компактная строка номеров: [10,11,12,15] -> "10-12, 15"."""
    nums = sorted(nums)
    parts = []
    i = 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        parts.append(str(nums[i]) if j == i else f"{nums[i]}-{nums[j]}")
        i = j + 1
    return ", ".join(parts)


def parse_laser_selection(s_raw, rows):
    """Разбирает поле «Ввести значения» для прямой лазерного режима.

    Возвращает ("auto", None) | ("rows", [позиции точек]) |
               ("pairs", [(x, y), ...]) | ("error", сообщение).
    rows — номера строк таблицы (колонка №) для каждой заполненной пары.
    """
    s = (s_raw or "").strip()
    if not s:
        return "auto", None
    if ";" in s:
        pts = []
        for tok in s.split():
            parts = tok.split(";")
            if len(parts) != 2:
                return "error", (f"Некорректная пара «{tok}»: в режиме координат каждая "
                                 f"пара имеет вид «x;y» (напр. 2.45;70), пары разделяются пробелом.")
            try:
                pts.append((float(parts[0].replace(",", ".")),
                            float(parts[1].replace(",", "."))))
            except ValueError:
                return "error", f"Некорректная пара «{tok}» (нужны числа)."
        if len(pts) < 2:
            return "error", "Для прямой лазерного режима нужно минимум 2 точки."
        return "pairs", pts
    row_set = set()
    try:
        toks = s.replace(",", " ").split()
        for tok in toks:
            if "-" in tok:
                a_s, b_s = tok.split("-")
                a, b = int(a_s), int(b_s)
                if a > b or a < 1:
                    return "error", f"Некорректный диапазон «{tok}»."
                row_set.update(range(a, b + 1))
            else:
                m = int(tok)
                if m < 1:
                    return "error", "Номера строк должны быть положительными."
                row_set.add(m)
    except ValueError:
        return "error", ("Не удалось разобрать список номеров строк. Примеры: «10», "
                         "«10-18», «10,12,14-18»; координаты задаются парами «x;y».")
    if len(row_set) == 1 and len(toks) == 1:
        # одно число — «первая точка лазерного режима»: берём до конца таблицы
        m = row_set.pop()
        row_set = set(range(m, max(rows) + 1))
    idx = [i for i, r in enumerate(rows) if r in row_set]
    if len(idx) < 2:
        return "error", ("Выбрано меньше 2 заполненных точек для линии лазерного "
                         "режима. Проверьте номера строк (колонка №).")
    return "rows", idx


def lab1_intensity_fit(i_pn, i_fp, rows, laser_str):
    """Двухпрямая аппроксимация I_ФП(I_pn) с выбором точек лазерного режима.

    Возвращает dict: k_sv, b_sv, k_laz, b_laz, i_por, n, r2_sv, r2_laz,
    mode_desc, led_mask / laz_mask (маски точек данных), laz_pts_extra
    (точки, заданные координатами), split (индекс для авто-режима).
    """
    x = np.asarray(i_pn, dtype=float)
    y = np.asarray(i_fp, dtype=float)
    n = len(x)
    mode, sel = parse_laser_selection(laser_str, rows)
    if mode == "error":
        return {"error": sel}
    led_mask = np.zeros(n, dtype=bool)
    laz_mask = np.zeros(n, dtype=bool)
    laz_pts_extra = []
    split = None
    if mode == "auto":
        base = piecewise_linear_fit(x, y, None)
        if "error" in base:
            return {"error": base["error"]}
        split = base["split"]
        led_mask[:split] = True
        laz_mask[split:] = True
        mode_desc = f"авто-поиск: № 1–{split} светодиодный, № {split + 1}–{n} лазерный"
    elif mode == "rows":
        laz_mask[sel] = True
        first_row = min(rows[i] for i in sel)
        for i, r in enumerate(rows):
            if r < first_row:
                led_mask[i] = True
        mode_desc = (f"№ {_ranges_str([rows[i] for i in sel])} "
                     f"({len(sel)} точек); светодиодный — № < {first_row}")
    else:  # pairs
        first_x = min(p[0] for p in sel)
        for i in range(n):
            if x[i] < first_x:
                led_mask[i] = True
        laz_pts_extra = sel
        mode_desc = f"точки заданы координатами ({len(sel)} шт.); светодиодный — I_pn < {first_x:g} мА"
    if int(led_mask.sum()) < 2:
        return {"error": ("Для прямой светодиодного режима нужно минимум 2 точки "
                          "до выбранных точек лазерного режима.")}
    k1, c1, r2_1 = _linfit_r2(x[led_mask], y[led_mask])
    if mode == "pairs":
        k2, c2, r2_2 = _linfit_r2([p[0] for p in sel], [p[1] for p in sel])
    else:
        k2, c2, r2_2 = _linfit_r2(x[laz_mask], y[laz_mask])
    i_por = None if abs(k2 - k1) < 1e-12 else float((c1 - c2) / (k2 - k1))
    return {"k_sv": k1, "b_sv": c1, "k_laz": k2, "b_laz": c2, "i_por": i_por,
            "n": int(n), "r2_sv": r2_1, "r2_laz": r2_2, "mode_desc": mode_desc,
            "led_mask": led_mask, "laz_mask": laz_mask,
            "laz_pts_extra": laz_pts_extra, "split": split}


def make_intensity_plot(x, y, fit):
    """График I_ФП(I_pn): точки с подписями значений, линии режимов, I_пор.

    Точки, использованные в подгонках, выделяются цветом: зелёные —
    светодиодный режим, жёлтые с красной каймой — лазерный (как в Excel).
    """
    fig, ax = plt.subplots(figsize=(9.5, 6.0), constrained_layout=True)
    x = np.asarray(x, dtype=float)
    y = np.asarray(y, dtype=float)
    ax.scatter(x, y, s=44, color="#1f4e79", zorder=5, label="Эксперимент (не в подгонке)")
    led_mask = fit.get("led_mask")
    laz_mask = fit.get("laz_mask")
    if led_mask is not None and led_mask.any():
        ax.scatter(x[led_mask], y[led_mask], s=52, facecolor="#66bb6a",
                   edgecolor="#1b5e20", linewidth=1.2, zorder=6,
                   label="Точки светодиодного режима")
    if laz_mask is not None and laz_mask.any():
        ax.scatter(x[laz_mask], y[laz_mask], s=58, facecolor="#fdd835",
                   edgecolor="#c62828", linewidth=1.4, zorder=7,
                   label="Точки лазерного режима (красная линия)")
    if laz_mask is not None and fit.get("laz_pts_extra"):
        px = [p[0] for p in fit["laz_pts_extra"]]
        py = [p[1] for p in fit["laz_pts_extra"]]
        ax.scatter(px, py, s=58, facecolor="#fdd835", edgecolor="#c62828",
                   linewidth=1.4, zorder=7,
                   label="Заданные точки лазерного режима")
    # подписи значений у каждой экспериментальной точки
    for i, (xi, yi) in enumerate(zip(x, y)):
        dy = 7 if i % 2 == 0 else -13
        ax.annotate(f"{xi:.2f}; {yi:.0f}", xy=(xi, yi),
                    xytext=(6, dy), textcoords="offset points",
                    fontsize=7.2, color="#444", zorder=9)
    if fit and "error" not in fit:
        x_int = fit["i_por"]
        led_xmax = float(x[led_mask].max()) if (led_mask is not None and led_mask.any()) else float(x.min())
        x_left_end = x_int if (x_int is not None and x_int <= led_xmax) else led_xmax
        xl = np.linspace(float(x.min()), x_left_end, 50)
        ax.plot(xl, fit["k_sv"] * xl + fit["b_sv"], "--", color="#2e7d32", linewidth=1.8,
                label=(f"Светодиодный режим: y = {fit['k_sv']:.2f}·x {fit['b_sv']:+.2f}, "
                       f"R² = {fit['r2_sv']:.4f}"))
        span = float(x.max() - x.min()) or 1.0
        # начало красной линии — там, где она пересекает ось I_ФП = 0 (колено)
        x0 = (-fit["b_laz"] / fit["k_laz"]) if fit["k_laz"] != 0 else float(x.min())
        x_start = min(max(x0, float(x.min()) - 0.10 * span), float(x.max()))
        x_end = float(x.max()) + 0.03 * span
        xr = np.linspace(x_start, x_end, 50)
        ax.plot(xr, fit["k_laz"] * xr + fit["b_laz"], "-", color="#c62828", linewidth=2.2,
                label=(f"Лазерный режим: y = {fit['k_laz']:.1f}·x {fit['b_laz']:+.1f}, "
                       f"R² = {fit['r2_laz']:.4f}"))
        if x_int is not None:
            ax.axvline(x_int, color="#888", linestyle=":", linewidth=1.3)
            y_int = fit["k_sv"] * x_int + fit["b_sv"]
            y_top = float(ax.get_ylim()[1])
            ax.annotate(f"I$_{{пор}}$ ≈ {x_int:.3f} мА",
                        xy=(x_int, y_int), xytext=(x_int, y_top * 0.92),
                        fontsize=10, color="#333", ha="left",
                        bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#bbb", alpha=0.9))
    ax.set_xlabel("I$_{pn}$, мА", fontsize=12)
    ax.set_ylabel("I$_{ФП}$, мкА", fontsize=12)
    ax.set_title("Зависимость интенсивности излучения от тока через p-n переход", fontsize=12)
    ax.grid(True, alpha=0.3)
    ax.legend(loc="best", fontsize=8.5)
    return fig_to_base64(fig)


def make_polar_plot(x, p, i_por=None):
    """График P(I_pn) с разметкой режимов: положение I_пор, области
    светодиодного (P мал) и лазерного (P возрастает) режимов.

    i_por берётся из последнего построения части 1 (аппроксимация I_ФП).
    """
    fig, ax = plt.subplots(figsize=(9.5, 6.0), constrained_layout=True)
    x = np.asarray(x, dtype=float)
    p = np.asarray(p, dtype=float)
    lo, hi = float(x.min()), float(x.max())
    span = (hi - lo) or 1.0
    marked = False
    if i_por is not None and (lo - 0.15 * span) <= i_por <= (hi + 0.15 * span):
        marked = True
        x_m = float(min(max(i_por, lo), hi))  # граница областей внутри графика
        ax.axvspan(lo, x_m, color="#2e7d32", alpha=0.10, zorder=0)
        ax.axvspan(x_m, hi, color="#c62828", alpha=0.10, zorder=0)
        ax.axvline(i_por, color="#333", linestyle=":", linewidth=1.5, zorder=4)
        y_top = 1.0
        ax.annotate(f"I$_{{пор}}$ ≈ {i_por:.3f} мА (из части 1)",
                    xy=(i_por, y_top), xycoords=("data", "axes fraction"),
                    xytext=(6, -14), textcoords="offset points",
                    fontsize=10, color="#333", ha="left", va="top",
                    bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#bbb", alpha=0.92))
        # подписи областей: слева сверху (кривая P там внизу), справа снизу
        xc_l = lo + 0.45 * (x_m - lo)
        xc_r = x_m + 0.55 * (hi - x_m)
        ax.text(xc_l, 0.78, "Светодиодный режим\n(P мал, ≈ const)",
                transform=ax.get_xaxis_transform(), fontsize=9.5,
                color="#1b5e20", ha="center", va="center",
                bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="#2e7d32", alpha=0.85))
        ax.text(xc_r, 0.07, "Лазерный режим\n(P возрастает)",
                transform=ax.get_xaxis_transform(), fontsize=9.5,
                color="#b71c1c", ha="center", va="bottom",
                bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="#c62828", alpha=0.85))
    ax.plot(x, p, "-o", color="#6a1b9a", markersize=7, linewidth=1.6, zorder=6,
            label="Степень поляризации P")
    ax.set_xlabel("I$_{pn}$, мА", fontsize=12)
    ax.set_ylabel("P = (I$_{max}$ − I$_{min}$) / (I$_{max}$ + I$_{min}$)", fontsize=12)
    ax.set_title("Зависимость степени поляризации излучения от тока через p-n переход", fontsize=12)
    ax.grid(True, alpha=0.3)
    ax.legend(loc="upper left", fontsize=10)
    return fig_to_base64(fig)


def _fx_or_val(cond, formula, value, nd=5):
    """Formula (если диапазоны непрерывны) либо числовое значение."""
    return formula if cond else (round(value, nd) if value is not None else "—")


def make_intensity_xlsx(x, y, fit):
    """Excel-файл I_ФП(I_pn): данные + ЖИВЫЕ ФОРМУЛЫ расчёта + диаграмма.

    Параметры подгонки записываются формулами НАКЛОН/ОТРЕЗОК/RSQ
    (в русском Excel =НАКЛОН(...), =ОТРЕЗОК(...)) по диапазонам ячеек,
    I_пор — формулой пересечения прямых, колонки линий — формулами
    k·x+b со ссылками на ячейки параметров.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Данные"
    wb.calculation.fullCalcOnLoad = True  # пересчёт формул при открытии
    n = len(x)
    led_mask = fit.get("led_mask")
    laz_mask = fit.get("laz_mask")
    led_pos = [i for i in range(n) if led_mask[i]]
    laz_pos = [i for i in range(n) if laz_mask[i]]
    # непрерывность блоков в отсортированных данных (для формул с диапазонами)
    led_cont = bool(led_pos) and led_pos == list(range(led_pos[0], led_pos[-1] + 1))
    laz_cont = bool(laz_pos) and laz_pos == list(range(laz_pos[0], laz_pos[-1] + 1))

    ws.append(["№", "I_pn, мА", "I_ФП, мкА (эксп.)",
               "I_ФП (св. режим)", "I_ФП (лаз. режим)"])
    for i, (xi, yi) in enumerate(zip(x, y), start=1):
        ws.append([i, round(float(xi), 4), round(float(yi), 4)])
    # строки Excel: данные 2..n+1; блок параметров с n+3
    prow = n + 3
    r_ks, r_bs = prow + 1, prow + 2
    r_kl, r_bl = prow + 3, prow + 4
    r_por, r_r2s, r_r2l = prow + 5, prow + 6, prow + 7
    r_desc = prow + 8
    # колонки линий: формулы k·x + b со ссылками на ячейки параметров
    for row in range(2, n + 2):
        ws.cell(row=row, column=4, value=f"=$B${r_ks}*B{row}+$B${r_bs}")
        ws.cell(row=row, column=5, value=f"=$B${r_kl}*B{row}+$B${r_bl}")

    a1, b1 = led_pos[0] + 2, led_pos[-1] + 2 if led_pos else 2
    ws.cell(row=prow, column=1, value="Параметр")
    ws.cell(row=prow, column=2, value="Значение")
    ws.cell(row=r_ks, column=1, value="k_св, мкА/мА (НАКЛОН)")
    ws.cell(row=r_ks, column=2,
            value=_fx_or_val(led_cont, f"=SLOPE(C{a1}:C{b1},B{a1}:B{b1})", fit["k_sv"]))
    ws.cell(row=r_bs, column=1, value="b_св, мкА (ОТРЕЗОК)")
    ws.cell(row=r_bs, column=2,
            value=_fx_or_val(led_cont, f"=INTERCEPT(C{a1}:C{b1},B{a1}:B{b1})", fit["b_sv"]))
    ws.cell(row=r_r2s, column=1, value="R² св. режим (RSQ)")
    ws.cell(row=r_r2s, column=2,
            value=_fx_or_val(led_cont, f"=RSQ(C{a1}:C{b1},B{a1}:B{b1})", fit["r2_sv"]))
    if fit.get("laz_pts_extra"):
        # точки лазерного режима заданы координатами — отдельная мини-таблица
        mt = r_desc + 2
        ws.cell(row=mt, column=1, value="Точки прямой лазерного режима (заданы координатами)")
        ws.cell(row=mt + 1, column=1, value="I_pn, мА")
        ws.cell(row=mt + 1, column=2, value="I_ФП, мкА")
        pts = fit["laz_pts_extra"]
        for j, (px_, py_) in enumerate(pts):
            ws.cell(row=mt + 2 + j, column=1, value=round(px_, 4))
            ws.cell(row=mt + 2 + j, column=2, value=round(py_, 4))
        ma, mb = mt + 2, mt + 1 + len(pts)
        ws.cell(row=r_kl, column=1, value="k_лаз, мкА/мА (НАКЛОН)")
        ws.cell(row=r_kl, column=2, value=f"=SLOPE(B{ma}:B{mb},A{ma}:A{mb})")
        ws.cell(row=r_bl, column=1, value="b_лаз, мкА (ОТРЕЗОК)")
        ws.cell(row=r_bl, column=2, value=f"=INTERCEPT(B{ma}:B{mb},A{ma}:A{mb})")
        ws.cell(row=r_r2l, column=1, value="R² лаз. режим (RSQ)")
        ws.cell(row=r_r2l, column=2, value=f"=RSQ(B{ma}:B{mb},A{ma}:A{mb})")
    else:
        a2, b2 = (laz_pos[0] + 2, laz_pos[-1] + 2) if laz_pos else (2, 2)
        ws.cell(row=r_kl, column=1, value="k_лаз, мкА/мА (НАКЛОН)")
        ws.cell(row=r_kl, column=2,
                value=_fx_or_val(laz_cont, f"=SLOPE(C{a2}:C{b2},B{a2}:B{b2})", fit["k_laz"]))
        ws.cell(row=r_bl, column=1, value="b_лаз, мкА (ОТРЕЗОК)")
        ws.cell(row=r_bl, column=2,
                value=_fx_or_val(laz_cont, f"=INTERCEPT(C{a2}:C{b2},B{a2}:B{b2})", fit["b_laz"]))
        ws.cell(row=r_r2l, column=1, value="R² лаз. режим (RSQ)")
        ws.cell(row=r_r2l, column=2,
                value=_fx_or_val(laz_cont, f"=RSQ(C{a2}:C{b2},B{a2}:B{b2})", fit["r2_laz"]))
    ws.cell(row=r_por, column=1, value="I_пор, мА (пересечение прямых)")
    ws.cell(row=r_por, column=2,
            value="— (линии параллельны)" if fit["i_por"] is None
            else f"=(B{r_bs}-B{r_bl})/(B{r_kl}-B{r_ks})")
    ws.cell(row=r_desc, column=1, value="Точки лазерного режима")
    ws.cell(row=r_desc, column=2, value=fit.get("mode_desc", ""))
    for col in range(1, 6):
        ws.column_dimensions[get_column_letter(col)].width = 22

    chart = ScatterChart()
    chart.title = "Зависимость I_ФП(I_pn)"
    chart.x_axis.title = "I_pn, мА"
    chart.y_axis.title = "I_ФП, мкА"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.height = 12
    chart.width = 22
    xref = Reference(ws, min_col=2, min_row=2, max_row=n + 1)

    s_exp = Series(Reference(ws, min_col=3, min_row=2, max_row=n + 1), xref, title="Эксперимент")
    s_exp.marker = Marker(symbol="circle", size=7)
    s_exp.graphicalProperties = GraphicalProperties()
    s_exp.graphicalProperties.line.noFill = True
    chart.series.append(s_exp)

    s_led = Series(Reference(ws, min_col=4, min_row=2, max_row=n + 1), xref, title="Светодиодный режим")
    s_led.marker = Marker(symbol="none")
    s_led.graphicalProperties = GraphicalProperties()
    s_led.graphicalProperties.line = LineProperties(solidFill="2E7D32", w=28000)
    chart.series.append(s_led)

    s_laz = Series(Reference(ws, min_col=5, min_row=2, max_row=n + 1), xref, title="Лазерный режим")
    s_laz.marker = Marker(symbol="none")
    s_laz.graphicalProperties = GraphicalProperties()
    s_laz.graphicalProperties.line = LineProperties(solidFill="C62828", w=28000)
    chart.series.append(s_laz)

    ws.add_chart(chart, "G2")
    return wb_to_base64(wb)


def make_polar_xlsx(recs, r1, i_por=None):
    """Excel-файл P(I_pn) с живыми формулами: I_pn = U_ПЛ/R1, P = (Imax−Imin)/(Imax+Imin).

    recs — [(i_pn, p, u, imax, imin), ...], отсортированные по I_pn.
    i_por — пороговый ток из части 1 (записывается в блок параметров).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Данные"
    wb.calculation.fullCalcOnLoad = True
    ws.append(["№", "U_ПЛ, мВ", "I_pn, мА", "I_max, мкА", "I_min, мкА", "P"])
    n = len(recs)
    prow = n + 3
    r_r1 = prow + 1
    for i, (ipn, p, u, imax, imin) in enumerate(recs, start=2):
        ws.append([i - 1, round(float(u), 3),
                   f"=B{i}/$B${r_r1}",
                   round(float(imax), 3), round(float(imin), 3),
                   f"=(D{i}-E{i})/(D{i}+E{i})"])
    ws.cell(row=prow, column=1, value="Параметр")
    ws.cell(row=prow, column=2, value="Значение")
    ws.cell(row=r_r1, column=1, value="R_1, Ом")
    ws.cell(row=r_r1, column=2, value=float(r1))
    ws.cell(row=r_r1 + 1, column=1, value="I_пор, мА (из части 1)")
    ws.cell(row=r_r1 + 1, column=2,
            value=(round(float(i_por), 4) if i_por is not None else "— (постройте часть 1)"))
    ws.column_dimensions["A"].width = 8
    for col in range(2, 7):
        ws.column_dimensions[get_column_letter(col)].width = 16

    chart = ScatterChart()
    chart.title = "Зависимость P(I_pn)"
    chart.x_axis.title = "I_pn, мА"
    chart.y_axis.title = "P"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.height = 12
    chart.width = 22
    xref = Reference(ws, min_col=3, min_row=2, max_row=n + 1)
    s = Series(Reference(ws, min_col=6, min_row=2, max_row=n + 1), xref, title="Степень поляризации P")
    s.marker = Marker(symbol="circle", size=7)
    s.graphicalProperties = GraphicalProperties()
    s.graphicalProperties.line = LineProperties(solidFill="6A1B9A", w=20000)
    chart.series.append(s)
    ws.add_chart(chart, "H2")
    return wb_to_base64(wb)


# Табличное значение r_33 для ниобата лития (м/В), λ=633 нм (из методички).
R33_TABLE = 30.8e-12


def lab2_transform(u_list, i_list, i0, lam, d, l, n0):
    """Формирует строки таблицы: U, I, I/I0, arcsin√(I/I0).
    Пропускает точки с I/I0 > 1 (физически невозможно — clipping)."""
    rows = []
    for u, iv in zip(u_list, i_list):
        if iv is None or u is None:
            continue
        ratio = iv / i0 if i0 else None
        a = None
        if ratio is not None and 0.0 <= ratio <= 1.0:
            a = float(np.arcsin(np.sqrt(ratio)))
        elif ratio is not None and ratio > 1.0:
            a = None
        rows.append({"u": float(u), "i": float(iv), "ratio": ratio, "arcsin": a})
    return rows


def lab2_fit(rows):
    """Линейная регрессия arcsin√(I/I0) = m·U + b по точкам с валидным arcsin."""
    xs, ys = [], []
    for r in rows:
        if r["arcsin"] is not None:
            xs.append(r["u"])
            ys.append(r["arcsin"])
    if len(xs) < 2:
        return {"error": "Недостаточно точек с валидным arcsin√(I/I0) для линейной аппроксимации (нужно минимум 2; проверьте I0 и отсутствие значений I>I0)."}
    x = np.array(xs)
    y = np.array(ys)
    m, b0 = np.polyfit(x, y, 1)
    yfit = m * x + b0
    sse = float(np.sum((yfit - y) ** 2))
    sst = float(np.sum((y - np.mean(y)) ** 2)) if len(y) > 1 else 1.0
    r2 = 1.0 - sse / sst if sst > 0 else 1.0
    return {"m": float(m), "b": float(b0), "r2": float(r2),
            "n": int(len(x)), "x": x.tolist(), "y": y.tolist(), "yfit": yfit.tolist()}


def lab2_results(fit, lam, d, l, n0):
    """Полуволновое напряжение U_λ/2 и электрооптический коэффициент r."""
    if "error" in fit:
        return fit
    m = fit["m"]
    m_abs = abs(m)
    n3 = n0 ** 3
    # знак m — направление ветви (m<0: I убывает с ростом U); физические
    # величины (U_λ/2, r) определяются модулем наклона
    u_half = float(np.pi / (2.0 * m_abs)) if m_abs > 0 else None
    # r = 2 |m| λ d / (π n0^3 l)  [м/В], если λ,d,l — в одинаковых единицах (м)
    r = float(2.0 * m_abs * lam * d / (np.pi * n3 * l)) if (m_abs > 0 and n3 != 0 and l != 0) else None
    return {
        "m": m, "m_abs": m_abs, "u_half": u_half, "r": r,
        "r_table": R33_TABLE,
        "r_ratio": (r / R33_TABLE) if (r is not None and R33_TABLE) else None,
    }


def make_IU_plot(u, i, i0, fit):
    """График I(U) — экспериментальные точки + теоретическая sin²(½ΓU)."""
    fig, ax = plt.subplots(figsize=(9, 5.8), constrained_layout=True)
    ax.plot(u, i, "o", color="#1f4e79", markersize=6, label="Эксперимент")
    if "error" not in fit:
        m = fit["m"]
        b = fit["b"]
        umin, umax = float(min(u)), float(max(u))
        xs = np.linspace(umin, umax, 400)
        # обратное преобразование линеаризации: I = I0 · sin²(m·U + b);
        # знак m и свободный член b задают ветвь (при m<0 кривая убывает от I0)
        ax.plot(xs, i0 * np.sin(m * xs + b) ** 2, "-", color="#c62828", linewidth=1.8,
                label="Аппроксимация: I = I$_0$·sin²(m·U + b)")
        # отметка U_λ/2
        u_half = float(np.pi / (2.0 * abs(m))) if m != 0 else None
        if u_half is not None and umin <= u_half <= umax:
            ax.axvline(u_half, color="#2e7d32", linestyle="--", linewidth=1.3)
            ax.annotate(f"U$_{{λ/2}}$ ≈ {u_half:.1f} В", xy=(u_half, i0),
                        xytext=(8, -18), textcoords="offset points", fontsize=10,
                        color="#2e7d32",
                        bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#2e7d32", alpha=0.9))
    ax.set_xlabel("U, В", fontsize=12)
    ax.set_ylabel("I, мкА", fontsize=12)
    ax.set_title("Зависимость интенсивности света от напряжения на кристалле", fontsize=12)
    ax.grid(True, alpha=0.3)
    ax.legend(loc="best", fontsize=10)
    return fig_to_base64(fig)


def make_IU_xlsx(u, i, i0, fit, res, lam_nm, d_mm, l_mm, n0, ne=None):
    """Excel-файл I(U) с живыми формулами: I/I0, arcsin√(I/I0), m/b/R², U_λ/2, r_33.

    Расчётные столбцы и параметры — формулы Excel со ссылками на ячейки
    (НАКЛОН/ОТРЕЗОК/RSQ по столбцу arcsin; r_33 — по формуле методички
    с конверсией нм/мм → м). В русском Excel: =НАКЛОН(...), =ОТРЕЗОК(...).
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Данные"
    wb.calculation.fullCalcOnLoad = True
    ws.append(["№", "U, В", "I, мкА", "I/I0",
               "arcsin√(I/I0), рад", "arcsin√(I/I0), град"])
    n = len(u)
    prow = n + 3
    r_i0 = prow + 1
    r_lam, r_d, r_l, r_n0, r_ne, r_rt = prow + 2, prow + 3, prow + 4, prow + 5, prow + 6, prow + 7
    r_m, r_b, r_r2, r_uh, r_r, r_ratio = prow + 8, prow + 9, prow + 10, prow + 11, prow + 12, prow + 13
    has_fit = fit and "error" not in fit
    for k, (uk, ik) in enumerate(zip(u, i), start=2):
        ws.append([k - 1, round(float(uk), 3), round(float(ik), 3),
                   f"=C{k}/$B${r_i0}",
                   f'=IF(D{k}<=1,ASIN(SQRT(D{k})),"")',
                   f'=IF(E{k}="","",DEGREES(E{k}))'])
    ws.cell(row=prow, column=1, value="Параметр")
    ws.cell(row=prow, column=2, value="Значение")
    ws.cell(row=r_i0, column=1, value="I0, мкА")
    ws.cell(row=r_i0, column=2, value=float(i0))
    ws.cell(row=r_lam, column=1, value="λ, нм")
    ws.cell(row=r_lam, column=2, value=float(lam_nm))
    ws.cell(row=r_d, column=1, value="d, мм")
    ws.cell(row=r_d, column=2, value=float(d_mm))
    ws.cell(row=r_l, column=1, value="l, мм")
    ws.cell(row=r_l, column=2, value=float(l_mm))
    ws.cell(row=r_n0, column=1, value="n0")
    ws.cell(row=r_n0, column=2, value=float(n0))
    ws.cell(row=r_ne, column=1, value="n_e")
    ws.cell(row=r_ne, column=2, value=(float(ne) if ne is not None else "— (не задано)"))
    ws.cell(row=r_rt, column=1, value="r_33 (табл.), пм/В")
    ws.cell(row=r_rt, column=2, value=30.8)
    if has_fit:
        e_rng, b_rng = f"E2:E{n + 1}", f"B2:B{n + 1}"
        ws.cell(row=r_m, column=1, value="m (наклон), рад/В (НАКЛОН)")
        ws.cell(row=r_m, column=2, value=f"=SLOPE({e_rng},{b_rng})")
        ws.cell(row=r_b, column=1, value="b (смещение), рад (ОТРЕЗОК)")
        ws.cell(row=r_b, column=2, value=f"=INTERCEPT({e_rng},{b_rng})")
        ws.cell(row=r_r2, column=1, value="R² (RSQ)")
        ws.cell(row=r_r2, column=2, value=f"=RSQ({e_rng},{b_rng})")
        ws.cell(row=r_uh, column=1, value="U_λ/2 (расч.), В")
        ws.cell(row=r_uh, column=2, value=f"=PI()/(2*ABS(B{r_m}))")
        # r_33 = 2·m·λ·d / (π·n0³·l);  λ в нм, d/l в мм → результат в пм/В
        ws.cell(row=r_r, column=1, value="r_33 (эксп.), пм/В")
        ws.cell(row=r_r, column=2,
                value=(f"=1000*2*ABS(B{r_m})*B{r_lam}*B{r_d}/(PI()*B{r_n0}^3*B{r_l})"))
        ws.cell(row=r_ratio, column=1, value="r_эксп / r_табл")
        ws.cell(row=r_ratio, column=2, value=f"=B{r_r}/B{r_rt}")
    else:
        for rr, name in ((r_m, "m (наклон), рад/В"), (r_b, "b (смещение), рад"),
                         (r_r2, "R²"), (r_uh, "U_λ/2 (расч.), В"),
                         (r_r, "r_33 (эксп.), пм/В"), (r_ratio, "r_эксп / r_табл")):
            ws.cell(row=rr, column=1, value=name)
            ws.cell(row=rr, column=2, value="— (нет аппроксимации)")
    for col in range(1, 7):
        ws.column_dimensions[get_column_letter(col)].width = 22

    chart = ScatterChart()
    chart.title = "Зависимость I(U)"
    chart.x_axis.title = "U, В"
    chart.y_axis.title = "I, мкА"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.height = 12
    chart.width = 22
    xref = Reference(ws, min_col=2, min_row=2, max_row=n + 1)
    s = Series(Reference(ws, min_col=3, min_row=2, max_row=n + 1), xref, title="I(U)")
    s.marker = Marker(symbol="circle", size=7)
    s.graphicalProperties = GraphicalProperties()
    s.graphicalProperties.line = LineProperties(solidFill="1F4E79", w=20000)
    chart.series.append(s)
    ws.add_chart(chart, "I2")
    return wb_to_base64(wb)


def make_arcsin_plot(rows, fit):
    """График arcsin(I/I0) = f(U) — эксперимент + линия МНК."""
    fig, ax = plt.subplots(figsize=(9, 5.8), constrained_layout=True)
    xs = [r["u"] for r in rows if r["arcsin"] is not None]
    ys = [r["arcsin"] for r in rows if r["arcsin"] is not None]
    ax.plot(xs, ys, "o", color="#1f4e79", markersize=6, label="arcsin√(I/I$_0$) (эксп.)")
    if "error" not in fit and xs:
        m, b0 = fit["m"], fit["b"]
        xa = np.array([float(min(xs)), float(max(xs))])
        ax.plot(xa, m * xa + b0, "-", color="#c62828", linewidth=1.8,
                label=f"МНК: arcsin√(I/I$_0$) = {m:.4g}·U {'+' if b0 >= 0 else '−'} {abs(b0):.4g}")
    ax.set_xlabel("U, В", fontsize=12)
    ax.set_ylabel("arcsin√(I/I$_0$), рад", fontsize=12)
    ax.set_title("Линеаризованная зависимость arcsin√(I/I$_0$) = f(U)", fontsize=12)
    ax.grid(True, alpha=0.3)
    ax.legend(loc="best", fontsize=10)
    return fig_to_base64(fig)


def make_arcsin_xlsx(rows, fit):
    """Excel-файл arcsin(I/I0)=f(U): данные + параметры + диаграмма с линией МНК."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Линеаризация"
    ws.append(["№", "U, В", "arcsin√(I/I0), рад"])
    nrow = 0
    for r in rows:
        if r["arcsin"] is None:
            continue
        nrow += 1
        ws.append([nrow, round(r["u"], 3), round(r["arcsin"], 5)])
    # линия МНК — опорные точки формулами m·x + b (ссылки на ячейки параметров ниже)
    fit_start_row = nrow + 1
    fit_end_row = fit_start_row
    has_fit = "error" not in fit and nrow >= 2
    if has_fit:
        m, b0 = fit["m"], fit["b"]
        xa = [float(min(r["u"] for r in rows if r["arcsin"] is not None)),
              float(max(r["u"] for r in rows if r["arcsin"] is not None))]
        prow = nrow + 6
        r_m, r_b, r_r2 = prow + 1, prow + 2, prow + 3
        ws.append(["", round(xa[0], 3), f"=$B${r_m}*B{nrow + 2}+$B${r_b}"])
        ws.append(["", round(xa[1], 3), f"=$B${r_m}*B{nrow + 3}+$B${r_b}"])
        fit_end_row = fit_start_row + 2
    else:
        prow = max(nrow + 4, fit_start_row + 2)
        r_m = r_b = r_r2 = None
    ws.cell(row=prow, column=1, value="Параметр")
    ws.cell(row=prow, column=2, value="Значение")
    ws.cell(row=prow + 1, column=1, value="m (наклон), рад/В (НАКЛОН)")
    ws.cell(row=prow + 2, column=1, value="b (смещение), рад (ОТРЕЗОК)")
    ws.cell(row=prow + 3, column=1, value="R² (RSQ)")
    ws.cell(row=prow + 4, column=1, value="число точек")
    if has_fit:
        e_rng, b_rng = f"C2:C{nrow + 1}", f"B2:B{nrow + 1}"
        ws.cell(row=prow + 1, column=2, value=f"=SLOPE({e_rng},{b_rng})")
        ws.cell(row=prow + 2, column=2, value=f"=INTERCEPT({e_rng},{b_rng})")
        ws.cell(row=prow + 3, column=2, value=f"=RSQ({e_rng},{b_rng})")
    else:
        ws.cell(row=prow + 1, column=2, value="—")
        ws.cell(row=prow + 2, column=2, value="—")
        ws.cell(row=prow + 3, column=2, value="—")
    ws.cell(row=prow + 4, column=2, value=nrow)
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 22

    chart = ScatterChart()
    chart.title = "arcsin√(I/I0) = f(U)"
    chart.x_axis.title = "U, В"
    chart.y_axis.title = "arcsin√(I/I0), рад"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.height = 12
    chart.width = 22
    xref = Reference(ws, min_col=2, min_row=2, max_row=nrow + 1)
    s_exp = Series(Reference(ws, min_col=3, min_row=2, max_row=nrow + 1), xref, title="arcsin (эксп.)")
    s_exp.marker = Marker(symbol="circle", size=7)
    s_exp.graphicalProperties = GraphicalProperties()
    s_exp.graphicalProperties.line.noFill = True
    chart.series.append(s_exp)
    if fit_end_row > fit_start_row:
        xfit = Reference(ws, min_col=2, min_row=fit_start_row + 1, max_row=fit_end_row)
        yfit = Reference(ws, min_col=3, min_row=fit_start_row + 1, max_row=fit_end_row)
        s_fit = Series(yfit, xfit, title="МНК")
        s_fit.marker = Marker(symbol="none")
        s_fit.graphicalProperties = GraphicalProperties()
        s_fit.graphicalProperties.line = LineProperties(solidFill="C62828", w=28000)
        chart.series.append(s_fit)
    ws.add_chart(chart, "E2")
    return wb_to_base64(wb)



def _f3(v):
    """Безопасный парсинг числа из входа JSON."""
    s = str(v).strip() if v is not None else ""
    if s == "" or s in ("None", "—"):
        return None
    try:
        return float(s)
    except (TypeError, ValueError):
        return None


def _fmt3(v, nd=4):
    """Формат числа для JS: округлённое число или '—'."""
    return round(float(v), nd) if v is not None else "—"


def lab3_linear_range(ipn, tau, tol_rel=0.05):
    """Определяет линейный диапазон по постоянству τ.

    Линейный диапазон — непрерывный участок от минимального I_pn до точки,
    где τ впервые отклоняется от опорного значения (медиана по первым
    min(3, n) точкам) более чем на tol_rel·|τ_ref|. Возвращает кортеж
    (idx_sorted, ipn_min, ipn_max, tau_ref, tol).
    """
    n = len(ipn)
    if n == 0:
        return [], None, None, None, 0.0
    k0 = min(3, n)
    tau_ref = float(np.median(tau[:k0]))
    tol = abs(tau_ref) * tol_rel if tau_ref != 0 else tol_rel
    idx = []
    for i in range(n):
        if abs(tau[i] - tau_ref) <= tol:
            idx.append(i)
        else:
            break
    if not idx:
        idx = [0]
    return idx, float(ipn[idx[0]]), float(ipn[idx[-1]]), tau_ref, tol


def lab3_fit(ipn, u_fd, idx):
    """Линейная аппроксимация U_ФД = a·I_pn + b по точкам с индексами idx."""
    x = np.array([ipn[i] for i in idx], dtype=float)
    y = np.array([u_fd[i] for i in idx], dtype=float)
    if len(x) < 2:
        a, b = (float(y[0]) / x[0], 0.0) if x[0] != 0 else (0.0, 0.0)
        return {"a": a, "b": b, "r2": 1.0, "x": x.tolist(), "yfit": (a * x + b).tolist()}
    a, b = np.polyfit(x, y, 1)
    yfit = a * x + b
    sse = float(np.sum((yfit - y) ** 2))
    sst = float(np.sum((y - np.mean(y)) ** 2))
    r2 = 1.0 - sse / sst if sst > 0 else 1.0
    return {"a": float(a), "b": float(b), "r2": float(r2),
            "x": x.tolist(), "yfit": yfit.tolist()}


def make_detector_plot(ipn, u_fd, tau, fit, lr, title):
    """График U_ФД(I_pn) (левая ось) и τ(I_pn) (правая ось) с линейным диапазоном."""
    fig, ax1 = plt.subplots(figsize=(9, 5.8), constrained_layout=True)
    ax2 = ax1.twinx()
    ipn_min, ipn_max = lr[1], lr[2]
    if ipn_min is not None and ipn_max is not None:
        ax1.axvspan(ipn_min, ipn_max, alpha=0.13, color="#2e7d32", zorder=0)
    ax1.plot(ipn, u_fd, "o", color="#1f4e79", markersize=6, zorder=5,
             label="U$_{ФД}$ (эксп.)")
    if fit and "error" not in fit:
        xf = np.array([float(min(ipn)), float(max(ipn))])
        b_sign = "+" if fit["b"] >= 0 else "−"
        ax1.plot(xf, fit["a"] * xf + fit["b"], "--", color="#1f4e79", linewidth=1.6,
                 label=f"Лин. аппр.: U = {fit['a']:.2f}·I {b_sign} {abs(fit['b']):.2f}")
    ax1.set_xlabel("I$_{pn}$, мА", fontsize=12)
    ax1.set_ylabel("U$_{ФД}$, мВ", fontsize=12, color="#1f4e79")
    ax1.tick_params(axis="y", labelcolor="#1f4e79")
    ax1.grid(True, alpha=0.3)
    ax2.plot(ipn, tau, "s", color="#c62828", markersize=6, zorder=5,
             label="τ = U*$_{ФД}$/U$_{ФД}$")
    ax2.set_ylabel("τ = U*$_{ФД}$ / U$_{ФД}$", fontsize=12, color="#c62828")
    ax2.tick_params(axis="y", labelcolor="#c62828")
    h1, l1 = ax1.get_legend_handles_labels()
    h2, l2 = ax2.get_legend_handles_labels()
    ax1.legend(h1 + h2, l1 + l2, loc="best", fontsize=9)
    ax1.set_title(title, fontsize=12)
    if ipn_min is not None and ipn_max is not None:
        ax1.text(0.03, 0.95, f"Линейный диапазон:\nI$_{{pn}}$ = {ipn_min:.2f} … {ipn_max:.2f} мА",
                 transform=ax1.transAxes, fontsize=9, color="#2e7d32", va="top",
                 bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#2e7d32", alpha=0.9))
    return fig_to_base64(fig)


def make_detector_xlsx(ipn, u_fd, us_fd, tau, fit, lr, idx, title):
    """Excel-файл фотодетектора с живыми формулами: τ = U*/U, лин. аппроксимация,
    a/b/R² (НАКЛОН/ОТРЕЗОК/RSQ по линейному диапазону), τ опорное (МЕДИАНА)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Данные"
    wb.calculation.fullCalcOnLoad = True
    ws.append(["№", "I_pn, мА", "U_ФД, мВ", "U*_ФД, мВ", "τ = U*_ФД/U_ФД",
               "U_ФД (лин. аппр.)"])
    n = len(ipn)
    has_fit = bool(fit) and "error" not in fit
    prow = n + 3
    r_a, r_b, r_r2 = prow + 1, prow + 2, prow + 3
    r_ipn1, r_ipn2, r_tref = prow + 4, prow + 5, prow + 6
    for k, (xi, yi, si, ti) in enumerate(zip(ipn, u_fd, us_fd, tau), start=2):
        row = [k - 1, round(float(xi), 4), round(float(yi), 4), round(float(si), 4),
               f"=D{k}/C{k}"]
        row.append(f"=$B${r_a}*B{k}+$B${r_b}" if has_fit else "—")
        ws.append(row)
    ws.cell(row=prow, column=1, value="Параметр")
    ws.cell(row=prow, column=2, value="Значение")
    ws.cell(row=r_ipn1, column=1, value="I_pn мин (лин.), мА")
    ws.cell(row=r_ipn1, column=2, value=(round(lr[1], 4) if lr[1] is not None else "—"))
    ws.cell(row=r_ipn2, column=1, value="I_pn макс (лин.), мА")
    ws.cell(row=r_ipn2, column=2, value=(round(lr[2], 4) if lr[2] is not None else "—"))
    # τ опорное = медиана τ по первым min(3, n) точкам (колонка τ = E)
    k0 = min(3, n)
    ws.cell(row=r_tref, column=1, value="τ опорное (МЕДИАНА)")
    ws.cell(row=r_tref, column=2, value=f"=MEDIAN(E2:E{k0 + 1})")
    if has_fit and idx and lr[1] is not None:
        # линейный диапазон — непрерывный блок строк idx[0]+2 .. idx[-1]+2
        a1, b1 = idx[0] + 2, idx[-1] + 2
        rng = f"C{a1}:C{b1}"
        rngx = f"B{a1}:B{b1}"
        ws.cell(row=r_a, column=1, value="Чувствительность a, мВ/мА (НАКЛОН)")
        ws.cell(row=r_a, column=2, value=f"=SLOPE({rng},{rngx})")
        ws.cell(row=r_b, column=1, value="Смещение b, мВ (ОТРЕЗОК)")
        ws.cell(row=r_b, column=2, value=f"=INTERCEPT({rng},{rngx})")
        ws.cell(row=r_r2, column=1, value="R² лин. диапазон (RSQ)")
        ws.cell(row=r_r2, column=2, value=f"=RSQ({rng},{rngx})")
    else:
        ws.cell(row=r_a, column=1, value="Чувствительность a, мВ/мА")
        ws.cell(row=r_a, column=2, value="—")
        ws.cell(row=r_b, column=1, value="Смещение b, мВ")
        ws.cell(row=r_b, column=2, value="—")
        ws.cell(row=r_r2, column=1, value="R² лин. диапазон")
        ws.cell(row=r_r2, column=2, value="—")
    for col in range(1, 7):
        ws.column_dimensions[get_column_letter(col)].width = 22

    # диаграмма 1: U_ФД от I_pn (эксп. + линия аппроксимации)
    chart = ScatterChart()
    chart.title = title
    chart.x_axis.title = "I_pn, мА"
    chart.y_axis.title = "U_ФД, мВ"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.height = 11
    chart.width = 22
    xref = Reference(ws, min_col=2, min_row=2, max_row=n + 1)
    s_exp = Series(Reference(ws, min_col=3, min_row=2, max_row=n + 1), xref, title="U_ФД (эксп.)")
    s_exp.marker = Marker(symbol="circle", size=7)
    s_exp.graphicalProperties = GraphicalProperties()
    s_exp.graphicalProperties.line.noFill = True
    chart.series.append(s_exp)
    if has_fit:
        s_fit = Series(Reference(ws, min_col=6, min_row=2, max_row=n + 1), xref, title="Лин. аппр.")
        s_fit.marker = Marker(symbol="none")
        s_fit.graphicalProperties = GraphicalProperties()
        s_fit.graphicalProperties.line = LineProperties(solidFill="1F4E79", w=20000)
        chart.series.append(s_fit)
    ws.add_chart(chart, "H2")

    # диаграмма 2: τ от I_pn
    chart2 = ScatterChart()
    chart2.title = title + " — τ"
    chart2.x_axis.title = "I_pn, мА"
    chart2.y_axis.title = "τ"
    chart2.x_axis.delete = False
    chart2.y_axis.delete = False
    chart2.height = 9
    chart2.width = 22
    s_tau = Series(Reference(ws, min_col=5, min_row=2, max_row=n + 1), xref, title="τ")
    s_tau.marker = Marker(symbol="square", size=7)
    s_tau.graphicalProperties = GraphicalProperties()
    s_tau.graphicalProperties.line = LineProperties(solidFill="C62828", w=18000)
    chart2.series.append(s_tau)
    ws.add_chart(chart2, "H26")
    return wb_to_base64(wb)


def build_lab3_analysis(results, configs, r1, r2, u_sm, tol_pct):
    """HTML-панель с таблицей диапазонов линейности и выводами."""
    html = '<h4>Определение диапазона линейности фотодетектора</h4>'
    html += ('<table class="res"><tr><th>Режим</th><th>Чувствительность a, '
             'мВ/мА</th><th>R²</th><th>Линейный диапазон I<sub>pn</sub>, мА</th>'
             '<th>τ опорное</th><th>Точек</th></tr>')
    ranges = {}
    for key, ipn, u, tau, title in configs:
        res = results.get(key, {})
        if "error" in res:
            html += f'<tr><td>{title}</td><td colspan="5">{res["error"]}</td></tr>'
            continue
        fit = res["fit"]
        lr = res["lr"]
        ranges[key] = (lr[1], lr[2])
        b_sign = "+" if fit["b"] >= 0 else "−"
        html += (f'<tr><td>{title}</td><td>{fit["a"]:.3f}</td>'
                 f'<td>{fit["r2"]:.4f}</td>'
                 f'<td>{lr[1]:.3f} … {lr[2]:.3f}</td>'
                 f'<td>{lr[3]:.4f}</td><td>{res["n"]}</td></tr>')
    html += '</table>'
    html += (f'<p class="note" style="color:#555;margin-top:8px;">Критерий линейности: '
             f'τ = U*<sub>ФД</sub>/U<sub>ФД</sub> отклоняется от опорного значения '
             f'(медиана по точкам с малым I<sub>pn</sub>) не более чем на '
             f'{tol_pct:g}%. Линейный диапазон — от минимального I<sub>pn</sub> '
             f'до точки, где τ начинает расти (насыщение фотодиода).</p>')

    # Выводы (сравнение диапазонов)
    html += '<h4 style="margin-top:14px;">Выводы</h4><ul style="font-size:13.5px;line-height:1.6;">'
    g1r = ranges.get("g1")
    g2r = ranges.get("g2")
    g3r = ranges.get("g3")
    if g1r and g2r:
        if g2r[1] < g1r[1]:
            html += (f'<li>В фотогальваническом режиме с <b>большим</b> сопротивлением '
                     f'нагрузки (R₂={r2:g} Ом) линейный диапазон '
                     f'(I<sub>pn</sub> ≤ {g2r[1]:.2f} мА) уже, чем с малым '
                     f'(R₁={r1:g} Ом, I<sub>pn</sub> ≤ {g1r[1]:.2f} мА): бóльшая нагрузка '
                     f'переводит фотодиод ближе к режиму холостого хода, где напряжение '
                     f'насыщается при меньшей интенсивности излучения.</li>')
        else:
            html += (f'<li>В фотогальваническом режиме линейные диапазоны для R₁={r1:g} Ом '
                     f'(до {g1r[1]:.2f} мА) и R₂={r2:g} Ом (до {g2r[1]:.2f} мА) близки.</li>')
    if g3r and g1r:
        if g3r[1] >= g1r[1]:
            html += (f'<li>В <b>фотодиодном режиме</b> (обратное смещение U<sub>см</sub>={u_sm:g} В) '
                     f'линейный диапазон (I<sub>pn</sub> ≤ {g3r[1]:.2f} мА) шире, чем в '
                     f'фотогальваническом: обратное смещение расширяет запирающий слой, и '
                     f'фотодиод работает как генератор тока в широком динамическом диапазоне.</li>')
        else:
            html += (f'<li>В фотодиодном режиме линейный диапазон (до {g3r[1]:.2f} мА) '
                     f'сопоставим с фотогальваническим (R₁, до {g1r[1]:.2f} мА).</li>')
    html += ('<li>Для измерения интенсивности лазерного излучения в широком диапазоне '
             'следует использовать <b>фотодиодный режим</b> с обратным смещением и малым '
             'сопротивлением нагрузки; для снижения шумов в ряде случаев выгоден '
             'фотогальванический режим с малой нагрузкой.</li>')
    html += '</ul>'
    return html


# ===========================================================================
#  Маршруты страниц курса
# ===========================================================================

KIOE_STUB_HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Лабораторная работа не найдена</title>
<style>
  body { margin:0; font-family:'Segoe UI', system-ui, Arial, sans-serif; background:#f4f6f9; color:#222; }
  header { background:#1f4e79; color:#fff; padding:22px 20px; }
  main { max-width:1000px; margin:40px auto; padding:0 16px; }
  .box { background:#fff; border:1px solid #d8dee9; border-left:4px solid #d4a72c; border-radius:8px; padding:26px 30px; }
</style>
</head>
<body>
<header><h1>Квантовая и оптическая электроника</h1></header>
<main><div class="box">
  <h2>Лабораторная работа не найдена</h2>
  <p>В курсе «Квантовая и оптическая электроника» реализованы работы № 1–3. Вернитесь на <a href="/">главную страницу</a>.</p>
</div></main>
</body>
</html>
"""


@bp.route("/kioe/lab/<int:lab_id>")
def kioe_lab(lab_id):
    if lab_id == 1:
        return LAB1_HTML
    if lab_id == 2:
        return LAB2_HTML
    if lab_id == 3:
        return LAB3_HTML
    return KIOE_STUB_HTML


# ===========================================================================
#  API лабораторных работ
# ===========================================================================



@bp.route("/api/kioe/lab1/calc_ipn", methods=["POST"])
def api_calc_ipn():
    try:
        d = request.get_json(force=True)
        r1 = float(d["r1"])
        if r1 <= 0:
            return jsonify({"error": "R1 должно быть положительным."}), 400
        u = [float(v) for v in d["u"] if str(v).strip() not in ("", "None")]
        if not u:
            return jsonify({"error": "Нет значений U_ПЛ для расчёта."}), 400
        i_pn = [round(uv / r1, 4) for uv in u]
        samples = [{"u": u[k], "r1": r1, "i": i_pn[k]} for k in range(min(3, len(u)))]
        wrows = [(f"I<sub>pn,{k}</sub> = U<sub>ПЛ,{k}</sub> / R<sub>1</sub>",
                  f"{fmt_n(u[k])} / {fmt_n(r1)}",
                  f"{fmt_n(i_pn[k])} мА") for k in range(len(u))]
        work = work_block(wrows, note="U<sub>ПЛ</sub> в мВ, R<sub>1</sub> в Ом → "
                          "I<sub>pn</sub> в мА (мВ/Ом = мА). Расчёт по всем заполненным точкам.")
        return jsonify({"i_pn": i_pn, "samples": samples, "r1": r1, "work_html": work})
    except (ValueError, KeyError, TypeError) as e:
        return jsonify({"error": f"Ошибка данных: {e}"}), 400


@bp.route("/api/kioe/lab1/plot_intensity", methods=["POST"])
def api_plot_intensity():
    try:
        d = request.get_json(force=True)
        i_pn = [float(v) for v in d["i_pn"]]
        i_fp = [float(v) for v in d["i_fp"]]
        rows = [int(r) for r in (d.get("rows") or list(range(1, len(i_pn) + 1)))]
        n = min(len(i_pn), len(i_fp), len(rows))
        i_pn, i_fp, rows = i_pn[:n], i_fp[:n], rows[:n]
        if n < 4:
            return jsonify({"error": "Нужно минимум 4 точки для построения и аппроксимации."}), 400
        fit = lab1_intensity_fit(i_pn, i_fp, rows, str(d.get("laser") or ""))
        if "error" in fit:
            return jsonify({"error": fit["error"]}), 400
        # сортировка по I_pn — маски точек следуют за своими точками
        order = list(np.argsort(np.array(i_pn)))
        xs = [i_pn[i] for i in order]
        ys = [i_fp[i] for i in order]
        fit["led_mask"] = fit["led_mask"][order]
        fit["laz_mask"] = fit["laz_mask"][order]
        img = make_intensity_plot(xs, ys, fit)
        xlsx = make_intensity_xlsx(xs, ys, fit)

        def _sums_block(xv, yv):
            x = np.asarray(xv, dtype=float)
            y = np.asarray(yv, dtype=float)
            n_ = len(x)
            Sx, Sy = float(x.sum()), float(y.sum())
            Sxy, Sxx = float((x * y).sum()), float((x * x).sum())
            k_ = (n_ * Sxy - Sx * Sy) / (n_ * Sxx - Sx * Sx)
            b_ = (Sy - k_ * Sx) / n_
            yfit = k_ * x + b_
            sse = float(((y - yfit) ** 2).sum())
            sst = float(((y - y.mean()) ** 2).sum())
            r2_ = 1.0 - sse / sst if sst > 0 else 1.0
            return n_, Sx, Sy, Sxy, Sxx, k_, b_, sse, sst, r2_

        led_x = np.asarray(xs)[np.asarray(fit["led_mask"], dtype=bool)]
        led_y = np.asarray(ys)[np.asarray(fit["led_mask"], dtype=bool)]
        if fit.get("laz_pts_extra"):
            laz_x = [p[0] for p in fit["laz_pts_extra"]]
            laz_y = [p[1] for p in fit["laz_pts_extra"]]
        else:
            laz_x = np.asarray(xs)[np.asarray(fit["laz_mask"], dtype=bool)]
            laz_y = np.asarray(ys)[np.asarray(fit["laz_mask"], dtype=bool)]
        wrows = []
        for tag, xv, yv in (("св", led_x, led_y), ("лаз", laz_x, laz_y)):
            n_, Sx, Sy, Sxy, Sxx, k_, b_, sse, sst, r2_ = _sums_block(xv, yv)
            wrows.append((f"k<sub>{tag}</sub> = (n·Σ(I<sub>pn</sub>·I<sub>ФП</sub>) − ΣI<sub>pn</sub>·ΣI<sub>ФП</sub>) / (n·ΣI<sub>pn</sub>² − (ΣI<sub>pn</sub>)²)",
                          f"({n_}·{fmt_g(Sxy)} − {fmt_g(Sx)}·{fmt_g(Sy)}) / ({n_}·{fmt_g(Sxx)} − ({fmt_g(Sx)})<sup>2</sup>)",
                          f"{fmt_n(k_)} мкА/мА"))
            wrows.append((f"b<sub>{tag}</sub> = (ΣI<sub>ФП</sub> − k<sub>{tag}</sub>·ΣI<sub>pn</sub>) / n",
                          f"({fmt_g(Sy)} − {fmt_n(k_)}·{fmt_g(Sx)}) / {n_}",
                          f"{fmt_n(b_)} мкА"))
            wrows.append((f"R²<sub>{tag}</sub> = 1 − Σ(y − ŷ)² / Σ(y − ȳ)²",
                          f"1 − {fmt_g(sse)} / {fmt_g(sst)}",
                          fmt_n(r2_)))
        if fit["i_por"] is not None:
            wrows.append(("I<sub>пор</sub> = (b<sub>св</sub> − b<sub>лаз</sub>) / (k<sub>лаз</sub> − k<sub>св</sub>)",
                          f"({fmt_n(fit['b_sv'])} − {fmt_n(fit['b_laz'])}) / ({fmt_n(fit['k_laz'])} − {fmt_n(fit['k_sv'])})",
                          f"{fmt_n(fit['i_por'])} мА"))
        work = work_block(wrows, note="МНК-прямая y = k·x + b по точкам соответствующего режима "
                          "(у = I<sub>ФП</sub>, x = I<sub>pn</sub>); I<sub>пор</sub> — пересечение прямых. "
                          "Режимы: " + fit["mode_desc"] + ".")
        return jsonify({
            "image": img,
            "xlsx": xlsx,
            "k_sv": round(fit["k_sv"], 5),
            "b_sv": round(fit["b_sv"], 5),
            "k_laz": round(fit["k_laz"], 5),
            "b_laz": round(fit["b_laz"], 5),
            "r2_sv": round(fit["r2_sv"], 5),
            "r2_laz": round(fit["r2_laz"], 5),
            "i_por": (round(fit["i_por"], 4) if fit["i_por"] is not None else None),
            "mode_desc": fit["mode_desc"],
            "n": fit["n"],
            "work_html": work,
        })
    except (ValueError, KeyError, TypeError) as e:
        return jsonify({"error": f"Ошибка данных: {e}. Проверьте таблицу и поле "
                                 f"«Ввести значения»."}), 400
    except Exception as e:
        return jsonify({"error": f"Ошибка: {e}"}), 400


@bp.route("/api/kioe/lab1/calc_polar", methods=["POST"])
def api_calc_polar():
    try:
        d = request.get_json(force=True)
        r1 = float(d["r1"])
        if r1 <= 0:
            return jsonify({"error": "R1 должно быть положительным."}), 400
        rows_in = d["rows"]
        out = []
        for r in rows_in:
            u_s = str(r.get("u", "")).strip()
            imax_s = str(r.get("imax", "")).strip()
            imin_s = str(r.get("imin", "")).strip()
            u = float(u_s) if u_s else None
            imax = float(imax_s) if imax_s else None
            imin = float(imin_s) if imin_s else None
            if u is None and imax is None and imin is None:
                continue
            ipn = round(u / r1, 4) if u is not None else None
            p = None
            if imax is not None and imin is not None and (imax + imin) != 0:
                p = round((imax - imin) / (imax + imin), 4)
            out.append({"u": u, "ipn": ipn, "imax": imax, "imin": imin, "p": p})
        samples = out[:3]
        w_ipn = [(f"I<sub>pn,{k}</sub> = U<sub>ПЛ,{k}</sub> / R<sub>1</sub>",
                  f"{fmt_n(r['u'])} / {fmt_n(r1)}",
                  f"{fmt_n(r['ipn'])} мА")
                 for k, r in enumerate(out, 1) if r["ipn"] is not None]
        w_p = [(f"P<sub>{k}</sub> = (I<sub>max</sub> − I<sub>min</sub>) / (I<sub>max</sub> + I<sub>min</sub>)",
                f"({fmt_n(r['imax'])} − {fmt_n(r['imin'])}) / ({fmt_n(r['imax'])} + {fmt_n(r['imin'])})",
                fmt_n(r["p"]))
               for k, r in enumerate(out, 1) if r["p"] is not None]
        if not w_ipn:
            w_ipn = [("I<sub>pn</sub> = U<sub>ПЛ</sub> / R<sub>1</sub>", "—", "нет заполненных строк")]
        if not w_p:
            w_p = [("P = (I<sub>max</sub> − I<sub>min</sub>) / (I<sub>max</sub> + I<sub>min</sub>)",
                    "—", "нет заполненных пар I<sub>max</sub>/I<sub>min</sub>")]
        work_ipn = work_block(w_ipn, note="U<sub>ПЛ</sub> в мВ, R<sub>1</sub> в Ом → I<sub>pn</sub> в мА.")
        work_p = work_block(w_p, note="I<sub>max</sub>, I<sub>min</sub> — прибор М2, мкА.")
        return jsonify({"rows": out, "samples": samples, "r1": r1,
                        "work_html_ipn": work_ipn, "work_html_p": work_p})
    except Exception as e:
        return jsonify({"error": f"Ошибка: {e}"}), 400


@bp.route("/api/kioe/lab1/plot_polar", methods=["POST"])
def api_plot_polar():
    try:
        d = request.get_json(force=True)
        r1 = float(d["r1"])
        if r1 <= 0:
            return jsonify({"error": "R1 должно быть положительным."}), 400
        recs = []
        for r in d["rows"]:
            u = float(r["u"])
            imax = float(r["imax"])
            imin = float(r["imin"])
            if imax + imin == 0:
                return jsonify({"error": "I_max + I_min не должно быть равно нулю."}), 400
            recs.append((u / r1, (imax - imin) / (imax + imin), u, imax, imin))
        if len(recs) < 2:
            return jsonify({"error": "Нужно минимум 2 точки для построения графика."}), 400
        recs.sort(key=lambda t: t[0])
        i_por = d.get("i_por")
        i_por = float(i_por) if i_por not in (None, "", "None") else None
        img = make_polar_plot([r[0] for r in recs], [r[1] for r in recs], i_por)
        xlsx = make_polar_xlsx(recs, r1, i_por)
        return jsonify({"image": img, "xlsx": xlsx, "n": len(recs), "i_por": i_por})
    except (ValueError, KeyError, TypeError) as e:
        return jsonify({"error": f"Ошибка данных: {e}. Проверьте U_ПЛ, I_max, I_min и R1."}), 400
    except Exception as e:
        return jsonify({"error": f"Ошибка: {e}"}), 400



@bp.route("/api/kioe/lab2/process", methods=["POST"])
def api_lab2_process():
    try:
        d = request.get_json(force=True)
        lam_nm = float(d["lambda_nm"])
        d_mm = float(d["d_mm"])
        l_mm = float(d["l_mm"])
        n0 = float(d["n0"])
        i0 = float(d["i0"])
        ne = _f3(d.get("ne"))
        if i0 <= 0:
            return jsonify({"error": "I0 должно быть положительным."}), 400
        if lam_nm <= 0 or d_mm <= 0 or l_mm <= 0 or n0 <= 0:
            return jsonify({"error": "Параметры λ, d, l, n0 должны быть положительными."}), 400
        if ne is not None and ne <= 0:
            return jsonify({"error": "n_e должно быть положительным."}), 400
        rtable_pm = d.get("rtable")
        rtable_pm = float(rtable_pm) if rtable_pm not in (None, "") else None

        rows_in = d["rows"]
        u_list, i_list = [], []
        for r in rows_in:
            us = str(r.get("u", "")).strip()
            is_ = str(r.get("i", "")).strip()
            if us == "" or is_ == "":
                continue
            u_list.append(float(us))
            i_list.append(float(is_))
        if len(u_list) < 3:
            return jsonify({"error": "Введите минимум 3 точки (U, I)."}), 400

        # Единицы: λ, d, l — приведём к метрам для расчёта r в м/В.
        lam = lam_nm * 1e-9
        d_m = d_mm * 1e-3
        l_m = l_mm * 1e-3

        order = np.argsort(u_list)
        u_sorted = np.array(u_list)[order].tolist()
        i_sorted = np.array(i_list)[order].tolist()

        rows = lab2_transform(u_sorted, i_sorted, i0, lam, d_m, l_m, n0)
        fit = lab2_fit(rows)
        if "error" in fit:
            return jsonify({"error": fit["error"]}), 400
        res = lab2_results(fit, lam, d_m, l_m, n0)

        # Таблица для JS (только ratio и arcsin как строки)
        table_out = []
        for r in rows:
            table_out.append({
                "ratio": (round(r["ratio"], 4) if r["ratio"] is not None else "—"),
                "arcsin": (round(r["arcsin"], 5) if r["arcsin"] is not None else "— (I>I0)"),
            })

        # Панель расчёта (формулы + первые точки)
        calc_html = '<p class="formula">Формулы обработки:</p>'
        calc_html += ('<p class="formula-eq">I = I<sub>0</sub>·sin²(m·U + b), &nbsp; '
                      'arcsin√(I/I<sub>0</sub>) = m·U + b</p>')
        calc_html += ('<p class="formula-eq">U<sub>λ/2</sub> = π / (2|m|), &nbsp; '
                      'r = 2·|m|·λ·d / (π·n<sub>0</sub>³·l)</p>')
        # предупреждение, если часть точек выкинута (I>I0)
        dropped = sum(1 for r in rows if r["arcsin"] is None)
        if dropped:
            calc_html += (f'<p class="note">Внимание: {dropped} точек исключены из '
                         'линеаризации, т.к. I &gt; I<sub>0</sub> (физически недопустимо). '
                         'Проверьте значение I<sub>0</sub> или исходные данные.</p>')

        # Панель результатов
        m = res["m"]
        u_half = res["u_half"]
        r_exp = res["r"]
        r_tab = res["r_table"]
        rratio = res["r_ratio"]
        def sci(x, d=4):
            if x is None:
                return "—"
            return f"{x:.{d}e}".replace("e", "·10").replace("·10+0", "·10⁰").replace("·10+", "·10⁺").replace("·10-", "·10⁻")
        # --- Расчетная часть: все расчёты (формула = подстановка = ответ) ---
        wrows = []
        for r in rows:
            if r["arcsin"] is None:
                continue
            wrows.append(("I / I<sub>0</sub>",
                          f"{fmt_n(r['i'])} / {fmt_n(i0)}",
                          fmt_n(r["ratio"])))
            wrows.append(("arcsin√(I / I<sub>0</sub>)",
                          f"arcsin√({fmt_n(r['i'])} / {fmt_n(i0)})",
                          f"{fmt_n(r['arcsin'])} рад"))
        xfitv = np.array([r["u"] for r in rows if r["arcsin"] is not None], dtype=float)
        yfitv = np.array([r["arcsin"] for r in rows if r["arcsin"] is not None], dtype=float)
        n_ = len(xfitv)
        Sx, Sy = float(xfitv.sum()), float(yfitv.sum())
        Sxy, Sxx = float((xfitv * yfitv).sum()), float((xfitv * xfitv).sum())
        sse = float(((yfitv - fit["m"] * xfitv - fit["b"]) ** 2).sum())
        sst = float(((yfitv - yfitv.mean()) ** 2).sum())
        wrows.append(("m = (n·Σ(U·y) − ΣU·Σy) / (n·ΣU² − (ΣU)²),&nbsp;&nbsp; y = arcsin√(I/I<sub>0</sub>)",
                      f"({n_}·{fmt_g(Sxy)} − {fmt_g(Sx)}·{fmt_g(Sy)}) / ({n_}·{fmt_g(Sxx)} − ({fmt_g(Sx)})<sup>2</sup>)",
                      f"{fmt_n(fit['m'])} рад/В"))
        wrows.append(("b = (Σy − m·ΣU) / n",
                      f"({fmt_g(Sy)} − {fmt_n(fit['m'])}·{fmt_g(Sx)}) / {n_}",
                      f"{fmt_n(fit['b'])} рад"))
        wrows.append(("R² = 1 − Σ(y − ŷ)² / Σ(y − ȳ)²",
                      f"1 − {fmt_g(sse)} / {fmt_g(sst)}",
                      fmt_n(fit["r2"])))
        wrows.append(("U<sub>λ/2</sub> = π / (2|m|)",
                      f"3.14159 / (2·{fmt_n(abs(fit['m']))})",
                      f"{fmt_n(res['u_half'])} В"))
        wrows.append(("r = 2·|m|·λ·d / (π·n<sub>0</sub>³·l)&nbsp;&nbsp;(λ, d, l — в СИ)",
                      (f"2·{fmt_n(abs(fit['m']))}·{fmt_n(lam_nm)}·10<sup>−9</sup>·{fmt_n(d_mm)}·10<sup>−3</sup> / "
                       f"(π·{fmt_n(n0)}<sup>3</sup>·{fmt_n(l_mm)}·10<sup>−3</sup>)"),
                      f"{fmt_n(res['r'] * 1e12)}·10<sup>−12</sup> м/В"))
        note_w = ("МНК-прямая y = m·U + b по всем валидным точкам (y = arcsin√(I/I<sub>0</sub>), рад); "
                  "кривая на графике I(U) — обратное преобразование I = I<sub>0</sub>·sin²(m·U + b). "
                  "U<sub>λ/2</sub> и r вычисляются по |m|: m &lt; 0 означает убывание "
                  "интенсивности с ростом U (ветвь от максимума при U = 0).")
        if dropped:
            note_w += f" Исключено точек: {dropped} (I &gt; I<sub>0</sub>)."
        work = work_block(wrows, note=note_w)

        res_html = "<h4>Результаты обработки</h4>"
        res_html += '<table class="res"><tr><th>Параметр</th><th>Значение</th></tr>'
        res_html += (f'<tr><td>Угловой коэффициент m (наклон; знак — направление ветви)</td>'
                     f'<td>{m:.6g} рад/В</td></tr>')
        res_html += f'<tr><td>Смещение b (МНК)</td><td>{fit["b"]:.6g} рад</td></tr>'
        res_html += f'<tr><td>Коэффициент детерминации R²</td><td>{fit["r2"]:.5f}</td></tr>'
        res_html += f'<tr><td>Число точек в МНК</td><td>{fit["n"]}</td></tr>'
        u_half_str = f"{u_half:.2f} В" if u_half is not None else "—"
        res_html += f'<tr><td><b>Полуволновое напряжение U<sub>λ/2</sub></b></td><td><b>{u_half_str}</b></td></tr>'
        res_html += f'<tr><td><b>Электрооптический коэффициент r (эксп.)</b></td><td><b>{sci(r_exp)} м/В</b></td></tr>'
        if rtable_pm is not None:
            rtab_m = rtable_pm * 1e-12
            r_exp_pm = (r_exp * 1e12) if r_exp is not None else None
            res_html += f'<tr><td>r<sub>33</sub> (табл., LiNbO<sub>3</sub>, λ=633 нм)</td><td>{sci(rtab_m)} м/В ({rtable_pm:g}·10⁻¹²)</td></tr>'
            if r_exp_pm is not None:
                res_html += f'<tr><td><b>r<sub>эксп</sub> / r<sub>табл</sub></b></td><td><b>{(r_exp_pm/rtable_pm):.4f} ({r_exp_pm:.2f} / {rtable_pm:g})</b></td></tr>'
        res_html += '</table>'
        res_html += ('<p class="note" style="color:#555;margin-top:8px;">U<sub>λ/2</sub> — '
                     'напряжение, при котором фазовый сдвиг равен π (максимальное пропускание модулятора). '
                     'r определён из наклона m линеаризованной зависимости arcsin√(I/I<sub>0</sub>) = f(U); '
                     'кривая аппроксимации — обратное преобразование I = I<sub>0</sub>·sin²(m·U + b), '
                     'поэтому она проходит через экспериментальные точки на любой ветви '
                     '(при m &lt; 0 — убывание от максимума при U = 0). U<sub>λ/2</sub> и r — по |m|. '
                     'Сравнение с табличным r<sub>33</sub> ниобата лития показывает точность эксперимента.</p>')

        # Графики и Excel
        iu_img = make_IU_plot(u_sorted, i_sorted, i0, fit)
        iu_xlsx = make_IU_xlsx(u_sorted, i_sorted, i0, fit, res, lam_nm, d_mm, l_mm, n0, ne)
        arcsin_img = make_arcsin_plot(rows, fit)
        arcsin_xlsx = make_arcsin_xlsx(rows, fit)

        return jsonify({
            "table": table_out,
            "calc_html": calc_html,
            "work_html": work,
            "res_html": res_html,
            "iu_image": iu_img,
            "iu_xlsx": iu_xlsx,
            "arcsin_image": arcsin_img,
            "arcsin_xlsx": arcsin_xlsx,
            "m": m,
            "b": fit["b"],
            "u_half": u_half,
            "r": r_exp,
        })
    except Exception as e:
        return jsonify({"error": f"Ошибка: {e}"}), 400



@bp.route("/api/kioe/lab3/process", methods=["POST"])
def api_lab3_process():
    try:
        d = request.get_json(force=True)
        r_pl = float(d["r_pl"])
        if r_pl <= 0:
            return jsonify({"error": "R_ПЛ должно быть положительным."}), 400
        r1 = float(d.get("r1", 180.0))
        r2 = float(d.get("r2", 1500.0))
        u_sm = d.get("u_sm")
        u_sm = float(u_sm) if u_sm not in (None, "") else None
        tol_pct = float(d["tol_pct"])
        if tol_pct <= 0:
            return jsonify({"error": "Допуск на τ должен быть положительным."}), 400
        tol_rel = tol_pct / 100.0

        # Таблица 3.1 (фотогальванический режим)
        t1_out = []
        wk_rows = []  # строки «Расчетной части»
        g1_ipn, g1_u, g1_tau, g1_us = [], [], [], []
        g2_ipn, g2_u, g2_tau, g2_us = [], [], [], []
        for r in d.get("table1", []):
            u_pl = _f3(r.get("u_pl"))
            u_fd1 = _f3(r.get("u_fd_r1"))
            us_fd1 = _f3(r.get("us_fd_r1"))
            u_fd2 = _f3(r.get("u_fd_r2"))
            us_fd2 = _f3(r.get("us_fd_r2"))
            ipn = round(u_pl / r_pl, 4) if u_pl is not None else None
            tau1 = round(us_fd1 / u_fd1, 5) if (u_fd1 and u_fd1 != 0 and us_fd1 is not None) else None
            tau2 = round(us_fd2 / u_fd2, 5) if (u_fd2 and u_fd2 != 0 and us_fd2 is not None) else None
            t1_out.append({"ipn": _fmt3(ipn) if ipn is not None else "—",
                           "tau1": _fmt3(tau1) if tau1 is not None else "—",
                           "tau2": _fmt3(tau2) if tau2 is not None else "—"})
            if ipn is not None:
                wk_rows.append(("I<sub>pn</sub> = U<sub>ПЛ</sub> / R<sub>ПЛ</sub>",
                                f"{fmt_n(u_pl)} / {fmt_n(r_pl)}",
                                f"{fmt_n(ipn)} мА"))
            if tau1 is not None:
                wk_rows.append(("τ<sub>1</sub> = U*<sub>ФД</sub> / U<sub>ФД</sub>&nbsp;&nbsp;(R<sub>1</sub>)",
                                f"{fmt_n(us_fd1)} / {fmt_n(u_fd1)}",
                                fmt_n(tau1)))
            if tau2 is not None:
                wk_rows.append(("τ<sub>2</sub> = U*<sub>ФД</sub> / U<sub>ФД</sub>&nbsp;&nbsp;(R<sub>2</sub>)",
                                f"{fmt_n(us_fd2)} / {fmt_n(u_fd2)}",
                                fmt_n(tau2)))
            if ipn is not None:
                if u_fd1 is not None and tau1 is not None:
                    g1_ipn.append(ipn); g1_u.append(u_fd1); g1_tau.append(tau1)
                    g1_us.append(us_fd1)
                if u_fd2 is not None and tau2 is not None:
                    g2_ipn.append(ipn); g2_u.append(u_fd2); g2_tau.append(tau2)
                    g2_us.append(us_fd2)

        # Таблица 3.2 (фотодиодный режим)
        t2_out = []
        g3_ipn, g3_u, g3_tau, g3_us = [], [], [], []
        for r in d.get("table2", []):
            u_pl = _f3(r.get("u_pl"))
            u_fd = _f3(r.get("u_fd"))
            us_fd = _f3(r.get("us_fd"))
            ipn = round(u_pl / r_pl, 4) if u_pl is not None else None
            tau = round(us_fd / u_fd, 5) if (u_fd and u_fd != 0 and us_fd is not None) else None
            t2_out.append({"ipn": _fmt3(ipn) if ipn is not None else "—",
                           "tau": _fmt3(tau) if tau is not None else "—"})
            if ipn is not None:
                wk_rows.append(("I<sub>pn</sub> = U<sub>ПЛ</sub> / R<sub>ПЛ</sub>",
                                f"{fmt_n(u_pl)} / {fmt_n(r_pl)}",
                                f"{fmt_n(ipn)} мА"))
            if tau is not None:
                wk_rows.append(("τ = U*<sub>ФД</sub> / U<sub>ФД</sub>&nbsp;&nbsp;(фотодиодный режим)",
                                f"{fmt_n(us_fd)} / {fmt_n(u_fd)}",
                                fmt_n(tau)))
            if ipn is not None and u_fd is not None and tau is not None:
                g3_ipn.append(ipn); g3_u.append(u_fd); g3_tau.append(tau)
                g3_us.append(us_fd)

        if not g1_ipn and not g2_ipn and not g3_ipn:
            return jsonify({"error": "Нет данных для построения: заполните хотя бы одну таблицу (U_ПЛ, U_ФД, U*_ФД)."}), 400

        # Сортировка по I_pn и сборка результатов
        def build_graph(ipn, u, us, tau, title):
            if len(ipn) < 2:
                return {"error": f"Недостаточно точек для «{title}» (нужно минимум 2)."}
            order = np.argsort(ipn)
            ipn_s = [ipn[i] for i in order]
            u_s = [u[i] for i in order]
            us_s = [us[i] for i in order]
            tau_s = [tau[i] for i in order]
            lr = lab3_linear_range(ipn_s, tau_s, tol_rel)
            idx = lr[0]
            fit = lab3_fit(ipn_s, u_s, idx)
            img = make_detector_plot(ipn_s, u_s, tau_s, fit, lr, title)
            xlsx = make_detector_xlsx(ipn_s, u_s, us_s, tau_s, fit, lr, idx, title)
            xf = np.asarray(ipn_s, dtype=float)[np.asarray(idx, dtype=int)]
            yf = np.asarray(u_s, dtype=float)[np.asarray(idx, dtype=int)]
            n_ = len(xf)
            Sx, Sy = float(xf.sum()), float(yf.sum())
            Sxy, Sxx = float((xf * yf).sum()), float((xf * xf).sum())
            sse = float(((yf - fit["a"] * xf - fit["b"]) ** 2).sum())
            sst = float(((yf - yf.mean()) ** 2).sum())
            tau_first = [tau_s[i] for i in idx[:min(3, len(idx))]]
            return {"fit": fit, "lr": lr, "n": len(ipn_s),
                    "sums": (n_, Sx, Sy, Sxy, Sxx, sse, sst),
                    "tau_first": tau_first,
                    "img": img, "xlsx": xlsx}

        configs = []
        results = {}
        if g1_ipn:
            title1 = f"Фотогальв. режим, R₁ = {r1:g} Ом"
            r1r = build_graph(g1_ipn, g1_u, g1_us, g1_tau, title1)
            results["g1"] = r1r
            configs.append(("g1", g1_ipn, g1_u, g1_tau, title1))
        if g2_ipn:
            title2 = f"Фотогальв. режим, R₂ = {r2:g} Ом"
            r2r = build_graph(g2_ipn, g2_u, g2_us, g2_tau, title2)
            results["g2"] = r2r
            configs.append(("g2", g2_ipn, g2_u, g2_tau, title2))
        if g3_ipn:
            sm_txt = f", U_см = {u_sm:g} В" if u_sm is not None else ""
            title3 = f"Фотодиодный режим{sm_txt}"
            r3r = build_graph(g3_ipn, g3_u, g3_us, g3_tau, title3)
            results["g3"] = r3r
            configs.append(("g3", g3_ipn, g3_u, g3_tau, title3))

        # МНК-подстановки по каждому режиму
        for key, lbl in (("g1", "R<sub>1</sub>"), ("g2", "R<sub>2</sub>"), ("g3", "фотодиодный")):
            rr = results.get(key)
            if not rr or "error" in rr or "sums" not in rr:
                continue
            n_, Sx, Sy, Sxy, Sxx, sse, sst = rr["sums"]
            a, b = rr["fit"]["a"], rr["fit"]["b"]
            lr = rr["lr"]
            wk_rows.append((f"a ({lbl}) = (n·Σ(I<sub>pn</sub>·U<sub>ФД</sub>) − ΣI<sub>pn</sub>·ΣU<sub>ФД</sub>) / (n·ΣI<sub>pn</sub>² − (ΣI<sub>pn</sub>)²)",
                            f"({n_}·{fmt_g(Sxy)} − {fmt_g(Sx)}·{fmt_g(Sy)}) / ({n_}·{fmt_g(Sxx)} − ({fmt_g(Sx)})<sup>2</sup>)",
                            f"{fmt_n(a)} мВ/мА"))
            wk_rows.append((f"b ({lbl}) = (ΣU<sub>ФД</sub> − a·ΣI<sub>pn</sub>) / n",
                            f"({fmt_g(Sy)} − {fmt_n(a)}·{fmt_g(Sx)}) / {n_}",
                            f"{fmt_n(b)} мВ"))
            wk_rows.append((f"R² ({lbl}) = 1 − Σ(y − ŷ)² / Σ(y − ȳ)²",
                            f"1 − {fmt_g(sse)} / {fmt_g(sst)}",
                            fmt_n(rr["fit"]["r2"])))
            med = "; ".join(fmt_n(t) for t in rr["tau_first"])
            wk_rows.append((f"τ<sub>оп</sub> ({lbl}) = медиана первых min(3, n) значений τ",
                            f"медиана({med})",
                            fmt_n(lr[3])))
            wk_rows.append((f"Линейный диапазон I<sub>pn</sub> ({lbl})",
                            f"критерий: |τ − τ<sub>оп</sub>| ≤ {tol_pct:g}% τ<sub>оп</sub>",
                            f"{fmt_n(lr[1])} … {fmt_n(lr[2])} мА"))
        if not wk_rows:
            wk_rows = [("—", "нет заполненных данных", "—")]

        # анализ + выводы
        analysis_html = build_lab3_analysis(results, configs, r1, r2, u_sm if u_sm is not None else 0, tol_pct)
        work = work_block(wk_rows, note="МНК-прямая U<sub>ФД</sub> = a·I<sub>pn</sub> + b по точкам "
                          "линейного диапазона; τ<sub>оп</sub> — медиана τ по первым min(3, n) точкам; "
                          "линейный диапазон — участок, где τ отличается от τ<sub>оп</sub> не более чем на "
                          f"{tol_pct:g}%.")

        resp = {"table1": t1_out, "table2": t2_out, "analysis_html": analysis_html,
                "work_html": work}
        for key in ("g1", "g2", "g3"):
            r = results.get(key)
            if r and "img" in r:
                resp[key + "_image"] = r["img"]
                resp[key + "_xlsx"] = r["xlsx"]
            elif r and "error" in r:
                resp[key + "_error"] = r["error"]
        return jsonify(resp)
    except Exception as e:
        return jsonify({"error": f"Ошибка: {e}"}), 400
