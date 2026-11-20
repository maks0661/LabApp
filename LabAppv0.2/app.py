# -*- coding: utf-8 -*-
"""
Лабораторный практикум — локальное веб-приложение по нескольким курсам.

app.py — общий файл: ядро Flask, главная страница со списком курсов
(выпадающие секции-аккордеоны) и запуск сервера. Содержание курсов
вынесено в отдельные модули:

    kioe.py — «Квантовая и оптическая электроника» (лабораторные работы № 1-3);
    vipe.py — «Вакуумная и плазменная электроника» (лабораторные работы № 1-3);
    common.py — общие утилиты (шрифты, кодирование графиков и Excel, аппроксимация).

Запуск:
    pip install -r requirements.txt
    python app.py
Затем открыть в браузере:  http://127.0.0.1:5000
(браузер откроется автоматически при запуске).
"""

import threading
import webbrowser

from flask import Flask

from kioe import bp as kioe_bp
from vipe import bp as vipe_bp

app = Flask(__name__)
app.register_blueprint(kioe_bp)
app.register_blueprint(vipe_bp)


# ===========================================================================
#  Главная страница
# ===========================================================================

INDEX_HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Лабораторный практикум</title>
<style>
  :root { --accent:#1f4e79; --bg:#f4f6f9; --card:#fff; --border:#d8dee9; }
  * { box-sizing: border-box; }
  body { margin:0; font-family:'Segoe UI', system-ui, Arial, sans-serif; background:var(--bg); color:#222; }
  header { background:var(--accent); color:#fff; padding:30px 20px; text-align:center; }
  header h1 { margin:0; font-size:26px; }
  header p { margin:8px 0 0; opacity:.85; font-size:14px; }
  main { max-width:1040px; margin:32px auto; padding:0 16px; }
  .intro { color:#555; font-size:14.5px; margin:0 0 20px; }
  details.course { background:var(--card); border:1px solid var(--border); border-radius:10px; margin:0 0 18px; overflow:hidden; }
  details.course summary { cursor:pointer; list-style:none; padding:18px 22px; background:#eef2f7;
          border-left:5px solid var(--accent); display:flex; align-items:baseline; gap:10px; flex-wrap:wrap; user-select:none; }
  details.course summary::-webkit-details-marker { display:none; }
  details.course summary:hover { background:#e6ecf5; }
  details.course summary .cname { font-size:18px; font-weight:700; color:var(--accent); }
  details.course summary .csub { font-size:13.5px; color:#777; }
  details.course summary .arrow { margin-left:auto; color:var(--accent); font-size:14px; transition:transform .15s; }
  details.course[open] summary .arrow { transform:rotate(90deg); }
  details.course .body { padding:20px 22px 24px; }
  .grid { display:grid; grid-template-columns:repeat(auto-fill, minmax(300px,1fr)); gap:20px; }
  .card { background:var(--card); border:1px solid var(--border); border-radius:10px;
          padding:22px; text-decoration:none; color:inherit; display:flex; flex-direction:column;
          transition:box-shadow .15s, transform .15s; }
  .card:hover { box-shadow:0 6px 18px rgba(0,0,0,.10); transform:translateY(-3px); }
  .card .num { font-size:12.5px; color:var(--accent); font-weight:600; letter-spacing:.5px; text-transform:uppercase; }
  .card h2 { font-size:17px; margin:10px 0 10px; line-height:1.35; }
  .card p { font-size:13.5px; color:#555; margin:0 0 16px; flex:1; line-height:1.5; }
  .card .badge { align-self:flex-start; font-size:12px; padding:3px 11px; border-radius:12px; background:#e3f0e3; color:#2e7d32; }
  .card.stub .badge { background:#f0f0f0; color:#888; }
  .card.stub h2, .card.stub p { color:#999; }
  .card .open { margin-top:10px; font-size:13px; color:var(--accent); font-weight:600; }
  footer { text-align:center; color:#888; font-size:12px; padding:24px; }
</style>
</head>
<body>
<header>
  <h1>Лабораторный практикум</h1>
  <p>Локальное приложение для ввода данных эксперимента, автоматического расчёта и построения графиков</p>
</header>
<main>
  <p class="intro">Выберите курс, затем лабораторную работу. Работа откроется в новой вкладке браузера.</p>

  <details class="course" open>
    <summary>
      <span class="cname">Квантовая и оптическая электроника</span>
      <span class="csub">&middot; ввод данных, расчёт и построение графиков</span>
      <span class="arrow">&#9654;</span>
    </summary>
    <div class="body">
      <div class="grid">
        <a class="card" href="/kioe/lab/1" target="_blank">
          <span class="num">Лабораторная работа № 1</span>
          <h2>Исследование основных параметров полупроводникового лазера</h2>
          <p>Измерение зависимости интенсивности излучения и степени поляризации от тока через p&#8209;n переход. Расчёт I<sub>pn</sub>, порогового тока I<sub>пор</sub>, коэффициентов k<sub>св</sub> и k<sub>лаз</sub>, построение графиков.</p>
          <span class="badge">Доступна</span>
          <span class="open">Открыть &rarr;</span>
        </a>
        <a class="card" href="/kioe/lab/2" target="_blank">
          <span class="num">Лабораторная работа № 2</span>
          <h2>Определение электрооптических параметров анизотропных кристаллов</h2>
          <p>Измерение зависимости интенсивности прошедшего света от напряжения на кристалле LiNbO<sub>3</sub>. Расчёт полуволнового напряжения U<sub>λ/2</sub>, линеаризация arcsin(I/I<sub>0</sub>), определение электрооптического коэффициента r<sub>33</sub>.</p>
          <span class="badge">Доступна</span>
          <span class="open">Открыть &rarr;</span>
        </a>
        <a class="card" href="/kioe/lab/3" target="_blank">
          <span class="num">Лабораторная работа № 3</span>
          <h2>Полупроводниковые детекторы оптического излучения</h2>
          <p>Исследование фотодиода (ФД-24К) в фотогальваническом и фотодиодном режимах. Расчёт I<sub>pn</sub>, пропускания τ светофильтра, определение диапазона линейности и чувствительности по трём графикам U<sub>ФД</sub>(I<sub>pn</sub>) и τ(I<sub>pn</sub>).</p>
          <span class="badge">Доступна</span>
          <span class="open">Открыть &rarr;</span>
        </a>
      </div>
    </div>
  </details>

  <details class="course">
    <summary>
      <span class="cname">Вакуумная и плазменная электроника</span>
      <span class="csub">&middot; термоэлектронная эмиссия и электровакуумные приборы</span>
      <span class="arrow">&#9654;</span>
    </summary>
    <div class="body">
      <div class="grid">
        <a class="card" href="/vipe/lab/1" target="_blank">
          <span class="num">Лабораторная работа № 1</span>
          <h2>Исследование термоэлектронной эмиссии</h2>
          <p>Семейство вольтамперных характеристик вакуумного диода при разных токах накала. Определение токов эмиссии и плотности тока, температура катода по удельной мощности накала, прямая Ричардсона (константы A и e&phi;), прямые Шоттки, зависимость j&nbsp;=&nbsp;f(T).</p>
          <span class="badge">Доступна</span>
          <span class="open">Открыть &rarr;</span>
        </a>
        <a class="card" href="/vipe/lab/2" target="_blank">
          <span class="num">Лабораторная работа № 2</span>
          <h2>Исследование электронно-лучевой трубки с магнитным управлением</h2>
          <p>Модуляционные характеристики при трёх U<sub>а</sub> (крутизна S, проницаемость модулятора Δ), ток фокусирующей катушки и фокусное расстояние магнитной линзы, отклонение пятна от тока катушек (чувствительность, мм/ампервиток) и от U<sub>а</sub>, яркость экрана.</p>
          <span class="badge">Доступна</span>
          <span class="open">Открыть &rarr;</span>
        </a>
        <a class="card stub" href="/vipe/lab/3" target="_blank">
          <span class="num">Лабораторная работа № 3</span>
          <h2>Исследование фотоэлектронных приборов</h2>
          <p>Раздел находится в разработке. Таблицы и графики будут добавлены позже.</p>
          <span class="badge">Заглушка</span>
          <span class="open">Открыть &rarr;</span>
        </a>
      </div>
    </div>
  </details>
</main>
<footer>Лабораторный практикум &middot; локальное приложение (Flask + matplotlib)</footer>
</body>
</html>
"""


@app.route("/")
def index():
    return INDEX_HTML


# ===========================================================================
#  Запуск
# ===========================================================================

if __name__ == "__main__":
    # Автоматически открыть браузер через 1.2 с после старта сервера.
    threading.Timer(1.2, lambda: webbrowser.open("http://127.0.0.1:5000/")).start()
    print("Сервер запущен: http://127.0.0.1:5000  (Ctrl+C для остановки)")
    app.run(host="127.0.0.1", port=5000, debug=False)
