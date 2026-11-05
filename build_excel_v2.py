# -*- coding: utf-8 -*-
"""
build_excel_v2.py — сборка snow_melt_v2.xlsx
Версия 2: расчёт формируется ПОСТЕПЕННО.
Сейчас: шапка + 13 исходных данных.
Колонки: B № | C имя | D символ | E значение | F ед. | G щель | H примечание.
(Пустые D/E/G старой раскладки v1 убраны, всё смещено влево.)
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Расчёт"
ws_k = wb.create_sheet("Константы")   # голубые константы — отдельный лист

# ---------------- стили (те же, что в v1) ----------------
F_TITLE = Font(name="Calibri", size=14, bold=True)
F_SECTION = Font(name="Calibri", size=11, bold=True, color="1F3864")
F_BOLD = Font(name="Calibri", size=11, bold=True)
F_NOTE = Font(name="Calibri", size=9, color="808080")

FILL_IN = PatternFill("solid", fgColor="FFF2CC")    # вводимые исходные данные
FILL_CONST = PatternFill("solid", fgColor="DDEBF7")  # физические/конструктивные константы
FILL_RES = PatternFill("solid", fgColor="E2EFDA")   # расчётные результаты (формулы)
FILL_SEC = PatternFill("solid", fgColor="D9D9D9")

THIN = Side(style="thin", color="BFBFBF")
B_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HAIR = Border(bottom=Side(style="hair", color="D0D0D0"))

AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
AL_R = Alignment(horizontal="right", vertical="center")

NUM1 = "0.0"
NUM2 = "0.00"
NUM3 = "0.000"

def put(row, col, value, font=None, fill=None, num=None, align=AL_L, border=False):
    c = ws.cell(row=row, column=col)
    c.value = value
    if font: c.font = font
    if fill: c.fill = fill
    if num: c.number_format = num
    c.alignment = align
    if border: c.border = B_ALL
    return c

def param(sheet, row, no, name, sym, value, unit, note, fill=None, num=NUM1):
    """№ | имя | символ | значение | ед. | примечание.
    fill=FILL_IN (по умолчанию) — вводимая величина (жёлтая);
    fill=FILL_CONST — константа (голубая)."""
    sheet.cell(row=row, column=2, value=no)
    sheet.cell(row=row, column=3, value=name)
    sheet.cell(row=row, column=4, value=sym)
    v = sheet.cell(row=row, column=5, value=value)
    sheet.cell(row=row, column=6, value=unit)
    sheet.cell(row=row, column=8, value=note)
    v.font = F_BOLD
    v.fill = fill if fill is not None else FILL_IN
    if num: v.number_format = num
    v.alignment = AL_R
    v.border = B_ALL
    for col in (2, 3, 4, 6, 8):
        sheet.cell(row=row, column=col).border = B_ALL

def section(sheet, row, text):
    for col in range(2, 9):
        c = sheet.cell(row=row, column=col, value=text if col == 2 else None)
        c.fill = FILL_SEC
        c.font = F_SECTION
    sheet.cell(row=row, column=2).value = text

# ---------------- ширины (компактная раскладка, без пустых D/E/G) ----------------
widths = {1: 2, 2: 8, 3: 46, 4: 8, 5: 12, 6: 10, 7: 2, 8: 62}
for shard in (ws, ws_k):
    for col, w in widths.items():
        shard.column_dimensions[get_column_letter(col)].width = w

# ---------------- заголовок ----------------
put(1, 2, "РАСЧЁТ СИСТЕМЫ СНЕГОТАЯНИЯ — v2", font=F_TITLE)
put(2, 2, "Версия 2.0 — формируется постепенно. Исходные данные добавляются по одной.",
    font=F_NOTE)
for r in (1, 2):
    ws.cell(row=r, column=2).alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
ws.row_dimensions[1].height = 20
ws.row_dimensions[2].height = 16

ws.freeze_panes = "A4"

# ================= 1. ИСХОДНЫЕ ДАННЫЕ (лист «Расчёт») =================
section(ws, 4, "1. ИСХОДНЫЕ ДАННЫЕ (ввод)")

param(ws, 5, 1, "Площадь обогреваемой площадки", "Fст", 50.0, "м²",
      "первый исходный параметр v2; v1: 14 м²")
param(ws, 6, 2, "Время стаивания", "τ", 120.0, "мин",
      "120 = режим на пределе; 300 = комфортный (trade-off, plan.md 8.9); вводится в минутах")
param(ws, 7, 3, "Температура наружного воздуха", "tв", -10.0, "°C",
      "температура воздуха при событии снегопада (СП подбирается по климатологии района; v1: −10 °C)")
param(ws, 8, 4, "Температура под обогреваемой конструкцией", "tн", 5.0, "°C",
      "диапазон грунта/подстилающих слоёв под песчаной подушкой; участвует в КПД (поток вниз)")
param(ws, 9, 5, "Температура поверхности площадки", "tст", 2.0, "°C",
      "начальная температура поверхности (стяжки) перед событием; v1: 2 °C")
param(ws, 10, 6, "Слой снега", "hсн", 2.0, "мм",
      "высота снежного покрова к моменту начала таяния; v1: 2 мм")
param(ws, 11, 7, "Скорость ветра", "vн", 3.0, "м/с",
      "вводится в αн = 8,7 + 2,6·vн; главный множитель конвективных потерь (plan.md 8.11)")
param(ws, 12, 8, "Шаг укладки трубы", "b", 0.20, "м",
      "расстояние между трубами петли; исходное = 200 мм, как в v1")
param(ws, 13, 9, "Температура подачи", "tпод", 55.0, "°C",
      "у подающей трубы; средняя tср = (tпод + tобр)/2 — свойства ПГ смотрим на листе «ПГ»; ПГ 40 %: tкрист ≈ −24 °C, к 55 °C запас есть")
param(ws, 14, 10, "Внутренний диаметр трубы", "dвн", 16.0, "мм",
      "PE-Xa 20×2 (стенка 2 мм) → dвн 16 мм; v1: то же; в формулах переводится в м (×10⁻³)", None, NUM1)

# ================= 2. РАСЧЁТ =================
section(ws, 16, "2. РАСЧЁТ")

# ссылки на ввод: hсн=E10 (мм), tв=E7, tст=E9; константы: ρл, cл, cв, λл — лист «Константы»
param(ws, 18, 1, "Нагрев слоя льда от tв до 0 °C", "Qсн",
      "=E10/1000*'Константы'!E5*'Константы'!E6*(0-E7)", "кДж/м²",
      "hсн·ρл·cл·(0−tв) = (hсн/1000, м)·900·2,1·(0−tв)",
      FILL_RES, NUM2)
param(ws, 19, 2, "Таяние слоя льда", "Qт",
      "=E10/1000*'Константы'!E5*'Константы'!E11", "кДж/м²",
      "hсн·ρл·λл — скрытая теплота плавления",
      FILL_RES, NUM2)
param(ws, 20, 3, "Нагрев талой воды от 0 до tст", "Qв",
      "=E10/1000*'Константы'!E5*'Константы'!E7*E9", "кДж/м²",
      "hсн·ρл·cв·(tст−0) — масса талой воды = масса льда: hсн·ρл/ρв = 2·900/1000 = 1,8 мм слоя воды той же массы",
      FILL_RES, NUM2)
param(ws, 21, 4, "Потери тепла излучением", "Qизл",
      "='Константы'!E8*'Константы'!E9*(POWER((0+273)/100,4)-POWER((E7+273)/100,4))*E6*60/1000", "кДж/м²",
      "ε·Co·[(273/100)⁴−((tв+273)/100)⁴]·τ·60/1000 — ЧИСТЫЙ обмен: встречное излучение окружения при tв; при tв=0 препятствий нет",
      FILL_RES, NUM2)
param(ws, 22, 5, "Коэффициент теплоотдачи конвекцией", "αн",
      "=8.7+2.6*E11", "Вт/(м²·К)",
      "αн = 8,7 + 2,6·vн — линейная зависимость от ветра; при vн=3 → 16,5; формула поверена (plan.md 8.11: ниже Nusselt–Jürges и McAdams)",
      FILL_RES, NUM2)
param(ws, 23, 6, "Потери тепла конвекцией", "Qкон",
      "=E22*(0-E7)*E6*60/1000", "кДж/м²",
      "αн·(0−tв)·τ·60/1000 — от поверхности 0 °C к воздуху tв, за весь период таивания",
      FILL_RES, NUM2)
# ================= ЛИСТ «КОНСТАНТЫ» =================
# заголовок листа
ws_k.cell(row=1, column=2, value="КОНСТАНТЫ — физические и справочные величины").font = F_TITLE
ws_k.cell(row=2, column=2, value="Значения не изменяются при каждом расчёте; редактируются осознанно.").font = F_NOTE
ws_k.row_dimensions[1].height = 20

section(ws_k, 4, "1. ФИЗИЧЕСКИЕ КОНСТАНТЫ")

param(ws_k, 5, 1, "Плотность льда", "ρл", 900.0, "кг/м³",
      "для расчёта нагрузки таяния в кДж/м²", FILL_CONST, NUM1)
param(ws_k, 6, 2, "Удельная теплоёмкость льда", "cл", 2.1, "кДж/(кг·К)",
      "нагрев льда от tв до 0 °C", FILL_CONST)
param(ws_k, 7, 3, "Удельная теплоёмкость воды", "cв", 4.2, "кДж/(кг·К)",
      "нагрев талой воды от 0 до tст", FILL_CONST)
param(ws_k, 8, 4, "Степень черноты поверхности", "ε", 0.90, "—",
      "снег/лёд/мокрое покрытие (0,85–0,98); 0,055 из исходника — фольга, не подходит тающей поверхности",
      FILL_CONST, "0.00")
param(ws_k, 9, 5, "Коэффициент излучения абс. чёрного тела", "Co", 5.67, "Вт/(м²·K⁴)",
      "Стефан–Больцман, современное значение; в v1 5,77 — старое учебное (разница 1,8 %)", FILL_CONST, NUM2)
param(ws_k, 10, 6, "Удельная теплоёмкость теплоносителя (ПГ 40 %, 30 °C)", "cж", 3.644, "кДж/(кг·К)",
      "оценка cp = 3,56 + 0,0028·T (точность ~1 %, hydr_calc); на «ПГ» зависимость от −20 до 80 °C меняет cp всего на ~8 %", FILL_CONST, NUM3)
param(ws_k, 11, 7, "Удельная теплота плавления льда", "λл", 330.0, "кДж/кг",
      "скрытая теплота фазового перехода лёд→вода при 0 °C", FILL_CONST, NUM1)
param(ws_k, 12, 8, "Плотность воды", "ρв", 1000.0, "кг/м³",
      "для перевода массы талой воды в объём (глубина слоя воды = масса/ρв); в тепловых строках масса считается по льду — см. примечание к Qв", FILL_CONST, NUM1)

# ================= ЛИСТ «ПГ» — теплоноситель пропиленгликоль 40 % =================
# Данные KasperCalc/Dow (как в hydr_calc.xlsx, D:\opencode\excel).
# ν в мм²/с (cSt); μ — динамическая вязкость, мПа·с = ρ·ν/1000.
# Таблица от −20 °C: точка кристаллизации ПГ 40 % ≈ −24 °C, ниже — не рабочая зона.
PG40 = [
    (-20, 1048.0, 46.66),
    (-15, 1046.3, 31.60),
    (-10, 1044.6, 22.12),
    (-5,  1042.8, 15.94),
    (0,   1041.0, 11.81),
    (5,   1039.1,  8.968),
    (10,  1037.1,  6.954),
    (15,  1035.0,  5.507),
    (20,  1032.7,  4.444),
    (25,  1030.2,  3.640),
    (30,  1027.5,  3.035),
    (35,  1024.6,  2.557),
    (40,  1021.5,  2.192),
    (45,  1018.3,  1.927),
    (50,  1015.0,  1.655),
    (55,  1011.6,  1.463),
    (60,  1008.2,  1.299),
    (65,  1004.7,  1.174),
    (70,  1001.2,  1.059),
    (75,   997.63, 0.9623),
    (80,   994.0,  0.8853),
]

ws_pg = wb.create_sheet("ПГ")
# заголовок листа
ws_pg.cell(row=1, column=2, value="ТЕПЛОНОСИТЕЛЬ — пропиленгликоль 40 %").font = F_TITLE
ws_pg.cell(row=2, column=2, value="Свойства раствора (40 % масс.) в зависимости от температуры; всегда один и тот же теплоноситель.").font = F_NOTE
ws_pg.cell(row=3, column=2, value="Источник: KasperCalc/Dow (как в hydr_calc.xlsx). Точка кристаллизации ≈ −24 °C — таблица начинается с −20 °C.").font = F_NOTE
ws_pg.row_dimensions[1].height = 20

for col, head in ((2, "T, °C"), (3, "ρ, кг/м³"), (4, "ν, мм²/с"), (5, "μ, мПа·с")):
    c = ws_pg.cell(row=5, column=col, value=head)
    c.fill = FILL_SEC
    c.font = F_BOLD
    c.alignment = AL_C
    c.border = B_ALL

r = 6
for T, rho, nu in PG40:
    a = ws_pg.cell(row=r, column=2, value=T)
    b = ws_pg.cell(row=r, column=3, value=rho)
    d = ws_pg.cell(row=r, column=4, value=nu)
    e = ws_pg.cell(row=r, column=5, value=round(rho * nu / 1000.0, 3))
    a.number_format = "0"
    b.number_format = "0.0"
    d.number_format = "0.000"
    e.number_format = "0.000"
    for cell in (a, b, d, e):
        cell.border = B_ALL
        cell.alignment = AL_R
    r += 1

pg_widths = {1: 2, 2: 10, 3: 12, 4: 12, 5: 12, 6: 2, 7: 52}
for col, w in pg_widths.items():
    ws_pg.column_dimensions[get_column_letter(col)].width = w

# примечание под таблицей
note = ws_pg.cell(row=r + 1, column=2,
    value="ν — кинематическая вязкость (мм²/с = сСт); μ = ρ·ν/1000 — динамическая (мПа·с).")
note.font = F_NOTE
note2 = ws_pg.cell(row=r + 2, column=2,
    value="Вязкость ПГ кратно выше воды: при холодном старте режим течения близок к ламинарному (Re ↓, учтём в гидравлике).")
note2.font = F_NOTE

wb.save("snow_melt_v2.xlsx")
print("OK: snow_melt_v2.xlsx saved")
