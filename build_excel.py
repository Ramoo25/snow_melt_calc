# -*- coding: utf-8 -*-
"""
build_excel.py — сборка snow_melt_new.xlsx
Система снеготаяния: тепловой расчёт, петли, рабочая точка насоса.
Структура листа «Расчёт» зафиксирована в plan.md (п. 5).
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Расчёт"

# ---------------- стили ----------------
F_TITLE = Font(name="Calibri", size=14, bold=True)
F_SECTION = Font(name="Calibri", size=11, bold=True, color="1F3864")
F_BOLD = Font(name="Calibri", size=11, bold=True)
F_NOTE = Font(name="Calibri", size=9, color="808080")
F_RES = Font(name="Calibri", size=11, bold=True, color="006100")

FILL_IN = PatternFill("solid", fgColor="FFF2CC")   # исходные данные (ввод)
FILL_CONST = PatternFill("solid", fgColor="DDEBF7")  # константы
FILL_RES = PatternFill("solid", fgColor="E2EFDA")  # результаты
FILL_SEC = PatternFill("solid", fgColor="D9D9D9")

THIN = Side(style="thin", color="BFBFBF")
B_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

AL_C = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL_L = Alignment(horizontal="left", vertical="center", wrap_text=True)
AL_R = Alignment(horizontal="right", vertical="center")

NUM3 = "0.000"
NUM2 = "0.00"
NUM1 = "0.0"

def put(row, col, value, font=None, fill=None, num=None, align=AL_L, border=False):
    c = ws.cell(row=row, column=col)
    c.value = value
    if font: c.font = font
    if fill: c.fill = fill
    if num: c.number_format = num
    c.alignment = align
    if border: c.border = B_ALL
    return c

def param(row, no, name, sym, value, unit, note, fill=None, num=None, formula=False):
    """Строка парamparameter: № | имя | (D,E) | симв. | (G) | знач. | (I) | ед. | (K) | прим.»"""
    ws.cell(row=row, column=2, value=no)          # B
    ws.cell(row=row, column=3, value=name)        # C
    ws.cell(row=row, column=6, value=sym)         # F
    v = ws.cell(row=row, column=8, value=value)   # H
    ws.cell(row=row, column=10, value=unit)       # J
    ws.cell(row=row, column=12, value=note)       # L
    v.font = F_BOLD
    if fill: v.fill = fill
    if num: v.number_format = num
    v.alignment = AL_R if not formula else AL_C
    v.border = B_ALL
    for col in (2, 3, 6, 8, 10, 12):
        if col != 8:
            ws.cell(row=row, column=col).border = B_ALL if fill else Border(bottom=Side(style="hair", color="D0D0D0"))

def section(row, text):
    for col in range(2, 13):
        c = ws.cell(row=row, column=col, value=text if col == 2 else None)
        c.fill = FILL_SEC
        c.font = F_SECTION
    ws.cell(row=row, column=2).value = text

# ---------------- ширины ----------------
widths = {1: 2, 2: 4, 3: 46, 4: 3, 5: 3, 6: 8, 7: 2, 8: 12, 9: 2, 10: 10, 11: 2, 12: 62}
for col, w in widths.items():
    ws.column_dimensions[get_column_letter(col)].width = w

# ---------------- заголовок ----------------
put(2, 2, "РАСЧЁТ СИСТЕМЫ СНЕГОТАЯНИЯ", font=F_TITLE)
put(3, 2, "snow_melt_new.xlsx — рабочая точка насоса и средняя длина петли", font=F_BOLD)

# ================= 1. ИСХОДНЫЕ ДАННЫЕ =================
section(5, "1. ИСХОДНЫЕ ДАННЫЕ")
put(6, 3, "Параметр", font=F_BOLD)
put(6, 6, "Обозн.", font=F_BOLD)
put(6, 8, "Значение", font=F_BOLD, align=AL_C)
put(6, 10, "Ед.", font=F_BOLD)
put(6, 12, "Примечание", font=F_BOLD)

param(7, 1, "Температура наружного воздуха", "tв", -10, "°C", "расчётная наружная температура", FILL_IN, NUM1)
param(8, 2, "Температура под обогреваемой конструкцией", "tн", 5, "°C", "", FILL_IN, NUM1)
param(9, 3, "Температура поверхности площадки", "tст", 2, "°C", "требуемая температура поверхности", FILL_IN, NUM1)
param(10, 4, "Слой снега/льда", "hсн", 2, "мм", "в плотном теле", FILL_IN, NUM1)
param(11, 5, "Температура антифриза в подающем коллекторе", "tпод", 55, "°C", "система коллекторная: подача задана напрямую", FILL_IN, NUM1)
param(12, 6, "Скорость ветра", "vн", 3, "м/с", "", FILL_IN, NUM1)
param(13, 7, "Площадь обогреваемой поверхности", "Fст", 14, "м²", "", FILL_IN, NUM1)
param(14, 8, "Время таяния снега", "τ", 60, "мин", "", FILL_IN, NUM1)
param(15, 9, "Модель насоса", "", "Grundfos UPS 25-20", "—", "выбрать из выпадающего списка (лист «Насосы-кривые», 6 моделей)", FILL_IN)

dv = DataValidation(type="list",
                    formula1="='Насосы-кривые'!$B$3:$G$3",
                    allow_blank=False, showDropDown=False)
ws.add_data_validation(dv)
dv.add(ws.cell(row=15, column=8))

# ================= 2. КОНСТАНТЫ =================
section(17, "2. КОНСТАНТЫ")
put(18, 3, "Параметр", font=F_BOLD)
put(18, 6, "Обозн.", font=F_BOLD)
put(18, 8, "Значение", font=F_BOLD, align=AL_C)
put(18, 10, "Ед.", font=F_BOLD)
put(18, 12, "Примечание", font=F_BOLD)

param(19, 1, "Шаг укладки трубы", "b", 0.2, "м", "конструкция", FILL_CONST)
param(20, 2, "Наружный диаметр трубы", "dн", 0.020, "м", "PE-Xa 20×2", FILL_CONST)
param(21, 3, "Толщина стенки трубы", "s", 0.002, "м", "", FILL_CONST)
param(22, 4, "Внутренний диаметр трубы", "dв", "=H20-2*H21", "м", "", FILL_CONST, formula=True)
param(23, 5, "Площадь живого сечения трубы", "ω", "=PI()*H22^2/4", "м²", "", FILL_CONST, formula=True)
param(24, 6, "Теплопроводность трубы (PE-X)", "λтр", 0.41, "Вт/(м·К)", "", FILL_CONST)
param(25, 7, "Стяжка (бетон)", "d", 0.060, "м", "слой НАД трубой", FILL_CONST)
param(26, 8, "Теплопроводность бетона/стяжки", "λ", 1.75, "Вт/(м·K)", "", FILL_CONST)
param(27, 9, "Керамическая плитка (+клей)", "d", 0.020, "м", "слой НАД трубой (клей входит в 20 мм)", FILL_CONST)
param(28, 10, "Теплопроводность плитки", "λ", 3.5, "Вт/(м·K)", "", FILL_CONST)
param(29, 11, "ЭППС (пеноплэкс)", "d", 0.050, "м", "слой ПОД трубой", FILL_CONST)
param(30, 12, "Теплопроводность ЭППС", "λ", 0.034, "Вт/(м·K)", "", FILL_CONST)
param(31, 13, "Бетонное основание", "d", 0.300, "м", "слой ПОД трубой", FILL_CONST)
param(32, 14, "Теплопроводность бетона", "λ", 1.75, "Вт/(м·K)", "", FILL_CONST)
param(33, 15, "Сопротивление уплотнённого песка", "Rпес", 2.1, "м²·К/Вт", "под бетонным основанием (зональная оценка, как в исходнике)", FILL_CONST)
param(34, 16, "Сопротивление слоёв НАД трубой", "Rвверх", "=H25/H26+H27/H28", "м²·К/Вт", "", FILL_CONST, formula=True)
param(35, 17, "Сопротивление слоёв ПОД трубой", "Rвниз", "=H29/H30+H31/H32+H33", "м²·К/Вт", "", FILL_CONST, formula=True)

put(36, 3, "Теплоноситель", font=F_SECTION)
param(37, 1, "Плотность пропиленгликоля 40%", "ρж", 1032, "кг/м³", "при 20 °C (glikoli.ru: 40% = 1,032 г/см³)", FILL_CONST, "0.0")
param(38, 2, "Удельная теплоёмкость антифриза", "cж", 3.64, "кДж/(кг·К)", "ПГ 40%, около 20–50 °C", FILL_CONST)
param(39, 3, "Температура кристаллизации", "", -25, "°C", "ПГ 40% (справочно; у 35% −20, у 45% −30)", FILL_CONST, "0.0")
param(40, 4, "Коэффициент антифриза (рабочий)", "kА", 1.39, "—", "как в исходнике (режим DIXIS); старт-коэффициенты не используются", FILL_CONST)

put(41, 3, "Материалы / излучение", font=F_SECTION)
param(42, 1, "Плотность льда", "ρл", 900, "кг/м³", "", FILL_CONST, "0.0")
param(43, 2, "Удельная теплоёмкость льда", "cл", 2.1, "кДж/(кг·К)", "", FILL_CONST)
param(44, 3, "Удельная теплота таяния льда", "λл", 330, "кДж/кг", "", FILL_CONST, "0.0")
param(45, 4, "Удельная теплоёмкость воды", "cв", 4.2, "кДж/(кг·К)", "", FILL_CONST)
param(46, 5, "Степень черноты", "ε", 0.055, "—", "как в исходнике (см. лист «Справка»)", FILL_CONST)
param(47, 6, "Коэффициент излучения чёрного тела", "Co", 5.77, "Вт/(м²·K⁴)", "", FILL_CONST)
param(48, 7, "Коэффициент теплоотдачи на внутренней поверхности трубы", "αв", 11.5, "Вт/(м²·К)", "", FILL_CONST)
param(49, 8, "Коэффициент теплоотдачи наружной поверхности", "αн", "=8.7+2.6*H12", "Вт/(м²·К)", "αн = 8,7 + 2,6·vн", FILL_CONST, formula=True)

put(50, 3, "Гидравлика — местные сопротивления и магистраль (СЕЙЧАС НЕ УЧАСТВУЮТ в расчёте, зарезервировано)", font=F_SECTION)
param(51, 1, "Кран шаровой (kvs)", "", 6, "—", "решение: kvs = 6", FILL_CONST, "0.0")
param(52, 2, "Количество кранов", "", 5, "шт", "", FILL_CONST, "0")
param(53, 3, "Фильтр-сетка (kvs)", "", 3.84, "—", "", FILL_CONST)
param(54, 4, "Количество фильтров", "", 1, "шт", "", FILL_CONST, "0")
param(55, 5, "Резкий поворот (ζ)", "", 1, "—", "", FILL_CONST, "0.0")
param(56, 6, "Количество поворотов", "", 6, "шт", "", FILL_CONST, "0")
param(57, 7, "Трёхходовик (kvs)", "", 4.5, "—", "", FILL_CONST)
param(58, 8, "Количество трёхходовиков", "", 1, "шт", "", FILL_CONST, "0")
param(59, 9, "Внутренний диаметр магистрали", "", 23.2, "мм", "не используется (магистраль пренебрежимо мала)", FILL_CONST)
param(60, 10, "Длина магистрали (подача + обратка)", "", 20, "м", "не используется", FILL_CONST, "0.0")
param(61, 11, "Коэффициент запаса (коллекторы и пр.)", "", 1.2, "—", "в текущей упрощённой схеме не используется (подключение напрямую в коллектор)", FILL_CONST)
param(62, 12, "Проектная скорость в петле", "vпр", 0.45, "м/с", "параметр компоновки (норма 0,2–0,5 м/с)", FILL_CONST)

# ================= 3. ТЕПЛОВОЙ БАЛАНС (на 1 м² поверхности) =================
section(64, "3. ТЕПЛОВОЙ БАЛАНС — на 1 м² обогреваемой поверхности (шаги 1–5 исходника)")
put(65, 3, "Составляющая", font=F_BOLD)
put(65, 8, "Значение", font=F_BOLD, align=AL_C)
put(65, 10, "Ед.", font=F_BOLD)
put(65, 12, "Формула", font=F_BOLD)

param(66, 1, "Нагрев льда от tв до 0 °C", "Q1", "=H10/1000*H42*H43*(0-H7)", "кДж/м²", "h·ρл·сл·(0−tв)", None, NUM2, formula=True)
param(67, 2, "Таяние льда", "Q2", "=H10/1000*H42*H44", "кДж/м²", "h·ρл·λл", None, NUM2, formula=True)
param(68, 3, "Нагрев талой воды от 0 до tст", "Q3", "=H10/1000*H42*H45*(H9-0)", "кДж/м²", "h·ρл·св·tст", None, NUM2, formula=True)
param(69, 4, "Потери тепла излучением", "Qизл", "=3600*H46*H47*POWER((H9+273)/100,4)/1000", "кДж/м²", "3600·ε·Co·((tст+273)/100)⁴/1000 (модель исходника, ε=0,055)", None, NUM2, formula=True)
param(70, 5, "Коэффициент конвекции", "αк", "=2.16*POWER(0-H7,1/3)", "Вт/(м²·К)", "2,16·(0−tв)^(1/3)", None, NUM2, formula=True)
param(71, 6, "Потери тепла конвекцией", "Qкон", "=1800*H70*(0-H7)/1000", "кДж/м²", "1800·αк·(0−tв)/1000 (коэф. 0,5 — как в исходнике)", None, NUM2, formula=True)
param(72, 7, "Суммарные затраты тепла", "ΣQ", "=H66+H67+H68+H69+H71", "кДж/м²", "Q1+Q2+Q3+Qизл+Qкон", None, NUM2, formula=True)

# ================= 4. ПРИВЕДЕННОЕ СОПРОТИВЛЕНИЕ =================
section(74, "4. ПРИВЕДЕННОЕ СОПРОТИВЛЕНИЕ «от теплоносителя до поверхности» (шаг 11)")
param(75, 1, "Приведенное сопротивление", "Rв", "=1/H48+H34+1/H49+H21/H24", "м²·К/Вт", "1/αв + Rвверх + 1/αн + s/λтр", None, "0.0000", formula=True)

# ================= 5. РАСЧЁТ ПРОХОДАМИ (tср вычисляется) =================
section(77, "5. РАСЧЁТ ПРОХОДАМИ — средняя температура теплоносителя вычисляется (решение п. 8.5)")
put(78, 3, "Каждый проход: tср → η → n1 → tмин → новое tср. Проход 1 стартует с грубой оценки.", font=F_NOTE)

def eta_formula(t_ref):
    return "=(({t}-H7)*H35)/((({t}-H7)*H35)+(({t}-H8)*H34))".format(t=t_ref)

# Проход 1: строки 79-83
param(79, "П1", "Средняя температура (оценка)", "tср₁", "=(H11+H9)/2", "°C", "начальная оценка: (tпод + tст)/2", None, NUM2, formula=True)
param(80, "П1", "КПД (доля тепла вверх)", "η₁", eta_formula("H79"), "—", "шаг 6 исходника", None, "0.0000", formula=True)
param(81, "П1", "Удельная мощность", "n1₁", "=H72/(H80*H14*60)", "кВт/м²", "шаг 7: ΣQ/(η·τ), τ в с", None, NUM3, formula=True)
param(82, "П1", "Минимальная температура в петле", "tмин₁", "=1000*H81*H75+H9", "°C", "шаг 11: 1000·n1·Rв + tст", None, NUM2, formula=True)
param(83, "П1", "Новое среднее значение", "tср₂", "=(H11+H82)/2", "°C", "(tпод + tмин)/2", None, NUM2, formula=True)

# Проход 2: строки 84-88
param(84, "П2", "Средняя температура", "tср₂", "=H83", "°C", "", None, NUM2, formula=True)
param(85, "П2", "КПД (доля тепла вверх)", "η₂", eta_formula("H84"), "—", "", None, "0.0000", formula=True)
param(86, "П2", "Удельная мощность", "n1₂", "=H72/(H85*H14*60)", "кВт/м²", "", None, NUM3, formula=True)
param(87, "П2", "Минимальная температура в петле", "tмин₂", "=1000*H86*H75+H9", "°C", "", None, NUM2, formula=True)
param(88, "П2", "Новое среднее значение", "tср₃", "=(H11+H87)/2", "°C", "", None, NUM2, formula=True)

# Проход 3 (окончательный): строки 89-93
param(89, "П3", "Средняя температура", "tср₃", "=H88", "°C", "", None, NUM2, formula=True)
param(90, "П3", "КПД (доля тепла вверх)", "η₃", eta_formula("H89"), "—", "", None, "0.0000", formula=True)
param(91, "П3", "Удельная мощность (расчётная)", "n1₃", "=H72/(H90*H14*60)", "кВт/м²", "", None, NUM3, formula=True)
param(92, "П3", "Минимальная температура в петле", "tмин₃", "=1000*H91*H75+H9", "°C", "в конце предельной петли", None, NUM2, formula=True)
param(93, "П3", "Контрольное новое среднее", "tср₄", "=(H11+H92)/2", "°C", "сходимость |Δtср| < 0,1 °C", None, NUM2, formula=True)

# ================= 6. РЕЗУЛЬТАТЫ ТЕПЛОВОГО РАСЧЁТА =================
section(95, "6. ПЕТЛИ И РАСХОД (шаги 8–19)")
param(96, 1, "Средняя температура теплоносителя", "tср", "=H89", "°C", "из прохода 3", None, NUM2, formula=True)
param(97, 2, "КПД", "η", "=H90", "—", "", None, "0.0000", formula=True)
param(98, 3, "Удельная мощность", "n1", "=H91", "кВт/м²", "", None, NUM3, formula=True)
param(99, 4, "Суммарная мощность", "N1", "=H13*H98", "кВт", "n1·Fст", None, NUM3, formula=True)
param(100, 5, "Минимальная температура в петле", "tмин", "=H92", "°C", "нижний предел теплоносителя", None, NUM2, formula=True)
param(101, 6, "Допустимый перепад по петле", "tдоп", "=H11-H100", "К", "tпод − tмин (коллектор рядом с петлями)", None, NUM2, formula=True)
param(102, 7, "Падение температуры на 1 м трубы", "Δt", "=H98*H19/(H62*H23*H37*H38)", "К/м", "b·n1/(vпр·ω·ρж·cж)", None, NUM3, formula=True)
param(103, 8, "Падение температуры на 1 м² поверхности", "t1", "=H102/H19", "К/м²", "Δt/b", None, NUM3, formula=True)
param(104, 9, "Максимальная площадь петли", "fст", "=H101/H103", "м²", "tдоп/t1", None, NUM2, formula=True)
param(105, 10, "Максимальная длина петли", "Lмакс", "=H104/H19", "м", "fст/b", None, NUM2, formula=True)
param(106, 11, "Количество петель", "i", "=ROUNDUP(H13/H104,0)", "шт", "ROUNDUP(Fст/fст)", None, "0", formula=True)
param(107, 12, "Фактическая площадь петли", "f", "=H13/H106", "м²", "Fст/i", None, NUM2, formula=True)
param(108, 13, "СРЕДНЯЯ ДЛИНА ПЕТЛИ", "Lср", "=H107/H19", "м", "f/b — целевой выход", None, NUM2, formula=True)
param(109, 14, "Расход на одну петлю (проектный)", "Gп", "=3600*H62*H23", "м³/ч", "vпр·ω·3600", None, NUM3, formula=True)
param(110, 15, "Общий расход (проектный)", "Gv", "=H106*H109", "м³/ч", "i·Gп", None, NUM3, formula=True)
param(111, 16, "Перепад в фактической петле", "Δtф", "=H98*H107/(H62*H23*H37*H38)", "К", "n1·f/(vпр·ω·ρж·cж)", None, NUM2, formula=True)
param(112, 17, "Обратная температура (проектная)", "tобр", "=H11-H111", "°C", "", None, NUM2, formula=True)
put(113, 3, " Средняя температура (проектная, контроль) ")
ws.cell(row=113, column=6, value="tсрп")
ws.cell(row=113, column=8, value="=(H11+H112)/2").number_format = NUM2
ws.cell(row=113, column=10, value="°C")
ws.cell(row=113, column=12, value="ср. от проектного расхода — сравнить с проходом 3")

# ================= 7. РАБОЧАЯ ТОЧКА НАСОСА =================
put(115, 2, "7. РАБОЧАЯ ТОЧКА НАСОСА")
for col in range(2, 13):
    ws.cell(row=115, column=col).fill = FILL_SEC
    ws.cell(row=115, column=col).font = F_SECTION
put(116, 3, "Схема: насос → подающий коллектор (без сопротивления) → петли → обратный коллектор → магистраль (пренебрежимо) → котёл. В расчёте — только потери по длине петли.", font=F_NOTE)

CAT = "'Насосы-кривые'"
knots = ['$C$122', '$D$122', '$E$122', '$F$122', '$G$122', '$H$122', '$I$122']   # x0..x6
pvals = ['$C$123', '$D$123', '$E$123', '$F$123', '$G$123', '$H$123', '$I$123']   # h0..h6

put(122, 2, "Узлы каталожных точек Q, м³/ч:", font=F_BOLD)
for j, kx in enumerate([0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0]):
    put(122, 3 + j, kx, font=F_BOLD, num="0.0", align=AL_C)
put(123, 2, "Напор H выбранного насоса, м:", font=F_BOLD)
for k in range(1, 8):
    f = "=INDEX({}!$B$4:$G$10,{},MATCH($H$15,{}!$B$3:$G$3,0))".format(CAT, k, CAT)
    put(123, 2 + k, f, num="0.00", align=AL_C)

# заголовок таблицы РТ — упрощённая схема: только потери по длине в петлях
hdr = ["Q, м³/ч", "v петли, м/с", "ΔP длина петли, Па", "H сети, м", "H насоса, м", "Δ = Hнас − Hсети"]
put(125, 2, "Таблица характеристик (Q — общий расход; схема: насос → коллектор → петли → магистраль → котёл, все узлы кроме петель — без сопротивления):", font=F_BOLD)
for j, h in enumerate(hdr):
    put(126, 2 + j, h, font=F_BOLD, align=AL_C, border=True, fill=FILL_SEC)

REF_I = "$H$106"     # количество петель
REF_L = "$H$108"     # средняя длина петли

def pump_formula(q_cell):
    # кусочно-линейная интерполяция каталожной кривой выбранного насоса
    parts = []
    for j in range(1, 7):
        e = "{pjm1}+({pj}-{pjm1})*({q}-{xjm1})/({xdc}-{xjm1})".format(
            pjm1=pvals[j-1], pj=pvals[j], q=q_cell, xjm1=knots[j-1], xdc=knots[j])
        parts.append(e)
    s = "=" + "IF({q}<={k}, {e}, ".format(q=q_cell, k=knots[1], e=parts[0])
    for j in range(2, 7):
        s += "IF({q}<={k}, {e}, ".format(q=q_cell, k=knots[j], e=parts[j-1])
    s += "0" + ")" * 6
    return s

r0 = 127   # первая строка сетки
NQ = 25    # 0 .. 3.0 шаг 0.125
for idx in range(NQ):
    r = r0 + idx
    q = round(0.125 * idx, 3)
    put(r, 2, q, num="0.000", align=AL_C, border=True)
    put(r, 3, "=B{r}/({i}*3600*$H$23)".format(r=r, i=REF_I), num="0.000", align=AL_C, border=True)
    put(r, 4, "=$H$40*32000*POWER(C{r},1.79)/POWER($H$22*1000,1.29)*{L}".format(r=r, L=REF_L),
        num="0", align=AL_C, border=True)
    put(r, 5, "=D{r}/($H$37*9.81)".format(r=r), num="0.000", align=AL_C, border=True, fill=FILL_RES)
    put(r, 6, pump_formula("B{}".format(r)), num="0.000", align=AL_C, border=True)
    put(r, 7, "=F{r}-E{r}".format(r=r), num="0.000", align=AL_C, border=True)

r_last = r0 + NQ - 1   # 151

put(153, 3, "Строк с Δ<0 (диагностика)")
ws.cell(row=153, column=8, value='=COUNTIF(G127:G151,"<0")').number_format = "0"

# рабочая точка
rng_Q  = "$B$127:$B$151"
rng_D  = "$G$127:$G$151"   # Δ
rng_E  = "$E$127:$E$151"   # H сети
rng_F  = "$F$127:$F$151"   # H насоса

put(154, 3, "РАСХОД рабочей точки")
put(154, 12, "первая Δ<0: строка 25−n; интерполяция между последней Δ≥0 и первой Δ<0", font=F_NOTE)
ws.cell(row=154, column=8).value = ('=IFERROR(IF($H$153=0,"пересечение не найдено — выберите другой насос",'
    'INDEX({Q},25-$H$153)+INDEX({D},25-$H$153)*(INDEX({Q},26-$H$153)-INDEX({Q},25-$H$153))/(INDEX({D},25-$H$153)-INDEX({D},26-$H$153))),"—")').format(Q=rng_Q, D=rng_D)
ws.cell(row=154, column=8).number_format = NUM3
ws.cell(row=154, column=8).font = F_RES
ws.cell(row=154, column=8).fill = FILL_RES
ws.cell(row=154, column=8).border = B_ALL
ws.cell(row=154, column=10, value="м³/ч")

put(155, 3, "НАПОР рабочей точки (насос)")
ws.cell(row=155, column=8).value = ('=IFERROR(INDEX({F},25-$H$153)+(INDEX({F},26-$H$153)-INDEX({F},25-$H$153))*($H$154-INDEX({Q},25-$H$153))/(INDEX({Q},26-$H$153)-INDEX({Q},25-$H$153)),"—")').format(F=rng_F, Q=rng_Q)
ws.cell(row=155, column=8).number_format = NUM2
ws.cell(row=155, column=8).font = F_RES
ws.cell(row=155, column=8).fill = FILL_RES
ws.cell(row=155, column=8).border = B_ALL
ws.cell(row=155, column=10, value="м")

put(156, 3, "Напор сети в рабочей точке (контроль)")
ws.cell(row=156, column=8).value = ('=IFERROR(INDEX({E},25-$H$153)+(INDEX({E},26-$H$153)-INDEX({E},25-$H$153))*($H$154-INDEX({Q},25-$H$153))/(INDEX({Q},26-$H$153)-INDEX({Q},25-$H$153)),"—")').format(E=rng_E, Q=rng_Q)
ws.cell(row=156, column=8).number_format = NUM2
ws.cell(row=156, column=10, value="м (≈ H насоса)")

put(157, 3, "Скорость в петле (фактическая по РТ)")
ws.cell(row=157, column=8).value = '=IFERROR($H$154/($H$106*3600*$H$23),"—")'
ws.cell(row=157, column=8).number_format = NUM3
ws.cell(row=157, column=10, value="м/с")

put(158, 3, "Перепад в петле при РТ")
ws.cell(row=158, column=8).value = '=IFERROR($H$98*$H$107/($H$157*$H$23*$H$37*$H$38),"—")'
ws.cell(row=158, column=8).number_format = NUM2
ws.cell(row=158, column=10, value="К")

put(159, 3, "Обратная температура при РТ")
ws.cell(row=159, column=8).value = '=IFERROR($H$11-$H$158,"—")'
ws.cell(row=159, column=8).number_format = NUM2
ws.cell(row=159, column=10, value="°C")

put(160, 3, "Средняя температура при РТ")
ws.cell(row=160, column=8).value = '=IFERROR(($H$11+$H$159)/2,"—")'
ws.cell(row=160, column=8).number_format = NUM2
ws.cell(row=160, column=10, value="°C")

put(161, 3, "КПД при РТ")
ws.cell(row=161, column=8).value = '=IFERROR(' + eta_formula("H160")[1:] + ',"—")'
ws.cell(row=161, column=8).number_format = "0.0000"

put(162, 3, "Удельная мощность при РТ")
ws.cell(row=162, column=8).value = '=IFERROR($H$72/($H$161*$H$14*60),"—")'
ws.cell(row=162, column=8).number_format = NUM3
ws.cell(row=162, column=10, value="кВт/м²")

put(163, 3, "Мощность системы при РТ")
ws.cell(row=163, column=8).value = '=IFERROR($H$162*$H$13,"—")'
ws.cell(row=163, column=8).number_format = NUM2
ws.cell(row=163, column=10, value="кВт")

put(164, 3, "Статус скорости в петле")
ws.cell(row=164, column=8).value = ('=IF(ISNUMBER($H$157),IF(AND($H$157>=0.2,$H$157<=0.5),"OK","ВНИМАНИЕ: вне диапазона 0,2–0,5 м/с"),"—")')

# ================= 8. ИТОГ =================
section(166, "8. ИТОГ")
param(167, 1, "СРЕДНЯЯ ДЛИНА ПЕТЛИ", "Lср", "=H108", "м", "целевой выход № 2", None, NUM2, formula=True)
param(168, 2, "Количество петель", "i", "=H106", "шт", "", None, "0", formula=True)
param(169, 3, "РАСХОД рабочей точки", "Gрт", "=H154", "м³/ч", "целевой выход № 1", None, NUM3, formula=True)
param(170, 4, "НАПОР рабочей точки", "Hрт", "=H155", "м", "", None, NUM2, formula=True)
param(171, 5, "Скорость в петле при РТ", "v", "=H157", "м/с", "норма 0,2–0,5", None, NUM3, formula=True)
param(172, 6, "Мощность системы при РТ", "N1", "=H163", "кВт", "", None, NUM2, formula=True)
param(173, 7, "Статус скорости", "", "=H164", "", "", None, None, formula=True)

# ================= ЛИСТ «Насосы-кривые» =================
wsc = wb.create_sheet("Насосы-кривые")
ws = wsc  # put()/section() пишут в глобальный ws — переключаем на лист каталога
put(1, 2, "КАТАЛОЖНЫЕ ХАРАКТЕРИСТИКИ НАСОСОВ H(Q)", font=F_TITLE)
put(2, 2, "Значения — справочные аппроксимации каталожных кривых по узловым точкам. Уточняйте по паспорту насоса.", font=F_NOTE)

prs = [
    "Grundfos UPS 25-20",
    "Grundfos UPS 25-40",
    "Grundfos UPS 25-50",
    "Grundfos UPS 25-60",
    "Wilo Yonos Picos2 25/1-4",
    "Wilo Yonos Picos2 25/1-6",
]
HQ = {
    "Grundfos UPS 25-20":        [2.0, 1.8, 1.35, 0.85, 0.3, 0.0, 0.0],
    "Grundfos UPS 25-40":        [4.0, 3.6, 2.9, 2.0, 1.0, 0.2, 0.0],
    "Grundfos UPS 25-50":        [5.0, 4.6, 3.8, 2.7, 1.5, 0.5, 0.0],
    "Grundfos UPS 25-60":        [6.0, 5.5, 4.8, 3.9, 2.8, 1.5, 0.5],
    "Wilo Yonos Picos2 25/1-4":  [4.0, 3.75, 3.3, 2.7, 1.9, 1.0, 0.3],
    "Wilo Yonos Picos2 25/1-6":  [6.0, 5.6, 5.0, 4.2, 3.2, 2.0, 0.8],
}
qk = [0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0]

put(3, 1, "Q, м³/ч", font=F_BOLD, align=AL_C, border=True, fill=FILL_SEC)
for j, name in enumerate(prs):
    put(3, 2 + j, name, font=F_BOLD, align=AL_C, border=True, fill=FILL_SEC)
for k, q in enumerate(qk):
    rr = 4 + k
    put(rr, 1, q, num="0.0", align=AL_C, border=True)
    for j, name in enumerate(prs):
        put(rr, 2 + j, HQ[name][k], num="0.00", align=AL_C, border=True, fill=FILL_CONST)
wsc.column_dimensions["A"].width = 10
for j in range(len(prs)):
    wsc.column_dimensions[get_column_letter(2 + j)].width = 24

# ================= ЛИСТ «Справка» =================
wsh = wb.create_sheet("Справка")
ws = wsh  # переключаем на лист «Справка»
wsh.column_dimensions["A"].width = 3
wsh.column_dimensions["B"].width = 140
lines = [
    ("СИСТЕМА СНЕГОТАЯНИЯ — справка, допущения, источники", F_TITLE),
    ("", None),
    ("Методика: развитие файла snow_melt.xlsx (20 шагов). Новое: слои по новому ТЗ, вычисляемая tср (проходами),", None),
    ("рабочая точка насоса = пересечение характеристики сети и каталожной кривой выбранного насоса.", None),
    ("", None),
    ("Допущения и решения (см. plan.md, п. 8):", F_BOLD),
    ("1. Степень черноты ε=0,055 — модель исходника (занижает излучение; у снега/льда фактически 0,96–0,99).", None),
    ("2. R песка = 2,1 м²·К/Вт (зональная оценка, как в исходнике). Rвниз = ЭППС + бетон + песок ≈ 3,742 м²·К/Вт.", None),
    ("3. Kvs шарового крана = 6 (в исходнике ошибочно 44,38).", None),
    ("4. Коэффициент антифриза kА = 1,39 — как в исходнике (рабочий режим).", None),
    ("5. Средняя температура tср ВЫЧИСЛЯЕТСЯ: tср = (tпод + tмин)/2, цепочка tср→η→n1→tмин повторяется до сходимости", None),
    ("   (3 прохода; фактически сходимость уже после 2-го).", None),
    ("6. Гидравлика рабочей точки (УПРОЩЕННАЯ СХЕМА): насос → подающий коллектор (без сопротивления) → петли →", None),
    ("   обратный коллектор → магистральная труба (пренебрежимо мала) → котёл (без сопротивления).", None),
    ("   Учитываются ТОЛЬКО потери по длине петли: R = 3,2·10⁴·v^1,79/d^1,29 (Па/м, d в мм), с коэффициентом антифриза kА.", None),
    ("   Местные сопротивления (краны, фильтр, повороты, трёхходовик) и запас 1,2 — зарезервированы в константах,", None),
    ("   в расчёт РТ не входят (легко включить обратно).", None),
    ("", None),
    ("Свойства теплоносителя (пропиленгликоль 40%):", F_BOLD),
    ("ρ = 1032 кг/м³ (20 °C), tкристаллизации = −25 °C (у 30% −13, 35% −20, 45% −30).", None),
    ("Source: glikoli.ru, справочник «Плотность и температура замерзания водных растворов гликолей».", None),
    ("", None),
    ("Рабочая точка насоса:", F_BOLD),
    ("Характеристика сети строится по сетке Q (0…3 м³/ч, шаг 0,125): расход через петлю Q/i даёт скорость в петле;", None),
    ("перепад — только по длине петли (местные сопротивления и магистраль нулевые).", None),
    ("Кривая насоса — кусочно-линейная интерполяция каталожных узлов листа «Насосы-кривые».", None),
    ("Пересечение: первая строка, где Δ = Hнас − Hсети < 0; рабочая точка — линейная интерполяция с предыдущей строки.", None),
    ("", None),
    ("Проверки:", F_BOLD),
    ("Скорость в петле в рабочей точке должна быть в диапазоне 0,2–0,5 м/с (ячейка «Статус скорости»).", None),
    ("Проектная скорость в петле (vпр = 0,45 м/с) — параметр компоновки в константах; при неудачном подборе насоса", None),
    ("можно перебрать выбор модели выпадающим списком.", None),
    ("", None),
    ("Сверка (verify.py): ТОЛЬКО физические критерии — сложение сопротивлений слоёв, разумная мощность,", None),
    ("скорость в петле 0,2–0,5 м/с при рабочей точке, tмин внутри (tст; tпод), рост допустимой длины петли при сильном ветре.", None),
    ("Значения исходного файла НЕ используются как эталон — в них возможны ошибки.", None),
]
for i, (txt, f) in enumerate(lines, start=2):
    put(i, 2, txt, font=f) if f else put(i, 2, txt, font=Font(size=10))

# ---------------- сохранение ----------------
wb.save(r"D:\opencode\snow_melt\snow_melt_new.xlsx")
print("OK: snow_melt_new.xlsx saved")



