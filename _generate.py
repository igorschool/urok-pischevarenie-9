# -*- coding: utf-8 -*-
"""Урок 10. Пищеварение и ферменты. Презентация и PDF-раздатки."""
from pathlib import Path

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

OUT = Path(__file__).resolve().parent
ASSETS = OUT / "assets"
PDF = OUT / "pdf"
PDF.mkdir(exist_ok=True)

pdfmetrics.registerFont(TTFont("A", r"C:\Windows\Fonts\arial.ttf"))
pdfmetrics.registerFont(TTFont("B", r"C:\Windows\Fonts\arialbd.ttf"))

INK = "1C2B2E"
TEAL = "0E3B43"
TEAL2 = "1F6F78"
SAND = "F6F1EA"
CORAL = "C4553A"
GOLD = "C4A035"
WHITE = "FFFFFF"
MINT = "E7F3F1"
LINE = "D5E4E2"


def rgb(h):
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_run_font(run, name="Calibri"):
    run.font.name = name
    rpr = run._r.get_or_add_rPr()
    for tag in ("latin", "ea", "cs"):
        el = rpr.find(qn(f"a:{tag}"))
        if el is None:
            el = etree.SubElement(rpr, qn(f"a:{tag}"))
        el.set("typeface", name)


def rect(slide, l, t, w, h, color, radius=False):
    shape = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    s = slide.shapes.add_shape(shape, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(color)
    s.line.fill.background()
    if radius:
        s.adjustments[0] = 0.08
    return s


def tb(slide, l, t, w, h, text, size, bold=False, color=INK, align="left", anchor="top"):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    anc = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE, "bottom": MSO_ANCHOR.BOTTOM}[anchor]
    tf.paragraphs[0].alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[align]
    box.text_frame.paragraphs[0].space_before = Pt(0)
    box.text_frame.paragraphs[0].space_after = Pt(0)
    # anchor
    tf._txBody.bodyPr.set("anchor", {"top": "t", "middle": "ctr", "bottom": "b"}[anchor])
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[align]
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        r = p.add_run()
        r.text = line
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = rgb(color)
        set_run_font(r)
    return box


def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def footer(slide, page, total):
    rect(slide, 0, 7.22, 13.333, 0.28, TEAL)
    tb(slide, 0.35, 7.22, 9, 0.28, "9 класс  ·  Урок 10  ·  Пищеварение и ферменты", 11, color=WHITE, anchor="middle")
    tb(slide, 11.2, 7.22, 1.8, 0.28, f"{page} / {total}", 11, color=WHITE, align="right", anchor="middle")


def new(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, 13.333, 7.5, SAND)
    return s


def pic(slide, name, l, t, w, h):
    path = ASSETS / name
    return slide.shapes.add_picture(str(path), Inches(l), Inches(t), Inches(w), Inches(h))


TOTAL = 19


def build_pptx():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "9 класс. Урок 10. Пищеварение и ферменты"
    page = 0

    def foot(s):
        nonlocal page
        page += 1
        footer(s, page, TOTAL)

    # 1 title
    s = new(prs)
    rect(s, 0, 0, 7.15, 7.5, TEAL)
    pic(s, "clinic_wide.jpg", 7.15, 0, 6.19, 7.5)
    tb(s, 0.55, 1.35, 6.2, 0.4, "ОТКРЫТЫЙ УРОК  ·  9 КЛАСС", 16, True, GOLD)
    tb(s, 0.55, 1.9, 6.2, 2.2, "Сегодня вы —\nгастроэнтерологи", 40, True, WHITE)
    tb(s, 0.55, 4.35, 6.2, 1.2, "Путь пищи и фермент,\nкоторый её расщепляет", 22, color="D7EBE8")
    tb(s, 0.55, 6.15, 6.2, 0.5, "9.1.2.1    ·    9.1.2.2", 16, True, GOLD)
    foot(s)
    notes(s, "До звонка. Цели вслух не читайте. Одна фраза: сегодня вы ведёте приём и ищете, какой фермент за какое вещество отвечает.")

    # 2 route
    s = new(prs)
    rect(s, 0, 0, 13.333, 1.15, TEAL)
    tb(s, 0.5, 0.28, 12, 0.65, "Как идёт приём", 32, True, WHITE)
    steps = [
        ("1", "Приём", "диалог у доски"),
        ("2", "Вывод", "зачем нам карта"),
        ("3", "Объяснение", "путь и ферменты"),
        ("4", "Группа", "4 роли, ответ"),
        ("5", "Группы", "жалоба Алины"),
        ("6", "Игра", "рука и рефлексия"),
    ]
    for i, (n, title, sub) in enumerate(steps):
        x = 0.4 + i * 2.15
        rect(s, x, 2.1, 2.0, 3.5, WHITE, radius=True)
        tb(s, x + 0.15, 2.3, 1.7, 0.7, n, 32, True, CORAL)
        tb(s, x + 0.15, 3.15, 1.7, 1.0, title, 20, True, TEAL)
        tb(s, x + 0.15, 4.3, 1.7, 0.9, sub, 14, color=INK)
    foot(s)
    notes(s, "20 секунд. Можно не задерживаться. Дальше сразу игра «Приём».")

    # 3 open game
    s = new(prs)
    rect(s, 0, 0, 13.333, 1.15, TEAL)
    tb(s, 0.5, 0.28, 12, 0.65, "К доске выходит врач", 32, True, WHITE)
    pic(s, "doctor_gastro.jpg", 0.45, 1.55, 3.3, 4.4)
    pic(s, "patient_alina.jpg", 3.95, 1.55, 3.3, 4.4)
    rect(s, 7.5, 1.7, 5.35, 4.15, WHITE, radius=True)
    tb(s, 7.75, 1.95, 4.95, 1.3, "Откройте\n«Приём Алины»", 28, True, TEAL)
    tb(s, 7.75, 3.5, 4.95, 1.8, "Ученик говорит за врача.\nВы говорите за Алину.\nТаблетки не назначаем.", 18, color=INK)
    foot(s)
    notes(s, "Браузер: http://127.0.0.1:8770/igra_priem.html  Один ученик у доски, он врач. Вы читаете Алину обычным голосом, как девочка на приёме. Кнопка «Дальше» ведёт к жалобам. После итога класс называет тему, следующий «Дальше» открывает тему и цели. Потом вернитесь на следующий слайд.")

    # 4 conclusion
    s = new(prs)
    rect(s, 0, 0, 0.22, 7.5, CORAL)
    tb(s, 0.7, 0.45, 12, 0.4, "ЧТО МЫ УСЛЫШАЛИ", 16, True, CORAL)
    tb(s, 0.7, 1.2, 12, 2.6, "Хлеб Алина ест спокойно.\nПосле молока бывает вздутие.\nМясо стоит вверху живота.\nОт жирного тошнит.", 32, True, TEAL)
    tb(s, 0.7, 4.7, 12, 1.5, "Еда застревает в разных местах.\nСейчас разберём, где и чем.", 26, color=INK)
    foot(s)
    notes(s, "Это уже вы классу, не врач. Спросите: почему хлеб проходит, а молоко нет? Ответ не говорите.")

    # 5 question
    s = new(prs)
    rect(s, 0, 0, 13.333, 7.5, TEAL)
    tb(s, 0.8, 1.8, 11.5, 1.0, "На сегодня", 20, True, GOLD)
    tb(s, 0.8, 2.6, 11.5, 2.6, "Куда девается обед\nи почему от молока пучит,\nа от хлеба нет?", 36, True, WHITE)
    foot(s)
    notes(s, "Не отвечайте сами. Пауза. Дальше показываете путь пищи.")

    # 6 path
    s = new(prs)
    rect(s, 0, 0, 13.333, 1.1, TEAL)
    tb(s, 0.45, 0.25, 8, 0.65, "Ход пищи", 32, True, WHITE)
    pic(s, "tract.jpg", 0.35, 1.35, 4.3, 5.6)
    organs = [
        ("1", "Рот", "измельчение и первая химия"),
        ("2", "Пищевод", "только продвижение"),
        ("3", "Желудок", "белки"),
        ("4", "Тонкая кишка", "основное расщепление и всасывание"),
        ("5", "Толстая кишка", "вода, формирование каловых масс"),
    ]
    for i, (n, name, sub) in enumerate(organs):
        y = 1.4 + i * 1.08
        rect(s, 5.0, y, 7.8, 0.95, WHITE, radius=True)
        tb(s, 5.2, y + 0.12, 0.7, 0.7, n, 26, True, CORAL, anchor="middle")
        tb(s, 6.0, y + 0.08, 6.4, 0.42, name, 22, True, TEAL)
        tb(s, 6.0, y + 0.48, 6.4, 0.38, sub, 14, color=INK)
    foot(s)
    notes(s, "Рисунок — опора. Названия читайте по списку, не ищите на рисунке мелкие детали. Подчеркните: пищевод сам пищу не расщепляет.")

    # 7 mouth stomach
    s = new(prs)
    rect(s, 0, 0, 13.333, 1.1, TEAL)
    tb(s, 0.45, 0.25, 12, 0.65, "Два первых фермента", 32, True, WHITE)
    rect(s, 0.4, 1.5, 6.15, 5.3, WHITE, radius=True)
    tb(s, 0.7, 1.75, 5.6, 0.4, "РОТ", 14, True, CORAL)
    tb(s, 0.7, 2.2, 5.6, 0.7, "Амилаза", 32, True, TEAL)
    tb(s, 0.7, 3.1, 5.6, 2.8, "Крахмал  →  мальтоза\n\nЗубы измельчают пищу.\nАмилаза слюны начинает\nрасщепление крахмала.", 22, color=INK)
    rect(s, 6.8, 1.5, 6.15, 5.3, WHITE, radius=True)
    tb(s, 7.1, 1.75, 5.6, 0.4, "ЖЕЛУДОК", 14, True, CORAL)
    tb(s, 7.1, 2.2, 5.6, 0.7, "Пепсин", 32, True, TEAL)
    tb(s, 7.1, 3.1, 5.6, 2.8, "Белки  →  пептиды\n\nСоляная кислота включает пепсин.\nСлизь защищает стенку желудка.", 22, color=INK)
    foot(s)
    notes(s, "Пример на доске, «я показываю»: крахмал — амилаза. Спросите: почему мясо «стоит вверху живота», а хлеб часто нет? Не раскрывайте молоко — это задание групп.")

    # 8 intestine
    s = new(prs)
    rect(s, 0, 0, 13.333, 1.1, TEAL)
    tb(s, 0.45, 0.25, 12, 0.65, "Тонкая кишка — главное место", 32, True, WHITE)
    rows = [
        ("Амилаза поджелудочной", "остаток крахмала → мальтоза"),
        ("Трипсин", "белки и пептиды → короче"),
        ("Липаза", "жиры → жирные кислоты и глицерин"),
        ("Лактаза", "лактоза → глюкоза и галактоза"),
        ("Мальтаза и пептидазы", "мальтоза → глюкоза; пептиды → аминокислоты"),
    ]
    for i, (a, b) in enumerate(rows):
        y = 1.4 + i * 1.05
        rect(s, 0.45, y, 12.4, 0.92, WHITE, radius=True)
        tb(s, 0.7, y + 0.18, 5.3, 0.6, a, 22, True, TEAL, anchor="middle")
        tb(s, 6.2, y + 0.18, 6.3, 0.6, b, 20, color=INK, anchor="middle")
    foot(s)
    notes(s, "Проговорите медленно. Желчь упомянете на следующем слайде одной фразой: это не фермент, лабораторная по желчи — следующий урок.")

    # 9 bile + absorption
    s = new(prs)
    rect(s, 0, 0, 13.333, 1.1, TEAL)
    tb(s, 0.45, 0.25, 12, 0.65, "Что ещё важно в пути", 32, True, WHITE)
    rect(s, 0.4, 1.5, 6.2, 5.3, WHITE, radius=True)
    tb(s, 0.7, 1.8, 5.6, 0.5, "ЖЕЛЧЬ", 14, True, CORAL)
    tb(s, 0.7, 2.4, 5.6, 1.5, "Это не фермент.", 32, True, TEAL)
    tb(s, 0.7, 4.1, 5.6, 2.2, "Она дробит жир на капли,\nчтобы липаза могла работать.\nРасщепляет жир липаза.", 20, color=INK)
    rect(s, 6.85, 1.5, 6.05, 5.3, WHITE, radius=True)
    tb(s, 7.15, 1.8, 5.5, 0.5, "ВСАСЫВАНИЕ", 14, True, CORAL)
    tb(s, 7.15, 2.4, 5.5, 3.6, "Ворсинки тонкой кишки.\n\nСахара и аминокислоты — в кровь.\nПродукты жиров — в лимфу.\n\nТолстая кишка забирает воду.", 20, color=INK)
    foot(s)
    notes(s, "Один устный вопрос: что всасывается в тонкой кишке? Коротко поправьте. Потом группы по 4. Желчь глубоко не разбирайте — это урок 11–12.")

    # 10 pair
    s = new(prs)
    rect(s, 0, 0, 13.333, 1.1, TEAL)
    tb(s, 0.45, 0.25, 12, 0.65, "Группа из 4", 32, True, WHITE)
    pic(s, "food_bread.jpg", 0.45, 1.55, 2.5, 2.5)
    tb(s, 3.2, 1.5, 9.5, 2.5, "Один лист на группу.\nУ каждого своя роль.\nПотом один общий ответ.\n\nФерменты пока не пишем.", 22, color=INK)
    rect(s, 0.45, 4.35, 12.4, 2.4, WHITE, radius=True)
    tb(s, 0.75, 4.55, 11.8, 2.0, "1  Порядок органов.\n2  Пищевод.   3  Желудок.   4  Тонкая кишка.\nНа доске только правильные ответы.", 22, color=TEAL)
    foot(s)
    notes(s, "Раздайте один групповой лист на четырёх. Роли 1–4 каждый пишет сам. Потом вместе собирают общий ответ. На доске http://127.0.0.1:8770/pary_doska.html только правильные ответы: сначала роли, общий ответ откройте, когда группа свою фразу уже сказала. Ферменты не пишите.")

    # 11 check reminder
    s = new(prs)
    rect(s, 0, 0, 13.333, 7.5, TEAL)
    tb(s, 0.8, 2.2, 11.5, 1.5, "Сверяем лист", 40, True, WHITE)
    tb(s, 0.8, 4.0, 11.5, 1.2, "Смотрите на доску.\nГде не так — поправьте ручкой.", 24, color="D7EBE8")
    foot(s)
    notes(s, "На доске кнопки «Показать ответ». Сначала роли, общий ответ — когда группа уже собрала свою фразу. Где не так, поправляют ручкой.")

    # 12 AI 1
    s = new(prs)
    rect(s, 0, 0, 13.333, 1.1, TEAL)
    tb(s, 0.45, 0.22, 8, 0.7, "Мне в чате написали", 30, True, WHITE)
    tb(s, 9.2, 0.28, 3.6, 0.55, "к доске 1 ученик", 16, True, GOLD, align="right")
    rect(s, 0.7, 1.8, 11.9, 3.3, WHITE, radius=True)
    tb(s, 1.0, 2.15, 11.3, 2.6, "«Молоко расщепляет пепсин\nв желудке».", 36, True, CORAL)
    tb(s, 0.7, 5.4, 12, 1.2, "Алина бы с этим согласилась?\nСкажите, где путаница.", 24, color=INK)
    foot(s)
    notes(s, "Для гостей это и есть момент с помощником: ответ из чата, ученик его проверяет. Слово «ИИ» на слайд не выносите. Верный смысл: молоко — лактоза, её расщепляет лактаза в тонкой кишке. Пепсин в желудке расщепляет белки. Следующий слайд после ответа ученика.")

    # 13 AI 1 key
    s = new(prs)
    rect(s, 0, 0, 0.22, 7.5, TEAL2)
    tb(s, 0.7, 0.5, 12, 0.4, "КАК СКАЗАТЬ АЛИНЕ", 16, True, TEAL2)
    tb(s, 0.7, 1.2, 12, 2.2, "Молоко расщепляет не пепсин.\nЛактозу расщепляет лактаза,\nи уже в тонкой кишке.", 32, True, TEAL)
    tb(s, 0.7, 4.2, 12, 1.6, "Пепсин в желудке расщепляет мясо, белок.\nИз лактозы получаются глюкоза и галактоза.", 22, color=INK)
    foot(s)
    notes(s, "Если ученик сказал верно — подтвердите и идите дальше. Если нет — эта карточка и есть подсказка, не оценка.")

    # 14 AI 2
    s = new(prs)
    rect(s, 0, 0, 13.333, 1.1, TEAL)
    tb(s, 0.45, 0.22, 8.5, 0.7, "В чате написали ещё", 30, True, WHITE)
    tb(s, 8.6, 0.28, 4.2, 0.55, "к доске 2-й ученик", 16, True, GOLD, align="right")
    rect(s, 0.7, 1.8, 11.9, 3.3, WHITE, radius=True)
    tb(s, 1.0, 2.15, 11.3, 2.6, "«Желчь — это фермент,\nона расщепляет жиры».", 34, True, CORAL)
    tb(s, 0.7, 5.4, 12, 1.2, "Где тут неправда?", 24, color=INK)
    foot(s)
    notes(s, "Желчь не фермент. Жир режет липаза. Желчь только дробит его на капли. В лабораторную по желчи не уходите. Следующий слайд после ответа.")

    s = new(prs)
    rect(s, 0, 0, 0.22, 7.5, TEAL2)
    tb(s, 0.7, 0.5, 12, 0.4, "КАК СКАЗАТЬ", 16, True, TEAL2)
    tb(s, 0.7, 1.2, 12, 2.2, "Желчь не фермент.\nЖир режет липаза.", 36, True, TEAL)
    tb(s, 0.7, 4.1, 12, 1.8, "Желчь только дробит жир на капли,\nчтобы липазе было что резать.", 24, color=INK)
    foot(s)
    notes(s, "Коротко. Не уходите в лабораторную по желчи — она на следующем уроке.")

    # 15 groups
    s = new(prs)
    rect(s, 0, 0, 13.333, 1.1, TEAL)
    tb(s, 0.45, 0.25, 12, 0.65, "Группы: одна жалоба", 32, True, WHITE)
    cards = [
        ("Молоко", "вздутие", "food_milk.jpg"),
        ("Мясо", "вверху живота", "food_meat.jpg"),
        ("Жирное", "тяжесть", "food_fat.jpg"),
        ("Хлеб", "почти норма", "food_bread.jpg"),
    ]
    for i, (title, sub, img) in enumerate(cards):
        x = 0.4 + i * 3.25
        rect(s, x, 1.5, 3.05, 5.25, WHITE, radius=True)
        pic(s, img, x + 0.2, 1.7, 2.65, 2.5)
        tb(s, x + 0.15, 4.35, 2.75, 0.7, title, 24, True, TEAL, align="center")
        tb(s, x + 0.15, 5.15, 2.75, 0.8, sub, 16, color=INK, align="center")
    foot(s)
    notes(s, "4 группы, карточка PDF каждой. 6 минут. Заполняют: вещество, фермент, орган, продукт, одно предложение «потому что». Если класса мало, те же карточки у меньших групп. Потом группа «Молоко» выходит к доске.")

    # 16 board path
    s = new(prs)
    rect(s, 0, 0, 13.333, 1.1, TEAL)
    tb(s, 0.45, 0.25, 12, 0.65, "У доски: проведите молоко Алины", 28, True, WHITE)
    steps2 = [
        "Рот. Лактоза здесь уже распадается?",
        "Пищевод. Тут есть фермент?",
        "Желудок. Пепсин сейчас про молоко или про мясо?",
        "Тонкая кишка. Что расщепляет лактозу и что получается?",
        "Ворсинки. Куда уходит глюкоза?",
    ]
    for i, line in enumerate(steps2):
        y = 1.4 + i * 1.05
        rect(s, 0.5, y, 12.3, 0.92, WHITE, radius=True)
        tb(s, 0.75, y + 0.18, 11.8, 0.6, f"{i+1}   {line}", 22, color=TEAL, anchor="middle")
    foot(s)
    notes(s, "Это результат, который должны увидеть гости. Ученик говорит сам. Ориентир: в рту лактоза цела; пищевод только везёт; в желудке пепсин про белки, не про лактозу; в тонкой кишке лактаза → глюкоза и галактоза; ворсинки → в кровь. Следующий слайд — после ответа, если нужно опереться.")

    # 17 game + reflection cue
    s = new(prs)
    rect(s, 0, 0, 13.333, 1.1, TEAL)
    tb(s, 0.45, 0.25, 12, 0.65, "Что расщепляет еду", 32, True, WHITE)
    tb(s, 0.55, 1.6, 12, 1.6, "Держите палец на карточке.\nЕсли камера молчит — нажмите мышкой.", 26, color=INK)
    tb(s, 0.55, 3.5, 12, 1.2, "http://127.0.0.1:8770/igra_ruki.html", 28, True, TEAL)
    tb(s, 0.55, 5.1, 12, 1.4, "Двое у доски по очереди.\nОстальные следят и поправляют вслух.", 22, color=INK)
    foot(s)
    notes(s, "Сервер должен быть запущен файлом запуск_сервер.bat, иначе камера часто молчит. 4–5 минут. Потом рефлексия на бумаге.")

    # 18 reflection
    s = new(prs)
    rect(s, 0, 0, 13.333, 1.1, TEAL)
    tb(s, 0.45, 0.25, 12, 0.65, "Две минуты, на листочке", 32, True, WHITE)
    lines = [
        "Могу провести обед по органам: да / ещё путаюсь",
        "Фермент, который объясню соседу: …",
        "Какую еду он расщепляет: …",
        "Что сегодня помогло понять: …",
    ]
    for i, line in enumerate(lines):
        y = 1.55 + i * 1.25
        rect(s, 0.5, y, 12.3, 1.08, WHITE, radius=True)
        tb(s, 0.8, y + 0.25, 11.7, 0.6, line, 24, color=TEAL, anchor="middle")
    foot(s)
    notes(s, "Лист рефлексии уже на парте или раздайте сейчас. Вслух ничего сдавать не нужно. Соберите листы. На этом приём закрыт.")

    out = OUT / "Презентация_Пищеварение.pptx"
    prs.save(out)
    print("pptx", out, "slides", page)


def wrap(c, text, font, size, x, y, max_w, leading, color=(0.11, 0.17, 0.18)):
    c.setFillColorRGB(*color)
    c.setFont(font, size)
    words = text.split()
    line = ""
    for w in words:
        trial = (line + " " + w).strip()
        if c.stringWidth(trial, font, size) <= max_w:
            line = trial
        else:
            c.drawString(x, y, line)
            y -= leading
            line = w
    if line:
        c.drawString(x, y, line)
        y -= leading
    return y


def header_bar(c, title, sub):
    c.setFillColorRGB(0.055, 0.23, 0.26)
    c.rect(0, 277 * mm, 210 * mm, 20 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("B", 14)
    c.drawString(14 * mm, 286 * mm, title)
    c.setFont("A", 9)
    c.drawString(14 * mm, 280 * mm, sub)


def footer_bar(c, page):
    c.setFillColorRGB(0.93, 0.35, 0.22)
    c.rect(0, 0, 210 * mm, 10 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("A", 8)
    c.drawString(14 * mm, 3.5 * mm, "9 класс  ·  Урок 10  ·  Пищеварение")
    c.drawRightString(196 * mm, 3.5 * mm, str(page))


def role_box(c, x, y, w, h, num, title, lines):
    c.setStrokeColorRGB(0.05, 0.23, 0.26)
    c.setFillColorRGB(1, 1, 1)
    c.roundRect(x, y, w, h, 3 * mm, fill=1, stroke=1)
    c.setFillColorRGB(0.05, 0.23, 0.26)
    c.roundRect(x, y + h - 8 * mm, w, 8 * mm, 3 * mm, fill=1, stroke=0)
    c.rect(x, y + h - 8 * mm, w, 4 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("B", 11)
    c.drawString(x + 3 * mm, y + h - 5.5 * mm, num + "  " + title)
    c.setFillColorRGB(0.11, 0.17, 0.18)
    c.setFont("A", 10)
    ty = y + h - 13 * mm
    for line in lines:
        c.drawString(x + 3 * mm, ty, line)
        ty -= 5.2 * mm


def build_pair_pdf():
    path = PDF / "01_Групповой_лист.pdf"
    c = canvas.Canvas(str(path), pagesize=A4)
    c.setFillColorRGB(1, 0.96, 0.86)
    c.rect(0, 0, A4[0], A4[1], fill=1, stroke=0)
    c.setFillColorRGB(0.93, 0.35, 0.22)
    c.rect(0, 277 * mm, 210 * mm, 20 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("B", 16)
    c.drawString(14 * mm, 286 * mm, "Групповой лист  ·  4 участника")
    c.setFont("A", 9)
    c.drawString(14 * mm, 280 * mm, "Один лист на группу. Сначала каждый свою роль. Потом вместе один общий ответ. Ферменты не пишите.")
    y = 268 * mm
    c.setFillColorRGB(0.11, 0.17, 0.18)
    c.setFont("A", 11)
    c.drawString(14 * mm, y, "Имена:  1 ____________________   2 ____________________")
    y -= 6 * mm
    c.drawString(14 * mm, y, "         3 ____________________   4 ____________________")
    role_box(c, 12 * mm, 178 * mm, 92 * mm, 72 * mm, "1", "Порядок", [
        "Органы вразброс. Поставьте номера.",
        "желудок, рот, толстая кишка,",
        "пищевод, тонкая кишка.",
        "",
        "1 ____________________",
        "2 ____________________",
        "3 ____________________",
        "4 ____________________",
        "5 ____________________",
    ])
    role_box(c, 108 * mm, 178 * mm, 90 * mm, 72 * mm, "2", "Пищевод", [
        "Пищевод пищу не переваривает.",
        "Допишите, что он делает.",
        "",
        "Он только",
        "____________________",
        "____________________",
    ])
    role_box(c, 12 * mm, 100 * mm, 92 * mm, 74 * mm, "3", "Желудок", [
        "Желудок пищу",
        "____________________",
        "",
        "Мясо стоит вверху,",
        "под рёбрами.",
        "Это орган:",
        "____________________",
    ])
    role_box(c, 108 * mm, 100 * mm, 90 * mm, 74 * mm, "4", "Тонкая кишка", [
        "В тонкой кишке пища",
        "____________________",
        "",
        "После молока вздутие",
        "по всему животу.",
        "Это орган:",
        "____________________",
    ])
    c.setFillColorRGB(0.05, 0.23, 0.26)
    c.roundRect(12 * mm, 16 * mm, 186 * mm, 80 * mm, 3 * mm, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("B", 12)
    c.drawString(16 * mm, 86 * mm, "Общий ответ. Пишут все четверо, когда роли готовы.")
    c.setFont("A", 11)
    lines = [
        "Пища идёт: __________ → __________ → __________ → __________ → __________",
        "",
        "Пищевод ________________________________.",
        "Желудок ________________________________. Там стоит мясо.",
        "В тонкой кишке ________________________________.",
        "Там бывает вздутие от молока.",
    ]
    ty = 78 * mm
    c.setFillColorRGB(1, 1, 1)
    for line in lines:
        c.drawString(16 * mm, ty, line)
        ty -= 8 * mm
    footer_bar(c, "группа")
    c.showPage()
    c.save()
    print("pdf", path)


def card_page(c, title, complaint, img, lines):
    header_bar(c, "Карточка группы  ·  " + title, "Жалоба Алины. Заполняет вся группа, говорит один.")
    c.drawImage(str(ASSETS / img), 14 * mm, 222 * mm, 32 * mm, 32 * mm, preserveAspectRatio=True, anchor="c", mask="auto")
    c.setFillColorRGB(0.055, 0.23, 0.26)
    c.setFont("B", 16)
    c.drawString(52 * mm, 244 * mm, title)
    wrap(c, complaint, "A", 11, 52 * mm, 236 * mm, 140 * mm, 5 * mm)
    y = 208 * mm
    y = wrap(
        c,
        "Прочитайте жалобу вслух. Болезнь не пишите и «надо к врачу» не пишите. Заполните четыре строки. Внизу одно обычное предложение: почему при такой еде бывает именно так. Потом один из вас говорит это классу.",
        "A", 10, 14 * mm, y, 182 * mm, 4.3 * mm,
    )
    y -= 4 * mm
    for label in lines:
        c.setFillColorRGB(0.055, 0.23, 0.26)
        c.setFont("B", 12)
        c.drawString(14 * mm, y, label)
        y -= 6 * mm
        c.setStrokeColorRGB(0.7, 0.78, 0.77)
        c.line(14 * mm, y, 196 * mm, y)
        y -= 11 * mm
    c.setFillColorRGB(0.055, 0.23, 0.26)
    c.setFont("B", 12)
    c.drawString(14 * mm, y, "Почему так, одним предложением")
    y -= 6 * mm
    c.setStrokeColorRGB(0.7, 0.78, 0.77)
    c.line(14 * mm, y, 196 * mm, y)
    y -= 12 * mm
    c.line(14 * mm, y, 196 * mm, y)
    footer_bar(c, title)


def build_group_pdf():
    path = PDF / "02_Групповые_карточки.pdf"
    c = canvas.Canvas(str(path), pagesize=A4)
    cards = [
        ("Молоко", "«Молоко и мороженое — через час начинается вздутие, живот урчит.»", "food_milk.jpg"),
        ("Мясо", "«Мясо ем, и оно будто стоит вот тут, вверху живота.»", "food_meat.jpg"),
        ("Жирное", "«После бешбармака тошнит. Тяжесть долго не отпускает.»", "food_fat.jpg"),
        ("Хлеб", "«А хлеб ем спокойно.» Объясните, почему эта еда её не беспокоит.", "food_bread.jpg"),
    ]
    labels = ["Какая это еда", "Какой фермент её расщепляет", "Где это происходит", "Что получается"]
    for title, complaint, img in cards:
        card_page(c, title, complaint, img, labels)
        c.showPage()
    c.save()
    print("pdf", path)


def build_reflex_pdf():
    path = PDF / "03_Рефлексия.pdf"
    c = canvas.Canvas(str(path), pagesize=A4)
    for i, top in enumerate((148 * mm, 0)):
        c.setFillColorRGB(0.055, 0.23, 0.26)
        c.rect(0, top + 128 * mm, 210 * mm, 16 * mm, fill=1, stroke=0)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("B", 13)
        c.drawString(12 * mm, top + 134 * mm, "Две минуты, для себя")
        c.setFillColorRGB(0.11, 0.17, 0.18)
        c.setFont("A", 11)
        c.drawString(12 * mm, top + 118 * mm, "Имя: ______________________________")
        lines = [
            "Могу провести обед по органам:   да     /     ещё путаюсь",
            "Фермент, который объясню соседу: _______________________________",
            "Какую еду он расщепляет: ____________________________________________",
            "Что сегодня помогло понять: ____________________________________",
        ]
        y = top + 104 * mm
        for line in lines:
            c.setFont("A", 11)
            c.drawString(12 * mm, y, line)
            y -= 16 * mm
        if i == 0:
            c.setStrokeColorRGB(0.7, 0.7, 0.7)
            c.setDash(3, 3)
            c.line(8 * mm, 148 * mm, 202 * mm, 148 * mm)
            c.setDash()
            c.setFont("A", 8)
            c.setFillColorRGB(0.45, 0.45, 0.45)
            c.drawString(12 * mm, 150 * mm, "разрезать")
    c.save()
    print("pdf", path)


def build_key_pdf():
    path = PDF / "04_Ключ_учителю.pdf"
    c = canvas.Canvas(str(path), pagesize=A4)
    header_bar(c, "Ключ учителю", "Не раздавать")
    y = 262 * mm
    c.setFillColorRGB(0.77, 0.33, 0.23)
    c.setFont("B", 12)
    c.drawString(14 * mm, y, "НЕ РАЗДАВАТЬ")
    y -= 10 * mm
    blocks = [
        "Групповой лист, 4 роли, без ферментов. Ход: рот, пищевод, желудок, тонкая кишка, толстая кишка.",
        "Роль 2: пищевод только проводит пищу.",
        "Роль 3: желудок удерживает и перемешивает. Мясо вверху — желудок.",
        "Роль 4: в тонкой кишке пища всасывается. Вздутие от молока — тонкая кишка.",
        "Общий ответ: пища идёт рот → пищевод → желудок → тонкая кишка → толстая кишка. Пищевод только проводит. Желудок удерживает и перемешивает, там стоит мясо. В тонкой кишке пища всасывается, там вздутие от молока.",
        "Белки — пепсин — желудок — пептиды.",
        "Жиры — липаза — тонкая кишка — жирные кислоты и глицерин.",
        "Лактоза — лактаза — тонкая кишка — глюкоза и галактоза.",
        "Крахмал (образец) — амилаза — рот и тонкая кишка — мальтоза.",
        "Молоко у доски: в рту лактоза цела; пищевод только везёт; желудок — пепсин для белков; тонкая кишка — лактаза до глюкозы и галактозы; ворсинки — в кровь.",
        "Чат про молоко: пепсин молоко не расщепляет. Лактозу расщепляет лактаза в тонкой кишке.",
        "Чат про желчь: желчь не фермент. Жир режет липаза. Желчь только дробит капли.",
        "Группа «Хлеб»: хлеб не беспокоит, потому что амилаза расщепляет крахмал.",
        "Группа «Мясо»: в желудке белок расщепляет пепсин. Если не успевает, мясо «стоит» вверху живота.",
        "Группа «Жирное»: жир режет липаза в тонкой кишке. Желчь только дробит капли.",
        "Группа «Молоко»: лактаза. Если лактоза не расщепилась, она остаётся в кишке, отсюда урчание. Это объяснение для урока, не болезнь.",
    ]
    for text in blocks:
        y = wrap(c, text, "A", 11, 14 * mm, y, 180 * mm, 5.2 * mm)
        y -= 3 * mm
    footer_bar(c, "ключ")
    c.showPage()
    c.save()
    print("pdf", path)


if __name__ == "__main__":
    build_pptx()
    build_pair_pdf()
    build_group_pdf()
    build_reflex_pdf()
    build_key_pdf()
