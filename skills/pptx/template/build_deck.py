#!/usr/bin/env python3
"""Deck source: one function per slide.

Build with ./build.sh, review with the skill's render.nu and audit.py.
"""

from deck import Deck, NIX, NEUTRAL, OK_TINT, PYTHON, WARN, WARN_TINT

OUT = "deck.pptx"

deck = Deck(
    head="Georgia",
    body="Calibri",
    code="DejaVu Sans Mono",
)


def slide_title():
    slide = deck.title_slide(
        ["Название доклада", "и вторая строка, если нужна"],
        subtitle="подзаголовок",
        year="2026",
        theme=NEUTRAL,
    )
    deck.notes(slide, "Что сказать на титуле: одна фраза о задаче и одна о том, "
                      "что слушатель унесёт.")


def slide_problem():
    slide = deck.slide()
    deck.header(slide, "Проблема", NEUTRAL)
    deck.bullets(slide, [
        "первый тезис, без точки в конце",
        "второй тезис",
        "третий тезис",
    ], top=deck.top, size=17, lowercase=True)
    deck.key(slide, "Одна мысль, ради которой существует слайд.", NEUTRAL)
    deck.notes(slide, "Раскрыть каждый пункт одним предложением.")


def slide_code():
    slide = deck.slide()
    deck.header(slide, "Пример кода", PYTHON)
    deck.code(slide, "Python", '''def greet(name: str) -> str:
    return f"hello, {name}"

print(greet("world"))''', "python", left=deck.M, top=deck.top,
              width=deck.CW * 0.55, size=13, theme=PYTHON)
    deck.bullets(slide, [
        "типы проверяются при выполнении",
        "функция — такое же значение, как число",
    ], left=deck.M + deck.CW * 0.6, top=deck.top,
       width=deck.CW * 0.38, size=13, lowercase=True)
    deck.key(slide, "Вывод: подпись под кодом, а не ещё один абзац.", PYTHON)
    deck.notes(slide, "Проговорить, что именно делает код, и не читать его вслух.")


def slide_table():
    slide = deck.slide()
    deck.header(slide, "Таблица", NIX)
    deck.table(slide, [
        ["Язык", "Типизация", "Память"],
        ["C++", "статическая", "ручная"],
        ["Python", "динамическая", "подсчёт ссылок"],
        ["Nix", "динамическая", "неизменяемые значения"],
    ], col_fracs=[0.2, 0.4, 0.4], font_size=14, row_height=0.62, theme=NIX)
    deck.key(slide, "Таблица отвечает на один вопрос, а не пересказывает текст.", NIX)
    deck.notes(slide, "Пройти по строкам и назвать контрпример.")


def slide_section():
    slide = deck.slide()
    deck.section(slide, "Раздел", NIX, number="01",
                 subtitle="Одна фраза о том, зачем этот раздел.",
                 notes="Переход: что было и что сейчас будет.")


def slide_compare():
    slide = deck.slide()
    deck.header(slide, "Два подхода", NEUTRAL)
    box_w = deck.CW / 2 - 0.15
    deck.box(slide, "Подход A\n\nчто он даёт\nи чем платит", deck.M, 1.7,
             box_w, 2.4, fill=OK_TINT, color=NEUTRAL.accent, valign="top")
    deck.box(slide, "Подход B\n\nчто он даёт\nи чем платит",
             deck.M + deck.CW / 2 + 0.15, 1.7,
             box_w, 2.4, fill=WARN_TINT, color=WARN, valign="top")
    deck.key(slide, "Сравнение без победителя: разные ограничения.", NEUTRAL)
    deck.notes(slide, "Назвать, какое ограничение снимает каждый подход и чем платит.")


def slide_closing():
    slide = deck.slide()
    deck.bar(slide, NEUTRAL)
    deck.text(slide, "Один вывод, который должен остаться",
              deck.M, 1.1, deck.CW, 1.0, size=28, color=NEUTRAL.accent,
              bold=True, font=deck.head)
    deck.bullets(slide, [
        "что теперь слушатель знает",
        "что с этим можно сделать",
    ], top=3.0, size=18, lowercase=True)
    deck.key(slide, "Спасибо.", NEUTRAL)
    deck.notes(slide, "Повторить один вывод и остановиться. Вопросы — после.")


def build():
    slide_title()
    slide_problem()
    slide_section()
    slide_code()
    slide_table()
    slide_compare()
    slide_closing()

    deck.properties(
        title="Название доклада",
        subject="о чём доклад",
        author="Имя Фамилия",
        keywords="ключевые слова",
    )
    deck.save(OUT, thumbnail=True, page_numbers=True)
    print(f"saved: {OUT} ({len(deck.prs.slides._sldIdLst)} slides)")


if __name__ == "__main__":
    build()
