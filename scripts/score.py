#!/usr/bin/env python3
"""Считает итоговый балл идей из data/ideas.json и печатает сводную таблицу.

Запуск:
    python3 scripts/score.py            # markdown-таблица, отсортированная по баллу
    python3 scripts/score.py --json     # те же данные с баллом в JSON

Чтобы пересчитать рейтинг под свои приоритеты, поменяйте meta.weights в data/ideas.json.
"""
import json
import math
import sys
from fractions import Fraction
from pathlib import Path

DATA = Path(__file__).resolve().parent.parent / "data" / "ideas.json"

COMPETITION_OPENNESS = {"Нет": Fraction(1), "Низкая": Fraction(3, 4), "Средняя": Fraction(1, 2), "Высокая": Fraction(1, 4)}
REG_SAFETY = {"Низкий": Fraction(1), "Средний": Fraction(3, 5), "Высокий": Fraction(1, 5)}
COST_LABEL = {1: "до 0,3 млн ₽", 2: "0,3–1 млн ₽", 3: "1–3 млн ₽", 4: "3+ млн ₽"}
TYPE_LABEL = {"world": "🌍 Есть в мире", "new": "💡 Новое"}


def score(idea, weights):
    parts = {
        "profit": Fraction(idea["profit"], 5),
        "popularity": Fraction(idea["popularity"], 5),
        "competition": COMPETITION_OPENNESS[idea["competition"]],
        "ease": Fraction(6 - idea["complexity"], 5),
        "cheapness": Fraction(4 - idea["cost_tier"], 3),
        "reg_safety": REG_SAFETY[idea["reg_risk"]],
    }
    w = {k: Fraction(str(v)) for k, v in weights.items()}
    raw = 100 * sum(w[k] * v for k, v in parts.items()) / sum(w.values())
    return math.floor(raw + Fraction(1, 2))  # точное округление «половина вверх», как ROUND в Excel


def dots(n, full="●", empty="○"):
    return full * n + empty * (5 - n)


def main():
    data = json.loads(DATA.read_text(encoding="utf-8"))
    weights = data["meta"]["weights"]
    ideas = data["ideas"]
    for idea in ideas:
        idea["score"] = score(idea, weights)
    ideas.sort(key=lambda i: (-i["score"], i["id"]))

    if "--json" in sys.argv:
        print(json.dumps(ideas, ensure_ascii=False, indent=2))
        return

    header = (
        "| # | ID | Продукт | Тип | Аналог в мире | Сложность | Стоимость MVP | Срок MVP "
        "| Прибыльность | Популярность | Конкуренция в РФ | Регуляторный риск | Монетизация | Балл |"
    )
    print(header)
    print("|" + "---|" * 14)
    for rank, i in enumerate(ideas, 1):
        print(
            f"| {rank} | [{i['id']}](#{i['id'].lower()}) | {i['name']} | {TYPE_LABEL[i['type']]} | {i['analog']} "
            f"| {dots(i['complexity'])} {i['complexity']} | {'💰' * i['cost_tier']} {COST_LABEL[i['cost_tier']]} "
            f"| {i['mvp_time']} | {dots(i['profit'], '★', '☆')} | {dots(i['popularity'], '★', '☆')} "
            f"| {i['competition']} | {i['reg_risk']} | {i['monetization']} | **{i['score']}** |"
        )


if __name__ == "__main__":
    main()
