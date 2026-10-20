#!/usr/bin/env python3
"""Считает итоговый балл идей из data/ideas.json и печатает сводную таблицу.

Запуск:
    python3 scripts/score.py            # markdown-таблица, отсортированная по баллу
    python3 scripts/score.py --json     # те же данные с баллом в JSON

Чтобы пересчитать рейтинг под свои приоритеты, поменяйте meta.weights в data/ideas.json.
"""
import json
import sys
from pathlib import Path

DATA = Path(__file__).resolve().parent.parent / "data" / "ideas.json"

COMPETITION_OPENNESS = {"Нет": 1.0, "Низкая": 0.75, "Средняя": 0.5, "Высокая": 0.25}
REG_SAFETY = {"Низкий": 1.0, "Средний": 0.6, "Высокий": 0.2}
COST_LABEL = {1: "до 0,3 млн ₽", 2: "0,3–1 млн ₽", 3: "1–3 млн ₽", 4: "3+ млн ₽"}
TYPE_LABEL = {"world": "🌍 Есть в мире", "new": "💡 Новое"}


def score(idea, weights):
    parts = {
        "profit": idea["profit"] / 5,
        "popularity": idea["popularity"] / 5,
        "competition": COMPETITION_OPENNESS[idea["competition"]],
        "ease": (6 - idea["complexity"]) / 5,
        "cheapness": (4 - idea["cost_tier"]) / 3,
        "reg_safety": REG_SAFETY[idea["reg_risk"]],
    }
    return round(100 * sum(weights[k] * v for k, v in parts.items()) / sum(weights.values()))


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
