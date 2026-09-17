"""
Генератор очереди на основе истории.

Идея: чем позже студент стоял в очереди в прошлом, тем выше его шанс
оказаться ближе к началу в новой очереди. Это реализовано через
взвешенную случайную выборку без возврата (weighted random sampling
without replacement): каждому студенту присваивается "вес" по средней
относительной позиции в истории, и очередь строится пошагово —
каждый следующий стоящий выбирается случайно, но с вероятностью,
пропорциональной весу.

Формат входного xlsx:
    - Один лист.
    - Каждый столбец = одна историческая очередь.
    - Заголовок столбца — произвольное имя (например, "Очередь 1", дата и т.п.).
    - В ячейках столбца — имена студентов, стоявших в очереди сверху вниз
      (первая строка после заголовка = первый в очереди).
    - Пустые ячейки допустимы (если очереди разной длины).

Результат:
    - В таблицу добавляется новый столбец с новой сгенерированной очередью.
    - Файл сохраняется (по умолчанию — поверх исходного, либо в указанный output).

Это MVP: обработка ошибок минимальна (например, не проверяется дублирование
имён внутри одного столбца, кодировки, битые файлы и т.п.).
"""

import argparse
import random
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook


def read_history(path: str, sheet_name=0) -> pd.DataFrame:
    """Читает историю очередей из xlsx. Каждый столбец — одна прошлая очередь."""
    return pd.read_excel(path, sheet_name=sheet_name)


def compute_weights(df: pd.DataFrame, neutral_weight: float = 0.5, epsilon: float = 0.1,
                     power: float = 1.0) -> tuple[dict, list]:
    """
    Вычисляет вес каждого студента на основе истории.

    Для каждого появления студента в очереди считается его нормализованная
    позиция (0.0 — начало очереди, 1.0 — конец). Вес студента — это среднее
    таких позиций по всем найденным очередям, возведённое в степень `power`
    (чтобы при желании усиливать/ослаблять эффект истории) плюс небольшая
    epsilon, чтобы вес никогда не был нулевым (иначе студент, всегда
    стоявший первым, никогда не попал бы в новую очередь).

    Студенты, отсутствующие в истории (новенькие), получают нейтральный вес
    `neutral_weight` — как если бы они всегда стояли "в середине".

    Возвращает (словарь вес по студенту, список всех студентов в порядке
    первого появления).
    """
    position_lists: dict[str, list[float]] = {}
    all_students: list[str] = []

    for col in df.columns:
        queue = [str(x).strip() for x in df[col].dropna().tolist()]
        length = len(queue)
        if length == 0:
            continue
        denom = max(length - 1, 1)  # чтобы не делить на 0 при очереди из 1 человека
        for pos, student in enumerate(queue):
            if student not in position_lists:
                position_lists[student] = []
                all_students.append(student)
            position_lists[student].append(pos / denom)

    weights: dict[str, float] = {}
    for student in all_students:
        avg_pos = sum(position_lists[student]) / len(position_lists[student])
        weights[student] = (avg_pos ** power) + epsilon

    return weights, all_students


def generate_queue(weights: dict, students: list, neutral_weight: float = 0.5,
                    epsilon: float = 0.1) -> list:
    """
    Строит новую очередь взвешенной случайной выборкой без возврата.

    На каждом шаге среди ещё не расставленных студентов случайно выбирается
    один, с вероятностью, пропорциональной его весу (чем выше вес — тем выше
    шанс попасть РАНЬШЕ, т.к. выбор идёт по одному "с начала очереди").
    """
    pool = list(students)
    w = dict(weights)
    queue = []

    while pool:
        total = sum(w.get(s, neutral_weight + epsilon) for s in pool)
        r = random.uniform(0, total)
        upto = 0.0
        chosen = pool[-1]  # запасной вариант на случай ошибок округления float
        for s in pool:
            upto += w.get(s, neutral_weight + epsilon)
            if upto >= r:
                chosen = s
                break
        queue.append(chosen)
        pool.remove(chosen)

    return queue


def append_new_column(path: str, new_queue: list, output_path: str = None,
                       sheet_name=None, column_name: str = None) -> str:
    """Дописывает новую очередь как новый столбец в xlsx и сохраняет файл."""
    wb = load_workbook(path)
    ws = wb[sheet_name] if sheet_name else wb.active

    next_col = ws.max_column + 1
    if column_name is None:
        # timestamp вместо порядкового номера — чтобы не пересекаться
        # с уже существующими названиями столбцов ("Очередь 1", "Очередь 2", ...)
        column_name = f"Очередь {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    ws.cell(row=1, column=next_col, value=column_name)
    for i, student in enumerate(new_queue, start=2):
        ws.cell(row=i, column=next_col, value=student)

    save_path = output_path or path
    wb.save(save_path)
    return save_path


def main():
    parser = argparse.ArgumentParser(
        description="Генерация новой очереди на основе истории из xlsx."
    )
    parser.add_argument("input", help="Путь к xlsx-файлу с историей очередей")
    parser.add_argument("-o", "--output", help="Куда сохранить результат "
                                                "(по умолчанию — перезаписать входной файл)")
    parser.add_argument("--sheet", default=0, help="Имя или индекс листа (по умолчанию первый)")
    parser.add_argument("--column-name", default=None,
                         help="Название нового столбца (по умолчанию 'Очередь N')")
    parser.add_argument("--power", type=float, default=1.0,
                         help="Степень усиления эффекта истории (>1 — сильнее тянет "
                              "тех, кто был в конце, к началу; <1 — слабее). По умолчанию 1.0")
    parser.add_argument("--seed", type=int, default=None,
                         help="Фиксировать случайность (для повторяемости/тестов)")
    args = parser.parse_args()

    if args.seed is not None:
        random.seed(args.seed)

    df = read_history(args.input, sheet_name=args.sheet)
    weights, students = compute_weights(df, power=args.power)
    new_queue = generate_queue(weights, students)

    saved_to = append_new_column(
        args.input, new_queue,
        output_path=args.output,
        sheet_name=args.sheet if isinstance(args.sheet, str) else None,
        column_name=args.column_name,
    )

    print(f"Новая очередь сохранена в: {saved_to}")
    print("Порядок:")
    for i, student in enumerate(new_queue, start=1):
        print(f"  {i}. {student}")


if __name__ == "__main__":
    main()