import logging

import pandas as pd


def validate_dataframe(
    filename,
    name,
    df: pd.DataFrame,
    forward_km: str,
    reverse_km: str,
    tolerance: float = 1e-5,
) -> None:
    """
    Проверка на правильность заполнения табл. "Количество рейсов и пробег транспортных средств" в приложении к контракту
    """
    logging.basicConfig(
        filename=filename,
        filemode="a",
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    logging.info(f"{name}")
    df = df.map(lambda x: float(x.replace(",", ".")))
    forward_km, reverse_km = float(forward_km.replace(",", ".")), float(
        reverse_km.replace(",", ".")
    )
    expected_ratios = {"Прямое": f"{forward_km}", "Обратное": f"{reverse_km}"}

    df = df.map(lambda x: float(str(x).replace(",", ".")))

    for direction in df.index[
        :-1
    ]:  # Проходим по всем строкам, кроме последней ("ИТОГО")
        expected_ratio = expected_ratios.get(direction)
        expected_ratio = float(expected_ratio.replace(",", "."))

        if expected_ratio is None:
            logging.error(
                f"Ожидаемое отношение для '{direction}' не задано в expected_ratios."
            )
            raise ValueError(
                f"Ожидаемое отношение для '{direction}' не задано в expected_ratios."
            )

        # Найти столбцы с "Пробег, км" и "Количество рейсов" в MultiIndex
        mileage_columns = df.columns[df.columns.get_level_values(1) == "Пробег, км"]
        count_columns = df.columns[
            df.columns.get_level_values(1) == "Количество рейсов"
        ]

        if len(mileage_columns) != len(count_columns):
            logging.error(
                "Количество столбцов 'Пробег, км' не соответствует количеству столбцов 'Количество рейсов'."
            )
            raise ValueError(
                "Количество столбцов 'Пробег, км' не соответствует количеству столбцов 'Количество рейсов'."
            )

        for m_col, c_col in zip(mileage_columns, count_columns):
            mileage_value = df.loc[direction, m_col]
            count_value = df.loc[direction, c_col]

            ratio = round(mileage_value / count_value, 2)
            logging.info(
                f"Направление: {direction}, КМ: {ratio}, Ожидаемые КМ: {expected_ratio}"
            )

            if abs(ratio - expected_ratio) > tolerance:
                logging.error(
                    f"Ошибка: отношение в строке '{direction}' между столбцами "
                    f"'{m_col}' и '{c_col}' "
                    f"равно {ratio}, но должно быть {expected_ratio}."
                )
                raise ValueError(
                    f"Ошибка: отношение в строке '{direction}' между столбцами "
                    f"'{m_col}' и '{c_col}' "
                    f"равно {ratio}, но должно быть {expected_ratio}."
                )
        # Проверка правильности строки "ИТОГО"
    calculated_totals = df.iloc[:-1].sum()
    logging.info(
        f"Ожидаемые итоги {calculated_totals.T.to_markdown(buf=None, mode='wt', index=True)}"
    )
    if not (
        abs(calculated_totals - df.loc["ИТОГО"]) <= tolerance
    ).all():  # Проверяем на допустимое отклонение
        logging.error(
            f"Ошибка: данные в строке 'ИТОГО' некорректны! {calculated_totals}, {df.loc['ИТОГО']}"
        )
        raise ValueError(
            f"Ошибка: данные в строке 'ИТОГО' некорректны! {calculated_totals}, {df.loc['ИТОГО']}"
        )

    # Логируем успешное завершение всех проверок

    # with pd.option_context('display.max_rows', None, 'display.max_columns', None, 'display.expand_frame_repr', False):
    df_str = df.to_markdown(buf=None, mode="wt", index=True)
    logging.info(f"Сравниваемые данные:\n{df_str}")
    logging.info("Все проверки пройдены успешно.\n")
    # print("Все проверки пройдены успешно.")
