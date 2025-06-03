import os
from datetime import datetime

import pandas as pd
from tqdm import tqdm

from src.args_parser import parse_args
from src.calculate_kilometers_and_price import (load_special_routes,
                                                process_routes_in_period)
from src.data_loader import load_data


def main():
    # Константы
    args = parse_args()

    AGG_NUM = args.AGG_NUMS
    ADD_NUMBER = args.ADD_NUMBERS
    START_DATE = os.environ["START_DATE"]
    END_DATE = os.environ["END_DATE"]

    start_months = pd.date_range(start=START_DATE, end=END_DATE, freq="3MS").union(
        [datetime.strptime("2028-07-01", "%Y-%m-%d")]
    )
    end_months = pd.date_range(start="2022-12-31", end=END_DATE, freq="3M").union(
        [datetime.strptime(END_DATE, "%Y-%m-%d")]
    )

    coefs_df, capacities_df, routes, plan_flights_df, plan_dists_df, holidays_df = (
        load_data(AGG_NUM, ADD_NUMBER)
    )
    special_routes = load_special_routes(holidays_df)

    res_df = pd.DataFrame(
        columns=[
            "Маршрут",
            "Период",
            "1",
            "Кол-во рейсов",
            "КМ от НП",
            "КМ от КП",
            "КМ",
            "Стоимость",
        ]
    )

    for idx_period, (start_period, end_period) in tqdm(
        enumerate(zip(start_months, end_months), start=2), total=len(start_months)
    ):
        period = pd.date_range(start_period, end_period, freq="D")
        period_str = (
            f"{start_period.strftime('%d.%m.%Y')}-{end_period.strftime('%d.%m.%Y')}"
        )

        # Обработка маршрутов в текущем периоде
        res_df = process_routes_in_period(
            routes,
            period,
            period_str,
            coefs_df,
            capacities_df,
            plan_flights_df,
            plan_dists_df,
            special_routes,
            res_df,
            idx_period,
        )

    # Группировка результатов и запись в файлы
    grouped = res_df.groupby("Период").sum().reset_index()
    grouped = grouped.loc[
        pd.to_datetime(grouped["Период"].str.split("-").str[0], format="%d.%m.%Y")
        .sort_values()
        .index
    ]

    grouped.to_excel(
        f"results/по кварталам гк{AGG_NUM} дс{ADD_NUMBER}.xlsx", index=False
    )
    res_df.to_excel(
        f"results/по маршрутам гк{AGG_NUM} дс{ADD_NUMBER}.xlsx", index=False
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
