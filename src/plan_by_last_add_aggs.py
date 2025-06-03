import pandas as pd

from src.calculate_kilometers_and_price import (calculate_kilometers_and_flights,
                                                load_special_routes)

from src.data_loader import load_data
from src.args_parser import parse_args


def main():
    args = parse_args()
    AGG_NUM = args.AGG_NUMS
    ADD_NUMBER = args.ADD_NUMBERS

    result = pd.DataFrame()
    for AGG_NUM, ADD_NUMBER in zip(AGG_NUM, ADD_NUMBER):

        coefs_df, capacities_df, routes, plan_flights_df, plan_dists_df, holidays_df = (
            load_data(AGG_NUM, ADD_NUMBER)
        )
        special_routes = load_special_routes(holidays_df)

        res_df = pd.DataFrame(columns=["Маршрут", "Кол-во рейсов", "КМ"])
        start, end = args.first_n_last_days_of_month.split("-")
        real_period = pd.date_range(start, end, freq="D")
        for route in routes:
            if route == "101":
                continue

            km_forward, km_reverse, km_total, flights_total = (
                calculate_kilometers_and_flights(
                    route, real_period, plan_flights_df, plan_dists_df, special_routes
                )
            )
            res_df.loc[len(res_df)] = [route, flights_total, km_total]

        result = pd.concat(
            [result, res_df[["Маршрут", "Кол-во рейсов", "КМ"]]], ignore_index=True
        )

    result.to_excel("results/план-факт по последним ДС.xlsx", index=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# python plan_by_last_add_aggs.py --ADD_NUMBERS 32 32 30 --AGG_NUMS 219 220 222 --first_n_last_days_of_month "07.01.2024-07.31.2024"
