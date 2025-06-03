from datetime import datetime, timedelta

import pandas as pd

from consts.special_cases import (special_cases_wo_year,
                                  special_cases,
                                  special_cases_wo_year_w_condition)


def get_value_for_date(df, route, date, forward_key, reverse_key):
    route_data = df[df["Маршрут"] == route].sort_values(by="Дата", ascending=False)
    for _, row in route_data.iterrows():
        date_route = (
            pd.to_datetime(row["Дата"], dayfirst=True)
            if isinstance(row["Дата"], str)
            else row["Дата"]
        )
        if date >= date_route:
            return row[forward_key], row[reverse_key]
    return 0, 0


def load_special_routes(holidays_df):
    special_routes = {}

    for _, row in holidays_df.iterrows():
        weekday = int(row["День недели"])
        date_str = row["Дата"]
        routes = row["Маршруты"].split(", ")

        # Проверяем, является ли дата периодом
        if "-" in date_str:
            start_date_str, end_date_str = date_str.split("-")
            start_date = datetime.strptime(start_date_str.strip(), "%d.%m.%Y").date()
            end_date = datetime.strptime(end_date_str.strip(), "%d.%m.%Y").date()

            # Проходим по каждому дню в периоде
            current_date = start_date
            while current_date <= end_date:
                special_routes.setdefault(weekday, {}).setdefault(
                    current_date, []
                ).extend(routes)
                current_date += timedelta(days=1)
        else:
            # Если дата не является периодом, просто обрабатываем ее как один день
            date = datetime.strptime(date_str.strip(), "%d.%m.%Y").date()
            special_routes.setdefault(weekday, {}).setdefault(date, []).extend(routes)

    return special_routes


def determine_weekday(date, route, special_routes):
    def check_special_cases(special_cases, date_format):
        for weekday, dates in special_cases.items():
            if date.strftime(date_format) in dates:
                return int(weekday)
        return None



    # Проверка на основе файла с особыми маршрутами
    for weekday, dates_routes in special_routes.items():
        if date in dates_routes and route in dates_routes[date]:
            return weekday

    # Проверка случаев с годом
    weekday = check_special_cases(special_cases, "%d.%m.%Y")
    if weekday is not None:
        return weekday

    # Проверка случаев без учета года
    weekday = check_special_cases(special_cases_wo_year, "%d.%m")
    if weekday is not None:
        return weekday

    # Проверка, если дата приходится на рабочий день
    if any(
        date.strftime("%d.%m") in dates
        for dates in special_cases_wo_year_w_condition.values()
    ):
        return 5 if date.weekday() + 1 < 5 else date.weekday() + 1
    return date.weekday() + 1


def calculate_kilometers_and_flights(
    route, period, plan_flights_df, plan_dists_df, special_routes
):
    km_forward, km_reverse = 0, 0
    flights_total = 0
    for date in period:
        weekday = determine_weekday(date.date(), route, special_routes)

        flights_forward, flights_reverse = get_value_for_date(
            plan_flights_df, route, date, f"Прямо - {weekday}", f"Обратно - {weekday}"
        )
        dist_forward, dist_reverse = get_value_for_date(
            plan_dists_df, route, date, "от НП", "от КП"
        )

        km_forward += flights_forward * dist_forward
        km_reverse += flights_reverse * dist_reverse
        flights_total += flights_forward + flights_reverse

    km_total = km_forward + km_reverse
    return km_forward, km_reverse, km_total, flights_total


def calculate_price(route, km_total, coefs_period, capacities_df):
    coef = coefs_period.query(f"Маршрут == '{route}'").iloc[0, 1]
    cap = capacities_df.query(f"Маршрут == '{route}'").iloc[0, 1]
    price = km_total * coef * cap
    return price


def process_routes_in_period(
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
):
    coefs_period = coefs_df[["Маршрут", period_str]]
    for route in routes:
        if route == "101":
            continue

        km_forward, km_reverse, km_total, flights_total = (
            calculate_kilometers_and_flights(
                route, period, plan_flights_df, plan_dists_df, special_routes
            )
        )

        # Вычисление стоимости
        price = calculate_price(route, km_total, coefs_period, capacities_df)

        res_df.loc[len(res_df)] = [
            route,
            period_str,
            idx_period,
            flights_total,
            km_forward,
            km_reverse,
            km_total,
            price,
        ]

    return res_df
