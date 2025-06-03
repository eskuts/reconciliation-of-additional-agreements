import os

import pandas as pd
from dotenv import load_dotenv


def load_data(AGG_NUM, ADD_NUMBER):
    load_dotenv()
    plans_path = (
        os.environ["plans"]
        .replace("AGG_NUM", str(AGG_NUM))
        .replace("ADD_NUMBER", str(ADD_NUMBER))
    )
    capacities_path = os.environ["path_to_cap"]
    routes_path = os.environ["routes"]
    coefs_path = os.environ["path_to_coefs"]
    holidays = os.environ["path_to_holidays"]

    coefs_df = pd.read_excel(coefs_path, sheet_name="Sheet1").assign(
        Маршрут=(lambda df: df["Маршрут"].astype(str).str.upper())
    )
    capacities_df = pd.read_excel(capacities_path).assign(
        Маршрут=(lambda df: df["Маршрут"].astype(str).str.upper())
    )
    routes_df = pd.read_excel(routes_path).query(f"ГК == {AGG_NUM}")
    routes = routes_df["Маршрут"].astype(str).str.upper().unique().tolist()

    # plan_df = pd.read_excel(plans_path, sheet_name=None)
    # plan_flights_df = plan_df["Кол-во рейсов"].assign(Маршрут=(lambda df: df["Маршрут"].astype(str).str.upper()))
    # plan_dists_df = plan_df["Протяженности"].assign(Маршрут=(lambda df: df["Маршрут"].astype(str).str.upper()))

    plan_df = pd.read_excel(plans_path, sheet_name="data")
    plan_df["Дата"] = pd.to_datetime(plan_df["Дата"], format="%d.%m.%Y")
    plan_dists_df = plan_df.iloc[:, :5].assign(
        Маршрут=(lambda df: df["Маршрут"].astype(str).str.upper())
    )
    plan_flights_df = pd.concat(
        [plan_df.iloc[:, :2], plan_df.iloc[:, -14:]], axis=1
    ).assign(Маршрут=(lambda df: df["Маршрут"].astype(str).str.upper()))

    holidays_df = (
        pd.read_excel(holidays, sheet_name=str(AGG_NUM), dtype=str).fillna(0)
        if str(AGG_NUM) in pd.ExcelFile(holidays).sheet_names
        else pd.DataFrame()
    )

    return coefs_df, capacities_df, routes, plan_flights_df, plan_dists_df, holidays_df