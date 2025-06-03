import argparse


def parse_args() -> argparse.Namespace:
    """
    аргументы, передаваемые через терминал
    """
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--ADD_NUMBERS", required=True, nargs="+", type=int, help="Номера ДС"
    )
    parser.add_argument(
        "--AGG_NUMS", required=True, nargs="+", type=int, help="Номера ГК"
    )
    parser.add_argument(
        "--first_n_last_days_of_month",
        default=None,
        required=False,
        type=str,
        help="Первый и последний дни месяца, для которого подсчитывается план по последним ДС",
    )
    return parser.parse_args()