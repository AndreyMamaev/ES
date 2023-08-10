from datetime import date as date_class
from datetime import timedelta, datetime


def get_quarter(p_date: date_class) -> int:
    return (p_date.month - 1) // 3 + 1


def get_first_day_of_the_quarter(p_date: date_class):
    return datetime(p_date.year, 3 * ((p_date.month - 1) // 3) + 1, 1)


def get_last_day_of_the_quarter(p_date: date_class):
    quarter = get_quarter(p_date)
    return (datetime(p_date.year + 3 * quarter // 12, 3 * quarter % 12 + 1, 1) + timedelta(days=-1))
