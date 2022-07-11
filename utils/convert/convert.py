from datetime import date
from typing import Union


def to_str(value: Union[date, str, int, None]) -> str:
    if value is None:
        return ""
    elif type(value) is str:
        return value
    else:
        return str(value)
