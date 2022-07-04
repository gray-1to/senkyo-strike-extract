from typing import Any


def to_str(value: Any) -> str:
    if value is None:
        return ""
    elif type(value) is str:
        return value
    else:
        return str(value)
