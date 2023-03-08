import re


def transform_to_date(x):
    x = int(x)
    hour = int(x / 60)
    minute = int(x % 60)
    return f"{hour} saat {minute} dakika"


def get_name(x):
    splitted = x.split(',')
    if len(splitted) == 1:
        return x
    else:
        return splitted[-1]


def get_last_name(x):
    splitted = x.split(',')
    if len(splitted) == 1:
        return x
    if len(splitted) > 2:
        return splitted[-2]
    else:
        return splitted[-2]


def match_string(string, strings_to_match):
    regex_string = ".*|.*".join(strings_to_match)
    regex_string = f".*{regex_string}.*"
    x = re.match(regex_string, string, re.IGNORECASE)
    if x:
        return True
    else:
        return False
