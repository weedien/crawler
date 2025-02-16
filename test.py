import re


def strip_only_once(s: str, ch: chr) -> str:
    if s.startswith(ch) and s.endswith(ch):
        return s[1:-1]
    else:
        return s


str1 = '"\\"Awaken, My Love!\\""'
result1 = strip_only_once(re.sub(r"\\*", "", str1), '"')
print(result1)
