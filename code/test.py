import re

s = "示例 1.2．3 文本"

result = re.sub('[0-9\.．\s]', '', s)
# . . ．
print(result)

exit()

import re


def compare_versions(version):
    # 将版本号解析成级别的数字列表
    levels = re.split(r'\.', version)
    levels = [int(level) for level in levels]
    print(levels)
    return levels


def valid_distance(x, y):  # x -> y
    x_level, y_level = x.split("."), y.split('.')
    x_levels, y_levels = [int(x_item) for x_item in x_level], [int(y_item) for y_item in y_level]
    # result = False
    min_length = min(len(x_levels), len(y_levels))
    if len(x_levels) >= len(y_levels):
        if x_levels[min_length - 1] + 1 == y_levels[min_length - 1]:
            return True
        else:
            return False
    else:
        result = True
        for i in range(min_length):
            if x_levels[i] != y_levels[i]:
                result = False
        if len(x_levels) + 1 != len(y_levels):
            return False
        if y_levels[-1] != 1:
            result = False
        return result

    # return result


a = "1.1.1"
b = "1.1"

result = valid_distance(a, b)
print(result)
exit()

a = ['1.2', '1.3', '2', '2.3', '2.10', '3', '1.5', '6', '1.1.1']

sorted_a = sorted(a, key=compare_versions)
print(sorted_a)
