import re

pattern = r'\d+(\.\d+)*'


def find_number_sequences(text):
    sequences = []
    match = re.search(pattern, text)
    print(match.group())
    while match:
        sequences.append(match.group())
        text = text[match.end():]
        match = re.search(pattern, text)
    return sequences


# 测试
text = "这是一些数字序列: 1, 1.1, 1.1.1，还有其他的文本。"
sequences = find_number_sequences(text)
print(sequences)
