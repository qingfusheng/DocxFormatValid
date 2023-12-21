import re


def split_to_sentences(temp_content):
    temp_lines = temp_content.split("\n")
    temp_result = []
    reg = re.compile(r'(.*?[。？])')
    for line in temp_lines:
        temp = []
        reg_line = reg.findall(line)
        if reg_line:
            temp.extend(reg_line)
        else:
            temp.append(line.strip())
        temp_result.append(temp)
    return temp_result


if __name__ == "__main__":
    with open("content.txt", "r", encoding="utf-8") as f:
        content = f.read()
    result = split_to_sentences(content)
    result_text = ""
    for each in result:
        for each2 in each:
            result_text += each2 + "\n"
        result_text += "---------------------------------------------------\n"
    with open("content3.txt", "w", encoding="utf-8") as f:
        f.write(result_text)
