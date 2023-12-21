from Util import Util
import xml.dom.minidom

if __name__ == "__main__":
    tools: Util = Util('肖露露毕业论文-正文部分.docx')
    body = tools.docx_body
    content = ""
    for elem in body.childNodes:
        elem: xml.dom.minidom.Element
        if elem.tagName != "w:p":
            continue
        content += tools.getFullText(elem)+"\n"
        # content
    with open("content.txt", "w", encoding="utf-8") as f:
        f.write(content)