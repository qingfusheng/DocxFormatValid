# the first target is extract content from docx
from Util import Util
if __name__ == "__main__":
    tools = Util('毕业论文2.docx')
    tools = Util("张军毕业论文.docx")
    # tools = Util("理工科-硕士-华中科技大学学位论文参考模板 - 副本.docx")
    # tools = Util("理工科-硕士-华中科技大学学位论文参考模板.docx")

    result = tools.getFullContext()
    for each in result:
        print(each)