# the first target is extract content from docx
from Util import Util
if __name__ == "__main__":
    tools = Util('肖露露毕业论文.docx')
    # tools = Util("胡广浩的毕业论文（定稿）v2.docx")
    # tools = Util("标注测试文档2.docx")
    # tools = Util("理工科-硕士-华中科技大学学位论文参考模板 - 副本.docx")
    # tools = Util("理工科-硕士-华中科技大学学位论文参考模板.docx")
    # tools = Util("硕士论文方晓亮 v2.docx")

    # result = tools.getFullContext()
    # for each in result:
    #     if each["type"] == "title":
    #         print(each)
    tools.test_method()
    