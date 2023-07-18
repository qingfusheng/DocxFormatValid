# the first target is extract content from docx
from Util import Util
if __name__ == "__main__":
    tools = Util('四川大学论文模板.docx')
    # tools = Util("理工科-硕士-华中科技大学学位论文参考模板.docx")

    # print(tools.getFullContext())
    # tools.test_method()