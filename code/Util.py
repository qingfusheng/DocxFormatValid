import zipfile
import os
import re
from lxml import etree
import sys
from xml.dom import minidom

project_dir = os.path.abspath("../../DocxFormatValid/")


class Util:
    def __init__(self, file_name):
        self.docx_dir = os.path.join(project_dir, "code", "./DocxFilter")
        self.workflow_dir = os.path.join(project_dir, "code", "./WorkFlowFilter")
        if not os.path.exists(self.docx_dir):
            os.mkdir(self.docx_dir)
        if not os.path.exists(self.workflow_dir):
            os.mkdir(self.workflow_dir)

        self.docx_filename = file_name
        self.new_docx_file = "new" + self.docx_filename
        self.unzip()
        self.doc = minidom.parse(os.path.join(self.workflow_dir, self.docx_filename, 'word', 'document.xml'))
        self.styles = minidom.parse(os.path.join(self.workflow_dir, self.docx_filename, 'word', 'styles.xml'))
        self.themes = minidom.parse(os.path.join(self.workflow_dir, self.docx_filename, 'word', 'theme', 'theme1.xml'))
        self.numbering = minidom.parse(os.path.join(self.workflow_dir, self.docx_filename, 'word', 'numbering.xml'))

        # self.content_text_list = []
        # self.getContent()

        return

    def unzip(self):
        f = zipfile.ZipFile(os.path.join(self.docx_dir, self.docx_filename))  # 打开需要修改的docx文件
        f.extractall(os.path.join(self.workflow_dir, self.docx_filename))  # 提取要修改的docx文件里的所有文件到workfolder文件夹
        f.close()
        return

    def getFullText(self, p) -> str:
        # 获取一个节点的所有文字
        text = ''
        for t in p.getElementsByTagName('w:t'):
            text += t.childNodes[0].data  # why childNode
        # 此处应该是将某些标号【1】,【2】等替换为空串
        # return text
        return re.sub(r'(【.*?】)', '', text)  # 匹配替换为选择的文本

    def getNodeText(self, node) -> str:
        text = ""
        for child_node in node.childNodes:
            if child_node.nodeType == child_node.TEXT_NODE:
                text += child_node.data
            elif child_node.nodeType == child_node.ELEMENT_NODE:
                text += self.getNodeText(child_node)
        return text

    def getContent(self):
        begin = False
        lst = []
        for p in self.doc.getElementsByTagName('w:p'):  # 获取所有的段落标签
            text = self.getFullText(p)  # 获取段落标签的文本
            if text.replace(' ', '') == '目录':
                # 当检测到目录文本时，开始获取content
                begin = True
                continue
            if begin:
                for t in lst:
                    if self.getFullText(p).replace(' ', '') in t.replace(' ', ''):
                        # 若当前获取到的文本在之前存在lst的列中，则将begin设为false，搭配后面的判断语句直接退出
                        begin = False
            if begin:
                lst.append(self.getFullText(p))
            if begin is False and len(lst) > 0:
                # 检测到正文标题和目录重合后退出，所以lst中只存的是目录
                break
        # self.content_text_list = lst
        return lst

    # @staticmethod
    # def levenshteinDistance(self, source, target):
    #     # Levenshtein 距离，又称编辑距离，指的是两个字符串之间，由一个转换成另一个所需的最少编辑操作次数
    #     def sub_cost(word1, word2, i, j):
    #         word1 = ' ' + word1
    #         word2 = ' ' + word2
    #         if word1[i] == word2[j]:
    #             return 0
    #         else:
    #             return 2
    #
    #     n = len(source)
    #     m = len(target)
    #     insert_cost = 1
    #     del_cost = 1
    #     lst = []
    #     tmp = 0
    #     tmp2 = 0
    #     for i in range(n + 1):
    #         if i == 0:
    #             for j in range(m + 1):
    #                 lst.append(j)
    #         else:
    #             for j in range(m + 1):
    #                 tmp2 = lst[j]
    #                 lst[j] = min(lst[j] + insert_cost, lst[j - 1] + del_cost if j - 1 >= 0 else i + insert_cost,
    #                              tmp + sub_cost(source, target, i, j) if j - 1 >= 0 else i - 1 + sub_cost(source,
    #                                                                                                       target, i, j))
    #                 tmp = tmp2
    #     return lst[m]

    def isTitle(self, p):
        # 判断是否为标题title
        # 问题是，type的含义是什么，type的取值是[0,1,2,3]，前面提到附录等一级标题的type为1，
        text = self.getFullText(p)
        type = 0
        # "25 ", "25.|25.|25，|25、|25或", "11.11", "1.1.11"
        reg = ["^[1-9][0-9]?\s*[a-zA-Z\u4E00-\u9FA5]{2,}", "^[1-9][0-9]*\.[1-9][0-9]*\s*[a-zA-Z\u4E00-\u9FA5]{2,}",
               "^[1-9][0-9]*\.[1-9][0-9]*\.[1-9][0-9]*\s*[a-zA-Z\u4E00-\u9FA5]{2,}"]
        for i in range(len(reg)):
            reg[i] = re.compile(reg[i])
        match = False
        if len(text) < 40:
            # 长度超过40的段落默认不是标题
            for i in ['附录', '致谢', '参考文献']:
                if i in text.replace(' ', '') and len(text) <= 20:
                    type = 0
                    match = True
                    break
            # i -> [3, 2, 1, 0]
            # 这里使用reverse的目的是,先匹配1.1.1， 再匹配1.1， 防止误将1.1.1匹配为一级或二级标签
            for i in reversed(range(len(reg))):
                # 1.1.11 -> type=3
                # 1.1 -> type=2
                # 1.|1.|1，|1、|1 -> type=1
                # '1 ' -> type=0
                if len(reg[i].findall(text)) == 1:
                    # print("match 1 , ", i)
                    # print(reg[i])
                    # print(text)
                    # print("-----------------")
                    #
                    type = i
                    match = True
                    break

                # if re.match(reg[i], text):
                #     print("match 1 , ", i)
                #     print(reg[i])
                #     print(text)
                #     print("-----------------")
                #     type = i
                #     match = True
                #     break
            # 若段落中有自动生成的标号，仿照大连理工的程序进行一系列判断，并利用与目录的最小编辑距离与辅助判断。
            # if p.getElementsByTagName('w:numPr'):
            #     print("***********************************************")
            #     numpr = p.getElementsByTagName('w:numPr')[0]
            #     ilvl = numpr.getElementsByTagName('w:ilvl')[0]
            #     numId = numpr.getElementsByTagName('w:numId')[0]
            #     for num in self.numbering.getElementsByTagName('w:num'):
            #         if num.getAttribute('w:numId') == numId.getAttribute('w:val'):
            #             abstract_numid = num.getElementsByTagName('w:abstractNumId')[0]
            #             for abstract_num in self.numbering.getElementsByTagName('w:abstractNum'):
            #                 if abstract_num.getAttribute('w:abstractNumId') == abstract_numid.getAttribute('w:val'):
            #                     for level in abstract_num.getElementsByTagName('w:lvl'):
            #                         if level.getAttribute('w:ilvl') == '0' and level.getElementsByTagName('w:lvlText')[
            #                             0].getAttribute('w:val') != '%1' and level.getElementsByTagName('w:lvlText')[
            #                             0].getAttribute('w:val') != '%1.':
            #                             match = False
            #                             break
            #                         elif level.getAttribute('w:ilvl') == '1' and \
            #                                 level.getElementsByTagName('w:lvlText')[0].getAttribute('w:val') != '%1.%2':
            #                             for s in self.content_text_list:
            #                                 if self.levenshteinDistance(text, re.sub('[0-9\.．\s]', '', s)) < len(
            #                                         re.sub('[0-9\.．\s]', '', s)) * 0.2:
            #                                     print("levenshteinDistance, True")
            #                                     match = True
            #                                     break
            #                             break
            #                         elif level.getAttribute('w:ilvl') == '2':
            #                             if level.getElementsByTagName('w:lvlText')[0].getAttribute(
            #                                     'w:val') != '%1.%2.%3':
            #                                 match = False
            #                                 break
            #                             match = True
            #
            #                     break
            #             break
            # 部分三级标题在目录中并没有出现,但是在正文中仍然祖籍为三级标题,为了正确匹配到这部分标题,我们需要将这段注释掉
            # if type == 1 or type == 0:  # 若匹配的样式是1. xxxx的样式，则查找目录是否有相同的内容，如果有，则判断为标题。因为有些字数少的正文的编号项目也有可能匹配成功。
            #     for t in self.content_text_list:
            #         if re.sub('[0-9\.．\s]', '', text) in re.sub('[0-9\.．\s]', '', t):
            #             break
            #     else:
            #         match = False

        # if type == 1:
        #     print(self.getFullText(p))
        # if match:
        #     print("match , ", type)
        return match, type

    @staticmethod
    def valid_distance(x, y):  # x -> y
        x_levels, y_levels = [int(x_item) for x_item in x.split(".")], [int(y_item) for y_item in y.split('.')]
        # result = False
        min_length = min(len(x_levels), len(y_levels))
        if len(x_levels) >= len(y_levels):
            # 1.2 -> 1.3
            # 1.1.1 -> 1.2|2
            if x_levels[min_length - 1] + 1 == y_levels[min_length - 1]:
                return True
            else:
                return False
        else:
            # 1.1 -> 1.1.1
            if len(x_levels) + 1 != len(y_levels):
                # 1 -> 1.1.1 判定为错，只能是逐级递增，比如1->1.1，1。 1->1.1.1
                return False
            result = True
            for i in range(min_length):
                if x_levels[i] != y_levels[i]:
                    result = False
            if y_levels[-1] != 1:
                result = False
            return result

    def getTitle(self):
        content_set = set()
        title_list = []
        title_began = False
        old_order = ""
        for p in self.doc.getElementsByTagName('w:p'):
            match, type = self.isTitle(p)
            if match:
                p_content = re.sub(re.compile(r'（\d+）'), "", self.getFullText(p))
                # print(p_content)
                if not title_began:
                    if p_content not in content_set:
                        content_set.add(p_content)
                    else:
                        title_began = True
                        # title_list.append((type, p_content))
                reg = re.compile(r'^\d+(\.\d+)*')
                reg_match = reg.search(p_content)
                if reg_match:
                    order = reg_match.group()
                else:
                    order = '0'
                if old_order != "" and order != '0':
                    valid_result = self.valid_distance(old_order, order)
                    if not valid_result:
                        continue
                old_order = order
                if title_began:
                    title_list.append((type, p_content, order))
        # for each in title_list:
        #     print(each)
        return title_list

    def getFullContext(self):
        content_over = False
        text_begin = False
        titleList = self.getTitle()
        pure_title = [each[1] for each in titleList]
        contentList = self.getContent()
        index = 0
        result = []
        for p in self.doc.getElementsByTagName('w:p'):
            if not text_begin and not content_over:
                if self.getFullText(p) == contentList[-1]:
                    content_over = True
                # else:
                #     continue
            if content_over:
                if self.getFullText(p) == pure_title[0]:
                    text_begin = True
                    content_over = False
                # else:
                #     continue
            if text_begin:
                # print(index)
                # 检测为title
                if index < len(titleList):
                    if self.getFullText(p) == titleList[index][1]:
                        result.append({
                            "type": "title",
                            "level": titleList[index][0],
                            "content": self.getFullText(p)
                        })
                        index += 1
                    else:
                        result.append({
                            "type": "text",
                            "level": -1,
                            "content": self.getFullText(p)
                        })
                else:
                    result.append({
                        "type": "text",
                        "level": -1,
                        "content": self.getFullText(p)
                    })
        for each in result:
            print(each)
        return




    def test_method(self):
        # document = self.doc.childNodes[0]
        # body = document.childNodes[0]
        # first_paragraph = body.childNodes[0]
        # print(len(body.childNodes))
        # print(first_paragraph.tagName)
        # print(len(first_paragraph.childNodes))
        # for each in first_paragraph.childNodes:
        #     print(each.tagName)
        # print()
        #
        # return
        self.getTitle()
        return
        for p in self.doc.getElementsByTagName('w:p'):
            # print("-----------------")
            # print(self.getFullText(p))
            temp_list = []
            if self.isTitle(p):
                temp_list.append(self.getFullText())
                print(self.getFullText(p))
                print("----------------------------------------------")
