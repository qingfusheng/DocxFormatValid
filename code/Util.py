import json
import xml.dom.minidom
import zipfile
import os
import re
from lxml import etree
import sys
from xml.dom import minidom
from xml.dom.minidom import parseString

project_dir = os.path.abspath("../../DocxFormatValid/")


class Style:
    def __init__(self):
        # 注意默认属性的影响
        self.font_ascii = "Times New Roman"
        self.font_eastAsia = ""
        self.font_sz = ""
        self.font_szCs = ""
        self.font_b = ""
        self.font_bCs = "0"
        self.font_i = "0"
        self.font_u = "0"
        self.font_color = ""
        self.font_shd = ""
        self.highlight = "0"

        self.jc = ""
        self.ind = "0"  # w:ind
        self.spacing = "240"

    def __str__(self):
        return json.dumps({
            "英文字体：": self.font_ascii,
            "中文字体：": self.font_eastAsia,
            "中文字号：": self.font_sz,
            "复杂字号：": self.font_szCs,
            "是否粗体：": self.font_b,
            "是否粗体2：": self.font_bCs,
            "是否斜体：": self.font_i,
            "是否下划线：": self.font_u,
            "颜色：": self.font_color,
            "背景颜色：": self.font_shd,
            "是否高亮：": self.highlight,
            "对齐方式": self.jc,
            "缩进设置：": self.ind,
            "行距设置": self.spacing
        }, ensure_ascii=False)


def get_pound_of_font_sz(font_sz: str):
    WordFontSizeDict = {
        "初号": "84",
        "小初": "72",
        "一号": "52",
        "小一": "48",
        "二号": "44",
        "小二": "36",
        "三号": "32",
        "小三": "30",
        "四号": "28",
        "小四": "24",
        "五号": "21",
        "小五": "18",
        "六号": "15",
        "小六": "13",
        "七号": "11",
        "八号": "10"
    }
    num_reg = re.compile(r'\d+')
    if num_reg.findall(font_sz):
        return str(2 * int(font_sz))
    if font_sz in WordFontSizeDict:
        return WordFontSizeDict[font_sz]
    else:
        raise Exception("请检查字号是否输入正确")


default_style_dict = {
    "font_ascii": "Times New Roman",
    "font_eastAsia": "宋体",
    "font_sz": get_pound_of_font_sz("小四"),
    "font_szCs": "24",
    "font_b": "0",
    "font_bCs": "0",
    "font_i": "0",
    "font_u": "0",
    "font_color": "auto",
    "font_shd": "000000",
    "highlight": "0",
    "jc": "left",
    "ind": "0",  # w:ind
    "spacing": "240"
}
default_style = Style()
for key, value in default_style_dict.items():
    setattr(default_style, key, value)


class Util:
    def __init__(self, file_name):
        self.comment_dict = None
        self.style_dict = None
        self.docx_dir = os.path.join(project_dir, "code", "./DocxFilter")
        self.workflow_dir = os.path.join(project_dir, "code", "./WorkFlowFilter")
        if not os.path.exists(self.docx_dir):
            os.mkdir(self.docx_dir)
        if not os.path.exists(self.workflow_dir):
            os.mkdir(self.workflow_dir)

        self.docx_filename = file_name
        self.new_docx_file = "new_" + self.docx_filename
        self.error_log_text = ""

        self.unzip()
        self.doc: xml.dom.minidom.Element = minidom.parse(
            os.path.join(self.workflow_dir, self.docx_filename, 'word', 'document.xml'))
        self.docx_body: xml.dom.minidom.Element = self.doc.childNodes[0].childNodes[0]
        self.styles: xml.dom.minidom.Element = minidom.parse(
            os.path.join(self.workflow_dir, self.docx_filename, 'word', 'styles.xml'))
        self.themes: xml.dom.minidom.Element = minidom.parse(
            os.path.join(self.workflow_dir, self.docx_filename, 'word', 'theme', 'theme1.xml'))
        self.numbering: xml.dom.minidom.Element = minidom.parse(
            os.path.join(self.workflow_dir, self.docx_filename, 'word', 'numbering.xml')) if os.path.exists(
            os.path.join(self.workflow_dir, self.docx_filename, 'word', 'numbering.xml')) else minidom.Document()
        self.comments: xml.dom.minidom.Element = minidom.parse(
            os.path.join(self.workflow_dir, self.docx_filename, 'word', 'comments.xml')) if os.path.exists(
            os.path.join(self.workflow_dir, self.docx_filename, 'word', 'comments.xml')) else minidom.Document()
        self.create_style_xml_index_by_styleId()
        self.create_comment_xml_index_by_commentId()
        print("comment_dict:", self.style_dict)
        return

    def unzip(self):
        f = zipfile.ZipFile(os.path.join(self.docx_dir, self.docx_filename))  # 打开需要修改的docx文件
        f.extractall(os.path.join(self.workflow_dir, self.docx_filename))  # 提取要修改的docx文件里的所有文件到workfolder文件夹
        f.close()
        return

    @staticmethod
    def create_empty_dom(self):
        dom = minidom.Document()
        return dom

    def create_comment_xml_index_by_commentId(self):
        self.comment_dict = {}
        comment_elements = self.comments.getElementsByTagName("w:comment")
        print("The Count of comments: ", len(comment_elements))
        for each_comment in comment_elements:
            comment_id = each_comment.getAttribute("w:id")
            comment_content = self.getFullText(each_comment)
            self.comment_dict[comment_id] = comment_content
        return

    @staticmethod
    def getFullText(p) -> str:
        # 获取一个节点的所有文字
        text = ''
        for t in p.getElementsByTagName('w:t'):
            text += t.childNodes[0].data  # why childNode
        # 此处应该是将某些标号【1】,【2】等替换为空串
        # return text
        return re.sub(r'(【.*?】)', '', text)  # 匹配替换为选择的文本

    def getFullTextAndCommentOfPara(self, paragraph):
        text = ""
        comments = []  # [(comment_id, comment_refer, comment_content]
        com_range_start, com_range_end = False, False
        comment_begin_index = 0
        comment_id = ""
        for child_node in paragraph.childNodes:
            if child_node.tagName == "w:commentRangeStart":
                comment_id = child_node.getAttribute("w:id")
                com_range_start = True
                com_range_end = False
                comment_begin_index = len(text)
                continue
            if child_node.tagName == "w:commentRangeEnd":
                # comment_end_index = len(text)
                comments.append([comment_id, (comment_begin_index, len(text)), self.comment_dict[comment_id]])
                com_range_start = False
                com_range_end = True
                continue
            if com_range_start and not com_range_end:
                text += self.getFullText(child_node)
            if not com_range_start and com_range_end:
                text += self.getFullText(child_node)
                pass
        print(text)
        print(comments)
        return text, comments

    def getNodeText(self, node) -> str:
        text = ""
        for child_node in node.childNodes:
            if child_node.nodeType == child_node.TEXT_NODE:
                text += child_node.data
            elif child_node.nodeType == child_node.ELEMENT_NODE:
                text += self.getNodeText(child_node)
        return text

    def create_style_xml_index_by_styleId(self):
        self.style_dict = {}
        for each_style in self.styles.getElementsByTagName("w:style"):
            style_id = each_style.getAttribute("w:styleId")
            self.style_dict[style_id] = each_style
        return

    def getPointsOfStyle(self, styleId):
        points = 0
        ilvl = ""
        each_style = self.style_dict[styleId]
        style_type = each_style.getAttribute("w:type")
        if style_type != "paragraph":
            return points, ilvl
        style_id = each_style.getAttribute("w:styleId")
        style_based_on = None if not each_style.getElementsByTagName("w:basedOn") else \
            each_style.getElementsByTagName("w:basedOn")[0].getAttribute("w:val")
        char_link_style = each_style.getElementsByTagName("w:link")[0] if each_style.getElementsByTagName(
            "w:link") else None
        # 字体大小，字体是否为粗体，是否存在自动编号

        # 这里int(ilvl)+1代表标题的级别，当ilvl为-1时表示没有自动编号
        paragraph_property = each_style.getElementsByTagName("w:pPr")[0] if each_style.getElementsByTagName(
            "w:pPr") else None
        if paragraph_property:
            num_pr = paragraph_property.getElementsByTagName("w:numPr")[
                0] if paragraph_property.getElementsByTagName("w:numPr") else None
            if num_pr:
                ilvl = num_pr.getElementsByTagName("w:ilvl")[0].getAttribute(
                    "w:val") if num_pr.getElementsByTagName("w:ilvl") else '0'
            else:
                ilvl = '-1'
        else:
            ilvl = '-1'
        run_property = each_style.getElementsByTagName("w:rPr")[0] if each_style.getElementsByTagName(
            "w:rPr") else None
        if run_property:
            rFont = run_property.getElementsByTagName("w:rFonts")[0].getAttribute(
                "w:eastAsia") if run_property.getElementsByTagName("w:rFonts") else None
            is_b = bool(run_property.getElementsByTagName("w:b")) or bool(run_property.getElementsByTagName("w:bCs"))
            size = run_property.getElementsByTagName("w:sz")[0].getAttribute(
                "w:val") if run_property.getElementsByTagName("w:sz") else '24'
            sizeCs = run_property.getElementsByTagName("w:szCs")[0].getAttribute(
                "w:val") if run_property.getElementsByTagName("w:szCs") else '24'
            if rFont == "黑体":
                points += 1
            if is_b:
                points += 1
            if int(size) > 24 or int(sizeCs) > 24:
                points += 1
        else:
            rFont = ""
            is_b = False
            size = '0'
            sizeCs = '0'
            points += 0
        points += int(ilvl)
        # print(ilvl, rFont, is_b, size, sizeCs)
        return points, ilvl

    def isTitle(self, p: xml.dom.minidom.Element):
        # 判断是否为标题title
        # 问题是，type的含义是什么，type的取值是[0,1,2,3]，前面提到附录等一级标题的type为1，
        text = self.getFullText(p)
        type = 0
        # "25 ", "25.|25.|25，|25、|25或", "11.11", "1.1.11"
        reg = ["^[1-9][0-9]?\s*[a-zA-Z\u4E00-\u9FA5 ]{2,}", "^[1-9][0-9]*\.[1-9][0-9]*\s*[a-zA-Z\u4E00-\u9FA5]{2,}",
               "^[1-9][0-9]*\.[1-9][0-9]*\.[1-9][0-9]*\s*[a-zA-Z\u4E00-\u9FA5]{2,}"]
        for i in range(len(reg)):
            reg[i] = re.compile(reg[i])
        match = False
        if len(text.strip()) == 0:
            return False, 0
        if len(text) < 40:
            # 长度超过40的段落默认不是标题
            for i in ['附录', '致谢', '参考文献']:
                if i in text.replace(' ', '') and len(text) <= 20:
                    type = 1
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
                    type = i + 1
                    match = True
                    break
            # print(self.getFullText(p), match, type)
            if not p.getElementsByTagName("w:pPr"):
                match = False
            else:
                points = 0
                pPr = p.getElementsByTagName("w:pPr")[0]
                rFont = pPr.getElementsByTagName("w:rFont")[0] if pPr.getElementsByTagName("w:rFont") else None
                bold = bool(pPr.getElementsByTagName("w:b") or pPr.getElementsByTagName("w:bCs"))
                sz = pPr.getElementsByTagName("w:sz")[0].getAttribute("w:val") if pPr.getElementsByTagName(
                    "w:sz") else "24"
                szCs = pPr.getElementsByTagName("w:szCs")[0].getAttribute("w:val") if pPr.getElementsByTagName(
                    "w:szCs") else "24"
                if rFont == "黑体":
                    points += 1
                if bold:
                    points += 1
                if int(sz) > 24 or int(szCs) > 24:
                    points += 1
                if points > 1:
                    match = True

                # if re.match(reg[i], text):
                #     print("match 1 , ", i)
                #     print(reg[i])
                #     print(text)
                #     print("-----------------")
                #     type = i
                #     match = True
                #     break

            if p.getElementsByTagName('w:pPr'):
                pPr = p.getElementsByTagName('w:pPr')[0]
                if pPr.getElementsByTagName("w:pStyle"):
                    pStyle = pPr.getElementsByTagName("w:pStyle")[0]
                    pStyleId = pStyle.getAttributeNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                                                     "val")
                    # print("pStyle:", pStyleId)
                    points, ilvl = self.getPointsOfStyle(pStyleId)
                    if points >= 2 and int(ilvl) >= 0:
                        # print("matched Successfully")
                        match = True
                        type = int(ilvl) + 1

        return match, type

    @staticmethod
    def valid_distance(x, y):  # x -> y
        x_levels, y_levels = [int(x_item) for x_item in x.split(".")], [int(y_item) for y_item in y.split('.')]
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

    def isTabTitle(self, p) -> (bool, int):
        text = self.getFullText(p)
        if len(text) < 30 and text.find(',') == -1 and text.find('，') == -1 and re.search('^[表][0-9]',
                                                                                          re.sub('\s', '', text)):
            # print(text)
            reg = re.compile(r'\d+([\.-]\d+)*')
            reg_match = reg.search(text)
            # print(reg_match)
            _order = reg_match.group()
            if '.' in _order:
                sig = '.'
            elif '-' in _order:
                sig = '-'
            type = len(_order.split(sig))
            return True, type
        return False, 0

    def isPicTitle(self, p) -> (bool, int):
        text = self.getFullText(p)
        if len(text) < 30 and text.find(',') == -1 and text.find('，') == -1 and re.search('^[图][0-9]',
                                                                                          re.sub('\s', '', text)):
            # print(text)
            reg = re.compile(r'\d+([\.-]\d+)*')
            reg_match = reg.search(text)
            # print(reg_match)
            _order = reg_match.group()
            if '.' in _order:
                sig = '.'
            elif '-' in _order:
                sig = '-'
            type = len(_order.split(sig))
            return True, type
        return False, 0

    def getFullContext(self):
        # 目录结束
        content_began = False
        content_over = False
        if not self.doc.getElementsByTagName("w:sdt"):
            content_began = False
            content_over = True
        # 正文文本开始
        # text_begin = False
        # title_list = self.getTitle()
        # pure_title = [each[1] for each in title_list]
        # content_list = self.getContent()
        # print("初始狀態：content_began為", content_began, "content_over為", content_over)
        title_index = 0
        result = []
        for p in self.doc.getElementsByTagName('w:p'):
            # print("初始狀態：content_began為", content_began, "content_over為", content_over)
            # 过滤到表格中的文本内容
            try:
                if p.parentNode.parentNode.parentNode.tagName == "w:tbl":
                    # print("识别到表格，已跳过")
                    continue
            except AttributeError as error:
                print("", end="")

            # 過濾空的paragraph
            if self.getFullText(p) == "":
                continue

            # print(self.getFullText(p))
            # 检测到目录的识别并未开始且并未结束
            if not content_began and not content_over:
                if p.getElementsByTagName('w:pPr'):
                    pPr = p.getElementsByTagName('w:pPr')[0]
                    if pPr.getElementsByTagName("w:pStyle"):
                        pStyle = pPr.getElementsByTagName("w:pStyle")[0]
                        pStyle_value = pStyle.getAttributeNS(
                            "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "val")
                        if "TOC" in pStyle_value or "toc" in pStyle_value:
                            content_began = True
                            # print("狀態變化1：content_began為", content_began, "content_over為", content_over)
                            continue
            # 检测到目录的识别已开始但未结束
            if content_began and not content_over:
                if p.getElementsByTagName('w:pPr'):
                    pPr = p.getElementsByTagName('w:pPr')[0]
                    if pPr.getElementsByTagName("w:pStyle"):
                        pStyle = pPr.getElementsByTagName("w:pStyle")[0]
                        pStyle_value = pStyle.getAttributeNS(
                            "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "val")
                        if "TOC" not in pStyle_value and 'toc' not in pStyle_value:
                            print(pStyle_value)
                            content_began = False
                            content_over = True
                            # print("狀態變化2：content_began為", content_began, "content_over為", content_over)
                    else:
                        content_began = False
                        content_over = True
                        # print("狀態變化3：content_began為", content_began, "content_over為", content_over)
                else:
                    # 检测到目录识别已经结束
                    content_began = False
                    content_over = True
                    # print("狀態變化4：content_began為", content_began, "content_over為", content_over)

                # if self.getFullText(p) == content_list[-1]:
                #     # 如果检测到当前的段落是目录的最后一个元素，就表示目录结束，并进入下一个分支
                #     content_over = True

            # if not content_began and content_over:
            #     # 检测到第一个标题（标题1,2,3）
            #     if self.getFullText(p) == pure_title[0]:
            #         text_begin = True
            #         content_over = False

            # 其实这里只需要判断content_over，content_began可能会有误判断的可能性
            if not content_began and content_over:
                is_title_result = self.isTitle(p)
                is_tab_title_result = self.isTabTitle(p)
                is_pic_title_result = self.isPicTitle(p)
                if is_title_result[0]:
                    result.append({
                        "type": "title",
                        "level": is_title_result[1],
                        "content": self.getFullText(p)
                    })
                elif is_tab_title_result[0]:
                    result.append({
                        "type": "table_title",
                        "level": is_tab_title_result[1],
                        "content": self.getFullText(p)
                    })
                elif is_pic_title_result[0]:
                    result.append({
                        "type": "pic_title",
                        "level": is_pic_title_result[1],
                        "content": self.getFullText(p)
                    })
                else:
                    result.append({
                        "type": "text",
                        "level": "-1",
                        "content": self.getFullText(p)
                    })

                # 检测为title
                # if title_index < len(title_list):
                #     if self.getFullText(p) == title_list[title_index][1]:
                #         result.append({
                #             "type": "title" if title_list[title_index][0] <= 3 else "pic_title" if
                #             title_list[title_index][0] == 4 else "tab_title",
                #             "level": title_list[title_index][0],
                #             "content": self.getFullText(p)
                #         })
                #         title_index += 1
                #         continue
                #     else:
                #         result.append({
                #             "type": "text",
                #             "level": -1,
                #             "content": self.getFullText(p)
                #         })
                # else:
                #     result.append({
                #         "type": "text",
                #         "level": -1,
                #         "content": self.getFullText(p)
                #     })
        # for each in result:
        #     print(each)
        return result

    def get_comment_method_2(self):
        # method 2 (适合只提取comments而不提取其他内容)
        comments = []  # [(comment_id, comment_refer, comment_content]
        for paragraph in self.doc.getElementsByTagName('w:p'):
            paragraph_txt = paragraph.toxml()
            reg = re.compile(r'<w:commentRangeStart w:id="(\d)+"/>(.*?)<w:commentRangeEnd w:id="(\d+)"/>')
            for elem in reg.findall(paragraph_txt):
                if elem[0] != elem[2]:
                    raise Exception("正则匹配失败，前后comment_id不匹配")
                comment_id = elem[0]
                try:
                    refer_run_dom = minidom.parseString(
                        "<root xmlns:w='http://example.com/namespace'>" + elem[1] + "</root>")
                except Exception as error:
                    refer_run_dom = minidom.parseString("<root xmlns:w='http://example.com/namespace'></root>")
                    print(error)
                refer_run_text = self.getFullText(refer_run_dom)
                comments.append([comment_id, refer_run_text, self.comment_dict[comment_id]])
                # {"comment_id": comment_id,"refer_text": refer_run_text,"comment_content": self.comment_dict[id]}
        print(comments)
        return comments

    def get_comment(self):
        comments = []  # [(标注的ID, 正文中标注引用的内容, 标注的内容]
        for paragraph in self.doc.getElementsByTagName('w:p'):
            com_text, com_list = self.getFullTextAndCommentOfPara(paragraph)
            for each_comment in com_list:
                comments.append([each_comment[0], com_text[each_comment[1][0]: each_comment[1][1]], each_comment[2]])
        print(comments)
        return

    @staticmethod
    def print_log(content: str):
        print(content)
        return

    def print_error(self, content: str, para: xml.dom.minidom.Element):
        print(content)
        return

    def get_style_from_styleId(self, style_id: str):
        # print("getting style from styleId:", style_id)
        style: xml.dom.minidom.Element = self.style_dict[style_id]
        format_style = Style()
        if style.getElementsByTagName("w:rPr"):
            run_property = style.getElementsByTagName("w:rPr")[0]
            if run_property.getElementsByTagName("w:rFonts"):
                if run_property.getElementsByTagName("w:rFonts")[0].getAttribute("w:eastAsia"):
                    format_style.font_eastAsia = run_property.getElementsByTagName("w:rFonts")[0].getAttribute(
                        "w:eastAsia")
                if run_property.getElementsByTagName("w:rFonts")[0].getAttribute("w:ascii"):
                    format_style.font_ascii = run_property.getElementsByTagName("w:rFonts")[0].getAttribute("w:ascii")
            if run_property.getElementsByTagName("w:sz"):
                format_style.font_sz = run_property.getElementsByTagName("w:sz")[0].getAttribute("w:val")
            if run_property.getElementsByTagName("w:szCs"):
                format_style.font_szCs = run_property.getElementsByTagName("w:szCs")[0].getAttribute("w:val")
            if run_property.getElementsByTagName("w:color"):
                format_style.font_color = run_property.getElementsByTagName("w:color")[0].getAttribute("w:val")
            if run_property.getElementsByTagName("w:b"):
                format_style.font_b = "1"
            if run_property.getElementsByTagName("w:bCs"):
                format_style.font_bCs = "1"

            if run_property.getElementsByTagName("w:i"):
                format_style.font_i = "1"
            if run_property.getElementsByTagName("w:u"):
                format_style.font_u = "1"

        if style.getElementsByTagName("w:pPr"):
            para_property = style.getElementsByTagName("w:pPr")[0]
            if para_property.getElementsByTagName("w:jc"):
                format_style.jc = para_property.getElementsByTagName("w:jc")[0].getAttribute("w:val")
            if para_property.getElementsByTagName("w:spacing"):
                format_style.spacing = para_property.getElementsByTagName("w:spacing")[0].getAttribute("w:line")
            if para_property.getElementsByTagName("w:ind"):
                format_style.spacing = para_property.getElementsByTagName("w:ind")[0].getAttribute("w:left")
        # print(format_style)
        return format_style

    def get_style_of_para(self, para: xml.dom.minidom.Element):
        style_class = Style()
        if para.getElementsByTagName("w:pPr"):
            para_property = para.getElementsByTagName("w:pPr")[0]
            if para_property.getElementsByTagName("w:jc"):
                style_class.js = para_property.getElementsByTagName("w:jc")[0].getAttribute("w:val")
            if para_property.getElementsByTagName("w:spacing"):
                style_class.spacing = para_property.getElementsByTagName("w:spacing")[0].getAttribute("w:line")
            if para_property.getElementsByTagName("w:rPr"):
                StyleOfrPr = self.get_style_of_run(para_property.getElementsByTagName("w:rPr")[0])
                style_class.font_eastAsia = StyleOfrPr.font_eastAsia if StyleOfrPr.font_eastAsia else style_class.font_eastAsia
                style_class.font_ascii = StyleOfrPr.font_ascii if StyleOfrPr.font_ascii else style_class.font_ascii
                style_class.font_b = StyleOfrPr.font_b if StyleOfrPr.font_b else style_class.font_b
                style_class.font_bCs = StyleOfrPr.font_bCs if StyleOfrPr.font_bCs else style_class.font_bCs
                style_class.font_sz = StyleOfrPr.font_sz if StyleOfrPr.font_sz else style_class.font_sz
                style_class.font_szCs = StyleOfrPr.font_szCs if StyleOfrPr.font_szCs else style_class.font_szCs
                style_class.font_i = StyleOfrPr.font_i if StyleOfrPr.font_i else style_class.font_i
                style_class.font_u = StyleOfrPr.font_u if StyleOfrPr.font_u else style_class.font_u
                style_class.font_color = StyleOfrPr.font_color if StyleOfrPr.font_color else style_class.font_color

            if para_property.getElementsByTagName("w:pStyle"):
                pStyle = para_property.getElementsByTagName("w:pStyle")[0]
                style_id = pStyle.getAttributeNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "val")
                # 首先检测是否有style_id,以及获取style_id的style属性
                style_class_of_id = self.get_style_from_styleId(style_id)
                style_name_list = ["font_ascii", "font_eastAsia", "font_sz", "font_szCs", "font_b", "font_bCs",
                                   "font_i", "font_u", "font_color", "font_shd", "highlight", "jc", "ind", "spacing"]
                null_style_class = Style()
                for style_name in style_name_list:
                    style_attr = getattr(style_class_of_id, style_name)
                    if style_attr != getattr(null_style_class, style_name):
                        setattr(style_class, style_name, style_attr)
        # print("getting style from parahraph")
        # print(style_class)
        return style_class

    def get_style_of_run(self, run: xml.dom.minidom.Element):
        # run的属性有几种，继承自pPr的rPr属性，内嵌的rPr属性，以及style中的rPr属性
        # 先检测w:rStyle，再检测w:rPr
        style_class = Style()
        if run.getElementsByTagName("w:rPr"):
            run_property = run.getElementsByTagName("w:rPr")[0]
            if run_property.getElementsByTagName("w:rStyle"):
                rStyle = run_property.getElementsByTagName("w:rStyle")[0]
                style_id = rStyle.getAttributeNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "val")
                style_class = self.get_style_from_styleId(style_id)

            if run_property.getElementsByTagName("w:rFonts"):
                if run_property.getElementsByTagName("w:rFonts")[0].getAttribute("w:eastAsia"):
                    style_class.font_eastAsia = run_property.getElementsByTagName("w:rFonts")[0].getAttribute(
                        "w:eastAsia")
                style_class.font_ascii = run_property.getElementsByTagName("w:rFonts")[0].getAttribute("w:ascii")
            if run_property.getElementsByTagName("w:b"):
                style_class.font_b = "1"
            if run_property.getElementsByTagName("w:bCs"):
                style_class.font_bCs = "1"
            if run_property.getElementsByTagName("w:sz"):
                style_class.font_sz = run_property.getElementsByTagName("w:sz")[0].getAttribute("w:val")
            if run_property.getElementsByTagName("w:szCs"):
                style_class.font_szCs = run_property.getElementsByTagName("w:szCs")[0].getAttribute("w:val")
            if run_property.getElementsByTagName("w:i"):
                style_class.font_i = "1"
            if run_property.getElementsByTagName("w:u"):
                style_class.font_u = "1"
            if run_property.getElementsByTagName("w:color"):
                style_class.font_color = run_property.getElementsByTagName("w:color")[0].getAttribute("w:val")
        return style_class

    def DetectPaper(self):
        def insert_index(temp_index: int):
            print("temp_index", temp_index)
            if self.getFullText(docx_body_childs[temp_index]):
                IndexList.append(temp_index)
            else:
                return

        CoverIndex = 0
        IndexList = [CoverIndex]
        # self.docx_body = self.doc.childNodes[0].childNodes[0]
        docx_body_childs = self.docx_body.childNodes
        title_list = ["独创性声明", "摘要", "Abstract", "目录", "绪论", "致谢", "参考文献",
                      "附录1  攻读硕士学位期间取得的学术成果", "附录2  其它附录"]
        if not self.doc.getElementsByTagName("w:sdt"):
            print("输入DOCX文档没有目录，请添加目录后重试1")
        for index in range(len(docx_body_childs)):
            paragraph_node: xml.dom.minidom.Element
            paragraph_node = docx_body_childs[index]
            if paragraph_node.nodeName == "w:sdt":
                for child_elem in paragraph_node.getElementsByTagName("w:p"):
                    if self.getFullText(child_elem).replace(" ", "") in title_list:
                        if IndexList[-1] != index:
                            insert_index(index)
                    else:
                        print("输入DOCX文档没有目录，请添加目录后重试2")
                        # exit()
                continue
            if paragraph_node.nodeName != "w:p":
                # 前面已经排除了是目录的情况
                continue

            """判断分页符"""
            # if "<w:br w:type=\"page\"/>" in paragraph_node.toxml():
            #     IndexList.append(index+1)
            #     continue
            if paragraph_node.getElementsByTagName("w:br"):
                if not self.getFullText(paragraph_node):
                    insert_index(index + 1)
                    continue
                else:
                    child_length = len(paragraph_node.childNodes)
                    for elem_index in range(child_length):
                        if paragraph_node.childNodes[elem_index].getElementsByTagName("w:br"):
                            if elem_index < child_length // 2:
                                insert_index(index)
                                continue
                            else:
                                insert_index(index + 1)
                                continue

            """判断分节符"""
            if paragraph_node.getElementsByTagName("w:sectPr"):
                if IndexList[-1] != index + 1:
                    insert_index(index + 1)
                    continue
            """判断是否含有标题标签"""
            if self.getFullText(paragraph_node).replace(" ", "") in title_list:
                if IndexList[-1] != index:
                    insert_index(index)
                    continue
        # insert_index(len(docx_body_childs))
        IndexList.append(len(docx_body_childs))
        print(IndexList)
        self.DetectCover(IndexList, 0, 2)
        self.DetectCopyright(IndexList[2], IndexList[3] - 1)
        self.DetectAbstract(IndexList, 3, 5)
        self.DetectCatalogue(IndexList[5], IndexList[6])
        text_end_index = 6
        for i in range(6, len(IndexList)):
            if self.getFullText(self.docx_body.childNodes[IndexList[i]]).replace(" ", "") == "致谢":
                text_end_index = i
                break
        self.DetectText(IndexList, 6, text_end_index)
        # self.DetectAcknowledge(IndexList[text_end_index], IndexList[text_end_index + 1])
        # self.DetectReference(IndexList[text_end_index + 1], IndexList[text_end_index + 2])
        # self.DetectAppendixes(IndexList, text_end_index + 2, len(IndexList))

    def DetectCover(self, index_list: list, index_start: int, index_end: int):
        self.print_log("Detecting Paper CCover")
        self.DetectChineseCover(index_list[index_start], index_list[index_start + 1])
        self.DetectEnglishCover(index_list[index_start + 1], index_list[index_end])
        pass

    def checkStyle(self, para, StyleDict):
        paragraph_style = Style()
        # print("-------------------------!!!!!!!!!!!!!!!!!!!!!!!!!!!!-------------------")
        print(self.getFullText(para))
        # print(StyleDict)
        if para.tagName == "w:p":
            paragraph_style = self.get_style_of_para(para)

        if para.tagName == "w:r":
            while not para.parentNode.tagName == "w:p":
                para = para.parentNode
            paragraph_style = self.get_style_of_para(para.parentNode)
        for elem in para.getElementsByTagName("w:r"):
            if elem.nodeName == "w:r":
                run_style = self.get_style_of_run(elem)
                # print(self.getFullText(elem))
                # print("para_style:", paragraph_style)
                # print("run_style:", run_style)
                style_class = Style()
                style_name_list = ["font_ascii", "font_eastAsia", "font_sz", "font_szCs", "font_b", "font_bCs",
                                   "font_i", "font_u", "font_color", "font_shd", "highlight", "jc", "ind", "spacing"]
                for style_name in style_name_list:
                    style_attr: str = getattr(style_class, style_name)
                    if getattr(paragraph_style, style_name) != "":
                        style_attr = getattr(paragraph_style, style_name)
                    if getattr(run_style, style_name) != "":
                        style_attr = getattr(run_style, style_name)
                    if style_attr == "":
                        # print(style_name, style_attr, end=" ")
                        style_attr = getattr(default_style, style_name)
                        # print(style_attr)
                    # print(style_name, style_attr)
                    setattr(style_class, style_name, style_attr)
                # print(paragraph_style)
                # print(run_style)
                # print(default_style)
                # print(style_class)
                # print("实际样式：", style_class, "\n规定样式：", StyleDict)
                if style_class.font_b != StyleDict["font_b"]:
                    self.print_error("粗体使用错误", elem)
                # if style_class.font_bCs != "1":
                #     self.print_error("请使用粗体", elem)
                if style_class.font_sz != get_pound_of_font_sz(StyleDict["font_sz"]):
                    self.print_error("字体大小错误，应当为" + StyleDict["font_sz"] + "，而实际上是" + style_class.font_sz,
                                     elem)
                if style_class.font_eastAsia != StyleDict["font_eastAsia"]:
                    self.print_error("中文字体使用错误，应为" + StyleDict[
                        "font_eastAsia"] + "，而实际上是" + style_class.font_eastAsia, elem)
                if StyleDict["font_ascii"] != "":
                    if style_class.font_ascii not in [StyleDict["font_ascii"], ""]:
                        self.print_error(
                            "英文字体使用错误，应为" + StyleDict["font_ascii"] + "，而实际上是" + style_class.font_ascii,
                            elem)
                if style_class.font_color not in ["", "000000", "auto"]:
                    self.print_error("颜色使用错误，应为" + "黑色", elem)
                if "jc" in StyleDict:
                    if style_class.jc != StyleDict["jc"]:
                        self.print_error(
                            "对齐方式使用错误，应为" + StyleDict["jc"] + "，而实际上是" + style_class.jc,
                            elem)
                # print("----over----")

    def DetectChineseCover(self, para_index_begin: int, para_index_end: int):
        def check_student_number(para: xml.dom.minidom.Element):
            student_number_reg = re.compile(r'M20\d{7}')
            if not student_number_reg.findall(self.getFullText(para).replace(" ", "")):
                self.print_error("学号不存在或位数错误", para)
            return

        def check_school_number(para: xml.dom.minidom.Element):
            if "10487" not in self.getFullText(para).replace(" ", ""):
                self.print_error("学校代码不存在或错误", para)
            return

        def check_paper_type_style(para: xml.dom.minidom.Element):
            print("----------checking type info----------")
            StyleDict = {
                "font_b": "1",
                "font_sz": "45",
                "font_eastAsia": "华文中宋",
                "font_ascii": "Times New Roman",
                "font_color": "000000"
            }
            self.checkStyle(para, StyleDict)
            print("--------------------over--------------------")

        def check_paper_title_style(para: xml.dom.minidom.Element):
            print("----------checking title info----------")
            StyleDict = {
                "font_b": "1",
                "font_sz": "一号",
                "font_eastAsia": "宋体",
                "font_ascii": "Times New Roman",
                "font_color": "000000"
            }
            self.checkStyle(para, StyleDict)
            print("--------------------over--------------------")
            return

        def check_paper_writer_style(para: xml.dom.minidom.Element):
            print("----------checking paper writer info----------")
            StyleDict = {
                "font_b": "1",
                "font_sz": "小三",
                "font_eastAsia": "宋体",
                "font_ascii": "Times New Roman",
                "font_color": "000000"
            }
            self.checkStyle(para, StyleDict)
            print("--------------------over--------------------")
            return

        checkList = [[check_student_number, "分类号"], [check_school_number, "学校代码"],
                     [check_paper_type_style, "硕士学位论文"], [check_paper_title_style, ""],
                     [check_paper_writer_style, "学位申请人"]]
        self.print_log("----------Detecting Paper Chinese Cover----------")
        check_index = 0
        for index in range(para_index_begin, para_index_end):
            if self.getFullText(self.docx_body.childNodes[index]).replace(" ", "") == "":
                continue
            if check_index < len(checkList):
                check_item = checkList[check_index]
                if check_item[1] in self.getFullText(self.docx_body.childNodes[index]).replace(" ", ""):
                    check_item[0](self.docx_body.childNodes[index])
                    check_index += 1
                    continue
                else:
                    if check_item[1] == "":
                        check_item[0](self.docx_body.childNodes[index])
                        check_index += 1
                        continue
        if check_index < len(checkList):
            self.print_log("匹配出错，部分标签未匹配")

            # print(self.docx_body.childNodes[index].tagName)
            # print(self.getFullText(self.docx_body.childNodes[index]), end="\n\n")

        self.print_log("--------------------------------------------------")

    def DetectEnglishCover(self, para_index_begin, para_index_end):
        self.print_log("----------Detecting Paper English Cover----------")
        # index = para_index_begin
        item_para_list = []
        if "Dissertation Submitted in Partial Fulfillment" not in self.getFullText(
                self.docx_body.childNodes[para_index_begin]):
            raise Exception("论文英文标题匹配错误")
        for index in range(para_index_begin, para_index_end):
            if self.getFullText(self.docx_body.childNodes[index]).replace(" ", "") != "":
                item_para_list.append(index)
        self.print_log("---!----")
        StyleDictOfEnglishType = {
            "font_b": "1",
            "font_sz": "小三",
            "font_eastAsia": "宋体",
            "font_ascii": "Times New Roman",
            "font_color": "000000"
        }
        self.checkStyle(self.docx_body.childNodes[item_para_list[0]], StyleDictOfEnglishType)
        StyleDictOfEnglishTitle = {
            "font_b": "1",
            "font_sz": "小二",
            "font_eastAsia": "宋体",
            "font_ascii": "Times New Roman",
            "font_color": "000000"
        }
        self.checkStyle(self.docx_body.childNodes[item_para_list[1]], StyleDictOfEnglishTitle)
        StyleDictOfEnglishWriterInfo = {
            "font_b": "1",
            "font_sz": "小三",
            "font_eastAsia": "宋体",
            "font_ascii": "Times New Roman",
            "font_color": "000000"
        }
        for index in range(2, 5):
            self.checkStyle(self.docx_body.childNodes[item_para_list[index]], StyleDictOfEnglishWriterInfo)
        StyleDictOfEnglishSchoolInfo = {
            "font_b": "1",
            "font_sz": "小三",
            "font_eastAsia": "宋体",
            "font_ascii": "Times New Roman",
            "font_color": "000000"
        }
        for index in range(5, 8):
            self.checkStyle(self.docx_body.childNodes[item_para_list[index]], StyleDictOfEnglishSchoolInfo)
        self.print_log("--------------------------------------------------")

    def DetectCopyright(self, para_index_begin, para_index_end):
        self.print_log("----------Detecting Paper Copyright----------")
        title_style_dict = {
            "font_b": "1",
            "font_sz": "三号",
            "font_eastAsia": "黑体",
            "font_ascii": "",
            "font_color": "000000"
        }
        text_style_dict = {
            "font_b": "0",
            "font_sz": "小四",
            "font_eastAsia": "宋体",
            "font_ascii": "",
            "font_color": "000000"
        }
        for index in range(para_index_begin, para_index_end):
            if self.getFullText(self.docx_body.childNodes[index]).replace(" ", "") == "":
                continue
            elif self.getFullText(self.docx_body.childNodes[index]).replace(" ",
                                                                            "") == "独创性声明" or self.getFullText(
                self.docx_body.childNodes[index]).replace(" ", "") == "学位论文版权使用授权书":
                self.checkStyle(self.docx_body.childNodes[index], title_style_dict)
            else:
                self.checkStyle(self.docx_body.childNodes[index], text_style_dict)
        self.print_log("--------------------------------------------------")

    def DetectAbstract(self, index_list, index_start, index_end):
        self.DetectChineseAbstract(index_list[index_start], index_list[index_start + 1])
        self.DetectEnglishAbstract(index_list[index_start + 1], index_list[index_end])

    def DetectChineseAbstract(self, para_index_begin, para_index_end):
        self.print_log("----------Detecting Paper Chinese Abstract----------")
        index = para_index_begin
        while self.getFullText(self.docx_body.childNodes[index]).replace(" ", "") == "":
            index += 1
        abstract_style_dict = {
            "font_b": "1",
            "font_sz": "16",
            "font_eastAsia": "黑体",
            "font_ascii": "",
            "font_color": "000000"
        }
        if self.getFullText(self.docx_body.childNodes[index]).replace(" ", "") == "摘要":
            self.checkStyle(self.docx_body.childNodes[index], abstract_style_dict)
            if not self.getFullText(self.docx_body.childNodes[index]).strip() == "摘  要":
                self.print_error("摘要二字中间应该要留两个空格", self.docx_body.childNodes[index])
        text_style_dict = {
            "font_b": "0",
            "font_sz": "小四",
            "font_eastAsia": "宋体",
            "font_ascii": "",
            "font_color": "000000"
        }
        index += 1
        # for index in range(temp_index, para_index_end):
        while index < para_index_end:
            if "关键词：" in self.getFullText(self.docx_body.childNodes[index]).strip():
                break
            self.checkStyle(self.docx_body.childNodes[index], text_style_dict)
            index += 1
        keyword_list = re.split(re.compile(r'[;；]'),
                                self.getFullText(self.docx_body.childNodes[index]).strip().replace("关键词：", ""))
        if not 3 <= len(keyword_list) <= 8:
            self.print_error("关键字数量一般为3~8个", self.docx_body.childNodes[index])
        keyword_title_style_dict = {
            "font_b": "1",
            "font_sz": "小四",
            "font_eastAsia": "黑体",
            "font_ascii": "",
            "font_color": "000000"
        }
        keyword_style_dict = {
            "font_b": "0",
            "font_sz": "小四",
            "font_eastAsia": "宋体",
            "font_ascii": "",
            "font_color": "000000"
        }
        print(self.getFullText(self.docx_body.childNodes[index]))
        bold_key_word = True
        text1 = ""
        for elem in self.docx_body.childNodes[index].childNodes:
            if elem.nodeName == "w:r":
                if bold_key_word:
                    self.checkStyle(elem, keyword_title_style_dict)
                    text1 += self.getFullText(elem).replace(" ", "").strip()
                    if text1 == "关键词：":
                        bold_key_word = False
                else:
                    self.checkStyle(elem, keyword_style_dict)
        self.print_log("--------------------------------------------------")

    def DetectEnglishAbstract(self, para_index_begin, para_index_end):
        self.print_log("----------Detecting Paper English Abstract----------")
        index = para_index_begin
        while self.getFullText(self.docx_body.childNodes[index]).replace(" ", "") == "":
            index += 1
        if self.getFullText(self.docx_body.childNodes[index]).replace(" ", "").strip() != "Abstract":
            self.print_error("此处应为Abstract", self.docx_body.childNodes[index])
        else:
            index += 1
        text_style_dict = {
            "font_b": "0",
            "font_sz": "小四",
            "font_eastAsia": "宋体",
            "font_ascii": "Times New Roman",
            "font_color": "000000"
        }
        # for index in range(temp_index, para_index_end):
        while index < para_index_end:
            if "Key words:" in self.getFullText(self.docx_body.childNodes[index]).strip():
                break
            self.checkStyle(self.docx_body.childNodes[index], text_style_dict)
            index += 1
        keyword_list = re.split(re.compile(r',\s*'),
                                self.getFullText(self.docx_body.childNodes[index]).strip().replace("Key words:", ""))
        if not 3 <= len(keyword_list) <= 8:
            self.print_error("关键字数量一般为3~8个", self.docx_body.childNodes[index])
        keyword_title_style_dict = {
            "font_b": "1",
            "font_sz": "小四",
            "font_eastAsia": "黑体",
            "font_ascii": "Times New Roman",
            "font_color": "000000"
        }
        keyword_style_dict = {
            "font_b": "0",
            "font_sz": "小四",
            "font_eastAsia": "宋体",
            "font_ascii": "Times New Roman",
            "font_color": "000000"
        }
        bold_key_word = True
        text1 = ""
        for elem in self.docx_body.childNodes[index].childNodes:
            if elem.nodeName == "w:r":
                if bold_key_word:
                    self.checkStyle(elem, keyword_title_style_dict)
                    text1 += self.getFullText(elem).replace(" ", "").strip()
                    if text1 == "关键词：":
                        bold_key_word = False
                else:
                    self.checkStyle(elem, keyword_style_dict)
        self.print_log("--------------------------------------------------")

    def DetectCatalogue(self, para_index_begin, para_index_end):
        self.print_log("----------Detecting Paper Catalogue----------")
        level0_reg = re.compile(r'^目\s*录$')
        level1_reg = re.compile(r'^摘\s*要|^Abstract|^致谢|^参考文献|^附录\d+\s+|^\d+\s*.+')
        level2_reg = re.compile(r'\d+\.\d+\s*[\u4e00-\u9fa5a-zA-Z]+.?')
        level0_style_dict = {
            "font_b": "1",
            "font_sz": "三号",
            "font_eastAsia": "黑体",
            "font_ascii": "Times New Roman",
            "font_color": "000000"
        }
        level1_style_dict = {
            "font_b": "1",
            "font_sz": "四号",
            "font_eastAsia": "黑体",
            "font_ascii": "Times New Roman",
            "font_color": "000000"
        }
        level2_style_dict = {
            "font_b": "0",
            "font_sz": "四号",
            "font_eastAsia": "宋体",
            "font_ascii": "Times New Roman",
            "font_color": "000000"
        }
        for index in range(para_index_begin, para_index_end):
            print("------------------------------------------------------------------------")
            if level2_reg.match(self.getFullText(self.docx_body.childNodes[index]).replace(" ", "").strip()):
                # print("2--------------------", self.getFullText(self.docx_body.childNodes[index]))
                self.checkStyle(self.docx_body.childNodes[index], level2_style_dict)
                continue
            if level1_reg.match(self.getFullText(self.docx_body.childNodes[index]).replace(" ", "").strip()):
                # print("1--------------------", self.getFullText(self.docx_body.childNodes[index]))
                self.checkStyle(self.docx_body.childNodes[index], level1_style_dict)
                continue
            if level0_reg.match(self.getFullText(self.docx_body.childNodes[index]).replace(" ", "").strip()):
                # print("0--------------------", self.getFullText(self.docx_body.childNodes[index]))
                self.checkStyle(self.docx_body.childNodes[index], level0_style_dict)
                continue
        self.print_log("--------------------------------------------------")
        pass

    def DetectText(self, index_list: list, index_start: int, index_end: int):
        self.print_log("----------Detecting Paper Text----------")
        title_style_dict = {
            "1": {
                "font_b": "1",
                "font_sz": "三号",
                "font_eastAsia": "黑体",
                "font_ascii": "Times New Roman",
                "font_color": "000000",
                "jc": "center"
            },
            "2": {
                "font_b": "1",
                "font_sz": "四号",
                "font_eastAsia": "黑体",
                "font_ascii": "Times New Roman",
                "font_color": "000000"
            },
            "3": {
                "font_b": "1",
                "font_sz": "小四",
                "font_eastAsia": "黑体",
                "font_ascii": "Times New Roman",
                "font_color": "000000"
            }
        }
        content_style_dict = {
            "font_b": "0",
            "font_sz": "小四",
            "font_eastAsia": "宋体",
            "font_ascii": "Times New Roman",
            "font_color": "000000"
        }
        tab_and_figure_title_style_dict = {
            "font_b": "0",
            "font_sz": "五号",
            "font_eastAsia": "宋体",
            "font_ascii": "Times New Roman",
            "font_color": "000000"
        }
        para_index_begin = index_list[index_start]
        para_index_end = index_list[index_end] - 1
        for index in range(para_index_begin, para_index_end):
            if self.docx_body.childNodes[index].tagName == "w:p":
                if self.getFullText(self.docx_body.childNodes[index]).strip() == "":
                    continue
                is_title, title_type = self.isTitle(self.docx_body.childNodes[index])
                is_tab_title, tab_title_type = self.isTabTitle(self.docx_body.childNodes[index])
                is_pic_title, pic_title_type = self.isPicTitle(self.docx_body.childNodes[index])
                if is_title:
                    print(title_type, self.getFullText(self.docx_body.childNodes[index]))
                    self.checkStyle(self.docx_body.childNodes[index], title_style_dict[str(title_type)])
                elif is_tab_title or is_pic_title:
                    self.checkStyle(self.docx_body.childNodes[index], tab_and_figure_title_style_dict)
                else:
                    self.checkStyle(self.docx_body.childNodes[index], content_style_dict)

        self.print_log("--------------------------------------------------")

    def DetectAcknowledge(self, para_index_begin, para_index_end):
        self.print_log("----------Detecting Paper Acknowledge----------")
        for index in range(para_index_begin, para_index_end):
            print(self.getFullText(self.docx_body.childNodes[index]))
        self.print_log("--------------------------------------------------")

    def DetectReference(self, para_index_begin, para_index_end):
        self.print_log("----------Detecting Paper Reference----------")
        for index in range(para_index_begin, para_index_end):
            print(self.getFullText(self.docx_body.childNodes[index]))
        self.print_log("--------------------------------------------------")

    def DetectAppendixes(self, index_list: list, index_start: int, index_end: int):
        self.print_log("----------Detecting Paper Appendixes----------")
        for i in range(index_start, index_end - 1):
            self.print_log("----------Detecting Paper Appendix" + str(i - index_start + 1) + "----------")
            self.DetectAppendix(index_list[i], index_list[i + 1])
        pass

    def DetectAppendix(self, para_index_begin, para_index_end):
        for index in range(para_index_begin, para_index_end):
            print(self.getFullText(self.docx_body.childNodes[index]))
        self.print_log("--------------------------------------------------")

    def test_method(self):
        self.DetectPaper()
        # document = self.doc.childNodes[0]
        # body = document.childNodes[0]
        # for each in body.childNodes:
        #     print(each.tagName)
        # first_paragraph = body.childNodes[0]
        # print(len(body.childNodes))
        # print(first_paragraph.tagName)
        # print(len(first_paragraph.childNodes))
        # for each in first_paragraph.childNodes:
        #     print(each.tagName)
        # for para in self.doc.getElementsByTagName("w:p"):
        #     self.getFullTextAndCommentOfPara(para)
        return
