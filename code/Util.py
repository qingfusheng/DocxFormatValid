import zipfile
import os
import re
from lxml import etree
import sys
from xml.dom import minidom

project_dir = os.path.abspath("../../DocxFormatValid/")


class Util:
    def __init__(self):
        self.docx_dir = os.path.join(project_dir, "code", "./DocxFilter")
        self.workflow_dir = os.path.join(project_dir, "code", "./WorkFlowFilter")
        if not os.path.exists(self.docx_dir):
            os.mkdir(self.docx_dir)
        if not os.path.exists(self.workflow_dir):
            os.mkdir(self.workflow_dir)

        self.docx_filename = '肖露露毕业论文.docx'
        self.new_docx_file = "new" + self.docx_filename
        self.unzip()
        self.doc = minidom.parse(os.path.join(self.workflow_dir, self.docx_filename, 'word', 'document.xml'))
        self.styles = minidom.parse(os.path.join(self.workflow_dir, self.docx_filename, 'word', 'styles.xml'))
        self.themes = minidom.parse(os.path.join(self.workflow_dir, self.docx_filename, 'word', 'theme', 'theme1.xml'))
        self.numbering = minidom.parse(os.path.join(self.workflow_dir, self.docx_filename, 'word', 'numbering.xml'))

    def unzip(self):
        f = zipfile.ZipFile(os.path.join(self.docx_dir, self.docx_filename))  # 打开需要修改的docx文件
        f.extractall(os.path.join(self.workflow_dir, self.docx_filename))  # 提取要修改的docx文件里的所有文件到workfolder文件夹
        f.close()

    
