import os.path
import sys

sys.path.append(os.path.abspath("../code"))
from Util import Util


class PaperFormatDt:
    def __init__(self):
        pass

    def run(self, filename):
        tools = Util(filename)
        tools.DetectPaper()
        return