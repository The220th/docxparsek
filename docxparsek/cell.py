from .line import Line
import re

# -*- coding: utf-8 -*-

class Cell:

    __lines = None
    __nlines = 0

    def __init__(self, xmlstr : str):
        xmlstr = str(xmlstr)
        print(xmlstr)

        self.__lines = []
        reres = re.findall(r"<w:p.*?<\/w:p>|<w:tbl>.*?<\/w:tbl>", xmlstr)
        for xmltag in reres:
            self.__lines.append(Line(xmltag, None))
        #print("Cell inited")

    def __iter__(self):
        self.__nlines = 0
        return self

    def __next__(self):
        if self.__nlines < len(self.__lines):
            result = self.__lines[self.__nlines]
            self.__nlines -=-1
            return result
        else:
            raise StopIteration