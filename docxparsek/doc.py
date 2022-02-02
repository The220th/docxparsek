# -*- coding: utf-8 -*-

import zipfile
from . import *
#from . import Line
import re

class Doc:
    __fileName = ""
    __z = None # zip
    __lines = []
    __nlines = 0

    def __init__(self, fileName : str):
        self.__fileName = fileName
        self.__z = zipfile.ZipFile(fileName)
        self.__initLines()
        #print(self.__z.namelist())
    
    def getDocXML(self) -> str:
        return self.__z.open("word/document.xml").read().decode("utf-8")

    def __initLines(self):
        s = self.getDocXML()
        reres = re.findall(r"<w:p.*?<\/w:p>|<w:tbl>.*?<\/w:tbl>", s)
        for xmltag in reres:
            #print("\n\n" + xmltag)
            oneline = Line(xmltag, self)
            self.__lines.append(oneline)

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

    def readFileText(self, path : str) -> str:
        return self.__z.open(path).read().decode("utf-8")

    def getImageBytes(self, image : Image) -> bytes:
        return image.getBytes()

    def readImage(self, pathImage : str) -> bytes:
        return self.__z.open(pathImage).read()
