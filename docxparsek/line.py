# -*- coding: utf-8 -*-

from typing import Final
import re
#from .doc import Doc
from .text import Text
from .image import Image
from .table import Table

class Line:
    #_TEXT : Final = "text"
    #_IMAGE : Final = "img"
    #_TABLE : Final = "table"
    
    @staticmethod
    def TEXT() -> str: return "=TEXT"

    @staticmethod
    def IMAGE() -> str: return "=IMG"

    @staticmethod
    def TABLE() -> str: return "=TABLE"

    @staticmethod
    def OTHER() -> str: return "=OTHER"

    __type = None

    __src = None

    __doc = None

    def __init__(self, xml : str, doc):

        self.__doc = doc

        if(xml.find("<w:tbl") != -1):     # Таблица
            self.__ifTable(xml)
            print("\n\n" + f"{xml}" + "\n\n")
        elif(xml.find("<pic:pic") != -1): # Картинка
            self.__ifPic(xml)
        elif(xml.find("<w:t") != -1):     # Текст
            self.__ifText(xml)
        else:                             # Что это?!
            self.__ifOther(xml)
            #raise Exception("Cannot determine the type... ")

    def __ifTable(self, xmltag : str) -> None:
        # pizza time
        self.__type = Line.TABLE()
        self.__src = Table(xmltag)
        return

    def __ifPic(self, xmltag : str) -> None:
        self.__type = Line.IMAGE()
        self.__src = Image(xmltag, self.__doc)
        return

    def __ifText(self, xmltag : str) -> None:
        self.__type = Line.TEXT()
        self.__src = Text(xmltag)
        return

    def __ifOther(self, xmltag : str) -> None:
        self.__type = Line.OTHER()
        self.__src = xmltag

    def getType(self) -> str:
        return self.__type

    def isText(self) -> bool:
        if(self.__type == Line.TEXT()):
            return True
        else:
            return False

    def isImage(self) -> bool:
        if(self.__type == Line.IMAGE()):
            return True
        else:
            return False

    def isTable(self) -> bool:
        if(self.__type == Line.TABLE()):
            return True
        else:
            return False

    def isOther(self) -> bool:
        if(self.__type == Line.OTHER()):
            return True
        else:
            return False

    def getSrc(self):
        return self.__src