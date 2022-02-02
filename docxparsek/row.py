from bs4 import BeautifulSoup
from . import Cell

# -*- coding: utf-8 -*-

class Row:

    __cells = None
    __ncells = 0

    __soup = None


    def __init__(self, xmlstr : str):
        xmlstr = str(xmlstr)
        print(xmlstr)
        self.__soup = BeautifulSoup(xmlstr, "lxml")
        blocks = self.__soup.find_all("w:tc")
        self.__cells = []
        for el in blocks:
            self.__cells.append(Cell(el))

        #print("Row inited")

    def __iter__(self):
        self.__ncells = 0
        return self

    def __next__(self):
        if self.__ncells < len(self.__cells):
            result = self.__cells[self.__ncells]
            self.__ncells += 1
            return result
        else:
            raise StopIteration