from bs4 import BeautifulSoup
from .row import Row

# -*- coding: utf-8 -*-

class Table:

    __rows = None
    __nrows = 0

    __soup = None


    def __init__(self, xmlstr : str):
        xmlstr = str(xmlstr)
        self.__soup = BeautifulSoup(xmlstr, "lxml")
        blocks = self.__soup.find_all("w:tr")
        self.__rows = []
        for el in blocks:
            self.__rows.append(Row(el))
    
    def __iter__(self):
        self.__nrows = 0
        return self

    def __next__(self):
        if self.__nrows < len(self.__rows):
            result = self.__rows[self.__nrows]
            self.__nrows += 1
            return result
        else:
            raise StopIteration