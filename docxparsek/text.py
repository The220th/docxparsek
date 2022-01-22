# -*- coding: utf-8 -*-

import re

class Text:
    __text = ""
    __bold = None
    __italic = None
    __color = None

    def __init__(self, xmlstr : str):
        ps = re.findall(r"<w:p.*?<\/w:p>", xmlstr)
        for p in ps:
            ws = re.findall(r"<w:t.*?</w:t>", p)
            for w in ws:
                w = w[w.find(">")+1:len(w)-len("</w:t>")]
                self.__text += w
        #print("Text inited")

    def getText(self) -> str:
        return self.__text