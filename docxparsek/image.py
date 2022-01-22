# -*- coding: utf-8 -*-

import re
#from .doc import Doc

class Image:

    __imageID = None

    __imagePath = None

    __doc = None

    def __init__(self, xmlstr : str, doc):

        self.__doc = doc

        self.__imageID = re.findall(r'<a:blip.+?embed=\s*".+?(?=\s*/>)', xmlstr)[0]
        self.__imageID = re.findall(r'(?<=").+?(?=")', self.__imageID)[0]
        #print(self.__imageID)
        self.__calcImagePath(self.__doc.readFileText("word/_rels/document.xml.rels"))
    
    def __calcImagePath(self, document_xml_rels : str):
        #print("Id\\s*=\\s*\"" + str(self.__imageID) + "\".+?(?=\"/>)")
        path = re.findall("Id\\s*=\\s*\"" + str(self.__imageID) + "\".+?(?=\"/>)", document_xml_rels)[0]
        path = re.findall(r'(?<=Target=").+', path)[0]
        path = "word/" + path
        #print(path)
        self.__imagePath = path

    def getBytes(self) -> bytes:
        return self.__doc.readImage(self.__imagePath)
