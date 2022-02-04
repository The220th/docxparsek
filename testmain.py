# -*- coding: utf-8 -*-

from docxparsek import Doc
from docxparsek import Line
from docxparsek import Text
from docxparsek import Image
from docxparsek import Table

if __name__ == "__main__":
    #doc = Doc("./table.docx")
    #doc = Doc("./table3.docx")
    #doc = Doc("./table4.docx")
    #doc = Doc("./text.docx")
    doc = Doc("./check.docx")

    #print(doc.getDocXML())
    image_j = 0
    li = 0
    for line in doc:
        if(line.isText()):
            print(f"{li}: {line.getSrc().getText()}")
        elif(line.isImage()):
            print(f"{li}: <IMAGE>")
            with open(f"image{image_j}.png", 'wb') as temp:
                temp.write(line.getSrc().getBytes())
            image_j-=-1
        elif(line.isTable()):
            print(f"{li}: <TABLE>")
            table = line.getSrc()
            lli, llj = 0, 0
            for row in table:
                llj = 0
                #print(*row)
                for cell in row:
                    for lline in cell:
                        if(lline.isText()):
                            print(f"{lli}:{llj}:{lline.getSrc().getText()}")
                        elif(lline.isImage()):
                            print(f"{lli}:{llj}: <IMAGE>")
                            with open(f"image{image_j}.png", 'wb') as temp:
                                temp.write(lline.getSrc().getBytes())
                            image_j-=-1
                        else:
                            print(f"{lli}:{llj}:{lline.getSrc()}")
                    llj += 1
                lli += 1
        else:
            sss = "do nothing pls"
            #print(f"{li}: {line.getSrc()}")
        li-=-1