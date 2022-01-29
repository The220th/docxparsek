# -*- coding: utf-8 -*-

from docxparsek import Doc
from docxparsek import Line
from docxparsek import Text
from docxparsek import Image
from docxparsek import Table

if __name__ == "__main__":
    doc = Doc("/home/the220th/git/oldgit/nOMOuKA/docx2gift/FOS_proverki_ostatochnykh_znaniy.docx")
    #print(doc.getDocXML())
    j = 0
    li = 0
    for line in doc:
        if(line.isText()):
            print(f"{li}: {line.getSrc().getText()}")
        elif(line.isImage()):
            print(f"{li}: <IMAGE>")
            with open(f"image{j}.png", 'wb') as temp:
                temp.write(line.getSrc().getBytes())
            j-=-1
        elif(line.isTable()):
            print(f"{li}: <TABLE>")
            table = line.getSrc()
            li, lj = 0
            for row in table:
                for cell in row:
                    for lline in cell:
                        print(f"{li}:{lj}:{lline.getSrc().getText()}")
                    lj += 1
                li += 1
        else:
            sss = "do nothing pls"
            print(f"{li}: {line.getSrc()}")
        li-=-1