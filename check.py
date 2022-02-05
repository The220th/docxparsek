# -*- coding: utf-8 -*-

from docxparsek import Doc
from docxparsek import Line
from docxparsek import Text
from docxparsek import Run
from docxparsek import Image
from docxparsek import Table

def getTextProp(text : 'Text') -> str:
    res = ""
    if(text.isBold()):
        res += "BOLD, "
    if(text.isItalic()):
        res += "ITALIC, "
    if(text.isUnderline()):
        res += "UNDERLINE, "
    if(text.isColored()):
        res += "COLORED="
        res += text.getColor()
    else:
        res += "STD COLOR"
    return res

def getRunProp(run : 'Run') -> str:
    res = ""
    if(run.isBold()):
        res += "BOLD, "
    if(run.isItalic()):
        res += "ITALIC, "
    if(run.isUnderline()):
        res += "UNDERLINE, "
    if(run.isColored()):
        res += "COLORED="
        res += run.getColor()
    else:
        res += "STD COLOR"
    return res

def printText(line_i : int, line : 'Line', offset=0):
    offchar = "\t"
    print(f"{offchar*offset}{line_i}:<TEXT>")
    text = line.getSrc()
    print(f"{offchar*offset}\t\"{text.getText()}\" - {getTextProp(text)}")
    print(f"{offchar*offset}\t<RUNS>")
    run_i = 0
    for run in text:
        print(f"{offchar*offset}\t\tRUN_{run_i}: \"{run.getText()}\" - {getRunProp(run)}")
        run_i+=1

def printImage(line_i : int, line : 'Line', offset=0):
    global image_i
    offchar = "\t"
    print(f"{offchar*offset}{line_i}: <IMAGE {image_i}>")
    image = line.getSrc()
    with open(f"image{image_i}.png", 'wb') as temp:
        temp.write(image.getBytes())
    image_i+=1

def printTable(line_i : int, line : 'Line', offset=0):
    global image_i
    global table_i
    offchar = "\t"
    print(f"{offchar*offset}{line_i}: <TABLE {table_i}>")
    table_i+=1

    table = line.getSrc()
    for row in table:
        for cell in row:
            print(f"{offchar*offset}\tCELL={cell.getPosition()} (vMerged={cell.is_vMerged()}, hMerged={cell.is_hMerged()}):")
            for lline in cell:
                if(lline.isText()):
                    printText("", lline, offset+2)
                elif(lline.isImage()):
                    printImage("", lline, offset+2)
                else: # Table or other
                    print(f"{lline.getSrc()}")


if __name__ == "__main__":
    doc = Doc("./check.docx")

    #print(doc.getDocXML())

    image_i = 0
    table_i = 0
    line_i = 0
    for line in doc:
        if(line.isText()):
            printText(line_i, line)
        elif(line.isImage()):
            printImage(line_i, line)
        elif(line.isTable()):
            printTable(line_i, line)
        elif(line.isOther()):
            sss = "do nothing"
        line_i+=1