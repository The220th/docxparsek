 
from bs4 import BeautifulSoup

text = ""
with open("document.xml", 'r', encoding="utf-8") as temp:
    text = temp.read()

soup = BeautifulSoup(text, "lxml")

blocks = soup.find_all("w:tr")

for el in blocks:
    #print("\n\n\n")
    #print(el)
    txts = el.find_all("w:tc")
    for txt in txts:
        print(txt)
    #print(el)
