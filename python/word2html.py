#coding=utf-8
from pydocx import PyDocX
import sys
def word2html(htmlfile,wordfile):
    html = PyDocX.to_html(wordfile)
    f = open(htmlfile, 'w', encoding="utf-8")
    f.write(html)
    f.close()
    print("data end")

if __name__ == '__main__':
    htmlfile=sys.argv[1]
    wordfile=sys.argv[2]
    # htmlfile=r"C:\Users\yilanqunzhi\Desktop\openvpn使用说明.html"
    # wordfile=r"C:\Users\yilanqunzhi\Desktop\openvpn使用说明.docx"
    word2html(htmlfile,wordfile)