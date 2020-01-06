#!/usr/bin/env python3

"""
 ODP Juicer is a tool to convert LibreOffice Impress Presentations in ODP format to Markdown compatible with pandoc
 """

import os
import zipfile
import argparse
import xml.dom.minidom as dom


class Slide:
    def __init__(self):
        self.name = 'Unknown'
        self.text = []
        self.notes = []
        self.media = []
        pass

class ODPJuicer:
    def __init__(self):
        self.slides = []
        self.doc = None
        self.currentSlide = None

    def debug(self,node):
        print("node ", node)

    def getTextFromNode(self,node):
        res = []
        for c in node.childNodes:
            if c.nodeType == c.TEXT_NODE:
                res.append(c.data)
        return res

    def hasAttributeWithValue(self,node,name,value):
        for attribute_name,attribute_value in node.attributes.items():
            if attribute_name == name and attribute_value == value:
                return True
        return False

    def handleTextSpan(self,node):
        res = self.getTextFromNode(node)
        print('span ',res)
        pass

    def handleTextParagraph(self,node):
        # iterate over textspans
        res = self.getTextFromNode(node)
        print('p',res)
        return res

    def handleTextListItem(self,node):
        for c in node.childNodes:
            if c.tagName == 'text:p':
                self.handleTextParagraph(c)
            if c.tagName == 'text:list':
                self.handleTextList(c)

    def handleTextList(self,node):
        for c in node.childNodes:
            if c.tagName == 'text:list-item':
                self.handleTextListItem(c)

    def handleTextBox(self,node):
        for c in node.childNodes:
            if c.tagName == 'text:p':
                self.handleTextParagraph(c)
            if c.tagName == 'text:list':
                self.handleTextList(c)

    def handleFrame(self,node):
        for c in node.childNodes:
            if c.tagName == 'draw:text-box':
                self.handleTextBox(c)

    def handlePage(self,node):
        for c in node.childNodes:
            if c.tagName == 'draw:text-box':
                self.handleTextBox(c)
            if c.tagName == 'draw:frame' and self.hasAttributeWithValue(c,'presentation:class','outline'):
                self.handleFrame(c)
        # for text_box in node.getElementsByTagName('draw:text-box'):
            # self.handleTextBox(text_box)

    def handleDocument(self,dom):
        pages = dom.getElementsByTagName("draw:page")
        for page in pages:
            # set new current slide
            self.currentSlide = Slide()
            # parse
            self.handlePage(page)
            # store
            self.slides.append(self.currentSlide)

        # debug
        print(self.slides)

    def open(self,fname):
        with zipfile.ZipFile(fname) as odp:
            info = odp.infolist()
            for i in info:
                # print(i)
                if (i.filename == 'content.xml'):
                    with odp.open('content.xml') as index:
                        doc = dom.parseString(index.read())
                        self.handleDocument(doc)

def main():
    parser = argparse.ArgumentParser(description='OpenDocument Presentation Parser')
    parser.add_argument("--input", default="",
                        help="ODP file to parse and extract")

    args = parser.parse_args()

    juicer = ODPJuicer()
    juicer.open(args.input)

if __name__ == '__main__':
    main()