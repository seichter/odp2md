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
        self.data = [{'p' : 'dadadad'},{'l' : ['blah', 'more']  }]

    def __str__(self):
        return " ".join([self.name," ".join(self.text)])

class ODPJuicer:
    def __init__(self):
        self.slides = []
        self.doc = None
        self.currentSlide = None

    def debug(self,node):
        print("node ", node)

    def getTextFromNode(self,node):
        for c in node.childNodes:
            if c.nodeType == c.TEXT_NODE:
                self.currentSlide.text + c.data

    def hasAttributeWithValue(self,node,name,value):
        for attribute_name,attribute_value in node.attributes.items():
            if attribute_name == name and attribute_value == value:
                return True
        return False

    def handleTextSpan(self,node):
        self.getTextFromNode(node)
        pass

    def handleTextParagraph(self,node):
        # iterate over textspans
        self.getTextFromNode(node)

    def handleTextListItem(self,node):
        for c in node.childNodes:
            if c.tagName == 'text:p':
                self.handleTextParagraph(c)
            if c.tagName == 'text:list':
                self.handleTextList(c)
    
    def handleTextBox(self,node):
        for c in node.childNodes:
            if c.tagName == 'text:p':
                self.handleTextParagraph(c)
            if c.tagName == 'text:list':
                self.handleTextList(c)

    def handleTextList(self,node):
        for c in node.childNodes:
            if c.tagName == 'text:list-item':
                self.handleTextListItem(c)

    def handleFrame(self,node):
        for c in node.childNodes:
            if c.tagName == 'draw:text-box':
                self.handleTextBox(c)

    def handlePage(self,node):
        # set new current slide
        self.currentSlide = Slide()
        self.currentSlide.name = node.attributes['draw:name']
        # parse
        self.handleNode(node)
        # store
        self.slides.append(self.currentSlide)

    
    def handleText(self,node):
        pass

    def handleNode(self,node):

        tag_functions = {
                #'draw:text-box' : self.handleTextBox,
                'draw:page' : self.handlePage,
                'text:p' : self.getTextFromNode
                }

        for c in node.childNodes:
            try:
                tag_functions[c.tagName](c)
                print('handled',c.tagName)
            except:
                try:
                    print('not handled:',c.tagName)
                except:
                    self.currentSlide.text.append(c.data)

                self.handleNode(c)

    def handleDocument(self,dom):
        # we only need the pages
        pages = dom.getElementsByTagName("draw:page")
        for page in pages:
            self.currentSlide = Slide()
            self.handleNode(page)
            self.slides.append(self.currentSlide)

        # debug
        for slide in self.slides:
            print(slide)

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