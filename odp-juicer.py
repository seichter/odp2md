#!/usr/bin/env python3

"""
 ODP Juicer is a tool to convert LibreOffice Impress Presentations in ODP format 
 to Pandoc Markdown
 """

import os
import zipfile
import argparse
from enum import Enum

import xml.dom.minidom as dom


class Slide:
    def __init__(self):
        self.name = 'Unknown'
        self.text = [[]]
        self.notes = []
        self.media = []
        #self.data = ['blah', ['more','blah blah']]

    def generateMarkdownBody(self,d,depth):
        out = ""
        for v in d:
            if isinstance(v,list):
                out += self.generateMarkdownBody(v,depth+1)
            else:
                out += "\t" * depth
                out += "- {0}\n".format(v)
        return out

    def generateMarkdown(self):
        
        out = "## {0}\n".format(self.name)

        out += self.generateMarkdownBody(self.text,0)
        
        return out

    def __str__(self):
        return self.generateMarkdown()

class Scope(Enum):
    NONE = 0
    TITLE = 1
    OUTLINE = 2
    NOTES = 3


class ODPJuicer:
    def __init__(self):
        self.slides = []
        self.doc = None
        self.currentSlide = None
        self.currentText = ""
        self.currentDepth = 0
        self.currentScope = Scope.NONE

    def getTextFromNode(self,node):
        if node.nodeType == node.TEXT_NODE and len(str(node.data)) > 0:
            return node.data
        return None

    def hasAttributeWithValue(self,node,name,value):
        for attribute_name,attribute_value in node.attributes.items():
            if attribute_name == name and attribute_value == value:
                return True
        return False

    def debugNode(self,node):
        print('node ', node.tagName)

    def handlePage(self,node):
        # set new current slide
        self.currentSlide = Slide()
        self.currentSlide.name = node.attributes['draw:name']
        # parse
        self.handleNode(node)
        # store
        self.slides.append(self.currentSlide)

    def handleNode(self,node):

        if self.hasAttributeWithValue(node,"presentation:class","title"):
            self.currentScope = Scope.TITLE
        elif self.hasAttributeWithValue(node,"presentation:class","outline"):
            self.currentScope = Scope.OUTLINE
        else:
            print("Unhandled Scope!")


        t = self.getTextFromNode(node)

        if t != None:
            self.currentText += (" " * self.currentDepth) + t

        for c in node.childNodes:
            self.currentDepth = self.currentDepth + 1
            self.handleNode(c)
            self.currentDepth = self.currentDepth - 1
                

    def handleDocument(self,dom):
        # we only need the pages
        pages = dom.getElementsByTagName("draw:page")
        # iterate pages
        for page in pages:
            self.currentSlide = Slide()
            self.handleNode(page)
            self.slides.append(self.currentSlide)

            print("text : ", self.currentText)
            self.currentText = ""

        # debug
        # for slide in self.slides:
        #     print(slide)

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