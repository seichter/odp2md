#!/usr/bin/env python3

"""

ODP Juicer 2020.1.0

ODP Juicer is a tiny tool to convert 
OpenDocument formatted presentations (ODP) 
to Pandocs' Markdown.

(c) Copyright 2019-2020 Hartmut Seichter

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.



Usage:

$> python odp-juicer --input <myslide.odp>

 """

import os
import zipfile
import argparse
from enum import Enum

import xml.dom.minidom as dom


class Slide:
    def __init__(self):
        self.title = 'Unknown'
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
        if node.attributes == None:
            return False
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
            pass
            # print("Unhandled Scope!")

        t = self.getTextFromNode(node)

        if t != None:
            if self.currentScope == Scope.OUTLINE:
                self.currentText += (" " * self.currentDepth) + t
            elif self.currentScope == Scope.TITLE:
                self.currentSlide.title += self.currentText
            

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