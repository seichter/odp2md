#!/usr/bin/env python3

"""

ODP2Pandoc 2020.1.0

ODP2Pandoc is a tiny tool to convert 
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

$> python odp2pandoc --input <myslide.odp>

 """

import os
import zipfile
import argparse
from enum import Enum

import xml.dom.minidom as dom


class Slide:
    def __init__(self):
        self.title = ''
        self.text = ""
        self.notes = ""
        self.media = []

    def generateMarkdown(self):
        out = "## {0}\n\n{1}\n".format(self.title,self.text)
        for m in self.media:
            out += "![]({0})\n".format(m)
        return out

    def __str__(self):
        return self.generateMarkdown()

class Scope(Enum):
    NONE = 0
    TITLE = 1
    OUTLINE = 2
    NOTES = 3
    IMAGES = 4


class Parser:
    def __init__(self):
        self.slides = []
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
        # elif == 'draw:image':
        #     self.currentScope = Scope.IMAGES
        else:
            pass
            # print("Unhandled Scope!")

        if node.nodeName in ['draw:image', 'draw:plugin']:
            for k,v in node.attributes.items():
                if k == 'xlink:href':
                    self.currentSlide.media.append(v)

            
        t = self.getTextFromNode(node)

        if t != None:
            if self.currentScope == Scope.OUTLINE:
                self.currentText += (" " * self.currentDepth) + t + "\n"
            elif self.currentScope == Scope.TITLE:
                self.currentSlide.title += t
            elif self.currentScope == Scope.IMAGES:
                print('image title ',t)     

        for c in node.childNodes:
            self.currentDepth += 1
            self.handleNode(c)
            self.currentDepth -= 1
                

    def handleDocument(self,dom):
        # we only need the pages
        pages = dom.getElementsByTagName("draw:page")
        # iterate pages
        for page in pages:

            self.currentDepth = 0
            self.currentSlide = Slide()
            self.handleNode(page)
            self.currentSlide.text = self.currentText
            self.slides.append(self.currentSlide)

            self.currentText = ""


    def open(self,fname):
        with zipfile.ZipFile(fname) as odp:
            info = odp.infolist()
            for i in info:
                # print(i)
                if (i.filename == 'content.xml'):
                    with odp.open('content.xml') as index:
                        doc = dom.parseString(index.read())
                        self.handleDocument(doc)
            # output markdown
            for slide in self.slides:
                print(slide)
            # generate files
            for slide in self.slides:
                for m in slide.media:
                    odp.extract(m,'.')


def main():
    argument_parser = argparse.ArgumentParser(description='OpenDocument Presentation converter')
    argument_parser.add_argument("--input", required=True,help="ODP file to parse and extract")

    args = argument_parser.parse_args()

    juicer = Parser()
    if 'input' in args:
        juicer.open(args.input)
    else:
        argument_parser.print_help()

if __name__ == '__main__':
    main()