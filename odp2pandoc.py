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
import sys
import re, unicodedata
import textwrap
from enum import Enum
import xml.dom.minidom as dom

class Slide:
    def __init__(self):
        self.title = ''
        self.text = ""
        self.notes = ""
        self.media = []

    def generateMarkdown(self):
        # fix identation
        self.text = textwrap.dedent(self.text)
        out = "## {0}\n\n{1}\n".format(self.title,self.text)
        for m,v in self.media:
            out += "![]({0})\n".format(v)
        return out
    
    # override string representation
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
        self.mediaDirectory = 'media'

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


    def slugify(self,value, allow_unicode=False):
        """
        Convert to ASCII if 'allow_unicode' is False. Convert spaces to hyphens.
        Remove characters that aren't alphanumerics, underscores, or hyphens.
        Convert to lowercase. Also strip leading and trailing whitespace.
        """
        value = str(value)
        if allow_unicode:
            value = unicodedata.normalize('NFKC', value)
        else:
            value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('ascii')
        value = re.sub(r'[^\w\s-]', '', value.lower()).strip()
        return re.sub(r'[-\s]+', '-', value)

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
                    # get the extension
                    name,ext = os.path.splitext(v)
                    ext = ext.lower()
                    # now we create a new slug name for conversion
                    slug = self.slugify(self.currentSlide.title)
                    if len(slug) < 1:
                        slug = "slide-" + str(len(self.slides)) + "-image"
                    slug += "-" + str(len(self.currentSlide.media))
                    print(self.mediaDirectory)
                    self.currentSlide.media.append((v,os.path.join(self.mediaDirectory,slug+ext)))

            
        t = self.getTextFromNode(node)

        if t != None:
            if self.currentScope == Scope.OUTLINE:
                self.currentText += (' ' * self.currentDepth) + '- ' + t + "\n"
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


    def open(self,fname,mediaDir='media'):
        self.mediaDirectory = mediaDir
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
                for m,v in slide.media:
                    odp.extract(m,'.')
                    if not os.path.exists(self.mediaDirectory):
                        os.makedirs(self.mediaDirectory)
                    os.rename(os.path.join('.',m),v)



def main():
    argument_parser = argparse.ArgumentParser(description='OpenDocument Presentation converter')
    argument_parser.add_argument("--input", required=True,help="ODP file to parse and extract")
    argument_parser.add_argument("--mediadir", required=False,default='media',help="output directory for linked media")

    args = argument_parser.parse_args()

    juicer = Parser()
    if 'input' in args:
        juicer.open(args.input,args.mediadir)
    else:
        argument_parser.print_help()
        return

if __name__ == '__main__':
    main()