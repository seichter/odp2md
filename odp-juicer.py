#!/usr/bin/env python3

import os
import zipfile
import argparse

class ODPJuicer:
    def __init__(self):
        self.slides = []

    def open(self,fname):
        with zipfile.ZipFile(fname) as odp:
            info = odp.infolist()
            for i in info:
                print(i)
                if (i.filename == 'content.xml'):
                    with odp.open('content.xml') as index:
                        print(index.read())

def main():
    parser = argparse.ArgumentParser(description='OpenDocument Presentation Parser')
    parser.add_argument("--input", default="",
                        help="ODP file to parse and extract")

    args = parser.parse_args()

    juicer = ODPJuicer()
    juicer.open(args.input)

if __name__ == '__main__':
    main()