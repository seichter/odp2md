#!/usr/bin/env python3

"""

odp2md 2021.5.0

A tiny tool to convert OpenDocument formatted presentations
(ODP) to Pandocs' Markdown.

(c) Copyright 2019-2025 Hartmut Seichter

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

$> python odp2md --input <myslide.odp>

"""

import os, sys, argparse

from odp2md.parser import Parser


class App:
    def __init__(self) -> None:
        pass

    def run(self):
        argument_parser = argparse.ArgumentParser(
            prog="odp2md",
            description="OpenDocument presentation to markdown converter",
        )

        argument_parser.add_argument(
            "-i",
            "--input",
            required=True,
            help="ODP file to parse and extract",
        )
        argument_parser.add_argument(
            "-m",
            "--markdown",
            help="generate Markdown files",
            action="store_true",
        )
        argument_parser.add_argument(
            "-b",
            "--blocks",
            help="generate pandoc blocks for video files",
            action="store_true",
        )
        argument_parser.add_argument(
            "-x",
            "--extract",
            help="extract media files",
            action="store_true",
        )
        argument_parser.add_argument(
            "--mediadir",
            required=False,
            default="media",
            help="output directory for linked media",
        )

        args = argument_parser.parse_args()

        parser = Parser()

        if "input" in args:
            parser.open(args.input, args.mediadir, args.markdown, args.extract)
        else:
            argument_parser.print_help()


# runner object
app = App()


def main_cli():
    app.run()
