#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = katran009"

import argparse
import os
import zipfile


def parse_args():
    """Parse command line arguements"""
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "text", help="the text searched for in a file", type=str)
    parser.add_argument(
        "--dir", help="the directory being searched", default=os.getcwd())
    return parser.parse_args()


def search_files(files, args):
    print("Searching directory " + args.dir +
          " for text '" + args.text + "' ...")
    search_counter = 0
    match_counter = 0
    for file in files:
        search_counter += 1
        path_ = os.path.join(args.dir, file)
        with zipfile.ZipFile(path_) as z:
            with z.open('word/document.xml', 'r') as doc:
                for line in doc:
                    line = line.decode('utf-8')
                    # line = str(line, "utf-8")
                    if line.find(args.text) != -1:
                        index = line.find(args.text)
                        print("Match found in file " + path_)
                        print("   ..." + line[index-40: index+40] + "...")
                        match_counter += 1
    print("Total dotm files searched: " + str(search_counter))
    print("Total dotm files matched: " + str(match_counter))


def main():
    args = parse_args()
    files = [file for file in os.listdir(args.dir) if file.endswith("dotm")]
    search_files(files, args)


if __name__ == '__main__':
    main()
