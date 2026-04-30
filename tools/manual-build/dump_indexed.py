#!/usr/bin/env python3
"""指定スライドの <a:t> を index 付きで出力する。"""
import sys
from lxml import etree

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}

def main():
    path = sys.argv[1]
    tree = etree.parse(path)
    for i, t in enumerate(tree.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t")):
        text = t.text or ""
        print(f"{i:3d} | {text}")

if __name__ == "__main__":
    main()
