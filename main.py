import difflib
from PIL import Image
from resizeimage import resizeimage
import re

try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile
import io
import os
from enum import Enum


class Type(Enum):
    PARAGRAPH = 0
    BOLD = 1
    ITALIC = 2


"""
Module that extract text from MS XML Word document (.docx).
(Inspired by python-docx <https://github.com/mikemaccana/python-docx>)
"""

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 'r'
TEXT_SECTION = WORD_NAMESPACE + 't'
BOLD = WORD_NAMESPACE + 'b'
ITALIC = WORD_NAMESPACE + 'i'
RPR = WORD_NAMESPACE + 'rPr'
BR = WORD_NAMESPACE + 'br'
ARTICLES_DIR = ''
GRAPHICS_DIR = ''
IMAGES_DIR = ''
GRAPHICS_MATCH_WORD = 'TYTULOWA'
title = ''


def get_docx_text(path, filename):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path + filename)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)

    paragraphs = []
    texts = []
    iterator = 0
    for paragraph in tree.iter(PARA):
        for node in paragraph.iter(TEXT):

            if node.findall(TEXT_SECTION):
                if node.find(RPR) and node.find(RPR).findall(BOLD):
                    paragraph_type = Type.BOLD
                elif node.find(RPR) and node.find(RPR).findall(ITALIC):
                    texts.append('<i>' + ''.join(node.find(TEXT_SECTION).text) + '</i>')
                    continue
                    paragraph_type = Type.ITALIC
                else:
                    paragraph_type = Type.PARAGRAPH

                all_text = node.findall(TEXT_SECTION)
                for textes in all_text:
                    texts.append("".join(textes.itertext()))
        if texts:
            if paragraph_type == Type.BOLD and iterator > 2:
                paragraphs.append('<h2>' + ''.join(texts) + '</h2>')
            elif iterator == 0:
                global title
                title = ''.join(texts)
                paragraphs.append('<h1>' + ''.join(texts) + '</h1><hr />')
            elif iterator == 1:
                paragraphs.append('<p class="author">' + ''.join(texts) + '</p>')
                image = addimage(filename)
                #image = ''
                if image:
                    resizeimg(IMAGES_DIR + image)
                    paragraphs.append('<img src="../Images/' + image + '">')
                    m = re.search("\[(.+?)\]", image)
                    if m:
                        paragraphs.append('<p class="credit">' + m.group(1) + '</p>')
            elif paragraph_type == Type.BOLD and iterator == 2:
                paragraphs.append('<p class="lead">' + ''.join(texts) + '</p>')
            elif texts and paragraph_type == Type.ITALIC:
                paragraphs.append('<i>' + ''.join(texts) + '</i>')
            elif texts and paragraph_type == Type.PARAGRAPH:
                paragraphs.append('<p>' + ''.join(texts) + '</p>')
            iterator += 1
        texts = []
    return '\n'.join(paragraphs)


def run():
    dirstr = ARTICLES_DIR
    directory = os.fsencode(dirstr)

    for file in os.listdir(directory):
        filename = os.fsdecode(file)
        if filename.endswith(".docx") or filename.endswith(".doc"):
            a = get_docx_text(dirstr, filename)
            with io.open(dirstr + 'gotowe/' + filename + '.html', "w", encoding="utf-8") as f:
                f.write(a)
        else:
            continue
    # print(a)


def addimage(article):
    dirstr = GRAPHICS_DIR
    directory = os.fsencode(dirstr)
    graphics = [os.fsdecode(file)
                for file in os.listdir(directory)]
    folderlist = difflib.get_close_matches(article, graphics, 1, 0.3)
    print(folderlist)
    dirstr = dirstr + next(iter(folderlist), "")
    directory = os.fsencode(dirstr)
    imageslist = [os.fsdecode(file)
                  for file in os.listdir(directory)]
    matchimageslist = difflib.get_close_matches(GRAPHICS_MATCH_WORD, imageslist, 1, 0.3)

    print(matchimageslist)
    if len(matchimageslist) > 0:
        os.rename(GRAPHICS_DIR + folderlist[0] + '/' + matchimageslist[0],
                  IMAGES_DIR + matchimageslist[0])
        return matchimageslist[0]
    else:
        return ""


def resizeimg(path):
    with open(path, 'r+b') as f:
        with Image.open(f) as image:
            cover = resizeimage.resize_cover(image, [600, 400])
            cover.save(path, image.format)


run()
