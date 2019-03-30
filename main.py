import difflib

try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile
import io
import os
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
ARTICLES_DIR = ''
GRAPHICS_DIR = ''
GRAPHICS_MATCH_WORD = 'TYTULOWA'

def get_docx_text(path):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)

    paragraphs = []
    texts = []
    iterator = 0 #0 nagłówek, 1 autor, 2 lead
    typ = 0 #0 akapit, 1 bold, 2 kursywa
    for paragraph in tree.iter(PARA):
        for node in paragraph.iter(TEXT):
            if node.findall(TEXT_SECTION):
                if node.find(RPR) and node.find(RPR).findall(BOLD):
                    typ = 1
                elif node.find(RPR) and node.find(RPR).findall(ITALIC):
                    typ = 2
                else:
                    typ = 0

                texts.append(node.find(TEXT_SECTION).text)
        if texts:
            if typ == 1 and iterator > 2:
                paragraphs.append('<h2>' + ''.join(texts) + '</h2>')
            elif iterator == 0:
                paragraphs.append('<h1>'+''.join(texts)+'</h1>')
            elif iterator == 1:
                paragraphs.append('<p class="author">' + ''.join(texts) + '</p>')
            elif typ == 1 and iterator == 2:
                paragraphs.append('<p class="lead">' + ''.join(texts) + '</p>')

            elif texts and typ == 2:
                paragraphs.append('<i>' + ''.join(texts) + '</i>')
            elif texts and typ == 0:
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
            a = get_docx_text(dirstr+filename)
            image = match(filename)
            a = a + '<img src="' + image + '">'
            with io.open(dirstr+'gotowe/'+filename+'.html', "w", encoding="utf-8") as f:
                f.write(a)
        else:
            continue
    # print(a)


def match(article):
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
        return matchimageslist[0]
    else:
        return ""



run()