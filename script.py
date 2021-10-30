from xml.etree.cElementTree import XML
import zipfile
import pandas as pd

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'

MyDir = r'T:\\Bearings\Private\New Job Reports\2013\NJK196_D1064755B_NASG\2013-11-19 - NJK196 - NASG - Rick Rogers - F13-21454 - D1064755B - 2015 C520 F Seat Side Member.docx'
MyDir2 = r'T:\\Bearings\Private\New Job Reports\2013\NJP316-2_D1066254_Dura Moberly\NJP316 002- Dura Moberly- D1066254- AA1150-P002AE.docx'

def get_docx_text(path):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)
    paragraphs = []
    for paragraph in tree.getiterator(PARA):
        texts = [node.text
                 for node in paragraph.getiterator(TEXT)
                 if node.text]
        if (texts):
            paragraphs.append(''.join(texts))
    delete_spaces(paragraphs)
    return (paragraphs)

def delete_spaces(text):
    count = 0
    for i in text:
        text[count] = re.sub(' +',' ', text[count])
        count += 1


def notes_index(notes):
    x = ''
    for icount in range(0, len(notes)-1):
        try: 
            if (str(notes[icount]).index("NOTES:")>=0):
                x = icount
                break
        except ValueError:
            continue
    return x
            
def review_index(notes):
    y = ''
    for icount in range(0, len(notes)-1):
        try:
            if (str(notes[icount]).index("Review Question")>=0):
                y = icount
                break
        except ValueError:
            continue
    return y

def extract_data(data):
    njr_col = ["NA","FROM:","CODE:","CUSTOMER:","SGPPL PART#:","DATE: ","CUSTOMER PART#:","SGPPL DRAWING#:","NA","AUTO MANUFACTURER:","APPLICATION:","NA","MATERIAL:","NA","CUSTOMER PURCHASE ORDER#:","SGPPL QUOTE#:","SELLING PRICE:","ANNUAL VOLUME:","NA","PPAP QUANTITY:","NA","PPAP DUE DATE:","NA","NA","PROTOTYPE BUILD START DATE:","NA","NA","NA","NA","PRODUCTION START DATE:","NA","NA","NA","NA","NA","NA","NA","NA","NA","NOTES:","NA","NA","NA","NA","NA","NA","NA","NA","NA","NA","NA"]
    ls = []
    notes = notes_index(data)
    range = review_index(data) - notes
    for njr_column in njr_col:
        if (njr_column == "NA"):
            ls.append('N/A')
        else:
            count = 0
            for info in data:
                try:
                    if(info.index(njr_column) >= 0):
                        if (notes == count):
                            ls.append(str(data[count]))
                        else:
                            ls.append(str(data[count+1]))
                        count += 1
                except ValueError:
                        count += 1
                        continue
    return ls

def extract_data_dict(data):
    njr_col = ["NA","FROM:","CODE:","CUSTOMER:","SGPPL PART#:","DATE: ","CUSTOMER PART#:","SGPPL DRAWING#:","NA","AUTO MANUFACTURER:","APPLICATION:","NA","MATERIAL:","NA","CUSTOMER PURCHASE ORDER#:","SGPPL QUOTE#:","SELLING PRICE:","ANNUAL VOLUME:","NA","PPAP QUANTITY:","NA","PPAP DUE DATE:","NA","NA","PROTOTYPE BUILD START DATE:","NA","NA","NA","NA","PRODUCTION START DATE:","NA","NA","NA","NA","NA","NA","NA","NA","NA","NOTES:","NA","NA","NA","NA","NA","NA","NA","NA","NA","NA","NA"]
    ls = {}
    notes = notes_index(data)
    range = review_index(data) - notes
    for njr_column in njr_col:
        if (njr_column == "NA"):
            ls.update({'N/A':'N/A'})
        else:
            count = 0
            for info in data:
                try:
                    if(info.index(njr_column) >= 0):
                        if (notes == count):
                            ls.update({njr_column:str(data[count])})
                        else:
                            ls.update({njr_column:str(data[count+1])})
                        count += 1
                except ValueError:
                        count += 1
                        continue
    return ls

