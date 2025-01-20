import argparse
import datetime
import xml.etree.ElementTree as ET
import os

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Mapping fifths value and pitch correction
MAP_CORRECTION = {
    "-7": 0,
    "-6": 4,
    "-5": 1,
    "-4": 5,
    "-3": 2,
    "-2": 6,
    "-1": 3,
    "0": 0,
    "1": -4,
    "2": -1,
    "3": -5,
    "4": -2,
    "5": -6,
    "6": -3,
    "7": 0,
}

def v1(v2, v3):
    v4 = v2["step"]
    v5 = v2['octave']
    v6 = v2['type']
    v7 = v2['duration']
    v8 = v2['dot_count']
    v9 = v2['accidental']

    v10 = v3['fifths']
    v11 = v3['divisions']

    v12 = ""
    v13 = 0

    if v4 == 'R':
        v13 = 0
        v12 = chr(48)
    elif v4 == 'C':
        v14 = 1 + int(MAP_CORRECTION[str(v10)])
        v13 = v14 % 7
        v13 = v13 if v13 != 0 else 7
        v12 = chr(48 + v13)
        if v14 <= 0:
            v5 -= 1
    elif v4 == 'D':
        v14 = 2 + int(MAP_CORRECTION[str(v10)])
        v13 = v14 % 7
        v13 = v13 if v13 != 0 else 7
        v12 = chr(48 + v13)
        if v14 <= 0:
            v5 -= 1
    elif v4 == 'E':
        v14 = 3 + int(MAP_CORRECTION[str(v10)])
        v13 = v14 % 7
        v13 = v13 if v13 != 0 else 7
        v12 = chr(48 + v13)
        if v14 <= 0:
            v5 -= 1
    elif v4 == 'F':
        v14 = 4 + int(MAP_CORRECTION[str(v10)])
        v13 = v14 % 7
        v13 = v13 if v13 != 0 else 7
        v12 = chr(48 + v13)
        if v14 <= 0:
            v5 -= 1
    elif v4 == 'G':
        v14 = 5 + int(MAP_CORRECTION[str(v10)])
        v13 = v14 % 7
        v13 = v13 if v13 != 0 else 7
        v12 = chr(48 + v13)
        if v14 <= 0:
            v5 -= 1
    elif v4 == 'A':
        v14 = 6 + int(MAP_CORRECTION[str(v10)])
        v13 = v14 % 7
        v13 = v13 if v13 != 0 else 7
        v12 = chr(48 + v13)
        if v14 <= 0:
            v5 -= 1
    elif v4 == 'B':
        v14 = 7 + int(MAP_CORRECTION[str(v10)])
        v13 = v14 % 7
        v13 = v13 if v13 != 0 else 7
        v12 = chr(48 + v13)
        if v14 <= 0:
            v5 -= 1
    else:
        return ""

    if v5 >= 4:
        if v5 == 5:
            v12 += chr(39)
        elif v5 == 6:
            v12 += chr(34)
        elif v5 == 7:
            v12 += chr(96)

        if v6 == 'whole' or v7 / v11 == 4:
            v12 += ' - - -'
        elif v6 == 'half' or v7 / v11 == 2:
            v12 += ' -'
        elif v6 == 'quarter' or v7 / v11 == 1:
            v12 += ''
        elif v6 == 'eighth' or v7 / v11 == 0.5:
            v12 += '_'
        elif v6 == "16th" or v7 / v11 == 0.25:
            v12 += "="
        elif v6 == "32nd" or v7 / v11 == 0.125:
            v12 += "/"
        elif v6 == "64th" or v7 / v11 == 1/16:
            v12 += "\\"

    elif v5 < 4:
        if v5 == 3:
            if v6 == 'whole' or v7 / v11 == 4:
                v12 += chr(113) + ' - - -'
            elif v6 == 'half' or v7 / v11 == 2:
                v12 += chr(113) + ' -'
            elif v6 == 'quarter' or v7 / v11 == 1:
                v12 += chr(113)
            elif v6 == 'eighth' or v7 / v11 == 0.5:
                v12 += chr(119)
            elif v6 == "16th" or v7 / v11 == 0.25:
                v12 += chr(101)
            elif v6 == "32nd" or v7 / v11 == 0.125:
                v12 += chr(114)
            elif v6 == "64th" or v7 / v11 == 1/16:
                v12 += chr(116)
        elif v5 == 2:
            if v6 == 'whole' or v7 / v11 == 4:
                v12 += chr(97) + ' - - -'
            elif v6 == 'half' or v7 / v11 == 2:
                v12 += chr(97) + ' -'
            elif v6 == 'quarter' or v7 / v11 == 1:
                v12 += chr(97)
            elif v6 == 'eighth' or v7 / v11 == 0.5:
                v12 += chr(115)
            elif v6 == "16th" or v7 / v11 == 0.25:
                v12 += chr(100)
            elif v6 == "32nd" or v7 / v11 == 0.125:
                v12 += chr(102)
            elif v6 == "64th" or v7 / v11 == 1/16:
                v12 += chr(103)
        elif v5 == 1:
            if v6 == 'whole' or v7 / v11 == 4:
                v12 += chr(122) + ' - - -'
            elif v6 == 'half' or v7 / v11 == 2:
                v12 += chr(122) + ' -'
            elif v6 == 'quarter' or v7 / v11 == 1:
                v12 += chr(122)
            elif v6 == 'eighth' or v7 / v11 == 0.5:
                v12 += chr(120)
            elif v6 == "16th" or v7 / v11 == 0.25:
                v12 += chr(99)
            elif v6 == "32nd" or v7 / v11 == 0.125:
                v12 += chr(118)
            elif v6 == "64th" or v7 / v11 == 1/16:
                v12 += chr(103)

    if v8 == 1:
        v12 += chr(46)
    elif v8 == 2:
        v12 += chr(46) + chr(44)

    if v9 is not None:
        if v9 == 'sharp':
            v12 = chr(105) + v12
        elif v9 == 'natural':
            v12 = chr(111) + v12
        elif v9 == 'flat':
            v12 = chr(112) + v12

    return v12


def parse(file_path) -> str:
    '''返回简谱排版所需要输入的字符串'''
    tree = ET.parse(file_path)
    root = tree.getroot()

    score = ''
    divisions = 1

    fifths = 0
    beam = False

    # 遍历小节(measure)
    for measure in root.findall('.//measure'):
        attributes = measure.find('attributes')
        notes = measure.findall('note') 
        if attributes is not None:
            # 获得乐谱属性
            div = attributes.find('divisions')
            if attributes.find('key/fifths') is not None:
                fifths = (attributes.find('key/fifths').text)
            if div is not None:
                divisions = int(div.text)
        
        # 遍历小节中的音符
        for note in notes:
            rest = note.find('rest')
            pitch = note.find('pitch')
            if rest is not None:
                # 特殊处理空拍
                duration = int(note.find('duration').text)
                note_type = None
                if note.find('type') is not None:
                    note_type = note.find('type').text
                dot_count = len(note.findall('dot'))

                score += v1(
                            {
                                'step': 'R',
                                'octave': 4,
                                'duration': duration,
                                'type': note_type,
                                'dot_count': dot_count,
                                'accidental': None,
                            }, 
                            {
                                'fifths': fifths,
                                'divisions': divisions,
                            }
                        ) + ' '
            elif pitch is not None:
                step = pitch.find('step').text
                octave = int(pitch.find('octave').text)
                duration = int(note.find('duration').text)
                note_type = note.find('type').text
                dot_count = len(note.findall('dot'))
                accidental = note.find('accidental')
                beam = note.find('beam')

                # 处理符杠
                spacing = ' '
                
                if beam is not None:
                    beam = beam.text
                if beam == "begin" or beam == "continue":
                    spacing = '' 
                
                # 升降号
                if accidental is not None:
                    accidental = accidental.text

                score += v1(
                            {
                                'step': step,
                                'octave': octave,
                                'duration': duration,
                                'type': note_type,
                                'dot_count': dot_count,
                                'accidental': accidental,
                            },
                            {
                                'fifths': fifths,
                                'divisions': divisions,
                            }
                        ) + spacing
        
        # 处理特殊小节线
        barline = measure.find('barline')
        if  barline is not None:
            if barline.find('bar-style').text == "light-light":
                score += "| | "
            elif barline.find('bar-style').text == "light-heavy":
                score += "+"
        else:
            score += "| " 
    return score
    

def create_doc(notes, output_doc):
    doc = Document()
    p = doc.add_paragraph()
    run = p.add_run(notes)
    run.font.name = 'jpfont-nds'
    r = run._element
    rFonts = r.find(qn('w:rPr')).find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        r.get_or_add('w:rPr').append(rFonts)
    rFonts.set(qn('w:eastAsia'), 'jpfont-nds')
    run.font.size = Pt(12)
    doc.save(output_doc)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('filename', type=str)
    args = parser.parse_args()

    musicxml_file = args.filename
    output_filename = os.path.splitext(os.path.basename(musicxml_file))[0]
    output_doc = 'outputs/' + output_filename + "_"+datetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S") + '.docx'

    notes = parse(musicxml_file)
    create_doc(notes, output_doc)

    print("Score saved to " + output_doc)