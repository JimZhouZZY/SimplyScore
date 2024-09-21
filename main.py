import xml.etree.ElementTree as ET
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

note_map = {'R': '0', 'C': '1', 'D': '2', 'E': '3', 'F': '4', 'G': '5', 'A': '6', 'B': '7'}
correction = {
    "-7": 0,
    "-6": -4,
    "-5": -1,
    "-4":  -5,
    "-3": -3,
    "-2": -6,
    "-1": -3,
    "0": 0,
    "1": -4,
    "2": -1,
    "3": -5,
    "4": -2,
    "5": -6,
    "6": -3,
    "7": 0,
}

def convert_to_jianpu(note, fifths):
    note_step = note["step"]
    octave = note['octave']
    duration_type = note['type']
    duration = note['duration']
    dot_count = note['dot_count']
    if note_step in note_map:
        if note_step == 'R':
            jianpu_note = '0' 
        else:
            note_step_cor = (int(note_map[note_step])-correction[str(fifths)])%7
            jianpu_note = str(note_step_cor) if note_step_cor != 0 else '7'
    else:
        return ""

    if octave > 4:
        jianpu_note += "'" * (octave - 4)  # 高音点
    elif octave < 4:
        # TODO
        jianpu_note = jianpu_note  # 低音点
    
    # 加入时值标识
    if duration_type is not None:
        if duration_type == 'whole':
            jianpu_note += ' - - -'
        elif duration_type == 'half':
            jianpu_note += ' -'
        elif duration_type == 'quarter':
            jianpu_note += ''
        elif duration_type == 'eighth':
            jianpu_note += '_'
        elif duration_type == "16th":
            jianpu_note += "="
        elif duration_type == "32nd":
            jianpu_note += "/"
        elif duration_type == "64th":
            jianpu_note += "\\"
    else: 
        if duration == 8:
            jianpu_note += ' - - -'
        elif duration == 4:
            jianpu_note += ' -'
        elif duration == 2:
            jianpu_note += ''
        elif duration == 1:
            jianpu_note += '_'
        elif duration == 0.5:
            jianpu_note += "="
        elif duration == 0.25:
            jianpu_note += "/"
        elif duration == 0.125:
            jianpu_note += "\\" 
        
    if dot_count == 1:
        jianpu_note += '.'
    elif dot_count == 2:
        jianpu_note += '.,'

    return jianpu_note


def parse(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()

    score = ''
    divisions = 1

    fifths = 0

    for measure in root.findall('.//measure'):
        attributes = measure.find('attributes')
        notes = measure.findall('note') 
        if attributes is not None:
            div = attributes.find('divisions')
            if attributes.find('key/fifths') is not None:
                fifths = (attributes.find('key/fifths').text)
            if div is not None:
                divisions = int(div.text)
        for note in notes:
            rest = note.find('rest')
            pitch = note.find('pitch')
            if rest is not None:
                duration = int(note.find('duration').text)
                note_type = None
                if note.find('type') is not None:
                    note_type = note.find('type').text
                dot_count = len(note.findall('dot'))

                score += convert_to_jianpu(
                            {
                                'step': 'R',
                                'octave': 4,
                                'duration': duration,
                                'type': note_type,
                                'dot_count': dot_count
                            }, 
                            fifths
                        ) + ' '

            if pitch is not None:
                step = pitch.find('step').text
                octave = int(pitch.find('octave').text)
                duration = int(note.find('duration').text)
                note_type = note.find('type').text
                dot_count = len(note.findall('dot'))

                score += convert_to_jianpu(
                            {
                                'step': step,
                                'octave': octave,
                                'duration': duration,
                                'type': note_type,
                                'dot_count': dot_count
                            },
                            fifths
                        ) + ' '
        score += "| " 
    return score
    


def create_word_with_notes(notes, output_doc):
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
    musicxml_file = 'music.xml'
    output_doc = 'output.docx'

    notes = parse(musicxml_file)

    create_word_with_notes(notes, output_doc)