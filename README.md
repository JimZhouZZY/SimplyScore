# SimplyScore 
The purpose of this project is to convert the internationally used MusicXML file into simplified score (Jian Pu), a simple musical notation format that is easy for Chinese people to understand.

## Guidance
Firstly, install the required font for simplified music score in `fonts/`.

Secondly, install the required libs:`pip3 install python-docx`.

Lastly, run `python3 main.py tests/musicxml/test_run.xml` and check the `outputs` folder for the compiled score.

## Acknowledgment

Thanks David Evillious and 怒独僧 for developing 'jp-font2' and allowing me to use the font under GPLv2 License. Please visit this [link](http://www.nuduseng.com/jianpu/) for detailed information about the font.
