# test_run.py
import subprocess

def test_main_script():
    result = subprocess.run(["python3", "main.py", "tests/musicxmls/test_sharp_flat.xml"], capture_output=True, text=True)
    assert result.returncode == 0
