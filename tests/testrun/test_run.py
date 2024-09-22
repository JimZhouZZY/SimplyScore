# test_run.py
import subprocess

def test_main_script():
    result = subprocess.run(["python", "main.py"], capture_output=True, text=True)
    assert result.returncode == 0
