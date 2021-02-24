import subprocess
from subprocess import PIPE

tess = 'tesseract.exe ./Untitled.jpg stdout --psm 6 -l eng -c tessedit_char_whitelist="ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789"'
cmd = "cmd"
returncode = subprocess.run(
    tess, shell=True, stdout=PIPE, stderr=PIPE, text=True)

print(returncode.stdout)
