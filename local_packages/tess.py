import subprocess
from subprocess import PIPE


tess = 'tesseract.exe'
image = './Untitled.jpg'
output_place = 'stdout'
tessedit_char_whitelist = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ 0123456789'
psm = 6
lang = 'eng'
config = 'tessedit_char_whitelist="'+tessedit_char_whitelist+'"'

tessArgs = image+' '+output_place+' --psm '+str(psm)+' -l '+lang+' -c '+config
cmd = 'cmd'
returncode = subprocess.run(
    tess+' '+tessArgs, shell=True, stdout=PIPE, stderr=PIPE, text=True)



print(returncode.stdout)
