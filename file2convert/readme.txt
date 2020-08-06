Converts one or more files into a different format. The files can be from 3 different categories:
Image files = bmp, gif, ico, jpg, jpeg, png, psd, psp, tga, tif, tiff, wmf, webp
Music files = wav, mp3, flac, ape, ogg
Document files = md, html, epub, txt, tex, xml

The utility is based on 
cwebp/dwebp: https://developers.google.com/speed/webp/docs/cwebp
irfanview: https://www.irfanview.com/
flac: https://xiph.org/flac/index.html
ogg: https://rarewares.org/ogg-oggenc.php
lame: https://lame.sourceforge.io/
ape: https://www.monkeysaudio.com/index.html
pandoc: https://pandoc.org/


Install instructions
* copy the folder convert into a folder of your choice
* make a new entry in the TC button bar (Configuration, Button Bar)
    - at the command line: fill in the path to the convert.vbs file
    - for parameters fill in (including quotes): "%P" "%T" %L
    - for icon file: fill in the path to the convert.ico file
    - for tooltip fill in the text: Convert image, music of document file(s)
    - see screenshot
* install pandoc from https://pandoc.org/

Usage
* Select one or more files or folder
* If the program recognizes the files, it will ask you the format to convert it in.
* If the category of files is not clear the program will ask you.
* Make sure you install pandoc from https://pandoc.org/. Without it the document conversion won't work. 
