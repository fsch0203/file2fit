# File2secret

Encrypts or decrypts 1 or more files based on the Blowfish algorithm. Each encrypted file has a base64-encoded text format. Encrypted files get the extension .mys.

The files are encrypted with the MySecret Blowfish Encryption Utility:

https://www.di-mgt.com.au/mysecret.html
http://www.schneier.com/index.html

MySecret.exe is small - 150 kB - and quick. It uses the Blowfish algorithm to create base64-encoded text output that can be easily transmitted over the Internet or stored on any computer system. For more information on Blowfish see the web site of its inventor, Bruce Schneier and his references to Products that Use Blowfish.


## Install instructions
* Copy the folder file2secret into a folder of your choice
* Make a new entry in the TC button bar (Configuration, Button Bar)
    - at the command line: fill in the path to the file2secret.vbs file
    - for parameters fill in: %L %T
    - for icon file: fill in the path to the file2secret.ico file
    - for tooltip fill in the text: Encrypt or decrypt with mysecret
    - see screenshot
