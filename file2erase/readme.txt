Secure delete one or more files. The utility is based on MicroSoft utility SDelete.exe:
https://docs.microsoft.com/en-us/sysinternals/downloads/sdelete
The only way to ensure that deleted files, as well as files that you encrypt with EFS, are safe from recovery is to use a secure delete application. Secure delete applications overwrite a deleted file's on-disk data using techiques that are shown to make disk data unrecoverable, even using recovery technology that can read patterns in magnetic media that reveal weakly deleted files. SDelete (Secure Delete) is such an application.


Install instructions
* copy the folder sdelete into a folder of your choice
* make a new entry in the TC button bar (Configuration, Button Bar)
    - at the command line: fill in the path to the sdelete.vbs file
    - for parameters fill in: %L
    - for icon file: fill in the path to the sdelete.ico file
    - for tooltip fill in the text: Secure delete file(s)
    - see screenshot

Usage
* Use with care: deleted files and folders cannot be recovered!
