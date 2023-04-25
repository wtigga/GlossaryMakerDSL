# **Game Code Snippets Tool**

### Description

    • This tool will extract all the code snippets defined by regex, find their corresponding meaning in source and target language, and save as a glossary in Excel or as a dictionary in DSL (GoldenDict-compatible).
    • The purpose is to have all the codes listed along with their actual translations, and to use in the GoldenDict (read below).
    • It was created specifically for the needs of the Russian Localization team of Genshin Impact, but can be used by other languages as well.
    • It runs locally and does not transfer any data outside of your computer.
    • It is written in Python, but it doesn't require Python to be installed on your machine. It is  compiled into an *.exe to run on any machine. It can be compiled for other platforms (macOS, Linux, etc.) from the source code.
    
![image](https://user-images.githubusercontent.com/7037184/234344793-3c510335-06fe-4eca-b9b0-5a0c1c4d64fd.png)

### What it does

It process all the *.xlsx files from a specified folder. First, it loops through the 'comment' column (usually 'EXTRA') and finds all the mentions of code snippets. It also extracts the Source equivalents of that code that is usually stored in the comments along with it.
Then, it searches for the translation of those words in the files.
Finally, it stores it as either bilingual Excel or a GoldenDict / Lingvo dictionary (DSL).

### How to use
    1. Download your source files with confirmed translations. You need files with translations  to be able to populate the translation part.
    2. Put them in the folder that you will select with 'Browse Folder' button.
    3. Select 'Source column' and 'Target column', they should be identical with the structure of your files. Case sensitive.
    4. Click 'Process files' and wait until you see a confirmation pop-up message. Larger files might take up a few minutes.
    5. Use 'Output Excel' and 'Output DSL' to choose where to store the results.
    6. Click on 'Save to Excel' and 'Save to DSL' to save the files.

### How to use DSL in GoldenDict
    1. Install GoldenDict, configure hotkey as per your liking (Ctrl+C+C seems to be comfortable choice)
    2. Define the folder where your DSL file is located, click 'Rescan now', and close settings.
    3. While you're in the browser CAT and stumble upon a code snippet, copy (Ctrl+C) the code and use the hotkey (like Ctrl+C+C defined above) to call GoldenDict:
    4. Click on the 'Close words':
    
 ### Notes for developers
 This script uses a bunch of dicts instead of fancy Pandas because I'm not a real developer. But the script gets the job done for my case, and that's what matters.
 To compile the file into the Windows executable:

 pyinstaller --onefile --noconsole --upx-dir "c:\Soft\upx-4.0.2-win64" --name gcg_glossary_maker main.py

(Replace the '--upx-dir' with the actual path to the UPX executable.)