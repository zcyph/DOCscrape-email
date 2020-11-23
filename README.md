DOCscrape-email: a script to extract emails from Microsoft Word documents. If you're having trouble manually copying a number of email addresess from .DOCX files, give this a try. I hope it helps! Apologies, but it will unfortunately NOT work with legacy .DOC formatted files. Please ensure you're working with .DOCX - you can save to the newer format, even with LibreOffice. :)

## Requirements

**Python**

This is a Python script, so of course you need to have Python installed to run it. The exact command may vary depending on your exact system configuration. You can skip this step if you already know Python is installed.

Debian-based distributions: `sudo apt-get install python`

Arch-based distributions: `sudo pacman -S install python`

Next, make sure you have the required packages:

**xerox**

`pip install xerox`

**xclip**

Debian-based distributions: `sudo apt-get install xclip`

Arch-based distributions: `sudo pacman -S xclip`

**docx2txt**

`pip install docx2txt`

## Usage

Give the script executable permissions:

`chmod +x DOCscrape-email.py`

Then run it with your Microsoft Word filename as a command-line argument:

`./DOCscrape-email.py yourDOCfile.docx`

Or run it without executable permissions:

`python DOCscrape-email.py yourDOCfile.docx`

If the file does not reside in the same folder as the script, enter the full absolute path including file name and extension. You can also omit the filename argument and enter the filename when prompted after running the script instead:

![](https://github.com/zcyph/DOCscrape-email/blob/master/screenshot.png)

The script will give you an option to copy the list of emails to clipboard, and an option to copy them to a text file.
