#!/usr/bin/python
# A simple script to scrape email addresses from Microsoft Word (.docx) files
# Unfortunately .doc is too antiquated and requires other workarounds, this will only read .docx
# If your document is a .doc, it can first be converted to .docx using openoffice
# Prerequisites: must install antiword for this to work

import re, sys, xerox
import docx2txt
from os import system, name

# Clear screen for a tidy start
system('clear')

# While loop with try & except to keep asking for file name until a valid one is provided
while True:
    try:
        # If filename not specified as commandline argument, ask user for filename
        if '.docx' not in str(sys.argv):
            SelectedFile = input("Please enter filename including extension (case sensitive): ")
        # Use the specified commandline argument as filename
        else:
            SelectedFile = str(sys.argv[1])
            print(f"You selected: {SelectedFile}. \n\nSearching...\n")

        # Open the specified file as "StoredContents"
        StoredContents = docx2txt.process(SelectedFile)

        # Use regex findall to grab email addresses from "StoredContents"
        emails = re.findall("([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)", StoredContents)

        # Put the list through a dictionary to remove duplicates
        emails = list(dict.fromkeys(emails))

        # Show how many and what email addresses were found
        print(f'\n{len(emails)} emails found: \n')
        for email in emails:
            print(email)

        # Ask user if they would like contents copied to clipboard
        CopyToClipboard = input("\nCopy to clipboard? (y/n) ")

        # Use Xerox to copy to clipboard (convert to string first using .join)
        if CopyToClipboard.lower() == 'y' or 'yes' in CopyToClipboard.lower() or 'yup' in CopyToClipboard.lower():
            clippy = '\n'.join(emails)
            xerox.copy(clippy)
            print("List of emails copied to clipboard.\n")
        else:
            pass

        # Ask user if they would like to save results to file
        SaveToFile = input("Save list to a text file? (y/n) ")

        # If the user chose 'y', write to file
        if SaveToFile.lower() == 'y' or 'yes' in SaveToFile.lower() or 'yup' in SaveToFile.lower():
            TextFile = open('DOCscraped.txt','w')
            for email in emails:
                TextFile.write(email)
                TextFile.write('\n')
            TextFile.close()
            print("List of emails saved to DOCscraped.txt")
        else:
            break
        break

    except:
        system('clear')
        print("That didn't seem to be a valid DOCX file, please try again.\n")
        continue
    else:
        break
