"""XLS 2 ANKI (TXT)

Python script that:
- Takes as a first input an XLS file that contains a table with an Anki deck
- It produces an Anki-ready TXT as an output

- The contents of the XLS file have to be transformed to TXT, noting that, in the XLS file:
--> the first row is to be omitted 
--> rows 2 to 5 are the metadata rows, they should be inserted first in the TXT file
--> row 6 is to be omitted 
--> rows 7 until the end of the file have to be included in the TXT in the following order:
----> first, the cell at column 1 (Anki GUID)
----> followed by the cell at column 3 (source language)
----> followed by the cell at column 2 (language being learned)
----> followed by the cell at column 4 (Anki Tags)
----> the separator to be used is the tab space

- Syntax:
./xls2anki.pyw "$$$THREE_LETTER_IDENTIFIER_OF_LANGUAGE$$$"

Author: jlnkls
"""

import pandas as pd
import sys
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
import csv
import random
import string


def generate_random_string():
    ''' Generates a random string of length 10 (alphanumeric + chars - quotation marks) '''

    # Chars
    chars = string.ascii_letters + string.digits + string.punctuation.replace('"', '').replace("'", "")

    # Random string of chars
    return ''.join(random.choice(chars) for _ in range(10))


def fill_empty_guid_cells(data):
    ''' Fill GUID cells with a new randomly generated GUID if cell is empty'''

    # Init used GUIDs
    used_guids = set()

    # Loop through dataframe
    for i in range(len(data)):
        # Cell is empty
        if data.iloc[i, 0] == '' or pd.isnull(data.iloc[i, 0]):
            # Generate random guid string until it is new
            while True:
                random_string = generate_random_string()
                if random_string not in used_guids:
                    data.iloc[i, 0] = random_string
                    used_guids.add(random_string)
                    break
        # Add existing GUID to used GUIDs
        else:
            used_guids.add(data.iloc[i, 0])


def xls2anki(path, vocab_list_name):
    ''' Produce an Anki-ready TXT file of a deck spreadsheet '''

    # Load the XLS file
    data = pd.read_excel(path + vocab_list_name + ".xlsm", header=None)

    # Extract metadata rows
    metadata = data.iloc[1:5, :]

    # Extract data rows and rearrange columns
    data = data.iloc[6:, [0, 2, 1, 3]]

    # Reset the column index to maintain the rearranged columns
    data.columns = range(data.shape[1])

    # Add GUIDs to notes that lack them
    fill_empty_guid_cells(data)

    # Escape hashes in GUIDs
    for i in range(len(data)):
        if '#' in str(data.iloc[i, 0]):
            data.iloc[i, 0] = '"' + str(data.iloc[i, 0]) + '"'

    # Combine metadata and data
    result = pd.concat([metadata, data])

    # Check and replace values starting with " =" (un-escape space-equal-sign combinations)
    result = result.applymap(lambda x: x.lstrip() if isinstance(x, str) and x.startswith(" =") else x)

    # Save the result to a TXT file with tab-separated values
    result.to_csv(path + vocab_list_name + ".txt", sep='\t', index=False, header=False, quoting=csv.QUOTE_NONE)




def main():
    # Root dir
    root_dir = "$$$ADD_YOUR_ROOT_DIR$$$/"

    # Checking language to process
    # Example languages provided
    if (len(sys.argv) < 2):
        exit()
    else:
        if (sys.argv[1] in "EUS"):
            lang_name = "Euskara/"
            vocab_list_name = "Hiztegia"
        elif (sys.argv[1] in "KR"):
            lang_name = "Hangugeo/"
            vocab_list_name = "eohwi"
        else:
            lang_name = "Suomi/"
            vocab_list_name = "Sanasto"

    lang_name += "Vocabulary/"

    # Path
    path = root_dir + lang_name

    # Anki TXT to XLS
    xls2anki(path, vocab_list_name)


# Main
main()