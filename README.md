
# Excel file to CSV transformer for CloudTalk
Importing contacts into CloudTalk from Excel isn't available. However, it's possible to [import from a CSV file](https://www.cloudtalk.io/contact-import/).

This project helps to generate CSV files from Excel (even if they are password-protected) compatible with CloudTalk.

The project was sponsored by [**Deidis** - a solution provider of CloudTalk](https://www.deidis.com/solutions/cloudtalk).

## How it works
When you run the transformer (see the installation instructions below) it will search for `.xlsx` files in the same folder and apply the column mappings defined in the ***config.txt*** file.

Follow the progress and instructions in the terminal window. The resulting `.csv` file(s) will be placed in the ***outputs*** folder. You can then import these files into CloudTalk, [here are the instructions](https://help.cloudtalk.io/en/articles/3422012-how-to-import-contacts-to-cloudtalk-from-a-file).

If the Excel file has multiple sheets, each of them will become a separate .csv file.

Try it step by step:

- Install the transformer (follow the instructions below)
- Put the ***test.xlsx*** file inside the installation folder (or any other .xlsx file)
- Run the transformer
- You should now have ***outputs/test-sheet1.csv*** and **outputs/test-sheet2.csv**

### config.txt
This file allows to map your Excel file columns to CloudTalk columns.  Instructions are inside the file.

## Installation for Windows
Download the `windows-cloudtalk-excel-to-csv.zip` file and extract it on your computer. It will have two files:

- ***transformer.exe***
- ***config.txt***

**To run the transformer simply double click `transformer.exe`**.

## Setup for MacOs / Linux
Currently no packaged version, simply use python. Step by step:

- Create a folder ***cloudtalk-excel-to-csv***
- Download ***transformer.py*** and ***config.txt*** to that folder
- Open the Terminal and cd to *cloudtalk-excel-to-csv*
- Run `...% python3 transformer.py`
- If there are dependencies missing, you'll need these:
    - `...% pip3 install openpyxl`
    - `...% pip3 install tabulate`
    - `...% pip3 install msoffcrypto` (for password-protected Excel files)

**To run the transformer simply run `...% python3 transformer.py` on the command line.**

# 
The project was sponsored by [**Deidis** - a solution provider of CloudTalk](https://www.deidis.com/solutions/cloudtalk).
