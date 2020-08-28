file: word2excel/ReadMe.txt
NOTE: This script and whole package is WIP


####
If you download the ZIP file and it only contains README, Requirements and word2excel, follow the following steps to prepare your coding environment

Preparation:
1. Be in Ubuntu and have Python3
2. Install python-venv package
    $ sudo apt install python3-venv
3. Create virtual environment (requires venv)
    $ python3 -m venv word2excel

To Activate virtual environment (requires venv):
$ source word2excel/bin/activate

To Install Requirements (requires pip):
$ pip install -r requirements.txt

To Run:
$ python3 word2excel.py <filename>

Example:
$ python3 word2excel.py contracts.docx

Output will be named output.xls