# Adobe Sign Form Tool
## Installation
1. Install [Python 3](https://www.python.org/downloads/)
2. Clone the github repository.
3. Create an empty folder called venv
4. Create a virtual environment by running `python3 -m venv .\venv`
5. Activate the venv by running `.\venv\Scripts\activate`
5. Install all packages required. Run `pip3 install -r requirements.txt`
6. Run `main.py` manually or by running `RUN.bat`.
7. Enter directory for excel file and default import directory. Use the downloads folder for the default import directory.
7. Make sure that no packages are missing. Install if neccesary.

# Instructions
## New Forms
1. Delegate and sign the form.
2. View the document in the adobe sign account.
3. Click *Download Form Field Data*. Make sure it is saved to the default import location (downloads) or you know the directory where it will be saved to.
![](https://github.com/calpoly-ee-tsg/adobesignformtool/blob/main/pic1.png)
4. Download as many agreements aas you want. It will import them all.
4. Run the python script and select option 1 (Import)
5. Press enter when it prompts you for the default import location
6. Wait for the excel sheet to open. 
7. Press Yes when it asks to repair.
8. Check it and extend the table all the way to the bottom.
9. Save it with the exact same filename that it had before.
## Returned Equipment
1. Use Excel search filters to find the person's equipment by date or name. Nice!
2. Select the last column (Returned).
3. Press `Ctrl-;` to auto type the current date (or simply type today's date).
4. Save the form
