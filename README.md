# Excel Comparator Tool
This tool is useful for comparing column elements from multiple Excel Sheet files.

# Background
I built this mini project during my internship at a Power Electronics company. When converting circuit schematics to actual hardware, the software you would use to design the schematic typically generate a Bill of Materials (BOM). These parts would then have to be ordered in stores like DigiKey. In my internship, I noticed that they would manually compare the data from the BOM sheet versus the DigiKey sheet to ensure that the parts being purchased were correct. This comparator is to automate that process.

# How To Use
1. Clone the repository 
```
git clone https://github.com/vulpix-repo/Excel-Comparator-Tool
```
2. Install dependencies
```
pip install requirements.txt
```
3. Run the tool 
```
python main.py
```
4. Once the GUI pops up, add the Excel files to compare. Select the reference column in both parts and the column to compare. Click on the compare button to begin. The results will show up and be saved in a new Excel sheet in the same directory as the uploaded sheets.

# Comparison Logic
The comparison mimics that of the VLOOKUP function in many spreadsheet applications. The tool uses the reference column as an ID, looks for that ID in both sheets, and compares the corresponding data in the comparison column set. It supports fuzzy matching to allow minimal typos and errors when doing the comparison.
