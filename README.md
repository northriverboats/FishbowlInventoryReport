# FishbowlInventoryReport
Fishbowl Inventory Valuation Summary Report for the Main Warehouse as an Excel spreadsheet

## Requirements
Any modern version of Python3.6+ should do. At the time of writing the current version is Python 3.11. The infomration and credentials for read only database access to the fishbowl server is also required.

## Setting up the development environment
```
git clone git@github.com:northriverboats/FishbowlInventoryReport.git
cd FishbowlInventoryReport

"\Program Files\python311\python" -m venv venv
venv\Scripts\activate
python -m pip install --upgrade pip
pip install wheel
pip install -r requirements.txt
```

## Doing Development Work
```
cd FishbowlInventoryReport

"\Program Files\python311\python" -m venv venv
venv\Scripts\activate

# edit and run code as needed
```

## Creating an Executable
Use:
```
pyinstaller.exe fishbowlinventoryreport.spec
```

Initial setup can be done with:
```
pyinstaller.exe --onefile --noconsole --add-data ".env;." -i Report.ico fishbowlinventoryreport.py
```