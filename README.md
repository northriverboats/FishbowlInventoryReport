# FishbowlInventoryReport
Fishbowl Inventory Valuation Summary Report for the Main Warehouse as an Excel spreadsheet

## Requirements
Any modern version of Python3.6+ should do. At the time of writing the current version is Python 3.11. The infomration and credentials for read only database access to the fishbowl server is also required. 

Also the Firebird 2.5.x client driver `fbclient.dll` is requered.

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


## Installing the fbclient.dll driver
The `fblcient.dll` is required for this program to run
1. The fishbowl server software installs this driver.
2. Versons `fbclient32.dll` and `fbclient64.dll` do not work.
3. No need to install the server version, just place the `fbclient.dll` in `C:\Prgram Files\Fishbowl\database\bin` folder and install the `Firebird.reg` file.
4. If this program crahses and keeps `fbclient.dll` open, if you run Fishbowl you will get an error that the Fisbhowl  server can not be ound.
5. If Fisbhbowl is running, then this script will fail to run properly.

The registry key that tells `fdb` where to find `fbclient.dll`:
```
SOFTWARE\\Firebird Project\\Firebird Server\\Instances
```

The contents of the `Firebird.reg` file:
```
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\Firebird Project\Firebird Server\Instances]
"DefaultInstance"="C:\\Program Files\\Fishbowl\\database\\"

```