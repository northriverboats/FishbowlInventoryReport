import os
import sys
from datetime import date, timedelta
from pathlib import Path
from PyQt6.QtWidgets import QApplication, QWidget, QMessageBox
from fishbowlinventoryreport_ui import Ui_Dialog
from decimal import Decimal
from datetime import datetime
from datetime import timedelta
from platform import system
import fdb
import itertools
from dotenv import load_dotenv
from excelopen import ExcelOpenDocument


FIELDS = [
    'Location',
    'Part',
    'Description',
    'Qty',
    'UOM',
    'Cost',
    'Extended',
]

FORMATS = [
    'General',
    'General',
    'General',
    '0.00',
    'General',
    r'[$$-409]#,##0.00;[RED]\-[$$-409]#,##0.00',
    r'[$$-409]#,##0.00;[RED]\-[$$-409]#,##0.00',
]

WIDTHS = [
    16.25,
    34.25,
    80.50,
    7.50,
    6.50,
    10,
    12.75,
]

def resource_path(relative_path):                                                                                                                                                          
    """ Get absolute path to resource, works for dev and for PyInstaller """                                                                                                               
    try:                                                                                                                                                                                   
        base_path = sys._MEIPASS                                                                                                                                                           
    except Exception:                                                                        
        base_path = os.path.abspath(".")                                                                                                                                                   
                                                                                                                                                                                           
    return os.path.join(base_path, relative_path)  

def filter_nonprintable(text):
    nonprintable = itertools.chain(range(0x00,0x20),range(0x7f,0xa0))
    # Use characters of control category
    # Use translate to remove all non-printable characters
    return text.translate({character:None for character in nonprintable})

def year_quarter(month, year):
    """Compute Corret Year and quarter for month/year"""
    quarter = 4
    if month < 3:
        year = year -1
    elif month < 6:
        quarter = 1
    elif month < 9:
        quarter = 2
    elif month < 12:
        quarter = 3
    return f"{year} Q{quarter}"

def default_filename():
    """Return recommended filename without extension"""
    current = date.today()
    quarter = year_quarter(current.month, current.year)
    return f"Inventory Valuation Summary {quarter}"

def read_firebird_database(include, exclude):
    """Create Inventory Value Summary from Fishbowl"""
    stock = []
    con = fdb.connect(
        host=os.getenv('HOST'),
        database=os.getenv('DATABASE'),
        user=os.getenv('USER'),
        password=os.getenv("PASSWORD"),
        charset='WIN1252'
    )

    select = """
    SELECT locationGroup.name AS "Group",
        COALESCE(partcost.avgcost, 0) AS averageunitcost,
        COALESCE(part.stdcost, 0) AS standardunitcost,
        locationgroup.name AS locationgroup,
        part.num AS partnumber,
        part.description AS partdescription,
        location.name AS location, asaccount.name AS inventoryaccount,
        uom.code AS uomcode, sum(tag.qty) AS qty, company.name AS company
    FROM part
        INNER JOIN partcost ON part.id = partcost.partid
        INNER JOIN tag ON part.id = tag.partid
        INNER JOIN location ON tag.locationid = location.id
        INNER JOIN locationgroup ON location.locationgroupid = locationgroup.id
        LEFT JOIN asaccount ON part.inventoryaccountid = asaccount.id
        LEFT JOIN uom ON uom.id = part.uomid
        JOIN company ON company.id = 1
    WHERE locationgroup.id IN (1)
    GROUP BY averageunitcost, standardunitcost, locationgroup, partnumber,
        partdescription, location, inventoryaccount, uomcode, company
    """

    cur = con.cursor()
    cur.execute(select)

    for (group, avgcost, stdcost, locationgroup, partnum, partdescription,
         location, invaccount, uom, qty, company) in cur:
        if location in exclude:
            continue
        if include and location not in include:
            continue
        stock.append([
            location,
            partnum,
            partdescription,
            str(Decimal(str(qty)).quantize(Decimal("1.00"))),
            uom,
            str(Decimal(str(avgcost)).quantize(Decimal("1.00"))),
        ])

    stock = sorted(stock, key=lambda k: (k[0], k[1]))
    return stock

def write_xlsx_file(rows, file_name):
    """write inventoty data to xlsx sheet"""
    excel = ExcelOpenDocument()
    excel.new(file_name)
    title_font = excel.font(name='Arial', size=10, bold=True)
    body_font = excel.font(name='Arial', size=10)

    for column, value in enumerate(FIELDS, start=1):
        excel.cell(row=1, column=column).value = value
        excel.cell(row=1, column=column).font = title_font

    for column, width in enumerate(WIDTHS, start=65):
        excel.set_width(chr(column), width)

    for row, all_fields in enumerate(rows, start=2):
        # remove ascii 0x00 - 0x1F
        all_fields[1] = filter_nonprintable(all_fields[1])
        all_fields[2] = filter_nonprintable(all_fields[2])
        for column, field in enumerate(all_fields, start=1):
            if FORMATS[column-1] == 'General':
                value = field
            else:
                value = float(field.replace(",", ""))
            cell = excel.cell(row=row, column=column)
            cell.value = value
            cell.number_format = FORMATS[column-1]
            cell.font = body_font

        excel.cell(row=row, column=7).value = "=SUM(D{}*F{})".format(row, row)
        excel.cell(row=row, column=7).font = body_font
        excel.cell(row=row, column=7).number_format = FORMATS[6]

    row = excel.max_row() + 2
    excel.cell(row=row, column=5).value = 'Grand Total:'
    excel.cell(row=row, column=5).font = title_font
    excel.cell(row=row, column=7).value = "=SUM(G2:G{})".format(row - 2)
    excel.cell(row=row, column=7).font = title_font
    excel.cell(row=row, column=7).number_format = FORMATS[6]

    excel.save()



class AppDialog(QWidget):
    def __init__(self):
        super().__init__()

        # desktop = Path.home() / 'Desktop'
        # full_path = desktop / default_filename()


        self.ui = Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.lineEdit.setText(default_filename())

        self.show()
    
    def accept(self):
        root = self.ui.lineEdit.text()
        file_name = Path.home() / 'Desktop'  / f"{root}.xlsx"
        if file_name.exists():
            button = QMessageBox.question(self, "File Already Exists", f"Do you want to overwrite {root}?")
            if button == QMessageBox.StandardButton.No.value:
                return
        # call create spreadsheet
        # exit program
        include = ""
        exclude = ""
        rows = read_firebird_database(include, exclude)
        write_xlsx_file(rows, str(file_name))
        print(file_name)
        self.close()


    def reject(self):
        print('Closing')
        self.close()


def main():
    load_dotenv(dotenv_path=resource_path(".env"))
    try:
        app = QApplication(sys.argv)
        app_dialog = AppDialog()
    except Exception as error:
        QMessageBox.critical(None, "Error", error)
    
    sys.exit(app.exec())


if __name__ == '__main__':
    main()