import encodings.idna

import json
import gspread

import xlsxwriter
import rfc3339

from xlsx import Workbook

#from xml.sax.saxutils import escape 
from xml.dom.minidom import Text, Element

import xml.etree.ElementTree as ET

import os
import sys
import winreg
import datetime
import codecs

import teparser

def export(projectName):
      
    print (projectName)
      
    key = "Software\\Mengine\\%s"%(projectName)
      
    try :
        root = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key)        
    except EnvironmentError :
        Error("project %s Variables %s not in Registry!"%(projectName, key))
        return
        pass
        
    try:
        MENGINE_PROJECT_PATH = winreg.QueryValueEx(root, "MENGINE_PROJECT_PATH")
        ProjectPath = MENGINE_PROJECT_PATH[0]
    except EnvironmentError :
        Error("project %s MENGINE_PARAMS_PATH not in Registry!"%(projectName))
        return
        pass          
        
    try:
        MENGINE_PARAMS_PATH = winreg.QueryValueEx(root, "MENGINE_PARAMS_PATH")
        ParamsPath = MENGINE_PARAMS_PATH[0]
    except EnvironmentError :
        Error("project %s MENGINE_PARAMS_PATH not in Registry!"%(projectName))
        return
        pass
    
    exportTextEncounterDataFromDocx(projectName, ProjectPath, ParamsPath, "Locale_RU", "Resources", "Scripts/TextEncounter", "PathTextEncounter.xlsx", createTextEncounter)

    exportData(projectName, ProjectPath, ParamsPath, "Database", "Resources", "Scripts/Database", "PathDatabase.xlsx", createDatabase2)
    exportData(projectName, ProjectPath, ParamsPath, "Database", "Resources", "Scripts/Database", "PathDatabaseORM.xlsx", createDatabaseORM)    
    exportData(projectName, ProjectPath, ParamsPath, "Scenario", "Resources", "Scripts/Scenario", "PathScenario.xlsx", createScenario)
    
    #exportGoogleSpreadsheet(projectName, ProjectPath, ParamsPath, "Database", "Resources", "Scripts/Database", "PathGoogleSpreadsheets.xlsx", createDatabaseORM)
    
    localeKit(projectName, ProjectPath, ParamsPath, "PathLocaleKit.xlsx")
    pass
   
    
def modification_date(filename):
    if os.path.exists(filename) is False:
        return None

    t = os.path.getmtime(filename)
    return datetime.datetime.fromtimestamp(t)
    pass    
    
    
def getXlsxData(filename):
    workbook = Workbook(filename)
    
    rowsList = []
    rows = workbook[1].rows()
    row_items = rows.items()
    #row_items.sort()

    for row, cells in row_items:
        cellsList = []
        length = len(cells)

        for i, cell in enumerate(cells):
            if cell.value is None:
                continue
                pass
        
            cell_value = cell.value.strip()
            
            if cell_value != "":
                cellsList.append(cell_value)
                pass
            pass

        rowsList.append(cellsList)
        pass

    return rowsList
    pass
    
def getXlsxSheetList(filename):
    workbook = Workbook(filename)
    
    sheetList = []
    for sheet in workbook:        
        rows = sheet.rows()
        row_items = rows.items()

        rowsList = []
        for row, cells in row_items:
            cellsList = []
            length = len(cells)

            for i, cell in enumerate(cells):
                cellsList.append(cell)
                pass

            i = 0
            for cell in cellsList:
                value = cell.value.strip(' ')
                i += len(value)
                pass
                
            if i == 0:
                continue
                pass

            rowsList.append(cellsList)
            pass
            
        sheetList.append((sheet.name, rowsList))
        pass

    return sheetList
    pass

def getSheetText(sheetList):
    Text = {}
    for name, sheetRows in sheetList:
        legends = sheetRows[0]
        
        def __findCellValue(row, name):
            for cell in row:
                value = cell.value.strip(' ')
                if value == name:
                    return cell
                    pass
                pass
                
            return None
            pass
        
        def __findCellColumn(row, column):
            for cell in row:
                if cell.column == column:
                    return cell
                    pass
                pass
                
            return None
            pass            
            
        CellKey = __findCellValue(legends, "Key")
        
        if CellKey is None:
            Error ("invalid find legend 'Key' in sheet '%s'"%(name))       
            return None
            pass
        
        CellValue = __findCellValue(legends, "Value")
        
        if CellValue is None:
            Error ("invalid find legend 'Value' in sheet '%s'"%(name))       
            return None
            pass        
            
        for i, row in enumerate(sheetRows[1:]):            
            Key = __findCellColumn(row, CellKey.column)
            Value = __findCellColumn(row, CellValue.column)
            
            if Key is None and Value is None:
                continue
                pass
                
            if Key is None:
                Error ("Sheet '%s' index %s KeyIndex %s ValueIndex %s not found Key [Value %s]"%(name, i, CellKey.column, CellValue.column, Value.value))
                continue
                pass

            if Value is None:
                Error ("Sheet '%s' index %s KeyIndex %s ValueIndex %s not found Value [Key %s]"%(name, i, CellKey.column, CellValue.column, Key.value))
                continue
                pass
                
            KeyValue = Key.value.strip(' ')
            ValueValue = Value.value.strip(' ')
            
            if len(KeyValue) == 0 and len(ValueValue) == 0:
                continue
                pass
                
            if len(KeyValue) == 0:
                Error ("index %s KeyIndex %s ValueIndex %s key is empty [Value %s]"%(i, CellKey.column, CellValue.column, Value.value))
                continue
                pass

            if len(ValueValue) == 0:
                Error ("index %s KeyIndex %s ValueIndex %s value is empty [Key %s]"%(i, CellKey.column, CellValue.column, Key.value))
                continue
                pass
                
            if KeyValue in Text:
                Error ("index %s KeyIndex %s ValueIndex %s key %s alredy exist! [prev value %s]"%(i, CellKey.column, CellValue.column, Key.value, Text[KeyValue]))
                continue
                pass
                
            ValueValue = ValueValue.replace("%SPACE%", " ")
                
            Text[KeyValue] = ValueValue
            pass
        pass
        
    return Text
    pass      
    
def getPrintValue(value):
    printValue = value
    
    if len(value) == 0:
        printValue = "None"
        pass
    elif value.upper() == "TRUE":
        printValue = "True"
        pass
    elif value.upper() == "FALSE": 
        printValue = "False"
        pass
    else:        
        try:
            float(value)
        except ValueError:
            printValue = "\"%s\""%(value)
            pass
        pass

    return printValue
    pass
    
def getPrintDictValue(d):
    p = ""
    comma = False
    for key, value in d.items():
        if comma is True:
            p += ", "
            pass
        else:
            comma = True
            pass                
            
        p += "%s = "%(key)
        
        if isinstance(value, dict) is False:
            p += "%s"%(value)
            pass
        else:
            pd = getPrintDictValue(value)
        
            p += "dict("            
            p += pd
            p += ")"
            pass
        pass
        
    return p    
    pass    
    
def writeValue(cell_value, cell_index, begin_index):
    result = ""

    if cell_index > begin_index:
        result += ", "
        pass

    result += getPrintValue(cell_value)

    return result
    pass
    
def getCellIndex(column):
    index = 0
    for i, c in enumerate(column):
        index += i * 26 + (ord(c) - ord('A'))
        pass
        
    return index
    pass

def findCell(index, bb_cell, cells):
    for cell in cells:
        cell_index = getCellIndex(cell.column)

        if index == cell_index:
            return cell
        pass

    return None
    pass      

def localeKit(projectName, ProjectPath, ParamsPath, LocaleKitPath):
    LocaleKitFullPath = ParamsPath + LocaleKitPath
    
    if os.path.isfile(LocaleKitFullPath) is False:
        Error ("localeKit %s file %s not exist in path '%s'"%(projectName, LocaleKitPath, ParamsPath))
        return
        pass
        
    resourceList = getXlsxData(LocaleKitFullPath)
    
    for row in resourceList:        
        if len(row) == 0:
            continue
            pass
        
        if len(row) < 2:
            Error ("Locale Kit %s invalid data for %s!"%(xlsxFileName, xlsx_name))
            continue
            pass
            
        xlsx_name = row[0]            
        locale_name = row[1]
        
        if len(row) >= 3:
            enable = row[2]
            
            if int(enable) == 0:
                continue
                pass
            pass
            
        xlsxFilePath = "%s%s"%(ParamsPath, xlsx_name)         
        
        if os.path.isfile(xlsxFilePath) is False:
            Error ("File with this path %s dont exist!"%(xlsxFilePath))
            continue
            pass
            
        
        
        xmlPath = "%s%s"%(ProjectPath, locale_name)
         
        print ("%s - LocaleKit export!"%(xlsxFilePath))
            
        sheetList = getXlsxSheetList(xlsxFilePath)
        
        if sheetList is None:
            return
            pass
        
        Text = getSheetText(sheetList)
        
        if Text is None:                
            return
            pass
        
        tree = None
        
        try:
            tree = ET.parse(xmlPath)
        except ET.ParseError as pe:
            Error ("%s - LocaleKit invalid export xml %s have error: %s"%(xlsxFilePath, xmlPath, pe))
            pass
            
        if tree is None:
            return
            pass
            
        root = tree.getroot()
        
        MissLocale = []
        for child in root:
            Key = child.attrib["Key"]
            if Key not in Text:
                MissLocale.append(Key)
                continue
                pass
              
                
            Value = Text.pop(Key)
            
            child.attrib["Value"] = Value
            pass
        
        if len(MissLocale) != 0:
            Error ("-----------------------------------------------------------------------")
            Error ("Miss Locale: ")
            for key in MissLocale:
                Error (key)
                pass
            Error ("-----------------------------------------------------------------------")
            pass
        
        if len(Text) != 0:
            Error ("-----------------------------------------------------------------------")
            Error ("Over Locale: ")
            for key in Text:
                Error (key)
                pass
            Error ("-----------------------------------------------------------------------")
            pass
            
        tree.write(xmlPath, encoding="utf-8")
        pass    
    pass

def exportGoogleSpreadsheet(projectName, ProjectPath, ParamsPath, DataType, ResourceDir, DataDir, DataFile, DataCreate):
    xlsPath = ParamsPath    
    xlsxFileName = DataFile

    resoureces = xlsPath + xlsxFileName
    
    if os.path.isfile(resoureces) is False:
        Error ("project %s file %s not exist in path '%s'"%(projectName, xlsxFileName, xlsPath))
        return
        pass
    
    resourceList = getXlsxData(resoureces)
        
    for row in resourceList:
        if len(row) == 0:
            continue
            pass
    
        pyPath = None
        if len(row) == 4:
            subproject_name = row[3]
            pyPath = ProjectPath + subproject_name + "/" + DataDir + "_" + subproject_name + "/"
        else:
            pyPath = ProjectPath + ResourceDir + "/" + DataDir + "/"
            pass
            
        directory = os.path.dirname(pyPath)
        if not os.path.exists(directory): 
            os.makedirs(directory)
            pass
            
        if not os.path.exists(pyPath + "__init__.py"):
            f = open(pyPath + "__init__.py", "w")
            f.write("#database")
            f.close()
            pass 
            
        docid = row[0]       
        username = row[1]
        password = row[2]
        
        json_key = None
        with open("Burritos.json") as json_burritos:
            json_key = json.load(json_burritos)
            pass
        
        scope = ['https://spreadsheets.google.com/feeds']

        credentials = SignedJwtAssertionCredentials(json_key['client_email'], bytes(json_key['private_key'], 'utf-8'), scope)
        
        client = gspread.authorize(credentials)
       
        #client = gspread.login(username, password)

        spreadsheet = client.open_by_key(docid)
        
        spreadsheet_title = spreadsheet.title
        
        GoogleSpreadsheetsCacheDir = "%sGoogleSpreadsheetsCache/"%(xlsPath)
        if not os.path.exists(GoogleSpreadsheetsCacheDir):
            os.makedirs(GoogleSpreadsheetsCacheDir)
            pass
            
        GoogleSpreadsheetsCacheTitleDir = "%sGoogleSpreadsheetsCache/%s/"%(xlsPath, spreadsheet_title)
        if not os.path.exists(GoogleSpreadsheetsCacheTitleDir):
            os.makedirs(GoogleSpreadsheetsCacheTitleDir)
            pass        
        
        for worksheet in spreadsheet.worksheets():       
            filename = "%s%s.xlsx"%(GoogleSpreadsheetsCacheTitleDir, worksheet.title)
            date_time = rfc3339.parse_datetime(worksheet.updated)
            # date_time.
            if os.path.exists(filename) :
                date_of_xls = datetime.datetime.fromtimestamp(os.path.getmtime(filename), date_time.tzinfo)
                if date_of_xls > date_time:
                    print ("file %s" % (filename), "Last Revision")
                    continue
                    pass
                pass
            pass
            
            workbook = xlsxwriter.Workbook(filename)
            excel_sheet = workbook.add_worksheet()
            vals = worksheet.get_all_values()

            for r, row in enumerate(vals):
                for c, cell in enumerate(row):
                    excel_sheet.write(r, c, cell)
                    pass
                pass

            workbook.close()            
            
            py_name = os.path.splitext(os.path.basename(filename))[0]
            py_file_name = "%s%s%s.py"%(pyPath, DataType, py_name)            
            
            print (filename, " - %s export!"%(DataType))
            DataCreate(filename, py_file_name, py_name)
            pass            
        pass    
    pass    
       
def exportData(projectName, ProjectPath, ParamsPath, DataType, ResourceDir, DataDir, DataFile, DataCreate):
    xlsPath = ParamsPath    
    xlsxFileName = DataFile

    resoureces = xlsPath + xlsxFileName
    
    if os.path.isfile(resoureces) is False:
        Error ("project %s file %s not exist in path '%s'"%(projectName, xlsxFileName, xlsPath))
        return
        pass
    
    resourceList = getXlsxData(resoureces)
        
    for row in resourceList:
        if len(row) == 0:
            continue
            pass
    
        xmlPath = None
        if len(row) == 3:
            subproject_name = row[2]            
            xmlPath = ProjectPath + subproject_name + "/" + DataDir + "_" + subproject_name + "/"
        else:
            xmlPath = ProjectPath + ResourceDir + "/" + DataDir + "/"
            pass
            
        directory = os.path.dirname(xmlPath)
        if not os.path.exists(directory): 
            os.makedirs(directory)
            pass
            
        if not os.path.exists(xmlPath + "__init__.py"):
            f = open(xmlPath + "__init__.py", "w")
            f.write("#database")
            f.close()
            pass 
       
        xlsx_path = row[0]
        
        xlsxFileName = "%s%s.xlsx"%(xlsPath, xlsx_path)
        date_xlsx = modification_date(xlsxFileName)
        
        if date_xlsx is None:
            Error ("File with this path %s dont exist!"%(xlsxFileName))
            continue
            pass
            
        if len(row) == 1:
            py_name = os.path.basename(row[0])
            pass
        else:
            py_name = os.path.basename(row[1])
            pass

        xmlFileName = "%s%s%s.py"%(xmlPath, DataType, py_name)
        date_xml = modification_date(xmlFileName)
        
        if date_xml is None or date_xml < date_xlsx:
            print (xlsxFileName, " - %s export!"%(DataType))
            DataCreate(xlsxFileName, xmlFileName, py_name)
            pass
        pass
    pass
     
def createDatabase2(xlsxFilename, paramsFilename, name):
    workbook = Workbook(xlsxFilename)
    
    rows = workbook[1].rows()
    
    if len(rows) == 0:
        for page in workbook:
            pagerows = page.rows()
            pagerows_items = pagerows.items()
            if len(pagerows_items) != 0:
                msg = "xlsx %s empty first page %s but have not empty other page %s"%(xlsxFilename, workbook[1].id, page.id)
                print(msg)
                return
                pass
            pass
        pass
    
    row_items = rows.items()

    indent = "    "
    
    result = ""
    
    result += "from GOAP2.Database import Database\n"
    result += "\n"
    result += "class Database%s(Database):\n"%(name)
    result += "    def __init__(self):\n"
    result += "        super(Database%s, self).__init__()\n"%(name)
    
    def makeRecordsValueORM(bb_cell, bb_ord, cells, legend_cells):
        DatabaseParams = []
        
        for legend in legend_cells:
            true_cell = [cell for cell in cells if cell.column == legend.column]
            
            legend_value = legend.value.strip()
                
            if legend_value[0] == "[":
                if len(true_cell) == 0:
                    params = "[]"
                
                    DatabaseParams.append(params)
                    continue
                    pass
                    
                true_cell = true_cell[0]
                index2 = cells.index(true_cell)
                for_cells = cells[index2:]
                
                printValue = ""

                for_comma = False

                for for_cell in for_cells:
                    for_cell_value = for_cell.value.strip()
                    
                    if for_cell_value == "":
                        break
                        pass
                        
                    if for_comma is True:
                        printValue += ", "
                        pass
                    else:
                        for_comma = True
                        pass

                    printValue += getPrintValue(for_cell_value)
                    pass
                
                if printValue != "":
                    params = "%s = [%s]"%(legend_value[1:-1], printValue)
                
                    DatabaseParams.append(params)
                    pass
                break
                pass
            else:
                cell = None if len(true_cell) == 0 else true_cell[0]
                
                if cell is None:
                    continue
                    pass
                    
                if cell.value is None:
                    continue
                    pass
                    
                cell_value = cell.value.strip()
                
                if cell.value == "":
                    # Empty string
                    continue
                    pass                
                
                printValue = getPrintValue(cell_value)
                
                params = "%s = %s"%(legend_value, printValue)
                
                DatabaseParams.append(params)
                pass
            pass

        if len(DatabaseParams) == 0:            
            return None
            pass
            
        printRecords = ""
        
        comma = False
        for value in DatabaseParams:
            if comma is True:
                printRecords += ", "
                pass
            else:
                comma = True
                pass

            printRecords += value
            pass
       
        return printRecords
        pass
        
    rows = []
    
    for row_index, row_zip in enumerate(row_items):
        row_empty = True
    
        for cell in row_zip[1]:
            if cell.value is None:
                continue
                pass
        
            cell_value = cell.value.strip()
            
            if len(cell_value) != 0:
                row_empty = False
                pass
            pass
    
        if row_empty is True:
            continue
            pass
    
        rows.append(row_zip)
        pass        
        
    if len(rows) != 0:        
        rows.sort(key = lambda x:x[0])
        
        legend_row, legend_cells = rows.pop(0)
               
        bb_cell = None
        bb_ord = None
        
        for row, cells in rows:
            if bb_cell is None:
                for cell in cells:
                    cell_value = cell.value.strip()
                    
                    if cell_value == "":
                        continue
                        pass

                    bb_ord = cell.column
                    bb_cell = getCellIndex(bb_ord)
                    break
                    pass
                pass        
            pass
            
        legend_cells = [cell for cell in legend_cells if cell.value.strip() != ""]
               
        for row, cells in rows:
            #print("row_index, row %s:%s"%(row_index, row))
            #result += indent
            
            cell_skip = False
            for cell in cells:
                value = cell.value.strip()
                if value != "":
                    if value == "#":
                        cell_skip = True
                        pass
                    else:
                        break
                        pass
                    pass
                pass
                
            if cell_skip is True:
                continue
                pass            
            
            printRecords = makeRecordsValueORM(bb_cell, bb_ord, cells, legend_cells)
                       
            if printRecords is None:
                continue
                pass
                
            result += "        self.addRecord(%s)\n"%(printRecords)            
            pass
        pass

    result += "        pass\n"
    result += "    pass\n"
    
    with codecs.open(paramsFilename, "w", "utf-8") as f:
        f.write(result)
        pass        
    pass      
    
def createDatabaseORM(xlsxFilename, paramsFilename, name):
    workbook = Workbook(xlsxFilename)
    
    rows = workbook[1].rows()
    
    if len(rows) == 0:
        for page in workbook:
            pagerows = page.rows()
            pagerows_items = pagerows.items()
            if len(pagerows_items) != 0:
                msg = "xlsx %s empty first page %s but have not empty other page %s"%(xlsxFilename, workbook[1].id, page.id)
                print(msg)
                return
                pass
            pass
        pass
    
    row_items = rows.items()

    indent = "    "
    
    result = ""
    
    result += "from GOAP2.Database import Database\n"
    result += "\n"
    result += "class Database%s(Database):\n"%(name)
    result += "    def __init__(self):\n"
    result += "        super(Database%s, self).__init__()\n"%(name)
    
    def makeRecordsValueORM(bb_cell, bb_ord, cells, legend_cells):
        DatabaseParams = []
        DatabaseContract = {}         
        DatabaseContractHas = False
        
        for legend in legend_cells:
            true_cell = [cell for cell in cells if cell.column == legend.column]
            
            legend_value = legend.value.strip()
                
            splitKey = legend_value.split('.')
            
            if len(splitKey) > 1:
                DatabaseContractHas = True
                
                if len(true_cell) == 0:
                    continue
                    pass
                    
                keyType = splitKey[0]
                keyResource = splitKey[1]
                
                contract = DatabaseContract
                for key in splitKey[:-1]:
                    if key not in contract:
                        contract[key] = {}
                        pass                                   
                    
                    contract = contract[key]
                    pass
                    
                lastKey = splitKey[-1]
                
                cell = true_cell[0]
                cell_value = cell.value.strip()
                
                if cell_value == "":
                    continue
                    pass
                
                printValue = getPrintValue(cell_value)
                
                contract[lastKey] = printValue
                pass               
            elif legend_value[0] == "[":
                if len(true_cell) == 0:
                    params = "[]"
                
                    DatabaseParams.append(params)
                    continue
                    pass
                    
                true_cell = true_cell[0]
                index2 = cells.index(true_cell)
                for_cells = cells[index2:]
                
                printValue = ""

                for_comma = False

                for for_cell in for_cells:
                    for_cell_value = for_cell.value.strip()
                    
                    if for_cell_value == "":
                        break
                        pass
                        
                    if for_comma is True:
                        printValue += ", "
                        pass
                    else:
                        for_comma = True
                        pass

                    printValue += getPrintValue(for_cell_value)
                    pass

                params = "[%s]"%(printValue)       
                
                DatabaseParams.append(params)
                break
                pass
            elif legend_value[0] == "!":
                cell_value = "" if len(true_cell) == 0 else true_cell[0].value.strip()
                
                #printValue = getPrintValue(cell_value)
                
                if cell_value == "0":
                    params = "False"
                    pass
                else:
                    params = "True"
                    pass
                
                DatabaseParams.append(params)
                pass
            elif legend_value[0] == "?":
                cell_value = "" if len(true_cell) == 0 else true_cell[0].value.strip()
                
                #printValue = getPrintValue(cell_value)
                
                if cell_value == "1":
                    params = "True"
                    pass
                else:
                    params = "False"
                    pass
                
                DatabaseParams.append(params)
                pass 
            else:
                cell_value = "" if len(true_cell) == 0 else true_cell[0].value.strip()
                
                printValue = getPrintValue(cell_value)
                
                params = "%s"%(printValue)
                
                DatabaseParams.append(params)
                pass
            pass

        if len(DatabaseParams) == 0 and len(DatabaseContract) == 0:            
            if DatabaseContractHas is True:
                return "{}"
                pass
        
            return None
            pass
            
        printRecords = ""
        
        comma = False
        for value in DatabaseParams:
            if comma is True:
                printRecords += ", "
                pass
            else:
                comma = True
                pass

            printRecords += value
            pass
            
        if len(DatabaseContract) == 0:
            if DatabaseContractHas is True:
                if len(DatabaseParams) != 0:
                    printRecords += ", "
                    pass            
            
                printRecords += "{}"
                pass
        
            return printRecords
            pass        
            
        if len(DatabaseParams) != 0:
            printRecords += ", "
            pass
                   
        printRecords += getPrintDictValue(DatabaseContract)    
        
        return printRecords
        pass
        
    def makeRecordsKeyORM(legend_cells):
        DatabaseParams = []
        DatabaseContract = []
        
        for cell in legend_cells:
            cell_value = cell.value.strip()      
                
            splitKey = cell_value.split('.')
            
            if len(splitKey) > 1:
                keyType = splitKey[0]
               
                if keyType in DatabaseContract:
                    continue
                    pass
                
                DatabaseContract.append(keyType)
                pass               
            elif cell_value[0] == "[":
                param = cell_value[1:-1]
                
                DatabaseParams.append(param)
                break
                pass
            elif cell_value[0] == "!":
                param = cell_value[1:]
                
                DatabaseParams.append(param)
                pass   
            elif cell_value[0] == "?":
                param = cell_value[1:]
                
                DatabaseParams.append(param)
                pass                  
            else:
                DatabaseParams.append(cell_value)
                pass
            pass

        if len(DatabaseParams) == 0 and len(DatabaseContract) == 0:
            return None
            pass
        
        return [(param, None) for param in DatabaseParams] + [(param, {}) for param in DatabaseContract]
        pass        
        
    rows = []
    
    for row_index, row_zip in enumerate(row_items):
        row_empty = True
    
        for cell in row_zip[1]:
            try:
                cell_value = cell.value.strip()
            except AttributeError as ex:
                print("row %s - %s %s %s invalid error %s"%(row_index, xlsxFilename, paramsFilename, name, cell.value))

                continue
                pass                
            
            if len(cell_value) != 0:
                row_empty = False
                pass
            pass
    
        if row_empty is True:
            continue
            pass
    
        rows.append(row_zip)
        pass        
        
    if len(rows) != 0:        
        rows.sort(key = lambda x:x[0])
        
        legend_row, legend_cells = rows.pop(0)
               
        bb_cell = None
        bb_ord = None
        
        for row, cells in rows:
            if bb_cell is None:
                for cell in cells:
                    cell_value = cell.value.strip()
                    
                    if cell_value == "":
                        continue

                    bb_ord = cell.column
                    bb_cell = getCellIndex(bb_ord)
                    break
                    pass
                pass        
            pass
            
        legend_cells = [cell for cell in legend_cells if cell.value.strip() != ""]
        
        legend_keys = makeRecordsKeyORM(legend_cells)
        
        result += "        class Record%s(object):\n"%(name)
        
        result += "            def __init__(self"   
        for key, default in legend_keys:
            if default is None:
                result += ", %s"%(key)
                pass
            else:
                result += ", %s = %s"%(key, default)
                pass
            pass
        result += "):\n"
        
        for key, default in legend_keys:
            result += "                self.%s = %s\n"%(key, key)
            pass
        
        result += "                pass\n"
        result += "            pass\n"
        result += "\n"
        
        for row, cells in rows:
            #print("row_index, row %s:%s"%(row_index, row))
            #result += indent
            
            cell_skip = False
            for cell in cells:
                value = cell.value.strip()
                if value != "":
                    if value == "#":
                        cell_skip = True
                        pass
                    else:
                        break
                        pass
                    pass
                pass
                
            if cell_skip is True:
                continue
                pass
            
            printRecords = makeRecordsValueORM(bb_cell, bb_ord, cells, legend_cells)
            
            if printRecords is None:
                continue
                pass
                
            result += "        self.addORM(Record%s(%s))\n"%(name, printRecords)
            pass
        pass

    result += "        pass\n"
    result += "    pass\n"
    
    with codecs.open(paramsFilename, "w", "utf-8") as f:
        f.write(result)
        pass        
    pass      
          
def createScenario(xlsxFilename, generateFilename, name):
    workbook = Workbook(xlsxFilename)

    rows = workbook[1].rows()
    row_items = rows.items()

    indent = "    "
    
    result = ""    
    
    result += "from GOAP2.Scenario import Scenario\n"
    result += "\n"
    result += "class Scenario%s(Scenario):\n"%(name)
    result += "    def __init__(self):\n"
    result += "        super(Scenario%s, self).__init__()\n"%(name)
    result += "        pass\n"
    result += "\n"
    result += "    def _onInitialize(self):\n"
            
#    bb_cell = None
#    bb_ord = None
        
    rows = []
    
    for row_index, row_zip in enumerate(row_items):
        rows.append(row_zip)
        pass
        
    if len(rows) != 0:        
        rows.sort(key = lambda x:x[0])
        
        row, cells = rows.pop(0)

        desc_cell = cells[0]
        bb_cell = getCellIndex(desc_cell.column)

        state_0 = "None"
        state_1 = "None"
        state_2 = "None"

        macro_count = None
        macro_id = None

        for row_index, (row, cells) in enumerate(rows):
#            print("row_index, row %s:%s"%(row_index, row))

            printValue = ""

            for cell_index in range(32):
#                print ("cell_index %d"%(cell_index))
                cell = findCell(cell_index, bb_cell, cells)

                cell_value = ""
                if cell is not None:
                    cell_value = cell.value.strip()
#                    print ("%s:%s = %s"%(row, cell.column, cell_value))
                    pass                
                
                if cell_index < 0:
#                    print("cell_index column - %d:%s - %d:%s"%(bb_cell, bb_ord, cell_index, cell.column))
                    continue
                    
#                print ("state %s %s %s"%(state_0, state_1, state_2))

                if cell_index == 0:
                    if cell_value == "Paragraph":
                    
                        if state_0 != "None":
                            result += "            pass\n"
                            pass

                        state_0 = "Paragraph"
                        state_1 = "Init"
                        state_2 = "None"
                    elif cell_value == "Repeat":
                        if state_0 != "None":
                            result += "            pass\n"
                            pass

                        state_0 = "Repeat"
                        state_1 = "Init"
                        state_2 = "None"
                        pass
                    elif cell_value == "Until":
                        if state_0 != "None":                        
                            result += "            pass\n"
                            pass

                        state_0 = "Until"
                        state_1 = "Init"
                        state_2 = "None"
                        pass
                    elif cell_value == "#":
                        break
                        pass
                    elif len(cell_value) >= 1 and cell_value[0] == "#":
                        break
                        pass                        
                    elif cell_value == "":
                        pass
                    else:
                        Error ("Invalid Scenario '%s' first level command '%s' is not supported [%s]"%(xlsxFilename, cell_value, ["Paragraph", "Repeat", "Until"]))
                        return False
                        pass

                if cell_index == 1:
                    if state_1 == "Init":
                        if state_0 == "Paragraph":
                            if cell_value == "":
                                break

                            printValue += writeValue(cell_value, cell_index, 1)
                        elif state_0 == "Repeat":
                            if cell_value == "":
                                break

                            printValue += writeValue(cell_value, cell_index, 1)
                        elif state_0 == "Until":
                            if cell_value == "":
                                break

                            printValue += writeValue(cell_value, cell_index, 1)
                        pass
                    elif state_1 == "None":
                        if cell_value == "Preparation":
                            state_1 = "Preparation"
                            state_2 = "Macro"
                        elif cell_value == "Initial":
                            state_1 = "Initial"
                            state_2 = "Macro"
                        elif cell_value == "Parallel":
                            state_1 = "Parallel"
                            state_2 = "Init"
                        elif cell_value == "Race":
                            state_1 = "Race"
                            state_2 = "Init"
                        elif cell_value == "ShiftCollect":
                            state_1 = "ShiftCollect"
                            state_2 = "Init"
                        elif cell_value == "":
                            state_1 = "Skip"
                            break
                        else:
                            state_1 = "Active"
                            state_2 = "Macro"
                            pass
                    elif state_1 in ["Parallel", "Race", "ShiftCollect"]:
                        if cell_value.isdigit() is True:
                            state_2 = "Macro"

                            macro_id = int(cell_value)
                            pass
                        elif cell_value == "":
                            state_2 = "Skip"
                            pass
                        else:
                            if cell_value in ["Parallel", "Race", "ShiftCollect"]:
                                state_1 = cell_value
                                state_2 = "Init"
                            else:
                                state_1 = "Active"
                                state_2 = "Macro"
                                pass
                            pass
                    else:
                        pass

                    if state_1 == "Active":
                        if cell_value == "":
                            break

                        printValue += writeValue(cell_value, cell_index, 1)
                        pass
                    pass

                if cell_index > 1:
                    if state_1 == "Init":
                        if state_0 == "Paragraph":
                            if cell_value == "":
                                break

                            printValue += writeValue(cell_value, cell_index, 1)
                            pass                    
                    
                    if state_1 == "Active":
                        if state_2 == "Macro":
                            if cell_value == "":
                                break

                            printValue += writeValue(cell_value, cell_index, 1)
                            pass
                        pass
                    else:
                        if state_2 == "Macro":
                            if cell_value == "":
                                break

                            printValue += writeValue(cell_value, cell_index, 2)
                            pass
                        pass
                    pass

                if cell_index == 2:
                    if state_1 in ["Parallel", "Race", "ShiftCollect"]:
                        if state_2 == "Init":
                            try:
                                macro_count = int(cell_value)
                            except ValueError:
                                Error ("Invalid %s Count %s"%(state_1, cell_value))
                                return False
                                pass
                            pass
                        pass
                    pass
                pass

#            print ("state_0 %s %s %s %s"%(state_0, state_1, state_2, printValue))

            if state_0 == "Paragraph":
                if state_1 == "Init":
                    result += "        with self.addParagraph(%d, %s) as paragraph:\n"%(row_index, printValue)
                    state_1 = "None"
                    pass
                elif state_1 == "Preparation":
                    result += "            paragraph.addPreparation(%d, %s)\n"%(row_index, printValue)
                    state_1 = "None"
                    pass
                elif state_1 == "Initial":
                    result += "            paragraph.addInitial(%d, %s)\n"%(row_index, printValue)
                    state_1 = "None"
                    pass
                elif state_1 == "Active":
                    result += "            paragraph.addActive(%d, %s)\n"%(row_index, printValue)
                    state_1 = "None"
                    pass
                elif state_1 == "Skip":
                    state_1 = "None"
                    pass
                elif state_1 == "Parallel":
                    if state_2 == "Init":
                        result += "            with paragraph.addParallel(%d, %d) as parallel:\n"%(row_index, macro_count)
                        state_2 = "None"
                        pass
                    elif state_2 == "Macro":
                        result += "                parallel.addActive(%d, %d, %s)\n"%(row_index, macro_id - 1, printValue)
                        state_2 = "None"
                        pass
                    pass
                elif state_1 == "Race":
                    if state_2 == "Init":
                        result += "            with paragraph.addRace(%d, %d) as race:\n"%(row_index, macro_count)
                        state_2 = "None"
                        pass
                    elif state_2 == "Macro":
                        result += "                race.addActive(%d, %d, %s)\n"%(row_index, macro_id - 1, printValue)
                        state_2 = "None"
                        pass
                    pass
                elif state_1 == "ShiftCollect":
                    if state_2 == "Init":
                        result += "            with paragraph.addShiftCollect(%d, %d) as shiftcollect:\n"%(row_index, macro_count)
                        state_2 = "None"
                        pass
                    elif state_2 == "Macro":
                        result += "                shiftcollect.addActive(%d, %d, %s)\n"%(row_index, macro_id - 1, printValue)
                        state_2 = "None"
                        pass
                    pass
                pass
            elif state_0 == "Repeat":
                if state_1 == "Init":
                    result += "        with self.addRepeat(%d, %s) as repeat:\n"%(row_index, printValue)
                    state_1 = "None"
                    pass
                if state_1 == "Preparation":
                    result += "            repeat.addPreparation(%d, %s)\n"%(row_index, printValue)
                    state_1 = "None"
                    pass
                elif state_1 == "Initial":
                    result += "            repeat.addInitial(%d, %s)\n"%(row_index, printValue)
                    state_1 = "None"
                    pass
                elif state_1 == "Active":
                    result += "            repeat.addActive(%d, %s)\n"%(row_index, printValue)
                    state_1 = "None"
                    pass
                elif state_1 == "Skip":
                    state_1 = "None"
                    pass
                elif state_1 == "Parallel":
                    if state_2 == "Init":
                        result += "            with repeat.addParallel(%d, %d) as parallel:\n"%(row_index, macro_count)
                        state_2 = "None"
                        pass
                    elif state_2 == "Macro":
                        result += "                parallel.addActive(%d, %d, %s)\n"%(row_index, macro_id - 1, printValue)
                        state_2 = "None"
                        pass
                    pass
                elif state_1 == "Race":
                    if state_2 == "Init":
                        result += "            with repeat.addRace(%d, %d) as race:\n"%(row_index, macro_count)
                        state_2 = "None"
                        pass
                    elif state_2 == "Macro":
                        result += "                race.addActive(%d, %d, %s)\n"%(row_index, macro_id - 1, printValue)
                        state_2 = "None"
                        pass
                    pass
                elif state_1 == "ShiftCollect":
                    if state_2 == "Init":
                        result += "            with repeat.addShiftCollect(%d, %d) as shiftcollect:\n"%(row_index, macro_count)
                        state_2 = "None"
                        pass
                    elif state_2 == "Macro":
                        result += "                shiftcollect.addActive(%d, %d, %s)\n"%(row_index, macro_id - 1, printValue)
                        state_2 = "None"
                        pass
                    pass
                pass
                pass
            elif state_0 == "Until":
                result += "            repeat.addUntil(%d, %s)\n"%(row_index, printValue)
                result += "            pass\n"

                state_0 = "None"
                state_1 = "None"
                state_2 = "None"
                pass
            pass

        if state_0 != "None":
            result += "            pass\n"
            pass
        pass



    result += "        pass\n"
    result += "    pass\n"

    directory = os.path.dirname(generateFilename)
    if not os.path.exists(directory): 
        os.makedirs(directory)
        pass
    
    f = open(generateFilename, "w")
    f.write(result)
    f.close()
    
    return True
    pass

# ----------------------------------------- Text Encounters -----------------------------------------
def exportTextEncounterDataFromDocx(projectName, ProjectPath, ParamsPath, LocaleDir, ResourceDir, DataDir, DataFile, DataCreate):
    """
    Call example:
    exportData(projectName, ProjectPath, ParamsPath, "TextEncounter", "Resources", "Scripts/TextEncounter", "PathTextEncounter.xlsx", createScenario)
    """

    xlsxFileName = DataFile

    resoureces = ParamsPath + xlsxFileName

    if os.path.isfile(resoureces) is False:
        Error ("project %s file %s not exist in path '%s'"%(projectName, xlsxFileName, ParamsPath))
        return
        pass

    resourceList = getXlsxData(resoureces)

    for row in resourceList:
        if len(row) == 0:
            continue
            pass

        scriptPath = ProjectPath + ResourceDir + "/" + DataDir + "/"

        directory = os.path.dirname(scriptPath)
        if not os.path.exists(directory):
            os.makedirs(directory)
            pass

        if not os.path.exists(scriptPath + "__init__.py"):
            f = open(scriptPath + "__init__.py", "w")
            # f.write("# text encounters")
            f.write("#database")
            f.close()
            pass

        docx_path = row[0]

        register_filename = os.path.join(ParamsPath, "TextEncounters.xlsx")
        # register_filename = os.path.join(ParamsPath, "TextEncountersRegister.xlsx")

        docxFileName = "%s%s.docx"%(ParamsPath, docx_path)

        texts_filename = ProjectPath + "/" + LocaleDir + "/TextEncounterTexts.xml"

        DataCreate(docxFileName, register_filename, texts_filename, scriptPath)
        pass
    pass

def createTextEncounter(docxFileName, register_filename, texts_filename, scriptFilePath):
    print ("   !!! {} {}".format(docxFileName, scriptFilePath))
    scripts, texts = teparser.process(docxFileName)

    workbook = xlsxwriter.Workbook(register_filename)
    excel_sheet = workbook.add_worksheet()

    excel_sheet.write(0, 0, "ID")
    excel_sheet.write(0, 1, "Module")
    excel_sheet.write(0, 2, "TypeName")

    for r, (te_id, text) in enumerate(scripts, start=1):

        type_name = "TextEncounter{}".format(te_id)
        generateFilename = os.path.join(scriptFilePath, type_name + ".py")

        with open(generateFilename, "w") as f:
            f.write(text)

        excel_sheet.write(r, 0, te_id)
        excel_sheet.write(r, 1, "TextEncounter")
        excel_sheet.write(r, 2, type_name)
        pass

    with open(texts_filename, "w") as f:
        f.write(texts)

    workbook.close()
    pass
