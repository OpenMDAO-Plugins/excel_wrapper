from openmdao.main.api import Component
from openmdao.lib.datatypes.api import Float, Int, Bool, Str
import os
import win32com.client
from xml.etree import ElementTree as ET

class ExcelWrapper(Component):
    """ An Excel Wrapper """

    def __init__(self, excelFile, xmlFile):
        super(ExcelWrapper, self).__init__()

        self.xmlFile = xmlFile
        try:
            tree = ET.parse(self.xmlFile)
        except:
            if not os.path.exists(self.xmlFile):
                print 'Cannot find the xml file at ' + self.xmlFile
        
        self.variables = tree.findall("Variable")
        for v in self.variables:
            name = v.attrib['name']
            kwargs = dict([(key, v.attrib[key]) for key in ('iotype', 'desc', 'units') if key in v.attrib])
            if v.attrib['iotype'] == 'in':
                if v.attrib['type'] == 'Float':
                    self.add(v.attrib['name'], Float(float(v.attrib['value']), **kwargs))
                elif v.attrib['type'] == 'Int':
                    self.add(v.attrib['name'], Int(int(v.attrib['value']), **kwargs))
                elif v.attrib['type'] == 'Bool':
                    self.add(v.attrib['name'], Bool(v.attrib['value'], **kwargs))
                elif v.attrib['type'] == 'Str':
                    self.add(v.attrib['name'], Str(v.attrib['value'], **kwargs))
            else:
                if v.attrib['type'] == 'Float':
                    self.add(v.attrib['name'], Float(**kwargs))
                elif v.attrib['type'] == 'Int':
                    self.add(v.attrib['name'], Int(**kwargs))
                elif v.attrib['type'] == 'Bool':
                    self.add(v.attrib['name'], Bool(**kwargs))
                elif v.attrib['type'] == 'Str':
                    self.add(v.attrib['name'], Str(**kwargs))
            # if v.attrib['iotype'] == 'in':
                # if v.attrib['type'] == 'Float':
                    # vars()[name] = Float(float(v.attrib['value']), **kwargs)
                # elif v.attrib['type'] == 'Int':
                    # vars()[name] = Int(int(v.attrib['value']), **kwargs)
                # elif v.attrib['type'] == 'Bool':
                    # vars()[name] = Bool(v.attrib['value'], **kwargs)
                # elif v.attrib['type'] == 'Str':
                    # vars()[name] = Str(v.attrib['value'], **kwargs)
            # else:
                # if v.attrib['type'] == 'Float':
                    # vars()[name] = Float(**kwargs)
                # elif v.attrib['type'] == 'Int':
                    # vars()[name] = Int(**kwargs)
                # elif v.attrib['type'] == 'Bool':
                    # vars()[name] = Bool(**kwargs)
                # elif v.attrib['type'] == 'Str':
                    # vars()[name] = Str(**kwargs)
        
        self.excelFile = excelFile
        self.xlInstance = None
        self.workbook = None
        self.ExcelConnectionIsValid = True
        if not os.path.exists(self.excelFile):
            print "Invalid file given"
            self.ExcelConnectionIsValid = False
        
        else:
            self.excelFile = os.path.abspath(self.excelFile)
            xl = self.openExcel()
            if xl is None:
                print "Connection to Excel failed."
                self.ExcelConnectionIsValid = False
            
            else:
                self.xlInstance = xl
                self.workbook = xl.Workbooks.Open(self.excelFile)
    # End __init__

    def __del__(self):
        if self.workbook is not None:
            self.workbook.Close(SaveChanges=False)
        
        if self.xlInstance is not None:
            del(self.xlInstance)
            self.xlInstance = None        
    # End __del__

    def openExcel(self):
        try:
            xl = win32com.client.Dispatch("Excel.Application")
        
        except:
            return None
        
        return xl
    # End openExcel

    def execute(self):
        if not self.ExcelConnectionIsValid or \
            self.xlInstance is None or \
            self.workbook is None:
            print "Aborted Execution of Bad ExcelWrapper Component Instance"
            return
        
        wb = self.workbook
        namelist = [x.name for x in wb.Names]

        for v in self.variables:
            name = v.attrib['name']

            if v.attrib['iotype'] == 'in':
                    self.xlInstance.Range(wb.Names(name).RefersToLocal).Value = v.attrib['value']        
            else:
                try:
                    excel_value = self.xlInstance.Range(wb.Names(name).RefersToLocal).Value
                except:
                    print 'Cannot retrieve values from the Excel file'
                    if name not in namelist:
                        print 'Error: ' + name + ' is not defined in ' + self.excelFile

                if v.attrib['type'] == 'Float':
                    vars(self)[name] = float(excel_value)
                elif v.attrib['type'] == 'Int':
                    vars(self)[name] = int(excel_value)
                elif v.attrib['type'] == 'Bool':
                    vars(self)[name] = excel_value
                elif v.attrib['type'] == 'Str':
                    vars(self)[name] = excel_value
    # End execute
# End ExcelWrapper class

if __name__ == '__main__':
    excelFile = r"C:\META Software\tool\excel_wrapper_test.xlsx"
    xmlFile = 'excel_wrapper.xml'
    ew = ExcelWrapper(excelFile, xmlFile)
    ew.execute()
    print ew.x, ew.y, ew.b, ew.bout, ew.s, ew.sout
    del(ew)
    os._exit(1)
# End excel_wrapper.py