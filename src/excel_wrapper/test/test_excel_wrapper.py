import unittest

from excel_wrapper.excel_wrapper import ExcelWrapper
import os

class ExcelWrapperTestCase(unittest.TestCase):

    def setUp(self):
        pass
        
    def tearDown(self):
        pass
        
    def test_ExcelWrapper(self):
        excelFile = r"excel_wrapper_test.xlsx"
        xmlFile = r"excel_wrapper_test.xml"
        ew = ExcelWrapper(excelFile, xmlFile)
        ew.execute()
        print '2.12345*%i = %f' %(ew.x, ew.y)
        print '~%s = %s' %(ew.b, ew.bout)
        print 'lower(%s) = %s' %(ew.s, ew.sout) 
        del(ew)
        os._exit(1)
        
if __name__ == "__main__":
    unittest.main()
    
