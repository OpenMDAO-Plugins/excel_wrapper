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
        ew = ExcelWrapper(excelFile)
        ew.execute()
        print '%f = 2.12345*%i' %(ew.y, ew.x)
        print '%s = ~%s' %(ew.bout, ew.b)
        print '%s = lower(%s)' %(ew.sout, ew.s) 
        del(ew)
        os._exit(1)
        
if __name__ == "__main__":
    unittest.main()
    
