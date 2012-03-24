import unittest
import glob
import nose
import sys
import logging
import os

from openmdao.util.testutil import assert_raises, assert_rel_error
from excel_wrapper.excel_wrapper import ExcelWrapper

class ExcelWrapperTestCase(unittest.TestCase):

    def setUp(self):
        if os.name != 'nt':
            raise nose.SkipTest('Currently, excel_wrapper works only on Windows.')
        if os.name == 'posix':
            raise nose.SkipTest('Currently, excel_wrapper works only on Windows.')
        
    def tearDown(self):
        for pattern in ('*.log'):
            for name in glob.glob(pattern):
                os.remove(name)
        
    def test_ExcelWrapper(self):
        logging.debug('')
        logging.debug('test_ExcelWrapper')

        excelFile = r"excel_wrapper_test.xlsx"
        xmlFile = r"excel_wrapper_test.xml"
        ew = ExcelWrapper(excelFile, xmlFile)
        ew.execute()
        assert_rel_error(self, ew.y, 2.12345, 0.0001)
        self.assertEqual(not ew.b, ew.bout, msg='excel_wrapper could not handle a Bool variable')
        self.assertEqual(ew.s.lower(), ew.sout, msg='excel_wrapper couldn not handle a Str variable.')
        del(ew)
        os._exit(1)
        
if __name__ == "__main__":
    sys.argv.append('--cover-package=excel_wrapper.')
    sys.argv.append('--cover-erase')
    nose.runmodule()