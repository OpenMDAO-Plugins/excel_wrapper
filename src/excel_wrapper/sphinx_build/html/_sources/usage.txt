
===========
Usage Guide
===========

Using excel_wrapper
========================================

Let's say we want to implement the following equations in Excel and wrap into OpenMDAO.

.. testcode:: equations

    y = 2.12345x
    bout = ~b
    sout = lower(s)

where x is an integer, y is a float, b and bout are booleans, and s and sout are strings. Create a new Excel file and
define the following names: x, y, b, bout, s, and sout.

Specify the attributes of the variables in an xml file as follows:

.. testcode:: xml

    <?xml version = "1.0"?>
    <Variables>
        <Variable name="x" type="Int" iotype="in" value="1" desc="integer x" units="kg"/>
        <Variable name="s" type="Str" iotype="in" value="Hello World!"/>
        <Variable name="b" type="Bool" iotype="in" value="False" desc="boolean b"/>
        <Variable name="y" type="Float" iotype="out" desc="float y" units="N"/>
        <Variable name="bout" type="Bool" iotype="out" desc="boolean bout"/>
        <Variable name="sout" type="Str" iotype="out" desc="string sout"/>
    </Variables>

Please note that the inputs should be followed by the outputs and that units and desc are optional.

Here is a python code that calls the wrapper and prints out the result.

.. testcode:: excel_wrapper_parts

    from excel_wrapper.excel_wrapper import ExcelWrapper
    import os

    excelFile = r"excel_wrapper_test.xlsx"
    xmlFile = r"excel_wrapper_test.xml"
    ew = ExcelWrapper(excelFile, xmlFile)
    ew.execute()
    print '2.12345*%i = %f' %(ew.x, ew.y)
    print '~%s = %s' %(ew.b, ew.bout)
    print 'lower(%s) = %s' %(ew.s, ew.sout) 
    del(ew)
    os._exit(1)
    
Place the excel file and the xml in the same directory as the above python code or specify the full parhs of the files.