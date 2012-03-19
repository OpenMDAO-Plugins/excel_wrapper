
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
define the variable names.

Specify the attributes of the variables in an xml file as follows:

.. testcode:: xml

    <?xml version = "1.0"?>
    <Variables>
        <Variable name="x" type="Int" iotype="in" value="1" desc="integer x"/>
        <Variable name="s" type="Str" iotype="in" value="Hello World!"/>
        <Variable name="b" type="Bool" iotype="in" value="False" desc="boolean b"/>
        <Variable name="y" type="Float" iotype="out" desc="float y"/>
        <Variable name="bout" type="Bool" iotype="out" desc="boolean bout"/>
        <Variable name="sout" type="Str" iotype="out" desc="string sout"/>
    </Variables>

Please note that the inputs should be followed by the outputs and that the xml file should be saved as excel_wrapper.xml.

Here is a code that executes the wrapper and prints out the result.

.. testcode:: excel_wrapper_parts

    from excel_wrapper.excel_wrapper import ExcelWrapper
    import os

    excelFile = r"excel_wrapper_test.xlsx"
    ew = ExcelWrapper(excelFile)
    ew.execute()
    print '%f = 2.12345*%i' %(ew.y, ew.x)
    print '%s = ~%s' %(ew.bout, ew.b)
    print '%s = lower(%s)' %(ew.sout, ew.s) 
    del(ew)
    os._exit(1)
    
Place excel_wrapper.xml in the same directory as the above python code.