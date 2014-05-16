"""
PyXLL Examples: Worksheet functions

The PyXLL Excel Addin is configured to load one or more
python modules when it's loaded. Functions are exposed
to Excel as worksheet functions by decorators declared in
the pyxll module.

Functions decorated with the xl_func decorator are exposed
to Excel as UDFs (User Defined Functions) and may be called
from cells in Excel.
"""

#
# 1) Basics - exposing functions to Excel
#

#
# xl_func is the main decorator and is used for exposing
# python functions to excel.
#
from pyxll import xl_func

#
# xl_func takes a string argument that is the signature of
# the function to be exposed to excel. This example takes
# three integers and returns an integer.
#

@xl_func("int x, int y, int z: int")
def basic_pyxll_function_1(x, y, z):
    """returns (x * y) ** z """
    return (x * y) ** z * 2

#
# there are a number of basic types that can be used in
# the function signature. These include:
#   int, float, bool and string
# There are more types that we'll come to later.
#

@xl_func("int x, float y, bool z: float")
def basic_pyxll_function_2(x, y, z):
    """if z return x, else return y"""
    if z:
        # we're returning an integer, but the signature
        # says we're returning a float.
        # PyXLL will convert the integer to a float for us.
        return x
    return y

#
# you can change the category the function appears under in
# Excel by using the optional argument 'category'.
#

@xl_func("int x: int", category="My new PyXLL Category")
def basic_pyxll_function_3(x):
    """docstrings appear as help text in Excel"""
    return x

#
# 2) The var type
#

#
# Another type is the var type. This can represent any
# of the basic types, depending on what type is passed to the
# function, or what type is returned.
#

@xl_func("var x: string")
def var_pyxll_function_1(x):
    """takes an float, bool, string, None or array"""
    # we'll return the type of the object passed to us, pyxll
    # will then convert that to a string when it's returned to
    # excel.
    return type(x)

#
# If var is the return type. PyXll will convert it to the
# most suitable basic type. If it's not a basic type and
# no suitable conversion can be found, it will be converted
# to a string and the string will be returned.
#

@xl_func("bool x: var")
def var_pyxll_function_2(x):
    """if x return string, else a number"""
    if x:
        return "var can be used to return different types"
    return 123.456

#
# 3) Arrays
#

#
# Arrays in PyXll are 2d arrays that correspond to the grid in
# Excel. In python, they are represented as lists of lists.
# Arrays of any type can be used, and the var type may be
# an array of vars.
#
# Arrays of floats are more efficient to marshall between
# python and Excel than other array types so should be used
# when possible instead of var.
#
# NumPy arrays are also supported. For those, see the
# next section.
#

@xl_func("float[] x: float")
def array_pyxll_function_1(x):
    """returns the sum of a range of floats"""
    total = 0.0
    # x is a list of lists - iterate through the rows:
    for row in x:
        # each row is a list of floats
        for element in row:
            total += element
    return total

#
# Functions can also return 2d arrays as lists of lists
# in python. These can be used as array formulas in excel
# to return a grid of data.
#

@xl_func("string[] array, string sep: string[]")
def array_pyxll_function_2(x, sep):
    """joins each row by 'sep' and returns a column of strings"""
    # result is a list of lists of strings
    result = []
    for row in x:
        s = sep.join(row)
        # the result is just one column wide
        result_row = [s]
        result.append(result_row)
    return result
    
#
# the var type may also be used to pass and return arrays, but
# the python function should do any necessary type checking.
#

@xl_func("var x: string[]")
def array_pyxll_function_3(x):
    """returns the types of the elements as strings"""
    # x may not be an array
    if not isinstance(x, list):
        return [[type(x)]]

    # x is a 2d array - list of lists.
    return [[type(e) for e in row] for row in x]

#
# var arrays may also be used
#

@xl_func("var[] x: string[]")
def array_pyxll_function_4(x):
    """returns the types of the elements as strings"""
    # x will always be a 2d array - list of lists.
    return [[type(e) for e in row] for row in x]

#
# xlfCaller can be used to get information about the
# calling cell or range
#
from pyxll import xlfCaller

@xl_func("var[] x: var[]")
def array_pyxll_function_5(x):
    """
    return the input array with row and col numbers.
    
    This example shows how to use xlfCaller to get the range
    of the cells the array function is being called by.
    """
    # get the size of the rect the array function was called over
    # i.e. the size of the array to be returned
    caller = xlfCaller()
    width = caller.rect.last_col - caller.rect.first_col + 1
    height = caller.rect.last_row - caller.rect.first_row + 1
    
    # check the input array is the same size
    assert len(x) == height
    assert len(x[0]) == width
    
    # construct the return value as a list of lists with the
    # same dimensions as the calling cells.
    result = []
    for i in range(height):
        row = []
        for j in range(width):
            row.append("%s (col=%d, row=%d)" % (x[i][j], j, i))
        result.append(row)

    return result

#
# 4) NumPy arrays
#
# the numpy_array type corresponds to the numpy.ndarray
# type.
#
# You must have numpy installed to be able to use the
# numpy_array type.
#

@xl_func("numpy_array x: numpy_array")
def numpy_array_function_1(x):
    # return the transpose of the array
    return x.transpose()

@xl_func("numpy_array<float_nan> x: numpy_array<float_nan>")
def numpy_array_function_2(x):
    # simply return the  array to demonstrate how errors from
    # excel may be passed to python as NaN
    return x

#
# As well as 2d arrays, 1d rows and columns may also be used
# as argument and return types.
#

@xl_func("numpy_row x: string")
def numpy_row_function_1(x):
    return str(x)

@xl_func("numpy_row x: numpy_column")
def numpy_row_function_2(x):
    return x.transpose()

@xl_func("numpy_column x: string")
def numpy_col_function_1(x):
    return str(x)

@xl_func("numpy_column x: numpy_row")
def numpy_col_function_2(x):
    return x.transpose()

#
# 5) Date and time types
#

#
# There are three date and time types: date, time, datetime
#
# Excel represents dates and times as floating point numbers.
# The pyxll datetime types convert the excel number to a
# python datetime.date, datetime.time and datetime.datetime
# object depending on what type you specify in the signature.
#
# dates and times may be returned using their type as the return
# type in the signature, or as the var type.
#

import datetime

@xl_func("date x: string")
def datetime_pyxll_function_1(x):
    """returns a string description of the date"""
    return "type=%s, date=%s" % (type(x), x)

@xl_func("time x: string")
def datetime_pyxll_function_2(x):
    """returns a string description of the time"""
    return "type=%s, time=%s" % (type(x), x)

@xl_func("datetime x: string")
def datetime_pyxll_function_3(x):
    """returns a string description of the datetime"""
    return "type=%s, datetime=%s" % (type(x), x)
    
@xl_func("datetime[] x: datetime")
def datetime_pyxll_function_4(x):
    """returns the max datetime"""
    m = datetime.datetime(1900, 1, 1)
    for row in x:
        m = max(m, max(row))
    return m

#
# 6) xl_cell
#
# The xl_cell type can be used to receive a cell
# object rather than a plain value. The cell object
# has the value, address, formula and note of the
# reference cell passed to the function.
#
# The function must be a macro sheet equivalent function
# in order to access the value, address, formula and note
# properties of the cell.
#

@xl_func("xl_cell cell : string", macro=True)
def xl_cell_example(cell):
    """a cell has a value, address, formula and note"""
    return "[value=%s, address=%s, formula=%s, note=%s]" % (cell.value, cell.address, cell.formula, cell.note)


@xl_func("int x: int")
def pyxll_test_1(x):
    """returns x * 2"""
    return x * 2
	
	
	
	