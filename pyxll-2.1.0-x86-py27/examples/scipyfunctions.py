#python functions exposed numpy and scipy functions to Excel


from pyxll import xl_func
import scipy
from scipy import stats
import numpy


# @xl_func("numpy_array x: numpy_array")
# def scipy_linregress(x):
    # # return the linear regression of this data
	# slope, intercept, r_value, p_value, std_err = stats.linregress(x)
	# ar = numpy.array([[slope, intercept, r_value, p_value, std_err]])
	# return ar
	
@xl_func("numpy_array x, numpy_array y: numpy_array")
def scipy_linregress(x, y):
    # return the linear regression of this data
	slope, intercept, r_value, p_value, std_err = stats.linregress(x, y)
	ar = numpy.array([[slope, intercept, r_value, p_value, std_err]])
	return ar

@xl_func("numpy_array A, numpy_array b: numpy_array")
def scipy_lstsqs(A, b):
    # Compute least-squares solution to equation Ax = b.
	x, residues, rank, s = scipy.linalg.lstsq(A, b)
	ar = numpy.array(x)
	return ar	
