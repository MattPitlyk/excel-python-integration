#python functions exposed numpy and scipy functions to Excel


from pyxll import xl_func
from scipy import stats
import numpy

# @xl_func("numpy_array x: float")
# def scipy_linregress(x):
    # # return the linear regression of this data
	# slope, intercept, r_value, p_value, std_err = stats.linregress(x)
	# return slope


# @xl_func("numpy_array x: numpy_array")
# def scipy_linregress(x):
    # # return the linear regression of this data
	# slope, intercept, r_value, p_value, std_err = stats.linregress(x)
	# ar = numpy.array([[slope, intercept, r_value, p_value, std_err]])
	# return ar
		
@xl_func("numpy_array x: float[]")
def scipy_linregress3(x):
    # return the linear regression of this data
	l = []
	l = list(stats.linregress(x))
	return l

	
@xl_func(": float")
def scipy_range():
    # return the linear regression of this data
	slope, intercept, r_value, p_value, std_err = stats.linregress(arange(10).reshape(2,5))
	return slope

	
@xl_func(": numpy_array")
def scipy_range2():
    # return the linear regression of this data
	
	return arange(10).reshape(2,5)