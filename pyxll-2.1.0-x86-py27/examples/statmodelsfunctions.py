
from pyxll import xl_func
import numpy as np
import statsmodels.stats.outliers_influence


@xl_func("numpy_array dm, int index: float")
def statsmodels_vif(dm,index):
	# expose the statmodels module function VIF
	return statsmodels.stats.outliers_influence.variance_inflation_factor(dm, index)