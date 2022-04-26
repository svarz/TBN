import pandas as pd

# 866 -> 1,111 total layers
# 325 -> 323 plans
# 19 missing plans from oldSheet -> newSheet

# C:\TBN\_CURRENT\Sorting_Project\_program\VETRO_Plan_vs_Layer_Analysis---OLD.xlsx , sheet_name = LAYERS
# keyCol = Plan_1

def remNullRows(excelDf) :
	"""Creates layerPlansDf DataFrame with the layers that have (a) matched plan(s). Layers with no plans are added to the noPlanDf DataFrame.
		Args:
			excelDf: the newer excel file.
		Returns:
			layerPlansDf: DataFrame containing layers that have a correlating plan assigned.
			noPlanDf: DataFrame containing layers that are marked "NO_PLAN" OR "MULTIPLE" under the column: NO_PLAN.
	"""

	excelDf.dropna(axis = 0 , subset = keyCol , inplace = True)
	return excelDf


def main(excelDf) :
	"""Main method for testing/initialization.
	"""

	test_3 = remNullRows(excelDf)
	test_3.to_excel('remNullRowsDf.xlsx')

if __name__ == "__main__":
	"""Main initializer to optimize for running from CLI. Helps pandas find curr dir.
	"""

	# gathers the file path name, sheet name, and column names through user input.
	print( "Input the full file path to the excel file within your machine below (Copy/paste recommended.):" )
	excelDfPath = input()
	print( "\nInput the sheet name of the excel file:" )
	sheetName = input()
	print( "\nInput the column name that you would like to parse for non-null values (Must be the exact column name. Copy/paste recommended.):" )
	keyCol = input()

	excelDf = pd.read_excel(excelDfPath, sheet_name = sheetName)

	main(excelDf)