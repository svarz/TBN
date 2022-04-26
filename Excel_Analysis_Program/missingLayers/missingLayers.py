import pandas as pd

# 866 -> 1,111 total layers
# 325 -> 323 plans
# 19 missing plans from oldSheet -> newSheet

# C:\TBN\_CURRENT\Sorting_Project\_program\VETRO_Plan_vs_Layer_Analysis---OLD.xlsx , sheet_name = 'LAYERS')
# C:\TBN\_CURRENT\Sorting_Project\_program\VETRO_Plan_vs_Layer_Analysis---NEW.xlsx , sheet_name = 'LAYERS')
# keyCol = Layer Names (VALUE)

def missingLayers(oldSheet , newSheet) :
	"""Creates missingLayersDf DataFrame with the layers that exist in oldSheet, but do not exist in newSheet. (Removed/modified layers).
		Args:
			oldSheet: the old excel file.
			newSheet: the newe excel file.
		Returns:
			newLayersDf: DataFrame containing the files that exist in oldSheet, but not in newSheet.
	"""

	missingLayersDf = ( oldSheet.merge( newSheet, on = keyCol, how = 'left', indicator = True )
		.query( '_merge == "left_only"' )
		.drop( '_merge', 1 ) )

	return missingLayersDf


def main(oldSheet , newSheet) :
	"""Main method for testing/initialization.
	"""
	
	test_2 = missingLayers(oldSheet , newSheet)
	test_2.to_excel('missingLayersDf.xlsx')

if __name__ == "__main__":
	"""Main initializer to optimize for running from CLI. Helps pandas find curr dir.
	"""

	print( "Input the full file path to the original excel within your machine below (Copy/paste recommended.):" )
	oldSheetPath = input()
	print( "\nInput the full file path to the new excel within your machine below(Copy/paste recommended.):" )
	newSheetPath = input()
	print( "\nInput the sheet name of the excel files that you would like to compare(The sheet name should be the same for both files):" )
	sheetName = input()
	print( "\nInput the column name that you would like to compare the excel files on (Must be the exact column name. Copy/paste recommended.):" )
	keyCol = input()

	oldSheet = pd.read_excel(oldSheetPath, sheet_name = sheetName)
	newSheet = pd.read_excel(newSheetPath, sheet_name = sheetName)

	main(oldSheet , newSheet)