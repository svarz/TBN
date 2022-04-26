import pandas as pd

# 866 -> 1,111 total layers
# 325 -> 323 plans
# 19 missing plans from oldSheet -> newSheet

# C:\TBN\_CURRENT\Sorting_Project\_program\VETRO_Plan_vs_Layer_Analysis---OLD.xlsx , sheet_name = LAYERS )
# C:\TBN\_CURRENT\Sorting_Project\_program\VETRO_Plan_vs_Layer_Analysis---NEW.xlsx , sheet_name = LAYERS )
# Layer Names (VALUE)

def newLayers( oldSheet , newSheet ) :
	"""Creates a DataFrame with the new Layers added in newSheet (compared to oldSheet) by performing a right-excluding DF merge.
    	Args:
        	oldSheet: the older excel file.
        	newSheet: the newer excel file.
		Returns:
			newLayersDB: DataFrame containing the files that exist in newSheet, but not in oldSheet.
	"""

	# left-join eliminating duplicates in newSheet, so that each row of oldSheet joins with exactly 1 row of newSheet. Parameter indicator
	# used to return an extra column indicating which table the row.

	newLayersDf = ( newSheet.merge( oldSheet, on = keyCol, how = 'left', indicator = True )
		.query( '_merge == "left_only"' )
		.drop( '_merge', 1 ) )

	return newLayersDf

def main( oldSheet , newSheet ) :
	"""Main method for testing/initialization.
	"""

	test_1 = newLayers(oldSheet, newSheet)

	# exports df to excel file
	test_1.to_excel('newLayersDf.xlsx')

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