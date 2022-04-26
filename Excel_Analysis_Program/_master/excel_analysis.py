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

def missingLayers(oldSheet , newSheet) :
	"""Creates missingLayersDf DataFrame with the layers that exist in oldSheet, but do not exist in newSheet. (Removed/modified layers).
		Args:
			oldSheet: the old excel file.
			newSheet: the newer excel file.
		Returns:
			newLayersDf: DataFrame containing the files that exist in oldSheet, but not in newSheet.
	"""

	missingLayersDf = ( oldSheet.merge( newSheet, on = keyCol, how = 'left', indicator = True )
		.query( '_merge == "left_only"' )
		.drop( '_merge', 1 ) )

	return missingLayersDf

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

def main( oldSheet , newSheet , desiredFunction) :
	"""Main method for testing/initialization.
	"""

	output = pd.DataFrame()

	if( desiredFunction == "newLayers" ) :
		output = newLayers( oldSheet, newSheet )

	elif( desiredFunction == "missingLayers" ) :
		output = missingLayers( oldSheet, newSheet )
	
	elif( desiredFunction == "remNullRows" ) :
		output = remNullRows( oldSheet )

	output.to_excel( str( desiredFunction ) + "Df.xlsx" )
	quit()


if __name__ == "__main__":
	"""Main initializer to optimize for running from CLI. Helps pandas find curr dir.
	"""

	print("Which program would you like to use?\nOptions:")
	print("newLayers: Gets the layers that are in the new excel, but not in the old.\nmissingLayers: Gets the layers that are in the old excel, but not in the new.\nremNullRows: Gets an excel that excludes the rows with null values in the selected column.")
	selectedProgram = input()

	# These statements check which program the user wants to run.
	if (selectedProgram == "newLayers") :
			desiredFunction = "newLayers"
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

			main(oldSheet, newSheet, desiredFunction)

	elif (selectedProgram == "missingLayers") :
		desiredFunction = "missingLayers"
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

		main(oldSheet, newSheet, desiredFunction)

	elif (selectedProgram == "remNullRows") :
		desiredFunction = "remNullRows"
		print( "Input the full file path to the excel file within your machine below (Copy/paste recommended.):" )
		excelDfPath = input()
		print( "\nInput the sheet name of the excel file:" )
		sheetName = input()
		print( "\nInput the column name that you would like to parse for non-null values (Must be the exact column name. Copy/paste recommended.):" )
		keyCol = input()

		excelDf = pd.read_excel(excelDfPath, sheet_name = sheetName)

		emptyDf = pd.DataFrame()

		main(excelDf, emptyDf , desiredFunction)

	else :
		print ("Invalid program selected. Please double check spelling and restart the program.")
		quit()


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