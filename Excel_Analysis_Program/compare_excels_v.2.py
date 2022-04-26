import pandas as pd
import numpy as np

# 866 -> 1,111 total layers
# 325 -> 323 plans
# 19 missing plans from oldSheet -> newSheet

# To use this program for different spreadsheets, replace the 2 paths below with the relevant .xlsx files. For csv files, replace '.read_excel()' with '.read_csv()'.
# Also replace the value of keyCol with the new DF's index column name.
oldSheet = pd.read_excel('C:\TBN\_CURRENT\Sorting_Project\_program\VETRO_Plan_vs_Layer_Analysis---OLD.xlsx' , sheet_name = 'LAYERS')
newSheet = pd.read_excel('C:\TBN\_CURRENT\Sorting_Project\_program\VETRO_Plan_vs_Layer_Analysis---NEW.xlsx' , sheet_name = 'LAYERS')
keyCol = 'Layer Names (VALUE)'

def newLayers(oldSheet , newSheet) :
    """Creates a DataFrame with the new Layers added in newSheet (compared to oldSheet) by performing a right-excluding DF merge.
        Args:
            oldSheet: the older excel/csv file
            newSheet: the newer excel/csv file
        Returns:
            newLayersDB: DataFrame containing the files that exist in newSheet, but not in oldSheet
    """
    
    newLayersDf = (oldSheet.merge(newSheet, on = keyCol, how = 'right', indicator = True)
        .query('_merge == "right_only"')
        .drop('_merge', 1))

    return newLayersDf
    #for later, make export new db to csv/excel in curr dir ('newLayersDf.xlsx)



def missingLayers(oldSheet , newSheet) :
    """Creates missingLayersDf DataFrame with the layers that exist in oldSheet, but do not exist in newSheet. (Removed/modified layers)
        Args:
            oldSheet: the older excel/csv file
            newSheet: the newer excel/csv file
        Returns:
            newLayersDf: DataFrame containing the files that exist in oldSheet, but not in newSheet
    """

    missingLayersDf = (oldSheet.merge(newSheet, on=keyCol, how='left', indicator=True)
        .query('_merge == "left_only"')
        .drop('_merge', 1))

    return missingLayersDf
    #for later, make export new db to csv/excel in curr dir ('missingLayers.xlsx)



def linkedLayers(newSheet) :
    """Creates layerPlansDf DataFrame with the layers that have (a) matched plan(s). Layers with no plans are added to the noPlanDf DataFrame.
        Args:
            newSheet: the newer excel/csv file
        Returns:
            layerPlansDf: DataFrame containing layers that have a correlating plan assigned
            noPlanDf: DataFrame containing layers that are marked "NO_PLAN" OR "MULTIPLE" under the column: NO_PLAN
    """
    layerPlansDf = newSheet[newSheet['Plan_1'] == 'NaN']
    noPlanDf = newSheet[newSheet['Plan_1'] != 'NaN']
    # for later, make export layerPlansDf and noPlanDf to csv/excel in curr dir (layerPlansDf.xlsx , noPlanDf.xlsx)
    print(layerPlansDf)


def main(oldSheet , newSheet) :
    """Main method for testing/initialization.
    """

    # test_1 = newLayers(oldSheet , newSheet)
    # print (test_1)

    test_2 = missingLayers(oldSheet , newSheet)
    print (test_2)

    # test_3 = linkedLayers(newSheet)
    # print (test_3)



if __name__ == "__main__":
    """Main initializer to optimize for running from CLI. Helps pandas find curr dir.
    """
    main(oldSheet , newSheet)