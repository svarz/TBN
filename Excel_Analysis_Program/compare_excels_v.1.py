import pandas as pd
import numpy as np

oldSheet = pd.read_excel('VETRO_Plan_vs_Layer_Analysis---OLD.xlsx' , sheet_name = 'LAYERS')
newSheet = pd.read_excel('VETRO_Plan_vs_Layer_Analysis---NEW.xlsx' , sheet_name = 'LAYERS')

joined_df = pd.merge(oldSheet , newSheet , on = 'Layer Names (VALUE)' , how = 'right' , indicator = True)
joined_df['NO_PLAN'] = np.where(joined_df['_merge'] == 'both', joined_df['NO_PLAN_x'], joined_df['NO_PLAN_y'])
joined_df['Plan_1'] = np.where(joined_df['_merge'] == 'both', joined_df['Plan_1_x'], joined_df['Plan_1_y'])
joined_df['Plan_2'] = np.where(joined_df['_merge'] == 'both', joined_df['Plan_2_x'], joined_df['Plan_2_y'])
joined_df['Plan_3'] = np.where(joined_df['_merge'] == 'both', joined_df['Plan_3_x'], joined_df['Plan_3_y'])
joined_df['Plan_4'] = np.where(joined_df['_merge'] == 'both', joined_df['Plan_4_x'], joined_df['Plan_4_y'])
joined_df['Plan_5'] = np.where(joined_df['_merge'] == 'both', joined_df['Plan_5_x'], joined_df['Plan_5_y'])
joined_df['Plan_6'] = np.where(joined_df['_merge'] == 'both', joined_df['Plan_6_x'], joined_df['Plan_6_y'])

joined_df = joined_df.drop(columns=['NO_PLAN_x', 'NO_PLAN_y', 'Plan_1_x', 'Plan_1_y' , 'Plan_2_x' , 'Plan_2_y' ,
'Plan_3_x' , 'Plan_3_y' , 'Plan_4_x' , 'Plan_4_y' , 'Plan_5_x' , 'Plan_5_y' , 'Plan_6_x' , 'Plan_6_y'] , axis = 1)

joined_df.to_excel(r'C:\TBN\_CURRENT\Sorting_Project\_program\VETRO_Plan_vs_Layer_Analysis---JOINED.xlsx' , index = False)