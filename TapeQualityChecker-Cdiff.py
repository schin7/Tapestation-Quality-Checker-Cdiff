import pandas as pd
import glob
import os
import os.path
import numpy as np
import xlsxwriter

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

PeakList = []

#Build up list of _compactPeakTable files:
for files in glob.glob("*compactPeakTable.csv"):
    PeakList.append(files) #filename with extension    


#Create dataframe for an individual peakTable in list
for peakTable in PeakList:
    #If QC_analyzed exists, skip
    if os.path.isfile(peakTable.replace(".csv","") + "QCAnalyzed.xlsx") == True:
        print peakTable + "peakTable file previously analyzed"
        pass
    else:
        #Open ANSI encoded file
        df = pd.read_csv(peakTable, encoding = 'cp1251')
        #Drop unneeded columns and rename unicode column to ul
        #Not using index for some cols as data exported is not consistent and columns do not always match
        df = df.drop('Peak', axis=1, errors='ignore')
        df = df.drop(df.columns[[0,5]], axis=1)
        
        df = df.drop('Peak Comment', axis=1, errors='ignore')
        df = df.drop('Peak Molarity [nmol/l]', axis=1, errors='ignore')
        
        df = df.drop('From [bp]', axis=1, errors='ignore')
        df = df.drop('To [bp]', axis=1, errors='ignore')
        df = df.drop('% Integrated Area', axis=1, errors='ignore')
        
        df = df.drop('Area', axis=1, errors='ignore')
        df = df.drop('% of Total', axis=1, errors='ignore')
        df = df.drop('Run Distance [%]', axis=1, errors='ignore')
        df = df.drop('From [%]', axis=1, errors='ignore')
        df = df.drop('To [%]', axis=1, errors='ignore')
        
        df.rename(columns = {list(df)[3]: 'Calibrated Conc. [ng/ul]'},inplace=True)
       
        df = df[~df.Observations.str.contains("Peak outside of Sizing Range", na=False)]
        #print(df)        
        #group samples by name and select first value then find string that is upper in observation and divide 
        #s = df.loc[df['Observations'].str.contains('Label'), "Height"]
        #print (df.loc[df["Height"].isin(s)])        
        
        
        lower = df.groupby('Sample Description')['Height'].transform('first')
        #print(lower)
        df['Ratio (Upper/Lower)'] = df.loc[df['Observations'].str.contains('Upper Marker', na=False)]['Height'].div(lower)  
        df['Ratio Check'] = np.where((df['Ratio (Upper/Lower)'] >= 1.95) &
                                     (df['Ratio (Upper/Lower)'] <= 3.1), 'Pass', 'Require Manual')
        df['Ratio Check'].mask(df['Ratio (Upper/Lower)'].isna(),np.nan,inplace=True)
        
        #print (df)
        
        #create separate df for upper values that correspond to QC check
        upperdf = df[df['Observations'].str.contains('Upper', na=False)]        
        upperdf = upperdf.drop(upperdf.columns[[2,3,4]], axis=1)
        #print(upperdf)
        
        #create additional copies of size and calibration for specifica tcda and tcdb
        df['Size [bp] tcdA']= df['Size [bp]']
        df['Calibrated Conc. [ng/ul] tcdA'] = df['Calibrated Conc. [ng/ul]']
        df['Size [bp] tcdB'] = df['Size [bp]']
        df['Calibrated Conc. [ng/ul] tcdB'] = df['Calibrated Conc. [ng/ul]']
              
        df['Tapestation Call tcdA'] = np.where(((df['Sample Description'].str.contains('tcdA')) &
                                                (df['Size [bp] tcdB'] >= 597.15) &
                                                (df['Size [bp] tcdB'] <= 729.85) &
                                                (df['Calibrated Conc. [ng/ul] tcdB'] > 1)), 'POS', 'NEG') 
        
        df['Tapestation Call tcdB'] = np.where(((df['Sample Description'].str.contains('tcdB')) &
                                                (df['Size [bp] tcdA'] >= 441.45) &
                                                (df['Size [bp] tcdA'] <= 539.55) &
                                                (df['Calibrated Conc. [ng/ul] tcdA'] > 1)), 'POS', 'NEG')        

        #create composite of composite dataframes
        compositedf1 = df[df['Tapestation Call tcdA'].str.contains('POS', na=False)]
        
        compositedf2 = df[df['Tapestation Call tcdB'].str.contains('POS', na=False)]
        
        #drop unneeded/conflicting columns        
        compositedf1 = compositedf1.drop(compositedf1.columns[[4,5,6,7,10,11,13]], axis=1)
        compositedf2 = compositedf2.drop(compositedf2.columns[[4,5,6,7,8,9,12]], axis=1) 
        
        #merge left twice on sample description
        df2 = upperdf.merge(compositedf1, on=['Sample Description'], how ='left')
        #print(df2)
        df3 = df2.merge(compositedf2, on=['Sample Description'], how ='left')
        
        df3 = df3.drop(df3.columns[[2,5,6,7,11,12,13]], axis=1)
        #print(df3)
        
        df3['Composite Tapestation Call'] = np.where(((df3['Tapestation Call tcdA'] == 'POS') | (df3['Tapestation Call tcdB'] == 'POS')), 'POS', 'NEG') 
                
        #fill NaN values with NA or NEG
        df3['Size [bp] tcdA'] = df3['Size [bp] tcdA'].fillna('NA')
        df3['Size [bp] tcdB'] = df3['Size [bp] tcdB'].fillna('NA')
        df3['Calibrated Conc. [ng/ul] tcdA'] = df3['Calibrated Conc. [ng/ul] tcdA'].fillna('NA')
        df3['Calibrated Conc. [ng/ul] tcdB'] = df3['Calibrated Conc. [ng/ul] tcdB'].fillna('NA')
        df3['Tapestation Call tcdA'] = df3['Tapestation Call tcdA'].fillna('NEG')
        df3['Tapestation Call tcdB'] = df3['Tapestation Call tcdB'].fillna('NEG')
        #print(df3)
        
        df3['Composite Tapestation Call'] = df3['Composite Tapestation Call'].fillna('NEG')
        df3['Well']=df3['Well_x']
        
        df3 = df3[['Well','Sample Description', 'Ratio (Upper/Lower)', 'Ratio Check', 'Size [bp] tcdA', 'Calibrated Conc. [ng/ul] tcdA', 'Size [bp] tcdB', 'Calibrated Conc. [ng/ul] tcdB', 'Tapestation Call tcdA', 'Tapestation Call tcdB', 'Composite Tapestation Call']]
        
    


        #Dataframe to excel
        path = os.getcwd()
        
        path = path +"/" + peakTable.replace("compactPeakTable.csv","") + "QCAnalyzed.xlsx"
            
        writer = pd.ExcelWriter(path, engine ='xlsxwriter')
        df3.to_excel(writer, sheet_name ='QC Analyzed', index = False)
        workbook = writer.book
        worksheet = writer.sheets['QC Analyzed']
        
        # Light red fill with dark red text.
        format1 = workbook.add_format({'bg_color':   '#FFC7CE',
                                       'font_color': '#9C0006'})
        
        # Green fill with dark green text.
        format2 = workbook.add_format({'bg_color':   '#C6EFCE',
                                       'font_color': '#006100'})
        # Grey fill with dark green text.
        format3 = workbook.add_format({'bg_color':   '#D3D3D3',
                                       'font_color': 'black'})        
        
        # Border formatting
        border_fmt = workbook.add_format({'border':1,
                                          'align':'left'})
        worksheet.conditional_format('A2:K1048576',{'type' : 'no_blanks', 'format' : border_fmt})
        
        for column in df3:
            column_width = max(df3[column].astype(str).map(len).max(), len(column))
            col_idx = df3.columns.get_loc(column)
            worksheet.set_column(col_idx, col_idx, column_width)
       
        #Conditional formatting for worksheet
        worksheet.conditional_format('D2:D1048576', {'type': 'cell',
                                            'criteria': 'equal to',
                                            'value': '"Pass"',
                                            'format': format2})   
        
        worksheet.conditional_format('D2:D1048576', {'type': 'cell',
                                            'criteria': 'equal to',
                                            'value': '"Require Manual"',
                                            'format': format1})          
       
        worksheet.conditional_format('A2:K1048576', {'type': 'cell',
                                            'criteria': 'equal to',
                                            'value': '"NEG"',
                                            'format': format1})
            
        worksheet.conditional_format('A2:K1048576', {'type': 'cell',
                                            'criteria': 'equal to',
                                            'value': '"POS"',
                                            'format': format2})
        
        worksheet.conditional_format('A2:K1048576', {'type': 'cell',
                                            'criteria': 'equal to',
                                            'value': '"NA"',
                                            'format': format3})

        
        writer.close()
        
print '-----------------'
print 'Analysis Complete'
print '-----------------'
raw_input()
    
