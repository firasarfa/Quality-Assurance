import pandas as pd 
import numpy as np 
import openpyxl 
import xlsxwriter
import matplotlib.pyplot as plt 
import statistics
import seaborn as sns
import math
from xlwings.constants import ColorIndex
import xlwings as xw
from main import auswertung



frequencies=[750,755,760,765,770,775,780,785,790,795,800,805,810,815,820,825,830,835,840,845,850,855,860,865,870,875,880,
    885,890,895,900,905,910,915,920,925,930,935,940,945,950,955,960,965,970,975,980,985,990,995,1000,1005,1010,1015,1020,1025,
    1030,1035,1040,1045,1050,1055,1060,1065,1070,1075,1080,1085,1090,1095,1100,1105,1110,1115,1120,1125,1130,1135,1140,1145,
    1150,1155,1160,1165,1170,1175,1180,1185,1190,1195,1200]



def read_file(filepath):
    if filepath.endswith('.csv'):
        return pd.read_csv(filepath)
    elif filepath.endswith('.json'):
        return pd.read_json(filepath)
    elif filepath.endswith('.xlsx'):
        return pd.read_excel(filepath)
    else:
        raise ValueError("File format not supported.")



#### this function returns the values of the measurement on a desired Frequency that the User defines 
def val_on_hz(file_path ,output_file,hz1,hz2,hz3,hz4):


    path = file_path

    df= read_file(path)
    
    first_col=list(df['Tagformance measurement file'])
    ##### we look for the Frequencies that the user gave us in the GUI (check gui) 
    hz_ind=[]
    hz1_ind=[i for i in range(len(first_col)) if first_col[i]==hz1 and i>460]
    hz2_ind=[i for i in range(len(first_col)) if first_col[i]==hz2 and i>460]
    hz3_ind=[i for i in range(len(first_col)) if first_col[i]==hz3 and i>460]
    hz4_ind=[i for i in range(len(first_col)) if first_col[i]==hz4 and i>500]
    

    #we create a list of sublists with the frequencies locations (index)
    hz_ind.append(hz1_ind)
    hz_ind.append(hz2_ind)
    hz_ind.append(hz3_ind)
    hz_ind.append(hz4_ind)
    print(hz_ind)
    print(len(hz1_ind))
    print(len(hz2_ind))
    print(len(hz3_ind))
    print(len(hz4_ind))
    
    d={}
    
    werte6=[]
    

    # iterate through the sublists in the original list and get the values if it'S POTF
    for sublist in hz_ind:
        new_sublist=[df.iat[i,6] for i in sublist]
        werte6.append(new_sublist)
    
    
    werte8=[]

    # iterate through the sublists in the original list and get the values if it'S TTRF
    for sublist in hz_ind:
        new_sublist=[df.iat[i,8] for i in sublist]
        werte8.append(new_sublist)
    
    # crearte a dataFrame with the frequencies and measurements 
    d={f"{hz1} MHZ":werte6[0] ,f"{hz2} MHZ":werte6[1] ,f"{hz3} MHZ":werte6[2] ,f"{hz4} MHZ":werte6[3] , }
    val_df6 = pd.DataFrame(data=d)
    d2={f"{hz1} MHZ":werte8[0] ,f"{hz2} MHZ":werte8[1] ,f"{hz3} MHZ":werte8[2] ,f"{hz4} MHZ":werte8[3] , }
    val_df8 = pd.DataFrame(data=d2)

    #import the table to the exel created in the Auswertung Function in the main.py file    
    with pd.ExcelWriter(output_file , mode='a' ,engine='openpyxl' , if_sheet_exists="overlay" ) as writer:
        val_df6.to_excel(writer, sheet_name= "POTF" , startcol= 9 ,index=False)
        val_df8.to_excel(writer, sheet_name= "TRRF" , startcol= 9 ,index=False)
        

#val_on_hz("fff.xlsx" ,"hffhhf.xlsx" , 800 ,900 ,1000,1100 )
##
def tot_avg(file_path):

    def Average(lst):
        return sum(lst) / len(lst)
    
    

    df= read_file(file_path)
    
    first_col=list(df['Tagformance measurement file'])
    
    frequencies=[750,755,760,765,770,775,780,785,790,795,800,805,810,815,820,825,830,835,840,845,850,855,860,865,870,875,880,
    885,890,895,900,905,910,915,920,925,930,935,940,945,950,955,960,965,970,975,980,985,990,995,1000,1005,1010,1015,1020,1025,
    1030,1035,1040,1045,1050,1055,1060,1065,1070,1075,1080,1085,1090,1095,1100,1105,1110,1115,1120,1125,1130,1135,1140,1145,
    1150,1155,1160,1165,1170,1175,1180,1185,1190,1195,1200]
    
    indexes=[]
    #we get the measurements at each frequency in the List so we can make the average at each frequency individually 
    for hzs in frequencies:
        new_sublist=[i for i in range(len(first_col)) if first_col[i]==hzs and i > 475]
        indexes.append(new_sublist)

    #print(indexes)
    #Get every measurement at a certian frequence and put them in seperate sublists POTF
    werte6=[]
    for sublist in indexes:
        new_sublist=[df.iat[i,6] for i in sublist]
        werte6.append(new_sublist)
    print(werte6)
    
    #Get every measurement at a certian frequence and put them in seperate sublists TTRF
    werte8=[]
    for sublist in indexes:
        new_sublist=[df.iat[i,8] for i in sublist]
        werte8.append(new_sublist)

    #deduce the average of measurement at each frequency for POTF
    averages6=[]
    for sublist in werte6:
        average6=round(Average(sublist),4)
        averages6.append(average6)
    #deduce the average of the measurements at each frequency TTRF
    
    averages8=[]
    for sublist in werte8:
        average8=round(Average(sublist),4)
        averages8.append(average8)
    
    #deduce the standard deviation of the measurements at each frequency POTF
    st_abws6=[]
    for sublist in werte6:
        st_dv=round(statistics.pstdev(sublist),4)
        st_abws6.append(st_dv)
    #deduce the standard deviation of the measurements at each frequency TTRF
    st_abws8=[]
    for sublist in werte8:
        st_dv=round(statistics.pstdev(sublist),4)
        st_abws8.append(st_dv)
    print(st_abws6)


    #Maximum tolerance limit POTF 
    obere6=[]
    #Minimum tolerance limit POTF 
    untere6=[]
    for g in range (0,len(averages6)):
        ob=round(averages6[g]+3*st_abws6[g],4)
        ut=round(averages6[g]-3*st_abws6[g],4)
        obere6.append(ob)
        untere6.append(ut)
    obere8=[]
    untere8=[]
    for n in range (0,len(averages8)):
        ob=round(averages8[n]+3*st_abws8[n],4)
        ut=round(averages8[n]-3*st_abws8[n],4)
        obere8.append(ob)
        untere8.append(ut)
    
    
    #print(obere)
    #print(untere)
        

    #create a data frame with the measurements' average , Standarddeviation at each Frequency for TTRF POTF
    d={ "Frequenzen": frequencies , "Durschnitt": averages6 , "Standard Abweichung":st_abws6 }
    df=pd.DataFrame(data=d)
    d2={ "Frequenzen": frequencies , "Durschnitt": averages8 , "Standard Abweichung":st_abws8 }
    df2=pd.DataFrame(data=d2)
    #print(averages)
    #print(len(averages))
    #print(len(frequencies))
    #chart=plt.plot(frequencies,averages6)
    #plt.show()
    
    
    #import the dataframes in the exel file we created in the Auswertungs function
    #with pd.ExcelWriter(output_file , mode='a' ,engine='openpyxl' , if_sheet_exists="overlay" ) as writer:
     #   df.to_excel(writer, sheet_name= "POTF" , startcol= 13 ,index=False)
      #  df2.to_excel(writer, sheet_name= "TRRF" , startcol= 13 ,index=False)

    return     (averages6,averages8 ,obere6 , untere6, werte6 , werte8 , st_abws8 , st_abws6 , obere8 , untere8)



def get_samples(file_path ):
    
    
    path = file_path
 
    df= read_file(path)
    ##Search for the word "Frequency" in the first column  as an indicator for the beginning of our Table
    first_col=list(df['Tagformance measurement file'])
    last=int (len(first_col))+7
    st_fi=[i for i in range(len(first_col)) if first_col[i]=="Frequency (MHz)"]
    st_fi.append(last)
    
    ### Lists we used in the function
   
    
    #d2={"Automatische":["Theoretical","forward"] , "UHF":["Read"] , "Auswertung":["Range"]}
    #df3=pd.DataFrame(data=d2)
    def get_sample(c):
        cf=df.iat[c-1,1]
        return cf 
    print(get_sample(486))
    
    ##After Looking for the word "Frequency" as an indicator for the beginning of our Table we try to find the boundaries of every table
    # Start just under the "Frequency" Cell in the exel table and the end 7 cells before the index or Cell of the next Frequency "Cell" 
    starts=[]
    ends=[]
    samples_names=[]
    ##get the index of where every table starts and ends  
    for p in range(1,len(st_fi)-1):
            starts.append(st_fi[p]+4)
            ends.append(st_fi[p+1]-7)
    for ind in range(2,len(st_fi)-1):
        samples_names.append(get_sample(st_fi[ind]))
    return(samples_names)


#######
###This is a function that represents all the mesurements graphically and adds up the maximum and the minimum tolerance 


def all_graphs(cc):
    a , b , c ,d , f , e , g ,h , l , m = tot_avg("fff.xlsx")
    
    df1=pd.DataFrame(f , columns= get_samples("fff.xlsx")  )
    df2=pd.DataFrame(e , columns= get_samples("fff.xlsx") )
    print(df1)
    #df['Samples'] = get_samples("fff.xlsx")
    #df["POTF Mittelwert"] = a
    #df["TTRF Mittelwert"] = b
    #df["Obere"] = c
    #df["untere"] = d

    #df.plot()
    #df.plot(y=['Obere', 'Untere'], style='--', linewidth=2.0)
    #plt.plot(frequencies, a ,linestyle='dotted')
    fig, (ax1, ax2) = plt.subplots(1, 2)
    ax1.plot(df1)
    ax1.plot(a, '--', linewidth=2.0 , label="POTF Mittelwert" , color='black' )
    ax1.plot(c, '-.', linewidth=2.0 , label="+3 sigma" , color='#5A5A5A' )
    ax1.plot(d,'-.', linewidth=2.0 , label="-3 sigma" , color='#694337')
    ax1.set_title("POTF")
    #ax1.legend(ncol=1, bbox_to_anchor=(1.05,1), loc='upper center') 
    #ax1.figure.savefig('ax1_plot.jpg')
    #ax1.clear()
    
    df2.plot(ax=ax2)   
    ax2.plot(b, '--', linewidth=2.0 , label="TTRF Mittelwert" , color='black' )
    ax2.plot(l, '-.', linewidth=2.0 , label="+3 sigma" , color='#5A5A5A' )
    ax2.plot(m,'-.', linewidth=2.0 , label="-3 sigma" , color='#694337')
    ax2.set_title("TRRF")
    #ax2.legend(ncol=1, bbox_to_anchor=(1.05,1), loc='upper center')
    #ax2.figure.savefig('ax2_plot.jpg')
    #plt.legend()

    #ax2.clear()
    
    plt.legend(ncol=2, bbox_to_anchor=(1.05,1), loc='upper center') 
    #x2.legend(fontsize='small')
    plt.savefig(cc)

    plt.show()
    

#all_graphs('vvvv')

def tableViz(out):
    
    a , b , c ,d , f , e , g ,h , l,m  = tot_avg(out)
    
    df1=pd.DataFrame(f , columns= get_samples(out))
    df2=pd.DataFrame(e , columns= get_samples(out))
    print(df1)
    refrence= pd.Series(a)
    df1['St_Ab']=h
    df1['reference']=a
    df2['St_Ab']=g
    df2['reference']=b
    df1.fillna(0)
    df2.fillna(0)
    with pd.ExcelWriter(out , mode='a' ,engine='openpyxl' , if_sheet_exists="overlay" ) as writer:
        df1.to_excel(writer, sheet_name= "TRRF MAP" , index=False)
        df2.to_excel(writer, sheet_name= "PTTF MAP" , index=False)

    



####this function helps us to classify the measurement and give the a color on excel to represent how good or bad the measurement is compared 
####to the average and other limits
def xlfunc2(original,ref,out,cl,num):
    

    ### prepare the refrence file that the user gave us   
    if num==1:

        a , b , c ,d , f , e , g ,h ,l , m   = tot_avg(ref)
        df1=pd.DataFrame(f , columns= get_samples(ref)  )
        df2=pd.DataFrame(e , columns= get_samples(ref) )
        
    #if the user didn't insert a refrence file than we continue with the analysed file 
    if num==2:

        a , b , c ,d , f , e , g ,h , l , m  = tot_avg(original)
        df1=pd.DataFrame(f , columns= get_samples(original)  )
        df2=pd.DataFrame(e , columns= get_samples(original) )
    
    ##
    ##add the refrence columns (average and standard deviation )
    df1.insert(0, 'Standard Deviation', h)
    df2.insert(0, 'Standard Deviation', g)
    df1['reference']=a
    df1.fillna(0)
    df2['reference']=b
    df2.fillna(0)


    #transfer the       
    print(df1)
    with pd.ExcelWriter(out , mode='a' ,engine='openpyxl' , if_sheet_exists="overlay" ) as writer:
            df1.to_excel(writer, sheet_name= 'TRRF MAP'  ,index=False)
            df2.to_excel(writer, sheet_name= 'POTF MAP'  , index=False )
    
    


    #Select the TRRF Sheet 
    # Connect to the Excel file
    wb = xw.Book(out)
    if cl==6:
    # Select the sheet with the data
        sheet = wb.sheets['TRRF MAP']
        

        ref_col = df1.iloc[:, -1]
        st_abw=df1.iloc[:, 0]
        # Loop through each column in the DataFrame (starting with the second column)
        for col_num in range(1, len(df1.columns)-1):

                # Select the column to compare
            col = df1.iloc[:, col_num]
            

                
                # Loop through each cell in the column
            for row_num, cell_value in enumerate(col):

                # Get the value of the corresponding cell in the reference column
                ref_value = ref_col[row_num]
                st_ab=st_abw[row_num]

                if not math.isnan(col[row_num]) and not math.isnan(ref_col[row_num]):    
                        # Compare the cell value to the reference value
                    if  ref_value-st_ab <cell_value < ref_value+st_ab:
                            # Set the cell color to green
                        sheet.range((row_num + 2, col_num + 1)).color = (76, 175, 80) 
                            # Add 2 to row_num to account for the header row
                        
                    elif ref_value-2*st_ab <cell_value < ref_value+2*st_ab: 
                        # Set the cell color to yellow    
                        sheet.range((row_num + 2, col_num + 1)).color = (255,255,0)
                            # Add 2 to row_num to account for the header row
                        
                    elif ref_value-3*st_ab <cell_value < ref_value+3*st_ab:
                            # Set the cell color to orange
                        sheet.range((row_num + 2, col_num + 1)).color = (236, 151, 6)
                            # Add 2 to row_num to account for the header row
                        
                        
                    elif ref_value-3*st_ab > cell_value  or cell_value > ref_value+3*st_ab:
                            # Set the cell color to red
                        sheet.range((row_num + 2, col_num + 1)).color = ((255,0,0))
    
    

    
    if cl==8:
    # Select the sheet with the data
        sheet = wb.sheets['POTF MAP']
    
        ref_col = df2.iloc[:, -1]
        st_abw=df2.iloc[:, 0]
        for col_num in range(1, len(df2.columns)-1):

                # Select the column to compare
            col = df2.iloc[:, col_num]
            

                
                # Loop through each cell in the column
            for row_num, cell_value in enumerate(col):

                # Get the value of the corresponding cell in the reference column
                ref_value = ref_col[row_num]
                st_ab=st_abw[row_num]

                if not math.isnan(col[row_num]) and not math.isnan(ref_col[row_num]):    
                        # Compare the cell value to the reference value
                    if  ref_value-st_ab <cell_value < ref_value+st_ab:
                            # Set the cell color to green
                        sheet.range((row_num + 2, col_num + 1)).color = (76, 175, 80) 
                            # Add 2 to row_num to account for the header row
                        
                    elif ref_value-2*st_ab <cell_value < ref_value+2*st_ab: 
                        # Set the cell color to yellow    
                        sheet.range((row_num + 2, col_num + 1)).color = (255,255,0)
                            # Add 2 to row_num to account for the header row
                        
                    elif ref_value-3*st_ab <cell_value < ref_value+3*st_ab:
                            # Set the cell color to orange
                        sheet.range((row_num + 2, col_num + 1)).color = (236, 151, 6)
                            # Add 2 to row_num to account for the header row
                        
                        
                    elif ref_value-3*st_ab > cell_value  or cell_value > ref_value+3*st_ab:
                            # Set the cell color to red
                        sheet.range((row_num + 2, col_num + 1)).color = ((255,0,0))




    wb.save(out)

    # Close the workbook
    wb.close()
            
        
#xlfunc2("fff.xlsx","fff.xlsx","xxxxx.xlsx",6,2)
#xlfunc2("fff.xlsx","fff.xlsx","xxxxx.xlsx",8,2)
#xlfunc()

'''
a , b , c ,d , f , e , g ,h , l, m  = tot_avg(out)
a: the average of measurement at each frequency for POTF
b: the average of measurement at each frequency for TRRF
c:#Maximum tolerance limit POTF 
d:#Minimum tolerance limit POTF 
e: measurement at a certian frequence and put them in seperate sublists POTF
f: measurement at a certian frequence and put them in seperate sublists TRRF
g: the standard deviation of the measurements at each frequency TRRF
h: the standard deviation of the measurements at each frequency POTF
l:#Maximum tolerance limit TRRF
m:#Minimum tolerance limit TRRF

'''
