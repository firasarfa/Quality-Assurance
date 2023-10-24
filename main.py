import pandas as pd 
import numpy as np 
import openpyxl 
import xlsxwriter
import seaborn as sns
import matplotlib.pyplot as plt 


def adjustable_range(start, stop=None, step=None):
        if not isinstance(start, int):
            raise TypeError('start')
        if stop is None:
            start, stop = 0, start
        elif not isinstance(stop, int):
            raise TypeError('stop')
        direction = stop - start
        positive, negative = direction > 0, direction < 0
        if step is None:
            step = +1 if positive else -1
        else:
            if not isinstance(step, int):
                raise TypeError('step')
            if positive and step < 0 or negative and step > 0:
                raise ValueError('step')
        if direction:
            valid = (lambda a, b: a < b) if positive else (lambda a, b: a > b)
            while valid(start, stop):
                message = yield start
                if message is not None:
                    if not isinstance(message, int):
                        raise ValueError('message')
                    stop = message
                start += step


##### This function is there to enble us to read differnt data types 
## such as xlsx , csv , json
def read_file(filepath):
    if filepath.endswith('.csv'):
        return pd.read_csv(filepath)
    elif filepath.endswith('.json'):
        return pd.read_json(filepath)
    elif filepath.endswith('.xlsx'):
        return pd.read_excel(filepath)
    else:
        raise ValueError("File format not supported.")

#### Look at line 216
def check_filename (name):
    lower_name=name.lower()
    if "500h" in lower_name and "init" in lower_name:
        return True
    if "168h" in lower_name and "init" in lower_name:
        return True
    elif "init" in lower_name:
        return True 
    else:
        return False
    



def auswertung(file_path  , output_file , column_id ):
    
    
    path = file_path
 
    df= read_file(path)
    ##Search for the word "Frequency" in the first column  as an indicator for the beginning of our Table
    first_col=list(df['measurement file'])
    last=int (len(first_col))+7
    st_fi=[i for i in range(len(first_col)) if first_col[i]=="Frequency (MHz)"]
    st_fi.append(last)
    
    ### Lists we used in the function
    peaks=[]
    ind_peaks=[]
    st_peaks=[]
    nd_peaks=[]
    rd_peaks=[]
    th_peaks=[]
    st_ind_peaks=[]
    nd_ind_peaks=[]
    rd_ind_peaks=[]
    th_ind_peaks=[]
    tables=[]
    hz_list=[]
    #d2={"Automatische":["Theoretical","forward"] , "UHF":["Read"] , "Auswertung":["Range"]}
    #df3=pd.DataFrame(data=d2)
    def get_sample(c):
        cf=df.iat[c-1,1]
        return cf 
    print(get_sample(486))
    
    
    def Average(lst):
        return sum(lst) / len(lst)
    
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

    ## if we remark a trend it means that there is a peak and we save the value of the peak and it's location (cell)
    ## a Peak/dip is when a value is bigger/smaller of the previous and the next value depending on the Column (TRRF or POTF)  
    if column_id==8:
        for p in range(1,len(starts)):
            for i in range(starts[p],ends[p]):          
                if df.iat[i-1,8] < df.iat[i,8] and df.iat[i,8] > df.iat[i+1,8] :
                        peaks.append(df.iat[i,8])
                        ind_peaks.append(i)
    if column_id==6:
        for p in range(1,len(starts)):
            for i in range(starts[p],ends[p]-3):          
                if df.iat[i-1,6] > df.iat[i,6] and df.iat[i,6] < df.iat[i+1,6]  :
                    peaks.append(df.iat[i,6])
                    ind_peaks.append(i)


    #from the location we deduce the frequency corresponding to the Dip/Peak where it happened 
    for k  in range (0,len(ind_peaks)):
        for z in range (2,len(st_fi)):
            
            if ind_peaks[k]<st_fi[z] and ind_peaks[k]>st_fi[z-1]  :
                print(f"{peaks[k]} ist in der Tabelle {z} mit der frequenz {df.iat[ind_peaks[k],0]} ") 
                hz_list.append(df.iat[ind_peaks[k],0])
                tables.append(z)
    print(tables)
    lists=[]
    p_tables=[]
    

    
    print(len(hz_list))
    print(len(peaks))
    print(len(tables))
    ## we remarked peaks and dips happening in the range (980 to 1035) happenning for no reason (it's not supposed to have a peak there)
    ## we eliminate the values happennig in that intervall by First looking which peaks exactly were detected in that intervall

    parasite_ind=[]
    
    for x in range(0,len(hz_list)):
        if 980<hz_list[x]<1036 : 
            parasite_ind.append(x)


    new_parasite=[]
    for p in parasite_ind :
        if p not in new_parasite:
            new_parasite.append(p)

    #Second delete the correspoding Frequency and value
    for g in reversed(new_parasite):
        del peaks[g]
        #peaks.remove(peaks[g])
        del hz_list[g]
        #hz_list.remove(hz_list[g])
        del tables[g]


    
    ### we used a variable table to keep track of which value is corresponding to which sample
    current_number=tables[0]
    current_sublist_peaks=[]
    current_sublist_hz=[]
    result_peaks=[]
    result_hz=[]
    ### using the variable Table we put the frequencies and the dips of each sample in seperate sublists  
    for i,x in enumerate(tables):
        if x==current_number:
            current_sublist_peaks.append(peaks[i])
            current_sublist_hz.append(hz_list[i])
        else:
            result_peaks.append(current_sublist_peaks)
            current_sublist_peaks=[peaks[i]]
            result_hz.append(current_sublist_hz)
            current_sublist_hz=[hz_list[i]]
            current_number=x
        
    result_peaks.append(current_sublist_peaks)
    result_hz.append(current_sublist_hz)
    
    print(result_peaks)
    print(result_hz)

    ### after a while in the Test some samples may lose a peak or two so we Auto-Fill them with 0 values
    ####Very Important :sometimes the code goes in an infinite loop while compiling and that's why there is a check name function 
    ###to fix it insert a (for) instead of the (if) when check_filename(filepath)==False . it has something to do with missing dips and peaks 
    ### it happens after a long time in Klima reliability Test


    #if check_filename(filepath) == True :
    if all(len(sublist)==4 for sublist in result_peaks): 
        for n in range (0,len(result_peaks)-1):
            if  len(result_peaks[n])<4:
                result_peaks[n].append(0)
                result_hz[n].append(0)

    




    #### I Kept those print lines to help you look for the bugs if there's any  
    print(result_peaks)
    print(result_hz)
    print(len(tables))
    print(len(hz_list))
    print(len(peaks))
    print(tables)
    print(hz_list)
    print(peaks)


    MAXS=[]
    MINS=[]
    MAX_ID=[]
    MIN_ID=[]
    auswert_2p=[]
    auswert_3p=[]
    auswert_4p=[]
    max=[]
    min=[]
    average=[]
    
    # We catigorize the values to First second third and fourth Peaks so we can analyze them individually 
    
    for m in range(0,len(result_hz)):
            st_peaks.append(result_peaks[m][0])
            st_ind_peaks.append(result_hz[m][0])
            nd_peaks.append(result_peaks[m][1])
            nd_ind_peaks.append(result_hz[m][1])
            rd_peaks.append(result_peaks[m][2])
            rd_ind_peaks.append(result_hz[m][2])
            th_peaks.append(result_peaks[m][3])
            th_ind_peaks.append(result_hz[m][3])             
    
    ### Create a dataFrame to with the Peaks and the FRequencies they accured at  

    d={"Sample": samples_names,"1st Peak Hz":st_ind_peaks,"1st Peak":st_peaks , "2nd Peak Hz":nd_ind_peaks,"2nd Peak":nd_peaks ,"3rd Peak Hz":rd_ind_peaks, "3rd Peak": rd_peaks , "4th Peak Hz":th_ind_peaks,"4th Peak": th_peaks}


    df2=pd.DataFrame(data=d)
    #print(df2)
    #sns.scatterplot(x="Sample" , y="1st Peak" , data=df2 , hue='1st Peak Hz')
    #plt.show()

    ###we deduce some informations based on those values like(MAX ; MIN ; AVG ) 

    average.append(df2["1st Peak"].mean())
    average.append(df2["2nd Peak"].mean())
    average.append(df2["3rd Peak"].mean())
    average.append(df2["4th Peak"].mean())
    min.append(df2["1st Peak"].idxmin())
    min.append(df2["2nd Peak"].idxmin())
    min.append(df2["3rd Peak"].idxmin())
    min.append(df2["4th Peak"].idxmin())
    max.append(df2["1st Peak"].idxmax())
    max.append(df2["2nd Peak"].idxmax())
    max.append(df2["3rd Peak"].idxmax())
    max.append(df2["4th Peak"].idxmax())
    for x in range (0,len(max)):
        MAX_ID.append(df2.iat[max[x], 0])
        MAXS.append(df2.iat[max[x], 2*(x+1)])
    for u in range (0,len(min)):
        MIN_ID.append(df2.iat[min[u],0])
        MINS.append(df2.iat[min[u],2*(u+1)])
    d2={"Maximum ID":MAX_ID, "Maximum Wert":MAXS , "Minimum ID":MIN_ID , "Minimum":MINS , "Durschnitt":average  }
    print(MAX_ID)
    print(MINS)
    print(MAXS)
    
    print(min)
    df3=pd.DataFrame(d2 , index=["1st Peak" , "2nd Peak" , "3rd Peak" , "4th Peak"])
    
    ##### create an exel file and Import the created dataframes in it    

    if column_id==8:
        with pd.ExcelWriter(output_file , mode='a' ,engine='openpyxl' , if_sheet_exists="overlay" ) as writer:
            df2.to_excel(writer, sheet_name= "TRRF"  ,index=False)
            df3.to_excel(writer, sheet_name= "TRRF" , startcol= 18 )
    if column_id==6:
        writer=pd.ExcelWriter(output_file , engine='xlsxwriter')
        #with pd.ExcelWriter(output_file , mode='a' ,engine='openpyxl' , if_sheet_exists="overlay" ) as writer:
        df2.to_excel(writer, sheet_name= "POTF"  ,index=False)
        df3.to_excel(writer, sheet_name= "POTF" , startcol= 18 )
        writer.save()
    return(samples_names)
    


