import tkinter as tk
import customtkinter
from tkinter import Tk, messagebox
from main import adjustable_range,auswertung,read_file
from Visualization import val_on_hz,tot_avg,xlfunc2,all_graphs
import pandas as pd
import matplotlib.pyplot as plt 
from tkinter import filedialog


customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

root = customtkinter.CTk()
root.geometry("1500*1500")
root.title('Auswertung')

eingabe_werte=[]
frequencies=[750,755,760,765,770,775,780,785,790,795,800,805,810,815,820,825,830,835,840,845,850,855,860,865,870,875,880,
    885,890,895,900,905,910,915,920,925,930,935,940,945,950,955,960,965,970,975,980,985,990,995,1000,1005,1010,1015,1020,1025,
    1030,1035,1040,1045,1050,1055,1060,1065,1070,1075,1080,1085,1090,1095,1100,1105,1110,1115,1120,1125,1130,1135,1140,1145,
    1150,1155,1160,1165,1170,1175,1180,1185,1190,1195,1200]
    
#### we ask the user for the input file , output file , frequencies and the graph name 
### we use the functions that we made to do the analysis and to show a graphical representation of the findings
### the generated graph will be saved as a picture 

def submit():
    no_ref=False
    file_path=path.get()
    ref_file_path=ref_path.get()
    outputfile=output.get()
    frq1=(fq1.get())
    frq2=(fq2.get())
    frq3=(fq3.get())
    frq4=(fq4.get())
    
    if not frq1.isdigit() or not frq2.isdigit() or not frq3.isdigit() or not frq4.isdigit():
        messagebox.showerror("Error", "Please enter valid numbers.")
        return
    frq1=int(fq1.get())
    frq2=int(fq2.get())
    frq3=int(fq3.get())
    frq4=int(fq4.get())
        # Check if entries are within the valid range
    if not (750 <= frq1 <= 1200) or not (750 <= frq2 <= 1200) or not (750 <= frq3 <= 1200) or not (750 <= frq4 <= 1200):
        messagebox.showerror("Error", "Please enter numbers between 750 and 1200.")
        return
    
    graph=grnm.get()

    eingabe_werte.append(file_path)
    eingabe_werte.append(outputfile)
    eingabe_werte.append(frq1 )
    eingabe_werte.append(frq2 )
    eingabe_werte.append(frq3 )
    eingabe_werte.append(frq4 )
    eingabe_werte.append(graph)
    eingabe_werte.append(ref_file_path)

    auswertung(eingabe_werte[0],eingabe_werte[1],6)
    auswertung(eingabe_werte[0],eingabe_werte[1],8)


    if ref_file_path == "":
        result = messagebox.askyesno("Confirmation", "Do you want to continue without a reference file?")
        if result:
            xlfunc2(eingabe_werte[0],eingabe_werte[7],eingabe_werte[1],6,2)
            xlfunc2(eingabe_werte[0],eingabe_werte[7],eingabe_werte[1],8,2)
        else:
            xlfunc2(eingabe_werte[0],eingabe_werte[7],eingabe_werte[1],6,1)
            xlfunc2(eingabe_werte[0],eingabe_werte[7],eingabe_werte[1],8,1)
            

    


    #auswertung(eingabe_werte[0],eingabe_werte[1])
    val_on_hz(eingabe_werte[0],eingabe_werte[1],eingabe_werte[2],eingabe_werte[3],eingabe_werte[4],eingabe_werte[5])
    all_graphs(eingabe_werte[6])
    #tot_avg(eingabe_werte[0])
    a , b , c ,d , f , e , g ,h , l ,m= tot_avg(eingabe_werte[0])



def browse_file():
    file_path = filedialog.askopenfilename()
    path.delete(0, tk.END)
    path.insert(tk.END, file_path)

## Those are the entry boxes with their labels . 
#there is a box for file that we want to analyze , the desired location and the name of the result file (outputfile)
#there is a box for the name of the image file created
#and boxes for frequencies we want to control 


frame = customtkinter.CTkFrame(master=root)
frame.pack(padx=2 , pady=6 , fill="both" , expand=True)

label1= customtkinter.CTkLabel(master=frame, text="File Path")
label1.pack()
path=customtkinter.CTkEntry(master=frame , placeholder_text="File Path")
path.pack()

button1 = customtkinter.CTkButton(master=frame,text="Browse" , command=browse_file)
button1.pack(pady=12,padx=10)


label8= customtkinter.CTkLabel(master=frame, text="Refrence File Path")
label8.pack()
ref_path=customtkinter.CTkEntry(master=frame , placeholder_text="Ref File Path")
ref_path.pack()

label2 = customtkinter.CTkLabel(master=frame, text="Output File ")
label2.pack()
output=customtkinter.CTkEntry(master=frame , placeholder_text="Output File")
output.pack()

label3 = customtkinter.CTkLabel(master=frame, text="First Frequence")
label3.pack()
fq1=customtkinter.CTkEntry(master=frame , placeholder_text="eg 845")
fq1.pack()

label4 = customtkinter.CTkLabel(master=frame, text="Second Frequence")
label4.pack()
fq2=customtkinter.CTkEntry(master=frame , placeholder_text="eg 870")
fq2.pack()

label5 = customtkinter.CTkLabel(master=frame, text="Third Frequence")
label5.pack()
fq3=customtkinter.CTkEntry(master=frame , placeholder_text="eg 890")
fq3.pack()

label6 = customtkinter.CTkLabel(master=frame, text="Fourth Frequence")
label6.pack()
fq4=customtkinter.CTkEntry(master=frame , placeholder_text="eg 1150")
fq4.pack()

label7 = customtkinter.CTkLabel(master=frame, text="Graph Name")
label7.pack()
grnm=customtkinter.CTkEntry(master=frame , placeholder_text="eg Graph UCode9 Chrome  ")
grnm.pack()


button = customtkinter.CTkButton(master=frame,text="Eingeben" , command=submit)
button.pack(pady=12,padx=10)

root.mainloop()