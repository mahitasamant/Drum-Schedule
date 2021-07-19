from tkinter import *
from numpy.lib.utils import _lookfor_generate_cache
import pandas as pd
from tkinter import messagebox
from tkinter import filedialog
from tkinter.ttk import Combobox
import openpyxl as pxl
main_df=pd.DataFrame()
main_drum_df=pd.DataFrame()
jk_df=pd.DataFrame(columns=['Service Voltage','Type','Number of kits'])
sheet_df=pd.DataFrame(columns=['Sheet Name','Sheet Content'])
cablesizing=''
num=1
filename=''
project_no=''
mainwindow=Tk()
mainwindow.geometry('1380x680+5+5')
mainwindow.title('Cable Manager')
mainwindow.configure(bg='white')
heading=Label(mainwindow,bg='white', fg='Black',font=('arial', 30,'bold'),text='DRUM SCHEDULE MANAGER')
heading.pack()
def calculatedrum(df,type,max_len,Lot,Area,voltage,cabletype,size):
 global main_df
 global main_drum_df
 global num
 global project_no
 global filename
 global jk_df
 max_len=max_len 
 inputdata=[]
 Lot=Lot
 Area=Area
 voltage=voltage
 cabletype=cabletype
 size=size
 for ind in df.index:
    item=[]
    item.append(df['Cable Tag'][ind])
    item.append(df['Cable Length'][ind])
    inputdata.append(item)
 #print(inputdata)
 morethanmax=[]
 lessthanmax=[]
 drumfinal1=[]
 
 n=0
 for i in inputdata:
    current=i[1]
    if current>=max_len:
      r=current%max_len
      if r>0:
        m=[]
        m.append(i[0])
        m.append(r)
        lessthanmax.append(m)
      p=int((current-r)/max_len)
      for z in range(p):
          q=[]
          q.append(i[0])
          q.append(max_len)
          drumfinal1.append(q)
    if current<=max_len:
        lessthanmax.append(i)
 lessthanmax.sort(reverse=True)
 #print(drumfinal1)
 #print(lessthanmax)
 p=0
 dummy_lessthanmax=lessthanmax
 flag=False
 drumfinal2=[]
 ignore = ['not an empty list']
 iterator = 0
 completed=True
 while (completed):
    
    #skip cable lengths which are already considered
    if iterator not in ignore:
        currentitem=[]
        currentLength = lessthanmax[iterator][1]
        currentitem.append(list(lessthanmax[iterator]))
        
        ignore.append(iterator)
        currentDrum=[currentLength]
        #print(currentDrum)
        difference = max_len - sum(currentDrum)
        #print(difference)
        for index,value in enumerate(lessthanmax):
            #print(value[1])
            #print(currentDrum)
            #skip cable lengths which are already considered but missed inadvertently
            if index not in ignore:
                if value[1] <= difference and (value[1]+sum(currentDrum))<=max_len:
                    currentDrum.append(value[1])
                    currentitem.append(value)
                    #since this index is used now, add to ignore array for next iteration
                    ignore.append(index)
            #print(currentitem)
        drumfinal2.append(currentitem)
    
    #increment iterator only till the last element of the array
    if iterator < len(lessthanmax):
        iterator +=1
    
    #End the loop as all the cable lengths are done
    if  len(ignore)>len(lessthanmax):
        completed = False  

 drumfinal=drumfinal1+drumfinal2      
#print(drumfinal)
 #print(drumfinal)
 print("************") 
    # Creating an empty dictionary
 freq = {}
 f=[]
 for item in drumfinal:
       p=[]
       try:
         if item[1]>0:
           p.append(item)
           f.append(p)
       except:
           f.append(item)
 q=1  
 #print(f)       
 for y in f:
    y.insert(0,q)
    q=q+1
 dict={}
 for l2 in f:
    dict[l2[0]] = l2[1:]
 df2 = pd.DataFrame(list(dict.items()),columns = ['column1','column2']) 
 column_names = ["Serial Drum Number", "Type", "Cables","Drum Length",'Cable and lengths','Drum Number']
 
 df3 = pd.DataFrame(columns = column_names)
 #print(df3)
 ind=0
 for l2 in f:
     #print("l2:",l2)
     #print(l2)
     list_cabletag=[] 
     drum_length=0
     drumno=l2[0]
     cable_lengths=l2
     if drumno<10:
         drumno="00"+str(drumno)
     elif drumno<100:
         drumno="0"+str(drumno)
     else:
         drumno=str(drumno)
     
     for l2_item in l2:
         try:
          list_cabletag.append(l2_item[0])
          drum_length=drum_length+l2_item[1]
         except:
             continue
            
     cs=list_cabletag
     drumlen=drum_length
     actualdrumno=str(project_no)+'-'+str(Lot)+'-'+str(voltage)+'-'+str(size)+'-'+str(cabletype)+'-'+str(drumlen)+'-D'+str(drumno)
     df3.loc[len(df3.index)] = [drumno,type,cs,drumlen,cable_lengths,actualdrumno]
 #df.insert('type')
 """
 for index_1 in df3.index:
     if df3['Drumno'][index_1]<10:
         df3['Drumno'][index_1]="00"+str(df3['Drumno'][index_1])
     elif df3['Drumno'][index_1]<100:
         df3['Drumno'][index_1]="0"+str(df3['Drumno'][index_1])
     else:
         df3['Drumno'][index_1]=df3['Drumno'][index_1] 
   """         
 #print(df3)
 list6=[]
 for ind in df3.index:
    list13=[] 
    
    
    list6.append(type)
    for i in df3['Cables'][ind]:
        
        
        for h in df.index:
            if df['Cable Tag'][h]==i:
                df['Consolidated Drum Serial Number'][h]=str(df['Consolidated Drum Serial Number'][h])+","+str(df3['Serial Drum Number'][ind])
                df['Consolidated Drum Number'][h]=str(df['Consolidated Drum Number'][h])+","+str(df3['Drum Number'][ind])
               
              
 
 print(df)   
 sheet_df.loc[len(sheet_df.index)] = [num,"Drum-"+type]
 num=num+1
 name=type
 excel_book = pxl.load_workbook(filename)
 main_drum_df=pd.concat([main_drum_df,df3])
 with pd.ExcelWriter(filename, engine='openpyxl') as writer:
    
    writer.book = excel_book

    writer.sheets = {worksheet.title: worksheet for worksheet in excel_book.worksheets}

    df3.to_excel(writer, "Sheet"+str(num), index=False)

    writer.save()
 

 
 #print(df)
 for value in df.index:
     currentstring=str(df['Consolidated Drum Serial Number'][value])
     currentstring=currentstring[4:]
     df['Consolidated Drum Serial Number'][value]=currentstring
     currentstring2=str(df['Consolidated Drum Number'][value])
     currentstring2=currentstring2[4:]
     df['Consolidated Drum Number'][value]=currentstring2
 for f in df.index:
        jk_count=str(df['Consolidated Drum Serial Number'][f]).count(",") 
        df['Jointing Kits'][f]=jk_count
 main_df=pd.concat([main_df,df])
 print(main_df)   
 sum_jk=0
 for u in df.index:
     sum_jk=sum_jk+df['Jointing Kits'][u]

 jk_df.loc[len(jk_df.index)] = [voltage,type,sum_jk]
 #print(jk_df)
 sheet_df.loc[len(sheet_df.index)] = [num,"Cable-"+type]
 num=num+1
 with pd.ExcelWriter(filename, engine='openpyxl') as writer:
    
    writer.book = excel_book

    writer.sheets = {worksheet.title: worksheet for worksheet in excel_book.worksheets}

    df.to_excel(writer, 'Sheet'+str(num), index=False)

    writer.save()

  

#Function to input Lv1,lv2 and lv-dc ratings
def Take_inputLvfreq():
            global project_no
            try: 
                project_no = str(Lvcablefrequencyrating.get('1.0', END))
            except(Exception):
                messagebox.showinfo('ERROR','PLEASE CHECK THE INPUT PROVIDED')
            
        
def del_inputLvfreq():
            global project_no       
              
            Lvcablefrequencyrating.delete('1.0', END)
                                       
Lvcablefrequency=Label(cablesizing,fg='black',text='Enter Project Code:')
Lvcablefrequency.place(x=50, y=100)
Lvcablefrequencyrating = Text(cablesizing, height = 2,width = 25,bg = "light yellow")
Lvcablefrequencyrating.place(x=200,y=100)   
go11 = Button(cablesizing, height = 2,width = 20,text ="ENTER",command = Take_inputLvfreq)
go11.place(x=450,y=100) 
go111 = Button(cablesizing, height = 2,width = 20,text ="RESET",command = del_inputLvfreq)
go111.place(x=630,y=100)

def browseFiles():
    global filename
    filename = filedialog.askopenfilename(initialdir = "/",
                                          title = "Select a File",
                                          filetypes = (
                                                       ("all files",
                                                        "*.*"),("Text files",
                                                        "*.txt*"),))
         
browseButton_Excel1 = Button(cablesizing,width=20,fg='blue',text='IMPORT EXCEL FILE', command=browseFiles, font=('Times New Roman', 16, 'bold'))
browseButton_Excel1.place(x=12, y=620)
def calc_drum(): 
    global filename 
    messagebox.showinfo("showinfo", "Your file is being processed, Kindly do not close the window or open the files being processed")
    df1=pd.read_excel(filename,sheet_name='input' ,engine='openpyxl')
    df1 = df1.dropna(subset = ['Cable Tag'])
    #df_len=pd.read_excel(filename,sheet_name='drum_length_data',engine='openpyxl') 
    main_list=list(df1['Type'])
    myset = set(main_list)
    print(myset)
    
    for i in myset:
     print(i)
     for p in df1.index:
        df=df1.where(df1['Type']==i)
        
        type=i
        
     df = df.dropna(subset = ['Cable Tag'])
     max_len_list=list(df['Max Drum Length'])
     max_len=max_len_list[0]
     LOT_list=list(df['LOT'])
     Lot=LOT_list[0]
     Area_list=list(df['Area'])
     Area=Area_list[0]
     voltage_list=list(df['Service Voltage'])
     voltage=voltage_list[0]
     cabletype_list=list(df['Cable Type'])
     cabletype=cabletype_list[0]
     size_list=list(df['Cable Size'])
     size=size_list[0]

     calculatedrum(df,type,max_len,Lot,Area,voltage,cabletype,size)
    #print(main_df)
    
    excel_book = pxl.load_workbook(filename)
    #print(main_df)
    #print(sheet_df)

    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
    # Your loaded workbook is set as the "base of work"
      writer.book = excel_book

    # Loop through the existing worksheets in the workbook and map each title to\
    # the corresponding worksheet (that is, a dictionary where the keys are the\
    # existing worksheets' names and the values are the actual worksheets)
      writer.sheets = {worksheet.title: worksheet for worksheet in excel_book.worksheets}

   
      sheet_df.to_excel(writer,'Sheet Index',index=False)
      jk_df.to_excel(writer, 'Jointing Kits', index=False)
      main_df.to_excel(writer, 'Cable Schedule', index=False)
      main_drum_df.to_excel(writer, 'Drum Schedule', index=False)
   
      writer.save()

    messagebox.showinfo("showinfo", "Your file is ready.")

calccablesize1=Button(cablesizing, width=10,text='RUN SIZING',bg='aliceblue',fg='blue4',font=('Times New Roman', 16, 'bold'),command=calc_drum)
calccablesize1.place(x=307,y=620)
mainwindow.mainloop()
#backend

print(project_no)
