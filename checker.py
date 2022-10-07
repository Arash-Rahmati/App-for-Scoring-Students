#absent finder / reading participat finder / conversation finder / show full insight for each class

from openpyxl import load_workbook
import re
import matplotlib.pyplot as pyplot
import arabic_reshaper
from bidi.algorithm import get_display
import seaborn as sns



alphabets =['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V']

def num_of_absences (index): #
    absent=0   
    for i in range(20):
        if(wss[alphabets[i]+str(index)].value=='a'): absent+=1
    return absent


def find_dates_of_absences (index): #
    list=[]
    for i in range(20):
        if(wss[alphabets[i]+str(index)].value=='a'): list.append(wss[alphabets[i]+'32'].value)
    return list

def readingTrue(index):
    if(wss['AH'+str(index)].value==None): return 0
    else: return 1

def ifexist(index):
    if(wss['B'+str(index)].value==None): return -1
    else: return wss['B'+str(index)].value
        

#---------------------------------------------------MAIN PROGRAM---------------------#
while 1<2:
    chh = input ('Which class 1, 2, 3, 4, 5, 6: ')
    chhh = input ('Press 1: For Absent  Press 2: For Reading Finder  Press 3: Total Class Activity   Press 4: Total Class Activity Stu num   Press 5: Histogram  ')
    wb=load_workbook('Class.xlsx',data_only=True,read_only=True)
    wss=wb[chh]
    if(int(chhh)==1):
        for i in range(20):
            k=num_of_absences(i+2)
            rightformatname1= arabic_reshaper.reshape(wss['B'+str(i+2)].value)
            rightformatname1= get_display(rightformatname1)                
            if (k>=4): print(rightformatname1, 'numbers: ',k, 'dates: ',find_dates_of_absences(i+2), '\n')

    elif (int(chhh)==2):
        for i in range(20):
            if(readingTrue(i+2)==0): 
                temp=wss['B'+str(i+2)].value
                if(temp!=None): 
                    temp=arabic_reshaper.reshape(wss['B'+str(i+2)].value)
                    temp=get_display(temp)
                    print(temp)
                

    elif (int(chhh)==3): 
        names=[]
        index=[]
        score=[]
        for i in range(25):
            studentname= ifexist(i+2)
            if(studentname!=-1): 
                studentname = arabic_reshaper.reshape(studentname)
                studentname = get_display(studentname)
                names.append(studentname)
                score.append(wss['AW'+str(i+2)].value)
                index.append(len(index)+1)
                
        pyplot.title('Students Activity in All Sessions Together: ' )
        pyplot.barh(index,score,tick_label=names,color='red')
        pyplot.subplots_adjust(0.3,bottom=0.05, right=None, top=0.92, wspace=None, hspace=None)
        pyplot.show()

    elif (int(chhh)==4): 
        names=[]
        index=[]
        score=[]
        for i in range(25):
            studentname= ifexist(i+2)
            if(studentname!=-1): 
                studentname = wss['AX'+str(i+2)].value
                names.append(studentname)
                score.append(wss['AW'+str(i+2)].value)
                index.append(len(index)+1)
                
        pyplot.title('Students Activity in All Sessions Together: ' )
        pyplot.barh(index,score,tick_label=names,color='red')
        pyplot.subplots_adjust(0.3,bottom=0.05, right=None, top=0.92, wspace=None, hspace=None)
        pyplot.show()

    elif (int(chhh)==5): 
                score=[]
                for i in range(25):
                    studentname= ifexist(i+2)
                    if(studentname!=-1): 
                        score.append(wss['AW'+str(i+2)].value)
                
                sns.displot(score, color='r')
                sns.displot(score, kind='kde', color='b')
                pyplot.show()       
    wb.close()


